# =============================================================
#  simulate.ps1 - 동기화 프로토콜 알고리즘 검증
#
#  modSync / modUtil 의 핵심 로직(이스케이프, 체크섬, 파트 분할,
#  조각 수신, 병합 규칙)을 그대로 옮겨 놓고 SPEC 의 검증 항목 1~7 을
#  실제로 돌려 본다.
#
#  Excel 없이 알고리즘이 맞는지 확인하는 용도다.
#  VBA 문법·엑셀 연동은 modTest.RunAllTests 로 따로 확인해야 한다.
#
#  실행:  powershell -ExecutionPolicy Bypass -File build\simulate.ps1
# =============================================================

param(
    # 지정하면 HTML 버전과 대조할 기준 파일(fixture)을 이 폴더에 쓴다
    [string]$FixtureDir = ''
)

$ErrorActionPreference = 'Stop'

$PROTO_TAG = 'WSCH1'
$PART_MAX  = 1500

$FD_NOHEADER    = 0
$FD_BADCHECKSUM = 1
$FD_PARTIAL     = 2
$FD_COMPLETE    = 3
$FD_DUPLICATE   = 4

$FIELDS = @('ID','문서번호','제목','내용','등록일','회신기한','작성자','수정시각','삭제여부')

# --- modUtil.Esc ---------------------------------------------
function Esc-Vba([string]$s) {
    $t = $s
    $t = $t.Replace('\', '\\')
    $t = $t.Replace('|', '\p')
    $t = $t.Replace("`r", '')
    $t = $t.Replace("`n", '\n')
    $t = $t.Replace("`t", '\t')
    return $t
}

# --- modUtil.Unesc (한 문자씩 훑는 방식) ----------------------
function Unesc-Vba([string]$s) {
    $n = $s.Length
    if ($n -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $n) {
        $ch = $s[$i]
        if ($ch -eq '\' -and $i -lt ($n - 1)) {
            $nx = $s[$i + 1]
            switch ($nx) {
                '\' { [void]$sb.Append('\') }
                'p' { [void]$sb.Append('|') }
                'n' { [void]$sb.Append("`n") }
                't' { [void]$sb.Append("`t") }
                default { [void]$sb.Append($nx) }
            }
            $i += 2
        } else {
            [void]$sb.Append($ch)
            $i += 1
        }
    }
    return $sb.ToString()
}

# --- modUtil.Checksum4 ---------------------------------------
function Checksum4([string]$s) {
    $sum = 0
    foreach ($c in $s.ToCharArray()) {
        $sum = (($sum * 31) + [int]$c) % 65536
    }
    return ('{0:X4}' -f $sum)
}

# --- modSync.EncodeShareLine ---------------------------------
function Encode-Line($rec) {
    $parts = @('S')
    foreach ($f in $FIELDS) { $parts += (Esc-Vba ([string]$rec[$f])) }
    return ($parts -join '|')
}

function Decode-Line([string]$ln) {
    $f = $ln -split '\|'
    if ($f.Count -lt 10) { return $null }
    if ($f[0] -ne 'S') { return $null }
    $rec = @{}
    for ($i = 0; $i -lt $FIELDS.Count; $i++) {
        $rec[$FIELDS[$i]] = Unesc-Vba $f[$i + 1]
    }
    return $rec
}

# --- modSync.SyncBuildParts ----------------------------------
function Build-Parts($records) {
    $lines = @()
    foreach ($r in $records) { $lines += (Encode-Line $r) }
    if ($lines.Count -eq 0) { return @() }

    $parts = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList
    $curLen = 0
    foreach ($ln in $lines) {
        if ($cur.Count -eq 0) { $addLen = $ln.Length } else { $addLen = $ln.Length + 1 }
        if ($cur.Count -gt 0 -and ($curLen + $addLen) -gt $PART_MAX) {
            [void]$parts.Add($cur)
            $cur = New-Object System.Collections.ArrayList
            [void]$cur.Add($ln)
            $curLen = $ln.Length
        } else {
            [void]$cur.Add($ln)
            $curLen += $addLen
        }
    }
    if ($cur.Count -gt 0) { [void]$parts.Add($cur) }

    $res = @()
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $body = ($parts[$i] -join "`n")
        $hdr = '###' + $PROTO_TAG + '|P' + ($i + 1) + '/' + $parts.Count + '|N=' + $parts[$i].Count + '###'
        $ftr = '###END|P' + ($i + 1) + '/' + $parts.Count + '|C=' + (Checksum4 $body) + '###'
        $res += ($hdr + "`r`n" + $body.Replace("`n", "`r`n") + "`r`n" + $ftr)
    }
    # 쉼표를 붙여 돌려줘야 PowerShell 이 1원소 배열을 문자열로 풀어버리지 않는다
    return ,$res
}

# --- 수신 버퍼 ------------------------------------------------
$script:Inbox = @{}
$script:Total = 0

function Reset-Inbox { $script:Inbox = @{}; $script:Total = 0 }

function Parse-PartNo([string]$s) {
    if ($s[0] -ne 'P') { return $null }
    $p = $s.IndexOf('/')
    if ($p -lt 0) { return $null }
    $a = $s.Substring(1, $p - 1)
    $b = $s.Substring($p + 1)
    $an = 0; $bn = 0
    if (-not [int]::TryParse($a, [ref]$an)) { return $null }
    if (-not [int]::TryParse($b, [ref]$bn)) { return $null }
    if ($an -lt 1 -or $bn -lt 1 -or $an -gt $bn) { return $null }
    return @{ No = $an; Tot = $bn }
}

function Parse-Header([string]$ln) {
    if (-not $ln.EndsWith('###')) { return $null }
    $core = $ln.Substring(3, $ln.Length - 6)
    $f = $core -split '\|'
    if ($f.Count -ne 3) { return $null }
    if ($f[0] -ne $PROTO_TAG) { return $null }
    $pn = Parse-PartNo $f[1]
    if ($pn -eq $null) { return $null }
    if (-not $f[2].StartsWith('N=')) { return $null }
    $n = 0
    if (-not [int]::TryParse($f[2].Substring(2), [ref]$n)) { return $null }
    return @{ No = $pn.No; Tot = $pn.Tot; N = $n }
}

function Parse-Footer([string]$ln) {
    if (-not $ln.EndsWith('###')) { return $null }
    $core = $ln.Substring(3, $ln.Length - 6)
    $f = $core -split '\|'
    if ($f.Count -ne 3) { return $null }
    if ($f[0] -ne 'END') { return $null }
    $pn = Parse-PartNo $f[1]
    if ($pn -eq $null) { return $null }
    if (-not $f[2].StartsWith('C=')) { return $null }
    $sum = $f[2].Substring(2)
    if ($sum.Length -ne 4) { return $null }
    return @{ No = $pn.No; Tot = $pn.Tot; Sum = $sum }
}

# --- modSync.SyncFeed ----------------------------------------
function Feed([string]$text) {
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    $lines = $text -split "`n"

    $inPart = $false
    $bodyLines = New-Object System.Collections.ArrayList
    $pNo = 0; $pTot = 0; $pN = 0
    $gotAny = $false; $gotNew = $false; $badSum = $false; $truncated = $false

    foreach ($raw in $lines) {
        $ln = $raw.Trim()
        if ($ln.Length -eq 0) { continue }

        if ($ln.StartsWith('###' + $PROTO_TAG + '|')) {
            if ($inPart) { $truncated = $true }
            $h = Parse-Header $ln
            if ($h -ne $null) {
                $inPart = $true
                $pNo = $h.No; $pTot = $h.Tot; $pN = $h.N
                $bodyLines = New-Object System.Collections.ArrayList
                $gotAny = $true
            } else {
                $inPart = $false
            }
        }
        elseif ($ln.StartsWith('###END|')) {
            if ($inPart) {
                $ft = Parse-Footer $ln
                if ($ft -eq $null) {
                    $badSum = $true
                } elseif ($ft.No -ne $pNo -or $ft.Tot -ne $pTot) {
                    $badSum = $true
                } else {
                    $body = ($bodyLines -join "`n")
                    if ((Checksum4 $body) -ne $ft.Sum) {
                        $badSum = $true
                    } elseif ($bodyLines.Count -ne $pN) {
                        $badSum = $true
                    } else {
                        if ($script:Total -ne $pTot) {
                            Reset-Inbox
                            $script:Total = $pTot
                        }
                        if (-not $script:Inbox.ContainsKey($pNo)) {
                            $script:Inbox[$pNo] = $body
                            $gotNew = $true
                        }
                    }
                }
                $inPart = $false
            }
        }
        elseif ($inPart) {
            [void]$bodyLines.Add($ln)
        }
    }
    if ($inPart) { $truncated = $true }

    if (-not $gotAny)                { return @{ St = $FD_NOHEADER;    Msg = '인식할 수 없는 텍스트' } }
    if ($truncated -or $badSum)      { return @{ St = $FD_BADCHECKSUM; Msg = '잘렸거나 변형됨' } }
    if ($script:Total -gt 0 -and $script:Inbox.Count -ge $script:Total) {
        return @{ St = $FD_COMPLETE; Msg = '모두 받음' }
    }
    if (-not $gotNew)                { return @{ St = $FD_DUPLICATE;   Msg = '이미 받은 조각' } }
    return @{ St = $FD_PARTIAL; Msg = ("" + $script:Inbox.Count + "/" + $script:Total) }
}

function Inbox-Records {
    $d = @{}
    for ($i = 1; $i -le $script:Total; $i++) {
        if (-not $script:Inbox.ContainsKey($i)) { continue }
        foreach ($ln in ($script:Inbox[$i] -split "`n")) {
            if ($ln.Trim().Length -eq 0) { continue }
            $rec = Decode-Line $ln
            if ($rec -eq $null) { continue }
            $id = $rec['ID']
            if ($id.Length -eq 0) { continue }
            if ($d.ContainsKey($id)) {
                if ($rec['수정시각'] -ge $d[$id]['수정시각']) { $d[$id] = $rec }
            } else {
                $d[$id] = $rec
            }
        }
    }
    return $d
}

# --- modSync.SyncApply ---------------------------------------
function Apply-Merge($local) {
    $inc = Inbox-Records
    $nAdd = 0; $nUpd = 0; $nSame = 0; $nDel = 0
    foreach ($id in $inc.Keys) {
        $rec = $inc[$id]
        $incDel = ($rec['삭제여부'] -eq '1')
        if (-not $local.ContainsKey($id)) {
            $local[$id] = $rec
            if ($incDel) { $nDel++ } else { $nAdd++ }
        } else {
            $cur = $local[$id]
            $locDel = ($cur['삭제여부'] -eq '1')
            if ($locDel) {
                $nSame++                      # 되살리지 않는다
            } elseif ($incDel) {
                $cur['삭제여부'] = '1'
                $cur['수정시각'] = $rec['수정시각']
                $nDel++
            } elseif ($rec['수정시각'] -gt $cur['수정시각']) {
                $local[$id] = $rec
                $nUpd++
            } else {
                $nSame++
            }
        }
    }
    Reset-Inbox
    return @{ Add = $nAdd; Upd = $nUpd; Same = $nSame; Del = $nDel }
}

function New-Rec($id, $docNo, $title, $content, $modAt, $del) {
    return @{
        'ID' = $id; '문서번호' = $docNo; '제목' = $title; '내용' = $content
        '등록일' = '2026-01-01'; '회신기한' = '2026-01-10'; '작성자' = '테스터'
        '수정시각' = $modAt; '삭제여부' = $del
    }
}

# =============================================================
$pass = 0
$fail = 0
function Chk($name, $cond, $detail) {
    if ($cond) {
        $script:pass++
        Write-Host ("  [통과] " + $name) -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host ("  [실패] " + $name + "  -> " + $detail) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "===== 동기화 프로토콜 시뮬레이션 =====" -ForegroundColor Cyan
Write-Host ""

# --- 1. 이스케이프 왕복 --------------------------------------
Write-Host "1. 이스케이프 왕복"
$src = "내용|파이프\백슬래시`n둘째줄`t탭끝"
$enc = Esc-Vba $src
Chk "원본 복원" ((Unesc-Vba $enc) -eq $src) ("[" + (Unesc-Vba $enc) + "]")
Chk "인코딩 결과에 구분자·줄바꿈·탭 없음" `
    (($enc.IndexOf('|') -lt 0) -and ($enc.IndexOf("`n") -lt 0) -and ($enc.IndexOf("`t") -lt 0)) $enc

# --- 2. 악성 입력 --------------------------------------------
Write-Host ""
Write-Host "2. 이스케이프 문자가 연속된 입력"
$cases = @('\\p', '\p', '\\', '\', 'a\\\\pb', '\n', '\\n', '|\\|', '\\\p', "끝에역슬래시\", '\t\\t',
           "여러줄`n`n연속", "탭`t`t연속", '||||', '\\\\\\')
$allOk = $true
foreach ($c in $cases) {
    $rt = Unesc-Vba (Esc-Vba $c)
    if ($rt -ne $c) { $allOk = $false; Write-Host ("     깨진 입력: [" + $c + "] -> [" + $rt + "]") -ForegroundColor Red }
}
Chk ("악성 입력 " + $cases.Count + "종 왕복") $allOk "위 목록 참조"

# 줄 전체 왕복 (필드 분리까지 포함)
$r = New-Rec 'X-1' 'A|B' "제목\백슬래시" "본문|파이프`n둘째줄`t탭 그리고 \\p 문자열" '2026-01-01 00:00:00' '0'
$back = Decode-Line (Encode-Line $r)
$fieldOk = $true
foreach ($f in $FIELDS) { if ($back[$f] -ne $r[$f]) { $fieldOk = $false; Write-Host ("     필드 불일치: " + $f) -ForegroundColor Red } }
Chk "레코드 한 줄 인코딩→디코딩 전 필드 일치" $fieldOk ""

# --- 3. 체크섬 -----------------------------------------------
Write-Host ""
Write-Host "3. 체크섬"
$a = 'S|홍길동-20260827-143052-8471|기획-2026-0451|상반기 실적 보고|본문|2026-08-27|2026-08-30|홍길동|2026-08-27 14:30:52|0'
$b = $a.Replace('0451', '0452')
$c = $a.Substring(0, $a.Length - 1)
Chk "한 글자 변형을 잡아냄" ((Checksum4 $a) -ne (Checksum4 $b)) ((Checksum4 $a) + " vs " + (Checksum4 $b))
Chk "잘린 문자열을 잡아냄"  ((Checksum4 $a) -ne (Checksum4 $c)) ((Checksum4 $a) + " vs " + (Checksum4 $c))
Chk "4자리 16진수"          ((Checksum4 $a).Length -eq 4) (Checksum4 $a)
# 한글 코드가 32767 을 넘어도 계산이 되는지
Chk "한글(코드 32767 초과)에서도 계산됨" ((Checksum4 '가나다힣') -match '^[0-9A-F]{4}$') (Checksum4 '가나다힣')

# --- 4. 3파트 역순 --------------------------------------------
Write-Host ""
Write-Host "4. 3파트로 나뉜 데이터를 역순으로 붙여넣기"
$recs = @(
    (New-Rec 'TEST-A' 'T-001' '가' ('가' * 1600) '2026-01-01 00:00:00' '0'),
    (New-Rec 'TEST-B' 'T-002' '나' ('나' * 1600) '2026-01-01 00:00:00' '0'),
    (New-Rec 'TEST-C' 'T-003' '다' ('다' * 1600) '2026-01-01 00:00:00' '0')
)
$parts = Build-Parts $recs
Chk "3건이 3개 파트로 나뉨" ($parts.Count -eq 3) ("파트 수 " + $parts.Count)

Reset-Inbox
$local = @{}
$st = $null
for ($i = $parts.Count - 1; $i -ge 0; $i--) { $st = Feed $parts[$i] }
Chk "역순으로 다 받으면 완료" ($st.St -eq $FD_COMPLETE) ("상태 " + $st.St)
$res = Apply-Merge $local
Chk "3건 병합됨" ($local.Count -eq 3) ("건수 " + $local.Count)
Chk "긴 내용 그대로 복원" ($local['TEST-B']['내용'].Length -eq 1600) ("길이 " + $local['TEST-B']['내용'].Length)

# 순서를 섞어도 되는지
Reset-Inbox
$local2 = @{}
Feed $parts[1] | Out-Null
Feed $parts[0] | Out-Null
$st = Feed $parts[2]
Chk "뒤섞인 순서(2,1,3)도 완료" ($st.St -eq $FD_COMPLETE) ("상태 " + $st.St)
Apply-Merge $local2 | Out-Null
Chk "뒤섞어도 3건" ($local2.Count -eq 3) ("건수 " + $local2.Count)

# --- 5. 멱등 -------------------------------------------------
Write-Host ""
Write-Host "5. 같은 조각을 두 번 붙여넣기"
$one = @( (New-Rec 'TEST-D' 'T-010' '멱등확인' '내용' '2026-01-01 00:00:00' '0') )
$p1 = Build-Parts $one
Chk "1건은 파트 1개" ($p1.Count -eq 1) ("파트 수 " + $p1.Count)

Reset-Inbox
$local3 = @{}
$s1 = Feed $p1[0]
$s2 = Feed $p1[0]
Chk "첫 투입에 완료" ($s1.St -eq $FD_COMPLETE) ("상태 " + $s1.St)
Chk "재투입해도 완료 유지" ($s2.St -eq $FD_COMPLETE) ("상태 " + $s2.St)
Apply-Merge $local3 | Out-Null
Chk "중복 생성 안 됨" ($local3.Count -eq 1) ("건수 " + $local3.Count)

Reset-Inbox
Feed $p1[0] | Out-Null
$res = Apply-Merge $local3
Chk "반영 후 또 넣어도 1건" ($local3.Count -eq 1) ("건수 " + $local3.Count)
Chk "두 번째 반영은 변화없음으로 계산" ($res.Same -eq 1 -and $res.Add -eq 0) ("add=" + $res.Add + " same=" + $res.Same)

# 3파트 중 같은 조각을 두 번 넣는 경우
Reset-Inbox
Feed $parts[0] | Out-Null
$dup = Feed $parts[0]
Chk "미완성 상태에서 같은 조각 재투입은 '이미 받음'" ($dup.St -eq $FD_DUPLICATE) ("상태 " + $dup.St)

# --- 6. 과거가 최신을 못 덮는다 --------------------------------
Write-Host ""
Write-Host "6. 수정시각이 과거인 레코드"
Reset-Inbox
$local4 = @{}
$local4['TEST-E'] = New-Rec 'TEST-E' 'T-020' '새제목' '새내용' '2026-01-02 00:00:00' '0'
$old = @( (New-Rec 'TEST-E' 'T-020' '옛제목' '옛내용' '2026-01-01 00:00:00' '0') )
Feed (Build-Parts $old)[0] | Out-Null
$res = Apply-Merge $local4
Chk "과거 레코드가 최신을 덮어쓰지 않음" ($local4['TEST-E']['제목'] -eq '새제목') ("제목 " + $local4['TEST-E']['제목'])
Chk "변화없음으로 계산" ($res.Same -eq 1) ("same=" + $res.Same)

# 반대: 최신이 들어오면 갱신
Reset-Inbox
$new = @( (New-Rec 'TEST-E' 'T-020' '최신제목' '최신내용' '2026-01-03 00:00:00' '0') )
Feed (Build-Parts $new)[0] | Out-Null
$res = Apply-Merge $local4
Chk "최신 레코드는 갱신됨" ($local4['TEST-E']['제목'] -eq '최신제목') ("제목 " + $local4['TEST-E']['제목'])

# --- 7. 삭제 플래그 -------------------------------------------
Write-Host ""
Write-Host "7. 삭제 플래그"
Reset-Inbox
$local5 = @{}
$local5['TEST-F'] = New-Rec 'TEST-F' 'T-030' '내가지움' '내용' '2026-01-05 00:00:00' '1'
$alive = @( (New-Rec 'TEST-F' 'T-030' '살아있음' '내용' '2026-01-09 00:00:00' '0') )
Feed (Build-Parts $alive)[0] | Out-Null
Apply-Merge $local5 | Out-Null
Chk "내가 지운 건은 더 최신 텍스트로도 되살아나지 않음" ($local5['TEST-F']['삭제여부'] -eq '1') ("삭제여부 " + $local5['TEST-F']['삭제여부'])

Reset-Inbox
$local6 = @{}
$local6['TEST-G'] = New-Rec 'TEST-G' 'T-031' '살아있음' '내용' '2026-01-09 00:00:00' '0'
$dead = @( (New-Rec 'TEST-G' 'T-031' '삭제됨' '내용' '2026-01-01 00:00:00' '1') )
Feed (Build-Parts $dead)[0] | Out-Null
Apply-Merge $local6 | Out-Null
Chk "남이 지운 표식은 수정시각이 과거여도 전달됨" ($local6['TEST-G']['삭제여부'] -eq '1') ("삭제여부 " + $local6['TEST-G']['삭제여부'])

# --- 8. 잘림 / 변형 감지 --------------------------------------
Write-Host ""
Write-Host "8. 잘리거나 변형된 텍스트"
Reset-Inbox
$broken = $p1[0].Replace('멱등확인', '멱등확읺')          # 본문 한 글자 변형
$st = Feed $broken
Chk "본문 변형을 체크섬이 잡아냄" ($st.St -eq $FD_BADCHECKSUM) ("상태 " + $st.St)

Reset-Inbox
$cut = ($p1[0] -split "`r`n")[0..1] -join "`r`n"           # 꼬리말이 잘림
$st = Feed $cut
Chk "꼬리말이 잘린 텍스트를 잡아냄" ($st.St -eq $FD_BADCHECKSUM) ("상태 " + $st.St)

Reset-Inbox
$st = Feed "그냥 아무 말이나 붙여넣었을 때"
Chk "머리말이 없으면 형식 불일치" ($st.St -eq $FD_NOHEADER) ("상태 " + $st.St)

# 메신저가 줄바꿈을 LF 로만 바꿔 보낸 경우
Reset-Inbox
$st = Feed ($p1[0].Replace("`r`n", "`n"))
Chk "줄바꿈이 LF 로 바뀌어도 통과" ($st.St -eq $FD_COMPLETE) ("상태 " + $st.St)

# 메신저가 각 줄에 공백을 덧붙인 경우
Reset-Inbox
$padded = (($p1[0] -split "`r`n") | ForEach-Object { $_ + '   ' }) -join "`r`n"
$st = Feed $padded
Chk "줄 끝 공백이 붙어도 통과" ($st.St -eq $FD_COMPLETE) ("상태 " + $st.St)

# --- 9. 파트 크기 ---------------------------------------------
Write-Host ""
Write-Host "9. 파트 크기"
$many = @()
for ($i = 1; $i -le 40; $i++) {
    $many += (New-Rec ("ID-" + $i) ("DOC-" + $i) ("제목 " + $i) ("본문 내용 " + $i) '2026-01-01 00:00:00' '0')
}
$mp = Build-Parts $many
$tooBig = $false
foreach ($p in $mp) {
    $body = ($p -split "`r`n")
    $bodyOnly = ($body[1..($body.Count - 2)] -join "`n")
    if ($bodyOnly.Length -gt $PART_MAX) { $tooBig = $true }
}
Chk ("40건이 " + $mp.Count + "개 파트로 나뉨") ($mp.Count -gt 1) ("파트 수 " + $mp.Count)
Chk "어떤 파트도 본문 1500자를 넘지 않음" (-not $tooBig) "초과 파트 있음"

Reset-Inbox
$localM = @{}
foreach ($p in $mp) { $st = Feed $p }
Chk "여러 파트를 순서대로 받으면 완료" ($st.St -eq $FD_COMPLETE) ("상태 " + $st.St)
Apply-Merge $localM | Out-Null
Chk "40건 전부 병합됨" ($localM.Count -eq 40) ("건수 " + $localM.Count)

# 레코드 하나가 1500자를 넘으면 단독 파트여야 한다
$huge = @( (New-Rec 'BIG-1' 'D1' '큰건' ('가' * 3000) '2026-01-01 00:00:00' '0'),
           (New-Rec 'SML-1' 'D2' '작은건' '짧음' '2026-01-01 00:00:00' '0') )
$hp = Build-Parts $huge
Chk "1500자를 넘는 레코드는 단독 파트로" ($hp.Count -eq 2) ("파트 수 " + $hp.Count)
Reset-Inbox
$localH = @{}
foreach ($p in $hp) { $st = Feed $p }
Apply-Merge $localH | Out-Null
Chk "큰 레코드도 그대로 복원" ($localH['BIG-1']['내용'].Length -eq 3000) ("길이 " + $localH['BIG-1']['내용'].Length)

# =============================================================
# HTML(JS) 버전과 대조할 기준 파일
#   같은 레코드로 만든 파트 텍스트가 바이트 단위로 같아야 한다.
#   이 PowerShell 구현은 VBA 와 같은 규칙임이 위 37건으로 확인됐으므로,
#   JS 결과가 여기에 일치하면 JS 도 엑셀 버전과 호환된다는 뜻이다.
# =============================================================
if ($FixtureDir -ne '') {
    Write-Host ""
    Write-Host "10. 기준 파일 생성" -ForegroundColor Cyan
    if (-not (Test-Path $FixtureDir)) { New-Item -ItemType Directory -Path $FixtureDir | Out-Null }

    $fx = @(
        (New-Rec 'FX-1' 'A|B\C' "제목 파이프|백슬래시\끝" ("본문 첫줄`n둘째줄`t탭" + '  그리고 \\p 같은 문자열') '2026-01-01 09:00:00' '0'),
        (New-Rec 'FX-2' 'DOC-002' '평범한 제목' '평범한 내용' '2026-02-03 10:11:12' '0'),
        (New-Rec 'FX-3' 'DOC-003' '삭제된 건' '내용' '2026-02-04 00:00:00' '1'),
        (New-Rec 'FX-4' 'DOC-004' '아주 긴 건' ('가' * 2000) '2026-02-05 13:14:15' '0'),
        (New-Rec 'FX-5' 'DOC-005' '한글 체크섬 힣뷁' '끝에 역슬래시\' '2026-02-06 23:59:59' '0')
    )
    $fxParts = Build-Parts $fx

    # ConvertTo-Json 은 원소가 1개면 배열로 안 만들어 준다. 직접 묶는다.
    function ToJsonArray($items) {
        return '[' + (($items | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 5 }) -join ',') + ']'
    }

    $encU8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $FixtureDir 'ps_parts.json'),   (ToJsonArray $fxParts), $encU8)
    [System.IO.File]::WriteAllText((Join-Path $FixtureDir 'ps_records.json'), (ToJsonArray $fx), $encU8)

    Write-Host ("  레코드 " + $fx.Count + "건 -> 파트 " + $fxParts.Count + "개")
    Write-Host ("  " + (Join-Path $FixtureDir 'ps_parts.json'))
}

# =============================================================
Write-Host ""
if ($fail -eq 0) {
    Write-Host ("===== 전부 통과 (" + $pass + "건) =====") -ForegroundColor Green
    exit 0
} else {
    Write-Host ("===== 성공 " + $pass + " / 실패 " + $fail + " =====") -ForegroundColor Red
    exit 1
}

# =============================================================
#  check.ps1 - Excel 없이 돌리는 정적 점검
#
#  실행:  powershell -ExecutionPolicy Bypass -File build\check.ps1
#
#  VBA 컴파일러를 대신하지는 못하지만, 실제로 자주 나는 실수를 잡는다.
#    1) 블록 짝 (Sub/End Sub, If/End If, For/Next, With/End With ...)
#    2) 같은 모듈 안의 프로시저 이름 중복
#    3) modXxx.프로시저 형태의 호출이 실제로 존재하는지
#    4) RecGet/RecSet/TB_Col 에 쓴 컬럼명이 스키마에 있는지  ★가장 중요
#    5) 폼 코드가 참조하는 컨트롤이 forms.ps1 에 정의돼 있는지
#    6) 폼 정의와 .frmcode 파일이 1:1로 맞는지
# =============================================================

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'src'

$script:errCount = 0
$script:warnCount = 0

function Bad($where, $msg) {
    Write-Host ("  [오류] " + $where + " : " + $msg) -ForegroundColor Red
    $script:errCount++
}
function Warn2($where, $msg) {
    Write-Host ("  [주의] " + $where + " : " + $msg) -ForegroundColor Yellow
    $script:warnCount++
}
function Sect($t) { Write-Host ""; Write-Host ("--- " + $t) -ForegroundColor Cyan }

function Read-Utf8([string]$p) {
    return [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
}

# 코드에서 주석과 문자열 리터럴을 지운 줄을 돌려준다
function Strip-Line([string]$line) {
    $out = New-Object System.Text.StringBuilder
    $inStr = $false
    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ($inStr) {
            if ($ch -eq '"') { $inStr = $false }
        } else {
            if ($ch -eq '"') { $inStr = $true }
            elseif ($ch -eq "'") { break }
            else { [void]$out.Append($ch) }
        }
    }
    return $out.ToString()
}

# 여러 줄로 이어진(_) 문장을 한 줄로 합친다
function Get-LogicalLines([string]$text) {
    $raw = $text -split "`r?`n"
    $res = New-Object System.Collections.ArrayList
    $buf = ""
    $startNo = 0
    for ($i = 0; $i -lt $raw.Count; $i++) {
        $l = $raw[$i]
        if ($buf -eq "") { $startNo = $i + 1 }
        if ($l -match '\s_\s*$') {
            $buf = $buf + ($l -replace '\s_\s*$', ' ')
        } else {
            $buf = $buf + $l
            [void]$res.Add(@{ No = $startNo; Text = $buf })
            $buf = ""
        }
    }
    if ($buf -ne "") { [void]$res.Add(@{ No = $startNo; Text = $buf }) }
    return $res
}

# =============================================================
Write-Host ""
Write-Host "===== 정적 점검 =====" -ForegroundColor Cyan

$vbaFiles = @()
$vbaFiles += Get-ChildItem $src -Filter *.bas
$vbaFiles += Get-ChildItem (Join-Path $src 'forms') -Filter *.frmcode
$vbaFiles += Get-ChildItem (Join-Path $src 'doc') -Filter *.txt

# --- 1. 블록 짝 --------------------------------------------
Sect "1. 블록 짝 확인"
foreach ($f in $vbaFiles) {
    $lines = Get-LogicalLines (Read-Utf8 $f.FullName)
    $stack = New-Object System.Collections.Stack
    foreach ($ln in $lines) {
        $t = (Strip-Line $ln.Text).Trim()
        if ($t -eq "") { continue }

        # 여는 블록
        if ($t -match '^\s*(Public |Private |Friend |Static )*Sub\s+\w+') { $stack.Push(@{K='Sub'; L=$ln.No}); continue }
        if ($t -match '^\s*(Public |Private |Friend |Static )*Function\s+\w+') { $stack.Push(@{K='Function'; L=$ln.No}); continue }
        if ($t -match '^\s*(Public |Private |Friend |Static )*Property\s+(Get|Let|Set)\s+\w+') { $stack.Push(@{K='Property'; L=$ln.No}); continue }
        if ($t -match '^\s*(Public |Private )*Enum\s+\w+') { $stack.Push(@{K='Enum'; L=$ln.No}); continue }
        if ($t -match '^\s*(Public |Private )*Type\s+\w+') { $stack.Push(@{K='Type'; L=$ln.No}); continue }
        if ($t -match '^\s*With\s+\S') { $stack.Push(@{K='With'; L=$ln.No}); continue }
        if ($t -match '^\s*Select\s+Case\s') { $stack.Push(@{K='Select'; L=$ln.No}); continue }
        if ($t -match '^\s*For\s+(Each\s+)?\S') { $stack.Push(@{K='For'; L=$ln.No}); continue }
        if ($t -match '^\s*Do(\s|$)' -and $t -notmatch '^\s*Do\s*$.*Loop') { $stack.Push(@{K='Do'; L=$ln.No}); continue }
        if ($t -match '^\s*If\s+.*\bThen\b\s*$') { $stack.Push(@{K='If'; L=$ln.No}); continue }
        if ($t -match '^\s*ElseIf\s+.*\bThen\b\s*$') { continue }

        # 닫는 블록
        if ($t -match '^\s*End\s+Sub\b')      { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Sub')      { Bad ($f.Name + ":" + $ln.No) "End Sub 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+Function\b') { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Function') { Bad ($f.Name + ":" + $ln.No) "End Function 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+Property\b') { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Property') { Bad ($f.Name + ":" + $ln.No) "End Property 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+Enum\b')     { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Enum')     { Bad ($f.Name + ":" + $ln.No) "End Enum 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+Type\b')     { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Type')     { Bad ($f.Name + ":" + $ln.No) "End Type 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+With\b')     { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'With')     { Bad ($f.Name + ":" + $ln.No) "End With 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+Select\b')   { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Select')   { Bad ($f.Name + ":" + $ln.No) "End Select 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*End\s+If\b')       { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'If')       { Bad ($f.Name + ":" + $ln.No) "End If 짝이 안 맞음 (스택 위: " + $(if($stack.Count -gt 0){$stack.Peek().K}else{'없음'}) + ")" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*Next\b')           { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'For')      { Bad ($f.Name + ":" + $ln.No) "Next 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
        if ($t -match '^\s*Loop\b')           { if ($stack.Count -eq 0 -or $stack.Peek().K -ne 'Do')       { Bad ($f.Name + ":" + $ln.No) "Loop 짝이 안 맞음" } else { [void]$stack.Pop() }; continue }
    }
    if ($stack.Count -gt 0) {
        $top = $stack.Peek()
        Bad $f.Name ("닫히지 않은 블록 " + $stack.Count + "개. 가장 위: " + $top.K + " (" + $top.L + "행)")
    }
}
if ($script:errCount -eq 0) { Write-Host "  블록 짝 이상 없음" -ForegroundColor Green }

# --- 2. 프로시저 목록 / 중복 --------------------------------
Sect "2. 프로시저 이름 중복"
$procs = @{}     # 모듈명 -> (이름 -> 종류)
foreach ($f in ($vbaFiles | Where-Object { $_.Extension -eq '.bas' })) {
    $mod = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $procs[$mod] = @{}
    $lines = Get-LogicalLines (Read-Utf8 $f.FullName)
    foreach ($ln in $lines) {
        $t = (Strip-Line $ln.Text).Trim()
        if ($t -match '^(?:Public\s+|Private\s+|Friend\s+|Static\s+)*(Sub|Function)\s+(\w+)') {
            $nm = $Matches[2]
            if ($procs[$mod].ContainsKey($nm)) { Bad ($mod + ":" + $ln.No) ("프로시저 이름 중복: " + $nm) }
            $procs[$mod][$nm] = $Matches[1]
        }
        elseif ($t -match '^(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Property\s+(?:Get|Let|Set)\s+(\w+)') {
            $procs[$mod][$Matches[1]] = 'Property'
        }
    }
}
Write-Host ("  모듈 " + $procs.Keys.Count + "개 / 프로시저 " + (($procs.Values | ForEach-Object { $_.Keys.Count } | Measure-Object -Sum).Sum) + "개")

# --- 3. modXxx.프로시저 호출 -------------------------------
Sect "3. 모듈명으로 한정한 호출"
$allText = @{}
foreach ($f in $vbaFiles) { $allText[$f.Name] = Read-Utf8 $f.FullName }
foreach ($k in $allText.Keys) {
    foreach ($m in ([regex]::Matches($allText[$k], '\b(mod[A-Z]\w*)\.(\w+)'))) {
        $mn = $m.Groups[1].Value
        $pn = $m.Groups[2].Value
        if (-not $procs.ContainsKey($mn)) { Bad $k ("없는 모듈을 참조: " + $mn); continue }
        if (-not $procs[$mn].ContainsKey($pn)) { Bad $k ($mn + "." + $pn + " 프로시저가 없습니다") }
    }
}
Write-Host "  확인 완료" -ForegroundColor Green

# --- 4. 컬럼명이 스키마에 있는지 ----------------------------
Sect "4. 컬럼명 대조 (가장 중요)"
$buildTxt = Read-Utf8 (Join-Path $PSScriptRoot 'build.ps1')
$schema = @{}    # 표이름 -> 컬럼 배열
foreach ($m in ([regex]::Matches($buildTxt, "New-DataSheet\s+'([^']+)'\s+'([^']+)'\s+@\(([^)]*)\)"))) {
    $tbl = $m.Groups[2].Value
    $cols = @()
    foreach ($cm in ([regex]::Matches($m.Groups[3].Value, "'([^']*)'"))) { $cols += $cm.Groups[1].Value }
    $schema[$tbl] = $cols
}
Write-Host ("  스키마에서 표 " + $schema.Keys.Count + "개를 읽었습니다")

$constMap = @{}  # TBL_TASK -> tbl_Task
$dataTxt = Read-Utf8 (Join-Path $src 'modData.bas')
foreach ($m in ([regex]::Matches($dataTxt, 'Public\s+Const\s+(TBL_\w+)\s+As\s+String\s*=\s*"([^"]+)"'))) {
    $constMap[$m.Groups[1].Value] = $m.Groups[2].Value
}

$checked = 0
foreach ($k in $allText.Keys) {
    # RecGet(rec, TBL_X, "컬럼")  /  RecSet(rec, TBL_X, "컬럼", ...)
    foreach ($m in ([regex]::Matches($allText[$k], 'Rec(?:Get|Set)\s*\(\s*[^,()]+,\s*(TBL_\w+)\s*,\s*"([^"]+)"'))) {
        $checked++
        $c = $constMap[$m.Groups[1].Value]
        if ($c -eq $null) { Bad $k ("모르는 표 상수: " + $m.Groups[1].Value); continue }
        if ($schema[$c] -notcontains $m.Groups[2].Value) {
            Bad $k ($c + " 에 '" + $m.Groups[2].Value + "' 컬럼이 없습니다")
        }
    }
    # TB_Col(TBL_X, "컬럼")
    foreach ($m in ([regex]::Matches($allText[$k], 'TB_Col\s*\(\s*(TBL_\w+)\s*,\s*"([^"]+)"'))) {
        $checked++
        $c = $constMap[$m.Groups[1].Value]
        if ($c -eq $null) { Bad $k ("모르는 표 상수: " + $m.Groups[1].Value); continue }
        if ($schema[$c] -notcontains $m.Groups[2].Value) {
            Bad $k ($c + " 에 '" + $m.Groups[2].Value + "' 컬럼이 없습니다")
        }
    }
}
Write-Host ("  컬럼 참조 " + $checked + "곳 확인")

# --- 5. 폼 컨트롤 대조 --------------------------------------
Sect "5. 폼 컨트롤 대조"
. (Join-Path $PSScriptRoot 'forms.ps1')
$vbKeywords = @('Me','True','False','Nothing','Err','Application','Format','Trim','Len','Left','Right','Mid','CStr','CLng','IsEmpty','IsNumeric','UBound','LBound','MsgBox','InputBox','Unload','Load','vbLf','vbCr','vbTab')

foreach ($def in $Global:FormDefs) {
    $codePath = Join-Path (Join-Path $src 'forms') ($def.Name + '.frmcode')
    if (-not (Test-Path $codePath)) { Bad $def.Name "짝이 되는 .frmcode 파일이 없습니다"; continue }
    $code = Read-Utf8 $codePath
    $ctlNames = @()
    foreach ($c in $def.Controls) { $ctlNames += $c.N }

    # 이벤트 핸들러 이름에서 컨트롤 추출: Private Sub xxx_Event(
    foreach ($m in ([regex]::Matches($code, '(?m)^\s*Private\s+Sub\s+([A-Za-z]\w*?)_(Click|Change|DblClick|KeyDown|Enter|Exit)\s*\('))) {
        $cn = $m.Groups[1].Value
        if ($cn -eq 'UserForm') { continue }
        if ($ctlNames -notcontains $cn) {
            Bad ($def.Name + '.frmcode') ("이벤트 핸들러가 가리키는 컨트롤 '" + $cn + "' 이 forms.ps1 에 없습니다")
        }
    }
    # 코드가 쓰는 컨트롤 이름 (txt/cmd/lst/lbl/opt 접두사만 검사)
    foreach ($m in ([regex]::Matches($code, '\b((?:txt|cmd|lst|lbl|opt|chk)[A-Za-z0-9]*)\b'))) {
        $cn = $m.Groups[1].Value
        if ($ctlNames -notcontains $cn) {
            Bad ($def.Name + '.frmcode') ("코드가 참조하는 컨트롤 '" + $cn + "' 이 forms.ps1 에 없습니다")
        }
    }
}
# .frmcode 파일이 정의 없이 떠 있지 않은지
$defNames = @()
foreach ($def in $Global:FormDefs) { $defNames += $def.Name }
foreach ($f in (Get-ChildItem (Join-Path $src 'forms') -Filter *.frmcode)) {
    $n = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    if ($defNames -notcontains $n) { Bad $f.Name "forms.ps1 에 이 폼의 레이아웃 정의가 없습니다" }
}
Write-Host ("  폼 " + $Global:FormDefs.Count + "개 확인")

# --- 6. modForms 의 UI_ 진입점 ------------------------------
Sect "6. UI 진입점"
$mainTxt = Read-Utf8 (Join-Path $src 'modMain.bas')
foreach ($m in ([regex]::Matches($mainTxt, 'RunUI\("(\w+)"\)'))) {
    $pn = $m.Groups[1].Value
    if (-not $procs['modForms'].ContainsKey($pn)) {
        Bad 'modMain.bas' ("RunUI(""" + $pn + """) 에 해당하는 프로시저가 modForms 에 없습니다")
    }
}
$formsTxt = Read-Utf8 (Join-Path $src 'modForms.bas')
foreach ($m in ([regex]::Matches($formsTxt, '(?m)^\s*(\w+)\.Show'))) {
    if ($defNames -notcontains $m.Groups[1].Value) {
        Bad 'modForms.bas' ($m.Groups[1].Value + " 폼이 forms.ps1 에 정의돼 있지 않습니다")
    }
}
Write-Host "  확인 완료" -ForegroundColor Green

# --- 7. Option Explicit -------------------------------------
Sect "7. Option Explicit"
foreach ($f in $vbaFiles) {
    $t = Read-Utf8 $f.FullName
    if ($t -notmatch '(?m)^\s*Option\s+Explicit\s*$') { Bad $f.Name "Option Explicit 이 없습니다" }
}
Write-Host "  확인 완료" -ForegroundColor Green

# =============================================================
Write-Host ""
if ($script:errCount -eq 0) {
    Write-Host ("===== 점검 통과 (주의 " + $script:warnCount + "건) =====") -ForegroundColor Green
    exit 0
} else {
    Write-Host ("===== 오류 " + $script:errCount + "건 / 주의 " + $script:warnCount + "건 =====") -ForegroundColor Red
    exit 1
}

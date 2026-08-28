# =============================================================
#  build.ps1 - Scheduler.xlsm 를 처음부터 만든다
#
#  실행:  powershell -ExecutionPolicy Bypass -File build\build.ps1
#
#  전제:  Excel 설치 + [옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정]
#         "VBA 프로젝트 개체 모델에 대한 액세스 신뢰" 체크
#
#  위 옵션이 정책상 잠겨 있으면 시트와 표만 갖춘 파일을 만들고
#  dist\manual_import 폴더에 수동 임포트용 소스를 CP949 로 내보낸다.
# =============================================================

param(
    # auto  : Excel 을 먼저 찾고, 없으면 한셀로 넘어간다
    # excel : Excel 만 쓴다
    # hcell : 한셀만 쓴다 (테스트용)
    [ValidateSet('auto','excel','hcell')]
    [string]$Engine = 'auto'
)

$ErrorActionPreference = 'Stop'

$root  = Split-Path -Parent $PSScriptRoot
$src   = Join-Path $root 'src'
$dist  = Join-Path $root 'dist'
$step  = '시작'
$excel = $null
$wb    = $null

function Say($msg)  { Write-Host $msg }
function Ok($msg)   { Write-Host ("  [완료] " + $msg) -ForegroundColor Green }
function Note($msg) { Write-Host ("  [안내] " + $msg) -ForegroundColor Yellow }

function Cleanup {
    if ($script:wb -ne $null) { try { $script:wb.Close($false) } catch {} }
    if ($script:excel -ne $null) {
        try { $script:excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:excel) } catch {}
    }
    $script:wb = $null
    $script:excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Die($msg) {
    Write-Host ""
    Write-Host ("빌드 실패 - 단계: " + $script:step) -ForegroundColor Red
    Write-Host ("  원인: " + $msg) -ForegroundColor Red
    Write-Host ""
    Cleanup
    exit 1
}

# 소스 파일을 읽어 VBE 에 넣을 문자열로 만든다.
#   파일은 전부 UTF-8 로 관리하고, CodeModule.AddFromString 은 유니코드로 받으므로
#   코드페이지 문제 없이 한글 주석·문자열이 그대로 들어간다.
function Get-SrcText([string]$path) {
    if (-not (Test-Path $path)) { Die ("소스 파일이 없습니다: " + $path) }
    $t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $t = $t -replace "(?m)^Attribute VB_Name.*\r?\n", ""
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace "`n", "`r`n"
    return $t
}

# COM 개체의 속성을 이름으로 설정
function Set-Prop($obj, [string]$name, $value) {
    try {
        [void]$obj.GetType().InvokeMember($name, 'SetProperty', $null, $obj, @($value))
    } catch {
        Die ("속성 설정 실패: " + $name + " = " + $value + " / " + $_.Exception.Message)
    }
}

# =============================================================
Say ""
Say "===== 업무 스케줄러 빌드 ====="
Say ""

# --- 0. 환경 확인 --------------------------------------------
$step = '0. 환경 확인'
$acp = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
if ($acp -ne 949) {
    Note ("이 PC의 ANSI 코드페이지가 " + $acp + " 입니다 (한국어 Windows 는 949).")
    Note "빌드는 유니코드로 넣으므로 문제 없지만, 수동 임포트용 파일의 한글이 깨질 수 있습니다."
}

# --- 1. 스프레드시트 프로그램 실행 -----------------------------
$step = '1. 스프레드시트 프로그램 실행'
$isHCell = $false

function Start-HCellIfNeeded {
    # 한셀은 COM 개체를 만들기 전에 프로그램이 한 번 떠 있어야 한다.
    # 안 그러면 New-Object 가 응답 없이 멈춘다.
    if (Get-Process -Name HCell -ErrorAction SilentlyContinue) { return $true }
    $exe = @(
        'C:\Program Files (x86)\Hnc\Office 2022\HOffice120\Bin\HCell.exe',
        'C:\Program Files\Hnc\Office 2022\HOffice120\Bin\HCell.exe'
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($exe -eq $null) {
        $found = Get-ChildItem 'C:\Program Files (x86)\Hnc','C:\Program Files\Hnc' -Recurse -Filter HCell.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found -ne $null) { $exe = $found.FullName }
    }
    if ($exe -eq $null) { return $false }
    Note "한셀을 먼저 띄웁니다 (COM 자동화 전제 조건)"
    Start-Process $exe
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        $p = Get-Process -Name HCell -ErrorAction SilentlyContinue
        if ($p -ne $null -and $p.MainWindowTitle -ne '') { Start-Sleep -Seconds 2; return $true }
        Start-Sleep -Milliseconds 1000
    }
    return (Get-Process -Name HCell -ErrorAction SilentlyContinue) -ne $null
}

if ($Engine -ne 'hcell') {
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
}
if ($excel -eq $null -and $Engine -ne 'excel') {
    if (Start-HCellIfNeeded) {
        try {
            $excel = New-Object -ComObject HCell.Application
            $isHCell = $true
        } catch { $excel = $null }
    }
}
if ($excel -eq $null) {
    if ($Engine -eq 'excel') { Die "Excel 을 실행할 수 없습니다. 이 PC에 Excel 이 설치돼 있는지 확인해 주세요." }
    if ($Engine -eq 'hcell') { Die "한셀을 실행할 수 없습니다. 한셀을 직접 한 번 실행한 뒤 다시 시도해 주세요." }
    Die "Excel 도 한셀도 실행할 수 없습니다. 둘 중 하나가 설치돼 있어야 합니다."
}

if ($isHCell) {
    # 한셀은 Visible = $false 로 두면 자동화가 불안정하다. 띄워 둔 채로 쓴다.
    try { $excel.Visible = $true } catch {}
    Note "한셀로 빌드합니다. 한셀에서는 VBA 코드를 자동으로 넣을 수 없어"
    Note "시트와 표만 만들고, 코드는 수동으로 넣어야 합니다 (아래 안내 참고)."
} else {
    $excel.Visible = $false
}
try { $excel.DisplayAlerts = $false } catch {}
if ($isHCell) { Ok ("한셀 " + $excel.Version) } else { Ok ("Excel " + $excel.Version) }

# --- 2. 통합 문서 만들기 --------------------------------------
$step = '2. 통합 문서 만들기'
$wb = $excel.Workbooks.Add()
while ($wb.Worksheets.Count -gt 1) {
    [void]$wb.Worksheets.Item($wb.Worksheets.Count).Delete()
}

function Add-Sheet([string]$name) {
    $last = $script:wb.Worksheets.Item($script:wb.Worksheets.Count)
    $ws = $script:wb.Worksheets.Add([Type]::Missing, $last)
    $ws.Name = $name
    return $ws
}

$shCal = $wb.Worksheets.Item(1)
$shCal.Name = '달력'
$shToday = Add-Sheet '오늘'
$shSetup = Add-Sheet '설정'
$shPanel = Add-Sheet '입력'
Ok "보이는 시트 4개 (달력 / 오늘 / 설정 / 입력)"

# --- 3. 데이터 시트 + 표 --------------------------------------
$step = '3. 데이터 시트와 표 만들기'

function New-DataSheet([string]$sheetName, [string]$tableName, [string[]]$cols) {
    $ws = Add-Sheet $sheetName
    $nc = $cols.Count
    # 날짜·시간·플래그를 전부 문자열로 다루므로 열 전체를 텍스트 서식으로 둔다.
    # (0212 같은 내선번호가 212 로 바뀌는 것도 이걸로 막는다)
    $ws.Range($ws.Columns.Item(1), $ws.Columns.Item($nc)).NumberFormat = "@"
    for ($i = 0; $i -lt $nc; $i++) {
        $ws.Cells.Item(1, $i + 1).Value2 = $cols[$i]
    }
    # 머리글 + 빈 행 1개로 표를 만든다 (머리글만으로 만들면 버전에 따라 실패한다)
    $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(2, $nc))
    $lo = $ws.ListObjects.Add(1, $rng, [Type]::Missing, 1)
    $lo.Name = $tableName
    try { $lo.TableStyle = "" } catch {}
    $ws.Rows.Item(1).Font.Bold = $true
    # 한셀은 COM 으로 시트 숨기기를 지원하지 않는다. 실패해도 빌드는 계속한다.
    try { $ws.Visible = 2 } catch { $script:hideFailed = $true }     # xlSheetVeryHidden
    return $ws
}

# 표에 초기 데이터 넣기 (표는 A1 에서 시작한다고 가정)
function Set-TableRows($ws, [string]$tableName, $rows) {
    $lo = $ws.ListObjects.Item($tableName)
    $nc = $lo.ListColumns.Count
    $n  = $rows.Count
    $hdr = $lo.HeaderRowRange
    [void]$lo.Resize($ws.Range($hdr.Cells.Item(1, 1), $hdr.Cells.Item(1, $nc).Offset($n, 0)))
    for ($i = 0; $i -lt $n; $i++) {
        for ($c = 0; $c -lt $nc; $c++) {
            $ws.Cells.Item(2 + $i, 1 + $c).Value2 = $rows[$i][$c]
        }
    }
}

$wsTask  = New-DataSheet '_Task'     'tbl_Task'     @('ID','제목','내용','날짜','시간','완료여부','완료시각','작성자','수정시각','삭제여부')
$wsCont  = New-DataSheet '_Contact'  'tbl_Contact'  @('ID','이름','부서','내선캐시','날짜','시간','완료여부','작성자','수정시각','삭제여부')
$wsShare = New-DataSheet '_Share'    'tbl_Share'    @('ID','문서번호','제목','내용','등록일','회신기한','작성자','수정시각','삭제여부')
$wsLunch = New-DataSheet '_Lunch'    'tbl_Lunch'    @('ID','대상이름','날짜','시간','장소','작성자','수정시각','삭제여부')
$wsEmp   = New-DataSheet '_Employee' 'mst_Employee' @('이름','내선번호','부서','직급')
$wsCfg   = New-DataSheet '_Setting'  'cfg_Setting'  @('키','값')
$wsLog   = New-DataSheet '_Log'      'log_Merge'    @('시각','동작','ID','이전값요약','이후값요약','출처')
Ok "숨김 데이터 시트 7개 / 표 7개"

# --- 4. 보이는 시트 꾸미기 ------------------------------------
$step = '4. 화면 구성'

$CLR_HEAD = 15921906
$CLR_BTN  = 15130800
$CLR_IN   = 15918588

# ---- 달력 ----
$shCal.Columns.Item(1).ColumnWidth = 2
for ($c = 2; $c -le 8; $c++) { $shCal.Columns.Item($c).ColumnWidth = 15 }
$shCal.Range("B2").Value2 = [string][char]0x25C0        # 좌 화살표
$shCal.Range("F2").Value2 = [string][char]0x25B6        # 우 화살표
$shCal.Range("G2").Value2 = "[오늘로]"
[void]$shCal.Range("C2:E2").Merge()
$shCal.Range("C2").Font.Size = 14
$shCal.Range("B2:H2").Font.Bold = $true
$shCal.Range("B2").HorizontalAlignment = -4108
$shCal.Range("F2").HorizontalAlignment = -4108
$shCal.Range("G2").HorizontalAlignment = -4108
$shCal.Range("B2").Interior.Color = $CLR_BTN
$shCal.Range("F2").Interior.Color = $CLR_BTN
$shCal.Range("G2").Interior.Color = $CLR_BTN

$days = @('일','월','화','수','목','금','토')
for ($i = 0; $i -lt 7; $i++) {
    $cell = $shCal.Cells.Item(4, 2 + $i)
    $cell.Value2 = $days[$i]
    $cell.HorizontalAlignment = -4108
    $cell.Font.Bold = $true
    $cell.Interior.Color = $CLR_HEAD
}
$grid = $shCal.Range($shCal.Cells.Item(5, 2), $shCal.Cells.Item(10, 8))
$grid.Borders.LineStyle = 1
$grid.Borders.Weight = 2
$grid.VerticalAlignment = -4160      # xlTop
$grid.HorizontalAlignment = -4131    # xlLeft
$grid.WrapText = $true
$grid.Font.Size = 9
for ($r = 5; $r -le 10; $r++) { $shCal.Rows.Item($r).RowHeight = 52 }
Ok "달력 시트"

# ---- 오늘 ----
$shToday.Columns.Item(1).ColumnWidth = 2
$shToday.Columns.Item(2).ColumnWidth = 5
$shToday.Columns.Item(3).ColumnWidth = 36
$shToday.Columns.Item(4).ColumnWidth = 30
$shToday.Columns.Item(5).ColumnWidth = 10
$shToday.Columns.Item(6).ColumnWidth = 20
$shToday.Columns.Item(7).ColumnWidth = 14
$shToday.Columns.Item(14).Hidden = $true

$shToday.Range("B2").Value2 = [string][char]0x25C0
$shToday.Range("D2").Value2 = "[오늘]"
$shToday.Range("E2").Value2 = [string][char]0x25B6
$shToday.Range("C2").Font.Size = 13
$shToday.Range("B2:G2").Font.Bold = $true
$shToday.Range("B2").HorizontalAlignment = -4108
$shToday.Range("D2").HorizontalAlignment = -4108
$shToday.Range("E2").HorizontalAlignment = -4108
$shToday.Range("B2").Interior.Color = $CLR_BTN
$shToday.Range("D2").Interior.Color = $CLR_BTN
$shToday.Range("E2").Interior.Color = $CLR_BTN

[void]$shToday.Range("B3:C3").Merge()
$shToday.Range("B3").Value2 = "[업무 등록]"
$shToday.Range("D3").Value2 = "[연락 등록]"
$shToday.Range("E3").Value2 = "[점심 등록]"
$shToday.Range("F3").Value2 = "[공유 등록]"
$shToday.Range("G3").Value2 = "[새로고침]"
$shToday.Range("B3:G3").Font.Bold = $true
$shToday.Range("B3:G3").HorizontalAlignment = -4108
$shToday.Range("B3:G3").Interior.Color = $CLR_BTN
Ok "오늘 시트"

# ---- 설정 ----
$shSetup.Columns.Item(1).ColumnWidth = 2
$shSetup.Columns.Item(2).ColumnWidth = 24
$shSetup.Columns.Item(3).ColumnWidth = 52
$shSetup.Columns.Item(4).ColumnWidth = 26
$shSetup.Range("B2").Value2 = "설정"
$shSetup.Range("B2").Font.Size = 14
$shSetup.Range("B2").Font.Bold = $true
$shSetup.Range("B4").Value2 = "내 이름"
$shSetup.Range("B5").Value2 = "명부 파일 경로"
$shSetup.Range("B6").Value2 = "[명부 갱신]"
$shSetup.Range("B7").Value2 = "백업 보관일수"
$shSetup.Range("B8").Value2 = "마지막 동기화 시각"
$shSetup.Range("B9").Value2 = "입력 방식"
$shSetup.Range("B11").Value2 = "[전체 데이터 내보내기]"
$shSetup.Range("D11").Value2 = "[전체 데이터 가져오기]"
$shSetup.Range("B12").Value2 = "[백업 지금 실행]"
$shSetup.Range("B13").Value2 = "[화면 새로고침]"
$shSetup.Range("B4:B9").Font.Bold = $true
$shSetup.Range("B6").Interior.Color = $CLR_BTN
$shSetup.Range("B11:B13").Interior.Color = $CLR_BTN
$shSetup.Range("D11").Interior.Color = $CLR_BTN
$shSetup.Range("C4:C5").Interior.Color = $CLR_IN
$shSetup.Range("C7").Interior.Color = $CLR_IN
$shSetup.Range("C4:C5").Borders.LineStyle = 1
$shSetup.Range("C7").Borders.LineStyle = 1
$shSetup.Range("B15").Value2 = "* 회색 칸(내 이름 / 명부 경로 / 백업 보관일수)은 직접 고칠 수 있습니다."
$shSetup.Range("B16").Value2 = "* 업무·연락·점심은 이 PC에만 저장됩니다. 공유되는 것은 업무공유뿐입니다."
Ok "설정 시트"

# ---- 입력 (폼 폴백용 패널) ----
$shPanel.Columns.Item(1).ColumnWidth = 2
$shPanel.Columns.Item(2).ColumnWidth = 14
$shPanel.Columns.Item(3).ColumnWidth = 56
$shPanel.Columns.Item(4).ColumnWidth = 10
$shPanel.Columns.Item(5).ColumnWidth = 20
$shPanel.Columns.Item(6).ColumnWidth = 10
$shPanel.Columns.Item(14).Hidden = $true
$shPanel.Range("B2").Font.Size = 13
$shPanel.Range("B2").Font.Bold = $true
$shPanel.Range("C4:C10").WrapText = $true
$shPanel.Range("C4:C10").Interior.Color = $CLR_IN
$shPanel.Range("C4:C10").Borders.LineStyle = 1
$shPanel.Range("B12").Interior.Color = $CLR_BTN
$shPanel.Range("D12").Interior.Color = $CLR_BTN
$shPanel.Range("F12").Interior.Color = $CLR_BTN
$shPanel.Range("B12:F12").Font.Bold = $true
$shPanel.Range("B12:F12").HorizontalAlignment = -4108
$shPanel.Visible = 0        # xlSheetHidden
Ok "입력 패널 시트"

if ($script:hideFailed) {
    Note "이 프로그램은 시트 숨기기를 지원하지 않아 데이터 시트가 그대로 보입니다."
    Note "동작에는 지장이 없습니다. (엑셀로 빌드하면 숨겨집니다)"
}

# --- 5. VBA 프로젝트 접근 확인 ---------------------------------
$step = '5. VBA 프로젝트 접근 확인'
$vbp = $null
try {
    $vbp = $wb.VBProject
    $null = $vbp.VBComponents.Count
    # 한셀은 VBProject 개체는 주지만 VBComponents.Add 가 아무것도 돌려주지 않는다.
    # 실제로 넣어 보기 전에는 알 수 없으므로 여기서 한 번 시험한다.
    $probeComp = $vbp.VBComponents.Add(1)
    if ($probeComp -eq $null) {
        $vbp = $null
    } else {
        try { $vbp.VBComponents.Remove($probeComp) } catch {}
    }
} catch {
    $vbp = $null
}

$useForms = $false
if ($vbp -eq $null) {
    if ($isHCell) {
        Note "한셀은 VBA 코드를 자동으로 넣을 수 없습니다."
        Note "시트와 표만 만듭니다. 코드는 dist\manual_import 안내대로 직접 넣어 주세요."
    } else {
        Note "VBA 프로젝트 개체 모델에 접근할 수 없습니다."
        Note "Excel > 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 에서"
        Note "  'VBA 프로젝트 개체 모델에 대한 액세스 신뢰' 를 켠 뒤 다시 실행해 주세요."
        Note "정책상 켤 수 없다면 dist\manual_import 의 안내대로 수동으로 넣으면 됩니다."
    }
} else {
    Ok "VBA 프로젝트 접근 가능"
}

# --- 6. 표준 모듈 --------------------------------------------
$mods = @('modUtil','modData','modEmployee','modRecord','modSync','modCalendar','modToday','modPanel','modMain','modTest')

if ($vbp -ne $null) {
    $step = '6. 표준 모듈 넣기'
    foreach ($m in $mods) {
        $path = Join-Path $src ($m + '.bas')
        $code = Get-SrcText $path
        try {
            $comp = $vbp.VBComponents.Add(1)        # vbext_ct_StdModule
            $comp.Name = $m
            $comp.CodeModule.AddFromString($code)
        } catch {
            Die ($m + " 모듈을 넣지 못했습니다: " + $_.Exception.Message)
        }
    }
    Ok ("표준 모듈 " + $mods.Count + "개")
}

# --- 7. 시트 / ThisWorkbook 코드 -------------------------------
if ($vbp -ne $null) {
    $step = '7. 시트 이벤트 코드 넣기'
    $docs = @(
        @{ Comp = 'ThisWorkbook'; File = 'ThisWorkbook.txt' },
        @{ Comp = $shCal.CodeName;   File = 'Sheet_달력.txt' },
        @{ Comp = $shToday.CodeName; File = 'Sheet_오늘.txt' },
        @{ Comp = $shSetup.CodeName; File = 'Sheet_설정.txt' },
        @{ Comp = $shPanel.CodeName; File = 'Sheet_입력.txt' }
    )
    foreach ($d in $docs) {
        $path = Join-Path (Join-Path $src 'doc') $d.File
        try {
            $comp = $vbp.VBComponents.Item($d.Comp)
            $comp.CodeModule.AddFromString((Get-SrcText $path))
        } catch {
            Die ($d.File + " 를 " + $d.Comp + " 에 넣지 못했습니다: " + $_.Exception.Message)
        }
    }
    Ok "시트 이벤트 코드 5개"
}

# --- 8. UserForm ---------------------------------------------
if ($vbp -ne $null) {
    $step = '8. UserForm 만들기'
    . (Join-Path $PSScriptRoot 'forms.ps1')

    $progIds = @{
        'Label'         = 'Forms.Label.1'
        'TextBox'       = 'Forms.TextBox.1'
        'CommandButton' = 'Forms.CommandButton.1'
        'ListBox'       = 'Forms.ListBox.1'
        'OptionButton'  = 'Forms.OptionButton.1'
    }

    try {
        foreach ($def in $Global:FormDefs) {
            $comp = $vbp.VBComponents.Add(3)        # vbext_ct_MSForm
            $comp.Name = $def.Name
            $comp.Properties.Item('Caption').Value = $def.Caption
            $comp.Properties.Item('Width').Value   = $def.Width
            $comp.Properties.Item('Height').Value  = $def.Height

            $designer = $comp.Designer
            foreach ($c in $def.Controls) {
                $prog = $progIds[$c.T]
                if ($prog -eq $null) { Die ("알 수 없는 컨트롤 종류: " + $c.T) }
                $ctl = $designer.Controls.Add($prog, $c.N, $true)
                $ctl.Left   = $c.L
                $ctl.Top    = $c.Y
                $ctl.Width  = $c.W
                $ctl.Height = $c.H
                if ($c.ContainsKey('C')) { $ctl.Caption = $c.C }
                if ($c.ContainsKey('P')) {
                    foreach ($k in $c.P.Keys) { Set-Prop $ctl $k $c.P[$k] }
                }
            }

            $codePath = Join-Path (Join-Path $src 'forms') ($def.Name + '.frmcode')
            $comp.CodeModule.AddFromString((Get-SrcText $codePath))
        }
        $useForms = $true
        Ok ("UserForm " + $Global:FormDefs.Count + "개")
    } catch {
        Note ("UserForm 을 만들지 못했습니다: " + $_.Exception.Message)
        Note "시트 기반 입력 패널로 동작하도록 설정합니다. 기능은 모두 쓸 수 있습니다."
        $useForms = $false
    }

    # 폼이 만들어졌을 때만 modForms 를 넣는다.
    # (폼이 없는 프로젝트에 폼 이름이 남아 있으면 컴파일 자체가 안 된다)
    if ($useForms) {
        $step = '8b. modForms 넣기'
        try {
            $comp = $vbp.VBComponents.Add(1)
            $comp.Name = 'modForms'
            $comp.CodeModule.AddFromString((Get-SrcText (Join-Path $src 'modForms.bas')))
        } catch {
            Die ("modForms 를 넣지 못했습니다: " + $_.Exception.Message)
        }
        Ok "modForms"
    }
}

# --- 9. 초기 설정값 -------------------------------------------
$step = '9. 초기 설정값'
$mode = '패널'
if ($useForms) { $mode = '폼' }

Set-TableRows $wsCfg 'cfg_Setting' @(
    @('내이름',''),
    @('명부파일경로',''),
    @('명부갱신시각',''),
    @('백업보관일수','30'),
    @('마지막동기화시각',''),
    @('입력방식', $mode),
    @('조회날짜',''),
    @('달력연월','')
)
$shSetup.Range("C7").Value2 = '30'
$shSetup.Range("C9").Value2 = $mode
Ok ("입력 방식: " + $mode)

# --- 10. 수동 임포트용 소스 내보내기 ---------------------------
$step = '10. 수동 임포트용 소스 내보내기'
$manual = Join-Path $dist 'manual_import'
if (-not (Test-Path $dist))   { New-Item -ItemType Directory -Path $dist | Out-Null }
if (Test-Path $manual)        { Remove-Item $manual -Recurse -Force }
New-Item -ItemType Directory -Path $manual | Out-Null

$enc949 = [System.Text.Encoding]::GetEncoding(949)
foreach ($m in ($mods + @('modForms'))) {
    $p = Join-Path $src ($m + '.bas')
    if (Test-Path $p) {
        $t = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
        $t = $t -replace "`r`n", "`n"
        $t = $t -replace "`n", "`r`n"
        [System.IO.File]::WriteAllText((Join-Path $manual ($m + '.bas')), $t, $enc949)
    }
}
foreach ($f in (Get-ChildItem (Join-Path $src 'doc') -Filter *.txt)) {
    $t = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $manual $f.Name), $t, $enc949)
}

$guide = @'
수동으로 소스를 넣는 방법
=========================

이 폴더의 파일들은 한국어 Windows 기본 인코딩(CP949)으로 저장돼 있습니다.
VBE(Alt+F11) 에서 그대로 가져오면 한글이 깨지지 않습니다.

1) 파일을 열고 매크로 편집기를 엽니다.
   - Excel : Alt+F11
   - 한셀  : 도구 > 매크로 > 매크로 편집 (또는 Alt+F11)

2) 표준 모듈 넣기
   메뉴 [파일] > [파일 가져오기] 로 아래 .bas 를 하나씩 가져옵니다.
     modUtil.bas, modData.bas, modEmployee.bas, modRecord.bas,
     modSync.bas, modCalendar.bas, modToday.bas, modPanel.bas,
     modMain.bas, modTest.bas
   * modForms.bas 는 UserForm 이 있을 때만 넣습니다. 수동 설치에서는 넣지 마세요.

3) 시트 이벤트 코드 넣기
   왼쪽 프로젝트 창에서 각 시트를 더블클릭해 코드 창을 열고,
   아래 파일 내용을 메모장으로 열어 통째로 붙여넣습니다.
     ThisWorkbook.txt  ->  ThisWorkbook
     Sheet_달력.txt     ->  달력 시트
     Sheet_오늘.txt     ->  오늘 시트
     Sheet_설정.txt     ->  설정 시트
     Sheet_입력.txt     ->  입력 시트  (숨김 시트라 프로젝트 창에서만 보입니다)

4) 저장하고 파일을 다시 엽니다.
   입력 화면은 UserForm 대신 "입력" 시트가 대신합니다. 기능은 모두 동일합니다.

5) 검증
   VBE 에서 modTest 모듈을 열고 RunAllTests 안에 커서를 둔 뒤 F5 를 누르면
   10가지 검증이 전부 실행됩니다.
'@
[System.IO.File]::WriteAllText((Join-Path $manual '읽어보세요.txt'), $guide, $enc949)
Ok ("수동 임포트용 소스: " + $manual)

# --- 11. 저장 -------------------------------------------------
$step = '11. 저장'
# 한셀로 만든 것은 코드가 안 들어간 반쪽짜리다. 배포용과 절대 섞이지 않게 이름을 나눈다.
if ($isHCell) {
    $out = Join-Path $dist 'Scheduler_한셀테스트.xlsm'
} else {
    $out = Join-Path $dist 'Scheduler.xlsm'
}
if (Test-Path $out) {
    try { Remove-Item $out -Force } catch { Die ("기존 파일을 지울 수 없습니다. 열려 있는지 확인해 주세요: " + $out) }
}
try {
    $wb.SaveAs($out, 52)          # xlOpenXMLWorkbookMacroEnabled
} catch {
    Die ("저장하지 못했습니다: " + $_.Exception.Message)
}
Ok $out

Cleanup

Say ""
Say "===== 빌드 완료 ====="
Say ""
Say ("  결과물: " + $out)
if ($useForms) {
    Say "  입력 방식: UserForm"
} else {
    Say "  입력 방식: 시트 패널 (UserForm 없음)"
}
Say ""
if ($isHCell) {
    Write-Host "※ 이 파일은 한셀로 만든 테스트용입니다. 배포하지 마세요." -ForegroundColor Yellow
    Write-Host "   VBA 코드가 들어 있지 않습니다. 시트와 표만 있습니다." -ForegroundColor Yellow
    Say ""
    Say "한셀에서 이어서 할 일"
    Say ("  1) 한셀로 " + $out + " 열기")
    Say "  2) 도구 > 매크로 > 매크로 편집 (또는 Alt+F11) 으로 편집기 열기"
    Say ("  3) dist\manual_import 의 .bas 10개를 [파일 > 파일 가져오기] 로 가져오기")
    Say "     (modForms.bas 는 넣지 마세요. UserForm 이 없습니다)"
    Say "  4) 같은 폴더의 시트 이벤트 코드(.txt)를 각 시트에 붙여넣기"
    Say "  5) 저장하고 다시 열기 -> modTest.RunAllTests 를 F5 로 실행"
    Say ""
    Say "실제 배포본은 Excel 이 있는 PC에서 build\build.ps1 을 그냥 돌리면 됩니다."
} else {
    Say "다음 할 일"
    Say "  1) 파일이 있는 폴더를 Excel 의 '신뢰할 수 있는 위치' 로 등록 (README 참고)"
    Say "  2) 파일을 열고 이름을 입력"
    Say "  3) 설정 시트에서 [명부 갱신] 으로 직원 명부 CSV 를 읽어오기"
    Say "  4) VBE(Alt+F11) 에서 modTest.RunAllTests 를 F5 로 실행해 검증"
}
Say ""
exit 0

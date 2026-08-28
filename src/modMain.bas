Attribute VB_Name = "modMain"
Option Explicit
'==============================================================
' modMain - 진입점, 버튼 핸들러, 폼/패널 분기
'
' UserForm 이 만들어지지 못한 환경(VBA 프로젝트 개체 모델 접근이 잠긴 PC)에서도
' 프로그램이 동작해야 한다. 그래서 modMain 은 폼 이름을 직접 참조하지 않고
' Application.Run 으로 후기 호출하고, 실패하면 시트 패널로 넘어간다.
' (폼 이름을 직접 쓰면 폼이 없는 프로젝트는 컴파일 자체가 안 된다)
'==============================================================

' 폼/패널이 주고받는 값
Private mEditID As String
Private mDetailTitle As String
Private mDetailBody As String
Private mPickData As Variant
Private mPickResult As Long

'==============================================================
' 시작 / 종료
'==============================================================
Public Sub InitApp()
    On Error GoTo ErrHandler

    TB_ClearCache
    EmpInvalidate

    ' 처음 실행이면 이름을 받는다
    If Len(CfgGet("내이름", "")) = 0 Then
        Dim nm As String
        nm = InputBox("이름을 입력해 주세요." & vbLf & _
                      "업무공유를 주고받을 때 작성자로 기록됩니다.", APP_NAME, Environ$("USERNAME"))
        If Len(Trim$(nm)) > 0 Then CfgSet "내이름", Trim$(nm)
    End If

    DoBackup False

    SetupRender
    RenderCalendar
    RenderToday

    On Error Resume Next
    ThisWorkbook.Worksheets(SH_TODAY).Activate
    On Error GoTo 0
    Exit Sub

ErrHandler:
    ErrSay "modMain.InitApp"
End Sub

Public Sub RenderAll()
    On Error GoTo ErrHandler
    TB_ClearCache
    EmpInvalidate
    SetupRender
    RenderCalendar
    RenderToday
    Exit Sub
ErrHandler:
    ErrSay "modMain.RenderAll"
End Sub

'==============================================================
' 자동 백업
'   VBA 통합 문서는 실제로 손상되는 일이 있다. 이건 선택이 아니다.
'==============================================================
Public Sub DoBackup(ByVal verbose As Boolean)
    Dim folder As String, target As String
    Dim keepDays As Long, fname As String, fdate As Date
    Dim victims As Collection, i As Long

    On Error GoTo ErrHandler

    If Len(ThisWorkbook.Path) = 0 Then
        If verbose Then Warn "파일을 한 번 저장한 뒤에 백업할 수 있습니다."
        Exit Sub
    End If

    folder = ThisWorkbook.Path & "\backup"
    EnsureFolder folder
    target = folder & "\Scheduler_" & Format$(Date, "yyyymmdd") & ".xlsm"

    If Len(Dir$(target)) = 0 Then
        ThisWorkbook.SaveCopyAs target
        If verbose Then Say "백업했습니다." & vbLf & target
    Else
        If verbose Then Say "오늘 백업은 이미 만들어져 있습니다." & vbLf & target
    End If

    ' 보관일수 초과분 정리
    keepDays = CLng(Val(CfgGet("백업보관일수", "30")))
    If keepDays < 1 Then keepDays = 30

    Set victims = New Collection
    fname = Dir$(folder & "\Scheduler_*.xlsm")
    Do While Len(fname) > 0
        If Len(fname) >= 22 Then
            fdate = ParseYmd(Mid$(fname, 11, 8))
            If fdate > 0 And fdate < Date - keepDays Then victims.Add fname
        End If
        fname = Dir$()
    Loop
    For i = 1 To victims.Count
        On Error Resume Next
        Kill folder & "\" & victims(i)
        On Error GoTo ErrHandler
    Next i
    Exit Sub

ErrHandler:
    If verbose Then ErrSay "modMain.DoBackup" Else FastReset
End Sub

Public Sub BackupNow()
    DoBackup True
End Sub

Private Function ParseYmd(ByVal s0 As String) As Date
    ParseYmd = 0
    If Len(s0) <> 8 Then Exit Function
    If Not IsNumeric(s0) Then Exit Function
    On Error Resume Next
    ParseYmd = DateSerial(CLng(Left$(s0, 4)), CLng(Mid$(s0, 5, 2)), CLng(Right$(s0, 2)))
    If Err.Number <> 0 Then
        ParseYmd = 0
        Err.Clear
    End If
End Function

'==============================================================
' 폼/패널 분기
'==============================================================
Private Function UseForms() As Boolean
    UseForms = (CfgGet("입력방식", "패널") = "폼")
End Function

' 인자 없는 UI 프로시저를 후기 호출한다. 없으면 False.
Private Function RunUI(ByVal procName As String) As Boolean
    On Error GoTo Fail
    If Not UseForms() Then Exit Function
    Application.Run procName
    RunUI = True
    Exit Function
Fail:
    Err.Clear
    RunUI = False
End Function

' 폼/패널이 읽어가는 값들
Public Property Get EditID() As String
    EditID = mEditID
End Property

Public Property Get DetailTitle() As String
    DetailTitle = mDetailTitle
End Property

Public Property Get DetailBody() As String
    DetailBody = mDetailBody
End Property

Public Property Get PickerData() As Variant
    PickerData = mPickData
End Property

Public Property Let PickerResult(ByVal v As Long)
    mPickResult = v
End Property

'==============================================================
' 등록 / 수정
'==============================================================
Public Sub AddTask()
    EditTask ""
End Sub

Public Sub EditTask(ByVal id As String)
    mEditID = id
    If Not RunUI("UI_Task") Then modPanel.PanelOpen "T", id
End Sub

Public Sub AddContact()
    EditContact ""
End Sub

Public Sub EditContact(ByVal id As String)
    mEditID = id
    If Not RunUI("UI_Contact") Then modPanel.PanelOpen "C", id
End Sub

Public Sub AddShare()
    EditShare ""
End Sub

Public Sub EditShare(ByVal id As String)
    mEditID = id
    If Not RunUI("UI_Share") Then modPanel.PanelOpen "S", id
End Sub

Public Sub AddLunch()
    EditLunch ""
End Sub

Public Sub EditLunch(ByVal id As String)
    mEditID = id
    If Not RunUI("UI_Lunch") Then modPanel.PanelOpen "L", id
End Sub

'--------------------------------------------------------------
' 내용 보기
'--------------------------------------------------------------
Public Sub ShowDetail(ByVal title As String, ByVal body As String)
    mDetailTitle = title
    mDetailBody = body
    If Len(body) < 200 Then
        MsgBox body, vbInformation, title
        Exit Sub
    End If
    If Not RunUI("UI_Detail") Then
        MsgBox body, vbInformation, title
    End If
End Sub

'--------------------------------------------------------------
' 동명이인 선택
'   matches: (1..n, 1..4) = 이름, 부서, 직급, 내선
'   반환: 선택한 행 번호 (취소하면 0)
'--------------------------------------------------------------
Public Function PickEmployee(ByRef matches As Variant) As Long
    Dim i As Long, n As Long, msg As String, ans As String

    If IsEmpty(matches) Then Exit Function
    n = UBound(matches, 1)
    If n = 1 Then
        PickEmployee = 1                  ' 1건이면 팝업을 띄우지 않는다
        Exit Function
    End If

    mPickData = matches
    mPickResult = 0
    If RunUI("UI_Picker") Then
        PickEmployee = mPickResult
        Exit Function
    End If

    ' 폼이 없을 때: 번호를 매겨 고르게 한다
    msg = "같은 이름이 " & n & "명 있습니다. 번호를 입력해 주세요." & vbLf & vbLf
    For i = 1 To n
        msg = msg & i & ". " & matches(i, 1) & _
              IIf(Len(S(matches(i, 2))) > 0, " | " & matches(i, 2), "") & _
              IIf(Len(S(matches(i, 3))) > 0, " | " & matches(i, 3), "") & _
              " | 내선 " & matches(i, 4) & vbLf
    Next i
    ans = InputBox(msg, APP_NAME, "1")
    If Len(Trim$(ans)) = 0 Then Exit Function
    If Not IsNumeric(ans) Then Exit Function
    If CLng(ans) < 1 Or CLng(ans) > n Then Exit Function
    PickEmployee = CLng(ans)
End Function

'==============================================================
' 업무공유 내보내기 / 가져오기
'==============================================================
Public Sub ShareExport()
    Dim parts As Variant, cnt As Long, i As Long, ok As Boolean
    On Error GoTo ErrHandler

    If RunUI("UI_Export") Then Exit Sub

    ' 폼이 없을 때: 전체 범위로 파트를 만들어 하나씩 클립보드에 넣어 준다
    parts = SyncBuildParts(SCOPE_ALL, 30, cnt)
    If IsEmpty(parts) Then
        Say "내보낼 공유 건이 없습니다."
        Exit Sub
    End If
    For i = 1 To UBound(parts)
        ok = ClipSet(CStr(parts(i)))
        If Not ok Then
            Warn "클립보드에 복사하지 못했습니다."
            Exit Sub
        End If
        If i < UBound(parts) Then
            If Not Ask(cnt & "건 중 " & i & "/" & UBound(parts) & " 조각을 복사했습니다." & vbLf & _
                       "메신저에 붙여넣은 뒤 [예]를 누르면 다음 조각을 복사합니다.") Then Exit Sub
        Else
            Say cnt & "건 중 " & i & "/" & UBound(parts) & " 조각을 복사했습니다. (마지막 조각)" & vbLf & _
                "메신저에 붙여넣어 주세요."
        End If
    Next i
    Exit Sub

ErrHandler:
    ErrSay "modMain.ShareExport"
End Sub

Public Sub ShareImport()
    Dim txt As String, msg As String, st As Long
    On Error GoTo ErrHandler

    SyncResetInbox
    If RunUI("UI_Import") Then Exit Sub

    ' 폼이 없을 때: 클립보드에서 바로 읽어 들인다
    Do
        txt = ClipGet()
        If Len(txt) = 0 Then
            Warn "클립보드가 비어 있습니다. 받은 텍스트를 복사한 뒤 다시 눌러 주세요."
            Exit Sub
        End If
        st = SyncFeed(txt, msg)
        Select Case st
            Case FD_COMPLETE
                If Ask(msg & vbLf & vbLf & "반영할까요?") Then
                    Say SyncApply()
                    RenderAll
                End If
                Exit Sub
            Case FD_PARTIAL, FD_DUPLICATE
                If Not Ask(msg & vbLf & vbLf & _
                           "다음 조각을 복사한 뒤 [예]를 눌러 주세요.") Then Exit Sub
            Case Else
                Warn msg
                Exit Sub
        End Select
    Loop
    Exit Sub

ErrHandler:
    ErrSay "modMain.ShareImport"
End Sub

'==============================================================
' 설정 시트
'==============================================================
Public Sub SetupRender()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    If ws Is Nothing Then Exit Sub
    ' 이벤트를 직접 껐다 켜면 이 함수를 부른 쪽의 상태를 망가뜨린다.
    ' FastOn/FastOff 는 중첩을 세므로 안전하다.
    FastOn
    ws.Range("C4").Value = CfgGet("내이름", "")
    ws.Range("C5").Value = CfgGet("명부파일경로", "")
    ws.Range("C6").Value = EmpCount() & "명" & _
        IIf(Len(CfgGet("명부갱신시각", "")) > 0, "  (갱신 " & CfgGet("명부갱신시각", "") & ")", "")
    ws.Range("C7").Value = CfgGet("백업보관일수", "30")
    ws.Range("C8").Value = CfgGet("마지막동기화시각", "(없음)")
    ws.Range("C9").Value = CfgGet("입력방식", "패널")
    FastOff
End Sub

' 설정 시트에서 값을 고쳤을 때
Public Sub SetupChanged(ByVal addr As String, ByVal v As String)
    Select Case addr
        Case "$C$4": CfgSet "내이름", Trim$(v)
        Case "$C$5": CfgSet "명부파일경로", Trim$(v)
        Case "$C$7": CfgSet "백업보관일수", Trim$(v)
    End Select
End Sub

Public Sub SetupClick(ByVal r As Long, ByVal c As Long)
    On Error GoTo ErrHandler
    If c <> 2 And c <> 4 Then Exit Sub
    Select Case r
        Case 6:  If c = 2 Then RefreshEmployee
        Case 11: If c = 2 Then FullExport Else FullImport
        Case 12: If c = 2 Then BackupNow
        Case 13: If c = 2 Then RenderAll
    End Select
    Exit Sub
ErrHandler:
    ErrSay "modMain.SetupClick"
End Sub

'--------------------------------------------------------------
' 명부 갱신
'--------------------------------------------------------------
Public Sub RefreshEmployee()
    Dim path As String, cnt As Long, msg As String
    On Error GoTo ErrHandler

    path = Application.GetOpenFilename( _
        "명부 파일 (*.csv;*.xlsx;*.xls;*.xlsm),*.csv;*.xlsx;*.xls;*.xlsm", _
        1, "직원 명부 파일을 선택하세요")
    If VarType(path) = vbBoolean Then Exit Sub
    If Len(path) = 0 Then Exit Sub

    FastOn
    EmpRefreshFromFile path, cnt, msg
    FastOff

    If cnt > 0 Then
        SetupRender
        RenderToday
        Say msg
    Else
        Warn msg
    End If
    Exit Sub

ErrHandler:
    ErrSay "modMain.RefreshEmployee"
End Sub

'==============================================================
' 전체 데이터 내보내기 / 가져오기
'   나중에 프로그램을 v2 로 올릴 때의 데이터 이관 통로다.
'==============================================================
Private Const FULL_TAG As String = "###WSCHFULL1###"

Public Sub FullExport()
    Dim path As Variant, txt As String
    On Error GoTo ErrHandler

    path = Application.GetSaveAsFilename( _
        "스케줄러데이터_" & Format$(Now, "yyyymmdd_hhnn") & ".txt", _
        "텍스트 파일 (*.txt),*.txt", 1, "전체 데이터 내보내기")
    If VarType(path) = vbBoolean Then Exit Sub

    txt = FULL_TAG & vbCrLf
    txt = txt & DumpTable(TBL_TASK)
    txt = txt & DumpTable(TBL_CONTACT)
    txt = txt & DumpTable(TBL_SHARE)
    txt = txt & DumpTable(TBL_LUNCH)
    txt = txt & DumpTable(TBL_EMP)
    txt = txt & DumpTable(TBL_CFG)
    txt = txt & "#END" & vbCrLf

    WriteUtf8 CStr(path), txt
    Say "전체 데이터를 저장했습니다." & vbLf & path
    Exit Sub

ErrHandler:
    ErrSay "modMain.FullExport"
End Sub

Private Function DumpTable(ByVal tName As String) As String
    Dim lo As ListObject, data As Variant
    Dim r As Long, c As Long, ln As String, res As String
    Set lo = TB(tName)

    res = "#TABLE " & tName & vbCrLf & "#COLS "
    For c = 1 To lo.ListColumns.Count
        If c > 1 Then res = res & vbTab
        res = res & lo.ListColumns(c).Name
    Next c
    res = res & vbCrLf

    data = TB_Data(tName)
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            ln = ""
            For c = 1 To UBound(data, 2)
                If c > 1 Then ln = ln & vbTab
                ln = ln & Esc(S(data(r, c)))
            Next c
            If Len(Replace$(Replace$(ln, vbTab, ""), " ", "")) > 0 Then
                res = res & ln & vbCrLf
            End If
        Next r
    End If
    DumpTable = res
End Function

Public Sub FullImport()
    Dim path As Variant, txt As String, lines() As String
    Dim i As Long, ln As String, curTbl As String
    Dim f() As String, rec As Variant, c As Long, n As Long
    On Error GoTo ErrHandler

    path = Application.GetOpenFilename("텍스트 파일 (*.txt),*.txt", 1, "전체 데이터 가져오기")
    If VarType(path) = vbBoolean Then Exit Sub

    txt = ReadUtf8(CStr(path))
    If InStr(txt, FULL_TAG) = 0 Then
        Warn "이 프로그램이 만든 전체 데이터 파일이 아닙니다."
        Exit Sub
    End If
    If Not Ask("현재 데이터를 전부 지우고 파일 내용으로 덮어씁니다." & vbLf & _
               "계속할까요?") Then Exit Sub

    FastOn
    txt = Replace$(Replace$(txt, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If Left$(ln, 7) = "#TABLE " Then
            curTbl = Trim$(Mid$(ln, 8))
            ClearTable curTbl
        ElseIf Left$(ln, 6) = "#COLS " Then
            ' 컬럼 구성은 참고용. 순서는 현재 표 기준으로 그대로 넣는다.
        ElseIf Left$(ln, 1) = "#" Then
            ' 무시
        ElseIf Len(Trim$(ln)) > 0 And Len(curTbl) > 0 Then
            f = Split(ln, vbTab)
            rec = TB_NewRec(curTbl)
            n = TB_ColCount(curTbl)
            For c = 1 To n
                If c - 1 <= UBound(f) Then rec(c) = Unesc(f(c - 1))
            Next c
            TB_Insert curTbl, rec
        End If
    Next i
    FastOff

    TB_ClearCache
    EmpInvalidate
    RenderAll
    Say "전체 데이터를 가져왔습니다."
    Exit Sub

ErrHandler:
    ErrSay "modMain.FullImport"
End Sub

Private Sub ClearTable(ByVal tName As String)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = TB(tName)
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
End Sub

Public Sub WriteUtf8(ByVal path As String, ByVal txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2                 ' adSaveCreateOverWrite
    st.Close
End Sub

Public Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)
    st.Close
    If Len(ReadUtf8) > 0 Then
        If AscW(Left$(ReadUtf8, 1)) = &HFEFF Then ReadUtf8 = Mid$(ReadUtf8, 2)
    End If
End Function

'==============================================================
' 삭제 (오늘 시트에서 선택한 행)
'==============================================================
Public Sub DeleteSelected()
    Dim ws As Worksheet, meta As String, r As Long
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_TODAY)
    If ActiveSheet.Name <> SH_TODAY Then
        Warn "'오늘' 시트에서 지울 항목의 줄을 선택한 뒤 실행해 주세요."
        Exit Sub
    End If
    r = Selection.Row
    meta = S(ws.Cells(r, TD_META_COL).Value)
    If Len(meta) = 0 Then
        Warn "지울 항목의 줄을 선택해 주세요."
        Exit Sub
    End If
    Select Case Left$(meta, 1)
        Case "T": If DeleteRecord(TBL_TASK, Mid$(meta, 3)) Then RenderAll
        Case "C": If DeleteRecord(TBL_CONTACT, Mid$(meta, 3)) Then RenderAll
        Case "S": If DeleteRecord(TBL_SHARE, Mid$(meta, 3)) Then RenderAll
    End Select
    Exit Sub
ErrHandler:
    ErrSay "modMain.DeleteSelected"
End Sub

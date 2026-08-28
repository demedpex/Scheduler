Attribute VB_Name = "modPanel"
Option Explicit
'==============================================================
' modPanel - 시트 기반 입력 패널 (UserForm 폴백 경로)
'
' 사내 보안정책으로 "VBA 프로젝트 개체 모델에 대한 액세스 신뢰" 를 켤 수 없으면
' build.ps1 이 UserForm 을 만들지 못한다. 그때도 프로그램이 온전히 돌아가도록
' "입력" 시트를 폼처럼 쓴다.
'
' 시트 구조
'   B2        제목
'   B4:B10    항목 이름     C4:C10  입력값     E열  보조 버튼
'   B12 [저장]   D12 [취소]   F12 [삭제]
'   N1 종류(T/C/S/L)   N2 편집중인 ID   N4:N10 각 행의 필드명
'==============================================================

Public Const SH_PANEL As String = "입력"
Public Const PN_ROW_FIRST As Long = 4
Public Const PN_ROW_LAST As Long = 10
Public Const PN_ROW_BTN As Long = 12
Public Const PN_META_COL As Long = 14        ' N열

'--------------------------------------------------------------
' 패널 열기
'--------------------------------------------------------------
Public Sub PanelOpen(ByVal kind As String, ByVal id As String)
    Dim ws As Worksheet, rec As Variant
    Dim names As Variant, labels As Variant
    Dim i As Long, r As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)
    FastOn

    Select Case kind
        Case "T"
            labels = Array("제목", "내용", "날짜", "시간")
            names = Array("제목", "내용", "날짜", "시간")
            ws.Range("B2").Value = IIf(Len(id) = 0, "업무 등록", "업무 수정")
        Case "C"
            labels = Array("이름", "부서", "내선번호", "날짜", "시간")
            names = Array("이름", "부서", "내선캐시", "날짜", "시간")
            ws.Range("B2").Value = IIf(Len(id) = 0, "연락하기 등록", "연락하기 수정")
        Case "S"
            labels = Array("문서번호", "제목", "내용", "등록일", "회신기한")
            names = Array("문서번호", "제목", "내용", "등록일", "회신기한")
            ws.Range("B2").Value = IIf(Len(id) = 0, "업무공유 등록", "업무공유 수정")
        Case "L"
            labels = Array("대상이름", "날짜", "시간", "장소")
            names = Array("대상이름", "날짜", "시간", "장소")
            ws.Range("B2").Value = IIf(Len(id) = 0, "점심약속 등록", "점심약속 수정")
        Case Else
            FastOff
            Exit Sub
    End Select

    ' 초기화
    ws.Range(ws.Cells(PN_ROW_FIRST, 2), ws.Cells(PN_ROW_LAST, 6)).ClearContents
    ws.Range(ws.Cells(PN_ROW_FIRST, PN_META_COL), ws.Cells(PN_ROW_LAST, PN_META_COL)).ClearContents
    ws.Cells(1, PN_META_COL).Value = kind
    ws.Cells(2, PN_META_COL).Value = id

    rec = Empty
    If Len(id) > 0 Then rec = LoadRecord(TableOf(kind), id)

    For i = LBound(labels) To UBound(labels)
        r = PN_ROW_FIRST + i
        ws.Cells(r, 2).Value = labels(i)
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, PN_META_COL).Value = names(i)
        If Not IsEmpty(rec) Then
            ws.Cells(r, 3).Value = RecGet(rec, TableOf(kind), CStr(names(i)))
        End If
    Next i

    ' 기본값
    If Len(id) = 0 Then
        Select Case kind
            Case "T": SetField ws, "날짜", DateStr(ViewDate())
            Case "S"
                SetField ws, "등록일", TodayStr()
                SetField ws, "회신기한", DateStr(Date + 3)
            Case "L": SetField ws, "날짜", DateStr(ViewDate())
        End Select
    End If

    ' 연락하기에는 명부 검색 버튼을 붙인다
    If kind = "C" Then ws.Cells(PN_ROW_FIRST, 5).Value = "[명부에서 찾기]"

    ws.Cells(PN_ROW_BTN, 2).Value = "[저장]"
    ws.Cells(PN_ROW_BTN, 4).Value = "[취소]"
    If Len(id) > 0 Then
        ws.Cells(PN_ROW_BTN, 6).Value = "[삭제]"
    Else
        ws.Cells(PN_ROW_BTN, 6).ClearContents      ' 새 등록에는 삭제 버튼을 두지 않는다
    End If

    ' 시트 숨김/표시를 지원하지 않는 환경(한셀 등)에서도 입력은 되게 한다
    On Error Resume Next
    ws.Visible = xlSheetVisible
    On Error GoTo ErrHandler
    FastOff
    ws.Activate
    ws.Cells(PN_ROW_FIRST, 3).Select
    Exit Sub

ErrHandler:
    ErrSay "modPanel.PanelOpen"
End Sub

Private Function TableOf(ByVal kind As String) As String
    Select Case kind
        Case "T": TableOf = TBL_TASK
        Case "C": TableOf = TBL_CONTACT
        Case "S": TableOf = TBL_SHARE
        Case "L": TableOf = TBL_LUNCH
    End Select
End Function

Private Function GetField(ByRef ws As Worksheet, ByVal fieldName As String) As String
    Dim r As Long
    For r = PN_ROW_FIRST To PN_ROW_LAST
        If S(ws.Cells(r, PN_META_COL).Value) = fieldName Then
            GetField = S(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
End Function

Private Sub SetField(ByRef ws As Worksheet, ByVal fieldName As String, ByVal v As String)
    Dim r As Long
    For r = PN_ROW_FIRST To PN_ROW_LAST
        If S(ws.Cells(r, PN_META_COL).Value) = fieldName Then
            ws.Cells(r, 3).Value = v
            Exit Sub
        End If
    Next r
End Sub

'--------------------------------------------------------------
' 패널 클릭 처리
'--------------------------------------------------------------
Public Sub PanelClick(ByVal r As Long, ByVal c As Long)
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)

    If r = PN_ROW_BTN Then
        Select Case c
            Case 2: PanelSave
            Case 4: PanelClose
            Case 6: PanelDelete
        End Select
        Exit Sub
    End If

    If c = 5 And S(ws.Cells(r, 5).Value) = "[명부에서 찾기]" Then LookupEmployee
    Exit Sub

ErrHandler:
    ErrSay "modPanel.PanelClick"
End Sub

'--------------------------------------------------------------
' 저장
'--------------------------------------------------------------
Public Sub PanelSave()
    Dim ws As Worksheet, kind As String, id As String, msg As String
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)
    kind = S(ws.Cells(1, PN_META_COL).Value)
    id = S(ws.Cells(2, PN_META_COL).Value)

    Select Case kind
        Case "T"
            msg = SaveTask(id, GetField(ws, "제목"), GetField(ws, "내용"), _
                           GetField(ws, "날짜"), GetField(ws, "시간"))
        Case "C"
            msg = SaveContact(id, GetField(ws, "이름"), GetField(ws, "부서"), _
                              GetField(ws, "내선캐시"), GetField(ws, "날짜"), GetField(ws, "시간"))
        Case "S"
            msg = SaveShare(id, GetField(ws, "문서번호"), GetField(ws, "제목"), _
                            GetField(ws, "내용"), GetField(ws, "등록일"), GetField(ws, "회신기한"))
        Case "L"
            msg = SaveLunch(id, GetField(ws, "대상이름"), GetField(ws, "날짜"), _
                            GetField(ws, "시간"), GetField(ws, "장소"))
        Case Else
            Exit Sub
    End Select

    If msg = "-" Then Exit Sub                ' 사용자가 확인 창에서 취소함
    If Len(msg) > 0 Then
        Warn msg
        Exit Sub
    End If

    PanelClose
    Exit Sub

ErrHandler:
    ErrSay "modPanel.PanelSave"
End Sub

Public Sub PanelDelete()
    Dim ws As Worksheet, kind As String, id As String
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)
    kind = S(ws.Cells(1, PN_META_COL).Value)
    id = S(ws.Cells(2, PN_META_COL).Value)
    If Len(id) = 0 Then Exit Sub
    If DeleteRecord(TableOf(kind), id) Then PanelClose
    Exit Sub
ErrHandler:
    ErrSay "modPanel.PanelDelete"
End Sub

Public Sub PanelClose()
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)
    FastOn
    ws.Range(ws.Cells(PN_ROW_FIRST, 2), ws.Cells(PN_ROW_BTN, 6)).ClearContents
    ws.Cells(1, PN_META_COL).ClearContents
    ws.Cells(2, PN_META_COL).ClearContents
    ThisWorkbook.Worksheets(SH_TODAY).Activate
    On Error Resume Next
    ws.Visible = xlSheetHidden
    On Error GoTo ErrHandler
    FastOff
    RenderCalendar
    RenderToday
    Exit Sub
ErrHandler:
    ErrSay "modPanel.PanelClose"
End Sub

'--------------------------------------------------------------
' 명부에서 찾기 (동명이인 처리)
'--------------------------------------------------------------
Public Sub LookupEmployee()
    Dim ws As Worksheet, q As String, matches As Variant, pick As Long
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_PANEL)

    q = GetField(ws, "이름")
    If Len(q) = 0 Then
        Warn "이름을 먼저 입력해 주세요. 일부만 입력해도 찾습니다."
        Exit Sub
    End If

    matches = EmpSearch(q)
    If IsEmpty(matches) Then
        If Ask("'" & q & "' 은(는) 명부에 없습니다." & vbLf & _
               "이름과 내선번호를 직접 입력하시겠습니까?") Then
            SetField ws, "내선캐시", InputBox("내선번호를 입력해 주세요.", APP_NAME, GetField(ws, "내선캐시"))
        End If
        Exit Sub
    End If

    pick = PickEmployee(matches)
    If pick = 0 Then Exit Sub

    SetField ws, "이름", CStr(matches(pick, 1))
    SetField ws, "부서", CStr(matches(pick, 2))
    SetField ws, "내선캐시", ""
    Say matches(pick, 1) & " / " & matches(pick, 2) & " / 내선 " & matches(pick, 4)
    Exit Sub

ErrHandler:
    ErrSay "modPanel.LookupEmployee"
End Sub

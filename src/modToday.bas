Attribute VB_Name = "modToday"
Option Explicit
'==============================================================
' modToday - "오늘" 시트 렌더링
'
' 원칙: 입력은 폼, 조회·목록은 시트에 그린다.
' 시트는 셀 서식만 바꾸면 취소선·회색처리가 끝나고 인쇄도 된다.
'
' 시트 구조
'   2행  [◀] 날짜 [오늘] [▶]        [공유 내보내기] [공유 가져오기]
'   3행  [업무 등록] [연락 등록] [점심 등록] [공유 등록] [새로고침]
'   5행~ 섹션 3개 (업무 / 유관부서 연락하기 / 업무공유)
'   N열  숨김. 각 행이 어떤 레코드인지 "T|ID" 형태로 적어 둔다.
'        모듈 변수 대신 시트에 적어 두므로 VBA 프로젝트가 리셋돼도 클릭이 동작한다.
'==============================================================

Public Const TD_ROW_NAV As Long = 2
Public Const TD_ROW_BTN As Long = 3
Public Const TD_ROW_START As Long = 5
Public Const TD_META_COL As Long = 14        ' N열
Public Const TD_MAX_ROW As Long = 400

Private Const CLR_HEAD As Long = 15921906    ' 연회색 배경
Private Const CLR_GRAY As Long = 10921638    ' 완료 항목 글자색
Private Const CLR_RED As Long = 255
Private Const CLR_ORANGE As Long = 26367
Private Const CLR_BTN As Long = 15130800     ' 버튼처럼 보이는 셀

'--------------------------------------------------------------
' 조회 날짜
'--------------------------------------------------------------
Public Function ViewDate() As Date
    Dim d As Date
    d = ParseDate(CfgGet("조회날짜", ""))
    If d = 0 Then d = Date
    ViewDate = d
End Function

Public Sub SetViewDate(ByVal d As Date)
    CfgSet "조회날짜", DateStr(d)
End Sub

Public Sub TodayShiftDay(ByVal delta As Long)
    SetViewDate ViewDate() + delta
    RenderToday
End Sub

Public Sub TodayGoToday()
    SetViewDate Date
    RenderToday
End Sub

'==============================================================
' 렌더링
'==============================================================
Public Sub RenderToday()
    Dim ws As Worksheet, r As Long, d As Date
    On Error GoTo ErrHandler
    FastOn

    Set ws = ThisWorkbook.Worksheets(SH_TODAY)
    d = ViewDate()

    ClearBody ws

    ' 상단 날짜 표시
    ws.Cells(TD_ROW_NAV, 3).Value = Format$(d, "yyyy-mm-dd") & " (" & WeekdayKo(d) & ")"

    r = TD_ROW_START
    r = DrawTasks(ws, d, r)
    r = r + 1
    r = DrawContacts(ws, d, r)
    r = r + 1
    r = DrawShares(ws, d, r)

    FastOff
    Exit Sub
ErrHandler:
    ErrSay "modToday.RenderToday"
End Sub

Private Sub ClearBody(ByRef ws As Worksheet)
    With ws.Range(ws.Cells(TD_ROW_START, 2), ws.Cells(TD_MAX_ROW, TD_META_COL))
        .ClearContents
        .Font.Strikethrough = False
        .Font.Bold = False
        .Font.Color = vbBlack
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Function WeekdayKo(ByVal d As Date) As String
    WeekdayKo = Mid$("일월화수목금토", Weekday(d), 1)
End Function

'--------------------------------------------------------------
' 1) 업무
'--------------------------------------------------------------
Private Function DrawTasks(ByRef ws As Worksheet, ByVal d As Date, ByVal startRow As Long) As Long
    Dim data As Variant, idx As Collection
    Dim r As Long, row As Long, i As Long
    Dim rec As Variant, done As Boolean, tm As String
    Dim dstr As String

    dstr = DateStr(d)
    row = startRow
    SectionHeader ws, row, "■ 업무"
    row = row + 1

    data = TB_Data(TBL_TASK)
    Set idx = New Collection
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            rec = RowOf(data, r)
            If Not ToBool(RecGet(rec, TBL_TASK, "삭제여부")) Then
                If RecGet(rec, TBL_TASK, "날짜") = dstr Then idx.Add r
            End If
        Next r
    End If

    If idx.Count = 0 Then
        EmptyLine ws, row, "등록된 업무가 없습니다."
        DrawTasks = row + 1
        Exit Function
    End If

    Set idx = SortRows(data, idx, TB_Col(TBL_TASK, "시간"))

    For i = 1 To idx.Count
        rec = RowOf(data, idx(i))
        done = ToBool(RecGet(rec, TBL_TASK, "완료여부"))
        tm = RecGet(rec, TBL_TASK, "시간")

        ws.Cells(row, 2).Value = IIf(done, ChrW$(9745), ChrW$(9744))
        ws.Cells(row, 2).HorizontalAlignment = xlCenter
        ws.Cells(row, 3).Value = RecGet(rec, TBL_TASK, "제목")
        ws.Cells(row, 4).Value = tm
        ws.Cells(row, TD_META_COL).Value = "T|" & RecGet(rec, TBL_TASK, "ID")

        If done Then
            With ws.Range(ws.Cells(row, 2), ws.Cells(row, 7))
                .Font.Strikethrough = True
                .Font.Color = CLR_GRAY
            End With
        End If
        row = row + 1
        If row > TD_MAX_ROW - 10 Then Exit For
    Next i

    DrawTasks = row
End Function

'--------------------------------------------------------------
' 2) 유관부서 연락하기
'    내선번호는 저장값이 아니라 명부에서 매번 조회한 결과를 쓴다.
'--------------------------------------------------------------
Private Function DrawContacts(ByRef ws As Worksheet, ByVal d As Date, ByVal startRow As Long) As Long
    Dim data As Variant, idx As Collection
    Dim r As Long, row As Long, i As Long
    Dim rec As Variant, done As Boolean
    Dim nm As String, dept As String, ext As String, dstr As String
    Dim isToday As Boolean, recDate As String

    dstr = DateStr(d)
    isToday = (d = Date)
    row = startRow
    SectionHeader ws, row, "■ 유관부서 연락하기"
    row = row + 1

    data = TB_Data(TBL_CONTACT)
    Set idx = New Collection
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            rec = RowOf(data, r)
            If Not ToBool(RecGet(rec, TBL_CONTACT, "삭제여부")) Then
                recDate = RecGet(rec, TBL_CONTACT, "날짜")
                If recDate = dstr Then
                    idx.Add r
                ElseIf Len(recDate) = 0 And isToday And _
                       Not ToBool(RecGet(rec, TBL_CONTACT, "완료여부")) Then
                    ' 날짜를 안 정한 연락 건은 오늘 화면에서만 보여준다
                    idx.Add r
                End If
            End If
        Next r
    End If

    If idx.Count = 0 Then
        EmptyLine ws, row, "연락할 건이 없습니다."
        DrawContacts = row + 1
        Exit Function
    End If

    Set idx = SortRows(data, idx, TB_Col(TBL_CONTACT, "시간"))

    For i = 1 To idx.Count
        rec = RowOf(data, idx(i))
        done = ToBool(RecGet(rec, TBL_CONTACT, "완료여부"))
        nm = RecGet(rec, TBL_CONTACT, "이름")
        dept = RecGet(rec, TBL_CONTACT, "부서")

        ext = EmpExtension(nm, dept)
        If Len(ext) = 0 Then ext = RecGet(rec, TBL_CONTACT, "내선캐시")
        If Len(ext) = 0 Then
            ext = "(명부 없음)"
        Else
            ext = "내선 " & ext
        End If

        ws.Cells(row, 2).Value = IIf(done, ChrW$(9745), ChrW$(9744))
        ws.Cells(row, 2).HorizontalAlignment = xlCenter
        If Len(dept) > 0 Then
            ws.Cells(row, 3).Value = nm & "(" & dept & ")"
        Else
            ws.Cells(row, 3).Value = nm
        End If
        ws.Cells(row, 4).Value = ext
        ws.Cells(row, 5).Value = RecGet(rec, TBL_CONTACT, "시간")
        ws.Cells(row, TD_META_COL).Value = "C|" & RecGet(rec, TBL_CONTACT, "ID")

        If done Then
            With ws.Range(ws.Cells(row, 2), ws.Cells(row, 7))
                .Font.Strikethrough = True
                .Font.Color = CLR_GRAY
            End With
        End If
        row = row + 1
        If row > TD_MAX_ROW - 5 Then Exit For
    Next i

    DrawContacts = row
End Function

'--------------------------------------------------------------
' 3) 업무공유
'    회신기한이 (조회날짜 - 7일) 이후인 살아있는 건을 기한순으로 보여준다.
'--------------------------------------------------------------
Private Function DrawShares(ByRef ws As Worksheet, ByVal d As Date, ByVal startRow As Long) As Long
    Dim data As Variant, idx As Collection
    Dim r As Long, row As Long, i As Long
    Dim rec As Variant, due As Date, dd As Long

    row = startRow
    SectionHeader ws, row, "■ 업무공유"
    ws.Cells(row, 6).Value = "[공유 내보내기]"
    ws.Cells(row, 7).Value = "[공유 가져오기]"
    ws.Range(ws.Cells(row, 6), ws.Cells(row, 7)).Font.Bold = True
    ws.Range(ws.Cells(row, 6), ws.Cells(row, 7)).HorizontalAlignment = xlCenter
    row = row + 1

    data = TB_Data(TBL_SHARE)
    Set idx = New Collection
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            rec = RowOf(data, r)
            If Not ToBool(RecGet(rec, TBL_SHARE, "삭제여부")) Then
                due = ParseDate(RecGet(rec, TBL_SHARE, "회신기한"))
                If due = 0 Or due >= d - 7 Then idx.Add r
            End If
        Next r
    End If

    If idx.Count = 0 Then
        EmptyLine ws, row, "공유된 업무가 없습니다."
        DrawShares = row + 1
        Exit Function
    End If

    Set idx = SortRows(data, idx, TB_Col(TBL_SHARE, "회신기한"))

    For i = 1 To idx.Count
        rec = RowOf(data, idx(i))
        due = ParseDate(RecGet(rec, TBL_SHARE, "회신기한"))

        ws.Cells(row, 3).Value = RecGet(rec, TBL_SHARE, "문서번호")
        ws.Cells(row, 4).Value = RecGet(rec, TBL_SHARE, "제목")
        ws.Cells(row, 7).Value = RecGet(rec, TBL_SHARE, "작성자")
        ws.Cells(row, TD_META_COL).Value = "S|" & RecGet(rec, TBL_SHARE, "ID")

        If due > 0 Then
            dd = CLng(due - d)
            ws.Cells(row, 6).Value = "회신기한 " & Format$(due, "mm/dd")
            If dd = 0 Then
                ws.Cells(row, 5).Value = "D-DAY"
                ws.Cells(row, 5).Font.Bold = True
                ws.Cells(row, 5).Font.Color = CLR_RED
            ElseIf dd < 0 Then
                ws.Cells(row, 5).Value = "D+" & (-dd)
                ws.Cells(row, 5).Font.Color = CLR_RED
            ElseIf dd <= 3 Then
                ws.Cells(row, 5).Value = "D-" & dd
                ws.Cells(row, 5).Font.Color = CLR_ORANGE
            Else
                ws.Cells(row, 5).Value = "D-" & dd
            End If
        End If

        row = row + 1
        If row > TD_MAX_ROW Then Exit For
    Next i

    DrawShares = row
End Function

'--------------------------------------------------------------
' 공통 그리기 도우미
'--------------------------------------------------------------
Private Sub SectionHeader(ByRef ws As Worksheet, ByVal row As Long, ByVal title As String)
    With ws.Range(ws.Cells(row, 2), ws.Cells(row, 7))
        .Interior.Color = CLR_HEAD
        .Font.Bold = True
    End With
    ws.Cells(row, 2).Value = title
End Sub

Private Sub EmptyLine(ByRef ws As Worksheet, ByVal row As Long, ByVal msg As String)
    ws.Cells(row, 3).Value = msg
    ws.Cells(row, 3).Font.Color = CLR_GRAY
End Sub

'--------------------------------------------------------------
' 정렬
'   키가 있는 항목이 먼저(오름차순), 키가 빈 항목은 뒤에 등록순 그대로.
'--------------------------------------------------------------
Private Function SortRows(ByRef data As Variant, ByRef idx As Collection, _
                          ByVal keyCol As Long) As Collection
    Dim withKey As Collection, noKey As Collection, res As Collection
    Dim i As Long, j As Long, r As Long
    Dim arr() As Long, keys() As String, n As Long
    Dim tmpL As Long, tmpS As String

    Set withKey = New Collection
    Set noKey = New Collection
    For i = 1 To idx.Count
        r = idx(i)
        If keyCol > 0 And Len(S(data(r, keyCol))) > 0 Then
            withKey.Add r
        Else
            noKey.Add r
        End If
    Next i

    n = withKey.Count
    Set res = New Collection
    If n > 0 Then
        ReDim arr(1 To n)
        ReDim keys(1 To n)
        For i = 1 To n
            arr(i) = withKey(i)
            keys(i) = S(data(withKey(i), keyCol))
        Next i
        ' 삽입 정렬. 하루치 목록이라 항목 수가 적어 이걸로 충분하다.
        For i = 2 To n
            tmpL = arr(i)
            tmpS = keys(i)
            j = i - 1
            Do While j >= 1
                If keys(j) > tmpS Then
                    keys(j + 1) = keys(j)
                    arr(j + 1) = arr(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            keys(j + 1) = tmpS
            arr(j + 1) = tmpL
        Next i
        For i = 1 To n
            res.Add arr(i)
        Next i
    End If
    For i = 1 To noKey.Count
        res.Add noKey(i)
    Next i

    Set SortRows = res
End Function

'==============================================================
' 클릭 처리 (오늘 시트의 SelectionChange 에서 호출)
'==============================================================
Public Sub TodayClick(ByVal r As Long, ByVal c As Long)
    Dim ws As Worksheet, meta As String, kind As String, id As String
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_TODAY)

    '---- 상단 버튼 ----
    If r = TD_ROW_NAV Then
        Select Case c
            Case 2: TodayShiftDay -1: Exit Sub
            Case 4: TodayGoToday: Exit Sub
            Case 5: TodayShiftDay 1: Exit Sub
        End Select
    ElseIf r = TD_ROW_BTN Then
        Select Case c
            Case 2, 3: modMain.AddTask: Exit Sub
            Case 4: modMain.AddContact: Exit Sub
            Case 5: modMain.AddLunch: Exit Sub
            Case 6: modMain.AddShare: Exit Sub
            Case 7: RenderToday: Exit Sub
        End Select
    End If

    '---- 본문 ----
    meta = S(ws.Cells(r, TD_META_COL).Value)

    ' 업무공유 섹션 머리줄의 내보내기/가져오기 버튼
    If Len(meta) = 0 Then
        If S(ws.Cells(r, 6).Value) = "[공유 내보내기]" And c = 6 Then
            modMain.ShareExport
            Exit Sub
        ElseIf S(ws.Cells(r, 7).Value) = "[공유 가져오기]" And c = 7 Then
            modMain.ShareImport
            Exit Sub
        End If
        Exit Sub
    End If

    kind = Left$(meta, 1)
    id = Mid$(meta, 3)

    Select Case kind
        Case "T"
            If c = 2 Then
                ToggleTask id
            ElseIf c = 3 Or c = 4 Then
                ShowTaskDetail id
            End If
        Case "C"
            If c = 2 Then
                ToggleContact id
            ElseIf c >= 3 And c <= 5 Then
                ShowContactDetail id
            End If
        Case "S"
            If c >= 3 And c <= 7 Then ShowShareDetail id
    End Select
    Exit Sub

ErrHandler:
    ErrSay "modToday.TodayClick"
End Sub

' 더블클릭 = 수정
Public Sub TodayDblClick(ByVal r As Long)
    Dim ws As Worksheet, meta As String
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_TODAY)
    meta = S(ws.Cells(r, TD_META_COL).Value)
    If Len(meta) = 0 Then Exit Sub
    Select Case Left$(meta, 1)
        Case "T": modMain.EditTask Mid$(meta, 3)
        Case "C": modMain.EditContact Mid$(meta, 3)
        Case "S": modMain.EditShare Mid$(meta, 3)
    End Select
    Exit Sub
ErrHandler:
    ErrSay "modToday.TodayDblClick"
End Sub

'--------------------------------------------------------------
' 완료 토글
'--------------------------------------------------------------
Public Sub ToggleTask(ByVal id As String)
    Dim r As Long, rec As Variant, done As Boolean
    r = TB_FindRowByID(TBL_TASK, id)
    If r = 0 Then Exit Sub
    rec = TB_GetRow(TBL_TASK, r)
    done = Not ToBool(RecGet(rec, TBL_TASK, "완료여부"))
    RecSet rec, TBL_TASK, "완료여부", BoolStr(done)
    RecSet rec, TBL_TASK, "완료시각", IIf(done, NowStamp(), "")
    RecSet rec, TBL_TASK, "수정시각", NowStamp()
    TB_UpdateRow TBL_TASK, r, rec
    RenderToday
End Sub

Public Sub ToggleContact(ByVal id As String)
    Dim r As Long, rec As Variant, done As Boolean
    r = TB_FindRowByID(TBL_CONTACT, id)
    If r = 0 Then Exit Sub
    rec = TB_GetRow(TBL_CONTACT, r)
    done = Not ToBool(RecGet(rec, TBL_CONTACT, "완료여부"))
    RecSet rec, TBL_CONTACT, "완료여부", BoolStr(done)
    RecSet rec, TBL_CONTACT, "수정시각", NowStamp()
    TB_UpdateRow TBL_CONTACT, r, rec
    RenderToday
End Sub

'--------------------------------------------------------------
' 상세 보기
'--------------------------------------------------------------
Public Sub ShowTaskDetail(ByVal id As String)
    Dim r As Long, rec As Variant, body As String
    r = TB_FindRowByID(TBL_TASK, id)
    If r = 0 Then Exit Sub
    rec = TB_GetRow(TBL_TASK, r)
    body = "날짜: " & RecGet(rec, TBL_TASK, "날짜")
    If Len(RecGet(rec, TBL_TASK, "시간")) > 0 Then body = body & " " & RecGet(rec, TBL_TASK, "시간")
    body = body & vbLf & String$(30, "-") & vbLf & RecGet(rec, TBL_TASK, "내용")
    modMain.ShowDetail RecGet(rec, TBL_TASK, "제목"), body
End Sub

Public Sub ShowContactDetail(ByVal id As String)
    Dim r As Long, rec As Variant, body As String, ext As String
    r = TB_FindRowByID(TBL_CONTACT, id)
    If r = 0 Then Exit Sub
    rec = TB_GetRow(TBL_CONTACT, r)
    ext = EmpExtension(RecGet(rec, TBL_CONTACT, "이름"), RecGet(rec, TBL_CONTACT, "부서"))
    If Len(ext) = 0 Then ext = RecGet(rec, TBL_CONTACT, "내선캐시")
    If Len(ext) = 0 Then ext = "(명부에 없음)"
    body = "부서: " & RecGet(rec, TBL_CONTACT, "부서") & vbLf & _
           "내선: " & ext & vbLf & _
           "일시: " & RecGet(rec, TBL_CONTACT, "날짜") & " " & RecGet(rec, TBL_CONTACT, "시간")
    modMain.ShowDetail RecGet(rec, TBL_CONTACT, "이름"), body
End Sub

Public Sub ShowShareDetail(ByVal id As String)
    Dim r As Long, rec As Variant, body As String
    r = TB_FindRowByID(TBL_SHARE, id)
    If r = 0 Then Exit Sub
    rec = TB_GetRow(TBL_SHARE, r)
    body = "문서번호: " & RecGet(rec, TBL_SHARE, "문서번호") & vbLf & _
           "등록일: " & RecGet(rec, TBL_SHARE, "등록일") & _
           "   회신기한: " & RecGet(rec, TBL_SHARE, "회신기한") & vbLf & _
           "작성자: " & RecGet(rec, TBL_SHARE, "작성자") & vbLf & _
           String$(30, "-") & vbLf & RecGet(rec, TBL_SHARE, "내용")
    modMain.ShowDetail RecGet(rec, TBL_SHARE, "제목"), body
End Sub

Attribute VB_Name = "modCalendar"
Option Explicit
'==============================================================
' modCalendar - "달력" 시트 렌더링
'
' 7열 x 6행 그리드를 시트에 미리 그려두고(build.ps1) VBA 는 텍스트만 채운다.
' UserForm 에 컨트롤 42개를 동적 생성하는 방식은 느리고 코드가 지저분해진다.
'
' 시트 구조
'   2행   [◀]  2026년 8월  [▶]      [오늘로]
'   4행   일 월 화 수 목 금 토   (B4:H4)
'   5~10행 날짜 칸 (B5:H10)
'==============================================================

Public Const CAL_ROW_NAV As Long = 2
Public Const CAL_ROW_HEAD As Long = 4
Public Const CAL_ROW_START As Long = 5
Public Const CAL_COL_START As Long = 2       ' B열
Public Const CAL_WEEKS As Long = 6

Private Const CLR_TODAY As Long = 14083324   ' 오늘 칸 배경 (연파랑)
Private Const CLR_URGENT As Long = 13684991  ' 회신기한 임박 (연빨강)
Private Const CLR_OTHER As Long = 15921906   ' 이번 달이 아닌 칸
Private Const CLR_SUN As Long = 255
Private Const CLR_SAT As Long = 12611584

'--------------------------------------------------------------
' 표시 중인 연월
'--------------------------------------------------------------
Public Function CalMonth() As Date
    Dim s0 As String, y As Long, m As Long
    s0 = CfgGet("달력연월", "")
    If Len(s0) = 7 Then
        If IsNumeric(Left$(s0, 4)) And IsNumeric(Right$(s0, 2)) Then
            y = CLng(Left$(s0, 4))
            m = CLng(Right$(s0, 2))
            If y >= 1900 And y <= 2999 And m >= 1 And m <= 12 Then
                CalMonth = DateSerial(y, m, 1)
                Exit Function
            End If
        End If
    End If
    CalMonth = DateSerial(Year(Date), Month(Date), 1)
End Function

Public Sub SetCalMonth(ByVal d As Date)
    CfgSet "달력연월", Format$(d, "yyyy-mm")
End Sub

Public Sub CalShiftMonth(ByVal delta As Long)
    SetCalMonth DateAdd("m", delta, CalMonth())
    RenderCalendar
End Sub

Public Sub CalGoToday()
    SetCalMonth DateSerial(Year(Date), Month(Date), 1)
    RenderCalendar
End Sub

'==============================================================
' 렌더링
'==============================================================
Public Sub RenderCalendar()
    Dim ws As Worksheet
    Dim first As Date, d As Date
    Dim dTask As Object, dCont As Object, dShare As Object, dLunch As Object, dUrgent As Object
    Dim startCol As Long, w As Long, c As Long
    Dim cell As Range, k As String
    Dim txt As String, line2 As String, line3 As String
    Dim inMonth As Boolean

    On Error GoTo ErrHandler
    FastOn

    Set ws = ThisWorkbook.Worksheets(SH_CAL)
    first = CalMonth()

    ws.Cells(CAL_ROW_NAV, 3).Value = Format$(first, "yyyy년 m월")

    CountByDate dTask, dCont, dShare, dLunch, dUrgent

    ' 달력의 첫 칸에 들어갈 날짜 (그 주의 일요일)
    startCol = Weekday(first, vbSunday)              ' 1=일 .. 7=토
    d = first - (startCol - 1)

    With ws.Range(ws.Cells(CAL_ROW_START, CAL_COL_START), _
                  ws.Cells(CAL_ROW_START + CAL_WEEKS - 1, CAL_COL_START + 6))
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Color = vbBlack
    End With

    For w = 0 To CAL_WEEKS - 1
        For c = 0 To 6
            Set cell = ws.Cells(CAL_ROW_START + w, CAL_COL_START + c)
            k = DateStr(d)
            inMonth = (Month(d) = Month(first) And Year(d) = Year(first))

            line2 = JoinCount("업무", GetCnt(dTask, k), "연락", GetCnt(dCont, k))
            line3 = JoinCount("공유", GetCnt(dShare, k), "점심", GetCnt(dLunch, k))

            txt = CStr(Day(d))
            If Len(line2) > 0 Then txt = txt & vbLf & line2
            If Len(line3) > 0 Then txt = txt & vbLf & line3
            cell.Value = txt

            If Not inMonth Then
                cell.Font.Color = CLR_OTHER
            ElseIf c = 0 Then
                cell.Font.Color = CLR_SUN
            ElseIf c = 6 Then
                cell.Font.Color = CLR_SAT
            End If

            If d = Date Then
                cell.Interior.Color = CLR_TODAY
                cell.Font.Bold = True
            ElseIf inMonth And dUrgent.Exists(k) Then
                cell.Interior.Color = CLR_URGENT
            End If

            d = d + 1
        Next c
    Next w

    FastOff
    Exit Sub
ErrHandler:
    ErrSay "modCalendar.RenderCalendar"
End Sub

'--------------------------------------------------------------
' 각 표를 한 번씩만 순회하며 날짜별로 집계한다.
' 셀마다 COUNTIFS 를 쓰면 42번 x 표 개수만큼 재계산이 돌아 느려진다.
'--------------------------------------------------------------
Private Sub CountByDate(ByRef dTask As Object, ByRef dCont As Object, _
                        ByRef dShare As Object, ByRef dLunch As Object, _
                        ByRef dUrgent As Object)
    Dim data As Variant, r As Long
    Dim dc As Long, xc As Long, k As String
    Dim due As Date

    Set dTask = NewDict()
    Set dCont = NewDict()
    Set dShare = NewDict()
    Set dLunch = NewDict()
    Set dUrgent = NewDict()

    ' 업무
    data = TB_Data(TBL_TASK)
    If Not IsEmpty(data) Then
        dc = TB_Col(TBL_TASK, "날짜")
        xc = TB_Col(TBL_TASK, "삭제여부")
        For r = 1 To UBound(data, 1)
            If Not ToBool(data(r, xc)) Then Bump dTask, S(data(r, dc))
        Next r
    End If

    ' 연락
    data = TB_Data(TBL_CONTACT)
    If Not IsEmpty(data) Then
        dc = TB_Col(TBL_CONTACT, "날짜")
        xc = TB_Col(TBL_CONTACT, "삭제여부")
        For r = 1 To UBound(data, 1)
            If Not ToBool(data(r, xc)) Then Bump dCont, S(data(r, dc))
        Next r
    End If

    ' 점심약속
    data = TB_Data(TBL_LUNCH)
    If Not IsEmpty(data) Then
        dc = TB_Col(TBL_LUNCH, "날짜")
        xc = TB_Col(TBL_LUNCH, "삭제여부")
        For r = 1 To UBound(data, 1)
            If Not ToBool(data(r, xc)) Then Bump dLunch, S(data(r, dc))
        Next r
    End If

    ' 업무공유는 회신기한 기준으로 센다
    data = TB_Data(TBL_SHARE)
    If Not IsEmpty(data) Then
        dc = TB_Col(TBL_SHARE, "회신기한")
        xc = TB_Col(TBL_SHARE, "삭제여부")
        For r = 1 To UBound(data, 1)
            If Not ToBool(data(r, xc)) Then
                k = S(data(r, dc))
                Bump dShare, k
                due = ParseDate(k)
                If due > 0 Then
                    If due >= Date And due <= Date + 3 Then dUrgent(k) = 1
                End If
            End If
        Next r
    End If
End Sub

Private Sub Bump(ByRef d As Object, ByVal k As String)
    If Len(k) = 0 Then Exit Sub
    If d.Exists(k) Then
        d(k) = CLng(d(k)) + 1
    Else
        d(k) = 1
    End If
End Sub

Private Function GetCnt(ByRef d As Object, ByVal k As String) As Long
    If d.Exists(k) Then GetCnt = CLng(d(k)) Else GetCnt = 0
End Function

Private Function JoinCount(ByVal l1 As String, ByVal n1 As Long, _
                           ByVal l2 As String, ByVal n2 As Long) As String
    Dim s0 As String
    If n1 > 0 Then s0 = l1 & " " & n1
    If n2 > 0 Then
        If Len(s0) > 0 Then s0 = s0 & " " & ChrW$(183) & " "
        s0 = s0 & l2 & " " & n2
    End If
    JoinCount = s0
End Function

'==============================================================
' 클릭 처리 (달력 시트의 SelectionChange 에서 호출)
'==============================================================
Public Sub CalClick(ByVal r As Long, ByVal c As Long)
    Dim ws As Worksheet, txt As String, p As Long
    Dim dayNum As Long, first As Date, d As Date
    Dim w As Long, col As Long, offset As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_CAL)

    '---- 상단 버튼 ----
    If r = CAL_ROW_NAV Then
        Select Case c
            Case 2: CalShiftMonth -1: Exit Sub
            Case 6: CalShiftMonth 1: Exit Sub
            Case 7: CalGoToday: Exit Sub
        End Select
        Exit Sub
    End If

    '---- 날짜 칸 ----
    If r < CAL_ROW_START Or r > CAL_ROW_START + CAL_WEEKS - 1 Then Exit Sub
    If c < CAL_COL_START Or c > CAL_COL_START + 6 Then Exit Sub

    txt = S(ws.Cells(r, c).Value)
    If Len(txt) = 0 Then Exit Sub
    p = InStr(txt, vbLf)
    If p > 0 Then txt = Left$(txt, p - 1)
    If Not IsNumeric(txt) Then Exit Sub
    dayNum = CLng(txt)

    ' 칸 위치로 날짜를 역산한다 (앞뒤 달의 날짜도 정확히 나온다)
    first = CalMonth()
    w = r - CAL_ROW_START
    col = c - CAL_COL_START
    offset = w * 7 + col
    d = first - (Weekday(first, vbSunday) - 1) + offset

    If Day(d) <> dayNum Then Exit Sub       ' 화면과 어긋나면 무시

    SetViewDate d
    RenderToday
    ThisWorkbook.Worksheets(SH_TODAY).Activate
    Exit Sub

ErrHandler:
    ErrSay "modCalendar.CalClick"
End Sub

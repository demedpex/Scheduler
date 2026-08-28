Attribute VB_Name = "modUtil"
Option Explicit
'==============================================================
' modUtil - 공통 유틸리티
'   ID 생성 / 이스케이프 / 체크섬 / 클립보드 / 로그 / 화면제어
'==============================================================

Public Const APP_NAME As String = "업무 스케줄러"
Public Const PROTO_TAG As String = "WSCH1"
Public Const PART_MAX_CHARS As Long = 1500      ' 파트 본문 최대 글자수

' 클립보드용 MSForms.DataObject CLSID (참조 설정 없이 후기 바인딩)
Private Const CLSID_DATAOBJECT As String = "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"

Private mSeeded As Boolean
Private mScrPrev As Boolean
Private mEvtPrev As Boolean
Private mDepth As Long

'--------------------------------------------------------------
' 날짜/시각 문자열
'   저장은 전부 문자열로 한다. 엑셀 날짜 서식·지역설정에 휘둘리지 않기 위함.
'--------------------------------------------------------------
Public Function NowStamp() As String
    NowStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function DateStr(ByVal d As Date) As String
    DateStr = Format$(d, "yyyy-mm-dd")
End Function

Public Function TodayStr() As String
    TodayStr = Format$(Date, "yyyy-mm-dd")
End Function

' "yyyy-mm-dd" -> Date. 형식이 아니면 0 을 돌려준다.
Public Function ParseDate(ByVal s0 As String) As Date
    Dim t As String
    t = Trim$(s0)
    ParseDate = 0
    If Len(t) < 10 Then Exit Function
    If Not (IsNumeric(Left$(t, 4)) And IsNumeric(Mid$(t, 6, 2)) And IsNumeric(Mid$(t, 9, 2))) Then Exit Function
    On Error Resume Next
    ParseDate = DateSerial(CLng(Left$(t, 4)), CLng(Mid$(t, 6, 2)), CLng(Mid$(t, 9, 2)))
    If Err.Number <> 0 Then
        ParseDate = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

' 시간 문자열 정규화. "9:5" -> "09:05". 비어 있거나 형식이 아니면 빈 문자열.
Public Function NormTime(ByVal s0 As String) As String
    Dim t As String, p As Long, h As Long, m As Long
    t = Trim$(s0)
    NormTime = ""
    If Len(t) = 0 Then Exit Function
    p = InStr(t, ":")
    If p = 0 Then
        If IsNumeric(t) And Len(t) <= 2 Then
            h = CLng(t)
            m = 0
        ElseIf IsNumeric(t) And Len(t) = 4 Then
            h = CLng(Left$(t, 2))
            m = CLng(Right$(t, 2))
        Else
            Exit Function
        End If
    Else
        If Not IsNumeric(Left$(t, p - 1)) Or Not IsNumeric(Mid$(t, p + 1)) Then Exit Function
        h = CLng(Left$(t, p - 1))
        m = CLng(Mid$(t, p + 1))
    End If
    If h < 0 Or h > 23 Or m < 0 Or m > 59 Then Exit Function
    NormTime = Format$(h, "00") & ":" & Format$(m, "00")
End Function

'--------------------------------------------------------------
' ID 생성
'   형식: 작성자-yyyymmdd-hhnnss-4자리난수
'   자동증가 숫자는 절대 쓰지 않는다. PC마다 따로 증가해 반드시 충돌한다.
'--------------------------------------------------------------
Public Function NewID(ByVal author As String) As String
    If Not mSeeded Then
        Randomize
        mSeeded = True
    End If
    NewID = SanitizeKey(author) & "-" & Format$(Now, "yyyymmdd-hhnnss") & _
            "-" & Format$(Int(Rnd() * 10000), "0000")
End Function

' ID 안에 들어가면 곤란한 문자를 제거
Public Function SanitizeKey(ByVal s0 As String) As String
    Dim t As String
    t = Trim$(s0)
    If Len(t) = 0 Then t = "사용자"
    t = Replace$(t, "|", "")
    t = Replace$(t, vbTab, "")
    t = Replace$(t, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Replace$(t, "\", "")
    SanitizeKey = t
End Function

'--------------------------------------------------------------
' 이스케이프 / 언이스케이프
'   인코딩 순서(반드시 이 순서):
'     1) \  ->  \\
'     2) |  ->  \p
'     3) CR 제거, LF -> \n
'     4) Tab -> \t
'--------------------------------------------------------------
Public Function Esc(ByVal s0 As String) As String
    Dim t As String
    t = s0
    t = Replace$(t, "\", "\\")
    t = Replace$(t, "|", "\p")
    t = Replace$(t, vbCr, "")
    t = Replace$(t, vbLf, "\n")
    t = Replace$(t, vbTab, "\t")
    Esc = t
End Function

' 디코딩을 단순 Replace 역순으로 하면 "\\p" 같은 입력에서 깨진다.
' 앞에서부터 한 문자씩 훑으며 \ 를 만나면 다음 한 글자를 보고 판단한다.
Public Function Unesc(ByVal s0 As String) As String
    Dim i As Long, n As Long, ch As String, nx As String
    Dim buf As String, pos As Long
    n = Len(s0)
    If n = 0 Then
        Unesc = ""
        Exit Function
    End If
    buf = Space$(n)
    pos = 0
    i = 1
    Do While i <= n
        ch = Mid$(s0, i, 1)
        If ch = "\" And i < n Then
            nx = Mid$(s0, i + 1, 1)
            pos = pos + 1
            Select Case nx
                Case "\": Mid$(buf, pos, 1) = "\"
                Case "p": Mid$(buf, pos, 1) = "|"
                Case "n": Mid$(buf, pos, 1) = vbLf
                Case "t": Mid$(buf, pos, 1) = vbTab
                Case Else: Mid$(buf, pos, 1) = nx   ' 알 수 없는 이스케이프는 글자 그대로
            End Select
            i = i + 2
        Else
            pos = pos + 1
            Mid$(buf, pos, 1) = ch
            i = i + 1
        End If
    Loop
    Unesc = Left$(buf, pos)
End Function

'--------------------------------------------------------------
' 체크섬
'   sum = (sum * 31 + 문자코드) Mod 65536 -> 4자리 대문자 16진수
'   AscW 는 32767 을 넘는 코드를 음수로 돌려주므로 부호를 없애고 쓴다.
'   (보내는 쪽/받는 쪽이 같은 코드를 쓰므로 결과는 항상 일치한다)
'--------------------------------------------------------------
Public Function Checksum4(ByVal s0 As String) As String
    Dim i As Long, sum As Long, c As Long
    For i = 1 To Len(s0)
        c = AscW(Mid$(s0, i, 1))
        If c < 0 Then c = c + 65536
        sum = (sum * 31 + c) Mod 65536
    Next i
    Checksum4 = Right$("000" & UCase$(Hex$(sum)), 4)
End Function

'--------------------------------------------------------------
' 클립보드 (참조 설정 없이 동작)
'--------------------------------------------------------------
Public Function ClipSet(ByVal s0 As String) As Boolean
    Dim dob As Object
    On Error GoTo Fail
    Set dob = GetObject(CLSID_DATAOBJECT)
    dob.SetText s0
    dob.PutInClipboard
    ClipSet = True
    Exit Function
Fail:
    ClipSet = False
End Function

Public Function ClipGet() As String
    Dim dob As Object
    On Error GoTo Fail
    Set dob = GetObject(CLSID_DATAOBJECT)
    dob.GetFromClipboard
    ClipGet = dob.GetText
    Exit Function
Fail:
    ClipGet = ""
End Function

'--------------------------------------------------------------
' 화면 갱신 제어 (중첩 호출 대비 카운터 방식)
'--------------------------------------------------------------
Public Sub FastOn()
    If mDepth = 0 Then
        mScrPrev = Application.ScreenUpdating
        mEvtPrev = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    mDepth = mDepth + 1
End Sub

Public Sub FastOff()
    mDepth = mDepth - 1
    If mDepth <= 0 Then
        mDepth = 0
        Application.ScreenUpdating = mScrPrev
        Application.EnableEvents = mEvtPrev
    End If
End Sub

' 에러 핸들러에서 부르는 강제 복구
Public Sub FastReset()
    mDepth = 0
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub

'--------------------------------------------------------------
' 메시지
'--------------------------------------------------------------
Public Sub Say(ByVal msg As String)
    MsgBox msg, vbInformation, APP_NAME
End Sub

Public Sub Warn(ByVal msg As String)
    MsgBox msg, vbExclamation, APP_NAME
End Sub

Public Function Ask(ByVal msg As String) As Boolean
    Ask = (MsgBox(msg, vbYesNo + vbQuestion, APP_NAME) = vbYes)
End Function

Public Sub ErrSay(ByVal procName As String)
    Dim n As Long, d As String
    n = Err.Number
    d = Err.Description
    FastReset
    MsgBox "처리 중 오류가 발생했습니다." & vbLf & vbLf & _
           "위치: " & procName & vbLf & _
           "내용: " & d & " (" & n & ")", _
           vbCritical, APP_NAME
End Sub

'--------------------------------------------------------------
' 값 변환 도우미
'--------------------------------------------------------------
Public Function S(ByVal v As Variant) As String
    If IsObject(v) Then
        S = ""
    ElseIf IsError(v) Then
        S = ""
    ElseIf IsNull(v) Then
        S = ""
    ElseIf IsEmpty(v) Then
        S = ""
    Else
        S = Trim$(CStr(v))
    End If
End Function

Public Function ToBool(ByVal v As Variant) As Boolean
    Dim t As String
    t = S(v)
    ToBool = (t = "1" Or UCase$(t) = "TRUE" Or t = "-1")
End Function

Public Function BoolStr(ByVal b As Boolean) As String
    If b Then BoolStr = "1" Else BoolStr = "0"
End Function

Public Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                       ' TextCompare
    Set NewDict = d
End Function

Public Function Ellipsis(ByVal s0 As String, ByVal n As Long) As String
    Dim t As String
    t = Replace$(Replace$(s0, vbCr, " "), vbLf, " ")
    If Len(t) <= n Then
        Ellipsis = t
    Else
        Ellipsis = Left$(t, n) & "..."
    End If
End Function

' 폴더가 없으면 만든다 (중간 경로 포함)
Public Sub EnsureFolder(ByVal path As String)
    Dim fso As Object, parent As String
    If Len(Trim$(path)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    ' 루트까지 올라가면 멈춘다 (안 그러면 무한 재귀)
    If Len(parent) > 0 And StrComp(parent, path, vbTextCompare) <> 0 Then
        EnsureFolder parent
    End If
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

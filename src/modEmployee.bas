Attribute VB_Name = "modEmployee"
Option Explicit
'==============================================================
' modEmployee - 직원 명부 로드 / 검색
'
' 외부 링크 수식은 쓰지 않는다. 경로가 바뀌면 전부 깨지기 때문이다.
' [명부 갱신] 버튼으로 파일을 읽어 mst_Employee 시트에 통째로 복사한다.
'
' 내선번호는 반드시 텍스트로 다룬다. 앞자리 0(0212)이 사라지면 안 된다.
'==============================================================

Private Const COL_NAME As Long = 1       ' 이름
Private Const COL_EXT As Long = 2        ' 내선번호
Private Const COL_DEPT As Long = 3       ' 부서
Private Const COL_RANK As Long = 4       ' 직급

Private mEmp As Variant                  ' (1..n, 1..4) 캐시
Private mEmpLoaded As Boolean
Private mExact As Object                 ' "이름<TAB>부서" -> 내선번호

'--------------------------------------------------------------
' 캐시
'--------------------------------------------------------------
Public Sub EmpInvalidate()
    mEmp = Empty
    mEmpLoaded = False
    Set mExact = Nothing
End Sub

' 명부 전체를 2차원 배열로. 비어 있으면 Empty.
Public Function EmpArray() As Variant
    If Not mEmpLoaded Then
        mEmp = TB_Data(TBL_EMP)
        mEmpLoaded = True
        Set mExact = Nothing
    End If
    EmpArray = mEmp
End Function

' 이름이 들어 있는 행만 센다 (표를 만들 때 생기는 빈 행은 제외)
Public Function EmpCount() As Long
    Dim v As Variant, r As Long, n As Long
    v = EmpArray()
    If IsEmpty(v) Then Exit Function
    For r = 1 To UBound(v, 1)
        If Len(S(v(r, COL_NAME))) > 0 Then n = n + 1
    Next r
    EmpCount = n
End Function

'--------------------------------------------------------------
' 이름 부분일치 검색
'   반환: (1..n, 1..4) = 이름, 부서, 직급, 내선번호. 없으면 Empty.
'   셀을 한 칸씩 읽으면 수백 배 느리다. 반드시 배열로 통째 읽는다.
'--------------------------------------------------------------
Public Function EmpSearch(ByVal namePart As String) As Variant
    Dim v As Variant, r As Long, n As Long
    Dim hit() As Long, cnt As Long
    Dim res As Variant, i As Long
    Dim q As String

    q = Trim$(namePart)
    EmpSearch = Empty
    If Len(q) = 0 Then Exit Function

    v = EmpArray()
    If IsEmpty(v) Then Exit Function
    n = UBound(v, 1)
    ReDim hit(1 To n)

    ' 1차: 완전일치 우선
    For r = 1 To n
        If StrComp(S(v(r, COL_NAME)), q, vbTextCompare) = 0 Then
            cnt = cnt + 1
            hit(cnt) = r
        End If
    Next r

    ' 완전일치가 없을 때만 부분일치로 넓힌다
    If cnt = 0 Then
        For r = 1 To n
            If Len(S(v(r, COL_NAME))) > 0 Then
                If InStr(1, S(v(r, COL_NAME)), q, vbTextCompare) > 0 Then
                    cnt = cnt + 1
                    hit(cnt) = r
                End If
            End If
        Next r
    End If

    If cnt = 0 Then Exit Function

    ReDim res(1 To cnt, 1 To 4)
    For i = 1 To cnt
        r = hit(i)
        res(i, 1) = S(v(r, COL_NAME))
        res(i, 2) = S(v(r, COL_DEPT))
        res(i, 3) = S(v(r, COL_RANK))
        res(i, 4) = S(v(r, COL_EXT))
    Next i
    EmpSearch = res
End Function

'--------------------------------------------------------------
' 이름 + 부서로 내선번호 조회 (오늘 시트 렌더링에서 매번 호출됨)
'   명부에 사번이 없으므로 이 둘이 조회 키다.
'--------------------------------------------------------------
Public Function EmpExtension(ByVal empName As String, ByVal dept As String) As String
    Dim v As Variant, r As Long, k As String
    If mExact Is Nothing Then
        Set mExact = NewDict()
        v = EmpArray()
        If Not IsEmpty(v) Then
            For r = 1 To UBound(v, 1)
                k = S(v(r, COL_NAME)) & vbTab & S(v(r, COL_DEPT))
                If Not mExact.Exists(k) Then mExact(k) = S(v(r, COL_EXT))
            Next r
        End If
    End If
    k = Trim$(empName) & vbTab & Trim$(dept)
    If mExact.Exists(k) Then
        EmpExtension = CStr(mExact(k))
    Else
        ' 부서가 비어 있는 명부도 있으므로 이름만으로 한 번 더 시도
        EmpExtension = ""
        v = EmpArray()
        If Not IsEmpty(v) Then
            For r = 1 To UBound(v, 1)
                If StrComp(S(v(r, COL_NAME)), Trim$(empName), vbTextCompare) = 0 Then
                    EmpExtension = S(v(r, COL_EXT))
                    Exit Function
                End If
            Next r
        End If
    End If
End Function

'==============================================================
' 명부 갱신
'==============================================================
Public Sub EmpRefreshFromFile(ByVal path As String, ByRef outCount As Long, ByRef outMsg As String)
    Dim rows As Collection
    Dim ext As String

    outCount = 0
    outMsg = ""

    If Len(Dir$(path)) = 0 Then
        outMsg = "파일을 찾을 수 없습니다:" & vbLf & path
        Exit Sub
    End If

    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
        Case "csv", "txt"
            Set rows = ReadDelimitedFile(path)
        Case "xlsx", "xlsm", "xls", "xlsb"
            Set rows = ReadExcelFile(path)
        Case Else
            outMsg = "지원하지 않는 파일 형식입니다: ." & ext & vbLf & _
                     "CSV 또는 XLSX 파일을 선택해 주세요."
            Exit Sub
    End Select

    If rows Is Nothing Then
        outMsg = "명부 파일을 읽지 못했습니다."
        Exit Sub
    End If
    If rows.Count = 0 Then
        outMsg = "명부 파일에서 읽을 행이 없습니다."
        Exit Sub
    End If

    WriteEmployeeTable rows, outCount
    EmpInvalidate
    CfgSet "명부파일경로", path
    CfgSet "명부갱신시각", NowStamp()
    outMsg = "명부 " & outCount & "명을 불러왔습니다."
End Sub

'--------------------------------------------------------------
' 읽어온 행들을 mst_Employee 에 덮어쓴다
'   rows: 각 항목은 (1..4) 문자열 배열 = 이름, 내선번호, 부서, 직급
'--------------------------------------------------------------
Private Sub WriteEmployeeTable(ByRef rows As Collection, ByRef outCount As Long)
    Dim lo As ListObject, ws As Worksheet
    Dim arr As Variant, i As Long, c As Long, rec As Variant

    Set lo = TB(TBL_EMP)
    Set ws = lo.Parent

    ' 기존 내용 비우기
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    ReDim arr(1 To rows.Count, 1 To 4)
    For i = 1 To rows.Count
        rec = rows(i)
        For c = 1 To 4
            arr(i, c) = rec(c)
        Next c
    Next i

    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), _
                       lo.HeaderRowRange.Cells(1, 4).Offset(rows.Count, 0))
    lo.DataBodyRange.NumberFormat = "@"
    lo.DataBodyRange.Value = arr
    outCount = rows.Count
End Sub

'--------------------------------------------------------------
' CSV / 탭구분 텍스트 읽기
'   BOM 이 있으면 UTF-8, 없으면 시스템 기본(한국어 Windows 는 CP949)으로 읽는다.
'--------------------------------------------------------------
Private Function ReadDelimitedFile(ByVal path As String) As Collection
    Dim txt As String, lines() As String, i As Long
    Dim f() As String, rec As Variant, res As Collection
    Dim map As Variant, headerDone As Boolean
    Dim sep As String

    txt = ReadTextFile(path)
    If Len(txt) = 0 Then
        Set ReadDelimitedFile = Nothing
        Exit Function
    End If

    txt = Replace$(txt, vbCrLf, vbLf)
    txt = Replace$(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    ' 구분자 추정: 첫 줄에 탭이 쉼표보다 많으면 탭
    sep = ","
    If UBound(lines) >= 0 Then
        If CountChar(lines(0), vbTab) > CountChar(lines(0), ",") Then sep = vbTab
    End If

    Set res = New Collection
    map = Array(0, 1, 2, 3)                 ' 머리글이 없을 때의 기본 순서

    For i = LBound(lines) To UBound(lines)
        f = SplitCsvLine(lines(i), sep)
        If Not headerDone Then
            headerDone = True
            If LooksLikeHeader(f) Then
                map = BuildHeaderMap(f)
                GoTo ContinueLoop
            End If
        End If
        If IsBlankRow(f) Then Exit For          ' 중간 빈 행을 만나면 거기서 멈춘다
        rec = PickFields(f, map)
        If Len(rec(1)) > 0 Then res.Add rec     ' 이름 없는 행은 버린다
ContinueLoop:
    Next i

    Set ReadDelimitedFile = res
End Function

' BOM 을 보고 인코딩을 정해 파일 전체를 읽는다
Private Function ReadTextFile(ByVal path As String) As String
    Dim st As Object, b() As Byte, isUtf8 As Boolean
    Dim ff As Integer, sz As Long

    ' 앞 3바이트로 BOM 판정
    On Error GoTo Fallback
    ff = FreeFile
    Open path For Binary Access Read As #ff
    sz = LOF(ff)
    If sz >= 3 Then
        ReDim b(0 To 2)
        Get #ff, 1, b
        isUtf8 = (b(0) = &HEF And b(1) = &HBB And b(2) = &HBF)
    End If
    Close #ff

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                  ' adTypeText
    If isUtf8 Then
        st.Charset = "utf-8"
    Else
        st.Charset = "ks_c_5601-1987"            ' 한국어 Windows 기본 코드페이지
    End If
    st.Open
    st.LoadFromFile path
    ReadTextFile = st.ReadText(-1)
    st.Close
    ' UTF-8 BOM 이 문자로 남는 경우 제거
    If Len(ReadTextFile) > 0 Then
        If AscW(Left$(ReadTextFile, 1)) = &HFEFF Then ReadTextFile = Mid$(ReadTextFile, 2)
    End If
    Exit Function

Fallback:
    ' ADODB 를 못 쓰는 환경이면 기본 인코딩으로 통째 읽기
    On Error Resume Next
    Close #ff
    Err.Clear
    ff = FreeFile
    Open path For Input As #ff
    ReadTextFile = Input$(LOF(ff), ff)
    Close #ff
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' XLSX 읽기
'   내선번호가 숫자로 저장돼 있으면 .Text 로 표시값을 가져와 0212 를 지킨다.
'--------------------------------------------------------------
Private Function ReadExcelFile(ByVal path As String) As Collection
    Dim wb As Workbook, ws As Worksheet
    Dim lastR As Long, lastC As Long, r As Long, c As Long
    Dim f() As String, rec As Variant, res As Collection
    Dim map As Variant, headerDone As Boolean

    On Error GoTo CleanFail
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    ws.Columns.AutoFit

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastC < 1 Then lastC = 1
    If lastC > 20 Then lastC = 20

    Set res = New Collection
    map = Array(0, 1, 2, 3)

    For r = 1 To lastR
        ReDim f(0 To lastC - 1)
        For c = 1 To lastC
            f(c - 1) = Trim$(ws.Cells(r, c).Text)
        Next c
        If Not headerDone Then
            headerDone = True
            If LooksLikeHeader(f) Then
                map = BuildHeaderMap(f)
                GoTo ContinueLoop
            End If
        End If
        If IsBlankRow(f) Then Exit For
        rec = PickFields(f, map)
        If Len(rec(1)) > 0 Then res.Add rec
ContinueLoop:
    Next r

    wb.Close SaveChanges:=False
    Set ReadExcelFile = res
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Set ReadExcelFile = Nothing
End Function

'==============================================================
' 파싱 도우미
'==============================================================

' 따옴표를 고려한 한 줄 분리
Private Function SplitCsvLine(ByVal ln As String, ByVal sep As String) As String()
    Dim res() As String, cnt As Long
    Dim i As Long, n As Long, ch As String
    Dim cur As String, inQ As Boolean

    n = Len(ln)
    ReDim res(0 To 0)
    cnt = 0
    i = 1
    Do While i <= n
        ch = Mid$(ln, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(ln, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = sep Then
                ReDim Preserve res(0 To cnt)
                res(cnt) = Trim$(cur)
                cnt = cnt + 1
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
        i = i + 1
    Loop
    ReDim Preserve res(0 To cnt)
    res(cnt) = Trim$(cur)
    SplitCsvLine = res
End Function

Private Function CountChar(ByVal s0 As String, ByVal ch As String) As Long
    CountChar = Len(s0) - Len(Replace$(s0, ch, ""))
End Function

Private Function IsBlankRow(ByRef f() As String) As Boolean
    Dim i As Long
    For i = LBound(f) To UBound(f)
        If Len(Trim$(f(i))) > 0 Then
            IsBlankRow = False
            Exit Function
        End If
    Next i
    IsBlankRow = True
End Function

Private Function LooksLikeHeader(ByRef f() As String) As Boolean
    Dim i As Long, t As String
    For i = LBound(f) To UBound(f)
        t = Trim$(f(i))
        If t = "이름" Or t = "성명" Or t = "내선번호" Or t = "내선" Or _
           t = "부서" Or t = "부서명" Or t = "직급" Or t = "직위" Then
            LooksLikeHeader = True
            Exit Function
        End If
    Next i
End Function

' 머리글 이름을 보고 컬럼 위치를 정한다.
' 반환: map(0..3) = 원본 컬럼 인덱스(0부터). 없으면 -1.
Private Function BuildHeaderMap(ByRef f() As String) As Variant
    Dim map(0 To 3) As Long
    Dim i As Long, t As String
    For i = 0 To 3
        map(i) = -1
    Next i
    For i = LBound(f) To UBound(f)
        t = Trim$(f(i))
        Select Case t
            Case "이름", "성명":            If map(0) < 0 Then map(0) = i
            Case "내선번호", "내선", "전화번호", "연락처":  If map(1) < 0 Then map(1) = i
            Case "부서", "부서명", "소속":   If map(2) < 0 Then map(2) = i
            Case "직급", "직위":            If map(3) < 0 Then map(3) = i
        End Select
    Next i
    ' 못 찾은 항목은 순서대로 채운다
    For i = 0 To 3
        If map(i) < 0 Then map(i) = i
    Next i
    BuildHeaderMap = map
End Function

' map 에 따라 (1..4) = 이름, 내선번호, 부서, 직급 을 뽑는다
Private Function PickFields(ByRef f() As String, ByRef map As Variant) As Variant
    Dim rec(1 To 4) As String
    Dim i As Long, src As Long
    For i = 0 To 3
        src = CLng(map(i))
        If src >= LBound(f) And src <= UBound(f) Then
            rec(i + 1) = Trim$(f(src))
        Else
            rec(i + 1) = ""
        End If
    Next i
    PickFields = rec
End Function

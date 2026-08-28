Attribute VB_Name = "modTest"
Option Explicit
'==============================================================
' modTest - 검증 프로시저
'
' VBE 에서 RunAllTests 에 커서를 두고 F5 를 누르면 전부 실행된다.
' 결과는 MsgBox 와 직접 실행 창(Ctrl+G) 양쪽에 나온다.
'
' 주의: 4~10번 테스트는 tbl_Share / mst_Employee 를 잠시 비우고 쓴다.
'       시작할 때 통째로 저장해 두고 끝나면(오류가 나도) 반드시 되돌린다.
'==============================================================

Private mPass As Long
Private mFail As Long
Private mLog As String

Public Sub RunAllTests()
    Dim shareSnap As Variant, empSnap As Variant, cfgPath As String
    Dim restored As Boolean

    mPass = 0
    mFail = 0
    mLog = ""

    shareSnap = TB_Data(TBL_SHARE)
    empSnap = TB_Data(TBL_EMP)
    cfgPath = CfgGet("명부파일경로", "")

    On Error GoTo Restore
    FastOn

    T01_EscapeRoundTrip
    T02_NastyEscape
    T03_Checksum
    T04_ReverseOrderParts
    T05_Idempotent
    T06_OlderDoesNotWin
    T07_TombstoneStays
    T08_Homonym
    T09_LeadingZero
    T10_EmptyDeptRank

Restore:
    If Err.Number <> 0 Then
        Fail "실행중단", "예상치 못한 오류: " & Err.Description & " (" & Err.Number & ")"
        Err.Clear
    End If
    On Error Resume Next
    PutAll TBL_SHARE, shareSnap
    PutAll TBL_EMP, empSnap
    CfgSet "명부파일경로", cfgPath
    SyncResetInbox
    EmpInvalidate
    restored = True
    FastOff
    On Error GoTo 0

    Debug.Print mLog
    MsgBox "테스트 결과: 성공 " & mPass & " / 실패 " & mFail & vbLf & vbLf & mLog, _
           IIf(mFail = 0, vbInformation, vbExclamation), APP_NAME & " 검증"
End Sub

'==============================================================
' 1. 이스케이프 왕복
'==============================================================
Private Sub T01_EscapeRoundTrip()
    Dim src As String, enc As String, dec As String
    src = "내용|파이프\백슬래시" & vbLf & "둘째줄" & vbTab & "탭끝"
    enc = Esc(src)
    dec = Unesc(enc)
    Check "1. 이스케이프 왕복", dec = src, "원본과 다름: [" & dec & "]"
    ' 인코딩 결과에는 구분자나 줄바꿈이 남아 있으면 안 된다
    Check "1b. 인코딩 결과에 구분자 없음", _
          InStr(enc, "|") = 0 And InStr(enc, vbLf) = 0 And InStr(enc, vbTab) = 0, _
          "인코딩 결과에 |, 줄바꿈, 탭이 남아 있음"
End Sub

'==============================================================
' 2. 이스케이프 문자가 연속된 악성 입력
'==============================================================
Private Sub T02_NastyEscape()
    Dim cases As Variant, i As Long, src As String, ok As Boolean
    cases = Array("\\p", "\p", "\\", "\", "a\\\\pb", "\n", "\\n", "|\\|", _
                  "\\\p", "끝에역슬래시\", "\t\\t")
    ok = True
    For i = LBound(cases) To UBound(cases)
        src = CStr(cases(i))
        If Unesc(Esc(src)) <> src Then
            ok = False
            mLog = mLog & "   실패한 입력: [" & src & "] -> [" & Unesc(Esc(src)) & "]" & vbLf
        End If
    Next i
    Check "2. 연속 이스케이프 입력 파싱", ok, "위 목록 참조"
End Sub

'==============================================================
' 3. 체크섬이 한 글자 변형을 잡아내는가
'==============================================================
Private Sub T03_Checksum()
    Dim a As String, b As String, c As String
    a = "S|홍길동-20260827-143052-8471|기획-2026-0451|상반기 실적 보고|본문|2026-08-27|2026-08-30|홍길동|2026-08-27 14:30:52|0"
    b = Replace$(a, "0451", "0452")                       ' 한 글자 변형
    c = Left$(a, Len(a) - 1)                              ' 뒤가 잘림

    Check "3a. 같은 문자열은 같은 체크섬", Checksum4(a) = Checksum4(a), "재현성 없음"
    Check "3b. 한 글자 바뀌면 체크섬이 달라짐", Checksum4(a) <> Checksum4(b), _
          "둘 다 " & Checksum4(a)
    Check "3c. 잘린 문자열도 체크섬이 달라짐", Checksum4(a) <> Checksum4(c), _
          "둘 다 " & Checksum4(a)
    Check "3d. 체크섬은 4자리 16진수", Len(Checksum4(a)) = 4, "길이 " & Len(Checksum4(a))
End Sub

'==============================================================
' 4. 3파트로 나뉜 데이터를 역순으로 붙여넣어도 병합되는가
'==============================================================
Private Sub T04_ReverseOrderParts()
    Dim parts As Variant, cnt As Long, i As Long
    Dim st As Long, msg As String

    ClearAll TBL_SHARE
    InsertShare "TEST-A", "T-001", "가", String$(1600, "가"), "2026-01-01 00:00:00", "0"
    InsertShare "TEST-B", "T-002", "나", String$(1600, "나"), "2026-01-01 00:00:00", "0"
    InsertShare "TEST-C", "T-003", "다", String$(1600, "다"), "2026-01-01 00:00:00", "0"

    parts = SyncBuildParts(SCOPE_ALL, 0, cnt)
    If IsEmpty(parts) Then
        Fail "4. 역순 병합", "파트 생성 실패"
        Exit Sub
    End If
    Check "4a. 3건이 3개 파트로 나뉨", UBound(parts) = 3, "파트 수 " & UBound(parts)

    ' 받는 쪽 상황을 만든다
    ClearAll TBL_SHARE
    SyncResetInbox

    For i = UBound(parts) To 1 Step -1                    ' 역순으로 먹인다
        st = SyncFeed(CStr(parts(i)), msg)
    Next i
    Check "4b. 역순으로 다 받으면 완료 상태", st = FD_COMPLETE, "상태 " & st & " / " & msg

    SyncApply
    Check "4c. 3건이 모두 병합됨", TB_RowCount(TBL_SHARE) = 3, "행 수 " & TB_RowCount(TBL_SHARE)
    Check "4d. 긴 내용이 그대로 복원됨", _
          Len(GetShareField("TEST-B", "내용")) = 1600, _
          "길이 " & Len(GetShareField("TEST-B", "내용"))
End Sub

'==============================================================
' 5. 같은 파트를 두 번 붙여넣어도 중복되지 않는가 (멱등)
'==============================================================
Private Sub T05_Idempotent()
    Dim parts As Variant, cnt As Long, msg As String
    Dim st1 As Long, st2 As Long

    ClearAll TBL_SHARE
    InsertShare "TEST-D", "T-010", "멱등확인", "내용", "2026-01-01 00:00:00", "0"
    parts = SyncBuildParts(SCOPE_ALL, 0, cnt)
    Check "5a. 1건은 파트 1개", UBound(parts) = 1, "파트 수 " & UBound(parts)

    ClearAll TBL_SHARE
    SyncResetInbox
    st1 = SyncFeed(CStr(parts(1)), msg)
    st2 = SyncFeed(CStr(parts(1)), msg)                   ' 같은 조각을 한 번 더

    Check "5b. 1개짜리는 바로 완료", st1 = FD_COMPLETE, "상태 " & st1
    Check "5c. 다시 넣어도 완료 상태 유지", st2 = FD_COMPLETE, "상태 " & st2 & " / " & msg

    SyncApply
    Check "5d. 중복 생성되지 않음", TB_RowCount(TBL_SHARE) = 1, "행 수 " & TB_RowCount(TBL_SHARE)

    ' 이미 반영한 텍스트를 또 붙여넣는 경우
    SyncResetInbox
    SyncFeed CStr(parts(1)), msg
    SyncApply
    Check "5e. 반영 후 또 넣어도 1건", TB_RowCount(TBL_SHARE) = 1, "행 수 " & TB_RowCount(TBL_SHARE)
End Sub

'==============================================================
' 6. 수정시각이 더 과거인 레코드가 최신을 덮어쓰지 않는가
'==============================================================
Private Sub T06_OlderDoesNotWin()
    Dim parts As Variant, cnt As Long, msg As String, r As Long, rec As Variant

    ClearAll TBL_SHARE
    InsertShare "TEST-E", "T-020", "옛제목", "옛내용", "2026-01-01 00:00:00", "0"
    parts = SyncBuildParts(SCOPE_ALL, 0, cnt)             ' 과거 시각의 텍스트를 확보

    ' 내 쪽에서 더 최신으로 고쳐 둔다
    r = TB_FindRowByID(TBL_SHARE, "TEST-E")
    rec = TB_GetRow(TBL_SHARE, r)
    RecSet rec, TBL_SHARE, "제목", "새제목"
    RecSet rec, TBL_SHARE, "수정시각", "2026-01-02 00:00:00"
    TB_UpdateRow TBL_SHARE, r, rec

    SyncResetInbox
    SyncFeed CStr(parts(1)), msg
    SyncApply

    Check "6. 과거 레코드가 최신을 덮어쓰지 않음", _
          GetShareField("TEST-E", "제목") = "새제목", _
          "제목이 [" & GetShareField("TEST-E", "제목") & "] 로 바뀜"
End Sub

'==============================================================
' 7. 삭제 플래그가 병합 후 되살아나지 않는가
'==============================================================
Private Sub T07_TombstoneStays()
    Dim parts As Variant, cnt As Long, msg As String, r As Long, rec As Variant

    '---- 7-1. 내가 지운 건이 남의 텍스트로 되살아나면 안 된다 ----
    ClearAll TBL_SHARE
    InsertShare "TEST-F", "T-030", "살아있음", "내용", "2026-01-01 00:00:00", "0"
    parts = SyncBuildParts(SCOPE_ALL, 0, cnt)             ' 살아있는 상태의 텍스트

    TB_SoftDeleteByID TBL_SHARE, "TEST-F"                 ' 내 쪽에서 삭제
    SyncResetInbox
    SyncFeed CStr(parts(1)), msg
    SyncApply
    Check "7a. 내가 지운 건은 되살아나지 않음", _
          GetShareField("TEST-F", "삭제여부") = "1", _
          "삭제여부가 " & GetShareField("TEST-F", "삭제여부")

    '---- 7-2. 남이 지운 표식은 내 쪽에도 반영돼야 한다 ----
    ClearAll TBL_SHARE
    InsertShare "TEST-G", "T-031", "곧삭제", "내용", "2026-01-01 00:00:00", "1"
    parts = SyncBuildParts(SCOPE_ALL, 0, cnt)             ' 삭제 표식이 담긴 텍스트

    ClearAll TBL_SHARE
    InsertShare "TEST-G", "T-031", "곧삭제", "내용", "2026-01-01 00:00:00", "0"   ' 내 쪽은 살아있음
    SyncResetInbox
    SyncFeed CStr(parts(1)), msg
    SyncApply
    Check "7b. 남이 지운 표식이 전달됨", _
          GetShareField("TEST-G", "삭제여부") = "1", _
          "삭제여부가 " & GetShareField("TEST-G", "삭제여부")
End Sub

'==============================================================
' 8. 동명이인 3명 / 1명
'==============================================================
Private Sub T08_Homonym()
    Dim m As Variant

    SetEmployees Array( _
        Array("김지은", "1101", "기획팀", "대리"), _
        Array("김지은", "1102", "총무팀", "과장"), _
        Array("김지은", "1103", "영업팀", "사원"), _
        Array("박유일", "1200", "인사팀", "차장"))

    m = EmpSearch("김지은")
    Check "8a. 동명이인 3건을 찾음", _
          Not IsEmpty(m), "검색 결과 없음"
    If Not IsEmpty(m) Then
        Check "8b. 결과가 정확히 3건", UBound(m, 1) = 3, "건수 " & UBound(m, 1)
        Check "8c. 부서가 함께 나옴", _
              Len(S(m(1, 2))) > 0 And Len(S(m(1, 3))) > 0 And Len(S(m(1, 4))) > 0, _
              "부서/직급/내선 중 빈 값이 있음"
    End If

    m = EmpSearch("박유일")
    If IsEmpty(m) Then
        Fail "8d. 1명이면 자동 선택", "검색 결과 없음"
    Else
        Check "8d. 1건이면 팝업 없이 자동 선택", _
              UBound(m, 1) = 1 And PickEmployee(m) = 1, _
              "건수 " & UBound(m, 1)
    End If

    ' 부분일치도 되는지
    m = EmpSearch("지은")
    Check "8e. 이름 일부만 입력해도 찾음", _
          Not IsEmpty(m), "부분일치 실패"
    If Not IsEmpty(m) Then
        Check "8f. 부분일치 결과 3건", UBound(m, 1) = 3, "건수 " & UBound(m, 1)
    End If
End Sub

'==============================================================
' 9. 내선번호 앞자리 0 이 살아남는가
'==============================================================
Private Sub T09_LeadingZero()
    Dim p As String, cnt As Long, msg As String

    p = Environ$("TEMP") & "\wsch_test_emp.csv"
    WriteUtf8 p, "이름,내선번호,부서,직급" & vbCrLf & _
                 "영벌이,0212,전산팀,대리" & vbCrLf & _
                 "홍길동,1234,기획팀,과장" & vbCrLf

    EmpRefreshFromFile p, cnt, msg
    Check "9a. CSV 명부 2명 읽힘", cnt = 2, msg & " (cnt=" & cnt & ")"
    Check "9b. 0212 가 212 로 바뀌지 않음", _
          EmpExtension("영벌이", "전산팀") = "0212", _
          "읽힌 값: [" & EmpExtension("영벌이", "전산팀") & "]"

    On Error Resume Next
    Kill p
    On Error GoTo 0
End Sub

'==============================================================
' 10. 부서·직급이 완전히 빈 명부로도 동작하는가
'==============================================================
Private Sub T10_EmptyDeptRank()
    Dim p As String, cnt As Long, msg As String, m As Variant

    p = Environ$("TEMP") & "\wsch_test_emp2.csv"
    WriteUtf8 p, "이름,내선번호,부서,직급" & vbCrLf & _
                 "최소한,3001,," & vbCrLf & _
                 "김최소,3002,," & vbCrLf

    EmpRefreshFromFile p, cnt, msg
    Check "10a. 부서·직급이 비어도 읽힘", cnt = 2, msg & " (cnt=" & cnt & ")"

    m = EmpSearch("최소")
    Check "10b. 검색이 동작함", Not IsEmpty(m), "검색 결과 없음"

    Check "10c. 부서가 비어도 내선 조회됨", _
          EmpExtension("최소한", "") = "3001", _
          "읽힌 값: [" & EmpExtension("최소한", "") & "]"

    ' 오늘 시트 렌더링이 빈 부서에서도 죽지 않는지
    On Error Resume Next
    RenderToday
    Check "10d. 렌더링이 오류 없이 끝남", Err.Number = 0, "오류 " & Err.Description
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    Kill p
    On Error GoTo 0
End Sub

'==============================================================
' 테스트 도우미
'==============================================================
Private Sub Check(ByVal name As String, ByVal cond As Boolean, ByVal detail As String)
    If cond Then
        mPass = mPass + 1
        mLog = mLog & "[통과] " & name & vbLf
    Else
        Fail name, detail
    End If
End Sub

Private Sub Fail(ByVal name As String, ByVal detail As String)
    mFail = mFail + 1
    mLog = mLog & "[실패] " & name & "  -> " & detail & vbLf
End Sub

Private Sub InsertShare(ByVal id As String, ByVal docNo As String, ByVal title As String, _
                        ByVal content As String, ByVal modAt As String, ByVal delFlag As String)
    Dim rec As Variant
    rec = TB_NewRec(TBL_SHARE)
    RecSet rec, TBL_SHARE, "ID", id
    RecSet rec, TBL_SHARE, "문서번호", docNo
    RecSet rec, TBL_SHARE, "제목", title
    RecSet rec, TBL_SHARE, "내용", content
    RecSet rec, TBL_SHARE, "등록일", "2026-01-01"
    RecSet rec, TBL_SHARE, "회신기한", "2026-01-10"
    RecSet rec, TBL_SHARE, "작성자", "테스터"
    RecSet rec, TBL_SHARE, "수정시각", modAt
    RecSet rec, TBL_SHARE, "삭제여부", delFlag
    TB_Insert TBL_SHARE, rec
End Sub

Private Function GetShareField(ByVal id As String, ByVal colName As String) As String
    Dim r As Long, rec As Variant
    r = TB_FindRowByID(TBL_SHARE, id)
    If r = 0 Then
        GetShareField = "(없음)"
        Exit Function
    End If
    rec = TB_GetRow(TBL_SHARE, r)
    GetShareField = RecGet(rec, TBL_SHARE, colName)
End Function

Private Sub ClearAll(ByVal tName As String)
    Dim lo As ListObject
    Set lo = TB(tName)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

' 스냅샷을 그대로 되돌려 놓는다
Private Sub PutAll(ByVal tName As String, ByRef v As Variant)
    Dim lo As ListObject, ws As Worksheet, nc As Long
    Set lo = TB(tName)
    Set ws = lo.Parent
    nc = lo.ListColumns.Count
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If IsEmpty(v) Then Exit Sub
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), _
                       lo.HeaderRowRange.Cells(1, nc).Offset(UBound(v, 1), 0))
    lo.DataBodyRange.Value = v
End Sub

Private Sub SetEmployees(ByVal rowsArr As Variant)
    Dim lo As ListObject, ws As Worksheet
    Dim arr As Variant, i As Long, c As Long, n As Long
    Set lo = TB(TBL_EMP)
    Set ws = lo.Parent
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    n = UBound(rowsArr) - LBound(rowsArr) + 1
    ReDim arr(1 To n, 1 To 4)
    For i = 1 To n
        For c = 1 To 4
            arr(i, c) = rowsArr(LBound(rowsArr) + i - 1)(c - 1)
        Next c
    Next i
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), _
                       lo.HeaderRowRange.Cells(1, 4).Offset(n, 0))
    lo.DataBodyRange.NumberFormat = "@"
    lo.DataBodyRange.Value = arr
    EmpInvalidate
End Sub

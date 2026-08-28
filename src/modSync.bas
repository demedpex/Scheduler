Attribute VB_Name = "modSync"
Option Explicit
'==============================================================
' modSync - 텍스트 동기화 (내보내기 / 가져오기 / 병합)
'
' 메신저 본문에 복사·붙여넣기로만 데이터를 주고받는다.
' 첨부파일도 공유폴더도 없다는 전제이므로 이 모듈이 이 프로그램의 핵심이다.
'
' 포맷
'   ###WSCH1|P1/3|N=5###
'   S|ID|문서번호|제목|내용|등록일|회신기한|작성자|수정시각|삭제여부
'   ...
'   ###END|P1/3|C=A3F1###
'
' 체크섬은 "본문 줄들을 vbLf 로 이은 문자열" 에 대해 계산한다.
' 메신저가 줄바꿈을 CRLF/CR 로 바꿔도 결과가 달라지지 않게 하기 위함이다.
'==============================================================

' 내보내기 범위
Public Const SCOPE_ALL As Long = 1
Public Const SCOPE_MINE As Long = 2
Public Const SCOPE_RECENT As Long = 3
Public Const SCOPE_OPEN As Long = 4

' SyncFeed 결과 코드
Public Const FD_NOHEADER As Long = 0        ' 형식 불일치
Public Const FD_BADCHECKSUM As Long = 1     ' 체크섬 불일치 / 잘림
Public Const FD_PARTIAL As Long = 2         ' 조각 일부만 도착
Public Const FD_COMPLETE As Long = 3        ' 전부 도착
Public Const FD_DUPLICATE As Long = 4       ' 이미 받은 조각

Private Const REC_SHARE As String = "S"

' 받은 조각 버퍼: 파트번호(문자열) -> 본문(vbLf 로 이어진 줄들)
Private mInbox As Object
Private mTotal As Long

'==============================================================
' 내보내기
'==============================================================

' 조건에 맞는 tbl_Share 레코드를 파트 텍스트 배열로 만든다.
' 반환: 1부터 시작하는 String 배열. 대상이 없으면 Empty.
Public Function SyncBuildParts(ByVal scope As Long, ByVal days As Long, _
                               ByRef outRecCount As Long) As Variant
    Dim lines As Collection
    Dim parts As Collection, cur As Collection
    Dim data As Variant, r As Long, nr As Long
    Dim rec As Variant, ln As String
    Dim curLen As Long, addLen As Long
    Dim i As Long, j As Long
    Dim res() As String, body As String, tmp As Collection

    outRecCount = 0
    Set lines = New Collection

    data = TB_Data(TBL_SHARE)
    If IsEmpty(data) Then
        SyncBuildParts = Empty
        Exit Function
    End If
    nr = UBound(data, 1)

    For r = 1 To nr
        rec = RowOf(data, r)
        If Len(RecGet(rec, TBL_SHARE, "ID")) > 0 Then
            If MatchScope(rec, scope, days) Then
                lines.Add EncodeShareLine(rec)
            End If
        End If
    Next r

    outRecCount = lines.Count
    If lines.Count = 0 Then
        SyncBuildParts = Empty
        Exit Function
    End If

    ' 1차: 레코드 단위로 파트에 담는다. 레코드 중간에서 자르지 않는다.
    Set parts = New Collection
    Set cur = New Collection
    curLen = 0
    For i = 1 To lines.Count
        ln = lines(i)
        If cur.Count = 0 Then
            addLen = Len(ln)
        Else
            addLen = Len(ln) + 1                ' 줄바꿈 1글자
        End If
        If cur.Count > 0 And curLen + addLen > PART_MAX_CHARS Then
            parts.Add cur
            Set cur = New Collection
            curLen = Len(ln)
            cur.Add ln
        Else
            cur.Add ln
            curLen = curLen + addLen
        End If
    Next i
    If cur.Count > 0 Then parts.Add cur

    ' 2차: 전체 파트 수가 정해졌으니 머리말·꼬리말을 붙여 렌더링
    ReDim res(1 To parts.Count)
    For i = 1 To parts.Count
        Set tmp = parts(i)
        body = ""
        For j = 1 To tmp.Count
            If j > 1 Then body = body & vbLf
            body = body & tmp(j)
        Next j
        res(i) = "###" & PROTO_TAG & "|P" & i & "/" & parts.Count & "|N=" & tmp.Count & "###" & vbCrLf & _
                 Replace$(body, vbLf, vbCrLf) & vbCrLf & _
                 "###END|P" & i & "/" & parts.Count & "|C=" & Checksum4(body) & "###"
    Next i

    SyncBuildParts = res
End Function

Private Function MatchScope(ByRef rec As Variant, ByVal scope As Long, ByVal days As Long) As Boolean
    Dim d As Date, reg As Date
    Select Case scope
        Case SCOPE_MINE
            MatchScope = (StrComp(RecGet(rec, TBL_SHARE, "작성자"), MyName(), vbTextCompare) = 0)
        Case SCOPE_RECENT
            reg = ParseDate(RecGet(rec, TBL_SHARE, "등록일"))
            MatchScope = (reg > 0 And reg >= Date - days)
        Case SCOPE_OPEN
            d = ParseDate(RecGet(rec, TBL_SHARE, "회신기한"))
            MatchScope = (d > 0 And d >= Date And Not ToBool(RecGet(rec, TBL_SHARE, "삭제여부")))
        Case Else
            MatchScope = True
    End Select
End Function

Private Function EncodeShareLine(ByRef rec As Variant) As String
    EncodeShareLine = REC_SHARE & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "ID")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "문서번호")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "제목")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "내용")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "등록일")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "회신기한")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "작성자")) & "|" & _
        Esc(RecGet(rec, TBL_SHARE, "수정시각")) & "|" & _
        Esc(BoolStr(ToBool(RecGet(rec, TBL_SHARE, "삭제여부"))))
End Function

' 한 줄 -> tbl_Share 레코드. 형식이 아니면 Empty.
Private Function DecodeShareLine(ByVal ln As String) As Variant
    Dim f() As String, rec As Variant
    f = Split(ln, "|")
    If UBound(f) < 9 Then
        DecodeShareLine = Empty
        Exit Function
    End If
    If f(0) <> REC_SHARE Then
        DecodeShareLine = Empty
        Exit Function
    End If
    rec = TB_NewRec(TBL_SHARE)
    RecSet rec, TBL_SHARE, "ID", Unesc(f(1))
    RecSet rec, TBL_SHARE, "문서번호", Unesc(f(2))
    RecSet rec, TBL_SHARE, "제목", Unesc(f(3))
    RecSet rec, TBL_SHARE, "내용", Unesc(f(4))
    RecSet rec, TBL_SHARE, "등록일", Unesc(f(5))
    RecSet rec, TBL_SHARE, "회신기한", Unesc(f(6))
    RecSet rec, TBL_SHARE, "작성자", Unesc(f(7))
    RecSet rec, TBL_SHARE, "수정시각", Unesc(f(8))
    RecSet rec, TBL_SHARE, "삭제여부", BoolStr(ToBool(Unesc(f(9))))
    DecodeShareLine = rec
End Function

'==============================================================
' 가져오기 - 조각 수신 버퍼
'==============================================================

Public Sub SyncResetInbox()
    Set mInbox = NewDict()
    mTotal = 0
End Sub

Public Function SyncInboxCount() As Long
    If mInbox Is Nothing Then SyncResetInbox
    SyncInboxCount = mInbox.Count
End Function

Public Function SyncInboxTotal() As Long
    SyncInboxTotal = mTotal
End Function

' 아직 못 받은 파트 번호를 "1, 3" 형태로
Public Function SyncMissingList() As String
    Dim i As Long, s0 As String
    If mInbox Is Nothing Then SyncResetInbox
    For i = 1 To mTotal
        If Not mInbox.Exists(CStr(i)) Then
            If Len(s0) > 0 Then s0 = s0 & ", "
            s0 = s0 & CStr(i)
        End If
    Next i
    SyncMissingList = s0
End Function

'--------------------------------------------------------------
' 붙여넣은 텍스트를 먹인다. 여러 조각이 한꺼번에 들어와도 처리한다.
' 같은 조각을 두 번 먹여도 안전하다(멱등).
'--------------------------------------------------------------
Public Function SyncFeed(ByVal text As String, ByRef msg As String) As Long
    Dim lines() As String, i As Long
    Dim n As Long, ln As String
    Dim inPart As Boolean
    Dim pNo As Long, pTot As Long, pN As Long
    Dim bodyLines As Collection
    Dim gotAny As Boolean, gotNew As Boolean, badSum As Boolean, truncated As Boolean
    Dim body As String, j As Long
    Dim hdrNo As Long, hdrTot As Long

    If mInbox Is Nothing Then SyncResetInbox

    text = Replace$(text, vbCrLf, vbLf)
    text = Replace$(text, vbCr, vbLf)
    lines = Split(text, vbLf)

    inPart = False
    Set bodyLines = New Collection

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextLine

        If Left$(ln, 3 + Len(PROTO_TAG) + 1) = "###" & PROTO_TAG & "|" Then
            ' 새 머리말. 이전 파트가 꼬리말 없이 끝났다면 그건 잘린 것이다.
            If inPart Then truncated = True
            If ParseHeader(ln, hdrNo, hdrTot, pN) Then
                inPart = True
                pNo = hdrNo
                pTot = hdrTot
                Set bodyLines = New Collection
                gotAny = True
            Else
                inPart = False
            End If

        ElseIf Left$(ln, 7) = "###END|" Then
            If inPart Then
                Dim fNo As Long, fTot As Long, fSum As String
                If ParseFooter(ln, fNo, fTot, fSum) Then
                    If fNo <> pNo Or fTot <> pTot Then
                        badSum = True
                    Else
                        body = ""
                        For j = 1 To bodyLines.Count
                            If j > 1 Then body = body & vbLf
                            body = body & bodyLines(j)
                        Next j
                        If StrComp(Checksum4(body), fSum, vbTextCompare) <> 0 Then
                            badSum = True
                        ElseIf bodyLines.Count <> pN Then
                            badSum = True
                        Else
                            ' 정상 조각 접수
                            If mTotal <> pTot Then
                                ' 전체 파트 수가 달라졌다 = 다른 묶음이다. 버퍼를 새로 시작.
                                SyncResetInbox
                                mTotal = pTot
                            End If
                            If Not mInbox.Exists(CStr(pNo)) Then
                                mInbox(CStr(pNo)) = body
                                gotNew = True
                            End If
                        End If
                    End If
                Else
                    badSum = True
                End If
                inPart = False
            End If

        ElseIf inPart Then
            bodyLines.Add ln
        End If
NextLine:
    Next i

    If inPart Then truncated = True

    '---- 상태 판정 ----
    If Not gotAny Then
        If mInbox.Count > 0 Then
            msg = "인식할 수 없는 텍스트입니다. 머리말(###" & PROTO_TAG & ")이 포함됐는지 확인하세요." & _
                  " (현재 " & mInbox.Count & "/" & mTotal & " 조각 보관 중)"
        Else
            msg = "인식할 수 없는 텍스트입니다. 머리말(###" & PROTO_TAG & ")이 포함됐는지 확인하세요."
        End If
        SyncFeed = FD_NOHEADER
        Exit Function
    End If

    If truncated Or badSum Then
        msg = "복사 중 텍스트가 잘렸거나 변형됐습니다. 다시 복사해 주세요."
        SyncFeed = FD_BADCHECKSUM
        Exit Function
    End If

    If mTotal > 0 And mInbox.Count >= mTotal Then
        msg = "조각 " & mTotal & "개를 모두 받았습니다. 아래 목록을 확인하고 [반영]을 누르세요."
        SyncFeed = FD_COMPLETE
        Exit Function
    End If

    If Not gotNew Then
        msg = "이미 받은 조각입니다. " & mInbox.Count & "/" & mTotal & " 조각 보관 중입니다. " & _
              "남은 조각(" & SyncMissingList() & ")을 이어서 붙여넣으세요."
        SyncFeed = FD_DUPLICATE
        Exit Function
    End If

    msg = mInbox.Count & "/" & mTotal & " 조각까지 받았습니다. " & _
          "남은 조각 " & (mTotal - mInbox.Count) & "개(" & SyncMissingList() & ")를 이어서 붙여넣으세요."
    SyncFeed = FD_PARTIAL
End Function

' "###WSCH1|P1/3|N=5###"
Private Function ParseHeader(ByVal ln As String, ByRef pNo As Long, ByRef pTot As Long, _
                            ByRef pN As Long) As Boolean
    Dim core As String, f() As String
    ParseHeader = False
    If Right$(ln, 3) <> "###" Then Exit Function
    core = Mid$(ln, 4, Len(ln) - 6)                 ' 앞뒤 ### 제거
    f = Split(core, "|")
    If UBound(f) <> 2 Then Exit Function
    If f(0) <> PROTO_TAG Then Exit Function
    If Not ParsePartNo(f(1), pNo, pTot) Then Exit Function
    If Left$(f(2), 2) <> "N=" Then Exit Function
    If Not IsNumeric(Mid$(f(2), 3)) Then Exit Function
    pN = CLng(Mid$(f(2), 3))
    ParseHeader = True
End Function

' "###END|P1/3|C=A3F1###"
Private Function ParseFooter(ByVal ln As String, ByRef pNo As Long, ByRef pTot As Long, _
                            ByRef sum4 As String) As Boolean
    Dim core As String, f() As String
    ParseFooter = False
    If Right$(ln, 3) <> "###" Then Exit Function
    core = Mid$(ln, 4, Len(ln) - 6)
    f = Split(core, "|")
    If UBound(f) <> 2 Then Exit Function
    If f(0) <> "END" Then Exit Function
    If Not ParsePartNo(f(1), pNo, pTot) Then Exit Function
    If Left$(f(2), 2) <> "C=" Then Exit Function
    sum4 = Mid$(f(2), 3)
    If Len(sum4) <> 4 Then Exit Function
    ParseFooter = True
End Function

' "P1/3"
Private Function ParsePartNo(ByVal s0 As String, ByRef pNo As Long, ByRef pTot As Long) As Boolean
    Dim p As Long, a As String, b As String
    ParsePartNo = False
    If Left$(s0, 1) <> "P" Then Exit Function
    p = InStr(s0, "/")
    If p = 0 Then Exit Function
    a = Mid$(s0, 2, p - 2)
    b = Mid$(s0, p + 1)
    If Not IsNumeric(a) Or Not IsNumeric(b) Then Exit Function
    pNo = CLng(a)
    pTot = CLng(b)
    If pNo < 1 Or pTot < 1 Or pNo > pTot Then Exit Function
    ParsePartNo = True
End Function

'==============================================================
' 버퍼에 모인 조각 -> 레코드 목록
'   같은 ID 가 여러 번 나오면 수정시각이 최신인 것만 남긴다.
'==============================================================
Private Function InboxRecords() As Object
    Dim d As Object, i As Long, j As Long
    Dim body As String, lines() As String
    Dim rec As Variant, id As String, old As Variant

    Set d = NewDict()
    If mInbox Is Nothing Then
        Set InboxRecords = d
        Exit Function
    End If
    For i = 1 To mTotal
        If mInbox.Exists(CStr(i)) Then
            body = mInbox(CStr(i))
            lines = Split(body, vbLf)
            For j = LBound(lines) To UBound(lines)
                If Len(Trim$(lines(j))) > 0 Then
                    rec = DecodeShareLine(lines(j))
                    If Not IsEmpty(rec) Then
                        id = RecGet(rec, TBL_SHARE, "ID")
                        If Len(id) > 0 Then
                            If d.Exists(id) Then
                                old = d(id)
                                If RecGet(rec, TBL_SHARE, "수정시각") >= RecGet(old, TBL_SHARE, "수정시각") Then
                                    d(id) = rec
                                End If
                            Else
                                d(id) = rec
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    Set InboxRecords = d
End Function

'==============================================================
' 미리보기
'   반환: (1..n, 1..4) = 상태, 문서번호, 제목, ID
'==============================================================
Public Function SyncPreview() As Variant
    Dim d As Object, k As Variant, rec As Variant
    Dim res As Variant, i As Long

    Set d = InboxRecords()
    If d.Count = 0 Then
        SyncPreview = Empty
        Exit Function
    End If
    ReDim res(1 To d.Count, 1 To 4)
    i = 0
    For Each k In d.Keys
        i = i + 1
        rec = d(k)
        res(i, 1) = MergeVerdict(rec)
        res(i, 2) = RecGet(rec, TBL_SHARE, "문서번호")
        res(i, 3) = RecGet(rec, TBL_SHARE, "제목")
        res(i, 4) = CStr(k)
    Next k
    SyncPreview = res
End Function

' 이 레코드를 병합하면 어떻게 되는지 판정
Private Function MergeVerdict(ByRef rec As Variant) As String
    Dim r As Long, cur As Variant
    Dim incDel As Boolean, locDel As Boolean

    r = TB_FindRowByID(TBL_SHARE, RecGet(rec, TBL_SHARE, "ID"))
    If r = 0 Then
        If ToBool(RecGet(rec, TBL_SHARE, "삭제여부")) Then
            MergeVerdict = "[삭제]"          ' 없던 건의 삭제 표식 -> 묘비만 받아둔다
        Else
            MergeVerdict = "[신규]"
        End If
        Exit Function
    End If

    cur = TB_GetRow(TBL_SHARE, r)
    incDel = ToBool(RecGet(rec, TBL_SHARE, "삭제여부"))
    locDel = ToBool(RecGet(cur, TBL_SHARE, "삭제여부"))

    If locDel Then
        ' 내가 이미 지운 건은 되살리지 않는다 (tombstone 우선)
        MergeVerdict = "[변화없음]"
    ElseIf incDel Then
        MergeVerdict = "[삭제]"
    ElseIf RecGet(rec, TBL_SHARE, "수정시각") > RecGet(cur, TBL_SHARE, "수정시각") Then
        MergeVerdict = "[갱신]"
    Else
        MergeVerdict = "[변화없음]"
    End If
End Function

'==============================================================
' 병합 실행
'==============================================================
Public Function SyncApply() As String
    Dim d As Object, k As Variant, rec As Variant, cur As Variant
    Dim r As Long
    Dim nAdd As Long, nUpd As Long, nSame As Long, nDel As Long
    Dim dupList As String, src As String
    Dim incDel As Boolean, locDel As Boolean
    Dim docDict As Object, doc As String

    Set d = InboxRecords()
    If d.Count = 0 Then
        SyncApply = "반영할 내용이 없습니다."
        Exit Function
    End If

    ' 기존 문서번호 -> ID 색인 (중복 검출용)
    Set docDict = BuildDocNoIndex()

    For Each k In d.Keys
        rec = d(k)
        src = RecGet(rec, TBL_SHARE, "작성자")
        r = TB_FindRowByID(TBL_SHARE, CStr(k))
        incDel = ToBool(RecGet(rec, TBL_SHARE, "삭제여부"))

        If r = 0 Then
            TB_Insert TBL_SHARE, rec
            If incDel Then
                nDel = nDel + 1
                LogMerge "삭제표식수신", CStr(k), "", ShareSummary(rec), src
            Else
                nAdd = nAdd + 1
                LogMerge "신규", CStr(k), "", ShareSummary(rec), src
            End If
        Else
            cur = TB_GetRow(TBL_SHARE, r)
            locDel = ToBool(RecGet(cur, TBL_SHARE, "삭제여부"))
            If locDel Then
                nSame = nSame + 1                       ' 되살리지 않는다
            ElseIf incDel Then
                RecSet cur, TBL_SHARE, "삭제여부", "1"
                RecSet cur, TBL_SHARE, "수정시각", RecGet(rec, TBL_SHARE, "수정시각")
                TB_UpdateRow TBL_SHARE, r, cur
                nDel = nDel + 1
                LogMerge "삭제", CStr(k), ShareSummary(cur), "", src
            ElseIf RecGet(rec, TBL_SHARE, "수정시각") > RecGet(cur, TBL_SHARE, "수정시각") Then
                LogMerge "갱신", CStr(k), ShareSummary(cur), ShareSummary(rec), src
                TB_UpdateRow TBL_SHARE, r, rec
                nUpd = nUpd + 1
            Else
                nSame = nSame + 1
            End If
        End If

        ' 문서번호가 같은데 ID 가 다른 건 경고 대상
        doc = RecGet(rec, TBL_SHARE, "문서번호")
        If Len(doc) > 0 Then
            If docDict.Exists(doc) Then
                If StrComp(CStr(docDict(doc)), CStr(k), vbBinaryCompare) <> 0 Then
                    If InStr(dupList, doc) = 0 Then
                        If Len(dupList) > 0 Then dupList = dupList & ", "
                        dupList = dupList & doc
                    End If
                End If
            Else
                docDict(doc) = CStr(k)
            End If
        End If
    Next k

    CfgSet "마지막동기화시각", NowStamp()

    SyncApply = "신규 " & nAdd & "건, 갱신 " & nUpd & "건, 변화없음 " & nSame & "건" & _
                IIf(nDel > 0, ", 삭제 " & nDel & "건", "") & " 반영했습니다."
    If Len(dupList) > 0 Then
        SyncApply = SyncApply & vbLf & vbLf & _
                    "※ 문서번호가 중복된 건이 있습니다: " & dupList & vbLf & _
                    "   (병합은 그대로 했습니다. 어느 쪽이 맞는지 확인해 주세요.)"
    End If

    SyncResetInbox
End Function

' 살아있는 레코드의 문서번호 -> ID 색인
Private Function BuildDocNoIndex() As Object
    Dim d As Object, data As Variant, r As Long
    Dim dc As Long, ic As Long, xc As Long, doc As String
    Set d = NewDict()
    dc = TB_Col(TBL_SHARE, "문서번호")
    ic = TB_Col(TBL_SHARE, "ID")
    xc = TB_Col(TBL_SHARE, "삭제여부")
    data = TB_Data(TBL_SHARE)
    If IsEmpty(data) Then
        Set BuildDocNoIndex = d
        Exit Function
    End If
    For r = 1 To UBound(data, 1)
        If Not ToBool(data(r, xc)) Then
            doc = S(data(r, dc))
            If Len(doc) > 0 Then
                If Not d.Exists(doc) Then d(doc) = S(data(r, ic))
            End If
        End If
    Next r
    Set BuildDocNoIndex = d
End Function

'==============================================================
' 등록 시 문서번호 중복 검사 (차단하지 않고 경고만)
'==============================================================
Public Function DocNoExists(ByVal docNo As String, ByVal exceptID As String) As String
    Dim data As Variant, r As Long
    Dim dc As Long, ic As Long, xc As Long, ac As Long
    DocNoExists = ""
    If Len(Trim$(docNo)) = 0 Then Exit Function
    dc = TB_Col(TBL_SHARE, "문서번호")
    ic = TB_Col(TBL_SHARE, "ID")
    xc = TB_Col(TBL_SHARE, "삭제여부")
    ac = TB_Col(TBL_SHARE, "작성자")
    data = TB_Data(TBL_SHARE)
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If Not ToBool(data(r, xc)) Then
            If StrComp(S(data(r, dc)), docNo, vbTextCompare) = 0 Then
                If StrComp(S(data(r, ic)), exceptID, vbBinaryCompare) <> 0 Then
                    DocNoExists = S(data(r, ac))
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

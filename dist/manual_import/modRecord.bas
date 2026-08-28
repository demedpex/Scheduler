Attribute VB_Name = "modRecord"
Option Explicit
'==============================================================
' modRecord - 레코드 저장/삭제 규칙
'
' UserForm 경로와 시트 패널 경로가 같은 코드를 쓰도록 여기에 모아 둔다.
' 검증 실패 시 사용자에게 보여줄 메시지를 돌려주고, 성공하면 빈 문자열.
'==============================================================

'--------------------------------------------------------------
' 업무
'--------------------------------------------------------------
Public Function SaveTask(ByVal id As String, ByVal title As String, ByVal content As String, _
                         ByVal dateS As String, ByVal timeS As String) As String
    Dim rec As Variant, r As Long, d As Date, tm As String

    If Len(Trim$(title)) = 0 Then
        SaveTask = "제목을 입력해 주세요."
        Exit Function
    End If
    If Len(Trim$(dateS)) = 0 Then dateS = TodayStr()
    d = ParseDate(dateS)
    If d = 0 Then
        SaveTask = "날짜는 2026-08-27 형식으로 입력해 주세요."
        Exit Function
    End If
    tm = NormTime(timeS)
    If Len(Trim$(timeS)) > 0 And Len(tm) = 0 Then
        SaveTask = "시간은 14:30 형식으로 입력해 주세요. (비워 두셔도 됩니다)"
        Exit Function
    End If

    If Len(id) = 0 Then
        rec = TB_NewRec(TBL_TASK)
        RecSet rec, TBL_TASK, "ID", NewID(MyName())
        RecSet rec, TBL_TASK, "작성자", MyName()
        RecSet rec, TBL_TASK, "완료여부", "0"
        RecSet rec, TBL_TASK, "삭제여부", "0"
    Else
        r = TB_FindRowByID(TBL_TASK, id)
        If r = 0 Then
            SaveTask = "수정할 업무를 찾지 못했습니다."
            Exit Function
        End If
        rec = TB_GetRow(TBL_TASK, r)
    End If

    RecSet rec, TBL_TASK, "제목", Trim$(title)
    RecSet rec, TBL_TASK, "내용", content
    RecSet rec, TBL_TASK, "날짜", DateStr(d)
    RecSet rec, TBL_TASK, "시간", tm
    RecSet rec, TBL_TASK, "수정시각", NowStamp()

    If Len(id) = 0 Then
        TB_Insert TBL_TASK, rec
    Else
        TB_UpdateRow TBL_TASK, r, rec
    End If
    SaveTask = ""
End Function

'--------------------------------------------------------------
' 연락하기
'   내선번호는 저장하지 않는다. 명부가 갱신되면 자동으로 최신값이 되도록
'   조회할 때마다 이름+부서로 명부에서 끌어온다.
'   내선캐시는 명부에 없는 사람일 때만 채운다.
'--------------------------------------------------------------
Public Function SaveContact(ByVal id As String, ByVal empName As String, ByVal dept As String, _
                            ByVal extCache As String, ByVal dateS As String, ByVal timeS As String) As String
    Dim rec As Variant, r As Long, d As Date, tm As String, dstr As String

    If Len(Trim$(empName)) = 0 Then
        SaveContact = "이름을 입력해 주세요."
        Exit Function
    End If

    dstr = ""
    If Len(Trim$(dateS)) > 0 Then
        d = ParseDate(dateS)
        If d = 0 Then
            SaveContact = "날짜는 2026-08-27 형식으로 입력해 주세요. (비워 두셔도 됩니다)"
            Exit Function
        End If
        dstr = DateStr(d)
    End If

    tm = NormTime(timeS)
    If Len(Trim$(timeS)) > 0 And Len(tm) = 0 Then
        SaveContact = "시간은 14:30 형식으로 입력해 주세요. (비워 두셔도 됩니다)"
        Exit Function
    End If

    ' 명부에 있으면 내선캐시는 비워 둔다 (명부가 항상 우선)
    If Len(EmpExtension(empName, dept)) > 0 Then extCache = ""

    If Len(id) = 0 Then
        rec = TB_NewRec(TBL_CONTACT)
        RecSet rec, TBL_CONTACT, "ID", NewID(MyName())
        RecSet rec, TBL_CONTACT, "작성자", MyName()
        RecSet rec, TBL_CONTACT, "완료여부", "0"
        RecSet rec, TBL_CONTACT, "삭제여부", "0"
    Else
        r = TB_FindRowByID(TBL_CONTACT, id)
        If r = 0 Then
            SaveContact = "수정할 연락 건을 찾지 못했습니다."
            Exit Function
        End If
        rec = TB_GetRow(TBL_CONTACT, r)
    End If

    RecSet rec, TBL_CONTACT, "이름", Trim$(empName)
    RecSet rec, TBL_CONTACT, "부서", Trim$(dept)
    RecSet rec, TBL_CONTACT, "내선캐시", Trim$(extCache)
    RecSet rec, TBL_CONTACT, "날짜", dstr
    RecSet rec, TBL_CONTACT, "시간", tm
    RecSet rec, TBL_CONTACT, "수정시각", NowStamp()

    If Len(id) = 0 Then
        TB_Insert TBL_CONTACT, rec
    Else
        TB_UpdateRow TBL_CONTACT, r, rec
    End If
    SaveContact = ""
End Function

'--------------------------------------------------------------
' 업무공유 (동기화 대상)
'--------------------------------------------------------------
Public Function SaveShare(ByVal id As String, ByVal docNo As String, ByVal title As String, _
                          ByVal content As String, ByVal regS As String, ByVal dueS As String) As String
    Dim rec As Variant, r As Long, reg As Date, due As Date, owner As String

    If Len(Trim$(docNo)) = 0 Then
        SaveShare = "문서번호를 입력해 주세요."
        Exit Function
    End If
    If Len(Trim$(title)) = 0 Then
        SaveShare = "제목을 입력해 주세요."
        Exit Function
    End If
    If Len(Trim$(regS)) = 0 Then regS = TodayStr()
    reg = ParseDate(regS)
    If reg = 0 Then
        SaveShare = "등록일은 2026-08-27 형식으로 입력해 주세요."
        Exit Function
    End If
    due = ParseDate(dueS)
    If due = 0 Then
        SaveShare = "회신기한을 2026-08-30 형식으로 입력해 주세요. (필수 항목입니다)"
        Exit Function
    End If

    ' 문서번호 중복은 경고만 하고 막지는 않는다
    owner = DocNoExists(Trim$(docNo), id)
    If Len(owner) > 0 Then
        If Not Ask("문서번호 " & Trim$(docNo) & " 은(는) 이미 등록돼 있습니다." & vbLf & _
                   "(등록자: " & owner & ")" & vbLf & vbLf & "그래도 등록할까요?") Then
            SaveShare = "-"
            Exit Function
        End If
    End If

    If Len(id) = 0 Then
        rec = TB_NewRec(TBL_SHARE)
        RecSet rec, TBL_SHARE, "ID", NewID(MyName())
        RecSet rec, TBL_SHARE, "작성자", MyName()
        RecSet rec, TBL_SHARE, "삭제여부", "0"
    Else
        r = TB_FindRowByID(TBL_SHARE, id)
        If r = 0 Then
            SaveShare = "수정할 공유 건을 찾지 못했습니다."
            Exit Function
        End If
        rec = TB_GetRow(TBL_SHARE, r)
    End If

    RecSet rec, TBL_SHARE, "문서번호", Trim$(docNo)
    RecSet rec, TBL_SHARE, "제목", Trim$(title)
    RecSet rec, TBL_SHARE, "내용", content
    RecSet rec, TBL_SHARE, "등록일", DateStr(reg)
    RecSet rec, TBL_SHARE, "회신기한", DateStr(due)
    RecSet rec, TBL_SHARE, "수정시각", NowStamp()

    If Len(id) = 0 Then
        TB_Insert TBL_SHARE, rec
    Else
        TB_UpdateRow TBL_SHARE, r, rec
    End If
    SaveShare = ""
End Function

'--------------------------------------------------------------
' 점심약속 (저장 직전 충돌 검사 2가지)
'--------------------------------------------------------------
Public Function SaveLunch(ByVal id As String, ByVal target As String, ByVal dateS As String, _
                          ByVal timeS As String, ByVal place As String) As String
    Dim rec As Variant, r As Long, d As Date, tm As String
    Dim other As String

    If Len(Trim$(target)) = 0 Then
        SaveLunch = "약속 대상 이름을 입력해 주세요."
        Exit Function
    End If
    If Len(Trim$(dateS)) = 0 Then dateS = TodayStr()
    d = ParseDate(dateS)
    If d = 0 Then
        SaveLunch = "날짜는 2026-08-27 형식으로 입력해 주세요."
        Exit Function
    End If
    tm = NormTime(timeS)
    If Len(Trim$(timeS)) > 0 And Len(tm) = 0 Then
        SaveLunch = "시간은 12:00 형식으로 입력해 주세요. (비워 두셔도 됩니다)"
        Exit Function
    End If

    ' 1) 같은 사람 + 같은 날짜 중복
    If LunchDup(DateStr(d), Trim$(target), id) Then
        If Not Ask(Format$(d, "m/d") & "에 " & Trim$(target) & "님과의 약속이 이미 등록돼 있습니다." & vbLf & _
                   "중복 등록할까요?") Then
            SaveLunch = "-"
            Exit Function
        End If
    Else
        ' 2) 같은 날짜에 다른 약속이 있는지
        other = LunchOnDate(DateStr(d), id)
        If Len(other) > 0 Then
            If Not Ask(Format$(d, "m/d") & "에 " & other & "님과 약속이 있습니다." & vbLf & _
                       "그래도 등록할까요?") Then
                SaveLunch = "-"
                Exit Function
            End If
        End If
    End If

    If Len(id) = 0 Then
        rec = TB_NewRec(TBL_LUNCH)
        RecSet rec, TBL_LUNCH, "ID", NewID(MyName())
        RecSet rec, TBL_LUNCH, "작성자", MyName()
        RecSet rec, TBL_LUNCH, "삭제여부", "0"
    Else
        r = TB_FindRowByID(TBL_LUNCH, id)
        If r = 0 Then
            SaveLunch = "수정할 약속을 찾지 못했습니다."
            Exit Function
        End If
        rec = TB_GetRow(TBL_LUNCH, r)
    End If

    RecSet rec, TBL_LUNCH, "대상이름", Trim$(target)
    RecSet rec, TBL_LUNCH, "날짜", DateStr(d)
    RecSet rec, TBL_LUNCH, "시간", tm
    RecSet rec, TBL_LUNCH, "장소", Trim$(place)
    RecSet rec, TBL_LUNCH, "수정시각", NowStamp()

    If Len(id) = 0 Then
        TB_Insert TBL_LUNCH, rec
    Else
        TB_UpdateRow TBL_LUNCH, r, rec
    End If
    SaveLunch = ""
End Function

' 같은 날짜에 있는 다른 약속의 대상 이름 (없으면 빈 문자열)
Public Function LunchOnDate(ByVal dstr As String, ByVal exceptID As String) As String
    Dim data As Variant, r As Long
    Dim dc As Long, nc As Long, ic As Long, xc As Long
    LunchOnDate = ""
    dc = TB_Col(TBL_LUNCH, "날짜")
    nc = TB_Col(TBL_LUNCH, "대상이름")
    ic = TB_Col(TBL_LUNCH, "ID")
    xc = TB_Col(TBL_LUNCH, "삭제여부")
    data = TB_Data(TBL_LUNCH)
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If Not ToBool(data(r, xc)) Then
            If S(data(r, dc)) = dstr And StrComp(S(data(r, ic)), exceptID, vbBinaryCompare) <> 0 Then
                LunchOnDate = S(data(r, nc))
                Exit Function
            End If
        End If
    Next r
End Function

Public Function LunchDup(ByVal dstr As String, ByVal target As String, ByVal exceptID As String) As Boolean
    Dim data As Variant, r As Long
    Dim dc As Long, nc As Long, ic As Long, xc As Long
    dc = TB_Col(TBL_LUNCH, "날짜")
    nc = TB_Col(TBL_LUNCH, "대상이름")
    ic = TB_Col(TBL_LUNCH, "ID")
    xc = TB_Col(TBL_LUNCH, "삭제여부")
    data = TB_Data(TBL_LUNCH)
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If Not ToBool(data(r, xc)) Then
            If S(data(r, dc)) = dstr And _
               StrComp(S(data(r, nc)), target, vbTextCompare) = 0 And _
               StrComp(S(data(r, ic)), exceptID, vbBinaryCompare) <> 0 Then
                LunchDup = True
                Exit Function
            End If
        End If
    Next r
End Function

'--------------------------------------------------------------
' 삭제 (물리 삭제 금지, 플래그만)
'--------------------------------------------------------------
Public Function DeleteRecord(ByVal tName As String, ByVal id As String) As Boolean
    If Len(id) = 0 Then Exit Function
    If Not Ask("이 항목을 삭제할까요?" & vbLf & _
               "(되돌릴 수 없습니다. 공유 건은 다음 내보내기 때 상대방에게도 삭제로 전달됩니다.)") Then
        Exit Function
    End If
    DeleteRecord = TB_SoftDeleteByID(tName, id)
End Function

'--------------------------------------------------------------
' 레코드 한 건 읽기 (폼/패널에서 초기값 채울 때)
'--------------------------------------------------------------
Public Function LoadRecord(ByVal tName As String, ByVal id As String) As Variant
    Dim r As Long
    r = TB_FindRowByID(tName, id)
    If r = 0 Then
        LoadRecord = Empty
    Else
        LoadRecord = TB_GetRow(tName, r)
    End If
End Function

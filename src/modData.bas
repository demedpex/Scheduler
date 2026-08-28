Attribute VB_Name = "modData"
Option Explicit
'==============================================================
' modData - 시트 테이블 CRUD
'   데이터는 전부 숨김 시트의 ListObject(표) 에 문자열로 저장한다.
'   레코드는 1차원 Variant 배열(1..컬럼수)로 주고받는다.
'==============================================================

' 표 이름
Public Const TBL_TASK As String = "tbl_Task"
Public Const TBL_CONTACT As String = "tbl_Contact"
Public Const TBL_SHARE As String = "tbl_Share"
Public Const TBL_LUNCH As String = "tbl_Lunch"
Public Const TBL_EMP As String = "mst_Employee"
Public Const TBL_CFG As String = "cfg_Setting"
Public Const TBL_LOG As String = "log_Merge"

' 보이는 시트 이름
Public Const SH_CAL As String = "달력"
Public Const SH_TODAY As String = "오늘"
Public Const SH_SETUP As String = "설정"

Private mLoCache As Object          ' 표이름 -> ListObject
Private mColCache As Object         ' 표이름 -> Dictionary(컬럼명 -> 인덱스)

'--------------------------------------------------------------
' 표 찾기
'--------------------------------------------------------------
Public Function TB(ByVal tName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    Dim cached As ListObject, probe As String
    If mLoCache Is Nothing Then Set mLoCache = NewDict()
    If mLoCache.Exists(tName) Then
        Set cached = mLoCache(tName)
        ' 캐시된 개체가 아직 살아 있는지 확인 (파일을 다시 열면 죽어 있을 수 있다)
        On Error Resume Next
        probe = cached.Name
        If Err.Number <> 0 Then
            Err.Clear
            Set cached = Nothing
            mLoCache.Remove tName
        End If
        On Error GoTo 0
        If Not cached Is Nothing Then
            Set TB = cached
            Exit Function
        End If
    End If
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tName, vbTextCompare) = 0 Then
                Set mLoCache(tName) = lo
                Set TB = lo
                Exit Function
            End If
        Next lo
    Next ws
    Err.Raise vbObjectError + 513, "modData.TB", "표를 찾을 수 없습니다: " & tName
End Function

Public Sub TB_ClearCache()
    Set mLoCache = Nothing
    Set mColCache = Nothing
End Sub

'--------------------------------------------------------------
' 컬럼
'--------------------------------------------------------------
Public Function TB_ColCount(ByVal tName As String) As Long
    TB_ColCount = TB(tName).ListColumns.Count
End Function

' 컬럼명 -> 1부터 시작하는 인덱스. 없으면 0.
Public Function TB_Col(ByVal tName As String, ByVal colName As String) As Long
    Dim d As Object, lc As ListColumn
    If mColCache Is Nothing Then Set mColCache = NewDict()
    If Not mColCache.Exists(tName) Then
        Set d = NewDict()
        For Each lc In TB(tName).ListColumns
            d(Trim$(lc.Name)) = lc.Index
        Next lc
        Set mColCache(tName) = d
    End If
    Set d = mColCache(tName)
    If d.Exists(colName) Then
        TB_Col = d(colName)
    Else
        TB_Col = 0
    End If
End Function

'--------------------------------------------------------------
' 읽기
'--------------------------------------------------------------
Public Function TB_RowCount(ByVal tName As String) As Long
    Dim lo As ListObject
    Set lo = TB(tName)
    If lo.DataBodyRange Is Nothing Then
        TB_RowCount = 0
    Else
        TB_RowCount = lo.DataBodyRange.Rows.Count
    End If
End Function

' 전체 데이터를 2차원 배열(1..행, 1..열)로. 행이 없으면 Empty.
Public Function TB_Data(ByVal tName As String) As Variant
    Dim lo As ListObject
    Set lo = TB(tName)
    If lo.DataBodyRange Is Nothing Then
        TB_Data = Empty
    Else
        TB_Data = lo.DataBodyRange.Value2
    End If
End Function

' 2차원 배열의 한 행을 1차원 레코드로
Public Function RowOf(ByVal data As Variant, ByVal r As Long) As Variant
    Dim rec As Variant, c As Long, nc As Long
    If IsEmpty(data) Then
        RowOf = Empty
        Exit Function
    End If
    nc = UBound(data, 2)
    ReDim rec(1 To nc)
    For c = 1 To nc
        rec(c) = data(r, c)
    Next c
    RowOf = rec
End Function

' 빈 레코드 (모든 칸 빈 문자열)
Public Function TB_NewRec(ByVal tName As String) As Variant
    Dim rec As Variant, c As Long, nc As Long
    nc = TB_ColCount(tName)
    ReDim rec(1 To nc)
    For c = 1 To nc
        rec(c) = ""
    Next c
    TB_NewRec = rec
End Function

' 레코드에서 컬럼명으로 값 읽기
Public Function RecGet(ByRef rec As Variant, ByVal tName As String, ByVal colName As String) As String
    Dim i As Long
    i = TB_Col(tName, colName)
    If i = 0 Then
        RecGet = ""
    Else
        RecGet = S(rec(i))
    End If
End Function

' 레코드에 컬럼명으로 값 쓰기
Public Sub RecSet(ByRef rec As Variant, ByVal tName As String, ByVal colName As String, ByVal v As String)
    Dim i As Long
    i = TB_Col(tName, colName)
    If i > 0 Then rec(i) = v
End Sub

'--------------------------------------------------------------
' 쓰기
'--------------------------------------------------------------
' 새 행 추가. 표의 첫 행이 완전히 비어 있으면 그 행을 재사용한다.
Public Sub TB_Insert(ByVal tName As String, ByRef rec As Variant)
    Dim lo As ListObject, rg As Range
    Set lo = TB(tName)
    If lo.DataBodyRange Is Nothing Then
        Set rg = lo.ListRows.Add.Range
    ElseIf lo.DataBodyRange.Rows.Count = 1 And _
           Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
        Set rg = lo.DataBodyRange.Rows(1)
    Else
        Set rg = lo.ListRows.Add.Range
    End If
    rg.Value = rec
End Sub

' rowIdx 는 DataBodyRange 기준 1부터
Public Sub TB_UpdateRow(ByVal tName As String, ByVal rowIdx As Long, ByRef rec As Variant)
    Dim lo As ListObject
    Set lo = TB(tName)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIdx < 1 Or rowIdx > lo.DataBodyRange.Rows.Count Then Exit Sub
    lo.DataBodyRange.Rows(rowIdx).Value = rec
End Sub

Public Function TB_GetRow(ByVal tName As String, ByVal rowIdx As Long) As Variant
    Dim lo As ListObject, v As Variant, rec As Variant, c As Long
    Set lo = TB(tName)
    If lo.DataBodyRange Is Nothing Then
        TB_GetRow = Empty
        Exit Function
    End If
    v = lo.DataBodyRange.Rows(rowIdx).Value2
    ReDim rec(1 To UBound(v, 2))
    For c = 1 To UBound(v, 2)
        rec(c) = v(1, c)
    Next c
    TB_GetRow = rec
End Function

' ID 로 행 번호 찾기. 없으면 0.
Public Function TB_FindRowByID(ByVal tName As String, ByVal id As String) As Long
    Dim data As Variant, r As Long, ci As Long
    TB_FindRowByID = 0
    If Len(id) = 0 Then Exit Function
    ci = TB_Col(tName, "ID")
    If ci = 0 Then Exit Function
    data = TB_Data(tName)
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If StrComp(S(data(r, ci)), id, vbBinaryCompare) = 0 Then
            TB_FindRowByID = r
            Exit Function
        End If
    Next r
End Function

' 물리 삭제 금지. 삭제여부 플래그만 켜고 수정시각을 갱신한다.
Public Function TB_SoftDeleteByID(ByVal tName As String, ByVal id As String) As Boolean
    Dim r As Long, rec As Variant
    r = TB_FindRowByID(tName, id)
    If r = 0 Then
        TB_SoftDeleteByID = False
        Exit Function
    End If
    rec = TB_GetRow(tName, r)
    RecSet rec, tName, "삭제여부", "1"
    RecSet rec, tName, "수정시각", NowStamp()
    TB_UpdateRow tName, r, rec
    TB_SoftDeleteByID = True
End Function

' 살아있는(삭제되지 않은) 레코드인지
Public Function IsAlive(ByVal data As Variant, ByVal r As Long, ByVal delCol As Long) As Boolean
    If delCol = 0 Then
        IsAlive = True
    Else
        IsAlive = Not ToBool(data(r, delCol))
    End If
End Function

'--------------------------------------------------------------
' 설정 (cfg_Setting)
'--------------------------------------------------------------
Public Function CfgGet(ByVal key As String, Optional ByVal defVal As String = "") As String
    Dim data As Variant, r As Long, kc As Long, vc As Long
    kc = TB_Col(TBL_CFG, "키")
    vc = TB_Col(TBL_CFG, "값")
    data = TB_Data(TBL_CFG)
    CfgGet = defVal
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If StrComp(S(data(r, kc)), key, vbTextCompare) = 0 Then
            CfgGet = S(data(r, vc))
            If Len(CfgGet) = 0 Then CfgGet = defVal
            Exit Function
        End If
    Next r
End Function

Public Sub CfgSet(ByVal key As String, ByVal val As String)
    Dim data As Variant, r As Long, kc As Long, vc As Long, rec As Variant
    kc = TB_Col(TBL_CFG, "키")
    vc = TB_Col(TBL_CFG, "값")
    data = TB_Data(TBL_CFG)
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            If StrComp(S(data(r, kc)), key, vbTextCompare) = 0 Then
                TB(TBL_CFG).DataBodyRange.Cells(r, vc).Value = val
                Exit Sub
            End If
        Next r
    End If
    rec = TB_NewRec(TBL_CFG)
    rec(kc) = key
    rec(vc) = val
    TB_Insert TBL_CFG, rec
End Sub

' 동기화에서 작성자로 기록되는 이름
Public Function MyName() As String
    Dim n As String
    n = CfgGet("내이름", "")
    If Len(n) = 0 Then n = Environ$("USERNAME")
    If Len(n) = 0 Then n = "사용자"
    MyName = n
End Function

'--------------------------------------------------------------
' 병합 로그
'--------------------------------------------------------------
Public Sub LogMerge(ByVal action As String, ByVal id As String, _
                    ByVal beforeTxt As String, ByVal afterTxt As String, _
                    ByVal source As String)
    Dim rec As Variant
    On Error Resume Next
    rec = TB_NewRec(TBL_LOG)
    RecSet rec, TBL_LOG, "시각", NowStamp()
    RecSet rec, TBL_LOG, "동작", action
    RecSet rec, TBL_LOG, "ID", id
    RecSet rec, TBL_LOG, "이전값요약", Ellipsis(beforeTxt, 200)
    RecSet rec, TBL_LOG, "이후값요약", Ellipsis(afterTxt, 200)
    RecSet rec, TBL_LOG, "출처", source
    TB_Insert TBL_LOG, rec
    On Error GoTo 0
End Sub

'--------------------------------------------------------------
' 레코드 요약 문자열 (로그·미리보기용)
'--------------------------------------------------------------
Public Function ShareSummary(ByRef rec As Variant) As String
    ShareSummary = RecGet(rec, TBL_SHARE, "문서번호") & " / " & _
                   RecGet(rec, TBL_SHARE, "제목") & " / 기한 " & _
                   RecGet(rec, TBL_SHARE, "회신기한")
End Function

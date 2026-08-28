Attribute VB_Name = "modForms"
Option Explicit
'==============================================================
' modForms - UserForm 진입점
'
' 이 모듈은 폼이 실제로 만들어진 경우에만 프로젝트에 들어간다.
' modMain 은 폼 이름을 직접 쓰지 않고 Application.Run 으로 여기를 부른다.
' (폼이 없는 프로젝트에 폼 이름이 남아 있으면 컴파일 자체가 안 되기 때문)
'
' 인자는 넘기지 않는다. 필요한 값은 modMain 의 속성에서 읽어 간다.
'==============================================================

Public Sub UI_Task()
    frmTask.Show
End Sub

Public Sub UI_Contact()
    frmContact.Show
End Sub

Public Sub UI_Share()
    frmShare.Show
End Sub

Public Sub UI_Lunch()
    frmLunch.Show
End Sub

Public Sub UI_Detail()
    frmDetail.Show
End Sub

Public Sub UI_Picker()
    frmPicker.Show
End Sub

Public Sub UI_Export()
    frmExport.Show
End Sub

Public Sub UI_Import()
    frmImport.Show
End Sub

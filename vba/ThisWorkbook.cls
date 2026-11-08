Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    AutoRefreshOnOpen
    Me.Saved = True
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If gAllowWorkbookSave Then
        gAllowWorkbookSave = False
        Exit Sub
    End If

    If Not SaveAsUI Then
        Cancel = True
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Сохранение самого файла Excel заблокировано, чтобы избежать конфликтов в Git!" & vbCrLf & vbCrLf & _
                     "Записать изменения в YAML-файлы («Сохранить в YAML») прямо сейчас?", _
                     vbYesNo + vbInformation, "Защита от конфликтов Git")
        If ans = vbYes Then
            Application.OnTime Now, "SaveSchemaToYaml"
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If gAllowWorkbookSave Then Exit Sub

    If Not Me.Saved Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("На листе есть несохранённые правки." & vbCrLf & vbCrLf & _
                     "Сохранить изменения в YAML-файлы перед выходом?" & vbCrLf & _
                     "(Сам файл Excel изменён не будет — защита от конфликтов в Git)", _
                     vbYesNoCancel + vbQuestion, "Выход из ODS YAML Editor")
        If ans = vbCancel Then
            Cancel = True
            Exit Sub
        ElseIf ans = vbYes Then
            SaveSchemaToYaml
            If Not Me.Saved Then
                Cancel = True
                Exit Sub
            End If
        Else
            Me.Saved = True
        End If
    End If
    On Error GoTo 0
End Sub

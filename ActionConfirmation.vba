Sub User_Confirm()
    If MsgBox("Confirm action?", vbYesNo + vbQuestion) = vbYes Then
        Call ActionSub
    Else:
        Exit Sub
    End If
End Sub

Sub Restrict_Action()
    Dim TargetRange as Range
    Set TargetRange = "A1:B2"

    If Not Application.Intersect(Selection, TargetRange) Is Nothing Then
        'Action
    End if
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Run action on activeworkbook save
End Sub

Private Sub Workbook_Open()
    'Run action on workbook open
End Sub

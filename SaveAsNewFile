Sub SaveFile()
  Worksheets(Array("Sheet 1", "Sheet 2")).Copy
  ActiveWorkbook.SaveAs Filename:= ActiveWorkbook.Path & "FileName" & Format(Now, "MMDDYYYY") & ".xlsx"
  Application.DisplayAlerts = False
  ActiveWorkbook.Close Savechanges:=True
End Sub

Sub SaveCSV()
  Worksheets("Sheet 1").Copy
  ActiveWorkbook.SaveAs ActiveWorkbook.Path & "FileName" & Format(Now, "MMDDYYYY") & ".csv", xlCSV
  Application.DisplayAlerts = False
  ActiveWorkbook.Close Savechanges:=True
End Sub

Sub SaveFileOverwrite()
  Worksheets("Sheet 1").Copy
  ActiveWorkbook.SaveAs Filename:="C:\Users\Glenn Dalida\Desktop\Payments\AMTrust\Amtrust Monthly Transactions.xlsx", _
    AccessMode:=xlExclusive, _
    ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
End Sub

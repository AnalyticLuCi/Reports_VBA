Sub Sort_Data()

Dim pt, Col, Row, Ans As String
Dim Rng As Range
Dim ColNo As Integer

On Error GoTo Bookmark2
    
Bookmark1:
    
    Set Rng = ActiveCell
    pt = Rng.PivotTable.Name
    ColNo = Rng.Column
    Col = Cells(11, ColNo).Value
    Row = Rng.PivotCell.RowItems(1).Parent.SourceName
    
    If Rng.Rows.Count > 1 Or Rng.Columns.Count > 1 Then
        MsgBox "Select only one cell."
        GoTo Bookmark1
    End If

    Ans = MsgBox("Sort in Descending order?", vbYesNo)

    If Ans = vbYes Then
        ActiveSheet.PivotTables(pt).PivotFields(Row).AutoSort _
            xlDescending, Col
    Else
        ActiveSheet.PivotTables(pt).PivotFields(Row).AutoSort _
            xlAscending, Col
    End If
    
Bookmark2:

End Sub

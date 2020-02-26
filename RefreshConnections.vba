Sub Refresh_Data()

Application.Calculation = xlManual

With ActiveWorkbook.
  Connections("Connection 1").Refresh
  Connections("Connection 2").Refresh
  Connections("Query - Query 1").Refresh
  Connections("Query - Query 2").Refresh
end with

ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh 'refreshes all PT of same source
ActiveSheet.PivotTables("PivotTable2").RefreshTable 'for specific PT only
Application.Calculation = xlAutomatic

End Sub

Attribute VB_Name = "PF___PivotTable"
Public Sub CreatePivotTable(sourceRange As Range, destinationRange As Range, pivotTblName As String)
'Purpose: To easily create a Pivot Table with minimal input from the user.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
Dim wb As Workbook
    Set wb = destinationRange.Worksheet.Parent
Dim rowNum As Long, columnNum As Long
    rowNum = sourceRange.Rows.Count
    columnNum = sourceRange.Columns.Count

With wb
    .PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sourceRange.Worksheet.Name & "!R1C1:R" & rowNum & "C" & columnNum, Version:=6). _
        CreatePivotTable TableDestination:=destinationRange, TableName:=pivotTblName, DefaultVersion:=6
End With
End Sub


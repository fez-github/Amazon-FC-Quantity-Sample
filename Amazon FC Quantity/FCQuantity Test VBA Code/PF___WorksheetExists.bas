Attribute VB_Name = "PF___WorksheetExists"
Public Function WorksheetExists(wb As Workbook, shtName As String) As Boolean
'Purpose: To see if a worksheet name is already used before adding a new one.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
Dim sht As Worksheet
For Each sht In wb.Sheets
    If sht.Name = shtName Then
        WorksheetExists = True
        Exit Function
    End If
Next sht
    WorksheetExists = False
End Function


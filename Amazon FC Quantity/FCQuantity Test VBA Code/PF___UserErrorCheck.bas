Attribute VB_Name = "PF___UserErrorCheck"
Public Sub UserErrorCheck(rng As Range, criteria As String, message As String)
'Purpose: Check for errors in a specified column range, and give a message when found.
'Written by: Mark Hansen
'Last Updated: July 26 2021
Dim i As Integer, colNum As Integer
colNum = rng.Column
With rng.Parent
    For i = 2 To .UsedRange.Rows.Count '.Range("B" & ws.rows.Count).End(xlUp).row
        If IsError(.Cells(i, colNum).Value) = True Then
            .Cells.AutoFilter field:=colNum, Criteria1:=criteria
            .Activate
            MsgBox message
                Exit Sub
        End If
    Next i
    MsgBox "No errors found."
End With
End Sub

Attribute VB_Name = "A_Initiate_FC_Quantity"
Option Explicit
Sub Full_Initiator()
'Purpose: To allow the user to select all workbooks needed, then automatically run the main sub multiple times.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
'Sheets Required:
'           [Active]Amazon Fulfillment Quantity Summary(use a blank sheet if you don't have one)
'           Daily Inventory History report(Amazon, one day back)
'           All Items & skuNames(Database)

'Variables
    Dim confirmBox As Variant, questionBox As Variant, arrayElement As Variant
    Dim filePicker As FileDialog
        Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    Dim Today As String, skuName As String, filepath As String, dailyInvt() As String
        Today = Left(Format(Now, "MM-DD-YY"), 9)
    Dim fcWB As Workbook, skuNameWB As Workbook, dailyInvtWB As Workbook
        Set fcWB = ActiveWorkbook
    Dim i As Integer
'Confirmation that you are on the proper workbook.
    confirmBox = MsgBox("Once activated, this macro cannot be undone.  Please ensure you meet the following conditions: " & vbCrLf & _
        "-No changes have been made to the downloaded reports." & vbCrLf & _
        "-You have a backup copy of the unedited downloaded reports in case something goes wrong.", vbYesNo, "Warning")
        If confirmBox <> vbYes Then
            Exit Sub
        End If
'Clear out all 0 records from FC sheet before continuing the rest of the process.
    ClearFC0 fcWB.Sheets("AmzRecords")
'Select All Items & skuNames workbook.
    filepath = FileSelect("Please select the All Items & skuName workbook.", "Select the All Items & skuName workbook.")
        If filepath = "" Then
            Exit Sub
        End If
    Set skuNameWB = Workbooks.Open(filepath)
'Select all Daily Inventory workbooks.
    dailyInvt() = MultiFileSelect( _
        dailyInvt(), "Please select all of the Daily Inventory workbooks,", "Select all of the Daily Inventory workbooks.")
'Run FC Quantity Data Cleaning sub on each workbook in the array.
    For Each arrayElement In dailyInvt
        Set dailyInvtWB = Workbooks.Open(arrayElement)
        FC_Quantity_Data_Cleaning fcWB, dailyInvtWB.Name, skuNameWB.Name
    Next arrayElement
'Save FC Quantity book, then save as new file for export.
    questionBox = MsgBox("Would you like to save this workbook?", vbYesNo)
    If questionBox = vbYes Then
        fcWB.Save
        fcWB.SaveAs filename:="Amazon Fulfillment Center Quantity " & Today, FileFormat:=xlCSVUTF8
    End If
End Sub
Sub ClearFC0(sht As Worksheet)
'Purpose: Delete all values with 0 in it.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
Dim i As Long, LastRow As Long
With sht
    If IsEmpty(.Range("A2")) = False Then
        LastRow = .Range("B" & Rows.Count).End(xlUp).Row
        For i = LastRow To 2 Step -1
            If .Cells(i, 5) = 0 Then
                .Rows(i).EntireRow.Delete
            End If
        Next i
    End If
End With
End Sub
Sub FC_Quantity_Data_Cleaning(fcWB As Workbook, dailyInvt As String, skuNameWB As String)
'Purpose: To clean Daily Inventory History sheet before importing.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
'Sheets Required:
'   [Active]Amazon Fulfillment Quantity Summary(use a blank sheet if you don't have one)
'           Daily Inventory History report(Amazon, one day back)
'           All Items & skuNames Cost Summary(Database)

'Declare variables
Dim skuNameSht As String, newestDate As String
    skuNameSht = Workbooks(skuNameWB).Sheets(1).Name
Dim i As Integer, LastRow As Integer
With Workbooks(dailyInvt).Sheets(1)
    'Add External ID
        If IsEmpty(.Range("I1")) = True Then
            AutoFillNewColumn .Range("I1"), "Record ID", "=CONCAT(C2,""-"",F2,""-"",G2)"
        End If
    'Add Name
        If IsEmpty(.Range("J1")) = True Then
            AutoFillNewColumn .Range("J1"), "Name", "=CONCAT(F2,""-"",C2,""-"",G2)"
        End If
    'Add Parent Item
        If IsEmpty(.Range("K1")) = True Then
            AutoFillNewColumn .Range("K1"), "Parent Name", _
            "=XLOOKUP(C2,'[" & skuNameWB & "]" & skuNameSht & "'!$A:$A,'[" & skuNameWB & "]" & skuNameSht & "'!$B:$B,,0)"
        End If
    'Copy data onto CurrentWorkbook sheet
       Paste_New_Range .UsedRange, fcWB.Sheets("AmzRecords").Range("A1").End(xlDown).Offset(1, 0), True
End With
With fcWB.Sheets("AmzRecords")
'Sort, then remove duplicates
    .UsedRange.Sort Key1:=.UsedRange, order1:=xlDescending, Header:=xlYes
    .UsedRange.RemoveDuplicates Columns:=Array(3, 6, 7), Header:=xlYes
'Turn all old quantity to 0
    newestDate = Left(.Range("A2").Value, 10)
    LastRow = .Range("B" & Rows.Count).End(xlUp).Row
    For i = 2 To LastRow
        If Left(.Cells(i, 1), 10) <> newestDate Then
            .Cells(i, 5).Value = 0
        End If
    Next i
End With
End Sub
Sub ParentItemErrors()
'Purpose: To show all Parent Name fields with an error.  Allows user to check for errors before exporting.
'         Meant to be called via Userform/button.
'Written by: Mark Hansen
'Last Updated: July 26 2021
    ThisWorkbook.Sheets("AmzRecords").Range ("K1"), "#N/A", _
        "Item without Parent Name detected!" & vbCrLf & _
        "Examine the item and resolve the error" & vbCrLf & _
        "Once you are done, resume the procedure."
End Sub
Sub FC_Quantity_Export()
'Purpose: To summarize Fulfillment Quantity into a Pivot Table and export it into the database.
'Written by: Mark Hansen
'Last Updated: July 26, 2021
'Sheets Needed:
'    [Active]Amazon Fulfillment Quantity Summary

'Declare variables
Dim confirmBox As Variant
Dim pivotSht As Worksheet
    If WorksheetExists(ThisWorkbook, "Pivot") = True Then
        Set pivotSht = Sheets("Pivot")
    Else
        Set pivotSht = ThisWorkbook.Sheets.Add
        pivotSht.Name = "Pivot"
    End If
Dim firstSht As String, Today As String
    firstSht = ThisWorkbook.Sheets("AmzRecords").Name
    Today = Left(Format(Now, "MM-DD-YY"), 9)
Dim LastRow As Long
With ThisWorkbook
    'Create Pivot Table, with Parent Item rows, and Sum of Quantity value.
    If pivotSht.PivotTables.Count = 1 Then
        pivotSht.PivotTables(1).Clear
    End If
    CreatePivotTable Sheets(firstSht).UsedRange, pivotSht.Range("A3"), "FCPivot"

    'Copy Pivot Table data onto new sheet.
    With .Sheets("Pivot").PivotTables("FCPivot")
        'Create pivot fields
            With .PivotFields("Parent Item Name")
                    .Orientation = xlRowField
                    .Position = 1
            End With
        'Add values
            .AddDataField .PivotFields("quantity"), "Sum of quantity", xlSum
        'Disable Grand Totals
            .ColumnGrand = False
    End With

    'Paste data into new sheet.
        If WorksheetExists(ThisWorkbook, "Import") = False Then
            .Sheets.Add(After:=.Sheets("Pivot")).Name = "Import"
        End If
        LastRow = .Sheets("Pivot").Range("A" & Rows.Count).End(xlUp).Row
        .Sheets("Pivot").Range("A4:B" & LastRow).Copy
        With .Sheets("Import")
            .Range("A1").Value = "Name"
            .Range("B1").Value = "Grand Total"
            .Range("A2").PasteSpecial
        End With
End With
'Export sheet as CSV.
    ThisWorkbook.SaveAs filename:="Amazon FC Export " & Today, FileFormat:=xlCSVUTF8
End Sub

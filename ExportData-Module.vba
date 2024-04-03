Sub ExportDataToMainWorkbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lastSourceRow As Long
    Dim lastTargetRow As Long

    Set sourceWorkbook = ThisWorkbook
    
    ' Adjust the path to the location of your main workbook
    Dim mainWorkbookPath As String
    mainWorkbookPath = "\\SERVER\PATH\TO\EXCEL\EXCEL-FILE.xlsm"

    ' Open the main workbook
    On Error Resume Next
    Set targetWorkbook = Workbooks.Open(mainWorkbookPath)
    If targetWorkbook Is Nothing Then
        MsgBox "Failed to open the Database Excel File."
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Assuming data to be copied is in Sheet1; adjust as necessary
    Set sourceSheet = sourceWorkbook.Sheets("Sheet1")
    Set targetSheet = targetWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in the source sheet
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "C").End(xlUp).Row
    
    ' Find the last row in the target sheet to append data
    lastTargetRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).Row + 1
    
    ' Check if there is data to copy
    If lastSourceRow >= 4 Then
        ' Copy data from source to target
        sourceSheet.Range("C4:O" & lastSourceRow).Copy _
            Destination:=targetSheet.Range("C" & lastTargetRow)
    Else
        MsgBox "No data to export."
    End If
    
    ' Save and close the main workbook
    targetWorkbook.Close SaveChanges:=True
    
    MsgBox "Data exported successfully."
End Sub




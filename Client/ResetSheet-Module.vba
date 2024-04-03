Sub ResetSheetDataRange()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Specify your worksheet name here
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in columns C to O. Adjust the column as needed for your data range.
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' If your data starts from row 4 and goes down to the last row with data
    If lastRow >= 4 Then
        ws.Range("C4:O" & lastRow).ClearContents
    End If
    
    MsgBox "Sheet has been Reset Successfully.", vbInformation
End Sub

Sub ClearOldEntriesAndCompact()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim checkDate As Date
    Dim monthsDifference As Integer
    Dim dataToDelete As Boolean

    ' Initialize flag to False
    dataToDelete = False

    ' Set the worksheet to work on
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet

    On Error GoTo ErrorHandler
    ' Find the last row with data in column M
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' If lastRow is less than 4, there's no data in the expected range, skip processing
    If lastRow < 4 Then
        MsgBox "Database is empty, Skipping function..", vbInformation
        Exit Sub
    End If

    ' Loop from the last row up to row 4
    For i = lastRow To 4 Step -1
        ' Check if the cell contains a date
        If IsDate(ws.Cells(i, "M").Value) Then
            checkDate = ws.Cells(i, "M").Value
            ' Calculate the difference in months
            monthsDifference = DateDiff("m", checkDate, Date)
            
            ' If the date is more than 4 months before today's date, set the flag
            If monthsDifference > 4 Then
                dataToDelete = True
                Exit For ' Exit the loop as we found data that meets the criteria
            End If
        End If
    Next i

    ' Check the flag before proceeding to delete
    If dataToDelete Then
        For i = lastRow To 4 Step -1
            If IsDate(ws.Cells(i, "M").Value) Then
                checkDate = ws.Cells(i, "M").Value
                monthsDifference = DateDiff("m", checkDate, Date)
                If monthsDifference > 4 Then
                    ws.Rows(i).EntireRow.Delete
                End If
            End If
        Next i
        MsgBox "Entries older than 4 Monthes have been cleared and data compacted for readability.", vbInformation
    Else
        MsgBox "All Data meets Cleanup Criteria, Skipping Function..", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred. Please contact NAME HERE -> NAME@EMAIL.COM", vbCritical
End Sub

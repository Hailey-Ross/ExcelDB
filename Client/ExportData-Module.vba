Sub HailsExportData()
    On Error GoTo ErrorHandler
    
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim sourceWb As Workbook, targetWb As Workbook
    Dim lastSourceRow As Long, lastTargetRow As Long, colCount As Long
    Dim sourceRow As Long, targetRow As Long, col As Long
    Dim cellMatch As Boolean, rowMatch As Boolean
    Dim updatesMade As Boolean

    ' Setup workbooks and worksheets
    Set sourceWb = ThisWorkbook
    Set sourceWs = sourceWb.Sheets("Sheet1") ' Adjust your source sheet name
    Set targetWb = Workbooks.Open("\\SERVER\PATH\TO\EXCEL.xlsm")
    Set targetWs = targetWb.Sheets("Sheet1") ' Adjust your target sheet name

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    lastSourceRow = sourceWs.Cells(sourceWs.Rows.Count, "C").End(xlUp).Row
    lastTargetRow = targetWs.Cells(targetWs.Rows.Count, "C").End(xlUp).Row
    colCount = 13 ' Columns C to O (inclusive)

    ' Check if there's data to process
    If lastSourceRow < 4 Then
        MsgBox "No new data was found, No changes were made to the Database", vbExclamation
        GoTo CleanExit
    End If

    For sourceRow = 4 To lastSourceRow
        rowMatch = False
        updatesMade = False

        For targetRow = 4 To lastTargetRow
            cellMatch = True ' Assume a match until proven otherwise

            ' Check each column in the current row for a match
            For col = 3 To colCount + 2 ' Adjust column indexes for C to O
                If sourceWs.Cells(sourceRow, col).Value <> targetWs.Cells(targetRow, col).Value Then
                    cellMatch = False ' Found a difference
                    Exit For
                End If
            Next col

            If cellMatch Then ' Full match, skip this row
                rowMatch = True
                Exit For
            Else ' Check for partial match to update differences
                For col = 3 To colCount + 2
                    If sourceWs.Cells(sourceRow, col).Value <> targetWs.Cells(targetRow, col).Value Then
                        ' Update only differing cells
                        targetWs.Cells(targetRow, col).Value = sourceWs.Cells(sourceRow, col).Value
                        updatesMade = True
                    End If
                Next col
            End If

            If updatesMade Then Exit For ' Exit after updating a partially matched row
        Next targetRow

        ' If no match at all, append the row
        If Not rowMatch And Not updatesMade Then
            lastTargetRow = lastTargetRow + 1
            sourceWs.Rows(sourceRow).Copy Destination:=targetWs.Rows(lastTargetRow)
        End If
    Next sourceRow

    targetWb.Close SaveChanges:=True

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred. Please Contact NAME HERE at NAME@EMAIL.COM", vbCritical
    Resume CleanExit
End Sub

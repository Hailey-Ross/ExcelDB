Sub ImportDataFromMainWorkbook()
    Dim mainWorkbook As Workbook
    Dim userWorkbook As Workbook
    Dim mainSheet As Worksheet
    Dim userSheet As Worksheet
    Dim lastRowMain As Long
    Dim lastRowUser As Long
    
    Set userWorkbook = ThisWorkbook
    ' Adjust "Sheet1" to your actual user sheet's name where data needs to be imported
    Set userSheet = userWorkbook.Sheets("Sheet1")

    ' Attempt to open the main workbook
    On Error Resume Next
    Set mainWorkbook = Workbooks.Open("\\SERVER\PATH\TO\EXCEL\EXCEL-FILE.xlsm")
    If mainWorkbook Is Nothing Then
        MsgBox "Failed to open the main workbook. Check your network connection or contact NAME HERE at NAME@EMAIL.com"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Adjust "Sheet1" to your actual main workbook's sheet name from which data is being imported
    Set mainSheet = mainWorkbook.Sheets("Sheet1")
    
    ' Determine the last row with data in the main workbook from the range C to O
    lastRowMain = mainSheet.Cells(mainSheet.Rows.Count, "C").End(xlUp).Row
    
    ' Determine the last row with data in the user's workbook from the range C to O to be cleared
    lastRowUser = userSheet.Cells(userSheet.Rows.Count, "C").End(xlUp).Row
    
    ' Clear existing data in the user's workbook in the specified range from C4 downwards
    If lastRowUser >= 4 Then
        userSheet.Range("C4:O" & lastRowUser).ClearContents
    End If

    ' Check if there's data to import from the main workbook
    If lastRowMain >= 4 Then
        ' Copy data from the main workbook to the user workbook
        mainSheet.Range("C4:O" & lastRowMain).Copy Destination:=userSheet.Range("C4")
    Else
        MsgBox "No data found to import."
    End If
    
    ' Close the main workbook without saving any changes
    mainWorkbook.Close SaveChanges:=False
    
    MsgBox "Data imported successfully."
End Sub


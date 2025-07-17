Sub SaveAllSheetsAsCSV()
    ' This macro saves every visible worksheet in the workbook as a separate CSV file.
    
    Dim ws As Worksheet
    Dim folderPath As String
    Dim fileName As String
    
    ' --- SET THE FOLDER WHERE CSVS WILL BE SAVED ---
    ' By default, this saves them in the same folder as the Excel file.
    ' To save to a different folder, uncomment the line below and change the path.
    ' folderPath = "C:\Users\YourName\Desktop\MyCSVs\"
    
    If folderPath = "" Then
        ' Check if the workbook has been saved before
        If ThisWorkbook.Path <> "" Then
            folderPath = ThisWorkbook.Path & Application.PathSeparator
        Else
            MsgBox "Please save your Excel file first before running this macro.", vbExclamation, "Save Required"
            Exit Sub
        End If
    End If
    
    ' --- MAIN PROCESS ---
    ' Turn off screen updating and alerts to speed up the process
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through each worksheet in the current workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is visible. Hidden sheets will be skipped.
        If ws.Visible = xlSheetVisible Then
            
            ' Create a unique file name based on the sheet name
            fileName = folderPath & ws.Name & ".csv"
            
            ' Copy the sheet to a new temporary workbook
            ' This is the safest way to save a single sheet without affecting the original
            ws.Copy
            
            ' Save the new temporary workbook as a CSV file
            ActiveWorkbook.SaveAs fileName:=fileName, FileFormat:=xlCSV, CreateBackup:=False
            
            ' Close the temporary workbook without saving any changes
            ActiveWorkbook.Close SaveChanges:=False
            
        End If
    Next ws
    
    ' --- CLEANUP ---
    ' Turn screen updating and alerts back on
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Inform the user that the process is complete
    MsgBox "All visible sheets have been successfully saved as separate CSV files in:" & vbNewLine & folderPath, vbInformation, "Process Complete"

End Sub

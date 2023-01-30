' this file does NOT contain any company confidential information
' this is a generic macro to convert xls files to xlsx


Dim InputXlsFile As String
Dim MyOpenedXlsFile As Workbook
Dim ConvertedXlsxlFile As String

'Use Application.ActiveWorkbook.Path for just the path itself (without the workbook name) or Application.ActiveWorkbook.FullName for the path with the workbook name.'
Dim myPath As String
Dim folderPath As String

folderPath = Application.ActiveWorkbook.Path
myPath = Application.ActiveWorkbook.FullName


'Looping through only xls files in input file folder
InputXlsFile = Dir(folderPath & "\*.xls")

While InputXlsFile <> ""

If Right(InputXlsFile, 4) <> "xlsm" Then
                
        Set MyOpenedXlsFile = Workbooks.Open(folderPath & "\" & InputXlsFile)
        ConvertedXlsxlFile = folderPath & "\" & Replace(ActiveWorkbook.Name, "xls", "xlsx")
            
        'Save each excel file as pdf file, the newly pdf file will be located where original excel file was located
        ActiveWorkbook.ActiveSheet.Name = "sheet1"
        ActiveWorkbook.SaveAs Filename:=ConvertedXlsxlFile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        MyOpenedXlsFile.Close
        
End If

        InputXlsFile = Dir
    


Wend



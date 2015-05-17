'Open Excel Files sing a fileopen dialog box and add them as a excel sheet next to "copied" sheet. 
'Name the new sheet as exported(if a sheet named exported is available, delete it)


Sub Import()
FileToOpen = Application.GetOpenFilename _
(Title:="Please Choose the File To Check", _
FileFilter:="Excel Files *.xlsx (*.xlsx),")

If FileToOpen = False Then
    MsgBox "No file specified.", vbExclamation, "No File Selected" '
    Exit Sub
Else
    With Application
    .ScreenUpdating = False
    End With
    Dim wsSheet As Worksheet
    On Error Resume Next
    Set wsSheet = Sheets("exported")
    On Error GoTo 0
    If Not wsSheet Is Nothing Then
        wsSheet.Delete
    End If
   Dim opened As Workbook
   Set opened = Workbooks.Open(Filename:=FileToOpen)
   opened.Worksheets(1).name = "exported"
   opened.Worksheets(1).Copy After:=ThisWorkbook.Worksheets("copied")
   With Application
    .ScreenUpdating = True
    opened.Close (False)
    End With
End If
End Sub
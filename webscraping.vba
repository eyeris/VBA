'Web data scraping vba code
'Reads the url links in "A1:A18" range and go to internet explorer and get the url addresses 

Private Sub CommandButton1_Click()
    Dim i As Long
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
    Dim selec As Range
    Dim cel As Range
    Set selec = Range("A1:A18")
 
    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    ' You can uncoment Next line To see form results
    IE.Visible = True
    
    For Each cel In selec.Cells
               ' Send the form data To URL As POST binary request
               IE.Navigate cel.Value
            
               ' Statusbar
               Application.StatusBar = "Page is loading. Please wait..."
            
               ' Wait while IE loading...
               Do While IE.Busy
                   Application.Wait (Now + TimeValue("0:00:05"))
               Loop
               
               Application.StatusBar = "Get THE URL. Please wait..."
               cel.Offset(0, 3).Value = IE.LocationURL
               'cel.Offset(0, 5).Value = IE.StatusText
               'MsgBox IE.LocationURL
    Next cel
    ' Clean up
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
 
    Application.StatusBar = ""
End Sub
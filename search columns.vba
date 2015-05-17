'Select a value in columns from a predefined range from other sheet. Use Sheet "main" keywords and search them in sheet "exported" column 9 until the end 'of data.Copy the data and paste into the columns "A", "B","C","D" and "E" Respectively

Sub MatchData()
    Dim selected As Range
    Dim i As Range
    Set selected = Worksheets("main").Range("A1").CurrentRegion
    For Each i In selected
        With Worksheets("exported").Columns(9)
            Set c = .Find(What:=i.Value, LookIn:=xlValues)
                If Not c Is Nothing Then
                firstAddress = c.Address
                     Do
                     If c.Offset(0, -7).Value <> c.Offset(0, -6).Value And c.Offset(0, -4).Value > 0 Then
                        c.Offset(0, -6).Copy
                        Worksheets("copied").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial
                        c.Offset(0, -2).Copy
                        Worksheets("copied").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).PasteSpecial
                        c.Copy
                        Worksheets("copied").Cells(Rows.Count, "C").End(xlUp).Offset(1, 0).PasteSpecial
                        c.Offset(0, 1).Copy
                        Worksheets("copied").Cells(Rows.Count, "D").End(xlUp).Offset(1, 0).PasteSpecial
                        c.Offset(0, 2).Copy
                        Worksheets("copied").Cells(Rows.Count, "E").End(xlUp).Offset(1, 0).PasteSpecial
                     End If
                     Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
        End With
    Next i
End Sub
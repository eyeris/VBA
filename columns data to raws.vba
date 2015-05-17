'column sequence data to raws of another sheet
'This function search a column for word "machine" and when it find the word, the follwoing values are copied to raw format


Sub InsertData()
        With Worksheets("Input data here").Columns(1)
            Set c = .Find(What:="Employee", LookIn:=xlValues)
                If Not c Is Nothing Then
                firstAddress = c.Address
                     Do
                     If c.Offset(2, 0).Value = "Machine" And c.Offset(4, 0).Value = "Product" And c.Offset(6, 0).Value = "Operation" And c.Offset(8, 0).Value = "Hours" And c.Offset(10, 0).Value = "QTY" And c.Offset(12, 0).Value = "Scraps" Then
                        Worksheets("Final Data").Cells(Rows.count, "A").End(xlUp).Offset(1, 0).Value = c.Offset(1, 0).Value
                        Worksheets("Final Data").Cells(Rows.count, "B").End(xlUp).Offset(1, 0).Value = c.Offset(3, 0).Value
                        Worksheets("Final Data").Cells(Rows.count, "C").End(xlUp).Offset(1, 0).Value = StringCat1(c.Offset(5, 0).Value)
                        Worksheets("Final Data").Cells(Rows.count, "D").End(xlUp).Offset(1, 0).Value = StringCat2(c.Offset(5, 0).Value)
                     Else
                        c.Interior.ColorIndex = 3
                     End If
                     Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
        End With
End Sub
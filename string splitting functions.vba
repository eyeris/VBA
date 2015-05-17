'String Splitting Functions
'"U.TBN-43-R32-9 Button 10427.U"==>"TBN-43-R32-9 Button" and "10427"

Sub Test()
     Dim i, j, k As String
     i = Split(Worksheets("test").Range("A1").Value, " ")(0)
     j = Split(Worksheets("test").Range("A1").Value, " ")(1)
     k = Split(Worksheets("test").Range("A1").Value, " ")(2)
     Worksheets("test").Range("A2").Value = Split(i, ".")(1) & " " & j
     Worksheets("test").Range("A3").Value = Split(k, ".")(0)
End Sub

'"U.TBN-43-R32-9 Button 10427.U"==>"TBN-43-R32-9 Button" 
Function StringCat1(word As String) As String
     Dim i, j As String
     i = Split(word, " ")(0)
     j = Split(word, " ")(1)
     StringCat1 = Split(i, ".")(1) & " " & j
End Function

"U.TBN-43-R32-9 Button 10427.U"==>"10427"
Function StringCat2(word As String) As String
     Dim k As String
     k = Split(word, " ")(2)
     StringCat2 = Split(k, ".")(0)
End Function
Public Sub SALVAR()
Range("A7:E7").Select
Selection.Copy
Sheets(" Matriz Base").Select
    Range("A1048576").End(xlUp).Offset(1, 0).Select
Selection.PasteSpecial Paste:=xlPasteValues

End Sub

Public Sub SALVAR()

    Range("A7:f7").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("A1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Sheets("ES Forms").Select
    Range("A12:d12").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("G1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("A14:f14").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("k1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("A17:f17").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("q1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Sheets("ES Forms").Select
    Range("A19:f19").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("w1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues    

    Sheets("ES Forms").Select
    Range("A21:f21").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ac1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues    

    Sheets("ES Forms").Select
    Range("A23:f23").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ai1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues  

    Sheets("ES Forms").Select
    Range("A28:f28").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("A29:f29").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("A30:f30").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a31:f31").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a32:f32").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a33:f33").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a34:f34").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' hack
    Range("a1048576").End(xlUp).Select    
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 39).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("a1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a35:f35").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("ao1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    Sheets("ES Forms").Select
    Range("a39:b39").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("au1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Sheets("ES Forms").Select
    Range("f39").Select
    Selection.Copy
    Sheets(" Matriz Base").Select
    Range("aw1048576").End(xlUp).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' Local Endereço Complemento
    Range("au1048576").End(xlUp).Offset(0, 0).Select
    StartCell = ActiveCell.Address
    EndCell = ActiveCell.Offset(0, 3).Address
    Range(StartCell, EndCell).Select
    Selection.Copy
    Range("au1048576").End(xlUp).Offset(1, 0).Select    
    Selection.PasteSpecial Paste:=xlPasteValues    

    Dim i, Rows As Integer
    Rows = 5
    For i=0 To Rows
        Range("au1048576").End(xlUp).Offset(1, 0).Select    
        Selection.PasteSpecial Paste:=xlPasteValues    
    Next i

End Sub


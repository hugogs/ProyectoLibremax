Attribute VB_Name = "AcumMovCompras1"
Sub ActualizarCompras()
    Call ReplicaFormulasAcumCompra
    Call BorraFormulasMovCompras
    Call ReplicaFormulasMovCompras
End Sub
Sub ReplicaFormulasAcumCompra()
    Sheets("Acum-Compra").Select
    Dim NroFilaCompras As Integer
    Dim NroColumnaCompras As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaCompras = ActiveCell.Row
    NroColumnaCompras = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range("P2:R2").Select
    Selection.Copy
    Range(Cells(3, 16), Cells(NroFilaCompras, NroColumnaCompras + 17)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
End Sub
Sub BorraFormulasMovCompras()
    Sheets("Mov.COMPRAS").Select
    Range("D4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("D4").Select
End Sub
Sub ReplicaFormulasMovCompras()
    Sheets("Mov.COMPRAS").Select
    Dim NroFilaMovCom As Integer
    Dim NroColumnaMovCom As Integer
    Range("A3").Select
    Range("A3").End(xlDown).Select
    NroFilaMovCom = ActiveCell.Row
    NroColumnaMovCom = ActiveCell.Column
    'Selecciona el rango de formulas a replicar
    Range("D3:AH3").Select
    Selection.Copy
    Range(Cells(4, 4), Cells(NroFilaMovCom, NroColumnaMovCom + 33)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    'Selecciona el formato de celdas a replicar
    Range("D3:AH3").Select
    Selection.Copy
    Range(Cells(4, 4), Cells(NroFilaMovCom, NroColumnaMovCom + 33)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
End Sub

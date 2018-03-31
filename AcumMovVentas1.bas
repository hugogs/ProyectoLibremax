Attribute VB_Name = "AcumMovVentas1"
Sub ActualizarVentas()
    Call ReplicaFormulasAcumVentas
    Call BorraFormulasMovVentas
    Call ReplicaFormulasMovVentas
End Sub
Sub ReplicaFormulasAcumVentas()
    Sheets("Acum-VENTAS").Select
    Dim NroFilaVentas As Integer
    Dim NroColumnaVentas As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaVentas = ActiveCell.Row
    NroColumnaVentas = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range("K2:L2").Select
    Selection.Copy
    Range(Cells(3, 11), Cells(NroFilaVentas, NroColumnaVentas + 11)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
End Sub
Sub BorraFormulasMovVentas()
    Sheets("Mov.VENTAS").Select
    Range("D5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("D5").Select
End Sub
Sub ReplicaFormulasMovVentas()
    Sheets("Mov.VENTAS").Select
    Dim NroFilaMovVta As Integer
    Dim NroColumnaMovVta As Integer
    Range("A4").Select
    Range("A4").End(xlDown).Select
    NroFilaMovVta = ActiveCell.Row
    NroColumnaMovVta = ActiveCell.Column
    'Selecciona el rango de formulas a replicar
    Range("D4:BX4").Select
    Selection.Copy
    Range(Cells(5, 4), Cells(NroFilaMovVta, NroColumnaMovVta + 75)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    'Selecciona el formato de celdas a replicar
    Range("D4:BX4").Select
    Selection.Copy
    Range(Cells(5, 4), Cells(NroFilaMovVta, NroColumnaMovVta + 75)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A2").Select
    Application.CutCopyMode = False
End Sub

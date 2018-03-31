Attribute VB_Name = "Libremax1"
Sub ActualizarInventario()
    Call BorrarDatosInventario
    Call DatosInventario
    Call ReplicaFormulas
    Call BorraInfoCodigosVentas
    Call InfoCodigosVentas
    Call BorraInfoCodigosCompras
    Call InfoCodigosCompras
    ActiveWorkbook.Save
End Sub
Sub BorrarDatosInventario()
    Sheets("INVENTARIO").Select
    Dim NroFilaDatos As Integer
    Dim NroColumnaDatos As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaDatos = ActiveCell.Row
    NroColumnaDatos = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range(Cells(2, 1), Cells(NroFilaDatos, NroColumnaDatos + 2)).Select
    Selection.ClearContents
    Range(Cells(3, 5), Cells(NroFilaDatos, NroColumnaDatos + 8)).Select
    Selection.ClearContents
    Range("A2").Select
End Sub
Sub DatosInventario()
    Sheets("Base").Select
    Dim NroFila As Integer
    Dim NroColumna As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Selecciona el rango a copiar de CODIGOS
    Range(Cells(2, 1), Cells(NroFila, NroColumna)).Select
    Selection.Copy
    Sheets("INVENTARIO").Select
    Range("A2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar de DESCRIPCION
    Sheets("Base").Select
    Range(Cells(2, 3), Cells(NroFila, NroColumna + 2)).Select
    Selection.Copy
    Sheets("INVENTARIO").Select
    Range("B2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar de MARCA
    Sheets("Base").Select
    Range(Cells(2, 6), Cells(NroFila, NroColumna + 5)).Select
    Selection.Copy
    Sheets("INVENTARIO").Select
    Range("C2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets("Base").Select
    Range("A1").Select
End Sub
Sub ReplicaFormulas()
    Sheets("INVENTARIO").Select
    Dim NroFilaForm As Integer
    Dim NroColumnaForm As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaForm = ActiveCell.Row
    NroColumnaForm = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range("E2:J2").Select
    Selection.Copy
    Range(Cells(3, 5), Cells(NroFilaForm, NroColumnaForm + 9)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
End Sub
Sub BorraInfoCodigosVentas()
    Application.Sheets("InfoParaVentas").Visible = True
    Sheets("InfoParaVentas").Select
    Range("A2").Select
    Dim NroFilaDatos As Integer
    Dim NroColumnaDatos As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaDatos = ActiveCell.Row
    NroColumnaDatos = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range(Cells(2, 1), Cells(NroFilaDatos, NroColumnaDatos + 4)).Select
    Selection.ClearContents
    Range("A2").Select
End Sub
Sub InfoCodigosVentas()
    Sheets("BASE").Select
    Dim NroFilaInfoVta As Integer
    Dim NroColumnaInfoVta As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaInfoVta = ActiveCell.Row
    NroColumnaInfoVta = ActiveCell.Column
    'Selecciona el rango a copiar Columna A,B Y C
    Range(Cells(2, 1), Cells(NroFilaInfoVta, NroColumnaInfoVta + 2)).Select
    Selection.Copy
    Sheets("InfoParaVentas").Select
    Range("A2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar Columna E
    Sheets("BASE").Select
    Range(Cells(2, 5), Cells(NroFilaInfoVta, NroColumnaInfoVta + 4)).Select
    Selection.Copy
    Sheets("InfoParaVentas").Select
    Range("D2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar Columna A
    Sheets("BASE").Select
    Range(Cells(2, 1), Cells(NroFilaInfoVta, NroColumnaInfoVta)).Select
    Selection.Copy
    Sheets("InfoParaVentas").Select
    Range("E2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets("BASE").Select
    Range("A1").Select
    Application.Sheets("InfoParaVentas").Visible = False
End Sub
Sub BorraInfoCodigosCompras()
    Application.Sheets("InfoParaCompras").Visible = True
    Sheets("InfoParaCompras").Select
    Range("A2").Select
    Dim NroFilaDatos As Integer
    Dim NroColumnaDatos As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaDatos = ActiveCell.Row
    NroColumnaDatos = ActiveCell.Column
    'Selecciona el rango de datos a borrar
    Range(Cells(2, 1), Cells(NroFilaDatos, NroColumnaDatos + 5)).Select
    Selection.ClearContents
    Range("A2").Select
End Sub
Sub InfoCodigosCompras()
    Sheets("BASE").Select
    Dim NroFilaInfoVta As Integer
    Dim NroColumnaInfoVta As Integer
    Range("A2").Select
    Range("A2").End(xlDown).Select
    NroFilaInfoVta = ActiveCell.Row
    NroColumnaInfoVta = ActiveCell.Column
    'Selecciona el rango a copiar Columna A,B Y C
    Range(Cells(2, 1), Cells(NroFilaInfoVta, NroColumnaInfoVta + 2)).Select
    Selection.Copy
    Sheets("InfoParaCompras").Select
    Range("A2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar Columna E Y F
    Sheets("BASE").Select
    Range(Cells(2, 5), Cells(NroFilaInfoVta, NroColumnaInfoVta + 5)).Select
    Selection.Copy
    Sheets("InfoParaCompras").Select
    Range("D2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Selecciona el rango a copiar Columna A
    Sheets("BASE").Select
    Range(Cells(2, 1), Cells(NroFilaInfoVta, NroColumnaInfoVta)).Select
    Selection.Copy
    Sheets("InfoParaCompras").Select
    Range("F2").Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets("BASE").Select
    Range("A1").Select
    Application.Sheets("InfoParaCompras").Visible = False
End Sub


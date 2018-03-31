Attribute VB_Name = "CompraLibremax1"
Sub Consulta()
    Sheets("CONSULTA").Select
    Range("D7").Select
End Sub
Sub Venta()
    Sheets("CONSULTA").Select
    Range("D7").Select
    Selection.ClearContents
    Sheets("COMPRA").Select
    Range("D21").Select
End Sub
Sub Grabar()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call Guarda_Datos
    Call Borrar_Codigos
    Call Reiniciar_Valores
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Sub Cancelar()
    Call Borrar_Codigos
    Call Reiniciar_Valores
    Application.Sheets("Datos").Visible = False
End Sub
Sub Borrar_Codigos()
'Macro para borrar codigos
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Sheets("COMPRA").Select
    Range("E7,E9,E11").Select
    Selection.ClearContents
    Range("D21,D23,D25,D27,D29,D31,D33,D35,D37,D39,D41,D43,D45,D47,D49,D51,D53,D55,D57,D59,D61,D63,D65,D67,D69,D71,D73,D75,D77,D79").Select
    Selection.ClearContents
    Range("L21,L23,L25,L27,L29,L31,L33,L35,L37,L39,L41,L43,L45,L47,L49,L51,L53,L55,L57,L59,L61,L63,L65,L67,L69,L71,L73,L75,L77,L79").Select
    Selection.ClearContents
    Range("D79").Select
    Range("D21").Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
End Sub
Sub Guarda_Datos()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Application.Sheets("Datos").Visible = True
    Sheets("Datos").Select
    'Obtiene cantidad de filas a copiar
    Range("AO1").Select
    datoFilas = ActiveCell
    'Selecciona el rango a copiar
    Range(Cells(2, 27), Cells(datoFilas + 1, 40)).Select
    Selection.Copy
   'Declaro variables
    Dim NroFilaVenta As Integer
    Dim NroColumnaVenta As Integer
    'Selecciona la ultima celda(fila) con datos
    Range("A1").End(xlDown).Select
    'Obtengo fila y columna
    NroFilaVenta = ActiveCell.Row
    NroColumnaVenta = ActiveCell.Column
    'Selecciona la 1ra celda disponible a copiar(Parte inferior)
    Cells(NroFilaVenta + 1, NroColumnaVenta).Select
    'Pego como valores los datos copiados
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Call Agrega06Digitos
    
    Application.Sheets("Datos").Visible = False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
End Sub
Sub Reiniciar_Valores()
    valor = 1
    Sheets("COMPRA").Select
    Dim r As Integer
    For r = 0 To 29
        Cells(21 + (2 * r), 16) = valor
    Next r
End Sub
Sub Agrega06Digitos()
    Sheets("Datos").Select
    Range("N1").Select
    '=SI(LARGO(F2)=13,BUSCARV(F2,Detalle!B:F,5,0),F2)
    'Obtengo la fila y columna FIN
    Range("N1").End(xlDown).Select
    Dim NroFilaFin As Integer
    Dim NroColumnaFin As Integer
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'Copio formula de R11
    Range("O2").Select
    Selection.Copy
    'Pego la formula en todo el rango seleccionado
    Range(Cells(3, 15), Cells(NroFilaFin, NroColumnaFin + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    'A valores en todo el rango seleccionado
    'Range(Cells(3, 15), Cells(NroFilaFin, NroColumnaFin + 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
End Sub

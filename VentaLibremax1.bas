Attribute VB_Name = "VentaLibremax1"
Sub Consulta()
    Sheets("CONSULTA").Select
    Range("D7").Select
End Sub
Sub Venta()
    Sheets("CONSULTA").Select
    Range("D7").Select
    Selection.ClearContents
    Sheets("VENTA").Select
    Range("E21").Select
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
    
    Sheets("VENTA").Select
    Range("E21,E23,E25,E27,E29,E31,E33,E35,E37,E39,E41,E43,E45,E47,E49,E51,E53,E55,E57,E59").Select
    Selection.ClearContents
    Range("E59").Select
    Range("E21").Select
    
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
    Range("AJ1").Select
    datoFilas = ActiveCell
    'Selecciona el rango a copiar
    Range(Cells(2, 27), Cells(datoFilas + 1, 35)).Select
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
    Sheets("VENTA").Select
    Dim r As Integer
    For r = 0 To 19
        Cells(21 + (2 * r), 18) = valor
    Next r
End Sub
Sub Agrega06Digitos()
    Sheets("Datos").Select
    Range("I1").Select
    '=SI(LARGO(F2)=13,BUSCARV(F2,Detalle!B:F,5,0),F2)
    'Obtengo la fila y columna FIN
    Range("I1").End(xlDown).Select
    Dim NroFilaFin As Integer
    Dim NroColumnaFin As Integer
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'Copio formula de R11
    Range("J2").Select
    Selection.Copy
    'Pego la formula en todo el rango seleccionado
    Range(Cells(3, 10), Cells(NroFilaFin, NroColumnaFin + 1)).Select
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

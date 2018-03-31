Attribute VB_Name = "AcumMovCompras2"
Sub MovimientosCompras()
 
'Definir objetos a utilizar
Dim wbDestino As Workbook, _
    wsOrigen As Excel.Worksheet, _
    wsDestino As Excel.Worksheet, _
    rngOrigen As Excel.Range, _
    rngDestino As Excel.Range
     
'Indicar el libro de Excel destino
Set wbDestino = Workbooks.Open(ActiveWorkbook.Path & "\LIBREMAX V3.0.xlsm")
 
'Activar este libro
ThisWorkbook.Activate
 
'Indicar las hojas de origen y destino
Set wsOrigen = Worksheets("Acum-Compra")
Set wsDestino = wbDestino.Worksheets("COMPRAS")
 
'Indicar la celda de origen y destino
Const celdaOrigen = "A2"
Const celdaDestino = "A2"
 
'Inicializar los rangos de origen y destino
Set rngOrigen = wsOrigen.Range(celdaOrigen)
Set rngDestino = wsDestino.Range(celdaDestino)
 
'Seleccionar rango de celdas origen
rngOrigen.Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
 
'Pegar datos en celda destino
rngDestino.PasteSpecial xlPasteValues
Application.CutCopyMode = False
 
'Guardar y cerrar el libro de Excel destino
wbDestino.Save
wbDestino.Close
 
End Sub

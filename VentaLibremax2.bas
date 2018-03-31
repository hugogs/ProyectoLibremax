Attribute VB_Name = "VentaLibremax2"
Sub CopiarVentasBD()
 
'Definir objetos a utilizar
Dim wbDestino As Workbook, _
    wsOrigen As Excel.Worksheet, _
    wsDestino As Excel.Worksheet, _
    rngOrigen As Excel.Range, _
    rngDestino As Excel.Range
     
'Indicar el libro de Excel destino
Set wbDestino = Workbooks.Open(ActiveWorkbook.Path & "\ACUM - MOV VENTAS V3.0.xlsm")
 
'Activar este libro
ThisWorkbook.Activate
 
'Indicar las hojas de origen y destino
Set wsOrigen = Worksheets("Datos")
Set wsDestino = wbDestino.Worksheets("Acum-VENTAS")
 
'Indicar la celda de origen y destino
Const celdaOrigen = "A3"
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


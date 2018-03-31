Attribute VB_Name = "Libremax2"
Sub AllComprasVentas()
    Call TodoCompras
    Call TodoVentas
End Sub
Sub TodoCompras()
    Application.Sheets("InfoParaCompras").Visible = True
    Call InfoParaCompras
    Call InfoAcumMovCompras
    Application.Sheets("InfoParaCompras").Visible = False
End Sub
Sub InfoParaCompras()
    Sheets("InfoParaCompras").Select
    Workbooks.Open Filename:="C:\Users\Leito\Desktop\COMPRA_Libremax V3.0.xlsm"
    Sheets("Detalle").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    Windows("LIBREMAX V3.0.xlsm").Activate
    Sheets("InfoParaCompras").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("COMPRA_Libremax V3.0.xlsm").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("COMPRA").Select
    Range("D21").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Sheets("BASE").Select
    Range("A1").Select
End Sub
Sub InfoAcumMovCompras()
    Sheets("INVENTARIO").Select
    Workbooks.Open Filename:="C:\Users\Leito\Desktop\ACUM - MOV COMPRAS V3.0.xlsm"
    Sheets("Mov.COMPRAS").Select
    Range("A3").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    Range("A3:C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Application.Run "'ACUM - MOV COMPRAS V3.0.xlsm'!BorraFormulasMovCompras"
    Range("A3").Select
    Windows("LIBREMAX V3.0.xlsm").Activate
    Sheets("INVENTARIO").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("ACUM - MOV COMPRAS V3.0.xlsm").Activate
    Range("A3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Acum-Compra").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Sheets("BASE").Select
    Range("A1").Select
End Sub
Sub TodoVentas()
    Application.Sheets("InfoParaVentas").Visible = True
    Call InfoParaVentas
    Call InfoAcumMovVentas
    Application.Sheets("InfoParaVentas").Visible = False
End Sub
Sub InfoParaVentas()
    Sheets("InfoParaVentas").Select
    Workbooks.Open Filename:="C:\Users\Leito\Desktop\VENTA_Libremax V3.0.xlsm"
    Sheets("Detalle").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    Windows("LIBREMAX V3.0.xlsm").Activate
    Sheets("InfoParaVentas").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VENTA_Libremax V3.0.xlsm").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("VENTA").Select
    Range("E21").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Sheets("BASE").Select
    Range("A1").Select
End Sub
Sub InfoAcumMovVentas()
    Sheets("INVENTARIO").Select
    Workbooks.Open Filename:="C:\Users\Leito\Desktop\ACUM - MOV VENTAS V3.0.xlsm"
    Sheets("Mov.VENTAS").Select
    Range("A4").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    Range("A4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Application.Run "'ACUM - MOV VENTAS V3.0.xlsm'!BorraFormulasMovVentas"
    Range("A4").Select
    Windows("LIBREMAX V3.0.xlsm").Activate
    Sheets("INVENTARIO").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("ACUM - MOV VENTAS V3.0.xlsm").Activate
    Range("A4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Acum-VENTAS").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Sheets("BASE").Select
    Range("A1").Select
End Sub

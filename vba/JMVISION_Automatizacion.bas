Attribute VB_Name = "JMVISION_Automatizacion"
Option Explicit

Public Sub SetupJMVisionWorkbook()
    Dim wsConfig As Worksheet, wsClientes As Worksheet, wsProductos As Worksheet
    Dim wsCot As Worksheet, wsHist As Worksheet, wsVentas As Worksheet, wsDash As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsConfig = EnsureSheet("CONFIG")
    Set wsClientes = EnsureSheet("CLIENTES")
    Set wsProductos = EnsureSheet("PRODUCTOS")
    Set wsCot = EnsureSheet("COTIZACION")
    Set wsHist = EnsureSheet("HISTORICO_COTIZACIONES")
    Set wsVentas = EnsureSheet("VENTAS")
    Set wsDash = EnsureSheet("DASHBOARD")

    SetupConfig wsConfig
    SetupClientes wsClientes
    SetupProductos wsProductos
    SetupCotizacion wsCot
    SetupHistorico wsHist
    SetupVentas wsVentas
    SetupDashboard wsDash
    SetupValidations wsCot, wsHist
    CreateActionButtons wsCot

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Configuración completada. Ahora carga clientes y productos y empieza la prueba.", vbInformation
End Sub

Public Sub GuardarCotizacionEnHistorico()
    Dim wsCot As Worksheet, wsHist As Worksheet
    Dim nextRow As Long
    Dim subtotal As Double, iva As Double, total As Double

    Set wsCot = ThisWorkbook.Worksheets("COTIZACION")
    Set wsHist = ThisWorkbook.Worksheets("HISTORICO_COTIZACIONES")

    If Trim(wsCot.Range("B2").Value) = "" Then
        MsgBox "No hay número de cotización.", vbExclamation
        Exit Sub
    End If

    subtotal = Application.WorksheetFunction.Sum(wsCot.Range("F12:F200"))
    iva = Application.WorksheetFunction.Sum(wsCot.Range("G12:G200"))
    total = Application.WorksheetFunction.Sum(wsCot.Range("H12:H200"))

    nextRow = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    wsHist.Cells(nextRow, "A").Value = wsCot.Range("B2").Value
    wsHist.Cells(nextRow, "B").Value = wsCot.Range("B3").Value
    wsHist.Cells(nextRow, "C").Value = wsCot.Range("B4").Value
    wsHist.Cells(nextRow, "D").Value = wsCot.Range("B5").Value
    wsHist.Cells(nextRow, "E").Value = subtotal
    wsHist.Cells(nextRow, "F").Value = iva
    wsHist.Cells(nextRow, "G").Value = total
    wsHist.Cells(nextRow, "H").Value = "Enviada"
    wsHist.Cells(nextRow, "K").Value = "Generada desde macro"

    MsgBox "Cotización guardada en histórico.", vbInformation
End Sub

Public Sub ConvertirCotizacionAVenta()
    Dim wsCot As Worksheet, wsVentas As Worksheet, wsHist As Worksheet
    Dim nextRow As Long, total As Double, cotizacion As String
    Dim estado As Variant

    Set wsCot = ThisWorkbook.Worksheets("COTIZACION")
    Set wsVentas = ThisWorkbook.Worksheets("VENTAS")
    Set wsHist = ThisWorkbook.Worksheets("HISTORICO_COTIZACIONES")

    cotizacion = Trim(wsCot.Range("B2").Value)
    If cotizacion = "" Then
        MsgBox "No hay cotización activa.", vbExclamation
        Exit Sub
    End If

    estado = Application.WorksheetFunction.XLookup(cotizacion, wsHist.Range("A:A"), wsHist.Range("H:H"), "NO_ENCONTRADA")
    If CStr(estado) <> "Aprobada" Then
        MsgBox "Para convertir a venta, primero marca la cotización como Aprobada en el histórico.", vbExclamation
        Exit Sub
    End If

    total = Application.WorksheetFunction.Sum(wsCot.Range("H12:H200"))

    nextRow = wsVentas.Cells(wsVentas.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    wsVentas.Cells(nextRow, "A").FormulaLocal = "=CONFIG!B10&""-""&AÑO(HOY())&""-""&TEXTO(CONTARA(VENTAS!A:A),""0000"")"
    wsVentas.Cells(nextRow, "B").Value = Date
    wsVentas.Cells(nextRow, "C").Value = cotizacion
    wsVentas.Cells(nextRow, "D").Value = wsCot.Range("B5").Value
    wsVentas.Cells(nextRow, "E").Value = total
    wsVentas.Cells(nextRow, "F").Value = "Transferencia"
    wsVentas.Cells(nextRow, "G").Value = "Pendiente"

    MsgBox "Cotización convertida a venta.", vbInformation
End Sub

Public Sub LimpiarCotizacion()
    Dim wsCot As Worksheet
    Set wsCot = ThisWorkbook.Worksheets("COTIZACION")

    wsCot.Range("B4").ClearContents
    wsCot.Range("B8").ClearContents
    wsCot.Range("J4").ClearContents
    wsCot.Range("A12:L200").ClearContents
    wsCot.Range("B3").Value = Date

    MsgBox "Formulario de cotización limpio.", vbInformation
End Sub

Private Function EnsureSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    Else
        EnsureSheet.Cells.Clear
    End If
End Function

Private Sub SetupConfig(ws As Worksheet)
    ws.Range("A1").Value = "Campo"
    ws.Range("B1").Value = "Valor"

    ws.Range("A2:A10").Value = Application.WorksheetFunction.Transpose(Array( _
        "Empresa", "NIT/ID fiscal", "Teléfono", "Correo ventas", "Moneda", "IVA estándar", "Días vigencia cotización", "Prefijo cotización", "Prefijo venta"))

    ws.Range("B2").Value = "JMVISION"
    ws.Range("B3").Value = "900123456"
    ws.Range("B4").Value = "+57 300 000 0000"
    ws.Range("B5").Value = "ventas@jmvision.com"
    ws.Range("B6").Value = "COP"
    ws.Range("B7").Value = 0.19
    ws.Range("B8").Value = 15
    ws.Range("B9").Value = "COT-JMV"
    ws.Range("B10").Value = "VTA-JMV"

    ws.Columns("A:B").AutoFit
    ws.Range("B7").NumberFormat = "0%"
End Sub

Private Sub SetupClientes(ws As Worksheet)
    ws.Range("A1:K1").Value = Array("ID_Cliente", "Tipo_Cliente", "Nombre_Cliente", "Documento/NIT", "Contacto", "Teléfono", "Correo", "Ciudad", "Dirección", "Canal_Origen", "Estado")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:K").AutoFit
End Sub

Private Sub SetupProductos(ws As Worksheet)
    ws.Range("A1:J1").Value = Array("SKU", "Producto_Servicio", "Categoría", "Unidad", "Precio_Lista", "Costo_Referencial", "Margen_%_Objetivo", "IVA_Aplica", "Stock", "Estado")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:J").AutoFit
    ws.Columns("E:F").NumberFormat = "$ #,##0.00"
    ws.Columns("G").NumberFormat = "0%"
End Sub

Private Sub SetupCotizacion(ws As Worksheet)
    ws.Range("A1").Value = "COTIZACIÓN"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A2:A9").Value = Application.WorksheetFunction.Transpose(Array("No_Cotización", "Fecha", "ID_Cliente", "Cliente", "Correo", "Vigencia hasta", "Asesor", "Moneda"))

    ws.Range("B2").FormulaLocal = "=CONFIG!B9&""-""&AÑO(HOY())&""-""&TEXTO(CONTARA(HISTORICO_COTIZACIONES!A:A)+1,""0000"")"
    ws.Range("B3").Value = Date
    ws.Range("B5").FormulaLocal = "=SI.ERROR(BUSCARX(B4,CLIENTES!A:A,CLIENTES!C:C),""Cliente no encontrado"")"
    ws.Range("B6").FormulaLocal = "=SI.ERROR(BUSCARX(B4,CLIENTES!A:A,CLIENTES!G:G),"""")"
    ws.Range("B7").FormulaLocal = "=B3+CONFIG!B8"
    ws.Range("B9").FormulaLocal = "=CONFIG!B6"

    ws.Range("J3").Value = "Buscar producto"
    ws.Range("J4").Value = "camaras"
    ws.Range("J5").Value = "SKU sugerido"
    ws.Range("J6").FormulaLocal = "=SI.ERROR(INDICE(FILTRAR(PRODUCTOS!A:A,ESNUMERO(HALLAR(MINUSC($J$4),MINUSC(PRODUCTOS!B:B)))),1),""SIN RESULTADO"")"

    ws.Range("A11:L11").Value = Array("SKU", "Descripción", "Cantidad", "Precio Unitario", "Desc_%", "Subtotal", "IVA", "Total Línea", "Categoría", "Unidad", "Stock", "Estado")
    ws.Rows(11).Font.Bold = True

    ws.Range("B12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!B:B),""")"
    ws.Range("D12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!E:E),0)"
    ws.Range("F12").FormulaLocal = "=C12*D12*(1-E12)"
    ws.Range("G12").FormulaLocal = "=SI(SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!H:H),""NO"")=""SI"",F12*CONFIG!B7,0)"
    ws.Range("H12").FormulaLocal = "=F12+G12"
    ws.Range("I12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!C:C),""")"
    ws.Range("J12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!D:D),""")"
    ws.Range("K12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!I:I),0)"
    ws.Range("L12").FormulaLocal = "=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!J:J),""")"

    ws.Range("B12:L12").AutoFill Destination:=ws.Range("B12:L200")

    ws.Range("A12").FormulaLocal = "=SI($J$6=""SIN RESULTADO"",""",$J$6)"

    ws.Range("N11").Value = "Subtotal"
    ws.Range("N12").FormulaLocal = "=SUMA(F12:F200)"
    ws.Range("N13").Value = "IVA"
    ws.Range("N14").FormulaLocal = "=SUMA(G12:G200)"
    ws.Range("N15").Value = "TOTAL"
    ws.Range("N16").FormulaLocal = "=SUMA(H12:H200)"

    ws.Columns("A:N").AutoFit
    ws.Columns("D:D").NumberFormat = "$ #,##0.00"
    ws.Columns("F:H").NumberFormat = "$ #,##0.00"
    ws.Range("E12:E200").NumberFormat = "0%"
End Sub

Private Sub SetupHistorico(ws As Worksheet)
    ws.Range("A1:K1").Value = Array("No_Cotización", "Fecha", "ID_Cliente", "Cliente", "Subtotal", "IVA", "Total", "Estado", "Fecha_Respuesta", "Motivo_Pérdida", "Observaciones")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:K").AutoFit
    ws.Columns("E:G").NumberFormat = "$ #,##0.00"

    With ws.Range("H2:H1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Aprobada""")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With ws.Range("H2:H1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Vencida""")
        .Interior.Color = RGB(255, 199, 206)
    End With
End Sub

Private Sub SetupVentas(ws As Worksheet)
    ws.Range("A1:H1").Value = Array("No_Venta", "Fecha_Venta", "No_Cotización", "Cliente", "Total_Venta", "Medio_Pago", "Estado_Pago", "Fecha_Cobro")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:H").AutoFit
    ws.Columns("E:E").NumberFormat = "$ #,##0.00"
End Sub

Private Sub SetupDashboard(ws As Worksheet)
    ws.Range("A1").Value = "DASHBOARD COMERCIAL (BASE)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "Cotizado mes actual"
    ws.Range("B3").FormulaLocal = "=SUMAR.SI.CONJUNTO(HISTORICO_COTIZACIONES!G:G,HISTORICO_COTIZACIONES!B:B,"">=""&FECHA(AÑO(HOY()),MES(HOY()),1),HISTORICO_COTIZACIONES!B:B,""<=""&FIN.MES(HOY(),0))"

    ws.Range("A4").Value = "Cotizaciones emitidas"
    ws.Range("B4").FormulaLocal = "=CONTAR.SI(HISTORICO_COTIZACIONES!H:H,""Enviada"")+CONTAR.SI(HISTORICO_COTIZACIONES!H:H,""Aprobada"")+CONTAR.SI(HISTORICO_COTIZACIONES!H:H,""Rechazada"")"

    ws.Range("A5").Value = "Tasa de aprobación"
    ws.Range("B5").FormulaLocal = "=SI.ERROR(CONTAR.SI(HISTORICO_COTIZACIONES!H:H,""Aprobada"")/MAX(1,CONTAR.SI(HISTORICO_COTIZACIONES!H:H,""Enviada"")),0)"

    ws.Range("A6").Value = "Ventas cerradas"
    ws.Range("B6").FormulaLocal = "=CONTARA(VENTAS!A:A)-1"

    ws.Range("A7").Value = "Ticket promedio"
    ws.Range("B7").FormulaLocal = "=SI.ERROR(PROMEDIO(VENTAS!E:E),0)"

    ws.Range("B5").NumberFormat = "0.00%"
    ws.Range("B3:B4,B6:B7").NumberFormat = "$ #,##0.00"
    ws.Columns("A:B").AutoFit
End Sub

Private Sub SetupValidations(wsCot As Worksheet, wsHist As Worksheet)
    With wsCot.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CLIENTES!$A$2:$A$1000"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With wsCot.Range("A12:A200").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PRODUCTOS!$A$2:$A$1000"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With wsHist.Range("H2:H1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Borrador,Enviada,Aprobada,Rechazada,Vencida"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub CreateActionButtons(ws As Worksheet)
    Dim s1 As Shape, s2 As Shape, s3 As Shape

    For Each s1 In ws.Shapes
        If s1.Name Like "btn_*" Then s1.Delete
    Next s1

    Set s1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 600, 20, 180, 30)
    s1.Name = "btn_guardar"
    s1.TextFrame.Characters.Text = "Guardar en histórico"
    s1.OnAction = "GuardarCotizacionEnHistorico"

    Set s2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 600, 60, 180, 30)
    s2.Name = "btn_venta"
    s2.TextFrame.Characters.Text = "Convertir a venta"
    s2.OnAction = "ConvertirCotizacionAVenta"

    Set s3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 600, 100, 180, 30)
    s3.Name = "btn_limpiar"
    s3.TextFrame.Characters.Text = "Limpiar formulario"
    s3.OnAction = "LimpiarCotizacion"
End Sub

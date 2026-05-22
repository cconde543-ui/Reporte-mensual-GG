Option Explicit

Public Sub CrearOActualizarPanelReportes()
    On Error GoTo EH

    Dim ws As Worksheet
    Dim meses As Variant
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo EH

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET_NAME
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "Reporte mensual Gerencia General"
    ws.Range("A3").Value = "Año": ws.Range("B3").Value = Year(Date)
    ws.Range("A4").Value = "Mes de cierre": ws.Range("B4").Value = "Enero"

    ws.Range("A1:F1").Merge
    ws.Range("A1:F1").HorizontalAlignment = xlLeft
    ws.Range("A1:F1").VerticalAlignment = xlCenter
    ws.Range("A1:F1").Font.Size = 16
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Font.Color = RGB(0, 32, 96)
    ws.Range("A3:A4").Font.Bold = True
    ws.Range("B3:B4").Interior.Color = RGB(242, 242, 242)
    ws.Range("B3:B4").Borders.LineStyle = xlContinuous

    meses = MesesES()
    With ws.Range("B4").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, Join(meses, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    On Error Resume Next
    ws.Shapes("btnGenerarReporteGG").Delete
    On Error GoTo EH

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D6").Left, ws.Range("D6").Top, ws.Range("F8").Left + ws.Range("F8").Width - ws.Range("D6").Left, ws.Range("F8").Top + ws.Range("F8").Height - ws.Range("D6").Top)
    shp.Name = "btnGenerarReporteGG"
    shp.TextFrame2.TextRange.Characters.Text = "Generar Reporte GG"
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.Fill.ForeColor.RGB = RGB(0, 112, 192)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "Generar_Reporte_GG_Desde_Panel"

    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:F").AutoFit
    Exit Sub
EH:
    Err.Raise Err.Number, "CrearOActualizarPanelReportes", "Error creando/actualizando panel: " & Err.Description
End Sub

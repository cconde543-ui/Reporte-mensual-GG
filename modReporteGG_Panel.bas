Option Explicit

Public Sub CrearOActualizarPanelReportes()
    Dim ws As Worksheet, meses As Variant, shp As Shape
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(PANEL_SHEET_NAME): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1)): ws.Name = PANEL_SHEET_NAME

    ws.Cells.Clear
    ws.Range("A1").Value = "Reporte mensual Gerencia General"
    ws.Range("A3").Value = "Año": ws.Range("B3").Value = 2026
    ws.Range("A4").Value = "Mes de cierre": ws.Range("B4").Value = "Enero"

    meses = MesesES()
    With ws.Range("B4").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, Join(meses, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    On Error Resume Next
    ws.Shapes("btnGenerarReporteGG").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D3").Left, ws.Range("D3").Top, ws.Range("F5").Left + ws.Range("F5").Width - ws.Range("D3").Left, ws.Range("F5").Top + ws.Range("F5").Height - ws.Range("D3").Top)
    shp.Name = "btnGenerarReporteGG"
    shp.TextFrame2.TextRange.Characters.Text = "Generar Reporte GG"
    shp.Fill.ForeColor.RGB = RGB(0, 112, 192)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "Generar_Reporte_GG_Desde_Panel"

    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:F").AutoFit
End Sub

Option Explicit

Public Sub CrearOActualizarPanelReportes()
    Dim ws As Worksheet, m As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET_NAME
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "Reporte mensual Gerencia General"
    ws.Range("A3").Value = "Año": ws.Range("B3").Value = Year(Date)
    ws.Range("A4").Value = "Mes de cierre": ws.Range("B4").Value = "Enero"
    m = MesesES()
    With ws.Range("B4").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, Join(m, ",")
    End With
    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:B").AutoFit
End Sub

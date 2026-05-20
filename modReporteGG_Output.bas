Option Explicit

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable, rng As Range
    On Error Resume Next: Application.DisplayAlerts = False: wbOut.Worksheets("Ejec. Mensual " & anio).Delete: Application.DisplayAlerts = True: On Error GoTo 0
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count)): ws.Name = "Ejec. Mensual " & anio

    ws.Range("A1").Value = "EJECUCIÓN " & anio
    ws.Range("A1").Resize(1, 6).Merge
    ws.Range("A1").Interior.Color = RGB(0, 112, 192)
    ws.Range("A1").Font.Color = vbWhite
    ws.Range("A1").Font.Bold = True

    Set rng = wsBase.Range("A1").CurrentRegion
    Set pc = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng.Address(True, True, xlR1C1, True))
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptGG")

    With pt
        .PivotFields("Financiamiento").Orientation = xlPageField
        .PivotFields("Nivel_1").Orientation = xlRowField
        .PivotFields("Nivel_2").Orientation = xlRowField
        .PivotFields("Nivel_3").Orientation = xlRowField
        .PivotFields("MesNombre").Orientation = xlColumnField
        .PivotFields("MesNum").Orientation = xlHidden
        .AddDataField .PivotFields("Importe"), "Suma de Importe", xlSum
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ShowDrillIndicators = True
    End With

    AplicarFormatoReporteGG ws, mesCierre
End Sub

Public Sub AplicarFormatoReporteGG(ByVal ws As Worksheet, ByVal mesCierre As Long)
    Dim lastRow As Long, lastCol As Long
    lastRow = UltimaFilaConDatos(ws): lastCol = UltimaColConDatos(ws)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol)).Interior.Color = RGB(0, 112, 192)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol)).Font.Color = vbWhite
    ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).NumberFormat = IIf(MOSTRAR_EN_MILES, "#,##0", "#,##0")
    ws.Columns.AutoFit
End Sub

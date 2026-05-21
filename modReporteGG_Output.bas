Option Explicit

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim etapaVisual As String
    etapaVisual = "iniciando"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim campo As String, objetoNothing As String
    On Error GoTo EH

    etapaVisual = "validando objetos de entrada"
    If wbOut Is Nothing Then Err.Raise vbObjectError + 720, "CrearTablaDinamicaOSalidaAgrupada", "Workbook de salida (wbOut) es Nothing."
    If wsBase Is Nothing Then Err.Raise vbObjectError + 721, "CrearTablaDinamicaOSalidaAgrupada", "Hoja base (wsBase) es Nothing."

    etapaVisual = "validando base agregada"
    ValidarBaseAgregada wsBase
    Set rg = wsBase.Range("A1").CurrentRegion

    etapaVisual = "creando hoja de reporte"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets("Ejec. Mensual " & anio).Delete
    Application.DisplayAlerts = True
    On Error GoTo EH

    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = "Ejec. Mensual " & anio
    ArmarEncabezadoVisual ws, anio, mesCierre

    etapaVisual = "creando pivot cache"
    Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg)
    If pivotCacheObj Is Nothing Then objetoNothing = "pivotCacheObj": Err.Raise vbObjectError + 722, "CrearTablaDinamicaOSalidaAgrupada", "No se pudo crear PivotCache."

    etapaVisual = "creando pivot table"
    Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("D7"), TableName:="ptGG")
    If pt Is Nothing Then objetoNothing = "pivotTableObj": Err.Raise vbObjectError + 723, "CrearTablaDinamicaOSalidaAgrupada", "No se pudo crear PivotTable."

    etapaVisual = "configurando campos"
    With pt
        .ManualUpdate = True
        campo = "Nivel_1": .PivotFields(campo).Orientation = xlRowField: .PivotFields(campo).Position = 1
        campo = "Nivel_2": .PivotFields(campo).Orientation = xlRowField: .PivotFields(campo).Position = 2
        campo = "Nivel_3": .PivotFields(campo).Orientation = xlRowField: .PivotFields(campo).Position = 3
        campo = "MesNombre": .PivotFields(campo).Orientation = xlColumnField: .PivotFields(campo).Position = 1
        campo = "MesNum": .PivotFields(campo).Orientation = xlHidden
        campo = "Importe": .AddDataField .PivotFields(campo), "Suma de Importe", xlSum
        .ManualUpdate = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ShowDrillIndicators = True
    End With

    etapaVisual = "ordenando meses"
    OrdenarMesesPivot pt, mesCierre

    etapaVisual = "aplicando formato"
    AplicarFormatoReporteGG ws, pt

    etapaVisual = "creando slicer"
    CrearSlicerFinanciamiento wbOut, ws, pt
    Exit Sub

EH:
    On Error Resume Next
    If Not pt Is Nothing Then pt.TableRange2.Clear
    On Error GoTo 0

    GenerarSalidaEstaticaAgrupada wbOut, wsBase, anio, mesCierre

    Err.Raise Err.Number, "CrearTablaDinamicaOSalidaAgrupada", _
              "Error creando reporte visual. Etapa visual: " & etapaVisual & _
              " | Campo: " & campo & _
              " | Objeto Nothing detectado: " & IIf(Len(objetoNothing) = 0, "(no identificado)", objetoNothing) & _
              " | Detalle: " & Err.Description
End Sub

Public Sub CrearSlicerFinanciamiento(ByVal wbOut As Workbook, ByVal wsReporte As Worksheet, ByVal pt As PivotTable)
    Dim sc As SlicerCache, sl As Slicer

    If wbOut Is Nothing Then Err.Raise vbObjectError + 740, "CrearSlicerFinanciamiento", "wbOut es Nothing."
    If wsReporte Is Nothing Then Err.Raise vbObjectError + 741, "CrearSlicerFinanciamiento", "wsReporte es Nothing."
    If pt Is Nothing Then Err.Raise vbObjectError + 742, "CrearSlicerFinanciamiento", "pivotTable es Nothing."

    On Error Resume Next
    wbOut.SlicerCaches("Slicer_Financiamiento").Delete
    wsReporte.Shapes("slFinanciamiento").Delete
    On Error GoTo EH

    On Error Resume Next
    Set sc = wbOut.SlicerCaches.Add2(pt, "Financiamiento", "Slicer_Financiamiento")
    If sc Is Nothing Then Set sc = wbOut.SlicerCaches.Add(pt, "Financiamiento", "Slicer_Financiamiento")
    On Error GoTo EH

    If sc Is Nothing Then Err.Raise vbObjectError + 743, "CrearSlicerFinanciamiento", "No fue posible crear SlicerCache de Financiamiento."

    Set sl = sc.Slicers.Add(wsReporte, , "slFinanciamiento", "Financiamiento", wsReporte.Range("A7").Left, wsReporte.Range("A7").Top, 165, 190)
    If sl Is Nothing Then Err.Raise vbObjectError + 744, "CrearSlicerFinanciamiento", "No fue posible crear el slicer de Financiamiento."

    sl.NumberOfColumns = 1
    sl.Style = "SlicerStyleLight2"
    Exit Sub
EH:
    On Error Resume Next
    pt.PivotFields("Financiamiento").Orientation = xlPageField
    pt.PivotFields("Financiamiento").Position = 1
    On Error GoTo 0
End Sub

Private Sub ValidarBaseAgregada(ByVal wsBase As Worksheet)
    Dim esperado As Variant, i As Long
    esperado = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")

    If wsBase Is Nothing Then Err.Raise vbObjectError + 700, "ValidarBaseAgregada", "wsBase es Nothing."
    If UltimaFilaConDatos(wsBase) < 2 Then Err.Raise vbObjectError + 701, "ValidarBaseAgregada", "No hay datos agregados para generar la tabla dinámica."
    If UltimaColConDatos(wsBase) < 7 Then Err.Raise vbObjectError + 702, "ValidarBaseAgregada", "La base agregada tiene menos de 7 columnas."

    For i = 1 To 7
        If StrComp(LimpiarTexto(CStr(wsBase.Cells(1, i).Value)), CStr(esperado(i - 1)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 703, "ValidarBaseAgregada", "Encabezado inválido en Base_Agregada. Se esperaba: " & Join(esperado, ", ")
        End If
    Next i
End Sub

Private Sub ArmarEncabezadoVisual(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim titulo As String, mes As String
    mes = UCase$(MesesES()(mesCierre - 1))
    ws.Range("A1:N1").Merge
    ws.Range("A1:N1").Interior.Color = RGB(47, 84, 150)
    ws.Range("A1:N1").RowHeight = 30
    ws.Range("N1").Value = "BPS": ws.Range("N1").Font.Bold = True: ws.Range("N1").Font.Color = vbWhite: ws.Range("N1").HorizontalAlignment = xlRight
    titulo = "Informe de Seguimiento Presupuestal " & mes & " " & anio & " - Ejecución mensual y acumulada" & SufijoUnidadTitulo()
    ws.Range("A3:N3").Merge
    ws.Range("A3").Value = titulo
End Sub

Private Sub OrdenarMesesPivot(ByVal pt As PivotTable, ByVal mesCierre As Long)
    Dim pf As PivotField, i As Long, m As Variant
    Set pf = pt.PivotFields("MesNombre")
    pf.AutoSort xlManual, pf.SourceName
    m = MesesESMin()
    On Error Resume Next
    For i = 0 To mesCierre - 1
        pf.PivotItems(CStr(m(i))).Position = i + 1
    Next i
    On Error GoTo 0
End Sub

Public Sub AplicarFormatoReporteGG(ByVal ws As Worksheet, ByVal pt As PivotTable)
    pt.TableStyle2 = "PivotStyleMedium9"
    If Not pt.DataBodyRange Is Nothing Then pt.DataBodyRange.NumberFormat = "#,##0"
    ws.Columns("A:B").ColumnWidth = 14
    ws.Columns("C:Z").AutoFit
End Sub

Private Sub GenerarSalidaEstaticaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, src As Variant, i As Long, outR As Long
    On Error Resume Next
    Set ws = wbOut.Worksheets("Ejec. Mensual " & anio)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
        ws.Name = "Ejec. Mensual " & anio
    End If

    ws.Cells.Clear
    ArmarEncabezadoVisual ws, anio, mesCierre
    ws.Range("D6:J6").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNombre", "MesNum", "Importe")

    src = wsBase.Range("A1").CurrentRegion.Value2
    outR = 7
    For i = 2 To UBound(src, 1)
        ws.Cells(outR, 4).Resize(1, 7).Value = Array(src(i, 1), src(i, 2), src(i, 3), src(i, 4), src(i, 6), src(i, 5), src(i, 7))
        outR = outR + 1
    Next i
    ws.Range("D6:J6").AutoFilter
    ws.Columns("D:J").AutoFit
End Sub

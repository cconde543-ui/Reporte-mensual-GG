Option Explicit

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim etapa As String, campo As String
    On Error GoTo EH

    ValidarBaseAgregada wsBase
    Set rg = wsBase.Range("A1").CurrentRegion

    On Error Resume Next: Application.DisplayAlerts = False: wbOut.Worksheets("Ejec. Mensual " & anio).Delete: Application.DisplayAlerts = True: On Error GoTo EH
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count)): ws.Name = "Ejec. Mensual " & anio
    ArmarEncabezadoVisual ws, anio, mesCierre

    Debug.Print "Base hoja: " & wsBase.Name
    Debug.Print "Rango fuente: " & rg.Address(External:=True)
    Debug.Print "Filas base: " & rg.Rows.Count
    Debug.Print "Columnas base: " & rg.Columns.Count
    Debug.Print "Encabezados base: " & Join(ObtenerEncabezadosBase(wsBase), " | ")
    Debug.Print "Financiamientos únicos: " & ObtenerValoresUnicosColumna(wsBase, 1)

    etapa = "crear cache": Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg)
    etapa = "crear tabla": Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("D7"), TableName:="ptGG")
    Debug.Print "Campos pivot disponibles: " & CamposPivotAsText(pt)

    With pt
        .ManualUpdate = True
        etapa = "configurar campo": campo = "Nivel_1": .PivotFields("Nivel_1").Orientation = xlRowField: .PivotFields("Nivel_1").Position = 1
        etapa = "configurar campo": campo = "Nivel_2": .PivotFields("Nivel_2").Orientation = xlRowField: .PivotFields("Nivel_2").Position = 2
        etapa = "configurar campo": campo = "Nivel_3": .PivotFields("Nivel_3").Orientation = xlRowField: .PivotFields("Nivel_3").Position = 3
        etapa = "configurar campo": campo = "MesNombre": .PivotFields("MesNombre").Orientation = xlColumnField: .PivotFields("MesNombre").Position = 1
        etapa = "configurar campo": campo = "MesNum": .PivotFields("MesNum").Orientation = xlHidden
        etapa = "agregar valor": campo = "Importe": .AddDataField .PivotFields("Importe"), "Suma de Importe", xlSum
        .ManualUpdate = False
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ShowDrillIndicators = True
    End With

    OrdenarMesesPivot pt, mesCierre
    AplicarFormatoReporteGG ws, pt
    CrearSlicerFinanciamiento wbOut, ws, pt
    Exit Sub
EH:
    Debug.Print "Error Pivot | etapa=" & etapa & " | campo=" & campo & " | nro=" & Err.Number & " | desc=" & Err.Description
    If Not pt Is Nothing Then On Error Resume Next: pt.TableRange2.Clear: On Error GoTo 0
    MsgBox "Falló tabla dinámica, se generará salida estática." & vbCrLf & _
           "Etapa: " & etapa & vbCrLf & _
           "Campo: " & campo & vbCrLf & _
           "Err: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Campos pivot: " & IIf(pt Is Nothing, "(sin PT)", CamposPivotAsText(pt)) & vbCrLf & _
           "Encabezados base: " & Join(ObtenerEncabezadosBase(wsBase), ", "), vbExclamation
    GenerarSalidaEstaticaAgrupada wbOut, wsBase, anio, mesCierre
End Sub

Public Sub CrearSlicerFinanciamiento(ByVal wbOut As Workbook, ByVal wsReporte As Worksheet, ByVal pt As PivotTable)
    Dim sc As SlicerCache, sl As Slicer
    On Error Resume Next
    wbOut.SlicerCaches("Slicer_Financiamiento").Delete
    wsReporte.Shapes("slFinanciamiento").Delete
    On Error GoTo 0

    On Error Resume Next
    Set sc = wbOut.SlicerCaches.Add2(pt, "Financiamiento", "Slicer_Financiamiento")
    If sc Is Nothing Then Set sc = wbOut.SlicerCaches.Add(pt, "Financiamiento", "Slicer_Financiamiento")
    On Error GoTo EH
    If sc Is Nothing Then Err.Raise vbObjectError + 750, , "No fue posible crear SlicerCache de Financiamiento."

    Set sl = sc.Slicers.Add(wsReporte, , "slFinanciamiento", "Financiamiento", wsReporte.Range("A7").Left, wsReporte.Range("A7").Top, 165, 190)
    sl.NumberOfColumns = 1
    sl.Style = "SlicerStyleLight2"
    Exit Sub
EH:
    MsgBox "Advertencia: no se pudo crear segmentación (slicer) de Financiamiento. Se mantiene la tabla dinámica y se aplicará filtro alternativo.", vbExclamation
    On Error Resume Next
    pt.PivotFields("Financiamiento").Orientation = xlPageField
    pt.PivotFields("Financiamiento").Position = 1
    On Error GoTo 0
End Sub

Private Sub ValidarBaseAgregada(ByVal wsBase As Worksheet)
    Dim esperado As Variant, i As Long
    esperado = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")
    If UltimaFilaConDatos(wsBase) < 2 Then Err.Raise vbObjectError + 700, , "No hay datos agregados para generar la tabla dinámica."
    For i = 1 To 7
        If StrComp(LimpiarTexto(CStr(wsBase.Cells(1, i).Value)), CStr(esperado(i - 1)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 701, , "Encabezado inválido en Base_Agregada. Se esperaba: " & Join(esperado, ", ")
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
    ws.Range("A3").Font.Size = 14
    ws.Range("A3").Font.Color = RGB(0, 32, 96)
    ws.Range("A4:N4").Interior.Color = RGB(47, 84, 150)
    ws.Range("D6").Value = "EJECUCIÓN " & anio
    ws.Range("D6").Interior.Color = RGB(68, 114, 196)
    ws.Range("D6").Font.Color = vbWhite
    ws.Range("D6").Font.Bold = True
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
    pt.DataBodyRange.NumberFormat = "#,##0"
    ws.Columns("A:B").ColumnWidth = 14
    ws.Columns("C:Z").AutoFit
End Sub

Private Function CamposPivotAsText(ByVal pt As PivotTable) As String
    Dim pf As PivotField, t As String
    For Each pf In pt.PivotFields: t = t & IIf(Len(t) > 0, " | ", "") & pf.Name: Next pf
    CamposPivotAsText = t
End Function
Private Function ObtenerEncabezadosBase(ByVal wsBase As Worksheet) As Variant
    Dim lastCol As Long, i As Long, arr() As String
    lastCol = UltimaColConDatos(wsBase): ReDim arr(0 To lastCol - 1)
    For i = 1 To lastCol: arr(i - 1) = CStr(wsBase.Cells(1, i).Value): Next i
    ObtenerEncabezadosBase = arr
End Function
Private Function ObtenerValoresUnicosColumna(ByVal ws As Worksheet, ByVal col As Long) As String
    Dim d As Object, arr As Variant, i As Long, t As String, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    arr = ws.Range(ws.Cells(2, col), ws.Cells(UltimaFilaConDatos(ws), col)).Value2
    For i = 1 To UBound(arr, 1)
        If Len(LimpiarTexto(CStr(arr(i, 1)))) > 0 Then d(LimpiarTexto(CStr(arr(i, 1)))) = 1
    Next i
    For Each k In d.Keys: t = t & IIf(Len(t) > 0, " | ", "") & CStr(k): Next k
    ObtenerValoresUnicosColumna = t
End Function

Private Sub GenerarSalidaEstaticaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, src As Variant, i As Long, outR As Long
    Set ws = wbOut.Worksheets("Ejec. Mensual " & anio)
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

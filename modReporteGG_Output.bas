Option Explicit

Private Const LOGO_BPS_PATH As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Logo_BPS.jpg"

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim etapaVisual As String
    etapaVisual = "iniciando"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim pfImporte As PivotField
    Dim errNumPivot As Long, errDescPivot As String
    Dim encabezadosBase As String, camposPivot As String
    Dim manualUpdateActivo As Boolean
    On Error GoTo EH

    etapaVisual = "validando objetos de entrada"
    If wbOut Is Nothing Then Err.Raise vbObjectError + 720, "CrearTablaDinamicaOSalidaAgrupada", "Workbook de salida (wbOut) es Nothing."
    If wsBase Is Nothing Then Err.Raise vbObjectError + 721, "CrearTablaDinamicaOSalidaAgrupada", "Hoja base (wsBase) es Nothing."

    etapaVisual = "validando base agregada"
    ValidarBaseAgregada wsBase
    encabezadosBase = EncabezadosBaseAgregada(wsBase)
    NormalizarBaseParaPivot wsBase
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

    etapaVisual = "creando pivot table"
    Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="ptGG")

    etapaVisual = "configurando campos"
    pt.ManualUpdate = True
    manualUpdateActivo = True

    ConfigurarCampoPivotSeguro pt, "Nivel_1", xlRowField, 1
    ConfigurarCampoPivotSeguro pt, "Nivel_2", xlRowField, 2
    ConfigurarCampoPivotSeguro pt, "Nivel_3", xlRowField, 3
    ConfigurarCampoPivotSeguro pt, "MesNombre", xlColumnField, 1
    ConfigurarCampoPivotSeguro pt, "MesNum", xlColumnField, 2
    ConfigurarCampoPivotSeguro pt, "Financiamiento", xlPageField, 1

    If Not PivotFieldExiste(pt, "Importe") Then
        Err.Raise vbObjectError + 724, "CrearTablaDinamicaOSalidaAgrupada", "No existe el campo 'Importe' en la PivotTable."
    End If
    Set pfImporte = pt.PivotFields("Importe")
    pt.AddDataField pfImporte, "EJECUCIÓN " & anio, xlSum

    pt.ManualUpdate = False
    manualUpdateActivo = False

    pt.RowAxisLayout xlCompactRow
    pt.RepeatAllLabels xlDoNotRepeatLabels
    pt.ShowDrillIndicators = True
    pt.DisplayFieldCaptions = False
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.NullString = ""
    pt.DataBodyRange.NumberFormat = "#.##0"

    OrdenarMesesPivot pt, mesCierre
    AplicarFormatoReporteGG ws, pt, anio
    CrearSlicerFinanciamiento wbOut, ws, pt
    Exit Sub

EH:
    errNumPivot = Err.Number
    errDescPivot = Err.Description
    camposPivot = CamposDisponiblesPivot(pt)

    If manualUpdateActivo Then
        On Error Resume Next
        pt.ManualUpdate = False
        On Error GoTo 0
    End If

    On Error Resume Next
    GenerarSalidaEstaticaAgrupada wbOut, wsBase, anio, mesCierre
    If Err.Number = 0 Then
        Debug.Print "[PIVOT] Fallback estático aplicado. Error original: " & errDescPivot
        Exit Sub
    End If

    Err.Raise errNumPivot, "CrearTablaDinamicaOSalidaAgrupada", _
              "Error creando reporte visual. Etapa visual: " & etapaVisual & _
              " | Campos disponibles pivot: " & camposPivot & _
              " | Encabezados Base_Agregada: " & encabezadosBase & _
              " | Detalle: " & errDescPivot
End Sub

Public Sub CrearSlicerFinanciamiento(ByVal wbOut As Workbook, ByVal wsReporte As Worksheet, ByVal pt As PivotTable)
    Dim sc As SlicerCache, sl As Slicer

    On Error GoTo EH
    On Error Resume Next
    wbOut.SlicerCaches("Slicer_Financiamiento").Delete
    wsReporte.Shapes("slFinanciamiento").Delete
    On Error GoTo EH

    On Error Resume Next
    Set sc = wbOut.SlicerCaches.Add2(pt, "Financiamiento", "Slicer_Financiamiento")
    If sc Is Nothing Then Set sc = wbOut.SlicerCaches.Add(pt, "Financiamiento", "Slicer_Financiamiento")
    On Error GoTo EH

    If sc Is Nothing Then Err.Raise vbObjectError + 743, "CrearSlicerFinanciamiento", "No fue posible crear SlicerCache de Financiamiento."

    Set sl = sc.Slicers.Add(wsReporte, , "slFinanciamiento", "Financiamiento", wsReporte.Range("A5").Left, wsReporte.Range("A5").Top, 165, 230)
    sl.NumberOfColumns = 1
    On Error Resume Next
    sl.Style = "SlicerStyleLight2"
    On Error GoTo EH
    Exit Sub
EH:
    Debug.Print "[ADVERTENCIA] No fue posible crear slicer de Financiamiento: " & Err.Description
    On Error Resume Next
    With wsReporte.Range("A5:A16")
        .Merge
        .Value = "Financiamiento"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri Light"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .Borders.LineStyle = xlContinuous
    End With
    On Error GoTo 0
End Sub

Private Sub ArmarEncabezadoVisual(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim titulo As String, mes As String, arrMeses As Variant
    arrMeses = MesesES()
    mes = UCase$(CStr(arrMeses(mesCierre - 1)))

    ws.Range("A3:M3").UnMerge
    ws.Range("A3:M3").Merge
    ws.Range("A3").Value = "Informe de Seguimiento Presupuestal " & mes & " " & anio & " - Ejecución mensual y acumulada"
    With ws.Range("A3")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Name = "Calibri Light"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(0, 32, 96)
    End With
    ws.Range("A3:M3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A3:M3").Borders(xlEdgeBottom).Weight = xlMedium

    InsertarLogoBPS ws
End Sub

Private Sub InsertarLogoBPS(ByVal ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    If Dir$(LOGO_BPS_PATH, vbNormal) = "" Then
        Debug.Print "[ADVERTENCIA] No se encontró logo BPS en: " & LOGO_BPS_PATH
        Exit Sub
    End If

    On Error GoTo EH
    Set shp = ws.Shapes.AddPicture(LOGO_BPS_PATH, msoFalse, msoTrue, ws.Range("L1").Left, ws.Range("A1").Top + 2, 120, 42)
    shp.Name = "imgLogoBPS"
    shp.LockAspectRatio = msoTrue
    Exit Sub
EH:
    Debug.Print "[ADVERTENCIA] No se pudo insertar logo BPS: " & Err.Description
End Sub

Private Sub OrdenarMesesPivot(ByVal pt As PivotTable, ByVal mesCierre As Long)
    Dim pfNom As PivotField, pfNum As PivotField, m As Variant, i As Long
    On Error Resume Next
    Set pfNom = pt.PivotFields("MesNombre")
    Set pfNum = pt.PivotFields("MesNum")
    On Error GoTo 0
    If pfNom Is Nothing Or pfNum Is Nothing Then Exit Sub

    m = MesesESMin()
    pfNum.AutoSort xlAscending, pfNum.SourceName

    On Error Resume Next
    For i = 0 To 11
        pfNom.PivotItems(CStr(m(i))).Visible = (i + 1 <= mesCierre)
    Next i
    On Error GoTo 0
End Sub

Public Sub AplicarFormatoReporteGG(ByVal ws As Worksheet, ByVal pt As PivotTable, ByVal anio As Long)
    Dim rng As Range
    pt.TableStyle2 = "PivotStyleMedium9"

    With ws.Cells
        .Font.Name = "Calibri Light"
        .Font.Size = 11
    End With

    On Error Resume Next
    Set rng = pt.TableRange1
    If Not rng Is Nothing Then
        rng.Columns(1).ColumnWidth = 28
        rng.Rows(1).Font.Color = RGB(255, 255, 255)
        rng.Rows(1).Interior.Color = RGB(0, 84, 147)
    End If
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:M").AutoFit
End Sub

Private Sub ValidarBaseAgregada(ByVal wsBase As Worksheet)
    Dim esperado As Variant, i As Long
    esperado = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")

    If wsBase Is Nothing Then Err.Raise vbObjectError + 700, "ValidarBaseAgregada", "wsBase es Nothing."
    If UltimaFilaConDatos(wsBase) < 2 Then Err.Raise vbObjectError + 701, "ValidarBaseAgregada", "No hay datos agregados para generar la tabla dinámica."
    If UltimaColConDatos(wsBase) < 7 Then Err.Raise vbObjectError + 702, "ValidarBaseAgregada", "La base agregada tiene menos de 7 columnas."

    For i = 1 To 7
        If StrComp(LimpiarTexto(CStr(wsBase.Cells(1, i).Value)), CStr(esperado(i - 1)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 703, "ValidarBaseAgregada", "Encabezado inválido en Base_Agregada."
        End If
    Next i
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
    ws.Range("B5:H5").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNombre", "MesNum", "Importe")

    src = wsBase.Range("A1").CurrentRegion.Value2
    outR = 6
    For i = 2 To UBound(src, 1)
        ws.Cells(outR, 2).Resize(1, 7).Value = Array(src(i, 1), src(i, 2), src(i, 3), src(i, 4), src(i, 6), src(i, 5), src(i, 7))
        outR = outR + 1
    Next i
End Sub

Private Sub ConfigurarCampoPivotSeguro(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    Dim pf As PivotField
    If pt Is Nothing Then Err.Raise vbObjectError + 730, "ConfigurarCampoPivotSeguro", "PivotTable es Nothing."
    If Not PivotFieldExiste(pt, nombreCampo) Then Err.Raise vbObjectError + 731, "ConfigurarCampoPivotSeguro", "No existe el campo '" & nombreCampo & "'."
    Set pf = pt.PivotFields(nombreCampo)
    pf.Orientation = orientacion
    If posicion > 0 Then pf.Position = posicion
End Sub

Private Function PivotFieldExiste(ByVal pt As PivotTable, ByVal nombreCampo As String) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nombreCampo)
    PivotFieldExiste = Not pf Is Nothing
    Set pf = Nothing
    On Error GoTo 0
End Function

Private Function CamposDisponiblesPivot(ByVal pt As PivotTable) As String
    Dim pf As PivotField
    Dim arr() As String
    Dim n As Long
    If pt Is Nothing Then Exit Function
    For Each pf In pt.PivotFields
        ReDim Preserve arr(0 To n)
        arr(n) = CStr(pf.Name)
        n = n + 1
    Next pf
    If n > 0 Then CamposDisponiblesPivot = Join(arr, ", ")
End Function

Private Function EncabezadosBaseAgregada(ByVal wsBase As Worksheet) As String
    Dim lastCol As Long, i As Long, arr() As String
    lastCol = UltimaColConDatos(wsBase)
    If lastCol < 1 Then Exit Function
    ReDim arr(1 To lastCol)
    For i = 1 To lastCol
        arr(i) = LimpiarTexto(CStr(wsBase.Cells(1, i).Value))
    Next i
    EncabezadosBaseAgregada = Join(arr, ", ")
End Function

Private Sub NormalizarBaseParaPivot(ByVal wsBase As Worksheet)
    Dim lastRow As Long, i As Long, valImp As Variant
    lastRow = UltimaFilaConDatos(wsBase)
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        If Len(Trim$(CStr(wsBase.Cells(i, 2).Value))) = 0 Then wsBase.Cells(i, 2).Value = "(Sin clasificar)"
        If Len(Trim$(CStr(wsBase.Cells(i, 3).Value))) = 0 Then wsBase.Cells(i, 3).Value = "(Sin clasificar)"
        If Len(Trim$(CStr(wsBase.Cells(i, 4).Value))) = 0 Then wsBase.Cells(i, 4).Value = "(Sin clasificar)"

        valImp = wsBase.Cells(i, 7).Value
        If IsError(valImp) Then
            wsBase.Cells(i, 7).Value = 0#
        ElseIf Len(Trim$(CStr(valImp))) = 0 Then
            wsBase.Cells(i, 7).Value = 0#
        ElseIf Not IsNumeric(valImp) Then
            wsBase.Cells(i, 7).Value = CDbl(Val(Replace(CStr(valImp), ",", ".")))
        End If
    Next i
End Sub

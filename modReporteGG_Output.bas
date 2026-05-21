Option Explicit

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim etapaVisual As String
    etapaVisual = "iniciando"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim campo As String, objetoNothing As String
    Dim campoActual As String, accionActual As String
    Dim pf As PivotField, pfImporte As PivotField
    Dim errNumPivot As Long, errDescPivot As String
    Dim encabezadosBase As String, camposPivot As String
    Dim manualUpdateActivo As Boolean
    On Error GoTo EH

    Debug.Print "[VISUAL] Entrando a CrearTablaDinamicaOSalidaAgrupada"
    Debug.Print "[VISUAL] wbOut Is Nothing: "; (wbOut Is Nothing)
    Debug.Print "[VISUAL] wsBase Is Nothing: "; (wsBase Is Nothing)
    Debug.Print "[VISUAL] anio: "; anio
    Debug.Print "[VISUAL] mesCierre: "; mesCierre

    etapaVisual = "validando objetos de entrada"
    If wbOut Is Nothing Then Err.Raise vbObjectError + 720, "CrearTablaDinamicaOSalidaAgrupada", "Workbook de salida (wbOut) es Nothing."
    If wsBase Is Nothing Then Err.Raise vbObjectError + 721, "CrearTablaDinamicaOSalidaAgrupada", "Hoja base (wsBase) es Nothing."

    etapaVisual = "validando base agregada"
    ValidarBaseAgregada wsBase
    encabezadosBase = EncabezadosBaseAgregada(wsBase)
    NormalizarBaseParaPivot wsBase
    Debug.Print "[VISUAL] Antes de crear rango base"
    Set rg = wsBase.Range("A1").CurrentRegion

    etapaVisual = "creando hoja de reporte"
    Debug.Print "[VISUAL] Antes de crear hoja de reporte"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets("Ejec. Mensual " & anio).Delete
    Application.DisplayAlerts = True
    On Error GoTo EH

    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = "Ejec. Mensual " & anio
    Debug.Print "[VISUAL] Despues de crear ws. ws Is Nothing: "; (ws Is Nothing)
    Debug.Print "[VISUAL] Antes de ArmarEncabezadoVisual"
    ArmarEncabezadoVisual ws, anio, mesCierre
    Debug.Print "[VISUAL] Despues de ArmarEncabezadoVisual"

    etapaVisual = "creando pivot cache"
    Debug.Print "[VISUAL] Antes de crear pivotCacheObj"
    Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg)
    If pivotCacheObj Is Nothing Then objetoNothing = "pivotCacheObj": Err.Raise vbObjectError + 722, "CrearTablaDinamicaOSalidaAgrupada", "No se pudo crear PivotCache."

    etapaVisual = "creando pivot table"
    Debug.Print "[VISUAL] Antes de crear pt"
    Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("D7"), TableName:="ptGG")
    If pt Is Nothing Then objetoNothing = "pivotTableObj": Err.Raise vbObjectError + 723, "CrearTablaDinamicaOSalidaAgrupada", "No se pudo crear PivotTable."

    etapaVisual = "configurando campos"
    Debug.Print "[VISUAL] Antes de configurar campos de pt"
    pt.ManualUpdate = True
    manualUpdateActivo = True

    campoActual = "Nivel_1"
    accionActual = "asignando Nivel_1 como fila"
    Debug.Print "[PIVOT] Configurando campo: " & campoActual
    ConfigurarCampoPivotSeguro pt, campoActual, xlRowField, 1
    Debug.Print "[PIVOT] OK campo: " & campoActual

    campoActual = "Nivel_2"
    accionActual = "asignando Nivel_2 como fila"
    Debug.Print "[PIVOT] Configurando campo: " & campoActual
    ConfigurarCampoPivotSeguro pt, campoActual, xlRowField, 2
    Debug.Print "[PIVOT] OK campo: " & campoActual

    campoActual = "Nivel_3"
    accionActual = "asignando Nivel_3 como fila"
    Debug.Print "[PIVOT] Configurando campo: " & campoActual
    ConfigurarCampoPivotSeguro pt, campoActual, xlRowField, 3
    Debug.Print "[PIVOT] OK campo: " & campoActual

    campoActual = "MesNombre"
    accionActual = "asignando MesNombre como columna"
    Debug.Print "[PIVOT] Configurando campo: " & campoActual
    ConfigurarCampoPivotSeguro pt, campoActual, xlColumnField, 1
    Debug.Print "[PIVOT] OK campo: " & campoActual

    campoActual = "Importe"
    accionActual = "agregando Importe como campo de valores"
    Debug.Print "[PIVOT] Configurando campo: " & campoActual
    If Not PivotFieldExiste(pt, campoActual) Then
        Err.Raise vbObjectError + 724, "CrearTablaDinamicaOSalidaAgrupada", "No existe el campo '" & campoActual & "' en la PivotTable."
    End If
    Set pfImporte = pt.PivotFields(campoActual)
    pt.AddDataField pfImporte, "Suma de Importe", xlSum
    Debug.Print "[PIVOT] OK campo: " & campoActual

    pt.ManualUpdate = False
    manualUpdateActivo = False

    accionActual = "aplicando RowAxisLayout"
    On Error Resume Next
    pt.RowAxisLayout xlTabularRow
    If Err.Number <> 0 Then
        Debug.Print "[PIVOT] Advertencia RowAxisLayout: " & Err.Description
        Err.Clear
    End If
    On Error GoTo EH

    accionActual = "aplicando RepeatAllLabels"
    On Error Resume Next
    pt.RepeatAllLabels xlRepeatLabels
    If Err.Number <> 0 Then
        Debug.Print "[PIVOT] Advertencia RepeatAllLabels: " & Err.Description
        Err.Clear
    End If
    On Error GoTo EH

    accionActual = "aplicando ShowDrillIndicators"
    On Error Resume Next
    pt.ShowDrillIndicators = True
    If Err.Number <> 0 Then
        Debug.Print "[PIVOT] Advertencia ShowDrillIndicators: " & Err.Description
        Err.Clear
    End If
    On Error GoTo EH

    etapaVisual = "ordenando meses"
    OrdenarMesesPivot pt, mesCierre

    etapaVisual = "aplicando formato"
    AplicarFormatoReporteGG ws, pt

    etapaVisual = "creando slicer"
    Debug.Print "[VISUAL] Antes de crear slicer"
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
    If Not pt Is Nothing Then pt.TableRange2.Clear
    On Error GoTo 0

    On Error Resume Next
    GenerarSalidaEstaticaAgrupada wbOut, wsBase, anio, mesCierre
    If Err.Number = 0 Then
        Debug.Print "[PIVOT] Fallback estático aplicado correctamente."
        Debug.Print "[PIVOT] Error pivot original. Etapa: " & etapaVisual & " | Campo: " & campoActual & " | Acción: " & accionActual & " | Err.Number: " & errNumPivot & " | Err.Description: " & errDescPivot
        Debug.Print "[PIVOT] Campos disponibles pivot: " & camposPivot
        Debug.Print "[PIVOT] Encabezados Base_Agregada: " & encabezadosBase
        Exit Sub
    End If

    Err.Raise errNumPivot, "CrearTablaDinamicaOSalidaAgrupada", _
              "Error creando reporte visual. Etapa visual: " & etapaVisual & _
              " | Campo actual: " & campoActual & _
              " | Acción actual: " & accionActual & _
              " | Campos disponibles pivot: " & camposPivot & _
              " | Encabezados Base_Agregada: " & encabezadosBase & _
              " | Objeto Nothing detectado: " & IIf(Len(objetoNothing) = 0, "(no identificado)", objetoNothing) & _
              " | Err.Number: " & errNumPivot & _
              " | Detalle: " & errDescPivot & _
              " | Error fallback: " & Err.Description
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
    Dim arrMeses As Variant
    Debug.Print "[ENCABEZADO] Entrando a ArmarEncabezadoVisual"
    Debug.Print "[ENCABEZADO] ws Is Nothing: "; (ws Is Nothing)
    Debug.Print "[ENCABEZADO] anio: "; anio
    Debug.Print "[ENCABEZADO] mesCierre: "; mesCierre

    If mesCierre < 1 Or mesCierre > 12 Then
        Err.Raise vbObjectError + 900, "ArmarEncabezadoVisual", "Mes de cierre inválido: " & mesCierre
    End If

    arrMeses = MesesES()
    Debug.Print "[ENCABEZADO] LBound MesesES: "; LBound(arrMeses)
    Debug.Print "[ENCABEZADO] UBound MesesES: "; UBound(arrMeses)
    mes = UCase$(CStr(arrMeses(mesCierre - 1)))
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
    If pt Is Nothing Then Exit Sub
    If Not PivotFieldExiste(pt, "MesNombre") Then Exit Sub

    On Error Resume Next
    Set pf = pt.PivotFields("MesNombre")
    If Err.Number <> 0 Then
        Debug.Print "[PIVOT] Advertencia al obtener MesNombre para ordenar: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    pf.AutoSort xlManual, pf.SourceName
    If Err.Number <> 0 Then
        Debug.Print "[PIVOT] Advertencia AutoSort MesNombre: " & Err.Description
        Err.Clear
    End If

    m = MesesESMin()
    For i = 0 To mesCierre - 1
        pf.PivotItems(CStr(m(i))).Position = i + 1
        If Err.Number <> 0 Then
            Debug.Print "[PIVOT] Advertencia ordenando mes '" & CStr(m(i)) & "': " & Err.Description
            Err.Clear
        End If
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

Private Sub ConfigurarCampoPivotSeguro(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    Dim pf As PivotField
    If pt Is Nothing Then Err.Raise vbObjectError + 730, "ConfigurarCampoPivotSeguro", "PivotTable es Nothing."
    If Not PivotFieldExiste(pt, nombreCampo) Then
        Err.Raise vbObjectError + 731, "ConfigurarCampoPivotSeguro", "No existe el campo '" & nombreCampo & "'. Campos disponibles: " & CamposDisponiblesPivot(pt)
    End If
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
    If pt Is Nothing Then
        CamposDisponiblesPivot = "(PivotTable Nothing)"
        Exit Function
    End If
    For Each pf In pt.PivotFields
        ReDim Preserve arr(0 To n)
        arr(n) = CStr(pf.Name)
        n = n + 1
    Next pf
    If n = 0 Then
        CamposDisponiblesPivot = "(sin campos)"
    Else
        CamposDisponiblesPivot = Join(arr, ", ")
    End If
End Function

Private Function EncabezadosBaseAgregada(ByVal wsBase As Worksheet) As String
    Dim lastCol As Long
    Dim i As Long
    Dim arr() As String
    lastCol = UltimaColConDatos(wsBase)
    If lastCol < 1 Then
        EncabezadosBaseAgregada = "(sin encabezados)"
        Exit Function
    End If
    ReDim arr(1 To lastCol)
    For i = 1 To lastCol
        arr(i) = LimpiarTexto(CStr(wsBase.Cells(1, i).Value))
    Next i
    EncabezadosBaseAgregada = Join(arr, ", ")
End Function

Private Sub NormalizarBaseParaPivot(ByVal wsBase As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim valImp As Variant

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

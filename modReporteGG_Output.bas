Option Explicit

Private Const LOGO_BPS_PATH As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Logo_BPS.jpg"
Private Const DUMMY_MARCA As String = "__DUMMY_VISUAL__"

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim etapaVisual As String
    etapaVisual = "iniciando"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim pfImporte As PivotField
    Dim campoNivel1 As String, campoNivel2 As String, campoNivel3 As String
    Dim campoMesNombre As String, campoMesNum As String
    Dim errNumPivot As Long, errDescPivot As String
    Dim encabezadosBase As String, camposPivot As String
    Dim sourceAddress As String
    Dim campoActual As String, accionActual As String
    Dim orientacionActual As XlPivotFieldOrientation
    Dim campoMesColumnaUsado As String
    Dim msgMesFallback As String
    On Error GoTo EH

    etapaVisual = "validando objetos de entrada"
    If wbOut Is Nothing Then Err.Raise vbObjectError + 720, "CrearTablaDinamicaOSalidaAgrupada", "Workbook de salida (wbOut) es Nothing."
    If wsBase Is Nothing Then Err.Raise vbObjectError + 721, "CrearTablaDinamicaOSalidaAgrupada", "Hoja base (wsBase) es Nothing."

    etapaVisual = "validando base agregada"
    ValidarBaseAgregada wsBase
    encabezadosBase = EncabezadosBaseAgregada(wsBase)
    NormalizarYValidarMesesBase wsBase
    Set rg = wsBase.Range("A1").CurrentRegion
    sourceAddress = rg.Address(ReferenceStyle:=xlR1C1, External:=True)
    DiagnosticarBaseAgregadaPrePivot wsBase, sourceAddress

    etapaVisual = "creando hoja de reporte"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets("Ejec. Mensual " & anio).Delete
    Application.DisplayAlerts = True
    On Error GoTo EH

    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = "Ejec. Mensual " & anio
    CrearHojaReporteVisual ws, anio, mesCierre

    etapaVisual = "creando pivot cache"
    Debug.Print "[PIVOT] SourceData=" & sourceAddress
    Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceAddress)
    pivotCacheObj.MissingItemsLimit = xlMissingItemsNone

    etapaVisual = "creando pivot table"
    Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="ptGG")

    etapaVisual = "refrescando cache y tabla"
    pivotCacheObj.Refresh
    pt.RefreshTable

    etapaVisual = "configurando campos"

    campoNivel1 = ObtenerCampoDisponible(pt, Array("Nivel_1", "Nivel 1"))
    campoNivel2 = ObtenerCampoDisponible(pt, Array("Nivel_2", "Nivel 2"))
    campoNivel3 = ObtenerCampoDisponible(pt, Array("Nivel_3", "Nivel 3"))
    campoMesNombre = ObtenerCampoDisponible(pt, Array("MesNombre", "Mes Nombre"))
    campoMesNum = ObtenerCampoDisponible(pt, Array("MesNum", "Mes Num"))

    If Not PivotFieldExiste(pt, "Importe") Then
        Err.Raise vbObjectError + 724, "CrearTablaDinamicaOSalidaAgrupada", "No existe el campo 'Importe' en la PivotTable."
    End If
    campoActual = "Importe"
    accionActual = "agregando campo de valores Importe"
    orientacionActual = xlDataField
    Set pfImporte = pt.PivotFields("Importe")
    pt.AddDataField pfImporte, "EJECUCIÓN " & anio, xlSum

    campoActual = campoNivel1
    accionActual = "asignando " & campoNivel1 & " como xlRowField"
    orientacionActual = xlRowField
    ConfigurarCampoPivotSeguro pt, campoNivel1, xlRowField, 1

    campoActual = campoNivel2
    accionActual = "asignando " & campoNivel2 & " como xlRowField"
    orientacionActual = xlRowField
    ConfigurarCampoPivotSeguro pt, campoNivel2, xlRowField, 2

    campoActual = campoNivel3
    accionActual = "asignando " & campoNivel3 & " como xlRowField"
    orientacionActual = xlRowField
    ConfigurarCampoPivotSeguro pt, campoNivel3, xlRowField, 3

    campoActual = campoMesNombre
    accionActual = "asignando " & campoMesNombre & " como xlColumnField"
    orientacionActual = xlColumnField
    campoMesColumnaUsado = ConfigurarCampoMesColumnaConFallback(pt, campoMesNombre, campoMesNum, msgMesFallback)
    If Len(msgMesFallback) > 0 Then Debug.Print msgMesFallback

    If PivotFieldExiste(pt, "Financiamiento") Then
        With pt.PivotFields("Financiamiento")
            .Orientation = xlPageField
            .Position = 1
        End With
    End If

    pt.ManualUpdate = False

    Debug.Print "[PIVOT] Filas base=" & UltimaFilaConDatos(wsBase) - 1 & " Cols base=" & UltimaColConDatos(wsBase)
    Debug.Print "[PIVOT] Campo de mes en columnas: " & campoMesColumnaUsado
    pt.RowAxisLayout xlCompactRow
    pt.RepeatAllLabels xlDoNotRepeatLabels
    pt.ShowDrillIndicators = True
    pt.DisplayFieldCaptions = False
    pt.ColumnGrand = True
    pt.RowGrand = True
    pt.NullString = ""
    pt.DisplayNullString = True

    If StrComp(campoMesColumnaUsado, campoMesNombre, vbTextCompare) = 0 Then
        OrdenarMesesPivot pt, campoMesNombre, campoMesNum
    End If
    ColapsarPivotInicial pt
    If Not pt.DataBodyRange Is Nothing Then pt.DataBodyRange.NumberFormat = "#,##0"

    AplicarFormatoReporteGG ws, pt, anio

    Debug.Print "[SLICER] omitido temporalmente hasta estabilizar PivotTable."
    Exit Sub

EH:
    errNumPivot = Err.Number
    errDescPivot = Err.Description
    camposPivot = CamposDisponiblesPivot(pt)

    Err.Raise errNumPivot, "CrearTablaDinamicaOSalidaAgrupada", _
              "Error creando reporte visual. Etapa visual: " & etapaVisual & _
              " | campoActual: " & campoActual & _
              " | accionActual: " & accionActual & _
              " | orientación solicitada: " & OrientacionPivotTexto(orientacionActual) & _
              " | Campos disponibles pivot: " & camposPivot & _
              " | Encabezados Base_Agregada: " & encabezadosBase & _
              " | Err.Number: " & CStr(errNumPivot) & _
              " | Err.Description: " & errDescPivot
End Sub

Public Function CrearSlicerFinanciamiento(ByVal wbOut As Workbook, ByVal wsReporte As Worksheet, ByVal pt As PivotTable) As Boolean
    Dim sc As SlicerCache, sl As Slicer
    Dim topPos As Double, leftPos As Double, ancho As Double, alto As Double
    Dim it As SlicerItem, nombreItem As String
    Dim nombreCampoFin As String
    Dim etapaSlicer As String

    CrearSlicerFinanciamiento = False
    etapaSlicer = "eliminando slicer anterior"
    On Error Resume Next
    Debug.Print "[SLICER] Etapa: " & etapaSlicer
    wbOut.SlicerCaches("Slicer_Financiamiento").Delete
    wsReporte.Shapes("slFinanciamiento").Delete
    wsReporte.Shapes("shpFinanciamientoFallback").Delete
    On Error GoTo 0

    nombreCampoFin = ""
    If PivotFieldExiste(pt, "Financiamiento") Then
        nombreCampoFin = "Financiamiento"
    ElseIf PivotFieldExiste(pt, "Financiamento") Then
        nombreCampoFin = "Financiamento"
    End If
    If Len(nombreCampoFin) = 0 Then
        Debug.Print "[SLICER] No existe campo Financiamiento/Financiamento."
        Exit Function
    End If

    On Error GoTo EH
    Set sc = Nothing
    etapaSlicer = "creando SlicerCache (Add sin nombre)"
    Debug.Print "[SLICER] Antes de SlicerCaches.Add | etapa=" & etapaSlicer & " | campo=" & nombreCampoFin
    Set sc = wbOut.SlicerCaches.Add(pt, nombreCampoFin)
    If sc Is Nothing Then
        etapaSlicer = "creando SlicerCache (Add con nombre)"
        Debug.Print "[SLICER] Reintento SlicerCaches.Add con nombre | etapa=" & etapaSlicer
        Set sc = wbOut.SlicerCaches.Add(pt, nombreCampoFin, "Slicer_Financiamiento")
    End If
    If sc Is Nothing Then
        Debug.Print "[SLICER] No fue posible crear SlicerCache."
        GoTo EH
    End If

    wsReporte.Columns("A").ColumnWidth = 28.14
    leftPos = wsReporte.Range("A5").Left
    topPos = wsReporte.Range("A5").Top
    ancho = wsReporte.Range("A5").Width
    alto = wsReporte.Range("A5:A14").Height

    etapaSlicer = "creando slicer visual"
    Debug.Print "[SLICER] Antes de sc.Slicers.Add | etapa=" & etapaSlicer
    Set sl = sc.Slicers.Add(wsReporte, , "slFinanciamiento", nombreCampoFin, leftPos, topPos, ancho, alto)

    etapaSlicer = "configurando propiedades de slicer"
    Debug.Print "[SLICER] Antes de configurar propiedades | etapa=" & etapaSlicer
    With sl
        .NumberOfColumns = 1
        .DisplayHeader = True
        .ColumnWidth = ancho - 12
        .RowHeight = 16
    End With

    For Each it In sc.SlicerItems
        nombreItem = Trim$(CStr(it.Name))
        If StrComp(nombreItem, "(Dummy)", vbTextCompare) = 0 Or StrComp(nombreItem, DUMMY_MARCA, vbTextCompare) = 0 Then
            On Error Resume Next
            it.Selected = False
            On Error GoTo EH
        End If
    Next it

    etapaSlicer = "configurando shape"
    Debug.Print "[SLICER] Antes de configurar shape | etapa=" & etapaSlicer
    With sl.Shape
        .Placement = xlMove
        .PrintObject = True
        .Locked = True
    End With

    etapaSlicer = "aplicando estilo"
    Debug.Print "[SLICER] Antes de aplicar estilo | etapa=" & etapaSlicer
    AplicarEstiloSlicerAzul sl
    Debug.Print "[SLICER] creado=SI campo=" & nombreCampoFin
    CrearSlicerFinanciamiento = True
    Exit Function
EH:
    Debug.Print "[SLICER] creado=NO | etapa=" & etapaSlicer & _
                " | Err.Number=" & Err.Number & _
                " | Err.Description=" & Err.Description & _
                " | TypeName(sc)=" & TypeName(sc) & _
                " | TypeName(sl)=" & TypeName(sl)
    Err.Clear
    On Error GoTo 0
    CrearSlicerFinanciamiento = False
End Function

Private Sub ArmarEncabezadoVisual(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim mes As String, arrMeses As Variant
    Dim rngBandaSuperior As Range, rngTitulo As Range, rngSubtitulo As Range

    arrMeses = MesesES()
    mes = UCase$(CStr(arrMeses(mesCierre - 1)))

    ws.Rows(1).RowHeight = 50.25
    ws.Rows(2).RowHeight = 15
    ws.Rows(3).RowHeight = 24

    Set rngBandaSuperior = ws.Range("A1:M1")
    Set rngTitulo = ws.Range("A3:M3")
    Set rngSubtitulo = ws.Range("A3:M3")

    rngBandaSuperior.UnMerge
    rngSubtitulo.UnMerge
    rngTitulo.Merge

    rngTitulo.ClearContents
    ws.Range("A2:M2").ClearContents

    CrearBandaAzulSuperior ws

    rngTitulo.Cells(1, 1).Value = "Informe de Seguimiento Presupuestal " & mes & " " & anio & " - Ejecución mensual y acumulada"
    With rngSubtitulo
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Name = "Calibri Light"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(0, 32, 96)
    End With

    With rngSubtitulo.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 32, 96)
    End With

    InsertarLogoBPS ws
End Sub


Private Sub CrearBandaAzulSuperior(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim rng As Range
    On Error Resume Next
    ws.Shapes("shpBandaAzul").Delete
    On Error GoTo 0
    Set rng = ws.Range("A1:M1")
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "shpBandaAzul"
    shp.Fill.ForeColor.RGB = RGB(0, 84, 147)
    shp.Line.Visible = msoFalse
    shp.Placement = xlMove
    shp.ZOrder msoSendToBack
End Sub

Private Sub InsertarLogoBPS(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim logoH As Double, logoW As Double, topPos As Double, leftPos As Double
    Dim rngBandaSuperior As Range

    On Error Resume Next
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    If Dir$(LOGO_BPS_PATH, vbNormal) = "" Then
        Debug.Print "[ADVERTENCIA] No se encontró logo BPS en: " & LOGO_BPS_PATH
        Exit Sub
    End If

    On Error GoTo EH
    logoH = ws.Rows(1).Height - 6
    Set shp = ws.Shapes.AddPicture(LOGO_BPS_PATH, msoFalse, msoTrue, 0, 0, -1, logoH)
    shp.Name = "imgLogoBPS"
    shp.LockAspectRatio = msoTrue
    logoW = shp.Width

    topPos = ws.Rows(1).Top + (ws.Rows(1).Height - shp.Height) / 2
    Set rngBandaSuperior = ws.Range("A1:M1")
    leftPos = rngBandaSuperior.Left + rngBandaSuperior.Width - logoW - 6
    shp.Top = topPos
    shp.Left = leftPos
    shp.Placement = xlMove
    Exit Sub
EH:
    Debug.Print "[ADVERTENCIA] No se pudo insertar logo BPS: " & Err.Description
End Sub

Private Sub OrdenarMesesPivot(ByVal pt As PivotTable, ByVal campoMesNombre As String, ByVal campoMesNum As String)
    Dim pfNom As PivotField, pfNum As PivotField, m As Variant, i As Long
    On Error Resume Next
    Set pfNom = pt.PivotFields(campoMesNombre)
    Set pfNum = pt.PivotFields(campoMesNum)
    On Error GoTo 0
    If pfNom Is Nothing Then Exit Sub

    m = MesesES()
    If Not pfNum Is Nothing Then
        On Error Resume Next
        pfNum.Orientation = xlHidden
        On Error GoTo 0
    End If

    On Error Resume Next
    pfNom.ShowAllItems = True
    pfNom.AutoSort xlManual, pfNom.SourceName
    For i = 0 To 11
        pfNom.PivotItems(CStr(m(i))).Position = i + 1
        pfNom.PivotItems(CStr(m(i))).Visible = True
    Next i
    On Error GoTo 0
End Sub

Public Sub AplicarFormatoReporteGG(ByVal ws As Worksheet, ByVal pt As PivotTable, ByVal anio As Long)
    Dim rng As Range

    On Error Resume Next
    pt.TableStyle2 = "PivotStyleMedium9"
    On Error GoTo 0

    ws.Cells.Interior.Color = vbWhite
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayGridlines = False

    With ws.Cells
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With

    On Error Resume Next
    Set rng = pt.TableRange1
    If Not rng Is Nothing Then
        rng.Columns(1).ColumnWidth = 28
        rng.Rows(1).Font.Color = RGB(255, 255, 255)
        rng.Rows(1).Interior.Color = RGB(0, 84, 147)
    End If
    If Not pt.DataBodyRange Is Nothing Then pt.DataBodyRange.NumberFormat = "#,##0"
    pt.PivotFields("Importe").NumberFormat = "#,##0"
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 28.14
    ws.Columns("B:M").AutoFit
    ws.Columns("B").ColumnWidth = 28

End Sub

Private Sub ColapsarPivotInicial(ByVal pt As PivotTable)
    Dim pf As PivotField, pi As PivotItem
    Dim campoNivel1 As String
    On Error Resume Next
    For Each pf In pt.RowFields
        pf.ShowDetail = False
    Next pf
    On Error GoTo 0

    campoNivel1 = ""
    On Error Resume Next
    campoNivel1 = ObtenerCampoDisponible(pt, Array("Nivel_1", "Nivel 1"))
    On Error GoTo 0
    If Len(campoNivel1) = 0 Then Exit Sub

    On Error Resume Next
    For Each pi In pt.PivotFields(campoNivel1).PivotItems
        pi.ShowDetail = False
    Next pi
    On Error GoTo 0

    On Error Resume Next
    pt.PivotFields(campoNivel1).PivotItems(DUMMY_MARCA).Visible = False
    On Error GoTo 0
End Sub

Private Sub AplicarEstiloSlicerAzul(ByVal sl As Slicer)
    On Error Resume Next
    sl.Style = "SlicerStyleLight2"
    If Err.Number <> 0 Then
        Err.Clear
        sl.Style = "SlicerStyleLight6"
    End If
    On Error GoTo 0
End Sub

Private Function CmToPt(ByVal cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function

Private Sub AgregarFilasDummyMeses(ByVal wsBase As Worksheet)
    Dim i As Long, lastRow As Long, m As Variant
    Dim existe As Boolean

    existe = False
    lastRow = UltimaFilaConDatos(wsBase)
    If lastRow >= 2 Then
        For i = 2 To lastRow
            If CStr(wsBase.Cells(i, 2).Value) = DUMMY_MARCA Then
                existe = True
                Exit For
            End If
        Next i
    End If
    If existe Then Exit Sub

    m = MesesES()
    For i = 0 To 11
        lastRow = lastRow + 1
        wsBase.Cells(lastRow, 1).Value = "(Dummy)"
        wsBase.Cells(lastRow, 2).Value = DUMMY_MARCA
        wsBase.Cells(lastRow, 3).Value = ""
        wsBase.Cells(lastRow, 4).Value = ""
        wsBase.Cells(lastRow, 5).Value = i + 1
        wsBase.Cells(lastRow, 6).Value = CStr(m(i))
        wsBase.Cells(lastRow, 7).Value = 0#
    Next i
End Sub

' --- resto sin cambios ---

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


Private Sub PrepararHojaReporte(ByVal ws As Worksheet)
    Dim shp As Shape

    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = vbWhite
    On Error GoTo 0
End Sub

Private Sub CrearHojaReporteVisual(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    PrepararHojaReporte ws
    ArmarEncabezadoVisual ws, anio, mesCierre
End Sub

Private Sub ConfigurarCampoPivotSeguro(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    Dim pf As PivotField
    Dim orientacionPrevia As String
    Dim sourceName As String
    Dim caption As String
    Dim posicionPrevia As String
    On Error GoTo EH
    If pt Is Nothing Then Err.Raise vbObjectError + 730, "ConfigurarCampoPivotSeguro", "PivotTable es Nothing."
    If Not PivotFieldExiste(pt, nombreCampo) Then Err.Raise vbObjectError + 731, "ConfigurarCampoPivotSeguro", "No existe el campo '" & nombreCampo & "'."
    Set pf = pt.PivotFields(nombreCampo)
    orientacionPrevia = OrientacionPivotTexto(pf.Orientation)
    sourceName = CStr(pf.SourceName)
    caption = CStr(pf.Caption)
    posicionPrevia = CStr(pf.Position)
    pf.Orientation = orientacion
    If posicion > 0 Then pf.Position = posicion
    Exit Sub
EH:
    Err.Raise Err.Number, "ConfigurarCampoPivotSeguro", _
        "No se pudo asignar orientación al campo '" & nombreCampo & _
        "'. Orientación solicitada: " & OrientacionPivotTexto(orientacion) & _
        ". Orientación actual previa: " & orientacionPrevia & _
        ". SourceName: " & sourceName & _
        ". Caption: " & caption & _
        ". Position previa: " & posicionPrevia & _
        ". Position solicitada: " & CStr(posicion) & _
        ". Err.Number: " & CStr(Err.Number) & _
        ". Detalle: " & Err.Description
End Sub

Private Function ConfigurarCampoMesColumnaConFallback(ByVal pt As PivotTable, ByVal campoMesNombre As String, ByVal campoMesNum As String, ByRef msgFallback As String) As String
    On Error GoTo FallbackMesNum
    ConfigurarCampoPivotSeguro pt, campoMesNombre, xlColumnField, 1
    ConfigurarCampoMesColumnaConFallback = campoMesNombre
    msgFallback = "[PIVOT] MesNombre funcionó como columna."
    Exit Function

FallbackMesNum:
    Dim errNumMesNombre As Long, errDescMesNombre As String
    errNumMesNombre = Err.Number
    errDescMesNombre = Err.Description
    Debug.Print "[PIVOT] MesNombre falló como columna. Probando MesNum."
    On Error GoTo EH
    Err.Clear
    ConfigurarCampoPivotSeguro pt, campoMesNum, xlColumnField, 1
    msgFallback = "[PIVOT] MesNombre falló (" & CStr(errNumMesNombre) & "). MesNum funcionó como columna técnica."
    ConfigurarCampoMesColumnaConFallback = campoMesNum
    Exit Function
EH:
    Err.Raise Err.Number, "ConfigurarCampoMesColumnaConFallback", _
              "MesNombre falló como columna (Err.Number=" & CStr(errNumMesNombre) & ", Detalle=" & errDescMesNombre & _
              "). MesNum también falló (Err.Number=" & CStr(Err.Number) & ", Detalle=" & Err.Description & ")."
End Function

Private Sub NormalizarYValidarMesesBase(ByVal wsBase As Worksheet)
    Dim colMesNum As Long, colMesNombre As Long, lastRow As Long, i As Long
    Dim mesTxt As String, mesNumVal As Variant
    Dim mesesValidos As Object
    Set mesesValidos = MesesValidosDict()

    colMesNum = 5
    colMesNombre = 6
    lastRow = UltimaFilaConDatos(wsBase)

    For i = 2 To lastRow
        mesTxt = LCase$(Trim$(CStr(wsBase.Cells(i, colMesNombre).Value)))
        mesTxt = Replace(mesTxt, "septiembre", "setiembre")
        wsBase.Cells(i, colMesNombre).Value = mesTxt

        If Len(mesTxt) = 0 Then
            Err.Raise vbObjectError + 760, "NormalizarYValidarMesesBase", "MesNombre vacío en fila " & CStr(i) & "."
        End If
        If Not mesesValidos.Exists(mesTxt) Then
            Err.Raise vbObjectError + 761, "NormalizarYValidarMesesBase", "MesNombre inválido ('" & mesTxt & "') en fila " & CStr(i) & "."
        End If

        mesNumVal = wsBase.Cells(i, colMesNum).Value
        If Not IsNumeric(mesNumVal) Then
            Err.Raise vbObjectError + 762, "NormalizarYValidarMesesBase", "MesNum no numérico en fila " & CStr(i) & "."
        End If
        If CLng(mesNumVal) < 1 Or CLng(mesNumVal) > 12 Then
            Err.Raise vbObjectError + 763, "NormalizarYValidarMesesBase", "MesNum fuera de rango (" & CStr(mesNumVal) & ") en fila " & CStr(i) & "."
        End If
    Next i
End Sub

Private Sub DiagnosticarBaseAgregadaPrePivot(ByVal wsBase As Worksheet, ByVal sourceAddress As String)
    Dim colMesNum As Long, colMesNombre As Long, lastRow As Long, lastCol As Long, i As Long
    Dim mesTxt As String, mesNumVal As Variant
    Dim dictMeses As Object
    Dim vaciosMesNombre As Long, invalidosMesNum As Long
    Dim claves As Variant, maxPreview As Long, j As Long
    Set dictMeses = CreateObject("Scripting.Dictionary")

    colMesNum = 5
    colMesNombre = 6
    lastRow = UltimaFilaConDatos(wsBase)
    lastCol = UltimaColConDatos(wsBase)

    For i = 2 To lastRow
        mesTxt = LCase$(Trim$(CStr(wsBase.Cells(i, colMesNombre).Value)))
        If Len(mesTxt) = 0 Then
            vaciosMesNombre = vaciosMesNombre + 1
        ElseIf Not dictMeses.Exists(mesTxt) Then
            dictMeses.Add mesTxt, 1
        End If

        mesNumVal = wsBase.Cells(i, colMesNum).Value
        If (Not IsNumeric(mesNumVal)) Or (CLng(mesNumVal) < 1 Or CLng(mesNumVal) > 12) Then
            invalidosMesNum = invalidosMesNum + 1
        End If
    Next i

    Debug.Print "[PIVOT] Base_Agregada filas=" & CStr(lastRow - 1) & " columnas=" & CStr(lastCol)
    Debug.Print "[PIVOT] SourceData=" & sourceAddress
    Debug.Print "[PIVOT] MesNombre vacíos=" & CStr(vaciosMesNombre)
    Debug.Print "[PIVOT] MesNum inválidos=" & CStr(invalidosMesNum)
    If dictMeses.Count > 0 Then
        claves = dictMeses.Keys
        Debug.Print "[PIVOT] MesNombre únicos=" & Join(claves, ", ")
    Else
        Debug.Print "[PIVOT] MesNombre únicos=(sin valores)"
    End If

    maxPreview = Application.WorksheetFunction.Min(10, lastRow - 1)
    For j = 2 To maxPreview + 1
        Debug.Print "[PIVOT] Row" & CStr(j) & ": " & _
                    CStr(wsBase.Cells(j, 1).Value) & " | " & CStr(wsBase.Cells(j, 2).Value) & " | " & _
                    CStr(wsBase.Cells(j, 3).Value) & " | " & CStr(wsBase.Cells(j, 4).Value) & " | " & _
                    CStr(wsBase.Cells(j, 5).Value) & " | " & CStr(wsBase.Cells(j, 6).Value) & " | " & _
                    CStr(wsBase.Cells(j, 7).Value)
    Next j
End Sub

Private Function MesesValidosDict() As Object
    Dim d As Object, m As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each m In Array("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre")
        d(m) = True
    Next m
    Set MesesValidosDict = d
End Function

Private Function OrientacionPivotTexto(ByVal orientacion As XlPivotFieldOrientation) As String
    Select Case orientacion
        Case xlRowField: OrientacionPivotTexto = "xlRowField"
        Case xlColumnField: OrientacionPivotTexto = "xlColumnField"
        Case xlPageField: OrientacionPivotTexto = "xlPageField"
        Case xlDataField: OrientacionPivotTexto = "xlDataField"
        Case xlHidden: OrientacionPivotTexto = "xlHidden"
        Case Else: OrientacionPivotTexto = CStr(orientacion)
    End Select
End Function


Private Function PivotFieldExiste(ByVal pt As PivotTable, ByVal nombreCampo As String) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nombreCampo)
    PivotFieldExiste = Not pf Is Nothing
    Set pf = Nothing
    On Error GoTo 0
End Function

Private Function ObtenerCampoDisponible(ByVal pt As PivotTable, ByVal candidatos As Variant) As String
    Dim i As Long
    For i = LBound(candidatos) To UBound(candidatos)
        If PivotFieldExiste(pt, CStr(candidatos(i))) Then
            ObtenerCampoDisponible = CStr(candidatos(i))
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 741, "ObtenerCampoDisponible", "No se encontró ninguno de los campos candidatos: " & Join(candidatos, ", ")
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

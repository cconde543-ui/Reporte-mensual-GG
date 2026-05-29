Option Explicit

Private Const LOGO_BPS_PATH As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Logo_BPS.jpg"

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim etapaVisual As String
    etapaVisual = "iniciando"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet
    Dim pivotCacheObj As PivotCache
    Dim pt As PivotTable
    Dim rg As Range
    Dim sourceAddress As String
    Dim campoNivel1 As String, campoNivel2 As String, campoNivel3 As String
    Dim campoMesNombre As String, campoMesNum As String
    Dim campoMesColumnaUsado As String
    Dim pfImporte As PivotField
    Dim encabezadosBase As String, camposPivot As String
    Dim campoActual As String, accionActual As String
    Dim orientacionActual As XlPivotFieldOrientation
    Dim errNumPivot As Long, errDescPivot As String

    On Error GoTo EH

    etapaVisual = "validando objetos de entrada"
    If wbOut Is Nothing Then Err.Raise vbObjectError + 720, "CrearTablaDinamicaOSalidaAgrupada", "Workbook de salida (wbOut) es Nothing."
    If wsBase Is Nothing Then Err.Raise vbObjectError + 721, "CrearTablaDinamicaOSalidaAgrupada", "Hoja base (wsBase) es Nothing."

    etapaVisual = "validando base agregada"
    ValidarBaseAgregada wsBase
    encabezadosBase = EncabezadosBaseAgregada(wsBase)
    NormalizarYValidarMesesBase wsBase
    Set rg = wsBase.Range("A1").CurrentRegion

    etapaVisual = "creando PivotCache"
    sourceAddress = rg.Address(ReferenceStyle:=xlR1C1, External:=True)
    Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceAddress)
    pivotCacheObj.MissingItemsLimit = xlMissingItemsNone

    etapaVisual = "creando hoja de reporte"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets("Ejec. Mensual " & anio).Delete
    Application.DisplayAlerts = True
    On Error GoTo EH

    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = "Ejec. Mensual " & anio
    CrearHojaReporteVisual ws, anio, mesCierre

    etapaVisual = "creando PivotTable"
    Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="ptGG")

    etapaVisual = "refrescando PivotCache y PivotTable"
    pivotCacheObj.Refresh
    pt.RefreshTable

    etapaVisual = "resolviendo campos"
    campoNivel1 = ObtenerCampoDisponible(pt, Array("Nivel_1", "Nivel 1"))
    campoNivel2 = ObtenerCampoDisponible(pt, Array("Nivel_2", "Nivel 2"))
    campoNivel3 = ObtenerCampoDisponible(pt, Array("Nivel_3", "Nivel 3"))
    campoMesNombre = ObtenerCampoDisponible(pt, Array("MesNombre", "Mes Nombre"))
    campoMesNum = ObtenerCampoDisponible(pt, Array("MesNum", "Mes Num"))

    etapaVisual = "configurando filas"
    campoActual = campoNivel1: accionActual = "asignando fila nivel 1": orientacionActual = xlRowField
    ConfigurarCampoPivotSeguro pt, campoNivel1, xlRowField, 1

    campoActual = campoNivel2: accionActual = "asignando fila nivel 2": orientacionActual = xlRowField
    ConfigurarCampoPivotSeguro pt, campoNivel2, xlRowField, 2

    campoActual = campoNivel3: accionActual = "asignando fila nivel 3": orientacionActual = xlRowField
    ConfigurarNivel3ConFallback pt, campoNivel3, 3

    etapaVisual = "configurando columna de mes"
    campoActual = campoMesNombre: accionActual = "asignando columna mes": orientacionActual = xlColumnField
    campoMesColumnaUsado = ConfigurarCampoMesColumnaConFallback(pt, campoMesNombre, campoMesNum)

    etapaVisual = "agregando valores"
    If Not PivotFieldExiste(pt, "Importe") Then
        Err.Raise vbObjectError + 724, "CrearTablaDinamicaOSalidaAgrupada", "No existe el campo 'Importe' en la PivotTable."
    End If
    Set pfImporte = pt.PivotFields("Importe")
    pt.AddDataField pfImporte, "EJECUCIÓN " & anio, xlSum

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
    If Not pt.DataBodyRange Is Nothing Then pt.DataBodyRange.NumberFormat = "#,##0;-#,##0;;@"
    AgregarSlicerFinanciamiento wbOut, ws, pt
    AjustarEncabezadoVisualAlPivot ws, pt, "Informe de Seguimiento Presupuestal " & UCase$(MesesES()(mesCierre - 1)) & " " & anio & " - Ejecución mensual y acumulada"
    OcultarGridlinesHoja ws
    Exit Sub

EH:
    errNumPivot = Err.Number
    errDescPivot = Err.Description
    camposPivot = CamposDisponiblesPivot(pt)

    Dim detallePivot As String
    detallePivot = "Error creando PivotTable. Etapa: " & etapaVisual
    detallePivot = detallePivot & " | campoActual: " & campoActual
    detallePivot = detallePivot & " | accionActual: " & accionActual
    detallePivot = detallePivot & " | orientación solicitada: " & OrientacionPivotTexto(orientacionActual)
    detallePivot = detallePivot & " | Campos disponibles pivot: " & camposPivot
    detallePivot = detallePivot & " | Encabezados Base_Agregada: " & encabezadosBase
    detallePivot = detallePivot & " | Err.Number: " & CStr(errNumPivot)
    detallePivot = detallePivot & " | Err.Description: " & errDescPivot
    Err.Raise errNumPivot, "CrearTablaDinamicaOSalidaAgrupada", detallePivot
End Sub

Public Sub ConstruirBasePorcEjec(ByVal ws As Worksheet, ByVal dictAgg As Object, ByVal dictAsignado As Object, ByVal mesCierre As Long)
    Dim dEj As Object, k As Variant, p() As String, k4 As String, fila As Long, ejec As Double, asig As Double
    Set dEj = CreateObject("Scripting.Dictionary")
    ws.Range("A1:G1").Value = Array("Clasificación", "Tipo", "Concepto", "Ejecutado", "Asignado", "% ejecutado", "Financiamiento")
    For Each k In dictAgg.Keys
        p = Split(CStr(k), "|")
        If CLng(p(4)) <= mesCierre Then
            k4 = p(0) & "|" & p(1) & "|" & p(2) & "|" & p(3)
            If Not dEj.Exists(k4) Then dEj.Add k4, 0#
            dEj(k4) = dEj(k4) + CDbl(dictAgg(k))
        End If
    Next k
    fila = 2
    For Each k In dEj.Keys
        p = Split(CStr(k), "|")
        ejec = CDbl(dEj(k))
        If dictAsignado.Exists(CStr(k)) Then asig = CDbl(dictAsignado(k)) Else asig = 0#
        ws.Cells(fila, 1).Value = p(1): ws.Cells(fila, 2).Value = p(2): ws.Cells(fila, 3).Value = p(3)
        ws.Cells(fila, 4).Value = ejec: ws.Cells(fila, 5).Value = asig
        If asig = 0 Then ws.Cells(fila, 6).Value = 0 Else ws.Cells(fila, 6).Value = ejec / asig
        ws.Cells(fila, 7).Value = p(0)
        fila = fila + 1
    Next k
End Sub

Private Sub ConfigurarNivel3ConFallback(ByVal pt As PivotTable, ByVal campoNivel3 As String, ByVal posicion As Long)
    On Error GoTo FallbackSinPosicion
    ConfigurarCampoPivotSeguro pt, campoNivel3, xlRowField, posicion
    Exit Sub

FallbackSinPosicion:
    Dim errNumPosicion As Long, errDescPosicion As String
    errNumPosicion = Err.Number
    errDescPosicion = Err.Description
    Err.Clear

    On Error GoTo EH
    ConfigurarCampoPivotSeguro pt, campoNivel3, xlRowField, 0
    Exit Sub

EH:
    Dim detalleNivel3 As String
    detalleNivel3 = "Falló configuración de '" & campoNivel3 & "' como fila. "
    detalleNivel3 = detalleNivel3 & "Intento 1 (con posición=" & CStr(posicion) & ") -> Err.Number=" & CStr(errNumPosicion) & ", Err.Description=" & errDescPosicion & ". "
    detalleNivel3 = detalleNivel3 & "Intento 2 (solo Orientation, sin Position) -> Err.Number=" & CStr(Err.Number) & ", Err.Description=" & Err.Description
    Err.Raise Err.Number, "ConfigurarNivel3ConFallback", detalleNivel3
End Sub

Private Function ConfigurarCampoMesColumnaConFallback(ByVal pt As PivotTable, ByVal campoMesNombre As String, ByVal campoMesNum As String) As String
    On Error GoTo FallbackMesNum
    ConfigurarCampoPivotSeguro pt, campoMesNombre, xlColumnField, 1
    ConfigurarCampoMesColumnaConFallback = campoMesNombre
    Exit Function

FallbackMesNum:
    Dim errNumMesNombre As Long, errDescMesNombre As String
    errNumMesNombre = Err.Number
    errDescMesNombre = Err.Description
    Err.Clear

    On Error GoTo EH
    ConfigurarCampoPivotSeguro pt, campoMesNum, xlColumnField, 1
    ConfigurarCampoMesColumnaConFallback = campoMesNum
    Exit Function

EH:
    Dim detalleMesCol As String
    detalleMesCol = "MesNombre falló como columna (Err.Number=" & CStr(errNumMesNombre) & ", Err.Description=" & errDescMesNombre & "). "
    detalleMesCol = detalleMesCol & "MesNum también falló (Err.Number=" & CStr(Err.Number) & ", Err.Description=" & Err.Description & ")."
    Err.Raise Err.Number, "ConfigurarCampoMesColumnaConFallback", detalleMesCol
End Function

Private Sub ConfigurarCampoPivotSeguro(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    Dim pf As PivotField
    Dim orientacionPrevia As XlPivotFieldOrientation
    Dim sourceName As String
    Dim caption As String
    Dim errNum As Long, errDesc As String

    If pt Is Nothing Then Err.Raise vbObjectError + 730, "ConfigurarCampoPivotSeguro", "PivotTable es Nothing."
    If Not PivotFieldExiste(pt, nombreCampo) Then Err.Raise vbObjectError + 731, "ConfigurarCampoPivotSeguro", "No existe el campo '" & nombreCampo & "'."

    Set pf = pt.PivotFields(nombreCampo)
    orientacionPrevia = pf.Orientation
    sourceName = CStr(pf.SourceName)
    caption = CStr(pf.Caption)

    On Error GoTo EH_ORIENTATION
    pf.Orientation = orientacion

    If posicion > 0 Then
        On Error GoTo EH_POSITION
        pf.Position = posicion
    End If

    Exit Sub

EH_ORIENTATION:
    errNum = Err.Number
    errDesc = Err.Description
    Dim detalleOrientation As String
    detalleOrientation = "Error al asignar Orientation. nombreCampo=" & nombreCampo
    detalleOrientation = detalleOrientation & " | orientación solicitada=" & OrientacionPivotTexto(orientacion)
    detalleOrientation = detalleOrientation & " | orientación actual previa=" & OrientacionPivotTexto(orientacionPrevia)
    detalleOrientation = detalleOrientation & " | SourceName=" & sourceName
    detalleOrientation = detalleOrientation & " | Caption=" & caption
    detalleOrientation = detalleOrientation & " | Err.Number=" & CStr(errNum)
    detalleOrientation = detalleOrientation & " | Err.Description=" & errDesc
    Err.Raise errNum, "ConfigurarCampoPivotSeguro", detalleOrientation

EH_POSITION:
    errNum = Err.Number
    errDesc = Err.Description
    Dim detallePosition As String
    detallePosition = "Error al asignar Position. nombreCampo=" & nombreCampo
    detallePosition = detallePosition & " | orientación solicitada=" & OrientacionPivotTexto(orientacion)
    detallePosition = detallePosition & " | orientación actual previa=" & OrientacionPivotTexto(orientacionPrevia)
    detallePosition = detallePosition & " | SourceName=" & sourceName
    detallePosition = detallePosition & " | Caption=" & caption
    detallePosition = detallePosition & " | Position solicitada=" & CStr(posicion)
    detallePosition = detallePosition & " | Err.Number=" & CStr(errNum)
    detallePosition = detallePosition & " | Err.Description=" & errDesc
    Err.Raise errNum, "ConfigurarCampoPivotSeguro", detallePosition
End Sub

Private Sub ArmarEncabezadoVisual(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim mes As String, arrMeses As Variant
    Dim rngBandaSuperior As Range, rngTitulo As Range, rngSubtitulo As Range

    arrMeses = MesesES()
    mes = UCase$(CStr(arrMeses(mesCierre - 1)))

    ws.Rows(1).RowHeight = 50.25
    ws.Rows(2).RowHeight = 15
    ws.Rows(3).RowHeight = 24

    Set rngBandaSuperior = ws.Range("A1:O1")
    Set rngTitulo = ws.Range("A3:O3")
    Set rngSubtitulo = ws.Range("A3:O3")

    rngBandaSuperior.UnMerge
    rngSubtitulo.UnMerge
    rngTitulo.Merge

    rngTitulo.ClearContents
    ws.Range("A2:O2").ClearContents

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
    Set rng = ws.Range("A1:O1")
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "shpBandaAzul"
    shp.Fill.ForeColor.RGB = RGB(0, 84, 147)
    shp.Line.Visible = msoFalse
    shp.Placement = xlMove
    shp.ZOrder msoSendToBack
End Sub

Private Sub InsertarLogoBPS(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim logoH As Double
    Dim logoW As Double
    Dim topPos As Double
    Dim leftPos As Double
    Dim rngBandaSuperior As Range
    Dim rutaLogo As String

    On Error Resume Next
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    rutaLogo = RutaLogoBPSActiva()
    If Len(rutaLogo) = 0 Then Exit Sub

    On Error GoTo EH
    logoH = ws.Rows(1).Height - 6
    Set shp = ws.Shapes.AddPicture(rutaLogo, msoFalse, msoTrue, 0, 0, -1, logoH)
    shp.Name = "imgLogoBPS"
    shp.LockAspectRatio = msoTrue
    logoW = shp.Width

    topPos = ws.Rows(1).Top + (ws.Rows(1).Height - shp.Height) / 2
    Set rngBandaSuperior = ws.Range("A1:O1")
    leftPos = rngBandaSuperior.Left + rngBandaSuperior.Width - logoW - 6
    shp.Top = topPos
    shp.Left = leftPos
    shp.Placement = xlMove
    Exit Sub
EH:
    'Si falla el logo, no romper el reporte.
End Sub

Private Function ArchivoExisteSeguro(ByVal ruta As String) As Boolean
    On Error GoTo Salida

    Dim fso As Object

    If Len(Trim$(ruta)) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    ArchivoExisteSeguro = fso.FileExists(ruta)

Salida:
End Function

Private Function CarpetaPadreLocalOutput() As String
    On Error GoTo Salida

    Dim fso As Object

    If Len(ThisWorkbook.Path) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    CarpetaPadreLocalOutput = fso.GetParentFolderName(ThisWorkbook.Path)

Salida:
End Function

Private Function RutaLogoBPSActiva() As String
    Dim base As String
    Dim padre As String
    Dim candidatos As Collection
    Dim i As Long

    base = ThisWorkbook.Path
    padre = CarpetaPadreLocalOutput()

    Set candidatos = New Collection
    candidatos.Add LOGO_BPS_PATH
    candidatos.Add base & "\Logo_BPS.jpg"
    candidatos.Add base & "\Logo_BPS.png"
    candidatos.Add base & "\Logo BPS.jpg"
    candidatos.Add base & "\Logo BPS.png"
    candidatos.Add base & "\Recursos\Logo_BPS.jpg"
    candidatos.Add base & "\Recursos\Logo_BPS.png"
    candidatos.Add base & "\Imagenes\Logo_BPS.jpg"
    candidatos.Add base & "\Imagenes\Logo_BPS.png"
    candidatos.Add base & "\Imágenes\Logo_BPS.jpg"
    candidatos.Add base & "\Imágenes\Logo_BPS.png"
    candidatos.Add padre & "\Logo_BPS.jpg"
    candidatos.Add padre & "\Logo_BPS.png"
    candidatos.Add padre & "\Logo BPS.jpg"
    candidatos.Add padre & "\Logo BPS.png"
    candidatos.Add padre & "\Recursos\Logo_BPS.jpg"
    candidatos.Add padre & "\Recursos\Logo_BPS.png"
    candidatos.Add padre & "\Imagenes\Logo_BPS.jpg"
    candidatos.Add padre & "\Imagenes\Logo_BPS.png"
    candidatos.Add padre & "\Imágenes\Logo_BPS.jpg"
    candidatos.Add padre & "\Imágenes\Logo_BPS.png"

    For i = 1 To candidatos.Count
        If ArchivoExisteSeguro(CStr(candidatos(i))) Then
            RutaLogoBPSActiva = CStr(candidatos(i))
            Exit Function
        End If
    Next i
End Function

Private Sub OrdenarMesesPivot(ByVal pt As PivotTable, ByVal campoMesNombre As String, ByVal campoMesNum As String)
    Dim pfNom As PivotField, pfNum As PivotField, m As Variant, i As Long
    On Error Resume Next
    Set pfNom = pt.PivotFields(campoMesNombre)
    Set pfNum = pt.PivotFields(campoMesNum)
    On Error GoTo 0
    If pfNom Is Nothing Then Exit Sub

    m = MesesESMin()
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

Private Sub ColapsarPivotInicial(ByVal pt As PivotTable)
    Dim pf As PivotField
    Dim pi As PivotItem
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
    OcultarGridlinesHoja ws
    ws.Columns("A").ColumnWidth = 31
    ArmarEncabezadoVisual ws, anio, mesCierre
    OcultarGridlinesHoja ws
End Sub

Private Sub OcultarGridlinesHoja(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Function NombreObjetoSeguro(ByVal texto As String) As String
    Dim s As String
    s = texto
    s = Replace(s, " ", "_")
    s = Replace(s, "%", "pct")
    s = Replace(s, ".", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, "á", "a")
    s = Replace(s, "é", "e")
    s = Replace(s, "í", "i")
    s = Replace(s, "ó", "o")
    s = Replace(s, "ú", "u")
    s = Replace(s, "ñ", "n")
    NombreObjetoSeguro = s
End Function

Private Sub AgregarSlicerFinanciamiento(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal pt As PivotTable)
    Const CAMPO_FINANCIAMIENTO As String = "Financiamiento"

    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim topPos As Double
    Dim leftPos As Double
    Dim ancho As Double
    Dim alto As Double
    Dim margenIzq As Double
    Dim margenDer As Double
    Dim nombreSlicer As String
    Dim nombreCache As String

    On Error GoTo SalidaSilenciosa

    If wb Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If pt Is Nothing Then Exit Sub
    If Not PivotFieldExiste(pt, CAMPO_FINANCIAMIENTO) Then Exit Sub

    nombreSlicer = "slFinanciamiento_" & NombreObjetoSeguro(ws.Name)
    nombreCache = "scFinanciamiento_" & NombreObjetoSeguro(ws.Name)

    On Error Resume Next
    ws.Shapes(nombreSlicer).Delete
    On Error GoTo SalidaSilenciosa

    Set sc = Nothing
    On Error Resume Next
    Set sc = wb.SlicerCaches.Add2(pt, CAMPO_FINANCIAMIENTO, nombreCache)
    If sc Is Nothing Then
        Err.Clear
        Set sc = wb.SlicerCaches.Add(pt, CAMPO_FINANCIAMIENTO)
    End If
    On Error GoTo SalidaSilenciosa
    If sc Is Nothing Then Exit Sub

    topPos = ws.Range("A5").Top
    margenIzq = 3
    margenDer = 10
    leftPos = ws.Range("A5").Left + margenIzq
    ancho = ws.Range("A:A").Width - margenIzq - margenDer
    If ancho < 80 Then ancho = ws.Range("A:A").Width
    alto = ws.Range("A5:A15").Height

    Set sl = sc.Slicers.Add(ws, , nombreSlicer, "Financiamiento", topPos, leftPos, ancho, alto)

    With sl
        .Caption = "Financiamiento"
        .DisplayHeader = True
        .Top = topPos
        .Left = leftPos
        .Width = ancho
        .Height = ws.Range("A5:A15").Height
        On Error Resume Next
        .RowHeight = 16
        .NumberOfColumns = 1
        On Error GoTo SalidaSilenciosa
    End With

SalidaSilenciosa:
End Sub


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

        If Len(mesTxt) = 0 Then Err.Raise vbObjectError + 760, "NormalizarYValidarMesesBase", "MesNombre vacío en fila " & CStr(i) & "."
        If Not mesesValidos.Exists(mesTxt) Then Err.Raise vbObjectError + 761, "NormalizarYValidarMesesBase", "MesNombre inválido ('" & mesTxt & "') en fila " & CStr(i) & "."

        mesNumVal = wsBase.Cells(i, colMesNum).Value
        If Not IsNumeric(mesNumVal) Then Err.Raise vbObjectError + 762, "NormalizarYValidarMesesBase", "MesNum no numérico en fila " & CStr(i) & "."
        If CLng(mesNumVal) < 1 Or CLng(mesNumVal) > 12 Then Err.Raise vbObjectError + 763, "NormalizarYValidarMesesBase", "MesNum fuera de rango (" & CStr(mesNumVal) & ") en fila " & CStr(i) & "."
    Next i
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

Public Sub CrearHojaPorcEjecucion(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable, rg As Range

    On Error GoTo EH

    etapaVisual = "validando base % ejecución"
    ValidarBasePorcEjec wsBase

    etapaVisual = "creando hoja % ejecución"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets("% ejecución " & anio).Delete
    Application.DisplayAlerts = True
    On Error GoTo EH
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = "% ejecución " & anio
    PrepararHojaReporte ws
    OcultarGridlinesHoja ws
    ws.Columns("A").ColumnWidth = 31

    Set rg = wsBase.Range("A1").CurrentRegion
    etapaVisual = "creando PivotCache % ejecución"
    Set pc = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg.Address(ReferenceStyle:=xlR1C1, External:=True))
    etapaVisual = "creando PivotTable % ejecución"
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="ptPorcEjecGG")

    etapaVisual = "configurando filas % ejecución"
    ConfigurarCampoPivotSeguro pt, "Clasificación", xlRowField, 1
    ConfigurarCampoPivotSeguro pt, "Tipo", xlRowField, 2
    ConfigurarCampoPivotSeguro pt, "Concepto", xlRowField, 3
    etapaVisual = "validando campos de valores % ejecución"
    If Not PivotFieldExiste(pt, "Ejecutado") Then
        Err.Raise vbObjectError + 1301, "CrearHojaPorcEjecucion", "No existe el campo 'Ejecutado' en Base_Porc_Ejec."
    End If
    If Not PivotFieldExiste(pt, "Asignado") Then
        Err.Raise vbObjectError + 1302, "CrearHojaPorcEjecucion", "No existe el campo 'Asignado' en Base_Porc_Ejec."
    End If

    etapaVisual = "agregando valores % ejecución"
    Call AgregarDataFieldSeguro(pt, "Ejecutado", "Ejecutado ", xlSum, "#,##0")
    Call AgregarDataFieldSeguro(pt, "Asignado", "Asignado " & CStr(anio), xlSum, "#,##0")

    etapaVisual = "creando campo calculado % ejecución"
    Call AsegurarCampoCalculadoPctEjec(pt)

    etapaVisual = "agregando valor % ejecución"
    Call AgregarDataFieldCalculadoSeguro(pt, "PctEjec", " % ejec.", xlSum, "0.0%;-0.0%;;@")

    pt.DisplayErrorString = True
    pt.ErrorString = ""

    pt.RefreshTable

    If PivotFieldExiste(pt, "Financiamiento") Then
        On Error Resume Next
        pt.PivotFields("Financiamiento").Orientation = xlHidden
        On Error GoTo EH
    End If

    pt.RowAxisLayout xlCompactRow
    pt.DisplayFieldCaptions = False
    pt.ColumnGrand = True: pt.RowGrand = True
    pt.NullString = "": pt.DisplayNullString = True
    ColapsarPivotInicial pt

    etapaVisual = "agregando slicer % ejecución"
    AgregarSlicerFinanciamiento wbOut, ws, pt
    etapaVisual = "ajustando encabezado % ejecución"
    AjustarEncabezadoVisualAlPivot ws, pt, "Informe de Seguimiento Presupuestal " & UCase$(MesesES()(mesCierre - 1)) & " " & anio & " - % de ejec. acumulada sobre la asignación presupuestal"
    OcultarGridlinesHoja ws
    Exit Sub
EH:
    Dim detalleHojaPct As String
    detalleHojaPct = "Error creando hoja % ejecución. Etapa: " & etapaVisual
    detalleHojaPct = detalleHojaPct & " | Campos disponibles pivot: " & CamposDisponiblesPivot(pt)
    detalleHojaPct = detalleHojaPct & " | Err.Number: " & CStr(Err.Number)
    detalleHojaPct = detalleHojaPct & " | Err.Description: " & Err.Description
    Err.Raise Err.Number, "CrearHojaPorcEjecucion", detalleHojaPct
End Sub

Private Sub AsegurarCampoCalculadoPctEjec(ByVal pt As PivotTable)
    Const CAMPO_CALC As String = "PctEjec"

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1450, "AsegurarCampoCalculadoPctEjec", "PivotTable es Nothing."
    End If

    If Not PivotFieldExiste(pt, "Ejecutado") Then
        Err.Raise vbObjectError + 1451, "AsegurarCampoCalculadoPctEjec", "No existe el campo fuente 'Ejecutado'."
    End If

    If Not PivotFieldExiste(pt, "Asignado") Then
        Err.Raise vbObjectError + 1452, "AsegurarCampoCalculadoPctEjec", "No existe el campo fuente 'Asignado'."
    End If

    If Not CampoCalculadoExiste(pt, CAMPO_CALC) Then
        On Error GoTo EH_ADD

        pt.CalculatedFields.Add CAMPO_CALC, "=Ejecutado/Asignado", True

        On Error GoTo 0
    End If

    If Not CampoCalculadoExiste(pt, CAMPO_CALC) Then
        Err.Raise vbObjectError + 1453, "AsegurarCampoCalculadoPctEjec", "No se pudo crear el campo calculado interno 'PctEjec'."
    End If

    Exit Sub

EH_ADD:
    Dim detallePctEjec As String
    detallePctEjec = "Falló pt.CalculatedFields.Add para 'PctEjec'. Fórmula usada: =Ejecutado/Asignado | Err.Number=" & CStr(Err.Number)
    detallePctEjec = detallePctEjec & " | Err.Description=" & Err.Description
    Err.Raise Err.Number, "AsegurarCampoCalculadoPctEjec", detallePctEjec
End Sub

Private Sub AsegurarCampoCalculadoPctVariacion( _
    ByVal pt As PivotTable, _
    ByVal campoEjecutadoActual As String, _
    ByVal campoEjecutadoAnteriorActualizado As String)

    Const CAMPO_CALC As String = "PctVariacion"

    Dim formulaCalc As String

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1470, "AsegurarCampoCalculadoPctVariacion", "PivotTable es Nothing."
    End If

    If Not PivotFieldExiste(pt, campoEjecutadoActual) Then
        Err.Raise vbObjectError + 1471, "AsegurarCampoCalculadoPctVariacion", "No existe el campo fuente actual '" & campoEjecutadoActual & "'."
    End If

    If Not PivotFieldExiste(pt, campoEjecutadoAnteriorActualizado) Then
        Err.Raise vbObjectError + 1472, "AsegurarCampoCalculadoPctVariacion", "No existe el campo fuente anterior actualizado '" & campoEjecutadoAnteriorActualizado & "'."
    End If

    formulaCalc = "=('" & campoEjecutadoActual & "'-'" & campoEjecutadoAnteriorActualizado & "')/'" & campoEjecutadoAnteriorActualizado & "'"

    If Not CampoCalculadoExiste(pt, CAMPO_CALC) Then
        On Error GoTo EH_ADD
        pt.CalculatedFields.Add CAMPO_CALC, formulaCalc, True
        On Error GoTo 0
    End If

    If Not CampoCalculadoExiste(pt, CAMPO_CALC) Then
        Err.Raise vbObjectError + 1473, "AsegurarCampoCalculadoPctVariacion", "No se pudo crear el campo calculado interno 'PctVariacion'. Fórmula: " & formulaCalc
    End If

    Exit Sub

EH_ADD:
    Dim detallePctVar As String
    detallePctVar = "Falló pt.CalculatedFields.Add para 'PctVariacion'. "
    detallePctVar = detallePctVar & "Campo actual: '" & campoEjecutadoActual & "'. "
    detallePctVar = detallePctVar & "Campo anterior actualizado: '" & campoEjecutadoAnteriorActualizado & "'. "
    detallePctVar = detallePctVar & "Fórmula usada: " & formulaCalc
    detallePctVar = detallePctVar & " | Err.Number=" & CStr(Err.Number)
    detallePctVar = detallePctVar & " | Err.Description=" & Err.Description
    Err.Raise Err.Number, "AsegurarCampoCalculadoPctVariacion", detallePctVar
End Sub

Private Sub AgregarDataFieldSeguro( _
    ByVal pt As PivotTable, _
    ByVal sourceFieldName As String, _
    ByVal caption As String, _
    ByVal funcion As XlConsolidationFunction, _
    ByVal formatoNumero As String)

    Dim nAntes As Long
    Dim pfSource As PivotField
    Dim pfData As PivotField

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1400, "AgregarDataFieldSeguro", "PivotTable es Nothing."
    End If

    If Not PivotFieldExiste(pt, sourceFieldName) Then
        Err.Raise vbObjectError + 1401, "AgregarDataFieldSeguro", "No existe el campo fuente '" & sourceFieldName & "'."
    End If

    Set pfSource = pt.PivotFields(sourceFieldName)

    nAntes = pt.DataFields.Count

    pt.AddDataField pfSource, caption, funcion

    If pt.DataFields.Count <= nAntes Then
        Err.Raise vbObjectError + 1402, "AgregarDataFieldSeguro", "No se agregó el campo de valores '" & caption & "'."
    End If

    Set pfData = pt.DataFields(pt.DataFields.Count)

    If Len(formatoNumero) > 0 Then
        pfData.NumberFormat = formatoNumero
    End If

End Sub

Private Function CampoCalculadoExiste(ByVal pt As PivotTable, ByVal nombreCampo As String) As Boolean
    Dim pf As PivotField

    If pt Is Nothing Then Exit Function

    On Error Resume Next
    Set pf = pt.PivotFields(nombreCampo)
    CampoCalculadoExiste = Not pf Is Nothing
    Set pf = Nothing
    On Error GoTo 0
End Function

Private Sub AgregarDataFieldCalculadoSeguro( _
    ByVal pt As PivotTable, _
    ByVal calculatedFieldName As String, _
    ByVal caption As String, _
    ByVal funcion As XlConsolidationFunction, _
    ByVal formatoNumero As String)

    Dim nAntes As Long
    Dim pfSource As PivotField
    Dim pfData As PivotField

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1460, "AgregarDataFieldCalculadoSeguro", "PivotTable es Nothing."
    End If

    If Not CampoCalculadoExiste(pt, calculatedFieldName) Then
        Err.Raise vbObjectError + 1461, "AgregarDataFieldCalculadoSeguro", "No existe el campo calculado '" & calculatedFieldName & "'."
    End If

    Set pfSource = pt.PivotFields(calculatedFieldName)

    nAntes = pt.DataFields.Count

    pt.AddDataField pfSource, caption, funcion

    If pt.DataFields.Count <= nAntes Then
        Err.Raise vbObjectError + 1462, "AgregarDataFieldCalculadoSeguro", "No se agregó el campo calculado como valor '" & caption & "'."
    End If

    Set pfData = pt.DataFields(pt.DataFields.Count)

    If Len(formatoNumero) > 0 Then
        pfData.NumberFormat = formatoNumero
    End If

End Sub


Private Sub ValidarBasePorcEjec(ByVal wsBase As Worksheet)
    Dim headers As Variant
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim vE As Variant
    Dim vA As Variant

    If wsBase Is Nothing Then
        Err.Raise vbObjectError + 1310, "ValidarBasePorcEjec", "La hoja Base_Porc_Ejec es Nothing."
    End If

    lastRow = UltimaFilaConDatos(wsBase)
    lastCol = UltimaColConDatos(wsBase)
    If lastRow < 2 Then
        Err.Raise vbObjectError + 1311, "ValidarBasePorcEjec", "Base_Porc_Ejec debe tener al menos 2 filas (encabezado + datos)."
    End If
    If lastCol < 7 Then
        Err.Raise vbObjectError + 1312, "ValidarBasePorcEjec", "Base_Porc_Ejec debe tener al menos 7 columnas."
    End If

    headers = Array("Clasificación", "Tipo", "Concepto", "Ejecutado", "Asignado", "% ejecutado", "Financiamiento")
    For i = 0 To 6
        If CStr(wsBase.Cells(1, i + 1).Value) <> CStr(headers(i)) Then
            Err.Raise vbObjectError + 1313, "ValidarBasePorcEjec", "Encabezado inválido en columna " & CStr(i + 1) & ". Esperado: '" & CStr(headers(i)) & "'. Encontrado: '" & CStr(wsBase.Cells(1, i + 1).Value) & "'."
        End If
    Next i

    For i = 2 To lastRow
        vE = wsBase.Cells(i, 4).Value
        vA = wsBase.Cells(i, 5).Value

        If Len(Trim$(CStr(vE))) > 0 And Not IsNumeric(vE) Then
            Err.Raise vbObjectError + 1314, "ValidarBasePorcEjec", "Valor no numérico en columna Ejecutado, fila " & CStr(i) & "."
        End If
        If Len(Trim$(CStr(vA))) > 0 And Not IsNumeric(vA) Then
            Err.Raise vbObjectError + 1315, "ValidarBasePorcEjec", "Valor no numérico en columna Asignado, fila " & CStr(i) & "."
        End If
    Next i
End Sub

Private Sub AjustarEncabezadoVisualAlPivot(ByVal ws As Worksheet, ByVal pt As PivotTable, ByVal titulo As String)
    Dim lastPivotCol As Long
    Dim rngBanda As Range, rngTitulo As Range, shp As Shape
    If pt Is Nothing Then Exit Sub
    lastPivotCol = pt.TableRange1.Column + pt.TableRange1.Columns.Count - 1
    Set rngBanda = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastPivotCol))
    Set rngTitulo = ws.Range(ws.Cells(3, 1), ws.Cells(3, lastPivotCol))

    ws.Rows(1).RowHeight = 50.25: ws.Rows(2).RowHeight = 15
    If InStr(1, ws.Name, "Comparativo", vbTextCompare) > 0 Then
        ws.Rows(3).RowHeight = 33.75
    ElseIf InStr(1, ws.Name, "% ejecución", vbTextCompare) > 0 Then
        ws.Rows(3).RowHeight = 42
    Else
        ws.Rows(3).RowHeight = 24
    End If
    If Not Intersect(rngTitulo, pt.TableRange2) Is Nothing Then
        Dim detalleEncabezado As String
        detalleEncabezado = "El rango del título " & rngTitulo.Address(False, False)
        detalleEncabezado = detalleEncabezado & " intersecta con la tabla dinámica " & pt.TableRange2.Address(False, False)
        detalleEncabezado = detalleEncabezado & ". No se puede fusionar el encabezado porque afectaría la PivotTable."
        Err.Raise vbObjectError + 1501, "AjustarEncabezadoVisualAlPivot", detalleEncabezado
    End If
    rngTitulo.UnMerge: rngTitulo.Merge
    rngTitulo.Value = titulo
    With rngTitulo
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True
        .Font.Name = "Calibri Light": .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(0, 32, 96)
    End With
    With rngTitulo.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 32, 96)
    End With

    On Error Resume Next
    ws.Shapes("shpBandaAzul").Delete
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, rngBanda.Left, rngBanda.Top, rngBanda.Width, rngBanda.Height)
    shp.Name = "shpBandaAzul": shp.Fill.ForeColor.RGB = RGB(0, 84, 147): shp.Line.Visible = msoFalse: shp.ZOrder msoSendToBack

    InsertarLogoBPS_EnRango ws, rngBanda
End Sub

Private Sub InsertarLogoBPS_EnRango(ByVal ws As Worksheet, ByVal rngBanda As Range)
    Dim shp As Shape
    Dim logoH As Double
    Dim rutaLogo As String

    If ws Is Nothing Then Exit Sub
    If rngBanda Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    rutaLogo = RutaLogoBPSActiva()
    If Len(rutaLogo) = 0 Then Exit Sub

    On Error GoTo EH

    logoH = ws.Rows(1).Height - 6
    Set shp = ws.Shapes.AddPicture(rutaLogo, msoFalse, msoTrue, 0, 0, -1, logoH)

    shp.Name = "imgLogoBPS"
    shp.LockAspectRatio = msoTrue
    shp.Top = ws.Rows(1).Top + (ws.Rows(1).Height - shp.Height) / 2
    shp.Left = rngBanda.Left + rngBanda.Width - shp.Width - 6
    shp.Placement = xlMove

    Exit Sub
EH:
    'Si falla el logo, no romper el reporte.
End Sub


Public Function ObtenerWorksheetSeguro(ByVal wb As Workbook, ByVal nombreHoja As String) As Worksheet
    On Error Resume Next
    Set ObtenerWorksheetSeguro = wb.Worksheets(nombreHoja)
    On Error GoTo 0
End Function

Public Sub OrdenarHojasVisualesReporteGG( _
    ByVal wbOut As Workbook, _
    ByVal anioActual As Long, _
    ByVal anioComparativo As Long)

    On Error GoTo EH

    Dim wsEjec As Worksheet
    Dim wsComp As Worksheet
    Dim wsPorc As Worksheet
    Dim nombreEjec As String
    Dim nombreComp As String
    Dim nombrePorc As String

    If wbOut Is Nothing Then
        Err.Raise vbObjectError + 2300, "OrdenarHojasVisualesReporteGG", "wbOut es Nothing."
    End If

    nombreEjec = "Ejec. Mensual " & CStr(anioActual)
    nombreComp = "Comparativo " & CStr(anioActual) & " vs. " & CStr(anioComparativo)
    nombrePorc = "% ejecución " & CStr(anioActual)

    Set wsEjec = ObtenerWorksheetSeguro(wbOut, nombreEjec)
    Set wsComp = ObtenerWorksheetSeguro(wbOut, nombreComp)
    Set wsPorc = ObtenerWorksheetSeguro(wbOut, nombrePorc)

    If wsEjec Is Nothing Then Err.Raise vbObjectError + 2301, "OrdenarHojasVisualesReporteGG", "No existe la hoja: " & nombreEjec
    If wsComp Is Nothing Then Err.Raise vbObjectError + 2302, "OrdenarHojasVisualesReporteGG", "No existe la hoja: " & nombreComp
    If wsPorc Is Nothing Then Err.Raise vbObjectError + 2303, "OrdenarHojasVisualesReporteGG", "No existe la hoja: " & nombrePorc

    wsEjec.Visible = xlSheetVisible
    wsComp.Visible = xlSheetVisible
    wsPorc.Visible = xlSheetVisible

    wsEjec.Move Before:=wbOut.Worksheets(1)
    wsComp.Move After:=wsEjec
    wsPorc.Move After:=wsComp

    Exit Sub

EH:
    Err.Raise Err.Number, "OrdenarHojasVisualesReporteGG", Err.Description
End Sub

Public Sub CrearHojaComparativoAnual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anioActual As Long, ByVal anioComparativo As Long, ByVal mesCierre As Long, ByRef etapaVisual As String)
    Dim ws As Worksheet, ptCache As PivotCache, pt As PivotTable, nombreHoja As String
    Dim rngBase As Range, titulo As String
    Dim campoActual As String
    Dim campoAnteriorAct As String
    nombreHoja = "Comparativo " & anioActual & " vs. " & anioComparativo
    On Error Resume Next
    Application.DisplayAlerts = False
    wbOut.Worksheets(nombreHoja).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(1))
    ws.Name = nombreHoja
    PrepararHojaReporte ws
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 31
    Set rngBase = wsBase.Range("A1").CurrentRegion
    Set ptCache = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngBase)
    Set pt = ws.PivotTables.Add(PivotCache:=ptCache, TableDestination:=ws.Range("B5"), TableName:="ptComparativoGG")
    campoActual = "Ejecutado " & CStr(anioActual) & "."
    campoAnteriorAct = "Ejecutado " & CStr(anioComparativo) & " a valores " & CStr(anioActual) & "."
    With pt
        ConfigurarCampoPivotSeguro pt, "Clasificación", xlRowField, 1
        ConfigurarCampoPivotSeguro pt, "Tipo", xlRowField, 2
        ConfigurarCampoPivotSeguro pt, "concepto", xlRowField, 3
    End With
    Call AgregarDataFieldSeguro(pt, campoActual, "Suma de " & campoActual, xlSum, "#,##0;-#,##0;;@")
    Call AgregarDataFieldSeguro(pt, campoAnteriorAct, "Suma de " & campoAnteriorAct, xlSum, "#,##0;-#,##0;;@")
    Call AsegurarCampoCalculadoPctVariacion(pt, campoActual, campoAnteriorAct)
    AgregarDataFieldCalculadoSeguro pt, "PctVariacion", "% de variación ", xlSum, "0.0%;-0.0%;;@"
    pt.DataFields(1).NumberFormat = "#,##0;-#,##0;;@"
    pt.DataFields(2).NumberFormat = "#,##0;-#,##0;;@"
    pt.RowAxisLayout xlCompactRow
    pt.DisplayFieldCaptions = False
    pt.ColumnGrand = True: pt.RowGrand = True
    pt.NullString = "": pt.DisplayNullString = True
    pt.DisplayErrorString = True: pt.ErrorString = ""
    ColapsarPivotInicial pt
    AgregarSlicerFinanciamiento wbOut, ws, pt
    titulo = "Informe de Seguimiento Presupuestal " & UCase$(MesesES()(mesCierre - 1)) & " " & anioActual & " - Ejecución acumulada y variación anual" & SufijoUnidadTitulo()
    AjustarEncabezadoVisualAlPivot ws, pt, titulo
End Sub

Public Sub CrearArchivoControlReporteGG(ByVal rutaControl As String, ByVal anioActual As Long, ByVal anioComparativo As Long, ByVal mesCierre As Long, ByVal dictControlEjecMensual As Object, ByVal dictControlCompActualPorClave As Object, ByVal dictControlCompAnteriorPorClave As Object, ByVal dictControlAsignadoPorClave As Object, ByVal dictControlEjecPorClave As Object)
    On Error GoTo EH

    Dim wbControl As Workbook
    Dim etapa As String

    etapa = "creando workbook control"
    Set wbControl = Workbooks.Add(xlWBATWorksheet)

    etapa = "creando hoja control ejecución mensual"
    CrearHojaControlEjecucionMensual wbControl, dictControlEjecMensual, anioActual

    etapa = "creando hoja control comparativo"
    CrearHojaControlComparativo wbControl, dictControlCompActualPorClave, dictControlCompAnteriorPorClave, anioActual, anioComparativo

    etapa = "creando hoja control % ejecución"
    CrearHojaControlPorcEjecucion wbControl, dictControlAsignadoPorClave, dictControlEjecPorClave, anioActual

    etapa = "eliminando hoja inicial vacía"
    Application.DisplayAlerts = False
    If wbControl.Worksheets.Count > 3 Then wbControl.Worksheets(1).Delete
    Application.DisplayAlerts = True

    etapa = "guardando archivo control"
    Application.DisplayAlerts = False
    wbControl.SaveAs Filename:=rutaControl, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    etapa = "cerrando archivo control"
    CerrarWorkbookSeguro wbControl, False, True
    Exit Sub

EH:
    Dim n As Long
    Dim d As String
    Dim src As String

    n = Err.Number
    d = Err.Description
    src = Err.Source

    On Error Resume Next
    Application.DisplayAlerts = True
    CerrarWorkbookSeguro wbControl, False
    On Error GoTo 0

    Err.Raise n, "CrearArchivoControlReporteGG", _
        "Error generando archivo de control." & vbCrLf & _
        "Etapa: " & etapa & vbCrLf & _
        "Ruta: " & rutaControl & vbCrLf & _
        "Err.Number original: " & CStr(n) & vbCrLf & _
        "Err.Source original: " & src & vbCrLf & _
        "Err.Description original: " & d
End Sub

Private Sub FormatearHojaControlBase(ByVal ws As Worksheet, ByVal filaHeader As Long, ByVal ultimaCol As Long, ByVal fmtImporteCols As String, Optional ByVal fmtPctCols As String = "")
    Dim lr As Long
    lr = UltimaFilaConDatos(ws)
    ws.Rows(filaHeader).Font.Bold = True
    ws.Range(ws.Cells(filaHeader, 1), ws.Cells(lr, ultimaCol)).AutoFilter
    On Error Resume Next
    ws.Parent.Activate
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Columns.AutoFit
    AplicarFormatoColumnasControl ws, fmtImporteCols, filaHeader + 1, lr, "#,##0.00"
    AplicarFormatoColumnasControl ws, fmtPctCols, filaHeader + 1, lr, "0.00%"
End Sub

Private Sub AplicarFormatoColumnasControl(ByVal ws As Worksheet, ByVal columnas As String, ByVal filaInicio As Long, ByVal filaFin As Long, ByVal formato As String)
    Dim partes() As String
    Dim limites() As String
    Dim i As Long
    Dim col As String

    If Len(Trim$(columnas)) = 0 Then Exit Sub
    If filaFin < filaInicio Then Exit Sub

    partes = Split(columnas, ",")
    For i = LBound(partes) To UBound(partes)
        col = Trim$(partes(i))
        If Len(col) > 0 Then
            If InStr(1, col, ":", vbTextCompare) > 0 Then
                limites = Split(col, ":")
                ws.Range(Trim$(limites(0)) & filaInicio & ":" & Trim$(limites(1)) & filaFin).NumberFormat = formato
            Else
                ws.Range(col & filaInicio & ":" & col & filaFin).NumberFormat = formato
            End If
        End If
    Next i
End Sub


Private Function NombreHojaSeguro(ByVal nombreBase As String) As String
    Dim s As String

    s = nombreBase
    s = Replace(s, ":", "-")
    s = Replace(s, "\", "-")
    s = Replace(s, "/", "-")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "(")
    s = Replace(s, "]", ")")

    If Len(s) > 31 Then s = Left$(s, 31)

    NombreHojaSeguro = s
End Function
Private Sub CrearHojaControlEjecucionMensual(ByVal wb As Workbook, ByVal d As Object, ByVal anio As Long)
    Dim ws As Worksheet, k As Variant, it As Object, f As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = NombreHojaSeguro("Ctrl_Ejec_" & anio)
    ws.Range("A1:K1").Value = Array("MesNum", "Mes", "Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Ejecutado", "Cantidad líneas origen", "Primera fila origen", "Última fila origen")
    f = 2
    For Each k In d.Keys
        Set it = d(k)
        ws.Cells(f, 1).Value = it("mesNum"): ws.Cells(f, 2).Value = it("mesNombre"): ws.Cells(f, 3).Value = it("clave")
        ws.Cells(f, 4).Value = it("financiamiento"): ws.Cells(f, 5).Value = it("nivel1"): ws.Cells(f, 6).Value = it("nivel2"): ws.Cells(f, 7).Value = it("nivel3")
        ws.Cells(f, 8).Value = it("ejecutado"): ws.Cells(f, 9).Value = it("cantidadLineas"): ws.Cells(f, 10).Value = it("primeraFilaOrigen"): ws.Cells(f, 11).Value = it("ultimaFilaOrigen")
        f = f + 1
    Next k
    ws.Cells(f, 7).Value = "TOTAL": ws.Cells(f, 8).Formula = "=SUM(H2:H" & f - 1 & ")": ws.Rows(f).Font.Bold = True
    FormatearHojaControlBase ws, 1, 11, "H"
End Sub

Private Sub CrearHojaControlComparativo(ByVal wb As Workbook, ByVal dAct As Object, ByVal dAnt As Object, ByVal anioActual As Long, ByVal anioComp As Long)
    Dim ws As Worksheet, keys As Object, k As Variant, f As Long, a As Variant, b As Variant, ejA As Double, ejB As Double, ratio As Double, ejBAct As Double
    Dim fin As String, n1 As String, n2 As String, n3 As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = NombreHojaSeguro("Ctrl_Comp_" & anioActual & "_" & anioComp)
    ws.Range("A1:S1").Value = Array("Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Ejecutado " & anioActual, "Ejecutado " & anioComp & " anterior original", "Indice", "Archivo índice", "Periodo índice base", "Valor índice base", "Periodo índice destino", "Valor índice destino", "Ratio actualización", "Ejecutado " & anioComp & " anterior actualizado a valores " & anioActual, "Diferencia", "% variación", "Cantidad líneas " & anioActual, "Cantidad líneas " & anioComp)
    Set keys = CreateObject("Scripting.Dictionary")
    For Each k In dAct.Keys: keys(k) = True: Next k
    For Each k In dAnt.Keys: keys(k) = True: Next k
    f = 2
    For Each k In keys.Keys
        ejA = 0#: ejB = 0#: ratio = 0#: ejBAct = 0#: fin = "": n1 = "": n2 = "": n3 = ""
        If dAct.Exists(k) Then
            a = dAct(k): ejA = CDbl(a(0))
            fin = CStr(a(2)): n1 = CStr(a(3)): n2 = CStr(a(4)): n3 = CStr(a(5))
        End If
        If dAnt.Exists(k) Then
            b = dAnt(k): ejB = CDbl(b(0)): ratio = CDbl(b(12)): ejBAct = ejB * ratio
            If Len(fin) = 0 Then fin = CStr(b(2)): n1 = CStr(b(3)): n2 = CStr(b(4)): n3 = CStr(b(5))
        End If
        ws.Cells(f, 1).Value = k: ws.Cells(f, 2).Value = fin: ws.Cells(f, 3).Value = n1: ws.Cells(f, 4).Value = n2: ws.Cells(f, 5).Value = n3
        ws.Cells(f, 6).Value = ejA: ws.Cells(f, 7).Value = ejB
        If dAnt.Exists(k) Then ws.Cells(f, 8).Resize(1, 6).Value = Array(b(6), b(7), b(8), b(9), b(10), b(11))
        ws.Cells(f, 14).Value = ratio: ws.Cells(f, 15).Value = ejBAct: ws.Cells(f, 16).Value = ejA - ejBAct
        If ejBAct <> 0 Then ws.Cells(f, 17).Value = (ejA - ejBAct) / ejBAct
        If dAct.Exists(k) Then ws.Cells(f, 18).Value = a(1) Else ws.Cells(f, 18).Value = 0
        If dAnt.Exists(k) Then ws.Cells(f, 19).Value = b(1) Else ws.Cells(f, 19).Value = 0
        f = f + 1
    Next k
    ws.Cells(f, 5).Value = "TOTAL": ws.Cells(f, 6).Formula = "=SUM(F2:F" & f - 1 & ")": ws.Cells(f, 7).Formula = "=SUM(G2:G" & f - 1 & ")"
    ws.Cells(f, 15).Formula = "=SUM(O2:O" & f - 1 & ")": ws.Cells(f, 16).Formula = "=F" & f & "-O" & f
    ws.Cells(f, 17).Formula = "=IF(O" & f & "=0,0,P" & f & "/O" & f & ")": ws.Rows(f).Font.Bold = True
    FormatearHojaControlBase ws, 1, 19, "F:G,K:M,O:P", "N,Q"
End Sub

Private Sub CrearHojaControlPorcEjecucion(ByVal wb As Workbook, ByVal dAsig As Object, ByVal dEj As Object, ByVal anio As Long)
    Dim ws As Worksheet, keys As Object, k As Variant, f As Long, a As Variant, e As Variant, asig As Double, ejec As Double
    Dim fin As String, n1 As String, n2 As String, n3 As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = NombreHojaSeguro("Ctrl_PctEjec_" & anio)
    ws.Range("A1:J1").Value = Array("Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Asignado", "Ejecutado", "% ejecución", "Cantidad líneas asignado", "Cantidad líneas ejecución")
    Set keys = CreateObject("Scripting.Dictionary")
    For Each k In dAsig.Keys: keys(k) = True: Next k
    For Each k In dEj.Keys: keys(k) = True: Next k
    f = 2
    For Each k In keys.Keys
        asig = 0#: ejec = 0#: fin = "": n1 = "": n2 = "": n3 = ""
        If dAsig.Exists(k) Then
            a = dAsig(k): asig = CDbl(a(0))
            fin = CStr(a(2)): n1 = CStr(a(3)): n2 = CStr(a(4)): n3 = CStr(a(5))
        End If
        If dEj.Exists(k) Then
            e = dEj(k): ejec = CDbl(e(0))
            If Len(fin) = 0 Then fin = CStr(e(2)): n1 = CStr(e(3)): n2 = CStr(e(4)): n3 = CStr(e(5))
        End If
        ws.Cells(f, 1).Value = k: ws.Cells(f, 2).Value = fin: ws.Cells(f, 3).Value = n1: ws.Cells(f, 4).Value = n2: ws.Cells(f, 5).Value = n3
        ws.Cells(f, 6).Value = asig: ws.Cells(f, 7).Value = ejec
        If asig <> 0 Then ws.Cells(f, 8).Value = ejec / asig
        If dAsig.Exists(k) Then ws.Cells(f, 9).Value = a(1) Else ws.Cells(f, 9).Value = 0
        If dEj.Exists(k) Then ws.Cells(f, 10).Value = e(1) Else ws.Cells(f, 10).Value = 0
        f = f + 1
    Next k
    ws.Cells(f, 5).Value = "TOTAL": ws.Cells(f, 6).Formula = "=SUM(F2:F" & f - 1 & ")": ws.Cells(f, 7).Formula = "=SUM(G2:G" & f - 1 & ")": ws.Cells(f, 8).Formula = "=IF(F" & f & "=0,0,G" & f & "/F" & f & ")": ws.Rows(f).Font.Bold = True
    FormatearHojaControlBase ws, 1, 10, "F:G", "H"
End Sub

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
    Exit Sub

EH:
    errNumPivot = Err.Number
    errDescPivot = Err.Description
    camposPivot = CamposDisponiblesPivot(pt)

    Err.Raise errNumPivot, "CrearTablaDinamicaOSalidaAgrupada", _
              "Error creando PivotTable. Etapa: " & etapaVisual & _
              " | campoActual: " & campoActual & _
              " | accionActual: " & accionActual & _
              " | orientación solicitada: " & OrientacionPivotTexto(orientacionActual) & _
              " | Campos disponibles pivot: " & camposPivot & _
              " | Encabezados Base_Agregada: " & encabezadosBase & _
              " | Err.Number: " & CStr(errNumPivot) & _
              " | Err.Description: " & errDescPivot
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
    Err.Raise Err.Number, "ConfigurarNivel3ConFallback", _
              "Falló configuración de '" & campoNivel3 & "' como fila. " & _
              "Intento 1 (con posición=" & CStr(posicion) & ") -> Err.Number=" & CStr(errNumPosicion) & ", Err.Description=" & errDescPosicion & ". " & _
              "Intento 2 (solo Orientation, sin Position) -> Err.Number=" & CStr(Err.Number) & ", Err.Description=" & Err.Description
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
    Err.Raise Err.Number, "ConfigurarCampoMesColumnaConFallback", _
              "MesNombre falló como columna (Err.Number=" & CStr(errNumMesNombre) & ", Err.Description=" & errDescMesNombre & "). " & _
              "MesNum también falló (Err.Number=" & CStr(Err.Number) & ", Err.Description=" & Err.Description & ")."
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
    Err.Raise errNum, "ConfigurarCampoPivotSeguro", _
              "Error al asignar Orientation. nombreCampo=" & nombreCampo & _
              " | orientación solicitada=" & OrientacionPivotTexto(orientacion) & _
              " | orientación actual previa=" & OrientacionPivotTexto(orientacionPrevia) & _
              " | SourceName=" & sourceName & _
              " | Caption=" & caption & _
              " | Err.Number=" & CStr(errNum) & _
              " | Err.Description=" & errDesc

EH_POSITION:
    errNum = Err.Number
    errDesc = Err.Description
    Err.Raise errNum, "ConfigurarCampoPivotSeguro", _
              "Error al asignar Position. nombreCampo=" & nombreCampo & _
              " | orientación solicitada=" & OrientacionPivotTexto(orientacion) & _
              " | orientación actual previa=" & OrientacionPivotTexto(orientacionPrevia) & _
              " | SourceName=" & sourceName & _
              " | Caption=" & caption & _
              " | Position solicitada=" & CStr(posicion) & _
              " | Err.Number=" & CStr(errNum) & _
              " | Err.Description=" & errDesc
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
    Dim logoH As Double, logoW As Double, topPos As Double, leftPos As Double
    Dim rngBandaSuperior As Range

    On Error Resume Next
    ws.Shapes("imgLogoBPS").Delete
    On Error GoTo 0

    If Dir$(LOGO_BPS_PATH, vbNormal) = "" Then Exit Sub

    On Error GoTo EH
    logoH = ws.Rows(1).Height - 6
    Set shp = ws.Shapes.AddPicture(LOGO_BPS_PATH, msoFalse, msoTrue, 0, 0, -1, logoH)
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
End Sub

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
    ws.Columns("A").ColumnWidth = 28
    ArmarEncabezadoVisual ws, anio, mesCierre
    ws.Activate
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayGridlines = False
End Sub

Private Sub AgregarSlicerFinanciamiento(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal pt As PivotTable)
    Const CAMPO_FINANCIAMIENTO As String = "Financiamiento"
    Const NOMBRE_SLICER As String = "slFinanciamiento"

    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim topPos As Double
    Dim leftPos As Double
    Dim ancho As Double
    Dim alto As Double

    On Error GoTo SalidaSilenciosa

    If wb Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If pt Is Nothing Then Exit Sub
    If Not PivotFieldExiste(pt, CAMPO_FINANCIAMIENTO) Then Exit Sub

    On Error Resume Next
    ws.Shapes(NOMBRE_SLICER).Delete
    On Error GoTo SalidaSilenciosa

    Set sc = CrearSlicerCacheParaCampo(wb, pt, CAMPO_FINANCIAMIENTO)
    If sc Is Nothing Then Exit Sub

    On Error Resume Next
    sc.PivotTables.AddPivotTable pt
    On Error GoTo SalidaSilenciosa

    topPos = ws.Range("A5").Top
    leftPos = ws.Range("A5").Left
    ancho = ws.Range("A:A").Width
    alto = 180

    Set sl = sc.Slicers.Add(ws, , NOMBRE_SLICER, "Financiamiento", topPos, leftPos, ancho, alto)

    With sl
        .Caption = "Financiamiento"
        .DisplayHeader = True
        .Top = topPos
        .Left = leftPos
        .Width = ancho
        .Height = alto
        .NumberOfColumns = 1
    End With

SalidaSilenciosa:
End Sub

Private Function CrearSlicerCacheParaCampo(ByVal wb As Workbook, ByVal pt As PivotTable, ByVal nombreCampo As String) As SlicerCache
    Dim sc As SlicerCache

    On Error Resume Next

    For Each sc In wb.SlicerCaches
        If StrComp(sc.SourceName, nombreCampo, vbTextCompare) = 0 Then
            Set CrearSlicerCacheParaCampo = sc
            Exit Function
        End If
    Next sc

    Err.Clear
    Set CrearSlicerCacheParaCampo = wb.SlicerCaches.Add2(pt, nombreCampo)

    If CrearSlicerCacheParaCampo Is Nothing Then
        Err.Clear
        Set CrearSlicerCacheParaCampo = wb.SlicerCaches.Add(pt, nombreCampo)
    End If

    On Error GoTo 0
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
    ws.Columns("A").ColumnWidth = 28

    Set rg = wsBase.Range("A1").CurrentRegion
    etapaVisual = "creando PivotCache % ejecución"
    Set pc = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg.Address(ReferenceStyle:=xlR1C1, External:=True))
    etapaVisual = "creando PivotTable % ejecución"
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="ptPorcEjecGG")

    etapaVisual = "configurando filas % ejecución"
    ConfigurarCampoPivotSeguro pt, "Clasificación", xlRowField, 1
    ConfigurarCampoPivotSeguro pt, "Tipo", xlRowField, 2
    ConfigurarCampoPivotSeguro pt, "Concepto", xlRowField, 3
    etapaVisual = "configurando filtro Financiamiento % ejecución"
    ConfigurarCampoPivotSeguro pt, "Financiamiento", xlPageField, 1

    etapaVisual = "validando campos de valores % ejecución"
    If Not PivotFieldExiste(pt, "Ejecutado") Then
        Err.Raise vbObjectError + 1301, "CrearHojaPorcEjecucion", "No existe el campo 'Ejecutado' en Base_Porc_Ejec."
    End If
    If Not PivotFieldExiste(pt, "Asignado") Then
        Err.Raise vbObjectError + 1302, "CrearHojaPorcEjecucion", "No existe el campo 'Asignado' en Base_Porc_Ejec."
    End If

    etapaVisual = "creando campo calculado % ejecución"
    If Not PivotFieldExiste(pt, "% de ejecución") Then
        pt.CalculatedFields.Add "% de ejecución", "=Ejecutado/Asignado"
    End If
    If Not PivotFieldExiste(pt, "% de ejecución") Then
        Err.Raise vbObjectError + 1303, "CrearHojaPorcEjecucion", "No se pudo crear el campo calculado '% de ejecución'."
    End If

    etapaVisual = "agregando valores % ejecución"
    Call AgregarDataFieldSeguro(pt, "Ejecutado", "Ejecutado ", xlSum, "#,##0")
    Call AgregarDataFieldSeguro(pt, "Asignado", "Asignado " & CStr(anio), xlSum, "#,##0")
    Call AgregarDataFieldSeguro(pt, "% de ejecución", " % ejec.", xlSum, "0.0%")

    pt.RefreshTable

    pt.RowAxisLayout xlCompactRow
    pt.DisplayFieldCaptions = False
    pt.ColumnGrand = True: pt.RowGrand = True
    pt.NullString = "": pt.DisplayNullString = True
    ColapsarPivotInicial pt

    etapaVisual = "agregando slicer % ejecución"
    AgregarSlicerFinanciamiento wbOut, ws, pt
    etapaVisual = "ajustando encabezado % ejecución"
    AjustarEncabezadoVisualAlPivot ws, pt, "Informe de Seguimiento Presupuestal " & UCase$(MesesES()(mesCierre - 1)) & " " & anio & " - % de ejec. acumulada sobre la asignación presupuestal"
    Exit Sub
EH:
    Err.Raise Err.Number, "CrearHojaPorcEjecucion", _
        "Error creando hoja % ejecución. Etapa: " & etapaVisual & _
        " | Campos disponibles pivot: " & CamposDisponiblesPivot(pt) & _
        " | Err.Number: " & CStr(Err.Number) & _
        " | Err.Description: " & Err.Description
End Sub

Private Function AgregarDataFieldSeguro( _
    ByVal pt As PivotTable, _
    ByVal sourceFieldName As String, _
    ByVal caption As String, _
    ByVal funcion As XlConsolidationFunction, _
    ByVal formatoNumero As String) As PivotField

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

    Set AgregarDataFieldSeguro = pfData
End Function


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
    lastPivotCol = pt.TableRange2.Column + pt.TableRange2.Columns.Count - 1
    Set rngBanda = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastPivotCol))
    Set rngTitulo = ws.Range(ws.Cells(3, 1), ws.Cells(3, lastPivotCol))

    ws.Rows(1).RowHeight = 50.25: ws.Rows(2).RowHeight = 15: ws.Rows(3).RowHeight = 24
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
    Dim shp As Shape, logoH As Double
    If Dir$(LOGO_BPS_PATH, vbNormal) = "" Then Exit Sub
    logoH = ws.Rows(1).Height - 6
    Set shp = ws.Shapes.AddPicture(LOGO_BPS_PATH, msoFalse, msoTrue, 0, 0, -1, logoH)
    shp.Name = "imgLogoBPS": shp.LockAspectRatio = msoTrue
    shp.Top = ws.Rows(1).Top + (ws.Rows(1).Height - shp.Height) / 2
    shp.Left = rngBanda.Left + rngBanda.Width - shp.Width - 6
End Sub

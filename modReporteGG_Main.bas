Option Explicit

Public Sub Generar_Reporte_GG_Desde_Panel()
    On Error GoTo EH

    Dim procedimiento As String
    Dim etapaActual As String
    Dim wsPanel As Worksheet
    Dim anio As Long
    Dim mesTxt As String
    Dim mesCierre As Long
    Dim archivoEjec As String
    Dim archivoCod As String
    Dim archivoAsignados As String
    Dim anioComparativo As Long
    Dim archivoEjecComparativo As String
    Dim archivoAsignadosComparativo As String
    Dim carpetaEjecActual As String
    Dim carpetaAsignadosActual As String
    Dim carpetaEjecComparativo As String
    Dim carpetaAsignadosComparativo As String
    Dim wbE As Workbook
    Dim wbC As Workbook
    Dim wbOut As Workbook
    Dim wbA As Workbook
    Dim wbEComp As Workbook
    Dim wbAComp As Workbook
    Dim wsE As Worksheet
    Dim wsC As Worksheet
    Dim wsBase As Worksheet
    Dim wsA As Worksheet
    Dim wsBasePorc As Worksheet
    Dim wsEComp As Worksheet
    Dim wsAComp As Worksheet
    Dim wsBaseComp As Worksheet
    Dim dictCod As Object
    Dim dictAgg As Object
    Dim dictLlavesCodiguera As Object
    Dim diag As Object
    Dim dictAsignado As Object
    Dim dictPorcEjec As Object
    Dim dictIndicePorClave As Object
    Dim dictCompActual As Object
    Dim dictCompAnteriorActualizado As Object
    Dim rutaFinal As String
    Dim etapaVisual As String
    Dim hojaReporteActual As String

    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errLine As Long
    Dim msg As String
    Dim etapaVisualMsg As String
    Dim wbOutMsg As String
    Dim existeBase As String
    Dim ultimaFilaBase As String
    Dim ultimaColBase As String
    Dim archivoEjecMsg As String
    Dim archivoCodMsg As String
    Dim archivoAsignadosMsg As String
    Dim salidaMsg As String
    Dim archivoEjecComparativoMsg As String
    Dim archivoAsignadosComparativoMsg As String
    Dim carpetaEjecActualMsg As String
    Dim carpetaAsignadosActualMsg As String
    Dim carpetaEjecComparativoMsg As String
    Dim carpetaAsignadosComparativoMsg As String
    Dim carpetaIndicesMsg As String
    Dim diagArchivosEjecActual As String
    Dim diagArchivosAsignadosActual As String
    Dim diagArchivosEjecComparativo As String
    Dim diagArchivosAsignadosComparativo As String

    procedimiento = "Generar_Reporte_GG_Desde_Panel"
    Debug.Print "Inicio Generar_Reporte_GG_Desde_Panel: " & Now

    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then ChDir ThisWorkbook.Path
    On Error GoTo EH

    etapaActual = "leyendo parámetros del panel"
    Set wsPanel = ObtenerHojaPanelReportes()

    etapaActual = "validando año"
    If Not IsNumeric(wsPanel.Range("B3").Value) Then
        Err.Raise vbObjectError + 100, procedimiento, "Año inválido en Panel Reportes!B3. Valor: '" & CStr(wsPanel.Range("B3").Value) & "'."
    End If
    anio = CLng(wsPanel.Range("B3").Value)
    If anio < 2000 Or anio > 2100 Then
        Err.Raise vbObjectError + 104, procedimiento, "Año fuera de rango esperado en B3: " & anio
    End If

    etapaActual = "validando mes"
    mesTxt = CStr(wsPanel.Range("B4").Value)
    mesCierre = MesTextoANumero(mesTxt)
    If mesCierre < 1 Or mesCierre > 12 Then
        Err.Raise vbObjectError + 101, procedimiento, "Mes inválido en Panel Reportes!B4. Use: Enero..Diciembre (Setiembre). Valor: '" & mesTxt & "'."
    End If

    Set dictCod = CreateObject("Scripting.Dictionary")
    Set dictAgg = CreateObject("Scripting.Dictionary")
    Set dictLlavesCodiguera = CreateObject("Scripting.Dictionary")
    Set diag = CreateObject("Scripting.Dictionary")
    Set dictAsignado = CreateObject("Scripting.Dictionary")
    Set dictPorcEjec = CreateObject("Scripting.Dictionary")
    Set dictIndicePorClave = CreateObject("Scripting.Dictionary")
    Set dictCompActual = CreateObject("Scripting.Dictionary")
    Set dictCompAnteriorActualizado = CreateObject("Scripting.Dictionary")
    anioComparativo = anio - 1

    etapaActual = "resolviendo carpeta de ejecuciones actual"
    carpetaEjecActual = RutaCarpetaEjecucionesAnioActiva(anio)
    carpetaEjecActualMsg = carpetaEjecActual
    diagArchivosEjecActual = DiagnosticoArchivosExcelCarpeta(carpetaEjecActual)
    If Len(Dir(carpetaEjecActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 102, procedimiento, "No existe la carpeta de ejecuciones del año " & anio & ": " & carpetaEjecActual
    End If

    etapaActual = "buscando archivo de ejecuciones actual"
    archivoEjec = ObtenerArchivoMasReciente(carpetaEjecActual)
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 103, procedimiento, "No se encontró archivo de ejecuciones del año " & anio & " en: " & carpetaEjecActual
    End If

    etapaActual = "buscando codiguera"
    archivoCod = ResolverArchivoCodiguera(RutaCodigueraActiva())
    If Len(archivoCod) = 0 Then
        Err.Raise vbObjectError + 103, procedimiento, "No se encontró archivo de codiguera en: " & RutaCodigueraActiva()
    End If

    etapaActual = "resolviendo carpeta de asignados actual"
    carpetaAsignadosActual = RutaCarpetaAsignadosGastosAnioActiva(anio)
    carpetaAsignadosActualMsg = carpetaAsignadosActual
    diagArchivosAsignadosActual = DiagnosticoArchivosExcelCarpeta(carpetaAsignadosActual)
    If Len(Dir(carpetaAsignadosActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 117, procedimiento, "No existe la carpeta de asignados del año " & anio & ": " & carpetaAsignadosActual
    End If

    etapaActual = "buscando asignados actuales por fecha de creación"
    archivoAsignados = ObtenerArchivoMasRecientePorFechaCreacion(carpetaAsignadosActual)

    If Len(archivoAsignados) = 0 Then
        Err.Raise vbObjectError + 118, procedimiento, "No se encontró archivo de asignados del año " & anio & " en: " & carpetaAsignadosActual
    End If

    archivoAsignadosMsg = archivoAsignados

    etapaActual = "abriendo archivo de ejecuciones"
    Set wbE = Workbooks.Open(archivoEjec, ReadOnly:=True)

    etapaActual = "abriendo codiguera"
    Set wbC = Workbooks.Open(archivoCod, ReadOnly:=False)
    If wbC.ReadOnly Then
        Err.Raise vbObjectError + 116, procedimiento, "No se pudo actualizar la codiguera porque está abierta en solo lectura o bloqueada por otro usuario."
    End If

    etapaActual = "abriendo asignados"
    Set wbA = Workbooks.Open(archivoAsignados, ReadOnly:=True)

    etapaActual = "leyendo hojas de origen"
    Set wsE = ObtenerHojaEjecuciones(wbE)
    Set wsC = ObtenerHojaCodiguera(wbC)
    Set wsA = ObtenerHojaAsignados(wbA)

    etapaActual = "leyendo codiguera"
    LeerCodiguera wsC, dictCod, dictLlavesCodiguera, diag
    LeerCodigueraIndices wsC, dictIndicePorClave

    etapaActual = "leyendo ejecuciones y acumulando"
    LeerEjecucionesYAcumular wsE, anio, mesCierre, dictCod, dictAgg, diag

    etapaActual = "leyendo asignados y acumulando"
    LeerAsignadosYAcumular wsA, dictCod, dictLlavesCodiguera, dictAsignado, diag, wbC, anio, archivoAsignados

    If diag.Exists("asignados_faltantes") Then
        etapaActual = "guardando codiguera por nuevas llaves de asignados actuales"
        wbC.Save
        EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

        If Not wbE Is Nothing Then wbE.Close False: Set wbE = Nothing
        If Not wbA Is Nothing Then wbA.Close False: Set wbA = Nothing
        If Not wbC Is Nothing Then wbC.Close True: Set wbC = Nothing

        MsgBox "Se agregaron nuevas llaves presupuestales del archivo de asignados actual a la codiguera. Debe clasificarlas, marcar Incluir_en_Informe cuando corresponda y volver a generar el reporte.", vbExclamation
        Exit Sub
    End If

    If dictAsignado.Count = 0 Then
        EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre
        Dim detalleAsignadosVacios As String
        detalleAsignadosVacios = "No se acumuló ningún asignado. Se escribió Diagnostico_Llaves con el resumen de asignados. Revise archivo de asignados, año, claves presupuestales e Incluir_en_Informe. Archivo: " & archivoAsignados & vbCrLf
        detalleAsignadosVacios = detalleAsignadosVacios & ResumenAsignadosParaError(diag)
        Err.Raise vbObjectError + 1250, procedimiento, detalleAsignadosVacios
    End If

    If SumaValoresDiccionario(dictAsignado) = 0 Then
        EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre
        Dim detalleAsignadosCero As String
        detalleAsignadosCero = "El total asignado acumulado es cero. Se escribió Diagnostico_Llaves con el resumen de asignados. Revise archivo de asignados, año, claves presupuestales e Incluir_en_Informe. Archivo: " & archivoAsignados & vbCrLf
        detalleAsignadosCero = detalleAsignadosCero & ResumenAsignadosParaError(diag)
        Err.Raise vbObjectError + 1251, procedimiento, detalleAsignadosCero
    End If

    etapaActual = "cerrando archivos actuales antes de abrir comparativo"
    If Not wbE Is Nothing Then
        wbE.Close False
        Set wbE = Nothing
    End If

    If Not wbA Is Nothing Then
        wbA.Close False
        Set wbA = Nothing
    End If

    etapaActual = "resolviendo carpeta ejecuciones comparativo"
    carpetaEjecComparativo = RutaCarpetaEjecucionesAnioActiva(anioComparativo)
    carpetaEjecComparativoMsg = carpetaEjecComparativo
    diagArchivosEjecComparativo = DiagnosticoArchivosExcelCarpeta(carpetaEjecComparativo)
    If Len(Dir(carpetaEjecComparativo, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 1970, procedimiento, "No existe la carpeta de ejecuciones comparativo: " & carpetaEjecComparativo
    End If

    etapaActual = "buscando archivo ejecuciones comparativo"
    archivoEjecComparativo = ObtenerArchivoMasReciente(carpetaEjecComparativo)
    archivoEjecComparativoMsg = archivoEjecComparativo
    If Len(archivoEjecComparativo) = 0 Then
        Err.Raise vbObjectError + 1971, procedimiento, "No se encontró archivo de ejecuciones comparativo en: " & carpetaEjecComparativo
    End If
    diag("archivo_ejec_comparativo") = archivoEjecComparativo
    diag("anio_comparativo") = anioComparativo

    etapaActual = "resolviendo carpeta asignados comparativo"
    carpetaAsignadosComparativo = RutaCarpetaAsignadosGastosAnioActiva(anioComparativo)
    carpetaAsignadosComparativoMsg = carpetaAsignadosComparativo
    diagArchivosAsignadosComparativo = DiagnosticoArchivosExcelCarpeta(carpetaAsignadosComparativo)
    If Len(Dir(carpetaAsignadosComparativo, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 1972, procedimiento, "No existe la carpeta de asignados comparativo: " & carpetaAsignadosComparativo
    End If

    etapaActual = "buscando archivo asignados comparativo"
    archivoAsignadosComparativo = ObtenerArchivoMasRecientePorFechaCreacion(carpetaAsignadosComparativo)
    archivoAsignadosComparativoMsg = archivoAsignadosComparativo
    If Len(archivoAsignadosComparativo) = 0 Then
        Err.Raise vbObjectError + 1973, procedimiento, "No se encontró archivo de asignados comparativo en: " & carpetaAsignadosComparativo
    End If
    diag("archivo_asignados_comparativo") = archivoAsignadosComparativo

    etapaActual = "abriendo archivo ejecuciones comparativo"
    Set wbEComp = Workbooks.Open(archivoEjecComparativo, ReadOnly:=True)
    If wbEComp Is Nothing Then
        Err.Raise vbObjectError + 1974, procedimiento, "Workbooks.Open devolvió Nothing para archivo de ejecuciones comparativo: " & archivoEjecComparativo
    End If

    etapaActual = "abriendo archivo asignados comparativo"
    Set wbAComp = Workbooks.Open(archivoAsignadosComparativo, ReadOnly:=True)
    If wbAComp Is Nothing Then
        Err.Raise vbObjectError + 1975, procedimiento, "Workbooks.Open devolvió Nothing para archivo de asignados comparativo: " & archivoAsignadosComparativo
    End If

    etapaActual = "obteniendo hoja ejecuciones comparativo"
    Set wsEComp = ObtenerHojaEjecuciones(wbEComp)
    If wsEComp Is Nothing Then
        Err.Raise vbObjectError + 1976, procedimiento, "No se pudo obtener hoja de ejecuciones comparativo en: " & archivoEjecComparativo
    End If

    etapaActual = "obteniendo hoja asignados comparativo"
    Set wsAComp = ObtenerHojaAsignados(wbAComp)
    If wsAComp Is Nothing Then
        Err.Raise vbObjectError + 1977, procedimiento, "No se pudo obtener hoja de asignados comparativo en: " & archivoAsignadosComparativo
    End If

    etapaActual = "validando asignados comparativo"
    ValidarAsignadosComparativoContraCodiguera wsAComp, dictCod, dictLlavesCodiguera, dictIndicePorClave, diag, wbC, anioComparativo, archivoAsignadosComparativo
    If diag.Exists("comparativo_asignados_faltantes") Then
        wbC.Save
        EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre
        If Not wbC Is Nothing Then wbC.Close True: Set wbC = Nothing
        If Not wbEComp Is Nothing Then wbEComp.Close False: Set wbEComp = Nothing
        If Not wbAComp Is Nothing Then wbAComp.Close False: Set wbAComp = Nothing
        MsgBox "Se agregaron nuevas llaves presupuestales del archivo de asignados comparativo a la codiguera. Debe clasificarlas, indicar Indice, marcar Incluir_en_Informe cuando corresponda y volver a generar el reporte.", vbExclamation
        Exit Sub
    End If

    ConstruirDictComparativoActualDesdeDictAgg dictAgg, mesCierre, dictCompActual
    LeerEjecucionesComparativoYAcumular wsEComp, anioComparativo, anio, mesCierre, dictCod, dictIndicePorClave, dictCompAnteriorActualizado, diag

    If Not wbEComp Is Nothing Then wbEComp.Close False: Set wbEComp = Nothing
    If Not wbAComp Is Nothing Then wbAComp.Close False: Set wbAComp = Nothing

    etapaActual = "completando meses faltantes"
    CompletarMesesAnioEnDictAgg dictAgg

    etapaActual = "creando workbook de salida"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBase = wbOut.Worksheets(1)
    wsBase.Name = "Base_Agregada"

    etapaActual = "construyendo base agregada"
    ConstruirBaseAgregadaReporte wsBase, dictAgg

    etapaActual = "creando reporte visual"
    hojaReporteActual = "Ejec. Mensual " & anio
    etapaActual = "creando hoja Ejec. Mensual"
    etapaVisual = "iniciando hoja Ejec. Mensual"
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre, etapaVisual
    Set wsBasePorc = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    wsBasePorc.Name = "Base_Porc_Ejec"
    ConstruirBasePorcEjec wsBasePorc, dictAgg, dictAsignado, mesCierre
    Set wsBaseComp = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    wsBaseComp.Name = "Ejec comparada datos"
    ConstruirBaseEjecComparada wsBaseComp, dictCompActual, dictCompAnteriorActualizado, anio, anioComparativo
    CrearHojaComparativoAnual wbOut, wsBaseComp, anio, anioComparativo, mesCierre, etapaVisual
    hojaReporteActual = "% ejecución " & anio
    etapaActual = "creando hoja % ejecución"
    etapaVisual = "iniciando hoja % ejecución"
    CrearHojaPorcEjecucion wbOut, wsBasePorc, anio, mesCierre, etapaVisual
    wsBase.Visible = xlSheetVeryHidden
    wsBasePorc.Visible = xlSheetVeryHidden
    wsBaseComp.Visible = xlSheetVeryHidden

    etapaActual = "guardando reporte liviano"
    rutaFinal = GuardarReporteLiviano(wbOut, anio, mesCierre)

    etapaActual = "escribiendo diagnóstico"
    EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

    etapaActual = "cerrando archivos"
    wbOut.Close False
    If Not wbE Is Nothing Then wbE.Close False
    If Not wbA Is Nothing Then wbA.Close False
    If Not wbC Is Nothing Then wbC.Close False
    If Not wbEComp Is Nothing Then wbEComp.Close False
    If Not wbAComp Is Nothing Then wbAComp.Close False

    MsgBox "Reporte generado: " & rutaFinal, vbInformation
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl

    If Len(etapaVisual) > 0 Then
        etapaVisualMsg = etapaVisual
    Else
        etapaVisualMsg = "(no aplica)"
    End If

    If wbOut Is Nothing Then
        wbOutMsg = "(Nothing)"
        existeBase = "(wbOut Nothing)"
        ultimaFilaBase = "(wbOut Nothing)"
        ultimaColBase = "(wbOut Nothing)"
    Else
        wbOutMsg = wbOut.Name
        If HojaExiste(wbOut, "Base_Agregada") Then
            existeBase = "SI"
            ultimaFilaBase = CStr(ObtenerUltimaFilaSegura(wbOut, "Base_Agregada"))
            ultimaColBase = CStr(ObtenerUltimaColSegura(wbOut, "Base_Agregada"))
        Else
            existeBase = "NO"
            ultimaFilaBase = "(no existe)"
            ultimaColBase = "(no existe)"
        End If
    End If

    If Len(archivoEjec) > 0 Then
        archivoEjecMsg = archivoEjec
    Else
        archivoEjecMsg = "(no detectado)"
    End If

    If Len(archivoCod) > 0 Then
        archivoCodMsg = archivoCod
    Else
        archivoCodMsg = "(no detectado)"
    End If

    If Len(archivoAsignados) > 0 Then
        archivoAsignadosMsg = archivoAsignados
    Else
        archivoAsignadosMsg = "(no detectado)"
    End If

    If Len(rutaFinal) > 0 Then
        salidaMsg = rutaFinal
    Else
        salidaMsg = RutaReportesGeneradosActiva()
    End If

    carpetaIndicesMsg = RutaCarpetaIndicesActiva()

    msg = "Error al generar reporte." & vbCrLf & vbCrLf
    msg = msg & "Procedimiento: " & procedimiento & vbCrLf
    msg = msg & "Etapa: " & etapaActual & vbCrLf
    msg = msg & "Etapa visual: " & etapaVisualMsg & vbCrLf
    msg = msg & "Err.Number: " & errNum & vbCrLf
    msg = msg & "Err.Description: " & errDesc & vbCrLf
    msg = msg & "Err.Source: " & errSource & vbCrLf
    msg = msg & "Erl: " & errLine & vbCrLf
    msg = msg & "Workbook salida: " & wbOutMsg & vbCrLf
    msg = msg & "Existe Base_Agregada: " & existeBase & vbCrLf
    msg = msg & "Última fila Base_Agregada: " & ultimaFilaBase & vbCrLf
    msg = msg & "Última columna Base_Agregada: " & ultimaColBase & vbCrLf
    msg = msg & "Nombre hoja reporte: " & IIf(Len(hojaReporteActual) > 0, hojaReporteActual, "(no determinada)") & vbCrLf
    msg = msg & "Archivo ejecuciones: " & archivoEjecMsg & vbCrLf
    msg = msg & "Archivo codiguera: " & archivoCodMsg & vbCrLf
    msg = msg & "Archivo asignados: " & archivoAsignadosMsg & vbCrLf
    msg = msg & "Carpeta ejecuciones actual: " & IIf(Len(carpetaEjecActualMsg) > 0, carpetaEjecActualMsg, "(no determinada)") & vbCrLf
    msg = msg & "Carpeta asignados actual: " & IIf(Len(carpetaAsignadosActualMsg) > 0, carpetaAsignadosActualMsg, "(no determinada)") & vbCrLf
    msg = msg & "Carpeta ejecuciones comparativo: " & IIf(Len(carpetaEjecComparativoMsg) > 0, carpetaEjecComparativoMsg, "(no determinada)") & vbCrLf
    msg = msg & "Archivo ejecuciones comparativo: " & IIf(Len(archivoEjecComparativoMsg) > 0, archivoEjecComparativoMsg, "(no detectado)") & vbCrLf
    msg = msg & "Carpeta asignados comparativo: " & IIf(Len(carpetaAsignadosComparativoMsg) > 0, carpetaAsignadosComparativoMsg, "(no determinada)") & vbCrLf
    msg = msg & "Archivo asignados comparativo: " & IIf(Len(archivoAsignadosComparativoMsg) > 0, archivoAsignadosComparativoMsg, "(no detectado)") & vbCrLf
    msg = msg & "Carpeta índices: " & carpetaIndicesMsg & vbCrLf
    msg = msg & "Salida: " & salidaMsg & vbCrLf
    msg = msg & "Carpeta base local: " & ThisWorkbook.Path & vbCrLf
    msg = msg & vbCrLf & DiagnosticoRutasActivas()

    msg = msg & vbCrLf & "Archivos candidatos ejecuciones actual:" & vbCrLf
    msg = msg & diagArchivosEjecActual & vbCrLf
    msg = msg & vbCrLf & "Archivos candidatos asignados actual:" & vbCrLf
    msg = msg & diagArchivosAsignadosActual & vbCrLf
    msg = msg & vbCrLf & "Archivos candidatos ejecuciones comparativo:" & vbCrLf
    msg = msg & diagArchivosEjecComparativo & vbCrLf
    msg = msg & vbCrLf & "Archivos candidatos asignados comparativo:" & vbCrLf
    msg = msg & diagArchivosAsignadosComparativo & vbCrLf
    msg = msg & vbCrLf & "Workbooks abiertos:" & vbCrLf
    msg = msg & DiagnosticoWorkbooksAbiertos() & vbCrLf

    Debug.Print String(100, "-")
    Debug.Print msg
    Debug.Print String(100, "-")

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
    If Not wbE Is Nothing Then wbE.Close False
    If Not wbA Is Nothing Then wbA.Close False
    If Not wbC Is Nothing Then wbC.Close False
    If Not wbEComp Is Nothing Then wbEComp.Close False
    If Not wbAComp Is Nothing Then wbAComp.Close False
    On Error GoTo 0

    MsgBox msg, vbCritical
End Sub

Public Sub LeerAsignadosYAcumular(ByVal ws As Worksheet, ByVal dictCod As Object, ByVal dictLlavesCodiguera As Object, ByRef dictAsignado As Object, ByRef diag As Object, Optional ByVal wbCodiguera As Workbook, Optional ByVal anioFiltro As Long = 0, Optional ByVal archivoAsignados As String = "")
    Dim arr As Variant, headers As Object, i As Long, clave As String, info As Variant, keyAgg As String, monto As Double
    Dim colAnio As Long, anioFila As Long
    Dim filasAsignadosLeidas As Long, filasAsignadosAnio As Long
    Dim filasAsignadosConClaveEnDictCod As Long, filasAsignadosClaveExistePeroNoIncluida As Long, filasAsignadosNuevas As Long
    Dim sumaAsignadoArchivo As Double, sumaAsignadoAnio As Double, sumaAsignadoAcumulado As Double

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    ValidarColumnasAsignados headers
    colAnio = ObtenerColumnaOpcional(headers, Array("año", "anio", "ejercicio", "ej"))

    For i = 2 To UBound(arr, 1)
        filasAsignadosLeidas = filasAsignadosLeidas + 1

        monto = CDbl(0 + arr(i, ObtenerColumna(headers, Array("asignado"))))
        sumaAsignadoArchivo = sumaAsignadoArchivo + monto

        If colAnio > 0 And anioFiltro > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                If anioFila <> anioFiltro Then GoTo SiguienteFila
            End If
        End If

        filasAsignadosAnio = filasAsignadosAnio + 1
        sumaAsignadoAnio = sumaAsignadoAnio + monto
        clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, ObtenerColumna(headers, Array("finac"))), arr(i, ObtenerColumna(headers, Array("der-f"))), arr(i, ObtenerColumna(headers, Array("pg"))), arr(i, ObtenerColumna(headers, Array("spg"))), arr(i, ObtenerColumna(headers, Array("proy"))), arr(i, ObtenerColumna(headers, Array("rubro"))), arr(i, ObtenerColumna(headers, Array("r. aux"))), arr(i, ObtenerColumna(headers, Array("ue"))), arr(i, ObtenerColumna(headers, Array("dep"))), arr(i, ObtenerColumna(headers, Array("obra"))), arr(i, ObtenerColumna(headers, Array("der. obra"))), arr(i, ObtenerColumna(headers, Array("serv"))), arr(i, ObtenerColumna(headers, Array("sniip"))))

        If dictCod.Exists(clave) Then
            info = dictCod(clave)
            keyAgg = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3))
            If Not dictAsignado.Exists(keyAgg) Then dictAsignado.Add keyAgg, 0#
            dictAsignado(keyAgg) = dictAsignado(keyAgg) + monto
            sumaAsignadoAcumulado = sumaAsignadoAcumulado + monto
            filasAsignadosConClaveEnDictCod = filasAsignadosConClaveEnDictCod + 1
        ElseIf dictLlavesCodiguera.Exists(clave) Then
            filasAsignadosClaveExistePeroNoIncluida = filasAsignadosClaveExistePeroNoIncluida + 1
            AgregarMuestraAsignadoNoAcumulado diag, "Existe en codiguera pero no está incluida en informe", i, clave, monto
        Else
            filasAsignadosNuevas = filasAsignadosNuevas + 1
            AgregarMuestraAsignadoNoAcumulado diag, "Nueva llave no encontrada en codiguera", i, clave, monto
            RegistrarYAgregarLlaveAsignadoFaltante wbCodiguera, diag, clave, arr, headers, i, archivoAsignados
            If Not dictLlavesCodiguera.Exists(clave) Then dictLlavesCodiguera.Add clave, True
        End If
SiguienteFila:
    Next i

    diag("asignados_resumen") = Array( _
        archivoAsignados, _
        filasAsignadosLeidas, _
        filasAsignadosAnio, _
        filasAsignadosConClaveEnDictCod, _
        filasAsignadosClaveExistePeroNoIncluida, _
        filasAsignadosNuevas, _
        sumaAsignadoArchivo, _
        sumaAsignadoAnio, _
        sumaAsignadoAcumulado, _
        dictAsignado.Count _
    )
End Sub

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictCod As Object, ByRef dictLlavesCodiguera As Object, ByRef diag As Object)
    On Error GoTo EH

    Dim arr As Variant, headers As Object, i As Long, incluir As String, clave As String, info As Variant
    Dim colTitular As Long, colClaveLlave As Long

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    colTitular = ObtenerColumna(headers, Array("titular"))
    colClaveLlave = ObtenerColumna(headers, Array("clave llave presupuestal"))
    If colTitular = 0 Then Err.Raise vbObjectError + 201, "LeerCodiguera", "Falta columna Titular en codiguera."

    For i = 2 To UBound(arr, 1)
        clave = NormalizarClaveCodigueraDesdeTexto(arr(i, colClaveLlave))
        If Len(clave) > 0 And Not dictLlavesCodiguera.Exists(clave) Then dictLlavesCodiguera.Add clave, True
        incluir = Replace(UCase$(Trim$(CStr(arr(i, ObtenerColumna(headers, Array("incluir_en_informe")))))), " ", "")
        If incluir = "SI" And Len(clave) > 0 Then
            info = Array(arr(i, colTitular), arr(i, ObtenerColumna(headers, Array("nivel_1"))), arr(i, ObtenerColumna(headers, Array("nivel_2"))), arr(i, ObtenerColumna(headers, Array("nivel_3"))))
            dictCod(clave) = info
        End If
    Next i
    Exit Sub
EH:
    Err.Raise Err.Number, "LeerCodiguera", "Error leyendo codiguera: " & Err.Description
End Sub

Public Sub LeerEjecucionesYAcumular(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByVal dictCod As Object, ByRef dictAgg As Object, ByRef diag As Object)
    On Error GoTo EH

    Dim arr As Variant, headers As Object, i As Long, fechaValor As Date, clave As String, info As Variant, mesNum As Long, aggregateKey As String, importeMN As Double
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    For i = 2 To UBound(arr, 1)
        If TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(headers, Array("fecha valor"))), fechaValor) Then
            If Year(fechaValor) = anio And Month(fechaValor) <= mesCierre Then
                clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, ObtenerColumna(headers, Array("finac código numérico"))), arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), arr(i, ObtenerColumna(headers, Array("pg código numérico"))), arr(i, ObtenerColumna(headers, Array("spg código numérico"))), arr(i, ObtenerColumna(headers, Array("proyecto", "proy"))), arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), arr(i, ObtenerColumna(headers, Array("ue código numérico"))), arr(i, ObtenerColumna(headers, Array("dep código numérico"))), arr(i, ObtenerColumna(headers, Array("obra código numérico"))), arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), arr(i, ObtenerColumna(headers, Array("serv código numérico"))), arr(i, ObtenerColumna(headers, Array("snip código numérico"))))
                If dictCod.Exists(clave) Then
                    info = dictCod(clave): mesNum = Month(fechaValor)
                    importeMN = CDbl(0 + arr(i, ObtenerColumna(headers, Array("importe moneda nacional"))))
                    aggregateKey = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3)) & "|" & CStr(mesNum)
                    If Not dictAgg.Exists(aggregateKey) Then dictAgg.Add aggregateKey, 0#
                    dictAgg(aggregateKey) = dictAgg(aggregateKey) + importeMN
                End If
            End If
        End If
    Next i
    Exit Sub
EH:
    Err.Raise Err.Number, "LeerEjecucionesYAcumular", "Error leyendo ejecuciones: " & Err.Description
End Sub


Private Sub CompletarMesesAnioEnDictAgg(ByRef dictAgg As Object)
    Dim dictRows As Object
    Dim k As Variant
    Dim partes() As String
    Dim rowKey As String
    Dim m As Long
    Dim fullKey As String

    Set dictRows = CreateObject("Scripting.Dictionary")

    For Each k In dictAgg.Keys
        partes = Split(CStr(k), "|")
        If UBound(partes) >= 4 Then
            rowKey = CStr(partes(0)) & "|" & CStr(partes(1)) & "|" & CStr(partes(2)) & "|" & CStr(partes(3))
            If Not dictRows.Exists(rowKey) Then dictRows.Add rowKey, True
        End If
    Next k

    For Each k In dictRows.Keys
        For m = 1 To 12
            fullKey = CStr(k) & "|" & CStr(m)
            If Not dictAgg.Exists(fullKey) Then dictAgg.Add fullKey, 0#
        Next m
    Next k
End Sub

Public Sub ConstruirBaseAgregadaReporte(ByVal ws As Worksheet, ByVal dictAgg As Object)
    Dim fila As Long, dictKey As Variant, partes() As String, importeSalida As Double, factor As Double
    Dim arrMesesMin As Variant
    ws.Range("A1:G1").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")
    arrMesesMin = MesesESMin()
    factor = FactorEscalaImporte()
    fila = 2
    For Each dictKey In dictAgg.Keys
        partes = Split(CStr(dictKey), "|")
        ws.Cells(fila, 1).Value = LimpiarTexto(CStr(partes(0)))
        ws.Cells(fila, 2).Value = LimpiarTexto(CStr(partes(1)))
        ws.Cells(fila, 3).Value = LimpiarTexto(CStr(partes(2)))
        ws.Cells(fila, 4).Value = LimpiarTexto(CStr(partes(3)))
        ws.Cells(fila, 5).Value = CLng(partes(4))
        ws.Cells(fila, 6).Value = arrMesesMin(CLng(partes(4)) - 1)
        importeSalida = CDbl(dictAgg(dictKey)) / factor
        ws.Cells(fila, 7).Value = importeSalida
        fila = fila + 1
    Next dictKey
End Sub

Public Function GuardarReporteLiviano(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mesNum As Long) As String
    Dim ruta As String, fileName As String
    Dim carpetaSalida As String
    On Error GoTo EH

    carpetaSalida = RutaReportesGeneradosActiva()
    AsegurarCarpetaExiste carpetaSalida
    fileName = "Informe_GG_Ejecucion_Mensual_" & anio & "_" & Format$(mesNum, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsx"
    ruta = CombinarRuta(carpetaSalida, fileName)
    wbOut.SaveAs ruta, xlOpenXMLWorkbook
    GuardarReporteLiviano = ruta
    Exit Function
EH:
    Err.Raise Err.Number, "GuardarReporteLiviano", "Error guardando reporte liviano: " & Err.Description & " | Ruta: " & ruta
End Function

Public Sub EscribirDiagnostico(ByVal wb As Workbook, ByVal diag As Object, ByVal archivoEjec As String, ByVal archivoCod As String, ByVal anio As Long, ByVal mesNum As Long)
    Dim ws As Worksheet, d As Object, k As Variant, f As Long, it As Variant
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(DIAG_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = DIAG_SHEET_NAME
    ws.Range("A1:B1").Value = Array("Campo", "Valor")
    ws.Cells(2, 1).Value = "Archivo ejecuciones": ws.Cells(2, 2).Value = archivoEjec
    ws.Cells(3, 1).Value = "Archivo codiguera": ws.Cells(3, 2).Value = archivoCod
    ws.Cells(4, 1).Value = "Año": ws.Cells(4, 2).Value = anio
    ws.Cells(5, 1).Value = "Mes cierre": ws.Cells(5, 2).Value = mesNum
    If diag.Exists("asignados_resumen") Then
        it = diag("asignados_resumen")
        ws.Cells(7, 1).Value = "Archivo asignados detectado": ws.Cells(7, 2).Value = it(0)
        ws.Cells(8, 1).Value = "Filas archivo asignados": ws.Cells(8, 2).Value = it(1)
        ws.Cells(9, 1).Value = "Filas usadas año seleccionado": ws.Cells(9, 2).Value = it(2)
        ws.Cells(10, 1).Value = "Filas con clave en dictCod": ws.Cells(10, 2).Value = it(3)
        ws.Cells(11, 1).Value = "Filas clave en codiguera no incluidas": ws.Cells(11, 2).Value = it(4)
        ws.Cells(12, 1).Value = "Filas nuevas": ws.Cells(12, 2).Value = it(5)
        ws.Cells(13, 1).Value = "Total Asignado leído": ws.Cells(13, 2).Value = it(6)
        ws.Cells(14, 1).Value = "Total Asignado del año": ws.Cells(14, 2).Value = it(7)
        ws.Cells(15, 1).Value = "Total Asignado acumulado": ws.Cells(15, 2).Value = it(8)
        ws.Cells(16, 1).Value = "Claves acumuladas dictAsignado": ws.Cells(16, 2).Value = it(9)
    End If
    If diag.Exists("asignados_faltantes") Then
        ws.Cells(17, 1).Value = "Llaves de asignados no encontradas en codiguera"
        ws.Range("A18:S18").Value = Array("Origen", "Archivo", "Fila origen", "Clave normalizada", "Llave presupuestal", "Finac", "Der-F", "PG", "Spg", "Proy", "Rubro", "R. Aux", "UE", "Dep", "Obra", "Der. Obra", "Serv", "SNIIP", "Estado")
        f = 19
        Set d = diag("asignados_faltantes")
        For Each k In d.Keys
            it = d(k)
            ws.Range("A" & f & ":S" & f).Value = it
            f = f + 1
        Next k
    End If
    If diag.Exists("asignados_muestra_no_acumulados") Then
        ws.Cells(17, 7).Value = "Muestra de asignados no acumulados"
        ws.Range("G18:J18").Value = Array("Estado", "Fila origen", "Clave normalizada", "Asignado")
        f = 19
        For Each it In diag("asignados_muestra_no_acumulados")
            ws.Cells(f, 7).Value = it(0)
            ws.Cells(f, 8).Value = it(1)
            ws.Cells(f, 9).Value = it(2)
            ws.Cells(f, 10).Value = it(3)
            f = f + 1
        Next it
    End If
    ws.Columns("A:B").AutoFit
    ws.Columns("G:J").AutoFit
    If diag.Exists("archivo_ejec_comparativo") Then ws.Cells(2, 4).Value = "Archivo ejecuciones comparativo": ws.Cells(2, 5).Value = diag("archivo_ejec_comparativo")
    If diag.Exists("archivo_asignados_comparativo") Then ws.Cells(3, 4).Value = "Archivo asignados comparativo": ws.Cells(3, 5).Value = diag("archivo_asignados_comparativo")
    If diag.Exists("anio_comparativo") Then ws.Cells(4, 4).Value = "Año comparativo": ws.Cells(4, 5).Value = diag("anio_comparativo")
End Sub

Public Sub LeerCodigueraIndices(ByVal ws As Worksheet, ByRef dictIndicePorClave As Object)
    Dim arr As Variant, headers As Object, i As Long, clave As String
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    For i = 2 To UBound(arr, 1)
        clave = NormalizarClaveCodigueraDesdeTexto(arr(i, ObtenerColumna(headers, Array("clave llave presupuestal"))))
        If Len(clave) > 0 Then dictIndicePorClave(clave) = Trim$(CStr(arr(i, ObtenerColumna(headers, Array("indice")))))
    Next i
End Sub

Public Function ClavePeriodoIndice(ByVal anio As Long, ByVal mes As Long) As String
    ClavePeriodoIndice = Format$(DateSerial(anio, mes, 1), "yyyy-mm")
End Function

Private Function ResumenAsignadosParaError(ByVal diag As Object) As String
    Dim it As Variant
    Dim mensaje As String

    If diag Is Nothing Then Exit Function
    If Not diag.Exists("asignados_resumen") Then Exit Function

    it = diag("asignados_resumen")

    mensaje = "Filas asignados leídas: " & CStr(it(1)) & vbCrLf
    mensaje = mensaje & "Filas asignados del año: " & CStr(it(2)) & vbCrLf
    mensaje = mensaje & "Claves acumuladas en dictCod: " & CStr(it(3)) & vbCrLf
    mensaje = mensaje & "Claves existentes pero no incluidas: " & CStr(it(4)) & vbCrLf
    mensaje = mensaje & "Claves nuevas: " & CStr(it(5)) & vbCrLf
    mensaje = mensaje & "Total Asignado leído: " & Format(it(6), "#,##0") & vbCrLf
    mensaje = mensaje & "Total Asignado del año: " & Format(it(7), "#,##0") & vbCrLf
    mensaje = mensaje & "Total Asignado acumulado: " & Format(it(8), "#,##0") & vbCrLf
    mensaje = mensaje & "Cantidad claves dictAsignado: " & CStr(it(9))

    If CLng(0 + it(2)) = 0 Then
        mensaje = mensaje & vbCrLf & "Atención: el archivo de asignados no tiene filas para el año seleccionado. Revise si el archivo corresponde al año del reporte."
    End If

    ResumenAsignadosParaError = mensaje
End Function

Private Function SumaValoresDiccionario(ByVal d As Object) As Double
    Dim k As Variant
    For Each k In d.Keys
        SumaValoresDiccionario = SumaValoresDiccionario + CDbl(d(k))
    Next k
End Function

Private Function ObtenerColumnaOpcional(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long, a As String
    For i = LBound(aliases) To UBound(aliases)
        a = LCase$(LimpiarTexto(CStr(aliases(i))))
        If headers.Exists(a) Then ObtenerColumnaOpcional = CLng(headers(a)): Exit Function
    Next i
End Function

Private Sub ValidarColumnasAsignados(ByVal headers As Object)
    Dim req As Variant, i As Long
    req = Array("finac", "der-f", "pg", "spg", "proy", "rubro", "r. aux", "ue", "dep", "obra", "der. obra", "serv", "sniip", "asignado")
    For i = LBound(req) To UBound(req)
        If Not headers.Exists(CStr(req(i))) Then
            If CStr(req(i)) = "asignado" Then
                Err.Raise vbObjectError + 1201, "LeerAsignadosYAcumular", "No se encontró la columna 'Asignado' en el archivo de asignados. No se puede generar la hoja % ejecución."
            Else
                Err.Raise vbObjectError + 1202, "LeerAsignadosYAcumular", "No se encontró la columna '" & CStr(req(i)) & "' en el archivo de asignados. No se puede construir la llave presupuestal."
            End If
        End If
    Next i
End Sub

Private Sub AgregarMuestraAsignadoNoAcumulado(ByRef diag As Object, ByVal estado As String, ByVal fila As Long, ByVal clave As String, ByVal monto As Double)
    Dim col As Collection
    If Not diag.Exists("asignados_muestra_no_acumulados") Then
        Set col = New Collection
        diag.Add "asignados_muestra_no_acumulados", col
    Else
        Set col = diag("asignados_muestra_no_acumulados")
    End If

    If col.Count < 20 Then
        col.Add Array(estado, fila, clave, monto)
    End If
End Sub

Private Function NormalizarClaveCodigueraDesdeTexto(ByVal valor As Variant) As String
    Dim partes() As String, i As Long
    Dim vacio(0 To 12) As Variant
    Dim componentes(0 To 12) As Variant
    Dim texto As String

    texto = Replace(CStr(valor), "|", "-")
    texto = LimpiarTexto(texto)
    If Len(texto) = 0 Then Exit Function
    partes = Split(texto, "-")
    If UBound(partes) = 12 Then
        For i = 0 To 12
            componentes(i) = partes(i)
        Next i
    Else
        For i = 0 To 12
            componentes(i) = vacio(i)
        Next i
        componentes(0) = texto
    End If
    NormalizarClaveCodigueraDesdeTexto = ConstruirClaveLlavePresupuestalCodiguera(componentes(0), componentes(1), componentes(2), componentes(3), componentes(4), componentes(5), componentes(6), componentes(7), componentes(8), componentes(9), componentes(10), componentes(11), componentes(12))
End Function

Private Sub RegistrarYAgregarLlaveAsignadoFaltante(ByVal wbCodiguera As Workbook, ByRef diag As Object, ByVal clave As String, ByRef arr As Variant, ByVal headers As Object, ByVal fila As Long, ByVal archivoAsignados As String)
    Dim wsC As Worksheet, hdrC As Object, lastR As Long, newR As Long, sec As Collection, rowInfo As Variant
    If Not diag.Exists("asignados_faltantes") Then Set diag("asignados_faltantes") = CreateObject("Scripting.Dictionary")
    If diag("asignados_faltantes").Exists(clave) Then Exit Sub
    rowInfo = Array("Asignados", archivoAsignados, fila, clave, "", arr(fila, ObtenerColumna(headers, Array("finac"))), arr(fila, ObtenerColumna(headers, Array("der-f"))), arr(fila, ObtenerColumna(headers, Array("pg"))), arr(fila, ObtenerColumna(headers, Array("spg"))), arr(fila, ObtenerColumna(headers, Array("proy"))), arr(fila, ObtenerColumna(headers, Array("rubro"))), arr(fila, ObtenerColumna(headers, Array("r. aux"))), arr(fila, ObtenerColumna(headers, Array("ue"))), arr(fila, ObtenerColumna(headers, Array("dep"))), arr(fila, ObtenerColumna(headers, Array("obra"))), arr(fila, ObtenerColumna(headers, Array("der. obra"))), arr(fila, ObtenerColumna(headers, Array("serv"))), arr(fila, ObtenerColumna(headers, Array("sniip"))), "Agregada a codiguera - pendiente clasificar")
    diag("asignados_faltantes")(clave) = rowInfo

    If wbCodiguera Is Nothing Then Exit Sub
    Set wsC = ObtenerHojaCodiguera(wbCodiguera)
    lastR = UltimaFilaConDatos(wsC): newR = lastR + 1
    wsC.Rows(lastR).Copy: wsC.Rows(newR).PasteSpecial xlPasteFormats: wsC.Rows(newR).PasteSpecial xlPasteValidation
    Application.CutCopyMode = False
    Set hdrC = MapearEncabezados(wsC.Range(wsC.Cells(1, 1), wsC.Cells(1, UltimaColConDatos(wsC))).Value2)
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("finac"))).Value = arr(fila, ObtenerColumna(headers, Array("finac")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("der-f"))).Value = arr(fila, ObtenerColumna(headers, Array("der-f")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("pg"))).Value = arr(fila, ObtenerColumna(headers, Array("pg")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("spg"))).Value = arr(fila, ObtenerColumna(headers, Array("spg")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("proy"))).Value = arr(fila, ObtenerColumna(headers, Array("proy")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("rubro"))).Value = arr(fila, ObtenerColumna(headers, Array("rubro")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("r. aux"))).Value = arr(fila, ObtenerColumna(headers, Array("r. aux")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("ue"))).Value = arr(fila, ObtenerColumna(headers, Array("ue")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("dep"))).Value = arr(fila, ObtenerColumna(headers, Array("dep")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("obra"))).Value = arr(fila, ObtenerColumna(headers, Array("obra")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("der. obra"))).Value = arr(fila, ObtenerColumna(headers, Array("der. obra")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("serv"))).Value = arr(fila, ObtenerColumna(headers, Array("serv")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("sniip"))).Value = arr(fila, ObtenerColumna(headers, Array("sniip")))
    wsC.Cells(newR, ObtenerColumna(hdrC, Array("clave llave presupuestal"))).Value = clave
    CompletarCodigoSiExiste wsC, hdrC, newR, "finac código numérico", arr(fila, ObtenerColumna(headers, Array("finac")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "der-f código numérico", arr(fila, ObtenerColumna(headers, Array("der-f")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "pg código numérico", arr(fila, ObtenerColumna(headers, Array("pg")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "spg código numérico", arr(fila, ObtenerColumna(headers, Array("spg")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "rubro código numérico", arr(fila, ObtenerColumna(headers, Array("rubro")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "r. aux código numérico", arr(fila, ObtenerColumna(headers, Array("r. aux")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "ue código numérico", arr(fila, ObtenerColumna(headers, Array("ue")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "dep código numérico", arr(fila, ObtenerColumna(headers, Array("dep")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "obra código numérico", arr(fila, ObtenerColumna(headers, Array("obra")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "der. obra código numérico", arr(fila, ObtenerColumna(headers, Array("der. obra")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "serv código numérico", arr(fila, ObtenerColumna(headers, Array("serv")))
    CompletarCodigoSiExiste wsC, hdrC, newR, "snip código numérico", arr(fila, ObtenerColumna(headers, Array("sniip")))
End Sub

Private Sub CompletarCodigoSiExiste(ByVal ws As Worksheet, ByVal headers As Object, ByVal fila As Long, ByVal nombreColumna As String, ByVal valor As Variant)
    Dim col As Long
    col = ObtenerColumnaOpcional(headers, Array(nombreColumna))
    If col > 0 Then ws.Cells(fila, col).Value = Split(ConstruirClaveLlavePresupuestalCodiguera(valor, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), "-")(0)
End Sub

Public Function LeerValorIndice(ByVal rutaArchivoIndice As String, ByVal anio As Long, ByVal mes As Long) As Double
    Dim wb As Workbook, ws As Worksheet, lr As Long, i As Long
    Dim periodo As String, periodoFila As String, vA As Variant, vB As Variant
    periodo = ClavePeriodoIndice(anio, mes)
    Set wb = Workbooks.Open(rutaArchivoIndice, ReadOnly:=True)
    Set ws = wb.Worksheets(1)
    lr = UltimaFilaConDatos(ws)
    For i = 1 To lr
        vA = ws.Cells(i, 1).Value
        If IsDate(vA) Then
            periodoFila = Format$(CDate(vA), "yyyy-mm")
        ElseIf IsNumeric(vA) And CDbl(vA) > 0 Then
            periodoFila = Format$(DateSerial(1899, 12, 30) + CDbl(vA), "yyyy-mm")
        Else
            periodoFila = Left$(Trim$(CStr(vA)), 7)
        End If
        If periodoFila = periodo Then
            vB = ws.Cells(i, 2).Value
            wb.Close False
            If Not IsNumeric(vB) Then Err.Raise vbObjectError + 1982, "LeerValorIndice", "El índice para " & periodo & " no es numérico. Archivo: " & rutaArchivoIndice
            LeerValorIndice = CDbl(vB)
            Exit Function
        End If
    Next i
    wb.Close False
    Err.Raise vbObjectError + 1981, "LeerValorIndice", "No se encuentra en la planilla el índice de " & LCase$(MesesES()(mes - 1)) & " de " & anio & ". Por favor, actualice el índice y vuelva a ejecutar. Archivo: " & rutaArchivoIndice
End Function

Public Function ObtenerFactorActualizacionIndice(ByVal tipoIndice As String, ByVal anioBase As Long, ByVal anioDestino As Long, ByVal mesCierre As Long, ByRef cacheFactores As Object) As Double
    Dim t As String, key As String, ruta As String, idxBase As Double, idxDestino As Double
    t = UCase$(Trim$(tipoIndice))
    If t = "IPC GRAL" Or t = "IPC GENERAL" Then t = "IPC"
    If t = "IMSN M B08" Then t = "IMSN"
    key = t & "|" & anioBase & "|" & anioDestino & "|" & mesCierre
    If cacheFactores.Exists(key) Then ObtenerFactorActualizacionIndice = cacheFactores(key): Exit Function
    ruta = ResolverArchivoIndice(t)
    idxDestino = LeerValorIndice(ruta, anioDestino, mesCierre)
    idxBase = LeerValorIndice(ruta, anioBase, mesCierre)
    If idxBase = 0 Then Err.Raise vbObjectError + 1983, "ObtenerFactorActualizacionIndice", "El índice base es 0 para " & t & " " & anioBase & "-" & Format$(mesCierre, "00") & "."
    cacheFactores.Add key, idxDestino / idxBase
    ObtenerFactorActualizacionIndice = cacheFactores(key)
End Function

Public Sub ConstruirDictComparativoActualDesdeDictAgg(ByVal dictAgg As Object, ByVal mesCierre As Long, ByRef dictCompActual As Object)
    Dim k As Variant, p() As String, keyComp As String, m As Long
    For Each k In dictAgg.Keys
        p = Split(CStr(k), "|")
        If UBound(p) >= 4 Then
            m = CLng(p(4))
            If m <= mesCierre Then
                keyComp = p(0) & "|" & p(1) & "|" & p(2) & "|" & p(3)
                If Not dictCompActual.Exists(keyComp) Then dictCompActual.Add keyComp, 0#
                dictCompActual(keyComp) = dictCompActual(keyComp) + CDbl(dictAgg(k))
            End If
        End If
    Next k
End Sub

Public Sub LeerEjecucionesComparativoYAcumular(ByVal ws As Worksheet, ByVal anioBase As Long, ByVal anioDestino As Long, ByVal mesCierre As Long, ByVal dictCod As Object, ByVal dictIndicePorClave As Object, ByRef dictCompAnteriorActualizado As Object, ByRef diag As Object)
    Dim arr As Variant, headers As Object, i As Long, fechaValor As Date, clave As String, info As Variant
    Dim keyAgg As String, importeMN As Double, tipoIndice As String, factor As Double
    Dim cacheFactores As Object
    Set cacheFactores = CreateObject("Scripting.Dictionary")
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    For i = 2 To UBound(arr, 1)
        If TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(headers, Array("fecha valor"))), fechaValor) Then
            If Year(fechaValor) = anioBase And Month(fechaValor) <= mesCierre Then
                clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, ObtenerColumna(headers, Array("finac código numérico"))), arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), arr(i, ObtenerColumna(headers, Array("pg código numérico"))), arr(i, ObtenerColumna(headers, Array("spg código numérico"))), arr(i, ObtenerColumna(headers, Array("proyecto", "proy"))), arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), arr(i, ObtenerColumna(headers, Array("ue código numérico"))), arr(i, ObtenerColumna(headers, Array("dep código numérico"))), arr(i, ObtenerColumna(headers, Array("obra código numérico"))), arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), arr(i, ObtenerColumna(headers, Array("serv código numérico"))), arr(i, ObtenerColumna(headers, Array("snip código numérico"))))
                If dictCod.Exists(clave) Then
                    If Not dictIndicePorClave.Exists(clave) Or Len(Trim$(CStr(dictIndicePorClave(clave)))) = 0 Then Err.Raise vbObjectError + 1990, "LeerEjecucionesComparativoYAcumular", "La llave " & clave & " está incluida en informe pero no tiene Indice informado en codiguera."
                    tipoIndice = CStr(dictIndicePorClave(clave))
                    factor = ObtenerFactorActualizacionIndice(tipoIndice, anioBase, anioDestino, mesCierre, cacheFactores)
                    info = dictCod(clave)
                    keyAgg = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3))
                    importeMN = CDbl(0 + arr(i, ObtenerColumna(headers, Array("importe moneda nacional")))) * factor
                    If Not dictCompAnteriorActualizado.Exists(keyAgg) Then dictCompAnteriorActualizado.Add keyAgg, 0#
                    dictCompAnteriorActualizado(keyAgg) = dictCompAnteriorActualizado(keyAgg) + importeMN
                End If
            End If
        End If
    Next i
End Sub

Public Sub ConstruirBaseEjecComparada(ByVal ws As Worksheet, ByVal dictCompActual As Object, ByVal dictCompAnteriorActualizado As Object, ByVal anioActual As Long, ByVal anioComparativo As Long)
    Dim d As Object, k As Variant, f As Long, p() As String, vA As Double, vE As Double, factor As Double
    Set d = CreateObject("Scripting.Dictionary")
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Clasificación", "Tipo", "concepto", "Ejecutado " & anioActual & ".", "Ejecutado " & anioComparativo & " a valores " & anioActual & ".", "Variación.", "Financiamiento")
    For Each k In dictCompActual.Keys: d(k) = True: Next k
    For Each k In dictCompAnteriorActualizado.Keys: d(k) = True: Next k
    factor = FactorEscalaImporte()
    f = 2
    For Each k In d.Keys
        p = Split(CStr(k), "|")
        vA = 0#: vE = 0#
        If dictCompActual.Exists(k) Then vA = CDbl(dictCompActual(k))
        If dictCompAnteriorActualizado.Exists(k) Then vE = CDbl(dictCompAnteriorActualizado(k))
        ws.Cells(f, 1).Value = p(1): ws.Cells(f, 2).Value = p(2): ws.Cells(f, 3).Value = p(3)
        ws.Cells(f, 4).Value = vA / factor
        ws.Cells(f, 5).Value = vE / factor
        If vE <> 0 Then ws.Cells(f, 6).Value = (vA - vE) / vE
        ws.Cells(f, 7).Value = p(0)
        f = f + 1
    Next k
End Sub

Public Sub ValidarAsignadosComparativoContraCodiguera(ByVal ws As Worksheet, ByVal dictCod As Object, ByVal dictLlavesCodiguera As Object, ByVal dictIndicePorClave As Object, ByRef diag As Object, ByVal wbCodiguera As Workbook, ByVal anioFiltro As Long, ByVal archivoAsignados As String)
    Dim arr As Variant, headers As Object, i As Long, clave As String, idx As String
    Dim colAnio As Long, anioFila As Long
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    ValidarColumnasAsignados headers
    colAnio = ObtenerColumnaOpcional(headers, Array("año", "anio", "ejercicio", "ej"))
    For i = 2 To UBound(arr, 1)
        If colAnio > 0 And anioFiltro > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                If anioFila <> anioFiltro Then GoTo SF
            End If
        End If
        clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, ObtenerColumna(headers, Array("finac"))), arr(i, ObtenerColumna(headers, Array("der-f"))), arr(i, ObtenerColumna(headers, Array("pg"))), arr(i, ObtenerColumna(headers, Array("spg"))), arr(i, ObtenerColumna(headers, Array("proy"))), arr(i, ObtenerColumna(headers, Array("rubro"))), arr(i, ObtenerColumna(headers, Array("r. aux"))), arr(i, ObtenerColumna(headers, Array("ue"))), arr(i, ObtenerColumna(headers, Array("dep"))), arr(i, ObtenerColumna(headers, Array("obra"))), arr(i, ObtenerColumna(headers, Array("der. obra"))), arr(i, ObtenerColumna(headers, Array("serv"))), arr(i, ObtenerColumna(headers, Array("sniip"))))
        If Not dictLlavesCodiguera.Exists(clave) Then
            RegistrarYAgregarLlaveAsignadoFaltante wbCodiguera, diag, clave, arr, headers, i, archivoAsignados
            If diag.Exists("asignados_faltantes") Then Set diag("comparativo_asignados_faltantes") = diag("asignados_faltantes")
            If Not dictLlavesCodiguera.Exists(clave) Then dictLlavesCodiguera.Add clave, True
        ElseIf Not dictCod.Exists(clave) Then
            AgregarMuestraAsignadoNoAcumulado diag, "Comparativo: existe en codiguera pero no incluida", i, clave, CDbl(0 + arr(i, ObtenerColumna(headers, Array("asignado"))))
        Else
            If Not dictIndicePorClave.Exists(clave) Or Len(Trim$(CStr(dictIndicePorClave(clave)))) = 0 Then
                Err.Raise vbObjectError + 1994, "ValidarAsignadosComparativoContraCodiguera", "La llave " & clave & " está incluida en informe pero no tiene Indice informado en codiguera."
            End If
            idx = UCase$(Trim$(CStr(dictIndicePorClave(clave))))
            If idx = "IPC GRAL" Or idx = "IPC GENERAL" Then idx = "IPC"
            If idx = "IMSN M B08" Then idx = "IMSN"
            If idx <> "IPC" And idx <> "IMSN" Then
                Err.Raise vbObjectError + 1995, "ValidarAsignadosComparativoContraCodiguera", "Índice inválido para llave " & clave & ": '" & dictIndicePorClave(clave) & "'. Archivo: " & archivoAsignados
            End If
        End If
SF:
    Next i
End Sub

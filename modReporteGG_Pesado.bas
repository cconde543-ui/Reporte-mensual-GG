Option Explicit

Private Const NOMBRE_HOJA_BASE_EJEC As String = "Base_Detalle_Ejecuciones"
Private Const NOMBRE_HOJA_BASE_ASIG As String = "Base_Detalle_Asignados"
Private Const NOMBRE_HOJA_BASE_TD As String = "Base_TD_Ejec_Asig"
Private Const NOMBRE_HOJA_BASE_COMP As String = "Base_Comparativo_Detalle"
Private Const NOMBRE_HOJA_TD_EJEC As String = "TD_Ejecucion_Detalle"
Private Const NOMBRE_HOJA_TD_COMBINADA As String = "TD_Ejec_Asig_Detalle"
Private Const NOMBRE_HOJA_TD_COMP As String = "TD_Comparativo_Anual"
Private Const NOMBRE_HOJA_RESUMEN_CONTROL As String = "Resumen_Control"
Private Const NOMBRE_HOJA_CONTROL_NO_INCLUIDAS As String = "Control_No_Incluidas"
Private Const NOMBRE_TABLA_BASE_EJEC As String = "tblBaseDetalleEjecGG"
Private Const NOMBRE_TABLA_BASE_ASIG As String = "tblBaseDetalleAsigGG"
Private Const NOMBRE_TABLA_BASE_TD As String = "tblBaseTDEjecAsigGG"
Private Const NOMBRE_TABLA_BASE_COMP As String = "tblBaseComparativoDetalleGG"
Private Const NOMBRE_TABLA_CONTROL_NO_INCLUIDAS As String = "tblControlNoIncluidasGG"
Private Const ORIGEN_EJECUCION As String = "Ejecución"
Private Const ORIGEN_ASIGNADO As String = "Asignado"
Private Const ORIGEN_COMPARATIVO_ANTERIOR As String = "Comparativo anterior"
Private Const MES_ASIGNADO_ANUAL As String = "Asignado anual"
Private Const MOTIVO_LLAVE_NO_ENCONTRADA As String = "LLAVE NO ENCONTRADA EN CODIGUERA"
Private Const MOTIVO_NO_INCLUIDA As String = "INCLUIR_EN_INFORME <> SI"

Private Const IDX_FINANCIAMIENTO As Long = 0
Private Const IDX_NIVEL1 As Long = 1
Private Const IDX_NIVEL2 As Long = 2
Private Const IDX_NIVEL3 As Long = 3
Private Const IDX_INCLUIR As Long = 4
Private Const IDX_ESTADO As Long = 5
Private Const IDX_INDICE As Long = 6

Public Sub Generar_TD_Detallada_GG_Desde_Panel()
    On Error GoTo EH

    Dim procedimiento As String
    Dim etapaActual As String
    Dim wsPanel As Worksheet
    Dim anio As Long
    Dim anioComparativo As Long
    Dim mesTxt As String
    Dim mesCierre As Long
    Dim archivoEjec As String
    Dim archivoCod As String
    Dim archivoAsignados As String
    Dim archivoEjecComparativo As String
    Dim carpetaEjecActual As String
    Dim carpetaAsignadosActual As String
    Dim carpetaEjecComparativo As String
    Dim wbE As Workbook
    Dim wbA As Workbook
    Dim wbC As Workbook
    Dim wbEComp As Workbook
    Dim wbAComp As Workbook
    Dim wbOut As Workbook
    Dim wsE As Worksheet
    Dim wsA As Worksheet
    Dim wsC As Worksheet
    Dim wsEComp As Worksheet
    Dim wsBaseEjec As Worksheet
    Dim wsBaseAsig As Worksheet
    Dim wsBaseTD As Worksheet
    Dim wsBaseComp As Worksheet
    Dim wsControlNoIncluidas As Worksheet
    Dim dictCodDetalle As Object
    Dim dictIndicePorClave As Object
    Dim resumen As Object
    Dim filasExcluidas As Collection
    Dim cacheDetallesIndice As Object
    Dim filaEjec As Long
    Dim filaAsig As Long
    Dim filaTD As Long
    Dim filaComp As Long
    Dim rutaFinal As String
    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errLine As Long
    Dim msg As String

    procedimiento = "Generar_TD_Detallada_GG_Desde_Panel"
    Debug.Print "Inicio Generar_TD_Detallada_GG_Desde_Panel: " & Now

    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then ChDir ThisWorkbook.Path
    On Error GoTo EH

    etapaActual = "leyendo parámetros del panel"
    Set wsPanel = ObtenerHojaPanelReportes()

    etapaActual = "validando año"
    If Not IsNumeric(wsPanel.Range("B3").Value) Then
        Err.Raise vbObjectError + 5100, procedimiento, "Año inválido en Panel Reportes!B3. Valor: '" & CStr(wsPanel.Range("B3").Value) & "'."
    End If
    anio = CLng(wsPanel.Range("B3").Value)
    If anio < 2000 Or anio > 2100 Then
        Err.Raise vbObjectError + 5101, procedimiento, "Año fuera de rango esperado en Panel Reportes!B3: " & anio
    End If
    anioComparativo = anio - 1

    etapaActual = "validando mes de cierre"
    mesTxt = CStr(wsPanel.Range("B4").Value)
    mesCierre = MesTextoANumero(mesTxt)
    If mesCierre < 1 Or mesCierre > 12 Then
        Err.Raise vbObjectError + 5102, procedimiento, "Mes inválido en Panel Reportes!B4. Use: Enero..Diciembre (Setiembre). Valor: '" & mesTxt & "'."
    End If

    etapaActual = "inicializando estructuras de trabajo"
    Set dictCodDetalle = CreateObject("Scripting.Dictionary")
    Set dictIndicePorClave = CreateObject("Scripting.Dictionary")
    Set resumen = CrearResumenDetallado()
    Set filasExcluidas = New Collection
    Set cacheDetallesIndice = CreateObject("Scripting.Dictionary")

    etapaActual = "resolviendo archivo de ejecuciones actual"
    carpetaEjecActual = RutaCarpetaEjecucionesAnioActiva(anio)
    If Len(Dir(carpetaEjecActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 5103, procedimiento, "No existe la carpeta de ejecuciones del año " & anio & ": " & carpetaEjecActual
    End If
    archivoEjec = ObtenerArchivoMasReciente(carpetaEjecActual)
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 5104, procedimiento, "No se encontró archivo de ejecuciones del año " & anio & " en: " & carpetaEjecActual
    End If

    etapaActual = "resolviendo archivo de asignados actual"
    carpetaAsignadosActual = RutaCarpetaAsignadosGastosAnioActiva(anio)
    If Len(Dir(carpetaAsignadosActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 5105, procedimiento, "No existe la carpeta de asignados del año " & anio & ": " & carpetaAsignadosActual
    End If
    archivoAsignados = ObtenerArchivoMasRecientePorFechaCreacion(carpetaAsignadosActual)
    If Len(archivoAsignados) = 0 Then
        Err.Raise vbObjectError + 5106, procedimiento, "No se encontró archivo de asignados del año " & anio & " en: " & carpetaAsignadosActual
    End If

    etapaActual = "resolviendo archivo de ejecuciones comparativo"
    carpetaEjecComparativo = RutaCarpetaEjecucionesAnioActiva(anioComparativo)
    If Len(Dir(carpetaEjecComparativo, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 5107, procedimiento, "No existe la carpeta de ejecuciones del año comparativo " & anioComparativo & ": " & carpetaEjecComparativo
    End If
    archivoEjecComparativo = ObtenerArchivoMasReciente(carpetaEjecComparativo)
    If Len(archivoEjecComparativo) = 0 Then
        Err.Raise vbObjectError + 5108, procedimiento, "No se encontró archivo de ejecuciones del año comparativo " & anioComparativo & " en: " & carpetaEjecComparativo
    End If

    etapaActual = "resolviendo archivo de codiguera"
    archivoCod = ResolverArchivoCodiguera(RutaCodigueraActiva())
    If Len(archivoCod) = 0 Then
        Err.Raise vbObjectError + 5109, procedimiento, "No se encontró archivo de codiguera en: " & RutaCodigueraActiva()
    End If

    etapaActual = "abriendo archivos fuente"
    Set wbE = Workbooks.Open(Filename:=archivoEjec, ReadOnly:=True, UpdateLinks:=False)
    Set wbA = Workbooks.Open(Filename:=archivoAsignados, ReadOnly:=True, UpdateLinks:=False)
    Set wbC = Workbooks.Open(Filename:=archivoCod, ReadOnly:=True, UpdateLinks:=False)
    Set wbEComp = Workbooks.Open(Filename:=archivoEjecComparativo, ReadOnly:=True, UpdateLinks:=False)

    etapaActual = "obteniendo hojas fuente"
    Set wsE = ObtenerHojaEjecuciones(wbE)
    Set wsA = ObtenerHojaAsignados(wbA)
    Set wsC = ObtenerHojaCodiguera(wbC)
    Set wsEComp = ObtenerHojaEjecuciones(wbEComp)

    etapaActual = "cargando codiguera detallada"
    LeerCodigueraDetalleDetallada wsC, dictCodDetalle
    LeerCodigueraIndices wsC, dictIndicePorClave

    etapaActual = "creando workbook detallado"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBaseEjec = wbOut.Worksheets(1)
    wsBaseEjec.Name = NOMBRE_HOJA_BASE_EJEC
    Set wsBaseAsig = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_BASE_ASIG)
    Set wsBaseTD = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_BASE_TD)
    Set wsBaseComp = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_BASE_COMP)
    Set wsControlNoIncluidas = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_CONTROL_NO_INCLUIDAS)

    etapaActual = "preparando encabezados"
    PrepararHojaDetalleConOriginales wsBaseEjec, wsE, EncabezadosBaseTDDetallada(), "Ejec_"
    PrepararHojaDetalleConOriginales wsBaseAsig, wsA, EncabezadosBaseTDDetallada(), "Asig_"
    EscribirEncabezados wsBaseTD, EncabezadosBaseTDDetallada()
    EscribirEncabezados wsBaseComp, EncabezadosBaseComparativoDetallada()
    EscribirEncabezados wsControlNoIncluidas, EncabezadosControlNoIncluidas()

    filaEjec = 2
    filaAsig = 2
    filaTD = 2
    filaComp = 2

    etapaActual = "construyendo bases de ejecuciones actuales"
    AgregarFilasEjecucionDetallada wsBaseEjec, wsBaseTD, wsBaseComp, filaEjec, filaTD, filaComp, wsE, dictCodDetalle, anio, mesCierre, archivoEjec, resumen, filasExcluidas

    etapaActual = "construyendo base de asignados actual"
    AgregarFilasAsignadoDetallada wsBaseAsig, wsBaseTD, filaAsig, filaTD, wsA, dictCodDetalle, anio, archivoAsignados, resumen, filasExcluidas

    etapaActual = "construyendo base comparativa anual"
    AgregarFilasComparativoAnteriorDetallada wsBaseComp, filaComp, wsEComp, dictCodDetalle, dictIndicePorClave, anioComparativo, anio, mesCierre, archivoEjecComparativo, resumen, filasExcluidas, cacheDetallesIndice

    etapaActual = "validando límites de filas"
    ValidarLimiteFilasHoja wsBaseEjec, NOMBRE_HOJA_BASE_EJEC
    ValidarLimiteFilasHoja wsBaseAsig, NOMBRE_HOJA_BASE_ASIG
    ValidarLimiteFilasHoja wsBaseTD, NOMBRE_HOJA_BASE_TD
    ValidarLimiteFilasHoja wsBaseComp, NOMBRE_HOJA_BASE_COMP
    ValidarLimiteFilasHoja wsControlNoIncluidas, NOMBRE_HOJA_CONTROL_NO_INCLUIDAS

    etapaActual = "creando detalle de líneas excluidas"
    CrearDetalleNoIncluidas wsControlNoIncluidas, filasExcluidas

    etapaActual = "creando tablas de bases"
    CrearTablaEnHoja wsBaseEjec, NOMBRE_TABLA_BASE_EJEC, True
    CrearTablaEnHoja wsBaseAsig, NOMBRE_TABLA_BASE_ASIG, True
    CrearTablaEnHoja wsBaseTD, NOMBRE_TABLA_BASE_TD, True
    CrearTablaEnHoja wsBaseComp, NOMBRE_TABLA_BASE_COMP, True
    CrearTablaEnHoja wsControlNoIncluidas, NOMBRE_TABLA_CONTROL_NO_INCLUIDAS, False

    etapaActual = "creando tablas dinámicas detalladas"
    CrearTablasDinamicasDetalladas wbOut

    etapaActual = "creando Resumen_Control"
    CrearResumenControlDetallado wbOut, resumen, archivoEjec, archivoAsignados, archivoEjecComparativo, archivoCod, anio, mesCierre, rutaFinal, filasExcluidas.Count

    etapaActual = "guardando archivo detallado"
    rutaFinal = GuardarReporteDetalladoGG(wbOut, anio, mesCierre)
    wbOut.Worksheets(NOMBRE_HOJA_RESUMEN_CONTROL).Range("B8").Value = rutaFinal
    wbOut.Save

    etapaActual = "cerrando archivo generado"
    CerrarWorkbookSeguro wbOut, False
    Set wbOut = Nothing

    etapaActual = "cerrando archivos auxiliares"
    CerrarWorkbookSeguro wbE, False
    Set wbE = Nothing
    CerrarWorkbookSeguro wbA, False
    Set wbA = Nothing
    CerrarWorkbookSeguro wbC, False
    Set wbC = Nothing
    CerrarWorkbookSeguro wbEComp, False
    Set wbEComp = Nothing
    If Not wbAComp Is Nothing Then CerrarWorkbookSeguro wbAComp, False
    Set wbAComp = Nothing

    MsgBox "Archivo detallado generado correctamente:" & vbCrLf & rutaFinal, vbInformation
    Debug.Print "Fin Generar_TD_Detallada_GG_Desde_Panel: " & Now & " | " & rutaFinal
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl

    On Error Resume Next
    If Not wbOut Is Nothing Then CerrarWorkbookSeguro wbOut, False
    CerrarWorkbookSeguro wbE, False
    CerrarWorkbookSeguro wbA, False
    CerrarWorkbookSeguro wbC, False
    CerrarWorkbookSeguro wbEComp, False
    If Not wbAComp Is Nothing Then CerrarWorkbookSeguro wbAComp, False
    On Error GoTo 0

    msg = "Error al generar TD detallada GG." & vbCrLf & vbCrLf
    msg = msg & "Procedimiento: " & procedimiento & vbCrLf
    msg = msg & "Etapa: " & etapaActual & vbCrLf
    msg = msg & "Err.Number: " & errNum & vbCrLf
    msg = msg & "Err.Description: " & errDesc & vbCrLf
    msg = msg & "Err.Source: " & errSource & vbCrLf
    msg = msg & "Erl: " & errLine & vbCrLf
    msg = msg & "Año: " & anio & vbCrLf
    msg = msg & "Mes cierre: " & mesCierre & vbCrLf
    msg = msg & "Archivo ejecuciones actual: " & IIf(Len(archivoEjec) > 0, archivoEjec, "(no detectado)") & vbCrLf
    msg = msg & "Archivo asignados actual: " & IIf(Len(archivoAsignados) > 0, archivoAsignados, "(no detectado)") & vbCrLf
    msg = msg & "Archivo ejecuciones comparativo: " & IIf(Len(archivoEjecComparativo) > 0, archivoEjecComparativo, "(no detectado)") & vbCrLf
    msg = msg & "Archivo codiguera: " & IIf(Len(archivoCod) > 0, archivoCod, "(no detectado)") & vbCrLf
    msg = msg & "Carpeta ejecuciones actual: " & IIf(Len(carpetaEjecActual) > 0, carpetaEjecActual, "(no determinada)") & vbCrLf
    msg = msg & "Carpeta asignados actual: " & IIf(Len(carpetaAsignadosActual) > 0, carpetaAsignadosActual, "(no determinada)") & vbCrLf
    msg = msg & "Carpeta ejecuciones comparativo: " & IIf(Len(carpetaEjecComparativo) > 0, carpetaEjecComparativo, "(no determinada)") & vbCrLf
    msg = msg & "Salida: " & IIf(Len(rutaFinal) > 0, rutaFinal, RutaReportesGeneradosActiva()) & vbCrLf
    msg = msg & vbCrLf & DiagnosticoRutasActivas() & vbCrLf
    msg = msg & vbCrLf & "Workbooks abiertos:" & vbCrLf & DiagnosticoWorkbooksAbiertos()

    Debug.Print String(100, "-")
    Debug.Print msg
    Debug.Print String(100, "-")
    MsgBox msg, vbCritical
End Sub

Public Sub Generar_TD_Pesada_GG_Desde_Panel()
    Generar_TD_Detallada_GG_Desde_Panel
End Sub

Private Function CrearResumenDetallado() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("total_ejecutado_incluido") = 0#
    d("total_asignado_incluido") = 0#
    d("total_comp_actual") = 0#
    d("total_comp_anterior_original") = 0#
    d("total_comp_anterior_actualizado") = 0#
    InicializarContadoresOrigen d, "ejec"
    InicializarContadoresOrigen d, "asig"
    InicializarContadoresOrigen d, "comp_ant"
    Set CrearResumenDetallado = d
End Function

Private Sub InicializarContadoresOrigen(ByVal d As Object, ByVal prefijo As String)
    d(prefijo & "_leidas") = 0&
    d(prefijo & "_incluidas") = 0&
    d(prefijo & "_no_incluidas") = 0&
    d(prefijo & "_llave_no_encontrada") = 0&
End Sub

Private Sub LeerCodigueraDetalleDetallada(ByVal ws As Worksheet, ByRef dictCodDetalle As Object)
    On Error GoTo EH

    Dim arr As Variant
    Dim headers As Object
    Dim i As Long
    Dim clave As String
    Dim info As Variant
    Dim colFinac As Long, colDerF As Long, colPg As Long, colSpg As Long, colProy As Long
    Dim colRubro As Long, colRAux As Long, colUe As Long, colDep As Long, colObra As Long, colDerObra As Long
    Dim colServ As Long, colSniip As Long, colTitular As Long, colN1 As Long, colN2 As Long, colN3 As Long
    Dim colIncluir As Long, colEstado As Long, colClaveTexto As Long, colIndice As Long

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    colFinac = ObtenerColumnaDetallada(headers, Array("finac", "finac código numérico"))
    colDerF = ObtenerColumnaDetallada(headers, Array("der-f", "der-f código numérico"))
    colPg = ObtenerColumnaDetallada(headers, Array("pg", "pg código numérico"))
    colSpg = ObtenerColumnaDetallada(headers, Array("spg", "spg código numérico"))
    colProy = ObtenerColumnaDetallada(headers, Array("proy", "proyecto", "proyecto código numérico", "proy código numérico"))
    colRubro = ObtenerColumnaDetallada(headers, Array("rubro", "rubro código numérico"))
    colRAux = ObtenerColumnaDetallada(headers, Array("r. aux", "r. aux código numérico"))
    colUe = ObtenerColumnaDetallada(headers, Array("ue", "ue código numérico"))
    colDep = ObtenerColumnaDetallada(headers, Array("dep", "dep código numérico"))
    colObra = ObtenerColumnaDetallada(headers, Array("obra", "obra código numérico"))
    colDerObra = ObtenerColumnaDetallada(headers, Array("der. obra", "der. obra código numérico"))
    colServ = ObtenerColumnaDetallada(headers, Array("serv", "serv código numérico"))
    colSniip = ObtenerColumnaDetallada(headers, Array("sniip", "snip", "sniip código numérico", "snip código numérico"))
    colTitular = ObtenerColumna(headers, Array("titular"))
    colN1 = ObtenerColumna(headers, Array("nivel_1"))
    colN2 = ObtenerColumna(headers, Array("nivel_2"))
    colN3 = ObtenerColumna(headers, Array("nivel_3"))
    colIncluir = ObtenerColumna(headers, Array("incluir_en_informe"))
    colEstado = ObtenerColumnaDetallada(headers, Array("estado_codiguera", "estado"))
    colClaveTexto = ObtenerColumnaDetallada(headers, Array("clave llave presupuestal"))
    colIndice = ObtenerColumnaDetallada(headers, Array("indice"))

    For i = 2 To UBound(arr, 1)
        If colFinac > 0 And colDerF > 0 And colPg > 0 And colSpg > 0 And colRubro > 0 And colRAux > 0 And colUe > 0 And colDep > 0 And colObra > 0 And colDerObra > 0 And colServ > 0 And colSniip > 0 Then
            clave = ConstruirClaveLlavePresupuestalCodiguera(ValorMatriz(arr, i, colFinac), ValorMatriz(arr, i, colDerF), ValorMatriz(arr, i, colPg), ValorMatriz(arr, i, colSpg), ValorMatriz(arr, i, colProy), ValorMatriz(arr, i, colRubro), ValorMatriz(arr, i, colRAux), ValorMatriz(arr, i, colUe), ValorMatriz(arr, i, colDep), ValorMatriz(arr, i, colObra), ValorMatriz(arr, i, colDerObra), ValorMatriz(arr, i, colServ), ValorMatriz(arr, i, colSniip))
        ElseIf colClaveTexto > 0 Then
            clave = NormalizarClaveTextoDetallada(arr(i, colClaveTexto))
        Else
            clave = ""
        End If

        If Len(clave) > 0 Then
            info = Array( _
                TextoSeguro(ValorMatriz(arr, i, colTitular)), _
                TextoSeguro(ValorMatriz(arr, i, colN1)), _
                TextoSeguro(ValorMatriz(arr, i, colN2)), _
                TextoSeguro(ValorMatriz(arr, i, colN3)), _
                NormalizarIncluirDetallado(ValorMatriz(arr, i, colIncluir)), _
                TextoSeguro(ValorMatriz(arr, i, colEstado)), _
                TextoSeguro(ValorMatriz(arr, i, colIndice)) _
            )
            dictCodDetalle(clave) = info
        End If
    Next i
    Exit Sub
EH:
    Err.Raise Err.Number, "LeerCodigueraDetalleDetallada", "Error leyendo codiguera completa: " & Err.Description
End Sub

Private Sub PrepararHojaDetalleConOriginales(ByVal ws As Worksheet, ByVal wsFuente As Worksheet, ByVal baseHeaders As Variant, ByVal prefijoOriginal As String)
    Dim arr As Variant
    Dim pref As Variant
    Dim allHeaders As Variant

    arr = wsFuente.Range(wsFuente.Cells(1, 1), wsFuente.Cells(UltimaFilaConDatos(wsFuente), UltimaColConDatos(wsFuente))).Value2
    pref = EncabezadosPrefijados(arr, prefijoOriginal)
    allHeaders = UnirEncabezados(baseHeaders, pref)
    EscribirEncabezados ws, allHeaders
End Sub

Private Sub AgregarFilasEjecucionDetallada(ByVal wsDetalle As Worksheet, ByVal wsTD As Worksheet, ByVal wsComp As Worksheet, ByRef filaDetalle As Long, ByRef filaTD As Long, ByRef filaComp As Long, ByVal wsE As Worksheet, ByVal dictCodDetalle As Object, ByVal anio As Long, ByVal mesCierre As Long, ByVal archivoOrigen As String, ByRef resumen As Object, ByRef filasExcluidas As Collection)
    Dim arr As Variant, headers As Object, i As Long, j As Long, fechaValor As Date, importeMN As Double
    Dim clave As String, info As Variant, incluir As String, motivo As String, meses As Variant
    Dim cols As Variant, baseVals As Variant, compVals As Variant, colImporte As Long, colOffsetOriginal As Long

    meses = MesesES()
    arr = wsE.Range(wsE.Cells(1, 1), wsE.Cells(UltimaFilaConDatos(wsE), UltimaColConDatos(wsE))).Value2
    Set headers = MapearEncabezados(arr)
    cols = ColumnasLlaveEjecucion(headers)
    colImporte = ObtenerColumna(headers, Array("importe moneda nacional"))
    colOffsetOriginal = UBound(EncabezadosBaseTDDetallada()) + 2

    For i = 2 To UBound(arr, 1)
        If Not TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(headers, Array("fecha valor"))), fechaValor) Then GoTo SiguienteFila
        If Year(fechaValor) <> anio Or Month(fechaValor) > mesCierre Then GoTo SiguienteFila

        resumen("ejec_leidas") = CLng(resumen("ejec_leidas")) + 1
        If Not TryDoubleSeguro(arr(i, colImporte), importeMN) Then importeMN = 0#
        clave = ClaveDesdeFila(arr, i, cols)
        info = InfoCodigueraParaClave(dictCodDetalle, clave)
        incluir = CStr(info(IDX_INCLUIR))
        motivo = MotivoExclusion(dictCodDetalle, clave, incluir)

        If Len(motivo) > 0 Then
            RegistrarExclusion resumen, filasExcluidas, "ejec", ORIGEN_EJECUCION, archivoOrigen, wsE.Name, i, anio, Month(fechaValor), meses(Month(fechaValor) - 1), clave, motivo, info, importeMN, importeMN, 0#
            GoTo SiguienteFila
        End If

        baseVals = BaseValsTD(ORIGEN_EJECUCION, archivoOrigen, wsE.Name, i, anio, Month(fechaValor), meses(Month(fechaValor) - 1), fechaValor, clave, info, arr, i, cols, importeMN, 0#, importeMN)
        ValidarFilaSalida filaDetalle, NOMBRE_HOJA_BASE_EJEC
        wsDetalle.Range(wsDetalle.Cells(filaDetalle, 1), wsDetalle.Cells(filaDetalle, UBound(baseVals) + 1)).Value = baseVals
        For j = 1 To UBound(arr, 2)
            wsDetalle.Cells(filaDetalle, colOffsetOriginal + j - 1).Value = arr(i, j)
        Next j
        filaDetalle = filaDetalle + 1

        ValidarFilaSalida filaTD, NOMBRE_HOJA_BASE_TD
        wsTD.Range(wsTD.Cells(filaTD, 1), wsTD.Cells(filaTD, UBound(baseVals) + 1)).Value = baseVals
        filaTD = filaTD + 1

        compVals = BaseValsComparativo("ACTUAL", ORIGEN_EJECUCION, archivoOrigen, wsE.Name, i, anio, anio, Month(fechaValor), meses(Month(fechaValor) - 1), fechaValor, clave, info, arr, i, cols, importeMN, "", "", "", "", "", "", 1#, importeMN, importeMN, 0#)
        ValidarFilaSalida filaComp, NOMBRE_HOJA_BASE_COMP
        wsComp.Range(wsComp.Cells(filaComp, 1), wsComp.Cells(filaComp, UBound(compVals) + 1)).Value = compVals
        filaComp = filaComp + 1

        resumen("ejec_incluidas") = CLng(resumen("ejec_incluidas")) + 1
        resumen("total_ejecutado_incluido") = CDbl(resumen("total_ejecutado_incluido")) + importeMN
        resumen("total_comp_actual") = CDbl(resumen("total_comp_actual")) + importeMN
SiguienteFila:
    Next i
End Sub

Private Sub AgregarFilasAsignadoDetallada(ByVal wsDetalle As Worksheet, ByVal wsTD As Worksheet, ByRef filaDetalle As Long, ByRef filaTD As Long, ByVal wsA As Worksheet, ByVal dictCodDetalle As Object, ByVal anio As Long, ByVal archivoOrigen As String, ByRef resumen As Object, ByRef filasExcluidas As Collection)
    Dim arr As Variant, headers As Object, i As Long, j As Long, monto As Double, clave As String
    Dim info As Variant, incluir As String, motivo As String, usarFila As Boolean, anioFila As Long
    Dim colAnio As Long, colAsignado As Long, cols As Variant, baseVals As Variant, colOffsetOriginal As Long

    arr = wsA.Range(wsA.Cells(1, 1), wsA.Cells(UltimaFilaConDatos(wsA), UltimaColConDatos(wsA))).Value2
    Set headers = MapearEncabezados(arr)
    colAnio = ObtenerColumnaDetallada(headers, Array("año", "anio", "ejercicio", "ej"))
    colAsignado = ObtenerColumna(headers, Array("asignado"))
    cols = ColumnasLlaveAsignado(headers)
    colOffsetOriginal = UBound(EncabezadosBaseTDDetallada()) + 2

    For i = 2 To UBound(arr, 1)
        usarFila = True
        If colAnio > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                usarFila = (anioFila = anio)
            End If
        End If
        If Not usarFila Then GoTo SiguienteFila

        resumen("asig_leidas") = CLng(resumen("asig_leidas")) + 1
        If Not TryDoubleSeguro(arr(i, colAsignado), monto) Then monto = 0#
        clave = ClaveDesdeFila(arr, i, cols)
        info = InfoCodigueraParaClave(dictCodDetalle, clave)
        incluir = CStr(info(IDX_INCLUIR))
        motivo = MotivoExclusion(dictCodDetalle, clave, incluir)

        If Len(motivo) > 0 Then
            RegistrarExclusion resumen, filasExcluidas, "asig", ORIGEN_ASIGNADO, archivoOrigen, wsA.Name, i, anio, 0, MES_ASIGNADO_ANUAL, clave, motivo, info, monto, 0#, monto
            GoTo SiguienteFila
        End If

        baseVals = BaseValsTD(ORIGEN_ASIGNADO, archivoOrigen, wsA.Name, i, anio, 0, MES_ASIGNADO_ANUAL, "", clave, info, arr, i, cols, 0#, monto, monto)
        ValidarFilaSalida filaDetalle, NOMBRE_HOJA_BASE_ASIG
        wsDetalle.Range(wsDetalle.Cells(filaDetalle, 1), wsDetalle.Cells(filaDetalle, UBound(baseVals) + 1)).Value = baseVals
        For j = 1 To UBound(arr, 2)
            wsDetalle.Cells(filaDetalle, colOffsetOriginal + j - 1).Value = arr(i, j)
        Next j
        filaDetalle = filaDetalle + 1

        ValidarFilaSalida filaTD, NOMBRE_HOJA_BASE_TD
        wsTD.Range(wsTD.Cells(filaTD, 1), wsTD.Cells(filaTD, UBound(baseVals) + 1)).Value = baseVals
        filaTD = filaTD + 1

        resumen("asig_incluidas") = CLng(resumen("asig_incluidas")) + 1
        resumen("total_asignado_incluido") = CDbl(resumen("total_asignado_incluido")) + monto
SiguienteFila:
    Next i
End Sub

Private Sub AgregarFilasComparativoAnteriorDetallada(ByVal wsComp As Worksheet, ByRef filaComp As Long, ByVal wsEComp As Worksheet, ByVal dictCodDetalle As Object, ByVal dictIndicePorClave As Object, ByVal anioBase As Long, ByVal anioDestino As Long, ByVal mesCierre As Long, ByVal archivoOrigen As String, ByRef resumen As Object, ByRef filasExcluidas As Collection, ByRef cacheDetallesIndice As Object)
    Dim arr As Variant, headers As Object, i As Long, fechaValor As Date, importeOriginal As Double, importeActualizado As Double
    Dim clave As String, info As Variant, incluir As String, motivo As String, meses As Variant, cols As Variant, compVals As Variant
    Dim colImporte As Long, tipoIndice As String, det As Variant

    meses = MesesES()
    arr = wsEComp.Range(wsEComp.Cells(1, 1), wsEComp.Cells(UltimaFilaConDatos(wsEComp), UltimaColConDatos(wsEComp))).Value2
    Set headers = MapearEncabezados(arr)
    cols = ColumnasLlaveEjecucion(headers)
    colImporte = ObtenerColumna(headers, Array("importe moneda nacional"))

    For i = 2 To UBound(arr, 1)
        If Not TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(headers, Array("fecha valor"))), fechaValor) Then GoTo SiguienteFila
        If Year(fechaValor) <> anioBase Or Month(fechaValor) > mesCierre Then GoTo SiguienteFila

        resumen("comp_ant_leidas") = CLng(resumen("comp_ant_leidas")) + 1
        If Not TryDoubleSeguro(arr(i, colImporte), importeOriginal) Then importeOriginal = 0#
        clave = ClaveDesdeFila(arr, i, cols)
        info = InfoCodigueraParaClave(dictCodDetalle, clave)
        incluir = CStr(info(IDX_INCLUIR))
        motivo = MotivoExclusion(dictCodDetalle, clave, incluir)

        If Len(motivo) > 0 Then
            RegistrarExclusion resumen, filasExcluidas, "comp_ant", ORIGEN_COMPARATIVO_ANTERIOR, archivoOrigen, wsEComp.Name, i, anioBase, Month(fechaValor), meses(Month(fechaValor) - 1), clave, motivo, info, importeOriginal, importeOriginal, 0#
            GoTo SiguienteFila
        End If

        tipoIndice = Trim$(TextoSeguro(info(IDX_INDICE)))
        If Len(tipoIndice) = 0 And dictIndicePorClave.Exists(clave) Then tipoIndice = Trim$(TextoSeguro(dictIndicePorClave(clave)))
        If Len(tipoIndice) = 0 Then
            Err.Raise vbObjectError + 5310, "AgregarFilasComparativoAnteriorDetallada", "La llave " & clave & " está incluida en informe pero tiene Indice vacío en codiguera."
        End If

        det = ObtenerDetalleActualizacionIndice(tipoIndice, anioBase, anioDestino, mesCierre, cacheDetallesIndice)
        importeActualizado = importeOriginal * CDbl(det(6))
        compVals = BaseValsComparativo("ANTERIOR_ACTUALIZADO", ORIGEN_COMPARATIVO_ANTERIOR, archivoOrigen, wsEComp.Name, i, anioBase, anioDestino, Month(fechaValor), meses(Month(fechaValor) - 1), fechaValor, clave, info, arr, i, cols, importeOriginal, det(0), det(1), det(2), det(3), det(4), det(5), det(6), importeActualizado, 0#, importeActualizado)
        ValidarFilaSalida filaComp, NOMBRE_HOJA_BASE_COMP
        wsComp.Range(wsComp.Cells(filaComp, 1), wsComp.Cells(filaComp, UBound(compVals) + 1)).Value = compVals
        filaComp = filaComp + 1

        resumen("comp_ant_incluidas") = CLng(resumen("comp_ant_incluidas")) + 1
        resumen("total_comp_anterior_original") = CDbl(resumen("total_comp_anterior_original")) + importeOriginal
        resumen("total_comp_anterior_actualizado") = CDbl(resumen("total_comp_anterior_actualizado")) + importeActualizado
SiguienteFila:
    Next i
End Sub

Private Sub RegistrarExclusion(ByRef resumen As Object, ByRef filasExcluidas As Collection, ByVal prefijoResumen As String, ByVal origen As String, ByVal archivoOrigen As String, ByVal hojaOrigen As String, ByVal filaOrigen As Long, ByVal anio As Long, ByVal mesNum As Long, ByVal mesNombre As String, ByVal clave As String, ByVal motivo As String, ByVal info As Variant, ByVal importeMN As Double, ByVal ejecutado As Double, ByVal asignado As Double)
    resumen(prefijoResumen & "_no_incluidas") = CLng(resumen(prefijoResumen & "_no_incluidas")) + 1
    If motivo = MOTIVO_LLAVE_NO_ENCONTRADA Then resumen(prefijoResumen & "_llave_no_encontrada") = CLng(resumen(prefijoResumen & "_llave_no_encontrada")) + 1
    filasExcluidas.Add Array(origen, NombreArchivoDesdeRutaDetallada(archivoOrigen), hojaOrigen, filaOrigen, anio, mesNum, mesNombre, clave, motivo, info(IDX_ESTADO), info(IDX_INCLUIR), info(IDX_FINANCIAMIENTO), info(IDX_NIVEL1), info(IDX_NIVEL2), info(IDX_NIVEL3), importeMN, ejecutado, asignado)
End Sub

Private Function MotivoExclusion(ByVal dictCodDetalle As Object, ByVal clave As String, ByVal incluir As String) As String
    If Not dictCodDetalle.Exists(clave) Then
        MotivoExclusion = MOTIVO_LLAVE_NO_ENCONTRADA
    ElseIf NormalizarIncluirDetallado(incluir) <> "SI" Then
        MotivoExclusion = MOTIVO_NO_INCLUIDA
    Else
        MotivoExclusion = ""
    End If
End Function

Private Function InfoCodigueraParaClave(ByVal dictCodDetalle As Object, ByVal clave As String) As Variant
    If dictCodDetalle.Exists(clave) Then
        InfoCodigueraParaClave = dictCodDetalle(clave)
    Else
        InfoCodigueraParaClave = Array("SIN CODIGUERA", "SIN CODIGUERA", "SIN CODIGUERA", "SIN CODIGUERA", "NO", "LLAVE NO ENCONTRADA", "")
    End If
End Function

Private Function ColumnasLlaveEjecucion(ByVal headers As Object) As Variant
    ColumnasLlaveEjecucion = Array( _
        ObtenerColumna(headers, Array("finac código numérico")), _
        ObtenerColumna(headers, Array("der-f código numérico")), _
        ObtenerColumna(headers, Array("pg código numérico")), _
        ObtenerColumna(headers, Array("spg código numérico")), _
        ObtenerColumna(headers, Array("proyecto", "proy", "proyecto código numérico", "proy código numérico")), _
        ObtenerColumna(headers, Array("rubro código numérico")), _
        ObtenerColumna(headers, Array("r. aux código numérico")), _
        ObtenerColumna(headers, Array("ue código numérico")), _
        ObtenerColumna(headers, Array("dep código numérico")), _
        ObtenerColumna(headers, Array("obra código numérico")), _
        ObtenerColumna(headers, Array("der. obra código numérico")), _
        ObtenerColumna(headers, Array("serv código numérico")), _
        ObtenerColumna(headers, Array("snip código numérico", "sniip código numérico", "snip", "sniip")) _
    )
End Function

Private Function ColumnasLlaveAsignado(ByVal headers As Object) As Variant
    ColumnasLlaveAsignado = Array( _
        ObtenerColumna(headers, Array("finac")), _
        ObtenerColumna(headers, Array("der-f")), _
        ObtenerColumna(headers, Array("pg")), _
        ObtenerColumna(headers, Array("spg")), _
        ObtenerColumna(headers, Array("proy")), _
        ObtenerColumna(headers, Array("rubro")), _
        ObtenerColumna(headers, Array("r. aux")), _
        ObtenerColumna(headers, Array("ue")), _
        ObtenerColumna(headers, Array("dep")), _
        ObtenerColumna(headers, Array("obra")), _
        ObtenerColumna(headers, Array("der. obra")), _
        ObtenerColumna(headers, Array("serv")), _
        ObtenerColumna(headers, Array("sniip", "snip")) _
    )
End Function

Private Function ClaveDesdeFila(ByRef arr As Variant, ByVal fila As Long, ByVal cols As Variant) As String
    ClaveDesdeFila = ConstruirClaveLlavePresupuestalCodiguera( _
        ValorMatriz(arr, fila, CLng(cols(0))), ValorMatriz(arr, fila, CLng(cols(1))), _
        ValorMatriz(arr, fila, CLng(cols(2))), ValorMatriz(arr, fila, CLng(cols(3))), _
        ValorMatriz(arr, fila, CLng(cols(4))), ValorMatriz(arr, fila, CLng(cols(5))), _
        ValorMatriz(arr, fila, CLng(cols(6))), ValorMatriz(arr, fila, CLng(cols(7))), _
        ValorMatriz(arr, fila, CLng(cols(8))), ValorMatriz(arr, fila, CLng(cols(9))), _
        ValorMatriz(arr, fila, CLng(cols(10))), ValorMatriz(arr, fila, CLng(cols(11))), _
        ValorMatriz(arr, fila, CLng(cols(12))))
End Function

Private Function BaseValsTD(ByVal origen As String, ByVal archivoOrigen As String, ByVal hojaOrigen As String, ByVal filaOrigen As Long, ByVal anio As Long, ByVal mesNum As Long, ByVal mesNombre As String, ByVal fechaValor As Variant, ByVal clave As String, ByVal info As Variant, ByRef arr As Variant, ByVal fila As Long, ByVal cols As Variant, ByVal ejecutado As Double, ByVal asignado As Double, ByVal importeMN As Double) As Variant
    BaseValsTD = Array(origen, NombreArchivoDesdeRutaDetallada(archivoOrigen), hojaOrigen, filaOrigen, _
        anio, mesNum, mesNombre, fechaValor, clave, info(IDX_FINANCIAMIENTO), info(IDX_NIVEL1), _
        info(IDX_NIVEL2), info(IDX_NIVEL3), ValorMatriz(arr, fila, CLng(cols(0))), _
        ValorMatriz(arr, fila, CLng(cols(1))), ValorMatriz(arr, fila, CLng(cols(2))), _
        ValorMatriz(arr, fila, CLng(cols(3))), ValorMatriz(arr, fila, CLng(cols(4))), _
        ValorMatriz(arr, fila, CLng(cols(5))), ValorMatriz(arr, fila, CLng(cols(6))), _
        ValorMatriz(arr, fila, CLng(cols(7))), ValorMatriz(arr, fila, CLng(cols(8))), _
        ValorMatriz(arr, fila, CLng(cols(9))), ValorMatriz(arr, fila, CLng(cols(10))), _
        ValorMatriz(arr, fila, CLng(cols(11))), ValorMatriz(arr, fila, CLng(cols(12))), _
        ejecutado, asignado, importeMN, "SI", info(IDX_ESTADO))
End Function

Private Function BaseValsComparativo(ByVal periodoComparativo As String, ByVal origen As String, _
        ByVal archivoOrigen As String, ByVal hojaOrigen As String, ByVal filaOrigen As Long, _
        ByVal anioOrigen As Long, ByVal anioDestino As Long, ByVal mesNum As Long, _
        ByVal mesNombre As String, ByVal fechaValor As Variant, ByVal clave As String, _
        ByVal info As Variant, ByRef arr As Variant, ByVal fila As Long, ByVal cols As Variant, _
        ByVal importeOriginal As Double, ByVal tipoIndice As Variant, ByVal archivoIndice As Variant, _
        ByVal periodoIndiceBase As Variant, ByVal valorIndiceBase As Variant, _
        ByVal periodoIndiceDestino As Variant, ByVal valorIndiceDestino As Variant, _
        ByVal factorActualizacion As Variant, ByVal importeActualizado As Double, _
        ByVal ejecutadoActual As Double, ByVal ejecutadoAnteriorActualizado As Double) As Variant
    BaseValsComparativo = Array(periodoComparativo, origen, NombreArchivoDesdeRutaDetallada(archivoOrigen), _
        hojaOrigen, filaOrigen, anioOrigen, anioDestino, mesNum, mesNombre, fechaValor, clave, _
        info(IDX_FINANCIAMIENTO), info(IDX_NIVEL1), info(IDX_NIVEL2), info(IDX_NIVEL3), _
        ValorMatriz(arr, fila, CLng(cols(0))), ValorMatriz(arr, fila, CLng(cols(1))), _
        ValorMatriz(arr, fila, CLng(cols(2))), ValorMatriz(arr, fila, CLng(cols(3))), _
        ValorMatriz(arr, fila, CLng(cols(4))), ValorMatriz(arr, fila, CLng(cols(5))), _
        ValorMatriz(arr, fila, CLng(cols(6))), ValorMatriz(arr, fila, CLng(cols(7))), _
        ValorMatriz(arr, fila, CLng(cols(8))), ValorMatriz(arr, fila, CLng(cols(9))), _
        ValorMatriz(arr, fila, CLng(cols(10))), ValorMatriz(arr, fila, CLng(cols(11))), _
        ValorMatriz(arr, fila, CLng(cols(12))), importeOriginal, tipoIndice, _
        NombreArchivoDesdeRutaDetallada(CStr(archivoIndice)), periodoIndiceBase, valorIndiceBase, _
        periodoIndiceDestino, valorIndiceDestino, factorActualizacion, importeActualizado, _
        ejecutadoActual, ejecutadoAnteriorActualizado, "SI", info(IDX_ESTADO))
End Function

Private Sub CrearDetalleNoIncluidas(ByVal ws As Worksheet, ByVal filasExcluidas As Collection)
    Dim i As Long
    Dim item As Variant

    For i = 1 To filasExcluidas.Count
        item = filasExcluidas(i)
        ValidarFilaSalida i + 1, NOMBRE_HOJA_CONTROL_NO_INCLUIDAS
        ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, UBound(item) + 1)).Value = item
    Next i
    ws.Columns.AutoFit
End Sub

Private Sub CrearTablaEnHoja(ByVal ws As Worksheet, ByVal nombreTabla As String, ByVal requerirDatos As Boolean)
    Dim lastRow As Long, lastCol As Long
    Dim lo As ListObject

    lastRow = UltimaFilaConDatos(ws)
    lastCol = UltimaColConDatos(ws)
    If requerirDatos And lastRow < 2 Then Err.Raise vbObjectError + 5200, "CrearTablaEnHoja", "La hoja " & ws.Name & " no tiene filas de datos."

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    lo.Name = nombreTabla
    lo.TableStyle = "TableStyleMedium2"
    ws.Columns.AutoFit
End Sub

Private Sub CrearTablasDinamicasDetalladas(ByVal wbOut As Workbook)
    On Error GoTo EH

    Dim etapaTD As String

    etapaTD = "creando TD_Ejecucion_Detalle"
    CrearPivotEjecucionDetallada wbOut

    etapaTD = "creando TD_Ejec_Asig_Detalle"
    CrearPivotCombinadaDetallada wbOut

    etapaTD = "creando TD_Comparativo_Anual"
    CrearPivotComparativoAnual wbOut

    Exit Sub

EH:
    Err.Raise Err.Number, "CrearTablasDinamicasDetalladas", _
        "Falló etapa interna: " & etapaTD & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Sub CrearPivotEjecucionDetallada(ByVal wbOut As Workbook)
    Dim ws As Worksheet, pt As PivotTable, pc As PivotCache
    Dim lo As ListObject

    Set lo = ObtenerTablaBasePivotDetallada(wbOut, NOMBRE_HOJA_BASE_EJEC, NOMBRE_TABLA_BASE_EJEC, _
        Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNombre", "Ejecutado"))
    Set ws = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_TD_EJEC)
    Set pc = CrearPivotCacheTabla(wbOut, lo)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptEjecucionDetalleGG")
    ConfigurarFilasClasificacion pt
    ConfigurarCampoPivotDetallado pt, "MesNombre", xlColumnField, 1
    OrdenarMesesPivotDetallado pt, "MesNombre"
    AgregarDataFieldDetallado pt, "Ejecutado", "Suma de Ejecutado", xlSum, "#,##0"
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Sub CrearPivotCombinadaDetallada(ByVal wbOut As Workbook)
    Dim ws As Worksheet, pt As PivotTable, pc As PivotCache
    Dim lo As ListObject

    Set lo = ObtenerTablaBasePivotDetallada(wbOut, NOMBRE_HOJA_BASE_TD, NOMBRE_TABLA_BASE_TD, _
        Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Incluir_en_Informe", "Ejecutado", "Asignado"))
    Set ws = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_TD_COMBINADA)
    Set pc = CrearPivotCacheTabla(wbOut, lo)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptEjecAsigDetalleGG")
    ConfigurarCampoPivotDetallado pt, "Incluir_en_Informe", xlPageField, 1
    FiltrarPivotCampoDetallado pt, "Incluir_en_Informe", "SI"
    ConfigurarFilasClasificacion pt
    AgregarDataFieldDetallado pt, "Ejecutado", "Suma de Ejecutado", xlSum, "#,##0"
    AgregarDataFieldDetallado pt, "Asignado", "Suma de Asignado", xlSum, "#,##0"
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Sub CrearPivotComparativoAnual(ByVal wbOut As Workbook)
    Dim ws As Worksheet, pt As PivotTable, pc As PivotCache
    Dim lo As ListObject

    Set lo = ObtenerTablaBasePivotDetallada(wbOut, NOMBRE_HOJA_BASE_COMP, NOMBRE_TABLA_BASE_COMP, _
        Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Ejecutado_Actual", "Ejecutado_Anterior_Actualizado"))
    Set ws = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_TD_COMP)
    Set pc = CrearPivotCacheTabla(wbOut, lo)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptComparativoAnualDetalleGG")
    ConfigurarFilasClasificacion pt
    AgregarDataFieldDetallado pt, "Ejecutado_Actual", "Suma de Ejecutado_Actual", xlSum, "#,##0"
    AgregarDataFieldDetallado pt, "Ejecutado_Anterior_Actualizado", "Suma de Ejecutado_Anterior_Actualizado", xlSum, "#,##0"
    AsegurarCampoCalculadoPctVariacionDetalle pt
    AgregarDataFieldDetallado pt, "PctVariacion", "% Variación", xlSum, "0.0%"
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Function ObtenerTablaBasePivotDetallada(ByVal wbOut As Workbook, ByVal nombreHoja As String, ByVal nombreTabla As String, ByVal camposRequeridos As Variant) As ListObject
    On Error GoTo EH

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim campo As Variant

    Set ws = wbOut.Worksheets(nombreHoja)
    Set lo = ws.ListObjects(nombreTabla)

    If lo.ListRows.Count = 0 Then
        Err.Raise vbObjectError + 5220, "ObtenerTablaBasePivotDetallada", _
            "La tabla base requerida para tabla dinámica no tiene filas de datos. Tabla=" & nombreTabla & _
            " | Hoja=" & nombreHoja
    End If

    For Each campo In camposRequeridos
        If Not ListColumnExisteDetallada(lo, CStr(campo)) Then
            Err.Raise vbObjectError + 5221, "ObtenerTablaBasePivotDetallada", _
                "No existe el campo requerido para tabla dinámica. Tabla=" & nombreTabla & _
                " | Hoja=" & nombreHoja & _
                " | Campo=" & CStr(campo) & _
                " | Campos disponibles=" & CamposDisponiblesTablaDetallada(lo)
        End If
    Next campo

    Set ObtenerTablaBasePivotDetallada = lo
    Exit Function

EH:
    Err.Raise Err.Number, "ObtenerTablaBasePivotDetallada", _
        "No se pudo validar la tabla base para tabla dinámica. Tabla=" & nombreTabla & _
        " | Hoja=" & nombreHoja & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Function

Private Function ListColumnExisteDetallada(ByVal lo As ListObject, ByVal nombreCampo As String) As Boolean
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nombreCampo, vbBinaryCompare) = 0 Then
            ListColumnExisteDetallada = True
            Exit Function
        End If
    Next lc
End Function

Private Function CamposDisponiblesTablaDetallada(ByVal lo As ListObject) As String
    Dim lc As ListColumn
    Dim res As String

    For Each lc In lo.ListColumns
        If Len(res) > 0 Then res = res & ", "
        res = res & lc.Name
    Next lc
    CamposDisponiblesTablaDetallada = res
End Function

Private Function CrearPivotCacheTabla(ByVal wb As Workbook, ByVal lo As ListObject) As PivotCache
    Dim src As String
    Dim n1 As Long
    Dim d1 As String
    Dim n2 As Long
    Dim d2 As String

    On Error GoTo EH_NOMBRE
    src = lo.Name
    Set CrearPivotCacheTabla = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=src)
    Exit Function

EH_NOMBRE:
    n1 = Err.Number
    d1 = Err.Description
    Err.Clear

    On Error GoTo EH_RANGO
    src = lo.Range.Address(ReferenceStyle:=xlR1C1, External:=True)
    Set CrearPivotCacheTabla = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=src)
    Exit Function

EH_RANGO:
    n2 = Err.Number
    d2 = Err.Description
    Err.Raise n2, "CrearPivotCacheTabla", _
        "No se pudo crear PivotCache. Tabla=" & lo.Name & _
        " | Hoja=" & lo.Parent.Name & _
        " | Rango=" & lo.Range.Address(False, False) & _
        " | Filas=" & lo.Range.Rows.Count & _
        " | Columnas=" & lo.Range.Columns.Count & _
        " | Primer intento con nombre falló: " & n1 & " - " & d1 & _
        " | Segundo intento falló: " & n2 & " - " & d2
End Function

Private Sub ConfigurarFilasClasificacion(ByVal pt As PivotTable)
    ConfigurarCampoPivotDetallado pt, "Financiamiento", xlRowField, 1
    ConfigurarCampoPivotDetallado pt, "Nivel_1", xlRowField, 2
    ConfigurarCampoPivotDetallado pt, "Nivel_2", xlRowField, 3
    ConfigurarCampoPivotDetallado pt, "Nivel_3", xlRowField, 4
End Sub

Private Sub ConfigurarCampoPivotDetallado(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    On Error GoTo EH

    If Not PivotFieldExisteDetallado(pt, nombreCampo) Then
        Err.Raise vbObjectError + 5230, "ConfigurarCampoPivotDetallado", _
            "No existe el campo de tabla dinámica a configurar. Campo=" & nombreCampo & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If

    With pt.PivotFields(nombreCampo)
        .Orientation = orientacion
        .Position = posicion
    End With
    Exit Sub

EH:
    Err.Raise Err.Number, "ConfigurarCampoPivotDetallado", _
        "No se pudo configurar campo de tabla dinámica. Campo=" & nombreCampo & _
        " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt) & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Sub FiltrarPivotCampoDetallado(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal valor As String)
    On Error GoTo EH

    If Not PivotFieldExisteDetallado(pt, nombreCampo) Then
        Err.Raise vbObjectError + 5231, "FiltrarPivotCampoDetallado", _
            "No existe el campo de tabla dinámica a filtrar. Campo=" & nombreCampo & _
            " | Valor=" & valor & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If

    With pt.PivotFields(nombreCampo)
        .ClearAllFilters
        .CurrentPage = valor
    End With
    Exit Sub

EH:
    Err.Raise Err.Number, "FiltrarPivotCampoDetallado", _
        "No se pudo filtrar campo de tabla dinámica. Campo=" & nombreCampo & _
        " | Valor=" & valor & _
        " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt) & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Sub OrdenarMesesPivotDetallado(ByVal pt As PivotTable, ByVal campoMesNombre As String)
    On Error Resume Next
    Dim meses As Variant, i As Long
    meses = MesesES()
    For i = LBound(meses) To UBound(meses)
        pt.PivotFields(campoMesNombre).PivotItems(CStr(meses(i))).Position = i + 1
    Next i
    On Error GoTo 0
End Sub

Private Sub AgregarDataFieldDetallado(ByVal pt As PivotTable, ByVal sourceFieldName As String, ByVal caption As String, ByVal funcion As XlConsolidationFunction, ByVal formatoNumero As String)
    On Error GoTo EH

    Dim pf As PivotField

    If Not PivotFieldExisteDetallado(pt, sourceFieldName) Then
        Err.Raise vbObjectError + 5232, "AgregarDataFieldDetallado", _
            "No existe el campo de origen para agregar a valores. Campo=" & sourceFieldName & _
            " | Caption=" & caption & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If

    Set pf = pt.AddDataField(pt.PivotFields(sourceFieldName), caption, funcion)
    FormatearDataField pf, formatoNumero
    Exit Sub

EH:
    Err.Raise Err.Number, "AgregarDataFieldDetallado", _
        "No se pudo agregar campo a valores de tabla dinámica. Campo=" & sourceFieldName & _
        " | Caption=" & caption & _
        " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt) & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Sub AsegurarCampoCalculadoPctVariacionDetalle(ByVal pt As PivotTable)
    On Error GoTo EH

    Dim formula As String

    formula = "=(Ejecutado_Actual-Ejecutado_Anterior_Actualizado)/Ejecutado_Anterior_Actualizado"

    If Not PivotFieldExisteDetallado(pt, "Ejecutado_Actual") Then
        Err.Raise vbObjectError + 5233, "AsegurarCampoCalculadoPctVariacionDetalle", _
            "No existe el campo fuente requerido para PctVariacion: Ejecutado_Actual" & _
            " | Fórmula=" & formula & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If

    If Not PivotFieldExisteDetallado(pt, "Ejecutado_Anterior_Actualizado") Then
        Err.Raise vbObjectError + 5234, "AsegurarCampoCalculadoPctVariacionDetalle", _
            "No existe el campo fuente requerido para PctVariacion: Ejecutado_Anterior_Actualizado" & _
            " | Fórmula=" & formula & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If

    If PivotFieldExisteDetallado(pt, "PctVariacion") Then Exit Sub

    pt.CalculatedFields.Add Name:="PctVariacion", Formula:=formula

    If Not PivotFieldExisteDetallado(pt, "PctVariacion") Then
        Err.Raise vbObjectError + 5235, "AsegurarCampoCalculadoPctVariacionDetalle", _
            "Excel no dejó disponible el campo calculado PctVariacion después de crearlo." & _
            " | Fórmula=" & formula & _
            " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt)
    End If
    Exit Sub

EH:
    Err.Raise Err.Number, "AsegurarCampoCalculadoPctVariacionDetalle", _
        "No se pudo crear o validar el campo calculado PctVariacion." & _
        " | Fórmula=" & formula & _
        " | Campos disponibles=" & CamposDisponiblesPivotDetallado(pt) & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Function PivotFieldExisteDetallado(ByVal pt As PivotTable, ByVal nombreCampo As String) As Boolean
    Dim pf As PivotField

    For Each pf In pt.PivotFields
        If StrComp(pf.Name, nombreCampo, vbBinaryCompare) = 0 Then
            PivotFieldExisteDetallado = True
            Exit Function
        End If
    Next pf
End Function

Private Function CamposDisponiblesPivotDetallado(ByVal pt As PivotTable) As String
    On Error GoTo EH

    Dim pf As PivotField
    Dim res As String

    For Each pf In pt.PivotFields
        If Len(res) > 0 Then res = res & ", "
        res = res & pf.Name
    Next pf

    CamposDisponiblesPivotDetallado = res
    Exit Function

EH:
    CamposDisponiblesPivotDetallado = "(no se pudieron leer los campos disponibles: " & Err.Number & " - " & Err.Description & ")"
End Function

Private Sub FormatearDataField(ByVal pf As PivotField, ByVal formato As String)
    On Error GoTo EH

    pf.NumberFormat = formato
    Exit Sub

EH:
    Err.Raise Err.Number, "FormatearDataField", _
        "No se pudo aplicar formato al campo de valores. Campo=" & pf.Name & _
        " | Formato=" & formato & _
        " | Err.Number=" & Err.Number & _
        " | Err.Description=" & Err.Description
End Sub

Private Function CrearHojaLimpiaDetallada(ByVal wb As Workbook, ByVal nombreHoja As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(nombreHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set CrearHojaLimpiaDetallada = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    CrearHojaLimpiaDetallada.Name = nombreHoja
End Function

Private Sub CrearResumenControlDetallado(ByVal wbOut As Workbook, ByVal resumen As Object, ByVal archivoEjec As String, ByVal archivoAsignados As String, ByVal archivoEjecComparativo As String, ByVal archivoCod As String, ByVal anio As Long, ByVal mesCierre As Long, ByVal rutaFinal As String, ByVal cantidadExcluidas As Long)
    Dim ws As Worksheet
    Set ws = CrearHojaLimpiaDetallada(wbOut, NOMBRE_HOJA_RESUMEN_CONTROL)

    ws.Range("A1").Value = "Resumen control TD detallada GG"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3:B31").Font.Name = "Calibri"
    ws.Range("A3").Value = "Año": ws.Range("B3").Value = anio
    ws.Range("A4").Value = "Mes cierre": ws.Range("B4").Value = mesCierre
    ws.Range("A5").Value = "Archivo ejecuciones actual": ws.Range("B5").Value = archivoEjec
    ws.Range("A6").Value = "Archivo asignados actual": ws.Range("B6").Value = archivoAsignados
    ws.Range("A7").Value = "Archivo ejecuciones comparativo": ws.Range("B7").Value = archivoEjecComparativo
    ws.Range("A8").Value = "Archivo salida": ws.Range("B8").Value = rutaFinal
    ws.Range("A9").Value = "Archivo codiguera": ws.Range("B9").Value = archivoCod
    ws.Range("A11").Value = "Total Ejecutado incluido": ws.Range("B11").Value = CDbl(resumen("total_ejecutado_incluido"))
    ws.Range("A12").Value = "Total Asignado incluido": ws.Range("B12").Value = CDbl(resumen("total_asignado_incluido"))
    ws.Range("A13").Value = "Total Ejecutado actual para comparativo": ws.Range("B13").Value = CDbl(resumen("total_comp_actual"))
    ws.Range("A14").Value = "Total Ejecutado anterior original": ws.Range("B14").Value = CDbl(resumen("total_comp_anterior_original"))
    ws.Range("A15").Value = "Total Ejecutado anterior actualizado": ws.Range("B15").Value = CDbl(resumen("total_comp_anterior_actualizado"))
    ws.Range("A17").Value = "Cantidad líneas ejecución leídas": ws.Range("B17").Value = CLng(resumen("ejec_leidas"))
    ws.Range("A18").Value = "Cantidad líneas ejecución incluidas": ws.Range("B18").Value = CLng(resumen("ejec_incluidas"))
    ws.Range("A19").Value = "Cantidad líneas ejecución no incluidas": ws.Range("B19").Value = CLng(resumen("ejec_no_incluidas"))
    ws.Range("A20").Value = "Cantidad líneas ejecución con llave no encontrada": ws.Range("B20").Value = CLng(resumen("ejec_llave_no_encontrada"))
    ws.Range("A21").Value = "Cantidad líneas asignados leídas": ws.Range("B21").Value = CLng(resumen("asig_leidas"))
    ws.Range("A22").Value = "Cantidad líneas asignados incluidas": ws.Range("B22").Value = CLng(resumen("asig_incluidas"))
    ws.Range("A23").Value = "Cantidad líneas asignados no incluidas": ws.Range("B23").Value = CLng(resumen("asig_no_incluidas"))
    ws.Range("A24").Value = "Cantidad líneas asignados con llave no encontrada": ws.Range("B24").Value = CLng(resumen("asig_llave_no_encontrada"))
    ws.Range("A25").Value = "Cantidad líneas comparativo anterior leídas": ws.Range("B25").Value = CLng(resumen("comp_ant_leidas"))
    ws.Range("A26").Value = "Cantidad líneas comparativo anterior incluidas": ws.Range("B26").Value = CLng(resumen("comp_ant_incluidas"))
    ws.Range("A27").Value = "Cantidad líneas comparativo anterior no incluidas": ws.Range("B27").Value = CLng(resumen("comp_ant_no_incluidas"))
    ws.Range("A28").Value = "Cantidad líneas comparativo anterior con llave no encontrada": ws.Range("B28").Value = CLng(resumen("comp_ant_llave_no_encontrada"))
    ws.Range("A30").Value = "Detalle compacto de líneas excluidas": ws.Range("B30").Value = NOMBRE_HOJA_CONTROL_NO_INCLUIDAS
    ws.Range("A31").Value = "Cantidad líneas excluidas registradas": ws.Range("B31").Value = cantidadExcluidas
    ws.Range("B11:B15").NumberFormat = "#,##0"
    ws.Columns("A:B").AutoFit
End Sub

Private Function GuardarReporteDetalladoGG(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mesNum As Long) As String
    Dim carpetaSalida As String
    Dim fileName As String
    Dim ruta As String
    On Error GoTo EH

    carpetaSalida = RutaReportesGeneradosActiva()
    AsegurarCarpetaExiste carpetaSalida
    fileName = "Informe_GG_Base_Detallada_" & anio & "_" & Format$(mesNum, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsb"
    ruta = CombinarRuta(carpetaSalida, fileName)
    wbOut.SaveAs Filename:=ruta, FileFormat:=xlExcel12
    GuardarReporteDetalladoGG = ruta
    Exit Function
EH:
    Err.Raise Err.Number, "GuardarReporteDetalladoGG", "Error guardando archivo detallado: " & Err.Description & " | Ruta: " & ruta
End Function

Private Function EncabezadosBaseTDDetallada() As Variant
    EncabezadosBaseTDDetallada = Array("Origen", "ArchivoOrigen", "HojaOrigen", "FilaOrigen", "Año", "MesNum", "MesNombre", "FechaValor", "Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Finac", "Der-F", "PG", "Spg", "Proy", "Rubro", "R. Aux", "UE", "Dep", "Obra", "Der. Obra", "Serv", "SNIIP", "Ejecutado", "Asignado", "ImporteMN", "Incluir_en_Informe", "Estado_Codiguera")
End Function

Private Function EncabezadosBaseComparativoDetallada() As Variant
    EncabezadosBaseComparativoDetallada = Array("PeriodoComparativo", "Origen", "ArchivoOrigen", "HojaOrigen", "FilaOrigen", "AñoOrigen", "AñoDestino", "MesNum", "MesNombre", "FechaValor", "Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Finac", "Der-F", "PG", "Spg", "Proy", "Rubro", "R. Aux", "UE", "Dep", "Obra", "Der. Obra", "Serv", "SNIIP", "ImporteOriginalMN", "TipoIndice", "ArchivoIndice", "PeriodoIndiceBase", "ValorIndiceBase", "PeriodoIndiceDestino", "ValorIndiceDestino", "FactorActualizacion", "ImporteActualizadoMN", "Ejecutado_Actual", "Ejecutado_Anterior_Actualizado", "Incluir_en_Informe", "Estado_Codiguera")
End Function

Private Function EncabezadosControlNoIncluidas() As Variant
    EncabezadosControlNoIncluidas = Array("Origen", "ArchivoOrigen", "HojaOrigen", "FilaOrigen", "Año", "MesNum", "MesNombre", "Clave Llave presupuestal", "MotivoExclusion", "Estado_Codiguera", "Incluir_en_Informe", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "ImporteMN", "Ejecutado", "Asignado")
End Function

Private Sub EscribirEncabezados(ByVal ws As Worksheet, ByVal headers As Variant)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).Value = headers
    ws.Rows(1).Font.Bold = True
End Sub

Private Function EncabezadosPrefijados(ByRef arr As Variant, ByVal prefijo As String) As Variant
    Dim res() As Variant
    Dim j As Long
    ReDim res(0 To UBound(arr, 2) - 1)
    For j = 1 To UBound(arr, 2)
        res(j - 1) = NombreEncabezadoUnico(prefijo & TextoSeguro(arr(1, j)), res, j - 1)
    Next j
    EncabezadosPrefijados = res
End Function

Private Function NombreEncabezadoUnico(ByVal candidato As String, ByRef existentes As Variant, ByVal cantidadExistentes As Long) As String
    Dim base As String, actual As String, suf As Long, i As Long, existe As Boolean
    base = LimpiarTexto(candidato)
    If Len(base) = 0 Then base = "Columna"
    actual = base
    suf = 1
    Do
        existe = False
        For i = 0 To cantidadExistentes - 1
            If StrComp(CStr(existentes(i)), actual, vbTextCompare) = 0 Then existe = True: Exit For
        Next i
        If Not existe Then Exit Do
        suf = suf + 1
        actual = base & "_" & CStr(suf)
    Loop
    NombreEncabezadoUnico = actual
End Function

Private Function UnirEncabezados(ByVal baseHeaders As Variant, ByVal pref As Variant) As Variant
    Dim total As Long, res() As Variant, i As Long, p As Long
    total = UBound(baseHeaders) + 1 + UBound(pref) + 1
    ReDim res(0 To total - 1)
    p = 0
    For i = LBound(baseHeaders) To UBound(baseHeaders): res(p) = baseHeaders(i): p = p + 1: Next i
    For i = LBound(pref) To UBound(pref): res(p) = pref(i): p = p + 1: Next i
    UnirEncabezados = res
End Function

Private Function ObtenerColumnaDetallada(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long, k As String
    For i = LBound(aliases) To UBound(aliases)
        k = LimpiarTexto(LCase$(CStr(aliases(i))))
        If headers.Exists(k) Then ObtenerColumnaDetallada = CLng(headers(k)): Exit Function
    Next i
End Function

Private Function ValorMatriz(ByRef arr As Variant, ByVal fila As Long, ByVal col As Long) As Variant
    If col <= 0 Then
        ValorMatriz = Empty
    Else
        ValorMatriz = arr(fila, col)
    End If
End Function

Private Function NormalizarIncluirDetallado(ByVal valor As Variant) As String
    Dim s As String
    s = Replace(UCase$(LimpiarTexto(TextoSeguro(valor))), " ", "")
    If s = "SI" Or s = "SÍ" Then
        NormalizarIncluirDetallado = "SI"
    Else
        NormalizarIncluirDetallado = "NO"
    End If
End Function

Private Function NormalizarClaveTextoDetallada(ByVal valor As Variant) As String
    Dim s As String
    s = TextoSeguro(valor)
    s = Replace(s, "|", "-")
    s = Replace(s, " ", "")
    NormalizarClaveTextoDetallada = s
End Function

Private Function NombreArchivoDesdeRutaDetallada(ByVal rutaArchivo As String) As String
    Dim p As Long
    p = InStrRev(rutaArchivo, "\")
    If p > 0 Then
        NombreArchivoDesdeRutaDetallada = Mid$(rutaArchivo, p + 1)
    Else
        NombreArchivoDesdeRutaDetallada = rutaArchivo
    End If
End Function

Private Sub ValidarFilaSalida(ByVal fila As Long, ByVal nombreHoja As String)
    If fila > Rows.Count Then
        Err.Raise vbObjectError + 5400, "ValidarFilaSalida", "La hoja " & nombreHoja & " supera el límite de Excel (" & Format$(Rows.Count, "#,##0") & " filas)."
    End If
End Sub

Private Sub ValidarLimiteFilasHoja(ByVal ws As Worksheet, ByVal nombreHoja As String)
    If UltimaFilaConDatos(ws) > Rows.Count Then
        Err.Raise vbObjectError + 5401, "ValidarLimiteFilasHoja", "La hoja " & nombreHoja & " supera el límite de Excel."
    End If
End Sub

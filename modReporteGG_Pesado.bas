Option Explicit

Private Const NOMBRE_HOJA_BASE_PESADA As String = "Base_Detallada"
Private Const NOMBRE_TABLA_BASE_PESADA As String = "tblBasePesadaGG"
Private Const NOMBRE_HOJA_TD_EJEC As String = "TD_Ejecucion_Detalle"
Private Const NOMBRE_HOJA_TD_ASIG As String = "TD_Asignado_Detalle"
Private Const NOMBRE_HOJA_TD_COMBINADA As String = "TD_Ejec_Asig_Detalle"
Private Const NOMBRE_HOJA_RESUMEN_PESADO As String = "Resumen_Control"
Private Const ORIGEN_EJECUCION As String = "Ejecución"
Private Const ORIGEN_ASIGNADO As String = "Asignado"
Private Const MES_ASIGNADO_ANUAL As String = "Asignado anual"

Public Sub Generar_TD_Pesada_GG_Desde_Panel()
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
    Dim carpetaEjecActual As String
    Dim carpetaAsignadosActual As String
    Dim wbE As Workbook
    Dim wbA As Workbook
    Dim wbC As Workbook
    Dim wbOut As Workbook
    Dim wsE As Worksheet
    Dim wsA As Worksheet
    Dim wsC As Worksheet
    Dim wsBase As Worksheet
    Dim dictCodDetalle As Object
    Dim resumen As Object
    Dim totalFilasDetalle As Long
    Dim rutaFinal As String
    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errLine As Long
    Dim msg As String

    procedimiento = "Generar_TD_Pesada_GG_Desde_Panel"
    Debug.Print "Inicio Generar_TD_Pesada_GG_Desde_Panel: " & Now

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

    etapaActual = "validando mes de cierre"
    mesTxt = CStr(wsPanel.Range("B4").Value)
    mesCierre = MesTextoANumero(mesTxt)
    If mesCierre < 1 Or mesCierre > 12 Then
        Err.Raise vbObjectError + 5102, procedimiento, "Mes inválido en Panel Reportes!B4. Use: Enero..Diciembre (Setiembre). Valor: '" & mesTxt & "'."
    End If

    etapaActual = "inicializando diccionarios de trabajo"
    Set dictCodDetalle = CreateObject("Scripting.Dictionary")
    Set resumen = CrearResumenPesado()

    etapaActual = "resolviendo archivos de entrada"
    carpetaEjecActual = RutaCarpetaEjecucionesAnioActiva(anio)
    If Len(Dir(carpetaEjecActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 5103, procedimiento, "No existe la carpeta de ejecuciones del año " & anio & ": " & carpetaEjecActual
    End If
    archivoEjec = ObtenerArchivoMasReciente(carpetaEjecActual)
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 5104, procedimiento, "No se encontró archivo de ejecuciones del año " & anio & " en: " & carpetaEjecActual
    End If

    etapaActual = "resolviendo archivos de entrada"
    carpetaAsignadosActual = RutaCarpetaAsignadosGastosAnioActiva(anio)
    If Len(Dir(carpetaAsignadosActual, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 5105, procedimiento, "No existe la carpeta de asignados del año " & anio & ": " & carpetaAsignadosActual
    End If
    archivoAsignados = ObtenerArchivoMasRecientePorFechaCreacion(carpetaAsignadosActual)
    If Len(archivoAsignados) = 0 Then
        Err.Raise vbObjectError + 5106, procedimiento, "No se encontró archivo de asignados del año " & anio & " en: " & carpetaAsignadosActual
    End If

    etapaActual = "resolviendo archivos de entrada"
    archivoCod = ResolverArchivoCodiguera(RutaCodigueraActiva())
    If Len(archivoCod) = 0 Then
        Err.Raise vbObjectError + 5107, procedimiento, "No se encontró archivo de codiguera en: " & RutaCodigueraActiva()
    End If

    etapaActual = "abriendo archivos fuente"
    Set wbE = Workbooks.Open(Filename:=archivoEjec, ReadOnly:=True, UpdateLinks:=False)
    Set wbA = Workbooks.Open(Filename:=archivoAsignados, ReadOnly:=True, UpdateLinks:=False)
    Set wbC = Workbooks.Open(Filename:=archivoCod, ReadOnly:=True, UpdateLinks:=False)

    etapaActual = "obteniendo hojas fuente"
    Set wsE = ObtenerHojaEjecuciones(wbE)
    Set wsA = ObtenerHojaAsignados(wbA)
    Set wsC = ObtenerHojaCodiguera(wbC)

    etapaActual = "cargando codiguera detallada"
    LeerCodigueraDetallePesada wsC, dictCodDetalle

    etapaActual = "contando filas detalladas"
    totalFilasDetalle = ContarFilasEjecucionPesada(wsE, anio, mesCierre) + ContarFilasAsignadoPesada(wsA, anio)
    If totalFilasDetalle + 1 > Rows.Count Then
        Err.Raise vbObjectError + 5108, procedimiento, "La base detallada tendría " & Format$(totalFilasDetalle, "#,##0") & " filas de datos y supera el límite de Excel (" & Format$(Rows.Count - 1, "#,##0") & "). Se requiere otra estrategia (por ejemplo Power Pivot, Power Query/Data Model o dividir la base)."
    End If

    etapaActual = "creando workbook pesado"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBase = wbOut.Worksheets(1)
    wsBase.Name = NOMBRE_HOJA_BASE_PESADA

    etapaActual = "construyendo base detallada pesada"
    ConstruirBaseDetalladaPesada wsBase, wsE, wsA, dictCodDetalle, anio, mesCierre, archivoEjec, archivoAsignados, resumen

    etapaActual = "creando tabla tblBasePesadaGG"
    CrearTablaBasePesada wsBase

    etapaActual = "creando tablas dinámicas pesadas"
    CrearTablasDinamicasPesadas wbOut, wsBase

    etapaActual = "creando Resumen_Control"
    CrearResumenControlPesado wbOut, resumen, archivoEjec, archivoAsignados, archivoCod, anio, mesCierre, rutaFinal

    etapaActual = "guardando archivo pesado"
    rutaFinal = GuardarReportePesadoGG(wbOut, anio, mesCierre)
    wbOut.Worksheets(NOMBRE_HOJA_RESUMEN_PESADO).Range("B8").Value = rutaFinal
    wbOut.Save

    etapaActual = "cerrando archivos auxiliares"
    CerrarWorkbookSeguro wbE, False
    CerrarWorkbookSeguro wbA, False
    CerrarWorkbookSeguro wbC, False

    MsgBox "Archivo pesado generado correctamente:" & vbCrLf & rutaFinal, vbInformation
    Debug.Print "Fin Generar_TD_Pesada_GG_Desde_Panel: " & Now & " | " & rutaFinal
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
    On Error GoTo 0

    msg = "Error al generar TD pesada GG." & vbCrLf & vbCrLf
    msg = msg & "Procedimiento: " & procedimiento & vbCrLf
    msg = msg & "Etapa: " & etapaActual & vbCrLf
    msg = msg & "Err.Number: " & errNum & vbCrLf
    msg = msg & "Err.Description: " & errDesc & vbCrLf
    msg = msg & "Err.Source: " & errSource & vbCrLf
    msg = msg & "Erl: " & errLine & vbCrLf
    msg = msg & "Año: " & anio & vbCrLf
    msg = msg & "Mes cierre: " & mesCierre & vbCrLf
    msg = msg & "Archivo ejecuciones: " & IIf(Len(archivoEjec) > 0, archivoEjec, "(no detectado)") & vbCrLf
    msg = msg & "Archivo asignados: " & IIf(Len(archivoAsignados) > 0, archivoAsignados, "(no detectado)") & vbCrLf
    msg = msg & "Archivo codiguera: " & IIf(Len(archivoCod) > 0, archivoCod, "(no detectado)") & vbCrLf
    msg = msg & "Carpeta ejecuciones: " & IIf(Len(carpetaEjecActual) > 0, carpetaEjecActual, "(no determinada)") & vbCrLf
    msg = msg & "Carpeta asignados: " & IIf(Len(carpetaAsignadosActual) > 0, carpetaAsignadosActual, "(no determinada)") & vbCrLf
    msg = msg & "Salida: " & IIf(Len(rutaFinal) > 0, rutaFinal, RutaReportesGeneradosActiva()) & vbCrLf
    msg = msg & vbCrLf & DiagnosticoRutasActivas() & vbCrLf
    msg = msg & vbCrLf & "Workbooks abiertos:" & vbCrLf & DiagnosticoWorkbooksAbiertos()

    Debug.Print String(100, "-")
    Debug.Print msg
    Debug.Print String(100, "-")
    MsgBox msg, vbCritical
End Sub

Private Function CrearResumenPesado() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d("total_ejecutado_informe") = 0#
    d("total_asignado_informe") = 0#
    d("lineas_ejecucion") = 0&
    d("lineas_asignados") = 0&
    d("lineas_llave_no_encontrada") = 0&
    d("lineas_no_incluidas") = 0&

    Set CrearResumenPesado = d
End Function

Private Sub LeerCodigueraDetallePesada(ByVal ws As Worksheet, ByRef dictCodDetalle As Object)
    On Error GoTo EH

    Dim arr As Variant
    Dim headers As Object
    Dim i As Long
    Dim clave As String
    Dim info As Variant
    Dim colFinac As Long, colDerF As Long, colPg As Long, colSpg As Long, colProy As Long
    Dim colRubro As Long, colRAux As Long, colUe As Long, colDep As Long, colObra As Long, colDerObra As Long
    Dim colServ As Long, colSniip As Long, colTitular As Long, colN1 As Long, colN2 As Long, colN3 As Long
    Dim colIncluir As Long, colEstado As Long, colClaveTexto As Long

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    colFinac = ObtenerColumnaPesada(headers, Array("finac", "finac código numérico"))
    colDerF = ObtenerColumnaPesada(headers, Array("der-f", "der-f código numérico"))
    colPg = ObtenerColumnaPesada(headers, Array("pg", "pg código numérico"))
    colSpg = ObtenerColumnaPesada(headers, Array("spg", "spg código numérico"))
    colProy = ObtenerColumnaPesada(headers, Array("proy", "proyecto", "proyecto código numérico", "proy código numérico"))
    colRubro = ObtenerColumnaPesada(headers, Array("rubro", "rubro código numérico"))
    colRAux = ObtenerColumnaPesada(headers, Array("r. aux", "r. aux código numérico"))
    colUe = ObtenerColumnaPesada(headers, Array("ue", "ue código numérico"))
    colDep = ObtenerColumnaPesada(headers, Array("dep", "dep código numérico"))
    colObra = ObtenerColumnaPesada(headers, Array("obra", "obra código numérico"))
    colDerObra = ObtenerColumnaPesada(headers, Array("der. obra", "der. obra código numérico"))
    colServ = ObtenerColumnaPesada(headers, Array("serv", "serv código numérico"))
    colSniip = ObtenerColumnaPesada(headers, Array("sniip", "snip", "sniip código numérico", "snip código numérico"))
    colTitular = ObtenerColumna(headers, Array("titular"))
    colN1 = ObtenerColumna(headers, Array("nivel_1"))
    colN2 = ObtenerColumna(headers, Array("nivel_2"))
    colN3 = ObtenerColumna(headers, Array("nivel_3"))
    colIncluir = ObtenerColumna(headers, Array("incluir_en_informe"))
    colEstado = ObtenerColumnaPesada(headers, Array("estado_codiguera", "estado"))
    colClaveTexto = ObtenerColumnaPesada(headers, Array("clave llave presupuestal"))

    For i = 2 To UBound(arr, 1)
        If colFinac > 0 And colDerF > 0 And colPg > 0 And colSpg > 0 And colRubro > 0 And colRAux > 0 And colUe > 0 And colDep > 0 And colObra > 0 And colDerObra > 0 And colServ > 0 And colSniip > 0 Then
            clave = ConstruirClaveLlavePresupuestalCodiguera(ValorMatriz(arr, i, colFinac), ValorMatriz(arr, i, colDerF), ValorMatriz(arr, i, colPg), ValorMatriz(arr, i, colSpg), ValorMatriz(arr, i, colProy), ValorMatriz(arr, i, colRubro), ValorMatriz(arr, i, colRAux), ValorMatriz(arr, i, colUe), ValorMatriz(arr, i, colDep), ValorMatriz(arr, i, colObra), ValorMatriz(arr, i, colDerObra), ValorMatriz(arr, i, colServ), ValorMatriz(arr, i, colSniip))
        ElseIf colClaveTexto > 0 Then
            clave = NormalizarClaveTextoPesada(arr(i, colClaveTexto))
        Else
            clave = ""
        End If

        If Len(clave) > 0 Then
            info = Array( _
                TextoSeguro(ValorMatriz(arr, i, colTitular)), _
                TextoSeguro(ValorMatriz(arr, i, colN1)), _
                TextoSeguro(ValorMatriz(arr, i, colN2)), _
                TextoSeguro(ValorMatriz(arr, i, colN3)), _
                NormalizarIncluirPesado(ValorMatriz(arr, i, colIncluir)), _
                TextoSeguro(ValorMatriz(arr, i, colEstado)) _
            )
            dictCodDetalle(clave) = info
        End If
    Next i
    Exit Sub
EH:
    Err.Raise Err.Number, "LeerCodigueraDetallePesada", "Error leyendo codiguera completa: " & Err.Description
End Sub

Private Function ContarFilasEjecucionPesada(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long) As Long
    Dim arr As Variant, headers As Object, i As Long, fechaValor As Date, colFecha As Long
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    colFecha = ObtenerColumna(headers, Array("fecha valor"))
    For i = 2 To UBound(arr, 1)
        If TryObtenerFechaValorSeguro(arr(i, colFecha), fechaValor) Then
            If Year(fechaValor) = anio And Month(fechaValor) <= mesCierre Then ContarFilasEjecucionPesada = ContarFilasEjecucionPesada + 1
        End If
    Next i
End Function

Private Function ContarFilasAsignadoPesada(ByVal ws As Worksheet, ByVal anio As Long) As Long
    Dim arr As Variant, headers As Object, i As Long, colAnio As Long, anioFila As Long
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    colAnio = ObtenerColumnaPesada(headers, Array("año", "anio", "ejercicio", "ej"))
    For i = 2 To UBound(arr, 1)
        If colAnio > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                If anioFila = anio Then ContarFilasAsignadoPesada = ContarFilasAsignadoPesada + 1
            Else
                ContarFilasAsignadoPesada = ContarFilasAsignadoPesada + 1
            End If
        Else
            ContarFilasAsignadoPesada = ContarFilasAsignadoPesada + 1
        End If
    Next i
End Function

Private Sub ConstruirBaseDetalladaPesada(ByVal wsBase As Worksheet, ByVal wsE As Worksheet, ByVal wsA As Worksheet, ByVal dictCodDetalle As Object, ByVal anio As Long, ByVal mesCierre As Long, ByVal archivoEjec As String, ByVal archivoAsignados As String, ByRef resumen As Object)
    On Error GoTo EH

    Dim arrE As Variant, arrA As Variant
    Dim headersE As Object, headersA As Object
    Dim prefE As Variant, prefA As Variant
    Dim baseHeaders As Variant, allHeaders As Variant
    Dim filaOut As Long
    Dim colOffsetE As Long, colOffsetA As Long

    arrE = wsE.Range(wsE.Cells(1, 1), wsE.Cells(UltimaFilaConDatos(wsE), UltimaColConDatos(wsE))).Value2
    arrA = wsA.Range(wsA.Cells(1, 1), wsA.Cells(UltimaFilaConDatos(wsA), UltimaColConDatos(wsA))).Value2
    Set headersE = MapearEncabezados(arrE)
    Set headersA = MapearEncabezados(arrA)

    prefE = EncabezadosPrefijados(arrE, "Ejec_")
    prefA = EncabezadosPrefijados(arrA, "Asig_")
    baseHeaders = EncabezadosBasePesada()
    allHeaders = UnirEncabezados(baseHeaders, prefE, prefA)

    wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(1, UBound(allHeaders) + 1)).Value = allHeaders
    colOffsetE = UBound(baseHeaders) + 2
    colOffsetA = colOffsetE + UBound(prefE) + 1

    filaOut = 2
    AgregarFilasEjecucionPesada wsBase, filaOut, arrE, headersE, dictCodDetalle, anio, mesCierre, archivoEjec, wsE.Name, colOffsetE, resumen
    AgregarFilasAsignadoPesada wsBase, filaOut, arrA, headersA, dictCodDetalle, anio, archivoAsignados, wsA.Name, colOffsetA, resumen

    wsBase.Columns.AutoFit
    Exit Sub
EH:
    Err.Raise Err.Number, "ConstruirBaseDetalladaPesada", "Error construyendo Base_Detallada: " & Err.Description
End Sub

Private Sub AgregarFilasEjecucionPesada(ByVal wsBase As Worksheet, ByRef filaOut As Long, ByRef arr As Variant, ByVal headers As Object, ByVal dictCodDetalle As Object, ByVal anio As Long, ByVal mesCierre As Long, ByVal archivoOrigen As String, ByVal hojaOrigen As String, ByVal colOffsetOriginal As Long, ByRef resumen As Object)
    Dim i As Long, j As Long, fechaValor As Date, importeMN As Double, clave As String, info As Variant, incluir As String
    Dim colFecha As Long, colFinac As Long, colDerF As Long, colPg As Long, colSpg As Long, colProy As Long
    Dim colRubro As Long, colRAux As Long, colUe As Long, colDep As Long, colObra As Long, colDerObra As Long, colServ As Long, colSniip As Long, colImporte As Long
    Dim baseVals As Variant, meses As Variant

    meses = MesesES()
    colFecha = ObtenerColumna(headers, Array("fecha valor"))
    colFinac = ObtenerColumna(headers, Array("finac código numérico"))
    colDerF = ObtenerColumna(headers, Array("der-f código numérico"))
    colPg = ObtenerColumna(headers, Array("pg código numérico"))
    colSpg = ObtenerColumna(headers, Array("spg código numérico"))
    colProy = ObtenerColumna(headers, Array("proyecto", "proy", "proyecto código numérico", "proy código numérico"))
    colRubro = ObtenerColumna(headers, Array("rubro código numérico"))
    colRAux = ObtenerColumna(headers, Array("r. aux código numérico"))
    colUe = ObtenerColumna(headers, Array("ue código numérico"))
    colDep = ObtenerColumna(headers, Array("dep código numérico"))
    colObra = ObtenerColumna(headers, Array("obra código numérico"))
    colDerObra = ObtenerColumna(headers, Array("der. obra código numérico"))
    colServ = ObtenerColumna(headers, Array("serv código numérico"))
    colSniip = ObtenerColumna(headers, Array("snip código numérico", "sniip código numérico", "snip", "sniip"))
    colImporte = ObtenerColumna(headers, Array("importe moneda nacional"))

    For i = 2 To UBound(arr, 1)
        If Not TryObtenerFechaValorSeguro(arr(i, colFecha), fechaValor) Then GoTo SiguienteFila
        If Year(fechaValor) <> anio Or Month(fechaValor) > mesCierre Then GoTo SiguienteFila
        If Not TryDoubleSeguro(arr(i, colImporte), importeMN) Then importeMN = 0#

        clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, colFinac), arr(i, colDerF), arr(i, colPg), arr(i, colSpg), arr(i, colProy), arr(i, colRubro), arr(i, colRAux), arr(i, colUe), arr(i, colDep), arr(i, colObra), arr(i, colDerObra), arr(i, colServ), arr(i, colSniip))
        info = InfoCodigueraParaClave(dictCodDetalle, clave)
        incluir = CStr(info(4))

        baseVals = Array(ORIGEN_EJECUCION, NombreArchivoDesdeRutaPesada(archivoOrigen), hojaOrigen, i, anio, Month(fechaValor), meses(Month(fechaValor) - 1), fechaValor, clave, info(0), info(1), info(2), info(3), arr(i, colFinac), arr(i, colDerF), arr(i, colPg), arr(i, colSpg), arr(i, colProy), arr(i, colRubro), arr(i, colRAux), arr(i, colUe), arr(i, colDep), arr(i, colObra), arr(i, colDerObra), arr(i, colServ), arr(i, colSniip), importeMN, 0#, importeMN, incluir, info(5))
        wsBase.Range(wsBase.Cells(filaOut, 1), wsBase.Cells(filaOut, UBound(baseVals) + 1)).Value = baseVals
        For j = 1 To UBound(arr, 2)
            wsBase.Cells(filaOut, colOffsetOriginal + j - 1).Value = arr(i, j)
        Next j

        resumen("lineas_ejecucion") = CLng(resumen("lineas_ejecucion")) + 1
        ActualizarResumenCodPesado resumen, incluir, clave, dictCodDetalle, importeMN, 0#
        filaOut = filaOut + 1
SiguienteFila:
    Next i
End Sub

Private Sub AgregarFilasAsignadoPesada(ByVal wsBase As Worksheet, ByRef filaOut As Long, ByRef arr As Variant, ByVal headers As Object, ByVal dictCodDetalle As Object, ByVal anio As Long, ByVal archivoOrigen As String, ByVal hojaOrigen As String, ByVal colOffsetOriginal As Long, ByRef resumen As Object)
    Dim i As Long, j As Long, monto As Double, clave As String, info As Variant, incluir As String
    Dim colAnio As Long, anioFila As Long, usarFila As Boolean
    Dim colFinac As Long, colDerF As Long, colPg As Long, colSpg As Long, colProy As Long
    Dim colRubro As Long, colRAux As Long, colUe As Long, colDep As Long, colObra As Long, colDerObra As Long, colServ As Long, colSniip As Long, colAsignado As Long
    Dim baseVals As Variant

    colAnio = ObtenerColumnaPesada(headers, Array("año", "anio", "ejercicio", "ej"))
    colFinac = ObtenerColumna(headers, Array("finac"))
    colDerF = ObtenerColumna(headers, Array("der-f"))
    colPg = ObtenerColumna(headers, Array("pg"))
    colSpg = ObtenerColumna(headers, Array("spg"))
    colProy = ObtenerColumna(headers, Array("proy"))
    colRubro = ObtenerColumna(headers, Array("rubro"))
    colRAux = ObtenerColumna(headers, Array("r. aux"))
    colUe = ObtenerColumna(headers, Array("ue"))
    colDep = ObtenerColumna(headers, Array("dep"))
    colObra = ObtenerColumna(headers, Array("obra"))
    colDerObra = ObtenerColumna(headers, Array("der. obra"))
    colServ = ObtenerColumna(headers, Array("serv"))
    colSniip = ObtenerColumna(headers, Array("sniip"))
    colAsignado = ObtenerColumna(headers, Array("asignado"))

    For i = 2 To UBound(arr, 1)
        usarFila = True
        If colAnio > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                usarFila = (anioFila = anio)
            End If
        End If
        If Not usarFila Then GoTo SiguienteFila
        If Not TryDoubleSeguro(arr(i, colAsignado), monto) Then monto = 0#

        clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, colFinac), arr(i, colDerF), arr(i, colPg), arr(i, colSpg), arr(i, colProy), arr(i, colRubro), arr(i, colRAux), arr(i, colUe), arr(i, colDep), arr(i, colObra), arr(i, colDerObra), arr(i, colServ), arr(i, colSniip))
        info = InfoCodigueraParaClave(dictCodDetalle, clave)
        incluir = CStr(info(4))

        baseVals = Array(ORIGEN_ASIGNADO, NombreArchivoDesdeRutaPesada(archivoOrigen), hojaOrigen, i, anio, 0, MES_ASIGNADO_ANUAL, "", clave, info(0), info(1), info(2), info(3), arr(i, colFinac), arr(i, colDerF), arr(i, colPg), arr(i, colSpg), arr(i, colProy), arr(i, colRubro), arr(i, colRAux), arr(i, colUe), arr(i, colDep), arr(i, colObra), arr(i, colDerObra), arr(i, colServ), arr(i, colSniip), 0#, monto, monto, incluir, info(5))
        wsBase.Range(wsBase.Cells(filaOut, 1), wsBase.Cells(filaOut, UBound(baseVals) + 1)).Value = baseVals
        For j = 1 To UBound(arr, 2)
            wsBase.Cells(filaOut, colOffsetOriginal + j - 1).Value = arr(i, j)
        Next j

        resumen("lineas_asignados") = CLng(resumen("lineas_asignados")) + 1
        ActualizarResumenCodPesado resumen, incluir, clave, dictCodDetalle, 0#, monto
        filaOut = filaOut + 1
SiguienteFila:
    Next i
End Sub

Private Sub ActualizarResumenCodPesado(ByRef resumen As Object, ByVal incluir As String, ByVal clave As String, ByVal dictCodDetalle As Object, ByVal ejecutado As Double, ByVal asignado As Double)
    If Not dictCodDetalle.Exists(clave) Then
        resumen("lineas_llave_no_encontrada") = CLng(resumen("lineas_llave_no_encontrada")) + 1
    ElseIf NormalizarIncluirPesado(incluir) <> "SI" Then
        resumen("lineas_no_incluidas") = CLng(resumen("lineas_no_incluidas")) + 1
    End If

    If NormalizarIncluirPesado(incluir) = "SI" Then
        resumen("total_ejecutado_informe") = CDbl(resumen("total_ejecutado_informe")) + ejecutado
        resumen("total_asignado_informe") = CDbl(resumen("total_asignado_informe")) + asignado
    End If
End Sub

Private Function InfoCodigueraParaClave(ByVal dictCodDetalle As Object, ByVal clave As String) As Variant
    If dictCodDetalle.Exists(clave) Then
        InfoCodigueraParaClave = dictCodDetalle(clave)
    Else
        InfoCodigueraParaClave = Array("SIN CODIGUERA", "SIN CODIGUERA", "SIN CODIGUERA", "SIN CODIGUERA", "NO", "LLAVE NO ENCONTRADA")
    End If
End Function

Private Sub CrearTablaBasePesada(ByVal wsBase As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim lo As ListObject

    lastRow = UltimaFilaConDatos(wsBase)
    lastCol = UltimaColConDatos(wsBase)
    If lastRow < 2 Then Err.Raise vbObjectError + 5200, "CrearTablaBasePesada", "Base_Detallada no tiene filas de datos."

    Set lo = wsBase.ListObjects.Add(xlSrcRange, wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(lastRow, lastCol)), , xlYes)
    lo.Name = NOMBRE_TABLA_BASE_PESADA
    lo.TableStyle = "TableStyleMedium2"
End Sub

Private Sub CrearTablasDinamicasPesadas(ByVal wbOut As Workbook, ByVal wsBase As Worksheet)
    Dim lo As ListObject
    Dim pc As PivotCache

    Set lo = wsBase.ListObjects(NOMBRE_TABLA_BASE_PESADA)
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 5201, "CrearTablasDinamicasPesadas", "La tabla " & NOMBRE_TABLA_BASE_PESADA & " existe, pero no tiene filas de datos para crear la PivotCache."
    End If
    Set pc = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=NOMBRE_TABLA_BASE_PESADA)

    CrearPivotEjecucionPesada wbOut, pc
    CrearPivotAsignadoPesada wbOut, pc
    CrearPivotCombinadaPesada wbOut, pc
End Sub

Private Sub CrearPivotEjecucionPesada(ByVal wbOut As Workbook, ByVal pc As PivotCache)
    Dim ws As Worksheet, pt As PivotTable
    Set ws = CrearHojaLimpiaPesada(wbOut, NOMBRE_HOJA_TD_EJEC)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptEjecucionDetalleGG")
    ConfigurarPivotBasePesada pt, ORIGEN_EJECUCION
    ConfigurarCampoPivotPesado pt, "MesNombre", xlColumnField, 1
    OrdenarMesesPivotPesado pt, "MesNombre"
    pt.AddDataField pt.PivotFields("Ejecutado"), "Suma de Ejecutado", xlSum
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Sub CrearPivotAsignadoPesada(ByVal wbOut As Workbook, ByVal pc As PivotCache)
    Dim ws As Worksheet, pt As PivotTable
    Set ws = CrearHojaLimpiaPesada(wbOut, NOMBRE_HOJA_TD_ASIG)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptAsignadoDetalleGG")
    ConfigurarPivotBasePesada pt, ORIGEN_ASIGNADO
    pt.AddDataField pt.PivotFields("Asignado"), "Suma de Asignado", xlSum
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Sub CrearPivotCombinadaPesada(ByVal wbOut As Workbook, ByVal pc As PivotCache)
    Dim ws As Worksheet, pt As PivotTable
    Set ws = CrearHojaLimpiaPesada(wbOut, NOMBRE_HOJA_TD_COMBINADA)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptEjecAsigDetalleGG")
    ConfigurarCampoPivotPesado pt, "Incluir_en_Informe", xlPageField, 1
    FiltrarPivotCampoPesado pt, "Incluir_en_Informe", "SI"
    ConfigurarCampoPivotPesado pt, "Financiamiento", xlRowField, 1
    ConfigurarCampoPivotPesado pt, "Nivel_1", xlRowField, 2
    ConfigurarCampoPivotPesado pt, "Nivel_2", xlRowField, 3
    ConfigurarCampoPivotPesado pt, "Nivel_3", xlRowField, 4
    pt.AddDataField pt.PivotFields("Ejecutado"), "Suma de Ejecutado", xlSum
    pt.AddDataField pt.PivotFields("Asignado"), "Suma de Asignado", xlSum
    pt.EnableDrilldown = True
    ws.Columns.AutoFit
End Sub

Private Sub ConfigurarPivotBasePesada(ByVal pt As PivotTable, ByVal origen As String)
    ConfigurarCampoPivotPesado pt, "Origen", xlPageField, 1
    ConfigurarCampoPivotPesado pt, "Incluir_en_Informe", xlPageField, 2
    FiltrarPivotCampoPesado pt, "Origen", origen
    FiltrarPivotCampoPesado pt, "Incluir_en_Informe", "SI"
    ConfigurarCampoPivotPesado pt, "Financiamiento", xlRowField, 1
    ConfigurarCampoPivotPesado pt, "Nivel_1", xlRowField, 2
    ConfigurarCampoPivotPesado pt, "Nivel_2", xlRowField, 3
    ConfigurarCampoPivotPesado pt, "Nivel_3", xlRowField, 4
End Sub

Private Sub ConfigurarCampoPivotPesado(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal orientacion As XlPivotFieldOrientation, ByVal posicion As Long)
    With pt.PivotFields(nombreCampo)
        .Orientation = orientacion
        .Position = posicion
    End With
End Sub

Private Sub FiltrarPivotCampoPesado(ByVal pt As PivotTable, ByVal nombreCampo As String, ByVal valor As String)
    On Error Resume Next
    With pt.PivotFields(nombreCampo)
        .ClearAllFilters
        .CurrentPage = valor
    End With
    On Error GoTo 0
End Sub

Private Sub OrdenarMesesPivotPesado(ByVal pt As PivotTable, ByVal campoMesNombre As String)
    On Error Resume Next
    Dim meses As Variant, i As Long
    meses = MesesES()
    For i = LBound(meses) To UBound(meses)
        pt.PivotFields(campoMesNombre).PivotItems(CStr(meses(i))).Position = i + 1
    Next i
    On Error GoTo 0
End Sub

Private Function CrearHojaLimpiaPesada(ByVal wb As Workbook, ByVal nombreHoja As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(nombreHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set CrearHojaLimpiaPesada = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    CrearHojaLimpiaPesada.Name = nombreHoja
End Function

Private Sub CrearResumenControlPesado(ByVal wbOut As Workbook, ByVal resumen As Object, ByVal archivoEjec As String, ByVal archivoAsignados As String, ByVal archivoCod As String, ByVal anio As Long, ByVal mesCierre As Long, ByVal rutaFinal As String)
    Dim ws As Worksheet
    Set ws = CrearHojaLimpiaPesada(wbOut, NOMBRE_HOJA_RESUMEN_PESADO)

    ws.Range("A1").Value = "Resumen control TD pesada GG"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3:B14").Font.Name = "Calibri"
    ws.Range("A3").Value = "Año": ws.Range("B3").Value = anio
    ws.Range("A4").Value = "Mes cierre": ws.Range("B4").Value = mesCierre
    ws.Range("A5").Value = "Archivo ejecuciones": ws.Range("B5").Value = archivoEjec
    ws.Range("A6").Value = "Archivo asignados": ws.Range("B6").Value = archivoAsignados
    ws.Range("A7").Value = "Archivo codiguera": ws.Range("B7").Value = archivoCod
    ws.Range("A8").Value = "Archivo salida": ws.Range("B8").Value = rutaFinal
    ws.Range("A10").Value = "Total Ejecutado (Incluir_en_Informe = SI)": ws.Range("B10").Value = CDbl(resumen("total_ejecutado_informe"))
    ws.Range("A11").Value = "Total Asignado (Incluir_en_Informe = SI)": ws.Range("B11").Value = CDbl(resumen("total_asignado_informe"))
    ws.Range("A12").Value = "Cantidad líneas ejecución": ws.Range("B12").Value = CLng(resumen("lineas_ejecucion"))
    ws.Range("A13").Value = "Cantidad líneas asignados": ws.Range("B13").Value = CLng(resumen("lineas_asignados"))
    ws.Range("A14").Value = "Cantidad líneas con llave no encontrada en codiguera": ws.Range("B14").Value = CLng(resumen("lineas_llave_no_encontrada"))
    ws.Range("A15").Value = "Cantidad líneas existentes pero no incluidas en informe": ws.Range("B15").Value = CLng(resumen("lineas_no_incluidas"))
    ws.Range("B10:B11").NumberFormat = "#,##0.00"
    ws.Columns("A:B").AutoFit
End Sub

Private Function GuardarReportePesadoGG(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mesNum As Long) As String
    Dim carpetaSalida As String
    Dim fileName As String
    Dim ruta As String
    On Error GoTo EH

    carpetaSalida = RutaReportesGeneradosActiva()
    AsegurarCarpetaExiste carpetaSalida
    fileName = "Informe_GG_Base_Detallada_Pesada_" & anio & "_" & Format$(mesNum, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsb"
    ruta = CombinarRuta(carpetaSalida, fileName)
    wbOut.SaveAs Filename:=ruta, FileFormat:=xlExcel12
    GuardarReportePesadoGG = ruta
    Exit Function
EH:
    Err.Raise Err.Number, "GuardarReportePesadoGG", "Error guardando archivo pesado: " & Err.Description & " | Ruta: " & ruta
End Function

Private Function EncabezadosBasePesada() As Variant
    EncabezadosBasePesada = Array("Origen", "ArchivoOrigen", "HojaOrigen", "FilaOrigen", "Año", "MesNum", "MesNombre", "FechaValor", "Clave Llave presupuestal", "Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Finac", "Der-F", "PG", "Spg", "Proy", "Rubro", "R. Aux", "UE", "Dep", "Obra", "Der. Obra", "Serv", "SNIIP", "Ejecutado", "Asignado", "ImporteMN", "Incluir_en_Informe", "Estado_Codiguera")
End Function

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

Private Function UnirEncabezados(ByVal baseHeaders As Variant, ByVal prefE As Variant, ByVal prefA As Variant) As Variant
    Dim total As Long, res() As Variant, i As Long, p As Long
    total = UBound(baseHeaders) + 1 + UBound(prefE) + 1 + UBound(prefA) + 1
    ReDim res(0 To total - 1)
    p = 0
    For i = LBound(baseHeaders) To UBound(baseHeaders): res(p) = baseHeaders(i): p = p + 1: Next i
    For i = LBound(prefE) To UBound(prefE): res(p) = prefE(i): p = p + 1: Next i
    For i = LBound(prefA) To UBound(prefA): res(p) = prefA(i): p = p + 1: Next i
    UnirEncabezados = res
End Function

Private Function ObtenerColumnaPesada(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long, k As String
    For i = LBound(aliases) To UBound(aliases)
        k = LimpiarTexto(LCase$(CStr(aliases(i))))
        If headers.Exists(k) Then ObtenerColumnaPesada = CLng(headers(k)): Exit Function
    Next i
End Function

Private Function ValorMatriz(ByRef arr As Variant, ByVal fila As Long, ByVal col As Long) As Variant
    If col <= 0 Then
        ValorMatriz = Empty
    Else
        ValorMatriz = arr(fila, col)
    End If
End Function

Private Function NormalizarIncluirPesado(ByVal valor As Variant) As String
    Dim s As String
    s = Replace(UCase$(LimpiarTexto(TextoSeguro(valor))), " ", "")
    If s = "SI" Or s = "SÍ" Then
        NormalizarIncluirPesado = "SI"
    Else
        NormalizarIncluirPesado = "NO"
    End If
End Function

Private Function NormalizarClaveTextoPesada(ByVal valor As Variant) As String
    Dim s As String
    s = TextoSeguro(valor)
    s = Replace(s, "|", "-")
    s = Replace(s, " ", "")
    NormalizarClaveTextoPesada = s
End Function

Private Function NombreArchivoDesdeRutaPesada(ByVal rutaArchivo As String) As String
    Dim p As Long
    p = InStrRev(rutaArchivo, "\")
    If p > 0 Then
        NombreArchivoDesdeRutaPesada = Mid$(rutaArchivo, p + 1)
    Else
        NombreArchivoDesdeRutaPesada = rutaArchivo
    End If
End Function

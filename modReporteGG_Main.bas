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
    Dim wbE As Workbook
    Dim wbC As Workbook
    Dim wbOut As Workbook
    Dim wbA As Workbook
    Dim wsE As Worksheet
    Dim wsC As Worksheet
    Dim wsBase As Worksheet
    Dim wsA As Worksheet
    Dim wsBasePorc As Worksheet
    Dim dictCod As Object
    Dim dictAgg As Object
    Dim dictLlavesCodiguera As Object
    Dim diag As Object
    Dim dictAsignado As Object
    Dim dictPorcEjec As Object
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

    etapaActual = "buscando archivo de ejecuciones"
    archivoEjec = ObtenerArchivoMasReciente(RutaCarpetaEjecucionesActiva())
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 102, procedimiento, "No se encontró archivo de ejecuciones en: " & RutaCarpetaEjecucionesActiva()
    End If

    etapaActual = "buscando codiguera"
    archivoCod = ResolverArchivoCodiguera(RutaCodigueraActiva())
    If Len(archivoCod) = 0 Then
        Err.Raise vbObjectError + 103, procedimiento, "No se encontró archivo de codiguera en: " & RutaCodigueraActiva()
    End If

    etapaActual = "buscando asignados por fecha de creación"
    archivoAsignados = ObtenerArchivoMasRecientePorFechaCreacion(RutaCarpetaAsignadosGastosActiva())

    If Len(archivoAsignados) = 0 Then
        Err.Raise vbObjectError + 117, procedimiento, "No se encontró archivo de asignados en: " & RutaCarpetaAsignadosGastosActiva()
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
    Set wsA = wbA.Worksheets(1)

    etapaActual = "leyendo codiguera"
    LeerCodiguera wsC, dictCod, dictLlavesCodiguera, diag

    etapaActual = "leyendo ejecuciones y acumulando"
    LeerEjecucionesYAcumular wsE, anio, mesCierre, dictCod, dictAgg, diag
    LeerAsignadosYAcumular wsA, dictCod, dictLlavesCodiguera, dictAsignado, diag, wbC, anio, archivoAsignados

    If diag.Exists("asignados_faltantes") Then
        etapaActual = "guardando codiguera por nuevas llaves de asignados"
        wbC.Save
        EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre
        wbE.Close False: wbA.Close False: wbC.Close True
        MsgBox "Se agregaron nuevas llaves presupuestales del archivo de asignados a la codiguera. Debe clasificarlas, marcar Incluir_en_Informe cuando corresponda y volver a generar el reporte.", vbExclamation
        Exit Sub
    End If

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
    hojaReporteActual = "% ejecución " & anio
    etapaActual = "creando hoja % ejecución"
    etapaVisual = "iniciando hoja % ejecución"
    CrearHojaPorcEjecucion wbOut, wsBasePorc, anio, mesCierre, etapaVisual
    wsBase.Visible = xlSheetVeryHidden
    wsBasePorc.Visible = xlSheetVeryHidden

    etapaActual = "guardando reporte liviano"
    rutaFinal = GuardarReporteLiviano(wbOut, anio, mesCierre)

    etapaActual = "escribiendo diagnóstico"
    EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

    etapaActual = "cerrando archivos"
    wbOut.Close False
    wbE.Close False
    wbA.Close False
    wbC.Close False

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

    msg = "Error al generar reporte." & vbCrLf & vbCrLf & _
          "Procedimiento: " & procedimiento & vbCrLf & _
          "Etapa: " & etapaActual & vbCrLf & _
          "Etapa visual: " & etapaVisualMsg & vbCrLf & _
          "Err.Number: " & errNum & vbCrLf & _
          "Err.Description: " & errDesc & vbCrLf & _
          "Err.Source: " & errSource & vbCrLf & _
          "Erl: " & errLine & vbCrLf & _
          "Workbook salida: " & wbOutMsg & vbCrLf & _
          "Existe Base_Agregada: " & existeBase & vbCrLf & _
          "Última fila Base_Agregada: " & ultimaFilaBase & vbCrLf & _
          "Última columna Base_Agregada: " & ultimaColBase & vbCrLf & _
          "Nombre hoja reporte: " & IIf(Len(hojaReporteActual) > 0, hojaReporteActual, "(no determinada)") & vbCrLf & _
          "Archivo ejecuciones: " & archivoEjecMsg & vbCrLf & _
          "Archivo codiguera: " & archivoCodMsg & vbCrLf & _
          "Archivo asignados: " & archivoAsignadosMsg & vbCrLf & _
          "Salida: " & salidaMsg & vbCrLf & _
          "Carpeta base local: " & ThisWorkbook.Path & vbCrLf & _
          DiagnosticoRutasActivas()

    Debug.Print String(100, "-")
    Debug.Print msg
    Debug.Print String(100, "-")

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
    If Not wbE Is Nothing Then wbE.Close False
    If Not wbA Is Nothing Then wbA.Close False
    If Not wbC Is Nothing Then wbC.Close False
    On Error GoTo 0

    MsgBox msg, vbCritical
End Sub

Public Sub LeerAsignadosYAcumular(ByVal ws As Worksheet, ByVal dictCod As Object, ByVal dictLlavesCodiguera As Object, ByRef dictAsignado As Object, ByRef diag As Object, Optional ByVal wbCodiguera As Workbook, Optional ByVal anioFiltro As Long = 0, Optional ByVal archivoAsignados As String = "")
    Dim arr As Variant, headers As Object, i As Long, clave As String, info As Variant, keyAgg As String, monto As Double
    Dim colAnio As Long, anioFila As Long
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    ValidarColumnasAsignados headers
    colAnio = ObtenerColumnaOpcional(headers, Array("año", "anio", "ejercicio", "ej"))

    For i = 2 To UBound(arr, 1)
        If colAnio > 0 And anioFiltro > 0 Then
            If IsNumeric(arr(i, colAnio)) Then
                anioFila = CLng(arr(i, colAnio))
                If anioFila <> anioFiltro Then GoTo SiguienteFila
            End If
        End If
        clave = ConstruirClaveLlavePresupuestalCodiguera(arr(i, ObtenerColumna(headers, Array("finac"))), arr(i, ObtenerColumna(headers, Array("der-f"))), arr(i, ObtenerColumna(headers, Array("pg"))), arr(i, ObtenerColumna(headers, Array("spg"))), arr(i, ObtenerColumna(headers, Array("proy"))), arr(i, ObtenerColumna(headers, Array("rubro"))), arr(i, ObtenerColumna(headers, Array("r. aux"))), arr(i, ObtenerColumna(headers, Array("ue"))), arr(i, ObtenerColumna(headers, Array("dep"))), arr(i, ObtenerColumna(headers, Array("obra"))), arr(i, ObtenerColumna(headers, Array("der. obra"))), arr(i, ObtenerColumna(headers, Array("serv"))), arr(i, ObtenerColumna(headers, Array("sniip"))))
        If dictCod.Exists(clave) Then
            info = dictCod(clave)
            keyAgg = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3))
            monto = CDbl(0 + arr(i, ObtenerColumna(headers, Array("asignado"))))
            If Not dictAsignado.Exists(keyAgg) Then dictAsignado.Add keyAgg, 0#
            dictAsignado(keyAgg) = dictAsignado(keyAgg) + monto
        ElseIf Not dictLlavesCodiguera.Exists(clave) Then
            RegistrarYAgregarLlaveAsignadoFaltante wbCodiguera, diag, clave, arr, headers, i, archivoAsignados
            If Not dictLlavesCodiguera.Exists(clave) Then dictLlavesCodiguera.Add clave, True
        End If
SiguienteFila:
    Next i
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
    If diag.Exists("asignados_faltantes") Then
        ws.Cells(8, 1).Value = "Llaves de asignados no encontradas en codiguera"
        ws.Range("A9:S9").Value = Array("Origen", "Archivo", "Fila origen", "Clave normalizada", "Llave presupuestal", "Finac", "Der-F", "PG", "Spg", "Proy", "Rubro", "R. Aux", "UE", "Dep", "Obra", "Der. Obra", "Serv", "SNIIP", "Estado")
        f = 10
        Set d = diag("asignados_faltantes")
        For Each k In d.Keys
            it = d(k)
            ws.Range("A" & f & ":S" & f).Value = it
            f = f + 1
        Next k
    End If
    ws.Columns("A:B").AutoFit
End Sub

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

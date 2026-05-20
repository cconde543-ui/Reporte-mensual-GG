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
    Dim wbE As Workbook
    Dim wbC As Workbook
    Dim wbOut As Workbook
    Dim wsE As Worksheet
    Dim wsC As Worksheet
    Dim wsBase As Worksheet
    Dim dictCod As Object
    Dim dictAgg As Object
    Dim diag As Object
    Dim rutaFinal As String

    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errLine As Long
    Dim msg As String

    procedimiento = "Generar_Reporte_GG_Desde_Panel"
    Debug.Print "Inicio Generar_Reporte_GG_Desde_Panel: " & Now

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
    Set diag = CreateObject("Scripting.Dictionary")

    etapaActual = "buscando archivo de ejecuciones"
    archivoEjec = ObtenerArchivoMasReciente(RUTA_CARPETA_EJECUCIONES)
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 102, procedimiento, "No se encontró archivo de ejecuciones en: " & RUTA_CARPETA_EJECUCIONES
    End If

    etapaActual = "buscando codiguera"
    archivoCod = ResolverArchivoCodiguera(RUTA_CODIGUERA)
    If Len(archivoCod) = 0 Then
        Err.Raise vbObjectError + 103, procedimiento, "No se encontró archivo de codiguera en: " & RUTA_CODIGUERA
    End If

    etapaActual = "abriendo archivo de ejecuciones"
    Set wbE = Workbooks.Open(archivoEjec, ReadOnly:=True)

    etapaActual = "abriendo codiguera"
    Set wbC = Workbooks.Open(archivoCod, ReadOnly:=True)

    etapaActual = "leyendo hojas de origen"
    Set wsE = ObtenerHojaEjecuciones(wbE)
    Set wsC = ObtenerHojaCodiguera(wbC)

    etapaActual = "leyendo codiguera"
    LeerCodiguera wsC, dictCod, diag

    etapaActual = "leyendo ejecuciones y acumulando"
    LeerEjecucionesYAcumular wsE, anio, mesCierre, dictCod, dictAgg, diag

    etapaActual = "creando workbook de salida"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBase = wbOut.Worksheets(1)
    wsBase.Name = "Base_Agregada"

    etapaActual = "construyendo base agregada"
    ConstruirBaseAgregadaReporte wsBase, dictAgg

    etapaActual = "creando reporte visual"
    CrearReporteEjecucionMensual wbOut, wsBase, anio, mesCierre
    wsBase.Visible = xlSheetVeryHidden

    etapaActual = "guardando reporte liviano"
    rutaFinal = GuardarReporteLiviano(wbOut, anio, mesCierre)

    etapaActual = "escribiendo diagnóstico"
    EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

    etapaActual = "cerrando archivos"
    wbOut.Close False
    wbE.Close False
    wbC.Close False

    MsgBox "Reporte generado: " & rutaFinal, vbInformation
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl

    msg = "Error al generar reporte." & vbCrLf & vbCrLf & _
          "Procedimiento: " & procedimiento & vbCrLf & _
          "Etapa: " & etapaActual & vbCrLf & _
          "Err.Number: " & errNum & vbCrLf & _
          "Err.Description: " & errDesc & vbCrLf & _
          "Err.Source: " & errSource & vbCrLf & _
          "Erl: " & errLine & vbCrLf & _
          "Archivo ejecuciones: " & IIf(Len(archivoEjec) > 0, archivoEjec, "(no detectado)") & vbCrLf & _
          "Archivo codiguera: " & IIf(Len(archivoCod) > 0, archivoCod, "(no detectado)") & vbCrLf & _
          "Salida: " & IIf(Len(rutaFinal) > 0, rutaFinal, RUTA_REPORTES_GENERADOS)

    Debug.Print String(100, "-")
    Debug.Print msg
    Debug.Print String(100, "-")

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
    If Not wbE Is Nothing Then wbE.Close False
    If Not wbC Is Nothing Then wbC.Close False
    On Error GoTo 0

    MsgBox msg, vbCritical
End Sub

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictCod As Object, ByRef diag As Object)
    On Error GoTo EH

    Dim arr As Variant, headers As Object, i As Long, incluir As String, clave As String, info As Variant
    Dim colTitular As Long

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)
    colTitular = ObtenerColumna(headers, Array("titular"))
    If colTitular = 0 Then Err.Raise vbObjectError + 201, "LeerCodiguera", "Falta columna Titular en codiguera."

    For i = 2 To UBound(arr, 1)
        incluir = Replace(UCase$(Trim$(CStr(arr(i, ObtenerColumna(headers, Array("incluir_en_informe")))))), " ", "")
        If incluir = "SI" Then
            clave = ConstruirClavePresupuestal(arr(i, ObtenerColumna(headers, Array("finac código numérico"))), arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), arr(i, ObtenerColumna(headers, Array("pg código numérico"))), arr(i, ObtenerColumna(headers, Array("spg código numérico"))), arr(i, ObtenerColumna(headers, Array("proy", "proyecto"))), arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), arr(i, ObtenerColumna(headers, Array("ue código numérico"))), arr(i, ObtenerColumna(headers, Array("dep código numérico"))), arr(i, ObtenerColumna(headers, Array("obra código numérico"))), arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), arr(i, ObtenerColumna(headers, Array("serv código numérico"))), arr(i, ObtenerColumna(headers, Array("snip código numérico"))))
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
                clave = ConstruirClavePresupuestal(arr(i, ObtenerColumna(headers, Array("finac código numérico"))), arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), arr(i, ObtenerColumna(headers, Array("pg código numérico"))), arr(i, ObtenerColumna(headers, Array("spg código numérico"))), arr(i, ObtenerColumna(headers, Array("proyecto", "proy"))), arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), arr(i, ObtenerColumna(headers, Array("ue código numérico"))), arr(i, ObtenerColumna(headers, Array("dep código numérico"))), arr(i, ObtenerColumna(headers, Array("obra código numérico"))), arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), arr(i, ObtenerColumna(headers, Array("serv código numérico"))), arr(i, ObtenerColumna(headers, Array("snip código numérico"))))
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

Public Sub ConstruirBaseAgregadaReporte(ByVal ws As Worksheet, ByVal dictAgg As Object)
    Dim fila As Long, dictKey As Variant, partes() As String, importeSalida As Double, factor As Double
    ws.Range("A1:G1").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")
    factor = FactorEscalaImporte()
    fila = 2
    For Each dictKey In dictAgg.Keys
        partes = Split(CStr(dictKey), "|")
        ws.Cells(fila, 1).Value = LimpiarTexto(CStr(partes(0)))
        ws.Cells(fila, 2).Value = LimpiarTexto(CStr(partes(1)))
        ws.Cells(fila, 3).Value = LimpiarTexto(CStr(partes(2)))
        ws.Cells(fila, 4).Value = LimpiarTexto(CStr(partes(3)))
        ws.Cells(fila, 5).Value = CLng(partes(4))
        ws.Cells(fila, 6).Value = MesesESMin()(CLng(partes(4)) - 1)
        importeSalida = CDbl(dictAgg(dictKey)) / factor
        ws.Cells(fila, 7).Value = importeSalida
        fila = fila + 1
    Next dictKey
End Sub

Public Function GuardarReporteLiviano(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mesNum As Long) As String
    Dim ruta As String, fileName As String
    On Error GoTo EH

    AsegurarCarpetaExiste RUTA_REPORTES_GENERADOS
    fileName = "Informe_GG_Ejecucion_Mensual_" & anio & "_" & Format$(mesNum, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsx"
    ruta = RUTA_REPORTES_GENERADOS & "\" & fileName
    wbOut.SaveAs ruta, xlOpenXMLWorkbook
    GuardarReporteLiviano = ruta
    Exit Function
EH:
    Err.Raise Err.Number, "GuardarReporteLiviano", "Error guardando reporte liviano: " & Err.Description & " | Ruta: " & ruta
End Function

Public Sub EscribirDiagnostico(ByVal wb As Workbook, ByVal diag As Object, ByVal archivoEjec As String, ByVal archivoCod As String, ByVal anio As Long, ByVal mesNum As Long)
    Dim ws As Worksheet
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
    ws.Columns("A:B").AutoFit
End Sub

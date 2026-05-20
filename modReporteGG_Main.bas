Option Explicit

Public Sub Generar_Reporte_GG_Desde_Panel()
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

    Set wsPanel = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)

    If Not IsNumeric(wsPanel.Range("B3").Value) Then
        Err.Raise vbObjectError + 100, , "Año inválido en B3"
    End If

    anio = CLng(wsPanel.Range("B3").Value)
    mesTxt = CStr(wsPanel.Range("B4").Value)
    mesCierre = MesTextoANumero(mesTxt)

    If mesCierre < 1 Or mesCierre > 12 Then
        Err.Raise vbObjectError + 101, , "Mes inválido en B4"
    End If

    Set dictCod = CreateObject("Scripting.Dictionary")
    Set dictAgg = CreateObject("Scripting.Dictionary")
    Set diag = CreateObject("Scripting.Dictionary")

    archivoEjec = ObtenerArchivoMasReciente(RUTA_CARPETA_EJECUCIONES)
    archivoCod = ResolverArchivoCodiguera(RUTA_CODIGUERA)

    Set wbE = Workbooks.Open(archivoEjec, ReadOnly:=True)
    Set wbC = Workbooks.Open(archivoCod, ReadOnly:=True)

    Set wsE = ObtenerHojaEjecuciones(wbE)
    Set wsC = ObtenerHojaCodiguera(wbC)

    LeerCodiguera wsC, dictCod, diag
    LeerEjecucionesYAcumular wsE, anio, mesCierre, dictCod, dictAgg, diag

    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBase = wbOut.Worksheets(1)
    wsBase.Name = "Base_Agregada"

    ConstruirBaseAgregadaReporte wsBase, dictAgg
    CrearReporteEjecucionMensual wbOut, wsBase, anio, mesCierre
    wsBase.Visible = xlSheetVeryHidden

    rutaFinal = GuardarReporteLiviano(wbOut, anio, mesCierre)
    EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

    wbOut.Close False
    wbE.Close False
    wbC.Close False

    MsgBox "Reporte generado: " & rutaFinal, vbInformation
End Sub

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictCod As Object, ByRef diag As Object)
    Dim arr As Variant
    Dim headers As Object
    Dim i As Long
    Dim incluir As String
    Dim clave As String
    Dim info As Variant

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    For i = 2 To UBound(arr, 1)
        incluir = Replace(UCase$(Trim$(CStr(arr(i, ObtenerColumna(headers, Array("incluir_en_informe")))))), " ", "")
        If incluir = "SI" Then
            clave = ConstruirClavePresupuestal( _
                arr(i, ObtenerColumna(headers, Array("finac código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("pg código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("spg código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("proy", "proyecto"))), _
                arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("ue código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("dep código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("obra código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("serv código numérico"))), _
                arr(i, ObtenerColumna(headers, Array("snip código numérico"))))

            info = Array( _
                arr(i, ObtenerColumna(headers, Array("finac"))), _
                arr(i, ObtenerColumna(headers, Array("nivel_1"))), _
                arr(i, ObtenerColumna(headers, Array("nivel_2"))), _
                arr(i, ObtenerColumna(headers, Array("nivel_3"))))

            dictCod(clave) = info
        End If
    Next i
End Sub

Public Sub LeerEjecucionesYAcumular( _
    ByVal ws As Worksheet, _
    ByVal anio As Long, _
    ByVal mesCierre As Long, _
    ByVal dictCod As Object, _
    ByRef dictAgg As Object, _
    ByRef diag As Object)

    Dim arr As Variant
    Dim headers As Object
    Dim i As Long
    Dim fechaValor As Date
    Dim clave As String
    Dim info As Variant
    Dim mesNum As Long
    Dim aggregateKey As String
    Dim importeMN As Double

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set headers = MapearEncabezados(arr)

    For i = 2 To UBound(arr, 1)
        If TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(headers, Array("fecha valor"))), fechaValor) Then
            If Year(fechaValor) = anio And Month(fechaValor) <= mesCierre Then
                clave = ConstruirClavePresupuestal( _
                    arr(i, ObtenerColumna(headers, Array("finac código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("der-f código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("pg código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("spg código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("proyecto", "proy"))), _
                    arr(i, ObtenerColumna(headers, Array("rubro código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("r. aux código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("ue código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("dep código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("obra código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("der. obra código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("serv código numérico"))), _
                    arr(i, ObtenerColumna(headers, Array("snip código numérico"))))

                If dictCod.Exists(clave) Then
                    info = dictCod(clave)
                    mesNum = Month(fechaValor)
                    importeMN = CDbl(0 + arr(i, ObtenerColumna(headers, Array("importe moneda nacional"))))

                    aggregateKey = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3)) & "|" & CStr(mesNum)

                    If Not dictAgg.Exists(aggregateKey) Then
                        dictAgg.Add aggregateKey, 0
                    End If

                    dictAgg(aggregateKey) = dictAgg(aggregateKey) + importeMN
                End If
            End If
        End If
    Next i
End Sub

Public Sub ConstruirBaseAgregadaReporte(ByVal ws As Worksheet, ByVal dictAgg As Object)
    Dim fila As Long
    Dim dictKey As Variant
    Dim partes() As String
    Dim importeSalida As Double

    ws.Range("A1:G1").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")

    fila = 2
    For Each dictKey In dictAgg.Keys
        partes = Split(CStr(dictKey), "|")

        ws.Cells(fila, 1).Value = partes(0)
        ws.Cells(fila, 2).Value = partes(1)
        ws.Cells(fila, 3).Value = partes(2)
        ws.Cells(fila, 4).Value = partes(3)
        ws.Cells(fila, 5).Value = CLng(partes(4))
        ws.Cells(fila, 6).Value = MesesES()(CLng(partes(4)) - 1)

        If MOSTRAR_EN_MILES Then
            importeSalida = CDbl(dictAgg(dictKey)) / 1000
        Else
            importeSalida = CDbl(dictAgg(dictKey))
        End If
        ws.Cells(fila, 7).Value = importeSalida

        fila = fila + 1
    Next dictKey
End Sub

Public Function GuardarReporteLiviano(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mesNum As Long) As String
    Dim ruta As String
    Dim fileName As String

    AsegurarCarpetaExiste RUTA_REPORTES_GENERADOS

    fileName = "Informe_GG_Ejecucion_Mensual_" & anio & "_" & Format$(mesNum, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsx"
    ruta = RUTA_REPORTES_GENERADOS & "\" & fileName

    wbOut.SaveAs ruta, xlOpenXMLWorkbook
    GuardarReporteLiviano = ruta
End Function

Public Sub EscribirDiagnostico( _
    ByVal wb As Workbook, _
    ByVal diag As Object, _
    ByVal archivoEjec As String, _
    ByVal archivoCod As String, _
    ByVal anio As Long, _
    ByVal mesNum As Long)

    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(DIAG_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = DIAG_SHEET_NAME

    ws.Range("A1").Value = "Campo"
    ws.Range("B1").Value = "Valor"

    ws.Cells(2, 1).Value = "Archivo ejecuciones"
    ws.Cells(2, 2).Value = archivoEjec
    ws.Cells(3, 1).Value = "Archivo codiguera"
    ws.Cells(3, 2).Value = archivoCod
    ws.Cells(4, 1).Value = "Hoja codiguera"
    ws.Cells(4, 2).Value = "Codiguera"
    ws.Cells(5, 1).Value = "Año"
    ws.Cells(5, 2).Value = anio
    ws.Cells(6, 1).Value = "Mes cierre"
    ws.Cells(6, 2).Value = mesNum

    ws.Columns("A:B").AutoFit
End Sub

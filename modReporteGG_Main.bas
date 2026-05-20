Option Explicit

Public Sub Generar_Reporte_GG_Desde_Panel()
    Dim wsPanel As Worksheet, anio As Long, mesTxt As String, mesCierre As Long
    Dim archivoEjec As String, archivoCod As String
    Dim wbE As Workbook, wbC As Workbook, wbOut As Workbook
    Dim wsE As Worksheet, wsC As Worksheet, wsBase As Worksheet
    Dim dictCod As Object, dictAgg As Object, diag As Object
    Dim rutaFinal As String

    Set wsPanel = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    If Not IsNumeric(wsPanel.Range("B3").Value) Then Err.Raise vbObjectError + 100, , "Año inválido en B3"
    anio = CLng(wsPanel.Range("B3").Value)
    mesTxt = CStr(wsPanel.Range("B4").Value)
    mesCierre = MesTextoANumero(mesTxt)
    If mesCierre < 1 Or mesCierre > 12 Then Err.Raise vbObjectError + 101, , "Mes inválido en B4"

    Set dictCod = CreateObject("Scripting.Dictionary")
    Set dictAgg = CreateObject("Scripting.Dictionary")
    Set diag = CreateObject("Scripting.Dictionary")

    archivoEjec = ObtenerArchivoMasReciente(RUTA_CARPETA_EJECUCIONES)
    archivoCod = ResolverArchivoCodiguera(RUTA_CODIGUERA)
    Set wbE = Workbooks.Open(archivoEjec, ReadOnly:=True)
    Set wbC = Workbooks.Open(archivoCod, ReadOnly:=True)
    Set wsE = wbE.Worksheets(1)
    Set wsC = wbC.Worksheets("Codiguera")

    LeerCodiguera wsC, dictCod, diag
    LeerEjecucionesYAcumular wsE, anio, mesCierre, dictCod, dictAgg, diag

    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsBase = wbOut.Worksheets(1): wsBase.Name = "Base_Agregada"
    ConstruirBaseAgregadaReporte wsBase, dictAgg
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre
    wsBase.Visible = xlSheetVeryHidden

    rutaFinal = GuardarReporteLiviano(wbOut, anio, mesCierre)
    EscribirDiagnostico ThisWorkbook, diag, archivoEjec, archivoCod, anio, mesCierre

    wbOut.Close False: wbE.Close False: wbC.Close False
    MsgBox "Reporte generado: " & rutaFinal, vbInformation
End Sub

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictCod As Object, ByRef diag As Object)
    Dim arr, h As Object, i As Long, inc As String, k As String, info As Variant
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)
    For i = 2 To UBound(arr, 1)
        inc = Replace(UCase$(Trim$(CStr(arr(i, ObtenerColumna(h, Array("incluir_en_informe")))))), " ", "")
        If inc = "SI" Then
            k = ConstruirClavePresupuestal(arr(i, ObtenerColumna(h, Array("finac código numérico"))), arr(i, ObtenerColumna(h, Array("der-f código numérico"))), arr(i, ObtenerColumna(h, Array("pg código numérico"))), arr(i, ObtenerColumna(h, Array("spg código numérico"))), arr(i, ObtenerColumna(h, Array("proy"))), arr(i, ObtenerColumna(h, Array("rubro código numérico"))), arr(i, ObtenerColumna(h, Array("r. aux código numérico"))), arr(i, ObtenerColumna(h, Array("ue código numérico"))), arr(i, ObtenerColumna(h, Array("dep código numérico"))), arr(i, ObtenerColumna(h, Array("obra código numérico"))), arr(i, ObtenerColumna(h, Array("der. obra código numérico"))), arr(i, ObtenerColumna(h, Array("serv código numérico"))), arr(i, ObtenerColumna(h, Array("snip código numérico"))))
            info = Array(arr(i, ObtenerColumna(h, Array("finac"))), arr(i, ObtenerColumna(h, Array("nivel_1"))), arr(i, ObtenerColumna(h, Array("nivel_2"))), arr(i, ObtenerColumna(h, Array("nivel_3"))))
            dictCod(k) = info
        End If
    Next i
End Sub

Public Sub LeerEjecucionesYAcumular(ByVal ws As Worksheet, ByVal anio As Long, ByVal mesCierre As Long, ByVal dictCod As Object, ByRef dictAgg As Object, ByRef diag As Object)
    Dim arr, h As Object, i As Long, fv As Date, k As String, info As Variant, mes As Long, akey As String, imp As Double
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)
    For i = 2 To UBound(arr, 1)
        If TryObtenerFechaValorSeguro(arr(i, ObtenerColumna(h, Array("fecha valor"))), fv) Then
            If Year(fv) = anio And Month(fv) <= mesCierre Then
                k = ConstruirClavePresupuestal(arr(i, ObtenerColumna(h, Array("finac código numérico"))), arr(i, ObtenerColumna(h, Array("der-f código numérico"))), arr(i, ObtenerColumna(h, Array("pg código numérico"))), arr(i, ObtenerColumna(h, Array("spg código numérico"))), arr(i, ObtenerColumna(h, Array("proyecto"))), arr(i, ObtenerColumna(h, Array("rubro código numérico"))), arr(i, ObtenerColumna(h, Array("r. aux código numérico"))), arr(i, ObtenerColumna(h, Array("ue código numérico"))), arr(i, ObtenerColumna(h, Array("dep código numérico"))), arr(i, ObtenerColumna(h, Array("obra código numérico"))), arr(i, ObtenerColumna(h, Array("der. obra código numérico"))), arr(i, ObtenerColumna(h, Array("serv código numérico"))), arr(i, ObtenerColumna(h, Array("snip código numérico"))))
                If dictCod.Exists(k) Then
                    info = dictCod(k): mes = Month(fv): imp = CDbl(0 + arr(i, ObtenerColumna(h, Array("importe moneda nacional"))))
                    akey = CStr(info(0)) & "|" & CStr(info(1)) & "|" & CStr(info(2)) & "|" & CStr(info(3)) & "|" & CStr(mes)
                    If Not dictAgg.Exists(akey) Then dictAgg.Add akey, 0#
                    dictAgg(akey) = dictAgg(akey) + imp
                End If
            End If
        End If
    Next i
End Sub

Public Sub ConstruirBaseAgregadaReporte(ByVal ws As Worksheet, ByVal dictAgg As Object)
    Dim r As Long, k As Variant, p() As String
    ws.Range("A1:G1").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")
    r = 2
    For Each k In dictAgg.Keys
        p = Split(CStr(k), "|")
        ws.Cells(r, 1).Value = p(0): ws.Cells(r, 2).Value = p(1): ws.Cells(r, 3).Value = p(2): ws.Cells(r, 4).Value = p(3)
        ws.Cells(r, 5).Value = CLng(p(4)): ws.Cells(r, 6).Value = MesesES()(CLng(p(4)) - 1)
        ws.Cells(r, 7).Value = IIf(MOSTRAR_EN_MILES, dictAgg(k) / 1000#, dictAgg(k))
        r = r + 1
    Next k
End Sub

Public Function GuardarReporteLiviano(ByVal wbOut As Workbook, ByVal anio As Long, ByVal mes As Long) As String
    Dim ruta As String, fn As String
    AsegurarCarpetaExiste RUTA_REPORTES_GENERADOS
    fn = "Informe_GG_Ejecucion_Mensual_" & anio & "_" & Format$(mes, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsx"
    ruta = RUTA_REPORTES_GENERADOS & "\" & fn
    wbOut.SaveAs ruta, xlOpenXMLWorkbook
    GuardarReporteLiviano = ruta
End Function

Public Sub EscribirDiagnostico(ByVal wb As Workbook, ByVal diag As Object, ByVal archivoEjec As String, ByVal archivoCod As String, ByVal anio As Long, ByVal mes As Long)
    Dim ws As Worksheet
    On Error Resume Next: Application.DisplayAlerts = False: wb.Worksheets(DIAG_SHEET_NAME).Delete: Application.DisplayAlerts = True: On Error GoTo 0
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = DIAG_SHEET_NAME
    ws.Range("A1:B1").Value = Array("Campo", "Valor")
    ws.Range("A2:B6").Value = Array(Array("Archivo ejecuciones", archivoEjec), Array("Archivo codiguera", archivoCod), Array("Hoja codiguera", "Codiguera"), Array("Año", anio), Array("Mes cierre", mes))
    ws.Columns("A:B").AutoFit
End Sub

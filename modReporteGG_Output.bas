Option Explicit

Public Sub CrearReporteEjecucionMensual(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    CrearTablaDinamicaOSalidaAgrupada wbOut, wsBase, anio, mesCierre
End Sub

Public Sub CrearTablaDinamicaOSalidaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, pivotCacheObj As PivotCache, pt As PivotTable, rg As Range
    Dim etapa As String, campo As String
    On Error GoTo EH

    If UltimaFilaConDatos(wsBase) < 2 Then Err.Raise vbObjectError + 700, , "No hay datos agregados para generar la tabla dinámica."
    LimpiarEncabezadosBase wsBase
    Debug.Print "Base hoja: " & wsBase.Name

    On Error Resume Next: Application.DisplayAlerts = False: wbOut.Worksheets("Ejec. Mensual " & anio).Delete: Application.DisplayAlerts = True: On Error GoTo EH
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count)): ws.Name = "Ejec. Mensual " & anio
    ws.Range("A1:F1").Merge: ws.Range("A1").Value = "EJECUCIÓN " & anio
    ws.Range("A1").Interior.Color = RGB(0, 112, 192): ws.Range("A1").Font.Color = vbWhite: ws.Range("A1").Font.Bold = True

    Set rg = wsBase.Range("A1").CurrentRegion
    Debug.Print "Rango fuente: " & rg.Address(External:=True)
    Debug.Print "Filas base: " & rg.Rows.Count & " | Columnas base: " & rg.Columns.Count
    Debug.Print "Encabezados base: " & Join(ObtenerEncabezadosBase(wsBase), " | ")

    etapa = "crear cache": Set pivotCacheObj = wbOut.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rg)
    etapa = "crear tabla": Set pt = pivotCacheObj.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ptGG")
    Debug.Print "Campos pivot: " & CamposPivotAsText(pt)

    ConfigurarCampoPivot pt, "Financiamiento", xlPageField, etapa, campo, wsBase
    ConfigurarCampoPivot pt, "Nivel_1", xlRowField, etapa, campo, wsBase
    ConfigurarCampoPivot pt, "Nivel_2", xlRowField, etapa, campo, wsBase
    ConfigurarCampoPivot pt, "Nivel_3", xlRowField, etapa, campo, wsBase
    ConfigurarCampoPivot pt, "MesNombre", xlColumnField, etapa, campo, wsBase
    If PivotFieldExiste(pt, "MesNum") Then pt.PivotFields("MesNum").Orientation = xlHidden
    If Not PivotFieldExiste(pt, "Importe") Then Err.Raise vbObjectError + 701, , "Falta campo Importe en tabla dinámica."
    pt.AddDataField pt.PivotFields("Importe"), "Suma de Importe", xlSum
    pt.RowAxisLayout xlTabularRow: pt.RepeatAllLabels xlRepeatLabels: pt.ShowDrillIndicators = True
    AplicarFormatoReporteGG ws, mesCierre
    Exit Sub
EH:
    Debug.Print "Error Pivot | etapa=" & etapa & " | campo=" & campo & " | nro=" & Err.Number & " | desc=" & Err.Description
    MsgBox "Falló tabla dinámica, se generará salida estática." & vbCrLf & "Etapa: " & etapa & vbCrLf & "Campo: " & campo & vbCrLf & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & "Campos pivot: " & IIf(pt Is Nothing, "(sin PT)", CamposPivotAsText(pt)) & vbCrLf & "Encabezados base: " & Join(ObtenerEncabezadosBase(wsBase), ", "), vbExclamation
    GenerarSalidaEstaticaAgrupada wbOut, wsBase, anio, mesCierre
End Sub

Private Sub ConfigurarCampoPivot(ByVal pt As PivotTable, ByVal nombre As String, ByVal orientation As XlPivotFieldOrientation, ByRef etapa As String, ByRef campo As String, ByVal wsBase As Worksheet)
    etapa = "configurar campo": campo = nombre
    If Not PivotFieldExiste(pt, nombre) Then
        Err.Raise vbObjectError + 702, , "Falta campo pivot: " & nombre & vbCrLf & "Campos pivot: " & CamposPivotAsText(pt) & vbCrLf & "Encabezados base: " & Join(ObtenerEncabezadosBase(wsBase), ", ")
    End If
    pt.PivotFields(nombre).Orientation = orientation
End Sub

Public Function PivotFieldExiste(ByVal pt As PivotTable, ByVal nombreCampo As String) As Boolean
    Dim pf As PivotField
    For Each pf In pt.PivotFields
        If StrComp(LimpiarTexto(pf.Name), LimpiarTexto(nombreCampo), vbTextCompare) = 0 Then PivotFieldExiste = True: Exit Function
    Next pf
End Function

Private Function CamposPivotAsText(ByVal pt As PivotTable) As String
    Dim pf As PivotField, t As String
    For Each pf In pt.PivotFields: t = t & IIf(Len(t) > 0, " | ", "") & pf.Name: Next pf
    CamposPivotAsText = t
End Function

Private Function ObtenerEncabezadosBase(ByVal wsBase As Worksheet) As Variant
    Dim lastCol As Long, i As Long, arr() As String
    lastCol = UltimaColConDatos(wsBase): ReDim arr(0 To lastCol - 1)
    For i = 1 To lastCol: arr(i - 1) = CStr(wsBase.Cells(1, i).Value): Next i
    ObtenerEncabezadosBase = arr
End Function

Private Sub LimpiarEncabezadosBase(ByVal wsBase As Worksheet)
    Dim i As Long, fixedHeaders As Variant
    fixedHeaders = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "MesNum", "MesNombre", "Importe")
    For i = 1 To 7: wsBase.Cells(1, i).Value = fixedHeaders(i - 1): Next i
End Sub

Private Sub GenerarSalidaEstaticaAgrupada(ByVal wbOut As Workbook, ByVal wsBase As Worksheet, ByVal anio As Long, ByVal mesCierre As Long)
    Dim ws As Worksheet, src As Variant, i As Long, outR As Long
    Set ws = wbOut.Worksheets("Ejec. Mensual " & anio)
    ws.Cells.Clear
    ws.Range("A1").Value = "EJECUCIÓN " & anio
    ws.Range("A3:H3").Value = Array("Financiamiento", "Nivel_1", "Nivel_2", "Nivel_3", "Mes", "MesNum", "Importe", "Total general")
    src = wsBase.Range("A1").CurrentRegion.Value2
    outR = 4
    For i = 2 To UBound(src, 1)
        ws.Cells(outR, 1).Resize(1, 7).Value = Array(src(i, 1), src(i, 2), src(i, 3), src(i, 4), src(i, 6), src(i, 5), src(i, 7))
        ws.Cells(outR, 8).FormulaR1C1 = "=RC[-1]"
        outR = outR + 1
    Next i
    ws.Range("A3:H3").AutoFilter
    ws.Columns.AutoFit
End Sub

Public Sub AplicarFormatoReporteGG(ByVal ws As Worksheet, ByVal mesCierre As Long)
    Dim lastRow As Long, lastCol As Long
    lastRow = UltimaFilaConDatos(ws): lastCol = UltimaColConDatos(ws)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol)).Interior.Color = RGB(0, 112, 192)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol)).Font.Color = vbWhite
    ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).NumberFormat = "#,##0"
    ws.Columns.AutoFit
End Sub

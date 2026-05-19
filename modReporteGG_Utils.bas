Option Explicit

Public Function ResolverArchivoCodiguera(ByVal rutaCodiguera As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(rutaCodiguera) Then
        ResolverArchivoCodiguera = rutaCodiguera
    ElseIf fso.FolderExists(rutaCodiguera) Then
        ResolverArchivoCodiguera = ObtenerArchivoMasReciente(rutaCodiguera)
    Else
        ResolverArchivoCodiguera = vbNullString
    End If
End Function

Public Function ObtenerArchivoMasReciente(ByVal carpeta As String) As String
    Dim fso As Object, folder As Object, archivo As Object
    Dim fechaMax As Date, candidato As String

    On Error GoTo ControlError
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function

    Set folder = fso.GetFolder(carpeta)
    fechaMax = #1/1/1900#

    For Each archivo In folder.Files
        If EsExtensionExcel(CStr(archivo.Name)) Then
            If archivo.DateLastModified > fechaMax Then
                fechaMax = archivo.DateLastModified
                candidato = CStr(archivo.Path)
            End If
        End If
    Next archivo

    ObtenerArchivoMasReciente = candidato
    Exit Function
ControlError:
    ObtenerArchivoMasReciente = vbNullString
End Function

Public Function EsExtensionExcel(ByVal nombreArchivo As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(nombreArchivo, InStrRev(nombreArchivo, ".") + 1))
    EsExtensionExcel = (ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb")
End Function

Public Function ObtenerHojaCodiguera(ByVal wb As Workbook) As Worksheet
    Set ObtenerHojaCodiguera = ObtenerHojaPorNombreExacto(wb, "Codiguera")
End Function

Public Function ObtenerHojaEjecuciones(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, score As Long, mejorScore As Long
    Dim arr As Variant, h As Object

    For Each ws In wb.Worksheets
        If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
            arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
            Set h = MapearEncabezados(arr)
            score = 0
            score = score + IIf(ExisteHeader(h, "Fecha valor"), 1, 0)
            score = score + IIf(ExisteHeader(h, "Importe moneda nacional"), 1, 0)
            score = score + IIf(ExisteHeader(h, "Proyecto"), 1, 0)
            score = score + IIf(ExisteHeader(h, "SNIP Código numérico"), 1, 0)
            If score > mejorScore Then
                mejorScore = score
                Set ObtenerHojaEjecuciones = ws
            End If
        End If
    Next ws
End Function

Public Function ExisteHeader(ByVal mapHeaders As Object, ByVal aliasName As String) As Boolean
    ExisteHeader = mapHeaders.Exists(NormalizarEncabezado(aliasName))
End Function

Public Function ObtenerHojaPorNombreExacto(ByVal wb As Workbook, ByVal nombreHoja As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombreHoja, vbTextCompare) = 0 Then Set ObtenerHojaPorNombreExacto = ws: Exit Function
    Next ws
End Function

Public Function MapearEncabezados(ByRef arrDatos As Variant) As Object
    Dim dict As Object, col As Long, nombre As String
    Set dict = CreateObject("Scripting.Dictionary")
    For col = 1 To UBound(arrDatos, 2)
        nombre = NormalizarEncabezado(CStr(arrDatos(1, col)))
        If Len(nombre) > 0 Then If Not dict.Exists(nombre) Then dict.Add nombre, col
    Next col
    Set MapearEncabezados = dict
End Function

Public Function ListarHeaders(ByVal mapHeaders As Object) As String
    If mapHeaders.Count = 0 Then
        ListarHeaders = "(sin encabezados)"
    Else
        ListarHeaders = Join(mapHeaders.Keys, " | ")
    End If
End Function

Public Function ObtenerColumna(ByVal mapHeaders As Object, ByVal aliases As Variant) As Long
    Dim i As Long, key As String, intentados As String
    For i = LBound(aliases) To UBound(aliases)
        key = NormalizarEncabezado(CStr(aliases(i)))
        If mapHeaders.Exists(key) Then ObtenerColumna = CLng(mapHeaders(key)): Exit Function
        If Len(intentados) > 0 Then intentados = intentados & ", "
        intentados = intentados & CStr(aliases(i))
    Next i
    Err.Raise vbObjectError + 3000, "ObtenerColumna", "Falta columna obligatoria. Alias intentados: [" & intentados & "] Encabezados detectados: [" & ListarHeaders(mapHeaders) & "]"
End Function

Public Function ObtenerColumnaOpcional(ByVal mapHeaders As Object, ByVal aliases As Variant) As Long
    Dim i As Long, key As String
    For i = LBound(aliases) To UBound(aliases)
        key = NormalizarEncabezado(CStr(aliases(i)))
        If mapHeaders.Exists(key) Then ObtenerColumnaOpcional = CLng(mapHeaders(key)): Exit Function
    Next i
    ObtenerColumnaOpcional = 0
End Function

Public Function NormalizarEncabezado(ByVal texto As String) As String
    Dim t As String
    t = LCase$(Trim$(texto))
    t = ReemplazarAcentos(t)
    t = Replace(t, ChrW$(160), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, "_", " ")
    t = Replace(t, ".", " ")
    t = Replace(t, "-", " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    NormalizarEncabezado = t
End Function

Public Function ReemplazarAcentos(ByVal texto As String) As String
    Dim t As String
    t = texto
    t = Replace(t, "á", "a"): t = Replace(t, "é", "e"): t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o"): t = Replace(t, "ú", "u"): t = Replace(t, "ü", "u")
    t = Replace(t, "ñ", "n")
    ReemplazarAcentos = t
End Function

Public Function UltimaFilaConDatos(ByVal ws As Worksheet) As Long
    Dim celda As Range
    Set celda = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    UltimaFilaConDatos = IIf(celda Is Nothing, 1, celda.Row)
End Function

Public Function UltimaColConDatos(ByVal ws As Worksheet) As Long
    Dim celda As Range
    Set celda = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    UltimaColConDatos = IIf(celda Is Nothing, 1, celda.Column)
End Function

Public Function ValorSeguro(ByRef arr As Variant, ByVal fila As Long, ByVal col As Long) As Variant
    If col <= 0 Then ValorSeguro = vbNullString Else ValorSeguro = arr(fila, col)
End Function

Public Function NormalizarCampoClave(ByVal valor As Variant) As String
    Dim t As String
    t = CStr(valor)
    t = Replace(t, ChrW$(160), "")
    t = Replace(t, vbTab, "")
    t = Trim$(t)
    If Len(t) = 0 Then NormalizarCampoClave = "0": Exit Function

    If IsNumeric(t) Then
        If InStr(t, ".") > 0 Or InStr(t, ",") > 0 Then
            If CDbl(Replace(t, ",", ".")) = Fix(CDbl(Replace(t, ",", "."))) Then
                NormalizarCampoClave = CStr(Fix(CDbl(Replace(t, ",", "."))))
            Else
                NormalizarCampoClave = CStr(CDbl(Replace(t, ",", ".")))
            End If
        Else
            NormalizarCampoClave = t
        End If
    Else
        NormalizarCampoClave = UCase$(t)
    End If
End Function

Public Function ConstruirClavePresupuestal(ByVal finac As Variant, ByVal derF As Variant, ByVal pg As Variant, ByVal spg As Variant, ByVal proy As Variant, ByVal rubro As Variant, ByVal rAux As Variant, ByVal ue As Variant, ByVal dep As Variant, ByVal obra As Variant, ByVal derObra As Variant, ByVal serv As Variant, ByVal snip As Variant, Optional ByVal incluirDep As Boolean = True) As String
    ConstruirClavePresupuestal = NormalizarCampoClave(finac) & "|" & _
                                 NormalizarCampoClave(derF) & "|" & _
                                 NormalizarCampoClave(pg) & "|" & _
                                 NormalizarCampoClave(spg) & "|" & _
                                 NormalizarCampoClave(proy) & "|" & _
                                 NormalizarCampoClave(rubro) & "|" & _
                                 NormalizarCampoClave(rAux) & "|" & _
                                 NormalizarCampoClave(ue) & "|" & _
                                 IIf(incluirDep, NormalizarCampoClave(dep) & "|", vbNullString) & _
                                 NormalizarCampoClave(obra) & "|" & _
                                 NormalizarCampoClave(derObra) & "|" & _
                                 NormalizarCampoClave(serv) & "|" & _
                                 NormalizarCampoClave(snip)
End Function

Public Function TryObtenerFecha(ByVal valor As Variant, ByRef fechaOut As Date) As Boolean
    On Error GoTo F
    If IsDate(valor) Then fechaOut = CDate(valor): TryObtenerFecha = True: Exit Function
F:
    TryObtenerFecha = False
End Function

Public Sub EliminarHojaSiExiste(ByVal wb As Workbook, ByVal nombreHoja As String)
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombreHoja, vbTextCompare) = 0 Then ws.Delete: Exit For
    Next ws
    Application.DisplayAlerts = True
End Sub

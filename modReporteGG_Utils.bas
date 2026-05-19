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

Public Function TryObtenerFechaValorSeguro(ByVal v As Variant, ByRef fechaOut As Date) As Boolean
    Dim t As String, n As Double
    Dim yy As Long, mm As Long, dd As Long

    On Error GoTo F

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbNull Then Exit Function

    If IsDate(v) Then
        fechaOut = CDate(v)
        TryObtenerFechaValorSeguro = True
        Exit Function
    End If

    If IsNumeric(v) Then
        n = CDbl(v)
        If n > 0# And n < 2958466# Then
            fechaOut = CDate(CDbl(DateSerial(1899, 12, 30)) + n)
            TryObtenerFechaValorSeguro = True
            Exit Function
        End If
    End If

    t = Trim$(CStr(v))
    If Len(t) = 0 Then Exit Function

    If ParseFechaISO(t, yy, mm, dd) Then
        fechaOut = DateSerial(yy, mm, dd)
        TryObtenerFechaValorSeguro = True
        Exit Function
    End If

    If ParseFechaLatam(t, yy, mm, dd) Then
        fechaOut = DateSerial(yy, mm, dd)
        TryObtenerFechaValorSeguro = True
        Exit Function
    End If

F:
    If Not TryObtenerFechaValorSeguro Then TryObtenerFechaValorSeguro = False
End Function

Public Function TipoVBADeValor(ByVal v As Variant) As String
    If IsObject(v) Then
        TipoVBADeValor = "Object"
    ElseIf IsError(v) Then
        TipoVBADeValor = "Error"
    Else
        TipoVBADeValor = TypeName(v) & " (VarType=" & CStr(VarType(v)) & ")"
    End If
End Function

Private Function ParseFechaISO(ByVal t As String, ByRef yy As Long, ByRef mm As Long, ByRef dd As Long) As Boolean
    Dim p() As String
    p = Split(t, "-")
    If UBound(p) <> 2 Then Exit Function
    If Not (EsEnteroPositivo(p(0)) And EsEnteroPositivo(p(1)) And EsEnteroPositivo(p(2))) Then Exit Function
    yy = CLng(p(0)): mm = CLng(p(1)): dd = CLng(p(2))
    If Not FechaValidaYMD(yy, mm, dd) Then Exit Function
    ParseFechaISO = True
End Function

Private Function ParseFechaLatam(ByVal t As String, ByRef yy As Long, ByRef mm As Long, ByRef dd As Long) As Boolean
    Dim p() As String
    p = Split(t, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (EsEnteroPositivo(p(0)) And EsEnteroPositivo(p(1)) And EsEnteroPositivo(p(2))) Then Exit Function
    dd = CLng(p(0)): mm = CLng(p(1)): yy = CLng(p(2))
    If Not FechaValidaYMD(yy, mm, dd) Then Exit Function
    ParseFechaLatam = True
End Function

Private Function EsEnteroPositivo(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    EsEnteroPositivo = True
End Function

Private Function FechaValidaYMD(ByVal yy As Long, ByVal mm As Long, ByVal dd As Long) As Boolean
    Dim f As Date
    On Error GoTo F
    If yy < 1900 Or yy > 9999 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function
    f = DateSerial(yy, mm, dd)
    If Year(f) = yy And Month(f) = mm And Day(f) = dd Then FechaValidaYMD = True
    Exit Function
F:
End Function

Public Sub EliminarHojaSiExiste(ByVal wb As Workbook, ByVal nombreHoja As String)
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombreHoja, vbTextCompare) = 0 Then ws.Delete: Exit For
    Next ws
    Application.DisplayAlerts = True
End Sub

Attribute VB_Name = "modReporteGG_Utils"
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
    Dim fechaMax As Date
    Dim candidato As String

    On Error GoTo ControlError

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function

    Set folder = fso.GetFolder(carpeta)

    fechaMax = #1/1/1900#
    candidato = vbNullString

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

Public Function ObtenerPrimeraHojaConDatos(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
            Set ObtenerPrimeraHojaConDatos = ws
            Exit Function
        End If
    Next ws
End Function

Public Function UltimaFilaConDatos(ByVal ws As Worksheet) As Long
    Dim celda As Range
    Set celda = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then
        UltimaFilaConDatos = 1
    Else
        UltimaFilaConDatos = celda.Row
    End If
End Function

Public Function UltimaColConDatos(ByVal ws As Worksheet) As Long
    Dim celda As Range
    Set celda = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If celda Is Nothing Then
        UltimaColConDatos = 1
    Else
        UltimaColConDatos = celda.Column
    End If
End Function

Public Function MapearEncabezados(ByRef arrDatos As Variant) As Object
    Dim dict As Object
    Dim col As Long
    Dim nombre As String

    Set dict = CreateObject("Scripting.Dictionary")

    For col = 1 To UBound(arrDatos, 2)
        nombre = NormalizarEncabezado(CStr(arrDatos(1, col)))
        If Len(nombre) > 0 Then
            If Not dict.Exists(nombre) Then
                dict.Add nombre, col
            End If
        End If
    Next col

    Set MapearEncabezados = dict
End Function

Public Function ObtenerColumna(ByVal mapHeaders As Object, ByVal aliases As Variant) As Long
    Dim i As Long, key As String
    For i = LBound(aliases) To UBound(aliases)
        key = NormalizarEncabezado(CStr(aliases(i)))
        If mapHeaders.Exists(key) Then
            ObtenerColumna = CLng(mapHeaders(key))
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 3000, "ObtenerColumna", "Falta columna obligatoria: " & CStr(aliases(LBound(aliases)))
End Function

Public Function ObtenerColumnaOpcional(ByVal mapHeaders As Object, ByVal aliases As Variant) As Long
    Dim i As Long, key As String
    For i = LBound(aliases) To UBound(aliases)
        key = NormalizarEncabezado(CStr(aliases(i)))
        If mapHeaders.Exists(key) Then
            ObtenerColumnaOpcional = CLng(mapHeaders(key))
            Exit Function
        End If
    Next i
    ObtenerColumnaOpcional = 0
End Function

Public Function NormalizarEncabezado(ByVal texto As String) As String
    Dim t As String
    t = LCase$(Trim$(texto))

    t = ReemplazarAcentos(t)
    t = Replace(t, vbTab, " ")
    t = Replace(t, "_", " ")
    t = Replace(t, ".", " ")
    t = Replace(t, "-", " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalizarEncabezado = t
End Function

Public Function ReemplazarAcentos(ByVal texto As String) As String
    Dim t As String
    t = texto
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ü", "u")
    t = Replace(t, "ñ", "n")
    ReemplazarAcentos = t
End Function

Public Function LimpiarTexto(ByVal valor As String) As String
    LimpiarTexto = Trim$(Replace(valor, vbTab, " "))
End Function

Public Function ValorSeguro(ByRef arr As Variant, ByVal fila As Long, ByVal col As Long) As Variant
    If col <= 0 Then
        ValorSeguro = vbNullString
    Else
        ValorSeguro = arr(fila, col)
    End If
End Function

Public Function EsFechaValida(ByVal valor As Variant) As Boolean
    On Error GoTo NoEsFecha
    If IsDate(valor) Then
        EsFechaValida = True
        Exit Function
    End If
NoEsFecha:
    EsFechaValida = False
End Function

Public Function EsNumeroValido(ByVal valor As Variant) As Boolean
    On Error GoTo NoEsNumero
    If Len(Trim$(CStr(valor))) = 0 Then GoTo NoEsNumero
    If IsNumeric(valor) Then
        EsNumeroValido = True
        Exit Function
    End If
NoEsNumero:
    EsNumeroValido = False
End Function

Public Function ConstruirLlavePresupuestal(ByVal finac As Variant, ByVal derF As Variant, ByVal pg As Variant, ByVal spg As Variant, _
                                           ByVal proy As Variant, ByVal rubro As Variant, ByVal rAux As Variant, ByVal ue As Variant, _
                                           ByVal dep As Variant, ByVal obra As Variant, ByVal derObra As Variant, ByVal serv As Variant, _
                                           ByVal sniip As Variant) As String

    ConstruirLlavePresupuestal = "F=" & NormalizarCodigo(finac) & _
                                 "|DF=" & NormalizarCodigo(derF) & _
                                 "|PG=" & NormalizarCodigo(pg) & _
                                 "|SPG=" & NormalizarCodigo(spg) & _
                                 "|PROY=" & NormalizarCodigo(proy) & _
                                 "|R=" & NormalizarCodigo(rubro) & _
                                 "|RA=" & NormalizarCodigo(rAux) & _
                                 "|UE=" & NormalizarCodigo(ue) & _
                                 "|DEP=" & NormalizarCodigo(dep) & _
                                 "|OB=" & NormalizarCodigo(obra) & _
                                 "|DOB=" & NormalizarCodigo(derObra) & _
                                 "|SV=" & NormalizarCodigo(serv) & _
                                 "|SN=" & NormalizarCodigo(sniip)
End Function

Public Function NormalizarCodigo(ByVal valor As Variant) As String
    Dim txt As String

    txt = Trim$(CStr(valor))
    If Len(txt) = 0 Then
        NormalizarCodigo = "0"
        Exit Function
    End If

    txt = Replace(txt, ",", ".")

    If IsNumeric(txt) Then
        NormalizarCodigo = CStr(CDec(txt))
    Else
        NormalizarCodigo = UCase$(txt)
    End If
End Function

Public Function ConstruirClaveCombo(ByVal nivel1 As String, ByVal nivel2 As String, ByVal subtipo As String) As String
    ConstruirClaveCombo = Trim$(nivel1) & "|" & Trim$(nivel2) & "|" & Trim$(subtipo)
End Function

Public Function ElegirRubroCodiguera(ByRef arr As Variant, ByVal fila As Long, ByVal colRubroNum As Long, ByVal colRubro As Long) As Variant
    Dim vNum As Variant
    vNum = ValorSeguro(arr, fila, colRubroNum)

    If colRubroNum > 0 And Len(Trim$(CStr(vNum))) > 0 Then
        ElegirRubroCodiguera = vNum
    Else
        ElegirRubroCodiguera = ValorSeguro(arr, fila, colRubro)
    End If
End Function

Public Function ElegirRAuxCodiguera(ByRef arr As Variant, ByVal fila As Long, ByVal colRAuxNum As Long, ByVal colRAux As Long) As Variant
    Dim vNum As Variant
    vNum = ValorSeguro(arr, fila, colRAuxNum)

    If colRAuxNum > 0 And Len(Trim$(CStr(vNum))) > 0 Then
        ElegirRAuxCodiguera = vNum
    Else
        ElegirRAuxCodiguera = ValorSeguro(arr, fila, colRAux)
    End If
End Function

Public Sub AcumularImporte(ByRef dictAcumulado As Object, ByVal comboKey As String, ByVal mes As Long, ByVal importe As Double)
    Dim arrMeses As Variant

    If mes < 1 Or mes > 12 Then Exit Sub

    If dictAcumulado.Exists(comboKey) Then
        arrMeses = dictAcumulado(comboKey)
    Else
        arrMeses = InicializarArregloMeses()
    End If

    arrMeses(mes) = CDbl(arrMeses(mes)) + importe
    dictAcumulado(comboKey) = arrMeses
End Sub

Public Function InicializarArregloMeses() As Variant
    Dim arr(1 To 12) As Double
    InicializarArregloMeses = arr
End Function

Public Sub EliminarHojaSiExiste(ByVal wb As Workbook, ByVal nombreHoja As String)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombreHoja, vbTextCompare) = 0 Then
            ws.Delete
            Exit For
        End If
    Next ws
End Sub

Public Function NombreArchivoDesdeRuta(ByVal ruta As String) As String
    NombreArchivoDesdeRuta = Mid$(ruta, InStrRev(ruta, "\") + 1)
End Function

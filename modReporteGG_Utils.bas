Option Explicit

Public Const MOSTRAR_EN_MILLONES As Boolean = True
Public Const MOSTRAR_EN_MILES As Boolean = False
Public Const PANEL_SHEET_NAME As String = "Panel Reportes"
Public Const DIAG_SHEET_NAME As String = "Diagnostico_Llaves"
Public Const CAMPO_FINANCIAMIENTO_CODIGUERA As String = "Titular"

Public Function FactorEscalaImporte() As Double
    If MOSTRAR_EN_MILLONES Then
        FactorEscalaImporte = 1000000#
    ElseIf MOSTRAR_EN_MILES Then
        FactorEscalaImporte = 1000#
    Else
        FactorEscalaImporte = 1#
    End If
End Function

Public Function SufijoUnidadTitulo() As String
    If MOSTRAR_EN_MILLONES Then
        SufijoUnidadTitulo = " (en millones de $)"
    ElseIf MOSTRAR_EN_MILES Then
        SufijoUnidadTitulo = " (en miles de $)"
    Else
        SufijoUnidadTitulo = ""
    End If
End Function

Public Function MesesES() As Variant
    MesesES = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                    "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Public Function MesesESMin() As Variant
    MesesESMin = Array("enero", "febrero", "marzo", "abril", "mayo", "junio", _
                       "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre")
End Function

Public Function MesTextoANumero(ByVal mesTexto As String) As Long
    Dim i As Long, meses As Variant
    meses = MesesES()
    For i = LBound(meses) To UBound(meses)
        If StrComp(LimpiarTexto(CStr(meses(i))), LimpiarTexto(mesTexto), vbTextCompare) = 0 Then MesTextoANumero = i + 1: Exit Function
    Next i
End Function

Public Function LimpiarTexto(ByVal valor As String) As String
    Dim t As String
    t = Replace(valor, ChrW$(160), " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Application.WorksheetFunction.Trim(t)
    LimpiarTexto = t
End Function

Public Function NormalizarCampoClave(ByVal valor As Variant) As String
    Dim texto As String, numero As Double
    texto = Replace(Replace(LimpiarTexto(CStr(valor)), vbTab, ""), " ", "")
    If Len(texto) = 0 Then NormalizarCampoClave = "0": Exit Function
    If IsNumeric(texto) Then
        numero = CDbl(Replace(texto, ",", "."))
        If numero = Fix(numero) Then NormalizarCampoClave = CStr(Fix(numero)) Else NormalizarCampoClave = CStr(numero)
    Else
        NormalizarCampoClave = UCase$(texto)
    End If
End Function

Public Function ConstruirClavePresupuestal(ByVal finac As Variant, ByVal derF As Variant, ByVal pg As Variant, ByVal spg As Variant, ByVal proy As Variant, ByVal rubro As Variant, ByVal rAux As Variant, ByVal ue As Variant, ByVal dep As Variant, ByVal obra As Variant, ByVal derObra As Variant, ByVal serv As Variant, ByVal snip As Variant) As String
    ConstruirClavePresupuestal = Join(Array(NormalizarCampoClave(finac), NormalizarCampoClave(derF), NormalizarCampoClave(pg), NormalizarCampoClave(spg), NormalizarCampoClave(proy), NormalizarCampoClave(rubro), NormalizarCampoClave(rAux), NormalizarCampoClave(ue), NormalizarCampoClave(dep), NormalizarCampoClave(obra), NormalizarCampoClave(derObra), NormalizarCampoClave(serv), NormalizarCampoClave(snip)), "|")
End Function

Public Function TryObtenerFechaValorSeguro(ByVal valorFuente As Variant, ByRef fechaOut As Date) As Boolean
    Dim texto As String, partes() As String, numero As Double
    On Error GoTo EH
    If IsDate(valorFuente) Then fechaOut = CDate(valorFuente): TryObtenerFechaValorSeguro = True: Exit Function
    If IsNumeric(valorFuente) Then
        numero = CDbl(valorFuente)
        If numero > 0 And numero < 2958466 Then fechaOut = CDate(DateSerial(1899, 12, 30) + numero): TryObtenerFechaValorSeguro = True: Exit Function
    End If
    texto = LimpiarTexto(CStr(valorFuente))
    If InStr(1, texto, "-") > 0 Then partes = Split(texto, "-"): If UBound(partes) = 2 Then fechaOut = DateSerial(CLng(partes(0)), CLng(partes(1)), CLng(partes(2))): TryObtenerFechaValorSeguro = True: Exit Function
    If InStr(1, texto, "/") > 0 Then partes = Split(texto, "/"): If UBound(partes) = 2 Then fechaOut = DateSerial(CLng(partes(2)), CLng(partes(1)), CLng(partes(0))): TryObtenerFechaValorSeguro = True: Exit Function
EH:
End Function

Public Function MapearEncabezados(ByRef matriz As Variant) As Object
    Dim d As Object, col As Long, h As String
    Set d = CreateObject("Scripting.Dictionary")
    For col = 1 To UBound(matriz, 2)
        h = LCase$(LimpiarTexto(CStr(matriz(1, col))))
        matriz(1, col) = LimpiarTexto(CStr(matriz(1, col)))
        If Len(h) > 0 And Not d.Exists(h) Then d.Add h, col
    Next col
    Set MapearEncabezados = d
End Function

Public Function ObtenerColumna(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long, a As String
    For i = LBound(aliases) To UBound(aliases)
        a = LCase$(LimpiarTexto(CStr(aliases(i))))
        If headers.Exists(a) Then ObtenerColumna = CLng(headers(a)): Exit Function
    Next i
    Err.Raise vbObjectError + 513, , "No se encontró columna obligatoria: " & Join(aliases, ", ")
End Function

Public Function UltimaFilaConDatos(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then UltimaFilaConDatos = 1 Else UltimaFilaConDatos = c.Row
End Function

Public Function UltimaColConDatos(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then UltimaColConDatos = 1 Else UltimaColConDatos = c.Column
End Function

Public Function HojaExiste(ByVal wb As Workbook, ByVal nombreHoja As String) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    HojaExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function ObtenerUltimaFilaSegura(ByVal wb As Workbook, ByVal nombreHoja As String) As Long
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    If Not HojaExiste(wb, nombreHoja) Then Exit Function
    Set ws = wb.Worksheets(nombreHoja)
    ObtenerUltimaFilaSegura = UltimaFilaConDatos(ws)
End Function

Public Function ObtenerUltimaColSegura(ByVal wb As Workbook, ByVal nombreHoja As String) As Long
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    If Not HojaExiste(wb, nombreHoja) Then Exit Function
    Set ws = wb.Worksheets(nombreHoja)
    ObtenerUltimaColSegura = UltimaColConDatos(ws)
End Function

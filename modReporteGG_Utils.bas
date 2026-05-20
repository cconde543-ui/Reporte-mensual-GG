Option Explicit

Public Const MOSTRAR_EN_MILES As Boolean = True
Public Const PANEL_SHEET_NAME As String = "Panel Reportes"
Public Const DIAG_SHEET_NAME As String = "Diagnostico_Llaves"

Public Function MesesES() As Variant
    MesesES = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Public Function MesTextoANumero(ByVal mesTexto As String) As Long
    Dim i As Long, m As Variant
    m = MesesES()
    For i = LBound(m) To UBound(m)
        If StrComp(Trim$(mesTexto), CStr(m(i)), vbTextCompare) = 0 Then MesTextoANumero = i + 1: Exit Function
    Next i
End Function

Public Function NormalizarCampoClave(ByVal valor As Variant) As String
    Dim t As String, d As Double
    t = Replace(Replace(Trim$(CStr(valor)), ChrW$(160), ""), vbTab, "")
    If Len(t) = 0 Then NormalizarCampoClave = "0": Exit Function
    If IsNumeric(t) Then
        d = CDbl(Replace(t, ",", "."))
        If d = Fix(d) Then NormalizarCampoClave = CStr(Fix(d)) Else NormalizarCampoClave = CStr(d)
    Else
        NormalizarCampoClave = UCase$(t)
    End If
End Function

Public Function ConstruirClavePresupuestal(ByVal finac As Variant, ByVal derF As Variant, ByVal pg As Variant, ByVal spg As Variant, ByVal proy As Variant, ByVal rubro As Variant, ByVal rAux As Variant, ByVal ue As Variant, ByVal dep As Variant, ByVal obra As Variant, ByVal derObra As Variant, ByVal serv As Variant, ByVal snip As Variant) As String
    ConstruirClavePresupuestal = Join(Array(NormalizarCampoClave(finac), NormalizarCampoClave(derF), NormalizarCampoClave(pg), NormalizarCampoClave(spg), NormalizarCampoClave(proy), NormalizarCampoClave(rubro), NormalizarCampoClave(rAux), NormalizarCampoClave(ue), NormalizarCampoClave(dep), NormalizarCampoClave(obra), NormalizarCampoClave(derObra), NormalizarCampoClave(serv), NormalizarCampoClave(snip)), "|")
End Function

Public Function TryObtenerFechaValorSeguro(ByVal v As Variant, ByRef fechaOut As Date) As Boolean
    Dim t As String, p() As String, n As Double
    On Error GoTo EH
    If IsDate(v) Then fechaOut = CDate(v): TryObtenerFechaValorSeguro = True: Exit Function
    If IsNumeric(v) Then
        n = CDbl(v)
        If n > 0 And n < 2958466 Then fechaOut = CDate(DateSerial(1899, 12, 30) + n): TryObtenerFechaValorSeguro = True: Exit Function
    End If
    t = Trim$(CStr(v))
    If InStr(1, t, "-") > 0 Then
        p = Split(t, "-")
        If UBound(p) = 2 Then fechaOut = DateSerial(CLng(p(0)), CLng(p(1)), CLng(p(2))): TryObtenerFechaValorSeguro = True: Exit Function
    End If
    If InStr(1, t, "/") > 0 Then
        p = Split(t, "/")
        If UBound(p) = 2 Then fechaOut = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0))): TryObtenerFechaValorSeguro = True: Exit Function
    End If
EH:
End Function

Public Function MapearEncabezados(ByRef arr As Variant) As Object
    Dim d As Object, c As Long
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(arr, 2)
        If Not d.Exists(LCase$(Trim$(CStr(arr(1, c))))) Then d.Add LCase$(Trim$(CStr(arr(1, c)))), c
    Next c
    Set MapearEncabezados = d
End Function

Public Function ObtenerColumna(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long, k As String
    For i = LBound(aliases) To UBound(aliases)
        k = LCase$(Trim$(CStr(aliases(i))))
        If headers.Exists(k) Then ObtenerColumna = CLng(headers(k)): Exit Function
    Next i
    Err.Raise vbObjectError + 513, , "No se encontró columna obligatoria."
End Function

Public Function UltimaFilaConDatos(ByVal ws As Worksheet) As Long
    UltimaFilaConDatos = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Public Function UltimaColConDatos(ByVal ws As Worksheet) As Long
    UltimaColConDatos = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

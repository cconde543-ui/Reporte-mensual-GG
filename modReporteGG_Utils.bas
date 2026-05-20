Option Explicit

Public Const MOSTRAR_EN_MILES As Boolean = True
Public Const PANEL_SHEET_NAME As String = "Panel Reportes"
Public Const DIAG_SHEET_NAME As String = "Diagnostico_Llaves"

Public Function MesesES() As Variant
    MesesES = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                    "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Public Function MesTextoANumero(ByVal mesTexto As String) As Long
    Dim i As Long
    Dim meses As Variant

    meses = MesesES()
    For i = LBound(meses) To UBound(meses)
        If StrComp(Trim$(mesTexto), CStr(meses(i)), vbTextCompare) = 0 Then
            MesTextoANumero = i + 1
            Exit Function
        End If
    Next i

    MesTextoANumero = 0
End Function

Public Function NormalizarCampoClave(ByVal valor As Variant) As String
    Dim texto As String
    Dim numero As Double

    texto = Replace(Replace(Trim$(CStr(valor)), ChrW$(160), ""), vbTab, "")
    If Len(texto) = 0 Then
        NormalizarCampoClave = "0"
        Exit Function
    End If

    If IsNumeric(texto) Then
        numero = CDbl(Replace(texto, ",", "."))
        If numero = Fix(numero) Then
            NormalizarCampoClave = CStr(Fix(numero))
        Else
            NormalizarCampoClave = CStr(numero)
        End If
    Else
        NormalizarCampoClave = UCase$(texto)
    End If
End Function

Public Function ConstruirClavePresupuestal( _
    ByVal finac As Variant, _
    ByVal derF As Variant, _
    ByVal pg As Variant, _
    ByVal spg As Variant, _
    ByVal proy As Variant, _
    ByVal rubro As Variant, _
    ByVal rAux As Variant, _
    ByVal ue As Variant, _
    ByVal dep As Variant, _
    ByVal obra As Variant, _
    ByVal derObra As Variant, _
    ByVal serv As Variant, _
    ByVal snip As Variant) As String

    ConstruirClavePresupuestal = Join(Array( _
        NormalizarCampoClave(finac), _
        NormalizarCampoClave(derF), _
        NormalizarCampoClave(pg), _
        NormalizarCampoClave(spg), _
        NormalizarCampoClave(proy), _
        NormalizarCampoClave(rubro), _
        NormalizarCampoClave(rAux), _
        NormalizarCampoClave(ue), _
        NormalizarCampoClave(dep), _
        NormalizarCampoClave(obra), _
        NormalizarCampoClave(derObra), _
        NormalizarCampoClave(serv), _
        NormalizarCampoClave(snip)), "|")
End Function

Public Function TryObtenerFechaValorSeguro(ByVal valorFuente As Variant, ByRef fechaOut As Date) As Boolean
    Dim texto As String
    Dim partes() As String
    Dim numero As Double

    On Error GoTo EH

    If IsDate(valorFuente) Then
        fechaOut = CDate(valorFuente)
        TryObtenerFechaValorSeguro = True
        Exit Function
    End If

    If IsNumeric(valorFuente) Then
        numero = CDbl(valorFuente)
        If numero > 0 And numero < 2958466 Then
            fechaOut = CDate(DateSerial(1899, 12, 30) + numero)
            TryObtenerFechaValorSeguro = True
            Exit Function
        End If
    End If

    texto = Trim$(CStr(valorFuente))

    If InStr(1, texto, "-") > 0 Then
        partes = Split(texto, "-")
        If UBound(partes) = 2 Then
            fechaOut = DateSerial(CLng(partes(0)), CLng(partes(1)), CLng(partes(2)))
            TryObtenerFechaValorSeguro = True
            Exit Function
        End If
    End If

    If InStr(1, texto, "/") > 0 Then
        partes = Split(texto, "/")
        If UBound(partes) = 2 Then
            fechaOut = DateSerial(CLng(partes(2)), CLng(partes(1)), CLng(partes(0)))
            TryObtenerFechaValorSeguro = True
            Exit Function
        End If
    End If

EH:
End Function

Public Function MapearEncabezados(ByRef matriz As Variant) As Object
    Dim diccHeaders As Object
    Dim col As Long
    Dim headerName As String

    Set diccHeaders = CreateObject("Scripting.Dictionary")
    For col = 1 To UBound(matriz, 2)
        headerName = LCase$(Trim$(CStr(matriz(1, col))))
        If Not diccHeaders.Exists(headerName) Then
            diccHeaders.Add headerName, col
        End If
    Next col

    Set MapearEncabezados = diccHeaders
End Function

Public Function ObtenerColumna(ByVal headers As Object, ByVal aliases As Variant) As Long
    Dim i As Long
    Dim aliasName As String

    For i = LBound(aliases) To UBound(aliases)
        aliasName = LCase$(Trim$(CStr(aliases(i))))
        If headers.Exists(aliasName) Then
            ObtenerColumna = CLng(headers(aliasName))
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 513, , "No se encontró columna obligatoria."
End Function

Public Function ObtenerColumnaOpcional(ByVal headers As Object, ByVal aliases As Variant, ByVal defaultCol As Long) As Long
    Dim i As Long
    Dim aliasName As String

    For i = LBound(aliases) To UBound(aliases)
        aliasName = LCase$(Trim$(CStr(aliases(i))))
        If headers.Exists(aliasName) Then
            ObtenerColumnaOpcional = CLng(headers(aliasName))
            Exit Function
        End If
    Next i

    ObtenerColumnaOpcional = defaultCol
End Function

Public Function UltimaFilaConDatos(ByVal ws As Worksheet) As Long
    UltimaFilaConDatos = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Public Function UltimaColConDatos(ByVal ws As Worksheet) As Long
    UltimaColConDatos = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

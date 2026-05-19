Option Explicit

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictLlaveACombo As Object, ByRef dictCombos As Object, ByRef diag As Object)
    Dim arr As Variant, h As Object, i As Long, key As String, keySinDep As String, comboKey As String
    Dim cIncluir As Long, cN1 As Long, cN2 As Long, cN3 As Long
    Dim cFin As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long, cRub As Long, cRA As Long
    Dim cUE As Long, cDep As Long, cOb As Long, cDOb As Long, cSrv As Long, cSn As Long
    Dim cClave As Long, cClaveSinDep As Long

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)

    cIncluir = ObtenerColumna(h, Array("Incluir_en_Informe"))
    cN1 = ObtenerColumna(h, Array("Nivel_1")): cN2 = ObtenerColumna(h, Array("Nivel_2")): cN3 = ObtenerColumna(h, Array("Nivel_3"))
    cFin = ObtenerColumna(h, Array("Finac Código numérico", "Finac"))
    cDerF = ObtenerColumna(h, Array("Der-F Código numérico", "Der-F"))
    cPG = ObtenerColumna(h, Array("PG Código numérico", "PG"))
    cSpg = ObtenerColumna(h, Array("Spg Código numérico", "Spg"))
    cProy = ObtenerColumna(h, Array("Proy"))
    cRub = ObtenerColumna(h, Array("Rubro Código numérico", "Rubro"))
    cRA = ObtenerColumna(h, Array("R. Aux Código numérico", "R. Aux"))
    cUE = ObtenerColumna(h, Array("UE Código numérico", "UE"))
    cDep = ObtenerColumna(h, Array("Dep Código numérico", "Dep"))
    cOb = ObtenerColumna(h, Array("Obra Código numérico", "Obra"))
    cDOb = ObtenerColumna(h, Array("Der. Obra Código numérico", "Der. Obra"))
    cSrv = ObtenerColumna(h, Array("Serv Código numérico", "Serv"))
    cSn = ObtenerColumna(h, Array("SNIP Código numérico"))
    cClave = ObtenerColumnaOpcional(h, Array("Clave Llave presupuestal"))
    cClaveSinDep = ObtenerColumnaOpcional(h, Array("Clave sin dep"))

    diag("cod_total") = UBound(arr, 1) - 1

    For i = 2 To UBound(arr, 1)
        If UCase$(NormalizarCampoClave(ValorSeguro(arr, i, cIncluir))) = "SI" Then
            diag("cod_si") = CLng(diag("cod_si")) + 1
            comboKey = CStr(ValorSeguro(arr, i, cN1)) & "|" & CStr(ValorSeguro(arr, i, cN2)) & "|" & CStr(ValorSeguro(arr, i, cN3))
            If Not dictCombos.Exists(comboKey) Then dictCombos.Add comboKey, Array(ValorSeguro(arr, i, cN1), ValorSeguro(arr, i, cN2), ValorSeguro(arr, i, cN3))

            key = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), True)
            keySinDep = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), False)

            If Not dictLlaveACombo.Exists(key) Then dictLlaveACombo.Add key, comboKey
            If Not diag("cod_set").Exists(key) Then diag("cod_set").Add key, True
            If Not diag("cod_set_sindep").Exists(keySinDep) Then diag("cod_set_sindep").Add keySinDep, True
            If diag("cod_keys").Count < 30 Then diag("cod_keys").Add key
            If cClave > 0 Then If Not diag("cod_set_clave").Exists(NormalizarCampoClave(ValorSeguro(arr, i, cClave))) Then diag("cod_set_clave").Add NormalizarCampoClave(ValorSeguro(arr, i, cClave)), True
            If cClaveSinDep > 0 Then If Not diag("cod_set_clavesindep").Exists(NormalizarCampoClave(ValorSeguro(arr, i, cClaveSinDep))) Then diag("cod_set_clavesindep").Add NormalizarCampoClave(ValorSeguro(arr, i, cClaveSinDep)), True

            If diag("cod_rows").Count < 30 Then diag("cod_rows").Add Array(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), ValorSeguro(arr, i, cN1), ValorSeguro(arr, i, cN2), ValorSeguro(arr, i, cN3), key)
        End If
    Next i
End Sub

Public Sub LeerEjecucionesYAcumular(ByVal ws As Worksheet, ByVal anioReporte As Long, ByVal dictLlaveACombo As Object, ByRef dictAcumulado As Object, ByRef diag As Object)
    Dim arr As Variant, h As Object, i As Long, key As String, keySinDep As String, combo As String
    Dim cFin As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long, cRub As Long, cRA As Long
    Dim cUE As Long, cDep As Long, cOb As Long, cDOb As Long, cSrv As Long, cSn As Long, cFv As Long, cImp As Long
    Dim fv As Date, importeMN As Variant

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)

    cFin = ObtenerColumna(h, Array("Finac Código numérico")): cDerF = ObtenerColumna(h, Array("Der-F Código numérico"))
    cPG = ObtenerColumna(h, Array("PG Código numérico")): cSpg = ObtenerColumna(h, Array("Spg Código numérico"))
    cProy = ObtenerColumna(h, Array("Proyecto")): cRub = ObtenerColumna(h, Array("Rubro Código numérico"))
    cRA = ObtenerColumna(h, Array("R. Aux Código numérico")): cUE = ObtenerColumna(h, Array("UE Código numérico"))
    cDep = ObtenerColumna(h, Array("Dep Código numérico")): cOb = ObtenerColumna(h, Array("Obra Código numérico"))
    cDOb = ObtenerColumna(h, Array("Der. Obra Código numérico")): cSrv = ObtenerColumna(h, Array("Serv Código numérico"))
    cSn = ObtenerColumna(h, Array("SNIP Código numérico")): cFv = ObtenerColumna(h, Array("Fecha valor"))
    cImp = ObtenerColumna(h, Array("Importe moneda nacional"))

    diag("ej_total") = UBound(arr, 1) - 1
    For i = 2 To UBound(arr, 1)
        If Not TryObtenerFecha(ValorSeguro(arr, i, cFv), fv) Then diag("ej_fecha_invalida") = CLng(diag("ej_fecha_invalida")) + 1: GoTo S
        If Year(fv) <> anioReporte Then GoTo S
        diag("ej_anio") = CLng(diag("ej_anio")) + 1

        importeMN = ValorSeguro(arr, i, cImp)
        If Not IsNumeric(importeMN) Then GoTo S
        diag("ej_importe_num") = CLng(diag("ej_importe_num")) + 1

        key = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), True)
        keySinDep = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), False)

        If Not diag("ej_set").Exists(key) Then diag("ej_set").Add key, True
        If Not diag("ej_set_sindep").Exists(keySinDep) Then diag("ej_set_sindep").Add keySinDep, True
        If diag("ej_keys").Count < 30 Then diag("ej_keys").Add key

        If diag("ej_rows").Count < 30 Then diag("ej_rows").Add Array(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), ValorSeguro(arr, i, cFv), ValorSeguro(arr, i, cImp), key, IIf(dictLlaveACombo.Exists(key), "SI", "NO"))

        If dictLlaveACombo.Exists(key) Then
            combo = dictLlaveACombo(key)
            If Not dictAcumulado.Exists(combo) Then dictAcumulado.Add combo, 0#
            dictAcumulado(combo) = CDbl(dictAcumulado(combo)) + CDbl(importeMN)
        End If
S:
    Next i
End Sub

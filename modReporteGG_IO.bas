Option Explicit

Public Sub LeerCodiguera(ByVal ws As Worksheet, ByRef dictLlaveACombo As Object, ByRef dictCombos As Object, ByRef diag As Object)
    Dim arr As Variant, h As Object, i As Long, key As String, keySinDep As String
    Dim cIncluir As Long, cN1 As Long, cN2 As Long, cN3 As Long
    Dim cFin As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long, cRub As Long, cRA As Long
    Dim cUE As Long, cDep As Long, cOb As Long, cDOb As Long, cSrv As Long, cSn As Long
    Dim cClave As Long, cClaveSinDep As Long, comboKey As String

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)

    Debug.Print "[DEBUG][Codiguera] Archivo abierto: " & ws.Parent.FullName
    Debug.Print "[DEBUG][Codiguera] Hoja usada: " & ws.Name
    Debug.Print "[DEBUG][Codiguera] Encabezados detectados: " & Join(h.Keys, " | ")

    cIncluir = ObtenerColumna(h, Array("Incluir_en_Informe"))
    cN1 = ObtenerColumna(h, Array("Nivel_1")): cN2 = ObtenerColumna(h, Array("Nivel_2")): cN3 = ObtenerColumna(h, Array("Nivel_3"))
    cFin = ObtenerColumna(h, Array("Finac Código numérico", "Finac"))
    cDerF = ObtenerColumna(h, Array("Der-F Código numérico", "Der-F"))
    cPG = ObtenerColumna(h, Array("PG Código numérico", "PG"))
    cSpg = ObtenerColumna(h, Array("Spg Código numérico", "Spg"))
    cProy = ObtenerColumna(h, Array("Proy", "Proyecto"))
    cRub = ObtenerColumna(h, Array("Rubro Código numérico", "Rubro"))
    cRA = ObtenerColumna(h, Array("R. Aux Código numérico", "R. Aux"))
    cUE = ObtenerColumna(h, Array("UE Código numérico", "UE"))
    cDep = ObtenerColumna(h, Array("Dep Código numérico", "Dep"))
    cOb = ObtenerColumna(h, Array("Obra Código numérico", "Obra"))
    cDOb = ObtenerColumna(h, Array("Der. Obra Código numérico", "Der. Obra"))
    cSrv = ObtenerColumna(h, Array("Serv Código numérico", "Serv"))
    cSn = ObtenerColumna(h, Array("SNIP Código numérico", "SNIIP", "SNIP"))
    cClave = ObtenerColumnaOpcional(h, Array("Clave Llave presupuestal"))
    cClaveSinDep = ObtenerColumnaOpcional(h, Array("Clave sin dep"))

    diag("cod_total") = UBound(arr, 1) - 1
    diag("cod_si") = 0

    For i = 2 To UBound(arr, 1)
        If UCase$(NormalizarCampoLlave(ValorSeguro(arr, i, cIncluir))) = "SI" Then
            diag("cod_si") = diag("cod_si") + 1
            comboKey = CStr(ValorSeguro(arr, i, cN1)) & "|" & CStr(ValorSeguro(arr, i, cN2)) & "|" & CStr(ValorSeguro(arr, i, cN3))
            If Not dictCombos.Exists(comboKey) Then dictCombos.Add comboKey, Array(ValorSeguro(arr, i, cN1), ValorSeguro(arr, i, cN2), ValorSeguro(arr, i, cN3))

            key = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), True)
            keySinDep = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), False)
            If Not dictLlaveACombo.Exists(key) Then dictLlaveACombo.Add key, comboKey
            If diag("cod_keys").Count < 30 Then diag("cod_keys").Add key
            If Not diag("cod_set").Exists(key) Then diag("cod_set").Add key, True
            If Not diag("cod_set_sindep").Exists(keySinDep) Then diag("cod_set_sindep").Add keySinDep, True
            If cClave > 0 Then If Not diag("cod_set_clave").Exists(NormalizarCampoLlave(ValorSeguro(arr, i, cClave))) Then diag("cod_set_clave").Add NormalizarCampoLlave(ValorSeguro(arr, i, cClave)), True
            If cClaveSinDep > 0 Then If Not diag("cod_set_clavesindep").Exists(NormalizarCampoLlave(ValorSeguro(arr, i, cClaveSinDep))) Then diag("cod_set_clavesindep").Add NormalizarCampoLlave(ValorSeguro(arr, i, cClaveSinDep)), True
        End If
    Next i
    Debug.Print "[DEBUG] Ejemplo llave codiguera: " & IIf(diag("cod_keys").Count > 0, diag("cod_keys")(1), "(sin llaves)")
End Sub

Public Sub LeerEjecucionesYAcumular(ByVal ws As Worksheet, ByVal anioReporte As Long, ByVal dictLlaveACombo As Object, ByRef dictAcumulado As Object, ByRef diag As Object)
    Dim arr As Variant, h As Object, i As Long, key As String, keySinDep As String, combo As String
    Dim cFin As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long, cRub As Long, cRA As Long
    Dim cUE As Long, cDep As Long, cOb As Long, cDOb As Long, cSrv As Long, cSn As Long, cFv As Long, cImp As Long
    Dim fv As Date, imp As Variant

    arr = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFilaConDatos(ws), UltimaColConDatos(ws))).Value2
    Set h = MapearEncabezados(arr)
    Debug.Print "[DEBUG][Ejec] Archivo abierto: " & ws.Parent.FullName
    Debug.Print "[DEBUG][Ejec] Hoja usada: " & ws.Name
    Debug.Print "[DEBUG][Ejec] Encabezados detectados: " & Join(h.Keys, " | ")

    cFin = ObtenerColumna(h, Array("Finac Código numérico", "Finac")): cDerF = ObtenerColumna(h, Array("Der-F Código numérico", "Der-F"))
    cPG = ObtenerColumna(h, Array("PG Código numérico", "PG")): cSpg = ObtenerColumna(h, Array("Spg Código numérico", "Spg"))
    cProy = ObtenerColumna(h, Array("Proyecto", "Proy")): cRub = ObtenerColumna(h, Array("Rubro Código numérico", "Rubro"))
    cRA = ObtenerColumna(h, Array("R. Aux Código numérico", "R. Aux")): cUE = ObtenerColumna(h, Array("UE Código numérico", "UE"))
    cDep = ObtenerColumna(h, Array("Dep Código numérico", "Dep")): cOb = ObtenerColumna(h, Array("Obra Código numérico", "Obra"))
    cDOb = ObtenerColumna(h, Array("Der. Obra Código numérico", "Der. Obra")): cSrv = ObtenerColumna(h, Array("Serv Código numérico", "Serv"))
    cSn = ObtenerColumna(h, Array("SNIP Código numérico", "SNIIP", "SNIP")): cFv = ObtenerColumna(h, Array("Fecha valor"))
    cImp = ObtenerColumna(h, Array("Importe moneda nacional"))

    diag("ej_total") = UBound(arr, 1) - 1
    For i = 2 To UBound(arr, 1)
        If Not TryObtenerFecha(ValorSeguro(arr, i, cFv), fv) Then
            diag("ej_fecha_invalida") = diag("ej_fecha_invalida") + 1
            GoTo S
        End If
        If Year(fv) <> anioReporte Then GoTo S
        diag("ej_2026") = diag("ej_2026") + 1
        imp = ValorSeguro(arr, i, cImp)
        If Not IsNumeric(imp) Then GoTo S
        diag("ej_importe_num") = diag("ej_importe_num") + 1

        key = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), True)
        keySinDep = ConstruirClavePresupuestal(ValorSeguro(arr, i, cFin), ValorSeguro(arr, i, cDerF), ValorSeguro(arr, i, cPG), ValorSeguro(arr, i, cSpg), ValorSeguro(arr, i, cProy), ValorSeguro(arr, i, cRub), ValorSeguro(arr, i, cRA), ValorSeguro(arr, i, cUE), ValorSeguro(arr, i, cDep), ValorSeguro(arr, i, cOb), ValorSeguro(arr, i, cDOb), ValorSeguro(arr, i, cSrv), ValorSeguro(arr, i, cSn), False)
        If diag("ej_keys").Count < 30 Then diag("ej_keys").Add key
        If Not diag("ej_set").Exists(key) Then diag("ej_set").Add key, True
        If Not diag("ej_set_sindep").Exists(keySinDep) Then diag("ej_set_sindep").Add keySinDep, True
        If dictLlaveACombo.Exists(key) Then
            combo = dictLlaveACombo(key)
            If Not dictAcumulado.Exists(combo) Then dictAcumulado.Add combo, 0#
            dictAcumulado(combo) = CDbl(dictAcumulado(combo)) + CDbl(imp)
        End If
S:
    Next i
    Debug.Print "[DEBUG] Ejemplo llave ejecución: " & IIf(diag("ej_keys").Count > 0, diag("ej_keys")(1), "(sin llaves)")
End Sub

Attribute VB_Name = "modReporteGG_IO"
Option Explicit

Public Sub LeerCodiguera(ByVal ws As Worksheet, _
                         ByRef dictLlaveACombo As Object, _
                         ByRef dictCombos As Object, _
                         ByRef llavesValidas As Long)

    Dim arrDatos As Variant
    Dim mapHeaders As Object
    Dim ultimaFila As Long, ultimaCol As Long
    Dim i As Long

    Dim cIncluir As Long, cNivel1 As Long, cNivel2 As Long, cSubtipo As Long
    Dim cFinac As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long
    Dim cRubroNum As Long, cRubro As Long, cRAuxNum As Long, cRAux As Long
    Dim cUE As Long, cDep As Long, cObra As Long, cDerObra As Long, cServ As Long, cSniip As Long

    Dim incluye As String
    Dim comboKey As String
    Dim llave As String

    ultimaFila = UltimaFilaConDatos(ws)
    ultimaCol = UltimaColConDatos(ws)

    If ultimaFila < 2 Then
        Err.Raise vbObjectError + 2000, "LeerCodiguera", "La codiguera no tiene filas de datos."
    End If

    arrDatos = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaCol)).Value2
    Set mapHeaders = MapearEncabezados(arrDatos)

    cIncluir = ObtenerColumna(mapHeaders, Array("Incluir_en_Informe"))
    cNivel1 = ObtenerColumna(mapHeaders, Array("Nivel_1"))
    cNivel2 = ObtenerColumna(mapHeaders, Array("Nivel_2"))
    cSubtipo = ObtenerColumna(mapHeaders, Array("Subtipo"))

    cFinac = ObtenerColumna(mapHeaders, Array("Finac"))
    cDerF = ObtenerColumna(mapHeaders, Array("Der-F", "Der F"))
    cPG = ObtenerColumna(mapHeaders, Array("PG"))
    cSpg = ObtenerColumna(mapHeaders, Array("Spg", "SPG"))
    cProy = ObtenerColumna(mapHeaders, Array("Proy", "Proyecto"))

    cRubroNum = ObtenerColumnaOpcional(mapHeaders, Array("Rubro_Num", "Rubro Num"))
    cRubro = ObtenerColumna(mapHeaders, Array("Rubro"))
    cRAuxNum = ObtenerColumnaOpcional(mapHeaders, Array("R. Aux_Num", "R Aux_Num", "R. Aux Num"))
    cRAux = ObtenerColumna(mapHeaders, Array("R. Aux", "R Aux"))

    cUE = ObtenerColumna(mapHeaders, Array("UE"))
    cDep = ObtenerColumnaOpcional(mapHeaders, Array("Dep", "DEP"))
    cObra = ObtenerColumna(mapHeaders, Array("Obra"))
    cDerObra = ObtenerColumna(mapHeaders, Array("Der. Obra", "Der Obra"))
    cServ = ObtenerColumna(mapHeaders, Array("Serv"))
    cSniip = ObtenerColumna(mapHeaders, Array("SNIIP", "SNIP"))

    llavesValidas = 0

    For i = 2 To UBound(arrDatos, 1)
        incluye = LimpiarTexto(CStr(ValorSeguro(arrDatos, i, cIncluir)))
        If UCase$(Replace(incluye, " ", "")) = "SI" Then

            comboKey = ConstruirClaveCombo(CStr(ValorSeguro(arrDatos, i, cNivel1)), _
                                           CStr(ValorSeguro(arrDatos, i, cNivel2)), _
                                           CStr(ValorSeguro(arrDatos, i, cSubtipo)))

            If Not dictCombos.Exists(comboKey) Then
                dictCombos.Add comboKey, Array(CStr(ValorSeguro(arrDatos, i, cNivel1)), _
                                               CStr(ValorSeguro(arrDatos, i, cNivel2)), _
                                               CStr(ValorSeguro(arrDatos, i, cSubtipo)))
            End If

            llave = ConstruirLlavePresupuestal( _
                    ValorSeguro(arrDatos, i, cFinac), _
                    ValorSeguro(arrDatos, i, cDerF), _
                    ValorSeguro(arrDatos, i, cPG), _
                    ValorSeguro(arrDatos, i, cSpg), _
                    ValorSeguro(arrDatos, i, cProy), _
                    ElegirRubroCodiguera(arrDatos, i, cRubroNum, cRubro), _
                    ElegirRAuxCodiguera(arrDatos, i, cRAuxNum, cRAux), _
                    ValorSeguro(arrDatos, i, cUE), _
                    IIf(cDep > 0, ValorSeguro(arrDatos, i, cDep), ValorSeguro(arrDatos, i, cUE)), _
                    ValorSeguro(arrDatos, i, cObra), _
                    ValorSeguro(arrDatos, i, cDerObra), _
                    ValorSeguro(arrDatos, i, cServ), _
                    ValorSeguro(arrDatos, i, cSniip))

            ' Supuesto documentado:
            ' Si la codiguera no trae columna DEP, se asume DEP = UE para construir la llave.
            If Not dictLlaveACombo.Exists(llave) Then
                dictLlaveACombo.Add llave, comboKey
                llavesValidas = llavesValidas + 1
            End If
        End If
    Next i
End Sub

Public Sub LeerEjecucionesYAcumular(ByVal ws As Worksheet, _
                                    ByVal anioReporte As Long, _
                                    ByRef dictLlaveACombo As Object, _
                                    ByRef dictAcumulado As Object, _
                                    ByRef registrosLeidos As Long)

    Dim arrDatos As Variant
    Dim mapHeaders As Object
    Dim ultimaFila As Long, ultimaCol As Long
    Dim i As Long

    Dim cFinac As Long, cDerF As Long, cPG As Long, cSpg As Long, cProy As Long
    Dim cRubro As Long, cRAux As Long, cUE As Long, cDep As Long
    Dim cObra As Long, cDerObra As Long, cServ As Long, cSniip As Long
    Dim cFechaValor As Long, cImporteMN As Long

    Dim fechaValor As Variant, importeMN As Variant
    Dim mes As Long
    Dim llave As String, comboKey As String

    registrosLeidos = 0

    ultimaFila = UltimaFilaConDatos(ws)
    ultimaCol = UltimaColConDatos(ws)

    If ultimaFila < 2 Then
        Err.Raise vbObjectError + 2100, "LeerEjecucionesYAcumular", "El archivo de ejecuciones no tiene filas de datos."
    End If

    arrDatos = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaCol)).Value2
    Set mapHeaders = MapearEncabezados(arrDatos)

    cFinac = ObtenerColumna(mapHeaders, Array("Finac Código numérico", "Finac"))
    cDerF = ObtenerColumna(mapHeaders, Array("Der-F Código numérico", "Der-F", "Der F"))
    cPG = ObtenerColumna(mapHeaders, Array("PG Código numérico", "PG"))
    cSpg = ObtenerColumna(mapHeaders, Array("Spg Código numérico", "Spg"))
    cProy = ObtenerColumna(mapHeaders, Array("Proyecto", "Proy"))
    cRubro = ObtenerColumna(mapHeaders, Array("Rubro Código numérico", "Rubro"))
    cRAux = ObtenerColumna(mapHeaders, Array("R. Aux Código numérico", "R. Aux", "R Aux"))
    cUE = ObtenerColumna(mapHeaders, Array("UE Código numérico", "UE"))
    cDep = ObtenerColumna(mapHeaders, Array("Dep Código numérico", "Dep"))
    cObra = ObtenerColumna(mapHeaders, Array("Obra Código numérico", "Obra"))
    cDerObra = ObtenerColumna(mapHeaders, Array("Der. Obra Código numérico", "Der. Obra", "Der Obra"))
    cServ = ObtenerColumna(mapHeaders, Array("Serv Código numérico", "Serv"))
    cSniip = ObtenerColumna(mapHeaders, Array("SNIP Código numérico", "SNIIP", "SNIP"))

    cFechaValor = ObtenerColumna(mapHeaders, Array("Fecha valor"))
    cImporteMN = ObtenerColumna(mapHeaders, Array("Importe moneda nacional"))

    For i = 2 To UBound(arrDatos, 1)
        registrosLeidos = registrosLeidos + 1

        fechaValor = ValorSeguro(arrDatos, i, cFechaValor)
        If Not EsFechaValida(fechaValor) Then
            GoTo SiguienteFila
        End If

        If Year(CDate(fechaValor)) <> anioReporte Then
            GoTo SiguienteFila
        End If

        importeMN = ValorSeguro(arrDatos, i, cImporteMN)
        If Not EsNumeroValido(importeMN) Then
            GoTo SiguienteFila
        End If

        mes = Month(CDate(fechaValor))

        llave = ConstruirLlavePresupuestal( _
                ValorSeguro(arrDatos, i, cFinac), _
                ValorSeguro(arrDatos, i, cDerF), _
                ValorSeguro(arrDatos, i, cPG), _
                ValorSeguro(arrDatos, i, cSpg), _
                ValorSeguro(arrDatos, i, cProy), _
                ValorSeguro(arrDatos, i, cRubro), _
                ValorSeguro(arrDatos, i, cRAux), _
                ValorSeguro(arrDatos, i, cUE), _
                ValorSeguro(arrDatos, i, cDep), _
                ValorSeguro(arrDatos, i, cObra), _
                ValorSeguro(arrDatos, i, cDerObra), _
                ValorSeguro(arrDatos, i, cServ), _
                ValorSeguro(arrDatos, i, cSniip))

        If dictLlaveACombo.Exists(llave) Then
            comboKey = CStr(dictLlaveACombo(llave))
            AcumularImporte dictAcumulado, comboKey, mes, CDbl(importeMN)
        End If

SiguienteFila:
    Next i
End Sub

Public Sub VolcarResultado(ByVal wbDestino As Workbook, _
                           ByVal anioReporte As Long, _
                           ByVal dictCombos As Object, _
                           ByVal dictAcumulado As Object, _
                           ByRef cantidadCombos As Long)

    Dim nombreHoja As String
    Dim ws As Worksheet
    Dim arrSalida() As Variant
    Dim encabezados As Variant

    Dim k As Variant
    Dim i As Long, j As Long
    Dim comboInfo As Variant
    Dim arrMeses As Variant
    Dim total As Double

    nombreHoja = "Ejec. Mensual " & CStr(anioReporte)

    EliminarHojaSiExiste wbDestino, nombreHoja
    Set ws = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
    ws.Name = nombreHoja

    encabezados = Array("Nivel_1", "Nivel_2", "Subtipo", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Total")

    cantidadCombos = dictCombos.Count
    ReDim arrSalida(1 To cantidadCombos + 1, 1 To 16)

    For j = 1 To 16
        arrSalida(1, j) = encabezados(j - 1)
    Next j

    i = 2
    For Each k In dictCombos.Keys
        comboInfo = dictCombos(k)

        arrSalida(i, 1) = comboInfo(0)
        arrSalida(i, 2) = comboInfo(1)
        arrSalida(i, 3) = comboInfo(2)

        If dictAcumulado.Exists(CStr(k)) Then
            arrMeses = dictAcumulado(CStr(k))
        Else
            arrMeses = InicializarArregloMeses()
        End If

        total = 0
        For j = 1 To 12
            arrSalida(i, 3 + j) = CDbl(arrMeses(j))
            total = total + CDbl(arrMeses(j))
        Next j

        arrSalida(i, 16) = total
        i = i + 1
    Next k

    With ws
        .Range(.Cells(1, 1), .Cells(UBound(arrSalida, 1), UBound(arrSalida, 2))).Value = arrSalida

        .Rows(1).Font.Bold = True
        .Range(.Cells(2, 4), .Cells(UBound(arrSalida, 1), 16)).NumberFormat = "#,##0.00"

        .Columns("A:C").AutoFit
        .Columns("D:P").ColumnWidth = 14

        ' Congelar en fila 2 y después de columnas descriptivas (A:C).
        ' Requiere activar la hoja para aplicar FreezePanes al window activo.
        .Activate
        ActiveWindow.SplitRow = 1
        ActiveWindow.SplitColumn = 3
        ActiveWindow.FreezePanes = True
    End With
End Sub

Option Explicit

Private Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Private Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Codiguera"

Public Sub Generar_Ejecucion_Mensual_GG()
    Dim anioReporte As Long: anioReporte = 2026
    Dim archivoEjec As String, archivoCod As String
    Dim wbEjec As Workbook, wbCod As Workbook, wsEjec As Worksheet, wsCod As Worksheet
    Dim dictLlaveACombo As Object, dictCombos As Object, dictAcum As Object, diag As Object
    Dim matches As Long, k As Variant

    Set diag = CrearDiagnosticoBase()
    Set dictLlaveACombo = CreateObject("Scripting.Dictionary")
    Set dictCombos = CreateObject("Scripting.Dictionary")
    Set dictAcum = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    archivoEjec = ObtenerArchivoMasReciente(RUTA_CARPETA_EJECUCIONES)
    archivoCod = ResolverArchivoCodiguera(RUTA_CODIGUERA)
    If Len(archivoEjec) = 0 Then Err.Raise vbObjectError + 1000, , "No se encontró archivo de ejecuciones."
    If Len(archivoCod) = 0 Then Err.Raise vbObjectError + 1001, , "No se encontró archivo de codiguera."

    Set wbEjec = Workbooks.Open(archivoEjec, ReadOnly:=True)
    Set wsEjec = ObtenerPrimeraHojaConDatos(wbEjec)
    Set wbCod = Workbooks.Open(archivoCod, ReadOnly:=True)
    Set wsCod = ObtenerHojaPorNombreExacto(wbCod, "Codiguera")
    If wsCod Is Nothing Then Err.Raise vbObjectError + 1002, , "No existe hoja Codiguera."

    diag("ruta_ejec") = archivoEjec: diag("archivo_ejec") = wbEjec.Name: diag("hoja_ejec") = wsEjec.Name
    diag("ruta_cod") = archivoCod: diag("archivo_cod") = wbCod.Name: diag("hoja_cod") = wsCod.Name

    LeerCodiguera wsCod, dictLlaveACombo, dictCombos, diag
    LeerEjecucionesYAcumular wsEjec, anioReporte, dictLlaveACombo, dictAcum, diag

    For Each k In diag("ej_set").Keys
        If diag("cod_set").Exists(CStr(k)) Then matches = matches + 1 Else If diag("ej_no_match").Count < 30 Then diag("ej_no_match").Add CStr(k)
    Next k
    diag("matches") = matches
    diag("no_matches") = diag("ej_set").Count - matches

    For Each k In diag("cod_set").Keys
        If Not diag("ej_set").Exists(CStr(k)) Then If diag("cod_no_match").Count < 30 Then diag("cod_no_match").Add CStr(k)
    Next k

    diag("match_sindep") = InterseccionCount(diag("cod_set_sindep"), diag("ej_set_sindep"))
    diag("match_clave_cod") = InterseccionCount(diag("cod_set_clave"), diag("ej_set"))
    diag("match_clave_sindep_cod") = InterseccionCount(diag("cod_set_clavesindep"), diag("ej_set_sindep"))

    EscribirDiagnostico ThisWorkbook, diag, anioReporte
    Debug.Print "[DEBUG] cantidad de matches: " & matches

    If diag("ej_2026") = 0 Then
        MsgBox "El archivo de ejecuciones más reciente no tiene datos del año " & anioReporte & " en 'Fecha valor'. Revisá Diagnostico_Llaves.", vbExclamation
    ElseIf matches = 0 Then
        MsgBox "Hay datos " & anioReporte & " pero no hubo coincidencias de llaves. Revisá Diagnostico_Llaves.", vbExclamation
    Else
        MsgBox "Proceso completado. Revisá Diagnostico_Llaves para detalle.", vbInformation
    End If

    wbEjec.Close False: wbCod.Close False
    Exit Sub
EH:
    On Error Resume Next
    EscribirDiagnostico ThisWorkbook, diag, anioReporte
    If Not wbEjec Is Nothing Then wbEjec.Close False
    If Not wbCod Is Nothing Then wbCod.Close False
    MsgBox "Error: " & Err.Description & vbCrLf & "Revisá Diagnostico_Llaves.", vbCritical
End Sub

Private Function CrearDiagnosticoBase() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Set d("cod_keys") = New Collection
    Set d("ej_keys") = New Collection
    Set d("ej_no_match") = New Collection
    Set d("cod_no_match") = New Collection

    Set d("cod_set") = CreateObject("Scripting.Dictionary")
    Set d("ej_set") = CreateObject("Scripting.Dictionary")
    Set d("cod_set_sindep") = CreateObject("Scripting.Dictionary")
    Set d("ej_set_sindep") = CreateObject("Scripting.Dictionary")
    Set d("cod_set_clave") = CreateObject("Scripting.Dictionary")
    Set d("cod_set_clavesindep") = CreateObject("Scripting.Dictionary")

    d("ruta_ejec") = vbNullString
    d("archivo_ejec") = vbNullString
    d("hoja_ejec") = vbNullString
    d("ruta_cod") = vbNullString
    d("archivo_cod") = vbNullString
    d("hoja_cod") = vbNullString

    d("cod_total") = 0
    d("cod_si") = 0
    d("ej_total") = 0
    d("ej_2026") = 0
    d("ej_importe_num") = 0
    d("ej_fecha_invalida") = 0
    d("matches") = 0
    d("no_matches") = 0
    d("match_sindep") = 0
    d("match_clave_cod") = 0
    d("match_clave_sindep_cod") = 0

    Set CrearDiagnosticoBase = d
End Function

Private Function InterseccionCount(ByVal a As Object, ByVal b As Object) As Long
    Dim k As Variant
    For Each k In a.Keys
        If b.Exists(CStr(k)) Then InterseccionCount = InterseccionCount + 1
    Next k
End Function

Private Sub EscribirDiagnostico(ByVal wb As Workbook, ByVal d As Object, ByVal anio As Long)
    Dim ws As Worksheet
    Dim r As Long

    EliminarHojaSiExiste wb, "Diagnostico_Llaves"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Diagnostico_Llaves"

    r = 1
    PutKV ws, r, "ruta completa del archivo de ejecuciones usado", d("ruta_ejec")
    PutKV ws, r, "nombre del archivo de ejecuciones usado", d("archivo_ejec")
    PutKV ws, r, "nombre de la hoja de ejecuciones usada", d("hoja_ejec")
    PutKV ws, r, "ruta completa del archivo de codiguera usado", d("ruta_cod")
    PutKV ws, r, "nombre del archivo de codiguera usado", d("archivo_cod")
    PutKV ws, r, "hoja de codiguera usada", d("hoja_cod")
    PutKV ws, r, "cantidad total de filas de codiguera", d("cod_total")
    PutKV ws, r, "cantidad de filas de codiguera con Incluir_en_Informe = SI", d("cod_si")
    PutKV ws, r, "cantidad de llaves únicas cargadas desde codiguera", d("cod_set").Count
    PutKV ws, r, "cantidad total de filas de ejecuciones", d("ej_total")
    PutKV ws, r, "cantidad de filas de ejecuciones con Fecha valor del año " & anio, d("ej_2026")
    PutKV ws, r, "cantidad de fechas inválidas en Fecha valor", d("ej_fecha_invalida")
    PutKV ws, r, "cantidad de filas de ejecuciones con importe numérico", d("ej_importe_num")
    PutKV ws, r, "cantidad de llaves únicas generadas desde ejecuciones para " & anio, d("ej_set").Count
    PutKV ws, r, "cantidad de coincidencias entre codiguera y ejecuciones", d("matches")
    PutKV ws, r, "cantidad de no coincidencias", d("no_matches")
    PutKV ws, r, "coincidencias alternativa sin Dep", d("match_sindep")
    PutKV ws, r, "coincidencias usando codiguera 'Clave Llave presupuestal'", d("match_clave_cod")
    PutKV ws, r, "coincidencias usando codiguera 'Clave sin dep'", d("match_clave_sindep_cod")

    r = r + 1
    PutLista ws, r, "Primeras 30 llaves codiguera", d("cod_keys")
    r = r + 1
    PutLista ws, r, "Primeras 30 llaves ejecuciones " & anio, d("ej_keys")
    r = r + 1
    PutLista ws, r, "Primeras 30 llaves ejecuciones no matchean", d("ej_no_match")
    r = r + 1
    PutLista ws, r, "Primeras 30 llaves codiguera SI no matchean", d("cod_no_match")

    ws.Columns("A:C").AutoFit
End Sub

Private Sub PutKV(ByVal ws As Worksheet, ByRef r As Long, ByVal k As String, ByVal v As Variant)
    ws.Cells(r, 1).Value = k
    ws.Cells(r, 2).Value = v
    r = r + 1
End Sub

Private Sub PutLista(ByVal ws As Worksheet, ByRef r As Long, ByVal titulo As String, ByVal col As Collection)
    Dim i As Long

    ws.Cells(r, 1).Value = titulo
    ws.Cells(r, 1).Font.Bold = True

    For i = 1 To col.Count
        ws.Cells(r + i, 2).Value = col(i)
    Next i
    r = r + IIf(col.Count > 0, col.Count + 1, 2)
End Sub

Option Explicit

Public Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Public Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Codiguera"
Public Const RUTA_REPORTES_GENERADOS As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\ReportesGenerados"
Public Const RUTA_CARPETA_ASIGNADOS_GASTOS As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\Asignados\Gastos"


Public Function RutaBaseLocalReporteGG() As String
    RutaBaseLocalReporteGG = ThisWorkbook.Path
End Function

Public Function CombinarRuta(ByVal carpeta As String, ByVal nombre As String) As String
    If Right$(carpeta, 1) = "\" Then
        CombinarRuta = carpeta & nombre
    Else
        CombinarRuta = carpeta & "\" & nombre
    End If
End Function

Private Function CarpetaExiste(ByVal ruta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CarpetaExiste = fso.FolderExists(ruta)
End Function

Private Function ArchivoExiste(ByVal ruta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ArchivoExiste = fso.FileExists(ruta)
End Function

Public Function BuscarArchivoLocalPorPatron(ByVal carpeta As String, ByVal patronContiene As String) As String
    Dim fso As Object
    Dim folderObj As Object
    Dim archivo As Object
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(carpeta) Then Exit Function

    Set folderObj = fso.GetFolder(carpeta)

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                If InStr(1, archivo.Name, patronContiene, vbTextCompare) > 0 Then
                    BuscarArchivoLocalPorPatron = archivo.Path
                    Exit Function
                End If
            End If
        End If
SiguienteArchivo:
    Next archivo
End Function

Public Function RutaCarpetaEjecucionesActiva() As String
    Dim candidata As String

    If CarpetaExiste(RUTA_CARPETA_EJECUCIONES) Then
        RutaCarpetaEjecucionesActiva = RUTA_CARPETA_EJECUCIONES
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "DatosDescargados\DetalleRegistros\Ejecuciones")
    If CarpetaExiste(candidata) Then
        RutaCarpetaEjecucionesActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "DetalleRegistros\Ejecuciones")
    If CarpetaExiste(candidata) Then
        RutaCarpetaEjecucionesActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "Ejecuciones")
    If CarpetaExiste(candidata) Then
        RutaCarpetaEjecucionesActiva = candidata
        Exit Function
    End If

    RutaCarpetaEjecucionesActiva = candidata
End Function

Public Function RutaCodigueraActiva() As String
    Dim candidata As String

    If CarpetaExiste(RUTA_CODIGUERA) Or ArchivoExiste(RUTA_CODIGUERA) Then
        RutaCodigueraActiva = RUTA_CODIGUERA
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "Codiguera")
    If CarpetaExiste(candidata) Or ArchivoExiste(candidata) Then
        RutaCodigueraActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "Copia de Llave presupuestal gastos (bps.budget.key.expense)(2).xlsx")
    If ArchivoExiste(candidata) Then
        RutaCodigueraActiva = candidata
        Exit Function
    End If

    candidata = BuscarArchivoLocalPorPatron(RutaBaseLocalReporteGG(), "Llave presupuestal")
    If Len(candidata) > 0 Then
        RutaCodigueraActiva = candidata
        Exit Function
    End If

    RutaCodigueraActiva = CombinarRuta(RutaBaseLocalReporteGG(), "Codiguera")
End Function

Public Function RutaCarpetaAsignadosGastosActiva() As String
    Dim candidata As String

    If CarpetaExiste(RUTA_CARPETA_ASIGNADOS_GASTOS) Then
        RutaCarpetaAsignadosGastosActiva = RUTA_CARPETA_ASIGNADOS_GASTOS
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "DatosDescargados\Asignados\Gastos")
    If CarpetaExiste(candidata) Then
        RutaCarpetaAsignadosGastosActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "Asignados\Gastos")
    If CarpetaExiste(candidata) Then
        RutaCarpetaAsignadosGastosActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(RutaBaseLocalReporteGG(), "Asignados")
    If CarpetaExiste(candidata) Then
        RutaCarpetaAsignadosGastosActiva = candidata
        Exit Function
    End If

    RutaCarpetaAsignadosGastosActiva = CombinarRuta(RutaBaseLocalReporteGG(), "Asignados\Gastos")
End Function

Public Function RutaCarpetaEjecucionesAnioActiva(ByVal anio As Long) As String
    RutaCarpetaEjecucionesAnioActiva = CombinarRuta(RutaCarpetaEjecucionesActiva(), CStr(anio))
End Function

Public Function RutaCarpetaAsignadosGastosAnioActiva(ByVal anio As Long) As String
    RutaCarpetaAsignadosGastosAnioActiva = CombinarRuta(RutaCarpetaAsignadosGastosActiva(), CStr(anio))
End Function

Public Function RutaCarpetaIndicesActiva() As String
    Dim candidata As String
    Dim rutaPadre As String

    candidata = "D:\Escritorio\Reporte GG\Indices"
    If CarpetaExiste(candidata) Then
        RutaCarpetaIndicesActiva = candidata
        Exit Function
    End If

    candidata = CombinarRuta(ThisWorkbook.Path, "Indices")
    If CarpetaExiste(candidata) Then
        RutaCarpetaIndicesActiva = candidata
        Exit Function
    End If

    rutaPadre = Left$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    If Len(rutaPadre) > 0 Then
        candidata = CombinarRuta(rutaPadre, "Indices")
        If CarpetaExiste(candidata) Then
            RutaCarpetaIndicesActiva = candidata
            Exit Function
        End If
    End If

    RutaCarpetaIndicesActiva = CombinarRuta(ThisWorkbook.Path, "Indices")
End Function

Public Function ResolverArchivoIndice(ByVal tipoIndice As String) As String
    Dim t As String
    Dim rutaArchivo As String

    t = UCase$(Trim$(tipoIndice))
    If t = "IPC GRAL" Or t = "IPC GENERAL" Then t = "IPC"
    If t = "IMSN M B08" Then t = "IMSN"

    Select Case t
        Case "IPC"
            rutaArchivo = CombinarRuta(RutaCarpetaIndicesActiva(), "IPC gral y variaciones_base 2022.xlsx")
        Case "IMSN"
            rutaArchivo = CombinarRuta(RutaCarpetaIndicesActiva(), "IMSN M B08.xlsx")
        Case Else
            Err.Raise vbObjectError + 1960, "ResolverArchivoIndice", "Tipo de índice no válido: '" & tipoIndice & "'. Debe ser IPC o IMSN."
    End Select

    If Not ArchivoExiste(rutaArchivo) Then
        Err.Raise vbObjectError + 1961, "ResolverArchivoIndice", "No se encontró el archivo de índice para '" & tipoIndice & "'. Archivo esperado: " & rutaArchivo
    End If

    ResolverArchivoIndice = rutaArchivo
End Function

Public Function RutaReportesGeneradosActiva() As String
    If CarpetaExiste(RUTA_REPORTES_GENERADOS) Then
        RutaReportesGeneradosActiva = RUTA_REPORTES_GENERADOS
    Else
        RutaReportesGeneradosActiva = CombinarRuta(RutaBaseLocalReporteGG(), "ReportesGenerados")
    End If
End Function

Public Function CarpetaSalidaReportesActiva() As String
    CarpetaSalidaReportesActiva = RutaReportesGeneradosActiva()
End Function

Public Function CarpetaControlesReporteActiva() As String
    Dim base As String
    Dim carpeta As String
    base = CarpetaSalidaReportesActiva()
    AsegurarCarpetaExiste base
    carpeta = CombinarRuta(base, "Controles")
    AsegurarCarpetaExiste carpeta
    CarpetaControlesReporteActiva = carpeta
End Function

Public Function RutaArchivoControlReporte(ByVal anio As Long, ByVal mesCierre As Long) As String
    RutaArchivoControlReporte = CombinarRuta( _
        CarpetaControlesReporteActiva(), _
        "Control_Reporte_GG_" & anio & "_" & Format$(mesCierre, "00") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".xlsx")
End Function

Public Function DiagnosticoRutasActivas() As String
    Dim rutaEjec As String, rutaCod As String, rutaAsig As String, rutaOut As String

    rutaEjec = RutaCarpetaEjecucionesActiva()
    rutaCod = RutaCodigueraActiva()
    rutaAsig = RutaCarpetaAsignadosGastosActiva()
    rutaOut = RutaReportesGeneradosActiva()

    DiagnosticoRutasActivas = "Rutas activas:" & vbCrLf
    DiagnosticoRutasActivas = DiagnosticoRutasActivas & "Base local: " & RutaBaseLocalReporteGG() & vbCrLf
    DiagnosticoRutasActivas = DiagnosticoRutasActivas & "Ejecuciones: " & rutaEjec & " | Existe: " & IIf(CarpetaExiste(rutaEjec), "SI", "NO") & vbCrLf
    DiagnosticoRutasActivas = DiagnosticoRutasActivas & "Codiguera: " & rutaCod & " | Existe: " & IIf(CarpetaExiste(rutaCod) Or ArchivoExiste(rutaCod), "SI", "NO") & vbCrLf
    DiagnosticoRutasActivas = DiagnosticoRutasActivas & "Asignados gastos: " & rutaAsig & " | Existe: " & IIf(CarpetaExiste(rutaAsig), "SI", "NO") & vbCrLf
    DiagnosticoRutasActivas = DiagnosticoRutasActivas & "Reportes generados: " & rutaOut & " | Existe: " & IIf(CarpetaExiste(rutaOut), "SI", "NO")
End Function

Public Function ObtenerHojaPanelReportes() As Worksheet
    On Error GoTo EH
    Set ObtenerHojaPanelReportes = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    Exit Function
EH:
    Err.Raise vbObjectError + 901, "ObtenerHojaPanelReportes", "No existe la hoja '" & PANEL_SHEET_NAME & "'. Ejecute CrearOActualizarPanelReportes."
End Function

Public Function ObtenerArchivoMasReciente(ByVal carpeta As String) As String
    On Error GoTo EH

    Dim fso As Object
    Dim archivo As Object
    Dim folderObj As Object
    Dim ultimaFecha As Date
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim$(carpeta)) = 0 Then
        Err.Raise vbObjectError + 1900, "ObtenerArchivoMasReciente", "La carpeta recibida está vacía."
    End If

    If Not fso.FolderExists(carpeta) Then
        Err.Raise vbObjectError + 1901, "ObtenerArchivoMasReciente", "La carpeta no existe: " & carpeta
    End If

    Set folderObj = fso.GetFolder(carpeta)
    ultimaFecha = #1/1/1900#

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))

        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                If archivo.Size > 0 Then
                    If archivo.DateLastModified > ultimaFecha Then
                        ultimaFecha = archivo.DateLastModified
                        ObtenerArchivoMasReciente = archivo.Path
                    End If
                End If
            End If
        End If
SiguienteArchivo:
    Next archivo

    If Len(ObtenerArchivoMasReciente) = 0 Then
        Err.Raise vbObjectError + 1902, "ObtenerArchivoMasReciente", "No se encontró archivo xls/xlsx/xlsm válido en: " & carpeta
    End If

    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerArchivoMasReciente", "Error buscando archivo más reciente en: " & carpeta & " | " & Err.Description
End Function

Public Function ResolverArchivoCodiguera(ByVal ruta As String) As String
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ruta) Then
        ResolverArchivoCodiguera = ruta
    Else
        ResolverArchivoCodiguera = ObtenerArchivoMasReciente(ruta)
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "ResolverArchivoCodiguera", "Error resolviendo archivo codiguera en: " & ruta & " | " & Err.Description
End Function

Public Function ObtenerHojaCodiguera(ByVal wb As Workbook) As Worksheet
    On Error GoTo EH
    If wb Is Nothing Then Err.Raise vbObjectError + 1100, "ObtenerHojaCodiguera", "El workbook recibido es Nothing."
    Set ObtenerHojaCodiguera = wb.Worksheets("Codiguera")
    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerHojaCodiguera", "No se encontró hoja 'Codiguera' en: " & wb.Name
End Function

Public Function ObtenerHojaEjecuciones(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim headers As Object
    Dim req As Variant
    Dim i As Long
    Dim lastCol As Long
    Dim arrHeader As Variant
    Dim hojasEncontradas As String
    Dim columnasRequeridas As String

    If wb Is Nothing Then
        Err.Raise vbObjectError + 1101, "ObtenerHojaEjecuciones", "El workbook recibido es Nothing."
    End If

    req = Array("fecha valor", "finac código numérico", "der-f código numérico", "pg código numérico", "spg código numérico", "rubro código numérico", "r. aux código numérico", "ue código numérico", "dep código numérico", "obra código numérico", "der. obra código numérico", "serv código numérico", "importe moneda nacional")
    columnasRequeridas = Join(req, ", ")

    For Each ws In wb.Worksheets
        If Len(hojasEncontradas) > 0 Then hojasEncontradas = hojasEncontradas & ", "
        hojasEncontradas = hojasEncontradas & ws.Name

        lastCol = UltimaColConDatos(ws)
        If lastCol <= 0 Then GoTo SiguienteHoja

        arrHeader = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2
        Set headers = MapearEncabezados(arrHeader)

        For i = LBound(req) To UBound(req)
            If Not headers.Exists(CStr(req(i))) Then GoTo SiguienteHoja
        Next i

        Set ObtenerHojaEjecuciones = ws
        Exit Function
SiguienteHoja:
    Next ws

    Dim detalleHojaEjecuciones As String
    detalleHojaEjecuciones = "No se encontró una hoja válida de ejecuciones en workbook: " & wb.Name
    detalleHojaEjecuciones = detalleHojaEjecuciones & " | Hojas encontradas: " & hojasEncontradas
    detalleHojaEjecuciones = detalleHojaEjecuciones & " | Columnas requeridas: " & columnasRequeridas
    Err.Raise vbObjectError + 1103, "ObtenerHojaEjecuciones", detalleHojaEjecuciones
End Function

Public Function ObtenerHojaAsignados(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim headers As Object
    Dim req As Variant
    Dim i As Long
    Dim lastCol As Long
    Dim arrHeader As Variant
    Dim hojasEncontradas As String
    Dim columnasRequeridas As String

    If wb Is Nothing Then
        Err.Raise vbObjectError + 1200, "ObtenerHojaAsignados", "El workbook recibido es Nothing."
    End If

    req = Array("finac", "der-f", "pg", "spg", "proy", "rubro", "r. aux", "ue", "dep", "obra", "der. obra", "serv", "sniip", "asignado")
    columnasRequeridas = Join(req, ", ")

    For Each ws In wb.Worksheets
        If Len(hojasEncontradas) > 0 Then hojasEncontradas = hojasEncontradas & ", "
        hojasEncontradas = hojasEncontradas & ws.Name

        lastCol = UltimaColConDatos(ws)
        If lastCol <= 0 Then GoTo SiguienteHoja

        arrHeader = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2
        Set headers = MapearEncabezados(arrHeader)
        For i = LBound(req) To UBound(req)
            If Not headers.Exists(CStr(req(i))) Then GoTo SiguienteHoja
        Next i
        Set ObtenerHojaAsignados = ws
        Exit Function
SiguienteHoja:
    Next ws

    Dim detalleHojaAsignados As String
    detalleHojaAsignados = "No se encontró una hoja válida de asignados en workbook: " & wb.Name
    detalleHojaAsignados = detalleHojaAsignados & " | Hojas encontradas: " & hojasEncontradas
    detalleHojaAsignados = detalleHojaAsignados & " | Columnas requeridas: " & columnasRequeridas
    Err.Raise vbObjectError + 1203, "ObtenerHojaAsignados", detalleHojaAsignados
End Function

Public Sub AsegurarCarpetaExiste(ByVal ruta As String)
    On Error GoTo EH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim$(ruta)) = 0 Then Exit Sub

    If Not fso.FolderExists(ruta) Then
        fso.CreateFolder ruta
    End If

    Exit Sub
EH:
    Err.Raise Err.Number, "AsegurarCarpetaExiste", _
        "No se pudo asegurar carpeta: " & ruta & " | " & Err.Description
End Sub


Public Function ObtenerArchivoMasRecientePorFechaCreacion(ByVal carpeta As String) As String
    On Error GoTo EH
    Dim fso As Object
    Dim archivo As Object
    Dim folderObj As Object
    Dim mejorArchivo As String
    Dim mejorCreacion As Date
    Dim mejorModificacion As Date
    Dim ext As String
    Dim fechaCreacion As Date
    Dim fechaModificacion As Date

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim$(carpeta)) = 0 Then
        Err.Raise vbObjectError + 1910, "ObtenerArchivoMasRecientePorFechaCreacion", "La carpeta recibida está vacía."
    End If

    If Not fso.FolderExists(carpeta) Then
        Err.Raise vbObjectError + 1911, "ObtenerArchivoMasRecientePorFechaCreacion", "La carpeta no existe: " & carpeta
    End If

    Set folderObj = fso.GetFolder(carpeta)

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))

        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                If archivo.Size = 0 Then GoTo SiguienteArchivo

                fechaCreacion = archivo.DateCreated
                fechaModificacion = archivo.DateLastModified

                If Len(mejorArchivo) = 0 Then
                    mejorArchivo = archivo.Path
                    mejorCreacion = fechaCreacion
                    mejorModificacion = fechaModificacion
                ElseIf fechaCreacion > mejorCreacion Then
                    mejorArchivo = archivo.Path
                    mejorCreacion = fechaCreacion
                    mejorModificacion = fechaModificacion
                ElseIf fechaCreacion = mejorCreacion Then
                    If fechaModificacion > mejorModificacion Then
                        mejorArchivo = archivo.Path
                        mejorCreacion = fechaCreacion
                        mejorModificacion = fechaModificacion
                    End If
                End If

            End If
        End If
SiguienteArchivo:
    Next archivo

    ObtenerArchivoMasRecientePorFechaCreacion = mejorArchivo

    If Len(ObtenerArchivoMasRecientePorFechaCreacion) = 0 Then
        Err.Raise vbObjectError + 1912, "ObtenerArchivoMasRecientePorFechaCreacion", "No se encontró archivo xls/xlsx/xlsm válido en: " & carpeta
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerArchivoMasRecientePorFechaCreacion", "Error buscando archivo por fecha de creación en: " & carpeta & " | " & Err.Description
End Function

Public Function DiagnosticoArchivosAsignados(ByVal carpeta As String) As String
    On Error GoTo EH
    Dim fso As Object
    Dim folderObj As Object
    Dim archivo As Object
    Dim mejorArchivo As String
    Dim mejorCreacion As Date
    Dim mejorModificacion As Date
    Dim ext As String
    Dim fechaCreacion As Date
    Dim fechaModificacion As Date
    Dim motivoSeleccion As String
    Dim detalle As String
    Dim seleccionado As String
    Dim motivoFila As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(carpeta) Then
        DiagnosticoArchivosAsignados = "Carpeta no existe: " & carpeta
        Exit Function
    End If

    Set folderObj = fso.GetFolder(carpeta)

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                fechaCreacion = archivo.DateCreated
                fechaModificacion = archivo.DateLastModified

                If Len(mejorArchivo) = 0 Then
                    mejorArchivo = archivo.Path
                    mejorCreacion = fechaCreacion
                    mejorModificacion = fechaModificacion
                    motivoSeleccion = "mayor DateCreated"
                ElseIf fechaCreacion > mejorCreacion Then
                    mejorArchivo = archivo.Path
                    mejorCreacion = fechaCreacion
                    mejorModificacion = fechaModificacion
                    motivoSeleccion = "mayor DateCreated"
                ElseIf fechaCreacion = mejorCreacion Then
                    If fechaModificacion > mejorModificacion Then
                        mejorArchivo = archivo.Path
                        mejorCreacion = fechaCreacion
                        mejorModificacion = fechaModificacion
                        motivoSeleccion = "desempate por DateLastModified"
                    End If
                End If
            End If
        End If
    Next archivo

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                seleccionado = IIf(StrComp(archivo.Path, mejorArchivo, vbTextCompare) = 0, "SI", "NO")
                motivoFila = ""
                If seleccionado = "SI" Then motivoFila = motivoSeleccion

                detalle = detalle & archivo.Name
                detalle = detalle & " | DateCreated=" & Format$(archivo.DateCreated, "yyyy-mm-dd hh:nn:ss")
                detalle = detalle & " | DateLastModified=" & Format$(archivo.DateLastModified, "yyyy-mm-dd hh:nn:ss")
                detalle = detalle & " | SELECCIONADO=" & seleccionado
                detalle = detalle & IIf(Len(motivoFila) > 0, " | Motivo=" & motivoFila, "")
                detalle = detalle & vbCrLf
            End If
        End If
    Next archivo

    DiagnosticoArchivosAsignados = "Carpeta: " & carpeta & vbCrLf
    DiagnosticoArchivosAsignados = DiagnosticoArchivosAsignados & "Archivo seleccionado: " & IIf(Len(mejorArchivo) > 0, mejorArchivo, "(ninguno)") & vbCrLf
    DiagnosticoArchivosAsignados = DiagnosticoArchivosAsignados & "Criterio: mayor DateCreated; si empata, mayor DateLastModified" & vbCrLf
    DiagnosticoArchivosAsignados = DiagnosticoArchivosAsignados & detalle
    Exit Function
EH:
    Err.Raise Err.Number, "DiagnosticoArchivosAsignados", "Error armando diagnóstico de asignados en: " & carpeta & " | " & Err.Description
End Function


Public Function DiagnosticoArchivosExcelCarpeta(ByVal carpeta As String) As String
    On Error GoTo EH

    Dim fso As Object
    Dim folderObj As Object
    Dim archivo As Object
    Dim ext As String
    Dim s As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim$(carpeta)) = 0 Then
        DiagnosticoArchivosExcelCarpeta = "Carpeta vacía."
        Exit Function
    End If

    If Not fso.FolderExists(carpeta) Then
        DiagnosticoArchivosExcelCarpeta = "Carpeta no existe: " & carpeta
        Exit Function
    End If

    Set folderObj = fso.GetFolder(carpeta)

    s = "Carpeta: " & carpeta & vbCrLf

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
            s = s & "- " & archivo.Name
            s = s & " | Creado: " & Format$(archivo.DateCreated, "yyyy-mm-dd hh:nn:ss")
            s = s & " | Modificado: " & Format$(archivo.DateLastModified, "yyyy-mm-dd hh:nn:ss")
            s = s & " | Size: " & CStr(archivo.Size)
            s = s & vbCrLf
        End If
    Next archivo

    If Len(s) = 0 Then s = "No se encontraron archivos Excel en: " & carpeta

    DiagnosticoArchivosExcelCarpeta = s
    Exit Function

EH:
    DiagnosticoArchivosExcelCarpeta = "Error diagnosticando carpeta: " & carpeta & " | " & Err.Description
End Function

Public Function RutaCompletaWorkbookSeguro(ByVal wb As Workbook) As String
    On Error Resume Next

    If wb Is Nothing Then
        RutaCompletaWorkbookSeguro = "(Nothing)"
    ElseIf Len(wb.Path) > 0 Then
        RutaCompletaWorkbookSeguro = wb.FullName
    Else
        RutaCompletaWorkbookSeguro = wb.Name
    End If

    On Error GoTo 0
End Function

Public Function DiagnosticoWorkbooksAbiertos() As String
    Dim wb As Workbook
    Dim s As String

    s = "Workbooks abiertos:" & vbCrLf

    For Each wb In Application.Workbooks
        s = s & "- Name: " & wb.Name
        s = s & " | FullName: " & RutaCompletaWorkbookSeguro(wb)
        s = s & vbCrLf
    Next wb

    DiagnosticoWorkbooksAbiertos = s
End Function

Public Function WorkbookAbiertoPorNombre(ByVal nombreArchivo As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nombreArchivo, vbTextCompare) = 0 Then
            Set WorkbookAbiertoPorNombre = wb
            Exit Function
        End If
    Next wb
End Function

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

Public Function RutaReportesGeneradosActiva() As String
    If CarpetaExiste(RUTA_REPORTES_GENERADOS) Then
        RutaReportesGeneradosActiva = RUTA_REPORTES_GENERADOS
    Else
        RutaReportesGeneradosActiva = CombinarRuta(RutaBaseLocalReporteGG(), "ReportesGenerados")
    End If
End Function

Public Function DiagnosticoRutasActivas() As String
    Dim rutaEjec As String, rutaCod As String, rutaAsig As String, rutaOut As String

    rutaEjec = RutaCarpetaEjecucionesActiva()
    rutaCod = RutaCodigueraActiva()
    rutaAsig = RutaCarpetaAsignadosGastosActiva()
    rutaOut = RutaReportesGeneradosActiva()

    DiagnosticoRutasActivas = "Rutas activas:" & vbCrLf & _
                             "Base local: " & RutaBaseLocalReporteGG() & vbCrLf & _
                             "Ejecuciones: " & rutaEjec & " | Existe: " & IIf(CarpetaExiste(rutaEjec), "SI", "NO") & vbCrLf & _
                             "Codiguera: " & rutaCod & " | Existe: " & IIf(CarpetaExiste(rutaCod) Or ArchivoExiste(rutaCod), "SI", "NO") & vbCrLf & _
                             "Asignados gastos: " & rutaAsig & " | Existe: " & IIf(CarpetaExiste(rutaAsig), "SI", "NO") & vbCrLf & _
                             "Reportes generados: " & rutaOut & " | Existe: " & IIf(CarpetaExiste(rutaOut), "SI", "NO")
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
    Dim fso As Object, archivo As Object, folderObj As Object, ultimaFecha As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function
    Set folderObj = fso.GetFolder(carpeta): ultimaFecha = #1/1/1900#
    For Each archivo In folderObj.Files
        If LCase$(fso.GetExtensionName(archivo.Name)) Like "xls*" Then
            If archivo.DateLastModified > ultimaFecha Then ultimaFecha = archivo.DateLastModified: ObtenerArchivoMasReciente = archivo.Path
        End If
    Next archivo
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
    Set ObtenerHojaCodiguera = wb.Worksheets("Codiguera")
    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerHojaCodiguera", "No se encontró hoja 'Codiguera' en: " & wb.Name
End Function

Public Function ObtenerHojaEjecuciones(ByVal wb As Workbook) As Worksheet
    On Error GoTo EH
    Set ObtenerHojaEjecuciones = wb.Worksheets(1)
    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerHojaEjecuciones", "No fue posible obtener hoja de ejecuciones en: " & wb.Name
End Function

Public Sub AsegurarCarpetaExiste(ByVal ruta As String)
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then fso.CreateFolder ruta
    Exit Sub
EH:
    Err.Raise Err.Number, "AsegurarCarpetaExiste", "No se pudo asegurar carpeta: " & ruta & " | " & Err.Description
End Sub


Public Function ObtenerArchivoMasRecientePorFechaCreacion(ByVal carpeta As String) As String
    On Error GoTo EH
    Dim fso As Object, archivo As Object, folderObj As Object
    Dim ultimaFecha As Date, ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function

    Set folderObj = fso.GetFolder(carpeta)
    ultimaFecha = #1/1/1900#

    For Each archivo In folderObj.Files
        ext = LCase$(fso.GetExtensionName(archivo.Name))
        If Left$(archivo.Name, 2) <> "~$" Then
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                If archivo.DateCreated > ultimaFecha Then
                    ultimaFecha = archivo.DateCreated
                    ObtenerArchivoMasRecientePorFechaCreacion = archivo.Path
                End If
            End If
        End If
    Next archivo

    If Len(ObtenerArchivoMasRecientePorFechaCreacion) = 0 Then
        Err.Raise vbObjectError + 1902, "ObtenerArchivoMasRecientePorFechaCreacion", "No se encontró archivo xls/xlsx/xlsm en carpeta de asignados: " & carpeta
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "ObtenerArchivoMasRecientePorFechaCreacion", "Error buscando archivo por fecha de creación en: " & carpeta & " | " & Err.Description
End Function

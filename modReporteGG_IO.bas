Option Explicit

Public Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Public Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Codiguera"
Public Const RUTA_REPORTES_GENERADOS As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\ReportesGenerados"

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

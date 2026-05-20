Option Explicit

Public Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Public Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Codiguera"
Public Const RUTA_REPORTES_GENERADOS As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\ReportesGenerados"

Public Function ObtenerArchivoMasReciente(ByVal carpeta As String) As String
    Dim fso As Object, archivo As Object, folderObj As Object, ultimaFecha As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function
    Set folderObj = fso.GetFolder(carpeta): ultimaFecha = #1/1/1900#
    For Each archivo In folderObj.Files
        If LCase$(fso.GetExtensionName(archivo.Name)) Like "xls*" Then
            If archivo.DateLastModified > ultimaFecha Then ultimaFecha = archivo.DateLastModified: ObtenerArchivoMasReciente = archivo.Path
        End If
    Next archivo
End Function
Public Function ResolverArchivoCodiguera(ByVal ruta As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ruta) Then ResolverArchivoCodiguera = ruta Else ResolverArchivoCodiguera = ObtenerArchivoMasReciente(ruta)
End Function
Public Function ObtenerHojaCodiguera(ByVal wb As Workbook) As Worksheet: Set ObtenerHojaCodiguera = wb.Worksheets("Codiguera"): End Function
Public Function ObtenerHojaEjecuciones(ByVal wb As Workbook) As Worksheet: Set ObtenerHojaEjecuciones = wb.Worksheets(1): End Function
Public Sub AsegurarCarpetaExiste(ByVal ruta As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then fso.CreateFolder ruta
End Sub

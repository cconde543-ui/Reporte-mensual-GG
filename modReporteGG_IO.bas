Option Explicit

Public Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Public Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\Codiguera"
Public Const RUTA_REPORTES_GENERADOS As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\Reporte GG\ReportesGenerados"

Public Function ObtenerArchivoMasReciente(ByVal carpeta As String) As String
    Dim fso As Object, f As Object, folder As Object, dt As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function
    Set folder = fso.GetFolder(carpeta)
    dt = #1/1/1900#
    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.Name)) Like "xls*" Then
            If f.DateLastModified > dt Then dt = f.DateLastModified: ObtenerArchivoMasReciente = f.Path
        End If
    Next f
End Function

Public Function ResolverArchivoCodiguera(ByVal ruta As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ruta) Then ResolverArchivoCodiguera = ruta Else ResolverArchivoCodiguera = ObtenerArchivoMasReciente(ruta)
End Function

Public Sub AsegurarCarpetaExiste(ByVal ruta As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then fso.CreateFolder ruta
End Sub

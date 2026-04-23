Attribute VB_Name = "modReporteGG_Main"
Option Explicit

Private Const RUTA_CARPETA_EJECUCIONES As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\DatosDescargados\DetalleRegistros\Ejecuciones"
Private Const RUTA_CODIGUERA As String = "\\estructura\Finanzas\AREA Contaduria\Adm Presupuestal\Prest y Recursos\SISTEMA DE CONTROL PRESUPUESTAL\SeguimientoPresupuestal\Reporte GG\Codiguera"

Public Sub Generar_Ejecucion_Mensual_GG()
    Const NOMBRE_MACRO As String = "Generar_Ejecucion_Mensual_GG"

    Dim anioReporte As Long
    Dim archivoEjec As String
    Dim archivoCodiguera As String

    Dim wbEjec As Workbook
    Dim wbCod As Workbook
    Dim wsEjec As Worksheet
    Dim wsCod As Worksheet

    Dim dictLlaveACombo As Object
    Dim dictCombos As Object
    Dim dictAcumulado As Object

    Dim registrosLeidos As Long
    Dim llavesValidasCodiguera As Long
    Dim combinacionesGeneradas As Long

    Dim fechaInicio As Double
    fechaInicio = Timer

    On Error GoTo ErrHandler

    anioReporte = 2026 ' <<< Ajustar aquí el año del informe

    Set dictLlaveACombo = CreateObject("Scripting.Dictionary")
    Set dictCombos = CreateObject("Scripting.Dictionary")
    Set dictAcumulado = CreateObject("Scripting.Dictionary")

    archivoEjec = ObtenerArchivoMasReciente(RUTA_CARPETA_EJECUCIONES)
    If Len(archivoEjec) = 0 Then
        Err.Raise vbObjectError + 1000, NOMBRE_MACRO, _
                  "No se encontró archivo de ejecuciones en: " & RUTA_CARPETA_EJECUCIONES
    End If

    archivoCodiguera = ResolverArchivoCodiguera(RUTA_CODIGUERA)
    If Len(archivoCodiguera) = 0 Then
        Err.Raise vbObjectError + 1001, NOMBRE_MACRO, _
                  "No se encontró archivo de codiguera en: " & RUTA_CODIGUERA
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbEjec = Workbooks.Open(Filename:=archivoEjec, ReadOnly:=True)
    Set wsEjec = ObtenerPrimeraHojaConDatos(wbEjec)
    If wsEjec Is Nothing Then
        Err.Raise vbObjectError + 1002, NOMBRE_MACRO, _
                  "El archivo de ejecuciones no contiene hojas con datos."
    End If

    Set wbCod = Workbooks.Open(Filename:=archivoCodiguera, ReadOnly:=True)
    Set wsCod = ObtenerPrimeraHojaConDatos(wbCod)
    If wsCod Is Nothing Then
        Err.Raise vbObjectError + 1003, NOMBRE_MACRO, _
                  "El archivo de codiguera no contiene hojas con datos."
    End If

    LeerCodiguera wsCod, dictLlaveACombo, dictCombos, llavesValidasCodiguera

    If llavesValidasCodiguera = 0 Then
        Err.Raise vbObjectError + 1004, NOMBRE_MACRO, _
                  "No hay filas con Incluir_en_Informe = SI en la codiguera."
    End If

    LeerEjecucionesYAcumular wsEjec, anioReporte, dictLlaveACombo, dictAcumulado, registrosLeidos

    If dictAcumulado.Count = 0 Then
        Err.Raise vbObjectError + 1005, NOMBRE_MACRO, _
                  "No hay coincidencias entre llaves de ejecuciones y codiguera para el año " & anioReporte & "."
    End If

    VolcarResultado ThisWorkbook, anioReporte, dictCombos, dictAcumulado, combinacionesGeneradas

    wbEjec.Close SaveChanges:=False
    wbCod.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Proceso completado correctamente." & vbCrLf & vbCrLf & _
           "Archivo ejecuciones: " & NombreArchivoDesdeRuta(archivoEjec) & vbCrLf & _
           "Archivo codiguera: " & NombreArchivoDesdeRuta(archivoCodiguera) & vbCrLf & _
           "Registros leídos (ejecuciones): " & Format$(registrosLeidos, "#,##0") & vbCrLf & _
           "Llaves válidas (codiguera): " & Format$(llavesValidasCodiguera, "#,##0") & vbCrLf & _
           "Combinaciones Nivel_1/Nivel_2/Subtipo: " & Format$(combinacionesGeneradas, "#,##0") & vbCrLf & _
           "Tiempo (seg): " & Format$(Timer - fechaInicio, "0.00"), vbInformation

    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not wbEjec Is Nothing Then wbEjec.Close SaveChanges:=False
    If Not wbCod Is Nothing Then wbCod.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Error en " & NOMBRE_MACRO & ":" & vbCrLf & Err.Description, vbCritical
End Sub

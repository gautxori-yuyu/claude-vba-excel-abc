Attribute VB_Name = "modAPPFolderWatcher"
' =====================================================
' MÓDULO DE UTILIDADES Y GESTIÓN DEL FOLDERWATCHER
' Reemplaza la funcionalidad del VBScript fw.vbs
' =====================================================

'@Folder "5-FolderWatcher"
Option Explicit
' BUG: El objeto COM FolderWatcher Aparentemente se queda residente una vez se cierra Excel, Causando potencialmente problemas de memoria. Hay que corregirlo
' =====================================================
' FUNCIONES DE CONFIGURACIÓN RÁPIDA
' =====================================================

'@Description: Configura monitoreo de subcarpetas de oportunidades
'@Scope: Friend (solo clsAplicacion)
Public Sub ConfigurarMonitoreoOportunidades(ByVal rutaBase As String, ByVal fw As clsFolderWatch)
Attribute ConfigurarMonitoreoOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    
    If Not RutaExiste(rutaBase) Then
        Debug.Print "[modAPPFolderWatcher] ADVERTENCIA: Ruta de oportunidades no existe: " & rutaBase
        Exit Sub
    End If
    
    ' Monitorear SOLO subcarpetas (no archivos individuales)
    fw.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=IIf(IsNetworkPath(rutaBase), 15, 5), _
        foldersOnly:=True
    
    ' Configurar filtro para solo carpetas
    fw.ConfigurarFiltroSoloCarpetas rutaBase
    
    Debug.Print "[modAPPFolderWatcher] Monitoreo de oportunidades configurado: " & rutaBase
    Exit Sub
    
ErrHandler:
    Debug.Print "[modAPPFolderWatcher.ConfigurarMonitoreoOportunidades] ERROR: " & Err.Description
End Sub

'@Description: Configura monitoreo de archivos de plantillas
Public Sub ConfigurarMonitoreoPlantillas(ByVal rutaBase As String, ByVal fw As clsFolderWatch)
Attribute ConfigurarMonitoreoPlantillas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    
    If Not RutaExiste(rutaBase) Then
        Debug.Print "[modAPPFolderWatcher] ADVERTENCIA: Ruta de plantillas no existe: " & rutaBase
        Exit Sub
    End If
    
    ' Monitorear cambios en archivos Excel
    fw.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=True, _
        filterPattern:="*.xlsx;*.xlsm", _
        eventsToWatch:=Array("Changed", "Created"), _
        inactivityMinutes:=10, _
        foldersOnly:=False
    
    Debug.Print "[modAPPFolderWatcher] Monitoreo de plantillas configurado: " & rutaBase
    Exit Sub
    
ErrHandler:
    Debug.Print "[modAPPFolderWatcher.ConfigurarMonitoreoPlantillas] ERROR: " & Err.Description
End Sub

'@Description: Configura monitoreo de archivos Gas (C-GAS-ING)
Public Sub ConfigurarMonitoreoGasVBNet(ByVal rutaBase As String, ByVal fw As clsFolderWatch)
Attribute ConfigurarMonitoreoGasVBNet.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    
    If Not RutaExiste(rutaBase) Then
        Debug.Print "[modAPPFolderWatcher] ADVERTENCIA: Ruta de Gas no existe: " & rutaBase
        Exit Sub
    End If
    
    ' Monitorear cambios en archivos Excel
    fw.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=True, _
        filterPattern:="*.xlsx;*.xlsm", _
        eventsToWatch:=Array("Changed"), _
        inactivityMinutes:=10, _
        foldersOnly:=False
    
    Debug.Print "[modAPPFolderWatcher] Monitoreo de Gas configurado: " & rutaBase
    Exit Sub
    
ErrHandler:
    Debug.Print "[modAPPFolderWatcher.ConfigurarMonitoreoGasVBNet] ERROR: " & Err.Description
End Sub

' =====================================================
' FUNCIONES DE CONSULTA Y DIAGNÓSTICO
' =====================================================

'@Description: Muestra estadísticas de monitoreo en MessageBox
Public Sub VerEstadisticasMonitoreo()
Attribute VerEstadisticasMonitoreo.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Then
        MsgBox "Aplicación no inicializada", vbExclamation
        Exit Sub
    End If
    
    If App.FolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Obtener carpetas activas
    Dim carpetas As Variant
    carpetas = App.FolderWatcher.CarpetasActivas()
    
    If Not IsArray(carpetas) Then
        MsgBox "No hay carpetas monitoreadas", vbInformation
        Exit Sub
    End If
    
    If UBound(carpetas) < LBound(carpetas) Then
        MsgBox "No hay carpetas monitoreadas", vbInformation
        Exit Sub
    End If
    
    ' Mostrar estadísticas de cada carpeta
    Dim msg As String
    msg = "CARPETAS MONITOREADAS:" & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = LBound(carpetas) To UBound(carpetas)
        Dim stats As Variant
        stats = App.FolderWatcher.ObtenerEstadisticas(CStr(carpetas(i)))
        
        If IsArray(stats) And UBound(stats) >= 8 Then
            msg = msg & "[i] " & stats(0) & vbCrLf
            msg = msg & "   Total eventos: " & stats(1) & vbCrLf
            msg = msg & "   Creados: " & stats(2) & " | "
            msg = msg & "Eliminados: " & stats(3) & " | "
            msg = msg & "Modificados: " & stats(4) & vbCrLf
            msg = msg & "   Renombrados: " & stats(5) & " | "
            msg = msg & "Errores: " & stats(8) & vbCrLf
            msg = msg & "   Eventos/hora: " & Format(stats(7), "0.0") & vbCrLf
            msg = msg & "   Última actividad: " & Format(stats(6), "dd/mm/yyyy hh:mm:ss") & vbCrLf
            msg = msg & vbCrLf
        End If
    Next i
    
    MsgBox msg, vbInformation, "Estadísticas FolderWatcher"
    Exit Sub
    
ErrHandler:
    MsgBox "Error al obtener estadísticas: " & Err.Description, vbCritical
End Sub

'@Description: Genera hoja de Excel con historial de eventos
Public Sub VerHistorialMonitoreo(Optional ByVal lastMinutes As Long = 60)
Attribute VerHistorialMonitoreo.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Or App.FolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Obtener historial
    Dim history As Variant
    history = App.FolderWatcher.ObtenerHistorial(lastMinutes, "")
    
    If Not IsArray(history) Then
        MsgBox "No hay eventos en el historial", vbInformation
        Exit Sub
    End If
    
    If UBound(history, 1) < 0 Then
        MsgBox "No hay eventos en el historial", vbInformation
        Exit Sub
    End If
    
    ' Crear hoja con historial
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "FW_Historial_" & Format(Now, "hhmmss")
    
    ' Encabezados
    With ws
        .Range("A1:E1").Value = Array("Timestamp", "Evento", "Carpeta", "Archivo", "Tamaño (bytes)")
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        
        ' Datos
        Dim rowCount As Long
        rowCount = UBound(history, 1) - LBound(history, 1) + 1
        .Range("A2").Resize(rowCount, 5).Value = history
        
        ' Formato
        .Columns("A:A").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Columns("E:E").NumberFormat = "#,##0"
        .Columns("A:E").AutoFit
        .Columns("C:C").ColumnWidth = 50
        .Columns("D:D").ColumnWidth = 30
    End With
    
    MsgBox "Historial generado: " & rowCount & " eventos", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error al generar historial: " & Err.Description, vbCritical
End Sub

'@Description: Limpia el historial de eventos
Public Sub LimpiarHistorialMonitoreo()
Attribute LimpiarHistorialMonitoreo.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Or App.FolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Deseas limpiar el historial de eventos?", vbQuestion + vbYesNo) = vbYes Then
        App.FolderWatcher.LimpiarHistorial
        MsgBox "Historial limpiado", vbInformation
    End If
End Sub

'@Description: Muestra información de configuración del watcher
Public Sub VerConfiguracionWatcher()
Attribute VerConfiguracionWatcher.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Or App.FolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    Dim config As String
    config = App.FolderWatcher.Configuracion
    
    Dim carpetas As Variant
    carpetas = App.FolderWatcher.CarpetasActivas()
    
    Dim msg As String
    msg = "CONFIGURACIÓN DEL FOLDERWATCHER" & vbCrLf & vbCrLf
    msg = msg & config & vbCrLf & vbCrLf
    msg = msg & "CARPETAS ACTIVAS:" & vbCrLf
    
    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            msg = msg & "  • " & carpetas(i) & vbCrLf
        Next i
    Else
        msg = msg & "  (ninguna)" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Configuración FolderWatcher"
End Sub

' =====================================================
' FUNCIONES DE TEST Y DEBUGGING
' =====================================================

'@Description: Test de monitoreo de subcarpetas
Sub Test_MonitoreoSubcarpetas()
Attribute Test_MonitoreoSubcarpetas.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Then
        MsgBox "Aplicación no inicializada", vbExclamation
        Exit Sub
    End If
    
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFolderWatcher"
    
    ' Crear carpeta de test si no existe
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Configurar monitoreo
    App.FolderWatcher.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=5, _
        foldersOnly:=True
    
    App.FolderWatcher.ConfigurarFiltroSoloCarpetas rutaTest
    
    MsgBox "Monitoreo de test iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Ahora crea, renombra o elimina subcarpetas para probar.", vbInformation
End Sub

'@Description: Test de filtros de tamaño
Sub Test_FiltroTamaño()
Attribute Test_FiltroTamaño.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Then
        MsgBox "Aplicación no inicializada", vbExclamation
        Exit Sub
    End If
    
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFiltros"
    
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Monitorear solo archivos mayores a 1 KB
    App.FolderWatcher.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*.*", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    App.FolderWatcher.ConfigurarFiltroTamaño rutaTest, 1024, 999999999
    
    MsgBox "Monitoreo con filtro de tamaño iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Solo detectará archivos mayores a 1 KB", vbInformation
End Sub

'@Description: Test de detección de rutas de red
Sub Test_DeteccionRutasRed()
Attribute Test_DeteccionRutasRed.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim rutasPrueba As Variant
    rutasPrueba = Array( _
                  "C:\Windows", _
                  "Z:\Compartido\Docs", _
                  "\\servidor\publico", _
                  "\\192.168.1.100\files" _
                  )
    
    Dim i As Long, ruta As Variant
    For Each ruta In rutasPrueba
        Debug.Print ruta & " ? " & IIf(IsNetworkPath(CStr(ruta)), "RED", "LOCAL")
    Next
    
    MsgBox "Ver resultados en Ventana Inmediato (Ctrl+G)", vbInformation
End Sub

'@Description: Test de acción automática de mover
Sub Test_AccionMover()
Attribute Test_AccionMover.VB_ProcData.VB_Invoke_Func = " \n0"
    If App Is Nothing Then
        MsgBox "Aplicación no inicializada", vbExclamation
        Exit Sub
    End If
    
    Dim rutaOrigen As String, rutaDestino As String
    rutaOrigen = Environ("TEMP") & "\TestOrigen"
    rutaDestino = Environ("TEMP") & "\TestDestino"
    
    ' Crear carpetas si no existen
    If Not RutaExiste(rutaOrigen) Then MkDir rutaOrigen
    If Not RutaExiste(rutaDestino) Then MkDir rutaDestino
    
    ' Configurar monitoreo con acción de mover
    App.FolderWatcher.IniciarMonitoreo _
        folderPath:=rutaOrigen, _
        includeSubdirs:=False, _
        filterPattern:="*.txt", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    App.FolderWatcher.ConfigurarAccionMover rutaOrigen, rutaDestino
    
    MsgBox "Test de acción automática iniciado:" & vbCrLf & _
           "Origen: " & rutaOrigen & vbCrLf & _
           "Destino: " & rutaDestino & vbCrLf & vbCrLf & _
           "Los archivos .txt se moverán automáticamente", vbInformation
End Sub

'@Description: Test completo del sistema
Sub Test_SistemaCompleto()
Attribute Test_SistemaCompleto.VB_ProcData.VB_Invoke_Func = " \n0"
    Debug.Print "==== TEST COMPLETO FOLDERWATCHER ===="
    Debug.Print "Configuración: " & App.FolderWatcher.Configuracion
    
    Dim carpetas As Variant
    carpetas = App.FolderWatcher.CarpetasActivas()
    
    Debug.Print "Carpetas monitoreadas: " & IIf(IsArray(carpetas), UBound(carpetas) + 1, 0)
    
    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            Debug.Print "  " & i & ": " & carpetas(i)
            
            Dim stats As Variant
            stats = App.FolderWatcher.ObtenerEstadisticas(CStr(carpetas(i)))
            
            If IsArray(stats) Then
                Debug.Print "     Eventos: " & stats(1) & " | Última actividad: " & stats(6)
            End If
        Next i
    End If
    
    Debug.Print "==== FIN TEST ===="
End Sub

' =====================================================
' FUNCIONES DE UTILIDAD
' =====================================================
'@Description: Determina si una ruta es de red (UNC o unidad mapeada)
'@Scope: Private (uso interno del módulo)
'@ArgumentDescriptions: ruta | Ruta completa a verificar
'@Returns: Boolean | True si es ruta de red
Private Function IsNetworkPath(ByVal ruta As String) As Boolean
    On Error Resume Next
    
    IsNetworkPath = False
    
    ' Normalizar ruta
    If Right(ruta, 1) = "\" Then
        ruta = Left(ruta, Len(ruta) - 1)
    End If
    
    ' 1. Detectar rutas UNC (\\servidor\compartido)
    If Left(ruta, 2) = "\\" Then
        IsNetworkPath = True
        Exit Function
    End If
    
    ' 2. Detectar unidades de red mapeadas (Z:, Y:, etc.)
    If Mid(ruta, 2, 1) = ":" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim drive As Object
        Dim driveLetter As String
        driveLetter = Left(ruta, 2)              ' Ej: "C:", "Z:"
        
        ' Verificar si la unidad existe
        If fso.DriveExists(driveLetter) Then
            Set drive = fso.GetDrive(driveLetter)
            
            ' DriveType: 0=Unknown, 1=Removable, 2=Fixed, 3=Network, 4=CDRom, 5=RamDisk
            If drive.DriveType = 3 Then          ' 3 = Network
                IsNetworkPath = True
            End If
        End If
        
        Set drive = Nothing
        Set fso = Nothing
    End If
    
    On Error GoTo 0
End Function

'@Description: Normaliza una ruta eliminando la barra final
Private Function NormalizarRuta(ByVal ruta As String) As String
    If Right(ruta, 1) = "\" Then
        NormalizarRuta = Left(ruta, Len(ruta) - 1)
    Else
        NormalizarRuta = ruta
    End If
End Function

'@Description: Obtiene el nombre de una carpeta de su ruta completa
Private Function ObtenerNombreCarpeta(ByVal rutaCompleta As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ObtenerNombreCarpeta = fso.GetFileName(rutaCompleta)
    On Error GoTo 0
    
    Set fso = Nothing
End Function



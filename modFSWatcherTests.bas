Attribute VB_Name = "modFSWatcherTests"
'@Folder "2-Infraestructura.4-Servicios.FSWatcher"
'@ModuleDescription "Funciones de test y depuracion del FolderWatcher. Solo para uso en desarrollo."
Option Explicit

' =====================================================
' FUNCIONES DE TEST Y DEBUGGING
' Uso: pasar App.FSMonitoring.FolderWatcher como parametro
'      o crear un clsFSWatcher independiente para tests aislados
' =====================================================

'@Description: Test de monitoreo de subcarpetas.
'              Crea una carpeta temporal y configura un watcher sobre ella.
Sub TestMonitoreoSubcarpetas(fw As clsFSWatcher)
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFolderWatcher"

    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If

    fw.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=5, _
        foldersOnly:=True

    fw.ConfigurarFiltroSoloCarpetas rutaTest

    MsgBox "Monitoreo de test iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Ahora crea, renombra o elimina subcarpetas para probar.", vbInformation
End Sub

'@Description: Test de filtros de tamaño.
'              Solo detecta archivos mayores a 1 KB.
Sub TestFiltroTamanio(fw As clsFSWatcher)
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFiltros"

    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If

    fw.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*.*", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5

    fw.ConfigurarFiltroTamaño rutaTest, 1024, 999999999

    MsgBox "Monitoreo con filtro de tamaño iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Solo detectara archivos mayores a 1 KB", vbInformation
End Sub

'@Description: Test de deteccion de rutas de red.
'              Imprime clasificacion local/red en el log.
Sub TestDeteccionRutasRed()
    Const PROC_NAME As String = "TestDeteccionRutasRed"

    Dim rutasPrueba As Variant
    rutasPrueba = Array( _
                  "C:\Windows", _
                  "Z:\Compartido\Docs", _
                  "\\servidor\publico", _
                  "\\192.168.1.100\files" _
                  )

    Dim ruta As Variant
    For Each ruta In rutasPrueba
        LogInfo "modFSWatcherTests", "[TestDeteccionRutasRed] " & ruta & " -> " & IIf(IsNetworkPath(CStr(ruta)), "RED", "LOCAL")
    Next

    MsgBox "Ver resultados en Ventana Inmediato (Ctrl+G)", vbInformation
End Sub

'@Description: Test de accion automatica de mover.
'              Los archivos .txt creados en origen se mueven a destino automaticamente.
Sub TestAccionMover(fw As clsFSWatcher)
    Dim rutaOrigen As String, rutaDestino As String
    rutaOrigen = Environ("TEMP") & "\TestOrigen"
    rutaDestino = Environ("TEMP") & "\TestDestino"

    If Not RutaExiste(rutaOrigen) Then MkDir rutaOrigen
    If Not RutaExiste(rutaDestino) Then MkDir rutaDestino

    fw.IniciarMonitoreo _
        folderPath:=rutaOrigen, _
        includeSubdirs:=False, _
        filterPattern:="*.txt", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5

    fw.ConfigurarAccionMover rutaOrigen, rutaDestino

    MsgBox "Test de accion automatica iniciado:" & vbCrLf & _
           "Origen: " & rutaOrigen & vbCrLf & _
           "Destino: " & rutaDestino & vbCrLf & vbCrLf & _
           "Los archivos .txt se moveran automaticamente", vbInformation
End Sub

'@Description: Test completo del sistema. Vuelca estado del watcher en uso al log.
Sub TestSistemaCompleto(fw As clsFSWatcher)
    Const LOG_SRC As String = "modFSWatcherTests"

    LogInfo LOG_SRC, "[TestSistemaCompleto] ==== TEST COMPLETO FOLDERWATCHER ===="
    LogInfo LOG_SRC, "[TestSistemaCompleto] Configuracion: " & fw.Configuracion

    Dim carpetas As Variant
    carpetas = fw.CarpetasActivas()

    LogInfo LOG_SRC, "[TestSistemaCompleto] Carpetas monitoreadas: " & IIf(IsArray(carpetas), UBound(carpetas) + 1, 0)

    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            LogInfo LOG_SRC, "[TestSistemaCompleto] " & i & ": " & carpetas(i)

            Dim stats As Variant
            stats = fw.ObtenerEstadisticas(CStr(carpetas(i)))

            If IsArray(stats) Then
                LogInfo LOG_SRC, "[TestSistemaCompleto]     Eventos: " & stats(1) & " | Ultima actividad: " & stats(6)
            End If
        Next i
    End If

    LogInfo LOG_SRC, "[TestSistemaCompleto] FIN TEST"
End Sub

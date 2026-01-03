Attribute VB_Name = "modAPPFSWatcher"
' =====================================================
' MÓDULO DE UTILIDADES Y GESTIÓN DEL FOLDERWATCHER
' Reemplaza la funcionalidad del VBScript fw.vbs
' =====================================================
'
' NOTA SOBRE RECURSOS COM:
' El componente FolderWatcherCOM.dll (.NET Framework 4.0) usa FileSystemWatcher
' que puede quedarse residente si no se libera correctamente.
' Ver clsFolderWatch.cls para documentacion detallada y recomendaciones
' para el codigo VB.NET del COM.
'
' SOLUCION IMPLEMENTADA:
' - clsFolderWatch.Dispose() libera todos los recursos
' - clsAplicacion.Terminate() llama a Dispose antes de Set = Nothing
' =====================================================

'@Folder "2-Servicios.Archivos.Supervision"
Option Explicit

Private Const MODULE_NAME As String = "modAPPFolderWatcher"

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemTemplateProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: FileSystemTemplateProvider
' ==================================================================
' Gestiona la detección de cambios en ficheros de plantilla.
' Suscribe a clsFSMonitoringCoord para eventos TemplateCreated/Changed.
'
' Responsabilidades:
' - Detectar creación/modificación de plantillas en carpeta monitorizada
' - Validar que el fichero detectado es una plantilla válida
' - Re-emitir eventos de dominio (TemplateDetected/TemplateModified)
' ==================================================================

'@Folder "2-Infraestructura.5-2-Providers.FS"
Option Explicit

Private Const MODULE_NAME As String = "FileSystemTemplateProvider"

' ==================================================================
' DEPENDENCIAS
' ==================================================================
Private mFileManager As clsFileManager
Private mConfiguration As clsConfiguration
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord

' ==================================================================
' EVENTOS (notificación de cambios en plantillas)
' ==================================================================
Public Event TemplateDetected(ByVal folder As String, ByVal fileName As String)
Public Event TemplateModified(ByVal folder As String, ByVal fileName As String)

' ==================================================================
' INICIALIZACIÓN
' ==================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Set mFSMonitoringCoord = Nothing
    Set mFileManager = Nothing
    Set mConfiguration = Nothing
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'@Description: Inyecta dependencias
'@Param fm: Gestor de sistema de archivos
'@Param oConfiguration: Configuración de la aplicación (para ruta de plantillas)
'@Param oFSMonitoringCoord: Coordinador de monitoreo FS (para eventos en tiempo real)
Public Sub Initialize(fm As clsFileManager, _
                      oConfiguration As clsConfiguration, _
                      Optional oFSMonitoringCoord As clsFSMonitoringCoord = Nothing)
    Set mFileManager = fm
    Set mConfiguration = oConfiguration
    If Not oFSMonitoringCoord Is Nothing Then
        Set mFSMonitoringCoord = oFSMonitoringCoord
        LogDebug MODULE_NAME, "[Initialize] FSMonitoringCoord inyectado"
    End If
    LogDebug MODULE_NAME, "[Initialize] FileManager y Configuration inyectados"
End Sub

' ==================================================================
' PROPIEDADES PÚBLICAS
' ==================================================================

'@Description: Ruta base de plantillas configurada
Public Property Get BasePath() As String
    If Not mConfiguration Is Nothing Then
        BasePath = mConfiguration.RutaPlantillas
    End If
End Property

'@Description: Indica si la ruta de plantillas está disponible
Public Property Get IsReady() As Boolean
    IsReady = Len(BasePath) > 0
    If IsReady And Not mFileManager Is Nothing Then
        IsReady = mFileManager.ExisteCarpeta(BasePath)
    End If
End Property

' ==================================================================
' MANEJADORES DE EVENTOS clsFSMonitoringCoord
' Detecta cambios en ficheros de plantilla
' ==================================================================

Private Sub mFSMonitoringCoord_TemplateCreated(ByVal folder As String, ByVal fileName As String)
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[TemplateCreated] Plantilla detectada: " & fileName & " en " & folder
    RaiseEvent TemplateDetected(folder, fileName)
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mFSMonitoringCoord_TemplateCreated]"
End Sub

Private Sub mFSMonitoringCoord_TemplateChanged(ByVal folder As String, ByVal fileName As String)
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[TemplateChanged] Plantilla modificada: " & fileName & " en " & folder
    RaiseEvent TemplateModified(folder, fileName)
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mFSMonitoringCoord_TemplateChanged]"
End Sub

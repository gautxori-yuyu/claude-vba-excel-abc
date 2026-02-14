VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.4-Servicios.FSWatcher"
'@ModuleDescription "Registro de carpetas monitoreadas y clasificacion de rutas entrantes"
Option Explicit

Private Const MODULE_NAME As String = "clsFolderMgr"

' Enum: tipo semantico de carpeta configurada
Public Enum FolderType
    ftUnknown = 0
    ftOpportunities = 1
    ftTemplates = 2
    ftGasVBNet = 3
End Enum

' Diccionario: ruta_normalizada -> FolderType
' CompareMode = vbTextCompare -> claves insensibles a mayusculas
Private mFolderTypes As Object

' -------------------------------------------------------
' Ciclo de vida
' -------------------------------------------------------
Private Sub Class_Initialize()
    Set mFolderTypes = CreateObject("Scripting.Dictionary")
    mFolderTypes.CompareMode = 1 ' vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set mFolderTypes = Nothing
End Sub

' -------------------------------------------------------
' Initialize: carga el diccionario de configuracion
' @param oDicFolders: Dictionary CFG_RUTA_* -> ruta
' -------------------------------------------------------
Public Sub Initialize(ByVal oDicFolders As Object)
    Const PROC_NAME As String = "Initialize"
    On Error GoTo ErrHandler

    mFolderTypes.RemoveAll

    If oDicFolders Is Nothing Then Exit Sub

    If oDicFolders.Exists(CFG_RUTA_OPORTUNIDADES) Then
        RegisterPath oDicFolders(CFG_RUTA_OPORTUNIDADES), ftOpportunities
    End If
    If oDicFolders.Exists(CFG_RUTA_PLANTILLAS) Then
        RegisterPath oDicFolders(CFG_RUTA_PLANTILLAS), ftTemplates
    End If
    If oDicFolders.Exists(CFG_RUTA_GAS_VBNET) Then
        RegisterPath oDicFolders(CFG_RUTA_GAS_VBNET), ftGasVBNet
    End If

    LogInfo MODULE_NAME, "[Initialize] " & mFolderTypes.Count & " rutas registradas"
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
End Sub

' -------------------------------------------------------
' ClassifyPath: devuelve el tipo semantico de una ruta de evento
' Logica: comprueba si alguna ruta registrada es prefijo de 'folder'
' -------------------------------------------------------
Public Function ClassifyPath(ByVal folder As String) As FolderType
    ClassifyPath = ftUnknown

    Dim key As Variant
    For Each key In mFolderTypes.Keys
        If Len(key) > 0 Then
            If InStr(1, folder, CStr(key), vbTextCompare) > 0 Then
                ClassifyPath = mFolderTypes(key)
                Exit Function
            End If
        End If
    Next key
End Function

' -------------------------------------------------------
' GetPath: devuelve la ruta configurada para un tipo dado
' -------------------------------------------------------
Public Function GetPath(ByVal folderType As FolderType) As String
    Dim key As Variant
    For Each key In mFolderTypes.Keys
        If mFolderTypes(key) = folderType Then
            GetPath = CStr(key)
            Exit Function
        End If
    Next key
End Function

' -------------------------------------------------------
' RegisterPath: normaliza la ruta y la almacena
' -------------------------------------------------------
Private Sub RegisterPath(ByVal path As String, ByVal folderType As FolderType)
    If Len(path) = 0 Then Exit Sub

    ' Normalizar: eliminar barra final
    Dim normalized As String
    normalized = path
    If Right(normalized, 1) = "\" Or Right(normalized, 1) = "/" Then
        normalized = Left(normalized, Len(normalized) - 1)
    End If

    If mFolderTypes.Exists(normalized) Then
        mFolderTypes(normalized) = folderType
    Else
        mFolderTypes.Add normalized, folderType
    End If

    LogDebug MODULE_NAME, "[RegisterPath] " & normalized & " -> tipo " & folderType
End Sub

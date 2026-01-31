VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado relacionado con el archivo Excel activo
'@Category: State
'@Folder "5-Aplicac (Coord).Estado"
Option Explicit

Private Const MODULE_NAME As String = "clsFileState"

' ==========================================
' EVENTOS (para Mediadores)
' ==========================================

'@Description: Se dispara cuando cambia el archivo activo (solicitud de coordinación)
Public Event CurrentFileChanged(f As clsFileXLS, isOpportunityFile As Boolean)

' ==========================================
' ESTADO DE ARCHIVO
' ==========================================
Private mCurrentFile As clsFileXLS
Private mIsOpportunityFile As Boolean

' ==========================================
' INICIALIZACIÓN
' ==========================================

Private Sub Class_Initialize()
    Set mCurrentFile = Nothing
    mIsOpportunityFile = False
End Sub

Private Sub Class_Terminate()
    Set mCurrentFile = Nothing
End Sub

' ==========================================
' PROPIEDADES PÚBLICAS
' ==========================================

'@Description: Obtiene o establece el archivo Excel activo actual
Public Property Get CurrentFile() As clsFileXLS
    Set CurrentFile = mCurrentFile
End Property

Friend Property Set CurrentFile(f As clsFileXLS)
    Stop
    If f.Equals(mCurrentFile) Then Exit Property

    Set mCurrentFile = f

    ' Actualizar flag de archivo de oportunidad
    UpdateIsOpportunityFile f

    LogDebug MODULE_NAME, "[CurrentFile] Changed to: " & GetCurrentFileName() & _
                         ", IsOpportunity=" & mIsOpportunityFile

    ' Generar evento para MEDIADORES (solicitud de coordinación)
    RaiseEvent CurrentFileChanged(f, mIsOpportunityFile)
End Property

'@Description: Indica si el archivo actual es un archivo de oportunidad
Public Property Get isOpportunityFile() As Boolean
    isOpportunityFile = mIsOpportunityFile
End Property

' ==========================================
' LÓGICA INTERNA
' ==========================================

'@Description: Actualiza el flag IsOpportunityFile basándose en el archivo actual
Private Sub UpdateIsOpportunityFile(f As clsFileXLS)
    Stop
    If f Is Nothing Then
        mIsOpportunityFile = False
    Else
        On Error Resume Next
        mIsOpportunityFile = f.EsOportunidad
        If Err.Number <> 0 Then
            mIsOpportunityFile = False
        End If
        On Error GoTo 0
    End If
End Sub

' ==========================================
' CONSULTAS DE ESTADO
' ==========================================

'@Description: Indica si hay un archivo actualmente activo
Public Function HasCurrentFile() As Boolean
    HasCurrentFile = Not (mCurrentFile Is Nothing)
End Function

'@Description: Obtiene el nombre del archivo actual
'@Returns: String con el nombre del archivo o "(ninguno)"
Public Function GetCurrentFileName() As String
    If mCurrentFile Is Nothing Then
        GetCurrentFileName = "(ninguno)"
    Else
        On Error Resume Next
        GetCurrentFileName = mCurrentFile.Name
        If Err.Number <> 0 Then
            GetCurrentFileName = "(error)"
        End If
        On Error GoTo 0
    End If
End Function

'@Description: Obtiene el path completo del archivo actual
'@Returns: String con el path o cadena vacía
Public Function GetCurrentFilePath() As String
    If mCurrentFile Is Nothing Then
        GetCurrentFilePath = ""
    Else
        On Error Resume Next
        GetCurrentFilePath = mCurrentFile.Path
        If Err.Number <> 0 Then
            GetCurrentFilePath = ""
        End If
        On Error GoTo 0
    End If
End Function



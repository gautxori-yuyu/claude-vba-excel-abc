VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==================================================================
' INTERFAZ PARA ARCHIVOS TRACKED
' ==================================================================
' Define el contrato que deben cumplir todas las clases de archivo
' que puedan ser gestionadas por clsFileManager
' ==================================================================

'@Folder "4-Servicios.Archivos.Tipos de archivos"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IFile"

'@Description: Identificador único del archivo
'@Returns: Double - Puntero al objeto subyacente
Public Property Get ObjectKey() As Double
    ' Las interfaces en VBA deben lanzar error o no hacer nada
    Err.Raise 5, "IFile", "ObjectKey debe implementarse en clase concreta"
End Property

'@Description: Ruta completa del archivo
'@Returns: String - Path completo o "(cerrado)" si no disponible
Public Property Get Path() As String
    Err.Raise 5, "IFile", "Path debe implementarse en clase concreta"
End Property

'@Description: Nombre del archivo (sin ruta)
'@Returns: String - Nombre del archivo
Public Property Get Name() As String
    Err.Raise 5, "IFile", "Name debe implementarse en clase concreta"
End Property

' ==================================================================
' NOTAS DE IMPLEMENTACIÓN
' ==================================================================
' Para implementar esta interfaz en una clase:
'
' 1. Agregar al inicio de la clase:
'    Implements IFile
'
' 2. VBA generará automáticamente los stubs:
'    Private Property Get IFile_ObjectKey() As Double
'    Private Property Get IFile_Path() As String
'    Private Property Get IFile_Name() As String
'
' 3. Implementar cada método delegando a las propiedades públicas
'
' 4. Opcionalmente mantener las propiedades públicas originales

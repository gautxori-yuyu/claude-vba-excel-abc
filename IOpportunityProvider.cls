VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOpportunityProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' INTERFAZ DE DOMINIO: IOpportunityProvider
' ==================================================================
' Define el contrato para obtener oportunidades.
' El dominio usa esta interfaz sin conocer la implementacion concreta
' (sistema de archivos, base de datos, API, etc.)
'
' Implementada por: FileSystemOpportunityProvider (infraestructura)
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores"
Option Explicit

'@Description: Obtiene los nombres de las carpetas de oportunidad validas, ordenados para presentacion
'@Returns: Collection de String con nombres de oportunidades
Public Function GetOpportunityFolders() As Collection
End Function

'@Description: Verifica si una carpeta de oportunidad existe
'@Param folderName: Nombre de la carpeta (sin ruta base)
'@Returns: True si existe
Public Function FolderExists(ByVal folderName As String) As Boolean
End Function

'@Description: Crea una nueva carpeta de oportunidad desde plantilla
'@Param opportunityName: Nombre completo de la nueva oportunidad
'@Returns: True si se creo exitosamente
Public Function CreateOpportunityFolder(ByVal opportunityName As String) As Boolean
End Function

'@Description: Obtiene la ruta base donde se almacenan las oportunidades
'@Returns: Ruta completa del directorio base
Public Property Get BasePath() As String
End Property

'@Description: Construye la ruta completa de una oportunidad
'@Param folderName: Nombre de la carpeta de oportunidad
'@Returns: Ruta completa (BasePath + folderName)
Public Function GetFullPath(ByVal folderName As String) As String
End Function

'@Description: Genera el siguiente codigo de oportunidad disponible
'@Returns: Codigo de oportunidad (ej: "412602001")
Public Function GetNextOpportunityCode() As String
End Function

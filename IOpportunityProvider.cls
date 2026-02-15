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
' Define el contrato para obtener y gestionar oportunidades.
' El dominio usa esta interfaz sin conocer la implementacion concreta
' (sistema de archivos, base de datos, API, etc.)
'
' PRINCIPIO: Las firmas son agnosticas de infraestructura.
' No hay referencias a carpetas, rutas ni filesystem.
'
' Implementada por: FileSystemOpportunityProvider (capa 2 - Infraestructura)
' Consumida por:    clsOpportunitiesMgr (capa 3 - Dominio)
' ==================================================================

'@Folder "3-Dominio"
Option Explicit

'@Description: Obtiene los nombres de oportunidades disponibles, ordenados para presentacion
'@Returns: Collection de String con nombres de oportunidades
Public Function GetOpportunities() As Collection
End Function

'@Description: Verifica si una oportunidad con ese nombre existe
'@Param opportunityName: Nombre de la oportunidad (sin detalles de ruta)
'@Returns: True si existe
Public Function OpportunityExists(ByVal opportunityName As String) As Boolean
End Function

'@Description: Crea una nueva oportunidad con el nombre indicado
'@Param opportunityName: Nombre completo de la nueva oportunidad
'@Returns: True si se creo exitosamente
Public Function CreateOpportunity(ByVal opportunityName As String) As Boolean
End Function

'@Description: Genera el siguiente codigo de oportunidad disponible
'@Returns: Codigo de oportunidad (ej: "412602001")
Public Function GetNextOpportunityCode() As String
End Function

'@Description: Obtiene la ubicacion completa de una oportunidad a partir de su nombre
'@Param opportunityName: Nombre de la oportunidad
'@Returns: Ubicacion completa (implementacion concreta decide el formato)
Public Function GetOpportunityPath(ByVal opportunityName As String) As String
End Function


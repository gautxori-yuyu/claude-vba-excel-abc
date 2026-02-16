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

'@Description: Obtiene las oportunidades disponibles, ordenadas para presentacion
'@Returns: Collection de clsOpportunity con entidades ya construidas
Public Function GetOpportunities() As Collection
End Function

'@Description: Verifica si una oportunidad con ese nombre existe
'@Param opportunityName: Nombre de la oportunidad (sin detalles de ruta)
'@Returns: True si existe
Public Function OpportunityExists(ByVal opportunityName As String) As Boolean
End Function

'@Description: Crea una nueva oportunidad interactivamente (solicita datos al usuario,
'              valida, crea almacenamiento, construye entidad de dominio).
'@Returns: clsOpportunity creada, o Nothing si el usuario cancela o hay error
Public Function CreateNewOpportunity() As clsOpportunity
End Function

'@Description: Crea una oportunidad con el nombre indicado (sin interaccion de usuario)
'@Param opportunityName: Nombre completo de la nueva oportunidad
'@Returns: clsOpportunity creada, o Nothing si falla
Public Function CreateOpportunity(ByVal opportunityName As String) As clsOpportunity
End Function

'@Description: Genera el siguiente codigo de oportunidad disponible
'@Returns: Codigo de oportunidad (ej: "412602001")
Public Function GetNextOpportunityCode() As String
End Function

'@Description: Indica si el almacenamiento (carpeta base) esta disponible.
'              No re-consulta el filesystem; usa el flag inicializado al crear el proveedor.
'@Returns: True si la carpeta base era accesible al inicializar el proveedor
Public Function StorageAvailable() As Boolean
End Function

'@Description: Valida si un nombre cumple el patron de oportunidad valida
'@Param opportunityName: Nombre a validar
'@Returns: True si cumple el patron
Public Function ValidateOpportunityName(ByVal opportunityName As String) As Boolean
End Function

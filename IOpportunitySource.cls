VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOpportunitySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==================================================================
' INTERFAZ PARA FUENTE DE DATOS DE OPORTUNIDAD
' ==================================================================
' Define el contrato que deben cumplir todas las clases que provean
' acceso a datos de oportunidad, permitiendo que el dominio trabaje
' con oportunidades sin conocer la implementación concreta.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores.Interfaces"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IOpportunitySource"

'@Description: Lee los datos de la oportunidad actual
'@Returns: clsOpportunity - Objeto con los datos de la oportunidad
Public Function ReadOpportunityData() As clsOpportunity
    ' Las interfaces en VBA deben lanzar error o no hacer nada
    Err.Raise 5, "IOpportunitySource", "ReadOpportunityData debe implementarse en clase concreta"
End Function

'@Description: Lee datos técnicos de la oportunidad
'@Returns: Dictionary u objeto con datos técnicos
Public Function ReadTechnicalData() As Object
    Err.Raise 5, "IOpportunitySource", "ReadTechnicalData debe implementarse en clase concreta"
End Function

'@Description: Lee datos de cotización de la oportunidad
'@Returns: Dictionary u objeto con datos de cotización
Public Function ReadQuotationData() As Object
    Err.Raise 5, "IOpportunitySource", "ReadQuotationData debe implementarse en clase concreta"
End Function

'@Description: Guarda los resultados de cálculos o procesamiento
'@Param: resultado - Objeto con los resultados a guardar
Public Sub SaveResults(resultado As Object)
    Err.Raise 5, "IOpportunitySource", "SaveResults debe implementarse en clase concreta"
End Sub

'@Description: Valida si la fuente de datos es válida para la oportunidad
'@Returns: Boolean - True si la fuente es válida
Public Function IsValidSource() As Boolean
    Err.Raise 5, "IOpportunitySource", "IsValidSource debe implementarse en clase concreta"
End Function

'@Description: Obtiene el path o identificador de la oportunidad
'@Returns: String - Ruta o identificador único
Public Property Get OpportunityIdentifier() As String
    Err.Raise 5, "IOpportunitySource", "OpportunityIdentifier debe implementarse en clase concreta"
End Property

' ==================================================================
' NOTAS DE IMPLEMENTACIÓN
' ==================================================================
' Para implementar esta interfaz en una clase:
'
' 1. Agregar al inicio de la clase:
'    Implements IOpportunitySource
'
' 2. VBA generará automáticoamente los stubs con prefijo "IOpportunitySource_"
'
' 3. Implementar cada método delegando a la lógica específica de la clase
'
' 4. Usar esta interfaz para que el dominio no dependa de implementaciones concretas
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOpportunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==================================================================
' INTERFAZ DE DOMINIO: IOpportunity
' ==================================================================
' Define el contrato que el dominio necesita para trabajar con oportunidades.
' El dominio define QUE necesita; la infraestructura implementa COMO se obtiene.
'
' Esta interfaz sera implementada por adaptadores como ExcelOpportunitySource
' que traducen la tecnologia concreta (Excel) al contrato de dominio.
' ==================================================================

'@Folder "3-Dominio.Interfaces"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IOpportunity"

' ==================================================================
' IDENTIFICACION
' ==================================================================

'@Description: Identificador unico de la oportunidad
'@Returns: String - ID de la oportunidad (ej: "OP-2024-001234")
Public Property Get OpportunityId() As String
    Err.Raise 5, MODULE_NAME, "OpportunityId debe implementarse en clase concreta"
End Property

'@Description: Ruta base donde reside la oportunidad
'@Returns: String - Path de la carpeta o archivo principal
Public Property Get BasePath() As String
    Err.Raise 5, MODULE_NAME, "BasePath debe implementarse en clase concreta"
End Property

'@Description: Nombre descriptivo de la oportunidad
'@Returns: String - Nombre para mostrar en UI
Public Property Get DisplayName() As String
    Err.Raise 5, MODULE_NAME, "DisplayName debe implementarse en clase concreta"
End Property

' ==================================================================
' ESTADO Y VALIDACION
' ==================================================================

'@Description: Indica si la oportunidad es valida para operaciones
'@Returns: Boolean - True si la oportunidad tiene datos completos y consistentes
Public Function IsValid() As Boolean
    Err.Raise 5, MODULE_NAME, "IsValid debe implementarse en clase concreta"
End Function

'@Description: Indica si la oportunidad esta lista para generar oferta
'@Returns: Boolean - True si puede generarse una oferta
Public Function CanGenerateQuote() As Boolean
    Err.Raise 5, MODULE_NAME, "CanGenerateQuote debe implementarse en clase concreta"
End Function

' ==================================================================
' OPERACIONES DE LECTURA
' ==================================================================

'@Description: Lee los datos tecnicos de la oportunidad
'@Returns: Variant - Diccionario o array con datos tecnicos
'@Note: La estructura devuelta depende de la implementacion concreta
Public Function ReadTechnicalData() As Variant
    Err.Raise 5, MODULE_NAME, "ReadTechnicalData debe implementarse en clase concreta"
End Function

' ==================================================================
' NOTAS DE IMPLEMENTACION
' ==================================================================
' Para implementar esta interfaz en una clase de infraestructura:
'
' 1. Agregar al inicio de la clase:
'    Implements IOpportunity
'
' 2. VBA generara automaticamente los stubs:
'    Private Property Get IOpportunity_OpportunityId() As String
'    Private Property Get IOpportunity_BasePath() As String
'    Private Property Get IOpportunity_DisplayName() As String
'    Private Function IOpportunity_IsValid() As Boolean
'    Private Function IOpportunity_CanGenerateQuote() As Boolean
'    Private Function IOpportunity_ReadTechnicalData() As Variant
'
' 3. Implementar cada metodo delegando a la logica interna
'
' Ejemplo de implementacion (ExcelOpportunitySource):
'
'   Private mFileXLS As clsFileXLS
'
'   Private Property Get IOpportunity_OpportunityId() As String
'       IOpportunity_OpportunityId = mFileXLS.GetOpportunityId()
'   End Property
'
'   Private Function IOpportunity_IsValid() As Boolean
'       IOpportunity_IsValid = mFileXLS.EsOportunidad
'   End Function

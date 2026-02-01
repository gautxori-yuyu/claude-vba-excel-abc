VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clase desde la que se gestiona cada oportunidad, y los elementos (entidad compresor, ficheros, etc) que la constituyen

'@Folder "6-DOMINIO-Oportunidades y compresores"
Option Explicit
Implements IOpportunity

Private Const MODULE_NAME As String = "clsOpportunity"

' ==========================================
' ATRIBUTOS PRIVADOS
' ==========================================
Private mLabel As String        ' Nombre completo de la oportunidad (ej: "412601015 - Cliente - Proyecto")
Private mName As String         ' Nombre corto (ej: "412601015")
Private mPath As String         ' Ruta completa a la carpeta
Private mCustomer As String     ' Nombre del cliente
Private mQuoteNumber As String  ' N�mero de cotizaci�n
Private mCreatedDate As Date    ' Fecha de creaci�n

' Componentes futuros (no implementados a�n):
'Private mFiles As clsOpportunityFileCollection
'Private mCompressors As clsCompressorCollection
'Private mOffers As clsOfferCollection
'Private mDocuments As clsDocumentCollection
'Private mStatus As OpportunityStatus

' ==========================================
' INICIALIZACI�N
' ==========================================

Private Sub Class_Initialize()
    mLabel = ""
    mName = ""
    mPath = ""
    mCustomer = ""
    mQuoteNumber = ""
    mCreatedDate = 0
End Sub

Private Sub Class_Terminate()
    ' Limpiar referencias si necesario
End Sub

' ==========================================
' PROPIEDADES P�BLICAS
' ==========================================

'@Description: Obtiene o establece el label completo de la oportunidad
Public Property Get Label() As String
    Label = mLabel
End Property

Public Property Let Label(value As String)
    mLabel = value
End Property

'@Description: Obtiene o establece el nombre corto de la oportunidad
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(value As String)
    mName = value
End Property

'@Description: Obtiene o establece la ruta completa a la carpeta
Public Property Get Path() As String
    Path = mPath
End Property

Public Property Let Path(value As String)
    mPath = value
End Property

'@Description: Obtiene o establece el nombre del cliente
Public Property Get Customer() As String
    Customer = mCustomer
End Property

Public Property Let Customer(value As String)
    mCustomer = value
End Property

'@Description: Obtiene o establece el n�mero de cotizaci�n
Public Property Get QuoteNumber() As String
    QuoteNumber = mQuoteNumber
End Property

Public Property Let QuoteNumber(value As String)
    mQuoteNumber = value
End Property

'@Description: Obtiene o establece la fecha de creaci�n
Public Property Get CreatedDate() As Date
    CreatedDate = mCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    mCreatedDate = value
End Property

' ==========================================
' M�TODOS P�BLICOS
' ==========================================
Public Function Equals(op As clsOpportunity) As Boolean
    If op Is Nothing Then
        Equals = False
    Else
        Equals = (LCase(mPath) = LCase(op.Path))
    End If
End Function

'@Description: Valida si la oportunidad tiene datos m�nimos requeridos
Public Function IsValid() As Boolean
    IsValid = Len(mLabel) > 0 And Len(mPath) > 0
End Function

'@Description: Genera la estructura de carpetas para la oportunidad (pendiente implementaci�n)
Public Sub GenerateFolderStructure()
    ' TODO: Implementar creaci�n de estructura de carpetas
End Sub

'@Description: Obtiene la plantilla de presupuesto (pendiente implementaci�n)
Public Function GetBudgetTemplate() As clsFileXLS
    ' TODO: Implementar obtenci�n de plantilla
    Set GetBudgetTemplate = Nothing
End Function

'@Description: Sincroniza con el repositorio (pendiente implementacion)
Public Sub SyncWithRepository()
    ' TODO: Implementar sincronizacion
End Sub

' ==========================================
' IMPLEMENTACION DE IOpportunity
' ==========================================

Private Property Get IOpportunity_OpportunityId() As String
    IOpportunity_OpportunityId = mName
End Property

Private Property Get IOpportunity_BasePath() As String
    IOpportunity_BasePath = mPath
End Property

Private Property Get IOpportunity_DisplayName() As String
    IOpportunity_DisplayName = mLabel
End Property

Private Function IOpportunity_IsValid() As Boolean
    IOpportunity_IsValid = IsValid()
End Function

Private Function IOpportunity_CanGenerateQuote() As Boolean
    ' Por ahora, si es valida puede generar oferta
    IOpportunity_CanGenerateQuote = IsValid()
End Function

Private Function IOpportunity_ReadTechnicalData() As Variant
    ' Devolver datos basicos como diccionario
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("Name") = mName
    data("Label") = mLabel
    data("Path") = mPath
    data("Customer") = mCustomer
    data("QuoteNumber") = mQuoteNumber
    data("CreatedDate") = mCreatedDate
    Set IOpportunity_ReadTechnicalData = data
End Function

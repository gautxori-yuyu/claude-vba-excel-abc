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

Private Const MODULE_NAME As String = "clsOpportunity"

' ==========================================
' ATRIBUTOS PRIVADOS
' ==========================================
Private mLabel As String        ' Nombre completo de la oportunidad (ej: "412601015 - Cliente - Proyecto")
Private mName As String         ' Nombre corto (ej: "412601015")
Private mPath As String         ' Ruta completa a la carpeta
Private mCustomer As String     ' Nombre del cliente
Private mQuoteNumber As String  ' Número de cotización
Private mCreatedDate As Date    ' Fecha de creación

' Componentes futuros (no implementados aún):
'Private mFiles As clsOpportunityFileCollection
'Private mCompressors As clsCompressorCollection
'Private mOffers As clsOfferCollection
'Private mDocuments As clsDocumentCollection
'Private mStatus As OpportunityStatus

' ==========================================
' INICIALIZACIÓN
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
' PROPIEDADES PÚBLICAS
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

'@Description: Obtiene o establece el número de cotización
Public Property Get QuoteNumber() As String
    QuoteNumber = mQuoteNumber
End Property

Public Property Let QuoteNumber(value As String)
    mQuoteNumber = value
End Property

'@Description: Obtiene o establece la fecha de creación
Public Property Get CreatedDate() As Date
    CreatedDate = mCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    mCreatedDate = value
End Property

' ==========================================
' MÉTODOS PÚBLICOS
' ==========================================

'@Description: Valida si la oportunidad tiene datos mínimos requeridos
Public Function IsValid() As Boolean
    IsValid = Len(mLabel) > 0 And Len(mPath) > 0
End Function

'@Description: Genera la estructura de carpetas para la oportunidad (pendiente implementación)
Public Sub GenerateFolderStructure()
    ' TODO: Implementar creación de estructura de carpetas
End Sub

'@Description: Obtiene la plantilla de presupuesto (pendiente implementación)
Public Function GetBudgetTemplate() As clsFileXLS
    ' TODO: Implementar obtención de plantilla
    Set GetBudgetTemplate = Nothing
End Function

'@Description: Sincroniza con el repositorio (pendiente implementación)
Public Sub SyncWithRepository()
    ' TODO: Implementar sincronización
End Sub

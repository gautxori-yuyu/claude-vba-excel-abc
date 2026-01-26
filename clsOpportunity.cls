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

' Atributos básicos:
Public Name As String
Public Path As String
Public Customer As String
Public QuoteNumber As String
Public CreatedDate As Date
'Public Status As OpportunityStatus ' Enum

' Componentes:
'Public Files As clsOpportunityFileCollection ' Implementa IFile
'Public Compressors As clsCompressorCollection
'Public Offers As clsOfferCollection
'Public Documents As clsDocumentCollection

' Métodos:
Public Function IsValid() As Boolean

End Function
Public Sub GenerateFolderStructure()

End Sub
Public Function GetBudgetTemplate() As clsFileXLS

End Function
Public Sub SyncWithRepository()

End Sub

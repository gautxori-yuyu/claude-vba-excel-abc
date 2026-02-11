VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEconomicQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' INTERFAZ DE DOMINIO: IEconomicQuote
' ==================================================================
' Define el contrato para trabajar con UNA valoracion economica.
' El dominio usa esta interfaz sin conocer la implementacion concreta.
'
' Basado en la estructura identificada en VBScript:
' - QuoteNumber: Numero de oferta (9 digitos, ej: "412601015")
' - Revision: Numero de revision
' - Asociacion con calculos tecnicos
'
' Patron de nombre: \d{9}(?:[\-_]\d+)?(?:[ \-_]*rev\.?[ \-_]*\d+)?.*\.xls[xm]$
'
' Implementada por: ExcelEconomicQuoteSource (infraestructura)
' ==================================================================

'@Folder "3-Dominio"
Option Explicit

' ==================================================================
' PROPIEDADES DE IDENTIFICACION
' ==================================================================

'@Description: Numero de oferta/valoracion (9 digitos, ej: "412601015")
Public Property Get quoteNumber() As String
End Property

'@Description: Numero de compresor dentro de la oferta (ej: "01", "02")
Public Property Get CompressorIndex() As String
End Property

'@Description: Identificador completo (QuoteNumber + CompressorIndex si aplica)
Public Property Get quoteId() As String
End Property

'@Description: Numero de revision de la valoracion (0 si es original)
Public Property Get Revision() As Long
End Property

'@Description: Ruta completa del fichero de valoracion
Public Property Get FilePath() As String
End Property

'@Description: Nombre del fichero (sin ruta)
Public Property Get fileName() As String
End Property

' ==================================================================
' PROPIEDADES DE ESTADO
' ==================================================================

'@Description: Indica si la valoracion tiene datos validos
Public Property Get IsValid() As Boolean
End Property

'@Description: Indica si la valoracion ha sido modificada sin guardar
Public Property Get IsDirty() As Boolean
End Property

'@Description: Indica si es una valoracion nueva (sin guardar nunca)
Public Property Get IsNew() As Boolean
End Property

'@Description: Fecha de ultima modificacion
Public Property Get LastModified() As Date
End Property

' ==================================================================
' PROPIEDADES DE ASOCIACION
' ==================================================================

'@Description: Codigo del calculo tecnico asociado (si existe)
Public Property Get AssociatedCalcOption() As String
End Property

'@Description: Indica si tiene un calculo tecnico asociado
Public Property Get HasAssociatedCalc() As Boolean
End Property

' ==================================================================
' METODOS DE DATOS
' ==================================================================

'@Description: Obtiene los datos generales de la valoracion
'@Returns: Dictionary con datos (cliente, proyecto, fecha, etc.)
Public Function GetGeneralData() As Object
End Function

'@Description: Obtiene los datos de precios
'@Returns: Dictionary con precios y margenes
Public Function GetPricingData() As Object
End Function

'@Description: Obtiene el numero de "otros" (items adicionales)
Public Property Get OtrosCount() As Long
End Property

'@Description: Obtiene los datos de un item "otro"
'@Param index: Indice del item (base 0)
'@Returns: Dictionary con datos del item
Public Function GetOtroData(ByVal Index As Long) As Object
End Function

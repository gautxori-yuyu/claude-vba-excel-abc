VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==================================================================
' INTERFAZ DE DOMINIO: IOferta
' ==================================================================
' Define el contrato que el dominio necesita para trabajar con ofertas.
' El dominio define QUE necesita; la infraestructura implementa COMO se obtiene.
'
' Esta interfaz sera implementada por clases como clsOferta (dominio)
' y adaptadores de infraestructura como ExcelOfertaSource.
' ==================================================================

'@Folder "3-Dominio.Interfaces"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IOferta"

' ==================================================================
' IDENTIFICACION
' ==================================================================

'@Description: Identificador unico de la oferta
'@Returns: Long - ID de la oferta en base de datos (OFER_ID)
Public Property Get OfertaId() As Long
    Err.Raise 5, MODULE_NAME, "OfertaId debe implementarse en clase concreta"
End Property

'@Description: Numero de oferta para mostrar
'@Returns: String - Numero de oferta formateado (ej: "412601015-01")
Public Property Get NumeroOferta() As String
    Err.Raise 5, MODULE_NAME, "NumeroOferta debe implementarse en clase concreta"
End Property

'@Description: Fecha de la oferta
'@Returns: Date - Fecha de creacion/emision de la oferta
Public Property Get FechaOferta() As Date
    Err.Raise 5, MODULE_NAME, "FechaOferta debe implementarse en clase concreta"
End Property

' ==================================================================
' ESTADO Y VALIDACION
' ==================================================================

'@Description: Indica si la oferta tiene datos validos
'@Returns: Boolean - True si la oferta tiene datos completos y consistentes
Public Function IsValid() As Boolean
    Err.Raise 5, MODULE_NAME, "IsValid debe implementarse en clase concreta"
End Function

'@Description: Indica si la oferta tiene cambios pendientes de guardar
'@Returns: Boolean - True si hay modificaciones no persistidas
Public Property Get IsDirty() As Boolean
    Err.Raise 5, MODULE_NAME, "IsDirty debe implementarse en clase concreta"
End Property

'@Description: Indica si es una oferta nueva (no guardada aun en BD)
'@Returns: Boolean - True si es nueva
Public Property Get IsNew() As Boolean
    Err.Raise 5, MODULE_NAME, "IsNew debe implementarse en clase concreta"
End Property

' ==================================================================
' OPERACIONES DE LECTURA
' ==================================================================

'@Description: Obtiene los datos generales de la oferta
'@Returns: Variant - Estructura con datos generales
Public Function GetDatosGenerales() As Variant
    Err.Raise 5, MODULE_NAME, "GetDatosGenerales debe implementarse en clase concreta"
End Function

'@Description: Obtiene el numero de items "Otros" en la oferta
'@Returns: Long - Cantidad de items otros
Public Property Get OtrosCount() As Long
    Err.Raise 5, MODULE_NAME, "OtrosCount debe implementarse en clase concreta"
End Property

' ==================================================================
' NOTAS DE IMPLEMENTACION
' ==================================================================
' Para implementar esta interfaz en una clase:
'
' 1. Agregar al inicio de la clase:
'    Implements IOferta
'
' 2. VBA generara automaticamente los stubs:
'    Private Property Get IOferta_OfertaId() As Long
'    Private Property Get IOferta_NumeroOferta() As String
'    Private Property Get IOferta_FechaOferta() As Date
'    Private Function IOferta_IsValid() As Boolean
'    Private Property Get IOferta_IsDirty() As Boolean
'    Private Property Get IOferta_IsNew() As Boolean
'    Private Function IOferta_GetDatosGenerales() As Variant
'    Private Property Get IOferta_OtrosCount() As Long
'
' 3. Implementar cada metodo delegando a la logica interna

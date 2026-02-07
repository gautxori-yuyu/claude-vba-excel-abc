VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOfertaOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "6-DOMINIO-Oportunidades y compresores.d-Ofertas.Gestion"
Option Explicit

Private Const MODULE_NAME As String = "clsOfertaOtro"

Private m_OFOT_LINEA As Long
Private m_OFOT_DESCRIPCION As String
Private m_OFOT_PRE_COSTE As Double

Public Property Get OFOT_LINEA() As Long
    OFOT_LINEA = m_OFOT_LINEA
End Property
Public Property Let OFOT_LINEA(ByVal Value As Long)
    m_OFOT_LINEA = Value
End Property

Public Property Get OFOT_DESCRIPCION() As String
    OFOT_DESCRIPCION = m_OFOT_DESCRIPCION
End Property
Public Property Let OFOT_DESCRIPCION(ByVal Value As String)
    m_OFOT_DESCRIPCION = Value
End Property

Public Property Get OFOT_PRE_COSTE() As Double
    OFOT_PRE_COSTE = m_OFOT_PRE_COSTE
End Property
Public Property Let OFOT_PRE_COSTE(ByVal Value As Double)
    m_OFOT_PRE_COSTE = Value
End Property


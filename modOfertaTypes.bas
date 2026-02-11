Attribute VB_Name = "modOfertaTypes"
'@Folder "3-Dominio.d-Ofertas.Gestion"
Option Explicit

Private Const MODULE_NAME As String = "modOfertaTypes"

Public Type tOfertasDatosGenerales
    OFER_ID As String
    OFER_NUM_OFERTA As String
    OFER_FECHA As Date
    OFER_CLIENTE As String
    GASE_ID As Long
    OFER_OBSERVACIONES As String
End Type


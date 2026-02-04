VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "6-DOMINIO-Oportunidades y compresores.d-Ofertas.Gestion"
Option Explicit

Private Const MODULE_NAME As String = "clsOferta"

Private m_DatosGenerales As tOfertasDatosGenerales
Private m_Otros As Collection
'TODO: implementar mÃ©todos para asignar y usar estos campos, de mantenimiento del estado de un registro (Â¿se ha cambiado? / Â¿se ha escrito en db?)
Private m_IsDirty As Boolean
Private m_IsNew As Boolean
'TODO: implementar validaciones de los campos a asignar (OFER_FECHA no nula, GASE_ID vÃ¡lido, OFOT_LINEA incremental), antes de escribir a la db
'TODO:Lectura completa por JOIN, no N recordsets (Como en formulario de Access); Un SQL maestro + subrecordsets en memoria
'TODO:CatÃ¡logos cacheados: GASE_ID -> GASE_DESCRIPCION; Diccionarios cargados una vez
'TODO:Escritura controlada: Crear oferta desde Excel / Insert + update selectivo / Sin borrar y recrear salvo colecciones
'TODO:Control de concurrencia
'TODO:Logs de operaciones
'TODO:ValidaciÃ³n previa a escritura
'TODO:Eventual multiusuario

Private Sub Class_Initialize()
    Set m_Otros = New Collection
End Sub

Public Property Let DatosGenerales(dg As tOfertasDatosGenerales)
    On Error GoTo ErrHandler

    m_DatosGenerales = dg
    Exit Property
ErrHandler:
    Err.Raise Err.Number, , "Error al DatosGenerales: " & Err.Description
End Property

Public Property Get DatosGenerales() As tOfertasDatosGenerales
    DatosGenerales = m_DatosGenerales
End Property

Public Property Get Otros() As Collection
    Set Otros = m_Otros
End Property

Public Sub AgregarOtro(ByVal Otro As clsOfertaOtro)
    m_Otros.Add Otro
End Sub


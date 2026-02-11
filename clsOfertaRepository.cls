VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOfertaRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3-Dominio.d-Ofertas.Gestion"
Option Explicit

Private Const MODULE_NAME As String = "clsOfertaRepository"

Private m_Ctx As clsDBContext

Public Sub SetDBContext(ByVal ctx As clsDBContext)
    Set m_Ctx = ctx
End Sub
'=========================================================
'@Description: Lee todas las ofertas del repositorio
'@Scope: Acceso a datos (lectura masiva)
'@ArgumentDescriptions: -
'@Returns: Collection de clsOferta
'@Category: Repositorio
'=========================================================

Public Function LeerTodas() As Collection
    On Error GoTo ErrHandler

    Dim col As New Collection
    Dim rs As Object
    Dim Sql As String
    Dim of As clsOferta
    Dim OferID As String

    ' 1 Leer tabla maestra
    Sql = "SELECT [OFER_ID] FROM [OfertasDatosGenerales] ORDER BY [OFER_FECHA], [OFER_NUM_OFERTA]"
    Set rs = m_Ctx.AbrirRecordset(Sql)

    Do While Not rs.EOF
        OferID = NormalizarGUID(rs("OFER_ID"))
        ' 2 Reutilizamos el método ya probado
        Set of = LeerPorOferID(OferID)

        col.Add of
        rs.MoveNext
    Loop

    Set LeerTodas = col
    rs.Close
    
    Exit Function
ErrHandler:
    rs.Close
    Err.Raise Err.Number, , "Error al leer todas las ofertas: " & Err.Description
End Function

Public Function LeerPorOferID(ByVal OferID As String) As clsOferta
    Select Case True
        Case OferID = "", Not isGUID(OferID)
            Err.Raise Err.Number, , "Error LeerPorOferID. OferID inválido."
            Exit Function
    End Select
    
    On Error GoTo ErrHandler

    Dim of As New clsOferta
    CargarDatosGenerales of, OferID
    CargarOtros of, OferID
    Set LeerPorOferID = of
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, , "Error al LeerPorOferID: " & Err.Description
End Function

Private Sub CargarDatosGenerales(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler

    Dim rs As Object
    Dim dg As tOfertasDatosGenerales

    If Not m_Ctx.EstaConectado Then
            Err.Raise Err.Number, , "Error ADO, no hay conexion a la DB."
            Exit Sub
    End If
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT OFER_ID, OFER_NUM_OFERTA, OFER_FECHA, GASE_ID, OFER_CLIENTE, OFER_OBSERVACIONES " & _
                "FROM OfertasDatosGenerales WHERE OFER_ID={" & OferID & "}")

    If rs.EOF Then Err.Raise 5, , "Oferta no encontrada"

    With dg
        .OFER_ID = NormalizarGUID(rs("OFER_ID"))
        LogDebug MODULE_NAME, "[CargarDatosGenerales] OFER_ID: " & .OFER_ID
        .OFER_NUM_OFERTA = NzVBA(rs("OFER_NUM_OFERTA"))
        .OFER_FECHA = NzVBA(rs("OFER_FECHA"))
        .OFER_CLIENTE = NzVBA(rs("OFER_CLIENTE"))
        .GASE_ID = NzVBA(rs("GASE_ID"))
        .OFER_OBSERVACIONES = NzVBA(rs("OFER_OBSERVACIONES"))
    End With

    Oferta.DatosGenerales = dg
    rs.Close
    Exit Sub
ErrHandler:
    rs.Close
    Err.Raise Err.Number, , "Error al CargarDatosGenerales: " & Err.Description
End Sub

Private Sub CargarOtros(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler

    Dim rs As Object

    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasOtros WHERE OFER_ID={" & OferID & "} ORDER BY OFOT_LINEA")

    Do While Not rs.EOF
        Dim it As New clsOfertaOtro
        it.OFOT_LINEA = IIf(IsNull(rs("OFOT_LINEA")), 0, rs("OFOT_LINEA"))
        it.OFOT_DESCRIPCION = NzVBA(rs("OFOT_DESCRIPCION"))
        it.OFOT_PRE_COSTE = IIf(IsNull(rs("OFOT_PRE_COSTE")), 0, rs("OFOT_PRE_COSTE"))
        Oferta.AgregarOtro it
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
ErrHandler:
    rs.Close
    Err.Raise Err.Number, , "Error al CargarOtros: " & Err.Description
End Sub

Private Function NormalizarGUID(ByVal v As Variant) As String
    NormalizarGUID = Replace(Replace(v, "{", ""), "}", "")
End Function

Private Function NzVBA(ByVal v As Variant) As String
    On Error GoTo ErrHandler
    If IsNull(v) Then
        NzVBA = ""
    Else
        NzVBA = v
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, , "Error NzVBA: " & Err.Description
End Function



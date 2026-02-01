VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "4-Servicios.DBs"
Option Explicit

Private Const MODULE_NAME As String = "clsDBContext"

Private m_Cnn As Object ' ADODB.Connection

Public Sub Conectar(ByVal RutaBD As String)
    Select Case True
        Case Len(RutaBD) = 0, Dir(RutaBD) = ""
            Err.Raise Err.Number, , _
                "Error ADO al conectar a DB." & vbCrLf & _
                "Path a la DB inexistente:" & vbCrLf & RutaBD
            Exit Sub
    End Select
    
    Set m_Cnn = CreateObject("ADODB.Connection")
'    If Not EstaConectado Then
'            Err.Raise Err.Number, , "Error ADO, no hay conexion a la DB."
'            Exit Sub
'    End If
    On Error GoTo ErrorHandler
    m_Cnn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & RutaBD
    On Error GoTo 0
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , _
        "Error ADO al abrir conexión a DB: " & vbCrLf & _
        Err.Description
End Sub
Public Function EstaConectado() As Boolean
    EstaConectado = Not m_Cnn Is Nothing And m_Cnn.State = 1
End Function

Public Sub Desconectar()
    If Not m_Cnn Is Nothing Then
        If m_Cnn.State = 1 Then m_Cnn.Close
        Set m_Cnn = Nothing
    End If
End Sub

Public Function AbrirRecordset(ByVal Sql As String) As Object
    Select Case True
        Case Sql = ""
            Err.Raise Err.Number, , _
                "Error ADO al abrir Recordset." & vbCrLf & _
                "Function SQL inexistente"
            Exit Function
    End Select
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    If Not EstaConectado Then
            Err.Raise Err.Number, , "Error ADO, no hay conexion a la DB."
            Exit Function
    End If

    On Error GoTo ErrHandler
    rs.Open Sql, m_Cnn, 1, 1   ' adOpenKeyset, adLockReadOnly
    Set AbrirRecordset = rs
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[AbrirRecordset] - SQL ejecutado: " & Sql
    Err.Raise Err.Number, , _
        "Error ADO al abrir Recordset." & vbCrLf & _
        "SQL ejecutado:" & vbCrLf & Sql & vbCrLf & _
        Err.Description
End Function

Public Sub BeginTrans()
    m_Cnn.BeginTrans
End Sub

Public Sub CommitTrans()
    m_Cnn.CommitTrans
End Sub

Public Sub RollbackTrans()
    m_Cnn.RollbackTrans
End Sub


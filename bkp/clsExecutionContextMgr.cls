VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura"
'@Description: Gestor de contexto de ejecuciÛn que traduce eventos tÈcnicos a eventos sem·nticos de infraestructura
'@Note: Centraliza la lÛgica de traducciÛn de eventos tÈcnicos a sem·nticos de infraestructura
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextMgr"

' ==========================================
' EVENTOS SEM√ÅNTICOS DE INFRAESTRUCTURA - Emite solo este componente
' ==========================================
Public Event WorkbookSessionStarted(ByVal Path As String, ByVal readOnly As Boolean)
Public Event WorkbookSessionActivated(ByVal Path As String)
Public Event WorkbookSessionClosing(ByVal Path As String, ByRef Cancel As Boolean)
Public Event SheetSessionActivated(ByVal sheetName As String, ByVal sheetType As String)
Public Event SheetSessionDeactivated(ByVal sheetName As String, ByVal sheetType As String)
Public Event WindowSessionActivated(ByVal windowTitle As String)
Public Event WindowSessionDeactivated(ByVal windowTitle As String)
Public Event SelectionChangedInSession(ByVal rangeAddress As String)
Public Event SheetCalculatedInSession(ByVal sheetName As String)

' ==========================================
' COMPONENTES DE GESTI”N
' ==========================================
Private m_state As clsExecutionContextState

' ==========================================
' INICIALIZACI”N
' ==========================================
Public Sub Initialize(State As clsExecutionContextState)
    Set m_state = State
    LogInfo MODULE_NAME, "[Initialize] - ExecutionContextMgr inicializado"
End Sub

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
    Set m_state = Nothing
End Sub

' ==========================================
' M…TODOS PARA PROCESAR EVENTOS T…CNICOS (llamados desde ThisWorkbook)
' ==========================================

Public Sub OnWorkbookOpened(Wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookOpened] " & Wb.Name
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveWorkbook Wb
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionStarted(Wb.FullName, Wb.readOnly)
End Sub

Public Sub OnWorkbookActivated(Wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookActivated] " & Wb.Name
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveWorkbook Wb
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionActivated(Wb.FullName)
End Sub

Public Sub OnWorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean)
    LogDebug MODULE_NAME, "[OnWorkbookBeforeClose] " & Wb.Name
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionClosing(Wb.FullName, Cancel)
    
    ' Si no se cancela, limpiar estado
    If Not Cancel Then
        If m_state.IsWorkbookValid(Wb) Then
            ' Si es el workbook activo, limpiar estado
            If m_state.ActiveWorkbook Is Wb Then
                m_state.SetActiveWorkbook Nothing
            End If
        End If
    End If
End Sub

Public Sub OnSheetActivated(Sh As Object)
    LogDebug MODULE_NAME, "[OnSheetActivated] " & TypeName(Sh) & " '" & Sh.Name & "'"
    
    ' Actualizar estado tÈcnico
    If TypeOf Sh Is Worksheet Then
        m_state.SetActiveWorksheet Sh
    End If
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetSessionActivated(Sh.Name, TypeName(Sh))
End Sub

Public Sub OnSheetDeactivated(Sh As Object)
    LogDebug MODULE_NAME, "[OnSheetDeactivated] '" & Sh.Name & "'"
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetSessionDeactivated(Sh.Name, TypeName(Sh))
    
    ' Si es la hoja activa, limpiar estado
    If m_state.ActiveWorksheet Is Sh Then
        m_state.SetActiveWorksheet Nothing
    End If
End Sub

Public Sub OnWindowActivated(Wn As Window)
    LogDebug MODULE_NAME, "[OnWindowActivated] " & Wn.Caption
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WindowSessionActivated(Wn.Caption)
End Sub

Public Sub OnWindowDeactivated(Wn As Window)
    LogDebug MODULE_NAME, "[OnWindowDeactivated] " & Wn.Caption
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WindowSessionDeactivated(Wn.Caption)
End Sub

Public Sub OnSelectionChanged(Target As range)
    LogDebug MODULE_NAME, "[OnSelectionChanged] " & Target.Address
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveSelection Target
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SelectionChangedInSession(Target.Address)
End Sub

Public Sub OnSheetCalculated(Sh As Object)
    LogDebug MODULE_NAME, "[OnSheetCalculated] " & Sh.Name
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetCalculatedInSession(Sh.Name)
End Sub

' ==========================================
' M…TODOS DE ACCESO AL ESTADO
' ==========================================
Public Function GetState() As clsExecutionContextState
    Set GetState = m_state
End Function

' ==========================================
' M…TODOS DE UTILIDAD
' ==========================================
Public Function GetCurrentContextInfo() As String
    Dim info As String
    info = "Context Info:" & vbCrLf
    If Not m_state Is Nothing Then
        info = info & m_state.Diagnostics()
    Else
        info = info & "State not initialized"
    End If
    GetCurrentContextInfo = info
End Function

' ==========================================
' M…TODOS DE CONTROL DE CICLO DE VIDA
' ==========================================
Public Sub ResetState()
    If Not m_state Is Nothing Then
        m_state.ClearState
    End If
End Sub

Public Sub ReconnectEvents()
    ' MÈtodo para reconexiÛn de eventos en caso de reset
    LogInfo MODULE_NAME, "[ReconnectEvents] - Intentando reconexiÛn de eventos"
    ' La reconexiÛn real se har√≠a en ThisWorkbook autom·ticamente
End Sub

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
Public Event WorkbookSessionStarted(ByVal path As String, ByVal readOnly As Boolean)
Public Event WorkbookSessionActivated(ByVal path As String)
Public Event WorkbookSessionClosing(ByVal path As String, ByRef Cancel As Boolean)
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
Public Sub Initialize(state As clsExecutionContextState)
    Set m_state = state
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

Public Sub OnWorkbookOpened(wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookOpened] " & wb.Name
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveWorkbook wb
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionStarted(wb.FullName, wb.ReadOnly)
End Sub

Public Sub OnWorkbookActivated(wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookActivated] " & wb.Name
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveWorkbook wb
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionActivated(wb.FullName)
End Sub

Public Sub OnWorkbookBeforeClose(wb As Workbook, ByRef Cancel As Boolean)
    LogDebug MODULE_NAME, "[OnWorkbookBeforeClose] " & wb.Name
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent WorkbookSessionClosing(wb.FullName, Cancel)
    
    ' Si no se cancela, limpiar estado
    If Not Cancel Then
        If m_state.IsWorkbookValid(wb) Then
            ' Si es el workbook activo, limpiar estado
            If m_state.ActiveWorkbook Is wb Then
                m_state.SetActiveWorkbook Nothing
            End If
        End If
    End If
End Sub

Public Sub OnSheetActivated(sh As Object)
    LogDebug MODULE_NAME, "[OnSheetActivated] " & TypeName(sh) & " '" & sh.Name & "'"
    
    ' Actualizar estado tÈcnico
    If TypeOf sh Is Worksheet Then
        m_state.SetActiveWorksheet sh
    End If
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetSessionActivated(sh.Name, TypeName(sh))
End Sub

Public Sub OnSheetDeactivated(sh As Object)
    LogDebug MODULE_NAME, "[OnSheetDeactivated] '" & sh.Name & "'"
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetSessionDeactivated(sh.Name, TypeName(sh))
    
    ' Si es la hoja activa, limpiar estado
    If m_state.ActiveWorksheet Is sh Then
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

Public Sub OnSelectionChanged(Target As Range)
    LogDebug MODULE_NAME, "[OnSelectionChanged] " & Target.Address
    
    ' Actualizar estado tÈcnico
    m_state.SetActiveSelection Target
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SelectionChangedInSession(Target.Address)
End Sub

Public Sub OnSheetCalculated(sh As Object)
    LogDebug MODULE_NAME, "[OnSheetCalculated] " & sh.Name
    
    ' Emitir evento sem·ntico de infraestructura
    RaiseEvent SheetCalculatedInSession(sh.Name)
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
        m_state.ClearState()
    End If
End Sub

Public Sub ReconnectEvents()
    ' MÈtodo para reconexiÛn de eventos en caso de reset
    LogInfo MODULE_NAME, "[ReconnectEvents] - Intentando reconexiÛn de eventos"
    ' La reconexiÛn real se har√≠a en ThisWorkbook autom·ticamente
End Sub 
  
' =========================================='  
' MTODO DE CONSULTA DE ESTADO DE GRFICO'  
' =========================================='  
Public Function GetCachedChartInfo() As T_CachedChartInfo  
    GetCachedChartInfo = m_state.CachedChartInfo  
End Function 

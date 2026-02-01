VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "1-Inicio e Instalacion.Ciclo de vida"
Option Explicit

Private Const MODULE_NAME As String = "ThisWorkbook"

Private WithEvents mXLApp As Application
Attribute mXLApp.VB_VarHelpID = -1

Private mExecutionContextMgr As clsExecutionContextMgr

Private mApp As clsApplication

Friend Function App() As clsApplication
    If mApp Is Nothing Then
        Set mApp = New clsApplication
    End If
    Set App = mApp
End Function

' M�todo para que la aplicaci�n pueda registrar el gestor de contexto
Friend Sub SetExecutionContextMgr(executionContextMgr As clsExecutionContextMgr)
    Set mExecutionContextMgr = executionContextMgr
End Sub

Public Sub TerminateApp()
    If Not mApp Is Nothing Then
        Set mApp = Nothing
    End If
End Sub

Private Sub Workbook_Open()
    ' 0. Inicializar logger con salida a fichero
    removeOldLogs
    InitLogger LOG_DEBUG, True, ThisWorkbook.Path & "\ABC_Log_" & Format(Now, "yyyy-mm-dd_hh-nn") & ".txt"
    
    LogInfo MODULE_NAME, "[Open] - Abriendo Workbook " & ThisWorkbook.Name
    ' Sistema de instalaci�n existente
    AutoInstalador
    AutoRegistrarTodasLasUDFs
    ' Ctrl + Shift + R
    Application.OnKey "^+R", "ToggleRibbonTab"

    ' Iniciar aplicaci�n (lazy init del singleton)
    App
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Este evento NO debe desregistrar las UDFs normalmente,
    ' solo si detecta que el complemento fue deshabilitado manualmente

    On Error GoTo ErrHandler

    ' Si fue deshabilitado pero no se disparo AddinUninstall, limpiar
    If Not ComprobarSiInstalado Then
        LogWarning MODULE_NAME, "[Workbook_BeforeClose] - Complemento no instalado, realizando limpieza"
    End If

    ' Limpiar atajos de teclado
    Application.OnKey "^+R"

    ' Terminar aplicacion de forma segura
    TerminateApp

    ' Desregistrar UDFs
    DesregistrarTodasLasUDFs
    LogInfo MODULE_NAME, "[Workbook_BeforeClose] - UDFs desregistradas"
    
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Workbook_BeforeClose]"
    ' Intentar limpieza minima en caso de error
    On Error Resume Next
    Application.OnKey "^+R"
    TerminateApp
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Regenerar listas autom�ticamente antes de guardar
    'Call GenerarListasUnidades
End Sub

Private Sub Workbook_AddinUninstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario desmarca la casilla en Archivo > Opciones > Complementos
    ' 2. Se ejecuta ai.Installed = False desde el script VBS
    
    On Error Resume Next
    
    ' Mensaje de despedida (opcional)
    '    MsgBox "Complemento desinstalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas han sido eliminadas.", vbInformation, "Desinstalaci�n completada"
    
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario marca la casilla en Archivo > Opciones > Complementos
    ' 2. El script VBS completa la instalaci�n y marca installed=True
    
    On Error Resume Next
    
    ' Mensaje de bienvenida (opcional, puedes comentarlo si es molesto)
    '    MsgBox "Complemento 'ABC Ofertas M�quina Especial' instalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas ya est�n disponibles.", vbInformation, "Instalaci�n completada"
    
    On Error GoTo 0
End Sub
' Elimina los ficheros de log que tengan M�s de 3 d�as, si hay M�s de 3 ficheros de log
Private Sub removeOldLogs()
    Dim fso As Object
    Dim folderPath As String
    Dim oFolder As Object
    Dim oFile As Object
    Dim logFiles As Object ' Collection para filtrar solo logs de este tipo
    Dim fileDate As Date
    Dim fileNamePattern As String, i
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.Path & "\"
    fileNamePattern = "ABC_Log_????-??-??_??-??.txt"
    
    ' 1. Coleccionamos los archivos que coinciden con el patr�n
    Set logFiles = New Collection
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set oFolder = fso.GetFolder(folderPath)
    
    For Each oFile In oFolder.Files
        ' Verificamos si el nombre coincide con el patr�n ABC_Log_...
        If oFile.Name Like fileNamePattern Then
            If logFiles.Count = 0 Then
                logFiles.Add oFile
            ElseIf (oFile.DateLastAccessed < logFiles.Item(logFiles.Count).DateLastAccessed) Then
                logFiles.Add Item:=oFile, before:=logFiles.Count
            Else
                logFiles.Add oFile
            End If
        End If
    Next oFile
    
    ' 2. Si hay M�s de 3 ficheros en total, evaluamos el borrado
    Do While logFiles.Count > 3
        logFiles.Item(1).Delete True
        logFiles.Remove 1
    Loop
    For Each oFile In logFiles
        ' Calculamos la diferencia de d�as (DateDiff)
        ' Usamos DateLastModified para mayor precisi�n, o DateCreated
        fileDate = oFile.DateLastModified
        
        ' Requisito: M�s de 3 d�as de antiguedad
        If DateDiff("d", fileDate, Now) > 3 Then
            ' Re-verificamos el conteo antes de cada borrado para asegurar que siempre queden 3
            If logFiles.Count > 3 Then
                On Error Resume Next ' Por si el archivo est� abierto o bloqueado
                oFile.Delete True
                If Err.Number = 0 Then
                    ' Actualizar l�gica si es necesario, aunque al ser una colecci�n
                    ' se recorre una foto fija. El "If logFiles.Count > 3"
                    ' aqu� dentro es preventivo.
                End If
                On Error GoTo 0
            End If
        Else
            Exit For
        End If
    Next oFile
    
    ' Limpieza de objetos
    Set oFile = Nothing
    Set oFolder = Nothing
    Set fso = Nothing
    LogInfo MODULE_NAME, "[removeOldLogs] - Finalizado"
End Sub

' ==========================================
' EVENTOS DE APLICACIN - Captura de eventos tcnicos (sin semntica)
' ==========================================

Private Sub mXLApp_WorkbookOpen(ByVal wb As Workbook)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnWorkbookOpened wb
    End If
End Sub

Private Sub mXLApp_WorkbookActivate(ByVal wb As Workbook)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnWorkbookActivated wb
    End If
End Sub

Private Sub mXLApp_WorkbookBeforeClose(ByVal wb As Workbook, ByRef Cancel As Boolean)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnWorkbookBeforeClose wb, Cancel
    End If
End Sub

Private Sub mXLApp_SheetActivate(ByVal sh As Object)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnSheetActivated sh
    End If
End Sub

Private Sub mXLApp_SheetDeactivate(ByVal sh As Object)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnSheetDeactivated sh
    End If
End Sub

Private Sub mXLApp_WindowActivate(ByVal Wn As Window)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnWindowActivated Wn
    End If
End Sub

Private Sub mXLApp_WindowDeactivate(ByVal Wn As Window)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnWindowDeactivated Wn
    End If
End Sub

Private Sub mXLApp_SelectionChange(ByVal Target As Range)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnSelectionChanged Target
    End If
End Sub

Private Sub mXLApp_SheetCalculate(ByVal Sh As Object)
    ' Solo reenviar al gestor de contexto, sin lgica
    If Not mExecutionContextMgr Is Nothing Then
        mExecutionContextMgr.OnSheetCalculated Sh
    End If
End Sub





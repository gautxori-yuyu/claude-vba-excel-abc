VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "1-Aplicacion.3-Ciclo de vida"
Option Explicit

Private Const MODULE_NAME As String = "ThisWorkbook"

Private mApp As clsApplication
Private WithEvents mXLApp As Application
Attribute mXLApp.VB_VarHelpID = -1

Friend Function App() As clsApplication
    If mApp Is Nothing Then
        Set mApp = New clsApplication
    End If
    Set App = mApp
End Function

Public Sub TerminateApp()
    If Not mApp Is Nothing Then
        Set mApp = Nothing
    End If
End Sub

Private Sub Workbook_Open()
    On Error GoTo ErrHandler

    ' 0. Inicializar logger con salida a fichero
    InitLogger LOG_DEBUG, True, ThisWorkbook.path & "\ABC_Log_" & Format(Now, "yyyy-mm-dd_hh-nn") & ".txt"
    removeOldLogs

    LogInfo MODULE_NAME, "[Open] Abriendo Workbook " & ThisWorkbook.Name
    ' Sistema de instalación existente
    AutoInstalador
    AutoRegistrarTodasLasUDFs
    ' Ctrl + Shift + R
    Application.OnKey "^+R", "ToggleRibbonTab"

    ' Capturar eventos de Application (Listener) - CRÍTICO para eventos subsiguientes
    Set mXLApp = Application
    LogDebug MODULE_NAME, "[Open] mXLApp conectado a Application"

    ' Iniciar aplicación (lazy init del singleton)
    App

    ' Tras reset: si la app sobrevivió pero el contexto es inválido, reiniciar
    If Not mApp Is Nothing Then
        If Not App.ExecutionContextMgr Is Nothing Then
            If Not App.ExecutionContextMgr.State.IsContextValid Then
                App.ExecutionContextMgr.ReInitializeContext
            End If
        End If
    End If

    LogInfo MODULE_NAME, "[Open] Inicialización completada correctamente"
    Exit Sub

ErrHandler:
    ' Registrar el error pero NO propagarlo para no interrumpir Excel
    LogError MODULE_NAME, "[Workbook_Open] Error durante inicialización", Err.Number, Err.Description
    ' Intentar mantener la conexión de eventos aunque falle algo
    If mXLApp Is Nothing Then
        On Error Resume Next
        Set mXLApp = Application
        LogWarning MODULE_NAME, "[Workbook_Open] Reconectando mXLApp tras error"
        On Error GoTo 0
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Este evento NO debe desregistrar las UDFs normalmente,
    ' solo si detecta que el complemento fue deshabilitado manualmente

    On Error GoTo ErrHandler

    ' Si fue deshabilitado pero no se disparo AddinUninstall, limpiar
    If Not ComprobarSiInstalado Then
        LogWarning MODULE_NAME, "[Workbook_BeforeClose] Complemento no instalado, realizando limpieza"
    End If

    ' Limpiar atajos de teclado
    Application.OnKey "^+R"

    ' Shutdown del contexto de ejecución
    If Not mApp Is Nothing Then
        App.ExecutionContextMgr.Shutdown
    End If

    ' Terminar aplicacion de forma segura
    TerminateApp

    ' Desregistrar UDFs
    DesregistrarTodasLasUDFs
    LogInfo MODULE_NAME, "[Workbook_BeforeClose] UDFs desregistradas"
    
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
    ' Regenerar listas automáticamente antes de guardar
    'Call GenerarListasUnidades
End Sub

Private Sub Workbook_AddinUninstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario desmarca la casilla en Archivo > Opciones > Complementos
    ' 2. Se ejecuta ai.Installed = False desde el script VBS
    
    On Error Resume Next
    
    ' Mensaje de despedida (opcional)
    '    MsgBox "Complemento desinstalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas han sido eliminadas.", vbInformation, "Desinstalación completada"
    
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario marca la casilla en Archivo > Opciones > Complementos
    ' 2. El script VBS completa la instalación y marca installed=True
    
    On Error Resume Next
    
    ' Mensaje de bienvenida (opcional, puedes comentarlo si es molesto)
    '    MsgBox "Complemento 'ABC Ofertas Máquina Especial' instalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas ya están disponibles.", vbInformation, "Instalación completada"
    
    On Error GoTo 0
End Sub
' Elimina los ficheros de log más antiguos para conservar solo los últimos 3.
Private Sub removeOldLogs()
    Const MAX_LOG_FILES_TO_KEEP As Long = 3
    Const LOG_FILE_PATTERN As String = "ABC_Log_????-??-??_??-??.txt"
    
    Dim fso As Object
    Dim folderPath As String
    Dim logFolder As Object
    Dim logFile As Object
    Dim logFiles() As Object ' Array dinámico de ficheros
    Dim fileCount As Long
    Dim i As Long, j As Long

    On Error GoTo ErrHandler
    
    ' --- 1. Recolectar ficheros de log en un array ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.path
    
    If Not fso.FolderExists(folderPath) Then Exit Sub

    Set logFolder = fso.GetFolder(folderPath)
    fileCount = 0
    ReDim logFiles(1 To logFolder.files.Count) ' Sobredimensionar inicialmente

    For Each logFile In logFolder.files
        If logFile.Name Like LOG_FILE_PATTERN Then
            fileCount = fileCount + 1
            Set logFiles(fileCount) = logFile
        End If
    Next logFile
    
    If fileCount = 0 Then
        Set fso = Nothing
        Set logFolder = Nothing
        Exit Sub
    End If
    ReDim Preserve logFiles(1 To fileCount) ' Ajustar tamaño

    ' --- 2. Salir si no hay suficientes ficheros para limpiar ---
    If fileCount <= MAX_LOG_FILES_TO_KEEP Then
        LogInfo MODULE_NAME, "[removeOldLogs] " & fileCount & " logs encontrados, no se necesita limpieza."
        Exit Sub
    End If

    ' --- 3. Ordenar el array por fecha de modificación (del más antiguo al más nuevo) ---
    Dim temp As Object
    For i = 1 To fileCount - 1
        For j = i + 1 To fileCount
            If logFiles(i).DateLastModified > logFiles(j).DateLastModified Then
                Set temp = logFiles(j)
                Set logFiles(j) = logFiles(i)
                Set logFiles(i) = temp
            End If
        Next j
    Next i

    ' --- 4. Borrar los ficheros más antiguos ---
    Dim filesToDeleteCount As Long
    filesToDeleteCount = fileCount - MAX_LOG_FILES_TO_KEEP
    
    LogInfo MODULE_NAME, "[removeOldLogs] " & fileCount & " logs encontrados. Borrando " & filesToDeleteCount & " más antiguos."

    On Error Resume Next ' Si un fichero está bloqueado, continuar con el siguiente
    For i = 1 To filesToDeleteCount
        LogDebug MODULE_NAME, "[removeOldLogs] Borrando: " & logFiles(i).Name
        logFiles(i).Delete True
    Next i
    On Error GoTo 0
    
    LogInfo MODULE_NAME, "[removeOldLogs] Finalizado."
    
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[removeOldLogs]"
    ' No es necesario un 'Resume', la limpieza se gestiona al salir.
End Sub

' ==========================================
' LISTENER - Reenvio de eventos de Application
' ==========================================

Private Sub mXLApp_WorkbookOpen(ByVal Wb As Workbook)
    If mApp Is Nothing Then Exit Sub
    If mApp.ExecutionContextMgr Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    App.ExecutionContextMgr.OnWorkbookOpened Wb
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mXLApp_WorkbookOpen]"
    Err.Raise Err.Number, MODULE_NAME & "[mXLApp_WorkbookOpen]", _
              Err.Description
End Sub

Private Sub mXLApp_WorkbookActivate(ByVal Wb As Workbook)
    If mApp Is Nothing Then Exit Sub
    If mApp.ExecutionContextMgr Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    App.ExecutionContextMgr.OnWorkbookActivated Wb
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mXLApp_WorkbookActivate]"
    Err.Raise Err.Number, MODULE_NAME & "[mXLApp_WorkbookActivate]", _
              Err.Description
End Sub

Private Sub mXLApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If mApp Is Nothing Then Exit Sub
    If mApp.ExecutionContextMgr Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    App.ExecutionContextMgr.OnWorkbookBeforeClose Wb, Cancel
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mXLApp_WorkbookBeforeClose]"
    Err.Raise Err.Number, MODULE_NAME & "[mXLApp_WorkbookBeforeClose]", _
              Err.Description
End Sub

Private Sub mXLApp_SheetActivate(ByVal Sh As Object)
    If mApp Is Nothing Then Exit Sub
    If mApp.ExecutionContextMgr Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    App.ExecutionContextMgr.OnSheetActivated Sh
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mXLApp_SheetActivate]"
    Err.Raise Err.Number, MODULE_NAME & "[mXLApp_SheetActivate]", _
              Err.Description
End Sub

Private Sub mXLApp_SheetDeactivate(ByVal Sh As Object)
    If mApp Is Nothing Then Exit Sub
    If mApp.ExecutionContextMgr Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    App.ExecutionContextMgr.OnSheetDeactivated Sh
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mXLApp_SheetDeactivate]"
    Err.Raise Err.Number, MODULE_NAME & "[mXLApp_SheetDeactivate]", _
              Err.Description
End Sub




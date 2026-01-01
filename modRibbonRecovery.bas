Attribute VB_Name = "modRibbonRecovery"
'@Folder "2-Control de estado"
' ==========================================
' MODULO DE RECUPERACION DEL RIBBON
' ==========================================
' Proporciona mecanismos para detectar y recuperar el Ribbon
' cuando la referencia IRibbonUI se pierde.
'
' Causas comunes de perdida del Ribbon:
' - Uso de Stop en el codigo
' - Error no manejado que interrumpe la ejecucion
' - Deshabilitacion/habilitacion del complemento
' - Cierre inesperado de Excel
'
' Estrategia de recuperacion:
' 1. Detectar si el Ribbon esta perdido
' 2. Intentar metodos no invasivos primero
' 3. Toggle del add-in como ultimo recurso
' ==========================================

Option Explicit

Private Const MODULE_NAME As String = "modRibbonRecovery"

' Estado de recuperacion
Private mRecoveryAttempts As Long
Private mLastRecoveryTime As Date
Private Const MAX_RECOVERY_ATTEMPTS As Long = 3
Private Const RECOVERY_COOLDOWN_SECONDS As Long = 10

' ==========================================
' FUNCIONES DE DIAGNOSTICO
' ==========================================

'@Description: Verifica si el Ribbon esta disponible y funcional
'@Returns: Boolean | True si el Ribbon esta operativo
Public Function IsRibbonAvailable() As Boolean
    On Error Resume Next

    ' Verificar que App existe
    If App Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: App Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que Ribbon existe
    If App.Ribbon Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: App.Ribbon Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que ribbonUI existe
    If App.Ribbon.ribbonUI Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: ribbonUI Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Intentar una operacion simple para verificar que funciona
    Dim testResult As Boolean
    testResult = Not (TypeName(App.Ribbon.ribbonUI) = "Nothing")

    If Err.Number <> 0 Then
        LogWarning MODULE_NAME, "IsRibbonAvailable: Error al verificar - " & Err.Description
        IsRibbonAvailable = False
        Err.Clear
    Else
        IsRibbonAvailable = testResult
    End If

    On Error GoTo 0
End Function

'@Description: Obtiene informacion de diagnostico del estado del Ribbon
'@Returns: String | Descripcion del estado actual
Public Function GetRibbonDiagnostics() As String
    Dim info As String

    info = "=== DIAGNOSTICO DEL RIBBON ===" & vbCrLf
    info = info & "Fecha/Hora: " & Now & vbCrLf
    info = info & "Log Path: " & GetLogFilePath() & vbCrLf
    info = info & vbCrLf

    ' Estado de App
    If App Is Nothing Then
        info = info & "[X] App: Nothing (CRITICO)" & vbCrLf
        GetRibbonDiagnostics = info
        Exit Function
    Else
        info = info & "[OK] App: Disponible" & vbCrLf
    End If

    ' Estado de Ribbon (clsRibbonEvents)
    If App.Ribbon Is Nothing Then
        info = info & "[X] App.Ribbon: Nothing (ERROR)" & vbCrLf
    Else
        info = info & "[OK] App.Ribbon: Disponible" & vbCrLf

        ' Diagnostico detallado
        info = info & "    -> " & App.Ribbon.GetQuickDiagnostics() & vbCrLf

        ' Estado de ribbonUI (IRibbonUI)
        On Error Resume Next
        If App.Ribbon.ribbonUI Is Nothing Then
            info = info & "[X] ribbonUI: Nothing (PERDIDO)" & vbCrLf
            info = info & "    -> El Ribbon necesita recuperacion" & vbCrLf
        Else
            info = info & "[OK] ribbonUI: Conectado" & vbCrLf
            info = info & "    -> Tipo: " & TypeName(App.Ribbon.ribbonUI) & vbCrLf
        End If
        On Error GoTo 0
    End If

    ' Estado de RibbonState
    If App.RibbonState Is Nothing Then
        info = info & "[X] RibbonState: Nothing" & vbCrLf
    Else
        info = info & "[OK] RibbonState: " & App.RibbonState.RibbonStateDescription & vbCrLf
    End If

    ' Intentos de recuperacion
    info = info & vbCrLf
    info = info & "Intentos de recuperacion: " & mRecoveryAttempts & "/" & MAX_RECOVERY_ATTEMPTS & vbCrLf
    If mLastRecoveryTime > 0 Then
        info = info & "Ultima recuperacion: " & mLastRecoveryTime & vbCrLf
    End If

    GetRibbonDiagnostics = info
End Function

' ==========================================
' FUNCIONES DE RECUPERACION
' ==========================================

'@Description: Intenta recuperar el Ribbon automaticamente
'@Returns: Boolean | True si la recuperacion fue exitosa
Public Function TryRecoverRibbon() As Boolean
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "TryRecoverRibbon - Iniciando recuperacion..."

    ' Verificar cooldown
    If DateDiff("s", mLastRecoveryTime, Now) < RECOVERY_COOLDOWN_SECONDS And mLastRecoveryTime > 0 Then
        LogDebug MODULE_NAME, "TryRecoverRibbon - Cooldown activo (" & RECOVERY_COOLDOWN_SECONDS & "s)"
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Verificar intentos maximos
    If mRecoveryAttempts >= MAX_RECOVERY_ATTEMPTS Then
        LogWarning MODULE_NAME, "TryRecoverRibbon - Maximo de intentos alcanzado (" & MAX_RECOVERY_ATTEMPTS & ")"
        MsgBox "Se han agotado los intentos automaticos de recuperacion del Ribbon." & vbCrLf & vbCrLf & _
               "Opciones:" & vbCrLf & _
               "1. Ejecute 'RecuperarRibbonManual' (Alt+F8)" & vbCrLf & _
               "2. Cierre y reabra Excel" & vbCrLf & vbCrLf & _
               "Consulte el log: " & GetLogFilePath(), _
               vbExclamation, "Ribbon - Recuperacion Fallida"
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Incrementar contador
    mRecoveryAttempts = mRecoveryAttempts + 1
    mLastRecoveryTime = Now

    LogInfo MODULE_NAME, "TryRecoverRibbon - Intento " & mRecoveryAttempts & " de " & MAX_RECOVERY_ATTEMPTS

    ' METODO 1: Soft refresh (no invasivo)
    If RecoverBySoftRefresh() Then
        LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via SoftRefresh"
        ResetRecoveryCounters
        TryRecoverRibbon = True
        Exit Function
    End If

    ' METODO 2: UI Refresh
    If RecoverByUIRefresh() Then
        LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via UIRefresh"
        ResetRecoveryCounters
        TryRecoverRibbon = True
        Exit Function
    End If

    ' METODO 3: Toggle del add-in (ultimo recurso, solo en intento 2+)
    If mRecoveryAttempts >= 2 Then
        LogWarning MODULE_NAME, "TryRecoverRibbon - Intentando toggle del add-in (ultimo recurso)"
        If RecoverByAddinToggle() Then
            LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via AddinToggle"
            ResetRecoveryCounters
            TryRecoverRibbon = True
            Exit Function
        End If
    End If

    LogError MODULE_NAME, "TryRecoverRibbon - Recuperacion fallida en intento " & mRecoveryAttempts
    TryRecoverRibbon = False
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "TryRecoverRibbon - Error", Err.Number, Err.Description
    TryRecoverRibbon = False
End Function

'@Description: Intenta recuperar sin ningun reinicio
Private Function RecoverBySoftRefresh() As Boolean
    On Error Resume Next

    LogDebug MODULE_NAME, "RecoverBySoftRefresh - Verificando puntero"

    DoEvents

    Dim dummy As Object
    Set dummy = Nothing

    DoEvents

    RecoverBySoftRefresh = IsRibbonAvailable()

    If RecoverBySoftRefresh Then
        LogInfo MODULE_NAME, "RecoverBySoftRefresh - Puntero recuperado sin reinicio"
    End If

    On Error GoTo 0
End Function

'@Description: Intenta recuperar forzando redibujado de la UI
Private Function RecoverByUIRefresh() As Boolean
    On Error Resume Next

    LogDebug MODULE_NAME, "RecoverByUIRefresh - Forzando redibujado de UI"

    Application.ScreenUpdating = False
    DoEvents
    Application.ScreenUpdating = True
    DoEvents

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.Visible = True
    End If
    DoEvents

    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    RecoverByUIRefresh = IsRibbonAvailable()

    If RecoverByUIRefresh Then
        LogInfo MODULE_NAME, "RecoverByUIRefresh - Ribbon recuperado via UI refresh"
    End If

    On Error GoTo 0
End Function

'@Description: Recupera el Ribbon toggleando el estado del add-in
Private Function RecoverByAddinToggle() As Boolean
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "RecoverByAddinToggle - Iniciando toggle del add-in"

    Dim ai As AddIn
    Dim targetAddin As AddIn

    ' Buscar nuestro add-in
    For Each ai In Application.AddIns
        If ai.Name = APP_NAME & ".xlam" Then
            Set targetAddin = ai
            Exit For
        End If
    Next ai

    If targetAddin Is Nothing Then
        LogError MODULE_NAME, "RecoverByAddinToggle - Add-in no encontrado: " & APP_NAME & ".xlam"
        RecoverByAddinToggle = False
        Exit Function
    End If

    ' Solo proceder si esta instalado
    If Not targetAddin.Installed Then
        LogError MODULE_NAME, "RecoverByAddinToggle - Add-in no esta instalado"
        RecoverByAddinToggle = False
        Exit Function
    End If

    ' Toggle: desactivar y reactivar
    LogDebug MODULE_NAME, "RecoverByAddinToggle - Desactivando add-in..."
    targetAddin.Installed = False

    ' Pequeña pausa
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    LogDebug MODULE_NAME, "RecoverByAddinToggle - Reactivando add-in..."
    targetAddin.Installed = True

    ' Pausa para que se recargue
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 2)
    DoEvents

    ' Verificar si se recupero
    RecoverByAddinToggle = IsRibbonAvailable()

    If RecoverByAddinToggle Then
        LogInfo MODULE_NAME, "RecoverByAddinToggle - Ribbon recuperado via toggle"
    Else
        LogWarning MODULE_NAME, "RecoverByAddinToggle - Toggle completado pero Ribbon no disponible"
    End If

    Exit Function

ErrHandler:
    LogError MODULE_NAME, "RecoverByAddinToggle - Error", Err.Number, Err.Description
    RecoverByAddinToggle = False
End Function

'@Description: Resetea los contadores de recuperacion
Public Sub ResetRecoveryCounters()
    mRecoveryAttempts = 0
    mLastRecoveryTime = 0
    LogDebug MODULE_NAME, "ResetRecoveryCounters - Contadores reseteados"
End Sub

' ==========================================
' MACROS PUBLICAS (Accesibles por el usuario)
' ==========================================

'@Description: Macro publica para recuperar el Ribbon manualmente
'@Category: Ribbon / Recuperacion
Public Sub RecuperarRibbonManual()
    Dim result As VbMsgBoxResult

    LogInfo MODULE_NAME, "RecuperarRibbonManual - Solicitado por usuario"
    Debug.Print GetRibbonDiagnostics()

    ' Si ya esta disponible, no hacer nada
    If IsRibbonAvailable() Then
        MsgBox "El Ribbon ya esta funcionando correctamente." & vbCrLf & vbCrLf & _
               App.Ribbon.GetQuickDiagnostics(), vbInformation, "Ribbon OK"
        Exit Sub
    End If

    ' Confirmar con el usuario
    result = MsgBox("El Ribbon no esta disponible." & vbCrLf & vbCrLf & _
                    "Se intentara recuperar. Esto puede requerir" & vbCrLf & _
                    "recargar el complemento temporalmente." & vbCrLf & vbCrLf & _
                    "Consulte el log: " & GetLogFilePath() & vbCrLf & vbCrLf & _
                    "Desea continuar?", _
                    vbQuestion + vbYesNo, "Recuperar Ribbon")

    If result <> vbYes Then Exit Sub

    ' Resetear contadores para permitir recuperacion
    ResetRecoveryCounters

    ' Intentar recuperacion
    If TryRecoverRibbon() Then
        MsgBox "Ribbon recuperado exitosamente." & vbCrLf & vbCrLf & _
               App.Ribbon.GetQuickDiagnostics(), vbInformation, "Recuperacion Exitosa"
    Else
        MsgBox "No se pudo recuperar el Ribbon automaticamente." & vbCrLf & vbCrLf & _
               "Recomendaciones:" & vbCrLf & _
               "1. Cierre Excel completamente" & vbCrLf & _
               "2. Vuelva a abrir Excel" & vbCrLf & vbCrLf & _
               "Consulte el log: " & GetLogFilePath(), _
               vbExclamation, "Recuperacion Fallida"
    End If
End Sub

'@Description: Muestra el diagnostico del Ribbon en un cuadro de dialogo
Public Sub MostrarDiagnosticoRibbon()
    LogInfo MODULE_NAME, "MostrarDiagnosticoRibbon - Solicitado"
    MsgBox GetRibbonDiagnostics(), vbInformation, "Diagnostico del Ribbon"
End Sub

'@Description: Abre el fichero de log
Public Sub AbrirLogRibbon()
    On Error Resume Next
    Dim logPath As String
    logPath = GetLogFilePath()

    If Len(Dir(logPath)) > 0 Then
        Shell "notepad.exe """ & logPath & """", vbNormalFocus
    Else
        MsgBox "El fichero de log no existe aun: " & logPath, vbInformation
    End If
    On Error GoTo 0
End Sub

'@Description: Fuerza el reinicio completo de la aplicacion
Public Sub ReiniciarAplicacion()
    Dim result As VbMsgBoxResult

    result = MsgBox("Esto reiniciara completamente el complemento ABC." & vbCrLf & vbCrLf & _
                    "Se cerrara y volvera a inicializar la aplicacion." & vbCrLf & _
                    "¿Desea continuar?", _
                    vbQuestion + vbYesNo, "Reiniciar Aplicacion")

    If result <> vbYes Then Exit Sub

    LogInfo MODULE_NAME, "ReiniciarAplicacion - Reinicio solicitado por usuario"

    On Error Resume Next

    ' Terminar aplicacion actual
    ThisWorkbook.TerminateApp

    ' Reinicializar
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    ' Forzar reinicio llamando a App()
    Dim dummy As clsAplicacion
    Set dummy = App()

    On Error GoTo 0

    ' Verificar estado
    If IsRibbonAvailable() Then
        MsgBox "Aplicación reiniciada correctamente." & vbCrLf & vbCrLf & _
               App.Ribbon.GetQuickDiagnostics(), vbInformation, "Reinicio Exitoso"
    Else
        MsgBox "Aplicación reiniciada, pero el Ribbon puede requerir atención adicional." & vbCrLf & _
               "Ejecute 'RecuperarRibbonManual' si es necesario.", _
               vbExclamation, "Reinicio Parcial"
    End If
End Sub


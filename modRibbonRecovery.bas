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
' 2. Intentar recuperacion automatica (toggle del add-in)
' 3. Proveer macro manual como fallback
' ==========================================

Option Explicit

' Estado de recuperacion
Private mRecoveryAttempts As Long
Private mLastRecoveryTime As Date
Private Const MAX_RECOVERY_ATTEMPTS As Long = 3
Private Const RECOVERY_COOLDOWN_SECONDS As Long = 5

' ==========================================
' FUNCIONES DE DIAGNOSTICO
' ==========================================

'@Description: Verifica si el Ribbon esta disponible y funcional
'@Returns: Boolean | True si el Ribbon esta operativo
Public Function IsRibbonAvailable() As Boolean
    On Error Resume Next

    ' Verificar que App existe
    If App Is Nothing Then
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que Ribbon existe
    If App.Ribbon Is Nothing Then
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que ribbonUI existe
    If App.Ribbon.ribbonUI Is Nothing Then
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Intentar una operacion simple para verificar que funciona
    Dim testResult As Boolean
    testResult = Not (TypeName(App.Ribbon.ribbonUI) = "Nothing")

    If Err.Number <> 0 Then
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

    Debug.Print "[modRibbonRecovery.TryRecoverRibbon] Iniciando recuperacion..."

    ' Verificar cooldown
    If DateDiff("s", mLastRecoveryTime, Now) < RECOVERY_COOLDOWN_SECONDS And mLastRecoveryTime > 0 Then
        Debug.Print "[modRibbonRecovery] Cooldown activo, esperando..."
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Verificar intentos maximos
    If mRecoveryAttempts >= MAX_RECOVERY_ATTEMPTS Then
        Debug.Print "[modRibbonRecovery] Maximo de intentos alcanzado. Use RecuperarRibbonManual."
        MsgBox "Se han agotado los intentos automaticos de recuperacion del Ribbon." & vbCrLf & vbCrLf & _
               "Por favor, ejecute la macro 'RecuperarRibbonManual' o reinicie Excel.", _
               vbExclamation, "Ribbon - Recuperacion Fallida"
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Incrementar contador
    mRecoveryAttempts = mRecoveryAttempts + 1
    mLastRecoveryTime = Now

    Debug.Print "[modRibbonRecovery] Intento " & mRecoveryAttempts & " de " & MAX_RECOVERY_ATTEMPTS

    ' METODO 1: Toggle del complemento (mas efectivo)
    If RecoverByAddinToggle() Then
        Debug.Print "[modRibbonRecovery] Recuperacion exitosa via toggle"
        ResetRecoveryCounters
        TryRecoverRibbon = True
        Exit Function
    End If

    ' METODO 2: Forzar invalidacion del CustomUI
    If RecoverByCustomUIInvalidate() Then
        Debug.Print "[modRibbonRecovery] Recuperacion exitosa via CustomUI"
        ResetRecoveryCounters
        TryRecoverRibbon = True
        Exit Function
    End If

    Debug.Print "[modRibbonRecovery] Recuperacion fallida en intento " & mRecoveryAttempts
    TryRecoverRibbon = False
    Exit Function

ErrHandler:
    Debug.Print "[modRibbonRecovery.TryRecoverRibbon] Error: " & Err.Description
    TryRecoverRibbon = False
End Function

'@Description: Recupera el Ribbon toggleando el estado del add-in
Private Function RecoverByAddinToggle() As Boolean
    On Error GoTo ErrHandler

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
        Debug.Print "[RecoverByAddinToggle] Add-in no encontrado"
        RecoverByAddinToggle = False
        Exit Function
    End If

    ' Solo proceder si esta instalado
    If Not targetAddin.Installed Then
        Debug.Print "[RecoverByAddinToggle] Add-in no esta instalado"
        RecoverByAddinToggle = False
        Exit Function
    End If

    ' Toggle: desactivar y reactivar
    Debug.Print "[RecoverByAddinToggle] Desactivando add-in..."
    targetAddin.Installed = False

    ' Pequena pausa
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    Debug.Print "[RecoverByAddinToggle] Reactivando add-in..."
    targetAddin.Installed = True

    ' Pausa para que se recargue
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 2)
    DoEvents

    ' Verificar si se recupero
    RecoverByAddinToggle = IsRibbonAvailable()

    Exit Function

ErrHandler:
    Debug.Print "[RecoverByAddinToggle] Error: " & Err.Description
    RecoverByAddinToggle = False
End Function

'@Description: Intenta recuperar forzando invalidacion del CustomUI
Private Function RecoverByCustomUIInvalidate() As Boolean
    On Error GoTo ErrHandler

    ' Esta tecnica funciona si el objeto ribbonUI existe pero esta en estado inconsistente
    ' Forzamos a Excel a recargar la UI

    ' Forzar recalculo de la ventana
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

    DoEvents

    ' Verificar resultado
    RecoverByCustomUIInvalidate = IsRibbonAvailable()

    Exit Function

ErrHandler:
    Debug.Print "[RecoverByCustomUIInvalidate] Error: " & Err.Description
    RecoverByCustomUIInvalidate = False
End Function

'@Description: Resetea los contadores de recuperacion
Private Sub ResetRecoveryCounters()
    mRecoveryAttempts = 0
    mLastRecoveryTime = 0
End Sub

' ==========================================
' MACROS PUBLICAS (Accesibles por el usuario)
' ==========================================

'@Description: Macro publica para recuperar el Ribbon manualmente
'@Category: Ribbon / Recuperacion
Public Sub RecuperarRibbonManual()
Attribute RecuperarRibbonManual.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim result As VbMsgBoxResult

    ' Mostrar diagnostico primero
    Debug.Print GetRibbonDiagnostics()

    ' Si ya esta disponible, no hacer nada
    If IsRibbonAvailable() Then
        MsgBox "El Ribbon ya esta funcionando correctamente.", vbInformation, "Ribbon OK"
        Exit Sub
    End If

    ' Confirmar con el usuario
    result = MsgBox("El Ribbon no esta disponible." & vbCrLf & vbCrLf & _
                    "Esto cerrara y reabrira el complemento para restaurar el Ribbon." & vbCrLf & _
                    "Los datos no guardados en otros libros no se veran afectados." & vbCrLf & vbCrLf & _
                    "¿Desea continuar?", _
                    vbQuestion + vbYesNo, "Recuperar Ribbon")

    If result <> vbYes Then Exit Sub

    ' Resetear contadores para permitir recuperacion
    ResetRecoveryCounters

    ' Intentar recuperacion
    If TryRecoverRibbon() Then
        MsgBox "Ribbon recuperado exitosamente.", vbInformation, "Recuperacion Exitosa"
    Else
        MsgBox "No se pudo recuperar el Ribbon automaticamente." & vbCrLf & vbCrLf & _
               "Recomendaciones:" & vbCrLf & _
               "1. Cierre Excel completamente" & vbCrLf & _
               "2. Vuelva a abrir Excel" & vbCrLf & _
               "3. El Ribbon deberia cargarse automaticamente", _
               vbExclamation, "Recuperacion Fallida"
    End If
End Sub

'@Description: Muestra el diagnostico del Ribbon en un cuadro de dialogo
Public Sub MostrarDiagnosticoRibbon()
Attribute MostrarDiagnosticoRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    MsgBox GetRibbonDiagnostics(), vbInformation, "Diagnostico del Ribbon"
End Sub

'@Description: Fuerza el reinicio completo de la aplicacion
Public Sub ReiniciarAplicacion()
Attribute ReiniciarAplicacion.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim result As VbMsgBoxResult

    result = MsgBox("Esto reiniciara completamente el complemento ABC." & vbCrLf & vbCrLf & _
                    "Se cerrara y volvera a inicializar la aplicacion." & vbCrLf & _
                    "¿Desea continuar?", _
                    vbQuestion + vbYesNo, "Reiniciar Aplicacion")

    If result <> vbYes Then Exit Sub

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
        MsgBox "Aplicacion reiniciada correctamente.", vbInformation, "Reinicio Exitoso"
    Else
        MsgBox "Aplicacion reiniciada, pero el Ribbon puede requerir atencion adicional." & vbCrLf & _
               "Ejecute 'RecuperarRibbonManual' si es necesario.", _
               vbExclamation, "Reinicio Parcial"
    End If
End Sub

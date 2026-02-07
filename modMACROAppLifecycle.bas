Attribute VB_Name = "modMACROAppLifecycle"
Option Private Module
' ==========================================
' CICLO DE VIDA DE LA APLICACION
' ==========================================
' Funciones publicas para gestion de la aplicacion y Ribbon
' ==========================================

'@Folder "1-Inicio e Instalacion.Ciclo de vida"
Option Explicit

Private Const MODULE_NAME As String = "modMACROAppLifecycle"

' ==========================================
' DETECCION DE RESET VBA
' ==========================================
' Cuando VBA se resetea (error fatal, End en depuracion, etc.),
' todas las variables de modulo se reinicializan a 0/Nothing.
' Usamos una variable Static dentro de una funcion para detectar esto.
' ==========================================

Private mInitCounter As Long       ' Contador de inicializaciones (persistente en sesion)
Private mLastInitTime As Double    ' Timestamp de ultima inicializacion

'@Description: Detecta si ocurrio un reset de VBA desde la ultima llamada
'@Returns: True si es la primera llamada tras un reset (contador > 1)
'@Note: Esta funcion usa una variable Static que sobrevive entre llamadas
'       pero se reinicia si VBA hace reset. El patron detecta ese reset.
Public Function DetectVBAResetOccurred() As Boolean
Attribute DetectVBAResetOccurred.VB_ProcData.VB_Invoke_Func = " \n0"
    Static sInitFlag As Boolean

    If Not sInitFlag Then
        ' Primera vez que se ejecuta desde reset
        sInitFlag = True
        mInitCounter = mInitCounter + 1
        mLastInitTime = Timer

        ' Si contador > 1, hubo reset previo
        DetectVBAResetOccurred = (mInitCounter > 1)

        If DetectVBAResetOccurred Then
            LogWarning MODULE_NAME, "[DetectVBAResetOccurred] Reset detectado! Inicializacion #" & mInitCounter
        Else
            LogInfo MODULE_NAME, "[DetectVBAResetOccurred] Primera inicializacion de la sesion"
        End If
    Else
        ' Llamadas subsiguientes en la misma sesion - no hay reset
        DetectVBAResetOccurred = False
    End If
End Function

'@Description: Obtiene el numero de veces que se ha inicializado la aplicacion
'@Returns: Long - Contador de inicializaciones (1 = primera vez, >1 = hubo resets)
Public Property Get InitializationCount() As Long
    InitializationCount = mInitCounter
End Property

'@Description: Obtiene el timestamp de la ultima inicializacion
'@Returns: Double - Valor de Timer en la ultima inicializacion
Public Property Get LastInitializationTime() As Double
    LastInitializationTime = mLastInitTime
End Property

' ==========================================
' ACCESO A LA APLICACION
' ==========================================

Public Function App() As clsApplication
Attribute App.VB_ProcData.VB_Invoke_Func = " \n0"
    Set App = ThisWorkbook.App
End Function

'@Description: Fuerza el reinicio completo de la aplicacion
Public Sub ReiniciarAplicacion()
Attribute ReiniciarAplicacion.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim result As VbMsgBoxResult

    result = MsgBox("Esto reiniciara completamente el complemento ABC." & vbCrLf & vbCrLf & _
                    "Se cerrara y volvera a inicializar la aplicacion." & vbCrLf & _
                    "¿Desea continuar?", _
                    vbQuestion + vbYesNo, "Reiniciar Aplicacion")

    If result <> vbYes Then Exit Sub

    LogInfo MODULE_NAME, "[ReiniciarAplicacion] Reinicio solicitado por usuario"

    On Error Resume Next

    ' Terminar aplicacion actual
    ThisWorkbook.TerminateApp

    ' Reinicializar
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    ' Forzar reinicio llamando a App()
    Dim dummy As clsApplication
    Set dummy = App()

    On Error GoTo 0

    ' Verificar estado
    If IsRibbonAvailable() Then
        MsgBox "Aplicación reiniciada correctamente." & vbCrLf & vbCrLf & _
               App.ribbon.GetQuickDiagnostics(), vbInformation, "Reinicio Exitoso"
    Else
        MsgBox "Aplicación reiniciada, pero el Ribbon puede requerir atención adicional." & vbCrLf & _
               "Ejecute 'RecuperarRibbon' si es necesario.", _
               vbExclamation, "Reinicio Parcial"
    End If
End Sub

' ==========================================
' GESTION DEL COMPLEMENTO XLAM
' MACROS PUBLICAS (Accesibles por el usuario)
' ==========================================

'@Description: Activa temporalmente la visibilidad del XLAM para operaciones de copia
'              Muestra el libro que contiene este XLAM, haciéndolo visible en la interfaz de Excel.
'@Scope: Manipula el libro host del complemento XLAM cargado.
'@ArgumentDescriptions: (no tiene argumentos)
'@Returns: (ninguno)
'@Category: ComplementosExcel
Sub DesactivarModoAddin()
Attribute DesactivarModoAddin.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False     ' Hace que el libro se muestre
        LogInfo MODULE_NAME, "[DesactivarModoAddin] XLAM visible temporalmente"
    End If
ErrHandler:
    Err.Raise Err.Number, "modMACROBackups.DesactivarModoAddin", _
              "Error desactivando el modo de AddIn: " & Err.Description
End Sub

'@Description: Restaura el estado de IsAddin del XLAM
'              Oculta el libro que contiene este XLAM, dejando el complemento operativo pero sin mostrar su ventana.
'@Scope: Manipula el libro host del complemento XLAM cargado.
'@ArgumentDescriptions: (no tiene argumentos)
'@Returns: (ninguno)
'@Category: ComplementosExcel
Sub RestaurarModoAddin()
Attribute RestaurarModoAddin.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    ThisWorkbook.IsAddin = True
    LogInfo MODULE_NAME, "[RestaurarModoAddin] XLAM restaurado como Add-in"
ErrHandler:
    Err.Raise Err.Number, "modMACROBackups.DesactivarModoAddin", _
              "Error desactivando el modo de AddIn: " & Err.Description
End Sub

' ==========================================
' GESTION DEL RIBBON
' MACROS PUBLICAS (Accesibles por el usuario)
' ==========================================

'@Description: Procedimiento puente para el atajo de teclado
Public Sub ToggleRibbonTab()
Attribute ToggleRibbonTab.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler

    App.ribbon.ToggleModo

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ToggleRibbonTab] Error", Err.Number, Err.Description
    MsgBox "Error al cambiar modo del Ribbon: " & Err.Description, vbExclamation
End Sub

'@Description: Macro publica para recuperar el Ribbon manualmente
'@Note: Ejecutar esta macro si el Ribbon desaparece o no responde
'@Category: Ribbon / Recuperacion
Public Sub RecuperarRibbon()
Attribute RecuperarRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim result As VbMsgBoxResult

    LogInfo MODULE_NAME, "[RecuperarRibbon] Solicitado por usuario"
    LogDebug MODULE_NAME, "[RecuperarRibbon] Diagnosticos: " & GetRibbonDiagnostics()

    ' Si ya esta disponible, no hacer nada
    If IsRibbonAvailable() Then
        MsgBox "El Ribbon ya esta funcionando correctamente.", vbInformation, "Ribbon OK"
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

    ' Intentar recuperacion
    If TryRecoverRibbon() Then
        MsgBox "Ribbon recuperado exitosamente." & vbCrLf & vbCrLf & _
               App.ribbon.GetQuickDiagnostics(), vbInformation, "Recuperacion Exitosa"
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
Attribute MostrarDiagnosticoRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "MostrarDiagnosticoRibbon - Solicitado"
    MsgBox GetRibbonDiagnostics(), vbInformation, "Diagnostico del Ribbon"
End Sub

' ------------------------------------------
' FUNCIONES DE DIAGNOSTICO
' ------------------------------------------

'@Description: Obtiene informacion de diagnostico del estado del Ribbon
'@Returns: String | Descripcion del estado actual
Public Function GetRibbonDiagnostics() As String
Attribute GetRibbonDiagnostics.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim info As String

    info = "=== DIAGNOSTICO DEL RIBBON ===" & vbCrLf
    info = info & "Fecha/Hora: " & Now & vbCrLf
    info = info & "Log Path: " & GetLogFilePath() & vbCrLf
    info = info & vbCrLf

    ' Estado de App
    Dim mApp As clsApplication
    Set mApp = App
    If mApp Is Nothing Then
        info = info & "[X] App: Nothing (CRITICO)" & vbCrLf
        GetRibbonDiagnostics = info
        Exit Function
    Else
        info = info & "[OK] App: Disponible" & vbCrLf
    End If

    ' Estado de Ribbon (clsRibbon)
    If mApp.ribbon Is Nothing Then
        info = info & "[X] App.Ribbon: Nothing (ERROR)" & vbCrLf
    Else
        info = info & "[OK] App.Ribbon: Disponible" & vbCrLf

        ' Diagnostico detallado
        info = info & "    -> " & mApp.ribbon.GetQuickDiagnostics() & vbCrLf

        ' Estado de ribbonUI (IRibbonUI)
        On Error Resume Next
        If mApp.ribbon.ribbonUI Is Nothing Then
            info = info & "[X] ribbonUI: Nothing (PERDIDO)" & vbCrLf
            info = info & "    -> El Ribbon necesita recuperacion" & vbCrLf
        Else
            info = info & "[OK] ribbonUI: Conectado" & vbCrLf
            info = info & "    -> Tipo: " & TypeName(App.ribbon.ribbonUI) & vbCrLf
        End If
        On Error GoTo 0
    End If

    ' Estado de Ribbon.State
    If mApp.ribbon.State Is Nothing Then
        info = info & "[X] Ribbon State: Nothing" & vbCrLf
    Else
        info = info & "[OK] Ribbon State: " & mApp.ribbon.State.Description & vbCrLf
    End If

    GetRibbonDiagnostics = info
End Function

'@Description: Verifica si el Ribbon esta disponible y funcional
'    DESDE EL CONTEXTO GLOBAL
'@Returns: Boolean | True si el Ribbon esta operativo
Public Function IsRibbonAvailable() As Boolean
Attribute IsRibbonAvailable.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next
    Dim mApp As clsApplication
    ' Verificar que App existe
    Set mApp = App
    If mApp Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: App Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que Ribbon existe
    If mApp.ribbon Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: App.Ribbon Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Verificar que ribbonUI existe
    If mApp.ribbon.ribbonUI Is Nothing Then
        LogDebug MODULE_NAME, "IsRibbonAvailable: ribbonUI Is Nothing"
        IsRibbonAvailable = False
        Exit Function
    End If

    ' Intentar una operacion simple para verificar que funciona
    Dim testResult As Boolean
    testResult = Not (TypeName(mApp.ribbon.ribbonUI) = "Nothing")

    If Err.Number <> 0 Then
        LogWarning MODULE_NAME, "IsRibbonAvailable: Error al verificar - " & Err.Description
        IsRibbonAvailable = False
        Err.Clear
    Else
        IsRibbonAvailable = testResult
    End If

    On Error GoTo 0
End Function

' ------------------------------------------
' FUNCIONES DE RECUPERACION
' ------------------------------------------

'@Description: Intenta recuperar el Ribbon automaticamente
'@Returns: Boolean | True si la recuperacion fue exitosa
Public Function TryRecoverRibbon() As Boolean
Attribute TryRecoverRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "TryRecoverRibbon - Iniciando recuperacion..."

    ' METODO 1: Soft refresh (no invasivo)
    If RecoverBySoftRefresh() Then
        LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via SoftRefresh"
        TryRecoverRibbon = True
        Exit Function
    End If

    ' METODO 2: UI Refresh
    If RecoverByUIRefresh() Then
        LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via UIRefresh"
        TryRecoverRibbon = True
        Exit Function
    End If

    ' METODO 3: Toggle del add-in (ultimo recurso, solo en intento 2+)
    If False Then
        LogWarning MODULE_NAME, "TryRecoverRibbon - Intentando toggle del add-in (ultimo recurso)"
        If RecoverByAddinToggle() Then
            LogInfo MODULE_NAME, "TryRecoverRibbon - Exito via AddinToggle"
            TryRecoverRibbon = True
            Exit Function
        End If
    End If

    LogError MODULE_NAME, "TryRecoverRibbon - Recuperacion fallida"
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

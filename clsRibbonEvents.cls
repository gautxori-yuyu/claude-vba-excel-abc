VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================
' CLASE DE EVENTOS DEL RIBBON
' ==========================================
' Esta clase envuelve el objeto IRibbonUI y gestiona
' su ciclo de vida con proteccion y logging.
'
' CAMBIOS IMPORTANTES:
' - ribbonUI ahora es privado con Property Get/Set
' - Cada cambio del puntero se registra en log (modLogger)
' - mWasEverInitialized es informativo, no bloquea
' ==========================================

Private Const MODULE_NAME As String = "clsRibbonEvents"

' ==========================================
' RIBBON UI Y ESTADO (PRIVADO)
' ==========================================
Private mRibbonUI As IRibbonUI              ' Puntero al Ribbon (PROTEGIDO)
Private mIsRecovering As Boolean            ' Flag para evitar recursion durante recuperacion
Private mWasEverInitialized As Boolean      ' Flag informativo
Private mLastKnownState As String           ' Ultimo estado conocido del puntero
Private mInitializationTime As Date         ' Cuando se inicializo por ultima vez
Private mInvalidationCount As Long          ' Contador de invalidaciones exitosas

' ==========================================
' EVENTOS DE APLICACION
' ==========================================
Public Event GenerarGraficosDesdeCurvasRto()
Public Event InvertirEjes()
Public Event FormatearCGASING()
Public Event Configurador()
Public Event NuevaOportunidad()
Public Event ReplaceWithNamesInValidations()

' ==========================================
' PROPERTY GET/SET PARA ribbonUI (PROTEGIDO)
' ==========================================

'@Description: Obtiene el puntero al Ribbon (puede ser Nothing)
Public Property Get ribbonUI() As IRibbonUI
    Set ribbonUI = mRibbonUI
End Property

'@Description: Establece el puntero al Ribbon con logging
Public Property Set ribbonUI(ByVal newRibbon As IRibbonUI)
    Dim oldState As String
    Dim newState As String

    ' Capturar estado anterior
    If mRibbonUI Is Nothing Then
        oldState = "Nothing"
    Else
        On Error Resume Next
        oldState = TypeName(mRibbonUI)
        If Err.Number <> 0 Then
            oldState = "Invalid/Corrupted"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Capturar nuevo estado
    If newRibbon Is Nothing Then
        newState = "Nothing"
    Else
        On Error Resume Next
        newState = TypeName(newRibbon)
        If Err.Number <> 0 Then
            newState = "Unknown"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Log del cambio usando modLogger
    LogWarning MODULE_NAME, "RIBBON POINTER CHANGE: " & oldState & " -> " & newState

    ' Asignar el nuevo valor
    Set mRibbonUI = newRibbon
    mLastKnownState = newState

    ' Actualizar flags
    If Not newRibbon Is Nothing Then
        mWasEverInitialized = True
        mInitializationTime = Now
    End If
End Property

' ==========================================
' INICIALIZACION Y LIMPIEZA
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "Class_Initialize - Nueva instancia creada"
    mWasEverInitialized = False
    mIsRecovering = False
    mInvalidationCount = 0
    mLastKnownState = "NotInitialized"
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "Class_Terminate - Instancia destruida. Estado: " & mLastKnownState
End Sub

'@Description: Inicializa el Ribbon con el puntero proporcionado
Public Sub Init(ByRef ribbonObj As IRibbonUI)
    LogInfo MODULE_NAME, "Init - Recibiendo puntero IRibbonUI"
    Set Me.ribbonUI = ribbonObj

    If ribbonObj Is Nothing Then
        LogWarning MODULE_NAME, "Init - Se paso Nothing como puntero"
    Else
        LogInfo MODULE_NAME, "Init - Puntero establecido correctamente"
    End If
End Sub

'@Description: Libera el puntero al Ribbon
Public Sub StopEvents()
    LogInfo MODULE_NAME, "StopEvents - Liberando puntero"
    Set Me.ribbonUI = Nothing
End Sub

' ==========================================
' EVENTOS DE APLICACION
' ==========================================

Public Sub OnGenerarGraficosDesdeCurvasRto()
    RaiseEvent GenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes()
    RaiseEvent InvertirEjes
End Sub

Public Sub OnFormatearCGASING()
    RaiseEvent FormatearCGASING
End Sub

Public Sub OnConfigurador()
    RaiseEvent Configurador
End Sub

Public Sub OnNuevaOportunidad()
    RaiseEvent NuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations()
    RaiseEvent ReplaceWithNamesInValidations
End Sub

' ==========================================
' ACTIVACION DE TABS
' ==========================================

'@Description: Activa una pestaña del Ribbon por id
'@Scope: Friend (uso exclusivo desde clsAplicacion)
'@ArgumentDescriptions: tabId | id del tab a activar
'@Category: UI / Ribbon
Friend Sub ActivarTab(tabId As String)
    On Error Resume Next
    If IsRibbonUIAvailable() Then
        mRibbonUI.ActivateTab tabId
    End If
End Sub

' ==========================================
' GESTION DE ESTADO DE CONTROLES
' ==========================================

Public Function GetRibbonControlEnabled(control As IRibbonControl) As Boolean
    Dim enabled As Boolean
    Select Case control.id
    Case "btnGenerarGraficos"                    ' GetGraficoEnabled
        enabled = EsFicheroOportunidad() And EsValidoGenerarGrafico()
    Case "btnInvertirSeries"                     ' GetInvertirEjesEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = App.bChartActive ' comprueba que hay un grafico seleccionado
        If enabled Then enabled = EsValidoInvertirEjes()
        ' HAY QUE PASAR A ESTO: el orquestador determina el estado del ribbon:
        ' ... PERO DE MOMENTO, NO ESTA TERMINADO DE IMPLEMENTAR clsExcelFile y clsFileManager
        ' enabled = App.bCanInvertAxes
    Case "btnCGASING"                            ' GetCGASINGEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = IsDefaultCGasIngSheet()
    Case "btnNuevaOp"                            ' GetNuevaOportunidadEnabled
        enabled = True
    End Select
    GetRibbonControlEnabled = enabled
    'InvalidarControl control.id ' esto ESTA DE SOBRA!!
    Debug.Print "[GetRibbonControlEnabled] - ¿" & control.id & "?: " & CBool(enabled)
End Function

' ==========================================
' INVALIDACION DEL RIBBON
' ==========================================

'@Description: Invalida todo el Ribbon, forzando recarga de callbacks
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarRibbon()
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        ' Esto evita que durante Class_Initialize se intente recuperar un ribbon que
        ' aun no ha sido configurado (el callback RibbonOnLoad aun no ha terminado)
        LogCritical MODULE_NAME, "InvalidarRibbon - Puntero no disponible. Estado: " & mLastKnownState

        ' Intentar recuperacion automatica (solo si no estamos ya recuperando)
        If Not mIsRecovering Then
            LogInfo MODULE_NAME, "InvalidarRibbon - Iniciando recuperacion automatica"
            If TryAutoRecover() Then
                LogInfo MODULE_NAME, "InvalidarRibbon - Recuperacion exitosa"
            Else
                LogError MODULE_NAME, "InvalidarRibbon - Recuperacion fallida"
                Exit Sub
            End If
        Else
            LogDebug MODULE_NAME, "InvalidarRibbon - Recuperacion ya en curso, omitiendo"
            Exit Sub
        End If
    End If

    ' Invalidar el Ribbon
    mRibbonUI.Invalidate
    mInvalidationCount = mInvalidationCount + 1
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "InvalidarRibbon - Error", Err.Number, Err.Description
    ' No propagar el error, solo loguear
End Sub

'@Description: Invalida un control especifico del Ribbon
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        LogCritical MODULE_NAME, "InvalidarControl(" & idControl & ") - Puntero no disponible"

        ' Intentar recuperacion automatica
        If Not mIsRecovering Then
            Debug.Print "[clsRibbonEvents.InvalidarControl] Ribbon perdido, intentando recuperacion..."
            If Not TryAutoRecover() Then
                Debug.Print "[clsRibbonEvents.InvalidarControl] Recuperacion fallida para control: " & idControl
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Invalidar el control
    mRibbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "InvalidarControl(" & idControl & ") - Error", Err.Number, Err.Description
End Sub

' ==========================================
' VERIFICACION DEL ESTADO
' ==========================================

'@Description: Verifica si el objeto ribbonUI esta disponible y funcional
Private Function IsRibbonUIAvailable() As Boolean
    On Error Resume Next

    If mRibbonUI Is Nothing Then
        IsRibbonUIAvailable = False
        Exit Function
    End If

    ' Verificacion de tipo
    Dim testType As String
    testType = TypeName(mRibbonUI)

    If Err.Number <> 0 Then
        LogWarning MODULE_NAME, "IsRibbonUIAvailable - Puntero corrupto detectado"
        IsRibbonUIAvailable = False
        Err.Clear
        Exit Function
    End If

    If testType = "Nothing" Or testType = "Empty" Then
        IsRibbonUIAvailable = False
        Exit Function
    End If

    IsRibbonUIAvailable = True
    On Error GoTo 0
End Function

' ==========================================
' RECUPERACION AUTOMATICA
' ==========================================

'@Description: Intenta recuperar el Ribbon automaticamente
Private Function TryAutoRecover() As Boolean
    On Error GoTo ErrHandler

    mIsRecovering = True
    LogInfo MODULE_NAME, "TryAutoRecover - Iniciando"

    ' Usar el modulo de recuperacion
    TryAutoRecover = modRibbonRecovery.TryRecoverRibbon()

    mIsRecovering = False
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "TryAutoRecover - Error", Err.Number, Err.Description
    mIsRecovering = False
    TryAutoRecover = False
End Function

' ==========================================
' PROPIEDADES DE ESTADO
' ==========================================

'@Description: Indica si el Ribbon esta en proceso de recuperacion
Public Property Get IsRecovering() As Boolean
    IsRecovering = mIsRecovering
End Property

'@Description: Indica si el Ribbon fue inicializado alguna vez
Public Property Get WasEverInitialized() As Boolean
    WasEverInitialized = mWasEverInitialized
End Property

'@Description: Obtiene cuando fue la ultima inicializacion
Public Property Get LastInitializationTime() As Date
    LastInitializationTime = mInitializationTime
End Property

'@Description: Obtiene el contador de invalidaciones
Public Property Get InvalidationCount() As Long
    InvalidationCount = mInvalidationCount
End Property

'@Description: Obtiene el ultimo estado conocido del puntero
Public Property Get LastKnownState() As String
    LastKnownState = mLastKnownState
End Property

'@Description: Diagnostico rapido del estado del Ribbon
Public Function GetQuickDiagnostics() As String
    Dim info As String
    info = "RibbonUI: "

    If mRibbonUI Is Nothing Then
        info = info & "Nothing"
    Else
        On Error Resume Next
        info = info & TypeName(mRibbonUI)
        If Err.Number <> 0 Then
            info = info & "Corrupted"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    info = info & " | WasInit: " & mWasEverInitialized
    info = info & " | Recovering: " & mIsRecovering
    info = info & " | Invalidations: " & mInvalidationCount

    If mInitializationTime > 0 Then
        info = info & " | InitTime: " & Format(mInitializationTime, "hh:nn:ss")
    End If

    GetQuickDiagnostics = info
End Function


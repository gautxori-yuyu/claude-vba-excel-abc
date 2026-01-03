VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================
' CLASE DE EVENTOS DEL RIBBON
' ==========================================
' Esta clase envuelve el objeto IRibbonUI y gestiona
' su ciclo de vida con proteccion y logging.
'
' CAMBIOS IMPORTANTES:
' - ribbonUI ahora es privado con Property Get
' - mWasEverInitialized es informativo, no bloquea
' ==========================================

'@Folder "2-Servicios.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbonEvents"

' ==========================================
' RIBBON UI Y ESTADO
' ==========================================
Private mribbonUI As IRibbonUI
Private mIsRecovering As Boolean              ' Flag para evitar recursion durante recuperacion
Private mWasEverInitialized As Boolean        ' Flag para distinguir "nunca inicializado" de "perdido"

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
    Set ribbonUI = mribbonUI
End Property
' ==========================================
' INICIALIZACION Y LIMPIEZA
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    mWasEverInitialized = False
    mIsRecovering = False
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

Public Sub Init(ByRef ribbonObj As IRibbonUI)
    LogDebug MODULE_NAME, "[Init] - Recibiendo puntero IRibbonUI"
    Set mribbonUI = ribbonObj

    If ribbonObj Is Nothing Then
        LogWarning MODULE_NAME, "[Init] - Se paso Nothing como puntero"
    Else
        LogDebug MODULE_NAME, "[Init] - Puntero establecido correctamente"
    End If
    mWasEverInitialized = True
End Sub

Public Sub StopEvents()
    LogDebug MODULE_NAME, "[StopEvents] - Liberando puntero"
    Set mribbonUI = Nothing
End Sub

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

'@Description: Activa una pestaña del Ribbon por id
'@Scope: Friend (uso exclusivo desde clsAplicacion)
'@ArgumentDescriptions: tabId | id del tab a activar
'@Category: UI / Ribbon
Friend Sub ActivarTab(tabId As String)
    On Error Resume Next
    If IsRibbonUIAvailable() Then
        mribbonUI.ActivateTab tabId
    End If
End Sub

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
    LogDebug MODULE_NAME, "[GetRibbonControlEnabled] - ¿" & control.id & "?: " & CBool(enabled)
End Function

'@Description: Invalida todo el Ribbon, forzando recarga de callbacks
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarRibbon()
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        ' Esto evita que durante Class_Initialize se intente recuperar un ribbon que
        ' aun no ha sido configurado (el callback RibbonOnLoad aun no ha terminado)
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarRibbon] Ribbon aun no inicializado, omitiendo"
            Exit Sub
        End If

        ' Intentar recuperacion automatica (solo si no estamos ya recuperando)
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarRibbon] Ribbon perdido, intentando recuperacion..."
            If TryAutoRecover() Then
                LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion exitosa"
            Else
                LogError MODULE_NAME, "[InvalidarRibbon] Recuperacion fallida"
                Exit Sub
            End If
        Else
            LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion en curso, omitiendo"
            Exit Sub
        End If
    End If

    ' Invalidar el Ribbon
    mribbonUI.Invalidate
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarRibbon] Error", Err.Description
    ' No propagar el error, solo loguear
End Sub

'@Description: Invalida un control especifico del Ribbon
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarControl] Ribbon aun no inicializado, omitiendo: " & idControl
            Exit Sub
        End If

        ' Intentar recuperacion automatica
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarControl] Ribbon perdido, intentando recuperacion..."
            If Not TryAutoRecover() Then
                LogError MODULE_NAME, "[InvalidarControl] Recuperacion fallida para control: " & idControl
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Invalidar el control
    mribbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarControl] Error en control " & idControl, , Err.Description
End Sub

'@Description: Verifica si el objeto ribbonUI esta disponible y funcional
Private Function IsRibbonUIAvailable() As Boolean
    If mribbonUI Is Nothing Then IsRibbonUIAvailable = False: Exit Function
    
    On Error Resume Next
    Dim testType As String
    testType = TypeName(mribbonUI)
    If Err.Number <> 0 Then
        IsRibbonUIAvailable = False
        Err.Clear
    ElseIf testType = "Nothing" Or testType = "Empty" Then
        IsRibbonUIAvailable = False
    Else
        IsRibbonUIAvailable = True
    End If
End Function

'@Description: Intenta recuperar el Ribbon automaticamente
Private Function TryAutoRecover() As Boolean
    On Error GoTo ErrHandler

    mIsRecovering = True
    LogInfo MODULE_NAME, "[TryAutoRecover] - Iniciando"

    ' Usar el modulo de recuperacion
    TryAutoRecover = TryRecoverRibbon()

    mIsRecovering = False
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[TryAutoRecover] - Error", Err.Number, Err.Description
    mIsRecovering = False
    TryAutoRecover = False
End Function

'@Description: Indica si el Ribbon esta en proceso de recuperacion
Public Property Get IsRecovering() As Boolean
    IsRecovering = mIsRecovering
End Property


'@Description: Diagnostico rapido del estado del Ribbon
Public Function GetQuickDiagnostics() As String
    Dim info As String
    info = "RibbonUI: "

    If Not IsRibbonUIAvailable() Then
        info = info & "Nothing"
    Else
        On Error Resume Next
        info = info & TypeName(mribbonUI)
        If Err.Number <> 0 Then
            info = info & "Corrupted"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    info = info & " | WasInit: " & mWasEverInitialized
    info = info & " | Recovering: " & mIsRecovering

    GetQuickDiagnostics = info
End Function



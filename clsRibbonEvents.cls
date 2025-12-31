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
' RIBBON UI Y ESTADO
' ==========================================
Public ribbonUI As IRibbonUI
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

Public Sub Init(ByRef ribbonObj As IRibbonUI)
    Set ribbonUI = ribbonObj
    mWasEverInitialized = True
End Sub

Public Sub StopEvents()
    Set ribbonUI = Nothing
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
    If Not ribbonUI Is Nothing Then
        ribbonUI.ActivateTab tabId
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
    Debug.Print "[GetRibbonControlEnabled] - ¿" & control.id & "?: " & CBool(enabled)
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
            Debug.Print "[clsRibbonEvents.InvalidarRibbon] Ribbon aun no inicializado, omitiendo"
            Exit Sub
        End If

        ' Intentar recuperacion automatica (solo si no estamos ya recuperando)
        If Not mIsRecovering Then
            Debug.Print "[clsRibbonEvents.InvalidarRibbon] Ribbon perdido, intentando recuperacion..."
            If TryAutoRecover() Then
                Debug.Print "[clsRibbonEvents.InvalidarRibbon] Recuperacion exitosa"
            Else
                Debug.Print "[clsRibbonEvents.InvalidarRibbon] Recuperacion fallida"
                Exit Sub
            End If
        Else
            Debug.Print "[clsRibbonEvents.InvalidarRibbon] Recuperacion en curso, omitiendo"
            Exit Sub
        End If
    End If

    ' Invalidar el Ribbon
    ribbonUI.Invalidate
    Exit Sub

ErrHandler:
    Debug.Print "[clsRibbonEvents.InvalidarRibbon] Error: " & Err.Description
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
            Debug.Print "[clsRibbonEvents.InvalidarControl] Ribbon aun no inicializado, omitiendo: " & idControl
            Exit Sub
        End If

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
    ribbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    Debug.Print "[clsRibbonEvents.InvalidarControl] Error en control " & idControl & ": " & Err.Description
End Sub

'@Description: Verifica si el objeto ribbonUI esta disponible y funcional
Private Function IsRibbonUIAvailable() As Boolean
    On Error Resume Next

    If ribbonUI Is Nothing Then
        IsRibbonUIAvailable = False
    ElseIf TypeName(ribbonUI) = "Nothing" Or TypeName(ribbonUI) = "Empty" Then
        IsRibbonUIAvailable = False
    Else
        ' Intentar acceder al objeto para verificar que es valido
        Dim testType As String
        testType = TypeName(ribbonUI)
        If Err.Number <> 0 Then
            IsRibbonUIAvailable = False
            Err.Clear
        Else
            IsRibbonUIAvailable = True
        End If
    End If

    On Error GoTo 0
End Function

'@Description: Intenta recuperar el Ribbon automaticamente
Private Function TryAutoRecover() As Boolean
    On Error GoTo ErrHandler

    mIsRecovering = True

    ' Usar el modulo de recuperacion
    TryAutoRecover = modRibbonRecovery.TryRecoverRibbon()

    mIsRecovering = False
    Exit Function

ErrHandler:
    Debug.Print "[clsRibbonEvents.TryAutoRecover] Error: " & Err.Description
    mIsRecovering = False
    TryAutoRecover = False
End Function

'@Description: Indica si el Ribbon esta en proceso de recuperacion
Public Property Get IsRecovering() As Boolean
    IsRecovering = mIsRecovering
End Property

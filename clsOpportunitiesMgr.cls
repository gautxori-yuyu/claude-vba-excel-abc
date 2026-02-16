VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunitiesMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' Clase: clsOpportunitiesMgr
'--------------------------------------------------------------
' Gestiona la lista de "Oportunidades" del sistema.
' Trabaja con identificadores de dominio (nombres/paths),
' NO con indices de UI.
'
' La logica de infraestructura (FileSystem, patrones regex)
' esta delegada a IOpportunityProvider.
'==============================================================

'@Exposed
'@Folder "3-Dominio"
Option Explicit

Private Const MODULE_NAME As String = "clsOpportunitiesMgr"

' ==================================================================
' EVENTOS DE DOMINIO
' ==================================================================

'@Description: Se dispara cuando cambia la oportunidad actual
Public Event CurrentOpportunityChanged(ByVal opportunityName As String, ByVal opportunityPath As String)

'@Description: Se dispara cuando se actualiza la coleccion de oportunidades
Public Event OpportunityCollectionUpdated(ByVal reason As String)

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private mOpportunityProvider As IOpportunityProvider  ' Proveedor de oportunidades (inyectado)
Private mOpportunityState As clsOpportunityState      ' Estado de oportunidad (composicion)
Private mOpportunities As Collection                   ' Collection<clsOpportunity> - construidos por el provider

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    Set mOpportunityState = New clsOpportunityState
    Set mOpportunities = New Collection
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"
    Set mOpportunityState = Nothing
    Set mOpportunities = Nothing
    Set mOpportunityProvider = Nothing
End Sub

'@Description: Inyecta el proveedor de oportunidades y carga datos iniciales
Public Sub Initialize(provider As IOpportunityProvider)
    Set mOpportunityProvider = provider
    RefreshOpportunities
End Sub

' ==================================================================
' PROPIEDADES PUBLICAS
' ==================================================================

'@Description: Expone el estado de oportunidad (para inyeccion en ApplicationState)
Public Property Get State() As clsOpportunityState
    Set State = mOpportunityState
End Property

'@Description: Obtiene el numero de oportunidades disponibles
Public Property Get Count() As Long
    Count = mOpportunities.Count
End Property

'@Description: Obtiene la lista de nombres de oportunidades (para UI)
Public Property Get OpportunityNames() As Collection
    Set OpportunityNames = mOpportunities
End Property

' ==================================================================
' METODOS DE DOMINIO
' ==================================================================

'@Description: Actualiza la coleccion de oportunidades desde el Provider
Public Sub RefreshOpportunities()
    On Error GoTo ErrHandler

    If mOpportunityProvider Is Nothing Then
        LogError MODULE_NAME, "[RefreshOpportunities] Provider no ha sido inyectado"
        Exit Sub
    End If

    ' Obtener oportunidades del Provider como Collection<clsOpportunity> (ya ordenadas)
    ' Cada elemento es un clsOpportunity con Label y path ya inicializados por el provider.
    Set mOpportunities = mOpportunityProvider.GetOpportunities()

    LogInfo MODULE_NAME, "[RefreshOpportunities] " & mOpportunities.Count & " oportunidades cargadas"

    ' Notificar actualizacion
    RaiseEvent OpportunityCollectionUpdated("refresh")

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[RefreshOpportunities]"
End Sub

'@Description: Alias para compatibilidad - usar RefreshOpportunities preferiblemente
Public Function actualizarColeccionOportunidades()
    RefreshOpportunities
End Function

'@Description: Establece la oportunidad actual por nombre
Public Sub SetCurrentOpportunity(ByVal opportunityName As String)
    On Error GoTo ErrHandler

    ' Buscar en la coleccion (items son clsOpportunity pre-construidos por el provider)
    Dim op As clsOpportunity
    Set op = FindInCollection(opportunityName)

    If op Is Nothing Then
        LogWarning MODULE_NAME, "[SetCurrentOpportunity] Oportunidad no encontrada: " & opportunityName
        Exit Sub
    End If

    ' Verificar si es la misma (evitar evento redundante)
    If Not mOpportunityState.CurrentOpportunity Is Nothing Then
        If mOpportunityState.CurrentOpportunity.Label = opportunityName Then
            Exit Sub
        End If
    End If

    ' Actualizar estado con el objeto ya construido (path sellado por el provider)
    Set mOpportunityState.CurrentOpportunity = op

    ' Obtener indice para compatibilidad con UI
    Dim idx As Long
    idx = GetIndexByName(opportunityName)
    mOpportunityState.CurrentIndex = idx

    LogInfo MODULE_NAME, "[SetCurrentOpportunity] -> " & opportunityName

    ' Notificar cambio
    RaiseEvent CurrentOpportunityChanged(opportunityName, op.path)

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[SetCurrentOpportunity]"
End Sub

'@Description: Obtiene una oportunidad por nombre
Public Function GetOpportunityByName(ByVal opportunityName As String) As clsOpportunity
    Set GetOpportunityByName = FindInCollection(opportunityName)
End Function

'@Description: Obtiene una oportunidad cuya ruta coincide con opportunityPath
Public Function GetOpportunityByPath(ByVal opportunityPath As String) As clsOpportunity
    Dim Item As clsOpportunity
    For Each Item In mOpportunities
        If StrComp(Item.path, opportunityPath, vbTextCompare) = 0 Then
            Set GetOpportunityByPath = Item
            Exit Function
        End If
    Next Item
    ' Fallback: intentar por nombre extraido de la ruta
    Dim parts() As String
    parts = Split(opportunityPath, "\")
    If UBound(parts) >= 0 Then
        Set GetOpportunityByPath = FindInCollection(parts(UBound(parts)))
    End If
End Function

'@Description: Verifica si una oportunidad existe en la coleccion
Public Function OpportunityExists(ByVal opportunityName As String) As Boolean
    Dim Item As clsOpportunity
    For Each Item In mOpportunities
        If StrComp(Item.Label, opportunityName, vbTextCompare) = 0 Then
            OpportunityExists = True
            Exit Function
        End If
    Next Item
End Function

'@Description: Procesa cambios detectados por el watcher en carpetas de oportunidad
Public Sub ProcesarCambiosEnOportunidades(ByVal subfolderName As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] Verificando: " & subfolderName

    ' Verificar si el cambio corresponde a una oportunidad valida
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    If regEx.Test(subfolderName) Then
        LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] Cambio en oportunidad detectado"
        RefreshOpportunities
    End If
End Sub

'@Description: Procesa cambios en items dentro de una oportunidad
Public Sub ProcesarCambiosEnItemsOportunidad(ByVal cambios As String)
    ' TODO: Implementar cuando se necesite
End Sub

'@Description: Crea una nueva oportunidad
Public Sub CreateNewOpportunity()
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[CreateNewOpportunity] Iniciando creacion de nueva oportunidad"

    If mOpportunityProvider Is Nothing Then
        LogError MODULE_NAME, "[CreateNewOpportunity] Provider no ha sido inyectado"
        ShowTaskDialogError "Error Crítico", "Error interno: Provider no disponible", _
                            "La aplicación no puede continuar porque un componente esencial (IOpportunityProvider) no se ha iniciado."
        Exit Sub
    End If

    ' Verificar que el almacenamiento esta disponible (usa flag, no re-consulta el FS)
    If Not mOpportunityProvider.StorageAvailable() Then
        ShowTaskDialogError "Ruta no encontrada", "No se pudo abrir la ruta de oportunidades.", _
                            "La carpeta de oportunidades no esta accesible. Verifique la configuracion de rutas."
        Exit Sub
    End If

    ' Obtener siguiente codigo
    Dim strOpCode As String
    strOpCode = mOpportunityProvider.GetNextOpportunityCode()

    ' Solicitar nombre de cliente
    Dim strCustomer As String
    strCustomer = GetCustomerNameFromUser(strOpCode)

    If strCustomer = "" Then Exit Sub ' User cancelled

    ' Construir nombre completo
    Dim opportunityName As String
    opportunityName = strOpCode & " - " & strCustomer & " - XXX"

    ' Crear carpeta y refrescar la lista inmediatamente.
    ' Nota: el FSWatcher tambien disparara ProcesarCambiosEnOportunidades cuando detecte
    ' la nueva carpeta, lo que causara un segundo RefreshOpportunities (idempotente, aceptable).
    If mOpportunityProvider.CreateOpportunity(opportunityName) Then
        LogInfo MODULE_NAME, "[CreateNewOpportunity] Oportunidad creada: " & opportunityName
        RefreshOpportunities
    Else
        ShowTaskDialogError "Error de Creación", "Error al crear la carpeta de la oportunidad.", _
                            "VBA no pudo crear la carpeta en el sistema de archivos. Revise los permisos o el log para más detalles."
    End If

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[CreateNewOpportunity]"
    ShowTaskDialogError "Error Inesperado", "Ocurrió un error inesperado al crear la oportunidad.", _
                        "Error: " & Err.Description & " (Nº " & Err.Number & ")"
End Sub

' ==================================================================
' METODOS PARA COMPATIBILIDAD CON UI (TEMPORALES)
' ==================================================================
' Estos metodos mantienen compatibilidad con el Ribbon actual.
' En una refactorizacion futura, el Ribbon deberia trabajar solo con nombres.

'@Description: Obtiene el nombre de una oportunidad por indice (para UI)
Public Function GetOpportunityNameByIndex(ByVal idx As Long) As String
    If idx >= 0 And idx < mOpportunities.Count Then
        Dim op As clsOpportunity
        Set op = mOpportunities(idx + 1)
        GetOpportunityNameByIndex = op.Label
    Else
        GetOpportunityNameByIndex = "(Sin datos)"
    End If
End Function

'@Description: Obtiene la ruta de una oportunidad por indice (para UI)
Public Function GetOpportunityPathByIndex(ByVal idx As Long) As String
    If idx >= 0 And idx < mOpportunities.Count Then
        Dim op As clsOpportunity
        Set op = mOpportunities(idx + 1)
        GetOpportunityPathByIndex = op.path
    End If
End Function

'@Description: Establece la oportunidad actual por indice (para UI)
Public Sub SetCurrentOpportunityByIndex(ByVal idx As Long)
    If idx >= 0 And idx < mOpportunities.Count Then
        Dim op As clsOpportunity
        Set op = mOpportunities(idx + 1)
        SetCurrentOpportunity op.Label
    End If
End Sub

'@Description: Obtiene el indice de la oportunidad actual (para UI)
Public Property Get CurrentIndex() As Long
    CurrentIndex = mOpportunityState.CurrentIndex
End Property

' ==================================================================
' METODOS PRIVADOS
' ==================================================================

'@Description: Busca un clsOpportunity en la coleccion por nombre (Label)
'@Note: Sustituye a CreateOpportunityFromName - ya no construye el objeto,
'       lo recupera directamente de la coleccion pre-construida por el provider.
Private Function FindInCollection(ByVal opportunityName As String) As clsOpportunity
    Dim Item As clsOpportunity
    For Each Item In mOpportunities
        If StrComp(Item.Label, opportunityName, vbTextCompare) = 0 Then
            Set FindInCollection = Item
            Exit Function
        End If
    Next Item
    ' Devuelve Nothing si no encontrado
End Function

'@Description: Obtiene el indice de una oportunidad por nombre
Private Function GetIndexByName(ByVal opportunityName As String) As Long
    GetIndexByName = -1
    Dim i As Long
    Dim op As clsOpportunity
    For i = 1 To mOpportunities.Count
        Set op = mOpportunities(i)
        If StrComp(op.Label, opportunityName, vbTextCompare) = 0 Then
            GetIndexByName = i - 1
            Exit Function
        End If
    Next i
End Function

'@Description: Solicita al usuario el nombre del cliente/proyecto de forma interactiva y robusta.
Private Function GetCustomerNameFromUser(ByVal opCode As String) As String
    Dim strCustomer As String
    Dim IsValid As Boolean
    Dim userResponse As Long
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN

    Dim TaskDlg As cTaskDialog
    Set TaskDlg = New cTaskDialog
    With TaskDlg
        .Init
        .Title = "Crear Nueva Oportunidad"
        .MainInstruction = "Nombre del cliente y proyecto"
        .Content = "Introduce el nombre del cliente, o cliente - proyecto, separados por un guión. No utilizar caracteres especiales!"
        .Flags = TDF_INPUT_BOX
        .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
        .IconMain = TD_INFORMATION_ICON
        .ParenthWnd = Application.hwnd
                
        Do
            ' 1. Solicitar input al usuario
            .ShowDialog
        
            strCustomer = .ResultInput
            If .ResultMain <> TD_OK Then
                ' 2. Si el usuario cancela (InputBox devuelve ""), salir
                GetCustomerNameFromUser = ""
                Exit Function
            Else
                ' 3. Validar el nombre introducido
                IsValid = regEx.Test(opCode & " - " & strCustomer & " - XXX")
                ' Evitamos caracteres especiales del sistema de archivos
                IsValid = IsValid And Not (strCustomer Like "*[\/:*?""<>|]*")
                If Not IsValid Then
                    .MainInstruction = "Nombre NO VALIDO"
                    .IconMain = TD_WARNING_ICON
                Else
                    Exit Do
                End If
            End If
        Loop While Not IsValid
    End With
    GetCustomerNameFromUser = strCustomer
End Function

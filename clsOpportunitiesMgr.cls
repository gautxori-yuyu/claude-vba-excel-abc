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
'@Folder "6-DOMINIO-Oportunidades y compresores"
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
Private mOpportunities As Collection                   ' Lista de nombres de oportunidades

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

'@Description: Obtiene la ruta base de oportunidades
Public Property Get BaseFolder() As String
    If mOpportunityProvider Is Nothing Then Exit Property
    BaseFolder = mOpportunityProvider.BasePath
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

    ' Obtener oportunidades del Provider (ya ordenadas)
    Set mOpportunities = mOpportunityProvider.GetOpportunityFolders()

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

    ' Verificar que existe en la coleccion
    If Not OpportunityExists(opportunityName) Then
        LogWarning MODULE_NAME, "[SetCurrentOpportunity] Oportunidad no encontrada: " & opportunityName
        Exit Sub
    End If

    ' Verificar si es la misma (evitar evento redundante)
    If Not mOpportunityState.CurrentOpportunity Is Nothing Then
        If mOpportunityState.CurrentOpportunity.Label = opportunityName Then
            Exit Sub
        End If
    End If

    ' Crear objeto clsOpportunity
    Dim op As clsOpportunity
    Set op = CreateOpportunityFromName(opportunityName)

    ' Actualizar estado
    Set mOpportunityState.CurrentOpportunity = op

    ' Obtener indice para compatibilidad temporal
    Dim idx As Long
    idx = GetIndexByName(opportunityName)
    mOpportunityState.CurrentIndex = idx

    LogInfo MODULE_NAME, "[SetCurrentOpportunity] -> " & opportunityName

    ' Notificar cambio
    RaiseEvent CurrentOpportunityChanged(opportunityName, op.Path)

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[SetCurrentOpportunity]"
End Sub

'@Description: Obtiene una oportunidad por nombre
Public Function GetOpportunityByName(ByVal opportunityName As String) As clsOpportunity
    If OpportunityExists(opportunityName) Then
        Set GetOpportunityByName = CreateOpportunityFromName(opportunityName)
    End If
End Function

'@Description: Obtiene una oportunidad por su ruta completa
Public Function GetOpportunityByPath(ByVal opportunityPath As String) As clsOpportunity
    ' Extraer nombre de la ruta
    Dim parts() As String
    parts = Split(opportunityPath, "\")
    If UBound(parts) >= 0 Then
        Set GetOpportunityByPath = GetOpportunityByName(parts(UBound(parts)))
    End If
End Function

'@Description: Verifica si una oportunidad existe en la coleccion
Public Function OpportunityExists(ByVal opportunityName As String) As Boolean
    Dim item As Variant
    For Each item In mOpportunities
        If StrComp(CStr(item), opportunityName, vbTextCompare) = 0 Then
            OpportunityExists = True
            Exit Function
        End If
    Next item
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

    If mOpportunityProvider Is Nothing Then
        LogError MODULE_NAME, "[CreateNewOpportunity] Provider no ha sido inyectado"
        MsgBox "Error interno: Provider no disponible", vbCritical
        Exit Sub
    End If

    ' Verificar que la carpeta base existe
    If Not mOpportunityProvider.FolderExists("") Then
        MsgBox "No se pudo abrir la ruta de oportunidades: " & BaseFolder, vbExclamation
        Exit Sub
    End If

    ' Obtener siguiente codigo
    Dim strOpCode As String
    strOpCode = mOpportunityProvider.GetNextOpportunityCode()

    ' Solicitar nombre de cliente
    Dim strCustomer As String
    strCustomer = GetCustomerNameFromUser(strOpCode)

    If strCustomer = "" Then Exit Sub

    ' Construir nombre completo
    Dim opportunityName As String
    opportunityName = strOpCode & " - " & strCustomer & " - XXX"

    ' Crear carpeta
    If mOpportunityProvider.CreateOpportunityFolder(opportunityName) Then
        LogInfo MODULE_NAME, "[CreateNewOpportunity] Oportunidad creada: " & opportunityName
        RefreshOpportunities
    Else
        MsgBox "Error al crear la oportunidad. Revise el log para mas detalles.", vbExclamation
    End If

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[CreateNewOpportunity]"
    MsgBox "No se pudo crear la oportunidad: " & Err.Description, vbExclamation
End Sub

' ==================================================================
' METODOS PARA COMPATIBILIDAD CON UI (TEMPORALES)
' ==================================================================
' Estos metodos mantienen compatibilidad con el Ribbon actual.
' En una refactorizacion futura, el Ribbon deberia trabajar solo con nombres.

'@Description: Obtiene el nombre de una oportunidad por indice (para UI)
Public Function GetOpportunityNameByIndex(ByVal idx As Long) As String
    If idx >= 0 And idx < mOpportunities.Count Then
        GetOpportunityNameByIndex = mOpportunities(idx + 1)
    Else
        GetOpportunityNameByIndex = "(Sin datos)"
    End If
End Function

'@Description: Obtiene la ruta de una oportunidad por indice (para UI)
Public Function GetOpportunityPathByIndex(ByVal idx As Long) As String
    If idx >= 0 And idx < mOpportunities.Count Then
        GetOpportunityPathByIndex = mOpportunityProvider.GetFullPath(mOpportunities(idx + 1))
    End If
End Function

'@Description: Establece la oportunidad actual por indice (para UI)
Public Sub SetCurrentOpportunityByIndex(ByVal idx As Long)
    If idx >= 0 And idx < mOpportunities.Count Then
        SetCurrentOpportunity mOpportunities(idx + 1)
    End If
End Sub

'@Description: Obtiene el indice de la oportunidad actual (para UI)
Public Property Get CurrentIndex() As Long
    CurrentIndex = mOpportunityState.CurrentIndex
End Property

' ==================================================================
' METODOS PRIVADOS
' ==================================================================

'@Description: Crea un objeto clsOpportunity desde un nombre
Private Function CreateOpportunityFromName(ByVal opportunityName As String) As clsOpportunity
    Dim op As clsOpportunity
    Set op = New clsOpportunity

    op.Label = opportunityName
    op.Path = mOpportunityProvider.GetFullPath(opportunityName)

    ' Extraer informacion del nombre
    ' Formato tipico: "412601015 - Cliente - Proyecto - Modelo"
    On Error Resume Next
    Dim parts() As String
    parts = Split(opportunityName, " - ")
    If UBound(parts) >= 0 Then
        op.Name = Trim(parts(0))
    End If
    If UBound(parts) >= 1 Then
        op.Customer = Trim(parts(1))
    End If
    On Error GoTo 0

    Set CreateOpportunityFromName = op
End Function

'@Description: Obtiene el indice de una oportunidad por nombre
Private Function GetIndexByName(ByVal opportunityName As String) As Long
    GetIndexByName = -1
    Dim i As Long
    For i = 1 To mOpportunities.Count
        If StrComp(mOpportunities(i), opportunityName, vbTextCompare) = 0 Then
            GetIndexByName = i - 1
            Exit Function
        End If
    Next i
End Function

'@Description: Solicita al usuario el nombre del cliente/proyecto
Private Function GetCustomerNameFromUser(ByVal opCode As String) As String
    Dim strCustomer As String
    Dim strMsg As String

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN

    Do
        strCustomer = InputBox(strMsg & "Introduce el nombre del cliente, o cliente - proyecto:", _
                               "Cliente - Proyecto", strCustomer)
        strMsg = "NOMBRE NO VALIDO. "
    Loop While strCustomer <> "" And Not regEx.Test(opCode & " - " & strCustomer & " - XXX")

    GetCustomerNameFromUser = strCustomer
End Function

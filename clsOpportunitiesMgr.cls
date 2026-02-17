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
' La lógica de infraestructura (FileSystem, patrones regex)
' está delegada a FileSystemOpportunityProvider.
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

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private WithEvents mProvider As FileSystemProvider  ' Proveedor (para llamadas Y WithEvents)
Attribute mProvider.VB_VarHelpID = -1
Private mOpportunityState As clsOpportunityState      ' Estado de oportunidad (composición)
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
    Set mProvider = Nothing
End Sub

'@Description: Inyecta el proveedor de oportunidades y carga datos iniciales.
'              Se usa la clase concreta (no la interfaz) porque WithEvents
'              no funciona con interfaces en VBA.
Public Sub Initialize(provider As FileSystemProvider)
    Set mProvider = provider
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

    If mProvider Is Nothing Then
        LogError MODULE_NAME, "[RefreshOpportunities] Provider no ha sido inyectado"
        Exit Sub
    End If

    ' Obtener oportunidades del Provider como Collection<clsOpportunity> (ya ordenadas)
    ' Cada elemento es un clsOpportunity con Label y path ya inicializados por el provider.
    Set mOpportunities = mProvider.GetOpportunities()

    LogInfo MODULE_NAME, "[RefreshOpportunities] " & mOpportunities.Count & " oportunidades cargadas"

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[RefreshOpportunities]"
End Sub

'@Description: Alias para compatibilidad - usar RefreshOpportunities preferiblemente
Public Function actualizarColeccionOportunidades()
    RefreshOpportunities
End Function

'@Description: Agrega una oportunidad a la coleccion si no existe ya (idempotente por Number).
'              Llamado por el provider de infraestructura al detectar carpeta valida.
'              Mantiene el orden descendente por Number.
Public Sub AddOpportunity(ByVal op As clsOpportunity)
    Const PROC_NAME As String = "AddOpportunity"
    On Error GoTo ErrHandler
    If op Is Nothing Then Exit Sub

    ' Idempotencia: verificar si ya existe por Number
    If op.Number <> 0 Then
        Dim existing As clsOpportunity
        For Each existing In mOpportunities
            If existing.Number = op.Number Then
                LogDebug MODULE_NAME, "[" & PROC_NAME & "] Ya existe Number=" & op.Number & ", ignorada"
                Exit Sub
            End If
        Next existing
    End If

    ' Insercion ordenada: buscar posicion (coleccion es descendente por Number)
    Dim insertBefore As Long
    insertBefore = 0
    Dim i As Long
    For i = 1 To mOpportunities.Count
        Dim candidate As clsOpportunity
        Set candidate = mOpportunities(i)
        If candidate.Number < op.Number Then
            insertBefore = i
            Exit For
        End If
    Next i

    If insertBefore = 0 Then
        mOpportunities.Add op               ' Agregar al final (numero menor o igual a todos)
    Else
        mOpportunities.Add op, , insertBefore   ' Insertar antes de la primera de menor numero
    End If

    LogInfo MODULE_NAME, "[" & PROC_NAME & "] Anadida: " & op.Label
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[" & PROC_NAME & "]"
End Sub

'@Description: Elimina una oportunidad de la coleccion por su codigo de 9 digitos.
'              Si era la oportunidad actual, limpia el estado de seleccion.
Public Sub RemoveOpportunity(ByVal opportunityCode As String)
    Const PROC_NAME As String = "RemoveOpportunity"
    On Error GoTo ErrHandler
    If Len(opportunityCode) = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To mOpportunities.Count
        Dim op As clsOpportunity
        Set op = mOpportunities(i)
        If CStr(op.Number) = opportunityCode Then
            Dim removedLabel As String
            removedLabel = op.Label
            ' Si era la oportunidad actual, limpiar estado
            If Not mOpportunityState.CurrentOpportunity Is Nothing Then
                If mOpportunityState.CurrentOpportunity.Label = op.Label Then
                    Set mOpportunityState.CurrentOpportunity = Nothing
                    mOpportunityState.CurrentIndex = -1
                    LogInfo MODULE_NAME, "[" & PROC_NAME & "] Estado de oportunidad actual limpiado"
                End If
            End If
            mOpportunities.Remove i
            LogInfo MODULE_NAME, "[" & PROC_NAME & "] Eliminada: " & removedLabel
            Exit Sub
        End If
    Next i

    LogWarning MODULE_NAME, "[" & PROC_NAME & "] Codigo no encontrado en coleccion: " & opportunityCode
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[" & PROC_NAME & "]"
End Sub

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

'@Description: Crea una nueva oportunidad delegando al provider.
'              El provider gestiona: código, input de usuario, validación, creación FS, entidad.
'              El dominio solo recibe la entidad limpia y la agrega a la colección.
'              El provider suprime el evento duplicado del FSWatcher (mSuppressNextDetection).
Public Sub CreateNewOpportunity()
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[CreateNewOpportunity] Delegando al provider"

    If mProvider Is Nothing Then
        LogError MODULE_NAME, "[CreateNewOpportunity] Provider no ha sido inyectado"
        Exit Sub
    End If

    Dim op As clsOpportunity
    Set op = mProvider.CreateNewOpportunity()

    If op Is Nothing Then
        LogDebug MODULE_NAME, "[CreateNewOpportunity] Provider devolvio Nothing (cancelado o error)"
        Exit Sub
    End If

    AddOpportunity op
    LogInfo MODULE_NAME, "[CreateNewOpportunity] Oportunidad creada: " & op.Label

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[CreateNewOpportunity]"
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

' ==================================================================
' MANEJADORES DE EVENTOS DEL PROVIDER (via WithEvents)
' Recibe objetos de dominio ya validados y construidos por infraestructura.
' ==================================================================

'@Description: Nueva oportunidad detectada en el sistema de archivos.
'              El provider ya valido el patron y construyo la entidad.
Private Sub mProvider_OpportunityDetected(ByVal op As clsOpportunity)
    LogInfo MODULE_NAME, "[mProvider_OpportunityDetected] " & op.Label
    AddOpportunity op
End Sub

'@Description: Oportunidad eliminada del sistema de archivos.
Private Sub mProvider_OpportunityRemoved(ByVal opportunityCode As String)
    LogInfo MODULE_NAME, "[mProvider_OpportunityRemoved] Codigo: " & opportunityCode
    RemoveOpportunity opportunityCode
End Sub

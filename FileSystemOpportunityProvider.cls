VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemOpportunityProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: FileSystemOpportunityProvider
' ==================================================================
' Implementa IOpportunityProvider usando el sistema de archivos.
' Contiene toda la logica de:
' - Listado de subcarpetas
' - Filtrado por patrones regex (FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN)
' - Ordenacion para presentacion
' - Creacion de carpetas desde plantilla
' - Validacion de nombres de oportunidad
' - Solicitud interactiva de datos al usuario (via mTDHelper)
'
' El dominio no conoce estos detalles de implementacion.
' ==================================================================

'@Folder "2-Infraestructura.5-2-Providers.FS"
Option Explicit
Implements IOpportunityProvider

Private Const MODULE_NAME As String = "FileSystemOpportunityProvider"

' ==================================================================
' DEPENDENCIAS
' ==================================================================
Private mFileManager As clsFileManager
Private mConfiguration As clsConfiguration
Private mIsReady As Boolean   ' Flag: BasePath era accesible al inicializar
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord

' ==================================================================
' EVENTOS DE DOMINIO (emitidos hacia suscriptores WithEvents)
' ==================================================================
Public Event OpportunityDetected(ByVal op As clsOpportunity)
Public Event OpportunityRemoved(ByVal opportunityCode As String)

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    mIsReady = False
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"
    Set mFSMonitoringCoord = Nothing
    Set mFileManager = Nothing
    Set mConfiguration = Nothing
End Sub

'@Description: Inyecta las dependencias e inicializa el flag de disponibilidad
'@Param fm: Gestor de sistema de archivos
'@Param oConfiguration: Configuracion de la aplicacion
'@Param oFSMonitoringCoord: Coordinador de monitoreo FS (opcional, para eventos en tiempo real)
Public Sub Initialize(fm As clsFileManager, _
                      oConfiguration As clsConfiguration, _
                      Optional oFSMonitoringCoord As clsFSMonitoringCoord = Nothing)
    On Error GoTo ErrHandler

    Set mFileManager = fm
    Set mConfiguration = oConfiguration

    ' Comprobar disponibilidad UNA SOLA VEZ: evita re-consultas redundantes al FS
    Dim strBasePath As String
    strBasePath = mConfiguration.RutaOportunidades
    mIsReady = Len(strBasePath) > 0 And mFileManager.ExisteCarpeta(strBasePath)

    ' Suscribirse al coordinador de monitoreo para detectar cambios en tiempo real
    If Not oFSMonitoringCoord Is Nothing Then
        Set mFSMonitoringCoord = oFSMonitoringCoord
        LogDebug MODULE_NAME, "[Initialize] FSMonitoringCoord inyectado"
    End If

    LogDebug MODULE_NAME, "[Initialize] FileManager y Configuration inyectados. StorageAvailable=" & mIsReady
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    mIsReady = False
End Sub

' ==================================================================
' IMPLEMENTACION DE IOpportunityProvider
' Stubs privados que delegan en los metodos publicos concretos.
' ==================================================================

Private Function IOpportunityProvider_GetOpportunities() As Collection
    Set IOpportunityProvider_GetOpportunities = BuildOpportunityObjects()
End Function

Private Function IOpportunityProvider_OpportunityExists(ByVal opportunityName As String) As Boolean
    IOpportunityProvider_OpportunityExists = FolderExists(opportunityName)
End Function

Private Function IOpportunityProvider_CreateNewOpportunity() As clsOpportunity
    Set IOpportunityProvider_CreateNewOpportunity = CreateNewOpportunity()
End Function

Private Function IOpportunityProvider_CreateOpportunity(ByVal opportunityName As String) As clsOpportunity
    Set IOpportunityProvider_CreateOpportunity = CreateOpportunity(opportunityName)
End Function

Private Function IOpportunityProvider_GetNextOpportunityCode() As String
    IOpportunityProvider_GetNextOpportunityCode = GetNextOpportunityCode()
End Function

Private Function IOpportunityProvider_StorageAvailable() As Boolean
    IOpportunityProvider_StorageAvailable = StorageAvailable()
End Function

Private Function IOpportunityProvider_ValidateOpportunityName(ByVal opportunityName As String) As Boolean
    IOpportunityProvider_ValidateOpportunityName = ValidateOpportunityName(opportunityName)
End Function

' ==================================================================
' METODOS PUBLICOS (implementacion concreta)
' ==================================================================

'@Description: Indica si el almacenamiento (BasePath) estaba disponible al inicializar.
Public Function StorageAvailable() As Boolean
    StorageAvailable = mIsReady
End Function

'@Description: Valida si un nombre cumple el patron de oportunidad valida
Public Function ValidateOpportunityName(ByVal opportunityName As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    ValidateOpportunityName = IsValidOpportunityName(opportunityName, regEx)
End Function

'@Description: Crea una nueva oportunidad interactivamente.
'              Gestiona: codigo, input de usuario, validacion regex, creacion FS, construccion de entidad.
'@Returns: clsOpportunity creada, o Nothing si el usuario cancela o hay error
Public Function CreateNewOpportunity() As clsOpportunity
    Const PROC_NAME As String = "CreateNewOpportunity"
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[" & PROC_NAME & "] Iniciando creacion de nueva oportunidad"

    ' 1. Verificar disponibilidad del almacenamiento
    If Not mIsReady Then
        ShowTaskDialogError "Ruta no encontrada", _
                            "No se pudo abrir la ruta de oportunidades.", _
                            "La carpeta de oportunidades no esta accesible. Verifique la configuracion de rutas."
        Exit Function
    End If

    ' 2. Obtener siguiente codigo de oportunidad
    Dim strOpCode As String
    strOpCode = GetNextOpportunityCode()
    If Len(strOpCode) = 0 Then
        ShowTaskDialogError "Error", _
                            "No se pudo generar un codigo de oportunidad.", _
                            "Revise el log para mas detalles."
        Exit Function
    End If

    ' 3. Solicitar nombre de cliente al usuario (con validacion)
    Dim strCustomer As String
    strCustomer = GetCustomerNameFromUser(strOpCode)
    If Len(strCustomer) = 0 Then Exit Function  ' Usuario cancelo

    ' 4. Construir nombre completo y crear
    Dim opportunityName As String
    opportunityName = BuildOpportunityFolderName(strOpCode, strCustomer)

    Set CreateNewOpportunity = CreateOpportunity(opportunityName)

    Exit Function
ErrHandler:
    LogCurrentError MODULE_NAME, "[" & PROC_NAME & "]"
    ShowTaskDialogError "Error Inesperado", _
                        "Ocurrio un error inesperado al crear la oportunidad.", _
                        "Error: " & Err.Description & " (Nr " & Err.Number & ")"
End Function

'@Description: Crea una oportunidad con el nombre dado (sin interaccion de usuario).
'              Copia plantilla al destino y devuelve la entidad construida.
'@Returns: clsOpportunity creada, o Nothing si falla
Public Function CreateOpportunity(ByVal opportunityName As String) As clsOpportunity
    Const PROC_NAME As String = "CreateOpportunity"
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[" & PROC_NAME & "] Creando: " & opportunityName

    If Not mIsReady Then
        LogError MODULE_NAME, "[" & PROC_NAME & "] BasePath no disponible"
        Exit Function
    End If

    ' Obtener ruta de plantilla
    Dim rutaOrigen As String
    rutaOrigen = mConfiguration.RutaPlantillas
    If Right(rutaOrigen, 1) = "\" Then rutaOrigen = Left(rutaOrigen, Len(rutaOrigen) - 1)

    ' Construir ruta destino
    Dim rutaDestino As String
    rutaDestino = mFileManager.ConstruirRuta(BasePath, opportunityName)

    ' Copiar carpeta de plantilla
    If mFileManager.CopiarCarpeta(rutaOrigen, rutaDestino) Then
        LogInfo MODULE_NAME, "[" & PROC_NAME & "] Oportunidad creada: " & opportunityName
        Set CreateOpportunity = BuildSingleOpportunity(opportunityName)
    Else
        LogError MODULE_NAME, "[" & PROC_NAME & "] Fallo al copiar plantilla"
    End If

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[" & PROC_NAME & "]"
End Function

'@Description: Obtiene los nombres de carpetas de oportunidad validas, ordenados
Public Function GetOpportunityFolders() As Collection
    Dim result As Collection
    Set result = New Collection

    On Error GoTo ErrHandler

    If mFileManager Is Nothing Then
        LogError MODULE_NAME, "[GetOpportunityFolders] FileManager no ha sido inyectado"
        Set GetOpportunityFolders = result
        Exit Function
    End If

    ' Obtener todas las subcarpetas del directorio base
    Dim subcarpetas As Collection
    Set subcarpetas = mFileManager.ListarSubcarpetas(BasePath)

    If subcarpetas.Count = 0 Then
        LogWarning MODULE_NAME, "[GetOpportunityFolders] No se encontraron subcarpetas en: " & BasePath
        Set GetOpportunityFolders = result
        Exit Function
    End If

    ' Filtrar carpetas validas usando patron de oportunidad
    Dim arr() As String
    ReDim arr(subcarpetas.Count - 1)
    Dim i As Long
    i = 0

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    Dim nombreCarpeta As Variant
    For Each nombreCarpeta In subcarpetas
        If IsValidOpportunityName(CStr(nombreCarpeta), regEx) Then
            arr(i) = CStr(nombreCarpeta)
            i = i + 1
        End If
    Next nombreCarpeta

    ' Si hay carpetas validas, ordenar y agregar a la coleccion
    If i > 0 Then
        ReDim Preserve arr(i - 1)
        arr = SortFoldersDescending(arr)

        Dim j As Long
        For j = LBound(arr) To UBound(arr)
            result.Add arr(j)
        Next j
    End If

    LogDebug MODULE_NAME, "[GetOpportunityFolders] Encontradas " & result.Count & " oportunidades validas"

    Set GetOpportunityFolders = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetOpportunityFolders]"
    Set GetOpportunityFolders = result
End Function

'@Description: Verifica si una carpeta de oportunidad existe
Public Function FolderExists(ByVal folderName As String) As Boolean
    If mFileManager Is Nothing Then Exit Function
    FolderExists = mFileManager.ExisteCarpeta(GetFullPath(folderName))
End Function

'@Description: Obtiene la ruta base de oportunidades
Public Property Get BasePath() As String
    BasePath = mConfiguration.RutaOportunidades
End Property

'@Description: Construye la ruta completa de una oportunidad
Public Function GetFullPath(ByVal folderName As String) As String
    GetFullPath = BasePath & "\" & folderName
End Function

'@Description: Genera el siguiente codigo de oportunidad disponible
Public Function GetNextOpportunityCode() As String
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[GetNextOpportunityCode] Calculando siguiente codigo de oportunidad"

    Dim folders As Collection
    Set folders = GetOpportunityFolders()

    Dim arrOps() As String
    ReDim arrOps(folders.Count - 1)

    Dim i As Long
    Dim folder As Variant
    For Each folder In folders
        arrOps(i) = Left(CStr(folder), 9)
        i = i + 1
    Next folder

    ' Obtener la ultima secuencia
    Dim dblLastOpSeq As Double
    dblLastOpSeq = 0
    If folders.Count > 0 Then
        dblLastOpSeq = CDbl(Mid(arrOps(0), 7))
    End If

    ' Generar siguiente codigo
    Dim strOpName As String
    Do
        dblLastOpSeq = dblLastOpSeq + 1
        strOpName = mConfiguration.SAM & _
                    Mid(Year(Now), 3) & _
                    String(2 - Len(CStr(Month(Now))), "0") & Month(Now) & _
                    String(3 - Len(CStr(dblLastOpSeq)), "0") & dblLastOpSeq
    Loop While UBound(Filter(arrOps, strOpName)) >= 0

    GetNextOpportunityCode = strOpName
    LogDebug MODULE_NAME, "[GetNextOpportunityCode] Codigo generado: " & strOpName

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetNextOpportunityCode]"
    GetNextOpportunityCode = ""
End Function

'@Description: Construye la coleccion de oportunidades como objetos clsOpportunity
'@Note: Es la implementacion correcta de GetOpportunities() via IOpportunityProvider.
'       El path queda sellado dentro de cada clsOpportunity, evitando que el dominio
'       tenga que llamar GetOpportunityPath() para reconstruirlo.
Public Function BuildOpportunityObjects() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim folders As Collection
    Set folders = GetOpportunityFolders()

    Dim folderName As Variant
    For Each folderName In folders
        Dim op As clsOpportunity
        Set op = BuildSingleOpportunity(CStr(folderName))
        result.Add op
        Set op = Nothing
    Next folderName

    Set BuildOpportunityObjects = result
End Function

' ==================================================================
' METODOS PRIVADOS - LOGICA DE INFRAESTRUCTURA
' ==================================================================

'@Description: Verifica si un nombre de carpeta cumple el patron de oportunidad
Private Function IsValidOpportunityName(ByVal folderName As String, regEx As Object) As Boolean
    ' Primero probar patron completo (con modelo)
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    If regEx.Test(folderName) Then
        IsValidOpportunityName = True
        Exit Function
    End If

    ' Si no cumple, probar patron sin modelo
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN
    If regEx.Test(folderName) Then
        IsValidOpportunityName = True
    End If
End Function

'@Description: Ordena array de carpetas en orden numerico descendente
Private Function SortFoldersDescending(arr() As String) As String()
    Dim i As Long, j As Long, tmp As String
    Dim keyI As Double, keyJ As Double

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            keyI = ExtractNumericKey(arr(i))
            keyJ = ExtractNumericKey(arr(j))

            If keyI < keyJ Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            ElseIf keyI = keyJ Then
                If StrComp(arr(i), arr(j), vbTextCompare) < 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            End If
        Next j
    Next i

    SortFoldersDescending = arr
End Function

'@Description: Extrae el ultimo numero del nombre de la carpeta
Private Function ExtractNumericKey(ByVal folderName As String) As Double
    On Error GoTo ErrHandler

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    re.Global = True

    If re.Test(folderName) Then
        Dim matches As Object
        Set matches = re.Execute(folderName)
        ExtractNumericKey = CDbl(matches(matches.Count - 1).Value)
        Exit Function
    End If

ErrHandler:
    ExtractNumericKey = -1E+99
End Function

'@Description: Construye un clsOpportunity desde un nombre de carpeta (ya validado o no).
'              El path queda sellado dentro del objeto (patron Sealed Path).
Private Function BuildSingleOpportunity(ByVal folderName As String) As clsOpportunity
    Dim op As clsOpportunity
    Set op = New clsOpportunity
    op.Label = folderName
    op.path = GetFullPath(folderName)

    On Error Resume Next
    Dim parts() As String
    parts = Split(folderName, " - ")
    If UBound(parts) >= 0 Then op.Number = CLng(Trim(parts(0)))
    If UBound(parts) >= 1 Then op.Customer = Trim(parts(1))
    If UBound(parts) >= 2 Then op.Project = Trim(parts(2))
    On Error GoTo 0

    Set BuildSingleOpportunity = op
End Function

'@Description: Extrae los primeros 9 digitos del nombre de carpeta como codigo.
'              Fallback: devuelve el nombre completo si no hay segmento numerico inicial.
Private Function ExtractOpportunityCode(ByVal folderName As String) As String
    On Error GoTo Fallback
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{9}"
    If re.Test(folderName) Then
        ExtractOpportunityCode = re.Execute(folderName)(0).Value
        Exit Function
    End If
Fallback:
    ExtractOpportunityCode = folderName
End Function

'@Description: Construye el nombre de carpeta de oportunidad a partir de sus componentes.
'              El formato del nombre es convencion de infraestructura, no de dominio.
Private Function BuildOpportunityFolderName(ByVal opCode As String, ByVal customer As String) As String
    BuildOpportunityFolderName = opCode & " - " & customer & " - XXX"
End Function

'@Description: Solicita al usuario el nombre de cliente/proyecto de forma interactiva con validacion.
'              Usa cTaskDialog via mTDHelper. La validacion regex se aplica aqui en infraestructura.
'@Returns: Nombre de cliente introducido, o cadena vacia si el usuario cancela
Private Function GetCustomerNameFromUser(ByVal opCode As String) As String
    Const PROC_NAME As String = "GetCustomerNameFromUser"
    Dim strCustomer As String
    Dim IsValid As Boolean

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    On Error GoTo ErrHandler

    Dim TaskDlg As cTaskDialog
    Set TaskDlg = New cTaskDialog
    With TaskDlg
        .Init
        .Title = "Crear Nueva Oportunidad"
        .MainInstruction = "Nombre del cliente y proyecto"
        .Content = "Introduce el nombre del cliente, o cliente - proyecto, separados por un guion. No utilizar caracteres especiales!"
        .Flags = TDF_INPUT_BOX
        .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
        .IconMain = TD_INFORMATION_ICON
        .ParenthWnd = Application.hwnd

        Do
            .ShowDialog

            strCustomer = .ResultInput
            If .ResultMain <> TD_OK Then
                GetCustomerNameFromUser = ""
                Exit Function
            Else
                ' Validar con regex (logica de infraestructura, no de dominio)
                Dim testName As String
                testName = BuildOpportunityFolderName(opCode, strCustomer)
                IsValid = IsValidOpportunityName(testName, regEx)
                ' Evitar caracteres especiales del sistema de archivos
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

    LogDebug MODULE_NAME, "[" & PROC_NAME & "] Cliente (+ proyecto): " & strCustomer

    Exit Function
ErrHandler:
    LogCurrentError MODULE_NAME, "[" & PROC_NAME & "]"
    ShowTaskDialogError "Error Inesperado", _
                        "Ocurrio un error al solicitar el nombre de cliente.", _
                        "Error: " & Err.Description & " (Nr " & Err.Number & ")"
End Function

' ==================================================================
' MANEJADORES DE EVENTOS clsFSMonitoringCoord
' La validacion regex permanece aqui en infraestructura.
' El dominio recibe objetos clsOpportunity ya construidos.
' ==================================================================

'@Description: Detecta creacion de subcarpeta en ruta de oportunidades.
'              Valida patron, construye objeto de dominio y emite OpportunityDetected.
Private Sub mFSMonitoringCoord_OpportunityCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[OpportunityCreated] Evaluando: " & subfolderName

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    If Not IsValidOpportunityName(subfolderName, regEx) Then
        LogDebug MODULE_NAME, "[OpportunityCreated] Ignorado (no cumple patron): " & subfolderName
        Exit Sub
    End If

    Dim op As clsOpportunity
    Set op = BuildSingleOpportunity(subfolderName)
    LogInfo MODULE_NAME, "[OpportunityCreated] Emitiendo OpportunityDetected: " & subfolderName
    RaiseEvent OpportunityDetected(op)
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mFSMonitoringCoord_OpportunityCreated]"
End Sub

'@Description: Detecta eliminacion de subcarpeta en ruta de oportunidades.
'              Extrae codigo y emite OpportunityRemoved.
Private Sub mFSMonitoringCoord_OpportunityDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[OpportunityDeleted] Carpeta eliminada: " & subfolderName
    RaiseEvent OpportunityRemoved(ExtractOpportunityCode(subfolderName))
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mFSMonitoringCoord_OpportunityDeleted]"
End Sub

'@Description: Detecta renombrado de subcarpeta en ruta de oportunidades.
'              Emite OpportunityRemoved(oldCode) y, si el nuevo nombre es valido,
'              OpportunityDetected(newOp).
Private Sub mFSMonitoringCoord_OpportunityRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[OpportunityRenamed] " & oldName & " -> " & newName

    ' Emitir remocion del nombre antiguo
    RaiseEvent OpportunityRemoved(ExtractOpportunityCode(oldName))

    ' Solo emitir deteccion del nuevo nombre si cumple el patron
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    If IsValidOpportunityName(newName, regEx) Then
        RaiseEvent OpportunityDetected(BuildSingleOpportunity(newName))
    Else
        LogDebug MODULE_NAME, "[OpportunityRenamed] Nuevo nombre no cumple patron: " & newName
    End If
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mFSMonitoringCoord_OpportunityRenamed]"
End Sub

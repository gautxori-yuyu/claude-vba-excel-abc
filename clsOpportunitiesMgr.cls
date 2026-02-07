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
' Gestiona la lista de "Oportunidades" (subcarpetas) de un
' directorio base configurado en el sistema. Expone métodos
' para refrescar, enumerar y cambiar oportunidad actual, así
' como un evento para notificar cambio de oportunidad actual.
'==============================================================
'TODO Implementar filtrado de oportunidades por año
'TODO Anadir búsqueda incremental en dropdown
'TODO Cachear resultado de actualizarColeccionOportunidades (5 min)

'@Exposed
'@Folder "6-DOMINIO-Oportunidades y compresores"
'@IgnoreModule IIfSideEffect
Option Explicit

Private Const MODULE_NAME As String = "clsOpportunitiesMgr"

'--------------------------------------------------------------
' @Description: Evento que se dispara cuando cambia
' la oportunidad actual (el usuario en el desplegable
' del Ribbon, o por programa).
'--------------------------------------------------------------
Public Event currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
'--------------------------------------------------------------
' @Description: Evento que se dispara cuando se detectan cambios
' en la carpeta de oportunidades
'--------------------------------------------------------------
Public Event OpportunityCollectionUpdate(ByVal cambios As String)

'--------------------------------------------------------------
' Variables miembro
'--------------------------------------------------------------
Private mConfiguration As clsConfiguration
Private mFileManager As clsFileManager            ' Gestor de sistema de archivos (inyectado)
Private mOpportunityState As clsOpportunityState ' Estado de oportunidad (composición)
Private p_ColOpportunities As Collection         ' Lista de subcarpetas encontradas
Private p_CurrOpportunity As Long                 ' Índice en p_ColOpportunities de la carpeta actual

'--------------------------------------------------------------
' @Description: Inicializa la clase y carga la ruta base de
' oportunidades desde el registro del sistema o su valor por
' defecto.
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"

    ' Crear estado de oportunidad (composición - Manager posee el estado)
    Set mOpportunityState = New clsOpportunityState

    Set p_ColOpportunities = New Collection
    p_CurrOpportunity = -1
    
    ' cargar configuracion y rutas
    Set mConfiguration = New clsConfiguration
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"

    ' Liberar estado
    Set mOpportunityState = Nothing
End Sub

'@Description: Inicializa datos que requieren dependencias inyectadas (FileManager)
Public Sub Initialize()
    actualizarColeccionOportunidades
End Sub

'@Description: Expone el estado de oportunidad (para inyección en ApplicationState)
Public Property Get State() As clsOpportunityState
    Set State = mOpportunityState
End Property

Public Property Get Conf() As clsConfiguration
    Set Conf = mConfiguration
End Property

'@Description: Inyecta la dependencia de FileManager
'@Param fm: Instancia de clsFileManager
'@Category: Inicializacion
Friend Property Set FileManager(fm As clsFileManager)
    Set mFileManager = fm
End Property

Public Property Get BaseFolder() As String
    BaseFolder = mConfiguration.RutaOportunidades
End Property

'--------------------------------------------------------------
' @Description: Actualiza la colección p_ColOpportunities con
' las subcarpetas existentes en la ruta base configurada.
' El listado resultante se ordena en orden numérico inverso
' (de mayor a menor) según el número detectado en el nombre
' de cada carpeta. Si no se detectan números, se ordena en
' orden alfabético descendente.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function actualizarColeccionOportunidades()
    Set p_ColOpportunities = New Collection      ' limpia y reinicia

    LogInfo MODULE_NAME, "[actualizarColeccionOportunidades] regenerando la variable local que contiene la lista de carpetas de oportunidades"

    Dim arr() As String, i As Long

    ' Verificar que FileManager esta disponible
    If mFileManager Is Nothing Then
        LogError MODULE_NAME, "[actualizarColeccionOportunidades] FileManager no ha sido inyectado"
        GoTo Salida
    End If

    ' Delegar obtencion de subcarpetas al FileManager
    Dim subcarpetas As Collection
    Set subcarpetas = mFileManager.ListarSubcarpetas(BaseFolder)

    If subcarpetas.Count = 0 Then
        LogWarning MODULE_NAME, "[actualizarColeccionOportunidades] No se encontraron subcarpetas en: " & BaseFolder
        GoTo Salida
    End If

    ' Filtrar carpetas validas usando expresiones regulares (logica de dominio)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    ReDim arr(subcarpetas.Count - 1)

    Dim nombreCarpeta As Variant
    For Each nombreCarpeta In subcarpetas
        regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
        If Not regEx.Test(nombreCarpeta) Then
            regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN
            If Not regEx.Test(nombreCarpeta) Then
                ' Nombre de carpeta mal implementado. Trata de corregirlo a partir de ofertas creadas
            End If
        End If
        regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
        If regEx.Test(nombreCarpeta) Then
            arr(i) = CStr(nombreCarpeta)
            i = i + 1
        End If
    Next nombreCarpeta

    ' Orden numérico inverso
    If i > 0 Then
        ReDim Preserve arr(i - 1)
        arr = OrdenarCarpetasDesc(arr)
        For i = LBound(arr) To UBound(arr)
            p_ColOpportunities.Add arr(i)
        Next i
    End If

Salida:
    ' Ajustar índice si la colección se acortó y el índice quedó fuera de rango
    If p_CurrOpportunity >= p_ColOpportunities.Count Then
        p_CurrOpportunity = IIf(p_ColOpportunities.Count > 0, p_ColOpportunities.Count - 1, -1)
    ElseIf p_CurrOpportunity < 0 And p_ColOpportunities.Count > 0 Then
        p_CurrOpportunity = 0
    End If
    RaiseEvent OpportunityCollectionUpdate("refresh")
    actualizarColeccionOportunidades = arr
End Function

'--------------------------------------------------------------
' @Description: Procesa cambios detectados por el watcher EN CARPETAS DE OPORTUNIDAD:
'  Verifica que esos cambios realmente corresponden a cambios en oportunidades y lanza los procesos consiguientes
' @ArgumentDescriptions: cambios: nombres de carpetas separados por |
'--------------------------------------------------------------
Public Sub ProcesarCambiosEnOportunidades(ByVal subfolderName As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    Dim bCambiosEnOportunidades As Boolean
    
    LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] identificando si los cambios introducidos por el watcher son oportunidades"
    ' Verificar si los cambios afectan a oportunidades válidas
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    bCambiosEnOportunidades = regEx.Test(subfolderName)
    
    ' Si el cambio corresponde a una oportunidad, actualizar
    If bCambiosEnOportunidades Then
        LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] cambios identificados"
        Call actualizarColeccionOportunidades
    End If
End Sub

'TODO: PTE de implementar: separar del procesado de 'CARPETAS / OPORTUNIDADES'
' NO DEBE comunicarse con el File Manager... de eso se encarga clsApplication
Public Sub ProcesarCambiosEnItemsOportunidad(ByVal cambios As String)

End Sub

'--------------------------------------------------------------
' @Description: Devuelve el número de oportunidades cargadas.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function numOpportunities() As Variant
    numOpportunities = p_ColOpportunities.Count
End Function

'--------------------------------------------------------------
' @Description: Devuelve una oportunidad
' según el índice indicado.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Get OportunityLabel(Index As Integer) As String
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        OportunityLabel = p_ColOpportunities(Index + 1)
    Else
        OportunityLabel = "(Sin datos)"
    End If
End Property

'--------------------------------------------------------------
' @Description: Devuelve la ruta completa de una oportunidad
' según el índice indicado.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Get OportunityPath(Index As Long) As String
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        OportunityPath = BaseFolder & "\" & p_ColOpportunities(Index + 1)
    End If
End Property

'--------------------------------------------------------------
' @Description: Actualiza la oportunidad actual y dispara
' el evento currOpportunityChanged.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Let CurrOpportunity(Index As Long)
    If Index = p_CurrOpportunity Then Exit Property
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        p_CurrOpportunity = Index

        ' Crear objeto clsOpportunity desde índice
        Dim op As clsOpportunity
        Set op = CreateOpportunityFromIndex(Index)

        ' Actualizar estado (esto genera evento para mediadores)
        Set mOpportunityState.CurrentOpportunity = op
        mOpportunityState.CurrentIndex = Index

        LogInfo MODULE_NAME, "[CurrOpportunity] cambio de oportunidad actual -> promoviendo evento OpportunityChanged para poner al corriente a otros módulos"
        RaiseEvent currOpportunityChanged(Index, p_ColOpportunities(Index + 1))
    End If
End Property

Public Property Get CurrOpportunity() As Long
    CurrOpportunity = p_CurrOpportunity
End Property

'--------------------------------------------------------------
' @Description: Ordena un array de rutas en orden numérico
' inverso, detectando el último número en el nombre de cada
' carpeta. Si no hay número, orden alfabético descendente.
'--------------------------------------------------------------
Private Function OrdenarCarpetasDesc(arr() As String) As String()
    Dim i As Long, j As Long, tmp As String
    Dim keyI As Double, keyJ As Double
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            keyI = ExtraerClaveNumerica(arr(i))
            keyJ = ExtraerClaveNumerica(arr(j))
            If keyI < keyJ Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            ElseIf keyI = keyJ Then
                If StrComp(arr(i), arr(j), vbTextCompare) < 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            End If
        Next j
    Next i
    OrdenarCarpetasDesc = arr
End Function

'--------------------------------------------------------------
' @Description: Extrae el último número del nombre de la carpeta.
' Si no hay número, devuelve -1E+99 para forzar orden al final.
'--------------------------------------------------------------
Private Function ExtraerClaveNumerica(ruta As String) As Double
    Dim re As Object, matches As Object, s As String
    Dim nombre As String
    On Error GoTo ErrHandler
    nombre = Dir(ruta)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    re.Global = True
    If re.Test(nombre) Then
        Set matches = re.Execute(nombre)
        s = matches(matches.Count - 1).Value
        ExtraerClaveNumerica = CDbl(s)
        Exit Function
    End If
ErrHandler:
    ExtraerClaveNumerica = -1E+99
End Function

'--------------------------------------------------------------
' @Description: Crea un objeto clsOpportunity desde un índice de la colección
'--------------------------------------------------------------
Private Function CreateOpportunityFromIndex(Index As Long) As clsOpportunity
    Dim op As clsOpportunity
    Set op = New clsOpportunity

    If Index >= 0 And Index < p_ColOpportunities.Count Then
        ' Extraer información del nombre de la carpeta
        Dim folderName As String
        folderName = p_ColOpportunities(Index + 1)

        ' Asignar propiedades básicas
        op.Label = folderName
        op.Path = BaseFolder & "\" & folderName

        ' Intentar extraer información adicional del nombre
        ' Formato típico: "412601015 - Cliente - Proyecto"
        On Error Resume Next
        Dim parts() As String
        parts = Split(folderName, " - ")
        If UBound(parts) >= 0 Then
            op.Name = Trim(parts(0))  ' "412601015"
        End If
        If UBound(parts) >= 1 Then
            op.Customer = Trim(parts(1))  ' "Cliente"
        End If
        On Error GoTo 0
    End If

    Set CreateOpportunityFromIndex = op
End Function

Private Function getNextOpportunityCode()
    Dim arrOps() As String, i, dblLastOpSeq As Double, strOpName As String
    arrOps = actualizarColeccionOportunidades
    dblLastOpSeq = 0
    For i = 0 To UBound(arrOps)
        arrOps(i) = Left(arrOps(i), 9)
    Next
    If UBound(arrOps) >= 0 Then dblLastOpSeq = CDbl(Mid(arrOps(0), 7))
    Do
        dblLastOpSeq = dblLastOpSeq + 1
        strOpName = mConfiguration.SAM & Mid(Year(Now), 3) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(3 - Len(CStr(dblLastOpSeq)), "0") & dblLastOpSeq
    Loop While UBound(Filter(arrOps, strOpName)) >= 0
    getNextOpportunityCode = strOpName
End Function
Private Function getCustomerNameInputBox(strOpName As String)
    Dim strCustomer As String
    
    Dim regEx As Object, strMsg
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    Do
        strCustomer = InputBox(strMsg & "Introduce el nombre del cliente, o cliente - proyecto:", "Cliente - Proyecto", strCustomer)
        strMsg = "NOMBRE NO VALIDO. "
    Loop While strCustomer <> "" And Not regEx.Test(strOpName & " - " & strCustomer & " - XXX")
    
    getCustomerNameInputBox = strCustomer
End Function
'--------------------------------------------------------------
' @Description: Crea una nueva carpeta de oportunidad, a partir de la estructura de plantilla
'--------------------------------------------------------------
Public Sub CreaOportunidad()
    On Error GoTo ErrHandler


    ' Verificar que FileManager esta disponible
    If mFileManager Is Nothing Then
        LogError MODULE_NAME, "[CreaOportunidad] FileManager no ha sido inyectado"
        MsgBox "Error interno: FileManager no disponible", vbCritical
        Exit Sub
    End If

    ' Verificar que la carpeta base existe
    If Not mFileManager.ExisteCarpeta(BaseFolder) Then
        MsgBox "No se pudo abrir la ruta de oportunidades: " & BaseFolder, vbExclamation
        Exit Sub
    End If

    Dim strOpName As String, strCustomer As String
    
    strOpName = getNextOpportunityCode
    
    strCustomer = getCustomerNameInputBox(strOpName)
    
    If strCustomer <> "" Then
        ' Construir rutas usando FileManager
        Dim rutaOrigen As String
        Dim rutaDestino As String
        Dim nombreCarpetaDestino As String

        rutaOrigen = App.OpportunitiesMgr.Conf.RutaPlantillas
        ' Eliminar barra final si existe
        If Right(rutaOrigen, 1) = "\" Then rutaOrigen = Left(rutaOrigen, Len(rutaOrigen) - 1)

        nombreCarpetaDestino = strOpName & " - " & strCustomer & " - XXX"
        rutaDestino = mFileManager.ConstruirRuta(BaseFolder, nombreCarpetaDestino)

        ' Delegar copia de carpeta al FileManager
        If mFileManager.CopiarCarpeta(rutaOrigen, rutaDestino) Then
            LogInfo MODULE_NAME, "[CreaOportunidad] Nueva oportunidad creada: " & nombreCarpetaDestino
        Else
            LogError MODULE_NAME, "[CreaOportunidad] Fallo al copiar carpeta de plantilla"
            MsgBox "Error al crear la oportunidad. Revise el log para mas detalles.", vbExclamation
        End If

        ' Actualizar coleccion de oportunidades
        actualizarColeccionOportunidades
    End If
    
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[CreaOportunidad]"
    MsgBox "No se pudo crear la oportunidad: " & Err.Description, vbExclamation
End Sub

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
' directorio base configurado en el sistema. Expone m?todos
' para refrescar, enumerar y cambiar oportunidad actual, as?
' como un evento para notificar cambio de oportunidad actual.
'==============================================================
'TODO Implementar filtrado de oportunidades por a?o
'TODO Anadir b?squeda incremental en dropdown
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
Private p_ColOpportunities As Collection         ' Lista de subcarpetas encontradas
Private p_CurrOpportunity As Long                 ' ?ndice en p_ColOpportunities de la carpeta actual

Private WithEvents mXlApp As Application
Attribute mXlApp.VB_VarHelpID = -1

'--------------------------------------------------------------
' @Description: Inicializa la clase y carga la ruta base de
' oportunidades desde el registro del sistema o su valor por
' defecto.
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    
    Set p_ColOpportunities = New Collection
    p_CurrOpportunity = -1
    
    ' cargar configuracion y rutas
    Set mConfiguration = New clsConfiguration
    actualizarColeccionOportunidades
    
    ' gestion de eventos
    Set mXlApp = Application
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

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
' @Description: Actualiza la colecci?n p_ColOpportunities con
' las subcarpetas existentes en la ruta base configurada.
' El listado resultante se ordena en orden num?rico inverso
' (de mayor a menor) seg?n el n?mero detectado en el nombre
' de cada carpeta. Si no se detectan n?meros, se ordena en
' orden alfab?tico descendente.
'--------------------------------------------------------------
' @Category: Informaci?n de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function actualizarColeccionOportunidades()
    Set p_ColOpportunities = New Collection      ' limpia y reinicia

    LogInfo MODULE_NAME, "[actualizarColeccionOportunidades] - regenerando la variable local que contiene la lista de carpetas de oportunidades"

    Dim arr() As String, i As Long

    ' Verificar que FileManager esta disponible
    If mFileManager Is Nothing Then
        LogError MODULE_NAME, "[actualizarColeccionOportunidades] - FileManager no ha sido inyectado"
        GoTo Salida
    End If

    ' Delegar obtencion de subcarpetas al FileManager
    Dim subcarpetas As Collection
    Set subcarpetas = mFileManager.ListarSubcarpetas(BaseFolder)

    If subcarpetas.Count = 0 Then
        LogWarning MODULE_NAME, "[actualizarColeccionOportunidades] - No se encontraron subcarpetas en: " & BaseFolder
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

    ' Orden numerico inverso (logica de dominio)
    If i > 0 Then
        ReDim Preserve arr(i - 1)
        arr = OrdenarCarpetasDesc(arr)
        For i = LBound(arr) To UBound(arr)
            p_ColOpportunities.Add arr(i)
        Next i
    End If

Salida:
    p_CurrOpportunity = IIf(p_CurrOpportunity > 0, p_CurrOpportunity, IIf(p_ColOpportunities.Count > 0, 0, -1))
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
    
    LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] - identificando si los cambios introducidos por el watcher son oportunidades"
    ' Verificar si los cambios afectan a oportunidades v?lidas
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    bCambiosEnOportunidades = regEx.Test(subfolderName)
    
    ' Si el cambio corresponde a una oportunidad, actualizar
    If bCambiosEnOportunidades Then
        LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] - cambios identificados"
        Call actualizarColeccionOportunidades
        ' NO actualizar ribbon directamente - delegar al evento
        ' La clase NO debe conocer el ribbon
        ' Disparar evento para que otros componentes reaccionen
        LogInfo MODULE_NAME, "[ProcesarCambiosEnOportunidades] - promoviendo evento OpportunityCollectionUpdate para poner al corriente a otros m?dulos"
        RaiseEvent OpportunityCollectionUpdate(subfolderName)
    End If
End Sub

'TODO: PTE de implementar: separar del procesado de 'CARPETAS / OPORTUNIDADES'
' NO DEBE comunicarse con el File Manager... de eso se encarga clsApplication
Public Sub ProcesarCambiosEnItemsOportunidad(ByVal cambios As String)

End Sub

'--------------------------------------------------------------
' @Description: Devuelve el n?mero de oportunidades cargadas.
'--------------------------------------------------------------
' @Category: Informaci?n de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function numOpportunities() As Variant
    numOpportunities = p_ColOpportunities.Count
End Function

'--------------------------------------------------------------
' @Description: Devuelve una oportunidad
' seg?n el ?ndice indicado.
'--------------------------------------------------------------
' @Category: Informaci?n de archivo
' @ArgumentDescriptions: index: ?ndice (base 0) de la oportunidad
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
' seg?n el ?ndice indicado.
'--------------------------------------------------------------
' @Category: Informaci?n de archivo
' @ArgumentDescriptions: index: ?ndice (base 0) de la oportunidad
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
' @Category: Informaci?n de archivo
' @ArgumentDescriptions: index: ?ndice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Let CurrOpportunity(Index As Long)
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        p_CurrOpportunity = Index
        
        LogInfo MODULE_NAME, "[CurrOpportunity] - cambio de oportunidad actual -> promoviendo evento OpportunityChanged para poner al corriente a otros m?dulos"
        RaiseEvent currOpportunityChanged(Index, p_ColOpportunities(Index + 1))
    End If
End Property

Public Property Get CurrOpportunity() As Long
    CurrOpportunity = p_CurrOpportunity
End Property

'--------------------------------------------------------------
' @Description: Ordena un array de rutas en orden num?rico
' inverso, detectando el ?ltimo n?mero en el nombre de cada
' carpeta. Si no hay n?mero, orden alfab?tico descendente.
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
' @Description: Extrae el ?ltimo n?mero del nombre de la carpeta.
' Si no hay n?mero, devuelve -1E+99 para forzar orden al final.
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
' @Description: Crea una nueva carpeta de oportunidad, a partir de la estructura de plantilla
'--------------------------------------------------------------
'@Description: Crea una nueva carpeta de oportunidad a partir de la estructura de plantilla
'@Note: Solicita el nombre del cliente mediante InputBox y valida el formato
'@Category: Oportunidades
Public Sub CreaOportunidad()
    On Error GoTo ErrHandler

    Dim arrOps() As String, i, strSAM As String, strCustomer As String, strOpName As String, dblLastOpSeq As Double

    ' Verificar que FileManager esta disponible
    If mFileManager Is Nothing Then
        LogError MODULE_NAME, "[CreaOportunidad] - FileManager no ha sido inyectado"
        MsgBox "Error interno: FileManager no disponible", vbCritical
        Exit Sub
    End If

    ' Verificar que la carpeta base existe
    If Not mFileManager.ExisteCarpeta(BaseFolder) Then
        MsgBox "No se pudo abrir la ruta de oportunidades: " & BaseFolder, vbExclamation
        Exit Sub
    End If

    ' Actualizar lista de oportunidades y calcular siguiente secuencia
    arrOps = actualizarColeccionOportunidades
    dblLastOpSeq = 0
    For i = 0 To UBound(arrOps)
        arrOps(i) = Left(arrOps(i), 9)
    Next
    If UBound(arrOps) >= 0 Then dblLastOpSeq = CDbl(Mid(arrOps(0), 7))

    ' Generar nombre de oportunidad (logica de dominio)
    strSAM = App.OpportunitiesMgr.Conf.SAM
    Do
        dblLastOpSeq = dblLastOpSeq + 1
        strOpName = strSAM & Mid(Year(Now), 3) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(3 - Len(CStr(dblLastOpSeq)), "0") & dblLastOpSeq
    Loop While UBound(Filter(arrOps, strOpName)) >= 0

    ' Solicitar y validar nombre de cliente
    Dim regEx As Object, strMsg
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    Do
        strCustomer = InputBox(strMsg & "Introduce el nombre del cliente, o cliente - proyecto:", "Cliente - Proyecto", strCustomer)
        strMsg = "NOMBRE NO VALIDO. "
    Loop While strCustomer <> "" And Not regEx.Test(strOpName & " - " & strCustomer & " - XXX")

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
            LogInfo MODULE_NAME, "[CreaOportunidad] - Nueva oportunidad creada: " & nombreCarpetaDestino
        Else
            LogError MODULE_NAME, "[CreaOportunidad] - Fallo al copiar carpeta de plantilla"
            MsgBox "Error al crear la oportunidad. Revise el log para mas detalles.", vbExclamation
        End If
    End If

    ' Actualizar coleccion de oportunidades
    arrOps = actualizarColeccionOportunidades
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[CreaOportunidad]"
    MsgBox "No se pudo crear la oportunidad: " & Err.Description, vbExclamation
End Sub

' ==================================================================
' IMPLEMENTACION DE EVENTOS: CALLBACKS
' ==================================================================

Private Sub mXlApp_WorkbookActivate(ByVal Wb As Workbook)
    ' Verificar si el workbook est? en una oportunidad
    ' Actualizar p_CurrOpportunity si corresponde
End Sub


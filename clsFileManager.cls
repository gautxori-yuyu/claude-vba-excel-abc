VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' GESTOR GENÉRICO DE ARCHIVOS
' ==================================================================
' Responsabilidades:
' - Supervisar CUALQUIER tipo de archivo (Excel, PDF, Word, etc.)
'   relacionado con la gestión que hace "la aplicación": supervisa
'   por tanto tanto ficheros abiertos por Excel, como los ficheros
'   que pudieran estar En ciertas carpetas o árbol de carpetas
' - Mantener un índice de archivos supervisados, por su ObjectKey
' - Mantener "sincronizado" el archivo de Excel activo
' - Proveer análisis de archivos sin duplicar lógica
' ==================================================================

'@IgnoreModule MissingAnnotationArgument
'@Folder "2-Infraestructura"
Option Explicit

Private Const MODULE_NAME As String = "clsFileManager"

' ==========================================
' EVENTOS
' ==========================================

Public Event ActiveFileChanged(f As clsFileXLS)

Private mTrackedFiles As Object                 ' Diccionario de TODOS los ficheros de excel, en TODAS las oportunidades procesadas
Private mFileState As clsFileState              ' Estado de archivo activo (composición)

' ==========================================
' SUSCRIPCION A EVENTOS DE APPLICATION
' ==========================================
Private WithEvents mXlApp As Application
Attribute mXlApp.VB_VarHelpID = -1

' ==================================================================
' INICIALIZACIÓN Y LIMPIEZA
' ==================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"

    ' Crear estado de archivo (composición - Manager posee el estado)
    Set mFileState = New clsFileState

    Set mTrackedFiles = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"

    ' Liberar estado
    Set mFileState = Nothing

    ' Limpiar todas las referencias
    mTrackedFiles.RemoveAll
    Set mTrackedFiles = Nothing
End Sub

'@Description: Expone el estado de archivo (para inyección en ApplicationState)
Public Property Get State() As clsFileState
    Set State = mFileState
End Property

'@Description: Establece el archivo Excel activo y notifica el cambio
Public Sub SetActiveFile(f As clsFileXLS)
    ' Almacenar archivo en estado (memoria pura)
    Set mFileState.CurrentFile = f

    ' Calcular y almacenar IsOpportunityFile (decisión del Manager)
    If f Is Nothing Then
        mFileState.isOpportunityFile = False
    Else
        Dim isOpp As Boolean
        isOpp = False
        On Error Resume Next
        isOpp = f.EsOportunidad
        On Error GoTo 0
        mFileState.isOpportunityFile = isOpp
    End If

    LogDebug MODULE_NAME, "[SetActiveFile] " & mFileState.GetCurrentFileName() & _
                          ", IsOpportunity=" & mFileState.isOpportunityFile

    ' Notificar cambio (evento del Manager)
    RaiseEvent ActiveFileChanged(f)
End Sub

' -------------------------------------------------------------
' Init: GESTIONA LA COLECCIÓN DE WORKBOOKS
' (Es responsabilidad de clsApplication, NO de clsFileManager)
' -------------------------------------------------------------
Public Sub Init(oCurrWbks As Workbooks)
    LogInfo MODULE_NAME, "[Init]"
    
    On Error GoTo ErrHandler
    
    ' Trackear todos los workbooks abiertos
    Dim oWorkBook As Workbook
    For Each oWorkBook In oCurrWbks
        GetOrTrackWorkbook oWorkBook
    Next oWorkBook

    LogInfo MODULE_NAME, "[Init] Workbooks tracked: " & mTrackedFiles.Count
    
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[Init] Error inicializando tracking de workbooks", , Err.Description
    Err.Raise Err.Number, MODULE_NAME & "[Init]", _
              "Error inicializando tracking de workbooks: " & Err.Description
End Sub

' ==================================================================
' PROPIEDADES PÚBLICAS
' ==================================================================

'@Description: Obtiene el archivo Excel activo
'@Note: Devuelve clsFileXLS o Nothing si no hay archivo activo o no es Excel
'Public Property Get ActiveExcelFile() As clsFileXLS
'    Set ActiveExcelFile = GetOrTrackWorkbook(mXlApp.ActiveWorkbook)
'End Property

' ==================================================================
' MÉTODOS PÚBLICOS - TRACKING
' ==================================================================

'@Description: Obtiene o crea el tracker para un Workbook
'@Returns: clsFileXLS asociado al Workbook
Public Function GetOrTrackWorkbook(wb As Workbook) As clsFileXLS
    If wb Is Nothing Then Exit Function
    
    On Error GoTo ErrHandler
    
    ' Crear instancia temporal para obtener su clave
    Dim f As clsFileXLS
    Set f = New clsFileXLS
    f.BindTo wb
    
    Dim key As Double
    key = GetFileKey(f)
    
    ' Validar que la clave es válida
    If key = 0 Then
        LogError MODULE_NAME, "[GetOrTrackWorkbook] - ObjectKey inválido para " & wb.Name
        Exit Function
    End If
    
    ' Si ya existe, devolver la instancia existente (descartar temporal)
    If mTrackedFiles.Exists(key) Then
        Set GetOrTrackWorkbook = mTrackedFiles(key)
        Exit Function
    End If
    
    ' Agregar al diccionario
    mTrackedFiles.Add key, f
    
    LogDebug MODULE_NAME, "[GetOrTrackWorkbook] Nuevo archivo tracked: " & f.Path & " (total: " & mTrackedFiles.Count & ")"
    
    Set GetOrTrackWorkbook = f
    Exit Function
    
ErrHandler:
    LogCurrentError MODULE_NAME, "[GetOrTrackWorkbook]"
End Function

'@Description: Remueve un Workbook del tracking (interfaz específica Excel)
Public Sub UntrackWorkbook(wb As Workbook)
    If wb Is Nothing Then Exit Sub
    
    ' Buscar el archivo tracked correspondiente
    Dim f As clsFileXLS
    Set f = FindTrackedWorkbook(wb)
    
    If Not f Is Nothing Then
        UntrackFile f
    End If
End Sub

' ==================================================================
' MÉTODOS PÚBLICOS - TRACKING GENÉRICO
' ==================================================================

'@Description: Agrega un archivo al tracking (genérico)
'@Note: Acepta clsFileXLS, clsFilePDF, o cualquier clase con ObjectKey
Public Sub TrackFile(f As Object)
    If f Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    
    Dim key As Double
    key = GetFileKey(f)
    
    If mTrackedFiles.Exists(key) Then
        LogWarning MODULE_NAME, "[TrackFile] Advertencia: Archivo ya está tracked"
        Exit Sub
    End If
    
    mTrackedFiles.Add key, f
    LogDebug MODULE_NAME, "[TrackFile] Archivo tracked: " & GetFilePath(f) & " (total: " & mTrackedFiles.Count & ")"
    Exit Sub
    
ErrHandler:
    LogCurrentError MODULE_NAME, "[TrackFile]"
End Sub

'@Description: Remueve un archivo del tracking (genérico)
'@Note: Usa ObjPtr de la clase de archivo, no necesita conocer el tipo específico
Public Sub UntrackFile(f As Object)
    If f Is Nothing Then Exit Sub
    
    Dim key As Variant
    key = GetFileKey(f)
    
    If IsEmpty(key) Then Exit Sub
    
    If Not mTrackedFiles.Exists(key) Then
        LogWarning MODULE_NAME, "[UntrackFile] Advertencia: Intento de untrack de archivo no tracked"
        Exit Sub
    End If
    
    ' Remover del diccionario
    mTrackedFiles.Remove key
    
    LogInfo MODULE_NAME, "[UntrackFile] Archivo untracked (quedan: " & mTrackedFiles.Count & ")"
End Sub

' ==================================================================
' MÉTODOS PÚBLICOS - ANÁLISIS
' ==================================================================

'@Description: Analiza un workbook y devuelve su información
'@Returns: T_InfoArchivo con el análisis completo
Public Function AnalizarArchivo(fich As Object) As mod_ConstantsGlobals.T_InfoArchivo
    Dim info As mod_ConstantsGlobals.T_InfoArchivo
    Select Case TypeName(fich)
    Case "Nothing", "Empty"
        info.EsValido = False
        AnalizarArchivo = info
    Case "Workbook"
        Dim f As clsFileXLS, owb As Workbook
        Set owb = fich
        Set f = GetOrTrackWorkbook(owb)
        AnalizarArchivo = f.info
    Case "FileSystemObject"
        ' HABRIA QUE PROCESAR DOCUMENTOS DE WORD, DE PDF, etc.
    Case Else
        ' ... (puedo pasar enteros, strings, ...)
    End Select
End Function

'@Description: Analiza el workbook activo actual
'@Returns: T_InfoArchivo del workbook activo, o estructura vacía si no hay activo
'Public Function AnalizarArchivoActivo() As mod_ConstantsGlobals.T_InfoArchivo
'    AnalizarArchivoActivo = AnalizarArchivo(mXlApp.ActiveWorkbook)
'End Function

' ==================================================================
' MÉTODOS DE UTILIDAD
' ==================================================================

'@Description: Verifica si un workbook está siendo tracked
Public Function IsTracked(wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    IsTracked = Not (FindTrackedWorkbook(wb) Is Nothing)
End Function

'@Description: Busca el clsFileXLS correspondiente a un Workbook
'@Returns: clsFileXLS si está tracked, Nothing si no
Private Function FindTrackedWorkbook(wb As Workbook) As clsFileXLS
    If wb Is Nothing Then Exit Function
    
    ' Crear instancia temporal solo para obtener la clave
    Dim fTemp As clsFileXLS
    Set fTemp = New clsFileXLS
    Dim key As Double
    key = fTemp.CreateFromWorkbook(wb).ObjectKey
    
    If mTrackedFiles.Exists(key) Then
        Set FindTrackedWorkbook = mTrackedFiles(key)
    End If
End Function

'@Description: Obtiene la clave de un archivo que implementa IFile
'@Returns: ObjectKey o 0 si no tiene
'@Note: Solo acepta objetos que implementen IFile
Private Function GetFileKey(f As Object) As Double
    On Error GoTo ErrHandler

    ' Validar que el objeto implementa IFile
    Dim IFile As IFile
    Set IFile = f  ' Esto falla si no implementa IFile

    ' Usar ObjectKey de la interfaz
    GetFileKey = IFile.ObjectKey

    ' Validar que sea válido
    If GetFileKey = 0 Then
        LogWarning MODULE_NAME, "[GetFileKey] ObjectKey = 0 para " & TypeName(f)
    End If
    
    Exit Function
    
ErrHandler:
    LogError MODULE_NAME, "[GetFileKey] Error para tipo " & TypeName(f) & ". El objeto no implementa IFile correctamente"
    GetFileKey = 0
End Function

'@Description: Obtiene el path de un archivo que implementa IFile
'@Returns: Path o "(desconocido)" si no tiene
Private Function GetFilePath(f As Object) As String
    On Error GoTo ErrHandler
    
    Dim IFile As IFile
    Set IFile = f
    
    GetFilePath = IFile.Path
    Exit Function
    
ErrHandler:
    LogError MODULE_NAME, "[GetFilePath] Error para tipo " & TypeName(f)
    GetFilePath = "(desconocido)"
End Function

'@Description: Obtiene información de todos los archivos tracked (para debugging)
Public Function GetTrackedFilesInfo() As String
    Dim info As String
    info = "Archivos tracked: " & mTrackedFiles.Count & vbCrLf
    
    Dim f As Object
    Dim key As Variant
    
    For Each key In mTrackedFiles.Keys
        Set f = mTrackedFiles(key)
        info = info & "  - [" & TypeName(f) & "] " & GetFilePath(f) & vbCrLf
    Next
    
'    If Not ActiveExcelFile Is Nothing Then
'        info = info & vbCrLf & "Archivo activo: " & GetFilePath(ActiveExcelFile)
'    Else
'        info = info & vbCrLf & "Archivo activo: (ninguno)"
'    End If
    
    GetTrackedFilesInfo = info
End Function

' ==================================================================
' METODOS PUBLICOS - SISTEMA DE ARCHIVOS
' ==================================================================

'@Description: Lista las subcarpetas de una ruta base
'@Param rutaBase: Ruta del directorio a explorar
'@Returns: Collection de nombres de subcarpetas (solo nombres, no rutas completas)
'@Note: Devuelve Collection vacia si la ruta no existe o no es accesible
'@Category: Sistema de archivos
Public Function ListarSubcarpetas(rutaBase As String) As Collection
    Dim resultado As Collection
    Set resultado = New Collection

    On Error GoTo ErrHandler

    Dim fso As Object
    Dim carpeta As Object
    Dim subcarpeta As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rutaBase) Then
        LogWarning MODULE_NAME, "[ListarSubcarpetas] - Carpeta no existe: " & rutaBase
        Set ListarSubcarpetas = resultado
        Exit Function
    End If

    Set carpeta = fso.GetFolder(rutaBase)

    For Each subcarpeta In carpeta.SubFolders
        resultado.Add subcarpeta.Name
    Next subcarpeta

    LogDebug MODULE_NAME, "[ListarSubcarpetas] - Encontradas " & resultado.Count & " subcarpetas en: " & rutaBase

    Set ListarSubcarpetas = resultado
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[ListarSubcarpetas] Error listando carpetas", Err.Number, Err.Description
    Set ListarSubcarpetas = resultado
End Function

'@Description: Copia una carpeta completa de origen a destino
'@Param rutaOrigen: Ruta de la carpeta a copiar
'@Param rutaDestino: Ruta destino donde se copiara la carpeta
'@Returns: True si la copia fue exitosa, False en caso contrario
'@Category: Sistema de archivos
Public Function CopiarCarpeta(rutaOrigen As String, rutaDestino As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rutaOrigen) Then
        LogError MODULE_NAME, "[CopiarCarpeta] - Carpeta origen no existe: " & rutaOrigen
        CopiarCarpeta = False
        Exit Function
    End If

    If fso.FolderExists(rutaDestino) Then
        LogWarning MODULE_NAME, "[CopiarCarpeta] - Carpeta destino ya existe: " & rutaDestino
        CopiarCarpeta = False
        Exit Function
    End If

    fso.CopyFolder rutaOrigen, rutaDestino

    LogInfo MODULE_NAME, "[CopiarCarpeta] - Carpeta copiada exitosamente: " & rutaOrigen & " -> " & rutaDestino
    CopiarCarpeta = True
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[CopiarCarpeta] Error copiando carpeta", Err.Number, Err.Description
    CopiarCarpeta = False
End Function

'@Description: Verifica si una carpeta existe
'@Param ruta: Ruta de la carpeta a verificar
'@Returns: True si la carpeta existe, False en caso contrario
'@Category: Sistema de archivos
Public Function ExisteCarpeta(ruta As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ExisteCarpeta = fso.FolderExists(ruta)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[ExisteCarpeta] Error verificando existencia", Err.Number, Err.Description
    ExisteCarpeta = False
End Function

'@Description: Construye una ruta completa a partir de componentes
'@Param componentes: Array de componentes de ruta a combinar
'@Returns: String con la ruta combinada correctamente
'@Category: Sistema de archivos
'@Example: ConstruirRuta("C:\Base", "Subfolder", "File.txt") -> "C:\Base\Subfolder\File.txt"
Public Function ConstruirRuta(ParamArray componentes() As Variant) As String
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim rutaFinal As String
    Dim i As Long

    If UBound(componentes) < 0 Then
        ConstruirRuta = ""
        Exit Function
    End If

    rutaFinal = CStr(componentes(0))

    For i = 1 To UBound(componentes)
        rutaFinal = fso.BuildPath(rutaFinal, CStr(componentes(i)))
    Next i

    ConstruirRuta = rutaFinal
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[ConstruirRuta] Error construyendo ruta", Err.Number, Err.Description
    ConstruirRuta = ""
End Function

'@Description: Obtiene el objeto Folder de FSO para una ruta
'@Param ruta: Ruta de la carpeta
'@Returns: Objeto Folder o Nothing si no existe
'@Category: Sistema de archivos
Public Function ObtenerCarpeta(ruta As String) As Object
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(ruta) Then
        Set ObtenerCarpeta = Nothing
        Exit Function
    End If

    Set ObtenerCarpeta = fso.GetFolder(ruta)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[ObtenerCarpeta] Error obteniendo carpeta", Err.Number, Err.Description
    Set ObtenerCarpeta = Nothing
End Function


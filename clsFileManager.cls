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
' GESTOR GEN�RICO DE ARCHIVOS
' ==================================================================
' Responsabilidades:
' - Supervisar CUALQUIER tipo de archivo (Excel, PDF, Word, etc.)
'   relacionado con la gesti�n que hace "la aplicaci�n": supervisa
'   por tanto tanto ficheros abiertos por Excel, como los ficheros
'   que pudieran estar En ciertas carpetas o �rbol de carpetas
' - Mantener un �ndice de archivos supervisados, por su ObjectKey
' - Mantener "sincronizado" el archivo de Excel activo
' - Proveer an�lisis de archivos sin duplicar l�gica
' ==================================================================

'@IgnoreModule MissingAnnotationArgument
'@Folder "2-Servicios.Archivos"
Option Explicit

Private p_trackedFiles As Object                 ' Diccionario de TODOS los ficheros de excel, en TODAS las oportunidades procesadas
Private p_currExcelFile As clsExcelFile

Private WithEvents ctx As clsExecutionContext
Attribute ctx.VB_VarHelpID = -1

' ==================================================================
' INICIALIZACI�N Y LIMPIEZA
' ==================================================================

Private Sub Class_Initialize()
    LogInfo "clsFileManager", "[Class_Initialize]"
    Set p_trackedFiles = CreateObject("Scripting.Dictionary")
End Sub

'@Description: Inicializa el FileManager con el contexto de ejecución compartido
'@Note: Debe llamarse desde clsAplicacion después de crear el contexto
Public Sub Initialize(ByVal executionContext As clsExecutionContext)
    If executionContext Is Nothing Then
        LogError "clsFileManager", "[Initialize] Error: executionContext es Nothing"
        Exit Sub
    End If
    Set ctx = executionContext
    LogInfo "clsFileManager", "[Initialize] Contexto de ejecución vinculado"
End Sub

Private Sub Class_Terminate()
    LogInfo "clsFileManager", "[Class_Terminate]"
    
    ' Limpiar todas las referencias
    p_trackedFiles.RemoveAll
    Set p_trackedFiles = Nothing
    Set p_currExcelFile = Nothing
End Sub

' ==================================================================
' PROPIEDADES P�BLICAS
' ==================================================================

'@Description: Obtiene el archivo Excel activo
'@Note: Devuelve clsExcelFile o Nothing si no hay archivo activo o no es Excel
Public Property Get ActiveWb() As clsExcelFile
    If TypeName(p_currExcelFile) = "clsExcelFile" Then
        Set ActiveWb = p_currExcelFile
    End If
End Property

'@Description: Establece el archivo activo (con validaci�n)
'@Note: Acepta clsExcelFile
'TODO: revisar posible COLISION con los eventos del contexto de ejecucion ctx (en ppio, esta funcion
' NO deberia usarse, y todo cambio de  deberia hacerse mediante eventos de ctx)
Friend Property Set ActiveWb(f As clsExcelFile)
    If f Is Nothing Then
        LogWarning "clsFileManager", "[ActiveWb] Advertencia: Se intenta asignar Nothing como archivo activo"
        Set p_currExcelFile = Nothing
        Exit Property
    End If
    
    ' Obtener la clave del archivo
    Dim key As Double
    key = GetFileKey(f)
    
    If key = 0 Then
        LogError "clsFileManager", "[ActiveWb] Error: El archivo no provee ObjectKey v�lido"
        Exit Property
    End If
    
    ' Validar que el archivo est� tracked
    If Not p_trackedFiles.Exists(key) Then
        LogWarning "clsFileManager", "[ActiveWb] Advertencia: Se intenta activar un archivo no tracked"
        TrackFile f
    End If
    
    Set p_currExcelFile = f
    LogDebug "clsFileManager", "[ActiveWb] Archivo activo: " & GetFilePath(f)
End Property

'@Description: Obtiene el n�mero de archivos tracked
Public Property Get TrackedCount() As Long
    TrackedCount = p_trackedFiles.Count
End Property

' ==================================================================
' M�TODOS P�BLICOS - TRACKING
' ==================================================================

'@Description: Obtiene o crea el tracker para un Workbook
'@Returns: clsExcelFile asociado al Workbook
Public Function GetOrTrackWorkbook(wb As Workbook) As clsExcelFile
    If wb Is Nothing Then Exit Function
    
    On Error GoTo ErrHandler
    
    ' Crear instancia temporal para obtener su clave
    Dim f As clsExcelFile
    Set f = New clsExcelFile
    f.BindTo wb
    
    Dim key As Double
    key = GetFileKey(f)
    
    ' Validar que la clave es v�lida
    If key = 0 Then
        LogError "clsFileManager", "[GetOrTrackWorkbook] Error: ObjectKey inv�lido para " & wb.Name
        Exit Function
    End If
    
    ' Si ya existe, devolver la instancia existente (descartar temporal)
    If p_trackedFiles.Exists(key) Then
        Set GetOrTrackWorkbook = p_trackedFiles(key)
        Exit Function
    End If
    
    ' Agregar al diccionario
    p_trackedFiles.Add key, f
    
    LogDebug "clsFileManager", "[GetOrTrackWorkbook] Nuevo archivo tracked: " & f.Path & " (total: " & p_trackedFiles.Count & ")"
    
    Set GetOrTrackWorkbook = f
    Exit Function
    
ErrHandler:
    LogError "clsFileManager", "[GetOrTrackWorkbook] Error", Err.Number, Err.Description
End Function

'@Description: Remueve un Workbook del tracking (interfaz espec�fica Excel)
Public Sub UntrackWorkbook(wb As Workbook)
    If wb Is Nothing Then Exit Sub
    
    ' Buscar el archivo tracked correspondiente
    Dim f As clsExcelFile
    Set f = FindTrackedWorkbook(wb)
    
    If Not f Is Nothing Then
        UntrackFile f
    End If
End Sub

' ==================================================================
' M�TODOS P�BLICOS - TRACKING GEN�RICO
' ==================================================================

'@Description: Agrega un archivo al tracking (gen�rico)
'@Note: Acepta clsExcelFile, clsPDFFile, o cualquier clase con ObjectKey
Public Sub TrackFile(f As Object)
    If f Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    
    Dim key As Double
    key = GetFileKey(f)
    
    If p_trackedFiles.Exists(key) Then
        LogWarning "clsFileManager", "[TrackFile] Advertencia: Archivo ya est� tracked"
        Exit Sub
    End If
    
    p_trackedFiles.Add key, f
    LogDebug "clsFileManager", "[TrackFile] Archivo tracked: " & GetFilePath(f) & " (total: " & p_trackedFiles.Count & ")"
    Exit Sub
    
ErrHandler:
    LogError "clsFileManager", "[TrackFile] Error", Err.Number, Err.Description
End Sub

'@Description: Remueve un archivo del tracking (gen�rico)
'@Note: Usa ObjPtr de la clase de archivo, no necesita conocer el tipo espec�fico
Public Sub UntrackFile(f As Object)
    If f Is Nothing Then Exit Sub
    
    Dim key As Variant
    key = GetFileKey(f)
    
    If IsEmpty(key) Then Exit Sub
    
    If Not p_trackedFiles.Exists(key) Then
        LogWarning "clsFileManager", "[UntrackFile] Advertencia: Intento de untrack de archivo no tracked"
        Exit Sub
    End If
    
    ' Limpiar referencia al archivo activo si es necesario
    If Not p_currExcelFile Is Nothing Then
        If p_currExcelFile.ObjectKey = key Then
            Set p_currExcelFile = Nothing
            LogDebug "clsFileManager", "[UntrackFile] Archivo activo limpiado"
        End If
    End If
    
    ' Remover del diccionario
    p_trackedFiles.Remove key
    
    LogInfo "clsFileManager", "[UntrackFile] Archivo untracked (quedan: " & p_trackedFiles.Count & ")"
End Sub

' ==================================================================
' M�TODOS P�BLICOS - AN�LISIS
' ==================================================================

'@Description: Analiza un workbook y devuelve su informaci�n
'@Returns: T_InfoArchivo con el an�lisis completo
Public Function AnalizarArchivo(fich As Object) As mod_ConstantsGlobals.T_InfoArchivo
    Dim info As mod_ConstantsGlobals.T_InfoArchivo
    Select Case TypeName(fich)
    Case "Nothing", "Empty"
        info.EsValido = False
        AnalizarArchivo = info
    Case "Workbook"
        Dim f As clsExcelFile, owb As Workbook
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
'@Returns: T_InfoArchivo del workbook activo, o estructura vac�a si no hay activo
Public Function AnalizarArchivoActivo() As mod_ConstantsGlobals.T_InfoArchivo
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    If wb Is Nothing Then
        ' Devolver estructura vac�a
        Dim emptyInfo As mod_ConstantsGlobals.T_InfoArchivo
        emptyInfo.EsValido = False
        AnalizarArchivoActivo = emptyInfo
        Exit Function
    End If
    
    ' Sincronizar p_currExcelFile si es necesario
    If p_currExcelFile Is Nothing Or p_currExcelFile.ObjectKey <> ObjPtr(wb) Then
        Set p_currExcelFile = GetOrTrackWorkbook(wb)
    End If
    
    ' Devolver el an�lisis
    AnalizarArchivoActivo = p_currExcelFile.info
End Function

' ==================================================================
' M�TODOS DE UTILIDAD
' ==================================================================

'@Description: Verifica si un workbook est� siendo tracked
Public Function IsTracked(wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    IsTracked = Not (FindTrackedWorkbook(wb) Is Nothing)
End Function

'@Description: Busca el clsExcelFile correspondiente a un Workbook
'@Returns: clsExcelFile si est� tracked, Nothing si no
Private Function FindTrackedWorkbook(wb As Workbook) As clsExcelFile
    If wb Is Nothing Then Exit Function
    
    ' Crear instancia temporal solo para obtener la clave
    Dim fTemp As clsExcelFile
    Set fTemp = New clsExcelFile
    Dim key As Double
    key = fTemp.CreateFromWorkbook(wb).ObjectKey
    
    If p_trackedFiles.Exists(key) Then
        Set FindTrackedWorkbook = p_trackedFiles(key)
    End If
End Function

'@Description: Obtiene la clave de un archivo que implementa IFile
'@Returns: ObjectKey del archivo
'@Note: Solo acepta objetos que implementen IFile
Private Function GetFileKey(f As Object) As Double
    On Error GoTo ErrHandler

    ' Validar que el objeto implementa IFile
    Dim ifile As IFile
    Set ifile = f  ' Esto falla si no implementa IFile

    ' Usar ObjectKey de la interfaz
    GetFileKey = ifile.ObjectKey

    ' Validar que sea válido
    If GetFileKey = 0 Then
        LogWarning "clsFileManager", "[GetFileKey] ObjectKey = 0 para " & TypeName(f)
    End If

    Exit Function

ErrHandler:
    LogError "clsFileManager", "[GetFileKey] Error para tipo " & TypeName(f), Err.Number, Err.Description
    LogWarning "clsFileManager", "[GetFileKey] El objeto no implementa IFile correctamente"
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
    LogError "clsFileManager", "[GetFilePath] Error para tipo " & TypeName(f)
    GetFilePath = "(desconocido)"
End Function

'@Description: Obtiene informaci�n de todos los archivos tracked (para debugging)
Public Function GetTrackedFilesInfo() As String
    Dim info As String
    info = "Archivos tracked: " & p_trackedFiles.Count & vbCrLf
    
    Dim f As Object
    Dim key As Variant
    
    For Each key In p_trackedFiles.Keys
        Set f = p_trackedFiles(key)
        info = info & "  - [" & TypeName(f) & "] " & GetFilePath(f) & vbCrLf
    Next
    
    If Not p_currExcelFile Is Nothing Then
        info = info & vbCrLf & "Archivo activo: [" & TypeName(p_currExcelFile) & "] " & GetFilePath(p_currExcelFile)
    Else
        info = info & vbCrLf & "Archivo activo: (ninguno)"
    End If
    
    GetTrackedFilesInfo = info
End Function

' ==================================================================
' IMPLEMENTACION DE EVENTOS: CALLBACKS
' ==================================================================

'@Description: Callback cuando se activa un workbook (desde clsExecutionContext)
Private Sub ctx_WorkbookActivated(ByVal wb As Workbook)
    LogDebug "clsFileManager", "[ctx_WorkbookActivated] " & wb.Name
    Set p_currExcelFile = GetOrTrackWorkbook(wb)
End Sub

'@Description: Callback cuando se abre un workbook (desde clsExecutionContext)
Private Sub ctx_WorkbookOpened(ByVal wb As Workbook)
    LogDebug "clsFileManager", "[ctx_WorkbookOpened] " & wb.Name
    GetOrTrackWorkbook wb
End Sub

'@Description: Callback cuando se cierra un workbook (desde clsExecutionContext)
Private Sub ctx_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug "clsFileManager", "[ctx_WorkbookBeforeClose] " & wb.Name
    UntrackWorkbook wb
End Sub


Attribute VB_Name = "modActivationContext"
' =====================================================
' MÓDULO DE ACTIVATION CONTEXT API
' =====================================================
' Permite cargar componentes COM .NET sin registro en Windows.
' Usa manifests para activar el contexto de ejecución correcto.
'
' Uso:
'   1. Llamar a InicializarActivationContext con la ruta al manifest
'   2. Llamar a ActivarContexto antes de crear el objeto COM
'   3. Crear el objeto COM con CreateObject
'   4. Llamar a DesactivarContexto después de crear el objeto
'
' Ejemplo:
'   If InicializarActivationContext(rutaManifest) Then
'       If ActivarContexto() Then
'           Set obj = CreateObject("FolderWatcher.Monitor")
'           DesactivarContexto
'       End If
'   End If
' =====================================================

'@Folder "1-Inicio e Instalacion"
Option Explicit

Private Const MODULE_NAME As String = "modActivationContext"

' =====================================================
' DECLARACIONES DE API DE WINDOWS
' =====================================================

#If VBA7 Then
    Private Declare PtrSafe Function CreateActCtxW Lib "kernel32" _
        (ByRef pActCtx As actCtx) As LongPtr

    Private Declare PtrSafe Function ActivateActCtx Lib "kernel32" _
        (ByVal hActCtx As LongPtr, ByRef lpCookie As LongPtr) As Long

    Private Declare PtrSafe Function DeactivateActCtx Lib "kernel32" _
        (ByVal dwFlags As Long, ByVal ulCookie As LongPtr) As Long

    Private Declare PtrSafe Sub ReleaseActCtx Lib "kernel32" _
        (ByVal hActCtx As LongPtr)

    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
    Private Declare Function CreateActCtxW Lib "kernel32" _
        (ByRef pActCtx As ACTCTX) As Long

    Private Declare Function ActivateActCtx Lib "kernel32" _
        (ByVal hActCtx As Long, ByRef lpCookie As Long) As Long

    Private Declare Function DeactivateActCtx Lib "kernel32" _
        (ByVal dwFlags As Long, ByVal ulCookie As Long) As Long

    Private Declare Sub ReleaseActCtx Lib "kernel32" _
        (ByVal hActCtx As Long)

    Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

' =====================================================
' ESTRUCTURA ACTCTX
' =====================================================

Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Long = &H4

#If VBA7 Then
    Private Type actCtx
        cbSize As Long
        dwFlags As Long
        lpSource As LongPtr          ' Puntero a la ruta del manifest
        wProcessorArchitecture As Integer
        wLangId As Integer
        lpAssemblyDirectory As LongPtr
        lpResourceName As LongPtr
        lpApplicationName As LongPtr
        hModule As LongPtr
    End Type
#Else
    Private Type actCtx
        cbSize As Long
        dwFlags As Long
        lpSource As Long
        wProcessorArchitecture As Integer
        wLangId As Integer
        lpAssemblyDirectory As Long
        lpResourceName As Long
        lpApplicationName As Long
        hModule As Long
    End Type
#End If

' =====================================================
' CONSTANTES
' =====================================================

#If VBA7 Then
    Private Const INVALID_HANDLE_VALUE As LongPtr = -1
#Else
    Private Const INVALID_HANDLE_VALUE As Long = -1
#End If

' =====================================================
' VARIABLES DE MÓDULO
' =====================================================

#If VBA7 Then
    Private mhActCtx As LongPtr          ' Handle del contexto de activación
    Private mCookie As LongPtr           ' Cookie de activación
#Else
    Private mhActCtx As Long
    Private mCookie As Long
#End If

Private mContextoActivo As Boolean       ' Indica si el contexto está activo
Private mContextoInicializado As Boolean ' Indica si el contexto fue creado
Private mRutaManifest As String          ' Ruta al archivo manifest
Private mRutaDLL As String               ' Ruta a la DLL del COM

' =====================================================
' FUNCIONES PÚBLICAS
' =====================================================

'@Description: Inicializa el Activation Context con la ruta al manifest
'@Returns: True si se inicializó correctamente
Public Function InicializarActivationContext(ByVal rutaManifest As String) As Boolean
    On Error GoTo ErrHandler

    Dim actCtx As actCtx
    Dim fso As Object

    ' Verificar que el manifest existe
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rutaManifest) Then
        LogError MODULE_NAME, "Manifest no encontrado: " & rutaManifest
        InicializarActivationContext = False
        Exit Function
    End If

    ' Si ya hay un contexto, liberarlo primero
    If mContextoInicializado Then
        LiberarActivationContext
    End If

    mRutaManifest = rutaManifest
    mRutaDLL = fso.GetParentFolderName(rutaManifest) & "\" & _
               fso.GetBaseName(rutaManifest) & ".dll"

    ' Verificar que la DLL existe
    If Not fso.FileExists(mRutaDLL) Then
        ' Intentar sin la extensión .manifest
        mRutaDLL = fso.GetParentFolderName(rutaManifest) & "\FolderWatcherCOM.dll"
        If Not fso.FileExists(mRutaDLL) Then
            LogError MODULE_NAME, "DLL no encontrada: " & mRutaDLL
            InicializarActivationContext = False
            Exit Function
        End If
    End If

    ' Configurar estructura ACTCTX
    actCtx.cbSize = LenB(actCtx)
    actCtx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID
    actCtx.lpSource = StrPtr(rutaManifest)
    actCtx.lpAssemblyDirectory = StrPtr(fso.GetParentFolderName(rutaManifest))

    ' Crear el contexto de activación
    mhActCtx = CreateActCtxW(actCtx)

    If mhActCtx = INVALID_HANDLE_VALUE Then
        Dim lastErr As Long
        lastErr = GetLastError()
        LogError MODULE_NAME, "Error creando Activation Context. Código: " & lastErr
        InicializarActivationContext = False
        Exit Function
    End If

    mContextoInicializado = True
    mContextoActivo = False

    LogInfo MODULE_NAME, "Activation Context inicializado: " & rutaManifest
    InicializarActivationContext = True

    Set fso = Nothing
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "Error en InicializarActivationContext", Err.Number, Err.Description
    InicializarActivationContext = False
End Function

'@Description: Activa el contexto antes de crear el objeto COM
'@Returns: True si se activó correctamente
Public Function ActivarContexto() As Boolean
    On Error GoTo ErrHandler

    If Not mContextoInicializado Then
        LogWarning MODULE_NAME, "Contexto no inicializado"
        ActivarContexto = False
        Exit Function
    End If

    If mContextoActivo Then
        ' Ya está activo
        ActivarContexto = True
        Exit Function
    End If

    Dim resultado As Long
    resultado = ActivateActCtx(mhActCtx, mCookie)

    If resultado = 0 Then
        Dim lastErr As Long
        lastErr = GetLastError()
        LogError MODULE_NAME, "Error activando contexto. Código: " & lastErr
        ActivarContexto = False
        Exit Function
    End If

    mContextoActivo = True
    LogDebug MODULE_NAME, "Contexto activado"
    ActivarContexto = True
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "Error en ActivarContexto", Err.Number, Err.Description
    ActivarContexto = False
End Function

'@Description: Desactiva el contexto después de crear el objeto COM
Public Sub DesactivarContexto()
    On Error Resume Next

    If Not mContextoActivo Then Exit Sub

    DeactivateActCtx 0, mCookie
    mContextoActivo = False
    mCookie = 0

    LogDebug MODULE_NAME, "Contexto desactivado"
End Sub

'@Description: Libera completamente el Activation Context
Public Sub LiberarActivationContext()
    On Error Resume Next

    ' Desactivar si está activo
    If mContextoActivo Then
        DesactivarContexto
    End If

    ' Liberar el handle
    If mContextoInicializado And mhActCtx <> INVALID_HANDLE_VALUE Then
        ReleaseActCtx mhActCtx
        mhActCtx = INVALID_HANDLE_VALUE
    End If

    mContextoInicializado = False
    mRutaManifest = ""
    mRutaDLL = ""

    LogInfo MODULE_NAME, "Activation Context liberado"
End Sub

'@Description: Crea un objeto COM usando el Activation Context
'@Returns: El objeto COM creado, o Nothing si falla
Public Function CrearObjetoCOM(ByVal progId As String) As Object
    On Error GoTo ErrHandler

    Set CrearObjetoCOM = Nothing

    If Not mContextoInicializado Then
        LogError MODULE_NAME, "Contexto no inicializado para crear " & progId
        Exit Function
    End If

    ' Activar contexto
    If Not ActivarContexto() Then
        LogError MODULE_NAME, "No se pudo activar contexto para " & progId
        Exit Function
    End If

    ' Crear el objeto
    Set CrearObjetoCOM = CreateObject(progId)

    ' Desactivar contexto (el objeto ya está creado y funcionará)
    DesactivarContexto

    LogInfo MODULE_NAME, "Objeto COM creado: " & progId
    Exit Function

ErrHandler:
    DesactivarContexto
    LogError MODULE_NAME, "Error creando objeto COM: " & progId, Err.Number, Err.Description
    Set CrearObjetoCOM = Nothing
End Function

' =====================================================
' PROPIEDADES DE CONSULTA
' =====================================================

Public Property Get EstaInicializado() As Boolean
    EstaInicializado = mContextoInicializado
End Property

Public Property Get EstaActivo() As Boolean
    EstaActivo = mContextoActivo
End Property

Public Property Get rutaManifest() As String
    rutaManifest = mRutaManifest
End Property

Public Property Get rutaDLL() As String
    rutaDLL = mRutaDLL
End Property

' =====================================================
' FUNCIONES DE UTILIDAD
' =====================================================

'@Description: Obtiene la ruta donde debería estar el COM (carpeta AddIns)
Public Function ObtenerRutaCOM() As String
    ObtenerRutaCOM = Application.UserLibraryPath & "FolderWatcherCOM.dll"
End Function

'@Description: Obtiene la ruta donde debería estar el manifest (carpeta AddIns)
Public Function ObtenerRutaManifest() As String
    ObtenerRutaManifest = Application.UserLibraryPath & "FolderWatcherCOM.dll.manifest"
End Function

'@Description: Verifica si los archivos del COM están presentes
Public Function ComprobarArchivosCOM() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ComprobarArchivosCOM = fso.FileExists(ObtenerRutaCOM()) And _
                           fso.FileExists(ObtenerRutaManifest())

    Set fso = Nothing
End Function

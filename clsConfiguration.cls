VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Control de estado"
'@IgnoreModule ProcedureNotUsed
'TODO Is this comment still valid? => TODO Migrar de Registry a archivo JSON para configuración
'TODO Is this comment still valid? => TODO Añadir validación de rutas al guardar
'TODO Is this comment still valid? => TODO Implementar importar/exportar configuración
Option Explicit

Private mRutaOportunidades As String
Private mRutaPlantillas As String
Private mRutaOfergas As String
Private mRutaGasVBNet As String
Private mRutaExcelCalcTempl As String
Private mArrComprImgs As Variant
Private mArrComprDrawPIDs As Variant
Private mSAM As Integer

Public Property Get RutaOportunidades() As String
    RutaOportunidades = mRutaOportunidades
End Property

Public Property Let RutaOportunidades(newRuta As String)
    mRutaOportunidades = newRuta
    GuardarRuta CFG_RUTA_OPORTUNIDADES, newRuta
End Property

Public Property Get RutaPlantillas() As String
    RutaPlantillas = mRutaPlantillas
End Property

Public Property Let RutaPlantillas(newRuta As String)
    mRutaPlantillas = newRuta
    GuardarRuta CFG_RUTA_PLANTILLAS, newRuta
End Property

Public Property Get RutaOfergas() As String
    RutaOfergas = mRutaOfergas
End Property

Public Property Let RutaOfergas(newRuta As String)
    mRutaOfergas = newRuta
    GuardarRuta CFG_RUTA_OFERGAS, newRuta
End Property

Public Property Get RutaGasVBNet() As String
    RutaGasVBNet = mRutaGasVBNet
End Property

Public Property Let RutaGasVBNet(newRuta As String)
    mRutaGasVBNet = newRuta
    GuardarRuta CFG_RUTA_GAS_VBNET, newRuta
End Property

Public Property Get RutaExcelCalcTempl() As String
    RutaExcelCalcTempl = mRutaExcelCalcTempl
End Property

Public Property Let RutaExcelCalcTempl(newRuta As String)
    mRutaExcelCalcTempl = newRuta
    GuardarRuta CFG_RUTA_PLCALCS, newRuta
End Property

Public Property Get ListComprImgs() As Variant
    ListComprImgs = mArrComprImgs
End Property

Public Property Let ListComprImgs(arrRutas As Variant)
    mArrComprImgs = arrRutas
    GuardarListaCarpetas arrRutas, CFG_RUTA_COMPRIMGS
End Property

Public Property Get ListComprDrawPIDs() As Variant
    ListComprDrawPIDs = mArrComprDrawPIDs
End Property

Public Property Let ListComprDrawPIDs(arrRutas As Variant)
    mArrComprDrawPIDs = arrRutas
    GuardarListaCarpetas arrRutas, CFG_RUTA_COMPRDRAWPID
End Property

Public Property Get SAM() As Integer
    SAM = mSAM
End Property

Public Property Let SAM(newSAM As Integer)
    mSAM = newSAM
    GuardarSAM newSAM
End Property

Private Sub Class_Initialize()
    Debug.Print vbTab & vbTab & "[clsConfiguracion Class_Initialize]"
    CargarConfiguracionDesdeRegistro
End Sub

Private Sub Class_Terminate()
    Debug.Print vbTab & vbTab & "[clsConfiguracion Class_Terminate]"
End Sub

Private Sub CargarConfiguracionDesdeRegistro()
    On Error GoTo ErrHandler
    
    ' Cargar cada ruta desde el registro y mostrarla
    Debug.Print vbTab & vbTab & "[clsConfiguracion CargarConfiguracionDesdeRegistro] - inicializacion de todas las rutas desde registro o valores por defecto"
    mRutaOportunidades = ObtenerRuta(CFG_RUTA_OPORTUNIDADES, CFG_RUTA_OPORTUNIDADES_DEFAULT)
    mRutaPlantillas = ObtenerRuta(CFG_RUTA_PLANTILLAS, CFG_RUTA_PLANTILLAS_DEFAULT)
    mRutaOfergas = ObtenerRuta(CFG_RUTA_OFERGAS, CFG_RUTA_OFERGAS_DEFAULT)
    mRutaGasVBNet = ObtenerRuta(CFG_RUTA_GAS_VBNET, CFG_RUTA_GAS_VBNET_DEFAULT)
    mRutaExcelCalcTempl = ObtenerRuta(CFG_RUTA_PLCALCS, CFG_RUTA_PLCALCS_DEFAULT)
    mArrComprImgs = ObtenerListaCarpetas(CFG_RUTA_COMPRIMGS, CFG_RUTA_COMPRIMGS_DEFAULT)
    mArrComprDrawPIDs = ObtenerListaCarpetas(CFG_RUTA_COMPRDRAWPID, CFG_RUTA_COMPRDRAWPID_DEFAULT)
    mSAM = ObtenerSAM()

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsConfiguration.CargarConfiguracionDesdeRegistro", _
              "Error en la carga de configuración desde registro: " & Err.Description
End Sub

'TODO: revisar el uso de esta función: Buscar otra forma de que el folder watcher tenga conocimiento de qué carpetas son supervisadas;
' posiblemente presentando de otra manera el formulario de configuración de carpetas, y creando en él opciones para su selección...
Public Property Get oDicFoldersToWatch() As Object
    Set oDicFoldersToWatch = CreateObject("scripting.dictionary")
    oDicFoldersToWatch.Add CFG_RUTA_OPORTUNIDADES, RutaOportunidades
    oDicFoldersToWatch.Add CFG_RUTA_PLANTILLAS, RutaPlantillas
    oDicFoldersToWatch.Add CFG_RUTA_OFERGAS, RutaOfergas
    oDicFoldersToWatch.Add CFG_RUTA_GAS_VBNET, RutaGasVBNet
    oDicFoldersToWatch.Add CFG_RUTA_PLCALCS, RutaExcelCalcTempl
End Property

'-------------------------------------------
' PERSISTENCIA (Registro de Windows)
'-------------------------------------------
Private Sub GuardarSAM(SAM As Integer)
    CreateObject("WScript.Shell").RegWrite CFG_PATH_SAM, CStr(SAM), "REG_SZ"
    Debug.Print vbTab & vbTab & "[clsConfiguracion GuardarSAM] - S.A.M. guardado en registro"
End Sub

Private Function ObtenerSAM() As Integer
    Dim valorLeido As Variant

    On Error Resume Next
    valorLeido = CreateObject("WScript.Shell").RegRead(CFG_PATH_SAM)
    Dim errNum As Long: errNum = Err.Number
    On Error GoTo 0

    ' Verificar si hubo error o valor invalido
    If errNum <> 0 Or IsEmpty(valorLeido) Or Not IsNumeric(valorLeido) Then
        Debug.Print vbTab & vbTab & "[clsConfiguracion ObtenerSAM] - no existente en registro o invalido, establecido valor por defecto"
        ObtenerSAM = CFG_SAM
        GuardarSAM CFG_SAM
    ElseIf CInt(valorLeido) = 0 Then
        Debug.Print vbTab & vbTab & "[clsConfiguration.ObtenerSAM] Valor cero en registro, usando valor por defecto"
        ObtenerSAM = CFG_SAM
        GuardarSAM CFG_SAM
    Else
        ObtenerSAM = CInt(valorLeido)
        Debug.Print vbTab & vbTab & "[clsConfiguration.ObtenerSAM] Leido del registro: " & ObtenerSAM
    End If
End Function


' Guardar una ruta en el registro
Private Sub GuardarRuta(nombreConfig As Variant, ruta As String)
    'SaveSetting APP_NAME & "\" & CFG_SECTION_RUTAS, nombreConfig, "", ruta
    CreateObject("WScript.Shell").RegWrite CFG_PATH_SECTION_RUTAS & nombreConfig & "\", ruta, "REG_SZ"
    Debug.Print vbTab & "[clsConfiguracion GuardarRuta] - ruta guardada en registro"
End Sub

' Recuperar una ruta del registro
Private Function ObtenerRuta(nombreConfig As Variant, strDefault As String) As String
    'ObtenerRuta = GetSetting(APP_NAME & "\" & CFG_SECTION_RUTAS, nombreConfig, "", "")
    On Error Resume Next
    ObtenerRuta = CreateObject("WScript.Shell").RegRead(CFG_PATH_SECTION_RUTAS & nombreConfig & "\")
    If ObtenerRuta = "" Or Not RutaExiste(ObtenerRuta) Then
        ObtenerRuta = strDefault
        GuardarRuta nombreConfig, ObtenerRuta
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion ObtenerRuta] - ruta no existente en registro o invalida, establecido valor por defecto"
    Else
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion ObtenerRuta] - ruta leida del registro"
    End If
    On Error GoTo 0
End Function

' Guarda la lista de Carpetas
Private Sub GuardarListaCarpetas(ByVal carpetas As Variant, nombreConfig As String)
    Dim lista As String
    Dim key As Variant
    
    If UBound(carpetas) < 0 Then
        Stop
        BorrarListaCarpetas nombreConfig
        Exit Sub
    End If
    
    For Each key In carpetas
        If lista <> "" Then lista = lista & "|"
        lista = lista & key
    Next key
    
    On Error Resume Next
    GuardarRuta nombreConfig, lista
    
    If Err.Number <> 0 Then
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion GuardarListaCarpetas] - No se pudo guardar lista en registro: " & Err.Description
    Else
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion GuardarListaCarpetas] - Lista de carpetas guardada en registro."
    End If
    
    On Error GoTo 0
End Sub

' Obtiene la lista de Carpetas
Private Function ObtenerListaCarpetas(nombreConfig As String, strDefault As String) As Variant
    On Error Resume Next
    ObtenerListaCarpetas = Split(CreateObject("WScript.Shell").RegRead(CFG_PATH_SECTION_RUTAS & nombreConfig & "\"), "|")
    
    If Err.Number <> 0 Then
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion ObtenerListaCarpetas] - No se encontró lista guardada en registro."
        ObtenerListaCarpetas = Split(strDefault, "|")
        Call GuardarListaCarpetas(ObtenerListaCarpetas, nombreConfig)
    Else
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion ObtenerListaCarpetas] - Lista de carpetas leida desde registro."
    End If
    On Error GoTo 0
End Function

' Borra la lista guardada
Private Sub BorrarListaCarpetas(nombreConfig As String)
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete CFG_PATH_SECTION_RUTAS & nombreConfig & "\"
    
    If Err.Number = 0 Then
        Debug.Print vbTab & vbTab & vbTab & "[clsConfiguracion BorrarListaCarpetas] - Lista de carpetas eliminada del registro."
    End If
    
    On Error GoTo 0
End Sub



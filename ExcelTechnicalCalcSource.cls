VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelTechnicalCalcSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: ExcelTechnicalCalcSource
' ==================================================================
' Implementa ITechnicalCalc leyendo datos de un fichero Excel de calculo.
' El fichero Excel ES la fuente de datos para esta entidad.
'
' Estructura tipica del fichero de calculo:
' - Hoja con datos del gas (composicion, propiedades)
' - Hoja con condiciones de proceso
' - Hoja con resultados del compresor
'
' Ficheros asociados identificados por patron:
' - *_calc.xlsx, *_calc_multi.xlsx (principal)
' - Antipul_*.xlsx (antipulsadores)
' - ABC_Gas_Cooler-*.xlsx, ABC_Aircooler-*.xlsx, etc. (auxiliares)
' ==================================================================

'@Folder "2-Infraestructura.5-1-Adaptadores.Excel"
Option Explicit
Implements ITechnicalCalc

Private Const MODULE_NAME As String = "ExcelTechnicalCalcSource"

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private mCalcNumber As String           ' ABC12345
Private mCalcOption As String           ' ABC12345_01
Private mRevision As Long               ' 0, 1, 2...
Private mFilePath As String             ' Ruta completa del fichero
Private mFileName As String             ' Solo nombre del fichero

Private mCompressorModel As String      ' Modelo resultante (2TEH-1-LG)
Private mIsProcessed As Boolean         ' Si se han leido los datos
Private mIsValid As Boolean             ' Si los datos son validos

Private mGasData As Object              ' Dictionary con datos del gas
Private mProcessData As Object          ' Dictionary con datos de proceso
Private mCompressorData As Object       ' Dictionary con datos del compresor
Private mAssociatedFiles As Collection  ' Ficheros relacionados

Private mHasAntipulsadorFile As Boolean
Private mHasAuxiliaryFiles As Boolean

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    Set mGasData = CreateObject("Scripting.Dictionary")
    Set mProcessData = CreateObject("Scripting.Dictionary")
    Set mCompressorData = CreateObject("Scripting.Dictionary")
    Set mAssociatedFiles = New Collection
End Sub

Private Sub Class_Terminate()
    Set mGasData = Nothing
    Set mProcessData = Nothing
    Set mCompressorData = Nothing
    Set mAssociatedFiles = Nothing
End Sub

'@Description: Inicializa el source con la informacion parseada del nombre de fichero
'@Param filePath: Ruta completa del fichero de calculo
'@Param parsedInfo: Dictionary con CalcOption, CalcNumber, Revision, etc.
Public Sub Initialize(ByVal filePath As String, parsedInfo As Object)
    On Error GoTo ErrHandler

    mFilePath = filePath

    ' Extraer nombre de fichero
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    mFileName = fso.GetFileName(filePath)

    ' Establecer datos parseados
    If Not parsedInfo Is Nothing Then
        If parsedInfo.Exists("CalcOption") Then mCalcOption = parsedInfo("CalcOption")
        If parsedInfo.Exists("CalcNumber") Then mCalcNumber = parsedInfo("CalcNumber")
        If parsedInfo.Exists("Revision") Then mRevision = parsedInfo("Revision")
    End If

    LogDebug MODULE_NAME, "[Initialize] " & mCalcOption & " rev" & mRevision
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
End Sub

' ==================================================================
' IMPLEMENTACION DE ITechnicalCalc
' ==================================================================

Private Property Get ITechnicalCalc_CalcNumber() As String
    ITechnicalCalc_CalcNumber = mCalcNumber
End Property

Private Property Get ITechnicalCalc_CalcOption() As String
    ITechnicalCalc_CalcOption = mCalcOption
End Property

Private Property Get ITechnicalCalc_Revision() As Long
    ITechnicalCalc_Revision = mRevision
End Property

Private Property Get ITechnicalCalc_FilePath() As String
    ITechnicalCalc_FilePath = mFilePath
End Property

Private Property Get ITechnicalCalc_FileName() As String
    ITechnicalCalc_FileName = mFileName
End Property

Private Property Get ITechnicalCalc_CompressorModel() As String
    If Not mIsProcessed Then LoadData
    ITechnicalCalc_CompressorModel = mCompressorModel
End Property

Private Property Get ITechnicalCalc_IsValid() As Boolean
    If Not mIsProcessed Then LoadData
    ITechnicalCalc_IsValid = mIsValid
End Property

Private Property Get ITechnicalCalc_IsProcessed() As Boolean
    ITechnicalCalc_IsProcessed = mIsProcessed
End Property

Private Function ITechnicalCalc_GetGasData() As Object
    If Not mIsProcessed Then LoadData
    Set ITechnicalCalc_GetGasData = mGasData
End Function

Private Function ITechnicalCalc_GetProcessData() As Object
    If Not mIsProcessed Then LoadData
    Set ITechnicalCalc_GetProcessData = mProcessData
End Function

Private Function ITechnicalCalc_GetCompressorData() As Object
    If Not mIsProcessed Then LoadData
    Set ITechnicalCalc_GetCompressorData = mCompressorData
End Function

Private Function ITechnicalCalc_GetAssociatedFiles() As Collection
    If Not mIsProcessed Then LoadData
    Set ITechnicalCalc_GetAssociatedFiles = mAssociatedFiles
End Function

Private Property Get ITechnicalCalc_HasAntipulsadorFile() As Boolean
    If Not mIsProcessed Then LoadData
    ITechnicalCalc_HasAntipulsadorFile = mHasAntipulsadorFile
End Property

Private Property Get ITechnicalCalc_HasAuxiliaryFiles() As Boolean
    If Not mIsProcessed Then LoadData
    ITechnicalCalc_HasAuxiliaryFiles = mHasAuxiliaryFiles
End Property

' ==================================================================
' PROPIEDADES PUBLICAS
' ==================================================================

Public Property Get calcNumber() As String
    calcNumber = mCalcNumber
End Property

Public Property Get calcOption() As String
    calcOption = mCalcOption
End Property

Public Property Get Revision() As Long
    Revision = mRevision
End Property

Public Property Get filePath() As String
    filePath = mFilePath
End Property

Public Property Get fileName() As String
    fileName = mFileName
End Property

Public Property Get CompressorModel() As String
    If Not mIsProcessed Then LoadData
    CompressorModel = mCompressorModel
End Property

Public Property Get IsValid() As Boolean
    If Not mIsProcessed Then LoadData
    IsValid = mIsValid
End Property

Public Property Get IsProcessed() As Boolean
    IsProcessed = mIsProcessed
End Property

' ==================================================================
' CARGA DE DATOS DEL FICHERO EXCEL
' ==================================================================

'@Description: Carga los datos del fichero Excel de calculo
Private Sub LoadData()
    On Error GoTo ErrHandler

    If mIsProcessed Then Exit Sub

    LogDebug MODULE_NAME, "[LoadData] Cargando datos de: " & mFileName

    ' Verificar que el fichero existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(mFilePath) Then
        LogWarning MODULE_NAME, "[LoadData] Fichero no existe: " & mFilePath
        mIsProcessed = True
        mIsValid = False
        Exit Sub
    End If

    ' Leer datos del Excel
    If Right(LCase(mFilePath), 5) = ".xlsx" Or Right(LCase(mFilePath), 5) = ".xlsm" Then
        LoadExcelData
    ElseIf Right(LCase(mFilePath), 4) = ".txt" Then
        LoadTextData
    ElseIf Right(LCase(mFilePath), 4) = ".rtf" Then
        LoadRtfData
    End If

    ' Buscar ficheros asociados
    FindAssociatedFiles

    mIsProcessed = True
    mIsValid = Len(mCompressorModel) > 0

    LogDebug MODULE_NAME, "[LoadData] Completado. Modelo: " & mCompressorModel & ", Valido: " & mIsValid
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadData]"
    mIsProcessed = True
    mIsValid = False
End Sub

'@Description: Carga datos de un fichero Excel
Private Sub LoadExcelData()
    On Error GoTo ErrHandler

    ' TODO: Implementar lectura de datos especificos del Excel
    ' Por ahora solo marcamos como procesado
    ' La implementacion completa dependera de la estructura del fichero

    LogDebug MODULE_NAME, "[LoadExcelData] Pendiente implementacion completa"

    ' Placeholder: intentar extraer modelo del nombre del fichero si no hay otra fuente
    ' Esto se mejorara cuando se implemente la lectura real del Excel

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadExcelData]"
End Sub

'@Description: Carga datos de un fichero de texto
Private Sub LoadTextData()
    On Error GoTo ErrHandler

    ' TODO: Implementar lectura de datos de fichero TXT
    LogDebug MODULE_NAME, "[LoadTextData] Pendiente implementacion"

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadTextData]"
End Sub

'@Description: Carga datos de un fichero RTF
Private Sub LoadRtfData()
    On Error GoTo ErrHandler

    ' TODO: Implementar lectura de datos de fichero RTF
    LogDebug MODULE_NAME, "[LoadRtfData] Pendiente implementacion"

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadRtfData]"
End Sub

'@Description: Busca ficheros asociados al calculo en la misma carpeta
Private Sub FindAssociatedFiles()
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(mFilePath) Then Exit Sub

    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(mFilePath)

    If Not fso.FolderExists(parentFolder) Then Exit Sub

    Dim folder As Object
    Set folder = fso.GetFolder(parentFolder)

    ' Patrones para ficheros asociados
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True

    ' Buscar ficheros que coincidan con la opcion de calculo
    Dim file As Object
    For Each file In folder.files
        ' Ignorar ficheros temporales
        If Left(file.Name, 2) = "~$" Then GoTo NextFile

        ' Ignorar el fichero principal
        If file.Name = mFileName Then GoTo NextFile

        ' Verificar si contiene la opcion de calculo
        If InStr(1, file.Name, mCalcOption, vbTextCompare) > 0 Then
            mAssociatedFiles.Add file.path

            ' Detectar tipos especiales
            If InStr(1, file.Name, "Antipul", vbTextCompare) > 0 Then
                mHasAntipulsadorFile = True
            End If

            If InStr(1, file.Name, "Gas_Cooler", vbTextCompare) > 0 Or _
               InStr(1, file.Name, "Aircooler", vbTextCompare) > 0 Or _
               InStr(1, file.Name, "Reducer", vbTextCompare) > 0 Or _
               InStr(1, file.Name, "Main Motor", vbTextCompare) > 0 Or _
               InStr(1, file.Name, "Instrumentation", vbTextCompare) > 0 Then
                mHasAuxiliaryFiles = True
            End If
        End If

NextFile:
    Next file

    LogDebug MODULE_NAME, "[FindAssociatedFiles] Encontrados " & mAssociatedFiles.Count & " ficheros asociados"

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[FindAssociatedFiles]"
End Sub

' ==================================================================
' METODOS PUBLICOS ADICIONALES
' ==================================================================

'@Description: Fuerza la recarga de datos del fichero
Public Sub Reload()
    mIsProcessed = False
    Set mGasData = CreateObject("Scripting.Dictionary")
    Set mProcessData = CreateObject("Scripting.Dictionary")
    Set mCompressorData = CreateObject("Scripting.Dictionary")
    Set mAssociatedFiles = New Collection
    mHasAntipulsadorFile = False
    mHasAuxiliaryFiles = False

    LoadData
End Sub

'@Description: Obtiene un resumen del calculo
Public Function GetSummary() As String
    If Not mIsProcessed Then LoadData

    Dim summary As String
    summary = "Calculo: " & mCalcOption
    If mRevision > 0 Then summary = summary & " rev" & mRevision
    summary = summary & vbCrLf
    summary = summary & "  Fichero: " & mFileName & vbCrLf
    summary = summary & "  Modelo: " & IIf(Len(mCompressorModel) > 0, mCompressorModel, "(pendiente)") & vbCrLf
    summary = summary & "  Valido: " & IIf(mIsValid, "Si", "No") & vbCrLf
    summary = summary & "  Ficheros asociados: " & mAssociatedFiles.Count

    GetSummary = summary
End Function

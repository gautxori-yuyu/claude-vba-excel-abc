VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelEconomicQuoteSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: ExcelEconomicQuoteSource
' ==================================================================
' Implementa IEconomicQuote leyendo datos de un fichero Excel de valoracion.
' El fichero Excel ES la fuente de datos para esta entidad.
'
' Estructura tipica de una valoracion:
' [Oportunidad]/3.VALORACION ECONOMICA/
'   +-- 123456789-01.xlsx (compresor 1)
'   +-- 123456789-02.xlsx (compresor 2)
'   +-- rev 1/
'   |   +-- 123456789-01 rev.1.xlsx
'   +-- rev 2/
'       +-- 123456789-01 rev.2.xlsx
'
' Patron de nombre: ^(\d{9})(?:[\-_](\d+))?(?:[ \-_]*rev\.?[ \-_]*(\d+))?.*\.xls[xm]$
' ==================================================================

'@Folder "2-Infraestructura.5-1-Adaptadores.Excel"
Option Explicit
Implements IEconomicQuote

Private Const MODULE_NAME As String = "ExcelEconomicQuoteSource"

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private mQuoteNumber As String          ' 123456789
Private mCompressorIndex As String      ' 01, 02...
Private mQuoteId As String              ' 123456789-01
Private mRevision As Long               ' 0, 1, 2...
Private mFilePath As String             ' Ruta completa del fichero
Private mFileName As String             ' Solo nombre del fichero

Private mIsValid As Boolean             ' Si los datos son validos
Private mIsDirty As Boolean             ' Si hay cambios sin guardar
Private mIsNew As Boolean               ' Si es nuevo (nunca guardado)
Private mLastModified As Date           ' Fecha de ultima modificacion

Private mAssociatedCalcOption As String ' Calculo asociado
Private mHasAssociatedCalc As Boolean

Private mGeneralData As Object          ' Dictionary con datos generales
Private mPricingData As Object          ' Dictionary con precios
Private mOtrosData As Collection        ' Collection de Dictionary para "otros"

Private mIsLoaded As Boolean            ' Si se han cargado los datos

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    Set mGeneralData = CreateObject("Scripting.Dictionary")
    Set mPricingData = CreateObject("Scripting.Dictionary")
    Set mOtrosData = New Collection
    mCompressorIndex = "01"  ' Default
End Sub

Private Sub Class_Terminate()
    Set mGeneralData = Nothing
    Set mPricingData = Nothing
    Set mOtrosData = Nothing
End Sub

'@Description: Inicializa el source con la informacion parseada del nombre de fichero
'@Param filePath: Ruta completa del fichero de valoracion
'@Param parsedInfo: Dictionary con QuoteNumber, CompressorIndex, Revision, etc.
Public Sub Initialize(ByVal filePath As String, parsedInfo As Object)
    On Error GoTo ErrHandler

    mFilePath = filePath

    ' Extraer nombre de fichero y fecha de modificacion
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        mFileName = fso.GetFileName(filePath)
        mLastModified = fso.GetFile(filePath).DateLastModified
        mIsNew = False
    Else
        mIsNew = True
    End If

    ' Establecer datos parseados
    If Not parsedInfo Is Nothing Then
        If parsedInfo.Exists("QuoteNumber") Then mQuoteNumber = parsedInfo("QuoteNumber")
        If parsedInfo.Exists("CompressorIndex") Then mCompressorIndex = parsedInfo("CompressorIndex")
        If parsedInfo.Exists("Revision") Then mRevision = parsedInfo("Revision")
        If parsedInfo.Exists("QuoteId") Then mQuoteId = parsedInfo("QuoteId")
    End If

    ' Construir QuoteId si no viene en parsedInfo
    If Len(mQuoteId) = 0 Then
        mQuoteId = mQuoteNumber
        If Len(mCompressorIndex) > 0 And mCompressorIndex <> "01" Then
            mQuoteId = mQuoteId & "-" & mCompressorIndex
        End If
    End If

    LogDebug MODULE_NAME, "[Initialize] " & mQuoteId & " rev" & mRevision
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
End Sub

' ==================================================================
' IMPLEMENTACION DE IEconomicQuote
' ==================================================================

Private Property Get IEconomicQuote_QuoteNumber() As String
    IEconomicQuote_QuoteNumber = mQuoteNumber
End Property

Private Property Get IEconomicQuote_CompressorIndex() As String
    IEconomicQuote_CompressorIndex = mCompressorIndex
End Property

Private Property Get IEconomicQuote_QuoteId() As String
    IEconomicQuote_QuoteId = mQuoteId
End Property

Private Property Get IEconomicQuote_Revision() As Long
    IEconomicQuote_Revision = mRevision
End Property

Private Property Get IEconomicQuote_FilePath() As String
    IEconomicQuote_FilePath = mFilePath
End Property

Private Property Get IEconomicQuote_FileName() As String
    IEconomicQuote_FileName = mFileName
End Property

Private Property Get IEconomicQuote_IsValid() As Boolean
    If Not mIsLoaded Then LoadData
    IEconomicQuote_IsValid = mIsValid
End Property

Private Property Get IEconomicQuote_IsDirty() As Boolean
    IEconomicQuote_IsDirty = mIsDirty
End Property

Private Property Get IEconomicQuote_IsNew() As Boolean
    IEconomicQuote_IsNew = mIsNew
End Property

Private Property Get IEconomicQuote_LastModified() As Date
    IEconomicQuote_LastModified = mLastModified
End Property

Private Property Get IEconomicQuote_AssociatedCalcOption() As String
    If Not mIsLoaded Then LoadData
    IEconomicQuote_AssociatedCalcOption = mAssociatedCalcOption
End Property

Private Property Get IEconomicQuote_HasAssociatedCalc() As Boolean
    If Not mIsLoaded Then LoadData
    IEconomicQuote_HasAssociatedCalc = mHasAssociatedCalc
End Property

Private Function IEconomicQuote_GetGeneralData() As Object
    If Not mIsLoaded Then LoadData
    Set IEconomicQuote_GetGeneralData = mGeneralData
End Function

Private Function IEconomicQuote_GetPricingData() As Object
    If Not mIsLoaded Then LoadData
    Set IEconomicQuote_GetPricingData = mPricingData
End Function

Private Property Get IEconomicQuote_OtrosCount() As Long
    If Not mIsLoaded Then LoadData
    IEconomicQuote_OtrosCount = mOtrosData.Count
End Property

Private Function IEconomicQuote_GetOtroData(ByVal Index As Long) As Object
    If Not mIsLoaded Then LoadData

    If Index >= 0 And Index < mOtrosData.Count Then
        Set IEconomicQuote_GetOtroData = mOtrosData(Index + 1)
    Else
        Set IEconomicQuote_GetOtroData = Nothing
    End If
End Function

' ==================================================================
' PROPIEDADES PUBLICAS
' ==================================================================

Public Property Get quoteNumber() As String
    quoteNumber = mQuoteNumber
End Property

Public Property Get CompressorIndex() As String
    CompressorIndex = mCompressorIndex
End Property

Public Property Get quoteId() As String
    quoteId = mQuoteId
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

Public Property Get IsValid() As Boolean
    If Not mIsLoaded Then LoadData
    IsValid = mIsValid
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get LastModified() As Date
    LastModified = mLastModified
End Property

Public Property Get AssociatedCalcOption() As String
    If Not mIsLoaded Then LoadData
    AssociatedCalcOption = mAssociatedCalcOption
End Property

Public Property Get HasAssociatedCalc() As Boolean
    If Not mIsLoaded Then LoadData
    HasAssociatedCalc = mHasAssociatedCalc
End Property

' ==================================================================
' CARGA DE DATOS DEL FICHERO EXCEL
' ==================================================================

'@Description: Carga los datos del fichero Excel de valoracion
Private Sub LoadData()
    On Error GoTo ErrHandler

    If mIsLoaded Then Exit Sub

    LogDebug MODULE_NAME, "[LoadData] Cargando datos de: " & mFileName

    ' Verificar que el fichero existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(mFilePath) Then
        LogWarning MODULE_NAME, "[LoadData] Fichero no existe: " & mFilePath
        mIsLoaded = True
        mIsValid = False
        Exit Sub
    End If

    ' Actualizar fecha de modificacion
    mLastModified = fso.GetFile(mFilePath).DateLastModified

    ' Cargar datos del Excel
    LoadExcelData

    mIsLoaded = True
    mIsValid = Len(mQuoteNumber) > 0

    LogDebug MODULE_NAME, "[LoadData] Completado. QuoteId: " & mQuoteId & ", Valido: " & mIsValid
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadData]"
    mIsLoaded = True
    mIsValid = False
End Sub

'@Description: Carga datos especificos del fichero Excel de valoracion
Private Sub LoadExcelData()
    On Error GoTo ErrHandler

    ' TODO: Implementar lectura de datos especificos del Excel de valoracion
    ' La estructura dependera del formato del fichero de valoracion

    ' Por ahora, intentar detectar el calculo asociado por nombre de fichero
    ' o por contenido de celdas especificas

    LogDebug MODULE_NAME, "[LoadExcelData] Pendiente implementacion completa"

    ' Placeholder: marcar como valido si el fichero existe
    mIsValid = True

    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LoadExcelData]"
End Sub

' ==================================================================
' METODOS PUBLICOS ADICIONALES
' ==================================================================

'@Description: Fuerza la recarga de datos del fichero
Public Sub Reload()
    mIsLoaded = False
    Set mGeneralData = CreateObject("Scripting.Dictionary")
    Set mPricingData = CreateObject("Scripting.Dictionary")
    Set mOtrosData = New Collection
    mAssociatedCalcOption = ""
    mHasAssociatedCalc = False
    mIsDirty = False

    LoadData
End Sub

'@Description: Marca la valoracion como modificada
Public Sub MarkDirty()
    mIsDirty = True
End Sub

'@Description: Marca la valoracion como guardada
Public Sub MarkSaved()
    mIsDirty = False
    mIsNew = False

    ' Actualizar fecha de modificacion
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(mFilePath) Then
        mLastModified = fso.GetFile(mFilePath).DateLastModified
    End If
End Sub

'@Description: Asocia la valoracion con un calculo tecnico
Public Sub AssociateWithCalculation(ByVal calcOption As String)
    mAssociatedCalcOption = calcOption
    mHasAssociatedCalc = Len(calcOption) > 0
    mIsDirty = True
End Sub

'@Description: Obtiene un resumen de la valoracion
Public Function GetSummary() As String
    If Not mIsLoaded Then LoadData

    Dim summary As String
    summary = "Valoracion: " & mQuoteId
    If mRevision > 0 Then summary = summary & " rev" & mRevision
    summary = summary & vbCrLf
    summary = summary & "  Fichero: " & mFileName & vbCrLf
    summary = summary & "  Compresor: " & mCompressorIndex & vbCrLf
    summary = summary & "  Calculo asociado: " & IIf(mHasAssociatedCalc, mAssociatedCalcOption, "(ninguno)") & vbCrLf
    summary = summary & "  Valido: " & IIf(mIsValid, "Si", "No") & vbCrLf
    summary = summary & "  Modificado: " & Format(mLastModified, "yyyy-mm-dd hh:nn:ss")

    GetSummary = summary
End Function

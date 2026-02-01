VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelOpportunitySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' ADAPTADOR DE OPORTUNIDAD PARA EXCEL
' ==================================================================
' Implementa la interfaz IOpportunitySource para permitir que
' el dominio acceda a datos de oportunidad a través de Excel
' sin conocer los detalles de implementación de Excel.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores.Adaptadores"
Option Explicit

Private Const MODULE_NAME As String = "clsExcelOpportunitySource"

' Implementa la interfaz
Implements IOpportunitySource

' Referencia al archivo Excel subyacente
Private mFileXLS As clsFileXLS
Private mPath As String

' ==========================================
' INICIALIZACIÓN
' ==========================================
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
    Set mFileXLS = Nothing
End Sub

'@Description: Inicializa el adaptador con un archivo Excel
Public Sub Initialize(filePath As String)
    mPath = filePath
    
    ' Crear instancia de clsFileXLS para manejar el archivo
    Set mFileXLS = New clsFileXLS
    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)
    mFileXLS.BindTo wb
    
    LogInfo MODULE_NAME, "[Initialize] Adaptador inicializado para: " & filePath
End Sub

'@Description: Inicializa el adaptador con un clsFileXLS existente
Public Sub InitializeWithFile(fileXLS As clsFileXLS)
    Set mFileXLS = fileXLS
    If Not mFileXLS Is Nothing Then
        mPath = mFileXLS.Path
    End If
    
    LogInfo MODULE_NAME, "[InitializeWithFile] Adaptador inicializado con clsFileXLS: " & mPath
End Sub

' ==========================================
' IMPLEMENTACIÓN DE LA INTERFAZ IOpportunitySource
' ==========================================

Private Function IOpportunitySource_ReadOpportunityData() As clsOpportunity
    Set IOpportunitySource_ReadOpportunityData = CreateOpportunityFromExcel()
End Function

Private Function IOpportunitySource_ReadTechnicalData() As Object
    Set IOpportunitySource_ReadTechnicalData = ReadTechnicalDataFromExcel()
End Function

Private Function IOpportunitySource_ReadQuotationData() As Object
    Set IOpportunitySource_ReadQuotationData = ReadQuotationDataFromExcel()
End Function

Private Sub IOpportunitySource_SaveResults(resultado As Object)
    Call SaveResultsToExcel(resultado)
End Sub

Private Function IOpportunitySource_IsValidSource() As Boolean
    IOpportunitySource_IsValidSource = ValidateExcelSource()
End Function

Private Property Get IOpportunitySource_OpportunityIdentifier() As String
    IOpportunitySource_OpportunityIdentifier = GetOpportunityIdentifier()
End Property

' ==========================================
' MÉTODOS PRIVADOS DE IMPLEMENTACIÓN
' ==========================================

'@Description: Crea un objeto clsOpportunity desde los datos en Excel
Private Function CreateOpportunityFromExcel() As clsOpportunity
    Set CreateOpportunityFromExcel = New clsOpportunity
    
    If mFileXLS Is Nothing Then
        LogWarning MODULE_NAME, "[CreateOpportunityFromExcel] mFileXLS es Nothing"
        Exit Function
    End If
    
    Dim wb As Workbook
    Set wb = mFileXLS.wb
    
    ' Extraer información de la oportunidad desde el workbook
    With CreateOpportunityFromExcel
        .Path = mFileXLS.Path
        .Name = mFileXLS.Name
        
        ' Extraer información adicional del nombre de la carpeta o del archivo
        Dim parts() As String
        parts = Split(mFileXLS.Name, " - ")
        If UBound(parts) >= 0 Then
            .Label = Trim(parts(0))
        End If
        If UBound(parts) >= 1 Then
            .Customer = Trim(parts(1))
        End If
    End With
    
    LogDebug MODULE_NAME, "[CreateOpportunityFromExcel] Oportunidad creada: " & CreateOpportunityFromExcel.Label
End Function

'@Description: Lee datos técnicos desde Excel
Private Function ReadTechnicalDataFromExcel() As Object
    Set ReadTechnicalDataFromExcel = CreateObject("Scripting.Dictionary")
    
    If mFileXLS Is Nothing Then
        LogWarning MODULE_NAME, "[ReadTechnicalDataFromExcel] mFileXLS es Nothing"
        Exit Function
    End If
    
    Dim wb As Workbook
    Set wb = mFileXLS.wb
    
    ' Aquí iría la lógica específica para leer datos técnicos desde las hojas de Excel
    ' Por ejemplo, buscar hojas específicas que contengan datos técnicos
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "TECNIC", vbTextCompare) > 0 Then
            ' Procesar hoja de datos técnicos
            ReadTechnicalDataFromExcel.Add "TechnicalSheet", ws.Name
            ' Aquí se leerían los datos específicos de la hoja
        End If
    Next ws
    
    LogDebug MODULE_NAME, "[ReadTechnicalDataFromExcel] Datos técnicos leídos"
End Function

'@Description: Lee datos de cotización desde Excel
Private Function ReadQuotationDataFromExcel() As Object
    Set ReadQuotationDataFromExcel = CreateObject("Scripting.Dictionary")
    
    If mFileXLS Is Nothing Then
        LogWarning MODULE_NAME, "[ReadQuotationDataFromExcel] mFileXLS es Nothing"
        Exit Function
    End If
    
    Dim wb As Workbook
    Set wb = mFileXLS.wb
    
    ' Aquí iría la lógica específica para leer datos de cotización desde las hojas de Excel
    ' Por ejemplo, buscar hojas específicas que contengan datos de cotización
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "OFERTA", vbTextCompare) > 0 Or InStr(1, ws.Name, "QUOTE", vbTextCompare) > 0 Then
            ' Procesar hoja de cotización
            ReadQuotationDataFromExcel.Add "QuotationSheet", ws.Name
            ' Aquí se leerían los datos específicos de la hoja
        End If
    Next ws
    
    LogDebug MODULE_NAME, "[ReadQuotationDataFromExcel] Datos de cotización leídos"
End Function

'@Description: Guarda resultados en Excel
Private Sub SaveResultsToExcel(resultado As Object)
    If mFileXLS Is Nothing Then
        LogWarning MODULE_NAME, "[SaveResultsToExcel] mFileXLS es Nothing"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Set wb = mFileXLS.wb
    
    ' Aquí iría la lógica específica para guardar resultados en Excel
    ' Por ejemplo, escribir en hojas específicas o rangos específicos
    LogDebug MODULE_NAME, "[SaveResultsToExcel] Resultados guardados en: " & mFileXLS.Path
End Sub

'@Description: Valida si la fuente de datos es válida
Private Function ValidateExcelSource() As Boolean
    ValidateExcelSource = False
    
    If mFileXLS Is Nothing Then
        LogWarning MODULE_NAME, "[ValidateExcelSource] mFileXLS es Nothing"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' Verificar que el archivo exista y esté accesible
    If mFileXLS.Path <> "(cerrado)" And mFileXLS.Path <> "" Then
        ' Verificar que el workbook esté disponible
        Dim dummy As String
        dummy = mFileXLS.Name  ' Esto causará un error si el workbook no está disponible
        
        ValidateExcelSource = True
        LogDebug MODULE_NAME, "[ValidateExcelSource] Fuente válida: " & mFileXLS.Path
    End If
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[ValidateExcelSource] Error validando fuente", Err.Number, Err.Description
    ValidateExcelSource = False
End Function

'@Description: Obtiene el identificador de la oportunidad
Private Function GetOpportunityIdentifier() As String
    If mFileXLS Is Nothing Then
        GetOpportunityIdentifier = ""
        Exit Function
    End If
    
    GetOpportunityIdentifier = mFileXLS.Path
End Function

' ==========================================
' MÉTODOS PÚBLICOS ADICIONALES
' ==========================================

'@Description: Obtiene el archivo Excel asociado
Public Property Get FileXLS() As clsFileXLS
    Set FileXLS = mFileXLS
End Property

'@Description: Obtiene la ruta del archivo
Public Property Get FilePath() As String
    FilePath = mPath
End Property
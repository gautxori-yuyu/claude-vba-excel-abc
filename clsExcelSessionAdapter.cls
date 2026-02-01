VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelSessionAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' ADAPTADOR DE SESIN DE EXCEL
' ==================================================================
' Implementa la interfaz IExcelSession para permitir que
' el dominio interacte con Excel sin conocer los detalles
' de implementacin concreta de Excel.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores.Adaptadores"
Option Explicit

Private Const MODULE_NAME As String = "clsExcelSessionAdapter"

' Implementa la interfaz
Implements IExcelSession

' Referencia al contexto de ejecucin
Private mExecutionContext As clsExecutionContext

' ==========================================
' INICIALIZACIN
' ==========================================
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
    Set mExecutionContext = Nothing
End Sub

'@Description: Inicializa el adaptador con un contexto de ejecucin
Public Sub Initialize(executionContext As clsExecutionContext)
    Set mExecutionContext = executionContext
    LogInfo MODULE_NAME, "[Initialize] Adaptador inicializado con ExecutionContext"
End Sub

' ==========================================
' IMPLEMENTACIN DE LA INTERFAZ IExcelSession
' ==========================================

Private Property Get IExcelSession_ActiveWorkbook() As Workbook
    Set IExcelSession_ActiveWorkbook = mExecutionContext.Workbook
End Property

Private Property Get IExcelSession_ActiveWorksheet() As Worksheet
    Set IExcelSession_ActiveWorksheet = mExecutionContext.Worksheet
End Property

Private Property Get IExcelSession_ActiveChart() As Chart
    Set IExcelSession_ActiveChart = mExecutionContext.Chart
End Property

Private Property Get IExcelSession_Selection() As Object
    Set IExcelSession_Selection = mExecutionContext.Selection
End Property

Private Sub IExcelSession_DirtyRecalc()
    Call DirtyRecalcWorkbook()
End Sub

Private Sub IExcelSession_EvalUDFs()
    Call EvaluateUDFs()
End Sub

Private Function IExcelSession_HasValidContext() As Boolean
    IExcelSession_HasValidContext = ValidateExcelContext()
End Function

Private Sub IExcelSession_SaveActiveWorkbook()
    Call SaveCurrentWorkbook()
End Sub

Private Sub IExcelSession_CloseActiveWorkbook(saveChanges As Boolean)
    Call CloseCurrentWorkbook(saveChanges)
End Sub

Private Function IExcelSession_OpenWorkbook(filePath As String) As Workbook
    Set IExcelSession_OpenWorkbook = OpenWorkbookAtPath(filePath)
End Function

Private Function IExcelSession_GetWorksheet(wb As Workbook, sheetName As String) As Worksheet
    Set IExcelSession_GetWorksheet = GetWorksheetByName(wb, sheetName)
End Function

Private Function IExcelSession_GetRange(ws As Worksheet, rangeAddress As String) As Range
    Set IExcelSession_GetRange = GetRangeByAddress(ws, rangeAddress)
End Function

Private Function IExcelSession_ReadRangeValues(rng As Range) As Variant
    IExcelSession_ReadRangeValues = ReadValuesFromRange(rng)
End Function

Private Sub IExcelSession_WriteRangeValues(rng As Range, values As Variant)
    Call WriteValuesToRange(rng, values)
End Sub

' ==========================================
' MTODOS PRIVADOS DE IMPLEMENTACIN
' ==========================================

'@Description: Ejecuta un clculo dirty en el workbook
Private Sub DirtyRecalcWorkbook()
    If mExecutionContext Is Nothing Then
        LogWarning MODULE_NAME, "[DirtyRecalcWorkbook] ExecutionContext es Nothing"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Set wb = mExecutionContext.Workbook
    
    If Not wb Is Nothing Then
        Application.CalculateFullRebuild
        LogDebug MODULE_NAME, "[DirtyRecalcWorkbook] Clculo completo ejecutado en: " & wb.Name
    End If
End Sub

'@Description: Evala las UDFs en el contexto actual
Private Sub EvaluateUDFs()
    If mExecutionContext Is Nothing Then
        LogWarning MODULE_NAME, "[EvaluateUDFs] ExecutionContext es Nothing"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Set wb = mExecutionContext.Workbook
    
    If Not wb Is Nothing Then
        ' Forzar recclculo de frmulas
        Application.Calculate
        LogDebug MODULE_NAME, "[EvaluateUDFs] UDFs evaluadas en: " & wb.Name
    End If
End Sub

'@Description: Valida si hay un contexto de Excel vlido
Private Function ValidateExcelContext() As Boolean
    ValidateExcelContext = False
    
    If mExecutionContext Is Nothing Then
        LogWarning MODULE_NAME, "[ValidateExcelContext] ExecutionContext es Nothing"
        Exit Function
    End If
    
    ' Verificar que tengamos acceso a los objetos bsicos
    On Error GoTo ErrorHandler
    
    If Not mExecutionContext.Workbook Is Nothing Then
        ValidateExcelContext = True
        LogDebug MODULE_NAME, "[ValidateExcelContext] Contexto vlido"
    End If
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[ValidateExcelContext] Error", Err.Number, Err.Description
    ValidateExcelContext = False
End Function

'@Description: Guarda el workbook activo
Private Sub SaveCurrentWorkbook()
    If mExecutionContext Is Nothing Then
        LogWarning MODULE_NAME, "[SaveCurrentWorkbook] ExecutionContext es Nothing"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Set wb = mExecutionContext.Workbook
    
    If Not wb Is Nothing Then
        wb.Save
        LogDebug MODULE_NAME, "[SaveCurrentWorkbook] Workbook guardado: " & wb.Name
    End If
End Sub

'@Description: Cierra el workbook activo
Private Sub CloseCurrentWorkbook(saveChanges As Boolean)
    If mExecutionContext Is Nothing Then
        LogWarning MODULE_NAME, "[CloseCurrentWorkbook] ExecutionContext es Nothing"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Set wb = mExecutionContext.Workbook
    
    If Not wb Is Nothing Then
        wb.Close saveChanges
        LogDebug MODULE_NAME, "[CloseCurrentWorkbook] Workbook cerrado: " & wb.Name
    End If
End Sub

'@Description: Abre un workbook desde una ruta
Private Function OpenWorkbookAtPath(filePath As String) As Workbook
    Set OpenWorkbookAtPath = Nothing
    
    On Error GoTo ErrorHandler
    
    Set OpenWorkbookAtPath = Workbooks.Open(filePath)
    LogDebug MODULE_NAME, "[OpenWorkbookAtPath] Workbook abierto: " & OpenWorkbookAtPath.Name
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[OpenWorkbookAtPath] Error abriendo workbook", Err.Number, Err.Description & " - Ruta: " & filePath
    Set OpenWorkbookAtPath = Nothing
End Function

'@Description: Obtiene una hoja especfica por nombre
Private Function GetWorksheetByName(wb As Workbook, sheetName As String) As Worksheet
    Set GetWorksheetByName = Nothing
    
    If wb Is Nothing Then
        LogWarning MODULE_NAME, "[GetWorksheetByName] Workbook es Nothing"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Set GetWorksheetByName = wb.Worksheets(sheetName)
    LogDebug MODULE_NAME, "[GetWorksheetByName] Hoja obtenida: " & sheetName & " en workbook: " & wb.Name
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[GetWorksheetByName] Error obteniendo hoja", Err.Number, Err.Description & " - Hoja: " & sheetName
    Set GetWorksheetByName = Nothing
End Function

'@Description: Obtiene un rango por direccin
Private Function GetRangeByAddress(ws As Worksheet, rangeAddress As String) As Range
    Set GetRangeByAddress = Nothing
    
    If ws Is Nothing Then
        LogWarning MODULE_NAME, "[GetRangeByAddress] Worksheet es Nothing"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Set GetRangeByAddress = ws.Range(rangeAddress)
    LogDebug MODULE_NAME, "[GetRangeByAddress] Rango obtenido: " & rangeAddress & " en hoja: " & ws.Name
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[GetRangeByAddress] Error obteniendo rango", Err.Number, Err.Description & " - Rango: " & rangeAddress
    Set GetRangeByAddress = Nothing
End Function

'@Description: Lee valores de un rango
Private Function ReadValuesFromRange(rng As Range) As Variant
    ReadValuesFromRange = Empty
    
    If rng Is Nothing Then
        LogWarning MODULE_NAME, "[ReadValuesFromRange] Range es Nothing"
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ReadValuesFromRange = rng.Value
    LogDebug MODULE_NAME, "[ReadValuesFromRange] Valores ledos del rango: " & rng.Address
    
    Exit Function
    
ErrorHandler:
    LogError MODULE_NAME, "[ReadValuesFromRange] Error leyendo valores", Err.Number, Err.Description & " - Rango: " & rng.Address
    ReadValuesFromRange = Empty
End Function

'@Description: Escribe valores en un rango
Private Sub WriteValuesToRange(rng As Range, values As Variant)
    If rng Is Nothing Then
        LogWarning MODULE_NAME, "[WriteValuesToRange] Range es Nothing"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    rng.Value = values
    LogDebug MODULE_NAME, "[WriteValuesToRange] Valores escritos en rango: " & rng.Address
    
    Exit Sub
    
ErrorHandler:
    LogError MODULE_NAME, "[WriteValuesToRange] Error escribiendo valores", Err.Number, Err.Description & " - Rango: " & rng.Address
End Sub

' ==========================================
' MTODOS PBLICOS ADICIONALES
' ==========================================

'@Description: Obtiene el contexto de ejecucin asociado
Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = mExecutionContext
End Property
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExcelSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==================================================================
' INTERFAZ PARA SESIÓN DE EXCEL
' ==================================================================
' Define el contrato que deben cumplir todas las clases que provean
' acceso a operaciones de Excel, permitiendo que el dominio interactúe
' con Excel sin conocer la implementación concreta.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores.Interfaces"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IExcelSession"

'@Description: Obtiene el workbook activo
'@Returns: Workbook - El workbook activo en la sesión
Public Property Get ActiveWorkbook() As Workbook
    ' Las interfaces en VBA deben lanzar error o no hacer nada
    Err.Raise 5, "IExcelSession", "ActiveWorkbook debe implementarse en clase concreta"
End Property

'@Description: Obtiene la hoja activa
'@Returns: Worksheet - La hoja activa en la sesión
Public Property Get ActiveWorksheet() As Worksheet
    Err.Raise 5, "IExcelSession", "ActiveWorksheet debe implementarse en clase concreta"
End Property

'@Description: Obtiene el gráfico activo
'@Returns: Chart - El gráfico activo en la sesión
Public Property Get ActiveChart() As Chart
    Err.Raise 5, "IExcelSession", "ActiveChart debe implementarse en clase concreta"
End Property

'@Description: Obtiene la selección actual
'@Returns: Object - La selección actual en la sesión
Public Property Get Selection() As Object
    Err.Raise 5, "IExcelSession", "Selection debe implementarse en clase concreta"
End Property

'@Description: Ejecuta un cálculo dirty en el workbook
Public Sub DirtyRecalc()
    Err.Raise 5, "IExcelSession", "DirtyRecalc debe implementarse en clase concreta"
End Sub

'@Description: Evalúa las UDFs en el contexto actual
Public Sub EvalUDFs()
    Err.Raise 5, "IExcelSession", "EvalUDFs debe implementarse en clase concreta"
End Sub

'@Description: Verifica si hay un contexto de Excel válido
'@Returns: Boolean - True si hay un contexto válido
Public Function HasValidContext() As Boolean
    Err.Raise 5, "IExcelSession", "HasValidContext debe implementarse en clase concreta"
End Function

'@Description: Guarda el workbook activo
Public Sub SaveActiveWorkbook()
    Err.Raise 5, "IExcelSession", "SaveActiveWorkbook debe implementarse en clase concreta"
End Sub

'@Description: Cierra el workbook activo
'@Param: saveChanges - Si se deben guardar los cambios antes de cerrar
Public Sub CloseActiveWorkbook(saveChanges As Boolean)
    Err.Raise 5, "IExcelSession", "CloseActiveWorkbook debe implementarse en clase concreta"
End Sub

'@Description: Abre un workbook desde una ruta
'@Param: filePath - Ruta del archivo a abrir
'@Returns: Workbook - El workbook abierto
Public Function OpenWorkbook(filePath As String) As Workbook
    Err.Raise 5, "IExcelSession", "OpenWorkbook debe implementarse en clase concreta"
End Function

'@Description: Obtiene una referencia a una hoja específica
'@Param: workbook - El workbook que contiene la hoja
'@Param: sheetName - Nombre de la hoja
'@Returns: Worksheet - La hoja solicitada
Public Function GetWorksheet(Workbook As Workbook, sheetName As String) As Worksheet
    Err.Raise 5, "IExcelSession", "GetWorksheet debe implementarse en clase concreta"
End Function

'@Description: Obtiene un rango específico
'@Param: worksheet - La hoja que contiene el rango
'@Param: rangeAddress - Dirección del rango (ej. "A1:B10")
'@Returns: Range - El rango solicitado
Public Function GetRange(Worksheet As Worksheet, rangeAddress As String) As range
    Err.Raise 5, "IExcelSession", "GetRange debe implementarse en clase concreta"
End Function

'@Description: Lee valores de un rango
'@Param: range - El rango del cual leer valores
'@Returns: Variant - Valores del rango
Public Function ReadRangeValues(range As range) As Variant
    Err.Raise 5, "IExcelSession", "ReadRangeValues debe implementarse en clase concreta"
End Function

'@Description: Escribe valores en un rango
'@Param: range - El rango donde escribir valores
'@Param: values - Valores a escribir
Public Sub WriteRangeValues(range As range, values As Variant)
    Err.Raise 5, "IExcelSession", "WriteRangeValues debe implementarse en clase concreta"
End Sub

' ==================================================================
' NOTAS DE IMPLEMENTACIÓN
' ==================================================================
' Para implementar esta interfaz en una clase:
'
' 1. Agregar al inicio de la clase:
'    Implements IExcelSession
'
' 2. VBA generará automáticoamente los stubs con prefijo "IExcelSession_"
'
' 3. Implementar cada método delegando a la lógica específica de la clase
'
' 4. Usar esta interfaz para que el dominio no dependa de implementaciones concretas de Excel

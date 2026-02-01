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
' INTERFAZ PARA SESI�N DE EXCEL
' ==================================================================
' Define el contrato que deben cumplir todas las clases que provean
' acceso a operaciones de Excel, permitiendo que el dominio interact�e
' con Excel sin conocer la implementaci�n concreta.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores.Interfaces"
'@Interface
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "IExcelSession"

'@Description: Obtiene el workbook activo
'@Returns: Workbook - El workbook activo en la sesi�n
Public Property Get ActiveWorkbook() As Workbook
    ' Las interfaces en VBA deben lanzar error o no hacer nada
    Err.Raise 5, "IExcelSession", "ActiveWorkbook debe implementarse en clase concreta"
End Property

'@Description: Obtiene la hoja activa
'@Returns: Worksheet - La hoja activa en la sesi�n
Public Property Get ActiveWorksheet() As Worksheet
    Err.Raise 5, "IExcelSession", "ActiveWorksheet debe implementarse en clase concreta"
End Property

'@Description: Obtiene el gr�fico activo
'@Returns: Chart - El gr�fico activo en la sesi�n
Public Property Get ActiveChart() As Chart
    Err.Raise 5, "IExcelSession", "ActiveChart debe implementarse en clase concreta"
End Property

'@Description: Obtiene la selecci�n actual
'@Returns: Object - La selecci�n actual en la sesi�n
Public Property Get Selection() As Object
    Err.Raise 5, "IExcelSession", "Selection debe implementarse en clase concreta"
End Property

'@Description: Ejecuta un c�lculo dirty en el workbook
Public Sub DirtyRecalc()
    Err.Raise 5, "IExcelSession", "DirtyRecalc debe implementarse en clase concreta"
End Sub

'@Description: Eval�a las UDFs en el contexto actual
Public Sub EvalUDFs()
    Err.Raise 5, "IExcelSession", "EvalUDFs debe implementarse en clase concreta"
End Sub

'@Description: Verifica si hay un contexto de Excel v�lido
'@Returns: Boolean - True si hay un contexto v�lido
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

'@Description: Obtiene una referencia a una hoja espec�fica
'@Param: workbook - El workbook que contiene la hoja
'@Param: sheetName - Nombre de la hoja
'@Returns: Worksheet - La hoja solicitada
Public Function GetWorksheet(Workbook As Workbook, sheetName As String) As Worksheet
    Err.Raise 5, "IExcelSession", "GetWorksheet debe implementarse en clase concreta"
End Function

'@Description: Obtiene un rango espec�fico
'@Param: worksheet - La hoja que contiene el rango
'@Param: rangeAddress - Direcci�n del rango (ej. "A1:B10")
'@Returns: Range - El rango solicitado
Public Function GetRange(Worksheet As Worksheet, rangeAddress As String) As Range
    Err.Raise 5, "IExcelSession", "GetRange debe implementarse en clase concreta"
End Function

'@Description: Lee valores de un rango
'@Param: range - El rango del cual leer valores
'@Returns: Variant - Valores del rango
Public Function ReadRangeValues(range As Range) As Variant
    Err.Raise 5, "IExcelSession", "ReadRangeValues debe implementarse en clase concreta"
End Function

'@Description: Escribe valores en un rango
'@Param: range - El rango donde escribir valores
'@Param: values - Valores a escribir
Public Sub WriteRangeValues(range As Range, values As Variant)
    Err.Raise 5, "IExcelSession", "WriteRangeValues debe implementarse en clase concreta"
End Sub

' ==================================================================
' NOTAS DE IMPLEMENTACI�N
' ==================================================================
' Para implementar esta interfaz en una clase:
'
' 1. Agregar al inicio de la clase:
'    Implements IExcelSession
'
' 2. VBA generar� autom�ticoamente los stubs con prefijo "IExcelSession_"
'
' 3. Implementar cada m�todo delegando a la l�gica espec�fica de la clase
'
' 4. Usar esta interfaz para que el dominio no dependa de implementaciones concretas de Excel

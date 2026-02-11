VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITechnicalCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' INTERFAZ DE DOMINIO: ITechnicalCalc
' ==================================================================
' Define el contrato para trabajar con UN calculo tecnico.
' El dominio usa esta interfaz sin conocer la implementacion concreta
' (ficheros Excel, base de datos, etc.)
'
' Basado en la estructura identificada en VBScript:
' - CalcNumber: Codigo del calculo (ej: "ABC12345")
' - CalcOption: Opcion del calculo (ej: "ABC12345_01")
' - Revision: Numero de revision (opcional)
' - Modelo de compresor resultante
'
' Implementada por: ExcelTechnicalCalcSource (infraestructura)
' ==================================================================

'@Folder "3-Dominio"
Option Explicit

' ==================================================================
' PROPIEDADES DE IDENTIFICACION
' ==================================================================

'@Description: Codigo del calculo (3 letras + 5 digitos, ej: "ABC12345")
Public Property Get calcNumber() As String
End Property

'@Description: Opcion del calculo (CalcNumber + "_" + 2 digitos, ej: "ABC12345_01")
Public Property Get calcOption() As String
End Property

'@Description: Numero de revision del calculo (0 si es original)
Public Property Get Revision() As Long
End Property

'@Description: Ruta completa del fichero principal del calculo
Public Property Get FilePath() As String
End Property

'@Description: Nombre del fichero (sin ruta)
Public Property Get fileName() As String
End Property

' ==================================================================
' PROPIEDADES DE RESULTADO
' ==================================================================

'@Description: Modelo de compresor resultante del calculo (ej: "2TEH-1-LG")
Public Property Get CompressorModel() As String
End Property

'@Description: Indica si el calculo tiene datos validos
Public Property Get IsValid() As Boolean
End Property

'@Description: Indica si el calculo ha sido procesado
Public Property Get IsProcessed() As Boolean
End Property

' ==================================================================
' PROPIEDADES DE DATOS TECNICOS
' ==================================================================

'@Description: Obtiene los datos del gas del calculo
'@Returns: Dictionary con propiedades del gas (composicion, temperatura, presion, etc.)
Public Function GetGasData() As Object
End Function

'@Description: Obtiene los datos de proceso del calculo
'@Returns: Dictionary con condiciones de proceso
Public Function GetProcessData() As Object
End Function

'@Description: Obtiene los datos del compresor resultante
'@Returns: Dictionary con especificaciones del compresor
Public Function GetCompressorData() As Object
End Function

' ==================================================================
' FICHEROS ASOCIADOS
' ==================================================================

'@Description: Obtiene la lista de ficheros asociados a este calculo
'@Returns: Collection de rutas de ficheros (.xlsx, .txt, .rtf, etc.)
Public Function GetAssociatedFiles() As Collection
End Function

'@Description: Indica si tiene fichero de antipulsadores
Public Property Get HasAntipulsadorFile() As Boolean
End Property

'@Description: Indica si tiene ficheros de componentes auxiliares (Gas_Cooler, Aircooler, etc.)
Public Property Get HasAuxiliaryFiles() As Boolean
End Property

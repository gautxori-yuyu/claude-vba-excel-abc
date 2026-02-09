VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITechnicalCalcProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' INTERFAZ DE DOMINIO: ITechnicalCalcProvider
' ==================================================================
' Define el contrato para obtener/buscar calculos tecnicos.
' El dominio usa esta interfaz sin conocer como se almacenan los calculos.
'
' Basado en la estructura de carpetas:
' [Oportunidad]/2.CALCULO TECNICO/
'   +-- ABC12345_01_calc.xlsx
'   +-- ABC12345_01.txt
'   +-- ABC12345_02_calc.xlsx
'   +-- Antipul_ABC12345_01.xlsx
'   +-- ABC_Gas_Cooler-ABC12345_01.xlsx
'
' Implementada por: FileSystemTechnicalCalcProvider (infraestructura)
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores"
Option Explicit

' ==================================================================
' CONSTANTES DE SUBCARPETA
' ==================================================================
' La carpeta de calculos tecnicos dentro de una oportunidad
' Public Const TECHNICAL_CALC_FOLDER As String = "2.CALCULO TECNICO"

' ==================================================================
' METODOS DE BUSQUEDA
' ==================================================================

'@Description: Obtiene todos los calculos tecnicos de una oportunidad
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de ITechnicalCalc
Public Function GetCalculations(ByVal opportunityPath As String) As Collection
End Function

'@Description: Obtiene un calculo por su numero
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Param calcNumber: Codigo del calculo (ej: "ABC12345")
'@Returns: Collection de ITechnicalCalc (puede haber varias opciones)
Public Function GetCalculationsByNumber(ByVal opportunityPath As String, ByVal calcNumber As String) As Collection
End Function

'@Description: Obtiene un calculo especifico por opcion
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Param calcOption: Opcion del calculo (ej: "ABC12345_01")
'@Returns: ITechnicalCalc o Nothing si no existe
Public Function GetCalculationByOption(ByVal opportunityPath As String, ByVal calcOption As String) As ITechnicalCalc
End Function

'@Description: Obtiene calculos agrupados por modelo de compresor
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Dictionary {modelName -> Collection of ITechnicalCalc}
Public Function GetCalculationsByCompressorModel(ByVal opportunityPath As String) As Object
End Function

' ==================================================================
' METODOS DE IDENTIFICACION
' ==================================================================

'@Description: Identifica los calculos tecnicos en una carpeta (sin procesar contenido)
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de nombres de ficheros de calculo encontrados
Public Function IdentifyCalculationFiles(ByVal opportunityPath As String) As Collection
End Function

'@Description: Verifica si un fichero es un calculo tecnico valido
'@Param fileName: Nombre del fichero a verificar
'@Returns: True si cumple el patron de nombre de calculo
Public Function IsValidCalculationFile(ByVal fileName As String) As Boolean
End Function

'@Description: Extrae informacion del nombre de un fichero de calculo
'@Param fileName: Nombre del fichero
'@Returns: Dictionary con CalcNumber, CalcOption, Revision, FileType
Public Function ParseCalculationFileName(ByVal fileName As String) As Object
End Function

' ==================================================================
' METODOS DE GESTION
' ==================================================================

'@Description: Mueve ficheros de calculo desde carpeta temporal a oportunidad
'@Param sourcePath: Ruta origen (carpeta de calculos temporal)
'@Param opportunityPath: Ruta de la oportunidad destino
'@Param calcNumber: Codigo del calculo a mover (opcional, vacio = todos los relacionados)
'@Returns: Numero de ficheros movidos
Public Function MoveCalculationsToOpportunity(ByVal sourcePath As String, _
                                               ByVal opportunityPath As String, _
                                               Optional ByVal calcNumber As String = "") As Long
End Function

'@Description: Obtiene la ruta de la carpeta de calculos de una oportunidad
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Ruta completa de la subcarpeta de calculos tecnicos
Public Function GetCalculationsFolder(ByVal opportunityPath As String) As String
End Function

'@Description: Verifica si existe la carpeta de calculos tecnicos
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: True si la subcarpeta existe
Public Function CalculationsFolderExists(ByVal opportunityPath As String) As Boolean
End Function

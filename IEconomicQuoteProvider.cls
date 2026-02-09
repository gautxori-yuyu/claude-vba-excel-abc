VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEconomicQuoteProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' INTERFAZ DE DOMINIO: IEconomicQuoteProvider
' ==================================================================
' Define el contrato para obtener/buscar valoraciones economicas.
' El dominio usa esta interfaz sin conocer como se almacenan.
'
' Basado en la estructura de carpetas:
' [Oportunidad]/3.VALORACION ECONOMICA/
'   +-- rev 1/
'   |   +-- 123456789-01.xlsx
'   |   +-- 123456789-02.xlsx
'   +-- rev 2/
'   |   +-- 123456789-01 rev.2.xlsx
'   |   +-- 123456789-02 rev.2.xlsx
'   +-- plantilla de referencia.xlsm
'
' Implementada por: FileSystemEconomicQuoteProvider (infraestructura)
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores"
Option Explicit

' ==================================================================
' METODOS DE BUSQUEDA
' ==================================================================

'@Description: Obtiene todas las valoraciones economicas de una oportunidad
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de IEconomicQuote
Public Function GetQuotes(ByVal opportunityPath As String) As Collection
End Function

'@Description: Obtiene valoraciones por numero de oferta
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Param quoteNumber: Numero de oferta (9 digitos)
'@Returns: Collection de IEconomicQuote (puede haber varios compresores)
Public Function GetQuotesByNumber(ByVal opportunityPath As String, ByVal quoteNumber As String) As Collection
End Function

'@Description: Obtiene una valoracion especifica
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Param quoteId: Identificador completo (ej: "123456789-01")
'@Returns: IEconomicQuote o Nothing si no existe
Public Function GetQuoteById(ByVal opportunityPath As String, ByVal quoteId As String) As IEconomicQuote
End Function

'@Description: Obtiene valoraciones agrupadas por revision
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Dictionary {revisionNumber -> Collection of IEconomicQuote}
Public Function GetQuotesByRevision(ByVal opportunityPath As String) As Object
End Function

'@Description: Obtiene la ultima revision de cada valoracion
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de IEconomicQuote (solo las mas recientes)
Public Function GetLatestQuotes(ByVal opportunityPath As String) As Collection
End Function

' ==================================================================
' METODOS DE IDENTIFICACION
' ==================================================================

'@Description: Identifica los ficheros de valoracion en una carpeta
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de nombres de ficheros encontrados
Public Function IdentifyQuoteFiles(ByVal opportunityPath As String) As Collection
End Function

'@Description: Verifica si un fichero es una valoracion economica valida
'@Param fileName: Nombre del fichero a verificar
'@Returns: True si cumple el patron de nombre de valoracion
Public Function IsValidQuoteFile(ByVal fileName As String) As Boolean
End Function

'@Description: Extrae informacion del nombre de un fichero de valoracion
'@Param fileName: Nombre del fichero
'@Returns: Dictionary con QuoteNumber, CompressorIndex, Revision
Public Function ParseQuoteFileName(ByVal fileName As String) As Object
End Function

' ==================================================================
' METODOS DE GESTION
' ==================================================================

'@Description: Obtiene la ruta de la carpeta de valoraciones de una oportunidad
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Ruta completa de la subcarpeta de valoraciones
Public Function GetQuotesFolder(ByVal opportunityPath As String) As String
End Function

'@Description: Verifica si existe la carpeta de valoraciones
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: True si la subcarpeta existe
Public Function QuotesFolderExists(ByVal opportunityPath As String) As Boolean
End Function

'@Description: Obtiene las subcarpetas de revision existentes
'@Param opportunityPath: Ruta de la carpeta de oportunidad
'@Returns: Collection de numeros de revision encontrados (ordenados)
Public Function GetRevisionFolders(ByVal opportunityPath As String) As Collection
End Function

'@Description: Asocia una valoracion con un calculo tecnico
'@Param quotePath: Ruta del fichero de valoracion
'@Param calcOption: Opcion del calculo a asociar (ej: "ABC12345_01")
'@Returns: True si la asociacion fue exitosa
Public Function AssociateWithCalculation(ByVal quotePath As String, ByVal calcOption As String) As Boolean
End Function

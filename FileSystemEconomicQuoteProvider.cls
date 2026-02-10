VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemEconomicQuoteProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: FileSystemEconomicQuoteProvider
' ==================================================================
' Implementa IEconomicQuoteProvider usando el sistema de archivos.
' Contiene toda la logica de:
' - Busqueda de ficheros de valoracion en carpeta de oportunidad
' - Identificacion por patrones regex
' - Parseo de nombres de fichero
' - Gestion de subcarpetas de revision
'
' Patron de VBScript:
' strQuoteNrPattern = "\d{9}(?:[\-_]\d+)?"
' strQuoteNrRevPattern = "(" & strQuoteNrPattern & ")(?:[ \-_]*rev\.?[ \-_]*\d+\b)?"
' ==================================================================

'@Folder "2-Infraestructura.Providers"
Option Explicit
Implements IEconomicQuoteProvider

Private Const MODULE_NAME As String = "FileSystemEconomicQuoteProvider"

' Nombre de la subcarpeta de valoraciones economicas
Private Const QUOTE_FOLDER_NAME As String = "3.VALORACION ECONOMICA"

' Patrones de expresion regular
Private Const QUOTE_NR_PATTERN As String = "\d{9}(?:[\-_]\d+)?"
Private Const QUOTE_FILE_PATTERN As String = "^(\d{9})(?:[\-_](\d+))?(?:[ \-_]*rev\.?[ \-_]*(\d+))?.*\.xls[xm]$"
Private Const REV_FOLDER_PATTERN As String = "^rev\.?\s*(\d+)$"

' ==================================================================
' DEPENDENCIAS
' ==================================================================
Private mFileManager As clsFileManager

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Set mFileManager = Nothing
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'@Description: Inyecta la dependencia de FileManager
Public Sub Initialize(fm As clsFileManager)
    Set mFileManager = fm
    LogDebug MODULE_NAME, "[Initialize] FileManager inyectado"
End Sub

' ==================================================================
' IMPLEMENTACION DE IEconomicQuoteProvider
' ==================================================================

Private Function IEconomicQuoteProvider_GetQuotes(ByVal opportunityPath As String) As Collection
    Set IEconomicQuoteProvider_GetQuotes = GetQuotes(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_GetQuotesByNumber(ByVal opportunityPath As String, ByVal quoteNumber As String) As Collection
    Set IEconomicQuoteProvider_GetQuotesByNumber = GetQuotesByNumber(opportunityPath, quoteNumber)
End Function

Private Function IEconomicQuoteProvider_GetQuoteById(ByVal opportunityPath As String, ByVal quoteId As String) As IEconomicQuote
    Set IEconomicQuoteProvider_GetQuoteById = GetQuoteById(opportunityPath, quoteId)
End Function

Private Function IEconomicQuoteProvider_GetQuotesByRevision(ByVal opportunityPath As String) As Object
    Set IEconomicQuoteProvider_GetQuotesByRevision = GetQuotesByRevision(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_GetLatestQuotes(ByVal opportunityPath As String) As Collection
    Set IEconomicQuoteProvider_GetLatestQuotes = GetLatestQuotes(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_IdentifyQuoteFiles(ByVal opportunityPath As String) As Collection
    Set IEconomicQuoteProvider_IdentifyQuoteFiles = IdentifyQuoteFiles(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_IsValidQuoteFile(ByVal fileName As String) As Boolean
    IEconomicQuoteProvider_IsValidQuoteFile = IsValidQuoteFile(fileName)
End Function

Private Function IEconomicQuoteProvider_ParseQuoteFileName(ByVal fileName As String) As Object
    Set IEconomicQuoteProvider_ParseQuoteFileName = ParseQuoteFileName(fileName)
End Function

Private Function IEconomicQuoteProvider_GetQuotesFolder(ByVal opportunityPath As String) As String
    IEconomicQuoteProvider_GetQuotesFolder = GetQuotesFolder(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_QuotesFolderExists(ByVal opportunityPath As String) As Boolean
    IEconomicQuoteProvider_QuotesFolderExists = QuotesFolderExists(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_GetRevisionFolders(ByVal opportunityPath As String) As Collection
    Set IEconomicQuoteProvider_GetRevisionFolders = GetRevisionFolders(opportunityPath)
End Function

Private Function IEconomicQuoteProvider_AssociateWithCalculation(ByVal quotePath As String, ByVal calcOption As String) As Boolean
    IEconomicQuoteProvider_AssociateWithCalculation = AssociateWithCalculation(quotePath, calcOption)
End Function

' ==================================================================
' METODOS PUBLICOS
' ==================================================================

'@Description: Obtiene todas las valoraciones de una oportunidad
Public Function GetQuotes(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim quoteFolder As String
    quoteFolder = GetQuotesFolder(opportunityPath)

    If Not mFileManager.ExisteCarpeta(quoteFolder) Then
        LogWarning MODULE_NAME, "[GetQuotes] Carpeta no existe: " & quoteFolder
        Set GetQuotes = result
        Exit Function
    End If

    ' Obtener ficheros de valoracion (raiz y subcarpetas de revision)
    Dim files As Collection
    Set files = IdentifyQuoteFiles(opportunityPath)

    ' Crear objetos IEconomicQuote para cada fichero
    Dim filePath As Variant
    Dim parsed As Object
    Dim quoteSource As ExcelEconomicQuoteSource
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each filePath In files
        Set parsed = ParseQuoteFileName(fso.GetFileName(CStr(filePath)))
        If Not parsed Is Nothing Then
            ' Crear ExcelEconomicQuoteSource
            Set quoteSource = New ExcelEconomicQuoteSource
            quoteSource.Initialize CStr(filePath), parsed

            ' Agregar a la coleccion usando QuoteId + revision como key
            Dim key As String
            key = parsed("QuoteId") & "_rev" & parsed("Revision")

            On Error Resume Next
            result.Add quoteSource, key
            If Err.Number <> 0 Then
                ' Ya existe, ignorar duplicado
                LogDebug MODULE_NAME, "[GetQuotes] Valoracion duplicada: " & key
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
    Next filePath

    LogDebug MODULE_NAME, "[GetQuotes] Encontradas " & result.Count & " valoraciones"
    Set GetQuotes = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetQuotes]"
    Set GetQuotes = result
End Function

'@Description: Obtiene valoraciones por numero de oferta
Public Function GetQuotesByNumber(ByVal opportunityPath As String, ByVal quoteNumber As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim allQuotes As Collection
    Set allQuotes = GetQuotes(opportunityPath)

    Dim quote As IEconomicQuote
    For Each quote In allQuotes
        If quote.QuoteNumber = quoteNumber Then
            result.Add quote
        End If
    Next quote

    Set GetQuotesByNumber = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetQuotesByNumber]"
    Set GetQuotesByNumber = result
End Function

'@Description: Obtiene una valoracion por ID
Public Function GetQuoteById(ByVal opportunityPath As String, ByVal quoteId As String) As IEconomicQuote
    On Error GoTo ErrHandler

    Dim allQuotes As Collection
    Set allQuotes = GetQuotes(opportunityPath)

    Dim quote As IEconomicQuote
    For Each quote In allQuotes
        If quote.QuoteId = quoteId Then
            Set GetQuoteById = quote
            Exit Function
        End If
    Next quote

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetQuoteById]"
End Function

'@Description: Agrupa valoraciones por revision
Public Function GetQuotesByRevision(ByVal opportunityPath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    Dim allQuotes As Collection
    Set allQuotes = GetQuotes(opportunityPath)

    Dim quote As IEconomicQuote
    Dim rev As Long
    For Each quote In allQuotes
        rev = quote.Revision
        If Not result.Exists(rev) Then
            result.Add rev, New Collection
        End If
        result(rev).Add quote
    Next quote

    Set GetQuotesByRevision = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetQuotesByRevision]"
    Set GetQuotesByRevision = result
End Function

'@Description: Obtiene las valoraciones mas recientes
Public Function GetLatestQuotes(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim byQuoteId As Object
    Set byQuoteId = CreateObject("Scripting.Dictionary")

    Dim allQuotes As Collection
    Set allQuotes = GetQuotes(opportunityPath)

    ' Agrupar por QuoteId y quedarse con la mayor revision
    Dim quote As IEconomicQuote
    For Each quote In allQuotes
        If Not byQuoteId.Exists(quote.QuoteId) Then
            byQuoteId.Add quote.QuoteId, quote
        Else
            If quote.Revision > byQuoteId(quote.QuoteId).Revision Then
                Set byQuoteId(quote.QuoteId) = quote
            End If
        End If
    Next quote

    ' Convertir a Collection
    Dim key As Variant
    For Each key In byQuoteId.Keys
        result.Add byQuoteId(key)
    Next key

    Set GetLatestQuotes = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetLatestQuotes]"
    Set GetLatestQuotes = result
End Function

'@Description: Identifica ficheros de valoracion
Public Function IdentifyQuoteFiles(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim quoteFolder As String
    quoteFolder = GetQuotesFolder(opportunityPath)

    If Not mFileManager.ExisteCarpeta(quoteFolder) Then
        Set IdentifyQuoteFiles = result
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(quoteFolder) Then
        Set IdentifyQuoteFiles = result
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(quoteFolder)

    ' Buscar en carpeta raiz
    Dim file As Object
    For Each file In folder.Files
        If IsValidQuoteFile(file.Name) Then
            result.Add file.Path
        End If
    Next file

    ' Buscar en subcarpetas de revision
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        If IsRevisionFolder(subfolder.Name) Then
            For Each file In subfolder.Files
                If IsValidQuoteFile(file.Name) Then
                    result.Add file.Path
                End If
            Next file
        End If
    Next subfolder

    LogDebug MODULE_NAME, "[IdentifyQuoteFiles] Encontrados " & result.Count & " ficheros"

    Set IdentifyQuoteFiles = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[IdentifyQuoteFiles]"
    Set IdentifyQuoteFiles = result
End Function

'@Description: Verifica si un fichero es una valoracion valida
Public Function IsValidQuoteFile(ByVal fileName As String) As Boolean
    On Error GoTo ErrHandler

    ' Ignorar ficheros temporales
    If Left(fileName, 2) = "~$" Then Exit Function

    ' Ignorar plantillas
    If InStr(1, fileName, "plantilla", vbTextCompare) > 0 Then Exit Function

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = QUOTE_FILE_PATTERN

    IsValidQuoteFile = regEx.Test(fileName)
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[IsValidQuoteFile]"
End Function

'@Description: Verifica si una carpeta es de revision
Private Function IsRevisionFolder(ByVal folderName As String) As Boolean
    On Error GoTo ErrHandler

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = REV_FOLDER_PATTERN

    IsRevisionFolder = regEx.Test(folderName)
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[IsRevisionFolder]"
End Function

'@Description: Parsea el nombre de un fichero de valoracion
Public Function ParseQuoteFileName(ByVal fileName As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = QUOTE_FILE_PATTERN

    If Not regEx.Test(fileName) Then
        Set ParseQuoteFileName = Nothing
        Exit Function
    End If

    Dim matches As Object
    Set matches = regEx.Execute(fileName)

    If matches.Count > 0 Then
        Dim m As Object
        Set m = matches(0)

        result.Add "QuoteNumber", m.SubMatches(0)  ' 123456789
        result.Add "CompressorIndex", IIf(Len(m.SubMatches(1)) > 0, m.SubMatches(1), "01")
        result.Add "Revision", IIf(Len(m.SubMatches(2)) > 0, CLng(m.SubMatches(2)), 0)
        result.Add "QuoteId", m.SubMatches(0) & IIf(Len(m.SubMatches(1)) > 0, "-" & m.SubMatches(1), "")

        Set ParseQuoteFileName = result
    Else
        Set ParseQuoteFileName = Nothing
    End If
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[ParseQuoteFileName]"
    Set ParseQuoteFileName = Nothing
End Function

'@Description: Obtiene la ruta de la carpeta de valoraciones
Public Function GetQuotesFolder(ByVal opportunityPath As String) As String
    GetQuotesFolder = opportunityPath & "\" & QUOTE_FOLDER_NAME
End Function

'@Description: Verifica si existe la carpeta de valoraciones
Public Function QuotesFolderExists(ByVal opportunityPath As String) As Boolean
    If mFileManager Is Nothing Then Exit Function
    QuotesFolderExists = mFileManager.ExisteCarpeta(GetQuotesFolder(opportunityPath))
End Function

'@Description: Obtiene las subcarpetas de revision
Public Function GetRevisionFolders(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim quoteFolder As String
    quoteFolder = GetQuotesFolder(opportunityPath)

    If Not mFileManager.ExisteCarpeta(quoteFolder) Then
        Set GetRevisionFolders = result
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(quoteFolder)

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = REV_FOLDER_PATTERN

    Dim subfolder As Object
    Dim revNum As Long
    For Each subfolder In folder.SubFolders
        If regEx.Test(subfolder.Name) Then
            revNum = CLng(regEx.Execute(subfolder.Name)(0).SubMatches(0))
            result.Add revNum
        End If
    Next subfolder

    ' TODO: Ordenar la coleccion

    Set GetRevisionFolders = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetRevisionFolders]"
    Set GetRevisionFolders = result
End Function

'@Description: Asocia una valoracion con un calculo
Public Function AssociateWithCalculation(ByVal quotePath As String, ByVal calcOption As String) As Boolean
    On Error GoTo ErrHandler

    ' TODO: Implementar logica de asociacion
    LogWarning MODULE_NAME, "[AssociateWithCalculation] No implementado aun"

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[AssociateWithCalculation]"
End Function

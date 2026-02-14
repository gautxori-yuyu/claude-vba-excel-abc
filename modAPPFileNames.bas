Attribute VB_Name = "modAPPFileNames"
'@IgnoreModule MissingAnnotationArgument
'@Folder "3-Dominio"
Option Explicit

Private Const MODULE_NAME As String = "modAPPFileNames"

Private Enum FNTag
    tCustomer
    tQuoteNr
    tQuoteRev
    tModel
    tFamily
    tCylinders
    tStages
End Enum

' Regex reutilizable a nivel de modulo (inicializada una sola vez, Late Binding)
Private mRegEx As Object

Private Function GetRegEx() As Object
    If mRegEx Is Nothing Then Set mRegEx = CreateObject("VBScript.RegExp")
    Set GetRegEx = mRegEx
End Function

'@Description: Devuelve el nombre del archivo actual (con extensión)
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Private Function getFileNameTag(tag As FNTag, fileName As String) As String
    Dim regEx As Object
    Dim matches As Object, sm As Integer

    On Error GoTo ErrorHandler

    Set regEx = GetRegEx()
    regEx.IgnoreCase = True
    
    Select Case tag
    Case tCustomer:
        regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN

            
        If regEx.Test(fileName) Then
            Set matches = regEx.Execute(fileName)
            getFileNameTag = matches(0).SubMatches(1)
        Else
            getFileNameTag = ""
        End If
    Case tQuoteNr:
        regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
            
        If regEx.Test(fileName) Then
            Set matches = regEx.Execute(fileName)
            getFileNameTag = matches(0).SubMatches(0)
        Else
            getFileNameTag = ""
        End If
    Case tQuoteRev:
        regEx.Pattern = QUOTENR_REV_PATTERN
            
        If regEx.Test(fileName) Then
            Set matches = regEx.Execute(fileName)
            getFileNameTag = matches(0).SubMatches(1)
        Else
            getFileNameTag = ""
        End If
    Case tModel, tFamily, tStages, tCylinders:
        regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
            
        If regEx.Test(fileName) Then
            Set matches = regEx.Execute(fileName)
            getFileNameTag = matches(0).SubMatches(3)
        Else
            getFileNameTag = ""
        End If
    Case Else:
        GoTo ErrorHandler
    End Select
    regEx.Pattern = MODEL_PATTERN
    Select Case tag
    Case tFamily: sm = 1
    Case tCylinders: sm = 2
    Case tStages: sm = 0
    End Select
    If regEx.Test(getFileNameTag) And tag > tFamily Then
        Set matches = regEx.Execute(getFileNameTag)
        getFileNameTag = matches(0).SubMatches(sm)
    End If
      
    Exit Function
ErrorHandler:
    getFileNameTag = "#ERROR: " & Err.Description
End Function

'@UDF
'@Description: Extrae el cliente del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions:
Public Function CustomerInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute CustomerInFileName.VB_Description = "[modAPPFileNames] Extrae el cliente del nombre de archivo, del workbook actual o el pasado como parametro"
Attribute CustomerInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    CustomerInFileName = getFileNameTag(tCustomer, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    CustomerInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de oferta del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions:
Public Function QuoteNrInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute QuoteNrInFileName.VB_Description = "[modAPPFileNames] Extrae el número de oferta del nombre de archivo, del workbook actual o el pasado como parametro"
Attribute QuoteNrInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    QuoteNrInFileName = getFileNameTag(tQuoteNr, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    QuoteNrInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de revisión de la oferta del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function QuoteRevInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute QuoteRevInFileName.VB_Description = "[modAPPFileNames] Extrae el número de revisión de la oferta del nombre de archivo, del workbook actual o el pasado como parametro"
Attribute QuoteRevInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    QuoteRevInFileName = getFileNameTag(tQuoteRev, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    QuoteRevInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae el modelo del compresor del nombre del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function ModelInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute ModelInFileName.VB_Description = "[modAPPFileNames] Extrae el modelo del compresor del nombre del nombre de archivo, del workbook actual o el pasado como parametro"
Attribute ModelInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    ModelInFileName = getFileNameTag(tModel, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    ModelInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae la familia del compresor (HA, HG, HP, HX) del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function FamilyInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute FamilyInFileName.VB_Description = "[modAPPFileNames] Extrae la familia del compresor (HA, HG, HP, HX) del modelo"
Attribute FamilyInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    FamilyInFileName = getFileNameTag(tFamily, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    FamilyInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de etapas del compresor del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function StagesInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute StagesInFileName.VB_Description = "[modAPPFileNames] Extrae el número de etapas del compresor del modelo"
Attribute StagesInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    StagesInFileName = getFileNameTag(tStages, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    StagesInFileName = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de cilindros del compresor del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function CylindersInFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute CylindersInFileName.VB_Description = "[modAPPFileNames] Extrae el número de cilindros del compresor del modelo"
Attribute CylindersInFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    CylindersInFileName = getFileNameTag(tCylinders, getContextWbkFileName(Wb))

    Exit Function
    
ErrorHandler:
    CylindersInFileName = "#ERROR"
End Function

'@UDF
'@Description: Devuelve el nombre de un File o un Workbook (segun contexto)
'@Category:
'@ArgumentDescriptions:
Public Function GetContextFileName(Optional Item As Object = Nothing) As Variant
Attribute GetContextFileName.VB_Description = "[modAPPFileNames] Devuelve el nombre de un File o un Workbook (segun contexto)"
Attribute GetContextFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Select Case True
        Case TypeName(Item) = "File"                 ' Se procesa un Path
            If EsLibroExcel(Item.Path) Then
                GetContextFileName = Item.Name
            End If
        Case TypeOf Item Is Workbook                   ' se procesa un Workbook
            GetContextFileName = Item.Name
        Case Item Is Nothing                           ' se procesa en contexto
            GetContextFileName = getContextWbkFileName(Item)
        Case Else
            Err.Raise vbObjectError + 513, "GetContextFileName", "No available file name"
    End Select
    Exit Function
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[GetContextFileName]"
    GetContextFileName = "#ERROR"
End Function

'@UDF
'@Description: Devuelve el nombre de un Workbook
'@Category:
'@ArgumentDescriptions:
Public Function getContextWbkFileName(Optional Wb As Workbook = Nothing) As Variant
Attribute getContextWbkFileName.VB_Description = "[modAPPFileNames] Devuelve el nombre de un Workbook"
Attribute getContextWbkFileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    Set Wb = GetContextWb(Wb)
    getContextWbkFileName = Wb.Name
    Exit Function
    
ErrorHandler:
    getContextWbkFileName = "#ERROR"
End Function

'@UDF
'@Description: Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
'@Category:
'@ArgumentDescriptions:
Public Function GetContextWb(Optional Wb As Workbook = Nothing) As Workbook
Attribute GetContextWb.VB_Description = "[modAPPFileNames] Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros. Aplica a: ActiveWorkbook|Cells Range"
Attribute GetContextWb.VB_ProcData.VB_Invoke_Func = " \n21"
    Select Case True
        Case Not Wb Is Nothing                       ' se procesa el parametro
            Set GetContextWb = Wb
        Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
            Set GetContextWb = Application.Caller.Worksheet.Parent
        Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
            Set GetContextWb = ActiveWorkbook
        Case Else
            Err.Raise vbObjectError + 513, "GetContextWb", "No available workbook"
    End Select
End Function

' Requiere referencia a "Microsoft Scripting Runtime" o usar Late Binding
Public Function EsLibroExcel(ByVal ruta As String) As Boolean
Attribute EsLibroExcel.VB_Description = "[modAPPFileNames] Requiere referencia a ""Microsoft Scripting Runtime"" o usar Late Binding"
Attribute EsLibroExcel.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim fso As Object
    Dim ext As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Verificar si el archivo existe
    If Not fso.FileExists(ruta) Then Exit Function
    
    ' 2. Extraer extensión (en minúsculas para comparar)
    ext = LCase(fso.GetExtensionName(ruta))
    
    ' 3. Lista de extensiones que Excel abre como Workbook nativo
    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls", "xltx", "xltm", "xlt", "xml"
            EsLibroExcel = True
        Case Else
            ' Nota: CSV y TXT se abren, pero técnicamente no son "Workbooks" nativos
            EsLibroExcel = False
    End Select
End Function

Attribute VB_Name = "modAPPFileNames"
'@IgnoreModule MissingAnnotationArgument
'@Folder "4-Oportunidades y compresores"
Option Explicit
Private Enum FNTag
    tCustomer
    tQuoteNr
    tQuoteRev
    tModel
    tFamily
    tCylinders
    tStages
End Enum

Public Function fileName(Optional wb As Workbook = Nothing) As Variant
Attribute fileName.VB_Description = "[UDFs_FileNames] file Name (función personalizada). Aplica a: ActiveWorkbook|Cells Range"
Attribute fileName.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case IsError(Application.Caller) And Not ActiveWorkbook Is Nothing ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Else
        Err.Raise vbObjectError + 513, "FileName", "No available workbook"
    End Select
    fileName = wb.Name
    Exit Function
    
ErrorHandler:
    fileName = "#ERROR"
End Function

'@Description: Devuelve el nombre del archivo actual (con extensión)
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Private Function getFileNameTag(tag As FNTag, Optional wb As Workbook = Nothing) As String
    Dim fileName As String
    Dim regEx As Object
    Dim matches As Object, sm As Integer
    
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto en VBA
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Customer", "No available workbook"
    End Select
    
    fileName = wb.Name
    
    Set regEx = CreateObject("VBScript.RegExp")
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
        regEx.Pattern = "^(?:\d+(?:[\-_]\d+)?)[ \-_]*rev\.?[ \-_]*(\d+)\s*\-"
            
        If regEx.Test(fileName) Then
            Set matches = regEx.Execute(fileName)
            getFileNameTag = matches(0).SubMatches(0)
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
    regEx.Pattern = "^(\d)\s?T?\s*E\s?(H[AGPX])\s?\-\s?(\d)\s?\-\s?[LGT]+"
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
Public Function Customer(Optional wb As Workbook = Nothing) As Variant
Attribute Customer.VB_Description = "[UDFs_FileNames] Extrae el cliente del nombre de archivo, del workbook actual o el pasado como parametro. Aplica a: ActiveWorkbook|Cells Range"
Attribute Customer.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Customer", "No available workbook"
    End Select
    
    Customer = getFileNameTag(tCustomer, wb)

    Exit Function
    
ErrorHandler:
    Customer = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de oferta del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function QuoteNr(Optional wb As Workbook = Nothing) As Variant
Attribute QuoteNr.VB_Description = "[UDFs_FileNames] Extrae el número de oferta del nombre de archivo, del workbook actual o el pasado como parametro. Aplica a: ActiveWorkbook|Cells Range"
Attribute QuoteNr.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "QuoteNr", "No available workbook"
    End Select
    
    QuoteNr = getFileNameTag(tQuoteNr, wb)

    Exit Function
    
ErrorHandler:
    QuoteNr = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de revisión de la oferta del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function QuoteRev(Optional wb As Workbook = Nothing) As Variant
Attribute QuoteRev.VB_Description = "[UDFs_FileNames] Extrae el número de revisión de la oferta del nombre de archivo, del workbook actual o el pasado como parametro. Aplica a: ActiveWorkbook|Cells Range"
Attribute QuoteRev.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "QuoteRev", "No available workbook"
    End Select
    
    QuoteRev = getFileNameTag(tQuoteRev, wb)

    Exit Function
    
ErrorHandler:
    QuoteRev = "#ERROR"
End Function

'@UDF
'@Description: Extrae el modelo del compresor del nombre del nombre de archivo, del workbook actual o el pasado como parametro
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function Model(Optional wb As Workbook = Nothing) As Variant
Attribute Model.VB_Description = "[UDFs_FileNames] Extrae el modelo del compresor del nombre del nombre de archivo, del workbook actual o el pasado como parametro. Aplica a: ActiveWorkbook|Cells Range"
Attribute Model.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Model", "No available workbook"
    End Select
    
    Model = getFileNameTag(tModel, wb)

    Exit Function
    
ErrorHandler:
    Model = "#ERROR"
End Function

'@UDF
'@Description: Extrae la familia del compresor (HA, HG, HP, HX) del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function Family(Optional wb As Workbook = Nothing) As Variant
Attribute Family.VB_Description = "[UDFs_FileNames] Extrae la familia del compresor (HA, HG, HP, HX) del modelo. Aplica a: ActiveWorkbook|Cells Range"
Attribute Family.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Family", "No available workbook"
    End Select
    
    Family = getFileNameTag(tFamily, wb)

    Exit Function
    
ErrorHandler:
    Family = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de etapas del compresor del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function Stages(Optional wb As Workbook = Nothing) As Variant
Attribute Stages.VB_Description = "[UDFs_FileNames] Extrae el número de etapas del compresor del modelo. Aplica a: ActiveWorkbook|Cells Range"
Attribute Stages.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Stages", "No available workbook"
    End Select
    
    Stages = getFileNameTag(tStages, wb)

    Exit Function
    
ErrorHandler:
    Stages = "#ERROR"
End Function

'@UDF
'@Description: Extrae el número de cilindros del compresor del modelo
'@Category: Información de Archivo
'@ArgumentDescriptions: (sin argumentos)
Public Function Cylinders(Optional wb As Workbook = Nothing) As Variant
Attribute Cylinders.VB_Description = "[UDFs_FileNames] Extrae el número de cilindros del compresor del modelo. Aplica a: ActiveWorkbook|Cells Range"
Attribute Cylinders.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    ' Para manejar correctamente el contexto, tanto en VBA, como al ser llamada como UDF, con y sin parametros
    Select Case True
    Case Not wb Is Nothing                       ' se procesa el parametro
    Case TypeOf Application.Caller Is Range      ' se procesa en contexto UDF
        Set wb = Application.Caller.Worksheet.Parent
    Case Not ActiveWorkbook Is Nothing           ' se procesa en contexto VBA
        Set wb = ActiveWorkbook
    Case Else
        Err.Raise vbObjectError + 513, "Cylinders", "No available workbook"
    End Select
    
    Cylinders = getFileNameTag(tCylinders, wb)

    Exit Function
    
ErrorHandler:
    Cylinders = "#ERROR"
End Function

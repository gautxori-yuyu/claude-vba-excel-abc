VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemTechnicalCalcProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' IMPLEMENTACION DE INFRAESTRUCTURA: FileSystemTechnicalCalcProvider
' ==================================================================
' Implementa ITechnicalCalcProvider usando el sistema de archivos.
' Contiene toda la logica de:
' - Busqueda de ficheros de calculo en carpeta de oportunidad
' - Identificacion por patrones regex
' - Parseo de nombres de fichero
'
' Patrones de VBScript:
' CalcFNamePattern = "(([A-Z]{3}\d{5})_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?.+?\.(?:txt|rtf|xlsx)$"
' GasVBNetExportedFNamesPattern = (complejo, multiples tipos de fichero)
' ==================================================================

'@Folder "2-Infraestructura.Providers"
Option Explicit
Implements ITechnicalCalcProvider

Private Const MODULE_NAME As String = "FileSystemTechnicalCalcProvider"

' Nombre de la subcarpeta de calculos tecnicos
Private Const CALC_FOLDER_NAME As String = "2.CALCULO TECNICO"

' Patrones de expresion regular
Private Const CALC_FNAME_PATTERN As String = "(([A-Z]{3}\d{5})_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?.+?\.(?:txt|rtf|xlsx)$"
Private Const GAS_VBNET_PATTERN As String = "(?:Antipul_|ABC_(Gas_Cooler|Aircooler|Reducer|Main Motor|Instrumentation|Gas_Filter|Frequency Converter|Cooling Water Pump|Dryer|Piston_rider_ring_selection|Cooling Water Tower|Pressure_Safety_Valve|Valves_selection)\-)?([A-Z]{3}\d{5}_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?(_calc(?:_multi)?)?.*?(\.(?:xlsx|rtf|txt))$"

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
' IMPLEMENTACION DE ITechnicalCalcProvider
' ==================================================================

Private Function ITechnicalCalcProvider_GetCalculations(ByVal opportunityPath As String) As Collection
    Set ITechnicalCalcProvider_GetCalculations = GetCalculations(opportunityPath)
End Function

Private Function ITechnicalCalcProvider_GetCalculationsByNumber(ByVal opportunityPath As String, ByVal calcNumber As String) As Collection
    Set ITechnicalCalcProvider_GetCalculationsByNumber = GetCalculationsByNumber(opportunityPath, calcNumber)
End Function

Private Function ITechnicalCalcProvider_GetCalculationByOption(ByVal opportunityPath As String, ByVal calcOption As String) As ITechnicalCalc
    Set ITechnicalCalcProvider_GetCalculationByOption = GetCalculationByOption(opportunityPath, calcOption)
End Function

Private Function ITechnicalCalcProvider_GetCalculationsByCompressorModel(ByVal opportunityPath As String) As Object
    Set ITechnicalCalcProvider_GetCalculationsByCompressorModel = GetCalculationsByCompressorModel(opportunityPath)
End Function

Private Function ITechnicalCalcProvider_IdentifyCalculationFiles(ByVal opportunityPath As String) As Collection
    Set ITechnicalCalcProvider_IdentifyCalculationFiles = IdentifyCalculationFiles(opportunityPath)
End Function

Private Function ITechnicalCalcProvider_IsValidCalculationFile(ByVal fileName As String) As Boolean
    ITechnicalCalcProvider_IsValidCalculationFile = IsValidCalculationFile(fileName)
End Function

Private Function ITechnicalCalcProvider_ParseCalculationFileName(ByVal fileName As String) As Object
    Set ITechnicalCalcProvider_ParseCalculationFileName = ParseCalculationFileName(fileName)
End Function

Private Function ITechnicalCalcProvider_MoveCalculationsToOpportunity(ByVal sourcePath As String, ByVal opportunityPath As String, Optional ByVal calcNumber As String = "") As Long
    ITechnicalCalcProvider_MoveCalculationsToOpportunity = MoveCalculationsToOpportunity(sourcePath, opportunityPath, calcNumber)
End Function

Private Function ITechnicalCalcProvider_GetCalculationsFolder(ByVal opportunityPath As String) As String
    ITechnicalCalcProvider_GetCalculationsFolder = GetCalculationsFolder(opportunityPath)
End Function

Private Function ITechnicalCalcProvider_CalculationsFolderExists(ByVal opportunityPath As String) As Boolean
    ITechnicalCalcProvider_CalculationsFolderExists = CalculationsFolderExists(opportunityPath)
End Function

' ==================================================================
' METODOS PUBLICOS
' ==================================================================

'@Description: Obtiene todos los calculos de una oportunidad
Public Function GetCalculations(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim calcFolder As String
    calcFolder = GetCalculationsFolder(opportunityPath)

    If Not mFileManager.ExisteCarpeta(calcFolder) Then
        LogWarning MODULE_NAME, "[GetCalculations] Carpeta no existe: " & calcFolder
        Set GetCalculations = result
        Exit Function
    End If

    ' Obtener ficheros de calculo
    Dim files As Collection
    Set files = IdentifyCalculationFiles(opportunityPath)

    ' Crear objetos ITechnicalCalc para cada fichero principal (_calc)
    Dim fileName As Variant
    Dim parsed As Object
    Dim calcSource As ExcelTechnicalCalcSource
    Dim filePath As String

    For Each fileName In files
        Set parsed = ParseCalculationFileName(CStr(fileName))
        If Not parsed Is Nothing Then
            If parsed("IsMainCalc") Then
                ' Crear ExcelTechnicalCalcSource
                filePath = calcFolder & "\" & CStr(fileName)
                Set calcSource = New ExcelTechnicalCalcSource
                calcSource.Initialize filePath, parsed

                ' Agregar a la coleccion usando CalcOption como key
                On Error Resume Next
                result.Add calcSource, parsed("CalcOption")
                If Err.Number <> 0 Then
                    ' Ya existe, ignorar duplicado
                    LogDebug MODULE_NAME, "[GetCalculations] Calculo duplicado: " & parsed("CalcOption")
                    Err.Clear
                End If
                On Error GoTo ErrHandler
            End If
        End If
    Next fileName

    LogDebug MODULE_NAME, "[GetCalculations] Encontrados " & result.Count & " calculos principales"
    Set GetCalculations = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetCalculations]"
    Set GetCalculations = result
End Function

'@Description: Obtiene calculos por numero
Public Function GetCalculationsByNumber(ByVal opportunityPath As String, ByVal calcNumber As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim allCalcs As Collection
    Set allCalcs = GetCalculations(opportunityPath)

    Dim calc As ITechnicalCalc
    For Each calc In allCalcs
        If calc.CalcNumber = calcNumber Then
            result.Add calc
        End If
    Next calc

    Set GetCalculationsByNumber = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetCalculationsByNumber]"
    Set GetCalculationsByNumber = result
End Function

'@Description: Obtiene un calculo por opcion
Public Function GetCalculationByOption(ByVal opportunityPath As String, ByVal calcOption As String) As ITechnicalCalc
    On Error GoTo ErrHandler

    Dim allCalcs As Collection
    Set allCalcs = GetCalculations(opportunityPath)

    Dim calc As ITechnicalCalc
    For Each calc In allCalcs
        If calc.CalcOption = calcOption Then
            Set GetCalculationByOption = calc
            Exit Function
        End If
    Next calc

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetCalculationByOption]"
End Function

'@Description: Agrupa calculos por modelo de compresor
Public Function GetCalculationsByCompressorModel(ByVal opportunityPath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    Dim allCalcs As Collection
    Set allCalcs = GetCalculations(opportunityPath)

    Dim calc As ITechnicalCalc
    Dim model As String
    For Each calc In allCalcs
        model = calc.CompressorModel
        If Len(model) > 0 Then
            If Not result.Exists(model) Then
                result.Add model, New Collection
            End If
            result(model).Add calc
        End If
    Next calc

    Set GetCalculationsByCompressorModel = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[GetCalculationsByCompressorModel]"
    Set GetCalculationsByCompressorModel = result
End Function

'@Description: Identifica ficheros de calculo en la carpeta
Public Function IdentifyCalculationFiles(ByVal opportunityPath As String) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler

    Dim calcFolder As String
    calcFolder = GetCalculationsFolder(opportunityPath)

    If Not mFileManager.ExisteCarpeta(calcFolder) Then
        Set IdentifyCalculationFiles = result
        Exit Function
    End If

    ' Obtener lista de ficheros
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(calcFolder) Then
        Set IdentifyCalculationFiles = result
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(calcFolder)

    Dim file As Object
    For Each file In folder.Files
        If IsValidCalculationFile(file.Name) Then
            result.Add file.Name
        End If
    Next file

    LogDebug MODULE_NAME, "[IdentifyCalculationFiles] Encontrados " & result.Count & " ficheros en " & calcFolder

    Set IdentifyCalculationFiles = result
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[IdentifyCalculationFiles]"
    Set IdentifyCalculationFiles = result
End Function

'@Description: Verifica si un fichero es un calculo valido
Public Function IsValidCalculationFile(ByVal fileName As String) As Boolean
    On Error GoTo ErrHandler

    ' Ignorar ficheros temporales
    If Left(fileName, 2) = "~$" Then Exit Function

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True

    ' Probar patron principal
    regEx.Pattern = CALC_FNAME_PATTERN
    If regEx.Test(fileName) Then
        IsValidCalculationFile = True
        Exit Function
    End If

    ' Probar patron Gas VBNet
    regEx.Pattern = GAS_VBNET_PATTERN
    If regEx.Test(fileName) Then
        IsValidCalculationFile = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[IsValidCalculationFile]"
End Function

'@Description: Parsea el nombre de un fichero de calculo
Public Function ParseCalculationFileName(ByVal fileName As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = True

    ' Probar patron Gas VBNet (mas especifico)
    regEx.Pattern = GAS_VBNET_PATTERN
    If regEx.Test(fileName) Then
        Dim matches As Object
        Set matches = regEx.Execute(fileName)
        If matches.Count > 0 Then
            Dim m As Object
            Set m = matches(0)

            result.Add "CalcOption", m.SubMatches(1)  ' ABC12345_01
            result.Add "CalcNumber", Left(m.SubMatches(1), 8)  ' ABC12345
            result.Add "Revision", IIf(Len(m.SubMatches(2)) > 0, CLng(m.SubMatches(2)), 0)
            result.Add "IsMainCalc", (m.SubMatches(3) = "_calc" Or m.SubMatches(3) = "_calc_multi")
            result.Add "IsMultiCalc", (m.SubMatches(3) = "_calc_multi")
            result.Add "FileType", m.SubMatches(4)
            result.Add "AuxType", m.SubMatches(0)  ' Antipul_, ABC_Gas_Cooler-, etc.

            Set ParseCalculationFileName = result
            Exit Function
        End If
    End If

    ' Probar patron basico
    regEx.Pattern = CALC_FNAME_PATTERN
    If regEx.Test(fileName) Then
        Set matches = regEx.Execute(fileName)
        If matches.Count > 0 Then
            Set m = matches(0)

            result.Add "CalcOption", m.SubMatches(0)  ' ABC12345_01
            result.Add "CalcNumber", m.SubMatches(1)  ' ABC12345
            result.Add "Revision", IIf(Len(m.SubMatches(2)) > 0, CLng(m.SubMatches(2)), 0)
            result.Add "IsMainCalc", (InStr(fileName, "_calc") > 0)
            result.Add "IsMultiCalc", (InStr(fileName, "_calc_multi") > 0)
            result.Add "FileType", "." & Split(fileName, ".")(UBound(Split(fileName, ".")))
            result.Add "AuxType", ""

            Set ParseCalculationFileName = result
            Exit Function
        End If
    End If

    ' No coincide con ningun patron
    Set ParseCalculationFileName = Nothing
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[ParseCalculationFileName]"
    Set ParseCalculationFileName = Nothing
End Function

'@Description: Mueve ficheros de calculo a oportunidad
Public Function MoveCalculationsToOpportunity(ByVal sourcePath As String, ByVal opportunityPath As String, Optional ByVal calcNumber As String = "") As Long
    Dim movedCount As Long
    On Error GoTo ErrHandler

    ' TODO: Implementar logica de mover ficheros
    ' Basado en la logica de cOp_CalcsTecn.identificaCalcTecnicos del VBScript
    LogWarning MODULE_NAME, "[MoveCalculationsToOpportunity] No implementado aun"

    MoveCalculationsToOpportunity = movedCount
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[MoveCalculationsToOpportunity]"
    MoveCalculationsToOpportunity = 0
End Function

'@Description: Obtiene la ruta de la carpeta de calculos
Public Function GetCalculationsFolder(ByVal opportunityPath As String) As String
    GetCalculationsFolder = opportunityPath & "\" & CALC_FOLDER_NAME
End Function

'@Description: Verifica si existe la carpeta de calculos
Public Function CalculationsFolderExists(ByVal opportunityPath As String) As Boolean
    If mFileManager Is Nothing Then Exit Function
    CalculationsFolderExists = mFileManager.ExisteCarpeta(GetCalculationsFolder(opportunityPath))
End Function

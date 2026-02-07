Attribute VB_Name = "modUTILSShellCmd"
' ==========================================
' FUNCIONES AUXILIARES
' ==========================================
'@Folder "Funciones auxiliares"

Option Explicit

Private Const MODULE_NAME As String = "modUTILSShellCmd"

Sub testFindImagesInFolder()
Attribute testFindImagesInFolder.VB_ProcData.VB_Invoke_Func = " \n0"
    Call FindImagesInFolder
End Sub

' Lanza una ventana de explorador con resultados de la busqueda indexada, de un patron de ficheros, en un path
Sub FindImagesInFolder(Optional strImgPattern As String, Optional strFolderPath As String)
Attribute FindImagesInFolder.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim strcmd
    If strFolderPath = "" Then strFolderPath = ActiveSheet.Range("A1").Value2
    If strImgPattern = "" Then strImgPattern = ActiveSheet.Range("A2").Value2
    strcmd = "explorer.exe " & Chr(34) & "search-ms:query=*" & strImgPattern & "*&crumb=location:""" & strFolderPath & """" & Chr(34)
    LogInfo MODULE_NAME, "[FindImagesInFolder] " & strcmd
    Call Shell(strcmd, vbNormalFocus)
End Sub

' @ArgumentDescriptions: ...cmdLineParams:parametros de linea de comandos al script; una cadena, con los argumentos debidamente separados;
' o un array, cuyos argumentos se separan debidamente en el script
Public Sub EjecutarScript(strOptB64Script As String, strScriptName As String, cmdLineParams As Variant, Optional bB64 As Boolean)
Attribute EjecutarScript.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim rutaTemp As String: rutaTemp = Environ("TEMP") & "\" & strScriptName
    Call ExtraerScriptVBScript(strOptB64Script, rutaTemp, bB64)
    Dim comando As String
    comando = "wscript """ & rutaTemp & """ "
    If IsArray(cmdLineParams) Then
        comando = comando & """" & Join(cmdLineParams, """ """) & """"
    ElseIf CStr(cmdLineParams) <> "" Then
        comando = comando & cmdLineParams
    End If
    Shell comando, vbHide
End Sub

Public Sub ExtraerScriptVBScript(strScript As String, rutaDestino As String, Optional bB64 As Boolean)
Attribute ExtraerScriptVBScript.VB_ProcData.VB_Invoke_Func = " \n0"
    ' el script está almacenado como cadena Base64, PENDIENTE añadir encriptacion RC4
    Dim fso As Object, archivo As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.CreateTextFile(rutaDestino, True)
    If bB64 Then
        archivo.Write Base64Decode(strScript)
    Else
        archivo.Write strScript
    End If
    archivo.Close
End Sub

'@Description: Comprime una carpeta completa en un archivo ZIP usando 7-Zip o Shell.Application (fallback)
'@Scope: Privado
'@ArgumentDescriptions: rutaCarpeta: Carpeta a comprimir | rutaZipDestino: Ruta completa del ZIP a crear
'@Returns: Boolean | True si se creó correctamente
Function ComprimirCarpetaAZip(rutaCarpeta As String, rutaZipDestino As String) As Boolean
Attribute ComprimirCarpetaAZip.VB_Description = "[modUTILSShellCmd] Comprime una carpeta completa en un archivo ZIP usando 7-Zip o Shell.Application (fallback)"
Attribute ComprimirCarpetaAZip.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim metodo As String
    Dim resultado As Boolean
    
    On Error GoTo ErrorHandler
    
    ' MÉTODO 1: Intentar con 7-Zip (más rápido y robusto)
    resultado = ComprimirCon7Zip(rutaCarpeta, rutaZipDestino)
    
    If resultado Then
        metodo = "7-Zip"
        ComprimirCarpetaAZip = True
    Else
        ' MÉTODO 2: Fallback a Shell.Application (método nativo mejorado)
        resultado = ComprimirConShellApplication(rutaCarpeta, rutaZipDestino)
        
        If resultado Then
            metodo = "Shell.Application"
            ComprimirCarpetaAZip = True
        Else
            ComprimirCarpetaAZip = False
            Exit Function
        End If
    End If
    
    LogInfo MODULE_NAME, "[ComprimirCarpetaAZip] Compresión exitosa usando: " & metodo
    Exit Function
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[ComprimirCarpetaAZip]"
    ComprimirCarpetaAZip = False
End Function

' ==========================================
' MÉTODO 1: COMPRESIÓN CON 7-ZIP
' ==========================================

'@Description: Intenta comprimir usando 7-Zip si está instalado
'@Scope: Privado
'@ArgumentDescriptions: rutaCarpeta: Carpeta a comprimir | rutaZipDestino: Ruta del ZIP
'@Returns: Boolean | True si 7-Zip funcionó correctamente
Function ComprimirCon7Zip(rutaCarpeta As String, rutaZipDestino As String) As Boolean
Attribute ComprimirCon7Zip.VB_Description = "[modUTILSShellCmd] MÉTODO 1: COMPRESIÓN CON 7-ZIP Intenta comprimir usando 7-Zip si está instalado"
Attribute ComprimirCon7Zip.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ruta7Zip As String
    Dim comando As String
    Dim fso As Object
    Dim wsh As Object
    Dim exitCode As Long
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Buscar 7-Zip en ubicaciones comunes
     ruta7Zip = Buscar7Zip()
    
    If ruta7Zip = "" Then
        ' 7-Zip no encontrado
        ruta7Zip = "7Z.exe"
    End If
    
    ' Eliminar ZIP destino si ya existe
    If fso.FileExists(rutaZipDestino) Then
        fso.DeleteFile rutaZipDestino, True
    End If
    
    ' Construir comando
    ' Sintaxis: 7z.exe a -tzip "destino.zip" "carpeta\*" -r
    comando = """" & ruta7Zip & """ a -tzip """ & rutaZipDestino & """ """ & rutaCarpeta & "\*"" -r"
    
    ' Ejecutar 7-Zip de forma sincrónica
    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run(comando, 0, True)  ' 0 = ventana oculta, True = esperar
    
    ' Verificar resultado
    If exitCode = 0 And fso.FileExists(rutaZipDestino) Then
        ' Verificar que el archivo tiene contenido
        If fso.GetFile(rutaZipDestino).SIZE > 100 Then  ' Más de 100 bytes (cabecera mínima)
            ComprimirCon7Zip = True
            LogInfo MODULE_NAME, "[ComprimirCon7Zip] Compresión exitosa con 7-Zip"
        Else
            ComprimirCon7Zip = False
        End If
    Else
        ComprimirCon7Zip = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[ComprimirCon7Zip]"
    ComprimirCon7Zip = False
End Function

'@Description: Busca el ejecutable de 7-Zip en ubicaciones comunes
'@Scope: Privado
'@Returns: String | Ruta completa a 7z.exe o "" si no se encuentra
Private Function Buscar7Zip() As String
    Dim fso As Object
    Dim rutas() As Variant
    Dim i As Long
    Dim ruta As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
'    ruta = ObtenerRutaEjecutable("7z.exe")
'    If fso.FileExists(ruta) Then
'        Buscar7Zip = ruta
'        Exit Function
'    End If
    
    ' Ubicaciones comunes de 7-Zip
    rutas = Array( _
        Environ("ProgramFiles") & "\7-Zip\7z.exe", _
        Environ("ProgramFiles(x86)") & "\7-Zip\7z.exe", _
        Environ("ProgramW6432") & "\7-Zip\7z.exe", _
        "C:\Program Files\7-Zip\7z.exe", _
        "C:\Program Files (x86)\7-Zip\7z.exe" _
    )
    
    ' Buscar en cada ubicación
    For i = LBound(rutas) To UBound(rutas)
        ruta = CStr(rutas(i))
        If fso.FileExists(ruta) Then
            Buscar7Zip = ruta
            Exit Function
        End If
    Next i
    
    ' No encontrado
    Buscar7Zip = ""
End Function

' ==========================================
' MÉTODO 2: COMPRESIÓN CON SHELL.APPLICATION (MEJORADO)
' ==========================================

'@Description: Comprime usando Shell.Application con sincronización robusta (basado en código de Gustav Brock)
'@Scope: Privado
'@ArgumentDescriptions: rutaCarpeta: Carpeta a comprimir | rutaZipDestino: Ruta del ZIP
'@Returns: Boolean | True si la compresión funcionó
Function ComprimirConShellApplication(rutaCarpeta As String, rutaZipDestino As String) As Boolean
Attribute ComprimirConShellApplication.VB_Description = "[modUTILSShellCmd] MÉTODO 2: COMPRESIÓN CON SHELL.APPLICATION (MEJORADO) Comprime usando Shell.Application con sincronización robusta (basado en código de Gustav Brock)"
Attribute ComprimirConShellApplication.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim fso As Object
    Dim shellApp As Object
    Dim carpetaOrigen As Object
    Dim zipTemp As String
    Dim zipFinal As String
    Dim zipHeader As String
    Dim isRemovable As Boolean
    Dim contador As Long
    Dim maxIntentos As Long
    Dim numArchivosOrigen As Long
    Dim numArchivosZip As Long
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellApp = CreateObject("Shell.Application")
    
    ' Verificar que la carpeta existe
    If Not fso.FolderExists(rutaCarpeta) Then
        ComprimirConShellApplication = False
        Exit Function
    End If
    
    Set carpetaOrigen = fso.GetFolder(rutaCarpeta)
    
    ' Determinar si es unidad removible o ruta de red
    isRemovable = EsRutaRemovibleORed(rutaCarpeta)
    
    ' Crear archivo ZIP temporal
    ' Si es unidad removible o red, usar carpeta TEMP local
    If isRemovable Then
        zipTemp = fso.BuildPath(Environ("TEMP"), fso.GetBaseName(fso.GetTempName()) & ".zip")
    Else
        zipTemp = fso.BuildPath(fso.GetParentFolderName(rutaZipDestino), fso.GetBaseName(fso.GetTempName()) & ".zip")
    End If
    
    ' Eliminar ZIP destino si existe
    If fso.FileExists(rutaZipDestino) Then
        fso.DeleteFile rutaZipDestino, True
    End If
    
    ' Crear archivo ZIP vacío con cabecera correcta
    ' Header proporcionado por Stuart McLachlan
    zipHeader = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, vbNullChar)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open zipTemp For Binary As #fileNum
    Put #fileNum, , zipHeader
    Close #fileNum
    
    ' Pequeña pausa para asegurar que el archivo se creó
    Sleep 200
    DoEvents
    
    ' Resolver rutas absolutas
    zipTemp = fso.GetAbsolutePathName(zipTemp)
    rutaCarpeta = fso.GetAbsolutePathName(rutaCarpeta)
    
    ' Contar archivos en carpeta origen (recursivamente)
    numArchivosOrigen = ContarArchivosRecursivo(carpetaOrigen)
    
    LogInfo MODULE_NAME, "[ComprimirConShellApplication] Iniciando compresión de " & numArchivosOrigen & " archivos..."
    
    ' Copiar archivos al ZIP usando Shell.Application
    On Error Resume Next
    shellApp.Namespace(CVar(zipTemp)).CopyHere shellApp.Namespace(CVar(rutaCarpeta)).Items, 16  ' 16 = Responder Sí a todo
    On Error GoTo ErrorHandler
    
    DoEvents
    
    ' SINCRONIZACIÓN ROBUSTA: Esperar hasta que todos los archivos estén en el ZIP
    contador = 0
    maxIntentos = 200  ' 200 * 250ms = 50 segundos máximo
    
    On Error Resume Next  ' Ignorar errores al consultar el ZIP mientras se crea
    
    Do While contador < maxIntentos
        DoEvents
        Sleep 250  ' Pausa de 250ms entre verificaciones
        
        ' Intentar contar archivos en el ZIP
        numArchivosZip = ContarItemsEnZip(shellApp, zipTemp)
        
        ' Verificar si ya tenemos todos los archivos
        If numArchivosZip > 0 And numArchivosZip >= numArchivosOrigen Then
            ' Esperar un poco más para asegurar que terminó
            Sleep 500
            Exit Do
        End If
        
        ' Debug cada 4 intentos (cada segundo)
        If contador Mod 4 = 0 Then
            LogInfo MODULE_NAME, "[ComprimirConShellApplication] Progreso: " & numArchivosZip & "/" & numArchivosOrigen
        End If
        
        contador = contador + 1
    Loop
    
    On Error GoTo ErrorHandler
    
    ' Verificación final
    numArchivosZip = ContarItemsEnZip(shellApp, zipTemp)
    
    If numArchivosZip < numArchivosOrigen Then
        LogWarning MODULE_NAME, "[ComprimirConShellApplication] ADVERTENCIA: ZIP incompleto (" & numArchivosZip & "/" & numArchivosOrigen & ")"
    Else
        LogInfo MODULE_NAME, "[ComprimirConShellApplication] Compresión completada (" & numArchivosZip & " archivos)"
    End If
    
    ' Mover/Renombrar ZIP temporal al destino final
    ' Usar bucle robusto como en código de Gustav Brock
    Const ErrorFileNotFound As Long = 53
    Const ErrorFileExists As Long = 58
    Const ErrorNoPermission As Long = 70
    
    On Error Resume Next
    
    Do
        DoEvents
        fso.MoveFile zipTemp, rutaZipDestino
        Sleep 50
        
        Select Case Err.Number
            Case ErrorFileExists, ErrorNoPermission
                ' Continuar intentando
                LogWarning MODULE_NAME, "[ComprimirConShellApplication] Reintentando mover archivo..."
            Case 0
                ' Éxito
                Exit Do
            Case ErrorFileNotFound
                ' El archivo ya se movió
                Exit Do
            Case Else
                ' Error inesperado
                LogError MODULE_NAME, "[ComprimirConShellApplication] Error al mover: ", , Err.Description
                Exit Do
        End Select
    Loop Until Err.Number = ErrorFileNotFound Or contador > 20
    
    On Error GoTo ErrorHandler
    
    ' Verificar que el archivo final existe y tiene tamaño razonable
    If fso.FileExists(rutaZipDestino) Then
        If fso.GetFile(rutaZipDestino).SIZE > 100 Then
            ComprimirConShellApplication = True
        Else
            ComprimirConShellApplication = False
        End If
    Else
        ComprimirConShellApplication = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[ComprimirConShellApplication]"
    ComprimirConShellApplication = False
End Function

'@Description: Cuenta archivos en una carpeta recursivamente
'@Scope: Privado
'@ArgumentDescriptions: carpeta: Objeto Folder
'@Returns: Long | Número total de archivos
Private Function ContarArchivosRecursivo(carpeta As Object) As Long
    Dim archivo As Object
    Dim subcarpeta As Object
    Dim total As Long
    
    On Error Resume Next
    
    total = 0
    
    ' Contar archivos en esta carpeta
    For Each archivo In carpeta.Files
        total = total + 1
    Next archivo
    
    ' Contar archivos en subcarpetas
    For Each subcarpeta In carpeta.SubFolders
        total = total + ContarArchivosRecursivo(subcarpeta)
    Next subcarpeta
    
    ContarArchivosRecursivo = total
    
    On Error GoTo 0
End Function

'@Description: Cuenta items en un archivo ZIP usando Shell.Application
'@Scope: Privado
'@ArgumentDescriptions: shellApp: Objeto Shell.Application | rutaZip: Ruta al ZIP
'@Returns: Long | Número de items en el ZIP (0 si error)
Private Function ContarItemsEnZip(shellApp As Object, rutaZip As String) As Long
    On Error Resume Next
    
    ContarItemsEnZip = shellApp.Namespace(CVar(rutaZip)).Items.Count
    
    If Err.Number <> 0 Then
        ContarItemsEnZip = 0
    End If
    
    On Error GoTo 0
End Function

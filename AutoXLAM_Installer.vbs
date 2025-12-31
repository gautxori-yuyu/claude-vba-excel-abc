' =====================================================
' SCRIPT DE INSTALACIÓN/DESINSTALACIÓN AUTOMÁTICA
' =====================================================
' Este script gestiona:
' 1. Copia del XLAM a la carpeta de complementos
' 2. Extracción del COM (FolderWatcherCOM.dll) desde dentro del XLAM
' 3. Registro/desregistro del complemento en Excel
'
' El XLAM es un fichero ZIP que contiene:
'   - xl/embeddings/FolderWatcherCOM.dll
'   - xl/embeddings/FolderWatcherCOM.dll.manifest
' =====================================================

Option Explicit

Const COM_DLL_NAME = "FolderWatcherCOM.dll"
Const COM_MANIFEST_NAME = "FolderWatcherCOM.dll.manifest"
Const COM_CONFIG_NAME = "FolderWatcherCOM.dll.config"
Const COM_EMBED_PATH = "xl\embeddings\"

Dim fso, args, modo, archivo, destino, nombre
Dim rutaFinal, excel, ai, vers

Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments

If args.Count < 4 Then
    MsgBox "Faltan parámetros en linea de comandos para poder completar la instalación." & vbcrlf & _
			"Uso: AutoXLAM_Installer.vbs /install|/uninstall archivo destino nombre", vbCritical
    WScript.Quit 1
End If

modo = args(0)
archivo = args(1)
destino = args(2)
nombre = args(3)

rutaFinal = destino & "\" & nombre & ".xlam"

' Esperar a que Excel libere los archivos
WScript.Sleep 4000

If modo = "/install" Then
    DoInstall
ElseIf modo = "/uninstall" Then
    DoUninstall
Else
    MsgBox "Modo de instalación no reconocido: " & modo & ", la instalación no se puede completar", vbCritical
    WScript.Quit 1
End If

' Limpiar: eliminar este script
On Error Resume Next
fso.DeleteFile WScript.ScriptFullName
On Error GoTo 0

WScript.Quit 0

' =====================================================
' INSTALACIÓN
' =====================================================
Sub DoInstall()
    If Not fso.FileExists(archivo) Then
        MsgBox "Error de instalación: no existe '" & archivo & "'", vbCritical
        WScript.Quit 1
    End If

    ' 1. Eliminar XLAM anterior si existe
    RemoveAddinInDestino rutaFinal

    ' 2. Extraer COM del XLAM origen ANTES de copiar
    '    (porque después de copiar el XLAM estará en uso por Excel)
    If Not ExtractCOMFromXLAM(archivo, destino) Then
        ' Si falla la extracción del COM, continuar de todos modos
        ' El complemento funcionará pero sin FolderWatcher
        WScript.Echo "Advertencia: No se pudo extraer el componente COM del XLAM. La vigilancia de carpetas no estará disponible."
    End If

    ' 3. Copiar XLAM al destino
    fso.CopyFile archivo, rutaFinal, True

    ' 4. Registrar en Excel
    Set excel = CreateObject("Excel.Application")
    excel.Visible = False

    For Each ai In excel.AddIns
        If LCase(ai.Name) = LCase(nombre & ".xlam") Then
            ai.Installed = True
            Exit For
        End If
    Next

    WScript.Sleep 1000

    If ai Is Nothing Then
        MsgBox "No ha sido posible completar la instalación. Por favor, habilita el complemento desde el menú de complementos de Excel.", vbCritical
    ElseIf Not ai.Installed Then
        MsgBox "No ha sido posible completar la instalación. Por favor, habilita el complemento desde el menú de complementos de Excel.", vbCritical
    Else
        MsgBox "Instalación completada, reinicia Excel.", vbInformation
    End If

    excel.Quit
    Set excel = Nothing
End Sub

' =====================================================
' DESINSTALACIÓN
' =====================================================
Sub DoUninstall()
    ' 1. Eliminar archivos COM primero (antes de que Excel los bloquee)
    RemoveCOMFiles destino

    ' 2. Eliminar XLAM
    RemoveAddinInDestino rutaFinal

    ' 3. Desregistrar de Excel
    Set excel = CreateObject("Excel.Application")
    vers = excel.Application.Version
    excel.Visible = False

    For Each ai In excel.AddIns
        If LCase(ai.Name) = LCase(nombre & ".xlam") Then
            ai.Installed = False
            Exit For
        End If
    Next

    Dim uninstallOK
    uninstallOK = True
    If Not ai Is Nothing Then
        If ai.Installed Then uninstallOK = False
    End If

    If Not uninstallOK Then
        MsgBox "No ha sido posible completar la desinstalación. Por favor, reinténtalo de nuevo o deshabilita el complemento desde el menú de complementos de Excel.", vbCritical
    Else
        MsgBox "Desinstalación completada, reinicia Excel.", vbInformation
    End If

    excel.Quit
    Set excel = Nothing

    ' 4. Limpiar registro
    CleanRegistry vers, nombre
End Sub

' =====================================================
' EXTRACCIÓN DEL COM DESDE EL XLAM (ZIP)
' =====================================================
Function ExtractCOMFromXLAM(xlamPath, destFolder)
    ExtractCOMFromXLAM = False

    On Error Resume Next

    ' Intentar primero con 7zip (más rápido y fiable)
    If TryExtractWith7Zip(xlamPath, destFolder) Then
        ExtractCOMFromXLAM = True
        Exit Function
    End If

    ' Si no hay 7zip, usar Shell.Application (Windows nativo)
    If TryExtractWithShell(xlamPath, destFolder) Then
        ExtractCOMFromXLAM = True
        Exit Function
    End If

    On Error GoTo 0
End Function

' Extracción usando 7-Zip
Function TryExtractWith7Zip(xlamPath, destFolder)
    TryExtractWith7Zip = False

    Dim shell, exec, sevenZipPath
    Set shell = CreateObject("WScript.Shell")

    ' Buscar 7z.exe en el PATH
    On Error Resume Next
    Set exec = shell.Exec("where 7z.exe")
    If Err.Number = 0 Then
        Do While exec.Status = 0
            WScript.Sleep 100
        Loop
        sevenZipPath = Trim(exec.StdOut.ReadLine)
    End If
    On Error GoTo 0

    If sevenZipPath = "" Or Not fso.FileExists(sevenZipPath) Then
        ' 7zip no encontrado
        Exit Function
    End If

    ' Extraer solo los archivos COM
    Dim cmd, dllPath, manifestPath, configPath
    dllPath = COM_EMBED_PATH & COM_DLL_NAME
    manifestPath = COM_EMBED_PATH & COM_MANIFEST_NAME
    configPath = COM_EMBED_PATH & COM_CONFIG_NAME

    ' Extraer DLL
    cmd = """" & sevenZipPath & """ e """ & xlamPath & """ -o""" & destFolder & """ """ & dllPath & """ -y"
    shell.Run cmd, 0, True

    ' Extraer Manifest
    cmd = """" & sevenZipPath & """ e """ & xlamPath & """ -o""" & destFolder & """ """ & manifestPath & """ -y"
    shell.Run cmd, 0, True

    ' Extraer Config
    cmd = """" & sevenZipPath & """ e """ & xlamPath & """ -o""" & destFolder & """ """ & configPath & """ -y"
    shell.Run cmd, 0, True

    ' Verificar que se extrajeron
    If fso.FileExists(destFolder & "\" & COM_DLL_NAME) And _
       fso.FileExists(destFolder & "\" & COM_CONFIG_NAME) And _
       fso.FileExists(destFolder & "\" & COM_MANIFEST_NAME) Then
        TryExtractWith7Zip = True
    End If

    Set shell = Nothing
End Function

' Extracción usando Shell.Application (Windows nativo)
Function TryExtractWithShell(xlamPath, destFolder)
    TryExtractWithShell = False

    On Error Resume Next

    ' Crear copia temporal como .zip
    Dim tempZip
    tempZip = fso.GetSpecialFolder(2) & "\" & fso.GetTempName() & ".zip"
    fso.CopyFile xlamPath, tempZip, True

    If Err.Number <> 0 Then Exit Function

    ' Usar Shell.Application para explorar el ZIP
    Dim shell, zipFolder, destFolderObj
    Set shell = CreateObject("Shell.Application")
    Set zipFolder = shell.NameSpace(tempZip)
    Set destFolderObj = shell.NameSpace(destFolder)

    If zipFolder Is Nothing Or destFolderObj Is Nothing Then
        fso.DeleteFile tempZip
        Exit Function
    End If

    ' Buscar la carpeta xl\embeddings dentro del ZIP
    Dim item, embedFolder
    Set embedFolder = Nothing

    ' Navegar a xl\embeddings
    Dim xlFolder
    For Each item In zipFolder.Items
        If LCase(item.Name) = "xl" Then
            Set xlFolder = shell.NameSpace(item.Path)
            Exit For
        End If
    Next

    If xlFolder Is Nothing Then
        fso.DeleteFile tempZip
        Exit Function
    End If

    For Each item In xlFolder.Items
        If LCase(item.Name) = "embeddings" Then
            Set embedFolder = shell.NameSpace(item.Path)
            Exit For
        End If
    Next

    If embedFolder Is Nothing Then
        fso.DeleteFile tempZip
        Exit Function
    End If

    ' Extraer los archivos COM
    Dim dllItem, manifestItem, configItem
    For Each item In embedFolder.Items
        If LCase(item.Name) = LCase(COM_DLL_NAME) Then
            Set dllItem = item
        ElseIf LCase(item.Name) = LCase(COM_MANIFEST_NAME) Then
            Set manifestItem = item
        ElseIf LCase(item.Name) = LCase(COM_CONFIG_NAME) Then
            Set configItem = item
        End If
    Next

    ' Copiar archivos al destino (16 = No mostrar diálogo, 1024 = No confirmar)
    If Not dllItem Is Nothing Then
        destFolderObj.CopyHere dllItem, 16 + 1024
        WScript.Sleep 500
    End If

    If Not manifestItem Is Nothing Then
        destFolderObj.CopyHere manifestItem, 16 + 1024
        WScript.Sleep 500
    End If

    If Not configItem Is Nothing Then
        destFolderObj.CopyHere configItem, 16 + 1024
        WScript.Sleep 500
    End If

    ' Limpiar
    fso.DeleteFile tempZip

    ' Verificar
    If fso.FileExists(destFolder & "\" & COM_DLL_NAME) And _
       fso.FileExists(destFolder & "\" & COM_CONFIG_NAME) And _
       fso.FileExists(destFolder & "\" & COM_MANIFEST_NAME) Then
        TryExtractWithShell = True
    End If

    On Error GoTo 0
End Function

' =====================================================
' ELIMINACIÓN DE ARCHIVOS COM
' =====================================================
Sub RemoveCOMFiles(folder)
    On Error Resume Next

    Dim dllPath, manifestPath, configPath
    dllPath = folder & "\" & COM_DLL_NAME
    manifestPath = folder & "\" & COM_MANIFEST_NAME
    configPath = folder & "\" & COM_CONFIG_NAME

    If fso.FileExists(dllPath) Then
        fso.DeleteFile dllPath, True
    End If

    If fso.FileExists(manifestPath) Then
        fso.DeleteFile manifestPath, True
    End If

    If fso.FileExists(configPath) Then
        fso.DeleteFile configPath, True
    End If

    On Error GoTo 0
End Sub

' =====================================================
' ELIMINACIÓN DEL XLAM EXISTENTE
' =====================================================
Sub RemoveAddinInDestino(rutaFinal)
    If Not fso.FileExists(rutaFinal) Then Exit Sub

    On Error Resume Next
    fso.DeleteFile rutaFinal, True
    On Error GoTo 0

    If Not fso.FileExists(rutaFinal) Then Exit Sub

    ' El archivo sigue existiendo, posiblemente bloqueado
    Dim objWMIService, colProcesses, answer, objProcess
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")

    If colProcesses.Count > 0 Then
        answer = MsgBox("Excel está en ejecución y puede estar bloqueando el archivo del complemento en destino. ¿Deseas cerrar Excel?", vbYesNo + vbQuestion)
        If answer = vbYes Then
            For Each objProcess In colProcesses
                objProcess.Terminate
            Next

            ' Esperar a que Excel cierre
            WScript.Sleep 3000

            ' Reintentar eliminar
            On Error Resume Next
            fso.DeleteFile rutaFinal, True
            On Error GoTo 0

            If fso.FileExists(rutaFinal) Then
                MsgBox "No ha sido posible completar el proceso. Por favor, cierra Excel manualmente y elimina el fichero" & vbCr & "'" & rutaFinal & "'.", vbCritical
                WScript.Quit 1
            End If
        Else
            MsgBox "No es posible completar el proceso. Por favor, cierra Excel manualmente e inténtalo de nuevo.", vbCritical
            WScript.Quit 1
        End If
    End If
End Sub

' =====================================================
' LIMPIEZA DEL REGISTRO
' =====================================================
Sub CleanRegistry(vers, nombre)
    On Error Resume Next

    Dim WshShell, i, clave, valor
    Set WshShell = CreateObject("WScript.Shell")

    For i = 1 To 50
        clave = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & vers & "\Excel\Options\OPEN" & i
        valor = WshShell.RegRead(clave)

        If Err.Number = 0 Then
            If InStr(1, valor, nombre & ".xlam", vbTextCompare) > 0 Then
                WshShell.RegDelete clave
                Exit For
            End If
        Else
            Err.Clear
            Exit For
        End If
    Next

    Set WshShell = Nothing
    On Error GoTo 0
End Sub

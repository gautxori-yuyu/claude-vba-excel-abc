Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
modo = args(0)
archivo = args(1)
destino = args(2)
nombre = args(3)

rutaFinal = destino & "\" & nombre & ".xlam"
WScript.Sleep (4000)
If modo = "/install" Then
    If Not fso.FileExists (archivo) Then MsgBox ("Error de instalacion: no existe '" & archivo & "'") : WScript.Quit
	RemoveAddinInDestino rutaFinal
    fso.CopyFile archivo, rutaFinal
	Set excel = CreateObject("Excel.Application")
	excel.Visible = False
    For Each ai In excel.AddIns
        If ai.Name = nombre & ".xlam" Then
            ai.Installed = True
            Exit For
        End If
    Next
    WScript.Sleep 1000
    If Not ai.Installed Then
    	MsgBox "No ha sido posible completar la instalación. Por favor, habilita el complemento desde el menú de complementos de Excel.", vbCritical
    Else
    	MsgBox "Instalación completada, reinicia Excel.", vbInformation
    End if
	excel.Quit
ElseIf modo = "/uninstall" Then
    RemoveAddinInDestino rutaFinal
	Set excel = CreateObject("Excel.Application")
	vers = Excel.Application.Version
	excel.Visible = False
    For Each ai In excel.AddIns
        If ai.Name = nombre & ".xlam" Then
            ai.Installed = False
            Exit For
        End If
    Next
    If ai.Installed Then
    	MsgBox "No ha sido posible completar la desinstalación. Por favor, reinténtalo de nuevo o deshabilita el complemento desde el menú de complementos de Excel.", vbCritical
    Else
    	MsgBox "Desinstalación completada, reinicia Excel.", vbInformation
    End If
	excel.Quit
	' borrar las marcas del registro
	Set WshShell = CreateObject("WScript.Shell")
	For i = 1 To 50
		clave = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & vers & "\Excel\Options\" & "OPEN" & i
        On Error Resume Next
		valor = WshShell.RegRead(clave)
        If Err Then Stop : Exit for
        On Error GoTo 0
		If InStr(1, valor, nombre & ".xlam", vbTextCompare) > 0 Then
			WshShell.RegDelete clave
			Exit For
		End If
	Next
    On Error GoTo 0
End If

fso.DeleteFile WScript.ScriptFullName

Sub RemoveAddinInDestino (rutaFinal)
    If fso.FileExists(rutaFinal) Then
        On Error Resume Next
        fso.DeleteFile rutaFinal
        On Error GoTo 0
    End If
    If fso.FileExists(rutaFinal) Then
        ' Check if Excel is running and offer to close it
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
        If colProcesses.Count > 0 Then
            answer = MsgBox("Excel está en ejecución y puede estar bloqueando el archivo del complemento en destino. ¿Deseas cerrar Excel?", vbYesNo + vbQuestion)
            If answer = vbYes Then
                For Each objProcess in colProcesses
                    objProcess.Terminate
                Next
                ' Wait a moment to ensure Excel has closed
                WScript.Sleep(3000)
                ' Try deleting the file again
                On Error Resume Next
                fso.DeleteFile rutaFinal
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
    End If
End Sub
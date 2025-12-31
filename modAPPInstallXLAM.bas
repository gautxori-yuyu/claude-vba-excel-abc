Attribute VB_Name = "modAPPInstallXLAM"
' ==========================================
' INSTALACIÓN Y DESINSTALACIÓN AUTOMÁTICA DEL COMPLEMENTO XLAM
' ==========================================
' Este módulo contiene la lógica de auto-instalación / auto-desinstalación
' del complemento XLAM en la carpeta de complementos del usuario, apoyándose
' en un script externo (VBScript) codificado en Base64 + RC4.
'
' El VBScript (AutoXLAM_Installer.vbs) gestiona:
'   1. Copia del XLAM a la carpeta de complementos
'   2. Extracción del COM desde dentro del XLAM (que es un ZIP):
'      - xl/embeddings/FolderWatcherCOM.dll
'      - xl/embeddings/FolderWatcherCOM.dll.manifest
'   3. Registro/desregistro del complemento en Excel
'
' IMPORTANTE: El COM debe estar embebido dentro del XLAM (carpeta xl/embeddings)
' ==========================================

'@Folder "1-Inicio e Instalacion"
'@IgnoreModule ProcedureNotUsed

Option Private Module
Option Explicit

' ---------------------------------------------------------------------
' CONSTANTES DE INSTALACIÓN
' ---------------------------------------------------------------------

' Constantes asociadas a la instalación del XLAM
Public Const SCRIPT_NOMBRE As String = "AutoXLAM_Installer.vbs"

' Constantes para el componente COM FolderWatcher
Private Const COM_DLL_NOMBRE As String = "FolderWatcherCOM.dll"
Private Const COM_MANIFEST_NOMBRE As String = "FolderWatcherCOM.dll.manifest"

' ---------------------------------------------------------------------
' UTILIDADES DE PREPARACIÓN DEL SCRIPT
' ---------------------------------------------------------------------

'@Description: Codifica el script de instalación VBScript a Base64 utilizando RC4 y lo transforma en una función VBA embebida.
'@Scope: Manipula archivos temporales del sistema y genera código embebido.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: Instalación XLAM
Sub archivoInstScriptToBase64RC4()
    ScriptToFunctionBase64RC4 _
        Replace(Environ$("TEMP") & "\" & "AutoXLAM_Installer.vbs", "\\", "\"), _
        Replace(Environ$("TEMP") & "\" & "AutoXLAM_Installer.Base64", "\\", "\"), _
        "INSTALLSCRIPT_B64RC4"
End Sub

' ---------------------------------------------------------------------
' FUNCIONES COM (DEPRECATED - El VBScript ahora gestiona la instalación)
' ---------------------------------------------------------------------
' NOTA: Estas funciones se mantienen como fallback pero ya no se usan
' directamente. El VBScript extrae el COM desde dentro del XLAM.

'@Description: [DEPRECATED] Instala los archivos COM desde una carpeta externa
'@Note: Ya no se usa. El VBScript extrae el COM del XLAM.
Private Function InstalarCOM(ByVal rutaOrigen As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim rutaDestino As String
    Dim rutaDLLOrigen As String
    Dim rutaManifestOrigen As String
    Dim rutaDLLDestino As String
    Dim rutaManifestDestino As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    rutaDestino = Application.UserLibraryPath

    ' Rutas de origen
    rutaDLLOrigen = rutaOrigen & COM_DLL_NOMBRE
    rutaManifestOrigen = rutaOrigen & COM_MANIFEST_NOMBRE

    ' Rutas de destino
    rutaDLLDestino = rutaDestino & COM_DLL_NOMBRE
    rutaManifestDestino = rutaDestino & COM_MANIFEST_NOMBRE

    ' Verificar que existen los archivos de origen
    If Not fso.FileExists(rutaDLLOrigen) Then
        Debug.Print "[InstalarCOM] - DLL no encontrada: " & rutaDLLOrigen
        InstalarCOM = False
        GoTo Cleanup
    End If

    If Not fso.FileExists(rutaManifestOrigen) Then
        Debug.Print "[InstalarCOM] - Manifest no encontrado: " & rutaManifestOrigen
        InstalarCOM = False
        GoTo Cleanup
    End If

    ' Eliminar archivos existentes si los hay
    On Error Resume Next
    If fso.FileExists(rutaDLLDestino) Then fso.DeleteFile rutaDLLDestino, True
    If fso.FileExists(rutaManifestDestino) Then fso.DeleteFile rutaManifestDestino, True
    On Error GoTo ErrHandler

    ' Copiar DLL
    fso.CopyFile rutaDLLOrigen, rutaDLLDestino, True
    Debug.Print "[InstalarCOM] - DLL copiada a: " & rutaDLLDestino

    ' Copiar Manifest
    fso.CopyFile rutaManifestOrigen, rutaManifestDestino, True
    Debug.Print "[InstalarCOM] - Manifest copiado a: " & rutaManifestDestino

    InstalarCOM = True
    Debug.Print "[InstalarCOM] - Componente COM instalado correctamente"

Cleanup:
    Set fso = Nothing
    Exit Function

ErrHandler:
    Debug.Print "[InstalarCOM] - Error: " & Err.Number & " - " & Err.Description
    InstalarCOM = False
    Resume Cleanup
End Function

'@Description: [DEPRECATED] Desinstala los archivos COM de la carpeta AddIns
'@Note: Ya no se usa. El VBScript elimina el COM durante desinstalación.
Private Function DesinstalarCOM() As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim rutaDestino As String
    Dim rutaDLL As String
    Dim rutaManifest As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    rutaDestino = Application.UserLibraryPath
    rutaDLL = rutaDestino & COM_DLL_NOMBRE
    rutaManifest = rutaDestino & COM_MANIFEST_NOMBRE

    ' Eliminar DLL si existe
    If fso.FileExists(rutaDLL) Then
        fso.DeleteFile rutaDLL, True
        Debug.Print "[DesinstalarCOM] - DLL eliminada: " & rutaDLL
    End If

    ' Eliminar Manifest si existe
    If fso.FileExists(rutaManifest) Then
        fso.DeleteFile rutaManifest, True
        Debug.Print "[DesinstalarCOM] - Manifest eliminado: " & rutaManifest
    End If

    DesinstalarCOM = True
    Debug.Print "[DesinstalarCOM] - Componente COM desinstalado correctamente"

Cleanup:
    Set fso = Nothing
    Exit Function

ErrHandler:
    Debug.Print "[DesinstalarCOM] - Error: " & Err.Number & " - " & Err.Description
    DesinstalarCOM = False
    Resume Cleanup
End Function

'@Description: Verifica si los archivos COM están instalados en la carpeta AddIns
'@Returns: Boolean | True si ambos archivos (DLL y manifest) existen
'@Category: Instalación COM
Public Function ComprobarCOMInstalado() As Boolean
    Dim fso As Object
    Dim rutaDestino As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaDestino = Application.UserLibraryPath

    ComprobarCOMInstalado = fso.FileExists(rutaDestino & COM_DLL_NOMBRE) And _
                            fso.FileExists(rutaDestino & COM_MANIFEST_NOMBRE)

    Set fso = Nothing
End Function

' ---------------------------------------------------------------------
' FLUJO PRINCIPAL DE AUTO-INSTALACIÓN / DESINSTALACIÓN
' ---------------------------------------------------------------------

'@Description: Gestiona automáticamente la instalación o desinstalación del complemento XLAM según su estado actual.
'@Scope: Manipula el libro actual, complementos de Excel y ejecuta scripts externos.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: Instalación XLAM
Public Sub AutoInstalador()

    ' Validar que se está ejecutando desde un XLAM
    If Not (ThisWorkbook.FileFormat = xlOpenXMLAddIn Or ThisWorkbook.FileFormat = xlAddIn) Then Exit Sub

    Dim rutaActual As String
    Dim rutaDestino As String

    rutaActual = ThisWorkbook.Path & "\"
    rutaDestino = Application.UserLibraryPath

    ' Si ya se ejecuta desde la carpeta destino, no hacer nada
    If rutaActual = rutaDestino Then
        Debug.Print "[AutoInstalador] - el complemento se inicia desde la ruta destino de instalación, NO se ejecuta el proceso de instalación / desinstalación"
        Exit Sub
    End If

    ' Si NO está instalado
    If Not ComprobarSiInstalado() Then

        ' Evitar sobrescribir un XLAM con el mismo nombre final
        If LCase$(ThisWorkbook.Name) = LCase$(APP_NAME & ".xlam") Then

            Debug.Print "[AutoInstalador] - XLAM no es posible instalarlo"
            MsgBox "El nombre del fichero a instalar tiene que ser diferente de '" & APP_NAME & ".xlam" & "'. Cámbialo si quieres hacer la instalación."

        ElseIf MsgBox("¿Deseas instalar este complemento?", vbYesNo + vbQuestion) = vbYes Then

            Debug.Print "[AutoInstalador] - ejecutando script de instalación"

            ' El VBScript extrae el COM desde dentro del XLAM y lo copia al destino
            EjecutarScript _
                INSTALLSCRIPT_B64RC4, _
                SCRIPT_NOMBRE, _
                Array("/install", ThisWorkbook.FullName, Application.UserLibraryPath, APP_NAME), _
                True

            If Application.Workbooks.Count <= 1 Then Application.Quit
            ThisWorkbook.Close SaveChanges:=False

        End If

    ' Si YA está instalado
    Else

        If MsgBox("Este complemento ya está instalado. ¿Deseas desinstalarlo?", vbYesNo + vbQuestion) = vbYes Then

            Debug.Print "[AutoInstalador] - ejecutando script de desinstalación"

            ' El VBScript elimina el COM y el XLAM del destino
            EjecutarScript _
                INSTALLSCRIPT_B64RC4, _
                SCRIPT_NOMBRE, _
                Array("/uninstall", ThisWorkbook.FullName, Application.UserLibraryPath, APP_NAME), _
                True

            If Application.Workbooks.Count <= 1 Then Application.Quit
            ThisWorkbook.Close SaveChanges:=False

        End If

    End If

End Sub

' ---------------------------------------------------------------------
' COMPROBACIÓN DE ESTADO DE INSTALACIÓN
' ---------------------------------------------------------------------

'@Description: Comprueba si el complemento XLAM está instalado correctamente en Excel y sincroniza su estado si hay inconsistencias.
'@Scope: Manipula la colección Application.AddIns y verifica archivos en el sistema.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Boolean | True si el XLAM está instalado; False en caso contrario.
'@Category: Instalación XLAM
Public Function ComprobarSiInstalado() As Boolean

    Dim ai As AddIn
    Dim bFExists As Boolean

    ' Verificar existencia física del XLAM
    bFExists = Dir(Application.UserLibraryPath & APP_NAME & ".xlam", vbNormal) <> ""

    For Each ai In Application.AddIns
        If ai.Name = APP_NAME & ".xlam" Then

            ' Estado inconsistente: marcado como instalado pero el fichero no existe
            If Not bFExists And ai.Installed Then
                Debug.Print "[ComprobarSiInstalado] - XLAM marcado como instalado, pero inexistente: forzando el proceso de desinstalación"
                ai.Installed = False
            End If

            ComprobarSiInstalado = ai.Installed
            Debug.Print "[ComprobarSiInstalado] - XLAM " & IIf(ComprobarSiInstalado, "", "no ") & "instalado"
            Exit Function

        End If
    Next ai

End Function

Function INSTALLSCRIPT_B64RC4() As String
    INSTALLSCRIPT_B64RC4 = _
        "JyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PQ0KJyBTQ1JJUFQgREUgSU5TVEFMQUNJw5NOL0RFU0lOU1RBTEFDScOTTiBBVVRPTcOBVElD" & _
        "QQ0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PQ0KJyBFc3RlIHNjcmlwdCBnZXN0aW9uYToNCicgMS4gQ29waWEgZGVsIFhMQU0gYSBs" & _
        "YSBjYXJwZXRhIGRlIGNvbXBsZW1lbnRvcw0KJyAyLiBFeHRyYWNjacOzbiBkZWwgQ09NIChG" & _
        "b2xkZXJXYXRjaGVyQ09NLmRsbCkgZGVzZGUgZGVudHJvIGRlbCBYTEFNDQonIDMuIFJlZ2lz" & _
        "dHJvL2Rlc3JlZ2lzdHJvIGRlbCBjb21wbGVtZW50byBlbiBFeGNlbA0KJw0KJyBFbCBYTEFN" & _
        "IGVzIHVuIGZpY2hlcm8gWklQIHF1ZSBjb250aWVuZToNCicgICAtIHhsL2VtYmVkZGluZ3Mv" & _
        "Rm9sZGVyV2F0Y2hlckNPTS5kbGwNCicgICAtIHhsL2VtYmVkZGluZ3MvRm9sZGVyV2F0Y2hl" & _
        "ckNPTS5kbGwubWFuaWZlc3QNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT0NCg0KT3B0aW9uIEV4cGxpY2l0DQoNCkNvbnN0IENPTV9E" & _
        "TExfTkFNRSA9ICJGb2xkZXJXYXRjaGVyQ09NLmRsbCINCkNvbnN0IENPTV9NQU5JRkVTVF9O" & _
        "QU1FID0gIkZvbGRlcldhdGNoZXJDT00uZGxsLm1hbmlmZXN0Ig0KQ29uc3QgQ09NX0NPTkZJ" & _
        "R19OQU1FID0gIkZvbGRlcldhdGNoZXJDT00uZGxsLmNvbmZpZyINCkNvbnN0IENPTV9FTUJF" & _
        "RF9QQVRIID0gInhsXGVtYmVkZGluZ3NcIg0KDQpEaW0gZnNvLCBhcmdzLCBtb2RvLCBhcmNo" & _
        "aXZvLCBkZXN0aW5vLCBub21icmUNCkRpbSBydXRhRmluYWwsIGV4Y2VsLCBhaSwgdmVycw0K" & _
        "DQpTZXQgZnNvID0gQ3JlYXRlT2JqZWN0KCJTY3JpcHRpbmcuRmlsZVN5c3RlbU9iamVjdCIp" & _
        "DQpTZXQgYXJncyA9IFdTY3JpcHQuQXJndW1lbnRzDQoNCklmIGFyZ3MuQ291bnQgPCA0IFRo" & _
        "ZW4NCiAgICBNc2dCb3ggIkZhbHRhbiBwYXLDoW1ldHJvcyBlbiBsaW5lYSBkZSBjb21hbmRv" & _
        "cyBwYXJhIHBvZGVyIGNvbXBsZXRhciBsYSBpbnN0YWxhY2nDs24uIiAmIHZiY3JsZiAmIF8N"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "CgkJCSJVc286IEF1dG9YTEFNX0luc3RhbGxlci52YnMgL2luc3RhbGx8L3VuaW5zdGFsbCBh" & _
        "cmNoaXZvIGRlc3Rpbm8gbm9tYnJlIiwgdmJDcml0aWNhbA0KICAgIFdTY3JpcHQuUXVpdCAx" & _
        "DQpFbmQgSWYNCg0KbW9kbyA9IGFyZ3MoMCkNCmFyY2hpdm8gPSBhcmdzKDEpDQpkZXN0aW5v" & _
        "ID0gYXJncygyKQ0Kbm9tYnJlID0gYXJncygzKQ0KDQpydXRhRmluYWwgPSBkZXN0aW5vICYg" & _
        "IlwiICYgbm9tYnJlICYgIi54bGFtIg0KDQonIEVzcGVyYXIgYSBxdWUgRXhjZWwgbGliZXJl" & _
        "IGxvcyBhcmNoaXZvcw0KV1NjcmlwdC5TbGVlcCA0MDAwDQoNCklmIG1vZG8gPSAiL2luc3Rh" & _
        "bGwiIFRoZW4NCiAgICBEb0luc3RhbGwNCkVsc2VJZiBtb2RvID0gIi91bmluc3RhbGwiIFRo" & _
        "ZW4NCiAgICBEb1VuaW5zdGFsbA0KRWxzZQ0KICAgIE1zZ0JveCAiTW9kbyBkZSBpbnN0YWxh" & _
        "Y2nDs24gbm8gcmVjb25vY2lkbzogIiAmIG1vZG8gJiAiLCBsYSBpbnN0YWxhY2nDs24gbm8g" & _
        "c2UgcHVlZGUgY29tcGxldGFyIiwgdmJDcml0aWNhbA0KICAgIFdTY3JpcHQuUXVpdCAxDQpF" & _
        "bmQgSWYNCg0KJyBMaW1waWFyOiBlbGltaW5hciBlc3RlIHNjcmlwdA0KT24gRXJyb3IgUmVz" & _
        "dW1lIE5leHQNCmZzby5EZWxldGVGaWxlIFdTY3JpcHQuU2NyaXB0RnVsbE5hbWUNCk9uIEVy" & _
        "cm9yIEdvVG8gMA0KDQpXU2NyaXB0LlF1aXQgMA0KDQonID09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQonIElOU1RBTEFDScOTTg0KJyA9" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0K" & _
        "U3ViIERvSW5zdGFsbCgpDQogICAgSWYgTm90IGZzby5GaWxlRXhpc3RzKGFyY2hpdm8pIFRo" & _
        "ZW4NCiAgICAgICAgTXNnQm94ICJFcnJvciBkZSBpbnN0YWxhY2nDs246IG5vIGV4aXN0ZSAn" & _
        "IiAmIGFyY2hpdm8gJiAiJyIsIHZiQ3JpdGljYWwNCiAgICAgICAgV1NjcmlwdC5RdWl0IDEN" & _
        "CiAgICBFbmQgSWYNCg0KICAgICcgMS4gRWxpbWluYXIgWExBTSBhbnRlcmlvciBzaSBleGlz" & _
        "dGUNCiAgICBSZW1vdmVBZGRpbkluRGVzdGlubyBydXRhRmluYWwNCg0KICAgICcgMi4gRXh0"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "cmFlciBDT00gZGVsIFhMQU0gb3JpZ2VuIEFOVEVTIGRlIGNvcGlhcg0KICAgICcgICAgKHBv" & _
        "cnF1ZSBkZXNwdcOpcyBkZSBjb3BpYXIgZWwgWExBTSBlc3RhcsOhIGVuIHVzbyBwb3IgRXhj" & _
        "ZWwpDQogICAgSWYgTm90IEV4dHJhY3RDT01Gcm9tWExBTShhcmNoaXZvLCBkZXN0aW5vKSBU" & _
        "aGVuDQogICAgICAgICcgU2kgZmFsbGEgbGEgZXh0cmFjY2nDs24gZGVsIENPTSwgY29udGlu" & _
        "dWFyIGRlIHRvZG9zIG1vZG9zDQogICAgICAgICcgRWwgY29tcGxlbWVudG8gZnVuY2lvbmFy" & _
        "w6EgcGVybyBzaW4gRm9sZGVyV2F0Y2hlcg0KICAgICAgICBXU2NyaXB0LkVjaG8gIkFkdmVy" & _
        "dGVuY2lhOiBObyBzZSBwdWRvIGV4dHJhZXIgZWwgY29tcG9uZW50ZSBDT00gZGVsIFhMQU0u" & _
        "IExhIHZpZ2lsYW5jaWEgZGUgY2FycGV0YXMgbm8gZXN0YXLDoSBkaXNwb25pYmxlLiINCiAg" & _
        "ICBFbmQgSWYNCg0KICAgICcgMy4gQ29waWFyIFhMQU0gYWwgZGVzdGlubw0KICAgIGZzby5D" & _
        "b3B5RmlsZSBhcmNoaXZvLCBydXRhRmluYWwsIFRydWUNCg0KICAgICcgNC4gUmVnaXN0cmFy" & _
        "IGVuIEV4Y2VsDQogICAgU2V0IGV4Y2VsID0gQ3JlYXRlT2JqZWN0KCJFeGNlbC5BcHBsaWNh" & _
        "dGlvbiIpDQogICAgZXhjZWwuVmlzaWJsZSA9IEZhbHNlDQoNCiAgICBGb3IgRWFjaCBhaSBJ" & _
        "biBleGNlbC5BZGRJbnMNCiAgICAgICAgSWYgTENhc2UoYWkuTmFtZSkgPSBMQ2FzZShub21i" & _
        "cmUgJiAiLnhsYW0iKSBUaGVuDQogICAgICAgICAgICBhaS5JbnN0YWxsZWQgPSBUcnVlDQog" & _
        "ICAgICAgICAgICBFeGl0IEZvcg0KICAgICAgICBFbmQgSWYNCiAgICBOZXh0DQoNCiAgICBX" & _
        "U2NyaXB0LlNsZWVwIDEwMDANCg0KICAgIElmIGFpIElzIE5vdGhpbmcgVGhlbg0KICAgICAg" & _
        "ICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJsZSBjb21wbGV0YXIgbGEgaW5zdGFsYWNpw7Nu" & _
        "LiBQb3IgZmF2b3IsIGhhYmlsaXRhIGVsIGNvbXBsZW1lbnRvIGRlc2RlIGVsIG1lbsO6IGRl" & _
        "IGNvbXBsZW1lbnRvcyBkZSBFeGNlbC4iLCB2YkNyaXRpY2FsDQogICAgRWxzZUlmIE5vdCBh" & _
        "aS5JbnN0YWxsZWQgVGhlbg0KICAgICAgICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJsZSBj"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "b21wbGV0YXIgbGEgaW5zdGFsYWNpw7NuLiBQb3IgZmF2b3IsIGhhYmlsaXRhIGVsIGNvbXBs" & _
        "ZW1lbnRvIGRlc2RlIGVsIG1lbsO6IGRlIGNvbXBsZW1lbnRvcyBkZSBFeGNlbC4iLCB2YkNy" & _
        "aXRpY2FsDQogICAgRWxzZQ0KICAgICAgICBNc2dCb3ggIkluc3RhbGFjacOzbiBjb21wbGV0" & _
        "YWRhLCByZWluaWNpYSBFeGNlbC4iLCB2YkluZm9ybWF0aW9uDQogICAgRW5kIElmDQoNCiAg" & _
        "ICBleGNlbC5RdWl0DQogICAgU2V0IGV4Y2VsID0gTm90aGluZw0KRW5kIFN1Yg0KDQonID09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQon" & _
        "IERFU0lOU1RBTEFDScOTTg0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PQ0KU3ViIERvVW5pbnN0YWxsKCkNCiAgICAnIDEuIEVsaW1p" & _
        "bmFyIGFyY2hpdm9zIENPTSBwcmltZXJvIChhbnRlcyBkZSBxdWUgRXhjZWwgbG9zIGJsb3F1" & _
        "ZWUpDQogICAgUmVtb3ZlQ09NRmlsZXMgZGVzdGlubw0KDQogICAgJyAyLiBFbGltaW5hciBY" & _
        "TEFNDQogICAgUmVtb3ZlQWRkaW5JbkRlc3Rpbm8gcnV0YUZpbmFsDQoNCiAgICAnIDMuIERl" & _
        "c3JlZ2lzdHJhciBkZSBFeGNlbA0KICAgIFNldCBleGNlbCA9IENyZWF0ZU9iamVjdCgiRXhj" & _
        "ZWwuQXBwbGljYXRpb24iKQ0KICAgIHZlcnMgPSBleGNlbC5BcHBsaWNhdGlvbi5WZXJzaW9u" & _
        "DQogICAgZXhjZWwuVmlzaWJsZSA9IEZhbHNlDQoNCiAgICBGb3IgRWFjaCBhaSBJbiBleGNl" & _
        "bC5BZGRJbnMNCiAgICAgICAgSWYgTENhc2UoYWkuTmFtZSkgPSBMQ2FzZShub21icmUgJiAi" & _
        "LnhsYW0iKSBUaGVuDQogICAgICAgICAgICBhaS5JbnN0YWxsZWQgPSBGYWxzZQ0KICAgICAg" & _
        "ICAgICAgRXhpdCBGb3INCiAgICAgICAgRW5kIElmDQogICAgTmV4dA0KDQogICAgRGltIHVu" & _
        "aW5zdGFsbE9LDQogICAgdW5pbnN0YWxsT0sgPSBUcnVlDQogICAgSWYgTm90IGFpIElzIE5v" & _
        "dGhpbmcgVGhlbg0KICAgICAgICBJZiBhaS5JbnN0YWxsZWQgVGhlbiB1bmluc3RhbGxPSyA9" & _
        "IEZhbHNlDQogICAgRW5kIElmDQoNCiAgICBJZiBOb3QgdW5pbnN0YWxsT0sgVGhlbg0KICAg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ICAgICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJsZSBjb21wbGV0YXIgbGEgZGVzaW5zdGFs" & _
        "YWNpw7NuLiBQb3IgZmF2b3IsIHJlaW50w6ludGFsbyBkZSBudWV2byBvIGRlc2hhYmlsaXRh" & _
        "IGVsIGNvbXBsZW1lbnRvIGRlc2RlIGVsIG1lbsO6IGRlIGNvbXBsZW1lbnRvcyBkZSBFeGNl" & _
        "bC4iLCB2YkNyaXRpY2FsDQogICAgRWxzZQ0KICAgICAgICBNc2dCb3ggIkRlc2luc3RhbGFj" & _
        "acOzbiBjb21wbGV0YWRhLCByZWluaWNpYSBFeGNlbC4iLCB2YkluZm9ybWF0aW9uDQogICAg" & _
        "RW5kIElmDQoNCiAgICBleGNlbC5RdWl0DQogICAgU2V0IGV4Y2VsID0gTm90aGluZw0KDQog" & _
        "ICAgJyA0LiBMaW1waWFyIHJlZ2lzdHJvDQogICAgQ2xlYW5SZWdpc3RyeSB2ZXJzLCBub21i" & _
        "cmUNCkVuZCBTdWINCg0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PQ0KJyBFWFRSQUNDScOTTiBERUwgQ09NIERFU0RFIEVMIFhMQU0g" & _
        "KFpJUCkNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT0NCkZ1bmN0aW9uIEV4dHJhY3RDT01Gcm9tWExBTSh4bGFtUGF0aCwgZGVzdEZv" & _
        "bGRlcikNCiAgICBFeHRyYWN0Q09NRnJvbVhMQU0gPSBGYWxzZQ0KDQogICAgT24gRXJyb3Ig" & _
        "UmVzdW1lIE5leHQNCg0KICAgICcgSW50ZW50YXIgcHJpbWVybyBjb24gN3ppcCAobcOhcyBy" & _
        "w6FwaWRvIHkgZmlhYmxlKQ0KICAgIElmIFRyeUV4dHJhY3RXaXRoN1ppcCh4bGFtUGF0aCwg" & _
        "ZGVzdEZvbGRlcikgVGhlbg0KICAgICAgICBFeHRyYWN0Q09NRnJvbVhMQU0gPSBUcnVlDQog" & _
        "ICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAgICcgU2kgbm8gaGF5IDd6" & _
        "aXAsIHVzYXIgU2hlbGwuQXBwbGljYXRpb24gKFdpbmRvd3MgbmF0aXZvKQ0KICAgIElmIFRy" & _
        "eUV4dHJhY3RXaXRoU2hlbGwoeGxhbVBhdGgsIGRlc3RGb2xkZXIpIFRoZW4NCiAgICAgICAg" & _
        "RXh0cmFjdENPTUZyb21YTEFNID0gVHJ1ZQ0KICAgICAgICBFeGl0IEZ1bmN0aW9uDQogICAg" & _
        "RW5kIElmDQoNCiAgICBPbiBFcnJvciBHb1RvIDANCkVuZCBGdW5jdGlvbg0KDQonIEV4dHJh"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "Y2Npw7NuIHVzYW5kbyA3LVppcA0KRnVuY3Rpb24gVHJ5RXh0cmFjdFdpdGg3WmlwKHhsYW1Q" & _
        "YXRoLCBkZXN0Rm9sZGVyKQ0KICAgIFRyeUV4dHJhY3RXaXRoN1ppcCA9IEZhbHNlDQoNCiAg" & _
        "ICBEaW0gc2hlbGwsIGV4ZWMsIHNldmVuWmlwUGF0aA0KICAgIFNldCBzaGVsbCA9IENyZWF0" & _
        "ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQoNCiAgICAnIEJ1c2NhciA3ei5leGUgZW4gZWwg" & _
        "UEFUSA0KICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQogICAgU2V0IGV4ZWMgPSBzaGVsbC5F" & _
        "eGVjKCJ3aGVyZSA3ei5leGUiKQ0KICAgIElmIEVyci5OdW1iZXIgPSAwIFRoZW4NCiAgICAg" & _
        "ICAgRG8gV2hpbGUgZXhlYy5TdGF0dXMgPSAwDQogICAgICAgICAgICBXU2NyaXB0LlNsZWVw" & _
        "IDEwMA0KICAgICAgICBMb29wDQogICAgICAgIHNldmVuWmlwUGF0aCA9IFRyaW0oZXhlYy5T" & _
        "dGRPdXQuUmVhZExpbmUpDQogICAgRW5kIElmDQogICAgT24gRXJyb3IgR29UbyAwDQoNCiAg" & _
        "ICBJZiBzZXZlblppcFBhdGggPSAiIiBPciBOb3QgZnNvLkZpbGVFeGlzdHMoc2V2ZW5aaXBQ" & _
        "YXRoKSBUaGVuDQogICAgICAgICcgN3ppcCBubyBlbmNvbnRyYWRvDQogICAgICAgIEV4aXQg" & _
        "RnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAgICcgRXh0cmFlciBzb2xvIGxvcyBhcmNoaXZv" & _
        "cyBDT00NCiAgICBEaW0gY21kLCBkbGxQYXRoLCBtYW5pZmVzdFBhdGgsIGNvbmZpZ1BhdGgN" & _
        "CiAgICBkbGxQYXRoID0gQ09NX0VNQkVEX1BBVEggJiBDT01fRExMX05BTUUNCiAgICBtYW5p" & _
        "ZmVzdFBhdGggPSBDT01fRU1CRURfUEFUSCAmIENPTV9NQU5JRkVTVF9OQU1FDQogICAgY29u" & _
        "ZmlnUGF0aCA9IENPTV9FTUJFRF9QQVRIICYgQ09NX0NPTkZJR19OQU1FDQoNCiAgICAnIEV4" & _
        "dHJhZXIgRExMDQogICAgY21kID0gIiIiIiAmIHNldmVuWmlwUGF0aCAmICIiIiBlICIiIiAm" & _
        "IHhsYW1QYXRoICYgIiIiIC1vIiIiICYgZGVzdEZvbGRlciAmICIiIiAiIiIgJiBkbGxQYXRo" & _
        "ICYgIiIiIC15Ig0KICAgIHNoZWxsLlJ1biBjbWQsIDAsIFRydWUNCg0KICAgICcgRXh0cmFl" & _
        "ciBNYW5pZmVzdA0KICAgIGNtZCA9ICIiIiIgJiBzZXZlblppcFBhdGggJiAiIiIgZSAiIiIg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "JiB4bGFtUGF0aCAmICIiIiAtbyIiIiAmIGRlc3RGb2xkZXIgJiAiIiIgIiIiICYgbWFuaWZl" & _
        "c3RQYXRoICYgIiIiIC15Ig0KICAgIHNoZWxsLlJ1biBjbWQsIDAsIFRydWUNCg0KICAgICcg" & _
        "RXh0cmFlciBDb25maWcNCiAgICBjbWQgPSAiIiIiICYgc2V2ZW5aaXBQYXRoICYgIiIiIGUg" & _
        "IiIiICYgeGxhbVBhdGggJiAiIiIgLW8iIiIgJiBkZXN0Rm9sZGVyICYgIiIiICIiIiAmIGNv" & _
        "bmZpZ1BhdGggJiAiIiIgLXkiDQogICAgc2hlbGwuUnVuIGNtZCwgMCwgVHJ1ZQ0KDQogICAg" & _
        "JyBWZXJpZmljYXIgcXVlIHNlIGV4dHJhamVyb24NCiAgICBJZiBmc28uRmlsZUV4aXN0cyhk" & _
        "ZXN0Rm9sZGVyICYgIlwiICYgQ09NX0RMTF9OQU1FKSBBbmQgXw0KICAgICAgIGZzby5GaWxl" & _
        "RXhpc3RzKGRlc3RGb2xkZXIgJiAiXCIgJiBDT01fQ09ORklHX05BTUUpIEFuZCBfDQogICAg" & _
        "ICAgZnNvLkZpbGVFeGlzdHMoZGVzdEZvbGRlciAmICJcIiAmIENPTV9NQU5JRkVTVF9OQU1F" & _
        "KSBUaGVuDQogICAgICAgIFRyeUV4dHJhY3RXaXRoN1ppcCA9IFRydWUNCiAgICBFbmQgSWYN" & _
        "Cg0KICAgIFNldCBzaGVsbCA9IE5vdGhpbmcNCkVuZCBGdW5jdGlvbg0KDQonIEV4dHJhY2Np" & _
        "w7NuIHVzYW5kbyBTaGVsbC5BcHBsaWNhdGlvbiAoV2luZG93cyBuYXRpdm8pDQpGdW5jdGlv" & _
        "biBUcnlFeHRyYWN0V2l0aFNoZWxsKHhsYW1QYXRoLCBkZXN0Rm9sZGVyKQ0KICAgIFRyeUV4" & _
        "dHJhY3RXaXRoU2hlbGwgPSBGYWxzZQ0KDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCg0K" & _
        "ICAgICcgQ3JlYXIgY29waWEgdGVtcG9yYWwgY29tbyAuemlwDQogICAgRGltIHRlbXBaaXAN" & _
        "CiAgICB0ZW1wWmlwID0gZnNvLkdldFNwZWNpYWxGb2xkZXIoMikgJiAiXCIgJiBmc28uR2V0" & _
        "VGVtcE5hbWUoKSAmICIuemlwIg0KICAgIGZzby5Db3B5RmlsZSB4bGFtUGF0aCwgdGVtcFpp" & _
        "cCwgVHJ1ZQ0KDQogICAgSWYgRXJyLk51bWJlciA8PiAwIFRoZW4gRXhpdCBGdW5jdGlvbg0K" & _
        "DQogICAgJyBVc2FyIFNoZWxsLkFwcGxpY2F0aW9uIHBhcmEgZXhwbG9yYXIgZWwgWklQDQog" & _
        "ICAgRGltIHNoZWxsLCB6aXBGb2xkZXIsIGRlc3RGb2xkZXJPYmoNCiAgICBTZXQgc2hlbGwg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "PSBDcmVhdGVPYmplY3QoIlNoZWxsLkFwcGxpY2F0aW9uIikNCiAgICBTZXQgemlwRm9sZGVy" & _
        "ID0gc2hlbGwuTmFtZVNwYWNlKHRlbXBaaXApDQogICAgU2V0IGRlc3RGb2xkZXJPYmogPSBz" & _
        "aGVsbC5OYW1lU3BhY2UoZGVzdEZvbGRlcikNCg0KICAgIElmIHppcEZvbGRlciBJcyBOb3Ro" & _
        "aW5nIE9yIGRlc3RGb2xkZXJPYmogSXMgTm90aGluZyBUaGVuDQogICAgICAgIGZzby5EZWxl" & _
        "dGVGaWxlIHRlbXBaaXANCiAgICAgICAgRXhpdCBGdW5jdGlvbg0KICAgIEVuZCBJZg0KDQog" & _
        "ICAgJyBCdXNjYXIgbGEgY2FycGV0YSB4bFxlbWJlZGRpbmdzIGRlbnRybyBkZWwgWklQDQog" & _
        "ICAgRGltIGl0ZW0sIGVtYmVkRm9sZGVyDQogICAgU2V0IGVtYmVkRm9sZGVyID0gTm90aGlu" & _
        "Zw0KDQogICAgJyBOYXZlZ2FyIGEgeGxcZW1iZWRkaW5ncw0KICAgIERpbSB4bEZvbGRlcg0K" & _
        "ICAgIEZvciBFYWNoIGl0ZW0gSW4gemlwRm9sZGVyLkl0ZW1zDQogICAgICAgIElmIExDYXNl" & _
        "KGl0ZW0uTmFtZSkgPSAieGwiIFRoZW4NCiAgICAgICAgICAgIFNldCB4bEZvbGRlciA9IHNo" & _
        "ZWxsLk5hbWVTcGFjZShpdGVtLlBhdGgpDQogICAgICAgICAgICBFeGl0IEZvcg0KICAgICAg" & _
        "ICBFbmQgSWYNCiAgICBOZXh0DQoNCiAgICBJZiB4bEZvbGRlciBJcyBOb3RoaW5nIFRoZW4N" & _
        "CiAgICAgICAgZnNvLkRlbGV0ZUZpbGUgdGVtcFppcA0KICAgICAgICBFeGl0IEZ1bmN0aW9u" & _
        "DQogICAgRW5kIElmDQoNCiAgICBGb3IgRWFjaCBpdGVtIEluIHhsRm9sZGVyLkl0ZW1zDQog" & _
        "ICAgICAgIElmIExDYXNlKGl0ZW0uTmFtZSkgPSAiZW1iZWRkaW5ncyIgVGhlbg0KICAgICAg" & _
        "ICAgICAgU2V0IGVtYmVkRm9sZGVyID0gc2hlbGwuTmFtZVNwYWNlKGl0ZW0uUGF0aCkNCiAg" & _
        "ICAgICAgICAgIEV4aXQgRm9yDQogICAgICAgIEVuZCBJZg0KICAgIE5leHQNCg0KICAgIElm" & _
        "IGVtYmVkRm9sZGVyIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBmc28uRGVsZXRlRmlsZSB0" & _
        "ZW1wWmlwDQogICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAgICcgRXh0" & _
        "cmFlciBsb3MgYXJjaGl2b3MgQ09NDQogICAgRGltIGRsbEl0ZW0sIG1hbmlmZXN0SXRlbSwg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "Y29uZmlnSXRlbQ0KICAgIEZvciBFYWNoIGl0ZW0gSW4gZW1iZWRGb2xkZXIuSXRlbXMNCiAg" & _
        "ICAgICAgSWYgTENhc2UoaXRlbS5OYW1lKSA9IExDYXNlKENPTV9ETExfTkFNRSkgVGhlbg0K" & _
        "ICAgICAgICAgICAgU2V0IGRsbEl0ZW0gPSBpdGVtDQogICAgICAgIEVsc2VJZiBMQ2FzZShp" & _
        "dGVtLk5hbWUpID0gTENhc2UoQ09NX01BTklGRVNUX05BTUUpIFRoZW4NCiAgICAgICAgICAg" & _
        "IFNldCBtYW5pZmVzdEl0ZW0gPSBpdGVtDQogICAgICAgIEVsc2VJZiBMQ2FzZShpdGVtLk5h" & _
        "bWUpID0gTENhc2UoQ09NX0NPTkZJR19OQU1FKSBUaGVuDQogICAgICAgICAgICBTZXQgY29u" & _
        "ZmlnSXRlbSA9IGl0ZW0NCiAgICAgICAgRW5kIElmDQogICAgTmV4dA0KDQogICAgJyBDb3Bp" & _
        "YXIgYXJjaGl2b3MgYWwgZGVzdGlubyAoMTYgPSBObyBtb3N0cmFyIGRpw6Fsb2dvLCAxMDI0" & _
        "ID0gTm8gY29uZmlybWFyKQ0KICAgIElmIE5vdCBkbGxJdGVtIElzIE5vdGhpbmcgVGhlbg0K" & _
        "ICAgICAgICBkZXN0Rm9sZGVyT2JqLkNvcHlIZXJlIGRsbEl0ZW0sIDE2ICsgMTAyNA0KICAg" & _
        "ICAgICBXU2NyaXB0LlNsZWVwIDUwMA0KICAgIEVuZCBJZg0KDQogICAgSWYgTm90IG1hbmlm" & _
        "ZXN0SXRlbSBJcyBOb3RoaW5nIFRoZW4NCiAgICAgICAgZGVzdEZvbGRlck9iai5Db3B5SGVy" & _
        "ZSBtYW5pZmVzdEl0ZW0sIDE2ICsgMTAyNA0KICAgICAgICBXU2NyaXB0LlNsZWVwIDUwMA0K" & _
        "ICAgIEVuZCBJZg0KDQogICAgSWYgTm90IGNvbmZpZ0l0ZW0gSXMgTm90aGluZyBUaGVuDQog" & _
        "ICAgICAgIGRlc3RGb2xkZXJPYmouQ29weUhlcmUgY29uZmlnSXRlbSwgMTYgKyAxMDI0DQog" & _
        "ICAgICAgIFdTY3JpcHQuU2xlZXAgNTAwDQogICAgRW5kIElmDQoNCiAgICAnIExpbXBpYXIN" & _
        "CiAgICBmc28uRGVsZXRlRmlsZSB0ZW1wWmlwDQoNCiAgICAnIFZlcmlmaWNhcg0KICAgIElm" & _
        "IGZzby5GaWxlRXhpc3RzKGRlc3RGb2xkZXIgJiAiXCIgJiBDT01fRExMX05BTUUpIEFuZCBf" & _
        "DQogICAgICAgZnNvLkZpbGVFeGlzdHMoZGVzdEZvbGRlciAmICJcIiAmIENPTV9DT05GSUdf" & _
        "TkFNRSkgQW5kIF8NCiAgICAgICBmc28uRmlsZUV4aXN0cyhkZXN0Rm9sZGVyICYgIlwiICYg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "Q09NX01BTklGRVNUX05BTUUpIFRoZW4NCiAgICAgICAgVHJ5RXh0cmFjdFdpdGhTaGVsbCA9" & _
        "IFRydWUNCiAgICBFbmQgSWYNCg0KICAgIE9uIEVycm9yIEdvVG8gMA0KRW5kIEZ1bmN0aW9u" & _
        "DQoNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT0NCicgRUxJTUlOQUNJw5NOIERFIEFSQ0hJVk9TIENPTQ0KJyA9PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KU3ViIFJlbW92ZUNP" & _
        "TUZpbGVzKGZvbGRlcikNCiAgICBPbiBFcnJvciBSZXN1bWUgTmV4dA0KDQogICAgRGltIGRs" & _
        "bFBhdGgsIG1hbmlmZXN0UGF0aCwgY29uZmlnUGF0aA0KICAgIGRsbFBhdGggPSBmb2xkZXIg" & _
        "JiAiXCIgJiBDT01fRExMX05BTUUNCiAgICBtYW5pZmVzdFBhdGggPSBmb2xkZXIgJiAiXCIg" & _
        "JiBDT01fTUFOSUZFU1RfTkFNRQ0KICAgIGNvbmZpZ1BhdGggPSBmb2xkZXIgJiAiXCIgJiBD" & _
        "T01fQ09ORklHX05BTUUNCg0KICAgIElmIGZzby5GaWxlRXhpc3RzKGRsbFBhdGgpIFRoZW4N" & _
        "CiAgICAgICAgZnNvLkRlbGV0ZUZpbGUgZGxsUGF0aCwgVHJ1ZQ0KICAgIEVuZCBJZg0KDQog" & _
        "ICAgSWYgZnNvLkZpbGVFeGlzdHMobWFuaWZlc3RQYXRoKSBUaGVuDQogICAgICAgIGZzby5E" & _
        "ZWxldGVGaWxlIG1hbmlmZXN0UGF0aCwgVHJ1ZQ0KICAgIEVuZCBJZg0KDQogICAgSWYgZnNv" & _
        "LkZpbGVFeGlzdHMoY29uZmlnUGF0aCkgVGhlbg0KICAgICAgICBmc28uRGVsZXRlRmlsZSBj" & _
        "b25maWdQYXRoLCBUcnVlDQogICAgRW5kIElmDQoNCiAgICBPbiBFcnJvciBHb1RvIDANCkVu" & _
        "ZCBTdWINCg0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PQ0KJyBFTElNSU5BQ0nDk04gREVMIFhMQU0gRVhJU1RFTlRFDQonID09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQpTdWIg" & _
        "UmVtb3ZlQWRkaW5JbkRlc3Rpbm8ocnV0YUZpbmFsKQ0KICAgIElmIE5vdCBmc28uRmlsZUV4" & _
        "aXN0cyhydXRhRmluYWwpIFRoZW4gRXhpdCBTdWINCg0KICAgIE9uIEVycm9yIFJlc3VtZSBO"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ZXh0DQogICAgZnNvLkRlbGV0ZUZpbGUgcnV0YUZpbmFsLCBUcnVlDQogICAgT24gRXJyb3Ig" & _
        "R29UbyAwDQoNCiAgICBJZiBOb3QgZnNvLkZpbGVFeGlzdHMocnV0YUZpbmFsKSBUaGVuIEV4" & _
        "aXQgU3ViDQoNCiAgICAnIEVsIGFyY2hpdm8gc2lndWUgZXhpc3RpZW5kbywgcG9zaWJsZW1l" & _
        "bnRlIGJsb3F1ZWFkbw0KICAgIERpbSBvYmpXTUlTZXJ2aWNlLCBjb2xQcm9jZXNzZXMsIGFu" & _
        "c3dlciwgb2JqUHJvY2Vzcw0KICAgIFNldCBvYmpXTUlTZXJ2aWNlID0gR2V0T2JqZWN0KCJ3" & _
        "aW5tZ210czpcXC5ccm9vdFxjaW12MiIpDQogICAgU2V0IGNvbFByb2Nlc3NlcyA9IG9ialdN" & _
        "SVNlcnZpY2UuRXhlY1F1ZXJ5KCJTZWxlY3QgKiBmcm9tIFdpbjMyX1Byb2Nlc3MgV2hlcmUg" & _
        "TmFtZSA9ICdFWENFTC5FWEUnIikNCg0KICAgIElmIGNvbFByb2Nlc3Nlcy5Db3VudCA+IDAg" & _
        "VGhlbg0KICAgICAgICBhbnN3ZXIgPSBNc2dCb3goIkV4Y2VsIGVzdMOhIGVuIGVqZWN1Y2nD" & _
        "s24geSBwdWVkZSBlc3RhciBibG9xdWVhbmRvIGVsIGFyY2hpdm8gZGVsIGNvbXBsZW1lbnRv" & _
        "IGVuIGRlc3Rpbm8uIMK/RGVzZWFzIGNlcnJhciBFeGNlbD8iLCB2Ylllc05vICsgdmJRdWVz" & _
        "dGlvbikNCiAgICAgICAgSWYgYW5zd2VyID0gdmJZZXMgVGhlbg0KICAgICAgICAgICAgRm9y" & _
        "IEVhY2ggb2JqUHJvY2VzcyBJbiBjb2xQcm9jZXNzZXMNCiAgICAgICAgICAgICAgICBvYmpQ" & _
        "cm9jZXNzLlRlcm1pbmF0ZQ0KICAgICAgICAgICAgTmV4dA0KDQogICAgICAgICAgICAnIEVz" & _
        "cGVyYXIgYSBxdWUgRXhjZWwgY2llcnJlDQogICAgICAgICAgICBXU2NyaXB0LlNsZWVwIDMw" & _
        "MDANCg0KICAgICAgICAgICAgJyBSZWludGVudGFyIGVsaW1pbmFyDQogICAgICAgICAgICBP" & _
        "biBFcnJvciBSZXN1bWUgTmV4dA0KICAgICAgICAgICAgZnNvLkRlbGV0ZUZpbGUgcnV0YUZp" & _
        "bmFsLCBUcnVlDQogICAgICAgICAgICBPbiBFcnJvciBHb1RvIDANCg0KICAgICAgICAgICAg" & _
        "SWYgZnNvLkZpbGVFeGlzdHMocnV0YUZpbmFsKSBUaGVuDQogICAgICAgICAgICAgICAgTXNn" & _
        "Qm94ICJObyBoYSBzaWRvIHBvc2libGUgY29tcGxldGFyIGVsIHByb2Nlc28uIFBvciBmYXZv"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ciwgY2llcnJhIEV4Y2VsIG1hbnVhbG1lbnRlIHkgZWxpbWluYSBlbCBmaWNoZXJvIiAmIHZi" & _
        "Q3IgJiAiJyIgJiBydXRhRmluYWwgJiAiJy4iLCB2YkNyaXRpY2FsDQogICAgICAgICAgICAg" & _
        "ICAgV1NjcmlwdC5RdWl0IDENCiAgICAgICAgICAgIEVuZCBJZg0KICAgICAgICBFbHNlDQog" & _
        "ICAgICAgICAgICBNc2dCb3ggIk5vIGVzIHBvc2libGUgY29tcGxldGFyIGVsIHByb2Nlc28u" & _
        "IFBvciBmYXZvciwgY2llcnJhIEV4Y2VsIG1hbnVhbG1lbnRlIGUgaW50w6ludGFsbyBkZSBu" & _
        "dWV2by4iLCB2YkNyaXRpY2FsDQogICAgICAgICAgICBXU2NyaXB0LlF1aXQgMQ0KICAgICAg" & _
        "ICBFbmQgSWYNCiAgICBFbmQgSWYNCkVuZCBTdWINCg0KJyA9PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KJyBMSU1QSUVaQSBERUwgUkVH" & _
        "SVNUUk8NCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT0NClN1YiBDbGVhblJlZ2lzdHJ5KHZlcnMsIG5vbWJyZSkNCiAgICBPbiBFcnJv" & _
        "ciBSZXN1bWUgTmV4dA0KDQogICAgRGltIFdzaFNoZWxsLCBpLCBjbGF2ZSwgdmFsb3INCiAg" & _
        "ICBTZXQgV3NoU2hlbGwgPSBDcmVhdGVPYmplY3QoIldTY3JpcHQuU2hlbGwiKQ0KDQogICAg" & _
        "Rm9yIGkgPSAxIFRvIDUwDQogICAgICAgIGNsYXZlID0gIkhLRVlfQ1VSUkVOVF9VU0VSXFNv" & _
        "ZnR3YXJlXE1pY3Jvc29mdFxPZmZpY2VcIiAmIHZlcnMgJiAiXEV4Y2VsXE9wdGlvbnNcT1BF" & _
        "TiIgJiBpDQogICAgICAgIHZhbG9yID0gV3NoU2hlbGwuUmVnUmVhZChjbGF2ZSkNCg0KICAg" & _
        "ICAgICBJZiBFcnIuTnVtYmVyID0gMCBUaGVuDQogICAgICAgICAgICBJZiBJblN0cigxLCB2" & _
        "YWxvciwgbm9tYnJlICYgIi54bGFtIiwgdmJUZXh0Q29tcGFyZSkgPiAwIFRoZW4NCiAgICAg" & _
        "ICAgICAgICAgICBXc2hTaGVsbC5SZWdEZWxldGUgY2xhdmUNCiAgICAgICAgICAgICAgICBF" & _
        "eGl0IEZvcg0KICAgICAgICAgICAgRW5kIElmDQogICAgICAgIEVsc2UNCiAgICAgICAgICAg" & _
        "IEVyci5DbGVhcg0KICAgICAgICAgICAgRXhpdCBGb3INCiAgICAgICAgRW5kIElmDQogICAg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "TmV4dA0KDQogICAgU2V0IFdzaFNoZWxsID0gTm90aGluZw0KICAgIE9uIEVycm9yIEdvVG8g" & _
        "MA0KRW5kIFN1Yg0K"
End Function



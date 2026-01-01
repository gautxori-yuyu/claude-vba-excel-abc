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
Attribute archivoInstScriptToBase64RC4.VB_ProcData.VB_Invoke_Func = " \n0"
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
Attribute AutoInstalador.VB_ProcData.VB_Invoke_Func = " \n0"
    
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
Attribute ComprobarSiInstalado.VB_ProcData.VB_Invoke_Func = " \n0"
    
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
        "PQ0KJyBTQ1JJUFQgREUgSU5TVEFMQUNJ004vREVTSU5TVEFMQUNJ004gQVVUT03BVElDQQ0K" & _
        "JyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PQ0KJyBFc3RlIHNjcmlwdCBnZXN0aW9uYToNCicgMS4gQ29waWEgZGVsIFhMQU0gYSBsYSBj" & _
        "YXJwZXRhIGRlIGNvbXBsZW1lbnRvcw0KJyAyLiBFeHRyYWNjafNuIGRlbCBDT00gKEZvbGRl" & _
        "cldhdGNoZXJDT00uZGxsKSBkZXNkZSBkZW50cm8gZGVsIFhMQU0NCicgMy4gUmVnaXN0cm8v" & _
        "ZGVzcmVnaXN0cm8gZGVsIGNvbXBsZW1lbnRvIGVuIEV4Y2VsDQonDQonIEVsIFhMQU0gZXMg" & _
        "dW4gZmljaGVybyBaSVAgcXVlIGNvbnRpZW5lOg0KJyAgIC0geGwvZW1iZWRkaW5ncy9Gb2xk" & _
        "ZXJXYXRjaGVyQ09NLmRsbA0KJyAgIC0geGwvZW1iZWRkaW5ncy9Gb2xkZXJXYXRjaGVyQ09N" & _
        "LmRsbC5tYW5pZmVzdA0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PQ0KDQpPcHRpb24gRXhwbGljaXQNCg0KQ29uc3QgQ09NX0RMTF9O" & _
        "QU1FID0gIkZvbGRlcldhdGNoZXJDT00uZGxsIg0KQ29uc3QgQ09NX01BTklGRVNUX05BTUUg" & _
        "PSAiRm9sZGVyV2F0Y2hlckNPTS5kbGwubWFuaWZlc3QiDQpDb25zdCBDT01fQ09ORklHX05B" & _
        "TUUgPSAiRm9sZGVyV2F0Y2hlckNPTS5kbGwuY29uZmlnIg0KQ29uc3QgQ09NX0VNQkVEX1BB" & _
        "VEggPSAieGxcZW1iZWRkaW5nc1wiDQoNCicgPT09PT09PT09PSBDT05TVEFOVEVTIFBBUkEg" & _
        "UkVHSVNUUk8gQ09NUE9ORU5URSBDT00gPT09PT09PT09PQ0KQ29uc3QgR1VJRF9DTFNJRCA9" & _
        "ICJ7QzNFNUY4QjItNTY3OC00Q0RFLUFCMTItMTIzNDU2Nzg5MEFEfSINCkNvbnN0IEdVSURf" & _
        "SW50ZXJmYWNlMSA9ICJ7OERBNUExNkEtRTBBMi0zNDQ4LTk1NUYtMkVFRTg3RkVCMEI0fSIN" & _
        "CkNvbnN0IEdVSURfSW50ZXJmYWNlMiA9ICJ7QjFEOUY3RTEtQUFBQS00Q0RFLUJDMTItMTIz" & _
        "NDU2Nzg5MEFDfSINCkNvbnN0IEdVSURfVHlwZUxpYiA9ICJ7RTBCQ0MwM0MtRDE1NS00RUEz"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "LUJDQjgtMUQwNzE3MTlFODU0fSINCg0KQ29uc3QgUFJPWFlTVFVCX0NMU0lEMSA9ICJ7MDAw" & _
        "MjA0MjQtMDAwMC0wMDAwLUMwMDAtMDAwMDAwMDAwMDQ2fSIgICcgUGFyYSBpbnRlcmZhY2Vz" & _
        "IG5vcm1hbGVzDQpDb25zdCBQUk9YWVNUVUJfQ0xTSUQyID0gInswMDAyMDQyMC0wMDAwLTAw" & _
        "MDAtQzAwMC0wMDAwMDAwMDAwNDZ9IiAgJyBQYXJhIGludGVyZmFjZXMgZGUgZXZlbnRvcw0K" & _
        "DQpDb25zdCBBU1NFTUJMWV9JTkZPID0gIkZvbGRlcldhdGNoZXJDT00sIFZlcnNpb249MS4w" & _
        "LjAuMCwgQ3VsdHVyZT1uZXV0cmFsLCBQdWJsaWNLZXlUb2tlbj0xZmIzZDY3ZGMzZWIyZTlm" & _
        "Ig0KQ29uc3QgUlVOVElNRV9WRVJTSU9OID0gInY0LjAuMzAzMTkiDQpDb25zdCBQUk9HX0lE" & _
        "ID0gIkZvbGRlcldhdGNoZXIuTW9uaXRvciINCkNvbnN0IENMQVNTX05BTUUgPSAiRm9sZGVy" & _
        "V2F0Y2hlckNPTS5Gb2xkZXJXYXRjaGVyIg0KDQpEaW0gZnNvLCBhcmdzLCBtb2RvLCBhcmNo" & _
        "aXZvLCBkZXN0aW5vLCBub21icmUNCkRpbSBydXRhRmluYWwsIGV4Y2VsLCBhaSwgdmVycw0K" & _
        "DQpTZXQgZnNvID0gQ3JlYXRlT2JqZWN0KCJTY3JpcHRpbmcuRmlsZVN5c3RlbU9iamVjdCIp" & _
        "DQpTZXQgYXJncyA9IFdTY3JpcHQuQXJndW1lbnRzDQoNCklmIGFyZ3MuQ291bnQgPCA0IFRo" & _
        "ZW4NCiAgICBNc2dCb3ggIkZhbHRhbiBwYXLhbWV0cm9zIGVuIGxpbmVhIGRlIGNvbWFuZG9z" & _
        "IHBhcmEgcG9kZXIgY29tcGxldGFyIGxhIGluc3RhbGFjafNuLiIgJiB2YmNybGYgJiBfDQoJ" & _
        "CQkiVXNvOiBBdXRvWExBTV9JbnN0YWxsZXIudmJzIC9pbnN0YWxsfC91bmluc3RhbGwgYXJj" & _
        "aGl2byBkZXN0aW5vIG5vbWJyZSIsIHZiQ3JpdGljYWwNCiAgICBXU2NyaXB0LlF1aXQgMQ0K" & _
        "RW5kIElmDQoNCm1vZG8gPSBhcmdzKDApDQphcmNoaXZvID0gYXJncygxKQ0KZGVzdGlubyA9" & _
        "IGFyZ3MoMikNCm5vbWJyZSA9IGFyZ3MoMykNCg0KcnV0YUZpbmFsID0gZGVzdGlubyAmICJc" & _
        "IiAmIG5vbWJyZSAmICIueGxhbSINCg0KJyBFc3BlcmFyIGEgcXVlIEV4Y2VsIGxpYmVyZSBs" & _
        "b3MgYXJjaGl2b3MNCldTY3JpcHQuU2xlZXAgNDAwMA0KDQpJZiBtb2RvID0gIi9pbnN0YWxs"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "IiBUaGVuDQogICAgRG9JbnN0YWxsDQpFbHNlSWYgbW9kbyA9ICIvdW5pbnN0YWxsIiBUaGVu" & _
        "DQogICAgRG9Vbmluc3RhbGwNCkVsc2UNCiAgICBNc2dCb3ggIk1vZG8gZGUgaW5zdGFsYWNp" & _
        "824gbm8gcmVjb25vY2lkbzogIiAmIG1vZG8gJiAiLCBsYSBpbnN0YWxhY2nzbiBubyBzZSBw" & _
        "dWVkZSBjb21wbGV0YXIiLCB2YkNyaXRpY2FsDQogICAgV1NjcmlwdC5RdWl0IDENCkVuZCBJ" & _
        "Zg0KDQonIExpbXBpYXI6IGVsaW1pbmFyIGVzdGUgc2NyaXB0DQpPbiBFcnJvciBSZXN1bWUg" & _
        "TmV4dA0KZnNvLkRlbGV0ZUZpbGUgV1NjcmlwdC5TY3JpcHRGdWxsTmFtZQ0KT24gRXJyb3Ig" & _
        "R29UbyAwDQoNCldTY3JpcHQuUXVpdCAwDQoNCicgPT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0NCicgSU5TVEFMQUNJ004NCicgPT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0NClN1YiBE" & _
        "b0luc3RhbGwoKQ0KICAgIElmIE5vdCBmc28uRmlsZUV4aXN0cyhhcmNoaXZvKSBUaGVuDQog" & _
        "ICAgICAgIE1zZ0JveCAiRXJyb3IgZGUgaW5zdGFsYWNp8246IG5vIGV4aXN0ZSAnIiAmIGFy" & _
        "Y2hpdm8gJiAiJyIsIHZiQ3JpdGljYWwNCiAgICAgICAgV1NjcmlwdC5RdWl0IDENCiAgICBF" & _
        "bmQgSWYNCg0KICAgICcgMS4gRWxpbWluYXIgWExBTSBhbnRlcmlvciBzaSBleGlzdGUNCiAg" & _
        "ICBSZW1vdmVBZGRpbkluRGVzdGlubyBydXRhRmluYWwNCg0KICAgICcgMi4gRXh0cmFlciBD" & _
        "T00gZGVsIFhMQU0gb3JpZ2VuIEFOVEVTIGRlIGNvcGlhcg0KICAgICcgICAgKHBvcnF1ZSBk" & _
        "ZXNwdelzIGRlIGNvcGlhciBlbCBYTEFNIGVzdGFy4SBlbiB1c28gcG9yIEV4Y2VsKQ0KICAg" & _
        "IElmIE5vdCBFeHRyYWN0Q09NRnJvbVhMQU0oYXJjaGl2bywgZGVzdGlubykgVGhlbg0KICAg" & _
        "ICAgICAnIFNpIGZhbGxhIGxhIGV4dHJhY2Np824gZGVsIENPTSwgY29udGludWFyIGRlIHRv" & _
        "ZG9zIG1vZG9zDQogICAgICAgICcgRWwgY29tcGxlbWVudG8gZnVuY2lvbmFy4SBwZXJvIHNp" & _
        "biBGb2xkZXJXYXRjaGVyDQogICAgICAgIFdTY3JpcHQuRWNobyAiQWR2ZXJ0ZW5jaWE6IE5v"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "IHNlIHB1ZG8gZXh0cmFlciBlbCBjb21wb25lbnRlIENPTSBkZWwgWExBTS4gTGEgdmlnaWxh" & _
        "bmNpYSBkZSBjYXJwZXRhcyBubyBlc3RhcuEgZGlzcG9uaWJsZS4iDQogICAgRW5kIElmDQog" & _
        "ICAgDQogICAgJyAzLiBJbnNlcnRhciBjbGF2ZXMgcGFyYSByZWdpc3RybyBkZWwgY29tcG9u" & _
        "ZW50ZSBjb20gRW4gSEtjVSAgDQogICAgUmVnaXN0cmFyQ2xhdmVzQ09NKCkNCg0KICAgICcg" & _
        "NC4gQ29waWFyIFhMQU0gYWwgZGVzdGlubw0KICAgIGZzby5Db3B5RmlsZSBhcmNoaXZvLCBy" & _
        "dXRhRmluYWwsIFRydWUNCg0KICAgICcgNS4gUmVnaXN0cmFyIGVuIEV4Y2VsDQogICAgU2V0" & _
        "IGV4Y2VsID0gQ3JlYXRlT2JqZWN0KCJFeGNlbC5BcHBsaWNhdGlvbiIpDQogICAgZXhjZWwu" & _
        "VmlzaWJsZSA9IEZhbHNlDQoNCiAgICBGb3IgRWFjaCBhaSBJbiBleGNlbC5BZGRJbnMNCiAg" & _
        "ICAgICAgSWYgTENhc2UoYWkuTmFtZSkgPSBMQ2FzZShub21icmUgJiAiLnhsYW0iKSBUaGVu" & _
        "DQogICAgICAgICAgICBhaS5JbnN0YWxsZWQgPSBUcnVlDQogICAgICAgICAgICBFeGl0IEZv" & _
        "cg0KICAgICAgICBFbmQgSWYNCiAgICBOZXh0DQoNCiAgICBXU2NyaXB0LlNsZWVwIDEwMDAN" & _
        "Cg0KICAgIElmIGFpIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBNc2dCb3ggIk5vIGhhIHNp" & _
        "ZG8gcG9zaWJsZSBjb21wbGV0YXIgbGEgaW5zdGFsYWNp824uIFBvciBmYXZvciwgaGFiaWxp" & _
        "dGEgZWwgY29tcGxlbWVudG8gZGVzZGUgZWwgbWVu+iBkZSBjb21wbGVtZW50b3MgZGUgRXhj" & _
        "ZWwuIiwgdmJDcml0aWNhbA0KICAgIEVsc2VJZiBOb3QgYWkuSW5zdGFsbGVkIFRoZW4NCiAg" & _
        "ICAgICAgTXNnQm94ICJObyBoYSBzaWRvIHBvc2libGUgY29tcGxldGFyIGxhIGluc3RhbGFj" & _
        "afNuLiBQb3IgZmF2b3IsIGhhYmlsaXRhIGVsIGNvbXBsZW1lbnRvIGRlc2RlIGVsIG1lbvog" & _
        "ZGUgY29tcGxlbWVudG9zIGRlIEV4Y2VsLiIsIHZiQ3JpdGljYWwNCiAgICBFbHNlDQogICAg" & _
        "ICAgIE1zZ0JveCAiSW5zdGFsYWNp824gY29tcGxldGFkYSwgcmVpbmljaWEgRXhjZWwuIiwg" & _
        "dmJJbmZvcm1hdGlvbg0KICAgIEVuZCBJZg0KDQogICAgZXhjZWwuUXVpdA0KICAgIFNldCBl"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "eGNlbCA9IE5vdGhpbmcNCkVuZCBTdWINCg0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KJyBERVNJTlNUQUxBQ0nTTg0KJyA9PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KU3Vi" & _
        "IERvVW5pbnN0YWxsKCkNCiAgICAnIDEuIEVsaW1pbmFyIGFyY2hpdm9zIENPTSBwcmltZXJv" & _
        "IChhbnRlcyBkZSBxdWUgRXhjZWwgbG9zIGJsb3F1ZWUpDQogICAgUmVtb3ZlQ09NRmlsZXMg" & _
        "ZGVzdGlubw0KDQogICAgJyAyLiBFbGltaW5hciBYTEFNDQogICAgUmVtb3ZlQWRkaW5JbkRl" & _
        "c3Rpbm8gcnV0YUZpbmFsDQogICAgDQogICAgJyAzLiBFbGltaW5hciBjbGF2ZXMgZGUgcmVn" & _
        "aXN0cm8gZGVsIGNvbXBvbmVudGUgY29tIEVuIEhLY1UgIA0KICAgIEVsaW1pbmFyQ2xhdmVz" & _
        "Q09NKCkNCg0KICAgICcgNC4gRGVzcmVnaXN0cmFyIGRlIEV4Y2VsDQogICAgU2V0IGV4Y2Vs" & _
        "ID0gQ3JlYXRlT2JqZWN0KCJFeGNlbC5BcHBsaWNhdGlvbiIpDQogICAgdmVycyA9IGV4Y2Vs" & _
        "LkFwcGxpY2F0aW9uLlZlcnNpb24NCiAgICBleGNlbC5WaXNpYmxlID0gRmFsc2UNCg0KICAg" & _
        "IEZvciBFYWNoIGFpIEluIGV4Y2VsLkFkZElucw0KICAgICAgICBJZiBMQ2FzZShhaS5OYW1l" & _
        "KSA9IExDYXNlKG5vbWJyZSAmICIueGxhbSIpIFRoZW4NCiAgICAgICAgICAgIGFpLkluc3Rh" & _
        "bGxlZCA9IEZhbHNlDQogICAgICAgICAgICBFeGl0IEZvcg0KICAgICAgICBFbmQgSWYNCiAg" & _
        "ICBOZXh0DQoNCiAgICBEaW0gdW5pbnN0YWxsT0sNCiAgICB1bmluc3RhbGxPSyA9IFRydWUN" & _
        "CiAgICBJZiBOb3QgYWkgSXMgTm90aGluZyBUaGVuDQogICAgICAgIElmIGFpLkluc3RhbGxl" & _
        "ZCBUaGVuIHVuaW5zdGFsbE9LID0gRmFsc2UNCiAgICBFbmQgSWYNCg0KICAgIElmIE5vdCB1" & _
        "bmluc3RhbGxPSyBUaGVuDQogICAgICAgIE1zZ0JveCAiTm8gaGEgc2lkbyBwb3NpYmxlIGNv" & _
        "bXBsZXRhciBsYSBkZXNpbnN0YWxhY2nzbi4gUG9yIGZhdm9yLCByZWludOludGFsbyBkZSBu" & _
        "dWV2byBvIGRlc2hhYmlsaXRhIGVsIGNvbXBsZW1lbnRvIGRlc2RlIGVsIG1lbvogZGUgY29t"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "cGxlbWVudG9zIGRlIEV4Y2VsLiIsIHZiQ3JpdGljYWwNCiAgICBFbHNlDQogICAgICAgIE1z" & _
        "Z0JveCAiRGVzaW5zdGFsYWNp824gY29tcGxldGFkYSwgcmVpbmljaWEgRXhjZWwuIiwgdmJJ" & _
        "bmZvcm1hdGlvbg0KICAgIEVuZCBJZg0KDQogICAgZXhjZWwuUXVpdA0KICAgIFNldCBleGNl" & _
        "bCA9IE5vdGhpbmcNCg0KICAgICcgNS4gTGltcGlhciBjbGF2ZXMgZGUgY29uZmlndXJhY2nz" & _
        "biBkZWwgWExBTSBlbiBlbCByZWdpc3RybyANCiAgICBDbGVhblJlZ2lzdHJ5IHZlcnMsIG5v" & _
        "bWJyZQ0KRW5kIFN1Yg0KDQonID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09DQonIEVYVFJBQ0NJ004gREVMIENPTSBERVNERSBFTCBYTEFN" & _
        "IChaSVApDQonID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09DQpGdW5jdGlvbiBFeHRyYWN0Q09NRnJvbVhMQU0oeGxhbVBhdGgsIGRlc3RG" & _
        "b2xkZXIpDQogICAgRXh0cmFjdENPTUZyb21YTEFNID0gRmFsc2UNCg0KICAgIE9uIEVycm9y" & _
        "IFJlc3VtZSBOZXh0DQoNCiAgICAnIEludGVudGFyIHByaW1lcm8gY29uIDd6aXAgKG3hcyBy" & _
        "4XBpZG8geSBmaWFibGUpDQogICAgSWYgVHJ5RXh0cmFjdFdpdGg3WmlwKHhsYW1QYXRoLCBk" & _
        "ZXN0Rm9sZGVyKSBUaGVuDQogICAgICAgIEV4dHJhY3RDT01Gcm9tWExBTSA9IFRydWUNCiAg" & _
        "ICAgICAgRXhpdCBGdW5jdGlvbg0KICAgIEVuZCBJZg0KDQogICAgJyBTaSBubyBoYXkgN3pp" & _
        "cCwgdXNhciBTaGVsbC5BcHBsaWNhdGlvbiAoV2luZG93cyBuYXRpdm8pDQogICAgSWYgVHJ5" & _
        "RXh0cmFjdFdpdGhTaGVsbCh4bGFtUGF0aCwgZGVzdEZvbGRlcikgVGhlbg0KICAgICAgICBF" & _
        "eHRyYWN0Q09NRnJvbVhMQU0gPSBUcnVlDQogICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBF" & _
        "bmQgSWYNCg0KICAgIE9uIEVycm9yIEdvVG8gMA0KRW5kIEZ1bmN0aW9uDQoNCicgRXh0cmFj" & _
        "Y2nzbiB1c2FuZG8gNy1aaXANCkZ1bmN0aW9uIFRyeUV4dHJhY3RXaXRoN1ppcCh4bGFtUGF0" & _
        "aCwgZGVzdEZvbGRlcikNCiAgICBUcnlFeHRyYWN0V2l0aDdaaXAgPSBGYWxzZQ0KDQogICAg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "RGltIHNoZWxsLCBleGVjLCBzZXZlblppcFBhdGgNCiAgICBTZXQgc2hlbGwgPSBDcmVhdGVP" & _
        "YmplY3QoIldTY3JpcHQuU2hlbGwiKQ0KDQogICAgJyBCdXNjYXIgN3ouZXhlIGVuIGVsIFBB" & _
        "VEgNCiAgICBPbiBFcnJvciBSZXN1bWUgTmV4dA0KICAgIFNldCBleGVjID0gc2hlbGwuRXhl" & _
        "Yygid2hlcmUgN3ouZXhlIikNCiAgICBJZiBFcnIuTnVtYmVyID0gMCBUaGVuDQogICAgICAg" & _
        "IERvIFdoaWxlIGV4ZWMuU3RhdHVzID0gMA0KICAgICAgICAgICAgV1NjcmlwdC5TbGVlcCAx" & _
        "MDANCiAgICAgICAgTG9vcA0KICAgICAgICBzZXZlblppcFBhdGggPSBUcmltKGV4ZWMuU3Rk" & _
        "T3V0LlJlYWRMaW5lKQ0KICAgIEVuZCBJZg0KICAgIE9uIEVycm9yIEdvVG8gMA0KDQogICAg" & _
        "SWYgc2V2ZW5aaXBQYXRoID0gIiIgT3IgTm90IGZzby5GaWxlRXhpc3RzKHNldmVuWmlwUGF0" & _
        "aCkgVGhlbg0KICAgICAgICAnIDd6aXAgbm8gZW5jb250cmFkbw0KICAgICAgICBFeGl0IEZ1" & _
        "bmN0aW9uDQogICAgRW5kIElmDQoNCiAgICAnIEV4dHJhZXIgc29sbyBsb3MgYXJjaGl2b3Mg" & _
        "Q09NDQogICAgRGltIGNtZCwgZGxsUGF0aCwgbWFuaWZlc3RQYXRoLCBjb25maWdQYXRoDQog" & _
        "ICAgZGxsUGF0aCA9IENPTV9FTUJFRF9QQVRIICYgQ09NX0RMTF9OQU1FDQogICAgbWFuaWZl" & _
        "c3RQYXRoID0gQ09NX0VNQkVEX1BBVEggJiBDT01fTUFOSUZFU1RfTkFNRQ0KICAgIGNvbmZp" & _
        "Z1BhdGggPSBDT01fRU1CRURfUEFUSCAmIENPTV9DT05GSUdfTkFNRQ0KDQogICAgJyBFeHRy" & _
        "YWVyIERMTA0KICAgIGNtZCA9ICIiIiIgJiBzZXZlblppcFBhdGggJiAiIiIgZSAiIiIgJiB4" & _
        "bGFtUGF0aCAmICIiIiAtbyIiIiAmIGRlc3RGb2xkZXIgJiAiIiIgIiIiICYgZGxsUGF0aCAm" & _
        "ICIiIiAteSINCiAgICBzaGVsbC5SdW4gY21kLCAwLCBUcnVlDQoNCiAgICAnIEV4dHJhZXIg" & _
        "TWFuaWZlc3QNCiAgICBjbWQgPSAiIiIiICYgc2V2ZW5aaXBQYXRoICYgIiIiIGUgIiIiICYg" & _
        "eGxhbVBhdGggJiAiIiIgLW8iIiIgJiBkZXN0Rm9sZGVyICYgIiIiICIiIiAmIG1hbmlmZXN0" & _
        "UGF0aCAmICIiIiAteSINCiAgICBzaGVsbC5SdW4gY21kLCAwLCBUcnVlDQoNCiAgICAnIEV4"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "dHJhZXIgQ29uZmlnDQogICAgY21kID0gIiIiIiAmIHNldmVuWmlwUGF0aCAmICIiIiBlICIi" & _
        "IiAmIHhsYW1QYXRoICYgIiIiIC1vIiIiICYgZGVzdEZvbGRlciAmICIiIiAiIiIgJiBjb25m" & _
        "aWdQYXRoICYgIiIiIC15Ig0KICAgIHNoZWxsLlJ1biBjbWQsIDAsIFRydWUNCg0KICAgICcg" & _
        "VmVyaWZpY2FyIHF1ZSBzZSBleHRyYWplcm9uDQogICAgSWYgZnNvLkZpbGVFeGlzdHMoZGVz" & _
        "dEZvbGRlciAmICJcIiAmIENPTV9ETExfTkFNRSkgQW5kIF8NCiAgICAgICBmc28uRmlsZUV4" & _
        "aXN0cyhkZXN0Rm9sZGVyICYgIlwiICYgQ09NX0NPTkZJR19OQU1FKSBBbmQgXw0KICAgICAg" & _
        "IGZzby5GaWxlRXhpc3RzKGRlc3RGb2xkZXIgJiAiXCIgJiBDT01fTUFOSUZFU1RfTkFNRSkg" & _
        "VGhlbg0KICAgICAgICBUcnlFeHRyYWN0V2l0aDdaaXAgPSBUcnVlDQogICAgRW5kIElmDQoN" & _
        "CiAgICBTZXQgc2hlbGwgPSBOb3RoaW5nDQpFbmQgRnVuY3Rpb24NCg0KJyBFeHRyYWNjafNu" & _
        "IHVzYW5kbyBTaGVsbC5BcHBsaWNhdGlvbiAoV2luZG93cyBuYXRpdm8pDQpGdW5jdGlvbiBU" & _
        "cnlFeHRyYWN0V2l0aFNoZWxsKHhsYW1QYXRoLCBkZXN0Rm9sZGVyKQ0KICAgIFRyeUV4dHJh" & _
        "Y3RXaXRoU2hlbGwgPSBGYWxzZQ0KDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCg0KICAg" & _
        "ICcgQ3JlYXIgY29waWEgdGVtcG9yYWwgY29tbyAuemlwDQogICAgRGltIHRlbXBaaXANCiAg" & _
        "ICB0ZW1wWmlwID0gZnNvLkdldFNwZWNpYWxGb2xkZXIoMikgJiAiXCIgJiBmc28uR2V0VGVt" & _
        "cE5hbWUoKSAmICIuemlwIg0KICAgIGZzby5Db3B5RmlsZSB4bGFtUGF0aCwgdGVtcFppcCwg" & _
        "VHJ1ZQ0KDQogICAgSWYgRXJyLk51bWJlciA8PiAwIFRoZW4gRXhpdCBGdW5jdGlvbg0KDQog" & _
        "ICAgJyBVc2FyIFNoZWxsLkFwcGxpY2F0aW9uIHBhcmEgZXhwbG9yYXIgZWwgWklQDQogICAg" & _
        "RGltIHNoZWxsLCB6aXBGb2xkZXIsIGRlc3RGb2xkZXJPYmoNCiAgICBTZXQgc2hlbGwgPSBD" & _
        "cmVhdGVPYmplY3QoIlNoZWxsLkFwcGxpY2F0aW9uIikNCiAgICBTZXQgemlwRm9sZGVyID0g" & _
        "c2hlbGwuTmFtZVNwYWNlKHRlbXBaaXApDQogICAgU2V0IGRlc3RGb2xkZXJPYmogPSBzaGVs"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "bC5OYW1lU3BhY2UoZGVzdEZvbGRlcikNCg0KICAgIElmIHppcEZvbGRlciBJcyBOb3RoaW5n" & _
        "IE9yIGRlc3RGb2xkZXJPYmogSXMgTm90aGluZyBUaGVuDQogICAgICAgIGZzby5EZWxldGVG" & _
        "aWxlIHRlbXBaaXANCiAgICAgICAgRXhpdCBGdW5jdGlvbg0KICAgIEVuZCBJZg0KDQogICAg" & _
        "JyBCdXNjYXIgbGEgY2FycGV0YSB4bFxlbWJlZGRpbmdzIGRlbnRybyBkZWwgWklQDQogICAg" & _
        "RGltIGl0ZW0sIGVtYmVkRm9sZGVyDQogICAgU2V0IGVtYmVkRm9sZGVyID0gTm90aGluZw0K" & _
        "DQogICAgJyBOYXZlZ2FyIGEgeGxcZW1iZWRkaW5ncw0KICAgIERpbSB4bEZvbGRlcg0KICAg" & _
        "IEZvciBFYWNoIGl0ZW0gSW4gemlwRm9sZGVyLkl0ZW1zDQogICAgICAgIElmIExDYXNlKGl0" & _
        "ZW0uTmFtZSkgPSAieGwiIFRoZW4NCiAgICAgICAgICAgIFNldCB4bEZvbGRlciA9IHNoZWxs" & _
        "Lk5hbWVTcGFjZShpdGVtLlBhdGgpDQogICAgICAgICAgICBFeGl0IEZvcg0KICAgICAgICBF" & _
        "bmQgSWYNCiAgICBOZXh0DQoNCiAgICBJZiB4bEZvbGRlciBJcyBOb3RoaW5nIFRoZW4NCiAg" & _
        "ICAgICAgZnNvLkRlbGV0ZUZpbGUgdGVtcFppcA0KICAgICAgICBFeGl0IEZ1bmN0aW9uDQog" & _
        "ICAgRW5kIElmDQoNCiAgICBGb3IgRWFjaCBpdGVtIEluIHhsRm9sZGVyLkl0ZW1zDQogICAg" & _
        "ICAgIElmIExDYXNlKGl0ZW0uTmFtZSkgPSAiZW1iZWRkaW5ncyIgVGhlbg0KICAgICAgICAg" & _
        "ICAgU2V0IGVtYmVkRm9sZGVyID0gc2hlbGwuTmFtZVNwYWNlKGl0ZW0uUGF0aCkNCiAgICAg" & _
        "ICAgICAgIEV4aXQgRm9yDQogICAgICAgIEVuZCBJZg0KICAgIE5leHQNCg0KICAgIElmIGVt" & _
        "YmVkRm9sZGVyIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBmc28uRGVsZXRlRmlsZSB0ZW1w" & _
        "WmlwDQogICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAgICcgRXh0cmFl" & _
        "ciBsb3MgYXJjaGl2b3MgQ09NDQogICAgRGltIGRsbEl0ZW0sIG1hbmlmZXN0SXRlbSwgY29u" & _
        "ZmlnSXRlbQ0KICAgIEZvciBFYWNoIGl0ZW0gSW4gZW1iZWRGb2xkZXIuSXRlbXMNCiAgICAg" & _
        "ICAgSWYgTENhc2UoaXRlbS5OYW1lKSA9IExDYXNlKENPTV9ETExfTkFNRSkgVGhlbg0KICAg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ICAgICAgICAgU2V0IGRsbEl0ZW0gPSBpdGVtDQogICAgICAgIEVsc2VJZiBMQ2FzZShpdGVt" & _
        "Lk5hbWUpID0gTENhc2UoQ09NX01BTklGRVNUX05BTUUpIFRoZW4NCiAgICAgICAgICAgIFNl" & _
        "dCBtYW5pZmVzdEl0ZW0gPSBpdGVtDQogICAgICAgIEVsc2VJZiBMQ2FzZShpdGVtLk5hbWUp" & _
        "ID0gTENhc2UoQ09NX0NPTkZJR19OQU1FKSBUaGVuDQogICAgICAgICAgICBTZXQgY29uZmln" & _
        "SXRlbSA9IGl0ZW0NCiAgICAgICAgRW5kIElmDQogICAgTmV4dA0KDQogICAgJyBDb3BpYXIg" & _
        "YXJjaGl2b3MgYWwgZGVzdGlubyAoMTYgPSBObyBtb3N0cmFyIGRp4WxvZ28sIDEwMjQgPSBO" & _
        "byBjb25maXJtYXIpDQogICAgSWYgTm90IGRsbEl0ZW0gSXMgTm90aGluZyBUaGVuDQogICAg" & _
        "ICAgIGRlc3RGb2xkZXJPYmouQ29weUhlcmUgZGxsSXRlbSwgMTYgKyAxMDI0DQogICAgICAg" & _
        "IFdTY3JpcHQuU2xlZXAgNTAwDQogICAgRW5kIElmDQoNCiAgICBJZiBOb3QgbWFuaWZlc3RJ" & _
        "dGVtIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBkZXN0Rm9sZGVyT2JqLkNvcHlIZXJlIG1h" & _
        "bmlmZXN0SXRlbSwgMTYgKyAxMDI0DQogICAgICAgIFdTY3JpcHQuU2xlZXAgNTAwDQogICAg" & _
        "RW5kIElmDQoNCiAgICBJZiBOb3QgY29uZmlnSXRlbSBJcyBOb3RoaW5nIFRoZW4NCiAgICAg" & _
        "ICAgZGVzdEZvbGRlck9iai5Db3B5SGVyZSBjb25maWdJdGVtLCAxNiArIDEwMjQNCiAgICAg" & _
        "ICAgV1NjcmlwdC5TbGVlcCA1MDANCiAgICBFbmQgSWYNCg0KICAgICcgTGltcGlhcg0KICAg" & _
        "IGZzby5EZWxldGVGaWxlIHRlbXBaaXANCg0KICAgICcgVmVyaWZpY2FyDQogICAgSWYgZnNv" & _
        "LkZpbGVFeGlzdHMoZGVzdEZvbGRlciAmICJcIiAmIENPTV9ETExfTkFNRSkgQW5kIF8NCiAg" & _
        "ICAgICBmc28uRmlsZUV4aXN0cyhkZXN0Rm9sZGVyICYgIlwiICYgQ09NX0NPTkZJR19OQU1F" & _
        "KSBBbmQgXw0KICAgICAgIGZzby5GaWxlRXhpc3RzKGRlc3RGb2xkZXIgJiAiXCIgJiBDT01f" & _
        "TUFOSUZFU1RfTkFNRSkgVGhlbg0KICAgICAgICBUcnlFeHRyYWN0V2l0aFNoZWxsID0gVHJ1" & _
        "ZQ0KICAgIEVuZCBJZg0KDQogICAgT24gRXJyb3IgR29UbyAwDQpFbmQgRnVuY3Rpb24NCg0K"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "JyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PQ0KJyBFTElNSU5BQ0nTTiBERSBBUkNISVZPUyBDT00NCicgPT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0NClN1YiBSZW1vdmVDT01GaWxl" & _
        "cyhmb2xkZXIpDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCg0KICAgIERpbSBkbGxQYXRo" & _
        "LCBtYW5pZmVzdFBhdGgsIGNvbmZpZ1BhdGgNCiAgICBkbGxQYXRoID0gZm9sZGVyICYgIlwi" & _
        "ICYgQ09NX0RMTF9OQU1FDQogICAgbWFuaWZlc3RQYXRoID0gZm9sZGVyICYgIlwiICYgQ09N" & _
        "X01BTklGRVNUX05BTUUNCiAgICBjb25maWdQYXRoID0gZm9sZGVyICYgIlwiICYgQ09NX0NP" & _
        "TkZJR19OQU1FDQoNCiAgICBJZiBmc28uRmlsZUV4aXN0cyhkbGxQYXRoKSBUaGVuDQogICAg" & _
        "ICAgIGZzby5EZWxldGVGaWxlIGRsbFBhdGgsIFRydWUNCiAgICBFbmQgSWYNCg0KICAgIElm" & _
        "IGZzby5GaWxlRXhpc3RzKG1hbmlmZXN0UGF0aCkgVGhlbg0KICAgICAgICBmc28uRGVsZXRl" & _
        "RmlsZSBtYW5pZmVzdFBhdGgsIFRydWUNCiAgICBFbmQgSWYNCg0KICAgIElmIGZzby5GaWxl" & _
        "RXhpc3RzKGNvbmZpZ1BhdGgpIFRoZW4NCiAgICAgICAgZnNvLkRlbGV0ZUZpbGUgY29uZmln" & _
        "UGF0aCwgVHJ1ZQ0KICAgIEVuZCBJZg0KDQogICAgT24gRXJyb3IgR29UbyAwDQpFbmQgU3Vi" & _
        "DQoNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT0NCicgRUxJTUlOQUNJ004gREVMIFhMQU0gRVhJU1RFTlRFDQonID09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQpTdWIgUmVtb3Zl" & _
        "QWRkaW5JbkRlc3Rpbm8ocnV0YUZpbmFsKQ0KICAgIElmIE5vdCBmc28uRmlsZUV4aXN0cyhy" & _
        "dXRhRmluYWwpIFRoZW4gRXhpdCBTdWINCg0KICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQog" & _
        "ICAgZnNvLkRlbGV0ZUZpbGUgcnV0YUZpbmFsLCBUcnVlDQogICAgT24gRXJyb3IgR29UbyAw" & _
        "DQoNCiAgICBJZiBOb3QgZnNvLkZpbGVFeGlzdHMocnV0YUZpbmFsKSBUaGVuIEV4aXQgU3Vi"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "DQoNCiAgICAnIEVsIGFyY2hpdm8gc2lndWUgZXhpc3RpZW5kbywgcG9zaWJsZW1lbnRlIGJs" & _
        "b3F1ZWFkbw0KICAgIERpbSBvYmpXTUlTZXJ2aWNlLCBjb2xQcm9jZXNzZXMsIGFuc3dlciwg" & _
        "b2JqUHJvY2Vzcw0KICAgIFNldCBvYmpXTUlTZXJ2aWNlID0gR2V0T2JqZWN0KCJ3aW5tZ210" & _
        "czpcXC5ccm9vdFxjaW12MiIpDQogICAgU2V0IGNvbFByb2Nlc3NlcyA9IG9ialdNSVNlcnZp" & _
        "Y2UuRXhlY1F1ZXJ5KCJTZWxlY3QgKiBmcm9tIFdpbjMyX1Byb2Nlc3MgV2hlcmUgTmFtZSA9" & _
        "ICdFWENFTC5FWEUnIikNCg0KICAgIElmIGNvbFByb2Nlc3Nlcy5Db3VudCA+IDAgVGhlbg0K" & _
        "ICAgICAgICBhbnN3ZXIgPSBNc2dCb3goIkV4Y2VsIGVzdOEgZW4gZWplY3VjafNuIHkgcHVl" & _
        "ZGUgZXN0YXIgYmxvcXVlYW5kbyBlbCBhcmNoaXZvIGRlbCBjb21wbGVtZW50byBlbiBkZXN0" & _
        "aW5vLiC/RGVzZWFzIGNlcnJhciBFeGNlbD8iLCB2Ylllc05vICsgdmJRdWVzdGlvbikNCiAg" & _
        "ICAgICAgSWYgYW5zd2VyID0gdmJZZXMgVGhlbg0KICAgICAgICAgICAgRm9yIEVhY2ggb2Jq" & _
        "UHJvY2VzcyBJbiBjb2xQcm9jZXNzZXMNCiAgICAgICAgICAgICAgICBvYmpQcm9jZXNzLlRl" & _
        "cm1pbmF0ZQ0KICAgICAgICAgICAgTmV4dA0KDQogICAgICAgICAgICAnIEVzcGVyYXIgYSBx" & _
        "dWUgRXhjZWwgY2llcnJlDQogICAgICAgICAgICBXU2NyaXB0LlNsZWVwIDMwMDANCg0KICAg" & _
        "ICAgICAgICAgJyBSZWludGVudGFyIGVsaW1pbmFyDQogICAgICAgICAgICBPbiBFcnJvciBS" & _
        "ZXN1bWUgTmV4dA0KICAgICAgICAgICAgZnNvLkRlbGV0ZUZpbGUgcnV0YUZpbmFsLCBUcnVl" & _
        "DQogICAgICAgICAgICBPbiBFcnJvciBHb1RvIDANCg0KICAgICAgICAgICAgSWYgZnNvLkZp" & _
        "bGVFeGlzdHMocnV0YUZpbmFsKSBUaGVuDQogICAgICAgICAgICAgICAgTXNnQm94ICJObyBo" & _
        "YSBzaWRvIHBvc2libGUgY29tcGxldGFyIGVsIHByb2Nlc28uIFBvciBmYXZvciwgY2llcnJh" & _
        "IEV4Y2VsIG1hbnVhbG1lbnRlIHkgZWxpbWluYSBlbCBmaWNoZXJvIiAmIHZiQ3IgJiAiJyIg" & _
        "JiBydXRhRmluYWwgJiAiJy4iLCB2YkNyaXRpY2FsDQogICAgICAgICAgICAgICAgV1Njcmlw"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "dC5RdWl0IDENCiAgICAgICAgICAgIEVuZCBJZg0KICAgICAgICBFbHNlDQogICAgICAgICAg" & _
        "ICBNc2dCb3ggIk5vIGVzIHBvc2libGUgY29tcGxldGFyIGVsIHByb2Nlc28uIFBvciBmYXZv" & _
        "ciwgY2llcnJhIEV4Y2VsIG1hbnVhbG1lbnRlIGUgaW506W50YWxvIGRlIG51ZXZvLiIsIHZi" & _
        "Q3JpdGljYWwNCiAgICAgICAgICAgIFdTY3JpcHQuUXVpdCAxDQogICAgICAgIEVuZCBJZg0K" & _
        "ICAgIEVuZCBJZg0KRW5kIFN1Yg0KDQonID09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09DQonIExJTVBJRVpBIERFTCBSRUdJU1RSTw0KJyA9" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0K" & _
        "U3ViIENsZWFuUmVnaXN0cnkodmVycywgbm9tYnJlKQ0KICAgIE9uIEVycm9yIFJlc3VtZSBO" & _
        "ZXh0DQoNCiAgICBEaW0gV3NoU2hlbGwsIGksIGNsYXZlLCB2YWxvcg0KICAgIFNldCBXc2hT" & _
        "aGVsbCA9IENyZWF0ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQoNCiAgICBGb3IgaSA9IDEg" & _
        "VG8gNTANCiAgICAgICAgY2xhdmUgPSAiSEtFWV9DVVJSRU5UX1VTRVJcU29mdHdhcmVcTWlj" & _
        "cm9zb2Z0XE9mZmljZVwiICYgdmVycyAmICJcRXhjZWxcT3B0aW9uc1xPUEVOIiAmIGkNCiAg" & _
        "ICAgICAgdmFsb3IgPSBXc2hTaGVsbC5SZWdSZWFkKGNsYXZlKQ0KDQogICAgICAgIElmIEVy" & _
        "ci5OdW1iZXIgPSAwIFRoZW4NCiAgICAgICAgICAgIElmIEluU3RyKDEsIHZhbG9yLCBub21i" & _
        "cmUgJiAiLnhsYW0iLCB2YlRleHRDb21wYXJlKSA+IDAgVGhlbg0KICAgICAgICAgICAgICAg" & _
        "IFdzaFNoZWxsLlJlZ0RlbGV0ZSBjbGF2ZQ0KICAgICAgICAgICAgICAgIEV4aXQgRm9yDQog" & _
        "ICAgICAgICAgICBFbmQgSWYNCiAgICAgICAgRWxzZQ0KICAgICAgICAgICAgRXJyLkNsZWFy" & _
        "DQogICAgICAgICAgICBFeGl0IEZvcg0KICAgICAgICBFbmQgSWYNCiAgICBOZXh0DQoNCiAg" & _
        "ICBTZXQgV3NoU2hlbGwgPSBOb3RoaW5nDQogICAgT24gRXJyb3IgR29UbyAwDQpFbmQgU3Vi" & _
        "DQoNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "PT09PT0NCicgIFNVQlJVVElOQSBQQVJBIFJFR0lTVFJBUiBDT01QT05FTlRFIENPTSANCicg" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0N" & _
        "ClN1YiBSZWdpc3RyYXJDbGF2ZXNDT00oKQ0KICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQog" & _
        "ICAgRGltIHNoZWxsLCBhcHBEYXRhUGF0aCwgYWRkaW5zUGF0aA0KICAgIA0KICAgIFNldCBz" & _
        "aGVsbCA9IENyZWF0ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQogICAgDQogICAgJyBPYnRl" & _
        "bmVyIHJ1dGEgZGVsIEFwcERhdGEgZGVsIHVzdWFyaW8gYWN0dWFsDQogICAgYXBwRGF0YVBh" & _
        "dGggPSBzaGVsbC5FeHBhbmRFbnZpcm9ubWVudFN0cmluZ3MoIiVBUFBEQVRBJSIpDQogICAg" & _
        "YWRkaW5zUGF0aCA9IGZzby5CdWlsZFBhdGgoYXBwRGF0YVBhdGgsICJNaWNyb3NvZnRcQWRk" & _
        "SW5zXCIpDQogICAgDQogICAgJyBDcmVhciBsYXMgY2xhdmVzIHByaW5jaXBhbGVzDQogICAg" & _
        "JyAxLiBDTFNJRCBwcmluY2lwYWwNCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FS" & _
        "RVxDbGFzc2VzXENMU0lEXCIgJiBHVUlEX0NMU0lEICYgIlwiLCBDTEFTU19OQU1FLCAiUkVH" & _
        "X1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURc" & _
        "IiAmIEdVSURfQ0xTSUQgJiAiXFByb2dJZFwiLCBQUk9HX0lELCAiUkVHX1NaIg0KICAgIA0K" & _
        "ICAgICcgMi4gSW1wbGVtZW50ZWQgQ2F0ZWdvcmllcw0KICAgIHNoZWxsLlJlZ1dyaXRlICJI" & _
        "S0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXEltcGxlbWVu" & _
        "dGVkIENhdGVnb3JpZXNcIiwgIiIsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhL" & _
        "Q1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW1wbGVtZW50" & _
        "ZWQgQ2F0ZWdvcmllc1x7NjJDOEZFNjUtNEVCQi00NWU3LUI0NDAtNkUzOUIyQ0RCRjI5fVwi" & _
        "LCAiIiwgIlJFR19TWiINCiAgICANCiAgICAnIDMuIElucHJvY1NlcnZlcjMyDQogICAgc2hl" & _
        "bGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJ"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "RCAmICJcSW5wcm9jU2VydmVyMzJcIiwgIm1zY29yZWUuZGxsIiwgIlJFR19TWiINCiAgICBz" & _
        "aGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXENMU0lEXCIgJiBHVUlEX0NM" & _
        "U0lEICYgIlxJbnByb2NTZXJ2ZXIzMlxUaHJlYWRpbmdNb2RlbCIsICJCb3RoIiwgIlJFR19T" & _
        "WiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXENMU0lEXCIg" & _
        "JiBHVUlEX0NMU0lEICYgIlxJbnByb2NTZXJ2ZXIzMlxDbGFzcyIsIENMQVNTX05BTUUsICJS" & _
        "RUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJ" & _
        "RFwiICYgR1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVyMzJcQXNzZW1ibHkiLCBBU1NFTUJM" & _
        "WV9JTkZPLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENs" & _
        "YXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXElucHJvY1NlcnZlcjMyXFJ1bnRpbWVW" & _
        "ZXJzaW9uIiwgUlVOVElNRV9WRVJTSU9OLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRl" & _
        "ICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXElucHJv" & _
        "Y1NlcnZlcjMyXENvZGVCYXNlIiwgImZpbGU6Ly8vIiAmIFJlcGxhY2UoYWRkaW5zUGF0aCwg" & _
        "IlwiLCAiLyIpICYgIkZvbGRlcldhdGNoZXJDT00uRExMIiwgIlJFR19TWiINCiAgICANCiAg" & _
        "ICAnIDQuIFZlcnNp824gZXNwZWPtZmljYQ0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNP" & _
        "RlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXElucHJvY1NlcnZlcjMy" & _
        "XDEuMC4wLjBcQ2xhc3MiLCBDTEFTU19OQU1FLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dy" & _
        "aXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXElu" & _
        "cHJvY1NlcnZlcjMyXDEuMC4wLjBcQXNzZW1ibHkiLCBBU1NFTUJMWV9JTkZPLCAiUkVHX1Na" & _
        "Ig0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAm" & _
        "IEdVSURfQ0xTSUQgJiAiXElucHJvY1NlcnZlcjMyXDEuMC4wLjBcUnVudGltZVZlcnNpb24i"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "LCBSVU5USU1FX1ZFUlNJT04sICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1Vc" & _
        "U09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVy" & _
        "MzJcMS4wLjAuMFxDb2RlQmFzZSIsICJmaWxlOi8vLyIgJiBSZXBsYWNlKGFkZGluc1BhdGgs" & _
        "ICJcIiwgIi8iKSAmICJGb2xkZXJXYXRjaGVyQ09NLkRMTCIsICJSRUdfU1oiDQogICAgDQog" & _
        "ICAgJyA1LiBQcm9nSWQNCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFz" & _
        "c2VzXCIgJiBQUk9HX0lEICYgIlwiLCBDTEFTU19OQU1FLCAiUkVHX1NaIg0KICAgIHNoZWxs" & _
        "LlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcIiAmIFBST0dfSUQgJiAiXENMU0lE" & _
        "XCIsIEdVSURfQ0xTSUQsICJSRUdfU1oiDQogICAgDQogICAgJyA2LiBJbnRlcmZhY2VzDQog" & _
        "ICAgUmVnaXN0cmFySW50ZXJmYXogR1VJRF9JbnRlcmZhY2UxLCAiX0ZvbGRlcldhdGNoZXIi" & _
        "LCBQUk9YWVNUVUJfQ0xTSUQxDQogICAgUmVnaXN0cmFySW50ZXJmYXogR1VJRF9JbnRlcmZh" & _
        "Y2UyLCAiSUZvbGRlcldhdGNoZXJFdmVudHMiLCBQUk9YWVNUVUJfQ0xTSUQyDQogICAgDQog" & _
        "ICAgJyA3LiBUeXBlTGliDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xh" & _
        "c3Nlc1xUeXBlTGliXCIgJiBHVUlEX1R5cGVMaWIgJiAiXCIsICIiLCAiUkVHX1NaIg0KICAg" & _
        "IHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcVHlwZUxpYlwiICYgR1VJ" & _
        "RF9UeXBlTGliICYgIlwxLjBcIiwgIkNvbXBvbmVudGUgQ09NIG1vbml0b3JpemFjafNuIGNh" & _
        "cnBldGFzIiwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxD" & _
        "bGFzc2VzXFR5cGVMaWJcIiAmIEdVSURfVHlwZUxpYiAmICJcMS4wXDBcIiwgIiIsICJSRUdf" & _
        "U1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xUeXBlTGli" & _
        "XCIgJiBHVUlEX1R5cGVMaWIgJiAiXDEuMFwwXHdpbjY0XCIsIGFkZGluc1BhdGggJiAiRm9s" & _
        "ZGVyV2F0Y2hlckNPTS50bGIiLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NV"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "XFNPRlRXQVJFXENsYXNzZXNcVHlwZUxpYlwiICYgR1VJRF9UeXBlTGliICYgIlwxLjBcRkxB" & _
        "R1NcIiwgIjAiLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJF" & _
        "XENsYXNzZXNcVHlwZUxpYlwiICYgR1VJRF9UeXBlTGliICYgIlwxLjBcSEVMUERJUlwiLCBh" & _
        "ZGRpbnNQYXRoLCAiUkVHX1NaIg0KICAgIA0KICAgICcgOC4gUmVnaXN0cm9zIFdPVzY0MzJO" & _
        "b2RlIChwYXJhIGNvbXBhdGliaWxpZGFkIDMyLWJpdCkNCiAgICBSZWdpc3RyYXJJbnRlcmZh" & _
        "eldPVzY0IEdVSURfSW50ZXJmYWNlMSwgIl9Gb2xkZXJXYXRjaGVyIiwgUFJPWFlTVFVCX0NM" & _
        "U0lEMQ0KICAgIFJlZ2lzdHJhckludGVyZmF6V09XNjQgR1VJRF9JbnRlcmZhY2UyLCAiSUZv" & _
        "bGRlcldhdGNoZXJFdmVudHMiLCBQUk9YWVNUVUJfQ0xTSUQyDQogICAgDQogICAgSWYgRXJy" & _
        "Lk51bWJlciA9IDAgVGhlbg0KICAgICAgICBXU2NyaXB0LkVjaG8gIlJlZ2lzdHJvIENPTSBj" & _
        "b21wbGV0YWRvIGV4aXRvc2FtZW50ZS4iDQogICAgRWxzZQ0KICAgICAgICBXU2NyaXB0LkVj" & _
        "aG8gIkVycm9yIGR1cmFudGUgZWwgcmVnaXN0cm86ICIgJiBFcnIuRGVzY3JpcHRpb24NCiAg" & _
        "ICBFbmQgSWYNCkVuZCBTdWINCg0KJyA9PT09PT09PT09IEZVTkNJ004gQVVYSUxJQVIgUEFS" & _
        "QSBJTlRFUkZBQ0VTID09PT09PT09PT0NClN1YiBSZWdpc3RyYXJJbnRlcmZheihndWlkLCBu" & _
        "b21icmUsIHByb3h5U3R1YkNsc2lkKQ0KICAgIERpbSBzaGVsbA0KICAgIFNldCBzaGVsbCA9" & _
        "IENyZWF0ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQogICAgDQogICAgc2hlbGwuUmVnV3Jp" & _
        "dGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIGd1aWQgJiAiXCIsIG5v" & _
        "bWJyZSwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFz" & _
        "c2VzXEludGVyZmFjZVwiICYgZ3VpZCAmICJcUHJveHlTdHViQ2xzaWQzMlwiLCBwcm94eVN0" & _
        "dWJDbHNpZCwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxD" & _
        "bGFzc2VzXEludGVyZmFjZVwiICYgZ3VpZCAmICJcVHlwZUxpYlwiLCBHVUlEX1R5cGVMaWIs"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xJ" & _
        "bnRlcmZhY2VcIiAmIGd1aWQgJiAiXFR5cGVMaWJcVmVyc2lvblwiLCAiMS4wIiwgIlJFR19T" & _
        "WiINCkVuZCBTdWINCg0KJyA9PT09PT09PT09IEZVTkNJ004gQVVYSUxJQVIgUEFSQSBXT1c2" & _
        "NDMyTm9kZSA9PT09PT09PT09DQpTdWIgUmVnaXN0cmFySW50ZXJmYXpXT1c2NChndWlkLCBu" & _
        "b21icmUsIHByb3h5U3R1YkNsc2lkKQ0KICAgIERpbSBzaGVsbA0KICAgIFNldCBzaGVsbCA9" & _
        "IENyZWF0ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQogICAgDQogICAgJyBEb3MgdWJpY2Fj" & _
        "aW9uZXMgZGlmZXJlbnRlcyBwYXJhIFdPVzY0DQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1Vc" & _
        "U09GVFdBUkVcQ2xhc3Nlc1xXT1c2NDMyTm9kZVxJbnRlcmZhY2VcIiAmIGd1aWQgJiAiXCIs" & _
        "IG5vbWJyZSwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxD" & _
        "bGFzc2VzXFdPVzY0MzJOb2RlXEludGVyZmFjZVwiICYgZ3VpZCAmICJcUHJveHlTdHViQ2xz" & _
        "aWQzMlwiLCBwcm94eVN0dWJDbHNpZCwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAi" & _
        "SEtDVVxTT0ZUV0FSRVxDbGFzc2VzXFdPVzY0MzJOb2RlXEludGVyZmFjZVwiICYgZ3VpZCAm" & _
        "ICJcVHlwZUxpYlwiLCBHVUlEX1R5cGVMaWIsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3Jp" & _
        "dGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xXT1c2NDMyTm9kZVxJbnRlcmZhY2VcIiAmIGd1" & _
        "aWQgJiAiXFR5cGVMaWJcVmVyc2lvblwiLCAiMS4wIiwgIlJFR19TWiINCiAgICANCiAgICBz" & _
        "aGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxXT1c2NDMyTm9kZVxDbGFzc2VzXEludGVy" & _
        "ZmFjZVwiICYgZ3VpZCAmICJcIiwgbm9tYnJlLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dy" & _
        "aXRlICJIS0NVXFNPRlRXQVJFXFdPVzY0MzJOb2RlXENsYXNzZXNcSW50ZXJmYWNlXCIgJiBn" & _
        "dWlkICYgIlxQcm94eVN0dWJDbHNpZDMyXCIsIHByb3h5U3R1YkNsc2lkLCAiUkVHX1NaIg0K" & _
        "ICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXFdPVzY0MzJOb2RlXENsYXNzZXNc"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "SW50ZXJmYWNlXCIgJiBndWlkICYgIlxUeXBlTGliXCIsIEdVSURfVHlwZUxpYiwgIlJFR19T" & _
        "WiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxXT1c2NDMyTm9kZVxDbGFz" & _
        "c2VzXEludGVyZmFjZVwiICYgZ3VpZCAmICJcVHlwZUxpYlxWZXJzaW9uXCIsICIxLjAiLCAi" & _
        "UkVHX1NaIg0KRW5kIFN1Yg0KDQonID09PT09PT09PT0gU1VCUlVUSU5BIFBBUkEgRUxJTUlO" & _
        "QVIgPT09PT09PT09PQ0KU3ViIEVsaW1pbmFyQ2xhdmVzQ09NKCkNCiAgICBPbiBFcnJvciBS" & _
        "ZXN1bWUgTmV4dA0KICAgIERpbSBzaGVsbA0KICAgIA0KICAgIFNldCBzaGVsbCA9IENyZWF0" & _
        "ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpDQogICAgDQogICAgJyBFbGltaW5hciBlbiBvcmRl" & _
        "biBpbnZlcnNvIChkZSBt4XMgZXNwZWPtZmljbyBhIG3hcyBnZW5lcmFsKQ0KICAgIA0KICAg" & _
        "ICcgMS4gRWxpbWluYXIgV09XNjQzMk5vZGUgZW50cmllcw0KICAgIEVsaW1pbmFyU2lFeGlz" & _
        "dGUgIkhLQ1VcU09GVFdBUkVcV09XNjQzMk5vZGVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIEdV" & _
        "SURfSW50ZXJmYWNlMSAmICJcIg0KICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdB" & _
        "UkVcV09XNjQzMk5vZGVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIEdVSURfSW50ZXJmYWNlMiAm" & _
        "ICJcIg0KICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xXT1c2" & _
        "NDMyTm9kZVxJbnRlcmZhY2VcIiAmIEdVSURfSW50ZXJmYWNlMSAmICJcIg0KICAgIEVsaW1p" & _
        "bmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xXT1c2NDMyTm9kZVxJbnRlcmZh" & _
        "Y2VcIiAmIEdVSURfSW50ZXJmYWNlMiAmICJcIg0KICAgIA0KICAgICcgMi4gRWxpbWluYXIg" & _
        "VHlwZUxpYg0KICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xU" & _
        "eXBlTGliXCIgJiBHVUlEX1R5cGVMaWIgJiAiXCINCiAgICANCiAgICAnIDMuIEVsaW1pbmFy" & _
        "IEludGVyZmFjZXMgbm9ybWFsZXMNCiAgICBFbGltaW5hclNpRXhpc3RlICJIS0NVXFNPRlRX" & _
        "QVJFXENsYXNzZXNcSW50ZXJmYWNlXCIgJiBHVUlEX0ludGVyZmFjZTEgJiAiXCINCiAgICBF"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "bGltaW5hclNpRXhpc3RlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcSW50ZXJmYWNlXCIgJiBH" & _
        "VUlEX0ludGVyZmFjZTIgJiAiXCINCiAgICANCiAgICAnIDQuIEVsaW1pbmFyIFByb2dJZA0K" & _
        "ICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1wiICYgUFJPR19J" & _
        "RCAmICJcIg0KICAgIA0KICAgICcgNS4gRWxpbWluYXIgQ0xTSUQgKGVzdG8gZWxpbWluYXLh" & _
        "IHRvZGEgbGEgamVyYXJxde1hKQ0KICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdB" & _
        "UkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcIg0KICAgIA0KICAgIElmIEVy" & _
        "ci5OdW1iZXIgPSAwIFRoZW4NCiAgICAgICAgV1NjcmlwdC5FY2hvICJFbGltaW5hY2nzbiBk" & _
        "ZSBjbGF2ZXMgQ09NIGNvbXBsZXRhZGEgZXhpdG9zYW1lbnRlLiINCiAgICBFbHNlDQogICAg" & _
        "ICAgIFdTY3JpcHQuRWNobyAiRXJyb3IgZHVyYW50ZSBsYSBlbGltaW5hY2nzbjogIiAmIEVy" & _
        "ci5EZXNjcmlwdGlvbg0KICAgIEVuZCBJZg0KRW5kIFN1Yg0KDQonID09PT09PT09PT0gRlVO" & _
        "Q0nTTiBBVVhJTElBUiBQQVJBIEVMSU1JTkFDSdNOIFNFR1VSQSA9PT09PT09PT09DQpTdWIg" & _
        "RWxpbWluYXJTaUV4aXN0ZShydXRhKQ0KICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQogICAg" & _
        "RGltIHNoZWxsDQogICAgU2V0IHNoZWxsID0gQ3JlYXRlT2JqZWN0KCJXU2NyaXB0LlNoZWxs" & _
        "IikNCiAgICANCiAgICAnIEludGVudGFyIGxlZXIgcGFyYSB2ZXIgc2kgZXhpc3RlDQogICAg" & _
        "c2hlbGwuUmVnUmVhZCBydXRhDQogICAgDQogICAgSWYgRXJyLk51bWJlciA9IDAgVGhlbg0K" & _
        "ICAgICAgICAnIExhIGNsYXZlIGV4aXN0ZSwgZWxpbe1uYWxhDQogICAgICAgIEVyci5DbGVh" & _
        "cg0KICAgICAgICBzaGVsbC5SZWdEZWxldGUgcnV0YQ0KICAgICAgICBJZiBFcnIuTnVtYmVy" & _
        "IDw+IDAgVGhlbg0KICAgICAgICAgICAgV1NjcmlwdC5FY2hvICIgIEFkdmVydGVuY2lhOiBO" & _
        "byBzZSBwdWRvIGVsaW1pbmFyICIgJiBydXRhDQogICAgICAgIEVuZCBJZg0KICAgIEVuZCBJ" & _
        "Zg0KICAgIEVyci5DbGVhcg0KRW5kIFN1Yg0KDQonID09PT09PT09PT0gRUpFTVBMTyBERSBV"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "U08gPT09PT09PT09PQ0KJyBQYXJhIHByb2JhciBsYXMgZnVuY2lvbmVzOg0KJyBSZWdpc3Ry" & _
        "YXJDbGF2ZXNDT00oKSAgICcgUGFyYSByZWdpc3RyYXINCicgRWxpbWluYXJDbGF2ZXNDT00o" & _
        "KSAgICAnIFBhcmEgZWxpbWluYXI="
End Function



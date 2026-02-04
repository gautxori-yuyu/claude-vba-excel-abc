Attribute VB_Name = "modAPPInstallXLAM"
' ==========================================
' INSTALACIÃN Y DESINSTALACIÃN AUTOMÃTICA DEL COMPLEMENTO XLAM
' ==========================================
' Este mÃ³dulo contiene la lÃ³gica de auto-instalaciÃ³n / auto-desinstalaciÃ³n
' del complemento XLAM en la carpeta de complementos del usuario, apoyÃ¡ndose
' en un script externo (VBScript) codificado en Base64 + RC4.
'
' El VBScript (AutoXLAM_Installer.vbs) gestiona:
'   1. Copia del XLAM a la carpeta de complementos
'   2. ExtracciÃ³n del COM desde dentro del XLAM (que es un ZIP):
'      - xl/embeddings/FolderWatcherCOM.dll
'      - xl/embeddings/FolderWatcherCOM.dll.manifest
'   3. Registro/desregistro del complemento en Excel
'
' IMPORTANTE: El COM debe estar embebido dentro del XLAM (carpeta xl/embeddings)
' ==========================================

'@Folder "1-Inicio e Instalacion.Instalacion"
'@IgnoreModule ProcedureNotUsed

Option Private Module
Option Explicit

Private Const MODULE_NAME As String = "modAPPInstallXLAM"

' ---------------------------------------------------------------------
' CONSTANTES DE INSTALACIÃN
' ---------------------------------------------------------------------

' Constantes asociadas a la instalaciÃ³n del XLAM
Public Const SCRIPT_NOMBRE As String = "AutoXLAM_Installer.vbs"

' Constantes para el componente COM FolderWatcher
Private Const COM_DLL_NOMBRE As String = "FolderWatcherCOM.dll"
Private Const COM_MANIFEST_NOMBRE As String = "FolderWatcherCOM.dll.manifest"

' ---------------------------------------------------------------------
' UTILIDADES DE PREPARACIÃN DEL SCRIPT
' ---------------------------------------------------------------------

'@Description: Codifica el script de instalaciÃ³n VBScript a Base64 utilizando RC4 y lo transforma en una funciÃ³n VBA embebida.
'@Scope: Manipula archivos temporales del sistema y genera cÃ³digo embebido.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: InstalaciÃ³n XLAM
Sub archivoInstScriptToBase64RC4()
Attribute archivoInstScriptToBase64RC4.VB_ProcData.VB_Invoke_Func = " \n0"
    ScriptToFunctionBase64RC4 _
        Replace(Environ$("TEMP") & "\" & "AutoXLAM_Installer.vbs", "\\", "\"), _
        Replace(Environ$("TEMP") & "\" & "AutoXLAM_Installer.Base64", "\\", "\"), _
        "INSTALLSCRIPT_B64RC4"
End Sub

' ---------------------------------------------------------------------
' FUNCIONES COM (DEPRECATED - El VBScript ahora gestiona la instalaciÃ³n)
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
        LogWarning "modAPPInstallXLAM", "[InstalarCOM] - DLL no encontrada: " & rutaDLLOrigen
        InstalarCOM = False
        GoTo CleanUp
    End If

    If Not fso.FileExists(rutaManifestOrigen) Then
        LogWarning "modAPPInstallXLAM", "[InstalarCOM] - Manifest no encontrado: " & rutaManifestOrigen
        InstalarCOM = False
        GoTo CleanUp
    End If

    ' Eliminar archivos existentes si los hay
    On Error Resume Next
    If fso.FileExists(rutaDLLDestino) Then fso.DeleteFile rutaDLLDestino, True
    If fso.FileExists(rutaManifestDestino) Then fso.DeleteFile rutaManifestDestino, True
    On Error GoTo ErrHandler

    ' Copiar DLL
    fso.CopyFile rutaDLLOrigen, rutaDLLDestino, True
    LogInfo "modAPPInstallXLAM", "[InstalarCOM] - DLL copiada a: " & rutaDLLDestino

    ' Copiar Manifest
    fso.CopyFile rutaManifestOrigen, rutaManifestDestino, True
    LogInfo "modAPPInstallXLAM", "[InstalarCOM] - Manifest copiado a: " & rutaManifestDestino

    InstalarCOM = True
CleanUp:
    Set fso = Nothing
    Exit Function

ErrHandler:
    LogError "modAPPInstallXLAM", "[InstalarCOM] - Error", Err.Number, Err.Description
    InstalarCOM = False
    Resume CleanUp
End Function

'@Description: [DEPRECATED] Desinstala los archivos COM de la carpeta AddIns
'@Note: Ya no se usa. El VBScript elimina el COM durante desinstalaciÃ³n.
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
        LogInfo "modAPPInstallXLAM", "[DesinstalarCOM] - DLL eliminada: " & rutaDLL
    End If

    ' Eliminar Manifest si existe
    If fso.FileExists(rutaManifest) Then
        fso.DeleteFile rutaManifest, True
        LogInfo "modAPPInstallXLAM", "[DesinstalarCOM] - Manifest eliminado: " & rutaManifest
    End If

    DesinstalarCOM = True

CleanUp:
    Set fso = Nothing
    Exit Function

ErrHandler:
    LogError "modAPPInstallXLAM", "[DesinstalarCOM] - Error", Err.Number, Err.Description
    DesinstalarCOM = False
    Resume CleanUp
End Function

'@Description: Verifica si los archivos COM estÃ¡n instalados en la carpeta AddIns
'@Returns: Boolean | True si ambos archivos (DLL y manifest) existen
'@Category: InstalaciÃ³n COM
Public Function ComprobarCOMInstalado() As Boolean
Attribute ComprobarCOMInstalado.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim fso As Object
    Dim rutaDestino As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaDestino = Application.UserLibraryPath

    ComprobarCOMInstalado = fso.FileExists(rutaDestino & COM_DLL_NOMBRE) And _
                            fso.FileExists(rutaDestino & COM_MANIFEST_NOMBRE)

    Set fso = Nothing
End Function

' ---------------------------------------------------------------------
' FLUJO PRINCIPAL DE AUTO-INSTALACIÃN / DESINSTALACIÃN
' ---------------------------------------------------------------------

'@Description: Gestiona automÃ¡ticamente la instalaciÃ³n o desinstalaciÃ³n del complemento XLAM segÃºn su estado actual.
'@Scope: Manipula el libro actual, complementos de Excel y ejecuta scripts externos.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: InstalaciÃ³n XLAM
Public Sub AutoInstalador()
Attribute AutoInstalador.VB_ProcData.VB_Invoke_Func = " \n0"
    
    ' Validar que se estÃ¡ ejecutando desde un XLAM
    If Not (ThisWorkbook.FileFormat = xlOpenXMLAddIn Or ThisWorkbook.FileFormat = xlAddIn) Then Exit Sub
    
    Dim rutaActual As String
    Dim rutaDestino As String
    
    rutaActual = ThisWorkbook.Path & "\"
    rutaDestino = Application.UserLibraryPath
    
    ' Si ya se ejecuta desde la carpeta destino, no hacer nada
    If rutaActual = rutaDestino Then
        LogInfo "modAPPInstallXLAM", "[AutoInstalador] - el complemento se inicia desde la ruta destino de instalaciÃ³n, NO se ejecuta el proceso de instalaciÃ³n / desinstalaciÃ³n"
        Exit Sub
    End If
    
    ' Si NO estÃ¡ instalado
    If Not ComprobarSiInstalado() Then
        
        ' Evitar sobrescribir un XLAM con el mismo nombre final
        If LCase$(ThisWorkbook.Name) = LCase$(APP_NAME & ".xlam") Then
            
            LogInfo "modAPPInstallXLAM", "[AutoInstalador] - XLAM no es posible instalarlo"
            MsgBox "El nombre del fichero a instalar tiene que ser diferente de '" & APP_NAME & ".xlam" & "'. CÃ¡mbialo si quieres hacer la instalaciÃ³n."
            
        ElseIf MsgBox("Â¿Deseas instalar este complemento?", vbYesNo + vbQuestion) = vbYes Then
            
            LogInfo "modAPPInstallXLAM", "[AutoInstalador] - ejecutando script de instalaciÃ³n"
            
            EjecutarScript _
                INSTALLSCRIPT_B64RC4, _
                SCRIPT_NOMBRE, _
                Array("/install", ThisWorkbook.FullName, Application.UserLibraryPath, APP_NAME), _
                True
            
            If Application.Workbooks.Count <= 1 Then Application.Quit
            ThisWorkbook.Close SaveChanges:=False
            
        End If
        
    ' Si YA estÃ¡ instalado
    Else
        
        If MsgBox("Este complemento ya estÃ¡ instalado. Â¿Deseas desinstalarlo?", vbYesNo + vbQuestion) = vbYes Then
            
            LogInfo "modAPPInstallXLAM", "[AutoInstalador] - ejecutando script de desinstalaciÃ³n"
            
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
' COMPROBACIÃN DE ESTADO DE INSTALACIÃN
' ---------------------------------------------------------------------

'@Description: Comprueba si el complemento XLAM estÃ¡ instalado correctamente en Excel y sincroniza su estado si hay inconsistencias.
'@Scope: Manipula la colecciÃ³n Application.AddIns y verifica archivos en el sistema.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Boolean | True si el XLAM estÃ¡ instalado; False en caso contrario.
'@Category: InstalaciÃ³n XLAM
Public Function ComprobarSiInstalado() As Boolean
Attribute ComprobarSiInstalado.VB_ProcData.VB_Invoke_Func = " \n0"
    
    Dim ai As AddIn
    Dim bFExists As Boolean
    
    ' Verificar existencia fÃ­sica del XLAM
    bFExists = Dir(Application.UserLibraryPath & APP_NAME & ".xlam", vbNormal) <> ""
    
    For Each ai In Application.AddIns
        If ai.Name = APP_NAME & ".xlam" Then
            
            ' Estado inconsistente: marcado como instalado pero el fichero no existe
            If Not bFExists And ai.Installed Then
                LogError "modAPPInstallXLAM", "[ComprobarSiInstalado] - XLAM marcado como instalado, pero inexistente: forzando el proceso de desinstalaciÃ³n"
                ai.Installed = False
            End If
            
            ComprobarSiInstalado = ai.Installed
            LogInfo "modAPPInstallXLAM", "[ComprobarSiInstalado] - XLAM " & IIf(ComprobarSiInstalado, "", "no ") & "instalado"
            Exit Function
            
        End If
    Next ai
    
End Function
Function INSTALLSCRIPT_B64RC4() As String
Attribute INSTALLSCRIPT_B64RC4.VB_ProcData.VB_Invoke_Func = " \n0"
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
        "ZW50ZSBjb20gRW4gSEtjVSAgDQogICAgJyBSZWdpc3RyYXJDbGF2ZXNDT00oKQ0KDQogICAg" & _
        "JyA0LiBDb3BpYXIgWExBTSBhbCBkZXN0aW5vDQogICAgZnNvLkNvcHlGaWxlIGFyY2hpdm8s" & _
        "IHJ1dGFGaW5hbCwgVHJ1ZQ0KDQogICAgJyA1LiBSZWdpc3RyYXIgZW4gRXhjZWwNCiAgICBT" & _
        "ZXQgZXhjZWwgPSBDcmVhdGVPYmplY3QoIkV4Y2VsLkFwcGxpY2F0aW9uIikNCiAgICBleGNl" & _
        "bC5WaXNpYmxlID0gRmFsc2UNCg0KICAgIEZvciBFYWNoIGFpIEluIGV4Y2VsLkFkZElucw0K" & _
        "ICAgICAgICBJZiBMQ2FzZShhaS5OYW1lKSA9IExDYXNlKG5vbWJyZSAmICIueGxhbSIpIFRo" & _
        "ZW4NCiAgICAgICAgICAgIGFpLkluc3RhbGxlZCA9IFRydWUNCiAgICAgICAgICAgIEV4aXQg" & _
        "Rm9yDQogICAgICAgIEVuZCBJZg0KICAgIE5leHQNCg0KICAgIFdTY3JpcHQuU2xlZXAgMTAw" & _
        "MA0KDQogICAgSWYgYWkgSXMgTm90aGluZyBUaGVuDQogICAgICAgIE1zZ0JveCAiTm8gaGEg" & _
        "c2lkbyBwb3NpYmxlIGNvbXBsZXRhciBsYSBpbnN0YWxhY2nzbi4gUG9yIGZhdm9yLCBoYWJp" & _
        "bGl0YSBlbCBjb21wbGVtZW50byBkZXNkZSBlbCBtZW76IGRlIGNvbXBsZW1lbnRvcyBkZSBF" & _
        "eGNlbC4iLCB2YkNyaXRpY2FsDQogICAgRWxzZUlmIE5vdCBhaS5JbnN0YWxsZWQgVGhlbg0K" & _
        "ICAgICAgICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJsZSBjb21wbGV0YXIgbGEgaW5zdGFs" & _
        "YWNp824uIFBvciBmYXZvciwgaGFiaWxpdGEgZWwgY29tcGxlbWVudG8gZGVzZGUgZWwgbWVu" & _
        "+iBkZSBjb21wbGVtZW50b3MgZGUgRXhjZWwuIiwgdmJDcml0aWNhbA0KICAgIEVsc2UNCiAg" & _
        "ICAgICAgTXNnQm94ICJJbnN0YWxhY2nzbiBjb21wbGV0YWRhLCByZWluaWNpYSBFeGNlbC4i" & _
        "LCB2YkluZm9ybWF0aW9uDQogICAgRW5kIElmDQoNCiAgICBleGNlbC5RdWl0DQogICAgU2V0"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "IGV4Y2VsID0gTm90aGluZw0KRW5kIFN1Yg0KDQonID09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQonIERFU0lOU1RBTEFDSdNODQonID09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQpT" & _
        "dWIgRG9Vbmluc3RhbGwoKQ0KICAgICcgMS4gRWxpbWluYXIgYXJjaGl2b3MgQ09NIHByaW1l" & _
        "cm8gKGFudGVzIGRlIHF1ZSBFeGNlbCBsb3MgYmxvcXVlZSkNCiAgICBSZW1vdmVDT01GaWxl" & _
        "cyBkZXN0aW5vDQoNCiAgICAnIDIuIEVsaW1pbmFyIFhMQU0NCiAgICBSZW1vdmVBZGRpbklu" & _
        "RGVzdGlubyBydXRhRmluYWwNCiAgICANCiAgICAnIDMuIEVsaW1pbmFyIGNsYXZlcyBkZSBy" & _
        "ZWdpc3RybyBkZWwgY29tcG9uZW50ZSBjb20gRW4gSEtjVSAgDQogICAgJyBFbGltaW5hckNs" & _
        "YXZlc0NPTSgpDQoNCiAgICAnIDQuIERlc3JlZ2lzdHJhciBkZSBFeGNlbA0KICAgIFNldCBl" & _
        "eGNlbCA9IENyZWF0ZU9iamVjdCgiRXhjZWwuQXBwbGljYXRpb24iKQ0KICAgIHZlcnMgPSBl" & _
        "eGNlbC5BcHBsaWNhdGlvbi5WZXJzaW9uDQogICAgZXhjZWwuVmlzaWJsZSA9IEZhbHNlDQoN" & _
        "CiAgICBGb3IgRWFjaCBhaSBJbiBleGNlbC5BZGRJbnMNCiAgICAgICAgSWYgTENhc2UoYWku" & _
        "TmFtZSkgPSBMQ2FzZShub21icmUgJiAiLnhsYW0iKSBUaGVuDQogICAgICAgICAgICBhaS5J" & _
        "bnN0YWxsZWQgPSBGYWxzZQ0KICAgICAgICAgICAgRXhpdCBGb3INCiAgICAgICAgRW5kIElm" & _
        "DQogICAgTmV4dA0KDQogICAgRGltIHVuaW5zdGFsbE9LDQogICAgdW5pbnN0YWxsT0sgPSBU" & _
        "cnVlDQogICAgSWYgTm90IGFpIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBJZiBhaS5JbnN0" & _
        "YWxsZWQgVGhlbiB1bmluc3RhbGxPSyA9IEZhbHNlDQogICAgRW5kIElmDQoNCiAgICBJZiBO" & _
        "b3QgdW5pbnN0YWxsT0sgVGhlbg0KICAgICAgICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJs" & _
        "ZSBjb21wbGV0YXIgbGEgZGVzaW5zdGFsYWNp824uIFBvciBmYXZvciwgcmVpbnTpbnRhbG8g" & _
        "ZGUgbnVldm8gbyBkZXNoYWJpbGl0YSBlbCBjb21wbGVtZW50byBkZXNkZSBlbCBtZW76IGRl"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "IGNvbXBsZW1lbnRvcyBkZSBFeGNlbC4iLCB2YkNyaXRpY2FsDQogICAgRWxzZQ0KICAgICAg" & _
        "ICBNc2dCb3ggIkRlc2luc3RhbGFjafNuIGNvbXBsZXRhZGEsIHJlaW5pY2lhIEV4Y2VsLiIs" & _
        "IHZiSW5mb3JtYXRpb24NCiAgICBFbmQgSWYNCg0KICAgIGV4Y2VsLlF1aXQNCiAgICBTZXQg" & _
        "ZXhjZWwgPSBOb3RoaW5nDQoNCiAgICAnIDUuIExpbXBpYXIgY2xhdmVzIGRlIGNvbmZpZ3Vy" & _
        "YWNp824gZGVsIFhMQU0gZW4gZWwgcmVnaXN0cm8gDQogICAgQ2xlYW5SZWdpc3RyeSB2ZXJz" & _
        "LCBub21icmUNCkVuZCBTdWINCg0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PQ0KJyBFWFRSQUNDSdNOIERFTCBDT00gREVTREUgRUwg" & _
        "WExBTSAoWklQKQ0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PQ0KRnVuY3Rpb24gRXh0cmFjdENPTUZyb21YTEFNKHhsYW1QYXRoLCBk" & _
        "ZXN0Rm9sZGVyKQ0KICAgIEV4dHJhY3RDT01Gcm9tWExBTSA9IEZhbHNlDQoNCiAgICBPbiBF" & _
        "cnJvciBSZXN1bWUgTmV4dA0KDQogICAgJyBJbnRlbnRhciBwcmltZXJvIGNvbiA3emlwICht" & _
        "4XMgcuFwaWRvIHkgZmlhYmxlKQ0KICAgIElmIFRyeUV4dHJhY3RXaXRoN1ppcCh4bGFtUGF0" & _
        "aCwgZGVzdEZvbGRlcikgVGhlbg0KICAgICAgICBFeHRyYWN0Q09NRnJvbVhMQU0gPSBUcnVl" & _
        "DQogICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAgICcgU2kgbm8gaGF5" & _
        "IDd6aXAsIHVzYXIgU2hlbGwuQXBwbGljYXRpb24gKFdpbmRvd3MgbmF0aXZvKQ0KICAgIElm" & _
        "IFRyeUV4dHJhY3RXaXRoU2hlbGwoeGxhbVBhdGgsIGRlc3RGb2xkZXIpIFRoZW4NCiAgICAg" & _
        "ICAgRXh0cmFjdENPTUZyb21YTEFNID0gVHJ1ZQ0KICAgICAgICBFeGl0IEZ1bmN0aW9uDQog" & _
        "ICAgRW5kIElmDQoNCiAgICBPbiBFcnJvciBHb1RvIDANCkVuZCBGdW5jdGlvbg0KDQonIEV4" & _
        "dHJhY2Np824gdXNhbmRvIDctWmlwDQpGdW5jdGlvbiBUcnlFeHRyYWN0V2l0aDdaaXAoeGxh" & _
        "bVBhdGgsIGRlc3RGb2xkZXIpDQogICAgVHJ5RXh0cmFjdFdpdGg3WmlwID0gRmFsc2UNCg0K"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ICAgIERpbSBzaGVsbCwgZXhlYywgc2V2ZW5aaXBQYXRoDQogICAgU2V0IHNoZWxsID0gQ3Jl" & _
        "YXRlT2JqZWN0KCJXU2NyaXB0LlNoZWxsIikNCg0KICAgICcgQnVzY2FyIDd6LmV4ZSBlbiBl" & _
        "bCBQQVRIDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCiAgICBTZXQgZXhlYyA9IHNoZWxs" & _
        "LkV4ZWMoIndoZXJlIDd6LmV4ZSIpDQogICAgSWYgRXJyLk51bWJlciA9IDAgVGhlbg0KICAg" & _
        "ICAgICBEbyBXaGlsZSBleGVjLlN0YXR1cyA9IDANCiAgICAgICAgICAgIFdTY3JpcHQuU2xl" & _
        "ZXAgMTAwDQogICAgICAgIExvb3ANCiAgICAgICAgc2V2ZW5aaXBQYXRoID0gVHJpbShleGVj" & _
        "LlN0ZE91dC5SZWFkTGluZSkNCiAgICBFbmQgSWYNCiAgICBPbiBFcnJvciBHb1RvIDANCg0K" & _
        "ICAgIElmIHNldmVuWmlwUGF0aCA9ICIiIE9yIE5vdCBmc28uRmlsZUV4aXN0cyhzZXZlblpp" & _
        "cFBhdGgpIFRoZW4NCiAgICAgICAgJyA3emlwIG5vIGVuY29udHJhZG8NCiAgICAgICAgc2V2" & _
        "ZW5aaXBQYXRoID0gIjd6LmV4ZSINCiAgICAgICAgJ0V4aXQgRnVuY3Rpb24NCiAgICBFbmQg" & _
        "SWYNCg0KICAgICcgRXh0cmFlciBzb2xvIGxvcyBhcmNoaXZvcyBDT00NCiAgICBEaW0gY21k" & _
        "LCBkbGxQYXRoLCBtYW5pZmVzdFBhdGgsIGNvbmZpZ1BhdGgNCiAgICBkbGxQYXRoID0gQ09N" & _
        "X0VNQkVEX1BBVEggJiBDT01fRExMX05BTUUNCiAgICBtYW5pZmVzdFBhdGggPSBDT01fRU1C" & _
        "RURfUEFUSCAmIENPTV9NQU5JRkVTVF9OQU1FDQogICAgY29uZmlnUGF0aCA9IENPTV9FTUJF" & _
        "RF9QQVRIICYgQ09NX0NPTkZJR19OQU1FDQoNCiAgICAnIEV4dHJhZXIgRExMDQogICAgY21k" & _
        "ID0gIiIiIiAmIHNldmVuWmlwUGF0aCAmICIiIiBlICIiIiAmIHhsYW1QYXRoICYgIiIiIC1v" & _
        "IiIiICYgZGVzdEZvbGRlciAmICIiIiAiIiIgJiBkbGxQYXRoICYgIiIiIC15Ig0KICAgIHNo" & _
        "ZWxsLlJ1biBjbWQsIDAsIFRydWUNCg0KICAgICcgRXh0cmFlciBNYW5pZmVzdA0KICAgIGNt" & _
        "ZCA9ICIiIiIgJiBzZXZlblppcFBhdGggJiAiIiIgZSAiIiIgJiB4bGFtUGF0aCAmICIiIiAt" & _
        "byIiIiAmIGRlc3RGb2xkZXIgJiAiIiIgIiIiICYgbWFuaWZlc3RQYXRoICYgIiIiIC15Ig0K"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "ICAgIHNoZWxsLlJ1biBjbWQsIDAsIFRydWUNCg0KICAgICcgRXh0cmFlciBDb25maWcNCiAg" & _
        "ICBjbWQgPSAiIiIiICYgc2V2ZW5aaXBQYXRoICYgIiIiIGUgIiIiICYgeGxhbVBhdGggJiAi" & _
        "IiIgLW8iIiIgJiBkZXN0Rm9sZGVyICYgIiIiICIiIiAmIGNvbmZpZ1BhdGggJiAiIiIgLXki" & _
        "DQogICAgc2hlbGwuUnVuIGNtZCwgMCwgVHJ1ZQ0KDQogICAgJyBWZXJpZmljYXIgcXVlIHNl" & _
        "IGV4dHJhamVyb24NCiAgICBJZiBmc28uRmlsZUV4aXN0cyhkZXN0Rm9sZGVyICYgIlwiICYg" & _
        "Q09NX0RMTF9OQU1FKSBBbmQgXw0KICAgICAgIGZzby5GaWxlRXhpc3RzKGRlc3RGb2xkZXIg" & _
        "JiAiXCIgJiBDT01fQ09ORklHX05BTUUpIEFuZCBfDQogICAgICAgZnNvLkZpbGVFeGlzdHMo" & _
        "ZGVzdEZvbGRlciAmICJcIiAmIENPTV9NQU5JRkVTVF9OQU1FKSBUaGVuDQogICAgICAgIFRy" & _
        "eUV4dHJhY3RXaXRoN1ppcCA9IFRydWUNCiAgICBFbmQgSWYNCg0KICAgIFNldCBzaGVsbCA9" & _
        "IE5vdGhpbmcNCkVuZCBGdW5jdGlvbg0KDQonIEV4dHJhY2Np824gdXNhbmRvIFNoZWxsLkFw" & _
        "cGxpY2F0aW9uIChXaW5kb3dzIG5hdGl2bykNCkZ1bmN0aW9uIFRyeUV4dHJhY3RXaXRoU2hl" & _
        "bGwoeGxhbVBhdGgsIGRlc3RGb2xkZXIpDQogICAgVHJ5RXh0cmFjdFdpdGhTaGVsbCA9IEZh" & _
        "bHNlDQoNCiAgICBPbiBFcnJvciBSZXN1bWUgTmV4dA0KDQogICAgJyBDcmVhciBjb3BpYSB0" & _
        "ZW1wb3JhbCBjb21vIC56aXANCiAgICBEaW0gdGVtcFppcA0KICAgIHRlbXBaaXAgPSBmc28u" & _
        "R2V0U3BlY2lhbEZvbGRlcigyKSAmICJcIiAmIGZzby5HZXRUZW1wTmFtZSgpICYgIi56aXAi" & _
        "DQogICAgZnNvLkNvcHlGaWxlIHhsYW1QYXRoLCB0ZW1wWmlwLCBUcnVlDQoNCiAgICBJZiBF" & _
        "cnIuTnVtYmVyIDw+IDAgVGhlbiBFeGl0IEZ1bmN0aW9uDQoNCiAgICAnIFVzYXIgU2hlbGwu" & _
        "QXBwbGljYXRpb24gcGFyYSBleHBsb3JhciBlbCBaSVANCiAgICBEaW0gc2hlbGwsIHppcEZv" & _
        "bGRlciwgZGVzdEZvbGRlck9iag0KICAgIFNldCBzaGVsbCA9IENyZWF0ZU9iamVjdCgiU2hl" & _
        "bGwuQXBwbGljYXRpb24iKQ0KICAgIFNldCB6aXBGb2xkZXIgPSBzaGVsbC5OYW1lU3BhY2Uo"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "dGVtcFppcCkNCiAgICBTZXQgZGVzdEZvbGRlck9iaiA9IHNoZWxsLk5hbWVTcGFjZShkZXN0" & _
        "Rm9sZGVyKQ0KDQogICAgSWYgemlwRm9sZGVyIElzIE5vdGhpbmcgT3IgZGVzdEZvbGRlck9i" & _
        "aiBJcyBOb3RoaW5nIFRoZW4NCiAgICAgICAgZnNvLkRlbGV0ZUZpbGUgdGVtcFppcA0KICAg" & _
        "ICAgICBFeGl0IEZ1bmN0aW9uDQogICAgRW5kIElmDQoNCiAgICAnIEJ1c2NhciBsYSBjYXJw" & _
        "ZXRhIHhsXGVtYmVkZGluZ3MgZGVudHJvIGRlbCBaSVANCiAgICBEaW0gaXRlbSwgZW1iZWRG" & _
        "b2xkZXINCiAgICBTZXQgZW1iZWRGb2xkZXIgPSBOb3RoaW5nDQoNCiAgICAnIE5hdmVnYXIg" & _
        "YSB4bFxlbWJlZGRpbmdzDQogICAgRGltIHhsRm9sZGVyDQogICAgRm9yIEVhY2ggaXRlbSBJ" & _
        "biB6aXBGb2xkZXIuSXRlbXMNCiAgICAgICAgSWYgTENhc2UoaXRlbS5OYW1lKSA9ICJ4bCIg" & _
        "VGhlbg0KICAgICAgICAgICAgU2V0IHhsRm9sZGVyID0gc2hlbGwuTmFtZVNwYWNlKGl0ZW0u" & _
        "UGF0aCkNCiAgICAgICAgICAgIEV4aXQgRm9yDQogICAgICAgIEVuZCBJZg0KICAgIE5leHQN" & _
        "Cg0KICAgIElmIHhsRm9sZGVyIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBmc28uRGVsZXRl" & _
        "RmlsZSB0ZW1wWmlwDQogICAgICAgIEV4aXQgRnVuY3Rpb24NCiAgICBFbmQgSWYNCg0KICAg" & _
        "IEZvciBFYWNoIGl0ZW0gSW4geGxGb2xkZXIuSXRlbXMNCiAgICAgICAgSWYgTENhc2UoaXRl" & _
        "bS5OYW1lKSA9ICJlbWJlZGRpbmdzIiBUaGVuDQogICAgICAgICAgICBTZXQgZW1iZWRGb2xk" & _
        "ZXIgPSBzaGVsbC5OYW1lU3BhY2UoaXRlbS5QYXRoKQ0KICAgICAgICAgICAgRXhpdCBGb3IN" & _
        "CiAgICAgICAgRW5kIElmDQogICAgTmV4dA0KDQogICAgSWYgZW1iZWRGb2xkZXIgSXMgTm90" & _
        "aGluZyBUaGVuDQogICAgICAgIGZzby5EZWxldGVGaWxlIHRlbXBaaXANCiAgICAgICAgRXhp" & _
        "dCBGdW5jdGlvbg0KICAgIEVuZCBJZg0KDQogICAgJyBFeHRyYWVyIGxvcyBhcmNoaXZvcyBD" & _
        "T00NCiAgICBEaW0gZGxsSXRlbSwgbWFuaWZlc3RJdGVtLCBjb25maWdJdGVtDQogICAgRm9y" & _
        "IEVhY2ggaXRlbSBJbiBlbWJlZEZvbGRlci5JdGVtcw0KICAgICAgICBJZiBMQ2FzZShpdGVt"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "Lk5hbWUpID0gTENhc2UoQ09NX0RMTF9OQU1FKSBUaGVuDQogICAgICAgICAgICBTZXQgZGxs" & _
        "SXRlbSA9IGl0ZW0NCiAgICAgICAgRWxzZUlmIExDYXNlKGl0ZW0uTmFtZSkgPSBMQ2FzZShD" & _
        "T01fTUFOSUZFU1RfTkFNRSkgVGhlbg0KICAgICAgICAgICAgU2V0IG1hbmlmZXN0SXRlbSA9" & _
        "IGl0ZW0NCiAgICAgICAgRWxzZUlmIExDYXNlKGl0ZW0uTmFtZSkgPSBMQ2FzZShDT01fQ09O" & _
        "RklHX05BTUUpIFRoZW4NCiAgICAgICAgICAgIFNldCBjb25maWdJdGVtID0gaXRlbQ0KICAg" & _
        "ICAgICBFbmQgSWYNCiAgICBOZXh0DQoNCiAgICAnIENvcGlhciBhcmNoaXZvcyBhbCBkZXN0" & _
        "aW5vICgxNiA9IE5vIG1vc3RyYXIgZGnhbG9nbywgMTAyNCA9IE5vIGNvbmZpcm1hcikNCiAg" & _
        "ICBJZiBOb3QgZGxsSXRlbSBJcyBOb3RoaW5nIFRoZW4NCiAgICAgICAgZGVzdEZvbGRlck9i" & _
        "ai5Db3B5SGVyZSBkbGxJdGVtLCAxNiArIDEwMjQNCiAgICAgICAgV1NjcmlwdC5TbGVlcCA1" & _
        "MDANCiAgICBFbmQgSWYNCg0KICAgIElmIE5vdCBtYW5pZmVzdEl0ZW0gSXMgTm90aGluZyBU" & _
        "aGVuDQogICAgICAgIGRlc3RGb2xkZXJPYmouQ29weUhlcmUgbWFuaWZlc3RJdGVtLCAxNiAr" & _
        "IDEwMjQNCiAgICAgICAgV1NjcmlwdC5TbGVlcCA1MDANCiAgICBFbmQgSWYNCg0KICAgIElm" & _
        "IE5vdCBjb25maWdJdGVtIElzIE5vdGhpbmcgVGhlbg0KICAgICAgICBkZXN0Rm9sZGVyT2Jq" & _
        "LkNvcHlIZXJlIGNvbmZpZ0l0ZW0sIDE2ICsgMTAyNA0KICAgICAgICBXU2NyaXB0LlNsZWVw" & _
        "IDUwMA0KICAgIEVuZCBJZg0KDQogICAgJyBMaW1waWFyDQogICAgZnNvLkRlbGV0ZUZpbGUg" & _
        "dGVtcFppcA0KDQogICAgJyBWZXJpZmljYXINCiAgICBJZiBmc28uRmlsZUV4aXN0cyhkZXN0" & _
        "Rm9sZGVyICYgIlwiICYgQ09NX0RMTF9OQU1FKSBBbmQgXw0KICAgICAgIGZzby5GaWxlRXhp" & _
        "c3RzKGRlc3RGb2xkZXIgJiAiXCIgJiBDT01fQ09ORklHX05BTUUpIEFuZCBfDQogICAgICAg" & _
        "ZnNvLkZpbGVFeGlzdHMoZGVzdEZvbGRlciAmICJcIiAmIENPTV9NQU5JRkVTVF9OQU1FKSBU" & _
        "aGVuDQogICAgICAgIFRyeUV4dHJhY3RXaXRoU2hlbGwgPSBUcnVlDQogICAgRW5kIElmDQoN"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "CiAgICBPbiBFcnJvciBHb1RvIDANCkVuZCBGdW5jdGlvbg0KDQonID09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQonIEVMSU1JTkFDSdNO" & _
        "IERFIEFSQ0hJVk9TIENPTQ0KJyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PQ0KU3ViIFJlbW92ZUNPTUZpbGVzKGZvbGRlcikNCiAgICBP" & _
        "biBFcnJvciBSZXN1bWUgTmV4dA0KDQogICAgRGltIGRsbFBhdGgsIG1hbmlmZXN0UGF0aCwg" & _
        "Y29uZmlnUGF0aA0KICAgIGRsbFBhdGggPSBmb2xkZXIgJiAiXCIgJiBDT01fRExMX05BTUUN" & _
        "CiAgICBtYW5pZmVzdFBhdGggPSBmb2xkZXIgJiAiXCIgJiBDT01fTUFOSUZFU1RfTkFNRQ0K" & _
        "ICAgIGNvbmZpZ1BhdGggPSBmb2xkZXIgJiAiXCIgJiBDT01fQ09ORklHX05BTUUNCg0KICAg" & _
        "IElmIGZzby5GaWxlRXhpc3RzKGRsbFBhdGgpIFRoZW4NCiAgICAgICAgZnNvLkRlbGV0ZUZp" & _
        "bGUgZGxsUGF0aCwgVHJ1ZQ0KICAgIEVuZCBJZg0KDQogICAgSWYgZnNvLkZpbGVFeGlzdHMo" & _
        "bWFuaWZlc3RQYXRoKSBUaGVuDQogICAgICAgIGZzby5EZWxldGVGaWxlIG1hbmlmZXN0UGF0" & _
        "aCwgVHJ1ZQ0KICAgIEVuZCBJZg0KDQogICAgSWYgZnNvLkZpbGVFeGlzdHMoY29uZmlnUGF0" & _
        "aCkgVGhlbg0KICAgICAgICBmc28uRGVsZXRlRmlsZSBjb25maWdQYXRoLCBUcnVlDQogICAg" & _
        "RW5kIElmDQoNCiAgICBPbiBFcnJvciBHb1RvIDANCkVuZCBTdWINCg0KJyA9PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KJyBFTElNSU5B" & _
        "Q0nTTiBERUwgWExBTSBFWElTVEVOVEUNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT0NClN1YiBSZW1vdmVBZGRpbkluRGVzdGlubyhy" & _
        "dXRhRmluYWwpDQogICAgSWYgTm90IGZzby5GaWxlRXhpc3RzKHJ1dGFGaW5hbCkgVGhlbiBF" & _
        "eGl0IFN1Yg0KDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCiAgICBmc28uRGVsZXRlRmls" & _
        "ZSBydXRhRmluYWwsIFRydWUNCiAgICBPbiBFcnJvciBHb1RvIDANCg0KICAgIElmIE5vdCBm"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "c28uRmlsZUV4aXN0cyhydXRhRmluYWwpIFRoZW4gRXhpdCBTdWINCg0KICAgICcgRWwgYXJj" & _
        "aGl2byBzaWd1ZSBleGlzdGllbmRvLCBwb3NpYmxlbWVudGUgYmxvcXVlYWRvDQogICAgRGlt" & _
        "IG9ialdNSVNlcnZpY2UsIGNvbFByb2Nlc3NlcywgYW5zd2VyLCBvYmpQcm9jZXNzDQogICAg" & _
        "U2V0IG9ialdNSVNlcnZpY2UgPSBHZXRPYmplY3QoIndpbm1nbXRzOlxcLlxyb290XGNpbXYy" & _
        "IikNCiAgICBTZXQgY29sUHJvY2Vzc2VzID0gb2JqV01JU2VydmljZS5FeGVjUXVlcnkoIlNl" & _
        "bGVjdCAqIGZyb20gV2luMzJfUHJvY2VzcyBXaGVyZSBOYW1lID0gJ0VYQ0VMLkVYRSciKQ0K" & _
        "DQogICAgSWYgY29sUHJvY2Vzc2VzLkNvdW50ID4gMCBUaGVuDQogICAgICAgIGFuc3dlciA9" & _
        "IE1zZ0JveCgiRXhjZWwgZXN04SBlbiBlamVjdWNp824geSBwdWVkZSBlc3RhciBibG9xdWVh" & _
        "bmRvIGVsIGFyY2hpdm8gZGVsIGNvbXBsZW1lbnRvIGVuIGRlc3Rpbm8uIL9EZXNlYXMgY2Vy" & _
        "cmFyIEV4Y2VsPyIsIHZiWWVzTm8gKyB2YlF1ZXN0aW9uKQ0KICAgICAgICBJZiBhbnN3ZXIg" & _
        "PSB2YlllcyBUaGVuDQogICAgICAgICAgICBGb3IgRWFjaCBvYmpQcm9jZXNzIEluIGNvbFBy" & _
        "b2Nlc3Nlcw0KICAgICAgICAgICAgICAgIG9ialByb2Nlc3MuVGVybWluYXRlDQogICAgICAg" & _
        "ICAgICBOZXh0DQoNCiAgICAgICAgICAgICcgRXNwZXJhciBhIHF1ZSBFeGNlbCBjaWVycmUN" & _
        "CiAgICAgICAgICAgIFdTY3JpcHQuU2xlZXAgMzAwMA0KDQogICAgICAgICAgICAnIFJlaW50" & _
        "ZW50YXIgZWxpbWluYXINCiAgICAgICAgICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQogICAg" & _
        "ICAgICAgICBmc28uRGVsZXRlRmlsZSBydXRhRmluYWwsIFRydWUNCiAgICAgICAgICAgIE9u" & _
        "IEVycm9yIEdvVG8gMA0KDQogICAgICAgICAgICBJZiBmc28uRmlsZUV4aXN0cyhydXRhRmlu" & _
        "YWwpIFRoZW4NCiAgICAgICAgICAgICAgICBNc2dCb3ggIk5vIGhhIHNpZG8gcG9zaWJsZSBj" & _
        "b21wbGV0YXIgZWwgcHJvY2Vzby4gUG9yIGZhdm9yLCBjaWVycmEgRXhjZWwgbWFudWFsbWVu" & _
        "dGUgeSBlbGltaW5hIGVsIGZpY2hlcm8iICYgdmJDciAmICInIiAmIHJ1dGFGaW5hbCAmICIn"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "LiIsIHZiQ3JpdGljYWwNCiAgICAgICAgICAgICAgICBXU2NyaXB0LlF1aXQgMQ0KICAgICAg" & _
        "ICAgICAgRW5kIElmDQogICAgICAgIEVsc2UNCiAgICAgICAgICAgIE1zZ0JveCAiTm8gZXMg" & _
        "cG9zaWJsZSBjb21wbGV0YXIgZWwgcHJvY2Vzby4gUG9yIGZhdm9yLCBjaWVycmEgRXhjZWwg" & _
        "bWFudWFsbWVudGUgZSBpbnTpbnRhbG8gZGUgbnVldm8uIiwgdmJDcml0aWNhbA0KICAgICAg" & _
        "ICAgICAgV1NjcmlwdC5RdWl0IDENCiAgICAgICAgRW5kIElmDQogICAgRW5kIElmDQpFbmQg" & _
        "U3ViDQoNCicgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT0NCicgTElNUElFWkEgREVMIFJFR0lTVFJPDQonID09PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQpTdWIgQ2xlYW5SZWdpc3Ry" & _
        "eSh2ZXJzLCBub21icmUpDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCg0KICAgIERpbSBX" & _
        "c2hTaGVsbCwgaSwgY2xhdmUsIHZhbG9yDQogICAgU2V0IFdzaFNoZWxsID0gQ3JlYXRlT2Jq" & _
        "ZWN0KCJXU2NyaXB0LlNoZWxsIikNCg0KICAgIEZvciBpID0gMSBUbyA1MA0KICAgICAgICBj" & _
        "bGF2ZSA9ICJIS0VZX0NVUlJFTlRfVVNFUlxTb2Z0d2FyZVxNaWNyb3NvZnRcT2ZmaWNlXCIg" & _
        "JiB2ZXJzICYgIlxFeGNlbFxPcHRpb25zXE9QRU4iICYgaQ0KICAgICAgICB2YWxvciA9IFdz" & _
        "aFNoZWxsLlJlZ1JlYWQoY2xhdmUpDQoNCiAgICAgICAgSWYgRXJyLk51bWJlciA9IDAgVGhl" & _
        "bg0KICAgICAgICAgICAgSWYgSW5TdHIoMSwgdmFsb3IsIG5vbWJyZSAmICIueGxhbSIsIHZi" & _
        "VGV4dENvbXBhcmUpID4gMCBUaGVuDQogICAgICAgICAgICAgICAgV3NoU2hlbGwuUmVnRGVs" & _
        "ZXRlIGNsYXZlDQogICAgICAgICAgICAgICAgRXhpdCBGb3INCiAgICAgICAgICAgIEVuZCBJ" & _
        "Zg0KICAgICAgICBFbHNlDQogICAgICAgICAgICBFcnIuQ2xlYXINCiAgICAgICAgICAgIEV4" & _
        "aXQgRm9yDQogICAgICAgIEVuZCBJZg0KICAgIE5leHQNCg0KICAgIFNldCBXc2hTaGVsbCA9" & _
        "IE5vdGhpbmcNCiAgICBPbiBFcnJvciBHb1RvIDANCkVuZCBTdWINCg0KJyA9PT09PT09PT09"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KJyAgU1VCUlVU" & _
        "SU5BIFBBUkEgUkVHSVNUUkFSIENPTVBPTkVOVEUgQ09NIA0KJyA9PT09PT09PT09PT09PT09" & _
        "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KU3ViIFJlZ2lzdHJhckNs" & _
        "YXZlc0NPTSgpDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCiAgICBEaW0gc2hlbGwsIGFw" & _
        "cERhdGFQYXRoLCBhZGRpbnNQYXRoDQogICAgDQogICAgU2V0IHNoZWxsID0gQ3JlYXRlT2Jq" & _
        "ZWN0KCJXU2NyaXB0LlNoZWxsIikNCiAgICANCiAgICAnIE9idGVuZXIgcnV0YSBkZWwgQXBw" & _
        "RGF0YSBkZWwgdXN1YXJpbyBhY3R1YWwNCiAgICBhcHBEYXRhUGF0aCA9IHNoZWxsLkV4cGFu" & _
        "ZEVudmlyb25tZW50U3RyaW5ncygiJUFQUERBVEElIikNCiAgICBhZGRpbnNQYXRoID0gZnNv" & _
        "LkJ1aWxkUGF0aChhcHBEYXRhUGF0aCwgIk1pY3Jvc29mdFxBZGRJbnNcIikNCiAgICANCiAg" & _
        "ICAnIENyZWFyIGxhcyBjbGF2ZXMgcHJpbmNpcGFsZXMNCiAgICAnIDEuIENMU0lEIHByaW5j" & _
        "aXBhbA0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURc" & _
        "IiAmIEdVSURfQ0xTSUQgJiAiXCIsIENMQVNTX05BTUUsICJSRUdfU1oiDQogICAgc2hlbGwu" & _
        "UmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAm" & _
        "ICJcUHJvZ0lkXCIsIFBST0dfSUQsICJSRUdfU1oiDQogICAgDQogICAgJyAyLiBJbXBsZW1l" & _
        "bnRlZCBDYXRlZ29yaWVzDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xh" & _
        "c3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW1wbGVtZW50ZWQgQ2F0ZWdvcmllc1wi" & _
        "LCAiIiwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFz" & _
        "c2VzXENMU0lEXCIgJiBHVUlEX0NMU0lEICYgIlxJbXBsZW1lbnRlZCBDYXRlZ29yaWVzXHs2" & _
        "MkM4RkU2NS00RUJCLTQ1ZTctQjQ0MC02RTM5QjJDREJGMjl9XCIsICIiLCAiUkVHX1NaIg0K" & _
        "ICAgIA0KICAgICcgMy4gSW5wcm9jU2VydmVyMzINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtD"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "VVxTT0ZUV0FSRVxDbGFzc2VzXENMU0lEXCIgJiBHVUlEX0NMU0lEICYgIlxJbnByb2NTZXJ2" & _
        "ZXIzMlwiLCAibXNjb3JlZS5kbGwiLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJI" & _
        "S0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAiXElucHJvY1Nl" & _
        "cnZlcjMyXFRocmVhZGluZ01vZGVsIiwgIkJvdGgiLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJl" & _
        "Z1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcQ0xTSURcIiAmIEdVSURfQ0xTSUQgJiAi" & _
        "XElucHJvY1NlcnZlcjMyXENsYXNzIiwgQ0xBU1NfTkFNRSwgIlJFR19TWiINCiAgICBzaGVs" & _
        "bC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXENMU0lEXCIgJiBHVUlEX0NMU0lE" & _
        "ICYgIlxJbnByb2NTZXJ2ZXIzMlxBc3NlbWJseSIsIEFTU0VNQkxZX0lORk8sICJSRUdfU1oi" & _
        "DQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYg" & _
        "R1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVyMzJcUnVudGltZVZlcnNpb24iLCBSVU5USU1F" & _
        "X1ZFUlNJT04sICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVc" & _
        "Q2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVyMzJcQ29kZUJh" & _
        "c2UiLCAiZmlsZTovLy8iICYgUmVwbGFjZShhZGRpbnNQYXRoLCAiXCIsICIvIikgJiAiRm9s" & _
        "ZGVyV2F0Y2hlckNPTS5ETEwiLCAiUkVHX1NaIg0KICAgIA0KICAgICcgNC4gVmVyc2nzbiBl" & _
        "c3BlY+1maWNhDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xD" & _
        "TFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVyMzJcMS4wLjAuMFxDbGFzcyIs" & _
        "IENMQVNTX05BTUUsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdB" & _
        "UkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJcSW5wcm9jU2VydmVyMzJcMS4w" & _
        "LjAuMFxBc3NlbWJseSIsIEFTU0VNQkxZX0lORk8sICJSRUdfU1oiDQogICAgc2hlbGwuUmVn" & _
        "V3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xDTFNJRFwiICYgR1VJRF9DTFNJRCAmICJc"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "SW5wcm9jU2VydmVyMzJcMS4wLjAuMFxSdW50aW1lVmVyc2lvbiIsIFJVTlRJTUVfVkVSU0lP" & _
        "TiwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2Vz" & _
        "XENMU0lEXCIgJiBHVUlEX0NMU0lEICYgIlxJbnByb2NTZXJ2ZXIzMlwxLjAuMC4wXENvZGVC" & _
        "YXNlIiwgImZpbGU6Ly8vIiAmIFJlcGxhY2UoYWRkaW5zUGF0aCwgIlwiLCAiLyIpICYgIkZv" & _
        "bGRlcldhdGNoZXJDT00uRExMIiwgIlJFR19TWiINCiAgICANCiAgICAnIDUuIFByb2dJZA0K" & _
        "ICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcIiAmIFBST0dfSUQg" & _
        "JiAiXCIsIENMQVNTX05BTUUsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1Vc" & _
        "U09GVFdBUkVcQ2xhc3Nlc1wiICYgUFJPR19JRCAmICJcQ0xTSURcIiwgR1VJRF9DTFNJRCwg" & _
        "IlJFR19TWiINCiAgICANCiAgICAnIDYuIEludGVyZmFjZXMNCiAgICBSZWdpc3RyYXJJbnRl" & _
        "cmZheiBHVUlEX0ludGVyZmFjZTEsICJfRm9sZGVyV2F0Y2hlciIsIFBST1hZU1RVQl9DTFNJ" & _
        "RDENCiAgICBSZWdpc3RyYXJJbnRlcmZheiBHVUlEX0ludGVyZmFjZTIsICJJRm9sZGVyV2F0" & _
        "Y2hlckV2ZW50cyIsIFBST1hZU1RVQl9DTFNJRDINCiAgICANCiAgICAnIDcuIFR5cGVMaWIN" & _
        "CiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXFR5cGVMaWJcIiAm" & _
        "IEdVSURfVHlwZUxpYiAmICJcIiwgIiIsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUg" & _
        "IkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xUeXBlTGliXCIgJiBHVUlEX1R5cGVMaWIgJiAiXDEu" & _
        "MFwiLCAiQ29tcG9uZW50ZSBDT00gbW9uaXRvcml6YWNp824gY2FycGV0YXMiLCAiUkVHX1Na" & _
        "Ig0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcVHlwZUxpYlwi" & _
        "ICYgR1VJRF9UeXBlTGliICYgIlwxLjBcMFwiLCAiIiwgIlJFR19TWiINCiAgICBzaGVsbC5S" & _
        "ZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXFR5cGVMaWJcIiAmIEdVSURfVHlwZUxp" & _
        "YiAmICJcMS4wXDBcd2luNjRcIiwgYWRkaW5zUGF0aCAmICJGb2xkZXJXYXRjaGVyQ09NLnRs"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "YiIsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nl" & _
        "c1xUeXBlTGliXCIgJiBHVUlEX1R5cGVMaWIgJiAiXDEuMFxGTEFHU1wiLCAiMCIsICJSRUdf" & _
        "U1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xUeXBlTGli" & _
        "XCIgJiBHVUlEX1R5cGVMaWIgJiAiXDEuMFxIRUxQRElSXCIsIGFkZGluc1BhdGgsICJSRUdf" & _
        "U1oiDQogICAgDQogICAgJyA4LiBSZWdpc3Ryb3MgV09XNjQzMk5vZGUgKHBhcmEgY29tcGF0" & _
        "aWJpbGlkYWQgMzItYml0KQ0KICAgIFJlZ2lzdHJhckludGVyZmF6V09XNjQgR1VJRF9JbnRl" & _
        "cmZhY2UxLCAiX0ZvbGRlcldhdGNoZXIiLCBQUk9YWVNUVUJfQ0xTSUQxDQogICAgUmVnaXN0" & _
        "cmFySW50ZXJmYXpXT1c2NCBHVUlEX0ludGVyZmFjZTIsICJJRm9sZGVyV2F0Y2hlckV2ZW50" & _
        "cyIsIFBST1hZU1RVQl9DTFNJRDINCiAgICANCiAgICBJZiBFcnIuTnVtYmVyID0gMCBUaGVu" & _
        "DQogICAgICAgIFdTY3JpcHQuRWNobyAiUmVnaXN0cm8gQ09NIGNvbXBsZXRhZG8gZXhpdG9z" & _
        "YW1lbnRlLiINCiAgICBFbHNlDQogICAgICAgIFdTY3JpcHQuRWNobyAiRXJyb3IgZHVyYW50" & _
        "ZSBlbCByZWdpc3RybzogIiAmIEVyci5EZXNjcmlwdGlvbg0KICAgIEVuZCBJZg0KRW5kIFN1" & _
        "Yg0KDQonID09PT09PT09PT0gRlVOQ0nTTiBBVVhJTElBUiBQQVJBIElOVEVSRkFDRVMgPT09" & _
        "PT09PT09PQ0KU3ViIFJlZ2lzdHJhckludGVyZmF6KGd1aWQsIG5vbWJyZSwgcHJveHlTdHVi" & _
        "Q2xzaWQpDQogICAgRGltIHNoZWxsDQogICAgU2V0IHNoZWxsID0gQ3JlYXRlT2JqZWN0KCJX" & _
        "U2NyaXB0LlNoZWxsIikNCiAgICANCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FS" & _
        "RVxDbGFzc2VzXEludGVyZmFjZVwiICYgZ3VpZCAmICJcIiwgbm9tYnJlLCAiUkVHX1NaIg0K" & _
        "ICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcSW50ZXJmYWNlXCIg" & _
        "JiBndWlkICYgIlxQcm94eVN0dWJDbHNpZDMyXCIsIHByb3h5U3R1YkNsc2lkLCAiUkVHX1Na" & _
        "Ig0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcSW50ZXJmYWNl"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "XCIgJiBndWlkICYgIlxUeXBlTGliXCIsIEdVSURfVHlwZUxpYiwgIlJFR19TWiINCiAgICBz" & _
        "aGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXEludGVyZmFjZVwiICYgZ3Vp" & _
        "ZCAmICJcVHlwZUxpYlxWZXJzaW9uXCIsICIxLjAiLCAiUkVHX1NaIg0KRW5kIFN1Yg0KDQon" & _
        "ID09PT09PT09PT0gRlVOQ0nTTiBBVVhJTElBUiBQQVJBIFdPVzY0MzJOb2RlID09PT09PT09" & _
        "PT0NClN1YiBSZWdpc3RyYXJJbnRlcmZheldPVzY0KGd1aWQsIG5vbWJyZSwgcHJveHlTdHVi" & _
        "Q2xzaWQpDQogICAgRGltIHNoZWxsDQogICAgU2V0IHNoZWxsID0gQ3JlYXRlT2JqZWN0KCJX" & _
        "U2NyaXB0LlNoZWxsIikNCiAgICANCiAgICAnIERvcyB1YmljYWNpb25lcyBkaWZlcmVudGVz" & _
        "IHBhcmEgV09XNjQNCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2Vz" & _
        "XFdPVzY0MzJOb2RlXEludGVyZmFjZVwiICYgZ3VpZCAmICJcIiwgbm9tYnJlLCAiUkVHX1Na" & _
        "Ig0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENsYXNzZXNcV09XNjQzMk5v" & _
        "ZGVcSW50ZXJmYWNlXCIgJiBndWlkICYgIlxQcm94eVN0dWJDbHNpZDMyXCIsIHByb3h5U3R1" & _
        "YkNsc2lkLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJlZ1dyaXRlICJIS0NVXFNPRlRXQVJFXENs" & _
        "YXNzZXNcV09XNjQzMk5vZGVcSW50ZXJmYWNlXCIgJiBndWlkICYgIlxUeXBlTGliXCIsIEdV" & _
        "SURfVHlwZUxpYiwgIlJFR19TWiINCiAgICBzaGVsbC5SZWdXcml0ZSAiSEtDVVxTT0ZUV0FS" & _
        "RVxDbGFzc2VzXFdPVzY0MzJOb2RlXEludGVyZmFjZVwiICYgZ3VpZCAmICJcVHlwZUxpYlxW" & _
        "ZXJzaW9uXCIsICIxLjAiLCAiUkVHX1NaIg0KICAgIA0KICAgIHNoZWxsLlJlZ1dyaXRlICJI" & _
        "S0NVXFNPRlRXQVJFXFdPVzY0MzJOb2RlXENsYXNzZXNcSW50ZXJmYWNlXCIgJiBndWlkICYg" & _
        "IlwiLCBub21icmUsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3JpdGUgIkhLQ1VcU09GVFdB" & _
        "UkVcV09XNjQzMk5vZGVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIGd1aWQgJiAiXFByb3h5U3R1" & _
        "YkNsc2lkMzJcIiwgcHJveHlTdHViQ2xzaWQsICJSRUdfU1oiDQogICAgc2hlbGwuUmVnV3Jp"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "dGUgIkhLQ1VcU09GVFdBUkVcV09XNjQzMk5vZGVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIGd1" & _
        "aWQgJiAiXFR5cGVMaWJcIiwgR1VJRF9UeXBlTGliLCAiUkVHX1NaIg0KICAgIHNoZWxsLlJl" & _
        "Z1dyaXRlICJIS0NVXFNPRlRXQVJFXFdPVzY0MzJOb2RlXENsYXNzZXNcSW50ZXJmYWNlXCIg" & _
        "JiBndWlkICYgIlxUeXBlTGliXFZlcnNpb25cIiwgIjEuMCIsICJSRUdfU1oiDQpFbmQgU3Vi" & _
        "DQoNCicgPT09PT09PT09PSBTVUJSVVRJTkEgUEFSQSBFTElNSU5BUiA9PT09PT09PT09DQpT" & _
        "dWIgRWxpbWluYXJDbGF2ZXNDT00oKQ0KICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQogICAg" & _
        "RGltIHNoZWxsDQogICAgDQogICAgU2V0IHNoZWxsID0gQ3JlYXRlT2JqZWN0KCJXU2NyaXB0" & _
        "LlNoZWxsIikNCiAgICANCiAgICAnIEVsaW1pbmFyIGVuIG9yZGVuIGludmVyc28gKGRlIG3h" & _
        "cyBlc3BlY+1maWNvIGEgbeFzIGdlbmVyYWwpDQogICAgDQogICAgJyAxLiBFbGltaW5hciBX" & _
        "T1c2NDMyTm9kZSBlbnRyaWVzDQogICAgRWxpbWluYXJTaUV4aXN0ZSAiSEtDVVxTT0ZUV0FS" & _
        "RVxXT1c2NDMyTm9kZVxDbGFzc2VzXEludGVyZmFjZVwiICYgR1VJRF9JbnRlcmZhY2UxICYg" & _
        "IlwiDQogICAgRWxpbWluYXJTaUV4aXN0ZSAiSEtDVVxTT0ZUV0FSRVxXT1c2NDMyTm9kZVxD" & _
        "bGFzc2VzXEludGVyZmFjZVwiICYgR1VJRF9JbnRlcmZhY2UyICYgIlwiDQogICAgRWxpbWlu" & _
        "YXJTaUV4aXN0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXFdPVzY0MzJOb2RlXEludGVyZmFj" & _
        "ZVwiICYgR1VJRF9JbnRlcmZhY2UxICYgIlwiDQogICAgRWxpbWluYXJTaUV4aXN0ZSAiSEtD" & _
        "VVxTT0ZUV0FSRVxDbGFzc2VzXFdPVzY0MzJOb2RlXEludGVyZmFjZVwiICYgR1VJRF9JbnRl" & _
        "cmZhY2UyICYgIlwiDQogICAgDQogICAgJyAyLiBFbGltaW5hciBUeXBlTGliDQogICAgRWxp" & _
        "bWluYXJTaUV4aXN0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXFR5cGVMaWJcIiAmIEdVSURf" & _
        "VHlwZUxpYiAmICJcIg0KICAgIA0KICAgICcgMy4gRWxpbWluYXIgSW50ZXJmYWNlcyBub3Jt" & _
        "YWxlcw0KICAgIEVsaW1pbmFyU2lFeGlzdGUgIkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xJbnRl"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "cmZhY2VcIiAmIEdVSURfSW50ZXJmYWNlMSAmICJcIg0KICAgIEVsaW1pbmFyU2lFeGlzdGUg" & _
        "IkhLQ1VcU09GVFdBUkVcQ2xhc3Nlc1xJbnRlcmZhY2VcIiAmIEdVSURfSW50ZXJmYWNlMiAm" & _
        "ICJcIg0KICAgIA0KICAgICcgNC4gRWxpbWluYXIgUHJvZ0lkDQogICAgRWxpbWluYXJTaUV4" & _
        "aXN0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXCIgJiBQUk9HX0lEICYgIlwiDQogICAgDQog" & _
        "ICAgJyA1LiBFbGltaW5hciBDTFNJRCAoZXN0byBlbGltaW5hcuEgdG9kYSBsYSBqZXJhcnF1" & _
        "7WEpDQogICAgRWxpbWluYXJTaUV4aXN0ZSAiSEtDVVxTT0ZUV0FSRVxDbGFzc2VzXENMU0lE" & _
        "XCIgJiBHVUlEX0NMU0lEICYgIlwiDQogICAgDQogICAgSWYgRXJyLk51bWJlciA9IDAgVGhl" & _
        "bg0KICAgICAgICBXU2NyaXB0LkVjaG8gIkVsaW1pbmFjafNuIGRlIGNsYXZlcyBDT00gY29t" & _
        "cGxldGFkYSBleGl0b3NhbWVudGUuIg0KICAgIEVsc2UNCiAgICAgICAgV1NjcmlwdC5FY2hv" & _
        "ICJFcnJvciBkdXJhbnRlIGxhIGVsaW1pbmFjafNuOiAiICYgRXJyLkRlc2NyaXB0aW9uDQog" & _
        "ICAgRW5kIElmDQpFbmQgU3ViDQoNCicgPT09PT09PT09PSBGVU5DSdNOIEFVWElMSUFSIFBB" & _
        "UkEgRUxJTUlOQUNJ004gU0VHVVJBID09PT09PT09PT0NClN1YiBFbGltaW5hclNpRXhpc3Rl" & _
        "KHJ1dGEpDQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCiAgICBEaW0gc2hlbGwNCiAgICBT" & _
        "ZXQgc2hlbGwgPSBDcmVhdGVPYmplY3QoIldTY3JpcHQuU2hlbGwiKQ0KICAgIA0KICAgICcg" & _
        "SW50ZW50YXIgbGVlciBwYXJhIHZlciBzaSBleGlzdGUNCiAgICBzaGVsbC5SZWdSZWFkIHJ1" & _
        "dGENCiAgICANCiAgICBJZiBFcnIuTnVtYmVyID0gMCBUaGVuDQogICAgICAgICcgTGEgY2xh" & _
        "dmUgZXhpc3RlLCBlbGlt7W5hbGENCiAgICAgICAgRXJyLkNsZWFyDQogICAgICAgIHNoZWxs" & _
        "LlJlZ0RlbGV0ZSBydXRhDQogICAgICAgIElmIEVyci5OdW1iZXIgPD4gMCBUaGVuDQogICAg" & _
        "ICAgICAgICBXU2NyaXB0LkVjaG8gIiAgQWR2ZXJ0ZW5jaWE6IE5vIHNlIHB1ZG8gZWxpbWlu" & _
        "YXIgIiAmIHJ1dGENCiAgICAgICAgRW5kIElmDQogICAgRW5kIElmDQogICAgRXJyLkNsZWFy"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
        "DQpFbmQgU3ViDQoNCicgPT09PT09PT09PSBFSkVNUExPIERFIFVTTyA9PT09PT09PT09DQon" & _
        "IFBhcmEgcHJvYmFyIGxhcyBmdW5jaW9uZXM6DQonIFJlZ2lzdHJhckNsYXZlc0NPTSgpICAg" & _
        "JyBQYXJhIHJlZ2lzdHJhcg0KJyBFbGltaW5hckNsYXZlc0NPTSgpICAgICcgUGFyYSBlbGlt" & _
        "aW5hcg=="
End Function

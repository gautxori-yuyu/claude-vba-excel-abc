Attribute VB_Name = "modAPPInstallXLAM"
' ==========================================
' INSTALACIÓN Y DESINSTALACIÓN AUTOMÁTICA DEL COMPLEMENTO XLAM
' ==========================================
' Este módulo contiene la lógica de auto-instalación / auto-desinstalación
' del complemento XLAM en la carpeta de complementos del usuario, apoyándose
' en un script externo codificado en Base64 + RC4.
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
Attribute INSTALLSCRIPT_B64RC4.VB_ProcData.VB_Invoke_Func = " \n0"
    INSTALLSCRIPT_B64RC4 = _
                         "U2V0IGZzbyA9IENyZWF0ZU9iamVjdCgiU2NyaXB0aW5nLkZpbGVTeXN0ZW1PYmplY3QiKQ0K" & _
                         "U2V0IGFyZ3MgPSBXU2NyaXB0LkFyZ3VtZW50cw0KbW9kbyA9IGFyZ3MoMCkNCmFyY2hpdm8g" & _
                         "PSBhcmdzKDEpDQpkZXN0aW5vID0gYXJncygyKQ0Kbm9tYnJlID0gYXJncygzKQ0KDQpydXRh" & _
                         "RmluYWwgPSBkZXN0aW5vICYgIlwiICYgbm9tYnJlICYgIi54bGFtIg0KV1NjcmlwdC5TbGVl" & _
                         "cCAoNDAwMCkNCklmIG1vZG8gPSAiL2luc3RhbGwiIFRoZW4NCiAgICBJZiBOb3QgZnNvLkZp" & _
                         "bGVFeGlzdHMgKGFyY2hpdm8pIFRoZW4gTXNnQm94ICgiRXJyb3IgZGUgaW5zdGFsYWNpb246" & _
                         "IG5vIGV4aXN0ZSAnIiAmIGFyY2hpdm8gJiAiJyIpIDogV1NjcmlwdC5RdWl0DQoJUmVtb3Zl" & _
                         "QWRkaW5JbkRlc3Rpbm8gcnV0YUZpbmFsDQogICAgZnNvLkNvcHlGaWxlIGFyY2hpdm8sIHJ1" & _
                         "dGFGaW5hbA0KCVNldCBleGNlbCA9IENyZWF0ZU9iamVjdCgiRXhjZWwuQXBwbGljYXRpb24i" & _
                         "KQ0KCWV4Y2VsLlZpc2libGUgPSBGYWxzZQ0KICAgIEZvciBFYWNoIGFpIEluIGV4Y2VsLkFk" & _
                         "ZElucw0KICAgICAgICBJZiBhaS5OYW1lID0gbm9tYnJlICYgIi54bGFtIiBUaGVuDQogICAg" & _
                         "ICAgICAgICBhaS5JbnN0YWxsZWQgPSBUcnVlDQogICAgICAgICAgICBFeGl0IEZvcg0KICAg" & _
                         "ICAgICBFbmQgSWYNCiAgICBOZXh0DQogICAgV1NjcmlwdC5TbGVlcCAxMDAwDQogICAgSWYg" & _
                         "Tm90IGFpLkluc3RhbGxlZCBUaGVuDQogICAgCU1zZ0JveCAiTm8gaGEgc2lkbyBwb3NpYmxl" & _
                         "IGNvbXBsZXRhciBsYSBpbnN0YWxhY2nzbi4gUG9yIGZhdm9yLCBoYWJpbGl0YSBlbCBjb21w" & _
                         "bGVtZW50byBkZXNkZSBlbCBtZW76IGRlIGNvbXBsZW1lbnRvcyBkZSBFeGNlbC4iLCB2YkNy" & _
                         "aXRpY2FsDQogICAgRWxzZQ0KICAgIAlNc2dCb3ggIkluc3RhbGFjafNuIGNvbXBsZXRhZGEs" & _
                         "IHJlaW5pY2lhIEV4Y2VsLiIsIHZiSW5mb3JtYXRpb24NCiAgICBFbmQgaWYNCglleGNlbC5R" & _
                         "dWl0DQpFbHNlSWYgbW9kbyA9ICIvdW5pbnN0YWxsIiBUaGVuDQogICAgUmVtb3ZlQWRkaW5J" & _
                         "bkRlc3Rpbm8gcnV0YUZpbmFsDQoJU2V0IGV4Y2VsID0gQ3JlYXRlT2JqZWN0KCJFeGNlbC5B" & _
                         "cHBsaWNhdGlvbiIpDQoJdmVycyA9IEV4Y2VsLkFwcGxpY2F0aW9uLlZlcnNpb24NCglleGNl" & _
                         "bC5WaXNpYmxlID0gRmFsc2UNCiAgICBGb3IgRWFjaCBhaSBJbiBleGNlbC5BZGRJbnMNCiAg" & _
                         "ICAgICAgSWYgYWkuTmFtZSA9IG5vbWJyZSAmICIueGxhbSIgVGhlbg0KICAgICAgICAgICAg"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
                           "YWkuSW5zdGFsbGVkID0gRmFsc2UNCiAgICAgICAgICAgIEV4aXQgRm9yDQogICAgICAgIEVu" & _
                           "ZCBJZg0KICAgIE5leHQNCiAgICBJZiBhaS5JbnN0YWxsZWQgVGhlbg0KICAgIAlNc2dCb3gg" & _
                           "Ik5vIGhhIHNpZG8gcG9zaWJsZSBjb21wbGV0YXIgbGEgZGVzaW5zdGFsYWNp824uIFBvciBm" & _
                           "YXZvciwgcmVpbnTpbnRhbG8gZGUgbnVldm8gbyBkZXNoYWJpbGl0YSBlbCBjb21wbGVtZW50" & _
                           "byBkZXNkZSBlbCBtZW76IGRlIGNvbXBsZW1lbnRvcyBkZSBFeGNlbC4iLCB2YkNyaXRpY2Fs" & _
                           "DQogICAgRWxzZQ0KICAgIAlNc2dCb3ggIkRlc2luc3RhbGFjafNuIGNvbXBsZXRhZGEsIHJl" & _
                           "aW5pY2lhIEV4Y2VsLiIsIHZiSW5mb3JtYXRpb24NCiAgICBFbmQgSWYNCglleGNlbC5RdWl0" & _
                           "DQoJJyBib3JyYXIgbGFzIG1hcmNhcyBkZWwgcmVnaXN0cm8NCglTZXQgV3NoU2hlbGwgPSBD" & _
                           "cmVhdGVPYmplY3QoIldTY3JpcHQuU2hlbGwiKQ0KCUZvciBpID0gMSBUbyA1MA0KCQljbGF2" & _
                           "ZSA9ICJIS0VZX0NVUlJFTlRfVVNFUlxTb2Z0d2FyZVxNaWNyb3NvZnRcT2ZmaWNlXCIgJiB2" & _
                           "ZXJzICYgIlxFeGNlbFxPcHRpb25zXCIgJiAiT1BFTiIgJiBpDQogICAgICAgIE9uIEVycm9y" & _
                           "IFJlc3VtZSBOZXh0DQoJCXZhbG9yID0gV3NoU2hlbGwuUmVnUmVhZChjbGF2ZSkNCiAgICAg" & _
                           "ICAgSWYgRXJyIFRoZW4gU3RvcCA6IEV4aXQgZm9yDQogICAgICAgIE9uIEVycm9yIEdvVG8g" & _
                           "MA0KCQlJZiBJblN0cigxLCB2YWxvciwgbm9tYnJlICYgIi54bGFtIiwgdmJUZXh0Q29tcGFy" & _
                           "ZSkgPiAwIFRoZW4NCgkJCVdzaFNoZWxsLlJlZ0RlbGV0ZSBjbGF2ZQ0KCQkJRXhpdCBGb3IN" & _
                           "CgkJRW5kIElmDQoJTmV4dA0KICAgIE9uIEVycm9yIEdvVG8gMA0KRW5kIElmDQoNCmZzby5E" & _
                           "ZWxldGVGaWxlIFdTY3JpcHQuU2NyaXB0RnVsbE5hbWUNCg0KU3ViIFJlbW92ZUFkZGluSW5E" & _
                           "ZXN0aW5vIChydXRhRmluYWwpDQogICAgSWYgZnNvLkZpbGVFeGlzdHMocnV0YUZpbmFsKSBU" & _
                           "aGVuDQogICAgICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQogICAgICAgIGZzby5EZWxldGVG" & _
                           "aWxlIHJ1dGFGaW5hbA0KICAgICAgICBPbiBFcnJvciBHb1RvIDANCiAgICBFbmQgSWYNCiAg" & _
                           "ICBJZiBmc28uRmlsZUV4aXN0cyhydXRhRmluYWwpIFRoZW4NCiAgICAgICAgJyBDaGVjayBp" & _
                           "ZiBFeGNlbCBpcyBydW5uaW5nIGFuZCBvZmZlciB0byBjbG9zZSBpdA0KICAgICAgICBTZXQg" & _
                           "b2JqV01JU2VydmljZSA9IEdldE9iamVjdCgid2lubWdtdHM6XFwuXHJvb3RcY2ltdjIiKQ0K"
    INSTALLSCRIPT_B64RC4 = INSTALLSCRIPT_B64RC4 & _
                           "ICAgICAgICBTZXQgY29sUHJvY2Vzc2VzID0gb2JqV01JU2VydmljZS5FeGVjUXVlcnkoIlNl" & _
                           "bGVjdCAqIGZyb20gV2luMzJfUHJvY2VzcyBXaGVyZSBOYW1lID0gJ0VYQ0VMLkVYRSciKQ0K" & _
                           "ICAgICAgICBJZiBjb2xQcm9jZXNzZXMuQ291bnQgPiAwIFRoZW4NCiAgICAgICAgICAgIGFu" & _
                           "c3dlciA9IE1zZ0JveCgiRXhjZWwgZXN04SBlbiBlamVjdWNp824geSBwdWVkZSBlc3RhciBi" & _
                           "bG9xdWVhbmRvIGVsIGFyY2hpdm8gZGVsIGNvbXBsZW1lbnRvIGVuIGRlc3Rpbm8uIL9EZXNl" & _
                           "YXMgY2VycmFyIEV4Y2VsPyIsIHZiWWVzTm8gKyB2YlF1ZXN0aW9uKQ0KICAgICAgICAgICAg" & _
                           "SWYgYW5zd2VyID0gdmJZZXMgVGhlbg0KICAgICAgICAgICAgICAgIEZvciBFYWNoIG9ialBy" & _
                           "b2Nlc3MgaW4gY29sUHJvY2Vzc2VzDQogICAgICAgICAgICAgICAgICAgIG9ialByb2Nlc3Mu" & _
                           "VGVybWluYXRlDQogICAgICAgICAgICAgICAgTmV4dA0KICAgICAgICAgICAgICAgICcgV2Fp" & _
                           "dCBhIG1vbWVudCB0byBlbnN1cmUgRXhjZWwgaGFzIGNsb3NlZA0KICAgICAgICAgICAgICAg" & _
                           "IFdTY3JpcHQuU2xlZXAoMzAwMCkNCiAgICAgICAgICAgICAgICAnIFRyeSBkZWxldGluZyB0" & _
                           "aGUgZmlsZSBhZ2Fpbg0KICAgICAgICAgICAgICAgIE9uIEVycm9yIFJlc3VtZSBOZXh0DQog" & _
                           "ICAgICAgICAgICAgICAgZnNvLkRlbGV0ZUZpbGUgcnV0YUZpbmFsDQogICAgICAgICAgICAg" & _
                           "ICAgT24gRXJyb3IgR29UbyAwDQogICAgICAgICAgICAgICAgSWYgZnNvLkZpbGVFeGlzdHMo" & _
                           "cnV0YUZpbmFsKSBUaGVuDQogICAgICAgICAgICAgICAgICAgIE1zZ0JveCAiTm8gaGEgc2lk" & _
                           "byBwb3NpYmxlIGNvbXBsZXRhciBlbCBwcm9jZXNvLiBQb3IgZmF2b3IsIGNpZXJyYSBFeGNl" & _
                           "bCBtYW51YWxtZW50ZSB5IGVsaW1pbmEgZWwgZmljaGVybyIgJiB2YkNyICYgIiciICYgcnV0" & _
                           "YUZpbmFsICYgIicuIiwgdmJDcml0aWNhbA0KICAgICAgICAgICAgICAgICAgICBXU2NyaXB0" & _
                           "LlF1aXQgMQ0KICAgICAgICAgICAgICAgIEVuZCBJZg0KICAgICAgICAgICAgRWxzZQ0KICAg" & _
                           "ICAgICAgICAgICAgIE1zZ0JveCAiTm8gZXMgcG9zaWJsZSBjb21wbGV0YXIgZWwgcHJvY2Vz" & _
                           "by4gUG9yIGZhdm9yLCBjaWVycmEgRXhjZWwgbWFudWFsbWVudGUgZSBpbnTpbnRhbG8gZGUg" & _
                           "bnVldm8uIiwgdmJDcml0aWNhbA0KICAgICAgICAgICAgICAgIFdTY3JpcHQuUXVpdCAxDQog" & _
                           "ICAgICAgICAgICBFbmQgSWYNCiAgICAgICAgRW5kIElmDQogICAgRW5kIElmDQpFbmQgU3Vi"
End Function



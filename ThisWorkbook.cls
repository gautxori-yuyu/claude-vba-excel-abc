VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "1-Inicio e Instalacion"
Option Explicit
Private mApp As clsAplicacion

Friend Function App() As clsAplicacion
    If mApp Is Nothing Then
        Set mApp = New clsAplicacion
    End If
    Set App = mApp
End Function

Public Sub TerminateApp()
    If Not mApp Is Nothing Then
        mApp.Terminate
        Set mApp = Nothing
    End If
End Sub

Private Sub Workbook_Open()
    Debug.Print "[ThisWorkbook - Open]"
    ' Sistema de instalación existente
    AutoInstalador
    AutoRegistrarTodasLasUDFs
    ' Ctrl + Shift + R
    Application.OnKey "^+R", "ToggleRibbonTab"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Este evento NO debe desregistrar las UDFs normalmente,
    ' solo si detecta que el complemento fue deshabilitado manualmente
    
    On Error Resume Next
    
    ' Si fue deshabilitado pero no se disparó AddinUninstall, limpiar
    If Not ComprobarSiInstalado Then
    End If
    
    Application.OnKey "^+R"
    
    TerminateApp
    
    DesregistrarTodasLasUDFs
    Debug.Print "BeforeClose: UDFs desregistradas (complemento deshabilitado)"

    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Regenerar listas automáticamente antes de guardar
    'Call GenerarListasUnidades
End Sub

Private Sub Workbook_AddinUninstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario desmarca la casilla en Archivo > Opciones > Complementos
    ' 2. Se ejecuta ai.Installed = False desde el script VBS
    
    On Error Resume Next
    
    ' Mensaje de despedida (opcional)
    '    MsgBox "Complemento desinstalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas han sido eliminadas.", vbInformation, "Desinstalación completada"
    
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    ' Este evento se dispara cuando:
    ' 1. El usuario marca la casilla en Archivo > Opciones > Complementos
    ' 2. El script VBS completa la instalación y marca installed=True
    
    On Error Resume Next
    
    ' Mensaje de bienvenida (opcional, puedes comentarlo si es molesto)
    '    MsgBox "Complemento 'ABC Ofertas Máquina Especial' instalado correctamente." & vbCrLf & _
    '           "Las funciones personalizadas ya están disponibles.", vbInformation, "Instalación completada"
    
    On Error GoTo 0
End Sub



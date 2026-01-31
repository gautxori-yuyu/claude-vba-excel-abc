VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "1-Inicio e Instalacion.Ciclo de vida"
Option Explicit

Private Const MODULE_NAME As String = "ThisWorkbook"

Private mApp As clsApplication

Friend Function App() As clsApplication
    If mApp Is Nothing Then
        Set mApp = New clsApplication
    End If
    Set App = mApp
End Function

Public Sub TerminateApp()
    If Not mApp Is Nothing Then
        Set mApp = Nothing
    End If
End Sub

Private Sub Workbook_Open()
    ' 0. Inicializar logger con salida a fichero
    removeOldLogs
    InitLogger LOG_DEBUG, True, ThisWorkbook.Path & "\ABC_Log_" & Format(Now, "yyyy-mm-dd_hh-nn") & ".txt"
    
    LogInfo MODULE_NAME, "[Open] - Abriendo Workbook " & ThisWorkbook.Name
    ' Sistema de instalación existente
    AutoInstalador
    AutoRegistrarTodasLasUDFs
    ' Ctrl + Shift + R
    Application.OnKey "^+R", "ToggleRibbonTab"

    ' Iniciar aplicación (lazy init del singleton)
    App
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Este evento NO debe desregistrar las UDFs normalmente,
    ' solo si detecta que el complemento fue deshabilitado manualmente

    On Error GoTo ErrHandler

    ' Si fue deshabilitado pero no se disparo AddinUninstall, limpiar
    If Not ComprobarSiInstalado Then
        LogWarning MODULE_NAME, "[Workbook_BeforeClose] - Complemento no instalado, realizando limpieza"
    End If

    ' Limpiar atajos de teclado
    Application.OnKey "^+R"

    ' Terminar aplicacion de forma segura
    TerminateApp

    ' Desregistrar UDFs
    DesregistrarTodasLasUDFs
    LogInfo MODULE_NAME, "[Workbook_BeforeClose] - UDFs desregistradas"
    
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[Workbook_BeforeClose]"
    ' Intentar limpieza minima en caso de error
    On Error Resume Next
    Application.OnKey "^+R"
    TerminateApp
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
' Elimina los ficheros de log que tengan más de 3 días, si hay más de 3 ficheros de log
Private Sub removeOldLogs()
    Dim fso As Object
    Dim folderPath As String
    Dim oFolder As Object
    Dim oFile As Object
    Dim logFiles As Object ' Collection para filtrar solo logs de este tipo
    Dim fileDate As Date
    Dim fileNamePattern As String, i
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.Path & "\"
    fileNamePattern = "ABC_Log_????-??-??_??-??.txt"
    
    ' 1. Coleccionamos los archivos que coinciden con el patrón
    Set logFiles = New Collection
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set oFolder = fso.GetFolder(folderPath)
    
    For Each oFile In oFolder.Files
        ' Verificamos si el nombre coincide con el patrón ABC_Log_...
        If oFile.Name Like fileNamePattern Then
            If logFiles.Count = 0 Then
                logFiles.Add oFile
            ElseIf (oFile.DateLastAccessed < logFiles.Item(logFiles.Count).DateLastAccessed) Then
                logFiles.Add Item:=oFile, before:=logFiles.Count
            Else
                logFiles.Add oFile
            End If
        End If
    Next oFile
    
    ' 2. Si hay más de 3 ficheros en total, evaluamos el borrado
    Do While logFiles.Count > 3
        logFiles.Item(1).Delete True
        logFiles.Remove 1
    Loop
    For Each oFile In logFiles
        ' Calculamos la diferencia de días (DateDiff)
        ' Usamos DateLastModified para mayor precisión, o DateCreated
        fileDate = oFile.DateLastModified
        
        ' Requisito: Más de 3 días de antigüedad
        If DateDiff("d", fileDate, Now) > 3 Then
            ' Re-verificamos el conteo antes de cada borrado para asegurar que siempre queden 3
            If logFiles.Count > 3 Then
                On Error Resume Next ' Por si el archivo está abierto o bloqueado
                oFile.Delete True
                If Err.Number = 0 Then
                    ' Actualizar lógica si es necesario, aunque al ser una colección
                    ' se recorre una foto fija. El "If logFiles.Count > 3"
                    ' aquí dentro es preventivo.
                End If
                On Error GoTo 0
            End If
        Else
            Exit For
        End If
    Next oFile
    
    ' Limpieza de objetos
    Set oFile = Nothing
    Set oFolder = Nothing
    Set fso = Nothing
    LogInfo MODULE_NAME, "[removeOldLogs] - Finalizado"
End Sub



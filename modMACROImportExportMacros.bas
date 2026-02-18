Attribute VB_Name = "modMACROImportExportMacros"
' ==============================================================================================================
' MÓDULO: modMACROImportExportMacros
' DESCRIPCIÓN: Módulo para exportar e importar componentes VBA (módulos, clases, formularios) desde y hacia
'              archivos físicos. Permite hacer backup del código VBA o transferirlo entre proyectos.
' REQUISITOS: - Referencia a "Microsoft Visual Basic for Applications Extensibility 5.3"
'             - Acceso al modelo de objetos VBA habilitado en las opciones de seguridad de Excel
' ==============================================================================================================

'@NOTE: Debes tener habilitado el acceso al modelo de objetos de VBA:
' - En el editor de VBA: ve a Herramientas > Referencias. Marca "Microsoft Visual Basic for Applications Extensibility 5.3".
' - en Excel: Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza
'       > Configuración de macros > marca "Confiar en el acceso al modelo de objetos del proyecto VBA".

'@Folder "9-Developer"
Option Explicit

Private Const MODULE_NAME As String = "modMACROImportExportMacros"

' -------------------------------------------------------------------------------------------------------------
' EXPORTACIÓN DE COMPONENTES VBA
' -------------------------------------------------------------------------------------------------------------

'@Description: Exporta todos los componentes VBA del libro seleccionado (módulos, clases, formularios)
'              a archivos individuales en la carpeta del libro
'@Scope: (muestra formulario de selección)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmImportExportMacros (formulario de selección de libro)
'@Note: Los archivos se guardan con extensiones: .bas (módulos), .cls (clases), .frm (formularios)
Sub ExportarComponentesVBA()
Attribute ExportarComponentesVBA.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim vbComp As Object
    Dim rutaExportacion As String
    Dim nombreArchivo As String
    Dim Wb As Workbook
    
    Dim frm As New frmImportExportMacros
    frm.Show vbModal
    If frm.WorkbookSeleccionado Is Nothing Then Exit Sub
    Set Wb = frm.WorkbookSeleccionado
    Unload frm
    If Wb Is Nothing Then Exit Sub               ' Cancelado o error
    
    ' Carpeta donde se guardarán los archivos exportados
    rutaExportacion = Wb.path
    
    ExportarFichsVBAaCarpeta Wb, rutaExportacion
    
    MsgBox "Componentes exportados a: " & rutaExportacion, vbInformation
End Sub

Sub ExportarComponentesVBAdesdeThisWorkbookXLAM()
Attribute ExportarComponentesVBAdesdeThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta donde se guardarán los archivos exportados
    Dim rutaExportacion As String
    rutaExportacion = ThisWorkbook.path
    
    ExportarFichsVBAaCarpeta ThisWorkbook, rutaExportacion
    
    MsgBox "Exportación completada en: " & rutaExportacion, vbInformation
End Sub

'@Description: Exporta componentes VBA sin mostrar mensajes al usuario
'@Scope: Privado
'@ArgumentDescriptions: wb: Workbook de origen | rutaDestino: Carpeta donde exportar
Public Sub ExportarFichsVBAaCarpeta(Wb As Workbook, rutaDestino As String)
    
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Asegurar que existe la carpeta
    If Not fso.FolderExists(rutaDestino) Then
        fso.CreateFolder rutaDestino
    End If
    
    Dim vbComp As Object
    Dim nombreArchivo As String
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In Wb.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: nombreArchivo = vbComp.Name & ".bas"  ' Módulo estándar
            Case 2, 100: nombreArchivo = vbComp.Name & ".cls"  ' Clase o documento
            Case 3: nombreArchivo = vbComp.Name & ".frm"  ' Formulario
            Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        
        ' Exportar solo si tiene código, Y SI NO ES CODIGO EN UNA HOJA DE CALCULO SIN RENOMBRAR
        ' (por evitar exportar codigo "de pruebas")
        If vbComp.CodeModule.CountOfLines = 0 And InStr(vbComp.Name, "Hoja") > 0 Then
        Else
            vbComp.Export rutaDestino & "\" & nombreArchivo
        End If
    Next vbComp
    
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------------------------------------------
' IMPORTACIÓN DE COMPONENTES VBA
' -------------------------------------------------------------------------------------------------------------

'@Description: Importa componentes VBA desde archivos físicos al libro seleccionado. Permite seleccionar
'              múltiples archivos (.bas, .cls, .frm)
'@Scope:  (muestra formularios de selección)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmImportExportMacros (formulario de selección de libro)
'@Note: Elimina el componente existente si ya existe uno con el mismo nombre antes de importar
Sub ImportarComponentesVBA()
Attribute ImportarComponentesVBA.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim fso As Object, carpeta As Object, archivo As Object
    Dim rutaImportacion As String
    Dim extension As String
    Dim Wb As Workbook
    
    Dim frm As New frmImportExportMacros
    frm.Show vbModal
    If frm.WorkbookSeleccionado Is Nothing Then Exit Sub
    Set Wb = frm.WorkbookSeleccionado
    Unload frm
    If Wb Is Nothing Then Exit Sub               ' Cancelado o error
    
    ' Carpeta desde donde se importarán los archivos
    rutaImportacion = Wb.path
    
    ImportarFichsVBAenCarpeta rutaImportacion
       
    MsgBox "Importación completada desde: " & rutaImportacion, vbInformation
End Sub

' ==========================================
' FUNCIÓN 5: UTILIDAD PARA RESTAURAR DESDE BACKUP
' ==========================================

'@Description: Restaura código VBA desde un archivo ZIP de backup
'@Scope: Público
'@ArgumentDescriptions: rutaZip: Ruta completa del archivo ZIP con el backup
Public Sub RestaurarBackupVBADesdeZip(Optional rutaZip As String)
Attribute RestaurarBackupVBADesdeZip.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim rutaTempDescompresion As String
    Dim timestampStr As String
    Dim fso As Object
    Dim shellApp As Object
    
    On Error GoTo ErrorHandler
    
    ' Si no se proporciona ruta, pedir al usuario
    If rutaZip = "" Then
        rutaZip = Application.GetOpenFilename("Archivos ZIP (*.zip), *.zip", , "Seleccionar backup ZIP")
        If rutaZip = "False" Then Exit Sub  ' Cancelado
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar que existe
    If Not fso.FileExists(rutaZip) Then
        MsgBox "El archivo ZIP no existe: " & rutaZip, vbExclamation, "Archivo no encontrado"
        Exit Sub
    End If
    
    ' Confirmar restauración
    If MsgBox("¿Desea restaurar el código VBA desde este backup?" & vbCrLf & vbCrLf & _
              "ADVERTENCIA: Se eliminarán todos los módulos actuales" & vbCrLf & _
              "y se cargarán los del backup.", vbExclamation + vbYesNo, "Confirmar restauración") <> vbYes Then
        Exit Sub
    End If
    
    Dim nfichs
    nfichs = ContarItemsEnZip(shellApp, rutaZip)
    
    ' Crear carpeta temporal para descomprimir
    timestampStr = Format(Now, "yyyymmdd_hhnnss")
    rutaTempDescompresion = Environ("TEMP") & "\VBA_Restore_" & timestampStr
    fso.CreateFolder rutaTempDescompresion
    
    ' Descomprimir
    Set shellApp = CreateObject("Shell.Application")
    shellApp.Namespace(rutaTempDescompresion).CopyHere shellApp.Namespace(rutaZip).Items
    
    ' Esperar a que termine la descompresión
    Dim intentos As Integer
    intentos = 0
    Do While ContarArchivosRecursivo(fso.GetFolder(rutaTempDescompresion)) < nfichs And intentos < 50
        DoEvents
        Sleep 200
        intentos = intentos + 1
    Loop
        
    ImportarFichsVBAenCarpeta rutaTempDescompresion
    
    ' Limpiar carpeta temporal
    On Error Resume Next
    fso.DeleteFolder rutaTempDescompresion, True
    On Error GoTo 0
        
    MsgBox "Restauración completada desde: " & rutaZip, vbInformation, "Restauración completada"
    
    Exit Sub
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[RestaurarBackupVBADesdeZip]"
    MsgBox "Error al restaurar backup: " & Err.Description, vbCritical, "Error"
End Sub

'@Description: Importa ficheros VBA al XLAM actual; incluye dependencias (inyeccion en el XLAM)
Sub ImportarComponentesVBAaThisWorkbookXLAM()
Attribute ImportarComponentesVBAaThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta desde donde se importarán los archivos
    Dim rutaImportacion As String
    rutaImportacion = ThisWorkbook.path
    
    ImportarFichsVBAenCarpeta rutaImportacion , True  ': REEMPLAZA LOS MODULOS QUE YA EXISTEN, sin confirmacion.
    
    If MsgBox("¿Importar Hojas de calculo desde XLAM ABC?", vbYesNo + vbQuestion, "Importar dependencias") = vbYes Then
        ImportarDependencias
    End If
    
    MsgBox "Importación completada desde: " & rutaImportacion & ". Falta, en el 'proyecto ABC', inyectar los ficheros a descomprimir (FSWatcher). Requiere hacerlo con Excel cerrado, usar VBScript.", vbInformation
End Sub

Private Sub ImportarDependencias()
    Dim rutaLibro As Variant
    Dim wbOrigen As Workbook
    Dim ws As Worksheet
    Dim nombreHojaBorrar As String: nombreHojaBorrar = "Hoja1"
    
    rutaLibro = Application.GetOpenFilename("Excel Files (*.xlam; *.xlsm), *.xlam; *.xlsm", , "Seleccionar versión previa")
    
    If rutaLibro = False Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Al abrir el libro con EnableEvents = False, NO se ejecutará su Workbook_Open
    Set wbOrigen = Workbooks.Open(fileName:=rutaLibro, ReadOnly:=True)
    
    ' Gestión de .xlam para permitir el copiado
    Dim bOrigenAddin As Boolean, bThisWBAddin As Boolean
    bOrigenAddin = wbOrigen.IsAddin
    If bOrigenAddin Then wbOrigen.IsAddin = False
    
    ' El metodo COPY NO FUNCIONA SI EL DESINATARIO ESTA ACTIVADO COMO ADDIN!
    bThisWBAddin = ThisWorkbook.IsAddin
    If bThisWBAddin Then ThisWorkbook.IsAddin = False
    
    For Each ws In wbOrigen.Worksheets
        If InStr(1, ws.Name, "Hoja", vbTextCompare) = 0 Then
            ' Al copiar la hoja con EnableEvents = False, NO se disparará Worksheet_Activate en el destino
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        End If
    Next ws
    
    ' Limpieza de hoja inicial
    On Error Resume Next
    If ThisWorkbook.Sheets.Count > 1 Then
        Set ws = ThisWorkbook.Sheets(nombreHojaBorrar)
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    End If
    On Error GoTo 0

CleanExit:
    If Not wbOrigen Is Nothing Then
        If bOrigenAddin Then wbOrigen.IsAddin = True
        wbOrigen.Close SaveChanges:=False
    End If
    If bThisWBAddin Then ThisWorkbook.IsAddin = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    LogCurrentError MODULE_NAME, "[ImportarDependencias]"
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error al importar"
    Resume CleanExit
End Sub

Private Sub ImportarFichsVBAenCarpeta(rutaImportacion As String, Optional bRemove As Boolean = False)
    Dim fso As Object, archivo As Object, carpeta As Object
    Dim extension As String, nombreComp As String
    Dim vbComp As Object
    Dim ts As Object, contenido As String
    
    ' Configuración de FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Validación de Carpeta
    If Not fso.FolderExists(rutaImportacion) Then
        MsgBox "La carpeta de importación no existe: " & rutaImportacion, vbExclamation
        Exit Sub
    End If
    Set carpeta = fso.GetFolder(rutaImportacion)

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' PRIMERO se añaden las DEPENDENCIAS de objetos COM "ajenos" (usar ListarReferenciasActuales para listarlas)
    ' las que añade esta aplicación:
    ' Descripción   GUID    Major | Minor   Ruta
    ' FolderWatcher: Componente COM monitorización carpetas {E0BCC03C-D155-4EA3-BCB8-1D071719E854}  1 | 0   C:\Users\Srey\AppData\Roaming\Microsoft\AddIns\FolderWatcherCOM.tlb
    ' Microsoft Visual Basic for Applications Extensibility 5.3 {0002E157-0000-0000-C000-000000000046}  5 | 3   C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    ' Microsoft Scripting Runtime   {420B2830-E718-11CF-893D-00A0C9054228}  1 | 0   C:\Windows\System32\scrrun.dll
    
    ' las siguientes,
    ' Descripción   GUID    Ruta
    ' Visual Basic For Applications {000204EF-0000-0000-C000-000000000046}  C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL
    ' Microsoft Excel 16.0 Object Library   {00020813-0000-0000-C000-000000000046}  C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE
    ' OLE Automation    {00020430-0000-0000-C000-000000000046}  C:\Windows\System32\stdole2.tlb
    ' Microsoft Office 16.0 Object Library  {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}  C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL
    ' Microsoft Forms 2.0 Object Library    {0D452EE1-E08F-101A-852E-02608C4D0BB4}  C:\WINDOWS\system32\FM20.DLL
    ' son estandares del sistema).
    Dim WshShell
    Set WshShell = CreateObject("Wscript.Shell")
    Call AgregarReferenciaPorRuta(WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns\FolderWatcherCOM.tlb", "FolderWatcher")
    Call AgregarReferenciaPorRuta(WshShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB", "VBE")
    Call AgregarReferenciaPorGUID("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0, "Scripting Runtime")

    For Each archivo In carpeta.files
        extension = LCase(fso.GetExtensionName(archivo.Name))
        nombreComp = fso.GetBaseName(archivo.Name)
        
        ' Solo procesamos archivos de código
        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            
            ' CASO A: Componente especial (ThisWorkbook o Hojas)
            If extension = "cls" And (nombreComp = "ThisWorkbook" Or nombreComp Like "Hoja*") Then
                Set vbComp = ThisWorkbook.VBProject.VBComponents(nombreComp)
                
                ' Leer contenido y limpiar encabezados de clase
                Set ts = fso.OpenTextFile(archivo.path, 1)
                contenido = ts.ReadAll
                ts.Close
                
                ' Limpiar atributos de clase (el encabezado que causa error)
                contenido = LimpiarAtributosClase(contenido)
                
                ' Decidir si sobreescribir
                If vbComp.CodeModule.CountOfLines <= 1 Then
                    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                    vbComp.CodeModule.AddFromString contenido
                Else
                    If bRemove Or MsgBox("¿Sobreescribir código en " & nombreComp & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Clase existente") = vbYes Then
                            ' Limpiar código existente antes de inyectar el nuevo
                        vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                            ' Insertar el contenido del archivo
                        vbComp.CodeModule.AddFromString contenido
                    End If
                End If
                
            ' CASO B: Módulos, Clases nuevas o Formularios
            Else
                On Error Resume Next
                Set vbComp = Nothing
                ' Intentamos asignar el componente a una variable
                Set vbComp = ThisWorkbook.VBProject.VBComponents(nombreComp)
                On Error GoTo ErrorHandler
                
                ' Si la variable es (Nothing), el componente no existe
                If vbComp Is Nothing Then
                    ' Importar
                    ThisWorkbook.VBProject.VBComponents.Import archivo.path
                Else
                    Select Case True
                        Case nombreComp = "mod_Logger", nombreComp = "modMACROImportExportMacros"
                            ' ESTOS MODULOS ESTAN EN EJECUCION, NO SE PUEDEN ACTUALIZAR al vuelo, HACERLO A MANO
                            MsgBox ("El modulo " & nombreComp & " hay que actualizarlo a mano, NO se puede actualizar automaticamente - tienen código en ejecución")
                        Case bRemove, MsgBox("¿Eliminar el componente " & nombreComp & "?", vbYesNo + vbDefaultButton2, "Clase existente") = vbYes
                            ThisWorkbook.VBProject.VBComponents.Remove vbComp
                                        
                            ' Importar
                            ThisWorkbook.VBProject.VBComponents.Import archivo.path
                    End Select
                End If
            End If
            
        End If
    Next archivo
    
CleanExit:
    On Error Resume Next ' Evita errores en la propia salida
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Set fso = Nothing
    Exit Sub
   
ErrorHandler:
    LogCurrentError MODULE_NAME, "[ImportarFichsVBAenCarpeta]"
    MsgBox "Error crítico al importar '" & archivo.Name & "':" & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Función auxiliar para limpiar el encabezado de los archivos .cls
Private Function LimpiarAtributosClase(ByVal texto As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = True
    re.Pattern = "VERSION 1\.0 CLASS[\s\r\n]*begin[\s\S]*?end\s*\r\n|Attribute.+\r\n"
    
    LimpiarAtributosClase = re.Replace(texto, "")
End Function

Sub ListarReferenciasActuales()
    Dim ref As Object
    For Each ref In ThisWorkbook.VBProject.References
        Debug.Print "Nombre: " & ref.Name
        Debug.Print "Descripción: " & ref.Description
        Debug.Print "GUID: " & ref.guid
        Debug.Print "Major: " & ref.major & " | Minor: " & ref.minor
        Debug.Print "Ruta: " & ref.fullPath
        Debug.Print "--------------------------"
    Next ref
End Sub

Sub AgregarReferenciaPorRuta(ruta As String, nombre As String)
    On Error Resume Next
    
    ThisWorkbook.VBProject.References.AddFromFile ruta
    
    If Err.Number = 0 Then
        LogInfo MODULE_NAME, "[AgregarReferenciaPorRuta] Referencia añadida con éxito: " & nombre
    ElseIf Err.Number = 32813 Then
        LogWarning MODULE_NAME, "[AgregarReferenciaPorRuta] La referencia ya estaba activada: " & nombre
    Else
        LogCurrentError MODULE_NAME, "[AgregarReferenciaPorRuta] " & nombre
    End If
    On Error GoTo 0
End Sub

Sub AgregarReferenciaPorGUID(guid As String, major As Long, minor As Long, nombre As String)
    
    On Error Resume Next ' Evita error si la referencia ya existe
    ThisWorkbook.VBProject.References.AddFromGuid guid, major:=1, minor:=0
    
    If Err.Number <> 0 And Err.Number <> 32813 Then
        ' Si falla (ej. error 429 o similar), intentamos cargarla sin forzar versión
        ' Nota: VBA a veces permite AddFromGuid con 0, 0 para "la última"
        ThisWorkbook.VBProject.References.AddFromGuid guid, 0, 0
    End If
    
    If Err.Number = 0 Then
        LogInfo MODULE_NAME, "[AgregarReferenciaPorGUID] Referencia añadida con éxito: " & nombre
    ElseIf Err.Number = 32813 Then
        LogWarning MODULE_NAME, "[AgregarReferenciaPorGUID] La referencia ya estaba activada: " & nombre
    Else
        LogCurrentError MODULE_NAME, "[AgregarReferenciaPorGUID] " & nombre
    End If
    On Error GoTo 0
End Sub


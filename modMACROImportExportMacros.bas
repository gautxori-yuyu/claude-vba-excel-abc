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

'@Folder "0-Developer"
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
    rutaExportacion = Wb.Path
    
    ' Crear carpeta si no existe
    If Dir(rutaExportacion, vbDirectory) = "" Then
        MkDir rutaExportacion
    End If
    
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In Wb.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1: nombreArchivo = vbComp.Name & ".bas" ' Módulo estándar
        Case 2, 100: nombreArchivo = vbComp.Name & ".cls" ' Clase
        Case 3: nombreArchivo = vbComp.Name & ".frm" ' Formulario
        Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        
        ' Exportar el componente
        If vbComp.CodeModule.CountOfLines = 0 And InStr(vbComp.Name, "Hoja") > 0 Then
        Else
            vbComp.Export rutaExportacion & "\" & nombreArchivo
        End If
    Next vbComp
    
    MsgBox "Componentes exportados a: " & rutaExportacion, vbInformation
End Sub

'@Description: Exporta componentes VBA sin mostrar mensajes al usuario
'@Scope: Privado
'@ArgumentDescriptions: wb: Workbook de origen | rutaDestino: Carpeta donde exportar
Sub ExportarComponentesVBASilencioso(Wb As Workbook, rutaDestino As String)
Attribute ExportarComponentesVBASilencioso.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim vbComp As Object
    Dim nombreArchivo As String
    Dim fso As Object
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Asegurar que existe la carpeta
    If Not fso.FolderExists(rutaDestino) Then
        fso.CreateFolder rutaDestino
    End If
    
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In Wb.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: nombreArchivo = vbComp.Name & ".bas"  ' Módulo estándar
            Case 2, 100: nombreArchivo = vbComp.Name & ".cls"  ' Clase o documento
            Case 3: nombreArchivo = vbComp.Name & ".frm"  ' Formulario
            Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        
        ' Exportar solo si tiene código
        If vbComp.CodeModule.CountOfLines > 0 Then
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
    rutaImportacion = Wb.Path
    
    If Dir(rutaImportacion, vbDirectory) = "" Then
        MsgBox "La carpeta de importación no existe: " & rutaImportacion, vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaImportacion)
    
    For Each archivo In carpeta.Files
        extension = LCase(fso.GetExtensionName(archivo.Name))
        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            Wb.VBProject.VBComponents.Import archivo.Path
        End If
    Next archivo
    
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
    Do While fso.GetFolder(rutaTempDescompresion).Files.Count = 0 And intentos < 50
        DoEvents
        Sleep 200
        intentos = intentos + 1
    Loop
    
    ' Confirmar restauración
    If MsgBox("¿Desea restaurar el código VBA desde este backup?" & vbCrLf & vbCrLf & _
              "ADVERTENCIA: Se eliminarán todos los módulos actuales" & vbCrLf & _
              "y se cargarán los del backup.", vbExclamation + vbYesNo, "Confirmar restauración") = vbYes Then
        
        ' Importar componentes
        Dim archivo As Object
        Dim extension As String
        
        For Each archivo In fso.GetFolder(rutaTempDescompresion).Files
            extension = LCase(fso.GetExtensionName(archivo.Name))
            If extension = "bas" Or extension = "cls" Or extension = "frm" Then
                ' Intentar eliminar componente existente
                On Error Resume Next
                Dim nombreComp As String
                nombreComp = fso.GetBaseName(archivo.Name)
                ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(nombreComp)
                On Error GoTo ErrorHandler
                
                ' Importar
                ThisWorkbook.VBProject.VBComponents.Import archivo.Path
            End If
        Next archivo
        
        MsgBox "Restauración completada desde: " & rutaZip, vbInformation, "Restauración completada"
    End If
    
    ' Limpiar carpeta temporal
    On Error Resume Next
    fso.DeleteFolder rutaTempDescompresion, True
    On Error GoTo 0
    
    Exit Sub
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[RestaurarBackupVBADesdeZip]"
    MsgBox "Error al restaurar backup: " & Err.Description, vbCritical, "Error"
End Sub

Sub ImportarComponentesVBAaThisWorkbookXLAM()
Attribute ImportarComponentesVBAaThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta desde donde se importarán los archivos
    Dim fso As Object, archivo As Object, carpeta As Object, rutaImportacion As String, extension As String
    rutaImportacion = ThisWorkbook.Path
    
    If Dir(rutaImportacion, vbDirectory) = "" Then
        MsgBox "La carpeta de importación no existe: " & rutaImportacion, vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaImportacion)
    
    For Each archivo In carpeta.Files
        extension = LCase(fso.GetExtensionName(archivo.Name))
        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            ThisWorkbook.VBProject.VBComponents.Import archivo.Path
        End If
    Next archivo
    
    MsgBox "Importación completada desde: " & rutaImportacion, vbInformation
End Sub

Sub ExportarComponentesVBAdesdeThisWorkbookXLAM()
Attribute ExportarComponentesVBAdesdeThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta donde se guardarán los archivos exportados
    Dim rutaExportacion As String, nombreArchivo As String, vbComp
    rutaExportacion = ThisWorkbook.Path
    
    ' Crear carpeta si no existe
    If Dir(rutaExportacion, vbDirectory) = "" Then
        MkDir rutaExportacion
    End If
    
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1: nombreArchivo = vbComp.Name & ".bas" ' Módulo estándar
        Case 2, 100: nombreArchivo = vbComp.Name & ".cls" ' Clase
        Case 3: nombreArchivo = vbComp.Name & ".frm" ' Formulario
        Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        If vbComp.CodeModule.CountOfLines > 0 Then vbComp.Export rutaExportacion & "\" & nombreArchivo
    Next vbComp
    
    MsgBox "Exportación completada en: " & rutaExportacion, vbInformation
End Sub

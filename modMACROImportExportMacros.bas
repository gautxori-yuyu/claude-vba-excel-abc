Attribute VB_Name = "modMACROImportExportMacros"
' ==============================================================================================================
' MÃDULO: modMACROImportExportMacros
' DESCRIPCIÃN: MÃ³dulo para exportar e importar componentes VBA (mÃ³dulos, clases, formularios) desde y hacia
'              archivos fÃ­sicos. Permite hacer backup del cÃ³digo VBA o transferirlo entre proyectos.
' REQUISITOS: - Referencia a "Microsoft Visual Basic for Applications Extensibility 5.3"
'             - Acceso al modelo de objetos VBA habilitado en las opciones de seguridad de Excel
' ==============================================================================================================

'@NOTE: Debes tener habilitado el acceso al modelo de objetos de VBA:
' - En el editor de VBA: ve a Herramientas > Referencias. Marca "Microsoft Visual Basic for Applications Extensibility 5.3".
' - en Excel: Archivo > Opciones > Centro de confianza > ConfiguraciÃ³n del Centro de confianza
'       > ConfiguraciÃ³n de macros > marca "Confiar en el acceso al modelo de objetos del proyecto VBA".

'@Folder "0-Developer"
Option Explicit

Private Const MODULE_NAME As String = "modMACROImportExportMacros"

' -------------------------------------------------------------------------------------------------------------
' EXPORTACIÃN DE COMPONENTES VBA
' -------------------------------------------------------------------------------------------------------------

'@Description: Exporta todos los componentes VBA del libro seleccionado (mÃ³dulos, clases, formularios)
'              a archivos individuales en la carpeta del libro
'@Scope: (muestra formulario de selecciÃ³n)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmImportExportMacros (formulario de selecciÃ³n de libro)
'@Note: Los archivos se guardan con extensiones: .bas (mÃ³dulos), .cls (clases), .frm (formularios)
Sub ExportarComponentesVBA()
Attribute ExportarComponentesVBA.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim vbComp As Object
    Dim rutaExportacion As String
    Dim nombreArchivo As String
    Dim wb As Workbook
    
    Dim frm As New frmImportExportMacros
    frm.Show vbModal
    If frm.WorkbookSeleccionado Is Nothing Then Exit Sub
    Set wb = frm.WorkbookSeleccionado
    Unload frm
    If wb Is Nothing Then Exit Sub               ' Cancelado o error
    
    ' Carpeta donde se guardarÃ¡n los archivos exportados
    rutaExportacion = wb.Path
    
    ' Crear carpeta si no existe
    If Dir(rutaExportacion, vbDirectory) = "" Then
        MkDir rutaExportacion
    End If
    
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In wb.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1: nombreArchivo = vbComp.Name & ".bas" ' MÃ³dulo estÃ¡ndar
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
Sub ExportarComponentesVBASilencioso(wb As Workbook, rutaDestino As String)
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
    For Each vbComp In wb.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: nombreArchivo = vbComp.Name & ".bas"  ' MÃ³dulo estÃ¡ndar
            Case 2, 100: nombreArchivo = vbComp.Name & ".cls"  ' Clase o documento
            Case 3: nombreArchivo = vbComp.Name & ".frm"  ' Formulario
            Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        
        ' Exportar solo si tiene cÃ³digo
        If vbComp.CodeModule.CountOfLines > 0 Then
            vbComp.Export rutaDestino & "\" & nombreArchivo
        End If
    Next vbComp
    
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------------------------------------------
' IMPORTACIÃN DE COMPONENTES VBA
' -------------------------------------------------------------------------------------------------------------

'@Description: Importa componentes VBA desde archivos fÃ­sicos al libro seleccionado. Permite seleccionar
'              mÃºltiples archivos (.bas, .cls, .frm)
'@Scope:  (muestra formularios de selecciÃ³n)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmImportExportMacros (formulario de selecciÃ³n de libro)
'@Note: Elimina el componente existente si ya existe uno con el mismo nombre antes de importar
Sub ImportarComponentesVBA()
Attribute ImportarComponentesVBA.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim fso As Object, carpeta As Object, archivo As Object
    Dim rutaImportacion As String
    Dim extension As String
    Dim wb As Workbook
    
    Dim frm As New frmImportExportMacros
    frm.Show vbModal
    If frm.WorkbookSeleccionado Is Nothing Then Exit Sub
    Set wb = frm.WorkbookSeleccionado
    Unload frm
    If wb Is Nothing Then Exit Sub               ' Cancelado o error
    
    ' Carpeta desde donde se importarÃ¡n los archivos
    rutaImportacion = wb.Path
    
    If Dir(rutaImportacion, vbDirectory) = "" Then
        MsgBox "La carpeta de importaciÃ³n no existe: " & rutaImportacion, vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaImportacion)
    
    For Each archivo In carpeta.Files
        extension = LCase(fso.GetExtensionName(archivo.Name))
        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            wb.VBProject.VBComponents.Import archivo.Path
        End If
    Next archivo
    
    MsgBox "ImportaciÃ³n completada desde: " & rutaImportacion, vbInformation
End Sub

' ==========================================
' FUNCIÃN 5: UTILIDAD PARA RESTAURAR DESDE BACKUP
' ==========================================

'@Description: Restaura cÃ³digo VBA desde un archivo ZIP de backup
'@Scope: PÃºblico
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
    
    ' Esperar a que termine la descompresiÃ³n
    Dim intentos As Integer
    intentos = 0
    Do While fso.GetFolder(rutaTempDescompresion).Files.Count = 0 And intentos < 50
        DoEvents
        Sleep 200
        intentos = intentos + 1
    Loop
    
    ' Confirmar restauraciÃ³n
    If MsgBox("Â¿Desea restaurar el cÃ³digo VBA desde este backup?" & vbCrLf & vbCrLf & _
              "ADVERTENCIA: Se eliminarÃ¡n todos los mÃ³dulos actuales" & vbCrLf & _
              "y se cargarÃ¡n los del backup.", vbExclamation + vbYesNo, "Confirmar restauraciÃ³n") = vbYes Then
        
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
        
        MsgBox "RestauraciÃ³n completada desde: " & rutaZip, vbInformation, "RestauraciÃ³n completada"
    End If
    
    ' Limpiar carpeta temporal
    On Error Resume Next
    fso.DeleteFolder rutaTempDescompresion, True
    On Error GoTo 0
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[RestaurarBackupVBADesdeZip] - Error: " & Err.Description
    MsgBox "Error al restaurar backup: " & Err.Description, vbCritical, "Error"
End Sub

Sub ImportarComponentesVBAaThisWorkbookXLAM()
Attribute ImportarComponentesVBAaThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta desde donde se importarÃ¡n los archivos
    Dim fso As Object, archivo As Object, carpeta As Object, rutaImportacion As String, extension As String
    rutaImportacion = ThisWorkbook.Path
    
    If Dir(rutaImportacion, vbDirectory) = "" Then
        MsgBox "La carpeta de importaciÃ³n no existe: " & rutaImportacion, vbExclamation
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
    
    MsgBox "ImportaciÃ³n completada desde: " & rutaImportacion, vbInformation
End Sub

Sub ExportarComponentesVBAdesdeThisWorkbookXLAM()
Attribute ExportarComponentesVBAdesdeThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta donde se guardarÃ¡n los archivos exportados
    Dim rutaExportacion As String, nombreArchivo As String, vbComp
    rutaExportacion = ThisWorkbook.Path
    
    ' Crear carpeta si no existe
    If Dir(rutaExportacion, vbDirectory) = "" Then
        MkDir rutaExportacion
    End If
    
    ' Recorrer todos los componentes del proyecto VBA
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
        Case 1: nombreArchivo = vbComp.Name & ".bas" ' MÃ³dulo estÃ¡ndar
        Case 2, 100: nombreArchivo = vbComp.Name & ".cls" ' Clase
        Case 3: nombreArchivo = vbComp.Name & ".frm" ' Formulario
        Case Else: nombreArchivo = vbComp.Name & ".txt"
        End Select
        If vbComp.CodeModule.CountOfLines > 0 Then vbComp.Export rutaExportacion & "\" & nombreArchivo
    Next vbComp
    
    MsgBox "ExportaciÃ³n completada en: " & rutaExportacion, vbInformation
End Sub

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
    rutaExportacion = Wb.Path
    
    ExportarFichsVBAaCarpeta Wb, rutaExportacion
    
    MsgBox "Componentes exportados a: " & rutaExportacion, vbInformation
End Sub

Sub ExportarComponentesVBAdesdeThisWorkbookXLAM()
Attribute ExportarComponentesVBAdesdeThisWorkbookXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Carpeta donde se guardarán los archivos exportados
    Dim rutaExportacion As String
    rutaExportacion = ThisWorkbook.Path
    
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
    rutaImportacion = Wb.Path
    
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
    rutaImportacion = ThisWorkbook.Path
    
    ImportarFichsVBAenCarpeta rutaImportacion
    
    If MsgBox("¿Importar Hojas de calculo desde XLAM ABC?", vbYesNo + vbQuestion, "Importar dependencias") = vbYes Then
        ImportarDependencias
    End If
    
    MsgBox "Importación completada desde: " & rutaImportacion & ". Falta, en el 'proyecto ABC', inyectar los ficheros a descomprimir (FSWatcher). Requiere hacerlo con Excel cerrado, usar VBScript.", vbInformation
End Sub

Private Sub ImportarDependencias()
    Dim rutaLibro As Variant
    Dim wbOrigen As Workbook
    Dim ws As Worksheet
    Dim nombreHojaBorrar As String: nombreHojaBorrar = "Hoja1" ' Nombre estándar
    
    rutaLibro = Application.GetOpenFilename("Excel Files (*.xlam; *.xlsm), *.xlam; *.xlsm", , "Seleccionar versión previa")
    
    If rutaLibro <> False Then
        Set wbOrigen = Workbooks.Open(fileName:=rutaLibro, ReadOnly:=True)
        
        For Each ws In wbOrigen.Worksheets
            If InStr(ws.Name, "Hoja") = 0 Then
                ' Copiar cada hoja RENOMBRADA al final del libro actual
                ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            End If
        Next ws
        
        wbOrigen.Close SaveChanges:=False
        
        ' 2. Intentar eliminar la hoja inicial si existe
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHojaBorrar)
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub ImportarFichsVBAenCarpeta(rutaImportacion As String, Optional bRemove As Boolean = False)
    Dim fso As Object, archivo As Object, carpeta As Object, extension As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(rutaImportacion)
    
    If Dir(rutaImportacion, vbDirectory) = "" Then
        MsgBox "La carpeta de importación no existe: " & rutaImportacion, vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Dim vbComp As Object
    
    For Each archivo In carpeta.files
        extension = LCase(fso.GetExtensionName(archivo.Name))
        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            If archivo.Name = "ThisWorkbook.cls" Then
                Set vbComp = ThisWorkbook.VBProject.VBComponents("ThisWorkbook")
                If vbComp.CodeModule.CountOfLines <= 1 Then
                    ' Insertar el contenido del archivo
                    vbComp.CodeModule.AddFromFile archivo.Path
                Else
                    Select Case True
                        Case bRemove, MsgBox("¿Eliminar el componente ThisWorkbook?", vbYesNo + vbDefaultButton2, "Clase existente") = vbYes
                            ' Limpiar código existente antes de inyectar el nuevo
                            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
                            
                            ' Insertar el contenido del archivo
                            vbComp.CodeModule.AddFromFile archivo.Path
                    End Select
                End If
            Else
                ' Intentar eliminar componente existente
                On Error Resume Next
                Dim nombreComp As String
                nombreComp = fso.GetBaseName(archivo.Name)
                
                On Error Resume Next
                ' Intentamos asignar el componente a una variable
                Set vbComp = ThisWorkbook.VBProject.VBComponents(nombreComp)
                On Error GoTo 0
                
                ' Si la variable es (Nothing), el componente no existe
                If Not vbComp Is Nothing Then
                    ' Importar
                    ThisWorkbook.VBProject.VBComponents.Import archivo.Path
                Else
                    Select Case True
                        Case bRemove, MsgBox("¿Eliminar el componente " & nombreComp & "?", vbYesNo + vbDefaultButton2, "Clase existente") = vbYes
                            ThisWorkbook.VBProject.VBComponents.Remove vbComp
                                        
                            ' Importar
                            ThisWorkbook.VBProject.VBComponents.Import archivo.Path
                    End Select
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next archivo
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "[ImportarFichsVBAenCarpeta]"
    MsgBox "Error al importar fichero: " & Err.Description, vbCritical, "Error"
End Sub

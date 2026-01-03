Attribute VB_Name = "UDFs_Backups"
' ==========================================
' FUNCIONES DE BACKUP PARA SINCRONIZACIÓN
' ==========================================
' PROPÓSITO:
' Crear copias de seguridad antes de modificar:
' 1. Código VBA (archivos .bas, .cls, .frm) -> ZIP
' 2. Hoja Excel "PROCEDIMIENTOS" -> Hoja duplicada con sufijo _bkp
'
' ==========================================

'@Folder "MACROS"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ==========================================
' FUNCIÓN 1: BACKUP DE CÓDIGO VBA A ZIP
' ==========================================

'@Description: Exporta todos los componentes VBA de ThisWorkbook a una carpeta temporal y la comprime en ZIP
'@Scope: Privado
'@ArgumentDescriptions: (sin argumentos)
'@Returns: String | Ruta completa del archivo ZIP creado, o "" si falla
Function CrearBackupCodigoVBA() As String
Attribute CrearBackupCodigoVBA.VB_Description = "[modMACROBackups] FUNCIÓN 1: BACKUP DE CÓDIGO VBA A ZIP Exporta todos los componentes VBA de ThisWorkbook a una carpeta temporal y la comprime en ZIP. Aplica a: ThisWorkbook\r\nM.D.:Privado"
Attribute CrearBackupCodigoVBA.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim rutaBackup As String
    Dim rutaTempExport As String
    Dim nombreZip As String
    Dim timestampStr As String
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Generar timestamp para nombre único
    timestampStr = Format(Now, "yyyymmdd_hhnnss")
    
    ' Rutas
    rutaBackup = ThisWorkbook.Path & "\Backups"
    nombreZip = "VBA_Backup_" & timestampStr & ".zip"
    rutaTempExport = Environ("TEMP") & "\VBA_Export_" & timestampStr
    
    ' Crear carpeta de backups si no existe
    If Not fso.FolderExists(rutaBackup) Then
        fso.CreateFolder rutaBackup
    End If
    
    ' Crear carpeta temporal para exportar
    If Not fso.FolderExists(rutaTempExport) Then
        fso.CreateFolder rutaTempExport
    End If
    
    ' Exportar todos los componentes VBA
    Call ExportarComponentesVBASilencioso(ThisWorkbook, rutaTempExport)
    
    ' Comprimir en ZIP
    Dim rutaZipCompleta As String
    rutaZipCompleta = rutaBackup & "\" & nombreZip
    
    If ComprimirCarpetaAZip(rutaTempExport, rutaZipCompleta) Then
        CrearBackupCodigoVBA = rutaZipCompleta
        
        ' Limpiar carpeta temporal
        On Error Resume Next
        fso.DeleteFolder rutaTempExport, True
        On Error GoTo ErrorHandler
    Else
        CrearBackupCodigoVBA = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "[CrearBackupCodigoVBA] - Error: " & Err.Description
    CrearBackupCodigoVBA = ""
End Function

' ==========================================
' FUNCIÓN 2: BACKUP DE HOJA EXCEL
' ==========================================

'@Description: Crea una copia de seguridad de una hoja Excel añadiendo sufijo _bkp (VERSIÓN PARA XLAM)
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet a duplicar
'@Returns: Boolean | True si se creó correctamente
Function CrearBackupHojaExcel(ws As Worksheet) As Boolean
Attribute CrearBackupHojaExcel.VB_Description = "[modMACROBackups] FUNCIÓN 2: BACKUP DE HOJA EXCEL Crea una copia de seguridad de una hoja Excel añadiendo sufijo _bkp (VERSIÓN PARA XLAM). Aplica a: Cells Range\r\nM.D.:Privado"
Attribute CrearBackupHojaExcel.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim nombreBackup As String
    Dim wsBackup As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim esAddin As Boolean
    Dim errorOcurrido As Boolean
    
    On Error GoTo ErrorHandler
    
    nombreBackup = ws.Name & "_bkp"
    errorOcurrido = False
    
    ' Verificar si ya existe una hoja de backup
    On Error Resume Next
    Set wsBackup = ws.Parent.Worksheets(nombreBackup)
    On Error GoTo ErrorHandler
    
    If Not wsBackup Is Nothing Then
        ' Ya existe, preguntar si eliminarla
        respuesta = MsgBox("Ya existe una copia de seguridad anterior: '" & nombreBackup & "'" & vbCrLf & vbCrLf & _
                          "¿Desea reemplazarla con una nueva copia?" & vbCrLf & vbCrLf & _
                          "Sí = Reemplazar (se eliminará la anterior)" & vbCrLf & _
                          "No = Cancelar operación", _
                          vbQuestion + vbYesNo, "Backup existente")
        
        If respuesta = vbNo Then
            CrearBackupHojaExcel = False
            Exit Function
        Else
            ' Eliminar backup antiguo
            Application.DisplayAlerts = False
            wsBackup.Delete
            Application.DisplayAlerts = True
            Set wsBackup = Nothing
        End If
    End If
    
    ' ========================================
    ' SOLUCIÓN PARA XLAM: Desactivar IsAddin temporalmente
    ' ========================================
    
    ' Guardar estado actual de IsAddin
    esAddin = ws.Parent.IsAddin
    
    ' Si es un Add-in, desactivarlo temporalmente para permitir copiar hojas
    If esAddin Then
        ws.Parent.IsAddin = False
        Debug.Print "[CrearBackupHojaExcel] - XLAM detectado, IsAddin desactivado temporalmente"
    End If
    
    ' Crear nueva copia de la hoja
    On Error GoTo ErrorHandlerRestaurar
    ws.Copy After:=ws
    
    ' Obtener referencia a la hoja recién creada
    Set wsBackup = ws.Parent.Worksheets(ws.Index + 1)
    wsBackup.Name = nombreBackup
    
    ' Añadir marca visual de que es backup
    With wsBackup.Range("A1")
        .Interior.Color = RGB(255, 200, 200)  ' Fondo rojo claro
        .Font.Bold = True
        
        If False Then
            ' Añadir comentario con fecha
            On Error Resume Next
            .ClearComments
            .AddComment "BACKUP creado el " & Format(Now, "dd/mm/yyyy hh:nn:ss")
            On Error GoTo ErrorHandlerRestaurar
        End If
    End With
    
    ' ========================================
    ' RESTAURAR IsAddin si era un Add-in
    ' ========================================
    If esAddin Then
        ws.Parent.IsAddin = True
        Debug.Print "[CrearBackupHojaExcel] - IsAddin restaurado"
    End If
    
    CrearBackupHojaExcel = True
    Exit Function
    
ErrorHandlerRestaurar:
    ' Error durante la copia, pero debemos restaurar IsAddin
    errorOcurrido = True
    Debug.Print "[CrearBackupHojaExcel] - Error al copiar: " & Err.Description
    
    ' Restaurar IsAddin antes de salir
    If esAddin Then
        On Error Resume Next
        ws.Parent.IsAddin = True
        Debug.Print "[CrearBackupHojaExcel] - IsAddin restaurado tras error"
        On Error GoTo 0
    End If
    
    CrearBackupHojaExcel = False
    Exit Function
    
ErrorHandler:
    Debug.Print "[CrearBackupHojaExcel] - Error: " & Err.Description
    CrearBackupHojaExcel = False
End Function

' ==========================================
' VERSIÓN ALTERNATIVA: USO MANUAL DEL ESTADO
' ==========================================

'@Description: Versión alternativa de CrearBackupHojaExcel con control manual de IsAddin
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet a duplicar
'@Returns: Boolean | True si se creó correctamente
Private Function CrearBackupHojaExcel_V2(ws As Worksheet) As Boolean
    Dim nombreBackup As String
    Dim wsBackup As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim bWasAddin As Boolean
    
    On Error GoTo ErrorHandler
    
    nombreBackup = ws.Name & "_bkp"
    
    ' Verificar si ya existe una hoja de backup
    On Error Resume Next
    Set wsBackup = ws.Parent.Worksheets(nombreBackup)
    On Error GoTo ErrorHandler
    
    If Not wsBackup Is Nothing Then
        respuesta = MsgBox("Ya existe una copia de seguridad anterior: '" & nombreBackup & "'" & vbCrLf & vbCrLf & _
                          "¿Desea reemplazarla con una nueva copia?" & vbCrLf & vbCrLf & _
                          "Sí = Reemplazar | No = Cancelar", _
                          vbQuestion + vbYesNo, "Backup existente")
        
        If respuesta = vbNo Then
            CrearBackupHojaExcel_V2 = False
            Exit Function
        Else
            Application.DisplayAlerts = False
            wsBackup.Delete
            Application.DisplayAlerts = True
            Set wsBackup = Nothing
        End If
    End If
    
    ' Desactivar modo Add-in temporalmente
    bWasAddin = ThisWorkbook.IsAddin
    DesactivarModoAddin
    
    ' Intentar copiar la hoja
    On Error GoTo ErrorConRestauracion
    
    ws.Copy After:=ws
    Set wsBackup = ws.Parent.Worksheets(ws.Index + 1)
    wsBackup.Name = nombreBackup
    
    ' Marca visual
    With wsBackup.Range("A1")
        .Interior.Color = RGB(255, 200, 200)
        .Font.Bold = True
        On Error Resume Next
        .ClearComments
        .AddComment "BACKUP creado el " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        On Error GoTo ErrorConRestauracion
    End With
    
    ' Restaurar estado Add-in
    If bWasAddin Then Call RestaurarModoAddin
    
    CrearBackupHojaExcel_V2 = True
    Exit Function
    
ErrorConRestauracion:
    ' Error, pero restaurar IsAddin antes de salir
    Call RestaurarModoAddin
    Debug.Print "[CrearBackupHojaExcel_V2] - Error: " & Err.Description
    CrearBackupHojaExcel_V2 = False
    Exit Function
    
ErrorHandler:
    Debug.Print "[CrearBackupHojaExcel_V2] - Error: " & Err.Description
    CrearBackupHojaExcel_V2 = False
End Function

' ==========================================
' INSTRUCCIONES DE USO
' ==========================================
'
' OPCIÓN 1 (RECOMENDADA):
' Reemplazar toda la función CrearBackupHojaExcel existente
' por la primera versión de este archivo
'
' OPCIÓN 2:
' Si prefieres tener control separado, usa:
' - DesactivarModoAddin()
' - RestaurarModoAddin()
' - CrearBackupHojaExcel_V2()
'
' ==========================================
'
' EXPLICACIÓN DEL PROBLEMA:
'
' Cuando ThisWorkbook.IsAddin = True:
' ? No se pueden copiar hojas (.Copy)
' ? No se pueden agregar hojas (.Add)
' ? No se pueden mover hojas (.Move)
' ? El libro no aparece en la UI de Excel
'
' Cuando ThisWorkbook.IsAddin = False:
' ? Se pueden copiar hojas
' ? Se pueden agregar hojas
' ? Se pueden mover hojas
' ? El libro aparece en la UI de Excel
'
' ==========================================
'
' FLUJO DE LA SOLUCIÓN:
'
' 1. Detectar si ThisWorkbook.IsAddin = True
' 2. Si es True ? Cambiar temporalmente a False
' 3. Copiar la hoja con .Copy After:=ws
' 4. Renombrar la copia
' 5. Añadir marca visual
' 6. SIEMPRE restaurar IsAddin al estado original
'    (incluso si hay error)
'
' ==========================================
'
' NOTAS IMPORTANTES:
'
' • El cambio de IsAddin es TEMPORAL (microsegundos)
' • El usuario NO verá el libro aparecer/desaparecer
' • Siempre se restaura el estado original
' • Funciona con GoTo ErrorHandlerRestaurar para
'   garantizar restauración incluso con errores
'
' ==========================================

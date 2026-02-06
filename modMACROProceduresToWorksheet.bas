Attribute VB_Name = "modMACROProceduresToWorksheet"
' ==========================================
' EXTENSIÓN PARA modUTILSProcedureParsing.bas
' SINCRONIZACIÓN BIDIRECCIONAL CON HOJA EXCEL "PROCEDIMIENTOS"
' ==========================================
' Reutiliza completamente la función ParsearProcsDelProyecto() existente
' y todas las propiedades de clsVBAProcedure
' ==========================================

'@Folder "1-Inicio e Instalacion.Gestion de modulos y procs"
Option Explicit

Private Const MODULE_NAME As String = "modMACROProceduresToWorksheet"

Private Const SHEET_NAME As String = "PROCEDIMIENTOS"

'@Description: Sincroniza procedimientos del proyecto con hoja Excel "PROCEDIMIENTOS". Crea la hoja si no existe, o sincroniza cambios bidireccionales si existe.
'@Scope: Público
'@Category: Sincronización
Public Sub WriteProcedimientosSheet()
Attribute WriteProcedimientosSheet.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim funciones As Object
    Dim bSheetExisted As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Parsear todos los procedimientos del proyecto
    Set funciones = ParsearProcsDelProyecto()
    If funciones Is Nothing Or funciones.Count = 0 Then
        MsgBox "No se encontraron procedimientos para procesar.", vbInformation, "Sin procedimientos"
        Exit Sub
    End If
    
    ' Verificar si existe la hoja
    bSheetExisted = SheetExists(ThisWorkbook, SHEET_NAME)
    
    If bSheetExisted Then
        ' Leer hoja existente y comparar
        Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
        Call SincronizarConHoja(ws, funciones)
    Else
        ' Crear hoja nueva y volcar datos
        Set ws = CrearHojaProcedimientos(ThisWorkbook, SHEET_NAME)
        Call VolcarProcedimientosAHoja(ws, funciones)
        MsgBox "Hoja '" & SHEET_NAME & "' creada con " & funciones.Count & " procedimientos.", vbInformation, "Hoja creada"
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "[WriteProcedimientosSheet] - Error: " & Err.Description
    MsgBox "Error al procesar hoja de procedimientos: " & Err.Description, vbCritical, "Error"
End Sub

' ==========================================
' FUNCIÓN 4: MODIFICACIÓN DE WriteProcedimientosSheet
' ==========================================

'@Description: VERSIÓN MODIFICADA que usa SincronizarConHoja_ConBackup en lugar de SincronizarConHoja
'@Scope: Público
'@Category: Sincronización
Public Sub WriteProcedimientosSheet_ConBackup()
Attribute WriteProcedimientosSheet_ConBackup.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim funciones As Object
    Dim bSheetExisted As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Parsear todos los procedimientos del proyecto
    Set funciones = ParsearProcsDelProyecto()
    If funciones Is Nothing Or funciones.Count = 0 Then
        MsgBox "No se encontraron procedimientos para procesar.", vbInformation, "Sin procedimientos"
        Exit Sub
    End If
    
    ' Verificar si existe la hoja
    bSheetExisted = SheetExists(ThisWorkbook, SHEET_NAME)
    
    If bSheetExisted Then
        ' Leer hoja existente y comparar (CON BACKUP)
        Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
        Call SincronizarConHoja_ConBackup(ws, funciones)  ' ? CAMBIO AQUÍ
    Else
        ' Crear hoja nueva y volcar datos (sin backup necesario)
        Set ws = CrearHojaProcedimientos(ThisWorkbook, SHEET_NAME)
        Call VolcarProcedimientosAHoja(ws, funciones)
        MsgBox "Hoja '" & SHEET_NAME & "' creada con " & funciones.Count & " procedimientos.", vbInformation, "Hoja creada"
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "[WriteProcedimientosSheet_ConBackup] - Error: " & Err.Description
    MsgBox "Error al procesar hoja de procedimientos: " & Err.Description, vbCritical, "Error"
End Sub

'@Description: Crea una nueva hoja con encabezados para procedimientos
'@Scope: Privado
'@ArgumentDescriptions: wb: Workbook donde crear | sheetName: Nombre de la hoja
'@Returns: Worksheet | Referencia a la hoja creada
Private Function CrearHojaProcedimientos(wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = sheetName
    
    Call CrearEncabezadosHoja(ws)
    
    Set CrearHojaProcedimientos = ws
End Function

'@Description: Crea los encabezados de columnas en la hoja de procedimientos
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet donde crear encabezados
Private Sub CrearEncabezadosHoja(ws As Worksheet)
    With ws
        .Cells(1, 1).value = "Módulo"
        .Cells(1, 2).value = "Firma del procedimiento"
        .Cells(1, 3).value = "Description"
        .Cells(1, 4).value = "Category"
        .Cells(1, 5).value = "Scope"
        .Cells(1, 6).value = "ArgumentDescriptions"
        .Cells(1, 7).value = "Returns"
        
        ' Formato de encabezados
        With .Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Ajustar anchos
        .Columns("A:A").ColumnWidth = 20  ' Módulo
        .Columns("B:B").ColumnWidth = 50  ' Firma
        .Columns("C:C").ColumnWidth = 70  ' Description
        .Columns("D:D").ColumnWidth = 25  ' Category
        .Columns("E:E").ColumnWidth = 25  ' Scope
        .Columns("F:F").ColumnWidth = 40  ' ArgumentDescriptions
        .Columns("G:G").ColumnWidth = 30  ' Returns
        
        .Columns("A:G").WrapText = True
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

'@Description: Vuelca todos los procedimientos parseados a la hoja Excel
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet destino | funciones: Dictionary con objetos clsVBAProcedure
Private Sub VolcarProcedimientosAHoja(ws As Worksheet, funciones As Object)
    Dim i As Long, col As Long, fila As Long
    Dim proc As clsVBAProcedure
    
    fila = 2 ' Fila inicial (después de encabezados)
    
    For i = 0 To funciones.Count - 1
        Set proc = funciones(i)
        
        With ws
            .Cells(fila, 1).value = proc.Module
            .Cells(fila, 2).value = proc.NormalizedSignature
            .Cells(fila, 3).value = proc.Description
            .Cells(fila, 4).value = IIf(proc.Category = DEFAULT_CATEGORY, "", proc.Category)
            .Cells(fila, 5).value = proc.Scope
            .Cells(fila, 6).value = IIf(proc.ArgumentDescriptions = DEFAULT_NOPARAMS, "", proc.ArgumentDescriptions)
            .Cells(fila, 7).value = proc.Returns
            
            ' Formato condicional para mejor lectura
            If i Mod 2 = 0 Then
                .Range(.Cells(fila, 1), .Cells(fila, 7)).Interior.Color = RGB(242, 242, 242)
                .Range(.Cells(fila, 1), .Cells(fila, 7)).HorizontalAlignment = xlLeft
                .Range(.Cells(fila, 1), .Cells(fila, 7)).VerticalAlignment = xlTop
            End If
            
            ' enfatiza celdas que tengan diferencias en valores deducidos vs metadatos procesados del proyecto
            For col = 3 To 7
                If InStr(.Cells(fila, col).value, "M.D.:") > 0 Then
                    If i Mod 2 = 0 Then
                        .Cells(fila, col).Interior.Color = RGB(240, 210, 120)
                    Else
                        .Cells(fila, col).Interior.Color = RGB(210, 240, 120)
                    End If
                End If
            Next
        End With
        
        fila = fila + 1
    Next i
    
    ' Aplicar bordes
    With ws.Range(ws.Cells(1, 1), ws.Cells(fila - 1, 7))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

'@Description: Gestiona la sincronización bidireccional entre hoja Excel y código VBA
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet con datos existentes | funciones: Dictionary con procedimientos actuales del código
Private Sub SincronizarConHoja(ws As Worksheet, funciones As Object)
    Dim dictHoja As Object ' Dictionary con clave compuesta: Módulo + "|" + Firma -> array de metadatos
    Dim dictCodigo As Object ' Dictionary con clave compuesta: Módulo + "|" + Firma -> objeto clsVBAProcedure
    Dim hayDiferencias As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim mensaje As String
    
    ' Crear diccionarios para comparación
    Set dictHoja = LeerMetadatosDeHoja(ws)
    Set dictCodigo = CrearDiccionarioProcedimientos(funciones)
    
    ' Comparar y detectar diferencias
    hayDiferencias = HayDiferenciasEnMetadatos(dictHoja, dictCodigo, mensaje)
    
    If Not hayDiferencias Then
        MsgBox "No hay diferencias entre la hoja y el código." & vbCrLf & vbCrLf & _
               "Total procedimientos: " & funciones.Count, vbInformation, "Sincronización"
        Exit Sub
    End If
    
    ' Preguntar al usuario qué hacer
    respuesta = MsgBox("Se encontraron diferencias entre la hoja PROCEDIMIENTOS y el código VBA:" & vbCrLf & vbCrLf & _
                       mensaje & vbCrLf & vbCrLf & _
                       "¿Desea actualizar el CÓDIGO con los datos de la hoja?" & vbCrLf & vbCrLf & _
                       "Sí = Actualizar código VBA desde Excel" & vbCrLf & _
                       "No = Actualizar hoja Excel desde código" & vbCrLf & _
                       "Cancelar = No hacer nada", _
                       vbYesNoCancel + vbQuestion, "Sincronizar Metadatos")
    
    Select Case respuesta
        Case vbYes
            ' Actualizar código VBA
            If MsgBox("ADVERTENCIA: Se modificarán los archivos de código VBA." & vbCrLf & _
                     "¿Está seguro de continuar?", vbExclamation + vbYesNo, "Confirmación") = vbYes Then
                Stop
                Call ActualizarCodigoVBA(dictHoja, dictCodigo)
                MsgBox "Código VBA actualizado correctamente." & vbCrLf & _
                       "Se recomienda revisar los cambios.", vbInformation, "Actualización completada"
            End If
            
        Case vbNo
            ' Actualizar hoja Excel
            ws.Cells.Clear
            Call CrearEncabezadosHoja(ws)
            Call VolcarProcedimientosAHoja(ws, funciones)
            MsgBox "Hoja Excel actualizada correctamente con " & funciones.Count & " procedimientos.", _
                   vbInformation, "Actualización completada"
            
        Case vbCancel
            ' No hacer nada
            MsgBox "Operación cancelada. No se realizaron cambios.", vbInformation, "Operación cancelada"
    End Select
End Sub

' ==========================================
' FUNCIÓN 3: MODIFICACIÓN DE SincronizarConHoja
' ==========================================

'@Description: Versión MODIFICADA de SincronizarConHoja que crea backups antes de modificar
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet con datos existentes | funciones: Dictionary con procedimientos actuales del código
Private Sub SincronizarConHoja_ConBackup(ws As Worksheet, funciones As Object)
    Dim dictHoja As Object
    Dim dictCodigo As Object
    Dim hayDiferencias As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim mensaje As String
    Dim rutaBackupVBA As String
    Dim backupHojaOK As Boolean
    
    ' Crear diccionarios para comparación
    Set dictHoja = LeerMetadatosDeHoja(ws)
    Set dictCodigo = CrearDiccionarioProcedimientos(funciones)
    
    ' Comparar y detectar diferencias
    hayDiferencias = HayDiferenciasEnMetadatos(dictHoja, dictCodigo, mensaje)
    
    If Not hayDiferencias Then
        MsgBox "No hay diferencias entre la hoja y el código." & vbCrLf & vbCrLf & _
               "Total procedimientos: " & funciones.Count, vbInformation, "Sincronización"
        Exit Sub
    End If
    
    ' Preguntar al usuario qué hacer
    respuesta = MsgBox("Se encontraron diferencias entre la hoja PROCEDIMIENTOS y el código VBA:" & vbCrLf & vbCrLf & _
                       mensaje & vbCrLf & vbCrLf & _
                       "¿Desea actualizar el CÓDIGO con los datos de la hoja?" & vbCrLf & vbCrLf & _
                       "Sí = Actualizar código VBA desde Excel" & vbCrLf & _
                       "No = Actualizar hoja Excel desde código" & vbCrLf & _
                       "Cancelar = No hacer nada", _
                       vbYesNoCancel + vbQuestion, "Sincronizar Metadatos")
    
    Select Case respuesta
        Case vbYes
            ' ============================================
            ' ACTUALIZAR CÓDIGO VBA (con backup)
            ' ============================================
            
            ' 1. CREAR BACKUP DE CÓDIGO VBA
            MsgBox "Creando copia de seguridad del código VBA...", vbInformation, "Backup en proceso"
            rutaBackupVBA = CrearBackupCodigoVBA()
            
            If rutaBackupVBA = "" Then
                If MsgBox("ADVERTENCIA: No se pudo crear la copia de seguridad del código VBA." & vbCrLf & vbCrLf & _
                         "¿Desea continuar SIN backup?", vbExclamation + vbYesNo, "Error de Backup") = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "Backup de código VBA creado:" & vbCrLf & rutaBackupVBA, vbInformation, "Backup creado"
            End If
            
            ' 2. CONFIRMAR Y ACTUALIZAR
            If MsgBox("ADVERTENCIA: Se modificarán los archivos de código VBA." & vbCrLf & _
                     "Backup guardado en: " & IIf(rutaBackupVBA <> "", rutaBackupVBA, "NO DISPONIBLE") & vbCrLf & vbCrLf & _
                     "¿Está seguro de continuar?", vbExclamation + vbYesNo, "Confirmación") = vbYes Then
                Stop
                Call ActualizarCodigoVBA(dictHoja, dictCodigo)
                
                MsgBox "Código VBA actualizado correctamente." & vbCrLf & _
                       "Backup guardado en: " & rutaBackupVBA, vbInformation, "Actualización completada"
            End If
            
        Case vbNo
            ' ============================================
            ' ACTUALIZAR HOJA EXCEL (con backup)
            ' ============================================
            
            ' 1. CREAR BACKUP DE HOJA EXCEL
            backupHojaOK = CrearBackupHojaExcel(ws)
            
            If Not backupHojaOK Then
                MsgBox "Operación cancelada. No se pudo crear backup de la hoja.", vbExclamation, "Cancelado"
                Exit Sub
            End If
            
            ' 2. ACTUALIZAR HOJA
            ws.Cells.Clear
            Call CrearEncabezadosHoja(ws)
            Call VolcarProcedimientosAHoja(ws, funciones)
            
            MsgBox "Hoja Excel actualizada correctamente con " & funciones.Count & " procedimientos." & vbCrLf & _
                   "Backup guardado como: '" & ws.Name & "_bkp'", vbInformation, "Actualización completada"
            
        Case vbCancel
            MsgBox "Operación cancelada. No se realizaron cambios.", vbInformation, "Operación cancelada"
    End Select
End Sub

'@Description: Lee los metadatos de procedimientos desde la hoja Excel
'@Scope: Privado
'@ArgumentDescriptions: ws: Worksheet de origen
'@Returns: Object | Dictionary con clave compuesta Módulo|Firma -> array(1 To 5) de metadatos
Private Function LeerMetadatosDeHoja(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim fila As Long
    Dim claveCompuesta As String
    Dim metadatos(1 To 5) As String
    Dim modulo As String, firma As String
    
    fila = 2 ' Primera fila de datos
    
    Do While ws.Cells(fila, 1).value <> ""
        modulo = ws.Cells(fila, 1).value
        firma = ws.Cells(fila, 2).value
        claveCompuesta = modulo & "|" & firma
        
        metadatos(1) = ws.Cells(fila, 3).value ' Description
        metadatos(2) = ws.Cells(fila, 4).value ' Category
        metadatos(3) = ws.Cells(fila, 5).value ' Scope
        metadatos(4) = ws.Cells(fila, 6).value ' ArgumentDescriptions
        metadatos(5) = ws.Cells(fila, 7).value ' Returns
        
        If Not dict.Exists(claveCompuesta) Then
            dict.Add claveCompuesta, metadatos
        Else
            Debug.Print "[LeerMetadatosDeHoja] - ADVERTENCIA: Procedimiento duplicado en hoja: " & claveCompuesta
        End If
        
        fila = fila + 1
    Loop
    
    Set LeerMetadatosDeHoja = dict
End Function

'@Description: Crea un diccionario de procedimientos indexado por clave compuesta Módulo|Firma
'@Scope: Privado
'@ArgumentDescriptions: funciones: Dictionary de ParsearProcsDelProyecto
'@Returns: Object | Dictionary con clave compuesta -> clsVBAProcedure
Private Function CrearDiccionarioProcedimientos(funciones As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim proc As clsVBAProcedure
    Dim claveCompuesta As String
    
    For i = 0 To funciones.Count - 1
        Set proc = funciones(i)
        claveCompuesta = proc.Module & "|" & proc.NormalizedSignature
        
        If Not dict.Exists(claveCompuesta) Then
            dict.Add claveCompuesta, proc
        Else
            Debug.Print "[CrearDiccionarioProcedimientos] - ADVERTENCIA: Procedimiento duplicado: " & claveCompuesta
        End If
    Next i
    
    Set CrearDiccionarioProcedimientos = dict
End Function

'@Description: Compara metadatos de hoja vs código y detecta diferencias
'@Scope: Privado
'@ArgumentDescriptions: dictHoja: Dictionary de Excel | dictCodigo: Dictionary del código | mensaje: String de salida con resumen
'@Returns: Boolean | True si hay diferencias
Private Function HayDiferenciasEnMetadatos(dictHoja As Object, dictCodigo As Object, ByRef mensaje As String) As Boolean
    Dim claveCompuesta As Variant
    Dim proc As clsVBAProcedure
    Dim metadatos As Variant
    Dim contadorDiferencias As Long
    Dim contadorNuevos As Long
    Dim contadorEliminados As Long
    
    HayDiferenciasEnMetadatos = False
    contadorDiferencias = 0
    contadorNuevos = 0
    contadorEliminados = 0
    
    ' Verificar cada procedimiento en el código
    For Each claveCompuesta In dictCodigo.Keys
        Set proc = dictCodigo(claveCompuesta)
        
        ' Si el procedimiento no está en la hoja, es nuevo
        If Not dictHoja.Exists(claveCompuesta) Then
            HayDiferenciasEnMetadatos = True
            contadorNuevos = contadorNuevos + 1
        Else
            ' Comparar metadatos
            metadatos = dictHoja(claveCompuesta)
            
            ' Normalizar valores para comparación
            Dim descCode As String, descHoja As String
            Dim catCode As String, catHoja As String
            Dim scopeCode As String, scopeHoja As String
            Dim argsCode As String, argsHoja As String
            Dim retCode As String, retHoja As String
            
            descCode = proc.Description
            descHoja = metadatos(1)
            
            catCode = IIf(proc.Category = DEFAULT_CATEGORY, "", proc.Category)
            catHoja = metadatos(2)
            
            scopeCode = proc.Scope
            scopeHoja = metadatos(3)
            
            argsCode = IIf(proc.ArgumentDescriptions = DEFAULT_NOPARAMS, "", proc.ArgumentDescriptions)
            argsHoja = metadatos(4)
            
            retCode = proc.Returns
            retHoja = metadatos(5)
            
            ' Comparar
            If descCode <> descHoja Or _
               catCode <> catHoja Or _
               scopeCode <> scopeHoja Or _
               argsCode <> argsHoja Or _
               retCode <> retHoja Then
                HayDiferenciasEnMetadatos = True
                contadorDiferencias = contadorDiferencias + 1
            End If
        End If
    Next claveCompuesta
    
    ' Verificar si hay procedimientos en la hoja que ya no existen en el código
    For Each claveCompuesta In dictHoja.Keys
        If Not dictCodigo.Exists(claveCompuesta) Then
            HayDiferenciasEnMetadatos = True
            contadorEliminados = contadorEliminados + 1
        End If
    Next claveCompuesta
    
    ' Generar mensaje resumen
    If HayDiferenciasEnMetadatos Then
        mensaje = "Diferencias detectadas:" & vbCrLf
        If contadorNuevos > 0 Then
            mensaje = mensaje & " " & contadorNuevos & " procedimiento(s) nuevo(s) en código" & vbCrLf
        End If
        If contadorEliminados > 0 Then
            mensaje = mensaje & " " & contadorEliminados & " procedimiento(s) eliminado(s) del código" & vbCrLf
        End If
        If contadorDiferencias > 0 Then
            mensaje = mensaje & " " & contadorDiferencias & " procedimiento(s) con metadatos diferentes" & vbCrLf
        End If
    Else
        mensaje = ""
    End If
End Function

'@Description: Actualiza los archivos de código VBA con los metadatos de la hoja Excel
'@Scope: Privado
'@ArgumentDescriptions: dictHoja: Dictionary de Excel | dictCodigo: Dictionary del código
Private Sub ActualizarCodigoVBA(dictHoja As Object, dictCodigo As Object)
    Dim claveCompuesta As Variant
    Dim proc As clsVBAProcedure
    Dim metadatos As Variant
    Dim vbComp As VBIDE.VBComponent
    Dim CodeModule As VBIDE.CodeModule
    Dim actualizados As Long
    
    On Error GoTo ErrorHandler
    
    actualizados = 0
    
    For Each claveCompuesta In dictHoja.Keys
        If dictCodigo.Exists(claveCompuesta) Then
            Set proc = dictCodigo(claveCompuesta)
            metadatos = dictHoja(claveCompuesta)
            
            ' Actualizar objeto en memoria
            proc.Description = metadatos(1)
            proc.Category = metadatos(2)
            proc.Scope = metadatos(3)
            proc.ArgumentDescriptions = metadatos(4)
            proc.Returns = metadatos(5)
            
            ' Actualizar archivo VBA
            Set vbComp = ThisWorkbook.VBProject.VBComponents(proc.Module)
            Set CodeModule = vbComp.CodeModule
            
            Call ActualizarMetadatosEnCodigo(CodeModule, proc)
            
            actualizados = actualizados + 1
        End If
    Next claveCompuesta
    
    Debug.Print "[ActualizarCodigoVBA] - " & actualizados & " procedimiento(s) actualizado(s) en código"
    
    Exit Sub
ErrorHandler:
    Debug.Print "[ActualizarCodigoVBA] - Error: " & Err.Description
    MsgBox "Error al actualizar código VBA: " & Err.Description & vbCrLf & _
           "Procedimiento: " & proc.Module & "." & proc.Name, vbCritical, "Error"
End Sub

'@Description: Actualiza los metadatos de un procedimiento específico en su módulo de código
'@Scope: Privado
'@ArgumentDescriptions: CodeModule: Módulo VBA donde está el procedimiento | proc: Objeto clsVBAProcedure con nuevos metadatos
Private Sub ActualizarMetadatosEnCodigo(CodeModule As VBIDE.CodeModule, proc As clsVBAProcedure)
    Dim i As Long
    Dim lineText As String
    Dim nuevosMetadatos As String
    Dim inicioMetadatos As Long
    Dim finMetadatos As Long
    Dim lineasEliminadas As Long
    
    On Error GoTo ErrorHandler
    
    ' Buscar y eliminar metadatos existentes (comentarios con @ antes de la firma)
    inicioMetadatos = proc.procStartLine
    finMetadatos = proc.procSignatureLine - 1
    lineasEliminadas = 0
    
    ' Eliminar metadatos antiguos de abajo hacia arriba para no desplazar índices
    For i = finMetadatos To inicioMetadatos Step -1
        lineText = Trim$(CodeModule.Lines(i, 1))
        If Left$(lineText, 1) = "'" And InStr(lineText, "@") > 0 Then
            CodeModule.DeleteLines i, 1
            lineasEliminadas = lineasEliminadas + 1
        End If
    Next i
    
    ' Generar nuevos metadatos
    nuevosMetadatos = GenerarMetadatosFormateados(proc)
    
    ' Insertar nuevos metadatos antes de la firma del procedimiento
    ' (ajustando la línea por las eliminaciones)
    If nuevosMetadatos <> "" Then
        Dim lineaInsercion As Long
        lineaInsercion = proc.procSignatureLine - lineasEliminadas
        CodeModule.InsertLines lineaInsercion, nuevosMetadatos
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "[ActualizarMetadatosEnCodigo] - Error en " & proc.Module & "." & proc.Name & ": " & Err.Description
End Sub

'@Description: Genera el texto formateado de los metadatos para insertar en el código
'@Scope: Privado
'@ArgumentDescriptions: proc: Objeto clsVBAProcedure con metadatos
'@Returns: String | Texto con metadatos formateados (incluye vbCrLf al final)
Private Function GenerarMetadatosFormateados(proc As clsVBAProcedure) As String
    Dim resultado As String
    
    resultado = ""
    
    ' Description (siempre incluir si no está vacío)
    If proc.Description <> "" Then
        resultado = resultado & "'@Description: " & proc.Description & vbCrLf
    End If
    
    ' Category (omitir si es la categoría por defecto)
    If proc.Category <> "" And proc.Category <> DEFAULT_CATEGORY Then
        resultado = resultado & "'@Category: " & proc.Category & vbCrLf
    End If
    
    ' Scope
    If proc.Scope <> "" Then
        resultado = resultado & "'@Scope: " & proc.Scope & vbCrLf
    End If
    
    ' ArgumentDescriptions (omitir si es "(sin parámetros)")
    If proc.ArgumentDescriptions <> "" And proc.ArgumentDescriptions <> DEFAULT_NOPARAMS Then
        resultado = resultado & "'@ArgumentDescriptions: " & proc.ArgumentDescriptions & vbCrLf
    End If
    
    ' Returns
    If proc.Returns <> "" Then
        resultado = resultado & "'@Returns: " & proc.Returns & vbCrLf
    End If
    
    GenerarMetadatosFormateados = resultado
End Function

' ==========================================
' FIN DE EXTENSIÓN
' ==========================================
'
' RESUMEN DE FUNCIONES CREADAS:
'
' 1. WriteProcedimientosSheet() - Procedimiento principal público
' 2. SheetExists() - Verifica existencia de hoja
' 3. CrearHojaProcedimientos() - Crea hoja nueva
' 4. CrearEncabezadosHoja() - Formatea encabezados
' 5. VolcarProcedimientosAHoja() - Vuelca datos a Excel
' 6. SincronizarConHoja() - Gestiona sincronización bidireccional
' 7. LeerMetadatosDeHoja() - Lee desde Excel
' 8. CrearDiccionarioProcedimientos() - Crea diccionario del código
' 9. HayDiferenciasEnMetadatos() - Detecta y reporta diferencias
' 10. ActualizarCodigoVBA() - Actualiza archivos VBA
' 11. ActualizarMetadatosEnCodigo() - Modifica metadatos en código
' 12. GenerarMetadatosFormateados() - Genera formato de metadatos
'
' MEJORAS RESPECTO A LA VERSIÓN ORIGINAL:
'  Clave compuesta Módulo|Firma para evitar conflictos
'  Columna adicional "Módulo" en Excel
'  Mensajes más informativos con contadores
'  Mejor formato visual en hoja Excel (colores alternados, bordes)
'  Confirmación adicional antes de modificar código
'  Normalización de valores para comparación correcta
'  Manejo de duplicados con advertencias en Debug
'  Anchos de columna optimizados
'
' ==========================================

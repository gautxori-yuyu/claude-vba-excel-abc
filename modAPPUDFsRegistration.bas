Attribute VB_Name = "modAPPUDFsRegistration"
' ==========================================
' SISTEMA DE AUTO-REGISTRO Y DESREGISTRO DE UDFs EN EXCEL
' ==========================================
' Este mÃ³dulo implementa un sistema completo para:
' - Detectar funciones UDF en el proyecto VBA
' - Registrar y desregistrar UDFs en Excel
' - Persistir el estado en el Registro de Windows
' - Soportar fallback dinÃ¡mico ante fallos
'
' Se apoya en ParsearProcsDelProyecto() y clsVBAProcedure
' ==========================================

'FIXME: revisarlo. Las clases clsVBAProcedure.cls,Â modAPPUDFsRegistration.bas yÂ modUTILSProcedureParsing.bas Pretenden
'  - la identificaciÃ³n de todos los procedimientos en el proyecto VBA (clsVBAProcedure.cls yÂ modUTILSProcedureParsing.bas)
'  - la identificaciÃ³n de todas las funciones Que pudieran ser udfs para registrarlas en  la carga del XLM (modAPPUDFsRegistration.bas).Â 
'
' [DONE 2026-01-01] Extendido el reconocimiento de atributos de documentacion:
'  - Soporta formatos: '@Tag: Valor', '@Tag("Valor")', '@Tag "Valor"'
'  - Nuevos atributos: @Example, @Raises, @Throws, @Dependencies
'  - Ver clsVBAProcedure.ParsearMetadataCompleta() para detalles

'@Folder "1-Inicio e Instalacion.Gestion de modulos y procs"
Option Explicit

Private Const MODULE_NAME As String = "modAPPUDFsRegistration"

Private bVerbose As Boolean
' ---------------------------------------------------------------------
' PUNTOS DE ENTRADA PÃBLICOS
' ---------------------------------------------------------------------

'@Description: Punto de entrada sin parÃ¡metros para auto-registrar todas las UDFs del proyecto.
'@Scope: Manipula el registro de funciones UDF en Excel.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: Registro UDF
Sub AutoRegistrarTodasLasUDFsNOPARAMS()
Attribute AutoRegistrarTodasLasUDFsNOPARAMS.VB_ProcData.VB_Invoke_Func = " \n0"
    Call AutoRegistrarTodasLasUDFs
End Sub

'@Description: Auto-registra todas las funciones UDF del proyecto VBA, opcionalmente filtrando por metadatos.
'@Scope: Analiza procedimientos del proyecto y registra funciones UDF en Excel.
'@ArgumentDescriptions: bOnlyWithMetadata: Solo registrar UDFs con metadatos explÃ­citos | bVerbose_: Mostrar salida detallada en ventana Inmediato
'@Returns: (ninguno)
'@Category: Registro UDF
Public Sub AutoRegistrarTodasLasUDFs(Optional bOnlyWithMetadata As Boolean = False, Optional bVerbose_ As Boolean = False)
Attribute AutoRegistrarTodasLasUDFs.VB_ProcData.VB_Invoke_Func = " \n0"
    
    On Error GoTo ErrorHandler
    
    Dim funciones As Object
    Dim metadata As clsVBAProcedure
    Dim key As Variant
    
    bVerbose = bVerbose_
    
    ' Parsear todas las UDFs del proyecto
    Set funciones = ParsearProcsDelProyecto()
    
    If funciones.Count > 0 Then
        
        Application.ScreenUpdating = False
        ThisWorkbook.Activate
        
        ' Registrar cada funciÃ³n
        For Each key In funciones
            Set metadata = funciones.Item(key)
            
            If metadata.ProcedureType <> udf Then
                funciones.Remove key
                
            ElseIf (Not bOnlyWithMetadata Or metadata.HasMetadata) Then
                If Not RegistrarUDF(metadata) Then funciones.Remove key
            End If
        Next key
        
        Application.ScreenUpdating = True
        
        ' Persistir lista para desinstalaciÃ³n posterior
        Call GuardarListaFuncionesRegistradas(funciones)
        
        LogInfo "modAPPUDFsRegistration", "[AutoRegistrarTodasLasUDFs] - " & funciones.Count & " funciones registradas correctamente."
        
    Else
        LogWarning "modAPPUDFsRegistration", "[AutoRegistrarTodasLasUDFs] - No se encontraron funciones UDF vÃ¡lidas."
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modAPPUDFsRegistration", "[AutoRegistrarTodasLasUDFs] - Error", , Err.Description
End Sub

'@Description: Desregistra todas las UDFs previamente registradas utilizando la informaciÃ³n persistida.
'@Scope: Elimina registros de funciones UDF en Excel.
'@ArgumentDescriptions: bVerbose_: Mostrar salida detallada en ventana Inmediato
'@Returns: (ninguno)
'@Category: Registro UDF
Public Sub DesregistrarTodasLasUDFs(Optional bVerbose_ As Boolean = False)
Attribute DesregistrarTodasLasUDFs.VB_ProcData.VB_Invoke_Func = " \n0"
    
    On Error GoTo ErrorHandler
    
    Dim listaFunciones As String
    Dim funciones() As String
    Dim i As Long
    
    bVerbose = bVerbose_
    
    ' Obtener lista persistida
    listaFunciones = ObtenerListaFuncionesRegistradas()
    
    If listaFunciones <> "" Then
        
        funciones = Split(listaFunciones, "|")
        
        For i = LBound(funciones) To UBound(funciones)
            If Trim(funciones(i)) <> "" Then
                DesregistrarUDF funciones(i)
            End If
        Next i
        
        ' Limpiar lista guardada
        BorrarListaFuncionesRegistradas
        
        LogInfo "modAPPUDFsRegistration", "[DesregistrarTodasLasUDFs] : " & (UBound(funciones) + 1) & " funciones desregistradas."
    Else
        LogWarning "modAPPUDFsRegistration", "[DesregistrarTodasLasUDFs] - No hay lista guardada, intentando desregistro manual."
        ' Fallback: desregistrar todas las funciones encontradas ahora
        DesregistrarTodasLasFuncionesActuales
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "modAPPUDFsRegistration", "[DesregistrarTodasLasUDFs] - Error", , Err.Description
    ' Intentar desregistro manual como respaldo
    DesregistrarTodasLasFuncionesActuales
End Sub

' ---------------------------------------------------------------------
' REGISTRO Y DESREGISTRO INDIVIDUAL DE UDFs
' ---------------------------------------------------------------------

'@Description: Registra una funciÃ³n UDF individual en Excel usando sus metadatos.
'@Scope: Llama a Application.MacroOptions para registrar funciones.
'@ArgumentDescriptions: metadata: Objeto clsVBAProcedure con informaciÃ³n de la funciÃ³n
'@Returns: Boolean | True si el registro fue exitoso
'@Category: Registro UDF
Private Function RegistrarUDF(metadata As clsVBAProcedure) As Boolean
    
    On Error Resume Next
    
    Dim strDescription As String
    Dim argArray() As String
    Dim i As Long
    
    strDescription = metadata.Description
    
    If metadata.Scope <> "" Then strDescription = strDescription & ". Aplica a: " & metadata.Scope
    If metadata.Returns <> "" Then strDescription = strDescription & ". Devuelve: " & metadata.Returns
    
    strDescription = Replace(strDescription, "..", ".")
    strDescription = Left("[" & metadata.Module & "] " & strDescription, 255)
    
    If metadata.ArgumentDescriptions <> "" Then
        ' Procesar argumentos separados por "|"
        argArray = Split(metadata.ArgumentDescriptions, "|")
        
        For i = LBound(argArray) To UBound(argArray)
            argArray(i) = Trim(argArray(i))
        Next i
        
        ' Registrar con argumentos ("'" & APP_NAME & ".xlam'!" & ?)
        Application.MacroOptions _
            Macro:=metadata.Name, _
            Description:=strDescription, _
            Category:=metadata.Category, _
            ArgumentDescriptions:=argArray
    Else
        ' Registrar sin descripciÃ³n de argumentos
        Application.MacroOptions _
            Macro:=metadata.Name, _
            Description:=strDescription, _
            Category:=metadata.Category
    End If
    
    If Err.Number <> 0 Then
        LogError "modAPPUDFsRegistration", "[RegistrarUDF] - Error registrando '" & metadata.Name & "'", , Err.Description
    Else
        If bVerbose Then LogInfo "modAPPUDFsRegistration", "[RegistrarUDF] - Registrada: " & metadata.Name
        RegistrarUDF = True
    End If
    
    On Error GoTo 0
End Function

'@Description: Elimina el registro de una funciÃ³n UDF individual en Excel.
'@Scope: Limpia metadatos de una funciÃ³n registrada.
'@ArgumentDescriptions: nombreFuncion: Nombre de la funciÃ³n UDF
'@Returns: Boolean | True si se desregistrÃ³ correctamente
'@Category: Registro UDF
Private Function DesregistrarUDF(nombreFuncion As String) As Boolean
    
    On Error Resume Next
    
    Application.MacroOptions _
        Macro:=Trim(nombreFuncion), _
        Description:=Empty, _
        Category:=Empty
    
    If Err.Number = 0 Then
        If bVerbose Then LogInfo "modAPPUDFsRegistration", "[DesregistrarUDF] - Desregistrada: " & nombreFuncion
        DesregistrarUDF = True
    End If
    
    On Error GoTo 0
    
End Function

' ---------------------------------------------------------------------
' PERSISTENCIA EN REGISTRO DE WINDOWS
' ---------------------------------------------------------------------

'@Description: Guarda en el registro de Windows la lista de UDFs registradas.
'@Scope: Escribe valores REG_SZ en el registro del sistema.
'@ArgumentDescriptions: funciones: Dictionary con objetos clsVBAProcedure
'@Returns: (ninguno)
'@Category: Persistencia
Private Sub GuardarListaFuncionesRegistradas(funciones As Object)
    
    Dim lista As Variant
    Dim key As Variant
    Dim metadata As clsVBAProcedure
    
    For Each key In funciones.Keys()
        Set metadata = funciones(key)
        If lista <> "" Then lista = lista & "|"
        lista = lista & metadata.Name
    Next key
    
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite CFG_RUTA_UDFS, lista, "REG_SZ"
    
    If Err.Number <> 0 Then
        LogError "modAPPUDFsRegistration", "[DesregistrarUDF] - No se pudo guardar lista en registro", , Err.Description
    Else
        If bVerbose Then LogInfo "modAPPUDFsRegistration", "[DesregistrarUDF] - Lista de UDFs guardada en registro."
    End If
    
    On Error GoTo 0
End Sub

'@Description: Obtiene del registro de Windows la lista de UDFs registradas.
'@Scope: Lee valores REG_SZ del registro del sistema.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: String | Lista de nombres de funciones separadas por "|"
'@Category: Persistencia
Private Function ObtenerListaFuncionesRegistradas() As String
    On Error Resume Next
    ObtenerListaFuncionesRegistradas = CreateObject("WScript.Shell").RegRead(CFG_RUTA_UDFS)
    
    If Err.Number <> 0 Then
        ObtenerListaFuncionesRegistradas = ""
        LogError "modAPPUDFsRegistration", "[ObtenerListaFuncionesRegistradas] - No se encontrÃ³ lista guardada en registro."
    End If
    
    On Error GoTo 0
End Function

'@Description: Elimina del registro de Windows la lista de UDFs almacenada.
'@Scope: Borra valores del registro del sistema.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: Persistencia
Private Sub BorrarListaFuncionesRegistradas()
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete CFG_RUTA_UDFS
    
    If Err.Number = 0 Then
        LogInfo "modAPPUDFsRegistration", "[BorrarListaFuncionesRegistradas] - Lista de UDFs eliminada del registro."
    End If
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
' FALLBACK: DESREGISTRO DINÃMICO
' ---------------------------------------------------------------------

'@Description: Desregistra dinÃ¡micamente todas las funciones encontradas en el proyecto.
'@Scope: Analiza procedimientos actuales y elimina su registro.
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Category: Registro UDF
Private Sub DesregistrarTodasLasFuncionesActuales()
    
    Dim funciones As Object
    Dim key As Variant
    Dim Count As Long
    
    Set funciones = ParsearProcsDelProyecto()
    
    For Each key In funciones
        If DesregistrarUDF(funciones(key).Name) Then Count = Count + 1
    Next key
    
    If Count > 0 Then
        If bVerbose Then LogInfo "modAPPUDFsRegistration", "[DesregistrarTodasLasFuncionesActuales] - Desregistro dinÃ¡mico: " & Count & " funciones procesadas."
    Else
        LogInfo "modAPPUDFsRegistration", "[DesregistrarTodasLasFuncionesActuales] - Desregistro dinÃ¡mico: No se encontraron funciones para desregistrar."
    End If
End Sub



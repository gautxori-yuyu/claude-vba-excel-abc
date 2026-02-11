VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Representa un Workbook abierto en el contexto de la aplicación.
' Es la única capa que conoce directamente al Workbook.
' Todas las operaciones que requieran wb deben pasar por aquí.

'@Folder "2-Infraestructura.FS"
Option Explicit

Private Const MODULE_NAME As String = "clsFileXLS"

Private p_wb As Workbook
Private p_PrivSheetsCol As Collection ' Colección de NOMBRES de hojas PRIVADAS (solo para uso interno)
Private p_PublSheetsCol As Collection ' Colección de NOMBRES de hojas PUBLICAS (SE PUEDEN COMPARTIR COMO DOCUMENTO EDITABLE)

Private p_ObjectKey As Double

Private p_WbTipoFichero As TipoArchivo

Private p_RibbonMode As eRibbonMode ' Estado del ribbon específico para este libro

'--------------------------------------------------------------
' @Description: Información generica de los archivos de Excel, relacionada con mi aplicación
'--------------------------------------------------------------
Private Type T_InfoArchivo
    EsValido As Boolean
    TipoDetectado As TipoArchivo
    Customer As String
    OpportunityNr As String
End Type

' ==================================================================
' IMPLEMENTA LA INTERFAZ IFile
' ==================================================================
Implements IFile

' ==================================================================
' PROPIEDADES PÚBLICAS (API pública de la clase)
' ==================================================================

Friend Property Get ObjectKey() As Double
    ObjectKey = p_ObjectKey
End Property

' @Description: Ruta completa (FullName), con fallback seguro.
Public Property Get Path() As String
    On Error Resume Next
    If p_ObjectKey <> 0 Then
        Path = p_wb.FullName
    Else
        Path = "(cerrado)"
    End If
    On Error GoTo 0
End Property

Public Property Get Name() As String
    On Error Resume Next
    If p_ObjectKey <> 0 Then
        Name = p_wb.Name
    Else
        Name = "(cerrado)"
    End If
    On Error GoTo 0
End Property

'--------------------------------------------------------------
' Estas propiedades Private son generadas automáticamente por VBA
' cuando escribes "Implements IFile"
' Delegan a las propiedades públicas
'--------------------------------------------------------------
Private Property Get IFile_ObjectKey() As Double
    IFile_ObjectKey = Me.ObjectKey
End Property

Private Property Get IFile_Path() As String
    IFile_Path = Me.Path
End Property

Private Property Get IFile_Name() As String
    IFile_Name = Me.Name
End Property

Private Property Get IFile_IsOpportunityFile() As Boolean
    IFile_IsOpportunityFile = Me.IsOpportunityFile
End Property

' ==================================================================
' Inicialización y estados asociados al fichero
' ==================================================================
Private Sub Class_Initialize()
    ' Estado por defecto del ribbon para este libro
    p_RibbonMode = DEFAULT_RIBBONMODE
End Sub

'@Description: Compara dos archivos para ver si son el mismo
Public Function Equals(f As clsFileXLS) As Boolean
    Equals = (ObjPtr(Me) = ObjPtr(f))
End Function

'--------------------------------------------------------------
' @Description: Control de estado del ribbon por libro
'--------------------------------------------------------------
'@Description: Acceso directo al modo del ribbon para este libro
Public Property Get RibbonMode() As eRibbonMode
    RibbonMode = p_RibbonMode
End Property

Public Property Let RibbonMode(Mode As eRibbonMode)
    p_RibbonMode = Mode
End Property

' ==================================================================
' Metodos publicos: vinculo entre Workbook y fichero
' ==================================================================
'--------------------------------------------------------------
' @Description: Asocia esta instancia a un Workbook.
'               Solo se puede llamar una vez por instancia.
'--------------------------------------------------------------
Public Sub BindTo(Wb As Workbook)
    If Not p_wb Is Nothing Then
        Err.Raise vbObjectError + 570, TypeName(Me), "La instancia ya está asociada a un Workbook."
    End If
    If Wb Is Nothing Then
        Err.Raise vbObjectError + 571, TypeName(Me), "No se puede asociar a Nothing."
    End If
    Set p_wb = Wb
    p_ObjectKey = ObjPtr(p_wb)
    AnalizarWorkBook
    LogDebug MODULE_NAME, "[BindTo] Asociado a: " & Me.Path & " (key: " & p_ObjectKey & ")"
End Sub
' Factory Method estático: "Creador especializado"
Public Static Function CreateFromWorkbook(Wb As Workbook) As clsFileXLS
    Dim instance As clsFileXLS
    Set instance = New clsFileXLS
    instance.BindTo Wb
    Set CreateFromWorkbook = instance
End Function
'--------------------------------------------------------------
' @Description: Acceso seguro al Workbook subyacente.
'               Lanza error si el wb ya no es válido.
'--------------------------------------------------------------
Public Property Get Wb() As Workbook
    If p_ObjectKey = 0 Then
        Err.Raise vbObjectError + 572, TypeName(Me), "Workbook ya no está disponible (cerrado o inválido)."
    End If
    Set Wb = p_wb
End Property

' ==================================================================
' IDENTIFICACION DEL TIPO DE FICHERO
' ==================================================================

'--------------------------------------------------------------
' @Description: Tipo detectado (con análisis bajo demanda).
'--------------------------------------------------------------
Public Property Get tipo() As TipoArchivo
    If p_ObjectKey = 0 Then Exit Property
    If p_WbTipoFichero = TipoArchivo.UnDef Or IsEmpty(p_WbTipoFichero) Then
        Call Me.AnalizarWorkBook
    End If
    tipo = p_WbTipoFichero
End Property

'--------------------------------------------------------------
' @Description: IMplementacion de interfaz
'--------------------------------------------------------------
Public Property Get IsOpportunityFile() As Boolean
    ' Si está en una carpeta de oportunidad... o es de gasnet, valoracion, oferta, ...
    IsOpportunityFile = (p_WbTipoFichero > TipoArchivo.Unknown)
End Property

Public Property Get EsOportunidad() As Boolean
    ' pte de implementar (DE MOMENTO no hay fichero de excel para una oportunidad...)
End Property

'--------------------------------------------------------------
' @Description: Realiza el análisis completo del Workbook.
'               Actualiza internamente `p_WbTipoFichero`.
'--------------------------------------------------------------
Public Sub AnalizarWorkBook(Optional ByVal Wb As Workbook = Nothing)
    If p_ObjectKey = 0 Then Exit Sub
    
    If Not Wb Is Nothing And p_wb Is Nothing Then
        BindTo Wb
    ElseIf Not Wb Is Nothing And Not (p_wb Is Wb) Then
        ' El libro ya estaba definido y se intenta analizar uno distinto
        LogError MODULE_NAME, "[AnalizarWorkBook] Intentando analizar un libro diferente al asociado"
        Err.Raise vbObjectError + 573, MODULE_NAME, _
                  "Esta instancia ya esta asociada a un Workbook diferente. (no debería pasar, revisa tu codigo)"
    End If
    
    ' Determinar tipo por nombre de fichero
    p_WbTipoFichero = TipoPorNombreFichero()
    
    ' Determinar tipo por CONTENIDO de fichero
    Stop
    Select Case p_WbTipoFichero
        Case TipoArchivo.CGASING_CurvasRendimiento, _
                TipoArchivo.CGASING_CalcOpc, _
                TipoArchivo.PlantillaBudget, _
                TipoArchivo.PlantillaOferta
        Case TipoArchivo.CGASING_CalcResults
            ' TODO: SE PUEDE SEGUIR FILTRANDO POR NOMBRES, ver en VBScript
            Stop
        Case TipoArchivo.oportunidad
            ' de momento NO HAY CRITERIO FIJO para el nombre de fichero
            ' que acote tipos con precisión (PTE)...
            ' y requeriría analizar contenido
            p_WbTipoFichero = TipoPorContenidoLibro
        Case Else
            ' Aqui puede haber de todo, PTE seguir fijando criterios... o descartando ficheros
            Stop
            ' p_WbTipoFichero = TipoPorContenidoLibro
    End Select
    
    ' TODO: Esto NO va a aqui, va en cada ExcelXXXSource
    ' Extraer información del nombre DE LA CARPETA DE OPORTUNIDAD, ... o de otros ficheros en ella!!!
    ' p_WbTipoFichero.Customer = Customer(p_wb)
    ' p_WbTipoFichero.OpportunityNr = OpportunityNr(p_wb)
End Sub

'--------------------------------------------------------------
' @Description: Determinan el tipo de archivo según nombre
'--------------------------------------------------------------
Private Function TipoPorNombreFichero() As TipoArchivo
    Dim nombreArchivo As String
    nombreArchivo = p_wb.Name
    
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    
    ' Verificar patrones en el nombre
    re.Pattern = GASVBNET_NAME_PATTERN
    If re.Test(p_wb.Name) And TipoPorNombreFichero = UnDef Then
        If InStr(1, nombreArchivo, "_calc", vbTextCompare) > 0 Then
            TipoPorNombreFichero = TipoArchivo.CGASING_CalcOpc
        Else
            ' Define un tippo genérico de calculo gasVBNet
            ' TODO: seguir fijando criterios para ficheros GasVBNet
            Stop
            TipoPorNombreFichero = TipoArchivo.CGASING_CalcResults
        End If
    End If
    
    re.Pattern = PLANTILLABUDGET_PATTERN
    If re.Test(p_wb.Name) And TipoPorNombreFichero = UnDef Then
        TipoPorNombreFichero = TipoArchivo.PlantillaBudget
    End If
    
    re.Pattern = PLANTILLAQUOTATION_PATTERN
    If re.Test(p_wb.Name) And TipoPorNombreFichero = UnDef Then
        TipoPorNombreFichero = TipoArchivo.PlantillaOferta
    End If
    
    If modAPPFileNames.QuoteNrInFileName(p_wb) <> "" And modAPPFileNames.CustomerInFileName(p_wb) <> "" _
             And TipoPorNombreFichero = UnDef Then
        Stop
        ' si tiene el patrón de oportunidad
        TipoPorNombreFichero = TipoArchivo.oportunidad
    Else
        TipoPorNombreFichero = TipoArchivo.Unknown
    End If
End Function

'--------------------------------------------------------------
' @Description: Determinan el tipo de archivo según contenido
'--------------------------------------------------------------
Private Function TipoPorContenidoLibro() As TipoArchivo
    Select Case p_WbTipoFichero
        Case TipoArchivo.CGASING_CalcResults
            ' Filtrar qué tipo de resultado es, según contenido
            If EsContenidoLibroTablasCurvasRto Then
                TipoPorContenidoLibro = TipoArchivo.CGASING_CurvasRendimiento
            End If
        Case Else
            ' Si NO tiene un contenido predefinido... habría que pasarlo por TODAS las comprobaciones??
    End Select
End Function

Private Function EsContenidoLibroTablasCurvasRto() As Boolean
    'On Error GoTo ErrorHandler
    Dim hoja As Worksheet
    Dim R As Range
    Dim encabezados As Range, datos As Range
    Dim formula As String
    
    For Each hoja In p_wb.Worksheets
        If Not IsNumeric(hoja.Name) Then GoTo Siguiente
        
        ' Validar que solo hay una tabla en la hoja
        If hoja.ListObjects.Count <> 1 Then GoTo ErrorHandler
        
        ' Validar que la tabla comience en A1
        If hoja.ListObjects(1).Range.Cells(1, 1).Address <> "$A$1" Then GoTo ErrorHandler
        
        Set R = hoja.Range("A1").CurrentRegion
        If R.Rows.Count < 2 Or R.Columns.Count < 2 Then GoTo Siguiente
        
        ' Validar encabezados (fila 1, desde la segunda columna)
        Set encabezados = R.Rows(1).Offset(0, 1).Resize(1, R.Columns.Count - 1)
        formula = "SUMPRODUCT(--ISNUMBER(SEARCH(""("", " & encabezados.Address(External:=True) & ")))"
        If Evaluate(formula) <> encabezados.Columns.Count Then GoTo ErrorHandler
        
        ' Validar datos numéricos (todo menos la primera fila y primera columna)
        Set datos = R.Offset(1, 1).Resize(R.Rows.Count - 1, R.Columns.Count - 1)
        formula = "SUMPRODUCT(--ISNUMBER(" & datos.Address(External:=True) & "))"
        If Evaluate(formula) <> datos.Cells.Count Then GoTo ErrorHandler
        
        EsContenidoLibroTablasCurvasRto = True
        Exit Function
Siguiente:
    Next hoja
    
ErrorHandler:
    EsContenidoLibroTablasCurvasRto = False
End Function

' ==================================================================
'
' ==================================================================
'--------------------------------------------------------------
' @Description: Determina si este libro debe moverse a una oportunidad dada.
'               Solo depende de su estado interno y la ruta objetivo.
'               NOTA: de momento usa funciones globales (futuro: encapsular aquí).
'--------------------------------------------------------------
Public Function DebeMoverseAOp(rutaOportunidadRelativa As String, rutaBaseOportunidades As String) As Boolean
    ' Usa Me.Path, Me.AnalizarWorkBook, OpportunityNr(Me.WB), etc.
    Dim Wb As Workbook, rutaOportunidad As String
    rutaOportunidad = rutaBaseOportunidades & "\" & rutaOportunidadRelativa
    Set Wb = p_wb
    
    ' 1. Verificar si es un fichero de oportunidad válido
    If Not EsFicheroOportunidad() Then Exit Function
    
    ' 2. Verificar si ya está en la carpeta correcta
    If InStr(1, Wb.Path, rutaOportunidad, vbTextCompare) > 0 Then
        ' Ya está en la carpeta correcta
        Exit Function
    End If
    
    ' 3. Si tiene un número de oportunidad válido pero no está en la carpeta correcta
    ' DE MOMENTO ESTA RESTRICCION SE ELIMINA: HAY QUE PERFECCIONAR LA IDENTIFICACION DE LOS FICHEROS DE OPORTUNIDAD.
'    Dim nrOportunidad As String
'    nrOportunidad = QuoteNr(Wb)
    'If nrOportunidad <> "" Then
    DebeMoverseAOp = True
    'End If
End Function

'--------------------------------------------------------------
' @Description: Mueve físicamente el libro abierto a la oportunidad indicada.
'               Usa SaveCopyAs + borrado del original (sin cerrar wb).
'               Usa lógica robusta de backup/truncado.
'--------------------------------------------------------------
Public Sub MoverAOp(rutaOportunidadRelativa As String, rutaBaseOportunidades As String)
    ' Usa toda la lógica robusta (SaveCopyAs + backup + truncado...)
    ' Actualiza internamente `Me.m_wb` (el path no cambia en VBA, pero sabemos que está en nueva ubicación)
    ' Opcional: invalidar cache en FileManager
    If p_ObjectKey = 0 Then Exit Sub
    
    Dim nombreArchivo As String
    nombreArchivo = p_wb.Name
    Dim respuesta As TDRESULT
    respuesta = ShowTaskDialogYesNo("Mover archivo a oportunidad", _
                                    "¿Deseas mover este archivo?", _
                                    "Se moverá '" & nombreArchivo & "' a la carpeta de la oportunidad seleccionada.")

    If respuesta <> vbYes Then Exit Sub
        
    On Error GoTo ErrorHandler
    
    ' 1. Guardar el libro si hay cambios
    If p_wb.Saved = False Then p_wb.Save
    
    ' 2. Construir nueva ruta
    Dim rutaOrigen As String: rutaOrigen = p_wb.FullName
    Dim nuevaRuta As String, rutaDestinoBase As String
    rutaDestinoBase = rutaBaseOportunidades & rutaOportunidadRelativa & "\2.CALCULO TECNICO"
    nuevaRuta = rutaDestinoBase & "\" & nombreArchivo
    
    ' 3. Mover el archivo
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 3.1 mover a old el fichero, si ya existe.
    If fso.FileExists(nuevaRuta) Then
        Dim rutaDestinoBackup As String
        Dim contador As Long: contador = 1
        Do
            rutaDestinoBackup = fso.BuildPath(rutaDestinoBase, _
                                              fso.GetBaseName(nombreArchivo) & "_old (" & contador & ")." & fso.GetExtensionName(nombreArchivo))
            contador = contador + 1
        Loop While fso.FileExists(rutaDestinoBackup)
        
        ' Renombrar el existente
        fso.MoveFile nuevaRuta, rutaDestinoBackup
        LogInfo MODULE_NAME, "[MoverAOp] Archivo existente renombrado a: " & rutaDestinoBackup
    End If
    
    ' === 3.2 Verificar longitud de ruta (MAX_PATH = 260 en Windows sin \\?\) ===
    If Len(nuevaRuta) > 259 Then
        ' Acortar nombre de archivo manteniendo extensión y sufijo "_trunc"
        Dim maxNombreLen As Long
        maxNombreLen = 259 - Len(rutaDestinoBase) - 1 ' -1 por barra
        If maxNombreLen < 6 Then
            Err.Raise vbObjectError + 574, "MoverAOp", _
                      "Ruta de destino demasiado larga incluso tras truncar el nombre."
        End If
        
        Dim nuevoNombre As String
        nuevoNombre = Left(fso.GetBaseName(nombreArchivo), maxNombreLen - Len(fso.GetExtensionName(nombreArchivo)) - 8) _
                    & fso.GetExtensionName(nombreArchivo)
        nuevaRuta = fso.BuildPath(rutaDestinoBase, nuevoNombre)
        ShowTaskDialogError "Nombre de archivo truncado", _
                            "Límite de ruta de Windows", _
                            "El nombre del archivo ha sido acortado a: " & nuevoNombre
    End If
    
    ' === 4. Guardar copia en destino ===
    'p_wb.SaveCopyAs nuevaRuta
    p_wb.SaveAs fileName:=nuevaRuta, FileFormat:=p_wb.FileFormat, Password:="", CreateBackup:=False, ReadOnlyRecommended:=False
    LogInfo MODULE_NAME, "[MoverAOp] Copia guardada en: " & nuevaRuta
    
    ' === 5. Borrar original (solo si la copia fue exitosa) ===
    If fso.FileExists(nuevaRuta) Then
        fso.DeleteFile rutaOrigen, True          ' True = Force
        LogInfo MODULE_NAME, "[MoverAOp] Original eliminado: " & rutaOrigen
    Else
        Err.Raise vbObjectError + 575, "MoverAOp", _
                  "No se pudo verificar la copia en destino. Operación cancelada."
    End If
    
    ' === 6. Reasignar .Path y .FullName (simbólicamente, VBA no lo permite directamente) ===
    ' El libro sigue abierto, pero su ruta "efectiva" ahora es la nueva.
    ' Excel lo manejará correctamente al guardar de nuevo.

    ShowTaskDialogError "Operación completada", _
                        "Archivo movido correctamente", _
                        "El archivo '" & nombreArchivo & "' se ha movido a la carpeta de la oportunidad."
    
    Exit Sub
    
ErrorHandler:
    LogCurrentError MODULE_NAME, "Error al mover el archivo"
    Err.Raise Err.Number, MODULE_NAME & " [MoverAOp]", Err.Description
    ' No hacer nada más: el libro sigue abierto y seguro en su ubicación original
End Sub



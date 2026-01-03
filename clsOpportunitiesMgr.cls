VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunitiesMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'==============================================================
' Clase: clsOpportunitiesMgr
'--------------------------------------------------------------
' Gestiona la lista de "Oportunidades" (subcarpetas) de un
' directorio base configurado en el sistema. Expone métodos
' para refrescar, enumerar y cambiar oportunidad actual, así
' como un evento para notificar cambio de oportunidad actual.
'==============================================================
'TODO Implementar filtrado de oportunidades por año
'TODO Anadir búsqueda incremental en dropdown
'TODO Cachear resultado de actualizarColeccionOportunidades (5 min)

'@Exposed
'@Folder "4-Oportunidades y compresores"
'@IgnoreModule IIfSideEffect
Option Explicit

'--------------------------------------------------------------
' @Description: Evento que se dispara cuando cambia
' la oportunidad actual (el usuario en el desplegable
' del Ribbon, o por programa).
'--------------------------------------------------------------
Public Event currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
'--------------------------------------------------------------
' @Description: Evento que se dispara cuando se detectan cambios
' en la carpeta de oportunidades
'--------------------------------------------------------------
Public Event OpportunityCollectionUpdate(ByVal cambios As String)

'--------------------------------------------------------------
' Variables miembro
'--------------------------------------------------------------
Private strOportunitiesBaseFolder As String
Private p_ColOpportunities As Collection         ' Lista de subcarpetas encontradas
Private p_CurrOpportunity As Long                 ' Índice en p_ColOpportunities de la carpeta actual
Private p_bEnabled As Boolean

Private WithEvents ctx As clsExecutionContext
Attribute ctx.VB_VarHelpID = -1

'--------------------------------------------------------------
' @Description: Inicializa la clase y carga la ruta base de
' oportunidades desde el registro del sistema o su valor por
' defecto.
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo "clsOpportunitiesMgr", "[Class_Initialize]"
    
    Set ctx = New clsExecutionContext
    
    Set p_ColOpportunities = New Collection
    ' NO cargar rutas aquí, se hará después
    p_CurrOpportunity = -1
End Sub

Private Sub Class_Terminate()
    LogInfo "clsOpportunitiesMgr", "[Class_Terminate]"
End Sub

Public Property Get bEnabled() As Boolean
    bEnabled = p_bEnabled
End Property

' inyección explícita de la ruta (sin globals)
Public Sub SetBaseFolder(ByVal ruta As String)
    If RutaExiste(ruta) Then
        strOportunitiesBaseFolder = ruta
        p_bEnabled = True
    Else
        ' Opcional: lanzar error, o usar fallback
        LogWarning "clsOpportunitiesMgr", "[SetBaseFolder] ADVERTENCIA: Ruta no existe: " & ruta & " (deberían desactivarse los controles de ribbon relativos a oportunidades)"
        strOportunitiesBaseFolder = ""
    End If
    
    actualizarColeccionOportunidades
End Sub

'--------------------------------------------------------------
' @Description: Actualiza la colección p_ColOpportunities con
' las subcarpetas existentes en la ruta base configurada.
' El listado resultante se ordena en orden numérico inverso
' (de mayor a menor) según el número detectado en el nombre
' de cada carpeta. Si no se detectan números, se ordena en
' orden alfabético descendente.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function actualizarColeccionOportunidades()
    Dim fso As Object, carpeta As Object, subcarpeta As Object
    
    Set p_ColOpportunities = New Collection      ' limpia y reinicia

    If Not p_bEnabled Then Exit Function

    LogInfo "clsOpportunitiesMgr", "[actualizarColeccionOportunidades] - regenerando la variable local que contiene la lista de carpetas de oportunidades"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(strOportunitiesBaseFolder)
    On Error GoTo 0
    
    If Not carpeta Is Nothing Then
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        Dim arr() As String, i As Long
        ReDim arr(carpeta.SubFolders.Count)
        ' Copiamos rutas en un array para poder ordenarlas
        For Each subcarpeta In carpeta.SubFolders
            regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
            If Not regEx.Test(subcarpeta.Name) Then
                regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN
                If Not regEx.Test(subcarpeta.Name) Then
                    ' Nombre de carpeta mal implementado. Trata de corregirlo a partir de ofertas creadas
                End If
            End If
            regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
            If regEx.Test(subcarpeta.Name) Then
                arr(i) = subcarpeta.Name
                i = i + 1
            End If
        Next subcarpeta
        
        ' Orden numérico inverso
        If i > 0 Then
            ReDim Preserve arr(i - 1)
            arr = OrdenarCarpetasDesc(arr)
            For i = LBound(arr) To UBound(arr)
                p_ColOpportunities.Add arr(i)
            Next i
        End If
    End If
    
    p_CurrOpportunity = IIf(p_CurrOpportunity > 0, p_CurrOpportunity, IIf(p_ColOpportunities.Count > 0, 0, -1))
    actualizarColeccionOportunidades = arr
End Function

'--------------------------------------------------------------
' @Description: Procesa cambios detectados por el watcher EN CARPETAS DE OPORTUNIDAD:
'  Verifica que esos cambios realmente corresponden a cambios en oportunidades y lanza los procesos consiguientes
' @ArgumentDescriptions: cambios: nombres de carpetas separados por |
'--------------------------------------------------------------
Public Sub ProcesarCambiosEnOportunidades(ByVal subfolderName As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    Dim carpeta As Variant
    Dim bCambiosEnOportunidades As Boolean

    If Not p_bEnabled Then Exit Sub
    
    LogInfo "clsOpportunitiesMgr", "[ProcesarCambiosEnOportunidades] - identificando si los cambios introducidos por el watcher son oportunidades"
    ' Verificar si los cambios afectan a oportunidades válidas
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    bCambiosEnOportunidades = regEx.Test(subfolderName)
    
    ' Si el cambio corresponde a una oportunidad, actualizar
    If bCambiosEnOportunidades Then
        LogInfo "clsOpportunitiesMgr", "[ProcesarCambiosEnOportunidades] - cambios identificados"
        Call actualizarColeccionOportunidades
        ' NO actualizar ribbon directamente - delegar al evento
        ' La clase NO debe conocer el ribbon
        ' Disparar evento para que otros componentes reaccionen
        LogInfo "clsOpportunitiesMgr", "[ProcesarCambiosEnOportunidades] - promoviendo evento OpportunityCollectionUpdate para poner al corriente a otros módulos"
        RaiseEvent OpportunityCollectionUpdate(subfolderName)
    End If
End Sub

'TODO: PTE de implementar: separar del procesado de 'CARPETAS / OPORTUNIDADES'
' NO DEBE comunicarse con el File Manager... de eso se encarga clsAplicacion
Public Sub ProcesarCambiosEnItemsOportunidad(ByVal cambios As String)

End Sub

'--------------------------------------------------------------
' @Description: Devuelve el número de oportunidades cargadas.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: (sin argumentos)
'--------------------------------------------------------------
Public Function numOpportunities() As Variant
    numOpportunities = p_ColOpportunities.Count
End Function

'--------------------------------------------------------------
' @Description: Devuelve una oportunidad
' según el índice indicado.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Get OportunityLabel(Index As Integer) As String
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        OportunityLabel = p_ColOpportunities(Index + 1)
    Else
        OportunityLabel = "(Sin datos)"
    End If
End Property

'--------------------------------------------------------------
' @Description: Devuelve la ruta completa de una oportunidad
' según el índice indicado.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Get OportunityPath(Index As Long) As String
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        OportunityPath = strOportunitiesBaseFolder & "\" & p_ColOpportunities(Index + 1)
    End If
End Property

'--------------------------------------------------------------
' @Description: Actualiza la oportunidad actual y dispara
' el evento currOpportunityChanged.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: index: índice (base 0) de la oportunidad
'--------------------------------------------------------------
Public Property Let CurrOpportunity(Index As Long)
    If Index >= 0 And Index < p_ColOpportunities.Count Then
        p_CurrOpportunity = Index
        
        LogInfo "clsOpportunitiesMgr", "[CurrOpportunity] - cambio de oportunidad actual -> promoviendo evento OpportunityChanged para poner al corriente a otros módulos"
        RaiseEvent currOpportunityChanged(Index, p_ColOpportunities(Index + 1))
    End If
End Property

Public Property Get CurrOpportunity() As Long
    CurrOpportunity = p_CurrOpportunity
End Property

'--------------------------------------------------------------
' @Description: Ordena un array de rutas en orden numérico
' inverso, detectando el último número en el nombre de cada
' carpeta. Si no hay número, orden alfabético descendente.
'--------------------------------------------------------------
Private Function OrdenarCarpetasDesc(arr() As String) As String()
    Dim i As Long, j As Long, tmp As String
    Dim keyI As Double, keyJ As Double
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            keyI = ExtraerClaveNumerica(arr(i))
            keyJ = ExtraerClaveNumerica(arr(j))
            If keyI < keyJ Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            ElseIf keyI = keyJ Then
                If StrComp(arr(i), arr(j), vbTextCompare) < 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            End If
        Next j
    Next i
    OrdenarCarpetasDesc = arr
End Function

'--------------------------------------------------------------
' @Description: Extrae el último número del nombre de la carpeta.
' Si no hay número, devuelve -1E+99 para forzar orden al final.
'--------------------------------------------------------------
Private Function ExtraerClaveNumerica(ruta As String) As Double
    Dim re As Object, matches As Object, s As String
    Dim nombre As String
    On Error GoTo ErrHandler
    nombre = Dir(ruta)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    re.Global = True
    If re.Test(nombre) Then
        Set matches = re.Execute(nombre)
        s = matches(matches.Count - 1).Value
        ExtraerClaveNumerica = CDbl(s)
        Exit Function
    End If
ErrHandler:
    ExtraerClaveNumerica = -1E+99
End Function

'--------------------------------------------------------------
' @Description: Crea una nueva carpeta de oportunidad, a partir de la estructura de plantilla
'--------------------------------------------------------------
Public Sub CreaOportunidad()
    Dim fso As Object, carpeta As Object
    Dim arrOps() As String, i, strSAM As String, strCustomer As String, strOpName As String, dblLastOpSeq As Double
    
    
    If Not p_bEnabled Then Exit Sub
    
    arrOps = actualizarColeccionOportunidades
    dblLastOpSeq = 0
    For i = 0 To UBound(arrOps)
        arrOps(i) = Left(arrOps(i), 9)
        'If CDbl(Mid(arrOps(i), 7)) > dblLastOpSeq Then dblLastOpSeq = CDbl(Mid(arrOps(i), 7))
    Next
    If UBound(arrOps) >= 0 Then dblLastOpSeq = CDbl(Mid(arrOps(0), 7))
    strSAM = App.Configuration.SAM
    Do
        dblLastOpSeq = dblLastOpSeq + 1
        strOpName = strSAM & Mid(Year(Now), 3) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(3 - Len(CStr(dblLastOpSeq)), "0") & dblLastOpSeq
    Loop While UBound(Filter(arrOps, strOpName)) >= 0
    
    Dim regEx As Object, strMsg
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN
    Do
        strCustomer = InputBox(strMsg & "Introduce el nombre del cliente, o cliente - proyecto:", "Cliente - Proyecto", strCustomer)
        strMsg = "NOMBRE NO VALIDO. "
    Loop While strCustomer <> "" And Not regEx.Test(strOpName & " - " & strCustomer & " - XXX")
    
    If strCustomer <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Set carpeta = fso.GetFolder(strOportunitiesBaseFolder)
        On Error GoTo 0
        If carpeta Is Nothing Then GoTo ErrHandler
        
        Dim templFold As String
        templFold = App.Configuration.RutaPlantillas
        If Right(templFold, 1) = "\" Then templFold = Left(templFold, Len(templFold) - 1)
        fso.CopyFolder templFold, _
                       fso.BuildPath(carpeta.Path, strOpName & " - " & strCustomer & " - XXX")
    End If

    arrOps = actualizarColeccionOportunidades
    Exit Sub

ErrHandler:
    MsgBox "No se pudo abrir la ruta de oportunidades, " & strOportunitiesBaseFolder & vbCrLf & Err.Description, vbExclamation
End Sub

' ==================================================================
' IMPLEMENTACION DE EVENTOS: CALLBACKS
' ==================================================================

Private Sub ctx_WorkbookChanged(ByVal wb As Workbook)
    ' Verificar si el workbook está en una oportunidad
    ' Actualizar p_CurrOpportunity si corresponde
End Sub


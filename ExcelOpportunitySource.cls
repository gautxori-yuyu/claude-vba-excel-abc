VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelOpportunitySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' ADAPTADOR DE INFRAESTRUCTURA: ExcelOpportunitySource
' ==================================================================
' Implementa la interfaz IOpportunity usando Excel como tecnologia concreta.
' Encapsula clsFileXLS y traduce sus operaciones al contrato de dominio.
'
' El dominio solo ve IOpportunity; esta clase sabe hablar Excel.
' ==================================================================

'@Folder "2-Infraestructura.5-1-Adaptadores.Excel"
Option Explicit
Implements IOpportunity

Private Const MODULE_NAME As String = "ExcelOpportunitySource"

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private mFileXLS As clsFileXLS
Private mOpportunityId As String  ' Cache del ID extraido

' ==================================================================
' INICIALIZACION
' ==================================================================

'@Description: Inicializa el adaptador con un archivo Excel trackeado
'@Param f: Instancia de clsFileXLS que representa el archivo de oportunidad
Public Sub Initialize(f As clsFileXLS)
    If f Is Nothing Then
        Err.Raise 5, MODULE_NAME, "Se requiere una instancia valida de clsFileXLS"
    End If
    Set mFileXLS = f
    ' Extraer y cachear el ID de oportunidad
    mOpportunityId = ExtractOpportunityId()
    LogDebug MODULE_NAME, "[Initialize] Adaptador creado para: " & f.path
End Sub

Private Sub Class_Terminate()
    Set mFileXLS = Nothing
End Sub

' ==================================================================
' IMPLEMENTACION DE IOpportunity
' ==================================================================

Private Property Get IOpportunity_OpportunityId() As String
    IOpportunity_OpportunityId = mOpportunityId
End Property

Private Property Get IOpportunity_BasePath() As String
    If mFileXLS Is Nothing Then
        IOpportunity_BasePath = ""
    Else
        IOpportunity_BasePath = mFileXLS.path
    End If
End Property

Private Property Get IOpportunity_DisplayName() As String
    If mFileXLS Is Nothing Then
        IOpportunity_DisplayName = "(sin archivo)"
    Else
        IOpportunity_DisplayName = mFileXLS.Name
    End If
End Property

Private Function IOpportunity_IsValid() As Boolean
    If mFileXLS Is Nothing Then
        IOpportunity_IsValid = False
    Else
        IOpportunity_IsValid = mFileXLS.EsOportunidad()
    End If
End Function

Private Function IOpportunity_CanGenerateQuote() As Boolean
    ' Una oportunidad puede generar oferta si es valida y tiene datos tecnicos
    If Not IOpportunity_IsValid() Then
        IOpportunity_CanGenerateQuote = False
        Exit Function
    End If

    ' Verificar que el workbook esta disponible
    If mFileXLS.Wb Is Nothing Then
        IOpportunity_CanGenerateQuote = False
        Exit Function
    End If

    ' Por ahora, si es oportunidad valida, puede generar oferta
    ' TODO: Agregar validaciones adicionales segun reglas de negocio
    IOpportunity_CanGenerateQuote = True
End Function

Private Function IOpportunity_ReadTechnicalData() As Variant
    ' Devuelve la info del archivo como datos tecnicos
End Function

' ==================================================================
' PROPIEDADES PUBLICAS (acceso directo, opcional)
' ==================================================================

'@Description: Acceso directo al ID de oportunidad
Public Property Get OpportunityId() As String
    OpportunityId = mOpportunityId
End Property

'@Description: Acceso directo al path
Public Property Get BasePath() As String
    BasePath = IOpportunity_BasePath
End Property

'@Description: Acceso directo al archivo subyacente (solo para infraestructura)
'@Note: El dominio NO debe usar esta propiedad; usar la interfaz IOpportunity
Friend Property Get FileXLS() As clsFileXLS
    Set FileXLS = mFileXLS
End Property

' ==================================================================
' LOGICA PRIVADA
' ==================================================================

'@Description: Extrae el ID de oportunidad del nombre/path del archivo
'@Returns: String con el ID o cadena vacia si no se puede extraer
Private Function ExtractOpportunityId() As String
    On Error GoTo ErrHandler

    If mFileXLS Is Nothing Then
        ExtractOpportunityId = ""
        Exit Function
    End If

    Dim fileName As String
    fileName = mFileXLS.Name

    ' Patron tipico: el ID esta en el nombre del archivo
    ' Ejemplo: "OP-2024-001234_TechCalc.xlsx" -> "OP-2024-001234"
    ' Por ahora, usar el nombre completo sin extension como ID
    If InStr(fileName, ".") > 0 Then
        ExtractOpportunityId = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        ExtractOpportunityId = fileName
    End If

    Exit Function
ErrHandler:
    LogCurrentError MODULE_NAME, "[ExtractOpportunityId]"
    ExtractOpportunityId = ""
End Function

' ==================================================================
' FACTORY METHOD (opcional, para crear instancias facilmente)
' ==================================================================

'@Description: Crea una instancia de ExcelOpportunitySource a partir de un clsFileXLS
'@Param f: Archivo Excel a adaptar
'@Returns: Nueva instancia inicializada, o Nothing si f es invalido
Public Function CreateFrom(f As clsFileXLS) As ExcelOpportunitySource
    If f Is Nothing Then
        Set CreateFrom = Nothing
        Exit Function
    End If

    Dim adapter As New ExcelOpportunitySource
    adapter.Initialize f
    Set CreateFrom = adapter
End Function


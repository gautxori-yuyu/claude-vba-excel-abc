VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompressor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' ENTIDAD DE DOMINIO: clsCompressor
' ==================================================================
' Representa UN compresor dentro de una oportunidad.
' Un compresor agrupa:
' - Uno o mas calculos tecnicos
' - Una o mas valoraciones economicas
' - Opcionalmente, budgets y cotizaciones
'
' Segun el documento de arquitectura:
' - Una oportunidad puede tener 1:N compresores
' - Cada compresor tiene minimo 1 calculo y 1 valoracion
'
' El modelo del compresor se deriva del resultado de los calculos.
' ==================================================================

'@Folder "6-DOMINIO-Oportunidades y compresores"
Option Explicit

Private Const MODULE_NAME As String = "clsCompressor"

' ==================================================================
' ESTADO INTERNO
' ==================================================================
Private mModel As String                    ' Modelo del compresor (ej: "2TEH-1-LG")
Private mIndex As String                    ' Indice dentro de la oportunidad (ej: "01", "02")
Private mCalculations As Collection         ' Coleccion de ITechnicalCalc
Private mQuotes As Collection               ' Coleccion de IEconomicQuote
Private mBudgets As Collection              ' Coleccion de budgets (ofertas tipo presupuesto)
Private mQuotations As Collection           ' Coleccion de cotizaciones formales

' ==================================================================
' INICIALIZACION
' ==================================================================

Private Sub Class_Initialize()
    Set mCalculations = New Collection
    Set mQuotes = New Collection
    Set mBudgets = New Collection
    Set mQuotations = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCalculations = Nothing
    Set mQuotes = Nothing
    Set mBudgets = Nothing
    Set mQuotations = Nothing
End Sub

' ==================================================================
' PROPIEDADES DE IDENTIFICACION
' ==================================================================

'@Description: Modelo del compresor (ej: "2TEH-1-LG", "4TEHA-2-GGT")
Public Property Get Model() As String
    Model = mModel
End Property

Public Property Let Model(ByVal value As String)
    mModel = value
End Property

'@Description: Indice del compresor dentro de la oportunidad (ej: "01")
Public Property Get Index() As String
    Index = mIndex
End Property

Public Property Let Index(ByVal value As String)
    mIndex = value
End Property

'@Description: Identificador unico del compresor (Model + Index)
Public Property Get CompressorId() As String
    If Len(mModel) > 0 Then
        CompressorId = mModel & "_" & mIndex
    Else
        CompressorId = "COMPR_" & mIndex
    End If
End Property

' ==================================================================
' PROPIEDADES DE CALCULOS
' ==================================================================

'@Description: Numero de calculos tecnicos asociados
Public Property Get CalculationsCount() As Long
    CalculationsCount = mCalculations.Count
End Property

'@Description: Coleccion de calculos tecnicos
Public Property Get Calculations() As Collection
    Set Calculations = mCalculations
End Property

'@Description: Agrega un calculo tecnico al compresor
'@Returns: True si se agrego exitosamente
Public Function AddCalculation(calc As ITechnicalCalc) As Boolean
    On Error GoTo ErrHandler

    If calc Is Nothing Then Exit Function

    ' Verificar que no este duplicado (por CalcOption)
    Dim existing As ITechnicalCalc
    For Each existing In mCalculations
        If existing.CalcOption = calc.CalcOption Then
            LogWarning MODULE_NAME, "[AddCalculation] Calculo ya existe: " & calc.CalcOption
            Exit Function
        End If
    Next existing

    mCalculations.Add calc, calc.CalcOption

    ' Actualizar modelo si no esta definido
    If Len(mModel) = 0 And Len(calc.CompressorModel) > 0 Then
        mModel = calc.CompressorModel
    End If

    AddCalculation = True
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[AddCalculation]"
End Function

'@Description: Obtiene un calculo por su opcion
Public Function GetCalculation(ByVal calcOption As String) As ITechnicalCalc
    On Error Resume Next
    Set GetCalculation = mCalculations(calcOption)
    On Error GoTo 0
End Function

' ==================================================================
' PROPIEDADES DE VALORACIONES
' ==================================================================

'@Description: Numero de valoraciones economicas asociadas
Public Property Get QuotesCount() As Long
    QuotesCount = mQuotes.Count
End Property

'@Description: Coleccion de valoraciones economicas
Public Property Get Quotes() As Collection
    Set Quotes = mQuotes
End Property

'@Description: Agrega una valoracion economica al compresor
'@Returns: True si se agrego exitosamente
Public Function AddQuote(quote As IEconomicQuote) As Boolean
    On Error GoTo ErrHandler

    If quote Is Nothing Then Exit Function

    ' Verificar que no este duplicado (por QuoteId + Revision)
    Dim key As String
    key = quote.QuoteId & "_rev" & quote.Revision

    Dim existing As IEconomicQuote
    For Each existing In mQuotes
        If existing.QuoteId & "_rev" & existing.Revision = key Then
            LogWarning MODULE_NAME, "[AddQuote] Valoracion ya existe: " & key
            Exit Function
        End If
    Next existing

    mQuotes.Add quote, key
    AddQuote = True
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[AddQuote]"
End Function

'@Description: Obtiene la valoracion mas reciente
Public Function GetLatestQuote() As IEconomicQuote
    If mQuotes.Count = 0 Then Exit Function

    Dim latest As IEconomicQuote
    Dim quote As IEconomicQuote

    For Each quote In mQuotes
        If latest Is Nothing Then
            Set latest = quote
        ElseIf quote.Revision > latest.Revision Then
            Set latest = quote
        ElseIf quote.Revision = latest.Revision And quote.LastModified > latest.LastModified Then
            Set latest = quote
        End If
    Next quote

    Set GetLatestQuote = latest
End Function

' ==================================================================
' PROPIEDADES DE BUDGETS Y COTIZACIONES
' ==================================================================

'@Description: Numero de budgets
Public Property Get BudgetsCount() As Long
    BudgetsCount = mBudgets.Count
End Property

'@Description: Coleccion de budgets
Public Property Get Budgets() As Collection
    Set Budgets = mBudgets
End Property

'@Description: Numero de cotizaciones formales
Public Property Get QuotationsCount() As Long
    QuotationsCount = mQuotations.Count
End Property

'@Description: Coleccion de cotizaciones
Public Property Get Quotations() As Collection
    Set Quotations = mQuotations
End Property

' ==================================================================
' METODOS DE VALIDACION
' ==================================================================

'@Description: Indica si el compresor tiene datos minimos requeridos
Public Property Get IsValid() As Boolean
    ' Segun documento: minimo 1 calculo y 1 valoracion
    IsValid = (mCalculations.Count >= 1) And (mQuotes.Count >= 1)
End Property

'@Description: Indica si el compresor tiene todos los datos completos
Public Property Get IsComplete() As Boolean
    If Not IsValid Then Exit Property

    ' Verificar que al menos un calculo este procesado
    Dim calc As ITechnicalCalc
    Dim hasProcessedCalc As Boolean
    For Each calc In mCalculations
        If calc.IsProcessed Then
            hasProcessedCalc = True
            Exit For
        End If
    Next calc

    ' Verificar que al menos una valoracion sea valida
    Dim quote As IEconomicQuote
    Dim hasValidQuote As Boolean
    For Each quote In mQuotes
        If quote.IsValid Then
            hasValidQuote = True
            Exit For
        End If
    Next quote

    IsComplete = hasProcessedCalc And hasValidQuote
End Property

'@Description: Obtiene un resumen del estado del compresor
Public Function GetStatusSummary() As String
    Dim summary As String
    summary = "Compresor: " & CompressorId & vbCrLf
    summary = summary & "  Modelo: " & IIf(Len(mModel) > 0, mModel, "(pendiente)") & vbCrLf
    summary = summary & "  Calculos: " & mCalculations.Count & vbCrLf
    summary = summary & "  Valoraciones: " & mQuotes.Count & vbCrLf
    summary = summary & "  Estado: " & IIf(IsValid, "Valido", "Incompleto")
    GetStatusSummary = summary
End Function

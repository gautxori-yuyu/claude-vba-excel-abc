Attribute VB_Name = "modMACROLeerOfertas"
Option Explicit
Const RUTA_BD As String = "C:\Program Files (x86)\Ofertas_Gas\BaseDatos\Ofertas_Gas.mdb"

Function isGUID(ByVal strGUID)
Attribute isGUID.VB_Description = "[modMACROLeerOfertas] is GUID (función personalizada)"
Attribute isGUID.VB_ProcData.VB_Invoke_Func = " \n23"
  If IsNull(strGUID) Then
    isGUID = False
    Exit Function
  End If
  Dim regEx
  Set regEx = New RegExp
  regEx.Pattern = "[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}"
  isGUID = regEx.Test(strGUID)
  Set regEx = Nothing
End Function

'=========================================================
'@Description: Lee una oferta desde Access y vuelca sus datos generales en Excel
'@Scope: Prueba desde Excel
'@ArgumentDescriptions: -
'@Returns: Nothing
'@Category: Test
'=========================================================

Public Sub Test_LeerOfertas()

    Const OFER_ID As String = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"

    Dim ctx As clsDBContext
    Dim repo As clsOfertaRepository
    Dim of As clsOferta
    Dim dg As tOfertasDatosGenerales
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Oferta")
    ws.Cells.Clear

    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD

    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx

    Set of = repo.LeerPorOferID(OFER_ID)
    dg = of.DatosGenerales

    ' Cabeceras
    ws.Range("A1:F1").Value = Array( _
        "OFER_ID", "OFER_NUM_OFERTA", "OFER_FECHA", _
        "OFER_CLIENTE", "GASE_ID", "OFER_OBSERVACIONES")

    ' Datos
    ws.Range("A2").Value = dg.OFER_ID
    ws.Range("B2").Value = dg.OFER_NUM_OFERTA
    ws.Range("C2").Value = dg.OFER_FECHA
    ws.Range("D2").Value = dg.OFER_CLIENTE
    ws.Range("E2").Value = dg.GASE_ID
    ws.Range("F2").Value = dg.OFER_OBSERVACIONES

    ctx.Desconectar

    MsgBox "Oferta cargada correctamente", vbInformation
End Sub

'=========================================================
'@Description: Lee una oferta y vuelca la tabla OfertasOtros en Excel
'@Scope: Prueba desde Excel
'@ArgumentDescriptions: -
'@Returns: Nothing
'@Category: Test
'=========================================================

Public Sub Test_LeerOfertaConOtros()

    Const OFER_ID As String = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"

    Dim ctx As clsDBContext
    Dim repo As clsOfertaRepository
    Dim of As clsOferta
    Dim ws As Worksheet
    Dim i As Long
    Dim it As clsOfertaOtro

    Set ws = ThisWorkbook.Worksheets("OfertasOtros")
    ws.Cells.Clear

    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD

    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx

    Set of = repo.LeerPorOferID(OFER_ID)

    ' Cabeceras
    ws.Range("A1:C1").Value = Array( _
        "OFOT_LINEA", "OFOT_DESCRIPCION", "OFOT_PRE_COSTE")

    ' Datos
    For i = 1 To of.Otros.Count
        Set it = of.Otros(i)
        ws.Cells(i + 1, 1).Value = it.OFOT_LINEA
        ws.Cells(i + 1, 2).Value = it.OFOT_DESCRIPCION
        ws.Cells(i + 1, 3).Value = it.OFOT_PRE_COSTE
    Next i

    ctx.Desconectar

    MsgBox "OfertasOtros volcadas correctamente", vbInformation
End Sub

'@Description: Lee todas las ofertas desde Access y las vuelca masivamente en una hoja de Excel
'@Scope: Excel VBA ? Base de datos Access (lectura)
'@ArgumentDescriptions: -
'@Returns: Nothing
'@Category: Exportación / Ofertas

Public Sub Test_VolcarTodasLasOfertasAExcel()


    Dim ctx As clsDBContext
    Dim repo As clsOfertaRepository
    Dim ofertas As Collection
    Dim of As clsOferta
    Dim dg As tOfertasDatosGenerales
    Dim ws As Worksheet
    Dim fila As Long

    '----------------------------------
    ' Preparar hoja destino
    '----------------------------------
    Set ws = ActiveWorkbook.Worksheets("Ofertas")
    ws.Cells.Clear

    ws.Range("A1:F1").Value = Array( _
        "OFER_ID", _
        "OFER_NUM_OFERTA", _
        "OFER_FECHA", _
        "OFER_CLIENTE", _
        "GASE_ID", _
        "OFER_OBSERVACIONES")

    fila = 2

    '----------------------------------
    ' Conectar a base de datos
    '----------------------------------
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD

    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx

    '----------------------------------
    ' Leer repositorio completo
    '----------------------------------
    Set ofertas = repo.LeerTodas()

    '----------------------------------
    ' Volcado masivo
    '----------------------------------
    For Each of In ofertas
        dg = of.DatosGenerales

        ws.Cells(fila, 1).Value = dg.OFER_ID
        ws.Cells(fila, 2).Value = dg.OFER_NUM_OFERTA
        ws.Cells(fila, 3).Value = dg.OFER_FECHA
        ws.Cells(fila, 4).Value = dg.OFER_CLIENTE
        ws.Cells(fila, 5).Value = dg.GASE_ID
        ws.Cells(fila, 6).Value = dg.OFER_OBSERVACIONES

        fila = fila + 1
    Next of

    '----------------------------------
    ' Limpieza
    '----------------------------------
    ctx.Desconectar

    MsgBox ofertas.Count & " ofertas volcadas correctamente.", vbInformation

End Sub


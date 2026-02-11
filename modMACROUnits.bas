Attribute VB_Name = "modMACROUnits"
'==========================================
' LAS PRIMERAS TRES MACROS SON PARA GENERAR NOMBRES Y/O RANGOS DE HOJAS OCULTAS, QUE AYUDEN A DEFINIR
' LAS VALIDACIONES DE DATOS EN LAS CELDAS DE SELECCION DE UNIDADES...
' De momento la unica solucion aceptable es CREAR LOS RANGOS, definir los nombres ES INUTIL, porque Excel NO es capaz de usarlos
'==========================================

'@Folder "3-Dominio.b-Calculos técnicos"
Option Explicit

Private Const MODULE_NAME As String = "modMACROUnits"

'==========================================
' INSTALAR VALIDACIONES CON DEPENDENCIA DEL XLAM
' Crea nombres que apuntan directamente al XLAM usando FILTRO
'==========================================
' Instala nombres de unidades en el libro destino con referencias al XLAM
' Esto crea dependencia: el XLAM debe estar disponible
' Uso: Call InstalarValidacionesUnidades(ThisWorkbook)
Public Sub InstalarValidacionesUnidades(libroDestino As Workbook)
Attribute InstalarValidacionesUnidades.VB_ProcData.VB_Invoke_Func = " \n0"
    
    Dim tipos() As Variant
    Dim i As Long
    Dim nombreRango As String
    Dim formulaReferencia As String
    Dim nombreXLAM As String
    
    MsgBox ("ESTE METODO NO FUNCIONA, EXCEL NO RECONOCE LOS NOMBRES"): Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Obtener nombre del XLAM actual
    nombreXLAM = ThisWorkbook.Name
    
    ' Definir tipos de unidades
    tipos = Array("Presión", "Temp", "Masa", "Peso molecular", "Potencia", "Caudal", "Distancia")
    
    ' Crear nombres definidos con fórmulas FILTRO que apuntan al XLAM
    For i = LBound(tipos) To UBound(tipos)
        nombreRango = "Unidades_" & Replace(tipos(i), " ", "_")
        
        ' Construir fórmula FILTRO
        ' =FILTRO('[XLAM]Unidades'!$B:$B,'[XLAM]Unidades'!$A:$A="Tipo")
        formulaReferencia = "=FILTER('[" & nombreXLAM & "]Unidades'!$B:$B," & _
                            "'[" & nombreXLAM & "]Unidades'!$A:$A=""" & tipos(i) & """)"
        
        ' Eliminar nombre si existe
        On Error Resume Next
        libroDestino.Names(nombreRango).Delete
        On Error GoTo 0
        
        ' Crear nuevo nombre con referencia al XLAM
        libroDestino.Names.Add Name:=nombreRango, RefersTo:=formulaReferencia
    Next i
    
    MsgBox "Validaciones de unidades instaladas correctamente en " & libroDestino.Name & vbCrLf & _
           "NOTA: Estas validaciones dependen del complemento " & nombreXLAM, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al instalar validaciones: " & Err.Description & vbCrLf & _
           "NOTA: Requiere Excel 365 o Excel 2021+ con función FILTRO", vbCritical
End Sub

'==========================================
' INSTALAR VALIDACIONES LOCALES (SIN DEPENDENCIA)
' Crea hoja oculta y nombres locales en el libro destino
'==========================================
' Instala los nombres de unidades en el libro destino de forma local
' Crea una hoja oculta con las listas y nombres que apuntan a ella
' NO crea dependencia del XLAM
' Uso: Call InstalarValidacionesLocalesUnidades(ThisWorkbook)
Public Sub InstalarValidacionesLocalesUnidades(libroDestino As Workbook)
Attribute InstalarValidacionesLocalesUnidades.VB_ProcData.VB_Invoke_Func = " \n0"
    
    Dim wsDestino As Worksheet
    Dim tipos() As Variant
    Dim i As Long
    Dim listaUnidades As Variant
    Dim numElementos As Long
    Dim nombreRango As String
    Dim ultimaFila As Long
    
    On Error GoTo ErrorHandler
    
    ' Verificar/crear hoja oculta en el libro destino
    On Error Resume Next
    Set wsDestino = libroDestino.Sheets("_ListasUnidades")
    On Error GoTo 0
    
    If wsDestino Is Nothing Then
        Set wsDestino = libroDestino.Sheets.Add(After:=libroDestino.Sheets(libroDestino.Sheets.Count))
        wsDestino.Name = "_ListasUnidades"
    Else
        wsDestino.Cells.Clear
    End If
    
    ' Definir tipos de unidades
    tipos = Array("Presión", "Temp", "Masa", "Peso molecular", "Potencia", "Caudal", "Distancia")
    
    ' Generar listas en el libro destino
    For i = LBound(tipos) To UBound(tipos)
        wsDestino.Cells(1, i + 1).Value = tipos(i)
        wsDestino.Cells(1, i + 1).Font.Bold = True
        
        listaUnidades = UdsPorTipo(tipos(i))
        
        If Not IsError(listaUnidades) Then
            If IsArray(listaUnidades) Then
                numElementos = UBound(listaUnidades, 1) - LBound(listaUnidades, 1) + 1
                wsDestino.Cells(2, i + 1).Resize(numElementos, 1).Value = listaUnidades
            End If
        End If
    Next i
    
    If MsgBox("QUIERES TAMBIEN CREAR NOMBRES PARA LOS RANGOS DE LAS TABLAS? (NO FUNCIONA en la validacion de celdas," & _
              "EXCEL NO RECONOCE LOS NOMBRES", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    
    ' Crear nombres definidos en el libro destino (referencias locales)
    For i = LBound(tipos) To UBound(tipos)
        ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, i + 1).End(xlUp).Row
        
        If ultimaFila > 1 Then
            nombreRango = "Unidades_" & Replace(tipos(i), " ", "_")
            
            ' Eliminar si existe
            On Error Resume Next
            libroDestino.Names(nombreRango).Delete
            On Error GoTo 0
            
            ' Crear nuevo nombre en el libro destino con referencia local
            libroDestino.Names.Add Name:=nombreRango, _
                                   RefersTo:=wsDestino.Range(wsDestino.Cells(2, i + 1), wsDestino.Cells(ultimaFila, i + 1))
        End If
    Next i
    
    ' Ocultar hoja
    wsDestino.Visible = xlSheetVeryHidden
    
    MsgBox "Validaciones locales de unidades instaladas correctamente en " & libroDestino.Name & vbCrLf & _
           "Las listas son independientes del complemento.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al instalar validaciones locales: " & Err.Description, vbCritical
End Sub

Public Sub GenerarListasUnidades()
Attribute GenerarListasUnidades.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejecutar esto una vez para crear las listas en el XLAM
    
    Dim wsListas As Worksheet
    Dim tipos() As Variant
    Dim i As Long
    Dim listaUnidades As Variant
    
    MsgBox ("ESTE METODO NO FUNCIONA, EN LAS CELDAS DE VALIDACION NO SE PUEDEN USAR REFERENCIAS AL XLAM... " & _
            "Y EXCEL NO RECONOCE LOS NOMBRES"): Exit Sub
    
    On Error Resume Next
    Set wsListas = ThisWorkbook.Sheets("ListasUnidades")
    On Error GoTo 0
    
    If wsListas Is Nothing Then
        Set wsListas = ThisWorkbook.Sheets.Add
        wsListas.Name = "ListasUnidades"
    Else
        wsListas.Cells.Clear
    End If
    
    ' Definir tipos de unidades
    tipos = Array("Presión", "Temp", "Masa", "Peso molecular", "Potencia", "Caudal", "Distancia")
    
    For i = LBound(tipos) To UBound(tipos)
        wsListas.Cells(1, i + 1).Value = tipos(i)
        wsListas.Cells(1, i + 1).Font.Bold = True
        
        listaUnidades = UdsPorTipo(tipos(i))
        
        If Not IsError(listaUnidades) Then
            If IsArray(listaUnidades) Then
                wsListas.Cells(2, i + 1).Resize(UBound(listaUnidades), 1).Value = listaUnidades
            End If
        End If
    Next i
    
    ' Crear nombres definidos para cada tipo
    Dim ultimaFila As Long
    Dim nombreRango As String
    
    For i = LBound(tipos) To UBound(tipos)
        ' Encontrar límite del rango en esa columna
        ultimaFila = wsListas.Cells(wsListas.Rows.Count, i + 1).End(xlUp).Row
        
        ' Solo crear nombre si hay datos (más de la fila de encabezado)
        If ultimaFila > 1 Then
            ' Crear nombre definido
            nombreRango = "Unidades_" & Replace(tipos(i), " ", "_")
            On Error Resume Next
            ThisWorkbook.Names(nombreRango).Delete
            On Error GoTo 0
            
            ThisWorkbook.Names.Add Name:=nombreRango, _
                                   RefersTo:=wsListas.Range(wsListas.Cells(2, i + 1), wsListas.Cells(ultimaFila, i + 1))
        End If
    Next i
    
    wsListas.Visible = xlSheetVeryHidden
    LogInfo MODULE_NAME, "[GenerarListasUnidades] Nombres de listas de unidades para validación de celdas generados correctamente"
End Sub



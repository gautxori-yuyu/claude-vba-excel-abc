Attribute VB_Name = "modMACROComparadorHojas"
' ==============================================================================================================
' MÓDULO: modMACROComparadorHojas
' DESCRIPCIÓN: Módulo para comparar dos hojas de Excel mostrándolas en paralelo y resaltando las diferencias
'              encontradas entre los rangos seleccionados. Incluye funcionalidad para deshacer la comparación
'              y restaurar los colores originales de las celdas.
' ==============================================================================================================
'TODO: QUE Al hacer la comparación **permita seleccionar una columna de los rangos de comparación**, y en tal caso
' las comparaciones no se hagan "línea a línea" sino identificando Los pares de líneas a comparar Por los elementos iguales En esa columna
' (Obviamente en tal caso para mostrar los resultados de la comparación Habría que generar Una nueva hoja Que permita emparejar las fila;
' Esa nueva hoja, La que establece el orden de las hilas en comparación, Será una copia de aquella en la que no se haya seleccionado
' "Marcar las diferencias con un color" )

'@Folder "MACROS"
Option Explicit

Private Const MODULE_NAME As String = "modMACROComparadorHojas"

' Diccionario para guardar: clave = dirección celda, valor = color original
' Si una celda está en este diccionario, significa que fue modificada
Private dictCeldasModificadas As Object

' Variable a nivel de modulo para mantener referencia al formulario no modal
Private mFrmComparador As frmComparadorHojas

' -------------------------------------------------------------------------------------------------------------
' INICIALIZACIÓN Y CONFIGURACIÓN
' -------------------------------------------------------------------------------------------------------------

'@Description: Inicializa el diccionario de celdas modificadas y muestra el formulario comparador
'              en modo no modal
'@Scope: Public
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmComparadorHojas (formulario)
'@Note: El diccionario se usa para rastrear las celdas cuyo color fue modificado durante la comparación
'        Usa instanciacion explicita para evitar problemas de memoria con la instancia predeterminada.
'        La referencia se mantiene a nivel de modulo para formularios no modales.
Sub MostrarComparador()
Attribute MostrarComparador.VB_ProcData.VB_Invoke_Func = " \n0"
    Set dictCeldasModificadas = CreateObject("Scripting.Dictionary")

    ' Si ya hay un formulario abierto, traerlo al frente
    If Not mFrmComparador Is Nothing Then
        On Error Resume Next
        mFrmComparador.Show
        If Err.Number = 0 Then Exit Sub
        On Error GoTo 0
    End If

    ' Crear nueva instancia
    Set mFrmComparador = New frmComparadorHojas
    mFrmComparador.Show vbModeless
End Sub
' -------------------------------------------------------------------------------------------------------------
' FUNCIONES DE COMPARACIÓN Y VISUALIZACIÓN
' -------------------------------------------------------------------------------------------------------------

'@Description: Compara dos rangos celda por celda y colorea las diferencias encontradas según los
'              parámetros especificados
'@Scope: Public
'@ArgumentDescriptions: rango1 (Range): Primer rango a comparar
'   | colorear1 (Boolean): Indica si se deben colorear las diferencias en el rango1
'   | rango2 (Range): Segundo rango a comparar
'   | colorear2 (Boolean): Indica si se deben colorear las diferencias en el rango2
'   | colorDiferencia (Long): Color RGB a aplicar en las celdas con diferencias
'   | soloFondoBlanco (Boolean): Si es True, solo colorea celdas con fondo blanco o sin color
'@Returns: (ninguno)
'@Dependencies: EsFondoBlanco, GuardarYModificar
'@Note: Compara solo la intersección de ambos rangos (menor dimensión). Los colores originales
'        se guardan en dictCeldasModificadas para poder deshacerlos posteriormente.
Sub CompararRangos(rango1 As Range, colorear1 As Boolean, _
                   rango2 As Range, colorear2 As Boolean, _
                   colorDiferencia As Long, soloFondoBlanco As Boolean)
Attribute CompararRangos.VB_ProcData.VB_Invoke_Func = " \n0"
    
    Dim fila As Long, col As Long
    Dim maxFilas As Long, maxCols As Long
    Dim celda1 As Range, celda2 As Range
    Dim valor1 As Variant, valor2 As Variant
    
    ' Limpiar diccionario previo
    dictCeldasModificadas.RemoveAll
    
    ' Calcular la intersección (menor dimensión de ambos rangos)
    maxFilas = Application.Min(rango1.Rows.Count, rango2.Rows.Count)
    maxCols = Application.Min(rango1.Columns.Count, rango2.Columns.Count)
    
    ' Recorrer la intersección
    For fila = 1 To maxFilas
        For col = 1 To maxCols
            Set celda1 = rango1.Cells(fila, col)
            Set celda2 = rango2.Cells(fila, col)
            
            ' Obtener valores
            valor1 = celda1.Value
            valor2 = celda2.Value
            
            ' Si son diferentes
            If valor1 <> valor2 Then
                ' Procesar Hoja 1
                If colorear1 Then
                    If soloFondoBlanco Then
                        ' Solo colorear si el fondo es blanco o sin color
                        If EsFondoBlanco(celda1) Then
                            GuardarYModificar celda1, colorDiferencia
                        End If
                    Else
                        GuardarYModificar celda1, colorDiferencia
                    End If
                End If
                
                ' Procesar Hoja 2
                If colorear2 Then
                    If soloFondoBlanco Then
                        ' Solo colorear si el fondo es blanco o sin color
                        If EsFondoBlanco(celda2) Then
                            GuardarYModificar celda2, colorDiferencia
                        End If
                    Else
                        GuardarYModificar celda2, colorDiferencia
                    End If
                End If
            End If
        Next col
    Next fila
    
End Sub

'@Description: Organiza dos hojas en vista paralela sincronizada, permitiendo comparar visualmente
'              los rangos especificados lado a lado
'@Scope: Public
'@ArgumentDescriptions: Hoja1 (Worksheet): Primera hoja a visualizar
'   | Hoja2 (Worksheet): Segunda hoja a visualizar
'   | rango1 (Range): Rango a visualizar en la Hoja1
'   | rango2 (Range): Rango a visualizar en la Hoja2
'@Returns: (ninguno)
'@Dependencies: Windows.CompareSideBySideWith, Sleep (API)
'@Note: Si ambas hojas pertenecen al mismo libro, crea una nueva ventana. Activa el desplazamiento
'        sincronizado para facilitar la comparación visual.
Sub VerHojasEnParalelo(Hoja1 As Worksheet, Hoja2 As Worksheet, rango1 As Range, rango2 As Range)
Attribute VerHojasEnParalelo.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim win1 As Window, win2 As Window
    Dim yaSincronizado As Boolean
    
    ' Activar primera hoja y crear/obtener su ventana
    Hoja1.Activate
    rango1.Cells(1, 1).Select                    ' Seleccionar primera celda del rango
    Application.GoTo rango1.Cells(1, 1), True    ' Asegurar que está visible
    Set win1 = ActiveWindow
    
    ' Si son del mismo libro, crear segunda ventana
    If Hoja1.Parent.Name = Hoja2.Parent.Name Then
        Hoja2.Parent.NewWindow
    End If
    Hoja2.Activate
    rango2.Cells(1, 1).Select                    ' Seleccionar primera celda del rango
    Application.GoTo rango2.Cells(1, 1), True    ' Asegurar que está visible
    Set win2 = ActiveWindow
    
    ' Asegurarse de que hay ventanas disponibles
    If Application.Windows.Count < 2 Then Exit Sub
    
    win1.Activate
    ActiveWindow.View = xlNormalView
    ' Organizar ventanas
    'Application.Windows.Arrange xlArrangeStyleVertical, True, True, True
    ' Activar vista en paralelo
    On Error Resume Next
    Windows.CompareSideBySideWith win2.Caption
    Windows.ResetPositionsSideBySide
    
    ' Verificar si ya está en modo comparación
    yaSincronizado = Windows.SyncScrollingSideBySide
    ' Solo activar si no está ya sincronizado
    If Not yaSincronizado Then
        'Windows.SyncScrollingSideBySide = True
        Sleep 200
        Application.CommandBars.ExecuteMso "SynchronousScrolling" ' Desplazamiento sincronizado
    End If
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES DE FORMATO Y COLOR
' -------------------------------------------------------------------------------------------------------------

'@Description: Verifica si una celda tiene fondo blanco o sin color (xlNone)
'@Scope: Public
'@ArgumentDescriptions: celda (Range): Celda a verificar
'@Returns: Boolean - True si el fondo es blanco (RGB 255,255,255) o sin color (xlNone), False en caso contrario
'@Dependencies: Ninguna
'@Note: xlNone tiene valor -4142. RGB blanco tiene valor 16777215
Function EsFondoBlanco(celda As Range) As Boolean
Attribute EsFondoBlanco.VB_Description = "[modMACROComparadorHojas] FUNCIONES AUXILIARES DE FORMATO Y COLOR Verifica si una celda tiene fondo blanco o sin color (xlNone) xlNone tiene valor -4142. RGB blanco tiene valor 16777215. Aplica a: Cells Range\r\nM.D.:Public"
Attribute EsFondoBlanco.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim colorInterior As Long
    colorInterior = celda.Interior.Color
    
    ' xlNone = -4142 (sin color)
    ' RGB(255, 255, 255) = 16777215 (blanco)
    EsFondoBlanco = (colorInterior = xlNone) Or (colorInterior = RGB(255, 255, 255))
End Function

'@Description: Guarda el color original de una celda en el diccionario y aplica un nuevo color
'@Scope: Public
'@ArgumentDescriptions: celda (Range): Celda a modificar
'   | nuevoColor (Long): Color RGB a aplicar
'@Returns: (ninguno)
'@Dependencies: dictCeldasModificadas (variable privada del módulo)
'@Note: Solo guarda el color si la celda no existe previamente en el diccionario, para preservar
'        el color original inicial
Sub GuardarYModificar(celda As Range, nuevoColor As Long)
Attribute GuardarYModificar.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim clave As String
    clave = celda.Address(External:=True)
    
    ' Solo guardar si no está ya en el diccionario
    If Not dictCeldasModificadas.Exists(clave) Then
        dictCeldasModificadas.Add clave, celda.Interior.Color
    End If
    
    ' Aplicar nuevo color
    celda.Interior.Color = nuevoColor
End Sub

' -------------------------------------------------------------------------------------------------------------
' FUNCIONES DE RESTAURACIÓN Y CONSULTA DE ESTADO
' -------------------------------------------------------------------------------------------------------------

'@Description: Deshace la última comparación realizada, restaurando los colores originales de todas
'              las celdas modificadas y cierra la vista en paralelo
'@Scope: Public
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: dictCeldasModificadas (variable privada del módulo)
'@Note: Si no hay comparación activa muestra un mensaje informativo. Limpia el diccionario al finalizar
Sub DeshacerComparacion()
Attribute DeshacerComparacion.VB_ProcData.VB_Invoke_Func = " \n0"
    
    Dim clave As Variant
    Dim colorOriginal As Long
    
    If dictCeldasModificadas.Count = 0 Then
        MsgBox "No hay comparación activa para deshacer.", vbInformation
        Exit Sub
    End If
    
    ' Recorrer SOLO las celdas que fueron modificadas
    For Each clave In dictCeldasModificadas.Keys
        colorOriginal = dictCeldasModificadas(clave)
        
        ' Restaurar color original
        On Error Resume Next
        ActiveSheet.Range(clave).Interior.Color = colorOriginal
        On Error GoTo 0
    Next
    
    ' Limpiar diccionario
    dictCeldasModificadas.RemoveAll
    
    Windows.BreakSideBySide
End Sub

'@Description: Verifica si existe una comparación activa en este momento
'@Scope: Public
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Boolean - True si hay celdas modificadas pendientes de restaurar, False en caso contrario
'@Dependencies: dictCeldasModificadas (variable privada del módulo)
'@Note: Útil para deshabilitar/habilitar botones en la interfaz según el estado de comparación
Public Function HayComparacionActiva() As Boolean
Attribute HayComparacionActiva.VB_Description = "[modMACROComparadorHojas] Verifica si existe una comparación activa en este momento Útil para deshabilitar/habilitar botones en la interfaz según el estado de comparación"
Attribute HayComparacionActiva.VB_ProcData.VB_Invoke_Func = " \n23"
    HayComparacionActiva = (dictCeldasModificadas.Count > 0)
End Function

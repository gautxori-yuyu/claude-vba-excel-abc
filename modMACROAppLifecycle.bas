Attribute VB_Name = "modMACROAppLifecycle"
Option Private Module
' ==========================================
' CICLO DE VIDA DE LA APLICACION
' ==========================================
' Funciones publicas para gestion de la aplicacion y Ribbon
' ==========================================

'@Folder "1-Aplicacion.3-Ciclo de vida"
Option Explicit

Private Const MODULE_NAME As String = "modMACROAppLifecycle"

' ==========================================
' RIBBON RECOVERY - ALMACENAMIENTO EN NOMBRES EXCEL4
' ==========================================
' Los nombres de Excel4 persisten mientras Excel este abierto,
' incluso tras un reset de VBA (a diferencia de variables de modulo).
' Esta es la ubicacion "mas estatica" disponible dentro del proceso Excel.
' ==========================================

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Private Const RIBBON_PTR_EXCEL_NAME As String = "ABC_RibbonPtr"

' RECUPERACION DE LA APLICACION
Private mInitCounter As Long       ' Contador de inicializaciones (persistente en sesion)
Private mLastInitTime As Double    ' Timestamp de ultima inicializacion

' ==========================================
' ACCESO A LA APLICACION
' ==========================================

Public Function App() As clsApplication
Attribute App.VB_ProcData.VB_Invoke_Func = " \n0"
    Set App = ThisWorkbook.App
End Function

' ================================================
' RECUPERACION DE LA APLICACION
' DETECCION DE RESET Y CONTROL INICIALIZACIONES
' ================================================
' Cuando VBA se resetea (error fatal, End en depuracion, etc.),
' todas las variables de modulo se reinicializan a 0/Nothing.
' Usamos una variable Static dentro de una funcion para detectar esto.
' ================================================

'@Description: Fuerza el reinicio completo de la aplicacion
Public Sub ReiniciarAplicacion()
Attribute ReiniciarAplicacion.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim result As TDRESULT

    result = ShowTaskDialogYesNo("Reiniciar Aplicacion", _
                                 "Reiniciar el complemento ABC?", _
                                 "Se cerrara y volvera a inicializar la aplicacion.")

    If result <> vbYes Then Exit Sub

    LogInfo MODULE_NAME, "[ReiniciarAplicacion] Reinicio solicitado por usuario"

    On Error Resume Next

    ' Terminar aplicacion actual
    ThisWorkbook.TerminateApp

    ' Reinicializar
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    ' Forzar reinicio llamando a App()
    Dim dummy As clsApplication
    Set dummy = App()

    On Error GoTo 0

    ' Verificar estado
    If IsRibbonAvailable() Then
        ShowTaskDialogError "Reinicio Exitoso", _
                            "Aplicacion reiniciada correctamente", _
                            GetRibbonDiagnostics()
    Else
        ShowTaskDialogError "Reinicio Parcial", _
                            "Aplicacion reiniciada con advertencias", _
                            "El Ribbon puede requerir atencion. Ejecute 'RecuperarRibbon' si es necesario."
    End If
End Sub

'@Description: Detecta si ocurrio un reset de VBA desde la ultima llamada
'@Returns: True si es la primera llamada tras un reset (contador > 1)
'@Note: Esta funcion usa una variable Static que sobrevive entre llamadas
'       pero se reinicia si VBA hace reset. El patron detecta ese reset.
Public Function DetectVBAResetOccurred() As Boolean
Attribute DetectVBAResetOccurred.VB_ProcData.VB_Invoke_Func = " \n0"
    Static sInitFlag As Boolean

    If Not sInitFlag Then
        ' Primera vez que se ejecuta desde reset
        sInitFlag = True
        mInitCounter = mInitCounter + 1
        mLastInitTime = Timer

        ' Si contador > 1, hubo reset previo
        DetectVBAResetOccurred = (mInitCounter > 1)

        If DetectVBAResetOccurred Then
            LogWarning MODULE_NAME, "[DetectVBAResetOccurred] Reset detectado! Inicializacion #" & mInitCounter
        Else
            LogInfo MODULE_NAME, "[DetectVBAResetOccurred] Primera inicializacion de la sesion"
        End If
    Else
        ' Llamadas subsiguientes en la misma sesion - no hay reset
        DetectVBAResetOccurred = False
    End If
End Function

'@Description: Obtiene el numero de veces que se ha inicializado la aplicacion
'@Returns: Long - Contador de inicializaciones (1 = primera vez, >1 = hubo resets)
Public Property Get InitializationCount() As Long
    InitializationCount = mInitCounter
End Property

'@Description: Obtiene el timestamp de la ultima inicializacion
'@Returns: Double - Valor de Timer en la ultima inicializacion
Public Property Get LastInitializationTime() As Double
    LastInitializationTime = mLastInitTime
End Property

' ==========================================
' GESTION DEL COMPLEMENTO XLAM
' MACROS PUBLICAS (Accesibles por el usuario)
' ==========================================

'@Description: Activa temporalmente la visibilidad del XLAM para operaciones de copia
'@Scope: Manipula el libro host del complemento XLAM cargado.
'@Category: ComplementosExcel
Sub DesactivarModoAddin()
Attribute DesactivarModoAddin.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        LogInfo MODULE_NAME, "[DesactivarModoAddin] XLAM visible temporalmente"
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modMACROAppLifecycle.DesactivarModoAddin", _
              "Error desactivando el modo de AddIn: " & Err.Description
End Sub

'@Description: Restaura el estado de IsAddin del XLAM
'@Category: ComplementosExcel
Sub RestaurarModoAddin()
Attribute RestaurarModoAddin.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    ThisWorkbook.IsAddin = True
    LogInfo MODULE_NAME, "[RestaurarModoAddin] XLAM restaurado como Add-in"
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modMACROAppLifecycle.RestaurarModoAddin", _
              "Error activando el modo de AddIn: " & Err.Description
End Sub

' ==========================================
' GESTION DEL RIBBON
' ==========================================

'@Description: Procedimiento puente para el atajo de teclado
Public Sub ToggleRibbonTab()
Attribute ToggleRibbonTab.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    If App.ribbon Is Nothing Then
        LogDebug MODULE_NAME, "[ToggleRibbonTab] Ribbon no disponible"
        Exit Sub
    End If
    App.ribbon.ToggleModo

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[ToggleRibbonTab]"
    Err.Raise Err.Number, MODULE_NAME & "[ToggleRibbonTab]", _
              "Error cambiando el modo del ribbon: " & Err.Description
End Sub

' ------------------------------------------
' ALMACENAMIENTO DEL PUNTERO - SET.NAME (Excel4)
' ------------------------------------------

'@Description: Almacena el puntero IRibbonUI en nombres de Excel4.
'              Llamar desde RibbonOnLoad para garantizar persistencia tras resets de VBA.
'@Note: Los nombres Excel4 sobreviven a resets de VBA porque viven en el
'       espacio de nombre del proceso Excel, no del proyecto VBA.
Public Sub StoreRibbonInExcelNames(ByRef ribbon As IRibbonUI)
    On Error GoTo ErrExit
    If ribbon Is Nothing Then Exit Sub

    Dim ptr As LongPtr
    ptr = ObjPtr(ribbon)
    Application.ExecuteExcel4Macro "SET.NAME(""" & RIBBON_PTR_EXCEL_NAME & """, """ & CStr(ptr) & """)"
    LogDebug MODULE_NAME, "[StoreRibbonInExcelNames] puntero guardado en nombre Excel4: " & CStr(ptr)
    Exit Sub

ErrExit:
    LogCurrentError MODULE_NAME, "[StoreRibbonInExcelNames]"
End Sub

'@Description: Recupera IRibbonUI desde nombres de Excel4.
'              Funciona incluso despues de un reset de VBA porque el nombre persiste.
'@Note: Si el puntero almacenado esta obsoleto (Excel invalido el ribbon), la llamada
'       fallara al acceder al objeto. El On Error del caller debe capturar esto.
'@Returns: IRibbonUI recuperado, o Nothing si no hay puntero valido almacenado
Public Function RecoverRibbonFromExcelNames() As IRibbonUI
    On Error GoTo ErrExit

    Dim ptrStr As String
    ptrStr = Application.ExecuteExcel4Macro(RIBBON_PTR_EXCEL_NAME)

    ' "False" es lo que devuelve ExecuteExcel4Macro cuando el nombre no existe
    If Len(ptrStr) = 0 Or ptrStr = "0" Or ptrStr = "False" Then Exit Function

    Dim ptr As LongPtr
    ptr = CLngPtr(ptrStr)
    If ptr = 0 Then Exit Function

    ' Reconstruir la referencia COM desde la direccion de memoria
    Dim obj As IRibbonUI
    CopyMemory obj, ptr, LenB(ptr)

    Set RecoverRibbonFromExcelNames = obj

    ' Prevenir doble-liberacion del refcount COM (es critico dejar el variable local a 0)
    Dim ptrZero As LongPtr
    CopyMemory obj, ptrZero, LenB(ptrZero)

    LogDebug MODULE_NAME, "[RecoverRibbonFromExcelNames] referencia reconstruida desde ptr=" & CStr(ptr)
    Exit Function

ErrExit:
    ' El puntero puede estar obsoleto: fallo silencioso, devolver Nothing
    Set RecoverRibbonFromExcelNames = Nothing
End Function

' ------------------------------------------
' RECUPERACION DEL RIBBON
' ------------------------------------------

'@Description: Macro publica para recuperar el Ribbon manualmente.
'@Note: Ejecutar si el Ribbon desaparece o no responde.
'@Category: Ribbon / Recuperacion
Public Sub RecuperarRibbon()
Attribute RecuperarRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[RecuperarRibbon] Solicitado por usuario"
    LogDebug MODULE_NAME, "[RecuperarRibbon] Diagnosticos: " & GetRibbonDiagnostics()

    ' Si ya esta disponible, no hacer nada
    If IsRibbonAvailable() Then
        ShowTaskDialogError "Ribbon OK", "Estado correcto", "El Ribbon ya esta funcionando correctamente."
        Exit Sub
    End If

    ' PASO 1: Intentar recuperar desde nombres Excel4 (rapido, no disruptivo)
    If TryRecoverRibbon() Then
        ShowTaskDialogError "Recuperacion Exitosa", "Ribbon recuperado", GetRibbonDiagnostics()
        Exit Sub
    End If

    ' PASO 2 (DESHABILITADO): Toggle del add-in
    ' Al deshabilitar el XLAM se pierde el control del programa y se produce un reinicio
    ' de la aplicacion. Este mecanismo es mas destructivo que util.
    ' Mantener como referencia pero NO ejecutar automaticamente.
    '
    ' Para recuperacion manual si el PASO 1 falla: cerrar y reabrir Excel.
    ShowTaskDialogError "Recuperacion Fallida", "Ribbon no recuperado", _
                        "La recuperacion automatica no tuvo exito." & vbCrLf & _
                        "Recomendaciones:" & vbCrLf & _
                        "1. Cierre Excel completamente" & vbCrLf & _
                        "2. Vuelva a abrir Excel"
End Sub

'@Description: Muestra el diagnostico del Ribbon en un cuadro de dialogo
Public Sub MostrarDiagnosticoRibbon()
Attribute MostrarDiagnosticoRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[MostrarDiagnosticoRibbon] Solicitado"
    ShowTaskDialogError "Diagnostico del Ribbon", "Estado actual del ribbon", GetRibbonDiagnostics()
End Sub

' ------------------------------------------
' FUNCIONES DE DIAGNOSTICO
' ------------------------------------------

'@Description: Obtiene informacion de diagnostico del estado del Ribbon
'@Returns: String - Descripcion del estado actual
Public Function GetRibbonDiagnostics() As String
Attribute GetRibbonDiagnostics.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim info As String
    info = "Fecha/Hora: " & Now & vbCrLf
    info = info & "Log Path: " & GetLogFilePath() & vbCrLf & vbCrLf

    ' Estado de App
    Dim mApp As clsApplication
    Set mApp = App
    If mApp Is Nothing Then
        info = info & "[X] App: Nothing (CRITICO)" & vbCrLf
        GetRibbonDiagnostics = info
        Exit Function
    Else
        info = info & "[OK] App: Disponible" & vbCrLf
    End If

    ' Estado de Ribbon (clsRibbon)
    If mApp.ribbon Is Nothing Then
        info = info & "[X] App.Ribbon: Nothing (ERROR)" & vbCrLf
    Else
        info = info & "[OK] App.Ribbon: Disponible" & vbCrLf

        ' Estado de ribbonUI (IRibbonUI)
        On Error Resume Next
        If mApp.ribbon.ribbonUI Is Nothing Then
            info = info & "[X] ribbonUI: Nothing (PERDIDO)" & vbCrLf
            info = info & "    -> puntero en Excel4: " & SafeGetExcelName() & vbCrLf
        Else
            info = info & "[OK] ribbonUI: Conectado (" & TypeName(mApp.ribbon.ribbonUI) & ")" & vbCrLf
        End If
        On Error GoTo 0
    End If

    ' Estado de Ribbon.State
    On Error Resume Next
    If mApp.ribbon Is Nothing Then
        info = info & "[X] Ribbon State: sin ribbon" & vbCrLf
    ElseIf mApp.ribbon.State Is Nothing Then
        info = info & "[X] Ribbon State: Nothing" & vbCrLf
    Else
        info = info & "[OK] Ribbon State: " & mApp.ribbon.State.Description & vbCrLf
    End If
    On Error GoTo 0

    ' Estado del nombre Excel4
    info = info & "Excel4 Name: " & SafeGetExcelName() & vbCrLf

    GetRibbonDiagnostics = info
End Function

'@Description: Verifica si el Ribbon esta disponible y funcional
'@Returns: Boolean - True si el Ribbon esta operativo
Public Function IsRibbonAvailable() As Boolean
Attribute IsRibbonAvailable.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next

    Dim mApp As clsApplication
    Set mApp = App
    If mApp Is Nothing Then IsRibbonAvailable = False: Exit Function
    If mApp.ribbon Is Nothing Then IsRibbonAvailable = False: Exit Function
    If mApp.ribbon.ribbonUI Is Nothing Then IsRibbonAvailable = False: Exit Function

    IsRibbonAvailable = (TypeName(mApp.ribbon.ribbonUI) <> "Nothing")
    If Err.Number <> 0 Then
        IsRibbonAvailable = False
        Err.Clear
    End If

    On Error GoTo 0
End Function

' ------------------------------------------
' FUNCIONES DE RECUPERACION (privadas)
' ------------------------------------------

'@Description: Intenta recuperar ribbon desde nombres Excel4 (no disruptivo).
'@Returns: Boolean - True si la recuperacion fue exitosa
Public Function TryRecoverRibbon() As Boolean
Attribute TryRecoverRibbon.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[TryRecoverRibbon] Iniciando recuperacion"

    ' Recuperar puntero desde nombres Excel4 (sobrevive resets de VBA)
    Dim recoveredRibbon As IRibbonUI
    Set recoveredRibbon = RecoverRibbonFromExcelNames()

    If recoveredRibbon Is Nothing Then
        LogWarning MODULE_NAME, "[TryRecoverRibbon] No hay puntero valido en nombres Excel4"
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Validar que el objeto es usable (no solo que el puntero no sea 0)
    On Error Resume Next
    Dim testType As String
    testType = TypeName(recoveredRibbon)
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo ErrHandler

    If errNum <> 0 Or Len(testType) = 0 Or testType = "Nothing" Then
        LogWarning MODULE_NAME, "[TryRecoverRibbon] Puntero obsoleto (ribbon invalido): err=" & errNum
        TryRecoverRibbon = False
        Exit Function
    End If

    ' Actualizar la referencia en clsRibbon
    If Not App.ribbon Is Nothing Then
        App.ribbon.Init recoveredRibbon
        LogInfo MODULE_NAME, "[TryRecoverRibbon] Ribbon recuperado correctamente desde nombres Excel4"
        TryRecoverRibbon = True
        Exit Function
    End If

    LogWarning MODULE_NAME, "[TryRecoverRibbon] App.ribbon no disponible"
    TryRecoverRibbon = False
    Exit Function

ErrHandler:
    TryRecoverRibbon = False
    LogCurrentError MODULE_NAME, "[TryRecoverRibbon]"
End Function

'@Description: Recupera el Ribbon toggleando el estado del add-in.
'              Esto fuerza una nueva llamada a onLoad con un IRibbonUI fresco.
'              Solo llamar bajo confirmacion explicita del usuario.
Private Function RecoverByAddinToggle() As Boolean
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[RecoverByAddinToggle] Iniciando toggle del add-in"

    Dim ai As AddIn
    Dim targetAddin As AddIn

    For Each ai In Application.AddIns
        If ai.Name = APP_NAME & ".xlam" Then
            Set targetAddin = ai
            Exit For
        End If
    Next ai

    If targetAddin Is Nothing Then
        LogError MODULE_NAME, "[RecoverByAddinToggle] Add-in no encontrado: " & APP_NAME & ".xlam"
        RecoverByAddinToggle = False
        Exit Function
    End If

    If Not targetAddin.Installed Then
        LogError MODULE_NAME, "[RecoverByAddinToggle] Add-in no esta instalado"
        RecoverByAddinToggle = False
        Exit Function
    End If

    ' Toggle: desactivar y reactivar (fuerza nueva llamada a onLoad)
    LogDebug MODULE_NAME, "[RecoverByAddinToggle] Desactivando add-in..."
    targetAddin.Installed = False

    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    LogDebug MODULE_NAME, "[RecoverByAddinToggle] Reactivando add-in..."
    targetAddin.Installed = True

    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 2)
    DoEvents

    RecoverByAddinToggle = IsRibbonAvailable()

    If RecoverByAddinToggle Then
        LogInfo MODULE_NAME, "[RecoverByAddinToggle] Ribbon recuperado via toggle"
    Else
        LogWarning MODULE_NAME, "[RecoverByAddinToggle] Toggle completado pero Ribbon no disponible"
    End If

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[RecoverByAddinToggle]"
    RecoverByAddinToggle = False
End Function

'@Description: Devuelve el valor almacenado en el nombre Excel4 (para diagnostico)
Private Function SafeGetExcelName() As String
    On Error Resume Next
    SafeGetExcelName = Application.ExecuteExcel4Macro(RIBBON_PTR_EXCEL_NAME)
    If Err.Number <> 0 Then SafeGetExcelName = "(error al leer)"
    On Error GoTo 0
End Function

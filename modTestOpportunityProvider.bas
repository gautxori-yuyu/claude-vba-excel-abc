Attribute VB_Name = "modTestOpportunityProvider"
' ==================================================================
' MODULO DE TEST: FileSystemOpportunityProvider
' ==================================================================
' Ejercita la clase FileSystemOpportunityProvider para verificar
' su funcionamiento y generar trazas en el log.
'
' COMO EJECUTAR:
'   1. Desde la Ventana Inmediato:  Call TestOpportunityProvider()
'   2. Desde el Editor VBA:         F5 con cursor en cualquier Sub
'   3. Macro:                        Alt+F8 -> TestOpportunityProvider
'
' Los resultados aparecen tanto en el log (mod_Logger) como
' en la Ventana Inmediato (Debug.Print).
' ==================================================================
'@Folder "1-Aplicacion.Tests"
Option Private Module
Option Explicit

Private Const MODULE_NAME As String = "modTestOpportunityProvider"

' ==================================================================
' ENTRY POINT PRINCIPAL
' Crea instancia fresca de FileSystemOpportunityProvider,
' inyecta dependencias desde App() y ejercita todos los metodos.
' ==================================================================
Public Sub TestOpportunityProvider()
    Const PROC_NAME As String = "TestOpportunityProvider"
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "================================================="
    LogInfo MODULE_NAME, "[TestOpportunityProvider] INICIO DEL TEST"
    LogInfo MODULE_NAME, "================================================="
    Debug.Print "=== TestOpportunityProvider ==="

    ' Verificar que App esta disponible
    Dim mApp As clsApplication
    Set mApp = App()
    If mApp Is Nothing Then
        LogError MODULE_NAME, "[TestOpportunityProvider] App() no disponible - inicializar el complemento primero"
        Debug.Print "ERROR: App() no disponible"
        Exit Sub
    End If

    If mApp.FileMgr Is Nothing Then
        LogError MODULE_NAME, "[TestOpportunityProvider] App.FileMgr no disponible"
        Exit Sub
    End If

    If mApp.Configuration Is Nothing Then
        LogError MODULE_NAME, "[TestOpportunityProvider] App.Configuration no disponible"
        Exit Sub
    End If

    ' Crear instancia fresca para test (independiente de la instancia de produccion)
    Dim provider As FileSystemOpportunityProvider
    Set provider = New FileSystemOpportunityProvider
    provider.Initialize mApp.FileMgr, mApp.Configuration

    ' ------------------------------------------------------------------
    ' TEST 1: BasePath
    ' ------------------------------------------------------------------
    LogInfo MODULE_NAME, "[Test1] BasePath = '" & provider.BasePath & "'"
    Debug.Print "  BasePath: " & provider.BasePath

    ' ------------------------------------------------------------------
    ' TEST 2: GetOpportunityFolders (= GetOpportunities en la interfaz)
    ' ------------------------------------------------------------------
    Dim folders As Collection
    Set folders = provider.GetOpportunityFolders()

    LogInfo MODULE_NAME, "[Test2] GetOpportunityFolders: " & folders.Count & " oportunidades encontradas"
    Debug.Print "  GetOpportunityFolders: " & folders.Count & " resultados"

    Dim i As Long
    For i = 1 To folders.Count
        LogDebug MODULE_NAME, "[Test2]   [" & i & "] " & folders(i)
        Debug.Print "    [" & i & "] " & folders(i)
        If i >= 5 Then
            LogDebug MODULE_NAME, "[Test2]   ... (mostrando primeras 5 de " & folders.Count & ")"
            Debug.Print "    ... (primeras 5 de " & folders.Count & ")"
            Exit For
        End If
    Next i

    ' ------------------------------------------------------------------
    ' TEST 3: GetNextOpportunityCode
    ' ------------------------------------------------------------------
    Dim nextCode As String
    nextCode = provider.GetNextOpportunityCode()
    LogInfo MODULE_NAME, "[Test3] GetNextOpportunityCode = '" & nextCode & "'"
    Debug.Print "  GetNextOpportunityCode: " & nextCode

    ' ------------------------------------------------------------------
    ' TEST 4: FolderExists con primera oportunidad real
    ' ------------------------------------------------------------------
    If folders.Count > 0 Then
        Dim firstOp As String
        firstOp = folders(1)

        Dim existeReal As Boolean
        existeReal = provider.FolderExists(firstOp)
        LogInfo MODULE_NAME, "[Test4] FolderExists('" & firstOp & "') = " & existeReal
        Debug.Print "  FolderExists('" & firstOp & "'): " & existeReal

        ' TEST 4b: FolderExists con nombre inventado
        Dim fakeOp As String
        fakeOp = "NO_EXISTE_ESTA_CARPETA_TEST_XYZ"
        Dim existeFake As Boolean
        existeFake = provider.FolderExists(fakeOp)
        LogInfo MODULE_NAME, "[Test4b] FolderExists('" & fakeOp & "') = " & existeFake & " (esperado: False)"
        Debug.Print "  FolderExists('" & fakeOp & "'): " & existeFake & " (esperado False)"

        ' TEST 5: GetFullPath
        Dim fullPath As String
        fullPath = provider.GetFullPath(firstOp)
        LogInfo MODULE_NAME, "[Test5] GetFullPath('" & firstOp & "') = '" & fullPath & "'"
        Debug.Print "  GetFullPath: " & fullPath
    Else
        LogWarning MODULE_NAME, "[Test4] Sin oportunidades - saltando tests 4 y 5"
        Debug.Print "  (sin oportunidades para tests 4 y 5)"
    End If

    ' ------------------------------------------------------------------
    ' TEST 6: Via interfaz IOpportunityProvider
    ' Ejercita el comportamiento polimorfico (stubs de delegacion)
    ' ------------------------------------------------------------------
    Dim iProvider As IOpportunityProvider
    Set iProvider = provider

    Dim viaInterface As Collection
    Set viaInterface = iProvider.GetOpportunities()
    LogInfo MODULE_NAME, "[Test6] IOpportunityProvider.GetOpportunities() = " & viaInterface.Count & " (via interfaz)"
    Debug.Print "  Via interfaz GetOpportunities(): " & viaInterface.Count

    Dim nextCodeInterface As String
    nextCodeInterface = iProvider.GetNextOpportunityCode()
    LogInfo MODULE_NAME, "[Test6] IOpportunityProvider.GetNextOpportunityCode() = '" & nextCodeInterface & "'"
    Debug.Print "  Via interfaz GetNextOpportunityCode(): " & nextCodeInterface

    If folders.Count > 0 Then
        Dim existsInterface As Boolean
        existsInterface = iProvider.OpportunityExists(folders(1))
        LogInfo MODULE_NAME, "[Test6] IOpportunityProvider.OpportunityExists('" & folders(1) & "') = " & existsInterface
        Debug.Print "  Via interfaz OpportunityExists: " & existsInterface

        Dim pathInterface As String
        pathInterface = iProvider.GetOpportunityPath(folders(1))
        LogInfo MODULE_NAME, "[Test6] IOpportunityProvider.GetOpportunityPath('" & folders(1) & "') = '" & pathInterface & "'"
        Debug.Print "  Via interfaz GetOpportunityPath: " & pathInterface
    End If

    ' ------------------------------------------------------------------
    ' RESULTADO FINAL
    ' ------------------------------------------------------------------
    LogInfo MODULE_NAME, "================================================="
    LogInfo MODULE_NAME, "[TestOpportunityProvider] FIN - " & folders.Count & " oportunidades, siguiente codigo: " & nextCode
    LogInfo MODULE_NAME, "================================================="
    Debug.Print "=== FIN TestOpportunityProvider ==="

    Set provider = Nothing
    Set iProvider = Nothing
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[TestOpportunityProvider]"
    Debug.Print "ERROR en TestOpportunityProvider: " & Err.Description
End Sub

' ==================================================================
' TEST DE STRESS: Llama a GetOpportunityFolders varias veces
' para verificar rendimiento y ausencia de efectos secundarios.
' ==================================================================
Public Sub TestOpportunityProviderStress()
    Const PROC_NAME As String = "TestOpportunityProviderStress"
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[TestOpportunityProviderStress] INICIO (10 iteraciones)"
    Debug.Print "=== Stress Test ==="

    Dim provider As FileSystemOpportunityProvider
    Set provider = New FileSystemOpportunityProvider
    provider.Initialize App.FileMgr, App.Configuration

    Dim startTime As Double
    startTime = Timer

    Dim j As Long
    Dim totalOps As Long
    For j = 1 To 10
        Dim result As Collection
        Set result = provider.GetOpportunityFolders()
        totalOps = totalOps + result.Count
    Next j

    Dim elapsed As Double
    elapsed = Timer - startTime

    LogInfo MODULE_NAME, "[TestOpportunityProviderStress] 10 llamadas en " & Format(elapsed, "0.000") & "s | Promedio ops: " & (totalOps / 10)
    Debug.Print "  10 llamadas en " & Format(elapsed, "0.000") & "s | Ops promedio: " & (totalOps / 10)
    Debug.Print "=== FIN Stress Test ==="

    Set provider = Nothing
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[TestOpportunityProviderStress]"
End Sub

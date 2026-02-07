Attribute VB_Name = "UDFs_FileSystem"
'@Folder "Funciones auxiliares"
Option Explicit

Private Const MODULE_NAME As String = "UDFs_FileSystem"

'@Description: Valida si una ruta de carpeta existe
Public Function RutaExiste(ruta As String) As Boolean
Attribute RutaExiste.VB_Description = "[UDFs_FileSystem] Valida si una ruta de carpeta existe"
Attribute RutaExiste.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error Resume Next
    RutaExiste = ruta <> "" And (Dir(ruta, vbDirectory) <> "")
    On Error GoTo 0
End Function
'@Description: Determina si una ruta es de red o unidad removible
'@Scope: Privado
'@ArgumentDescriptions: ruta: Ruta a verificar
'@Returns: Boolean | True si es ruta de red o removible
Function EsRutaRemovibleORed(ruta As String) As Boolean
Attribute EsRutaRemovibleORed.VB_Description = "[UDFs_FileSystem] Determina si una ruta es de red o unidad removible"
Attribute EsRutaRemovibleORed.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim fso As Object
    Dim drive As Object
    Dim driveLetter As String
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Caso 1: Ruta UNC (\\servidor\compartido)
    If Left$(ruta, 2) = "\\" Then
        EsRutaRemovibleORed = True
        Exit Function
    End If
    
    ' Caso 2: Verificar tipo de unidad
    driveLetter = fso.GetDriveName(ruta)
    
    If driveLetter <> "" And fso.DriveExists(driveLetter) Then
        Set drive = fso.GetDrive(driveLetter)
        
        ' DriveType: 0=Unknown, 1=Removable, 2=Fixed, 3=Network, 4=CDRom, 5=RamDisk
        If drive.DriveType = 1 Or drive.DriveType = 3 Then  ' Removable o Network
            EsRutaRemovibleORed = True
        Else
            EsRutaRemovibleORed = False
        End If
    Else
        ' No se pudo determinar, asumir que es removible por seguridad
        EsRutaRemovibleORed = True
    End If
    
    On Error GoTo 0
End Function

'@Description: Determina si una ruta es de red (UNC o unidad mapeada)
'@Scope: Private (uso interno del módulo)
'@ArgumentDescriptions: ruta | Ruta completa a verificar
'@Returns: Boolean | True si es ruta de red
Function IsNetworkPath(ByVal ruta As String) As Boolean
Attribute IsNetworkPath.VB_Description = "[UDFs_FileSystem] Determina si una ruta es de red (UNC o unidad mapeada)"
Attribute IsNetworkPath.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error Resume Next
    
    IsNetworkPath = False
    
    ' Normalizar ruta
    If Right(ruta, 1) = "\" Then
        ruta = Left(ruta, Len(ruta) - 1)
    End If
    
    ' 1. Detectar rutas UNC (\\servidor\compartido)
    If Left(ruta, 2) = "\\" Then
        IsNetworkPath = True
        Exit Function
    End If
    
    ' 2. Detectar unidades de red mapeadas (Z:, Y:, etc.)
    If Mid(ruta, 2, 1) = ":" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim drive As Object
        Dim driveLetter As String
        driveLetter = Left(ruta, 2)              ' Ej: "C:", "Z:"
        
        ' Verificar si la unidad existe
        If fso.DriveExists(driveLetter) Then
            Set drive = fso.GetDrive(driveLetter)
            
            ' DriveType: 0=Unknown, 1=Removable, 2=Fixed, 3=Network, 4=CDRom, 5=RamDisk
            If drive.DriveType = 3 Then          ' 3 = Network
                IsNetworkPath = True
            End If
        End If
        
        Set drive = Nothing
        Set fso = Nothing
    End If
    
    On Error GoTo 0
End Function

'@Description: Normaliza una ruta eliminando la barra final
Function NormalizarRuta(ByVal ruta As String) As String
Attribute NormalizarRuta.VB_Description = "[UDFs_FileSystem] Normaliza una ruta eliminando la barra final"
Attribute NormalizarRuta.VB_ProcData.VB_Invoke_Func = " \n23"
    If Right(ruta, 1) = "\" Then
        NormalizarRuta = Left(ruta, Len(ruta) - 1)
    Else
        NormalizarRuta = ruta
    End If
End Function

'@Description: Obtiene el nombre de una carpeta de su ruta completa
Function ObtenerNombreCarpeta(ByVal rutaCompleta As String) As String
Attribute ObtenerNombreCarpeta.VB_Description = "[UDFs_FileSystem] Obtiene el nombre de una carpeta de su ruta completa"
Attribute ObtenerNombreCarpeta.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ObtenerNombreCarpeta = fso.GetFileName(rutaCompleta)
    On Error GoTo 0
    
    Set fso = Nothing
End Function

Function ObtenerRutaEjecutable(nombreExe As String) As String
Attribute ObtenerRutaEjecutable.VB_Description = "[UDFs_FileSystem] Obtener Ruta Ejecutable (función personalizada)"
Attribute ObtenerRutaEjecutable.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim objShell As Object
    Dim objExec As Object
    Dim comando As String
    Dim resultado As String

    Set objShell = CreateObject("WScript.Shell")
    
    ' Construye el comando WHERE para el ejecutable indicado
    comando = "cmd.exe /u where " & nombreExe
    
    ' Ejecuta el comando de forma oculta y captura la salida
    Set objExec = objShell.Exec(comando)
    resultado = objExec.StdOut.ReadAll
    
    ' Limpiar saltos de línea y devolver solo la primera ruta encontrada
    If resultado <> "" Then
        ObtenerRutaEjecutable = Split(resultado, vbCrLf)(0)
    Else
        ObtenerRutaEjecutable = ""
    End If
End Function

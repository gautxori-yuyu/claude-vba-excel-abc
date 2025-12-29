Attribute VB_Name = "modLogger"
'@Folder "2-Control de estado"
' ==========================================
' MODULO DE LOGGING CENTRALIZADO
' ==========================================
' Proporciona funciones de logging consistentes para toda la aplicacion.
' Soporta diferentes niveles de log y salida a Debug.Print o archivo.
' ==========================================

Option Explicit

' Niveles de logging
Public Enum LogLevel
    LOG_DEBUG = 0
    LOG_INFO = 1
    LOG_WARNING = 2
    LOG_ERROR = 3
    LOG_CRITICAL = 4
End Enum

' Configuracion del logger
Private mMinLevel As LogLevel
Private mLogToFile As Boolean
Private mLogFilePath As String
Private mIncludeTimestamp As Boolean

' ==========================================
' INICIALIZACION
' ==========================================

Public Sub InitLogger(Optional ByVal minLevel As LogLevel = LOG_DEBUG, _
                      Optional ByVal logToFile As Boolean = False, _
                      Optional ByVal logFilePath As String = "")
    mMinLevel = minLevel
    mLogToFile = logToFile
    mIncludeTimestamp = True

    If logToFile And logFilePath = "" Then
        mLogFilePath = Environ("TEMP") & "\ABC_VBA_Log_" & Format(Date, "yyyy-mm-dd") & ".txt"
    Else
        mLogFilePath = logFilePath
    End If
End Sub

' ==========================================
' FUNCIONES PUBLICAS DE LOGGING
' ==========================================

'@Description: Registra un mensaje de debug (solo en modo desarrollo)
Public Sub LogDebug(ByVal source As String, ByVal message As String)
    WriteLog LOG_DEBUG, source, message
End Sub

'@Description: Registra un mensaje informativo
Public Sub LogInfo(ByVal source As String, ByVal message As String)
    WriteLog LOG_INFO, source, message
End Sub

'@Description: Registra una advertencia
Public Sub LogWarning(ByVal source As String, ByVal message As String)
    WriteLog LOG_WARNING, source, message
End Sub

'@Description: Registra un error
Public Sub LogError(ByVal source As String, ByVal message As String, _
                    Optional ByVal errNumber As Long = 0, _
                    Optional ByVal errDescription As String = "")
    Dim fullMessage As String
    fullMessage = message

    If errNumber <> 0 Then
        fullMessage = fullMessage & " [Error " & errNumber & ": " & errDescription & "]"
    End If

    WriteLog LOG_ERROR, source, fullMessage
End Sub

'@Description: Registra un error critico
Public Sub LogCritical(ByVal source As String, ByVal message As String, _
                       Optional ByVal errNumber As Long = 0, _
                       Optional ByVal errDescription As String = "")
    Dim fullMessage As String
    fullMessage = "CRITICO: " & message

    If errNumber <> 0 Then
        fullMessage = fullMessage & " [Error " & errNumber & ": " & errDescription & "]"
    End If

    WriteLog LOG_CRITICAL, source, fullMessage
End Sub

'@Description: Registra el error actual del objeto Err
Public Sub LogCurrentError(ByVal source As String, Optional ByVal additionalInfo As String = "")
    If Err.Number = 0 Then Exit Sub

    Dim message As String
    message = "Error capturado"
    If additionalInfo <> "" Then message = message & " - " & additionalInfo

    LogError source, message, Err.Number, Err.Description
End Sub

' ==========================================
' FUNCIONES PRIVADAS
' ==========================================

Private Sub WriteLog(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
    ' Verificar nivel minimo
    If level < mMinLevel Then Exit Sub

    ' Construir mensaje formateado
    Dim logMessage As String
    logMessage = FormatLogMessage(level, source, message)

    ' Salida a Debug.Print
    Debug.Print logMessage

    ' Salida a archivo si esta habilitado
    If mLogToFile Then
        WriteToFile logMessage
    End If
End Sub

Private Function FormatLogMessage(ByVal level As LogLevel, _
                                  ByVal source As String, _
                                  ByVal message As String) As String
    Dim prefix As String

    ' Prefijo segun nivel
    Select Case level
        Case LOG_DEBUG:    prefix = "[DEBUG]   "
        Case LOG_INFO:     prefix = "[INFO]    "
        Case LOG_WARNING:  prefix = "[WARNING] "
        Case LOG_ERROR:    prefix = "[ERROR]   "
        Case LOG_CRITICAL: prefix = "[CRITICAL]"
        Case Else:         prefix = "[UNKNOWN] "
    End Select

    ' Construir mensaje
    If mIncludeTimestamp Then
        FormatLogMessage = Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & prefix & " [" & source & "] " & message
    Else
        FormatLogMessage = prefix & " [" & source & "] " & message
    End If
End Function

Private Sub WriteToFile(ByVal message As String)
    On Error Resume Next

    Dim fileNum As Integer
    fileNum = FreeFile

    Open mLogFilePath For Append As #fileNum
    Print #fileNum, message
    Close #fileNum

    On Error GoTo 0
End Sub

' ==========================================
' UTILIDADES
' ==========================================

'@Description: Obtiene el nombre del nivel de log
Public Function GetLevelName(ByVal level As LogLevel) As String
    Select Case level
        Case LOG_DEBUG:    GetLevelName = "DEBUG"
        Case LOG_INFO:     GetLevelName = "INFO"
        Case LOG_WARNING:  GetLevelName = "WARNING"
        Case LOG_ERROR:    GetLevelName = "ERROR"
        Case LOG_CRITICAL: GetLevelName = "CRITICAL"
        Case Else:         GetLevelName = "UNKNOWN"
    End Select
End Function

'@Description: Limpia el archivo de log
Public Sub ClearLogFile()
    On Error Resume Next

    If mLogFilePath <> "" Then
        Kill mLogFilePath
    End If

    On Error GoTo 0
End Sub

'@Description: Obtiene la ruta del archivo de log actual
Public Function GetLogFilePath() As String
    GetLogFilePath = mLogFilePath
End Function

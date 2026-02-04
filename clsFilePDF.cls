VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFilePDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' EJEMPLO DE IMPLEMENTACIÃN FUTURA
' ==================================================================
' Demuestra cÃ³mo clsFileManager puede gestionar diferentes tipos
' de archivos sin modificar su cÃ³digo
' ==================================================================

' ==================================================================
' USO DESDE clsFileManager
' ==================================================================
' Dim pdfFile As clsFilePDF
' Set pdfFile = New clsFilePDF
' pdfFile.BindTo "C:\Oportunidades\ABC\documento.pdf"
'
' ' FileManager lo trackea sin saber que es PDF
' mFileMgr.TrackFile pdfFile
' ==================================================================

'@Folder "4-Servicios.Archivos.Tipos de archivos"
'@Implements IFile
Option Explicit

Private Const MODULE_NAME As String = "clsFilePDF"

Private p_filePath As String
Private p_fileInfo As Object                     ' PodrÃ­a ser metadatos del PDF
Private p_ObjectKey As Double

' ImplementaciÃ³n de IFile
Private Property Get IFile_ObjectKey() As Double
    IFile_ObjectKey = p_ObjectKey
End Property

Private Property Get IFile_Path() As String
    IFile_Path = p_filePath
End Property

Private Property Get IFile_Name() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IFile_Name = fso.GetFileName(p_filePath)
End Property

' Propiedades pÃºblicas
Public Property Get ObjectKey() As Double
    ObjectKey = p_ObjectKey
End Property

Public Property Get Path() As String
    Path = p_filePath
End Property

Public Property Get Name() As String
    Name = IFile_Name
End Property

' MÃ©todo especÃ­fico de PDF
Public Sub BindTo(filePath As String)
    If p_ObjectKey <> 0 Then
        Err.Raise vbObjectError + 580, TypeName(Me), "Ya estÃ¡ asociado a un archivo."
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        Err.Raise vbObjectError + 581, TypeName(Me), "El archivo no existe."
    End If
    
    p_filePath = filePath
    ' Para archivos en disco, usar hash del path como ObjectKey
    p_ObjectKey = GetPathHashCode(filePath)
End Sub

' Helper para generar ObjectKey consistente desde path
Private Function GetPathHashCode(Path As String) As Double
    ' ImplementaciÃ³n simple: sum de cÃ³digos ASCII
    Dim i As Long, hash As Double
    For i = 1 To Len(Path)
        hash = hash + Asc(Mid(Path, i, 1)) * i
    Next i
    GetPathHashCode = hash
End Function

' MÃ©todos especÃ­ficos de PDF
Public Function ExtractText() As String
    ' ImplementaciÃ³n futura: usar librerÃ­a de PDF
End Function

Public Function GetPageCount() As Long
    ' ImplementaciÃ³n futura
End Function


Attribute VB_Name = "modMACROListarProyectosVBA"
'@Folder "1-Inicio e Instalacion.Gestion de modulos y procs"
'@Ignore VariableNotUsed
Option Explicit

Sub ListarProyectosAlternativo()
Attribute ListarProyectosAlternativo.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Este método intenta acceder de forma indirecta
    Dim i As Integer
    Dim projName As String, vbap As Object
    
    On Error GoTo ErrorHandler
    
    Debug.Print "INTENTANDO ACCEDER A PROYECTOS VBA:"
    Debug.Print "=================================="
    
    For i = 1 To Application.VBE.VBProjects.Count
        Set vbap = Application.VBE.VBProjects(i)
        projName = Application.VBE.VBProjects(i).Name
        Debug.Print i & ". " & projName
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR: No se puede acceder a los proyectos VBA directamente"
    Debug.Print "Solución: Habilitar la referencia VBE Extensibility"
    Debug.Print "1. Ir al Editor VBA (ALT + F11)"
    Debug.Print "2. Tools > References"
    Debug.Print "3. Buscar 'Microsoft Visual Basic for Applications Extensibility 5.3'"
    Debug.Print "4. Marcar la casilla y hacer OK"
End Sub

Sub ListarProyectosVBAAuto()
Attribute ListarProyectosVBAAuto.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim vbProj As Object
    Dim contador As Integer
    
    ' Intentar habilitar la referencia automáticamente
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid _
        "{0002E157-0000-0000-C000-000000000046}", 5, 3 ' VBE Extensibility
    
    contador = 0
    Debug.Print "PROYECTOS VBA DETECTADOS:"
    Debug.Print "========================"
    
    For Each vbProj In Application.VBE.VBProjects
        contador = contador + 1
        Debug.Print contador & ". " & vbProj.Name
        Debug.Print "   File: " & vbProj.fileName
        Debug.Print "   Type: " & TypeName(vbProj)
        Debug.Print "------------------------"
    Next vbProj
    
    If contador = 0 Then
        Debug.Print "No se pudieron acceder a los proyectos VBA"
        Debug.Print "Verifica que tengas la referencia habilitada:"
        Debug.Print "Microsoft Visual Basic for Applications Extensibility 5.3"
    Else
        Debug.Print "Total proyectos: " & contador
    End If
End Sub

Sub ListarProyectosVBAIncluyendoXLAM()
Attribute ListarProyectosVBAIncluyendoXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim vbProj As VBIDE.VBProject
    Dim contador As Integer
    Dim tipoProyecto As String
    
    ' Referencia necesaria: Microsoft Visual Basic for Applications Extensibility 5.3
    ' Tools > References > Microsoft Visual Basic for Applications Extensibility 5.3
    
    On Error Resume Next                         ' Por si no tiene la referencia habilitada
    
    contador = 0
    Debug.Print "PROYECTOS VBA (INCLUYENDO XLAM NO INSTALADOS):"
    Debug.Print "=============================================="
    
    For Each vbProj In Application.VBE.VBProjects
        contador = contador + 1
        
        ' Determinar el tipo de proyecto
        Select Case vbProj.Protection
        Case vbext_pp_locked
            tipoProyecto = "Protegido"
        Case vbext_pp_none
            tipoProyecto = "No protegido"
        Case Else
            tipoProyecto = "Desconocido"
        End Select
        
        Debug.Print contador & ". " & vbProj.Name
        Debug.Print "   Descripción: " & vbProj.Description
        Debug.Print "   Tipo: " & tipoProyecto
        Debug.Print "   Archivo: " & vbProj.fileName
        Debug.Print "   HelpFile: " & vbProj.HelpFile
        Debug.Print "   Mode: " & vbProj.Mode
        Debug.Print "----------------------------------"
    Next vbProj
    
    Debug.Print "Total proyectos VBA: " & contador
End Sub



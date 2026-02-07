VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wshUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME As String = "wshUnidades"

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A:E")) Is Nothing Then
        Call GenerarListasUnidades
    End If
End Sub



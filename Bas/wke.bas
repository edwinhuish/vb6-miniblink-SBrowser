Attribute VB_Name = "wke"
Option Explicit
Private IsInitWkeApi As Boolean
Public NodeDllPath As String

Public Sub wke_api_init()
    If IsInitWkeApi = True Then Exit Sub
    
    NodeDllPath = "includes\node.dll"
    
    If Dir(NodeDllPath) = "" Then
        MsgBox "node.dll ²»´æÔÚ: " & NodeDllPath, vbSystemModal
        IsInitWkeApi = False
        Exit Sub
    End If
    
    IsInitWkeApi = True
End Sub

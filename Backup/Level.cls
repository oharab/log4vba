VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DisplayName As String
Public name As String
Public value As String

Private Sub Class_Initialize()
    With Me
        .DisplayName = "Off"
        .name = "Off"
        .value = 99
    End With
End Sub

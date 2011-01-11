VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NullLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Dbug(sMessage As String)
End Sub
Public Sub Info(sMessage As String)
End Sub
Public Sub Warn(sMessage As String)
End Sub
Public Sub Error(sMessage As String, Exception As ErrObject)
End Sub
Public Sub Fatal(sMessage As String, Optional Exception As ErrObject)
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Property Get Name() As String
End Property
Public Property Let Name(sName As String)
End Property

Public Function IsDbugEnabled() As Boolean
End Function
Public Function IsInfoEnabled() As Boolean
End Function
Public Function IsWarnEnabled() As Boolean
End Function
Public Function IsErrorEnabled() As Boolean
End Function
Public Function IsFatalEnabled() As Boolean
End Function


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

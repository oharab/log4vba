VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub OutputDebugString Lib "kernel32" _
                                      Alias "OutputDebugStringA" _
                                      (ByVal lpOutputString As String)
                                      
Public Sub Dbug(sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Dbug
    End If
End Sub
Public Sub Info(sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Info
    End If
End Sub
Public Sub Warn(sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Warn
    End If
End Sub
Public Sub Error(sMessage As String, Exception As ErrObject)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Error
    End If
End Sub
Public Sub Fatal(sMessage As String, Optional Exception As ErrObject)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Fatal
    End If
End Sub


Private Sub doLog(sMessage As String, oLevel As Level, Optional Exception As ErrObject)
    Dim s As String
    s = "[log4vba] " & oLevel.Name & vbTab & Now & vbTab & sMessage
    If Not Exception Is Nothing Then
        s = s & vbTab & Exception.Number & vbTab & Exception.Description & vbTab & Exception.Source
    End If

    OutputDebugString s
End Sub



Public Property Get Name() As String
    Name = "Internal Logger"
End Property

Public Property Let Name(sName As String)
End Property

Public Function IsDbugEnabled() As Boolean
    IsDbugEnabled = DBGLog4VBA
End Function
Public Function IsInfoEnabled() As Boolean
    IsInfoEnabled = DBGLog4VBA
End Function
Public Function IsWarnEnabled() As Boolean
    IsWarnEnabled = DBGLog4VBA
End Function
Public Function IsErrorEnabled() As Boolean
    IsErrorEnabled = DBGLog4VBA
End Function
Public Function IsFatalEnabled() As Boolean
    IsFatalEnabled = DBGLog4VBA
End Function


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

Public Sub Dbug(ByVal sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Dbug
    End If
End Sub
Public Sub Info(ByVal sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Info
    End If
End Sub
Public Sub Warn(ByVal sMessage As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Warn
    End If
End Sub
Public Sub Error(ByVal sMessage As String, Optional ByVal ExceptionNumber As Long, Optional ByVal ExceptionDescription As String, Optional ByVal ExceptionSource As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Error, ExceptionNumber, ExceptionDescription, ExceptionSource
    End If
End Sub
Public Sub Fatal(ByVal sMessage As String, Optional ByVal ExceptionNumber As Long, Optional ByVal ExceptionDescription As String, Optional ByVal ExceptionSource As String)
    If DBGLog4VBA Then
        doLog sMessage, Levels.Fatal, ExceptionNumber, ExceptionDescription, ExceptionSource
    End If
End Sub


Private Sub doLog(ByVal sMessage As String, ByVal oLevel As Level, Optional ByVal ExceptionNumber As Long, Optional ByVal ExceptionDescription As String, Optional ByVal ExceptionSource As String)
    Dim s As String
    s = "[log4vba] " & oLevel.name & vbTab & Now & vbTab & sMessage
    If ExceptionNumber <> 0 Then
        s = s & vbTab & ExceptionNumber & vbTab & ExceptionDescription & vbTab & ExceptionSource
    End If

    OutputDebugString s
End Sub



Public Property Get name() As String
    name = "Internal Logger"
End Property

Public Property Let name(sName As String)
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputDebugStringAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAppender


Private log As ILog

Private Declare Sub OutputDebugString Lib "kernel32" _
                                      Alias "OutputDebugStringA" _
                                      (ByVal lpOutputString As String)

Private m_name As String
Private m_threshold As Level

Private Sub Class_Initialize()
    Set m_threshold = Levels.Dbug
    m_name = "OutputDebugStringAppender"
End Sub

Public Property Get IAppender_Name() As String
    IAppender_Name = m_name
End Property
Public Property Let IAppender_Name(sName As String)
    m_name = sName
End Property

Public Sub IAppender_DoAppend(oLoggingEvent As LoggingEvent)
    With oLoggingEvent
        If .Level.value >= Me.Threshold.value Then
            Dim s As String
            s = "[" & .LoggerName & "] " & .Level.Name & vbTab & .TimeStamp & vbTab & oLoggingEvent.Message
            If .ErrorNumber <> 0 Then
                s = s & vbTab & .ErrorNumber & vbTab & .ErrorDescription & vbTab & .ErrorSource
            End If

            OutputDebugString s
        End If
    End With
End Sub

Public Property Get Threshold() As Level
    Set Threshold = m_threshold
End Property
Public Property Set Threshold(oThreshold As Level)
    Set m_threshold = oThreshold
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAppender

Private m_name As String
Private m_threshold As Level

Private Sub Class_Initialize()
    Set m_threshold = Levels.Dbug
    m_name = "ConsoleAppender"
End Sub

Public Property Get IAppender_Name() As String
    IAppender_Name = m_name
End Property
Public Property Let IAppender_Name(sName As String)
    m_name = sName
End Property

Public Sub IAppender_DoAppend(oLoggingEvent As LoggingEvent)
    With oLoggingEvent
        Dim s As String
        If .Level.value >= m_threshold.value Then
            s = "[" & .LoggerName & "] " & .Level.Name & vbTab & .TimeStamp & vbTab & oLoggingEvent.Message
            If .ErrorNumber <> 0 Then
                s = s & vbTab & .ErrorNumber & vbTab & .ErrorDescription & vbTab & .ErrorSource
            End If
            Debug.Print s
        End If
    End With
End Sub
Public Property Get Threshold() As Level
    Set Threshold = m_threshold
End Property
Public Property Set Threshold(oThreshold As Level)
    Set m_threshold = oThreshold
End Property

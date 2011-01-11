VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NullAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAppender

Private m_name As String
Private m_threshold As Level

Private Sub Class_Initialize()
    m_threshold = Levels.Off
    m_name = "NullAppender"
End Sub

Public Property Get IAppender_Name() As String
    IAppender_Name = m_name
End Property
Public Property Let IAppender_Name(sName As String)
    m_name = sName
End Property

Public Sub IAppender_DoAppend(oLoggingEvent As LoggingEvent)
End Sub

Public Property Get Threshold() As Level
    Set Threshold = m_threshold
End Property
Public Property Set Threshold(oThreshold As Level)
End Property

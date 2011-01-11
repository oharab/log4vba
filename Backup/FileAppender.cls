VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAppender

Private log As ILog


Private m_name As String
Private m_threshold As Level
Private m_file As String
Private m_lockingModel As LockingModel
Private m_appendToFile As Boolean

Private m_fileNum As Integer
Private m_openFile As Boolean
Public Enum LockingModel
    ExclusiveLock
    MinimalLock
End Enum

Private Sub Class_Initialize()
    Set m_threshold = Levels.Dbug
    m_name = "FileAppender"
    m_lockingModel = ExclusiveLock

End Sub

Private Sub Class_Terminate()
    If m_openFile Then Close #m_fileNum
End Sub

Public Property Get IAppender_Name() As String
    IAppender_Name = m_name
End Property
Public Property Let IAppender_Name(ByVal sName As String)
    m_name = sName
End Property

Public Sub IAppender_DoAppend(ByVal oLoggingEvent As LoggingEvent)
    With oLoggingEvent
        If .Level.value >= Me.Threshold.value Then
            If m_lockingModel = MinimalLock Then
                OpenFile
                writeFile oLoggingEvent
                CloseFile
            Else
                If Not m_openFile Then OpenFile
                writeFile oLoggingEvent
            End If
        End If
    End With
End Sub

Public Property Get Threshold() As Level
    Set Threshold = m_threshold
End Property
Public Property Set Threshold(ByVal oThreshold As Level)
    Set m_threshold = oThreshold
End Property

Public Property Let File(ByVal sFile As String)
    m_file = sFile
End Property
Public Property Let LockingModel(ByVal eLockingModel As LockingModel)
    m_lockingModel = eLockingModel
End Property
Public Property Let AppendToFile(ByVal bAppend As Boolean)
    m_appendToFile = bAppend
End Property

Private Sub OpenFile()
    m_fileNum = FreeFile
    Open m_file For Append As m_fileNum
    m_openFile = True
End Sub

Private Sub CloseFile()
    If m_openFile Then Close #m_fileNum
    m_openFile = False
End Sub

Private Sub writeFile(ByVal oLoggingEvent As LoggingEvent)
    With oLoggingEvent
        Dim s As String
            s = "[" & .LoggerName & "] " & .Level.name & vbTab & .TimeStamp & vbTab & oLoggingEvent.Message
            If .ErrorNumber <> 0 Then
                s = s & vbTab & .ErrorNumber & vbTab & .ErrorDescription & vbTab & .ErrorSource
                If .ErrorLine > 0 Then
                    s = s & vbTab & .ErrorLine
                End If
            End If
            Print #m_fileNum, s
    End With

End Sub

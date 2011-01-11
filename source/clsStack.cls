VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private log As ILog
Private profileLog As ILog

Private m_stackItem As StackItem


Public Sub Enter(sProcedureName As String)
    log.Dbug "Entering procedure " & sProcedureName
    Dim s As New StackItem
    s.Data = sProcedureName
    If m_stackItem Is Nothing Then
    Else
        Set s.Previous = m_stackItem
    End If
    Set m_stackItem = s
End Sub

Public Sub Leave(sProcedureName As String)
    log.Dbug "Exiting procedure " & sProcedureName
    If profileLog.IsInfoEnabled Then
        profileLog.Info "Time in " & sProcedureName & ": " & (CLng(Date) + Timer) * 1000 - m_stackItem.TimeStamp & "MS"
    End If
    If m_stackItem.Previous Is Nothing Then
        Set m_stackItem = Nothing
    Else
        Set m_stackItem = m_stackItem.Previous
    End If
End Sub

Private Sub Class_Initialize()
    Set log = LogManager.GetLogger("GlobalStack")
    Set profileLog = LogManager.GetLogger("GlobalStack.Profile")
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements ILogManager

Private m_log As New LogLog

Private m_loggers As HashTable
Private m_appenders As HashTable
Private m_objRootLevel As Level
Private m_rootAppenders As HashTable


Public Function ILogManager_GetLogger(sLoggerName As String) As ILog
    On Error GoTo GetLogger_Error
    m_log.Dbug "Looking for Logger " & sLoggerName
    If Not m_loggers.Exists(sLoggerName) Then
        m_log.Dbug "Logger """ & sLoggerName & """ not found. Creating new one."
        m_loggers.Add sLoggerName, New clsLogImpl
        With m_loggers(sLoggerName)
            .Name = sLoggerName

            If m_objRootLevel Is Nothing Then
            Else
                Set .Level = m_objRootLevel
            End If
            If m_rootAppenders.count > 0 Then
                Dim key As Variant
                For Each key In m_rootAppenders.Keys
                    .AddAppender m_rootAppenders(CStr(key))
                Next key
            End If
        End With
    End If
    Set ILogManager_GetLogger = m_loggers(sLoggerName)

GetLogger_Exit:
    Exit Function

GetLogger_Error:
    Select Case Err
    Case Else
        m_log.Error "Error getting logger: " & sLoggerName, Err
        Set ILogManager_GetLogger = New clsLogImpl
    End Select
    Resume GetLogger_Exit
    Resume
End Function

Private Sub Class_Initialize()
    Set m_loggers = New HashTable
    Set m_appenders = New HashTable
    Set m_rootAppenders = New HashTable
    Dim currentPath As String
    Dim applicationName As String
    Dim configFilePath As String
    Dim app As Object
    
    Set app = Application
    If app.Name = "Microsoft Excel" Then
        currentPath = app.ActiveWorkbook.Path
        applicationName = Left(app.ActiveWorkbook.Name, Len(app.ActiveWorkbook.Name) - 4)
    ElseIf app.Name = "Microsoft Access" Then
        currentPath = app.CurrentProject.Path
        applicationName = Left(app.CurrentProject.Name, Len(app.CurrentProject.Name) - 4)
    Else
        Err.Raise vbObjectError, "clsLogManager.Initialize", "Unknown application."
    End If
    m_log.Dbug "Application Name=" & applicationName
    Select Case True
    Case Dir(currentPath & "\" & applicationName & ".log4vba") <> ""
        configFilePath = currentPath & "\" & applicationName & ".log4vba"
    Case Dir(currentPath & "\log4vba.xml") <> ""
        configFilePath = currentPath & "\" & "log4vba.xml"
    Case Else
        configFilePath = ""
    End Select
    If Dir(configFilePath) <> "" Then
        Dim config As New XmlConfigurator
        config.File = configFilePath
        config.Configure Me
    End If


End Sub

Private Sub Class_Terminate()
    m_loggers.RemoveAll
    Set m_loggers = Nothing

    m_appenders.RemoveAll
    Set m_appenders = Nothing

    m_rootAppenders.RemoveAll
    Set m_rootAppenders = Nothing
    Set m_objRootLevel = Nothing
End Sub


Friend Property Get Appenders() As HashTable
    Set Appenders = m_appenders
End Property

Friend Property Get Loggers() As HashTable
    Set Loggers = m_loggers
End Property


'---------------------------------------------------------------------------------------
' Procedure : RootLevel
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RootLevel() As Level

    On Error GoTo RootLevel_Error
    GlobalStack.Enter "clsLogManager.RootLevel"
    Set RootLevel = m_objRootLevel


    On Error Resume Next
    GlobalStack.Leave "clsLogManager.RootLevel"
    Exit Property

RootLevel_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.RootLevel", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume RootLevel_Error

End Property

'---------------------------------------------------------------------------------------
' Procedure : RootLevel
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set RootLevel(objRootLevel As Level)

    On Error GoTo RootLevel_Error
    Set m_objRootLevel = objRootLevel


    On Error Resume Next
    Exit Property

RootLevel_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.RootLevel", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume RootLevel_Error

End Property

'---------------------------------------------------------------------------------------
' Procedure : AddRootAppender
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AddRootAppender(appender As IAppender)

    On Error GoTo AddRootAppender_Error
    m_rootAppenders.Add appender.Name, appender

    On Error Resume Next
    Exit Sub

AddRootAppender_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.AddRootAppender", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume AddRootAppender_Error
End Sub

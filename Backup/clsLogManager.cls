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
Private m_configured As Boolean

Public Function ILogManager_GetLogger(ByVal sLoggerName As String) As ILog
    On Error GoTo GetLogger_Error
    If Not m_configured Then AutoConfigure
    m_log.Dbug "Looking for Logger " & sLoggerName
    If Not m_loggers.Exists(sLoggerName) Then
        m_log.Dbug "Logger """ & sLoggerName & """ not found. Creating new one."
        m_loggers.Add sLoggerName, New clsLogImpl
        With m_loggers(sLoggerName)
            .name = sLoggerName

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
            If sLoggerName = "GlobalStack" Or sLoggerName = "GlobalStack.Profile" Then
                Set .Level = Levels.Off
            End If
        End With
    End If
    Set ILogManager_GetLogger = m_loggers(sLoggerName)

GetLogger_Exit:
    Exit Function

GetLogger_Error:
    Select Case Err
    Case Else
        m_log.Error "Error getting logger: " & sLoggerName, Err.Number, Err.Description, Err.Source
        Set ILogManager_GetLogger = New clsLogImpl
    End Select
    Resume GetLogger_Exit
    Resume
End Function

Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Error
    
    Set m_loggers = New HashTable
    Set m_appenders = New HashTable
    Set m_rootAppenders = New HashTable


Class_Initialize_Exit:
    On Error Resume Next
    Exit Sub

Class_Initialize_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.Class_Initialize", Err.Number, Err.Description, Err.Source
    End Select
    Resume Class_Initialize_Exit
    Resume

End Sub

Private Sub Class_Terminate()
    If Not m_loggers Is Nothing Then m_loggers.RemoveAll
    Set m_loggers = Nothing

    If Not m_appenders Is Nothing Then m_appenders.RemoveAll
    Set m_appenders = Nothing

    If Not m_rootAppenders Is Nothing Then m_rootAppenders.RemoveAll
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
    Set RootLevel = m_objRootLevel

RootLevel_Exit:
    On Error Resume Next
    Exit Property

RootLevel_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.RootLevel", Err.Number, Err.Description, Err.Source
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume RootLevel_Exit

End Property

'---------------------------------------------------------------------------------------
' Procedure : RootLevel
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set RootLevel(ByVal objRootLevel As Level)

    On Error GoTo RootLevel_Error
    Set m_objRootLevel = objRootLevel

RootLevel_Exit:
    On Error Resume Next
    Exit Property

RootLevel_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.RootLevel", Err.Number, Err.Description, Err.Source
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume RootLevel_Exit

End Property

'---------------------------------------------------------------------------------------
' Procedure : AddRootAppender
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AddRootAppender(ByVal appender As IAppender)

    On Error GoTo AddRootAppender_Error
    If Not m_rootAppenders.Exists(appender.name) Then
        m_rootAppenders.Add appender.name, appender
    End If

AddRootAppender_Exit:
    On Error Resume Next
    Exit Sub

AddRootAppender_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.AddRootAppender", Err.Number, Err.Description, Err.Description
    End Select
    Resume AddRootAppender_Exit
End Sub

Private Sub AutoConfigure()
    Dim currentPath As String
    Dim applicationName As String
    Dim configFilePath As String
    Dim app As Object

    On Error GoTo AutoConfigure_Error
    m_log.Info "Auto configuring LogManager"
    Set app = Application
    If app.name = "Microsoft Excel" Then
        currentPath = app.ActiveWorkbook.Path
        applicationName = Left(app.ActiveWorkbook.name, Len(app.ActiveWorkbook.name) - 4)
    ElseIf app.name = "Microsoft Access" Then
        currentPath = app.CurrentProject.Path
        applicationName = Left(app.CurrentProject.name, Len(app.CurrentProject.name) - 4)
    Else
        Err.Raise vbObjectError, "clsLogManager.Initialize", "Unknown application."
    End If
    m_log.Dbug "Application Name=" & applicationName
    Select Case True
    Case Dir(currentPath & "\" & applicationName & ".log4vba") <> ""
        configFilePath = currentPath & "\" & applicationName & ".log4vba"
    Case Dir(currentPath & "\log4vba.xml") <> ""
        configFilePath = currentPath & "\" & "log4vba.xml"
    Case Dir(currentPath & "\log4vba.config") <> ""
        configFilePath = currentPath & "\" & "log4vba.config"
    Case Else
        configFilePath = ""
    End Select
    ILogManager_Configure configFilePath
    
AutoConfigure_Exit:
    On Error Resume Next
    Exit Sub

AutoConfigure_Error:
    Select Case Err
    Case Else
        m_log.Error "Unhandled Error in clsLogManager.AutoConfigure", Err.Number, Err.Description, Err.Source
    End Select
    Resume AutoConfigure_Exit
    Resume

End Sub

Public Sub ILogManager_Configure(configFilePath As String)

    Dim config As String
    config = Dir(configFilePath & "\log4vba.config")
    If config = "" Then config = Dir(configFilePath & "\log4vba.xml")
    If config = "" Then config = Dir(configFilePath & "\*.log4vba")
    If Dir(configFilePath & "\" & config) <> "" Then
        Dim configuror As New XmlConfigurator
        configuror.File = configFilePath & "\" & config
        m_configured = configuror.Configure(Me)
        
    End If
End Sub

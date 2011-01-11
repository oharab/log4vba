VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppenderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : AppenderFactory
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Returns a new IAppender configured by the passed in xml config.
' Note      : As new appenders are created, this will need editing to configure them.
'---------------------------------------------------------------------------------------

Option Explicit

Private log As New LogLog

'---------------------------------------------------------------------------------------
' Procedure : Build
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Build(ByVal configNode As Object) As IAppender
    Dim app As Object, nd As Object
    On Error GoTo Build_Error
    log.Dbug "Building " & configNode.Attributes.getNamedItem("type").Text

    'Create the appender & apply specific configuration
    Select Case configNode.Attributes.getNamedItem("type").Text
    Case "ConsoleAppender"
        Set app = New ConsoleAppender
    Case "OutputDebugStringAppender"
        Set app = New OutputDebugStringAppender
    Case "MsgboxAppender"
        Set app = New MsgboxAppender
    Case "EmailAppender"
        Set app = New EmailAppender

        Set nd = configNode.selectSingleNode("server")
        If nd Is Nothing Then Err.Raise vbObjectError, "AppenderFactory", "Email Server not set in config."
        app.Server = nd.Attributes.getNamedItem("name").Text
        app.Port = nd.Attributes.getNamedItem("port").Text

        Set nd = configNode.selectSingleNode("subject")
        If nd Is Nothing Then Err.Raise vbObjectError, "AppenderFactory", "Email subject not set in config."
        app.Subject = nd.Attributes.getNamedItem("value").Text

        Set nd = configNode.selectSingleNode("sender")
        If nd Is Nothing Then Err.Raise vbObjectError, "AppenderFactory", "Email sender not set in config."
        app.Sender = nd.Attributes.getNamedItem("value").Text

        Set nd = configNode.selectSingleNode("receiver")
        If nd Is Nothing Then Err.Raise vbObjectError, "AppenderFactory", "Email receiver not set in config."
        app.Receiver = nd.Attributes.getNamedItem("value").Text
    Case "FileAppender"
        Set app = New FileAppender

        Set nd = configNode.selectSingleNode("file")
        If nd Is Nothing Then Err.Raise vbObjectError, "AppenderFactory", "FileAppender output not set in config."
        app.File = nd.Attributes.getNamedItem("value").Text

        Set nd = configNode.selectSingleNode("appendToFile")
        If Not nd Is Nothing Then
            Select Case nd.Attributes.getNamedItem("value").Text
            Case "false"
                app.AppendToFile = False
            Case "true"
                app.AppendToFile = True
            Case Else
                Err.Raise vbObjectError, "AppenderFactory", "Invalid value for appendToFile."
            End Select
        End If

        Set nd = configNode.selectSingleNode("lockingModel")
        If Not nd Is Nothing Then
            Select Case nd.Attributes.getNamedItem("value").Text
            Case "MinimalLock"
                app.LockingModel = LockingModel.MinimalLock
            Case "ExclusiveLock"
                app.LockingModel = LockingModel.ExclusiveLock
            Case Else
                Err.Raise vbObjectError, "AppenderFactory", "Invalid value for LockingModel."
            End Select
        End If

    Case Else
        Err.Raise vbObjectError, "AppenderFactory", "Unable to find Appender type " & configNode.Attributes("type").Text
    End Select

    app.IAppender_Name = configNode.Attributes.getNamedItem("name").Text
    If configNode.hasChildNodes Then
        'Set the individual threshold of this appender, if specified
        If Not configNode.selectSingleNode("threshold") Is Nothing Then
            Set app.Threshold = Levels.FromString(configNode.selectSingleNode("threshold").Attributes.getNamedItem("value").Text)
        End If
    End If
    Set Build = app


    On Error Resume Next
    Exit Function

Build_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in AppenderFactory.Build. Returning null appender", Err.Number, Err.Description, Err.Source
        Set Build = New NullAppender
    End Select
    Resume Build_Error
End Function

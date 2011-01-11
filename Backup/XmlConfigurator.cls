VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlConfigurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private log As New LogLog

Private m_file As String
Private m_logmgr As clsLogManager
Public Property Let File(sFile As String)
    m_file = sFile
End Property

Public Function Configure(logmgr As clsLogManager) As Boolean
    Dim doc As Object
    Dim conf As Object
    Dim appFactory As New AppenderFactory
    On Error GoTo Configure_Error


    Set doc = VBA.CreateObject("MSXML.DOMDocument")
    log.Info "Configuring LogManager"
    Set m_logmgr = logmgr
    If doc.Load(m_file) Then
        Set conf = doc.selectSingleNode("log4vba")
        If conf.Attributes.getNamedItem("debug") Is Nothing Then
        Else
            DBGLog4VBA = conf.Attributes.getNamedItem("debug").Text = "true"
        End If
        If conf Is Nothing Then
        Else
            Dim appNode As Object
            logmgr.Appenders.RemoveAll
            logmgr.Loggers.RemoveAll
            
            For Each appNode In conf.selectNodes("appender")
                logmgr.Appenders.Add appNode.Attributes.getNamedItem("name").Text, appFactory.Build(appNode)
            Next appNode
            Dim loggerNode As Object
            For Each loggerNode In conf.selectNodes("logger")
                logmgr.Loggers.Add loggerNode.Attributes.getNamedItem("name").Text, buildLogger(loggerNode)
            Next loggerNode
            If conf.selectSingleNode("root") Is Nothing Then
            Else
                prepareRootLogger conf.selectSingleNode("root")
            End If
        End If
        Configure = True
    Else
        log.Fatal "Error loading config file """ & m_file & """."
    End If
    
Configure_Exit:
    On Error Resume Next
           Exit Function

Configure_Error:
    Select Case Err
        Case Else
            log.Error "Unhandled Error in XmlConfigurator.Configure", Err.Number, Err.Description, Err.Source
    End Select
    Resume Configure_Exit
    Resume
End Function

Private Function buildLogger(conf As Object, Optional name As String) As ILog
    Dim l As New clsLogImpl
    If name = "" Then
        l.name = conf.Attributes.getNamedItem("name").Text
    Else
        l.name = name
    End If
    log.Dbug "Build Logger " & l.name
    Dim levelNode As Object
    Set levelNode = conf.selectSingleNode("level")
    If Not levelNode Is Nothing Then
        Set l.Level = Levels.FromString(levelNode.Attributes.getNamedItem("value").Text)
    End If
    Dim appenderref As Object
    For Each appenderref In conf.selectNodes("appender-ref")
        Dim ref As String
        ref = appenderref.Attributes.getNamedItem("ref").Text
        If m_logmgr.Appenders.Exists(ref) Then
            l.AddAppender m_logmgr.Appenders(ref)
        Else
            Err.Raise vbObjectError, "XmlConfigurator", "Unknown appender-ref "
        End If
    Next appenderref
    Set buildLogger = l
End Function

Private Sub prepareRootLogger(conf As Object)
    Dim levelNode As Object
    log.Dbug "Preparing root logger."
    Set levelNode = conf.selectSingleNode("level")
    If Not levelNode Is Nothing Then
        Set m_logmgr.RootLevel = Levels.FromString(levelNode.Attributes.getNamedItem("value").Text)
    End If
    Dim appenderref As Object
    For Each appenderref In conf.selectNodes("appender-ref")
        Dim ref As String
        ref = appenderref.Attributes.getNamedItem("ref").Text
        If m_logmgr.Appenders.Exists(ref) Then
            m_logmgr.AddRootAppender m_logmgr.Appenders(ref)
        Else
            Err.Raise vbObjectError, "XmlConfigurator", "Unknown appender-ref "
        End If
    Next appenderref
End Sub

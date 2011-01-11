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

Public Sub Configure(logmgr As clsLogManager)
    Dim doc As New DOMDocument
    Dim conf As IXMLDOMNode
    Dim appFactory As New AppenderFactory
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
            Dim appNode As IXMLDOMNode
            For Each appNode In conf.selectNodes("appender")
                logmgr.Appenders.Add appNode.Attributes.getNamedItem("name").Text, appFactory.Build(appNode)
            Next appNode
            Dim loggerNode As IXMLDOMNode
            For Each loggerNode In conf.selectNodes("logger")
                logmgr.Loggers.Add loggerNode.Attributes.getNamedItem("name").Text, buildLogger(loggerNode)
            Next loggerNode
            If conf.selectSingleNode("root") Is Nothing Then
            Else
                prepareRootLogger conf.selectSingleNode("root")
            End If
        End If
    Else
        log.Fatal "Error loading config file """ & m_file & """."
    End If
End Sub

Private Function buildLogger(conf As IXMLDOMNode, Optional Name As String) As ILog
    Dim l As New clsLogImpl
    If Name = "" Then
        l.Name = conf.Attributes.getNamedItem("name").Text
    Else
        l.Name = Name
    End If
    log.Dbug "Build Logger " & l.Name
    Dim levelNode As IXMLDOMNode
    Set levelNode = conf.selectSingleNode("level")
    If Not levelNode Is Nothing Then
        Set l.Level = Levels.FromString(levelNode.Attributes.getNamedItem("value").Text)
    End If
    Dim appenderref As IXMLDOMNode
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

Private Sub prepareRootLogger(conf As IXMLDOMNode)
    Dim levelNode As IXMLDOMNode
    log.Dbug "Preparing root logger."
    Set levelNode = conf.selectSingleNode("level")
    If Not levelNode Is Nothing Then
        Set m_logmgr.RootLevel = Levels.FromString(levelNode.Attributes.getNamedItem("value").Text)
    End If
    Dim appenderref As IXMLDOMNode
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

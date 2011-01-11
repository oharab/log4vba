VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAppender


Private log As ILog

Private m_name As String
Private m_threshold As Level
Private m_server As String
Private m_port As Integer
Private m_subject As String
Private m_sender As String
Private m_receiver As String


Private Sub Class_Initialize()
    Set m_threshold = Levels.Dbug
    m_name = "EmailAppender"
End Sub

Public Property Get IAppender_Name() As String
    IAppender_Name = m_name
End Property
Public Property Let IAppender_Name(sName As String)
    m_name = sName
End Property

Public Sub IAppender_DoAppend(oLoggingEvent As LoggingEvent)
    If oLoggingEvent.Level.value >= m_threshold.value Then
        Dim msg As Object
        Dim conf As Object
        Dim flds As Variant
        Set msg = CreateObject("CDO.Message")
        Set conf = CreateObject("CDO.Configuration")
        conf.Load -1    ' CDO Source Defaults
        Set flds = conf.fields
        With flds
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = m_server
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = m_port
            .Update
        End With



        With msg
            Set .Configuration = conf
            .To = m_receiver
            .CC = ""
            .BCC = ""
            .From = m_sender
            .Subject = m_subject
            Dim strBody As String
            With oLoggingEvent
                strBody = "<h1>[" & .Level.Name & "]" & vbTab & .LoggerName & _
                          "<table border=""1"">" & _
                          "<tr>" & _
                          "<td>Time:</td><td>" & .TimeStamp & "</td>" & _
                          "</tr>" & _
                          "<tr>" & _
                          "<td>Message</td><td>" & oLoggingEvent.Message & "</td>" & _
                          "</tr>" & _
                          "<tr>" & _
                          "<td>Machine Name:</td><td>" & .MachineName & "</td>" & _
                          "</tr>" & _
                          "<tr>" & _
                          "<td>Network User:</td><td>" & .NetworkUser & "</td>" & _
                          "</tr>" & _
                          "<tr>" & _
                          "<td>IP Adress:</td><td>" & .IPAddress & "</td>" & _
                          "</tr>"
                If .ErrorNumber <> 0 Then
                    strBody = strBody & "<tr>" & _
                              "<td>Error:</td><td>" & .ErrorNumber & "</td>" & _
                              "</tr>" & _
                              "<tr>" & _
                              "<td>Error Description:</td><td>" & .ErrorDescription & "</td>" & _
                              "</tr>" & _
                              "<tr>" & _
                              "<td>Error Source:</td><td>" & .ErrorSource & "</td>" & _
                              "</tr>"
                End If
                strBody = strBody & "</table>"

            End With

            .HTMLBody = strBody
            .send
        End With


    End If
End Sub
Public Property Get Threshold() As Level
    Set Threshold = m_threshold
End Property
Public Property Set Threshold(oThreshold As Level)
    Set m_threshold = oThreshold
End Property

Public Property Let Server(sServer As String)
    m_server = sServer
End Property
Public Property Let Port(iPort As Integer)
    m_port = iPort
End Property
Public Property Let Subject(sSubject As String)
    m_subject = sSubject
End Property
Public Property Let Sender(sSender As String)
    m_sender = sSender
End Property
Public Property Let Receiver(sReceiver As String)
    m_receiver = sReceiver
End Property

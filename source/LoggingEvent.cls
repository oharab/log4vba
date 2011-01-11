VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggingEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LoggerName As String
Public Level As Level
Public UserName As String
Public TimeStamp As Date
Public NetworkUser As String
Public MachineName As String
Public IPAddress As String
Public Message As String
Public ErrorNumber As Integer
Public ErrorDescription As String
Public ErrorSource As String



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get name() As String
End Property
Public Property Let name(ByVal sName As String)
End Property

Public Sub DoAppend(ByVal LoggingEvent As LoggingEvent)
End Sub


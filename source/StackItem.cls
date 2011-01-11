VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : StackItem
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------

Private log As ILog

Private m_objPrevious As StackItem
Private m_vData As Variant
Private m_dtTimeStamp As Long

Private Sub Class_Initialize()
    m_dtTimeStamp = (CLng(Date) + Timer) * 1000
End Sub

Private Sub Class_Terminate()

    Set m_objPrevious = Nothing
    Set log = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Previous
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Previous item in the stack.
'---------------------------------------------------------------------------------------
'
Public Property Get Previous() As StackItem

    On Error GoTo Previous_Error
    Set Previous = m_objPrevious

    On Error Resume Next
    Exit Property

Previous_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.GetPrevious", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Previous_Error

End Property
Public Property Set Previous(objPrevious As StackItem)

    On Error GoTo Previous_Error
    Set m_objPrevious = objPrevious

    On Error Resume Next
    Exit Property

Previous_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.SetPrevious", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Previous_Error

End Property

'---------------------------------------------------------------------------------------
' Procedure : Data
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Data to be stored in stack.
'---------------------------------------------------------------------------------------
'
Public Property Get Data() As Variant

    On Error GoTo Data_Error
    If IsObject(m_vData) Then
        Set Data = m_vData
    Else
        Data = m_vData
    End If

    On Error Resume Next
    Exit Property

Data_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.GetData", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Data_Error

End Property
Public Property Let Data(ByVal vData As Variant)

    On Error GoTo Data_Error
    m_vData = vData

    On Error Resume Next
    Exit Property

Data_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.Data", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Data_Error

End Property
Public Property Set Data(ByVal vData As Object)

    On Error GoTo Data_Error
    Set m_vData = vData

    On Error Resume Next
    Exit Property

Data_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.Data", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Data_Error

End Property


'---------------------------------------------------------------------------------------
' Procedure : TimeStamp
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TimeStamp() As Long

    On Error GoTo TimeStamp_Error
    TimeStamp = m_dtTimeStamp


    On Error Resume Next
    Exit Property

TimeStamp_Error:
    Select Case Err
    Case Else
        log.Error "Unhandled Error in StackItem.TimeStamp", Err
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume TimeStamp_Error

End Property

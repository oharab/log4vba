VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLevels
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Faux enumeration allowing statically typed objects to be returned.
'---------------------------------------------------------------------------------------
Option Explicit

Private log As New LogLog
Private m_Dbug As Level
Private m_Info As Level
Private m_Warn As Level
Private m_Error As Level
Private m_Fatal As Level
Private m_Off As Level

'---------------------------------------------------------------------------------------
' Procedure : Dbug
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Lowest granulation of logging.
'---------------------------------------------------------------------------------------
'
Public Function Dbug() As Level
    On Error GoTo Dbug_Error
    If m_Dbug Is Nothing Then
        Set m_Dbug = New Level
        With m_Dbug
            .DisplayName = "Debug"
            .name = "DEBUG"
            .value = 1
        End With
    End If
    Set Dbug = m_Dbug

Dbug_Exit:
    On Error Resume Next
    Exit Function

Dbug_Error:
    Select Case Err
    Case Else
        Set Dbug = Me.Off
        log.Error "Error Creating Dbug Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Dbug_Exit
End Function


'---------------------------------------------------------------------------------------
' Procedure : Info
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : General information that may be useful for problem solving
'---------------------------------------------------------------------------------------
'
Public Function Info() As Level
    On Error GoTo Info_Error
    If m_Info Is Nothing Then
        Set m_Info = New Level
        With m_Info
            .DisplayName = "Info"
            .name = "INFO"
            .value = 2
        End With
    End If
    Set Info = m_Info


    On Error Resume Next
    Exit Function

Info_Exit:
    On Error Resume Next
    Exit Function

Info_Error:
    Select Case Err
    Case Else
        Set Info = Me.Off
        log.Error "Error Creating Info Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Info_Exit
End Function


'---------------------------------------------------------------------------------------
' Procedure : Warn
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Warnings which aren't errors
'---------------------------------------------------------------------------------------
'
Public Function Warn() As Level
    On Error GoTo Warn_Error
    If m_Warn Is Nothing Then
        Set m_Warn = New Level
        With m_Warn
            .DisplayName = "Warn"
            .name = "WARN"
            .value = 3
        End With
    End If
    Set Warn = m_Warn

Warn_Exit:
    On Error Resume Next
    Exit Function

Warn_Error:
    Select Case Err
    Case Else
        Set Warn = Me.Off
        log.Error "Error Creating Warn Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Warn_Exit
End Function


'---------------------------------------------------------------------------------------
' Procedure : Error
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Unexpected errors.
'---------------------------------------------------------------------------------------
'
Public Function Error() As Level
    On Error GoTo Error_Error
    If m_Error Is Nothing Then
        Set m_Error = New Level
        With m_Error
            .DisplayName = "Error"
            .name = "ERROR"
            .value = 4
        End With
    End If
    Set Error = m_Error

Error_Exit:
    On Error Resume Next
    Exit Function

Error_Error:
    Select Case Err
    Case Else
        Set Error = Me.Off
        log.Error "Error Creating Error Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Error_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Fatal
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Error severe enough to render the operation uncontinuable
'---------------------------------------------------------------------------------------
'
Public Function Fatal() As Level
    On Error GoTo Fatal_Error
    If m_Fatal Is Nothing Then
        Set m_Fatal = New Level
        With m_Fatal
            .DisplayName = "Fatal"
            .name = "Fatal"
            .value = 5
        End With
    End If
    Set Fatal = m_Fatal


Fatal_Exit:
    On Error Resume Next
    Exit Function

Fatal_Error:
    Select Case Err
    Case Else
        Set Fatal = Me.Off
        log.Error "Error Creating Fatal Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Fatal_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Off
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : No logging should be performed
'---------------------------------------------------------------------------------------
'
Public Function Off() As Level
    On Error GoTo Off_Error
    If m_Off Is Nothing Then
        Set m_Off = New Level
    End If
    Set Off = m_Off


Off_Exit:
    On Error Resume Next
    Exit Function

Off_Error:
    Select Case Err
    Case Else
        Set Off = Me.Off
        log.Error "Error Creating Off Level", Err.Number, Err.Description, Err.Source
    End Select
    Resume Off_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : FromString
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Converts as string to an appropriate Level instances
'---------------------------------------------------------------------------------------
'
Public Function FromString(ByVal sLevel As String) As Level
    On Error GoTo FromString_Error
    Select Case UCase(sLevel)
    Case "DEBUG"
        Set FromString = Me.Dbug
    Case "INFO"
        Set FromString = Me.Info
    Case "WARN"
        Set FromString = Me.Warn
    Case "ERROR"
        Set FromString = Me.Error
    Case "FATAL"
        Set FromString = Me.Fatal
    Case "OFF"
        Set FromString = Me.Off
    Case Else
        Err.Raise vbObjectError, "clsLevels", "Unknown Level: " & sLevel
        Set FromString = Me.Off
    End Select

FromString_Exit:
    On Error Resume Next
    Exit Function

FromString_Error:
    Select Case Err
    Case Else
        'been passed an invalid string. Just log it and return an OFF
        log.Error Err.Description, Err.Number, Err.Description, Err.Source
        Set FromString = Me.Off
    End Select
    Resume FromString_Exit
End Function


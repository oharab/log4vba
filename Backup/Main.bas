Attribute VB_Name = "Main"
'---------------------------------------------------------------------------------------
' Module    : Main
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Workround to make instance available as global/static objects
'---------------------------------------------------------------------------------------
Option Explicit
Public DBGLog4VBA As Boolean
Private m_logmanager As ILogManager
Private m_levels As clsLevels


'---------------------------------------------------------------------------------------
' Procedure : LogManager
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Returns the LogManager which creates new configured instances of ILog
'---------------------------------------------------------------------------------------
'
Public Function LogManager() As ILogManager
    On Error GoTo LogManager_Error

    If m_logmanager Is Nothing Then
        Set m_logmanager = New clsLogManager
    End If
    Set LogManager = m_logmanager


    On Error Resume Next
    Exit Function

LogManager_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume LogManager_Error
    Resume
End Function

'---------------------------------------------------------------------------------------
' Procedure : Levels
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Allows access to a faux enumeration of objects.
'---------------------------------------------------------------------------------------
'
Public Function Levels() As clsLevels
    On Error GoTo Levels_Error
    If m_levels Is Nothing Then
        Set m_levels = New clsLevels
    End If
    Set Levels = m_levels

    On Error Resume Next
    Exit Function

Levels_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume Levels_Error
End Function


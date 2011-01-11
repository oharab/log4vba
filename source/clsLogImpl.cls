VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLogImpl
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : The implementation of ILog that is configured by the Log manager.
'---------------------------------------------------------------------------------------
Option Explicit

Implements ILog

Private log As ILog


Private m_name As String
Private m_appenders As HashTable
Private m_level As Level

'******** Code Start ********
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const AF_INET = 2

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(MAX_WSADescription) As Byte
    szSystemStatus(MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

' returns the standard host name for the local machine
Private Declare Function apiGetHostName _
                          Lib "wsock32" Alias "gethostname" _
                              (ByVal Name As String, _
                               ByVal nameLen As Long) _
                               As Long

' retrieves host information corresponding to a host name
' from a host database
Private Declare Function apiGetHostByName _
                          Lib "wsock32" Alias "gethostbyname" _
                              (ByVal hostname As String) _
                              As Long

' retrieves the host information corresponding to a network address
Private Declare Function apiGetHostByAddress _
                          Lib "wsock32" Alias "gethostbyaddr" _
                              (addr As Long, _
                               ByVal dwLen As Long, _
                               ByVal dwType As Long) _
                               As Long

' moves memory either forward or backward, aligned or unaligned,
' in 4-byte blocks, followed by any remaining bytes
Private Declare Sub sapiCopyMem _
                     Lib "kernel32" Alias "RtlMoveMemory" _
                         (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)

' converts a string containing an (Ipv4) Internet Protocol
' dotted address into a proper address for the IN_ADDR structure
Private Declare Function apiInetAddress _
                          Lib "wsock32" Alias "inet_addr" _
                              (ByVal cp As String) _
                              As Long

' function initiates use of Ws2_32.dll by a process
Private Declare Function apiWSAStartup _
                          Lib "wsock32" Alias "WSAStartup" _
                              (ByVal wVersionRequired As Integer, _
                               lpWsaData As WSADATA) _
                               As Long

Private Declare Function apilstrlen _
                          Lib "kernel32" Alias "lstrlen" _
                              (ByVal lpString As Long) _
                              As Long

Private Declare Function apilstrlenW _
                          Lib "kernel32" Alias "lstrlenW" _
                              (ByVal lpString As Long) _
                              As Long

' function terminates use of the Ws2_32.dll
Private Declare Function apiWSACleanup _
                          Lib "wsock32" Alias "WSACleanup" _
                              () As Long

Private Function fGetHostIPAddresses(strHostName As String) As Collection
'
' Resolves the English HostName and returns
' a collection with all the IPs bound to the card
'
    On Error GoTo ErrHandler
    Dim lngRet As Long
    Dim lpHostEnt As HOSTENT
    Dim strOut As String
    Dim colOut As Collection
    Dim lngIPAddr As Long
    Dim abytIPs() As Byte
    Dim I As Integer

    Set colOut = New Collection

    If fInitializeSockets() Then
        strOut = String$(255, vbNullChar)
        lngRet = apiGetHostByName(strHostName)
        If lngRet Then

            Call sapiCopyMem( _
                 lpHostEnt, _
                 ByVal lngRet, _
                 Len(lpHostEnt))

            Call sapiCopyMem( _
                 lngIPAddr, _
                 ByVal lpHostEnt.hAddrList, _
                 Len(lngIPAddr))

            Do While (lngIPAddr)
                With lpHostEnt
                    ReDim abytIPs(0 To .hLength - 1)
                    strOut = vbNullString
                    Call sapiCopyMem( _
                         abytIPs(0), _
                         ByVal lngIPAddr, _
                         .hLength)
                    For I = 0 To .hLength - 1
                        strOut = strOut & abytIPs(I) & "."
                    Next
                    strOut = Left$(strOut, Len(strOut) - 1)
                    .hAddrList = .hAddrList + Len(.hAddrList)
                    Call sapiCopyMem( _
                         lngIPAddr, _
                         ByVal lpHostEnt.hAddrList, _
                         Len(lngIPAddr))
                    If Len(Trim$(strOut)) Then colOut.Add strOut
                End With
            Loop
        End If
    End If
    Set fGetHostIPAddresses = colOut
ExitHere:
    Call apiWSACleanup
    Set colOut = Nothing
    Exit Function
ErrHandler:
    With Err
        MsgBox "Error: " & .Number & vbCrLf & .Description, _
               vbOKOnly Or vbCritical, _
               .Source
    End With
    Resume ExitHere
End Function

Private Function fGetHostName(strIPAddress As String) As String
'
' Looks up a given IP address and returns the
' machine name it's bound to
'
    On Error GoTo ErrHandler
    Dim lngRet As Long
    Dim lpAddress As Long
    Dim strOut As String
    Dim lpHostEnt As HOSTENT

    If fInitializeSockets() Then
        lpAddress = apiInetAddress(strIPAddress)
        lngRet = apiGetHostByAddress(lpAddress, 4, AF_INET)
        If lngRet Then
            Call sapiCopyMem( _
                 lpHostEnt, _
                 ByVal lngRet, _
                 Len(lpHostEnt))
            fGetHostName = fStrFromPtr(lpHostEnt.hName, False)
        End If
    End If
ExitHere:
    Call apiWSACleanup
    Exit Function
ErrHandler:
    With Err
        MsgBox "Error: " & .Number & vbCrLf & .Description, _
               vbOKOnly Or vbCritical, _
               .Source
    End With
    Resume ExitHere
End Function

Private Function fInitializeSockets() As Boolean
    Dim lpWsaData As WSADATA
    Dim wVersionRequired As Integer

    wVersionRequired = fMakeWord(2, 2)
    fInitializeSockets = ( _
                         apiWSAStartup(wVersionRequired, lpWsaData) = 0)

End Function

Private Function fMakeWord( _
        ByVal low As Integer, _
        ByVal hi As Integer) _
        As Integer
    Dim intOut As Integer
    Call sapiCopyMem( _
         ByVal VarPtr(intOut) + 1, _
         ByVal VarPtr(hi), _
         1)
    Call sapiCopyMem( _
         ByVal VarPtr(intOut), _
         ByVal VarPtr(low), _
         1)
    fMakeWord = intOut
End Function

Private Function fStrFromPtr( _
        pBuf As Long, _
        Optional blnIsUnicode As Boolean) _
        As String
    Dim lngLen As Long
    Dim abytBuf() As Byte

    If blnIsUnicode Then
        lngLen = apilstrlenW(pBuf) * 2
    Else
        lngLen = apilstrlen(pBuf)
    End If
    ' if it's not a ZLS
    If lngLen Then
        ReDim abytBuf(lngLen)
        ' return the buffer
        If blnIsUnicode Then
            'blnIsUnicode is True not tested
            Call sapiCopyMem(abytBuf(0), ByVal pBuf, lngLen)
            fStrFromPtr = abytBuf
        Else
            ReDim Preserve abytBuf(UBound(abytBuf) - 1)
            Call sapiCopyMem(abytBuf(0), ByVal pBuf, lngLen)
            fStrFromPtr = StrConv(abytBuf, vbUnicode)
        End If
    End If
End Function
' ******** Code End ********




Private Sub Class_Initialize()
    Set m_appenders = New HashTable
    Set m_level = Levels.Off
End Sub

Private Sub Class_Terminate()
    m_appenders.RemoveAll
    Set m_appenders = Nothing
    Set m_level = Nothing
    Set log = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ILog_Info
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Logs an "Info" message to all appenders.
'---------------------------------------------------------------------------------------
'
Public Sub ILog_Info(sMessage As String)
    On Error GoTo ILog_Info_Error
    If Me.ILog_IsInfoEnabled Then
        Dim key As Variant, a As IAppender, e As LoggingEvent
        Set e = BuildLoggingEvent(sMessage, Levels.Info)
        For Each key In m_appenders.Keys
            Set a = m_appenders(CStr(key))
            a.DoAppend e
        Next key
    End If


    On Error Resume Next
    Exit Sub

ILog_Info_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume ILog_Info_Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ILog_Dbug
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Logs a "Debug" message to all appenders
'---------------------------------------------------------------------------------------
'
Public Sub ILog_Dbug(sMessage As String)
    On Error GoTo ILog_Dbug_Error
    If Me.ILog_IsDbugEnabled Then
        Dim key As Variant, a As IAppender, e As LoggingEvent
        Set e = BuildLoggingEvent(sMessage, Levels.Dbug)
        For Each key In m_appenders.Keys
            Set a = m_appenders(CStr(key))
            a.DoAppend e
        Next key
    End If


    On Error Resume Next
    Exit Sub

ILog_Dbug_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume ILog_Dbug_Error

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ILog_Warn
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Logs a "Warn" message to all appenders
'---------------------------------------------------------------------------------------
'
Public Sub ILog_Warn(sMessage As String)
    On Error GoTo ILog_Warn_Error
    If Me.ILog_IsWarnEnabled Then
        Dim key As Variant, a As IAppender, e As LoggingEvent
        Set e = BuildLoggingEvent(sMessage, Levels.Warn)
        For Each key In m_appenders.Keys
            Set a = m_appenders(CStr(key))
            a.DoAppend e
        Next key
    End If


    On Error Resume Next
    Exit Sub

ILog_Warn_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume ILog_Warn_Error

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ILog_Error
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Logs an "Error" message to all appenders
'---------------------------------------------------------------------------------------
'
Public Sub ILog_Error(sMessage As String, Exception As ErrObject)
    On Error GoTo ILog_Error_Error
    If Me.ILog_IsErrorEnabled Then
        Dim key As Variant, a As IAppender, e As LoggingEvent
        Set e = BuildLoggingEvent(sMessage, Levels.Error, Exception)
        For Each key In m_appenders.Keys
            Set a = m_appenders(CStr(key))
            a.DoAppend e
        Next key
    End If


    On Error Resume Next
    Exit Sub

ILog_Error_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume ILog_Error_Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ILog_Fatal
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Logs a "Fatal" message to all appenders
'---------------------------------------------------------------------------------------
'
Public Sub ILog_Fatal(sMessage As String, Optional Exception As ErrObject)
    On Error GoTo ILog_Fatal_Error
    If Me.ILog_IsFatalEnabled Then
        Dim key As Variant, a As IAppender, e As LoggingEvent
        Set e = BuildLoggingEvent(sMessage, Levels.Fatal, Exception)
        For Each key In m_appenders.Keys
            Set a = m_appenders(CStr(key))
            a.DoAppend e
        Next key
    End If


    On Error Resume Next
    Exit Sub

ILog_Fatal_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume ILog_Fatal_Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BuildLoggingEvent
' Author    : bpo@robotparade.co.uk
' Date      : 22/11/2010
' Purpose   : Builds a Logging Event from the message and and global data.
'---------------------------------------------------------------------------------------
'
Private Function BuildLoggingEvent(sMessage As String, oLevel As Level, Optional Exception As ErrObject) As LoggingEvent
    On Error GoTo BuildLoggingEvent_Error
    Set BuildLoggingEvent = New LoggingEvent
    With BuildLoggingEvent
        .LoggerName = Me.Name
        .Message = sMessage
        Set .Level = oLevel
        If Not Exception Is Nothing Then
            .ErrorNumber = Exception.Number
            .ErrorDescription = Exception.Description
            .ErrorSource = Exception.Source
        End If
        .MachineName = Environ$("computername")
        .NetworkUser = Environ$("username")
        .TimeStamp = Now
        .IPAddress = fGetHostIPAddresses(.MachineName)(1)
    End With


    On Error Resume Next
    Exit Function

BuildLoggingEvent_Error:
    Select Case Err
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpContext
    End Select
    Resume BuildLoggingEvent_Error

End Function

Public Property Let ILog_Name(sName As String)
    m_name = sName
End Property
Public Property Get ILog_Name() As String
    ILog_Name = m_name
End Property

Public Property Let Name(sName As String)
    m_name = sName
End Property
Public Property Get Name() As String
    Name = m_name
End Property


Public Function ILog_IsInfoEnabled() As Boolean
    ILog_IsInfoEnabled = Me.Level.value <= Levels.Info.value
End Function
Public Function ILog_IsDbugEnabled() As Boolean
    ILog_IsDbugEnabled = Me.Level.value <= Levels.Dbug.value
End Function
Public Function ILog_IsErrorEnabled() As Boolean
    ILog_IsErrorEnabled = Me.Level.value <= Levels.Error.value
End Function
Public Function ILog_IsWarnEnabled() As Boolean
    ILog_IsWarnEnabled = Me.Level.value <= Levels.Warn.value
End Function
Public Function ILog_IsFatalEnabled() As Boolean
    ILog_IsFatalEnabled = Me.Level.value <= Levels.Fatal.value
End Function


Public Property Get Level() As Level
    Set Level = m_level
End Property
Public Property Set Level(oLevel As Level)
    Set m_level = oLevel
End Property

Public Sub AddAppender(appender As IAppender)
    m_appenders.Add appender.Name, appender
End Sub



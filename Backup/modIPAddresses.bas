Attribute VB_Name = "modIPAddresses"
Option Explicit

' ******** Code Start ********
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
    (ByVal name As String, _
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
    
Function fGetHostIPAddresses(strHostName As String) As Collection
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
Dim i As Integer

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
                    For i = 0 To .hLength - 1
                        strOut = strOut & abytIPs(i) & "."
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
 
Function fGetHostName(strIPAddress As String) As String
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



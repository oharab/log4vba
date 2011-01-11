Attribute VB_Name = "ImportExcel"
Option Explicit


Private Const BACKUP_LOCATION As String = "O:\Common\dev\log4vba\Backup\"

Public Sub Import()
Dim fso As Object, fld As Object, fl As Object
Dim p As VBProject

    On Error GoTo Import_Error
    Set p = ThisWorkbook.VBProject
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    Set fld = fso.getfolder(BACKUP_LOCATION)
    For Each fl In fld.Files
        Dim componentName As String, compontentType As Long
        Dim extPosition As Double, ext As String
        extPosition = VBA.InStrRev(fl.name, ".")
        ext = VBA.Mid(fl.name, extPosition + 1)
        Select Case ext
            Case "bas"
                compontentType = vbext_ct_StdModule
            Case "cls"
                compontentType = vbext_ct_ClassModule
        End Select
        componentName = Mid(fl.name, 1, extPosition - 1)
        Debug.Print fl.name, ext, componentName
        Dim vb As VBComponent
        Set vb = p.VBComponents.Import(fl.Path)
        'vb.Type = compontentType
        vb.name = componentName
    Next fl



Import_Exit:
    Set fl = Nothing
    Set fld = Nothing
    Set fso = Nothing
    Set p = Nothing
    On Error Resume Next
           Exit Sub

Import_Error:
    Select Case Err
        Case Else
            Debug.Print Err.Description
    End Select
    Resume Import_Exit
    Resume

End Sub

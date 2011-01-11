Attribute VB_Name = "BackupCode"
Option Explicit

Private Const BACKUP_LOCATION As String = "O:\Common\dev\log4vba\Backup\"

Public Function Backup()
    Dim vb As Object
    For Each vb In ThisWorkbook.VBProject.VBE.ActiveVBProject.VBComponents
        Debug.Print vb.name
        Select Case vb.Type
        Case 2    'vbext_ct_ClassModule
            vb.Export BACKUP_LOCATION & vb.name & ".cls"
        Case 1    'vbext_ct_StdModule
            vb.Export BACKUP_LOCATION & vb.name & ".bas"
        End Select
    Next vb
End Function

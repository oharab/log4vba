Attribute VB_Name = "Backup"
Option Compare Database
Option Explicit

Private Const BACKUP_LOCATION As String = "O:\Common\dev\log4vba\Backup\"

Public Function Backup()
    Dim cnt As DAO.Container, doc As DAO.Document
    For Each cnt In CurrentDb.Containers
        For Each doc In cnt.Documents
            Select Case cnt.name
            Case "Forms"
                Application.SaveAsText acForm, doc.name, BACKUP_LOCATION & doc.name & ".frm"
            Case "Modules"
                'Application.SaveAsText acModule, doc.name, BACKUP_LOCATION & doc.name & ".bas"
            Case "Relationships"
            Case "Reports"
                Application.SaveAsText acReport, doc.name, BACKUP_LOCATION & doc.name & ".rpt"
            Case "Scripts"
                Application.SaveAsText acMacro, doc.name, BACKUP_LOCATION & doc.name & ".mac"
            Case "Tables"

            Case Else
                'Debug.Print cnt.name, doc.name
            End Select
            'Debug.Print vbTab & doc.Name
        Next doc
    Next cnt
    Dim vb As Object
    For Each vb In VBE.ActiveVBProject.VBComponents
        Debug.Print vb.name
        Select Case vb.Type
        Case 2    'vbext_ct_ClassModule
            vb.Export BACKUP_LOCATION & vb.name & ".cls"
        Case 1    'vbext_ct_StdModule
            vb.Export BACKUP_LOCATION & vb.name & ".bas"
        End Select
    Next vb
End Function

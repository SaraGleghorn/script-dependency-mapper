VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn_findfiles_Click()

If populateFileList() = True Then
    MsgBox DCount("*", "script_filepath") & " script files found", vbOKOnly
Else
    MsgBox "There were errors finding SQL files. " _
    & "Check the VBA (Alt + F11) immediate window for error messages.", _
    vbCritical, "Search Failed"
End If

End Sub

Private Sub btn_getdependencies_Click()

Call sqlDependencyMapper
MsgBox "Done", vbOKOnly
DoCmd.OpenTable ("table_usage")

End Sub

Private Sub btn_searchfolders_Click()

DoCmd.OpenTable ("script_folder")


End Sub

Private Sub form_load()
' Obtain the version information upon load

Dim db  As DAO.Database
Dim rs  As DAO.Recordset
Dim strQry  As String

Set db = CurrentDb
strQry = "SELECT * FROM version_control WHERE version = (SELECT MAX(version) FROM version_control)"
Set rs = db.OpenRecordset(strQry)

If rs.RecordCount = 0 Then
    lbl_version.Caption = "Error retrieving version control info"
Else
    rs.MoveFirst
    lbl_version.Caption = rs!Version & " | " & rs!author & " | " & rs!release_date
End If

End Sub

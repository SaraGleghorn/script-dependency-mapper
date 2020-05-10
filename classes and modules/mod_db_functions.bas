Attribute VB_Name = "mod_db_functions"
Option Compare Database
Option Explicit

Function databaseObjectExists( _
    strObjectName As String, _
    Optional strObjectType As String, _
    Optional db As DAO.Database _
    ) As Boolean
' Purpose: ********************************************************************
' Returns true if 'strObjectName' exists as an Access object within 'db'
'   This includes tables and queries
' Requirements: None
' Version Control:
' Vers      Author          Date        Change
' 1.0.0     Sara Gleghorn   23/03/2020  Original
' *****************************************************************************
'' Expected Parameters:
'Dim strObjectName  As String   ' The name of the object we're looking for
'Dim strObjectType  As String   ' The type of object ("Table", "Query", "Form", "Report")
'Dim db             As DAO.Database ' The database we're looking in.
'                                ' Defaults to this db if null
Definitions:
Dim strFnName       As String   ' The name of this function (for debugging messages)
Dim strSection      As String   ' The name of the strSection (for debugging messages)
Dim tbl             As TableDef ' For looping through tables
Dim qry             As QueryDef ' For looping through queries
Dim frm             As Form     ' For looping through forms
Dim rprt            As Report   ' For looping through reports

On Error GoTo ErrorHandler
strFnName = "databaseObjectExists"

SetDefaults: ' ----------------------------------------------------------------
strSection = "SetDefaults"

If db Is Nothing Then
    Set db = CurrentDb
End If

SkipToType: ' -----------------------------------------------------------------
strSection = "SkipToType"

Select Case strObjectType
    Case "Table", "table", "tbl"
        GoTo SearchTables
    Case "Query", "query", "qry"
        GoTo SearchQueries
    Case "Form", "form", "frm"
        GoTo SearchForms
    Case "Report", "report", "rprt"
        GoTo SearchReports
    Case ""
        Debug.Print Now() & " " _
            & strFnName & ": strObjectType is blank. Searching all database objects."
    Case Else
        Debug.Print Now() & " " _
            & strFnName & ": strObjectType " & """" & strObjectType & """" _
            & " is not known. Searching all database objects."
End Select

SearchTables: ' ---------------------------------------------------------------
strSection = "SearchTables"

For Each tbl In db.TableDefs
    If tbl.Name = strObjectName Then
        databaseObjectExists = True
        Exit Function
    End If
Next

Select Case strObjectType
    Case "Table", "table", "tbl"
        Exit Function
End Select

SearchQueries: ' --------------------------------------------------------------
strSection = "SearchQueries"

For Each qry In db.QueryDefs
    If qry.Name = strObjectName Then
        databaseObjectExists = True
        Exit Function
    End If
Next

Select Case strObjectType
    Case "Query", "query", "qry"
        Exit Function
End Select

SearchForms: ' ----------------------------------------------------------------
strSection = "SearchForms"

For Each frm In Forms
    If frm.Name = strObjectName Then
        databaseObjectExists = True
        Exit Function
    End If
Next

Select Case strObjectType
    Case "Form", "form", "frm"
        Exit Function
End Select

SearchReports: ' --------------------------------------------------------------
strSection = "SearchReports"

For Each rprt In Reports
    If rprt.Name = strObjectName Then
        databaseObjectExists = True
        Exit Function
    End If
Next

Select Case strObjectType
    Case "Report", "report", "rprt"
        Exit Function
End Select

Cleanup: '---------------------------------------------------------------------
strSection = "Cleanup"
Debug.Print Now() & " " & strFnName & ": " _
    & "Object: " & strObjectName & " not found in any object type"
Exit Function

ErrorHandler:
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
End Function



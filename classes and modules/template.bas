Attribute VB_Name = "template"
Option Compare Database
Option Explicit

Function fn_template(foo As Variant) As Boolean
' Purpose: ********************************************************************
' A template, to keep the same error handler structure.
' Requirements:
'
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           dd/mm/yyyy  Original
' *****************************************************************************
' Expected Parameters:
'Dim Foo    As Bar  ' Description

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function (for debugging messages)
Dim strSection      As String           ' The name of the section (for debugging messages)

On Error GoTo ErrorHandler
CheckPrerequisites: ' ---------------------------------------------------------
strFnName = "fn_template"
strSection = "CheckPrerequisites"

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
fn_template = True
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function
End Function


Attribute VB_Name = "mod_errorhandlers"
Option Compare Database
Option Explicit

Function showErrorHandlerPopup( _
    strErrorFunction As String, _
    strErrorSection As String, _
    strErrorMessage As String, _
    Optional strActionMessage As String = vbNullString, _
    Optional vbaMsgBoxStyle As VbMsgBoxStyle = vbCritical _
    ) As VbMsgBoxResult
' Purpose: ********************************************************************
' Show a popup upon error.
' strActionMessage is included to make you think about making _actionable_
' error handlers.
' Err.Description.

' Requirements:
'
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           06/04/2020  Original
' 2     Sara Gleghorn  --           11/04/2020  Fixed a bug preventing message
'                                               box result being returned
' *****************************************************************************
'' Expected Parameters:
'Dim strErrorFunction    As String   ' The name of the errored function
'Dim strErrorSection     As String   ' The name of the section within the function
'Dim strErrorMessage     As String   ' What happened?
'Dim strActionMessage    As String   ' What do you want the end user to do about it?

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String   ' The name of this function (for debugging messages)
Dim strSection      As String   ' The name of the section (for debugging messages)
Dim strMsgBox       As String   ' Contents of the popup box

On Error GoTo ErrorHandler
CheckPrerequisites: ' ---------------------------------------------------------
strFnName = "showErrorHandlerPopup"
strSection = "CheckPrerequisites"

If strErrorFunction = vbNullString Then
    strErrorFunction = "Function name not passed to error handler"
ElseIf strErrorFunction = strFnName Then
    Exit Function ' Error handler has called error handler.
    ' This shouldn't happen, but exists to break the loop if it does.
End If

If strErrorMessage = vbNullString Then
    strErrorMessage = "Error message not passed to error handler"
End If

GeneratePopup: ' -------------------------------------------------------------
strSection = "CheckPrerequisites"

strMsgBox = "Function: " & strErrorFunction _
    & vbNewLine & "Section: " & strErrorSection _
    & vbNewLine _
    & vbNewLine & "Error: " & strErrorMessage _
    & vbNewLine _
    & vbNewLine & strActionMessage
    
showErrorHandlerPopup = MsgBox(strMsgBox, vbCritical, "Custom Error Handler")

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

End Function

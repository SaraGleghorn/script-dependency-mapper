Attribute VB_Name = "mod_filesystem"
Option Compare Database
Option Explicit

Function retrieveFilesInFolder( _
    strFolderPath As String, _
    bIncludeSubfolders As Boolean, _
    Optional strPatternMatch As String _
    ) As Collection
' Purpose: ********************************************************************
' Loop through each folder, and return the path of all files matching anything
' in the patternMatch array
' Requirements:
'   Reference -  Microsoft Scripting Runtime (for FileSystemObject)
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           23/03/2020  Original
' *****************************************************************************
' Expected Parameters:
'Dim folderPath      As String       ' The folder to search in
'Dim bIncludeSubfolders  As Boolean  ' Whether to include subfolders
'Dim patternMatch    As String       ' A pipe (|) separated set of patterns,
'                                    ' to match with VBA LIKE. e.g:
'                                    ' "*.sql" for anything ending in .sql
'                                    ' "*.xls?" for ending .xlsx or .xlsm
'                                    ' Leave blank to get all files

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function (for debugging messages)
Dim strSection      As String           ' The name of the section (for debugging messages)
Dim FSO             As FileSystemObject ' FileSystemObject
Dim fsoFolder       As Folder           ' Folders within file system object
Dim fsoSubfolder    As Folder
Dim fsoFile         As file             ' File objects
Dim colFolders      As Collection       ' Holds all our folders & subfolders
Dim i               As Integer          ' For looping through our matching criteria
Dim strPatterns()   As String           ' An array of our patterns for matching

On Error GoTo ErrorHandler
strFnName = "retrieveFilesInFolder"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set retrieveFilesInFolder = New Collection
Set colFolders = New Collection
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

If FSO.FolderExists(strFolderPath) = False Then GoTo ErrorBadFolder

SetDefaults: ' ----------------------------------------------------------------
strSection = "SetDefaults"
If strPatternMatch = "" Then
    strPatternMatch = "*.*"
End If

RetrieveFiles: ' --------------------------------------------------------------
strSection = "RetrieveFiles"

strPatterns() = Split(strPatternMatch, "|")

colFolders.Add strFolderPath

Do While colFolders.Count > 0
    Set fsoFolder = FSO.GetFolder(colFolders(1))
    colFolders.Remove 1
    
    ' Retrieve Subfolders
    If bIncludeSubfolders = True Then
        For Each fsoSubfolder In fsoFolder.SubFolders
            colFolders.Add fsoSubfolder.Path
        Next
    End If
    
    ' Retrieve Files
    For Each fsoFile In fsoFolder.Files
        For i = LBound(strPatterns) To UBound(strPatterns)
            If Left(fsoFile.Name, 1) <> "~" _
            And fsoFile.Name Like strPatterns(i) Then
                retrieveFilesInFolder.Add fsoFile.Path
            End If
        Next
    Next

Loop

Exit Function
ErrorHandler: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Call showErrorHandlerPopup(strFnName, strSection, Err.Description)
Exit Function

ErrorBadFolder: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & ": " & "Folder " & """" & strFolderPath _
    & """" & " does not exist."
Call showErrorHandlerPopup(strFnName, strSection, _
    "Folder " & """" & strFolderPath & """" & " does not exist.", _
    "Please check if the path in the script_folder table is valid and that " _
        & "the folder exists", vbCritical)
End Function

Function MakeNestedDirectory(strFolderPath As String) As Boolean
' Purpose: ********************************************************************
' MkDir and fso.CreateFolder only work if the parent folder exists.
' This makes the entire folder path if it's missing.
' Requirements:
'   Reference -  Microsoft Scripting Runtime (for FileSystemObject)
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           23/03/2020  Original
' *****************************************************************************
' Expected Parameters:
'Dim folderPath      As String       ' The folder to create (absolute directory)


Definitions: '-----------------------------------------------------------------
Dim strFnName               As String           ' The name of this function (for debugging messages)
Dim strSection              As String           ' The name of the section (for debugging messages)
Dim strFolderPathParts()    As String           ' Holds the split Folder Path directories
Dim strFolderPathPartial    As String           ' Holds partial folder path
Dim iStart                  As Integer          ' So we can skip generating the drive:\ or network \\
Dim i                       As Integer          ' For iterating through strFolderPath backslashes
Dim FSO                     As FileSystemObject ' So we don't break any global Dir() usage

On Error GoTo ErrorHandler
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

' Is folderPath anything?
If Len(strFolderPath) = 0 Then GoTo ErrorEmptyPath

' Have we specified a drive name or network?
If Mid(strFolderPath, 2, 2) = ":\" Then
    iStart = 1
ElseIf Left(strFolderPath, 2) = "\\" Then
    iStart = 3
Else
    GoTo ErrorInvalidPath
End If

GenerateFolder: ' -------------------------------------------------------------
strSection = "GenerateFolder"
Set FSO = New FileSystemObject
strFolderPathParts() = Split(strFolderPath, "\")
strFolderPathPartial = Left(strFolderPath, InStr(iStart, strFolderPath, "\"))

For i = iStart To UBound(strFolderPathParts)
    If strFolderPathParts(i) <> "" Then
        strFolderPathPartial = strFolderPathPartial & strFolderPathParts(i) & "\"
        Debug.Print strFolderPathPartial
        If FSO.FolderExists(strFolderPathPartial) = False Then
            FSO.CreateFolder (strFolderPathPartial)
        End If
    End If
Next

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
MakeNestedDirectory = True
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

ErrorEmptyPath: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "argument: folderPath, is nothing. Cannot create folder."
Exit Function

ErrorInvalidPath: ' -----------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "argument: folderPath, is not a valid path"
Exit Function
End Function

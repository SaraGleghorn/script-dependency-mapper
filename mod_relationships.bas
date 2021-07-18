Attribute VB_Name = "mod_relationships"
Option Compare Database
Option Explicit

Function sqlDependencyMapper() As Boolean
' Purpose: ********************************************************************
' Calls extractDependenciesFromFile for each file in script_filepath
' Version Control:
' Vers      Author          Date        Change
' 0.0.9     Sara Gleghorn   31/03/2020  Original
' 1.0.0     Sara Gleghorn   06/04/2020  Added showErrorHandlerPopup calls
'                                       Improved error handling around missing
'                                       files
' 1.2.0     Sara Gleghorn   15/04/2020  Added splitObjectNames call
'                                       Whitespace tidy
' *****************************************************************************
' Expected Parameters:
' None
Definitions: '-----------------------------------------------------------------
Dim strFnName   As String           ' The name of this function
                                    ' (for debugging messages)
Dim strSection  As String           ' The name of the section
                                    ' (for debugging messages)
                                    
Dim db          As DAO.Database     ' This database
Dim rsFileList  As DAO.Recordset    ' List of all the files to parse,
                                    ' populated by populateFileList()
Dim FSO         As FileSystemObject ' For navigating file system
Dim strMsg      As String           ' Messages for listing missing files.
Dim bSupress    As Boolean          ' Stop further popups about missing files?

On Error GoTo ErrorHandler

CheckPrerequisites: ' ---------------------------------------------------------
Set db = CurrentDb
strFnName = "fn_template"
strSection = "CheckPrerequisites"
Set FSO = New FileSystemObject

Set rsFileList = db.OpenRecordset("script_filepath")
If rsFileList.RecordCount = 0 Then GoTo ErrorNoFiles

ClearExistingData: ' -----------------------------------------------------------
strSection = "ClearExistingData"
db.Execute ("DELETE * FROM parse_errors")
db.Execute ("DELETE * FROM table_usage")
Call updateLog("Clear table_usage table")

LoopThroughFiles: ' -----------------------------------------------------------
strSection = "LoopThroughFiles"

Do Until rsFileList.EOF

    If FSO.FileExists(rsFileList!filepath) = False Then
        If bSupress = False Then
            If showErrorHandlerPopup(strFnName, strSection, _
                "Some files are missing. Would you like to continue?", _
                "(future messages will be suppressed.)", vbYesNo) = vbYes Then
                    bSupress = True
                    strMsg = "Missing Files: " _
                        & vbNewLine & "    " & strFnName
            Else
                GoTo ExitMissingFile
            End If
        Else
            strMsg = strMsg & vbNewLine & "    " & strFnName
        End If
    Else
        Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
            & vbNewLine; "    Preparing to parse: " & rsFileList!filepath
        Call extractDependenciesFromFile(rsFileList!filepath)
    End If
    rsFileList.MoveNext
Loop

If strMsg <> vbNullString Then GoTo ExitMissingFile

Set rsFileList = Nothing

SplitObjects: ' ----------------------------------------------------------------
strSection = "SplitObjects"
Call splitObjectNames

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
Call updateLog("Populate table_usage table")
sqlDependencyMapper = True
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Call showErrorHandlerPopup(strFnName, strSection, Err.Description)
Exit Function

ErrorNoFiles: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: There are no files to parse. Check the script_filepath table"
Call showErrorHandlerPopup(strFnName, strSection, _
    "There are no files to parse", _
    "Check that the script_folder table contains the correct folders, " _
    & "and that 'Find .SQL files' (function populateFileList) has run first." _
    , vbOKOnly)
Exit Function

ExitMissingFile: ' ------------------------------------------------------------
' User chose to quit after an error message
Debug.Print Now() & " " & strFnName & "." & strSection & ": " & strMsg

End Function

Function populateFileList() As Boolean
' Purpose: ********************************************************************
' Loop through each folder, and dump a list of filenames into our local table.
' Requirements:
'   Reference -  Microsoft Scripting Runtime (for FileSystemObject)
' Version Control:
' Vers  Author          Date        Change
' 0.0.9 Sara Gleghorn   24/03/2020  Original.
' 1.0.0 Sara Gleghorn   06/04/2020  Handles apostrophes in filenames.
'                                   Added updated error handler.
'                                   Removed unused variables (i, bLogInfo)
' 1.2.0 Sara Gleghorn   15/04/2020  Whitespace tidy
' *****************************************************************************
Definitions: ' ----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function
                                        ' (for debugging messages)
Dim strSection      As String           ' The name of the strSection
                                        ' (for debugging messages)
Dim db              As DAO.Database     ' This Database
Dim strSQL          As String           ' Inline query text
Dim rsFolderList    As DAO.Recordset    ' Will hold our list of folders
Dim strFolderPath   As String           ' The name of our folder
Dim colFiles        As Collection       ' Collection of filenames
Dim strFilePath     As String           ' Filepath currently being processed
Dim i               As Integer          ' For looping
Dim strMsg          As String           ' For error messaging
Dim FSO             As FileSystemObject ' For navigating files.
                                        ' Dir() is faster, but is global,
                                        ' so can interupt Dir() in calling
                                        ' or called functions
Dim fsoFolder       As Folder

On Error GoTo ErrorHandler
strFnName = "populateFileList"

CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

Set db = CurrentDb

' Check table exists
If databaseObjectExists("script_folder", "table", db) = False Then
    GoTo ErrorMissingTable
End If

' Check folders listed in our table exist
' This means if we've typo'd a folder, we quit early,
' instead of after spending time retrieving files
strSQL = "SELECT folder_path, include_subfolders FROM script_folder"
Set rsFolderList = db.OpenRecordset(strSQL)

rsFolderList.MoveFirst
Set FSO = New FileSystemObject
Do Until rsFolderList.EOF = True
    If FSO.FolderExists(rsFolderList!folder_path) = False Then
        strMsg = strMsg & "    " & rsFolderList!folder_path & "," & vbNewLine
    End If
    rsFolderList.MoveNext
Loop

If Len(strMsg) > 0 Then GoTo ErrorBadFolder

PrepareTables: ' --------------------------------------------------------------

If updateLog("Clear script_filepath table") = False Then GoTo ErrorMakingLog

strSection = "CheckPrerequisites"
If databaseObjectExists("script_filepath", "table", db) = False Then
    ' Generate make table code
    strSQL = "CREATE TABLE script_filepath (" _
        & vbNewLine & "filepath CHAR CONSTRAINT [Primary Key] PRIMARY KEY" _
        & vbNewLine & ");"
Else
    ' Generate clear table code
    strSQL = "DELETE * FROM script_filepath;"
End If

db.Execute strSQL, dbFailOnError
strSQL = vbNullString

RetrieveFiles: '---------------------------------------------------------------
strSection = "RetrieveFiles"

rsFolderList.MoveFirst
Do Until rsFolderList.EOF = True

    ' Add the trailing slash
    strFolderPath = rsFolderList!folder_path
    If Right(strFolderPath, 1) <> "\" Then strFolderPath = strFolderPath & "\"
       
    ' Retrieve the files
    Set colFiles = retrieveFilesInFolder( _
        strFolderPath, _
        rsFolderList!include_subfolders, _
        "*.SQL")
    
    ' If the strFilePath does not already exist in our table, insert it.
    Do While colFiles.Count > 0
        strFilePath = colFiles(1)
        colFiles.Remove 1
        ' Sanitise before insert
        strFilePath = Replace(strFilePath, "'", "''")
        If DCount("*", _
            "script_filepath", _
            "filepath = '" & strFilePath & "'") = 0 _
        Then
            strSQL = "INSERT INTO script_filepath VALUES ('" & strFilePath & "');"
            db.Execute strSQL
        End If
    Loop
    
    rsFolderList.MoveNext
Loop

Cleanup: ' -------------------------------------------------------------------
If updateLog("Populate script_filepath table") = False Then
    GoTo ErrorMakingLog
End If

populateFileList = True

Exit Function

ErrorHandler: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Call showErrorHandlerPopup(strFnName, strSection, Err.Description)
Exit Function

ErrorMissingTable: ' ----------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "search_folders table doesn't exist. " _
    & "Cannot retrieve list of folders to search. Quitting."
Call showErrorHandlerPopup(strFnName, strSection, _
    "Cannot locate the script_folder table. ", _
    "Check that this table has not been deleted or renamed.")
Exit Function

ErrorBadFolder: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Could not find the following folders: " _
    & vbNewLine & strMsg _
    & "Quitting."
Call showErrorHandlerPopup(strFnName, strSection, _
    "Could not find folders: " & vbNewLine & strMsg, _
    "Check the folder paths listed in script_folder table are correct, " _
    & "and that you have permissions to access the folder.")

Exit Function

ErrorMakingLog: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Errors updating log table. Quitting."
Call showErrorHandlerPopup(strFnName, strSection, _
    "Error updating log table: " & Err.Description, _
    "Check that the log_last_update table exists and is writable.")
    
Exit Function

End Function

Function extractDependenciesFromQuery(arQuery As Variant, _
    strSource As String, _
    strSourceType As String) As Boolean
' Purpose: ********************************************************************
'   Extract table dependencies from a single query, and inserts them into
'   the script dependencies table.
'   Recursive: for subqueries, calls this function again
'   This does assume that your query is valid.
' Requirements:
'
' Version Control:
' Vers  Author          Date        Change
' 0.0.9 Sara Gleghorn   01/04/2020  Original
' 1.0.0 Sara Gleghorn   06/04/2020  Bugfix to stop capturing aliases after
'                                   source subqueries as tablenames
' 1.1.0 Sara Gleghorn   11/04/2020  Bugfix to exclude ";" from table names
' 1.2.0 Sara Gleghorn   15/04/2020  Whitespace tidy
' *****************************************************************************
' Expected Parameters:
'Dim arQuery()       As Variant  ' The query to parse:
                                 ' [0, n] linenumber,
                                 ' [1, n] line content
'Dim strSource       As String   ' The query source or filepath
'Dim strSourceType   As String   ' Where this query exists
                                 ' e.g. SQL File, Access QueryDef
Definitions: '-----------------------------------------------------------------
Dim strFnName           As String           ' The name of this function
                                            ' (for debugging messages)
Dim strSection          As String           ' The name of the section
                                            ' (for debugging messages)

Dim db                  As DAO.Database     ' This db
Dim strQry              As String           ' Query to retrieve next expected
                                            ' SQL command
Dim strQryInsert        As String           ' Query for our insert statements,
                                            ' when a dependency is found.
Dim rsSeek              As DAO.Recordset    ' Recordset of phrases to look for
Dim iLine               As Integer          ' Current query line
Dim iChar               As Integer          ' Current character position
Dim strChar             As String           ' Current character being parsed
Dim bSkipChar           As Boolean          ' Skip adding current character to
                                            ' phrase capture
Dim bStringData         As Boolean          ' Is this character part of a literal
                                            ' text string? (e.g. 'things' in:
                                            '   "SELECT 'things'
                                            '   FROM stuff")
Dim bIdentifier         As Boolean          ' Is this part of an identifier,
                                            ' which may include spaces?
                                            ' (e.g. [my table] in
                                            '   "SELECT myColumn
                                            '   FROM [my table];"
Dim bPhrase             As Boolean          ' Do we have a completed word?
Dim strPhrase           As String           ' Our parsed phrase
Dim strCommand          As String           ' What this table name is used for
                                            ' e.g. SELECT, INSERT INTO, WITH
Dim bCaptureSourceMode  As Boolean          ' Are we capturing a source?
Dim bCaptureAliasMode   As Boolean          ' Are we expecting an alias?
                                            ' (prevent recording aliases as
                                            ' tables)
Dim bSubquery           As Boolean          ' Are we parsing a subquery?
Dim iBrackets           As Integer          ' To keep track of brackets
Dim arSubquery()        As Variant          ' Will hold our subquery.
Dim iSubqueryRow        As Integer          ' Row of our subquery array
Dim strSubquery         As String           ' For holding our subquery content
Dim strPhraseSubstitute As String           ' strPhrase with Regex breaking
                                            ' content removed.

On Error GoTo ErrorHandler
strFnName = "extractDependenciesFromQuery"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

Initialise: ' -----------------------------------------------------------------
strSection = "Initialise"
strCommand = vbNullString
Set db = CurrentDb

' Move some of our syntax tables into a recordset for speed
' Queries have a structure - we'll have a command or CTE first,
' then look for our table sources
strQry = "SELECT * FROM syntax_phrases WHERE parse_type IN ('command', 'common table expression')"
Set rsSeek = db.OpenRecordset(strQry)
' These should already be null, but assist when debugging.
strQry = vbNullString
strChar = vbNullString
bSkipChar = False
bStringData = False
bPhrase = False
bIdentifier = False
bSubquery = False
bCaptureSourceMode = False
bCaptureAliasMode = False
iBrackets = 0
iSubqueryRow = 0
strPhrase = vbNullString
ReDim arSubquery(1, 0)
strSubquery = vbNullString

Parse: '-----------------------------------------------------------------------
strSection = "Parse"
' We can't split the query into an array because split() will not detect
' when a split is inappropriate (e.g. within a comment or quote).
' So we will loop through per character.
' We are only interested in tables - columns and conditions will not be parsed.

For iLine = LBound(arQuery, 2) To UBound(arQuery, 2) ' Each line

    If bSubquery = True Then
        iSubqueryRow = iSubqueryRow + 1
        ReDim Preserve arSubquery(1, iSubqueryRow)
    End If

    For iChar = 1 To Len(arQuery(1, iLine)) ' Each Character
        bSkipChar = False
    
        ' Move the character into a variable for easier reading
        strChar = Mid(arQuery(1, iLine), iChar, 1)
        
        ' Detect special characters
        If bSubquery = True Then
            Select Case strChar
                Case "("
                    iBrackets = iBrackets + 1
                Case ")"
                    iBrackets = iBrackets - 1
                    If iBrackets = 0 Then
                        bSkipChar = True
                    End If
            End Select
        Else
        Select Case strChar
                Case "'"
                    If iChar > 1 Then
                        If Mid(arQuery(1, iLine), iChar - 1, 1) = "\" Then
                            ' Do nothing - the apostrophe is escaped
                        Else
                            bStringData = Not bStringData
                            bPhrase = False
                            bSkipChar = True
                        End If
                    Else
                        bStringData = Not bStringData
                        bPhrase = False
                        bSkipChar = True
                    End If
                Case "["
                    If bStringData = False Then
                        bIdentifier = True
                    End If
                Case "]"
                    If bStringData = False Then
                        bIdentifier = False
                        bPhrase = True
                    End If
                Case " ", vbTab ' End of a word, unless within an identifier
                    If bStringData = False And bIdentifier = False Then
                        bPhrase = True
                        bSkipChar = True
                    End If
                Case "," ' End of a word
                    If bStringData = False Then
                        bPhrase = True
                        bSkipChar = True
                    End If
                Case "("
                    bSubquery = True
                    iSubqueryRow = 0
                    bSkipChar = True
                    strSubquery = vbNullString
                    iBrackets = iBrackets + 1
                Case ";"
                    bSkipChar = True
                    bPhrase = True
            End Select
        End If
        
        ' Handle Subqueries
        If bSubquery = True Then
            If bSkipChar = True Then
                bSkipChar = False
            Else
                strSubquery = strSubquery + strChar
            End If
            
            If iBrackets = 0 _
            Or iChar = Len(arQuery(1, iLine)) Then
                ' It is either the end of the subquery or the end of the line.
                ' Add the line so far into an array
                arSubquery(0, iSubqueryRow) = arQuery(0, iLine)
                arSubquery(1, iSubqueryRow) = strSubquery
                strSubquery = vbNullString
            End If
            
            If iBrackets = 0 Then
                If extractDependenciesFromQuery( _
                    arSubquery, _
                    strSource, _
                    strSourceType) = False _
                Then
                    GoTo ErrorSubqueryParseFailed
                End If
                bSubquery = False
                strSubquery = vbNullString
                iSubqueryRow = 0
                ReDim arSubquery(1, iSubqueryRow)
                
                If bCaptureSourceMode = True Then
                    ' What follows should be either a comma or an alias
                    bCaptureAliasMode = True
                End If
            End If
        Else
            ' The end of a line will always be the end of a word
            If iChar = Len(arQuery(1, iLine)) Then
                bPhrase = True
            End If
            
            ' Add this character to our captured phrase
            If bStringData = True Then
                ' Do nothing
            ElseIf bSkipChar = True Then
                bSkipChar = False
            Else
                strPhrase = strPhrase + strChar
            End If
            
            ' Are we in source capture mode?
            If bCaptureSourceMode = True _
            And bPhrase = True _
            And strPhrase <> vbNullString _
            Then
                ' Don't capture aliases as table names
                If bCaptureAliasMode = False Then
                        
                    Debug.Print strPhrase
                    strQryInsert = "INSERT INTO table_usage (" _
                        & vbNewLine & "    " & "query_source_type, " _
                        & vbNewLine & "    " & "query_source, " _
                        & vbNewLine & "    " & "query_object, " _
                        & vbNewLine & "    " & "line_number, " _
                        & vbNewLine & "    " & "command " _
                        & vbNewLine & ") VALUES ( " _
                        & vbNewLine & "    " & "'" & strSourceType & "', " _
                        & vbNewLine & "    " & "'" & strSource & "', " _
                        & vbNewLine & "    " & "'" & strPhrase & "', " _
                        & vbNewLine & "    " & "'" & arQuery(0, iLine) & "'," _
                        & vbNewLine & "    " & "'" & strCommand & "' " _
                        & vbNewLine & ")"
                    db.Execute strQryInsert
                    strQryInsert = vbNullString
                    
                    bPhrase = False
                    strPhrase = vbNullString
                    bCaptureAliasMode = True
                    
                    If strCommand = "SELECT INTO" Then
                        strCommand = "SELECT"
                    End If
                End If
            End If
            
            
            If bPhrase = True Then
                If strPhrase = vbNullString Then
                    bPhrase = False
                Else
                    ' Find exact matches to this phrase
                    rsSeek.MoveFirst
                    ' Catch TSQL's "SELECT ... INTO"
                    If strCommand = "SELECT" And strPhrase = "INTO" Then
                        rsSeek.FindFirst "[parse_phrase] = 'SELECT INTO'"
                    Else
                        rsSeek.FindFirst "[parse_phrase] = '" & strPhrase & "'"
                    End If
                    
                    If Not rsSeek.NoMatch Then
                        ' Found an exact match
                        bPhrase = True
                        bCaptureSourceMode = rsSeek!capture_table_next
                        bCaptureAliasMode = False
                        
                        ' Manage types of phrase
                        Select Case rsSeek!parse_Type
                            Case "common table expression"
                                strCommand = rsSeek!parse_phrase
                                strQry = "SELECT * " _
                                    & "FROM syntax_phrases " _
                                    & "WHERE parse_type IN ('command')"
                            Case "command"
                                strCommand = rsSeek!parse_phrase
                                strQry = "SELECT * " _
                                    & "FROM syntax_phrases " _
                                    & "WHERE parse_type IN (" _
                                    & "'command','source','modify','union'," _
                                    & "'into','common table expression'," _
                                    & "'apply')"
                            Case "source"
                                strQry = "SELECT * " _
                                    & "FROM syntax_phrases " _
                                    & "WHERE parse_type IN ('command'," _
                                    & "'source','union','condition','group'," _
                                    & "'where', 'order')"
                                If strCommand = "SELECT INTO" Then
                                    strCommand = "SELECT"
                                End If
                        End Select
                        bPhrase = False
                        strPhrase = vbNullString
                    Else
                        ' Check if there are any partial matches
                        ' (e.g. INSERT INTO for INSERT)
                        rsSeek.MoveFirst
                        
                        strPhraseSubstitute = strPhrase
                        
                        Select Case strPhrase
                            Case "[", "?", "#", "*" ' TODO: Check this works
                                strPhraseSubstitute = Replace( _
                                    strPhrase, "[", "[[]")
                                strPhraseSubstitute = Replace( _
                                    strPhrase, "?", "[?]")
                                strPhraseSubstitute = Replace( _
                                    strPhrase, "#", "[#]")
                                strPhraseSubstitute = Replace( _
                                    strPhrase, "*", "[*]")
                        End Select
                        rsSeek.FindFirst "[parse_phrase] LIKE '" _
                            & strPhraseSubstitute & " *'"
                    
                        If Not rsSeek.NoMatch Then
                            ' We may have first word in a multi-word command.
                            bPhrase = False
                            If strChar = " " Or strChar = vbTab Then
                                strPhrase = strPhrase & " "
                            End If
                        Else
                            ' Matches nothing
                            strPhrase = vbNullString
                            bPhrase = False
                        End If
                        
                        strPhraseSubstitute = vbNullString
                        
                    End If
                    
                    If strQry <> vbNullString Then
                        rsSeek.Close
                        Set rsSeek = db.OpenRecordset(strQry)
                        strQry = vbNullString
                        bPhrase = False
                    End If
                End If
                
                ' Detect the end of alias
                If bCaptureAliasMode = True Then
                    If strChar = "," Then
                        bCaptureAliasMode = False
                    End If
                End If
            End If
        End If
                   

    Next 'iChar
Next ' iLine

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
extractDependenciesFromQuery = True
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

ErrorQueryEmpty: ' ------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "The passed argument 'strQuery' is empty. Nothing to parse. Quitting,"
Exit Function

ErrorSubqueryParseFailed: ' ---------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Failure when parsing subquery:"
For iLine = LBound(arSubquery, 2) To UBound(arSubquery, 2)
    Debug.Print "    " & arSubquery(0, iLine) & ":    " & arSubquery(1, iLine)
Next

End Function

Function extractDependenciesFromFile(strFilePath As String) As Boolean
' Purpose: ********************************************************************
' Splits an SQL file into separate queries, and sends them to
' extractDependenciesFromQuery.
'
' Requirements:
'   Reference: Microsoft ActiveX Object Library (for ADODB.Stream)
'
' Version Control:
' Vers  Author          Date        Change
' 0.9.0 Sara Gleghorn   31/03/2020  Original
' 1.0.0 Sara Gleghorn   06/04/2020  Add recognition for EXECUTE IMMEDIATE.
'                                   Added showErrorHandlerPopup calls
' 1.1.0 Sara Gleghorn   11/04/2020  Fixed error on logging errors in files
'                                   when there are apostrophes in filename.
' 1.2.0 Sara Gleghorn   15/04/2020  Whitespace and comments tidy
'                                   Fixed error on logging errors
' 1.3.0 Sara Gleghorn   10/05/2020  Changed ts from TextStream to ADODB.Stream
'                                   for greater flexibility on reading files.
'                                   Fixed issues with EXECUTE IMMEDIATE not
'                                   being captured if there was a previous
'                                   command (e.g WHENEVER...) that didn't end
'                                   with a semi colon.
'                                   Improved detection for end of file so we
'                                   don't need to repeat logic for catching
'                                   the last query if it doesn't have a semi-
'                                   colon.
'                                   Amended version control to be inline with
'                                   all other versioning in this file.
' 1.3.1 Sara Gleghorn   18/07/2021  Fixed an issue where SQL files that don't
'                                   use CRLF (but LF, in Linux) don't get
'                                   parsed correctly.
' *****************************************************************************
' Expected Parameters:
'Dim strFilePath    As String  ' The script to extract dependancies from

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function
                                        ' (for debugging messages)
Dim strSection      As String           ' The name of the section
                                        ' (for debugging messages)

Dim db              As DAO.Database     ' This database
Dim strErrorSQL     As String           ' String for error logging query
Dim strTempPath     As String           ' Path to copy our file to while
                                        ' reading it (prevent locking of
                                        ' operational scripts)
Dim FSO             As FileSystemObject ' Windows File System Object
Dim fileTemp        As file             ' Our copy of the file to be parsed
'Dim ts              As TextStream       ' Contents of fileTemp
Dim ts              As ADODB.Stream     ' Contents of fileTemp
Dim strLine         As String           ' Contents of one line of ts
Dim iFileLength     As Integer          ' Length of the text file
Dim iFileLine       As Integer          ' Line number counter
Dim iQueryLine      As Integer          ' Line number counter within one query
Dim iComment        As Integer          ' For tracking multiline comments, as
                                        ' some platforms allow nesting comments
Dim arQuery()       As Variant          ' Holds one query, for passing to
                                        ' extractDependenciesFromQuery()
                                        ' [0, n] iFileLine,
                                        ' [1, n] Query text on this line
Dim iChar           As Integer          ' Current character
Dim strCleanLine    As String           ' strLine after removing commented code
Dim bQueryEnd       As Boolean          ' We have reached the end of a query
Dim bLiteralString  As Boolean          ' We are inside a literal string,
                                        ' (So semi colons don't end the query)
Dim strRemainder    As String           ' Any code remaining after semicolon
Dim bTransaction    As Boolean          ' Are we inside an EXECUTE IMMEDIATE?
Dim bSkip           As Boolean          ' Whether to skip to the next query
                                        ' Used for some transactions

On Error GoTo ErrorHandler
strFnName = "extractScriptDependencies"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"
Set db = CurrentDb

' File Exists
Set FSO = New FileSystemObject
If FSO.FileExists(strFilePath) = False Then GoTo ErrorBadFile

' Temporary Directory Exists
strTempPath = DLookup( _
    "configuration", _
    "config", _
    "description = 'Temporary File Location'")
If IsNull(strTempPath) Then
    strTempPath = "C:\Temp\" ' Default
    Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
        ; "Could not locate filepath configuration from " _
        & """" & "config" & """" & " table. Using default location of " _
        & vbNewLine & "    " & """" & strTempPath & """"
End If

If Right(strTempPath, 1) <> "\" Then strTempPath = strTempPath & "\"
strTempPath = strTempPath & strFnName & "\"

If FSO.FolderExists(strTempPath) = False Then
    If MakeNestedDirectory(strTempPath) = False Then GoTo ErrorMakingDirectory
End If

CopyFileToTemp: ' -------------------------------------------------------------
' Copy the file to a location in our C:\ drive,
' so that we don't lock files for other users
strSection = "CopyFileToTemp"

strTempPath = strTempPath & Right( _
    strFilePath, _
    Len(strFilePath) - InStrRev(strFilePath, "\"))
If strFilePath = strTempPath Then GoTo ErrorSamePath
FSO.CopyFile strFilePath, strTempPath, True
Set fileTemp = FSO.GetFile(strTempPath)

ParseFile: ' -------------------------------------------------------------
strSection = "ParseFile"
' Set up an array to hold one query,
' to be passed to extractDependenciesFromQuery()
ReDim arQuery(1, 0)

' Open and loop through our temp file, extracting all references to tables
Set ts = New ADODB.Stream
With ts
    ts.Charset = "utf-8"
    .Type = adTypeText
    .LineSeparator = adCRLF ' Windows CRFL. This is the most likely scenario;
                            ' if you're using VBA you're probably on Windows
    .Open
    .LoadFromFile (fileTemp)
End With

GetFileLength:
' Get the number of Lines
iFileLength = 0
ts.Position = 0
Do While ts.EOS = False
    strLine = ts.ReadText(adReadLine)
    iFileLength = iFileLength + 1
Loop

' Check the number of lines.
' If we only retrieved one line, then it's possible that we're using the wrong
' line separator. This could happen if the file was made on another operating
' system.
If iFileLength <= 1 Then
    If ts.LineSeparator = adCRLF Then
        With ts
            .Close
            .LineSeparator = adLF ' Linux LF only
            .Open
            .LoadFromFile (fileTemp)
        End With
        GoTo GetFileLength
    ElseIf ts.LineSeparator = adLF Then
        With ts
            .Close
            .LineSeparator = adCR ' Mac CR only
            .Open
            .LoadFromFile (fileTemp)
        End With
        GoTo GetFileLength
    End If
End If

' Reset to the beginning and split the file into individual queries
ts.Position = 0

Do While ts.EOS = False
    
    ' If we ended a query midline,
    ' add the remainder to the beginning of the new query.
    If strRemainder = vbNullString Then
        If iQueryLine > UBound(arQuery, 2) Then
            ReDim Preserve arQuery(1, iQueryLine)
        End If
        strLine = Trim(Replace(ts.ReadText(adReadLine), vbTab, " "))
        iFileLine = iFileLine + 1
    Else
        ReDim arQuery(1, iQueryLine)
        strLine = strRemainder
        strRemainder = vbNullString
    End If
    
    ' Manage multiline comments
    strCleanLine = vbNullString
    If Len(strLine) > 0 Then
        If iComment > 0 Or InStr(1, strLine, "/*") <> 0 Then
            For iChar = 1 To Len(strLine)
                If Mid(strLine, iChar, 2) = "/*" Then
                    iComment = iComment + 1
                ElseIf Mid(strLine, iChar, 2) = "*/" Then
                    iComment = iComment - 1
                Else
                    If iComment = 0 Then
                        strCleanLine = strCleanLine & Mid( _
                            strLine, _
                            iChar + 1, _
                            1)
                        ' End comment is two characters, so skip 1
                        iChar = iChar + 1
                    End If
                End If
            Next
            
            strLine = Trim(strCleanLine)
            strCleanLine = vbNullString
        End If
    End If
    
    ' Strip single line comments
    If InStr(1, strLine, "--") <> 0 Then
        strLine = Left(strLine, InStr(1, strLine, "--") - 1)
    End If
    
    ' Strip out transactions
    Select Case strLine
        Case "BEGIN", "END;", "/", "COMMIT", "COMMIT;", "WHENEVER", "EXIT", "QUIT"
            strLine = vbNullString
    End Select

    If Left(strLine, 8) = "WHENEVER" Then
        strLine = vbNullString
    End If
    
    If InStr(1, strLine, "EXECUTE IMMEDIATE") <> 0 Then
        bTransaction = True
    End If
    
   
    ' Check for semi colons signalling the end of the query
    If bTransaction = False Then
        If InStr(strLine, ";") = 0 Then
            If InStr(strLine, "'") = 0 Then
            Else
                For iChar = 1 To Len(strLine)
                    If Mid(strLine, iChar, 1) = "'" Then
                        bLiteralString = Not bLiteralString
                    End If
                Next
            End If
        Else
            For iChar = 1 To Len(strLine)
                If Mid(strLine, iChar, 1) = "'" Then
                    bLiteralString = Not bLiteralString
                ElseIf Mid(strLine, iChar, 1) = ";" Then
                    If bLiteralString = False Then
                        bQueryEnd = True
                        strRemainder = Mid( _
                            strLine, _
                            iChar + 1, _
                            Len(strLine) - iChar)
                        strLine = Trim(Left(strLine, iChar))
                        Exit For
                    End If
                End If
            Next
        End If
    ElseIf bTransaction = True _
    And bSkip = False Then
        If InStr(strLine, "INTO") = 0 _
        And InStr(strLine, "USING") = 0 _
        And InStr(strLine, ";") = 0 Then
            If InStr(strLine, "'") = 0 Then
                ' Do nothing
            Else
                For iChar = 1 To Len(strLine)
                    If Mid(strLine, iChar, 1) = "'" Then
                        bLiteralString = Not bLiteralString
                    End If
                Next
            End If
        Else
            For iChar = 1 To Len(strLine)
                If Mid(strLine, iChar, 1) = "'" Then
                    bLiteralString = Not bLiteralString
                ElseIf Mid(strLine, iChar, 1) = ";" Then
                    If bLiteralString = False Then
                        bQueryEnd = True
                        strRemainder = Mid( _
                            strLine, _
                            iChar + 1, _
                            Len(strLine) - iChar)
                        strLine = Trim(Left(strLine, iChar))
                        If bTransaction = True Then
                            ' Chop off the "';" at the end
                            strLine = Left(strLine, Len(strLine) - 2)
                        End If
                        Exit For
                    End If
                End If
                
                If bLiteralString = False Then
                    If iChar <= Len(strLine) - Len("INTO") Then
                        If Mid(strLine, iChar, 4) = "INTO" Then
                            strLine = Left(strLine, iChar - 1)
                            bSkip = True
                        End If
                    End If
                    
                    If iChar <= Len(strLine) - Len("USING") Then
                        If Mid(strLine, iChar, 4) = "USING" Then
                            strLine = Left(strLine, iChar - 1)
                            bSkip = True
                        End If
                    End If
                End If
            Next
        End If
    ElseIf bTransaction = True _
    And bSkip = True Then
        If InStr(strLine, ";") = 0 Then
            If InStr(strLine, "'") = 0 Then
                ' Do Nothing
            Else
                For iChar = 1 To Len(strLine)
                    If Mid(strLine, iChar, 1) = "'" Then
                        bLiteralString = Not bLiteralString
                    End If
                Next
            End If
            strLine = vbNullString
        Else
            For iChar = 1 To Len(strLine)
                If Mid(strLine, iChar, 1) = "'" Then
                    bLiteralString = Not bLiteralString
                ElseIf Mid(strLine, iChar, 1) = ";" Then
                    If bLiteralString = False Then
                        bQueryEnd = True
                        strRemainder = Mid( _
                            strLine, _
                            iChar + 1, _
                            Len(strLine) - iChar)
                        strLine = ";"
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    
    If strLine <> vbNullString Then
        
        ' Manage transaction syntax
        If bTransaction = True Then
        
            If Left(strLine, 17) = "EXECUTE IMMEDIATE" Then
                ' Strip out the "EXECUTE IMMEDIATE" and the first apostrophe
                strLine = Trim(Replace(strLine, "EXECUTE IMMEDIATE", ""))
                strLine = Right(strLine, Len(strLine) - 1)
            End If
            strLine = Replace(strLine, "''", "'")
            
        End If
        
        arQuery(0, iQueryLine) = iFileLine
        arQuery(1, iQueryLine) = strLine
        iQueryLine = iQueryLine + 1
    End If
    
    ' When the end of a query is detected, extract its dependencies
    If bQueryEnd = True _
    Or ( _
        iFileLine = iFileLength _
        And strRemainder = vbNullString) _
    Then
    
        If bSkip = True Then
            If Right(arQuery(0, UBound(arQuery, 2)), 1) = "'" Then
                arQuery(0, UBound(arQuery, 2)) = _
                    Left(arQuery(0, UBound(arQuery, 2)), _
                        Len(arQuery(0, UBound(arQuery, 2))) - 1)
            End If
        End If
    
        If extractDependenciesFromQuery( _
            arQuery, _
            strFilePath, _
            "SQL File") = False _
        Then
            strErrorSQL = "INSERT INTO parse_errors VALUES ( " _
                & vbNewLine & "    " & "#" & Format( _
                    Now(), _
                    "dd-mmm-yyyy hh:nn:ss") & "#, " _
                & vbNewLine & "    " & "'" & Replace( _
                    strFilePath, _
                    "'", _
                    "''") & "'," _
                & vbNewLine & "    'Error parsing query within file.'," _
                & vbNewLine & "    '" & iFileLine & "')"
            db.Execute strErrorSQL
            strErrorSQL = vbNullString
        End If
        iQueryLine = 0
        ReDim arQuery(1, iQueryLine)
        bQueryEnd = False
        bLiteralString = False
        bSkip = False
    End If
Loop

CatchStragglingQuery: ' -------------------------------------------------------
strSection = "CatchStragglingQuery"
'If there is no semi colon on the end of the last query in a file, it's valid SQL,
' but the above will have missed it

If UBound(arQuery, 2) = 0 _
And IsEmpty(arQuery(0, 0)) Then
    ' Do Nothing
Else
    strErrorSQL = "INSERT INTO parse_errors VALUES ( " _
        & vbNewLine & "    " & "#" & Format( _
            Now(), _
            "dd-mmm-yyyy hh:nn:ss") & "#, " _
        & vbNewLine & "    " & "'" & Replace( _
            strFilePath, _
            "'", _
            "''") & "'," _
        & vbNewLine & "    'Missed last query in file.'," _
        & vbNewLine & "    '" & iFileLine & "')"
    db.Execute strErrorSQL
    strErrorSQL = vbNullString
End If

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
extractDependenciesFromFile = True
ts.Close

If FSO.FileExists(strTempPath) = True Then
    ' FileCopy will have inherited file permissions.
    ' Remove ReadOnly if it was inherited.
    If fileTemp.Attributes And ReadOnly Then
        fileTemp.Attributes = fileTemp.Attributes - ReadOnly
    End If
    FSO.DeleteFile (strTempPath)
End If

Set FSO = Nothing
Erase arQuery

Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Call showErrorHandlerPopup( _
    strFnName, _
    strSection, _
    Err.Description, _
    , _
    vbCritical)
Exit Function

ErrorBadFile: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Could not find the following file: " _
    & vbNewLine & "    " & """" & strFilePath _
    & vbNewLine & "    Quitting."
Call showErrorHandlerPopup(strFnName, strSection, "Could not find file: " _
    & vbNewLine & """" & strFilePath & """", _
    "Check the file still exists and you have access to it.", _
    vbCritical)
Exit Function

ErrorSamePath: '---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Original file and destination for temporary copy are the same: " _
    & vbNewLine & "    " & """" & strFilePath _
    & vbNewLine & "    Quitting."
Call showErrorHandlerPopup(strFnName, strSection, _
    "Original file and destination for temporary copy are the same: " _
    & vbNewLine & """" & strFilePath & """", _
    "Change the Temporary File Location as defined in the 'config' table.", _
    vbCritical)
Exit Function

ErrorMakingDirectory: ' -------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Errors while trying to make the directory: " _
    & vbNewLine & "    " & """" & strTempPath & """" _
    & vbNewLine & "    Quitting."
Call showErrorHandlerPopup( _
    strFnName, _
    strSection, _
    "Could not make directory: " _
    & vbNewLine & """" & strTempPath & """", _
    "Check whether you have permissions to the path in config", _
    vbCritical)
Exit Function
End Function

Function updateLog(strEventDescription As String) As Boolean
' Purpose: ********************************************************************
' Update the log table in this db
' Version Control:
' Vers  Author          Date        Change
' 1     Sara Gleghorn   23/03/2020  Original
' 1.0.1 Sara Gleghorn   15/04/2020  Whitespace tidy
' *****************************************************************************
' Expected Parameters:
'Dim strEventDescription    As String    ' What happened?

Definitions: ' ----------------------------------------------------------------
Dim strFnName               As String   ' The name of this function
                                        ' (for debugging messages)
Dim strSection              As String   ' The name of the strSection
                                        ' (for debugging messages)
Dim db                      As Database ' This Database
Dim strSQL                  As String   ' Update/Insert query text

On Error GoTo ErrorHandler
strFnName = "updateLog"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"
Set db = CurrentDb

updateLogTable: '---------------------------------------------------------------
strSection = "updateLogTable"

If DCount("*", _
    "log_last_update", _
    "event_description = '" & strEventDescription & "'") = 0 _
Then
    strSQL = "INSERT INTO log_last_update VALUES (" _
        & vbNewLine & "    '" & strEventDescription & "', " _
        & vbNewLine & "    #" & Format(Now(), "dd-mmm-yyyy hh:nn:ss") & "#, " _
        & vbNewLine & "    '" & Environ$("username") & "'" _
        & vbNewLine & ");"
Else
    strSQL = "UPDATE log_last_update SET" _
        & vbNewLine & "    event_date = #" & Format( _
            Now(), _
            "dd-mmm-yyyy hh:nn:ss") & "#, " _
        & vbNewLine & "    user_id = '" & Environ$("username") & "'" _
        & vbNewLine & "WHERE event_description = '" & strEventDescription & "'"
End If

db.Execute strSQL, dbFailOnError

updateLog = True

Exit Function

ErrorHandler: '----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
End Function

Function splitObjectNames() As Boolean
' Purpose: ********************************************************************
' Splits composite object references (schema.table or table@database) into
' separate fields.
' We could use select case / iif and manage everything on one mega query.
' But for readability, I have split into several separate queries.
' Version Control:
' Vers  Author          Date        Change
' 1     Sara Gleghorn   15/04/2020  Original
' *****************************************************************************
' Expected Parameters:
' None

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String       ' The name of this function
                                    ' (for debugging messages)
Dim strSection      As String       ' The name of the section
                                    ' (for debugging messages)
Dim db              As DAO.Database ' This database
Dim strSQL          As String       ' Update query text.

On Error GoTo ErrorHandler
CheckPrerequisites: ' ---------------------------------------------------------
strFnName = "splitObjectNames"
strSection = "CheckPrerequisites"

Set db = CurrentDb
If DCount("*", "table_usage") = 0 Then GoTo ErrorNothingToUpdate

SplitOracleLinkedDb: ' ---------------------------------------------------------
strSection = "SplitOracleDbLinks"
' We won't attempt to break the db link into server and db,
' because we do not know if Global Naming is in use.

' "Schema.Table@DatabaseLink" (Oracle linked database, with explicit schema)
strSQL = "UPDATE table_usage " _
    & vbNewLine & "SET object_database = RIGHT( " _
    & vbNewLine & "        query_object, " _
    & vbNewLine & "        LEN(query_object) - INSTR(query_object, '@')), " _
    & vbNewLine & "    object_schema = LEFT( " _
    & vbNewLine & "        query_object, " _
    & vbNewLine & "        INSTR(query_object, '.') - 1), " _
    & vbNewLine & "    object_table = MID( " _
    & vbNewLine & "        query_object, " _
    & vbNewLine & "        INSTR(1, query_object, '.') + 1, " _
    & vbNewLine & "        LEN(query_object) - InStr(query_object, '.') " _
    & vbNewLine & "        - (Len(query_object) - InStr(query_object, '@')) " _
    & vbNewLine & "        - 1) " _
    & vbNewLine & "WHERE query_object LIKE '*.*@*' " _
    & vbNewLine & "AND object_table IS NULL;"
db.Execute strSQL, dbFailOnError

' "Table@DatabaseLink" (Oracle linked database, implicit local schema)
strSQL = "UPDATE table_usage " _
    & vbNewLine & "SET object_database = RIGHT( " _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        LEN(query_object) - INSTR(query_object, '@'))," _
    & vbNewLine & "    object_table = LEFT(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        INSTR(query_object, '@') - 1) " _
    & vbNewLine & "WHERE query_object LIKE '*@*' " _
    & vbNewLine & "AND object_table IS NULL;"
db.Execute strSQL, dbFailOnError

SplitMSSSQLLinkedDb: ' --------------------------------------------------------
strSection = "SplitMSSSQLLinkedDb"

' Server.Database.Schema.Table (MSSQL linked database including server)
strSQL = "UPDATE table_usage " _
    & vbNewLine & "SET object_server = LEFT( " _
    & vbNewLine & "        query_object, " _
    & vbNewLine & "        INSTR(query_object, '.') - 1)," _
    & vbNewLine & "    object_database = MID(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        InStr(query_object, '.') + 1," _
    & vbNewLine & "        LEN(query_object) - INSTR(query_object, '.')" _
    & vbNewLine & "        - (LEN(query_object) - InStr(" _
    & vbNewLine & "                INSTR(query_object, '.')+1," _
    & vbNewLine & "                query_object," _
    & vbNewLine & "                '.') + 1))," _
    & vbNewLine & "    object_schema = MID(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        INSTR(" _
    & vbNewLine & "            INSTR(query_object, '.') + 1," _
    & vbNewLine & "            query_object,'.') + 1," _
    & vbNewLine & "        LEN(query_object) - INSTR(" _
    & vbNewLine & "            INSTR(query_object, '.')+1," _
    & vbNewLine & "            query_object," _
    & vbNewLine & "            '.')" _
    & vbNewLine & "        - (LEN(query_object) - INSTRREV("
strSQL = strSQL _
    & vbNewLine & "            query_object, " _
    & vbNewLine & "            '.') + 1))," _
    & vbNewLine & "    object_table = RIGHT(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        LEN(query_object) - INSTRREV(query_object,'.'))" _
    & vbNewLine & "WHERE query_object LIKE '*.*.*.*'" _
    & vbNewLine & "AND object_table IS NULL;"
db.Execute strSQL, dbFailOnError

' Database.Schema.Table (MSQL linked database)
strSQL = "Update table_usage " _
    & vbNewLine & "SET object_database = LEFT(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        INSTR(query_object, '.') - 1)," _
    & vbNewLine & "    object_schema = MID(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        INSTR(query_object, '.') + 1," _
    & vbNewLine & "        LEN(query_object) - INSTR(query_object, '.') - (" _
    & vbNewLine & "            LEN(query_object) - INSTR(" _
    & vbNewLine & "                INSTR(query_object, '.')+1," _
    & vbNewLine & "                query_object," _
    & vbNewLine & "                '.') + 1))," _
    & vbNewLine & "    object_table = RIGHT(query_object," _
    & vbNewLine & "        LEN(query_object) - InStrRev(query_object,'.'))" _
    & vbNewLine & "WHERE query_object LIKE '*.*.*'" _
    & vbNewLine & "AND object_table IS NULL;"
db.Execute strSQL, dbFailOnError

SplitSchemaTable: '------------------------------------------------------------
strSection = "SplitSchemaTable"

' Schema.Table
strSQL = "Update table_usage" _
    & vbNewLine & "SET object_schema = LEFT(" _
    & vbNewLine & "        query_object, INSTR(query_object,'.') - 1)," _
    & vbNewLine & "    object_table = RIGHT(" _
    & vbNewLine & "        query_object," _
    & vbNewLine & "        LEN(query_object) - INSTR(query_object,'.'))" _
    & vbNewLine & "WHERE query_object LIKE '*.*'" _
    & vbNewLine & "AND object_table IS NULL;"
db.Execute strSQL, dbFailOnError

OnlyTable: '-------------------------------------------------------------------
strSection = "OnlyTable"

strSQL = "UPDATE table_usage" _
    & vbNewLine & "SET object_table = query_object" _
    & vbNewLine & "WHERE object_table Is Null;"
db.Execute strSQL, dbFailOnError

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
splitObjectNames = True
Exit Function

ErrorHandler: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Call showErrorHandlerPopup(strFnName, strSection, Err.Description)
Exit Function

ErrorNothingToUpdate: '--------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: There is nothing in the table_usage table to update"
Call showErrorHandlerPopup(strFnName, strSection, "There are no results in the table_usage table.", "Check whether the files in 'script_filepath' table contain valid SQL code.")
Exit Function

End Function

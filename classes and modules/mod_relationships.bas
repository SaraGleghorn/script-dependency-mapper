Attribute VB_Name = "mod_relationships"
Option Compare Database
Option Explicit

Function sqlDependencyMapper() As Boolean
' Purpose: ********************************************************************
' Calls extractDependenciesFromFile for each file in script_filepath
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           31/03/2020  Original
' *****************************************************************************
' Expected Parameters:
' None
Definitions: '-----------------------------------------------------------------
Dim strFnName   As String           ' The name of this function (for debugging messages)
Dim strSection  As String           ' The name of the section (for debugging messages)

Dim db          As DAO.Database
Dim rsFileList  As DAO.Recordset
Dim FSO         As FileSystemObject

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
Call updateLog("Cleare table_usage table")

LoopThroughFiles: ' -----------------------------------------------------------
strSection = "LoopThroughFiles"

Do Until rsFileList.EOF

    If FSO.FileExists(rsFileList!filepath) = False Then
        Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
            & vbNewLine; "    Could not find "; rsFileList!filepath
        updateLog ("File not found: " & rsFileList!filepath)
    Else
        Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
            & vbNewLine; "    Preparing to parse: " & rsFileList!filepath
        extractDependenciesFromFile (rsFileList!filepath)
    End If

    rsFileList.MoveNext
Loop

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
Call updateLog("Populate table_usage table")
sqlDependencyMapper = True
Set rsFileList = Nothing
Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

ErrorNoFiles: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

End Function

Function populateFileList() As Boolean
' Purpose: ********************************************************************
' Loop through each folder, and dump a list of filenames into our local table.
' Requirements:
'   Reference -  Microsoft Scripting Runtime (for FileSystemObject)
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           24/03/2020  Original
' *****************************************************************************
' Expected Parameters:
Dim bLogInfo        As Boolean          ' Whether to log information
    bLogInfo = True

Definitions: ' ----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function (for debugging messages)
Dim strSection      As String           ' The name of the strSection (for debugging messages)
Dim db              As DAO.Database     ' This Database
Dim strSQL          As String           ' Text for strSQL query (easier debugging)
Dim rsFolderList    As DAO.Recordset    ' Will hold our list of folders
Dim strFolderPath   As String           ' The name of our folder
Dim colFiles        As Collection       ' Collection of filenames inside the folder
Dim strFilePath     As String           ' The name of our file as we iterate through our collection
Dim i               As Integer          ' For looping
Dim strMsg          As String           ' For error messaging
Dim FSO             As FileSystemObject ' For navigating files.
                                        ' Dir() is faster, but is global so can interupt Dir() in calling or called functions
Dim fsoFolder       As Folder

On Error GoTo ErrorHandler
strFnName = "populateFileList"

CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

Set db = CurrentDb

' Check table exists
If databaseObjectExists("script_folder", "table", db) = False Then GoTo ErrorMissingTable

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

RetrieveFiles: '---------------------------------------------------------------
strSection = "RetrieveFiles"

rsFolderList.MoveFirst
Do Until rsFolderList.EOF = True

    ' Add the trailing slash
    strFolderPath = rsFolderList!folder_path
    If Right(strFolderPath, 1) <> "\" Then strFolderPath = strFolderPath & "\"
       
    ' Retrieve the files
    Set colFiles = retrieveFilesInFolder(strFolderPath, rsFolderList!include_subfolders, "*.SQL")
    
    ' If the strFilePath does not already exist in our table, insert it.
    Do While colFiles.Count > 0
        strFilePath = colFiles(1)
        colFiles.Remove 1
        If DCount("*", "script_filepath", "filepath = '" & strFilePath & "'") = 0 Then
            strSQL = "INSERT INTO script_filepath VALUES ('" & strFilePath & "');"
            db.Execute strSQL
        End If
    Loop
    
    rsFolderList.MoveNext
Loop

Cleanup: ' -------------------------------------------------------------------
If updateLog("Populate script_filepath table") = False Then GoTo ErrorMakingLog

If bLogInfo = True Then Debug.Print Now() & " " & strFnName & ": " _
    & "Ended without errors."
Exit Function

ErrorHandler: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

ErrorMissingTable: ' ----------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "search_folders table doesn't exist. " _
    & "Cannot retrieve list of folders to search. Quitting."
Exit Function

ErrorBadFolder: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Could not find the following folders: " _
    & vbNewLine & strMsg _
    & "Quitting."
Exit Function

ErrorMakingLog: ' -------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Errors updating log table. Quitting."
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
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           01/04/2020  Original
' *****************************************************************************
' Expected Parameters:
'Dim arQuery()    As Variant  ' The query to parse, with the linenumber then the line content
'Dim strSource       As String   ' The query source or filepath
'Dim strSourceType   As String   ' What type of query source e.g. SQL File, Access QueryDef
Definitions: '-----------------------------------------------------------------
Dim strFnName           As String           ' The name of this function (for debugging messages)
Dim strSection          As String           ' The name of the section (for debugging messages)

Dim db                  As DAO.Database     ' This db
Dim strQry              As String           ' Query for our recordset
Dim strQryInsert        As String           ' Query for our insert statements, when a dependency is found.
Dim rsSeek              As DAO.Recordset    ' Recordset of phrases to look for
Dim iLine               As Integer          ' What line of our query are we on?
Dim iChar               As Integer          ' For looping through each character of our query
Dim strChar             As String           ' Holds the character we're currently lexing
Dim bSkipChar           As Boolean          ' Whether to skip adding this character to our phrase capture
Dim bStringData         As Boolean          ' Are we currently inside a text string (e.g. 'things' in: "select 'things' as stuff")
Dim bIdentifier         As Boolean          ' Are we currently inside an identifier string, where we need to extend our phrase beyond spaces? (e.g. [my column] in "SELECT [my column] FROM mytable;"
Dim bPhrase             As Boolean          ' Do we have a completed word?
Dim strPhrase           As String           ' Our parsed phrase
Dim strCommand          As String           ' What our query is e.g. SELECT, INSERT INTO, etc
Dim bCaptureSourceMode  As Boolean          ' Are we capturing a source?
Dim bCaptureAliasMode   As Boolean          ' Are we expecting an alias (prevents us capturing aliases as sources)
Dim bSubquery           As Boolean          ' Are we parsing a subquery?
Dim iBrackets           As Integer          ' How many brackets do we have open?
Dim arSubquery()        As Variant          ' Will hold our subquery.
Dim iSubqueryRow        As Integer
Dim strSubquery         As String           ' For holding our subquery content
Dim strPhraseSubstitute As String           ' If strPhrase contains characters that break Regex, then use this.

On Error GoTo ErrorHandler
strFnName = "extractDependenciesFromQuery"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

Initialise: ' -----------------------------------------------------------------
strSection = "Initialise"
strCommand = vbNullString
Set db = CurrentDb

' Move some of our syntax tables into a recordset for speed
' Queries have a structure - we'll have a command or CTE first, then look for our table sources
strQry = "SELECT * FROM syntax_phrases WHERE parse_type IN ('command', 'common table expression')"
Set rsSeek = db.OpenRecordset(strQry)
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
' We are only interested in tables - column names and conditions will not be parsed.

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
                ' It is either the end of the subquery or the end of the line. Add the line so far into an array
                arSubquery(0, iSubqueryRow) = arQuery(0, iLine)
                arSubquery(1, iSubqueryRow) = strSubquery
                strSubquery = vbNullString
            End If
            
            If iBrackets = 0 Then
                If extractDependenciesFromQuery(arSubquery, strSource, strSourceType) = False Then GoTo ErrorSubqueryParseFailed
                bSubquery = False
                strSubquery = vbNullString
                iSubqueryRow = 0
                ReDim arSubquery(1, iSubqueryRow)
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
                        & vbNewLine & "    " & "table_name, " _
                        & vbNewLine & "    " & "line_number, " _
                        & vbNewLine & "    " & "command " _
                        & vbNewLine & ") VALUES ( " _
                        & vbNewLine & "    " & "'" & strSourceType & "', " _
                        & vbNewLine & "    " & "'" & strSource & "', " _
                        & vbNewLine & "    " & "'" & strPhrase & "', " _
                        & vbNewLine & "    " & "'" & arQuery(0, iLine) & "', " _
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
                    Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
                    & strPhrase
                    
                    ' Find exact matches to this phrase
                    rsSeek.MoveFirst
                    If strCommand = "SELECT" And strPhrase = "INTO" Then
                        rsSeek.FindFirst "[parse_phrase] = 'SELECT INTO'" ' Seperable command
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
                                strQry = "SELECT * FROM syntax_phrases WHERE parse_type IN ('command')"
                            Case "command"
                                strCommand = rsSeek!parse_phrase
                                strQry = "SELECT * FROM syntax_phrases WHERE parse_type IN ('command','source','modify','union','into','common table expression','apply')"
                            Case "source"
                                strQry = "SELECT * FROM syntax_phrases WHERE parse_type IN ('command','source','union','condition','group', 'where', 'order')"
                                If strCommand = "SELECT INTO" Then strCommand = "SELECT"
                        End Select
                        bPhrase = False
                        strPhrase = vbNullString
                    Else
                        ' Check if there are any partial matches (e.g. INSERT INTO for INSERT)
                        rsSeek.MoveFirst
                        
                        strPhraseSubstitute = strPhrase
                        
                        Select Case strPhrase
                            Case "[", "?", "#", "*"
                                strPhraseSubstitute = Replace(strPhrase, "[", "[[]")
                                strPhraseSubstitute = Replace(strPhrase, "?", "[?]")
                                strPhraseSubstitute = Replace(strPhrase, "#", "[#]")
                                strPhraseSubstitute = Replace(strPhrase, "*", "[*]")
                        End Select
                        rsSeek.FindFirst "[parse_phrase] LIKE '" & strPhraseSubstitute & " *'" ' Include trailing space
                    
                        If Not rsSeek.NoMatch Then
                            ' This may be the first word in a multiple-word command.
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
'
' Requirements:
'
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           31/03/2020  Original
' *****************************************************************************
' Expected Parameters:
'Dim strFilePath    As String  ' The script to extract dependancies from

Definitions: '-----------------------------------------------------------------
Dim strFnName       As String           ' The name of this function (for debugging messages)
Dim strSection      As String           ' The name of the section (for debugging messages)
Dim db              As DAO.Database     ' This database
Dim strErrorSQL     As String           ' String for error logging query
Dim strTempPath     As String           ' The folder we'll store our copy in while parsing it.
Dim FSO             As FileSystemObject ' So we don't break any current global Dir() usage
Dim fileTemp        As file             ' For reading our file through FSO
Dim ts              As TextStream       ' Contents of our file
Dim strLine         As String           ' Holds one line of script
Dim iFileLine       As Integer          ' File linenumber
Dim iQueryLine      As Integer          ' Line number counter
Dim iComment        As Integer          ' For tracking multiline comments
Dim arQuery         As Variant          ' Holds our entire file
Dim iChar           As Integer
Dim strCleanLine    As String           ' For passing lines after multiline comments
Dim bQueryEnd       As Boolean
Dim bLiteralString  As Boolean          ' For escaping semicolons within code
Dim strRemainder    As String           ' Any code remaining after semicolon

On Error GoTo ErrorHandler
strFnName = "extractScriptDependencies"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"

' File Exists
Set FSO = New FileSystemObject
If FSO.FileExists(strFilePath) = False Then GoTo ErrorBadFile

' Temporary Directory Exists
strTempPath = DLookup("configuration", "config", "description = 'Temporary File Location'")
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
' Copy the file to a location in our C:\ drive so that we don't lock files for other users
strSection = "CopyFileToTemp"

strTempPath = strTempPath & Right(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\"))
FSO.CopyFile strFilePath, strTempPath, True
Set fileTemp = FSO.GetFile(strTempPath)

ParseFile: ' -------------------------------------------------------------

' Set an array to pass
ReDim arQuery(1, 0)
' Open and loop through our temp file, extracting all references to tables
Set ts = fileTemp.OpenAsTextStream(ForReading, TristateUseDefault)

Do While ts.AtEndOfStream = False
    ' Split into individual queries

    
    If strRemainder = vbNullString Then
        If iQueryLine > UBound(arQuery, 2) Then
            ReDim Preserve arQuery(1, iQueryLine)
        End If
        strLine = Trim(Replace(ts.ReadLine, vbTab, " "))
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
                        strCleanLine = strCleanLine & Mid(strLine, iChar + 1, 1)
                        iChar = iChar + 1 ' Because end comment takes up 2 characters
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
    ' Check for semi colons signalling the end of the query
    If InStr(strLine, ";") = 0 Then
        ' Do Nothing
    Else
        ' Check if it's inside a literal string
        If InStr(strLine, "'") = 0 Then
            bQueryEnd = True
            strRemainder = Trim(Mid(strLine, InStr(strLine, ";") + 1, Len(strLine) - (InStr(strLine, ";"))))
            strLine = Trim(Left(strLine, InStr(strLine, ";")))
        Else
            For iChar = 1 To Len(strLine)
                If Mid(strLine, iChar, 1) = "'" Then
                    bLiteralString = Not bLiteralString
                ElseIf Mid(strLine, iChar, 1) = ";" Then
                    If bLiteralString = False Then
                        bQueryEnd = True
                        strRemainder = Mid(strLine, iChar + 1, Len(strLine) - iChar)
                        strLine = Trim(Left(strLine, iChar))
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    
    If strLine <> vbNullString Then
        arQuery(0, iQueryLine) = iFileLine
        arQuery(1, iQueryLine) = strLine
        iQueryLine = iQueryLine + 1
    End If
    
    If bQueryEnd = True Then
        If extractDependenciesFromQuery(arQuery, strFilePath, "SQL File") = False Then
            strErrorSQL = "INSERT INTO parse_errors ( " _
                & vbNewLine & "    " & "#" & Format(Now(), "dd-mmm-yyyy hh:nn:ss") & "#, " _
                & vbNewLine & "    " & "'" & strFilePath & "'" _
                & vbNewLine & "    " & "Error parsing query within file." _
                & vbNewLine & "    " & "'" & iFileLine & "')"
            db.Execute strErrorSQL
            strErrorSQL = vbNullString
        End If
        iQueryLine = 0
        ReDim arQuery(1, iQueryLine)
        bQueryEnd = False
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
    ' Parse it
    If extractDependenciesFromQuery(arQuery, strFilePath, "SQL File") = False Then
        strErrorSQL = "INSERT INTO parse_errors ( " _
            & vbNewLine & "    " & "#" & Format(Now(), "dd-mmm-yyyy hh:nn:ss") & "#, " _
            & vbNewLine & "    " & "'" & strFilePath & "'" _
            & vbNewLine & "    " & "Error parsing query within file." _
            & vbNewLine & "    " & "'" & iFileLine & "')"
        db.Execute strErrorSQL
        strErrorSQL = vbNullString
    End If
    iQueryLine = 0
    ReDim arQuery(1, iQueryLine)
End If

Cleanup: ' --------------------------------------------------------------------
strSection = "Cleanup"
extractDependenciesFromFile = True

ts.Close
If FSO.FileExists(strTempPath) = True Then
    FSO.DeleteFile (strTempPath)
End If

Set FSO = Nothing
Erase arQuery

Exit Function

ErrorHandler: ' ----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
Exit Function

ErrorBadFile: ' ---------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Could not find the following file: " _
    & vbNewLine & "    " & """" & strFilePath _
    & vbNewLine & "    Quitting."
Exit Function

ErrorMakingDirectory:
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Errors while trying to make the directory: " _
    & vbNewLine & "    " & """" & strTempPath & """" _
    & vbNewLine & "    Quitting."
Exit Function
End Function


Function RegExTest() As Boolean
Dim regEx           As RegExp
Set regEx = New RegExp
regEx.Pattern = "\/\*.*\*\/"

Debug.Print regEx.Replace(" /* Comment /* Nested Comment */ End of line", vbNullString)

End Function


Function updateLog(strEventDescription As String) As Boolean
' Purpose: ********************************************************************
' Update the log table in this db
' Version Control:
' Vers  Author         Authoriser   Date        Change
' 1     Sara Gleghorn  --           23/03/2020  Original
' *****************************************************************************
' Expected Parameters:
'Dim strEventDescription    As String           ' A comma separated list of filetypes to record

Definitions: ' ----------------------------------------------------------------
Dim strFnName               As String           ' The name of this function (for debugging messages)
Dim strSection              As String           ' The name of the strSection (for debugging messages)
Dim db                      As Database         ' This Database
Dim strSQL                  As String           ' Update/Insert query text

On Error GoTo ErrorHandler
strFnName = "updateLog"
CheckPrerequisites: ' ---------------------------------------------------------
strSection = "CheckPrerequisites"
Set db = CurrentDb

updateLogTable: '---------------------------------------------------------------
strSection = "updateLogTable"

If DCount("*", "log_last_update", "event_description = '" & strEventDescription & "'") = 0 Then
    strSQL = "INSERT INTO log_last_update VALUES (" _
        & vbNewLine & "    '" & strEventDescription & "', " _
        & vbNewLine & "    #" & Format(Now(), "dd-mmm-yyyy hh:nn:ss") & "#, " _
        & vbNewLine & "    '" & Environ$("username") & "'" _
        & vbNewLine & ");"
Else
    strSQL = "UPDATE log_last_update SET" _
        & vbNewLine & "    event_date = #" & Format(Now(), "dd-mmm-yyyy hh:nn:ss") & "#, " _
        & vbNewLine & "    user_id = '" & Environ$("username") & "'" _
        & vbNewLine & "WHERE event_description = '" & strEventDescription & "';"
End If

db.Execute strSQL, dbFailOnError

updateLog = True

Exit Function

ErrorHandler: '----------------------------------------------------------------
Debug.Print Now() & " " & strFnName & "." & strSection & ": " _
    & "Error: " & Err.Description
End Function


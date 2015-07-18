Option Explicit

' How this is not another proof that doing VBA is a bad idea?
' Nevertheless, we'll try to make the scripts relying on Application.FileSearch works again.

' The interface of this YtoFileSearch class aims to stick to the original
' Application.FileSearch class interface.
' Cf is https://msdn.microsoft.com/en-us/library/office/aa219847(v=office.11).aspx

' For now it do not handle recursive search and only search for files.
' More precisely the following filters are not implemented:
' * SearchSubFolders
' * MatchTextExactly
' * FileType
' If that's something you need, please create an issue so we have a look at it.

' Our class attributes.
Private pDirectoryPath As String
Private pFileNameFilter As String
Private pFoundFiles As Collection

' Set the directory in which we will search.
Public Property Let LookIn(directoryPath As String)
    pDirectoryPath = directoryPath
End Property

' Allow to filter by file name.
Public Property Let fileName(fileName As String)
    pFileNameFilter = fileName
End Property

'Property to get all the found files.
Public Property Get FoundFiles() As Collection
    Set FoundFiles = pFoundFiles
End Property

' Reset the FileSearch object for a new search.
Public Sub NewSearch()
    'Reset the found files object.
    Set pFoundFiles = New Collection
    ' and the search criterions.
    pDirectoryPath = ""
    pFileNameFilter = ""
End Sub

' Launch the search and return the number of occurrences.
Public Function Execute() As Long
    'Lance la recherche
    doSearch

    Execute = pFoundFiles.Count
End Function

' Do the nasty work here.
Private Sub doSearch()
    Dim directoryPath As String
    Dim currentFile As String
    Dim filter As String
    
    directoryPath = pDirectoryPath
    If InStr(Len(pDirectoryPath), pDirectoryPath, "\") = 0 Then
        directoryPath = directoryPath & "\"
    End If

    ' If no directory is specified, abort the search.
    If Len(directoryPath) = 0 Then
        Exit Sub
    End If
    
    ' Check that directoryPath is a valid directory path.
    ' http://stackoverflow.com/questions/15480389/excel-vba-check-if-directory-exists-error
    If Dir(directoryPath, vbDirectory) = "" Then
        Debug.Print "Directory " & directoryPath & " does not exists"
        Exit Sub
    Else
        If (GetAttr(directoryPath) And vbDirectory) <> vbDirectory Then
            Debug.Print directoryPath & " is not a directory"
            Exit Sub
        End If
    End If
    
    ' We rely on the Dir() function for the search.
    ' cf https://msdn.microsoft.com/fr-fr/library/dk008ty4(v=vs.90).aspx
    
    ' Create the filter used with the Dir() function.
    filter = directoryPath

    If Len(pFileNameFilter) > 0 Then
        ' Add the file name filter.
        filter = filter & "*" & pFileNameFilter & "*"
    End If
    
    ' Start to search.
    currentFile = Dir(filter)
    Do While currentFile <> ""
        ' Use bitwise comparison to make sure currentFile is not a directory.
        If (GetAttr(directoryPath & currentFile) And vbDirectory) <> vbDirectory Then
            ' Add the entry to the list of found files.
            pFoundFiles.Add directoryPath & currentFile
        End If
        ' Get next entry.
        currentFile = Dir()
    Loop
End Sub

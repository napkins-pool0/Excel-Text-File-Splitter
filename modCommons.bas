Attribute VB_Name = "modCommons"
Option Explicit

'--- Function to build dictionary for use as variables---
Function buildTableDict() As Dictionary

    Dim ws As Worksheet                                                 'defines worksheet for variables source
    Dim dict As Dictionary                                              'defines dictionary to store variables
    Dim tbl As ListObject                                               'defines table source for variables
    Dim row As ListRow                                                  'defines row for loop
    Dim otable As clsTable                                              'defines object to define variables
    Dim tempArray As Variant                                            'defines array for holding prior to writing to file
    Dim tableName As String                                             'defines the names of table as key for dictionary
    
    'set variables and objects for use
    Set ws = ThisWorkbook.Sheets("Sheet1")                            'assigns worksheet source as variable
    Set tbl = ws.ListObjects(1)                                         'assigns table source as variable
    Set dict = New Dictionary                                           'creates dictionary for storing worksheet variables
    
    'loop through rows in table on Sheet1 to build definitions
    For Each row In tbl.ListRows
        tableName = row.Range.Columns(1)                                'assigns name of table as variable
        If tableName = "" Then GoTo nextRow
        Set otable = New clsTable                                       'creates new class object
        dict.Add tableName, otable                                      'stores new class object in dictionary
        otable.tblPrefix = row.Range.Columns(2)                         'stores table prefix that identifies table as part of class object
        otable.tblCtr = 0                                               'stores long as part of class object
        otable.tblName = row.Range.Columns(1)                           'stores retrievable name as part of class object
        ReDim tempArray(1 To 1) As Variant                              'creates temp array
        otable.tblArray = tempArray                                     'stores temp array as part of class object
nextRow:
    Next row
    
    Set buildTableDict = dict                                           'returns dictionary from function

End Function

' --- Creates empty text files, named from Sheet1, for loading data into ---
Sub fileCreator(folderpath, tableDict)

Dim oFolder As Object                                               'defines file object for loop
Dim ws As Worksheet                                                 'defines worksheet for variables source
Dim key As Variant                                                  'defines key used by dictionary
Dim otable As clsTable                                              'defines object to define variables
Dim ifile As Variant                                                'defines file number
    
    Call CheckFileExists(folderpath, "Table Files")
    MkDir (folderpath & "Table Files")
    
    'sets variables for use of loop
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(folderpath)      'assigns object to connect to folder for files requiring reading
    Set ws = ThisWorkbook.Sheets("Sheet1")                        'assigns worksheet for filename range
    
                              'create destination folder for table files
    'loop creates file in folder for each table name
    For Each key In tableDict
        Set otable = tableDict(key)                                 'set cls obj as variable
        ifile = FreeFile                                            'assigns variable with file number not already in use
        otable.tblPath = folderpath & "Table Files\" & otable.tblName & ".txt"      'assign cls obj variable as filepath for table file
        Open otable.tblPath For Output As #ifile                    'creates text file, referencing the file number variable
        Close #ifile                                                'closes file
    Next key

End Sub

' ---Counts the lines of the file, returning full line count of file---
Function countLines(fName As String) As Long
  countLines = CreateObject("Scripting.FileSystemObject").OpenTextFile(fName, 8, True).line

End Function

'---Pass cls obj array to table docs, clear array---
Function writeToTableText(tableDict)

    Dim key As Variant                                              'defines key used by dictionary
    Dim otable As clsTable                                          'defines object to define variables
    Dim ifile As Variant                                            'defines file number
    
    ReDim tempArray(1 To 1) As Variant                                              'create empty array
    For Each key In tableDict                                                       'cycle through each cls obj
        Set otable = tableDict(key)                                                 'set class obj from dict
        If otable.tblCtr = 0 Then GoTo nextKey                                      'check if cls obj array has been populated, if not move to next
        ifile = FreeFile                                                            'assigns variable with file number not already in use
        Open otable.tblPath For Append As #ifile                                    'open file for appending
        Print #ifile, Join$(otable.tblArray, vbLf)                                  'print array to file
        Close #ifile                                                                'close file
        otable.tblArray = tempArray                                                 'reset array in cls obj
        otable.tblCtr = 0                                                           'reset cls obj counter
        Set writeToTableText = tableDict
        Exit Function
nextKey:
    Next key
    
Set writeToTableText = tableDict

End Function

'---Loops through dict with line, looking for matching row---
Function lineCapture(tableDict, lineread)

    Dim key As Variant                                              'defines key used by dictionary
    Dim otable As clsTable                                          'defines object to define variables
    Dim tempArray As Variant                                        'defines array for holding prior to writing to file
    
    
    For Each key In tableDict                                                               'cycle through keys for matching key
        Set otable = tableDict(key)                                                         'set cls object from dict
        If Left(lineread, Len(otable.tblPrefix)) = otable.tblPrefix Then                    'if n char match the line prefix
            otable.tblCtr = otable.tblCtr + 1                                               'increment obj counter
            tempArray = otable.tblArray                                                     'set cls obj array as temp array
            ReDim Preserve tempArray(1 To otable.tblCtr) As Variant                         'redim temp array
            otable.tblArray = addToTable(lineread, otable.tblCtr, tempArray)                'set cls obj array to temp array
            Set lineCapture = tableDict                                                                       'next line
            Exit Function
        End If
    Next key                                                                                'else check next key
    Set lineCapture = tableDict
End Function

'--- Add line read to array---
Function addToTable(item, counter, arrays) As Variant

    arrays(counter) = item                                          'array point loaded with line
    addToTable = arrays                                             'return array with line

End Function


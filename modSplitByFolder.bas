Attribute VB_Name = "modSplitByFolder"
Option Explicit

'--- Split to tables ---
Sub splitFilesToTables(filepath As String, tableDict As Dictionary)

    Dim oFolder As Object, oFile As Object                          'defines folder and file objects for loop
    Dim filename As String                                          'defines file path in loop as string for searching
    Dim fileLines As Long, countBase As Long                        'defines file line count for parameters, how many lines read in each loop
    Dim readCycle As Long                                           'defines how many loops will occur
    Dim FSO As Object                                               'defines object for reading
    Dim i As Integer, fileCount As Integer, fileTotal As Integer, updateFigure As Integer    'defines integer for counting readcycles
    Dim lineCycleSize As Long                                       'defines count of lines to be read in loop
    Dim lineread As String                                          'defines string that is read from file
    Dim key As Variant                                              'defines key used by dictionary
    Dim otable As clsTable                                          'defines object to define variables
    Dim tempArray As Variant                                        'defines array for holding prior to writing to file
    Dim ifile As Variant                                            'defines file number

    
    'sets objects and files for use
    Call fileCreator(filepath, tableDict)                                                              'Creates tables for inputting into
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(filepath)
    fileTotal = oFolder.Files.Count
    ThisWorkbook.Sheets(1).Range("H11").Value = "0/" & fileTotal & " complete"
    
    'loops through each file in folder path, using only the files with Carriage Returns replaced
    For Each oFile In oFolder.Files
        filename = oFile.Path                                       'Assigns variable file path
        If ThisWorkbook.Sheets(1).Range("D3").Value = "No" Then
            GoTo skipCRCheck
        ElseIf InStr(filename, "CRReplaced") = 0 Then               'if variable isn't a file with carriage returns replaced
            GoTo nextiteration
skipCRCheck:
        End If
    
        'gets parameters of lines and read cycles in text file
        fileLines = countLines(filename)                            'assigns total lines count in file to variable
        countBase = ActiveWorkbook.Sheets("Sheet1").Range("E3").Value       'sets cycle size
        readCycle = Application.WorksheetFunction.RoundUp(fileLines / countBase, 0)     'gets read cycle (how many lines it will read before writing to file)
            
        'Opens file then loops through each line, copying to an array
        Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename)     'connects to text file
    
        For i = 1 To readCycle
            lineCycleSize = countBase * i                           'how many lines on this cycle to perform reading before loading
            If lineCycleSize > fileLines Then lineCycleSize = fileLines     'if last cycle, ensures doesn't look for non-existent lines
        
            Do While FSO.line < lineCycleSize
                lineread = FSO.Readline                             'assign variable the line of text
                Set tableDict = lineCapture(tableDict, lineread)    'add line read to table, return dict once complete
nextLoop:
            Loop
            
        Set tableDict = writeToTableText(tableDict)             'writes cls obj to text files, clears arrays
            
        If updateFigure < 25 Then
            If lineCycleSize > 0.25 * fileLines Then updateFigure = percIndicator(updateFigure)
        ElseIf updateFigure < 50 Then
            If lineCycleSize > 0.5 * fileLines Then updateFigure = percIndicator(updateFigure)
        ElseIf updateFigure < 75 Then
            If lineCycleSize > 0.75 * fileLines Then updateFigure = percIndicator(updateFigure)
        ElseIf updateFigure < 90 Then
            If lineCycleSize > 0.9 * fileLines Then updateFigure = percIndicator(updateFigure)
        End If
            
        Next i                                                      'next cycle
nextiteration:
    Set FSO = Nothing                                              'clear FSO obj variable
    fileCount = fileCount + 1
    ThisWorkbook.Sheets(1).Range("H11").Value = fileCount & "/" & fileTotal & " complete"
    Next oFile                                                      'next file
End Sub

'--- Function to remove all Carriage Returns ---
Function replaceFolderCRs()

    Dim oFolder As Object, oFile As Object                                                  'defines folder and file objects for loop
    Dim folder As String                                                                    'defines folder path
    Dim TSO As Object                                                                       'defines file in loop for reading
    Dim sText As String                                                                     'defines string text that function is performed on
    Dim ifile As Variant                                                                    'defines file number
    Dim filename As String                                                                  'defines filename in loop
    
    'calls function to navigate to folder and assign to variable
    folder = SelectFolder()                                                                 'folder containing split text of download

    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(folder)              'assigns object to connect to folder for files requiring reading
    If oFolder.Files.Count = 0 Then                                                         'Checks folder qty
        MsgBox "No files in folder"                                                         'informs no files in folder
        End                                                                                 'ends
    End If
    
    Call CheckFileExists(folder, "Output")
    MkDir (folder & "Output")
    folder = folder & "Output\"
    'searches through each folder for file to replace text on
    For Each oFile In oFolder.Files
        filename = oFile.Name
        If InStr(filename, "-CRReplaced") > 0 Then GoTo nextiteration                     'avoids attempting to perform full replace on already existing CR replaced docs
        If InStr(filename, ".txt") = 0 Then GoTo nextiteration                            'avoids attempting to perform full replace on files not wanted
        
        Set TSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(oFile)            'assigns object to read text file
        On Error Resume Next
        sText = TSO.ReadAll                                                                 'assigns variable with text file contents
        If sText = "" Then GoTo tSOErr
        On Error GoTo 0
        sText = Replace(sText, vbCr, "")                                                    'assigns variable with replaced carriage returns
        

        ifile = FreeFile                                                                    'assigns variable with file number not already in use

        Open folder & oFile.Name & "-CRReplaced" & ".txt" For Output As #ifile                       'creates text file, referencing the file number variable
        Print #ifile, sText                                                                 'writes CR replaced text to file
        Close #ifile                                                                        'closes file
tSOErr:
        Set TSO = Nothing
nextiteration:
    Next oFile
    
    replaceFolderCRs = folder                                                                     'returns folder path used to main routine
    Set oFolder = Nothing
End Function

'--- Function to select folder for use by tool ---

Function SelectFolder()
    Dim sFolder As String
    
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)                                  'With used for ease of commenting
        If .Show = -1 Then                                                                  'if OK is pressed
            sFolder = .SelectedItems(1)                                                     'assigns selected folder to variable
        End If
    End With
    
    If sFolder <> "" Then                                                                   'Checks if folder path obtained
        SelectFolder = sFolder & "\"                                                        'returns selected folder
    Else
        MsgBox "Folder not selected"                                                        'informs user folder not selected
        End                                                                                 'exits
    End If
    
End Function

Sub CheckFileExists(folder, folderName)

Dim strFileName As String
Dim strFileExists As String

    strFileName = folder & folderName & "\"
    strFileExists = Dir(strFileName)

   If strFileExists <> "" Then
        MsgBox "An existing folder that exists in the directory provided will cause conflict." & _
        " Delete the below folder to avoid conflict" & vbLf & vbLf & folderName
        End
    End If

End Sub

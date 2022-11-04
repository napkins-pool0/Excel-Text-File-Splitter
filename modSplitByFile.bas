Attribute VB_Name = "modSplitByFile"
Option Explicit

'--- Split to tables ---
Sub splitFileToTables(filename As String, tableDict As Dictionary)

    Dim oFolder As Object, oFile As Object                          'defines folder and file objects for loop
    Dim fileLines As Long, countBase As Long                        'defines file line count for parameters, how many lines read in each loop
    Dim readCycle As Long                                           'defines how many loops will occur
    Dim FSO As Object                                               'defines object for reading
    Dim i As Integer, ctr As Long, updateFigure As Integer          'defines integer for counting readcycles
    Dim lineCycleSize As Long                                       'defines count of lines to be read in loop
    Dim lineread As String                                          'defines string that is read from file
    Dim key As Variant                                              'defines key used by dictionary
    Dim otable As clsTable                                          'defines object to define variables
    Dim tempArray As Variant                                        'defines array for holding prior to writing to file
    Dim ifile As Variant                                            'defines file number
    Dim folderpath As String                                        'defines folder path to write files to
    Dim charCount As Integer, folderCount As Integer                'defines counts for determining folder path

    
    folderpath = Left(filename, InStrRev(filename, "\"))
    'sets objects and files for use
    Call fileCreator(folderpath, tableDict)                                                              'Creates tables for inputting into
    
    Set oFile = CreateObject("Scripting.FileSystemObject").GetFile(filename)

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
        Set FSO = Nothing                                           'clear FSO obj variable
End Sub

'--- Function to remove all Carriage Returns ---

Function replaceFileCRs()

    Dim oFile As Object                                                                     'defines objects for loop
    Dim file As String                                                                      'defines file path
    Dim TSO As Object                                                                       'defines file in loop for reading
    Dim sText As String                                                                     'defines string text that function is performed on
    Dim ifile As Variant                                                                    'defines file number
    Dim filename As String                                                                  'defines filename in loop

    'calls function to navigate to file and assign to variable
    file = selectFile()                                                                 'file containing split text of download

    Set oFile = CreateObject("Scripting.FileSystemObject").GetFile(file)              'assigns object to connect to file for files requiring reading
    filename = oFile.Name

    Set TSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(oFile)            'assigns object to read text file
    On Error Resume Next
    sText = TSO.ReadAll                                                                 'assigns variable with text file contents
    If sText = "" Then GoTo tSOErr
    On Error GoTo 0
    sText = Replace(sText, vbCr, "")                                                    'assigns variable with replaced carriage returns

    ifile = FreeFile                                                                    'assigns variable with file number not already in use

    Open oFile.Path & "-CRReplaced" & ".txt" For Output As #ifile                       'creates text file, referencing the file number variable
    Print #ifile, sText                                                                 'writes CR replaced text to file
    Close #ifile                                                                        'closes file

    Set TSO = Nothing
    Set oFile = Nothing
    replaceFileCRs = file                                                                     'returns file path used to main routine
    
    Exit Function
    
tSOErr:
    MsgBox "The file appears to be empty"
    End
End Function

'--- Function to select file for use by tool ---

Function selectFile()
    Dim sFile As String

    ' Open the select file prompt
    With Application.FileDialog(msoFileDialogFilePicker)                                  'With used for ease of commenting
        If .Show = -1 Then                                                                  'if OK is pressed
            sFile = .SelectedItems(1)                                                     'assigns selected file to variable
        End If
    End With

    If InStr(sFile, ".txt") = 0 Then
        MsgBox "File is not text file"
        End
    ElseIf ThisWorkbook.Sheets(1).Range("D3").Value = "Yes" Then
        If InStr(sFile, "-CRReplaced") > 0 Then
            MsgBox "File selected has already had carriage returns removed"
            End
        End If
    ElseIf sFile <> "" Then                                                                   'Checks if file path obtained
        selectFile = sFile                                                        'returns selected file
    Else
        MsgBox "File not selected"                                                        'informs user file not selected
        End                                                                                 'exits
    End If

End Function

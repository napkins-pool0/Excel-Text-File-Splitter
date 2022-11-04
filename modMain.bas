Attribute VB_Name = "modMain"
Option Explicit

'--- Main routine to execute tool ---
Sub mainSplit()
                                                                        'Defines answer from msgbox to fork
    Dim ans As Integer
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("H7:H11").Interior.ColorIndex = xlNone
    ws.Range("H11").Value = ""
    
    
    'queries if to conduct on single file, or folder
    ans = MsgBox("If you wish to perform this to a folder, press Yes. If to a single file, press No", (vbYesNo + vbQuestion))
    If ans = 6 Then Call byFolder Else Call byFile                      'forks to either perform by file or folder

End Sub

Sub byFolder()
    
    Dim folderpath As String                                            'defines folder for documents to be loaded into by tool
    Dim tableDict As Dictionary                                         'defines dictionary for variables

    
    If ThisWorkbook.Sheets(1).Range("D3") = "Yes" Then
        folderpath = replaceFolderCRs()                                     'removes carriage returns that interfere previously
    Else: folderpath = SelectFolder()
    End If
    
    Set tableDict = buildTableDict()                                    'builds dictionary
    Call splitFilesToTables(folderpath, tableDict)                      'allocates files to appropriate tables

End Sub

Sub byFile()

    Dim filepath As String                                              'defines file for documents to be loaded into by tool
    Dim tableDict As Dictionary                                         'defines dictionary for variables

    If ThisWorkbook.Sheets(1).Range("D3") = "Yes" Then
        filepath = replaceFileCRs()                                     'removes carriage returns that interfere previously
    Else: filepath = selectFile()
    End If
    
    Set tableDict = buildTableDict()                                    'builds dictionary
    Call splitFileToTables(filepath, tableDict)                         'allocates files to appropriate tables
    MsgBox "Done!", vbInformation
    
End Sub

Function percIndicator(updateFigure) As Integer

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    If updateFigure < 25 Then
        ws.Range("H7").Interior.ColorIndex = 43
        percIndicator = 25
    ElseIf updateFigure < 50 Then
        ws.Range("H8").Interior.ColorIndex = 43
        percIndicator = 50
    ElseIf updateFigure < 75 Then
        ws.Range("H9").Interior.ColorIndex = 43
        percIndicator = 75
    ElseIf updateFigure < 90 Then
        ws.Range("H10").Interior.ColorIndex = 43
        percIndicator = 90
    End If

End Function

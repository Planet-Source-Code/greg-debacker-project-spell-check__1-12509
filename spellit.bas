Attribute VB_Name = "SpellIt"
Option Explicit
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SW_SHOWNORMAL = 1

Public Declare Function PathFileExistsW Lib "shlwapi.dll" (ByVal pszPath As Long) As Boolean
Public Sub ParseForSpell(TheTreeView As TreeView, TheListView As ListView, TheProgMeter As ProgressBar)
    Dim sFileDir As String
    Dim sFileName As String
    Dim lFreeFile As Long
    Dim sThisLine As String
    Dim lSoFar As Long
    Dim lComntCt As Long
    Dim lTotComntCt As Long
    Dim lLineCt As Long
    Dim lWhiteCt As Long
    Dim lTotLineCt As Long
    Dim lTotWhiteCt As Long
    Dim iSubCt As Integer
    Dim iFuncCt As Integer
    Dim iTotSubCt As Integer
    Dim iTotFuncCt As Integer
    Dim iCtrlCt As Integer
    Dim iTotCtrlCt As Integer
    Dim sRoutineName As String
    Dim iStart As Integer
    Dim iFinish As Integer
    Dim bInRoutine As Boolean
    Dim bInControl As Boolean
    Dim bFoundFRX As Boolean
    Dim sCheckStr As String
    Dim sVarName As String
    Dim sType As String
    Dim sTemp As String
    Dim bShowFormName As Boolean
    Dim bShowSubName As Boolean
    Dim sDllName As String
    Dim lvwItem As ListItem
    Dim k As Integer
    Dim bGetComments As Boolean
    
    On Error Resume Next

    '
    'Dim sFuncTotal As String
    
    'sFuncTotal is not used at this point. There are lines of commented code that gather
    'all if the lines of code for each function. It would be nice to add the functions to
    'an array. Each function and sub could then be parsed to check for undeclared
    'variables or other anomalies.
    
    
    If MsgBox("Do you want to spell check comments?", vbYesNo, App.Title) = vbYes Then
        bGetComments = True
    Else
        bGetComments = False
    End If
    
    Const FILE_NAME = 0
    Const LINE_COUNT = 1            'Lines of code
    Const COMNT_COUNT = 2        'Lines of comments
    Const WHITE_COUNT = 3         'Blank lines
    Const SUB_COUNT = 4             'Number of subs in a file
    Const FUNC_COUNT = 5           'Number of functions in a file
    Const CONTROL_COUNT = 6     'Number of controls on a form
    Const DLLS = 7                         'DLLs used in the file
    
    'This array stores info about each of the files in the project
    ReDim sFileInfo(TheTreeView.Nodes.Count, DLLS + 1) As String
        
    Screen.MousePointer = 13
    TheListView.ListItems.Clear
    
    'Get the total size of all project files
    'Used for basic info and the progress bar
    For k = 1 To TheTreeView.Nodes.Count
        If PathFileExistsW(StrPtr(TheTreeView.Nodes(k).Key)) Then
            lLineCt = lLineCt + FileLen(TheTreeView.Nodes(k).Key)
        End If
    Next
    
    TheTreeView.Nodes.Add "properties", tvwChild, "totsize", "Total Size  - " & Bytes2KB(lLineCt), "info", "info"

    TheProgMeter.Min = 0
    TheProgMeter.Max = lLineCt
    
    lLineCt = 0
    
    'Loop through the TreeView control and go though each file
    For k = 1 To TheTreeView.Nodes.Count
        'The key contains the full path to the file
        If PathFileExistsW(StrPtr(TheTreeView.Nodes(k).Key)) Then
            sFileName = TheTreeView.Nodes(k).Key
            sFileInfo(k - 1, FILE_NAME) = sFileName
            
            lLineCt = 0
            lWhiteCt = 0
            lComntCt = 0
            'Because each file has properties it will be counted as a control.
            'Initialize the variable to -1 to account for this
            iCtrlCt = -1
            iSubCt = 0
            iFuncCt = 0
            bShowFormName = True
            bShowSubName = True
            bFoundFRX = False
            
            lFreeFile = FreeFile
            Open sFileName For Input As lFreeFile
                Do While Not EOF(lFreeFile)
                    Line Input #lFreeFile, sThisLine
                    
                    lSoFar = lSoFar + Len(sThisLine)
                    TheProgMeter.Value = lSoFar
                    
                    'Count comment lines, code lines, and white space
                    If Len(Trim$(sThisLine)) > 0 Then
                        If Left$(Trim$(sThisLine), 1) = Chr$(39) Then
                            lComntCt = lComntCt + 1
                        Else
                            lLineCt = lLineCt + 1
                            If InStr(sThisLine, Chr$(39)) > 0 Then
                                lComntCt = lComntCt + 1
                            End If
                        End If
                    Else
                        lWhiteCt = lWhiteCt + 1
                    End If
                    
                    'If bInRoutine = true then we are in a function or sub routine.
                    'Keep adding lines to sFuncTotal
                    
                    'If bInRoutine Then sFuncTotal = sFuncTotal & sThisLine
                    
                    sThisLine = Trim$(sThisLine)
                    'This is the start of a function - Get the name, increment a counter, and set some vars
                    If (Left$(sThisLine, 8) = "Function") Or (Left$(sThisLine, 16) = "Private Function") Or (Left$(sThisLine, 15) = "Public Function") Then
                        'sFuncTotal = ""
                        'sFuncTotal = sThisLine
                        
                        iFinish = InStr(1, sThisLine, "(")
                        iStart = InStrRev(sThisLine, " ", iFinish)
                        sRoutineName = Mid$(sThisLine, iStart, iFinish - iStart)
                        sType = "Function"
                        bInRoutine = True
                        bInControl = False
                        iFuncCt = iFuncCt + 1
                    'This is the of a sub routine
                    ElseIf (Left$(sThisLine, 4) = "Sub ") Or (Left$(sThisLine, 12) = "Private Sub ") Or (Left$(sThisLine, 11) = "Public Sub ") Then
                        'sFuncTotal = ""
                        'sFuncTotal = sThisLine
                        
                        iFinish = InStr(1, sThisLine, "(")
                        iStart = InStrRev(sThisLine, " ", iFinish)
                        sRoutineName = Mid$(sThisLine, iStart, iFinish - iStart)
                        sType = "Sub"
                        bInRoutine = True
                        bInControl = False
                        iSubCt = iSubCt + 1
                    
                    'This is the end of a sub or function
                    ElseIf (Left$(sThisLine, 7) = "End Sub") Or (Left$(sThisLine, 12) = "End Function") Then
                        'At this point we have reached the end of a function or sub routine.
                        'sFuncTotal contains the entire routine from start to finish.
                        'sFuncTotal
                        
                        bInRoutine = False
                        bShowSubName = True
                        
                    'The start of a control, menu, form, etc.
                    ElseIf (Left$(sThisLine, 6) = "Begin ") Then
                        iStart = InStrRev(sThisLine, " ", Len(sThisLine))
                        iFinish = Len(sThisLine) + 1
                        sRoutineName = Mid$(sThisLine, iStart, iFinish - iStart)
                        
                        iStart = InStr(sThisLine, " ") + 1
                        iFinish = InStr(iStart, sThisLine, " ")
                        sType = Mid$(sThisLine, iStart, iFinish - iStart)
                        If LCase$(Left$(sType, 3)) = "vb." Then sType = Right$(sType, Len(sType) - 3)
                        
                        bInRoutine = True
                        bInControl = True
                        iCtrlCt = iCtrlCt + 1
                        
                    'This is the end of a control, menu, form, etc.
                    ElseIf sThisLine = "End" Then
                        bInRoutine = False
                        bShowSubName = True
                        
                    'This is a DLL function declaration
                    ElseIf (Left$(sThisLine, 16) = "Private Declare ") Or (Left$(sThisLine, 15) = "Public Declare ") Or (Left$(sThisLine, 8) = "Declare ") Then
                        If InStr(sThisLine, " Lib ") > 0 Then
                            iStart = InStr(sThisLine, " Lib ") + 5
                            iFinish = InStr(iStart + 1, sThisLine, Chr$(34))
                            If iFinish > iStart Then
                                sDllName = Mid$(sThisLine, iStart + 1, iFinish - iStart - 1)
                                sFileInfo(k - 1, DLLS) = sFileInfo(k - 1, DLLS) & sDllName & Chr$(1)
                            End If
                        End If
                    Else
                    
                        ' Inside of a sub routine or function
                        If bInRoutine Then
                            sCheckStr = ""
                            
                            'This as a MsgBox call
                            'Get sVarName - In this case MsgBox
                            'Get sCheckStr - The assignment we will check the spelling on
                            If (Left$(sThisLine, 7) = "MsgBox(") Or (Left$(sThisLine, 7) = "MsgBox ") Then
                                iStart = InStr(sThisLine, "MsgBox ") + 8
                                iFinish = InStr(iStart, sThisLine, Chr$(34))
                                If iFinish > iStart Then
                                    sCheckStr = Mid$(sThisLine, iStart, iFinish - iStart)
                                    sVarName = Left$(sThisLine, InStr(sThisLine, " ") - 1)
                                End If
                                
                            'Next is a variable assignment or some other type of assignment
                            'It could be something like:
                            '   If (MyVar = "Blue") or (MyVar = "Red") then
                            'Or it could be variable that has multiple quoted parts separated by variables
                            '   sQuestion = "Is the color " & sClr & "? If so then bark."
                            'We loop though the line and get all of the quoted parts.
                            
                            'Get sVarName - The declared variable
                            'Get sCheckStr - 'The assignment we will check the spelling on
                            ElseIf InStr(sThisLine, " = ") > 0 Then
                                iStart = 1
                                sVarName = ""
                                If InStr(iStart, sThisLine, Chr$(34)) Then
                                    Do Until InStr(iStart, sThisLine, Chr$(34)) = 0
                                        iStart = InStr(iStart, sThisLine, Chr$(34)) + 1
                                        iFinish = InStr(iStart, sThisLine, Chr$(34))
                                        If iFinish > iStart Then
                                            sCheckStr = sCheckStr & " " & Mid$(sThisLine, iStart, iFinish - iStart)
                                            If Len(sVarName) = 0 Then
                                                If InStr(sThisLine, "MsgBox") Then
                                                    'This will deal with a line like
                                                    '   If MsgBox("Can you bark like a dog?", vbYesNo, App.Title) = vbYes Then
                                                    sVarName = "MsgBox"
                                                Else
                                                    sVarName = Trim$(Left$(sThisLine, InStr(sThisLine, " ") - 1))
                                                End If
                                            End If
                                        End If
                                        If iFinish > iStart Then iStart = iFinish + 1
                                    Loop
                                End If
                            End If
                            
                            'Here we are in a Control block
                            If (bInControl) Then
                                'Control blocks have their own comments
                                'Decrement the Comment line counter
                                If InStr(sThisLine, Chr$(39)) > 0 Then
                                    lComntCt = lComntCt - 1
                                End If
                            
                                'The list property of ComboBox or ListBox controls
                                'may contain list items add on the property page.
                                If (Trim$(sVarName) <> "Caption") And (Trim$(sVarName) <> "Text") Then
                                    If (sVarName = "List") And (InStr(sCheckStr, ".frx") > 0) Then
                                        If bFoundFRX = False Then
                                            bFoundFRX = True
                                            sVarName = Trim$(sCheckStr)
                                            'The GetFRXItems function parses the FRX file and returns and list items
                                            sFileDir = AddBackslash(GetDir(sFileName))
                                            sCheckStr = GetFRXItems(sFileDir & Trim$(sCheckStr))
                                        End If
                                    Else
                                        sCheckStr = ""
                                    End If
                                End If
                            End If
                                
                            'If the property assignment is an FRX file clear it.
                            If (InStr(sCheckStr, ".frx") > 0) Then sCheckStr = ""
                            
                            'This will get comments if the user said OK when the project was opened
                            If (Not bInControl) And (bGetComments) Then
                                If InStr(sThisLine, Chr$(39)) > 0 Then
                                    sCheckStr = Right$(sThisLine, Len(sThisLine) - InStr(sThisLine, Chr$(39)))
                                    If Len(Trim$(sCheckStr)) > 0 Then
                                        Set lvwItem = TheListView.ListItems.Add(, , " ")
                                        If bShowSubName Then
                                            bShowSubName = Not bShowSubName
                                            sTemp = sRoutineName
                                        Else
                                            sTemp = " "
                                            sType = ""
                                        End If
                                        
                                        'Add all of the other info
                                        lvwItem.SubItems(1) = sTemp
                                        lvwItem.SubItems(2) = sType
                                        lvwItem.SubItems(3) = "Comment"
                                        lvwItem.SubItems(4) = sCheckStr
                                        sCheckStr = ""
                                    End If
                                End If
                            End If
                            
                            'If sCheckStr has something then we will check the spelling on it.
                            If (Len(Trim$(sCheckStr)) > 0) Then
                                Select Case Asc(Left$(Trim$(sCheckStr), 1))
                                    Case 65 To 90, 97 To 122, 48 To 57, 32, 38
                                        'If bShowFormName = True then this is the first
                                        'time something from this form has been added
                                        If bShowFormName Then
                                            bShowFormName = Not bShowFormName
                                            sTemp = GetName(sFileName)
                                        Else
                                            sTemp = " "
                                        End If
                                        Set lvwItem = TheListView.ListItems.Add(, , sTemp)
                                        
                                        If Len(Trim$(sTemp)) > 0 Then
                                            Select Case Right$(LCase$(sTemp), 3)
                                                Case "frm"
                                                    lvwItem.SmallIcon = "form"
                                                Case "cls"
                                                    lvwItem.SmallIcon = "class"
                                                Case "bas"
                                                    lvwItem.SmallIcon = "module"
                                            End Select
                                        End If
                                        
                                        'If bShowSubName = True then this is the first time
                                        'something from this sub has been added
                                        If bShowSubName Then
                                            bShowSubName = Not bShowSubName
                                            sTemp = sRoutineName
                                        Else
                                            sTemp = " "
                                            sType = ""
                                        End If
                                        
                                        'Add all of the other info
                                        lvwItem.SubItems(1) = sTemp
                                        lvwItem.SubItems(2) = sType
                                        lvwItem.SubItems(3) = sVarName
                                        lvwItem.SubItems(4) = sCheckStr
                                End Select
                            End If
                        End If
                    End If
                Loop
            Close lFreeFile
            
            'Modules and classes, or forms with no controls will return iCtrlCt = -1
            If iCtrlCt < 0 Then iCtrlCt = 0
            
            'Store all of the counter info for this file.
            sFileInfo(k - 1, LINE_COUNT) = Str$(lLineCt)
            sFileInfo(k - 1, COMNT_COUNT) = Str$(lComntCt)
            sFileInfo(k - 1, WHITE_COUNT) = Str$(lWhiteCt)
            sFileInfo(k - 1, CONTROL_COUNT) = Str$(iCtrlCt)
            sFileInfo(k - 1, SUB_COUNT) = Str$(iSubCt)
            sFileInfo(k - 1, FUNC_COUNT) = Str$(iFuncCt)
        End If
    Next
    Screen.MousePointer = 0
    TheProgMeter.Value = 0
    
    'Add all of the crap we gathered from the files
    For k = 0 To TheTreeView.Nodes.Count - 1
        If Len(sFileInfo(k, FILE_NAME)) > 0 Then
            lLineCt = Val(sFileInfo(k, LINE_COUNT))
            lComntCt = Val(sFileInfo(k, COMNT_COUNT))
            lWhiteCt = Val(sFileInfo(k, WHITE_COUNT))
            iCtrlCt = Val(sFileInfo(k, CONTROL_COUNT))
            iSubCt = Val(sFileInfo(k, SUB_COUNT))
            iFuncCt = Val(sFileInfo(k, FUNC_COUNT))
            
            lTotLineCt = lTotLineCt + lLineCt
            lTotComntCt = lTotComntCt + lComntCt
            lTotWhiteCt = lTotWhiteCt + lWhiteCt
            iTotCtrlCt = iTotCtrlCt + iCtrlCt
            iTotSubCt = iTotSubCt + iSubCt
            iTotFuncCt = iTotFuncCt + iFuncCt
            
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "filesize", "Size  - " & Bytes2KB(FileLen(sFileInfo(k, FILE_NAME))), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "linecnt", "Code Lines - " & Format$(lLineCt, "#,###,##0"), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "cmntcnt", "Comment Lines  - " & Format$(lComntCt, "#,###,##0") & " - " & Format(lComntCt / lLineCt, "Percent"), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "whitect", "Blank Lines  - " & Format$(lWhiteCt, "#,###,##0") & " - " & Format(lWhiteCt / lLineCt, "Percent"), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "controls", "Controls  - " & Trim$(Str$(iCtrlCt)), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "subct", "Subs  - " & Trim$(Str$(iSubCt)), "info", "info"
            TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "func", "Functions - " & Trim$(Str$(iFuncCt)), "info", "info"
            
            'Add all of the DLL declarations. The same DLL may be used for several functions but it only
            'needs to be added once. Instead of checking for multiple calls to the same DLL file I just
            'let the On Error Resume Next work its magic. If the loop tries to add the same DLL name twice
            'an error will occur because it will not be a unique key property and the Add statement will fail.
            sTemp = sFileInfo(k, DLLS)
            If Len(sTemp) > 0 Then
                TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME), tvwChild, sFileInfo(k, FILE_NAME) & "dlls", "DLLs", "close", "open"
                Do Until Len(sTemp) = 0
                    sDllName = Left$(sTemp, InStr(sTemp, Chr$(1)) - 1)
                    
                    TheTreeView.Nodes.Add "dlls", tvwChild, sDllName & "dll", sDllName, "info", "info"
                    TheTreeView.Nodes.Add sFileInfo(k, FILE_NAME) & "dlls", tvwChild, sFileInfo(k, FILE_NAME) & sDllName, sDllName, "info", "info"
    
                    sTemp = Right$(sTemp, Len(sTemp) - Len(sDllName) - 1)
                Loop
            End If
       End If
    Next


    'Add all the totals to the Properties node
    TheTreeView.Nodes.Add "properties", tvwChild, "linecnt", "Code Lines  - " & Format$(lTotLineCt, "#,###,##0"), "info", "info"
    TheTreeView.Nodes.Add "properties", tvwChild, "cmntcnt", "Comment Lines - " & Format$(lTotComntCt, "#,###,##0") & " - " & Format(lTotComntCt / lTotLineCt, "Percent"), "info", "info"
    TheTreeView.Nodes.Add "properties", tvwChild, "whitect", "Blank Lines  - " & Format$(lTotWhiteCt, "#,###,##0") & " - " & Format(lTotWhiteCt / lTotLineCt, "Percent"), "info", "info"
    TheTreeView.Nodes.Add "properties", tvwChild, "controls", "Control Count  - " & Trim$(Str$(iTotCtrlCt)), "info", "info"
    TheTreeView.Nodes.Add "properties", tvwChild, "subcnt", "Sub Count  - " & Trim$(Str$(iTotSubCt)), "info", "info"
    TheTreeView.Nodes.Add "properties", tvwChild, "funcnt", "Function Count  - " & Trim$(Str$(iTotFuncCt)), "info", "info"

End Sub

Public Sub CheckSpelling(TheTreeView As TreeView, TheListView As ListView, TheProgMeter As ProgressBar)
    'Called from the toolbar of frmMain.
    
    If GetFileCount(TheTreeView) = 0 Then
        MsgBox "You need to open a project first.", 48, App.Title
        Exit Sub
    End If
    
    ParseForSpell TheTreeView, TheListView, TheProgMeter

    frmSpellIt.Show 1
End Sub
Public Sub ShowSpellReport(TheListView As ListView, TheComDlg As CommonDialog)
    
    Dim lFreeFile As Long
    Dim bNewFile As Boolean
    Dim sCurFile As String
    Dim bNewRoutine As Boolean
    Dim sCurRoutine As String
    Dim iRet As Long
    Dim DeskhWnd As Long
    Dim sMsg As String
    Dim bHaveOutPut As Boolean
    Dim k As Long
    
    'Display the final report.
    
    If TheListView.ListItems.Count = 0 Then
        MsgBox "Nothing to report.", 48, App.Title
        Exit Sub
    End If

    On Error Resume Next
    TheComDlg.FileName = "spell.log"
    TheComDlg.DialogTitle = "Save To Log File"
    TheComDlg.Filter = "Log Files|*.log|All Files|*.*"
    TheComDlg.ShowOpen
    
    If Err = 32755 Then Exit Sub
    
    If Len(Dir$(TheComDlg.FileName)) > 0 Then
        If MsgBox("This file exists. Do you want to overwrite it?", 36, App.Title) = 7 Then Exit Sub
    End If
   
    lFreeFile = FreeFile
    Open TheComDlg.FileName For Output As lFreeFile
        
        For k = 1 To TheListView.ListItems.Count
            If (Len(Trim$(TheListView.ListItems(k).Text)) > 0) Then
                bNewFile = True
                sCurFile = "************File: " & TheListView.ListItems(k).Text & " ************"
            End If
                
            If (Len(Trim$(TheListView.ListItems(k).SubItems(1))) > 0) Then
                bNewRoutine = True
                sCurRoutine = TheListView.ListItems(k).SubItems(2) & ": " & TheListView.ListItems(k).SubItems(1)
            End If
        
        
            'Only output info that has a check mark by it
            If TheListView.ListItems(k).Checked = True Then
                
                'Only output the File Name once
                If bNewFile Then
                    Print #lFreeFile, vbCrLf
                    Print #lFreeFile, sCurFile
                    bNewFile = Not bNewFile
                End If
                
                'Only output the routine or control name once
                If bNewRoutine Then
                    Print #lFreeFile, Tab(1); sCurRoutine
                    bNewRoutine = Not bNewRoutine
                End If
                
                'SubItems(3) is the Var/Contorl name
                'SubItems(4) is the original text
                'SubItems(5) is the corrected text
                If Len(Trim$(TheListView.ListItems(k).SubItems(3))) > 0 Then
                    Print #lFreeFile, "Variable/Property: " & TheListView.ListItems(k).SubItems(3)
                End If
                Print #lFreeFile, " Original Text: " & TheListView.ListItems(k).SubItems(4)
                Print #lFreeFile, "Corrected Text: " & TheListView.ListItems(k).SubItems(5)
                Print #lFreeFile, "_____________________________________________"
                bHaveOutPut = True
            End If
        Next
    Close lFreeFile

    
    If Not bHaveOutPut Then
        MsgBox "No changes were made. Nothing to report.", 64, App.Title
        Exit Sub
    End If
    
    DeskhWnd = GetDesktopWindow()
    
    'Shell to notepad with the file
    iRet = ShellExecute(DeskhWnd, vbNullString, "notepad.exe", TheComDlg.FileName, "", SW_SHOWNORMAL)
    If iRet < 32 Then
        
        Select Case iRet
            Case SE_ERR_FNF
                sMsg = "File not found."
            Case SE_ERR_PNF
                sMsg = "Path not found."
            Case SE_ERR_ACCESSDENIED
                sMsg = "Access denied."
            Case SE_ERR_OOM
                sMsg = "Out of memory."
            Case SE_ERR_DLLNOTFOUND
                sMsg = "DLL not found."
            Case SE_ERR_SHARE
                sMsg = "A sharing violation occurred."
            Case SE_ERR_ASSOCINCOMPLETE
                sMsg = "Incomplete or invalid file association."
            Case SE_ERR_DDETIMEOUT
                sMsg = "DDE Time out."
            Case SE_ERR_DDEFAIL
                sMsg = "DDE transaction failed."
            Case SE_ERR_DDEBUSY
                sMsg = "DDE busy."
            Case SE_ERR_NOASSOC
                sMsg = "No association for file extension."
            Case ERROR_BAD_FORMAT
                sMsg = "Invalid EXE file or error in EXE image."
            Case Else
                sMsg = "Unknown error"
        End Select
        If Len(sMsg) > 0 Then MsgBox sMsg, 48, App.Title
    End If


End Sub










Public Function GetFRXItems(ByVal sFRXFile As String) As String
    
    Dim lFreeFile As Long
    Dim sFileBody As String
    Dim sFinsihWord As String
    Dim bInWord  As Boolean
    Dim sThisLine As String
    Dim k As Integer
    
    'This function gets ListBox and ComboBox items
    'from the form files FRX file
    
    On Error GoTo ProblemFile
    'Try to read the whole file in pop
    'More often than not it generates an Input Past End Of File error
    'If so ProblemFile will read it line be line
    lFreeFile = FreeFile
    Open sFRXFile For Input As lFreeFile
        sFileBody = Input$(LOF(lFreeFile), lFreeFile)
    Close lFreeFile
        
    'The file has mostly binary info mixed with the readable listitems. I couldn't figure out a pattern
    'so I just go through it char by char and get the words. Crude but effective.
    For k = 1 To Len(sFileBody)
        Select Case Asc(Mid$(sFileBody, k, 1))
            'Alph/Numeric chars - everything but zero
            Case 65 To 90, 97 To 122, 49 To 57, 32
                sFinsihWord = sFinsihWord & Mid$(sFileBody, k, 1)
                bInWord = True
            Case Else
                'We've hit a non-printable char. Add a space to finish the word.
                If bInWord Then
                    sFinsihWord = sFinsihWord & " "
                    bInWord = Not bInWord
                End If
        End Select
    Next
    
    '"lt" is in the FRX files a lot. I don't know what it is but we don't want it.
    If Trim$(sFinsihWord) = "lt" Then
        sFinsihWord = ""
    Else
        If Left$(sFinsihWord, 2) = "lt" Then
            sFinsihWord = Right$(sFinsihWord, Len(sFinsihWord) - 2)
        End If
    End If
    
    'Return the space separated list of words
    GetFRXItems = sFinsihWord
    
    Exit Function
ProblemFile:
    Select Case Err
        Case 62
            Close lFreeFile
            lFreeFile = FreeFile
            Open sFRXFile For Input As lFreeFile
                Do While Not EOF(lFreeFile)
                    Line Input #lFreeFile, sThisLine
                    sFileBody = sFileBody & sThisLine
                Loop
            Close lFreeFile
            Resume Next
        Case Else
            MsgBox Error$
            Exit Function
    End Select

End Function



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSpellIt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spell Check"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgCount 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtfSpell 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSpellIt.frx":0000
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtSpell 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.ListBox lstCorrect 
      Height          =   1155
      IntegralHeight  =   0   'False
      Left            =   4080
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "Suggestions"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label 
      Caption         =   "Change To"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label 
      Caption         =   "Not In Dictionary"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   6
      Top             =   90
      Width           =   2295
   End
   Begin VB.Label lblFind 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmSpellIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private SpellIt As Word.Application
    Private SpDoc As Document
    Private SpErrors As SpellingSuggestions
    Private SplError As SpellingSuggestion
    Private bDontCheck As Boolean
    Private lStart As Long
    Dim sAlready() As String
    Dim iLineCt As Integer
    Dim sHotKey As String
Private Function AddHotKey(ByVal sCaption As String)
    Dim k As Integer
    If (Len(sHotKey) = 0) Or (InStr(sCaption, sHotKey) = 0) Then
        AddHotKey = sCaption
        Exit Function
    End If
    
    'If the original word had a hot key then we want to ad the ampersand into
    'the suggested words with the same hot key char.
    AddHotKey = Left$(sCaption, InStr(sCaption, sHotKey) - 1) & "&" & Right$(sCaption, Len(sCaption) - InStr(sCaption, sHotKey) + 1)
End Function
Private Function RemoveHotKey(ByVal sCaption As String)
    Dim k As Integer
    
    'Strip the ampersand and return the word
    If Len(sHotKey) = 0 Then
        RemoveHotKey = sCaption
        Exit Function
    End If
    
    RemoveHotKey = Left$(sCaption, InStr(sCaption, "&") - 1) & Right$(sCaption, Len(sCaption) - InStr(sCaption, "&"))
End Function

Private Sub CheckWords()
    Dim sCheckWord As String
    Dim lSpot As Long
    Dim lTempSpot As Long
    Dim lSpcSpot As Long
    Dim bLastWord As Boolean
    Dim lRetWords As Long


    If bDontCheck Then Exit Sub
    If Len(rtfSpell.Text) = 0 Then Exit Sub
    
    Screen.MousePointer = 13
    If lStart = 0 Then lStart = 1
    'Loop through the text box and get one word at a time and then spell check it
    If InStr(lStart, rtfSpell.Text, " ") Or (InStr(lStart, rtfSpell.Text, vbCrLf) > 0) Then
        Do Until (InStr(lStart, rtfSpell.Text, " ") = 0) And (InStr(lStart, rtfSpell.Text, vbCrLf) = 0)
            If lStart = 1 Then
                If Left$(rtfSpell.Text, 1) <> Chr$(32) Then
                    lSpot = 1
                Else
                    lSpot = InStr(lStart, rtfSpell.Text, " ") + 1
                End If
            Else
                lSpot = InStr(lStart, rtfSpell.Text, " ")
            End If
            lSpcSpot = InStr(lSpot + 1, rtfSpell.Text, " ")
            
            If (lSpcSpot = 0) Then
                lSpcSpot = Len(rtfSpell.Text) + 1
                bLastWord = True
            End If
            
            'Get the word
            sCheckWord = Mid$(rtfSpell.Text, lSpot, lSpcSpot - lSpot)
            
            'The FixWord function checks the word for a number of things
            lSpot = FixWord(lSpot, sCheckWord)
            
            If Len(Trim$(sCheckWord)) Then
                'GetSuggestions populates the list box with suggestions for the misspelled word
                lRetWords = GetSuggestions(sCheckWord)
                If (lRetWords > 0) Then
                    'If the list count is > 0 then the word was not found in the dictionary
                    If (lstCorrect.ListCount > 0) Then
                        'Show the word
                        lblFind = sCheckWord
                        'select it in the text box containing the var/or caption assignment
                        rtfSpell.SelStart = lSpot - 1
                        rtfSpell.SelLength = Len(sCheckWord) + Len(sHotKey)
                        'Save the end spot of the word in the text box
                        lStart = lSpcSpot - 1
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                ElseIf lRetWords < -1 Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                lSpcSpot = lSpcSpot + 1
            End If
            lStart = lSpcSpot
            sCheckWord = ""
        Loop
    End If
    
    If bLastWord Then GoTo FinishJob
    Screen.MousePointer = 0
    'If we got here then all of the words in the text box have been checked.
    'Click cmdStart_Click to select the next item and start the process again.
    cmdStart_Click
    Exit Sub
FinishJob:
    Screen.MousePointer = 0
    cmdStart_Click
    'MsgBox "Spell check is complete.", 64, App.Title
    

End Sub
Private Function FixWord(lSpot As Long, sWordToFix As String) As Long
'    On Error Resume Next
    Dim lCount As Long
    Dim bFoundOne As Boolean
    Dim k As Long
    
    
    If Len(sWordToFix) = 0 Then
        FixWord = lSpot
        Exit Function
    End If
    sHotKey = ""
    
    'If its a caption with a hot key then get the char
    'just after the ampersand
    If InStr(sWordToFix, "&") > 0 Then
        sHotKey = Mid$(sWordToFix, InStr(sWordToFix, "&") + 1, 1)
    End If
    
    'Get rid of puncuation, brackets, parenthesis, etc
    Select Case Asc(Right$(sWordToFix, 1))
        Case 33, 44, 46, 58, 59, 63, 125, 41, 93, 13, 10, 61
            sWordToFix = Left$(sWordToFix, Len(sWordToFix) - 1)
            'lSpot = lSpot - 1
    End Select
    
    If Len(sWordToFix) = 0 Then
        FixWord = lSpot
        Exit Function
    End If

    'Again for the other side
    Select Case Asc(Left$(sWordToFix, 1))
        Case 40, 91, 123, 13, 10
            sWordToFix = Right$(sWordToFix, Len(sWordToFix) - 1)
            'lSpot = lSpot + 1
    End Select
    
    If Len(sWordToFix) = 0 Then
        FixWord = lSpot
        Exit Function
    End If
    
    Select Case Asc(Left$(sWordToFix, 1))
        Case 32
            sWordToFix = Right$(sWordToFix, Len(sWordToFix) - 1)
            lSpot = lSpot + 1
    End Select
    
    'Strip any vbCrLf
    Do Until Left$(sWordToFix, 2) <> vbCrLf
        If Left$(sWordToFix, 2) = vbCrLf Then
            sWordToFix = Right$(sWordToFix, Len(sWordToFix) - 2)
            lSpot = lSpot + 2
        End If
    Loop
    
    'See if we've already confirmed this word to be spelled correctly
    lCount = UBound(sAlready)
    For k = 0 To lCount - 1
        If sAlready(k) = sWordToFix Then
            sWordToFix = ""
            bFoundOne = True
            Exit For
        End If
    Next
    
    'If not add it to the list.
    'If the word is incorrect it will be removed.
    If Not bFoundOne Then
        ReDim Preserve sAlready(lCount + 1)
        sAlready(lCount) = sWordToFix
    End If
        
    FixWord = lSpot
End Function


Private Function GetSuggestions(sWord As String) As Long
    lstCorrect.Clear
    txtSpell.Text = ""
    On Error GoTo NoWord2
    
    'Strip any ampersand from the word
    sWord = RemoveHotKey(sWord)
    'Check the Word 8 dictionary
    Set SpErrors = SpellIt.GetSpellingSuggestions(Word:=sWord)
    
    If InStr(sWord, "~-") > 0 Then GoTo NoWord2
    
    GetSuggestions = SpErrors.Count
    'The word was found in the dictionary and there are suggestions
    If SpErrors.Count > 0 Then
        'Loop through the words returned from the dictionary.
        'Add back an ampersand for the hot key if necessary.
        For Each SplError In SpErrors
          lstCorrect.AddItem AddHotKey(SplError.Name)
        Next SplError
        lstCorrect.Enabled = True
    'The word was spelled correctly - do nothing
    ElseIf SpellIt.CheckSpelling(Word:=sWord) Then
        GetSuggestions = -1
    'The word was not found in the dictionary and there was no suggestions
    Else
        GetSuggestions = 1
        lstCorrect.AddItem "(No Suggestion)"
        txtSpell.Text = sWord
        SelectIt txtSpell
        lstCorrect.Enabled = False
    End If
    
    'Select the first word in the list
    If (lstCorrect.ListCount) And (lstCorrect.Enabled) Then
        lstCorrect.ListIndex = 0
    End If
      
      
    Exit Function
NoWord2:
    Select Case Err
        Case -2147417851
            GetSuggestions = Err
            MsgBox "The file SPELL.Doc is missing or is corrupted. Reinstall this program, or recreate the file using Microsoft Word.", 48, App.Title
            Exit Function
        Case Else
            GetSuggestions = -1
            'MsgBox "Unexpected error.", 48, App.Title
            Exit Function
    End Select
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdChange_Click()
    Dim lCount As Long
    Dim k As Long
    lCount = UBound(sAlready)
    For k = 0 To lCount - 1
        If sAlready(k) = rtfSpell.SelText Then
            sAlready(k) = ""
            Exit For
        End If
    Next

    rtfSpell.SelText = txtSpell.Text
    frmMain.lvwOutPut.ListItems(iLineCt).SubItems(5) = rtfSpell.Text
    frmMain.lvwOutPut.ListItems(iLineCt).Checked = True
    CheckWords
End Sub


Private Sub cmdIgnore_Click()
    CheckWords
End Sub





Private Sub cmdStart_Click()
    'iLineCt equals the index of the current ListView item
    'Keep incrementing iLineCt until we've gone through the whole list.
    'Every time all of the words in the current list have been checked
    'the CheckWords sub calls cmdStart_Click to move to the next list item.
    
    iLineCt = iLineCt + 1
    If iLineCt > frmMain.lvwOutPut.ListItems.Count Then
        MsgBox "Finished", 64, App.Title
        Unload Me
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    
    prgCount.Value = iLineCt
    
    'Display the current file, sub, var
    If Len(Trim$(frmMain.lvwOutPut.ListItems(iLineCt).Text)) > 0 Then
        lblInfo(0).Caption = frmMain.lvwOutPut.ListItems(iLineCt).Text
    End If
        
    If Len(Trim$(frmMain.lvwOutPut.ListItems(iLineCt).SubItems(1))) > 0 Then
        lblInfo(1).Caption = frmMain.lvwOutPut.ListItems(iLineCt).SubItems(2) & " " & frmMain.lvwOutPut.ListItems(iLineCt).SubItems(1)
    End If
    
    If Len(Trim$(frmMain.lvwOutPut.ListItems(iLineCt).SubItems(3))) > 0 Then
        lblInfo(2).Caption = frmMain.lvwOutPut.ListItems(iLineCt).SubItems(3)
    End If
    
    'Populate the Spell Check text box with the current ListItem text
    rtfSpell.Text = ""
    rtfSpell.SelRTF = frmMain.lvwOutPut.ListItems(iLineCt).SubItems(4)
    DoEvents
    lStart = 0
    lblFind = ""
    'Check the words
    CheckWords

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    SpDoc.Close
    SpellIt.Quit
    Set frmSpellIt = Nothing
End Sub

Private Sub lstCorrect_Click()
    If lstCorrect.Text <> "(No Suggestion)" Then
        txtSpell.Text = lstCorrect.Text
    End If
End Sub

Private Sub lstCorrect_DblClick()
    cmdChange_Click
End Sub

Private Sub Form_Load()
    Dim sAppPath As String
    On Error GoTo NoWord
    CenterForm Me
    
    sAppPath = AddBackslash(App.Path)
    
    Set SpellIt = New Word.Application
    SpellIt.Visible = False
    Set SpDoc = SpellIt.Documents.Open(sAppPath & "spell.doc")
    
    prgCount.Min = 1
    prgCount.Max = frmMain.lvwOutPut.ListItems.Count
    
    'An array of common words
    'More words will be added
    'Check the array first before going to the dictionary
    ReDim Preserve sAlready(13)
    sAlready(0) = "a"
    sAlready(1) = "in"
    sAlready(2) = "the"
    sAlready(3) = "it"
    sAlready(4) = "with"
    sAlready(5) = "for"
    sAlready(6) = "is"
    sAlready(7) = "i"
    sAlready(8) = "of"
    sAlready(9) = "to"
    sAlready(10) = "get"
    sAlready(11) = "be"
    sAlready(12) = "not"
Exit Sub
NoWord:
Select Case Err
    Case 429
        bDontCheck = True
        MsgBox "The spell checker requires Microsoft Word 97 or greater.", 64, App.Title
        Unload Me
        Exit Sub
    Case Else
        Resume Next
End Select

End Sub

Private Sub txtSpell_Change()
    cmdChange.Enabled = Len(txtSpell.Text)
    cmdIgnore.Enabled = Len(txtSpell.Text)
End Sub



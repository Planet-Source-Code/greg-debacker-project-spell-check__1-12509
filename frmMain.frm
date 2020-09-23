VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project Spell Checker"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "spell"
            Object.ToolTipText     =   "Spell Check"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "report"
            Object.ToolTipText     =   "View Report"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPics 
      Left            =   6840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":065E
            Key             =   "report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0776
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AD6
            Key             =   "info"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E2A
            Key             =   "module"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":117E
            Key             =   "form"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D2
            Key             =   "class"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1826
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwFiles 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwOutPut 
      Height          =   5055
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlPics"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Control/Routine"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Variable/Property"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Value/Caption"
         Object.Width           =   15663
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Corrected Text"
         Object.Width           =   15663
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgFileOpen 
      Left            =   6360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Project"
      Filter          =   "Project Files|*.vbp|All Files|*.*"
   End
   Begin VB.Label lblProject 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub OpenProject()
    Dim lFreeFile As Long
    Dim sThisLine As String
    Dim iFormCt As Integer
    Dim iBasCt As Integer
    Dim iClassCt As Integer
    Dim sTemp As String
    Dim sProjPath As String
    Dim sProgName As String
    Dim sMajorVer As String
    Dim sMinorVer  As String
    Dim sRevisionVer  As String
            
    
    On Error Resume Next
    cdgFileOpen.FileName = ""
    cdgFileOpen.DialogTitle = "Open Project File"
    cdgFileOpen.Filter = "Project Files|*.vbp|All Files|*.*"
    
    cdgFileOpen.ShowOpen
    If Err = 32755 Then Exit Sub
    lblProject.Caption = cdgFileOpen.FileName
        
    'Clear the tree and add basic nodes.
    ResetNodes
    
    sProjPath = AddBackslash(GetDir(lblProject.Caption))
    lFreeFile = FreeFile
    
    'Open the project and get all the files and add them to the tree view.
    'The key of each node will have the full path to the file.
    Open lblProject.Caption For Input As lFreeFile
        Do While Not EOF(lFreeFile)
            Line Input #lFreeFile, sThisLine
            
            'Get Modules, classes, and forms
            If (Left$(sThisLine, 7) = "Module=") Then
                iBasCt = iBasCt + 1
                sTemp = Right$(sThisLine, Len(sThisLine) - InStr(sThisLine, "; ") - 1)
                tvwFiles.Nodes.Add "modules", tvwChild, sProjPath & sTemp, sTemp, "module", "module"
            ElseIf (Left$(sThisLine, 6) = "Class=") Then
                iClassCt = iClassCt + 1
                sTemp = Right$(sThisLine, Len(sThisLine) - InStr(sThisLine, "; ") - 1)
                tvwFiles.Nodes.Add "classes", tvwChild, sProjPath & sTemp, sTemp, "class", "class"
            ElseIf (Left$(sThisLine, 5) = "Form=") Then
                iFormCt = iFormCt + 1
                sTemp = Right$(sThisLine, Len(sThisLine) - InStr(sThisLine, "="))
                tvwFiles.Nodes.Add "forms", tvwChild, sProjPath & sTemp, GetName(sTemp), "form", "form"
            
            'Get references and objects
            ElseIf (Left$(sThisLine, 10) = "Reference=") Then
                If InStrRev(sThisLine, "#") > 0 Then
                    sThisLine = Right$(sThisLine, Len(sThisLine) - InStrRev(sThisLine, "#"))
                    tvwFiles.Nodes.Add "reference", tvwChild, sThisLine, sThisLine, "info", "info"
                End If
            ElseIf (Left$(sThisLine, 7) = "Object=") Then
                If InStrRev(sThisLine, ";") > 0 Then
                    sThisLine = Right$(sThisLine, Len(sThisLine) - InStrRev(sThisLine, ";"))
                    tvwFiles.Nodes.Add "object", tvwChild, sThisLine, Trim$(sThisLine), "info", "info"
                End If
                
            'Get the title and version info
            ElseIf (Left$(sThisLine, 6) = "Title=") Then
                sThisLine = Right$(sThisLine, Len(sThisLine) - 6)
                If Len(sThisLine) > 0 Then
                    sThisLine = StripQuotes(sThisLine)
                    sProgName = sThisLine
                End If
            ElseIf (Left$(sThisLine, 8) = "MajorVer") Then
                sMajorVer = Right$(sThisLine, Len(sThisLine) - 9)
            ElseIf (Left$(sThisLine, 8) = "MinorVer") Then
                sMinorVer = Right$(sThisLine, Len(sThisLine) - 9)
            ElseIf (Left$(sThisLine, 11) = "RevisionVer") Then
                sRevisionVer = Right$(sThisLine, Len(sThisLine) - 12)
            End If
        Loop
    Close lFreeFile

    'Add the title and version info if it was there
    If Len(sProgName) > 0 Then
        tvwFiles.Nodes(1).Text = sProgName & " " & sMajorVer & "." & sMinorVer & "." & sRevisionVer
    End If
    
    'Add in the counts of forms, classes, and modules
    tvwFiles.Nodes.Add "properties", tvwChild, "formcnt", Trim$(Str$(iFormCt)) & " Forms", "form", "form"
    tvwFiles.Nodes.Add "properties", tvwChild, "modcnt", Trim$(Str$(iBasCt)) & " Modules", "module", "module"
    tvwFiles.Nodes.Add "properties", tvwChild, "classcnt", Trim$(Str$(iClassCt)) & " Classes", "class", "class"
    
    tvwFiles.Nodes("classcnt").EnsureVisible
End Sub

Private Sub ResetNodes()
    tvwFiles.Nodes.Clear
     
    tvwFiles.ImageList = imlPics
    tvwFiles.Nodes.Add , , "project", "Project", "close", "open"

    tvwFiles.Nodes.Add "project", tvwChild, "forms", "Forms", "close", "open"
    tvwFiles.Nodes.Add "project", tvwChild, "modules", "Modules", "close", "open"
    tvwFiles.Nodes.Add "project", tvwChild, "classes", "Class Modules", "close", "open"
    tvwFiles.Nodes.Add "project", tvwChild, "properties", "Properties", "close", "open"
    tvwFiles.Nodes.Add "properties", tvwChild, "reference", "References", "close", "open"
    tvwFiles.Nodes.Add "properties", tvwChild, "object", "Objects", "close", "open"
    tvwFiles.Nodes.Add "properties", tvwChild, "dlls", "DLLs", "close", "open"
    
    tvwFiles.Nodes("forms").Sorted = True
    tvwFiles.Nodes("modules").Sorted = True
    tvwFiles.Nodes("classes").Sorted = True
    tvwFiles.Nodes("dlls").EnsureVisible

End Sub

Private Sub Form_Load()
    CenterForm Me
    ResetNodes
End Sub



Private Sub Form_Resize()
    Dim lTooHeight As Long
    On Error Resume Next
    
    lblProject.Width = Me.Width - 360
    prgProgress.Width = Me.Width - 360
    
    lTooHeight = tbrMain.Height + lblProject.Height + prgProgress.Height + 240
    tvwFiles.Move 120, lTooHeight, tvwFiles.Width, Me.Height - lTooHeight - 620
    lvwOutPut.Move tvwFiles.Width + 240, lTooHeight, Me.Width - tvwFiles.Width - 480, Me.Height - lTooHeight - 620
   
End Sub


Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "open": OpenProject
        Case "spell": CheckSpelling tvwFiles, lvwOutPut, prgProgress
        Case "report": ShowSpellReport lvwOutPut, cdgFileOpen
    End Select
End Sub

Private Sub tvwFiles_NodeClick(ByVal Node As MSComctlLib.Node)
    If PathFileExistsW(StrPtr(Node.Key)) Then MsgBox Node.Key
End Sub



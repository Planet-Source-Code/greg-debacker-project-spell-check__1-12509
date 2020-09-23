Attribute VB_Name = "Utils"
Option Explicit

Public Function AddBackslash(ByVal sPath As String) As String

    'Add a backslash to the end of a path is one does not exist.
    If Len(sPath) Then
       If Right$(sPath, 1) <> "\" Then
          AddBackslash = sPath & "\"
       Else
          AddBackslash = sPath
       End If
    Else
       AddBackslash = "\"
    End If
End Function

Public Function GetFileCount(TheTreeView As TreeView) As Integer
    'Checks all of the Node.Key property for file path and returns the count of files found
    Dim k As Integer
    Dim iCount As Integer
    For k = 1 To TheTreeView.Nodes.Count
        If PathFileExistsW(StrPtr(TheTreeView.Nodes(k).Key)) Then
            iCount = iCount + 1
        End If
    Next
    GetFileCount = iCount
End Function

Public Sub SelectIt(TxtBx As TextBox)
    'Select all of the text in a TextBox control
    TxtBx.SelStart = 0
    TxtBx.SelLength = Len(TxtBx.Text)
End Sub

Public Function Bytes2KB(ByVal lLongInt As Long) As String
    'Format bytes to kilobytes if the number of bytes are greater than 1024.
    
    If lLongInt > 0 Then
        lLongInt = lLongInt / 1024
        If lLongInt = 0 Then lLongInt = 1
    End If
    Bytes2KB = Format$(lLongInt, "#,###,###,##0KB")
End Function
Public Function StripQuotes(sMsg As String) As String
    'Strip the start and end quotes from a string
    sMsg = Trim$(sMsg)
    If Left$(sMsg, 1) = Chr$(34) Then sMsg = Right$(sMsg, Len(sMsg) - 1)
    If Right$(sMsg, 1) = Chr$(34) Then sMsg = Left$(sMsg, Len(sMsg) - 1)
    StripQuotes = sMsg
End Function
Public Function GetDir(ByVal sFileName As String) As String

    'Return the directory part of a path/file statement
    Dim k As Long
    If sFileName = "" Then
        GetDir = CurDir$
    Else
        For k = Len(sFileName) To 1 Step -1
            If Mid$(sFileName, k, 1) = "\" Then
                GetDir = Left$(sFileName, k)
                Exit For
            End If
        Next
    End If
End Function
Public Function GetName(sFileName As String) As String
    'Return the file name part of a file path
    
    Dim k As Integer
    GetName = sFileName
    k = InStrRev(sFileName, "\")
    If k > 0 Then GetName = Right$(sFileName, Len(sFileName) - k)
End Function
Public Sub CenterForm(F As Form)
    F.Left = (Screen.Width - F.Width) / 2
    F.Top = (Screen.Height - F.Height) / 2
End Sub

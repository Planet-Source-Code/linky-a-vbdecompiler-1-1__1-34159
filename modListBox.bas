Attribute VB_Name = "modListBox"
Option Explicit

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const LB_FINDSTRINGEXACT = &H1A2

Public Sub SaveListBox(theBox As ListBox, thePath As String, Optional bAppend As Boolean)
Dim i As Long, Freed As Integer, sLine As String
On Error GoTo NoFile
    Let Freed = FreeFile
    If bAppend Then
        Open thePath For Append As #Freed
    Else
        Open thePath For Output As #Freed
    End If
    For i = 0 To theBox.ListCount - 1
        sLine = theBox.List(i)
        sLine = Replace(Trim(sLine), Chr(0), vbNullString)
        If sLine <> vbNullString Then
            Print #Freed, sLine
        End If
    Next i
NoFile:
Close #Freed
End Sub

Public Sub LoadListBox(lstBox As ListBox, sPath As String)
On Error GoTo NoFile
Dim i As Long, iFreed As Integer
Dim sData As String
    iFreed = FreeFile
    Open sPath For Input As #iFreed
        Do While Not EOF(iFreed)
            Line Input #iFreed, sData
            If sData <> vbNullString Then
                sData = Trim(sData)
                lstBox.AddItem sData
            End If
        Loop
NoFile:
    Close #iFreed
End Sub

Public Function KillDupesAPI(lpBox As Control) As Long
On Error Resume Next
Dim nCount As Long, nPos1 As Long
Dim nPos2 As Long, nTotal As Long
    nTotal = lpBox.ListCount
    For nCount = 0 To nTotal
        Do
            nPos1 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nCount, lpBox.List(nCount))
            nPos2 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nPos1 + 1, lpBox.List(nCount))
            If Trim(lpBox.List(nCount)) = vbNullString Then lpBox.RemoveItem nCount
            If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
            lpBox.RemoveItem nCount
        Loop
    Next nCount
    KillDupesAPI = nTotal - lpBox.ListCount
End Function


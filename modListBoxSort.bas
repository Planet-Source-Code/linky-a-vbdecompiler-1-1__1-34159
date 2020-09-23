Attribute VB_Name = "modListBoxSort"
Option Explicit

Public Sub BubbleSort(myList As ListBox, Raising As Boolean)
    Do Until Sorted(myList, Raising)
        Bubble myList, Raising
    Loop
End Sub

Public Sub Sort(myList As ListBox, Optional Raising As Boolean)
    If IsMissing(Raising) Then Raising = True
    BubbleSort myList, Raising
End Sub

Public Function Sorted(myList As ListBox, Optional Raising As Boolean) As Boolean
    If IsMissing(Raising) Then Raising = True
    Dim StrCompVal As Integer
    Sorted = True
    
    If Raising Then
        StrCompVal = 1
    Else
        StrCompVal = -1
    End If
    
    Dim i As Integer
    For i = 0 To myList.ListCount - 2
        If StrComp(myList.List(i), myList.List(i + 1), vbTextCompare) = StrCompVal Then
            Sorted = False
            Exit Function
        End If
    Next i
End Function

Public Sub Bubble(myList As ListBox, Raising As Boolean)
    Dim i As Integer
    i = 0
    If Raising Then
        Do Until StrComp(myList.List(i), myList.List(i + 1), vbTextCompare) = 1
            i = i + 1
            If i > myList.ListCount - 2 Then
                Exit Sub
            End If
        Loop
    Else
        Do Until StrComp(myList.List(i), myList.List(i + 1), vbTextCompare) = -1
            i = i + 1
            If i > myList.ListCount - 2 Then
                Exit Sub
            End If
        Loop
    End If
    'do the bubble i and i+1
    Dim tmpStr As String
    tmpStr = myList.List(i)
    myList.RemoveItem i
    myList.AddItem tmpStr, i + 1
End Sub




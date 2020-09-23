Attribute VB_Name = "modVB5Comptibility"

Option Explicit

Public Function pInstrRev(sExpression As String, sFind As String, Optional ByVal iStart As Long = -1) As Long
Dim iPos As Long
Dim iPos2 As Long
Dim sInRev As String
    sInRev = sExpression
    If iStart < 1 Then iStart = Len(sInRev)
    sInRev = Left(sInRev, iStart)
    iPos2 = Len(sInRev)
    Do
        iPos = (InStr(iPos2, sInRev, sFind))
        iPos2 = iPos2 - 1
    Loop Until iPos > 0 Or iPos2 = 0
    pInstrRev = iPos
End Function

Public Function pReplace(sText As String, sFind As String, sReplace As String) As String
    Dim iStart As Long, iNextPos As Long
    iStart = 1: iNextPos = 1
    Do
        iStart = InStr(iNextPos, LCase(sText), LCase(sFind))
        If iStart <> 0 Then
            sText = Left$(sText, iStart - 1) _
            & sReplace & Mid$(sText, Len(sFind) + iStart)
        End If
        iNextPos = iStart + Len(sReplace) + 1
    Loop Until iStart = 0
    pReplace = sText
End Function




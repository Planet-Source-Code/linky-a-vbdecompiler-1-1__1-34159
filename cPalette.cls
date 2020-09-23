VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type RGBQUAD
   Red As Byte
   Green As Byte
   Blue As Byte
   Reserved As Byte
End Type
Private m_tPal() As RGBQUAD
Private m_iPalette As Long


Friend Sub fAdd(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    m_iPalette = m_iPalette + 1
    ReDim Preserve m_tPal(1 To m_iPalette) As RGBQUAD
    m_tPal(m_iPalette).Red = Red
    m_tPal(m_iPalette).Green = Green
    m_tPal(m_iPalette).Blue = Blue
End Sub

Public Property Get Count() As Long
' Number of colours in this palette:
   Count = m_iPalette
End Property
Public Property Get Red(ByVal iIndex As Long) As Byte
' Red value for palette index
   Red = m_tPal(iIndex).Red
End Property
Public Property Get Green(ByVal iIndex As Long) As Byte
' Green value for palette index
   Green = m_tPal(iIndex).Green
End Property
Public Property Get Blue(ByVal iIndex As Long) As Byte
' Blue value for palette index
   Blue = m_tPal(iIndex).Blue
End Property
Public Property Let Red(ByVal iIndex As Long, ByVal bRed As Byte)
' Red value for palette index
   m_tPal(iIndex).Red = bRed
End Property
Public Property Let Green(ByVal iIndex As Long, ByVal bGreen As Byte)
' Green value for palette index
   m_tPal(iIndex).Green = bGreen
End Property
Public Property Let Blue(ByVal iIndex As Long, ByVal bBlue As Byte)
' Blue value for palette index
   m_tPal(iIndex).Blue = bBlue
End Property

Friend Sub CreateWebSafe()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim l As Long, i As Long

   ' Websafe (IE 216 colour) palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256) As RGBQUAD
   p16ColourLow8 1
   p16ColourHigh8 248
   lIndex = 8
   For b = 0 To &HFF Step &H33
      For g = 0 To &HFF Step &H33
         For r = 0 To &HFF Step &H33
            ' ignore if the output is any combination of 0 and FF
            l = r + g + b
            If l = 0 Or l = &H2FD Then
               ' ignore
            ElseIf l = &H1FE And (r = 0 Or g = 0 Or b = 0) Then
               ' ignore
            ElseIf l = &HFF And ((r = 0 And g = 0) Or (r = 0 And b = 0) Or (g = 0 And b = 0)) Then
               ' ignore
            Else
               ' add
               lIndex = lIndex + 1
               With m_tPal(lIndex)
                  .Red = r: .Green = g: .Blue = b
               End With
            End If
         Next
      Next
   Next
   ' Fill the remain entries with gray shades:
   r = 8: g = 8: b = 8
   For i = 217 To 247
      With m_tPal(lIndex)
         .Red = r: .Green = g: .Blue = b
         r = r + 8: g = g + 8: b = b + 8
      End With
   Next i
   
End Sub
Friend Sub CreateHalfTone()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim rA As Long, gA As Long, bA As Long
Dim l As Long, i As Long

   ' Halftone 256 colour palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256) As RGBQUAD
   For b = 0 To &H100 Step &H40
      If b = &H100 Then
         bA = b - 1
      Else
         bA = b
      End If
      For g = 0 To &H100 Step &H40
         If g = &H100 Then
            gA = g - 1
         Else
            gA = g
         End If
         For r = 0 To &H100 Step &H40
            If r = &H100 Then
               rA = r - 1
            Else
               rA = r
            End If
            lIndex = lIndex + 1
            With m_tPal(lIndex)
               .Red = rA: .Green = gA: .Blue = bA
            End With
         Next r
      Next g
   Next b
   
End Sub
Friend Sub CreateMono()
   ' Monochrome palette
   m_iPalette = 2
   ReDim m_tPal(1 To 2) As RGBQUAD
   With m_tPal(2)
      .Blue = 255
      .Green = 255
      .Red = 255
   End With
End Sub
Private Sub p16ColourLow8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 1
   With m_tPal(lStartIndex + 2)
      .Red = &H80: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 3)
      .Red = 0: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 4)
      .Red = &H80: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 5)
      .Red = 0: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 6)
      .Red = &H80: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 7)
      .Red = 0: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 8)
      .Red = &HC0: .Green = &HC0: .Blue = &HC0
   End With

End Sub
Private Sub p16ColourHigh8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 9
   With m_tPal(lStartIndex + 9)
      .Red = &H80: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 10)
      .Red = &HFF: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 11)
      .Red = 0: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 12)
      .Red = &HFF: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 13)
      .Red = 0: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 14)
      .Red = &HFF: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 15)
      .Red = 0: .Green = &HFF: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 16)
      .Red = &HFF: .Green = &HFF: .Blue = &HFF
   End With
End Sub
Friend Sub Create16Colour()
   ' Standard EGA style 16 colour palette:
   m_iPalette = 16
   ReDim m_tPal(1 To 16) As RGBQUAD
   p16ColourLow8 1
   p16ColourHigh8 9
End Sub
Friend Property Get ClosestIndex( _
      ByVal Red As Long, _
      ByVal Green As Long, _
      ByVal Blue As Long _
   ) As Long
Dim lMinIndex As Long
   

Dim i As Long
Dim lER As Long, lEB As Long, lEG As Long
Dim lMinER As Long, lMinEB As Long, lMinEG As Long
   
   lMinER = 255: lMinEB = 255: lMinEG = 255
   For i = 1 To m_iPalette
      With m_tPal(i)
         If (Red = .Red) And (Blue = .Blue) And (Green = .Green) Then
            ClosestIndex = i
            Exit Property
         Else
            lER = Abs(Red - .Red)
            lEB = Abs(Blue - .Blue)
            lEG = Abs(Green - .Green)
            If (lER + lEB + lEG < lMinER + lMinEB + lMinEG) Then
               lMinER = lER
               lMinEB = lEB
               lMinEG = lEG
               lMinIndex = i
            End If
         End If
      End With
   Next i
   ClosestIndex = lMinIndex
      
End Property


Friend Sub GetPaletteFromIcon(ByRef cFI As cFileIcon, ByVal nIndex As Long)
   cFI.GetPalette nIndex, VarPtr(m_tPal(1))
End Sub
Friend Sub SetPaletteToIcon(ByRef cFI As cFileIcon, ByVal nIndex As Long)
   cFI.SetPalette nIndex, VarPtr(m_tPal(1))
End Sub

Public Function SaveToJASCFile(ByVal sFileName As String) As Boolean
Dim i As Long
Dim sJasc As String
Dim iFile As Long

   On Error Resume Next
   Kill sFileName
   
   sJasc = "JASC" & vbCrLf & "0100" & vbCrLf & m_iPalette
   For i = 1 To m_iPalette
      sJasc = sJasc & vbCrLf & m_tPal(i).Red & " " & m_tPal(i).Green & " " & m_tPal(i).Blue
   Next i
   
   On Error GoTo ErrorHandler
   iFile = FreeFile
   Open sFileName For Binary Access Write As #iFile
   Put #iFile, , sJasc
   Close #iFile
   iFile = 0
   Exit Function
   
ErrorHandler:
   Err.Raise Err.Number, App.EXEName & ".cPalette", Err.Description
   If (iFile <> 0) Then
      Close #iFile
   End If
   Exit Function
   
End Function

Public Function LoadFromJASCFile(ByVal sFileName As String) As Boolean
' Loads a JASC style .Pal file
' Not quick because it is a validating loader with error handling.
' The Native format is considerably quicker
Dim iFile As Long
Dim sBuf As String
Dim iPos As Long
Dim iNextPos As Long
Dim sLines() As String, iLineCount As Long, iLine As Long
Dim sParts() As String, iPartCount As Long
   
   m_iPalette = 0
   Erase m_tPal

   On Error GoTo LoadFromFileError
   ' Load it:
   iFile = FreeFile
   Open sFileName For Binary Access Read As #iFile
   sBuf = String$(LOF(iFile), 32)
   Get #iFile, , sBuf
   Close #iFile
   iFile = 0
   
   SplitDelimitedString sBuf, vbCrLf, sLines(), iLineCount
   For iLine = 1 To iLineCount
      SplitDelimitedString sLines(iLine), " ", sParts(), iPartCount
      If (iPartCount = 3) Then
         m_iPalette = m_iPalette + 1
         ReDim Preserve m_tPal(1 To m_iPalette) As RGBQUAD
         With m_tPal(m_iPalette)
            .Red = CByte(sParts(1))
            .Green = CByte(sParts(2))
            .Blue = CByte(sParts(3))
         End With
      End If
   Next iLine
   
   LoadFromJASCFile = True
   
   Exit Function

LoadFromFileError:
   Err.Raise Err.Number, App.EXEName & ".cPalette", Err.Description
   If (iFile <> 0) Then
      Close #iFile
      m_iPalette = 0
      Erase m_tPal
   End If
   Exit Function

End Function

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )

Dim iPos As Long
Dim iNextPos As Long
Dim iDelimLen As Long

    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub




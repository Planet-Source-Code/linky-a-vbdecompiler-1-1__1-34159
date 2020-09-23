Attribute VB_Name = "modFunctions"

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFileAPI Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Global cLanguage As New clsLangage

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Function ReadFile(PathName As String) As String
Dim hFile As Long
Dim RawData() As Byte
Dim ReadString As String
Dim FileLength  As Long
Dim ActualBytes As Long
Dim ret As Long
    
FileLength = FileLen(PathName)
ReDim RawData(FileLength)
    
hFile = CreateFile(PathName & vbNullChar, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)

If hFile = 0 Then
    MsgBox "Erreur d'ouverture"
    Exit Function
End If
    
ret = ReadFileAPI(hFile, RawData(0), FileLength, ActualBytes, 0)

If ret = 0 Or ActualBytes <> FileLength Then
    MsgBox "Erreur de lecture"
    CloseHandle hFile
    Exit Function
End If

CloseHandle hFile
ReadFile = StrConv(RawData, vbUnicode)
End Function
'-----------------------------------------------------------------------
Public Function FitText(ByRef Ctl As Control, _
                        ByVal strCtlCaption) As String
Dim lngCtlLeft As Long
Dim lngMaxWidth As Long
Dim lngTextWidth As Long
Dim lngX As Long

lngCtlLeft = Ctl.Left
lngMaxWidth = Ctl.Width
lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)


lngX = (Len(strCtlCaption) \ 2) - 2
While lngTextWidth > lngMaxWidth And lngX > 3
    strCtlCaption = Left(strCtlCaption, lngX) & "..." & _
                    Right(strCtlCaption, lngX)
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
    lngX = lngX - 1
Wend

FitText = strCtlCaption

End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, _
                               Optional ByVal strFormatMask As String) _
                               As String

Select Case dblFileSize
    Case 0 To 1023              ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575        ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823#       ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String

Select Case sglTime
    Case 0 To 59
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599
    FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function DiskFreeSpace(strDrive As String) As Double

Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim NumberOfFreeClusters As Long
Dim TotalNumberOfClusters As Long
Dim FreeBytes As Long
Dim spaceInt As Integer

strDrive = QualifyPath(strDrive)

GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector

End Function


Public Function QualifyPath(strPath As String) As String
QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")
End Function


Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))
End If

End Function



Public Function isFilename(sFileName As String, sExtention As String) As Boolean
On Error Resume Next
    Dim i As Integer, x As Integer
    Dim iLen As Integer
    
    If InStr(sFileName, "\") Then
        sFileName = Mid(sFileName, _
        InStr(sFileName, "\") + 1)
                                     
                                     
    End If
    
    
    For i = Len(sFileName) To 1 Step -1
        For x = 1 To 39
            If Mid$(sFileName, i, 1) = Chr$(x) _
            Or Mid$(sFileName, i, 1) = Chr$(96) Then
                sFileName = Mid$(sFileName, i + 1)
                Exit For
            End If
        Next x
    Next i
    For i = Len(sFileName) To 1 Step -1
        For x = 123 To 255
            If Mid$(sFileName, i, 1) = Chr$(x) _
            Or Mid$(sFileName, i, 1) = Chr$(96) Then
                sFileName = Mid$(sFileName, i + 1)
                Exit For
            End If
        Next x
    Next i
    
    iLen = Len(Left(sFileName, InStr( _
    sFileName, sExtention) - 1))
                                        
                                        
                                                                        
    If iLen < 20 And iLen > 1 Then
        isFilename = True
    Else
        isFilename = False
    End If

End Function


Public Function isFunky(sCheck As String) As Boolean
    Dim i As Integer, x As Integer
    For i = Len(sCheck) To 1 Step -1
        For x = 1 To 39
            If Mid(sCheck, i, 1) = Chr(x) Then
                isFunky = True
                Exit Function
            End If
        Next x
    Next i
    
    
    For i = Len(sCheck) To 1 Step -1
        For x = 123 To 255
            If Mid(sCheck, i, 1) = Chr(x) Then
                isFunky = True
                Exit Function
            End If
        Next x
    Next i
    

End Function

Public Function CharsIN(sText As String, sChar As String) As Long
    Dim iPos As Long, sNext As String
    sNext = sText
    Do
        iPos = InStr(sText, sChar)
        If iPos = 0 Then Exit Function
        sText = Mid(sText, iPos + 1)
        CharsIN = CharsIN + 1
    Loop
End Function

Public Function CharsPOS(sText As String, sChar As String, Optional ByVal iStart As Long = 1) As Long

    Dim iPos As Long, iCount As Long
    iCount = 1
    Do

        iPos = InStr(iPos + 1, sText, sChar)
        If iPos = 0 Then Exit Function
        If iCount = iStart Then
            CharsPOS = iPos
            Exit Do
        End If
        iCount = iCount + 1
    Loop
End Function

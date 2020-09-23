Attribute VB_Name = "mGetResource"
Private Type ACCEL_TABLE_ENTRY
   fFlags As Integer
   wASCII As Integer
   wID As Integer
   wPadding As Integer
End Type
Private Const FVIRTKEY = &H1
Private Const FNOINVERT = &H2
Private Const FSHIFT = &H4
Private Const FCONTROL = &H8
Private Const FALT = &H10

Private Type PictDesc
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_REALSIZE = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADTRANSPARENT = &H20

Public Const MAX_STRING = 260

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Const MIIM_STATE = &H1
Const MIIM_ID = &H2
Const MIIM_SUBMENU = &H4
Const MIIM_CHECKMARKS = &H8
Const MIIM_TYPE = &H10
Const MIIM_DATA = &H20
Const MFT_SEPARATOR = &H800
Const MFS_CHECKED = &H8
Private Declare Function LoadMenu& Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String)
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Dim sMenuText As String

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Private Type MESSAGE_RESOURCE_BLOCK
   LowId As Long
   HighId As Long
   OffsetToEntries As Long
End Type
Private Type MESSAGE_RESOURCE_ENTRY
  uLength As Integer
  iFlags As Integer
  sText As String
End Type
Private Type MESSAGE_RESOURCE_DATA
   NumberOfBlocks As Long
   mrb() As MESSAGE_RESOURCE_BLOCK
   mre() As MESSAGE_RESOURCE_ENTRY
End Type

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceByNum Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long

Public Function GetPicture(ByVal ResType As String, ByVal ResName As String) As StdPicture
   Dim hData As Long
   Dim arr() As Byte
   Select Case ResType
      Case "1", "3"
         arr = GetDataArray(ResType, ResName)
         hData = CreateIconFromResourceEx(arr(0), UBound(arr) + 1, CLng(ResType) - 1, &H30000, 0, 0, LR_LOADMAP3DCOLORS)
      Case "2"
         hData = LoadImage(hModule, ResName, 0, 0, 0, LR_LOADMAP3DCOLORS)
      Case "12"
         hData = LoadImage(hModule, ResName, 2, 0, 0, LR_LOADMAP3DCOLORS)
      Case "14"
         hData = LoadImage(hModule, ResName, 1, 0, 0, LR_LOADMAP3DCOLORS)
   End Select
   If hData = 0 Then Exit Function
   If ResType = "2" Then
      Set GetPicture = BitmapToPicture(hData)
   Else
      Set GetPicture = IconToPicture(hData)
   End If
End Function

Public Function GetPictureExt(ByVal ResType As String, ByVal ResName As String) As IPictureDisp
   Dim arr() As Byte
   Dim nFile As Integer
   Dim Pic As IPictureDisp
   Dim nWidth As Long, nHeight As Long
   Dim sType As String
   If Dir(TEMP_FILE_NAME) <> "" Then
      Call mciSendString("close video", 0&, 0, 0)
      Kill TEMP_FILE_NAME
   End If
   arr = GetDataArray(ResType, ResName)
   nFile = FreeFile
   Open TEMP_FILE_NAME For Binary As #nFile
      Put #nFile, , arr
   Close #nFile
   Set GetPictureExt = LoadPicture(TEMP_FILE_NAME)
   Kill TEMP_FILE_NAME
End Function

Private Function BitmapToPicture(ByVal hBmp As Long) As StdPicture
    Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .PicType = vbPicTypeBitmap
       .hImage = hBmp
    End With
    With IGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    Set BitmapToPicture = oNewPic
End Function

Private Function IconToPicture(ByVal hIcon As Long) As StdPicture
    If hIcon = 0 Then Exit Function
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As Guid
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .PicType = vbPicTypeIcon
       .hImage = hIcon
    End With
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    Set IconToPicture = oNewPic
End Function

Public Function GetString(ByVal ResName As String) As String
   Dim arr() As Byte
   Dim nPos As Long, wID As Long, uLength As Long
   Dim s As String, sText As String
   arr = GetDataArray("6", ResName)
   For wID = (CLng(Mid(ResName, 2)) - 1) * 16 To CLng(Mid(ResName, 2)) * 16 - 1
       Call CopyMemory(uLength, arr(nPos), 2)
       If uLength Then
          s = String(uLength, 0)
          CopyMemory ByVal StrPtr(s), arr(nPos + 2), uLength * 2
          s = CStr(wID) & ": " & s

          s = ReplaceStr(s, vbLf, vbNewLine)
          s = ReplaceStr(s, vbCr & vbNewLine, vbNewLine)
          sText = sText & TrimNULL(s) & vbNewLine
          nPos = nPos + uLength * 2 + 2
       Else
          nPos = nPos + 2
       End If
   Next wID
   GetString = sText
End Function

Public Function GetMenuText(ByVal ResName As String) As String
   Dim hMenu As Long
   sMenuText = ""
   hMenu = LoadMenu(hModule, ResName)
   GetMenuInfo hMenu, 0
   DestroyMenu hMenu
   GetMenuText = sMenuText
End Function

Private Sub GetMenuInfo(ByVal hMenu As Long, ByVal level As Long)
    Dim itemcount As Long
    Dim c As Long
    Dim mii As MENUITEMINFO
    Dim retval As Long
    itemcount = GetMenuItemCount(hMenu)
    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU Or MIIM_ID
        For c = 0 To itemcount - 1
            .dwTypeData = Space(256)
            .cch = 256
            retval = GetMenuItemInfo(hMenu, c, 1, mii)
            If mii.fType = MFT_SEPARATOR Then
               sMenuText = sMenuText & String(5 * level, ".") & "[MENU SEPARATOR]" & vbNewLine
            Else
               sMenuText = sMenuText & String(5 * level, ".") & Left(.dwTypeData, .cch)
               If (.fState And MFS_CHECKED) Then sMenuText = sMenuText & " (checked)"
               sMenuText = sMenuText & " (cmdID = " & .wID & ")" & vbNewLine
            End If
            If .hSubMenu <> 0 Then GetMenuInfo .hSubMenu, level + 1
        Next c
    End With
End Sub

Public Function GetAccelerators(ByVal ResName As String) As String
   Dim nItems As Long, i As Long
   Dim arr() As Byte
   Dim ate() As ACCEL_TABLE_ENTRY
   Dim sText As String
   arr = GetDataArray("9", ResName)
   nItems = (UBound(arr) + 1) \ 8
   ReDim ate(nItems - 1)
   Call CopyMemory(ate(0), arr(0), nItems * 8)
   For i = 0 To nItems - 1
       If (ate(i).fFlags And FSHIFT) Then sText = sText & "Shift+"
       If (ate(i).fFlags And FCONTROL) Then sText = sText & "Ctrl+"
       If (ate(i).fFlags And FALT) Then sText = sText & "Alt+"
       sText = sText & KeyName(ate(i).wASCII)
       sText = sText & " (cmdID = " & ate(i).wID & ")" & vbNewLine
   Next i
   GetAccelerators = sText
End Function

Private Function KeyName(ByVal key As Long) As String
  Select Case key
     Case vbKeyA To vbKeyZ, vbKey0 To vbKey9
          KeyName = Chr(key)
     Case vbKeyF1 To vbKeyF16
          KeyName = "F" & CStr(key - vbKeyF1 + 1)
     Case vbKeyCancel:   KeyName = "CANCEL"
     Case vbKeyBack:     KeyName = "BACKSPACE"
     Case vbKeyTab:      KeyName = "TAB"
     Case vbKeyClear:    KeyName = "CLEAR"
     Case vbKeyReturn:   KeyName = "ENTER"
     Case vbKeyShift:    KeyName = "SHIFT"
     Case vbKeyControl:  KeyName = "CTRL"
     Case vbKeyMenu:     KeyName = "MENU"
     Case vbKeyPause:    KeyName = "PAUSE"
     Case vbKeyCapital:  KeyName = "CAPS LOCK"
     Case vbKeyEscape:   KeyName = "ESC"
     Case vbKeySpace:    KeyName = "SPACEBAR"
     Case vbKeyPageUp:   KeyName = "PAGE UP"
     Case vbKeyPageDown: KeyName = "PAGE DOWN"
     Case vbKeyEnd:      KeyName = "END"
     Case vbKeyHome:     KeyName = "HOME"
     Case vbKeyLeft:     KeyName = "LEFT ARROW"
     Case vbKeyUp:       KeyName = "UP ARROW"
     Case vbKeyRight:    KeyName = "RIGHT ARROW"
     Case vbKeyDown:     KeyName = "DOWN ARROW"
     Case vbKeySelect:   KeyName = "SELECT"
     Case vbKeyPrint:    KeyName = "PRINT SCREEN"
     Case vbKeyExecute:  KeyName = "EXECUTE"
     Case vbKeySnapshot: KeyName = "SNAPSHOT"
     Case vbKeyInsert:   KeyName = "INS"
     Case vbKeyDelete:   KeyName = "DEL"
     Case vbKeyHelp:     KeyName = "HELP"
     Case vbKeyNumlock:  KeyName = "NUM LOCK"
     Case Else:          KeyName = "Virtual Key " & CStr(key)
  End Select
End Function

Public Function GetHexDump(ByVal ResType As String, ByVal ResName As String) As String
   Dim arr() As Byte
   Dim sText As String, sLine As String
   arr = GetDataArray(ResType, ResName)
   sText = Space$((UBound(arr) \ 16 + 1) * 79)
   If Len(sText) > 65534 Then
      GetHexDump = sText
      Exit Function
   End If
   On Error Resume Next
   For i = 0 To UBound(arr) - 1 Step 16
       sLine = ZeroPad(Hex(i), 8) & " | "
       For j = 0 To 15
           sLine = sLine & ZeroPad(Hex(arr(i + j)), 2) & " "
           If Err Then sLine = sLine & "   "
       Next j
       sLine = sLine & "| "
       For j = 0 To 15
           If arr(i + j) < 32 Then
              sLine = sLine & "."
           Else
              sLine = sLine & Chr(arr(i + j))
           End If
       Next j
       sLine = sLine & vbNewLine
       Mid(sText, (i \ 16) * 79 + 1, 79) = sLine
   Next i
   GetHexDump = sText
End Function

Public Function GetDataArray(ByVal ResType As String, ByVal ResName As String) As Variant
   Dim hRsrc As Long
   Dim hGlobal As Long
   Dim arrData() As Byte
   Dim lpData As Long
   Dim arrSize As Long
   If IsNumeric(ResType) Then hRsrc = FindResourceByNum(hModule, ResName, CLng(ResType))
   If hRsrc = 0 Then hRsrc = FindResource(hModule, ResName, ResType)
   If hRsrc = 0 Then Exit Function
   hGlobal = LoadResource(hModule, hRsrc)
   lpData = LockResource(hGlobal)
   arrSize = SizeofResource(hModule, hRsrc)
   If arrSize = 0 Then Exit Function
   ReDim arrData(arrSize - 1)
   Call CopyMemory(arrData(0), ByVal lpData, arrSize)
   Call FreeResource(hGlobal)
   GetDataArray = arrData
End Function

Public Function GetVersionInfo(ByVal ResName As String) As String
   Dim arrVerInfo() As Byte
   Dim arrInfoName As Variant
   Dim arrLang(3) As Byte
   Dim sLang As String
   Dim dwBytes As Long
   Dim lpBuffer As Long
   Dim s As String
   Dim sText As String
   s = String(MAX_STRING, 0)
   Call GetModuleFileName(hModule, s, MAX_STRING)
   s = TrimNULL(s)
   dwBytes = GetFileVersionInfoSize(s, lpBuffer)
   ReDim arrVerInfo(0 To dwBytes - 1)
   Call GetFileVersionInfo(s, 0, dwBytes, arrVerInfo(0))
   arrInfoName = Array("OriginalFilename", "InternalName", "FileVersion", "FileDescription", "ProductName", "ProductVersion", "CompanyName", "LegalCopyright")
   Call VerQueryValue(arrVerInfo(0), "\VarFileInfo\Translation", lpBuffer, dwBytes)
   Call CopyMemory(arrLang(0), ByVal lpBuffer, dwBytes)
   sLang = ZeroPad(Hex(arrLang(1)), 2) & ZeroPad(Hex(arrLang(0)), 2) & ZeroPad(Hex(arrLang(3)), 2) & ZeroPad(Hex(arrLang(2)), 2)
   For i = 0 To UBound(arrInfoName) - 1
       Call VerQueryValue(arrVerInfo(0), "\StringFileInfo\" & sLang & "\" & CStr(arrInfoName(i)), lpBuffer, dwBytes)
       s = StrFromPtrA(lpBuffer)
       If s <> "" Then sText = sText & arrInfoName(i) & ":" & vbCrLf & vbTab & s & vbNewLine
   Next i
   GetVersionInfo = Trim(sText)
End Function

Private Function ZeroPad(strValue As String, intLen As String) As String
    ZeroPad = Right$(String(intLen, "0") & strValue, intLen)
End Function

Public Function GetMessageTable(ByVal ResName As String) As String
   Dim arr() As Byte
   Dim nBlocks As Long
   Dim nPos As Long
   Dim i As Long, j As Long
   Dim s As String, sText As String
   Dim uLength As Long, uFlag As Long
   arr = GetDataArray("11", ResName)
   Call CopyMemory(nBlocks, arr(0), 4)
   If nBlocks = 0 Then Exit Function
   ReDim mrb(nBlocks - 1) As MESSAGE_RESOURCE_BLOCK
   Call CopyMemory(mrb(0), arr(4), 12 * nBlocks)
   For i = 0 To nBlocks - 1
       nPos = mrb(i).OffsetToEntries
       For j = mrb(i).LowId To mrb(i).HighId
           Call CopyMemory(uLength, arr(nPos), 2)
           If uLength Then
              Call CopyMemory(uFlag, arr(nPos + 2), 2)
              If uFlag = 1 Then
                 s = String(uLength, Chr$(0))
                 CopyMemory ByVal StrPtr(s), arr(nPos + 4), uLength - 4
                 s = CStr(j) & ": " & s
              Else
                 s = CStr(j) & ": " & StrFromPtrA(VarPtr(arr(nPos + 4)), uLength - 4)
              End If

              s = ReplaceStr(s, vbLf, vbNewLine)
              s = ReplaceStr(s, vbCr & vbNewLine, vbNewLine)
              sText = sText & TrimNULL(s)
              nPos = nPos + uLength
           End If
       Next j
   Next i
   GetMessageTable = sText
End Function

Public Function GetHTML(ByVal ResType As String, ByVal ResName As String) As String
  GetHTML = StrConv(GetDataArray(ResType, ResName), vbUnicode)
End Function

Public Function ResSize(ByVal ResType As String, ByVal ResName As String) As Long
   Dim hRsrc As Long, hGlobal As Long
   If IsNumeric(ResType) Then hRsrc = FindResourceByNum(hModule, ResName, CLng(ResType))
   If hRsrc = 0 Then hRsrc = FindResource(hModule, ResName, ResType)
   If hRsrc = 0 Then Exit Function
   hGlobal = LoadResource(hModule, hRsrc)
   ResSize = SizeofResource(hModule, hRsrc)
   Call FreeResource(hGlobal)
End Function



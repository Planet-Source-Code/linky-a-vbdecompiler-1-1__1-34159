Attribute VB_Name = "modMain"

Public Enum ResTypes
   RT_CURSOR = 1&
   RT_BITMAP = 2&
   RT_ICON = 3&
   RT_MENU = 4&
   RT_DIALOG = 5&
   RT_STRING = 6&
   RT_FONTDIR = 7&
   RT_FONT = 8&
   RT_ACCELERATOR = 9&
   RT_RCDATA = 10&
   RT_MESSAGETABLE = 11&
   RT_GROUP_CURSOR = 12&
   RT_GROUP_ICON = 14&
   RT_VERSION = 16&
   RT_DLGINCLUDE = 17&
   RT_PLUGPLAY = 19&
   RT_VXD = 20&
   RT_ANICURSOR = 21&
   RT_ANIICON = 22&
   RT_HTML = 23&
End Enum

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1
Public Const LOAD_LIBRARY_AS_DATAFILE = 2

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CopyStringA Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Public hModule As Long
Public picHeight As Long, picWidth As Long

Public Function InitResource(ByVal sLibName As String) As Boolean
  On Error Resume Next
  hModule = LoadLibraryEx(sLibName, 0, 1)
'  hModule = LoadLibrary(sLibName)
  InitResource = (hModule <> 0)
End Function

Public Sub ClearResource()
   If Dir(TEMP_FILE_NAME) <> "" Then
      Call mciSendString("close video", 0&, 0, 0)
      Kill TEMP_FILE_NAME
   End If
   If hDialog Then Call DestroyWindow(hDialog)
   If hModule Then FreeLibrary (hModule)
End Sub

Public Function ResTypeName(ByVal ResType As ResTypes) As String
   Select Case ResType
      Case RT_ACCELERATOR
         ResTypeName = "Accelerateur de donnée"
      Case RT_ANICURSOR
         ResTypeName = "Curseur animé"
      Case RT_ANIICON
         ResTypeName = "Icones animée"
      Case RT_BITMAP
         ResTypeName = "Images"
      Case RT_CURSOR
         ResTypeName = "HD Curseur"
      Case RT_DIALOG
         ResTypeName = "Dialogue"
      Case RT_DLGINCLUDE
         ResTypeName = ""
      Case RT_FONT
         ResTypeName = "Polices"
      Case RT_FONTDIR
         ResTypeName = "Dossier des polices"
      Case RT_GROUP_CURSOR
         ResTypeName = "Curseur resource"
      Case RT_GROUP_ICON
         ResTypeName = "Icones resource"
      Case RT_HTML
         ResTypeName = "Document HTML"
      Case RT_ICON
         ResTypeName = "H Icone resource"
      Case RT_MENU
         ResTypeName = "Menu"
      Case RT_MESSAGETABLE
         ResTypeName = "Strings"
      Case RT_PLUGPLAY
         ResTypeName = "Plug and play"
      Case RT_RCDATA
         ResTypeName = "RAW"
      Case RT_STRING
         ResTypeName = "String-table"
      Case RT_VERSION
         ResTypeName = "Version"
      Case RT_VXD
         ResTypeName = "VXD"
      Case Else
         ResTypeName = "Resource perso"
   End Select
End Function

Public Function StrFromPtrA(ByVal lpszA As Long, Optional nSize As Long = 0) As String
   Dim s As String, bTrim As Boolean
   If nSize = 0 Then
      nSize = lstrlenA(lpszA)
      bTrim = True
   End If
   s = String(nSize, Chr$(0))
   CopyStringA s, ByVal lpszA
   If bTrim Then s = TrimNULL(s)
   StrFromPtrA = s
End Function

Public Function StrFromPtrW(ByVal lpszW As Long, Optional nSize As Long = 0) As String
   Dim s As String, bTrim As Boolean
   If nSize = 0 Then
      nSize = lstrlenW(lpszW)
      bTrim = True
   End If
   s = String(nSize, Chr$(0))
   CopyMemory ByVal StrPtr(s), ByVal lpszW, nSize
   If bTrim Then s = TrimNULL(s)
   StrFromPtrW = s
End Function

Public Function TrimNULL(ByVal str As String) As String
    If InStr(str, Chr$(0)) > 0& Then
        TrimNULL = Left$(str, InStr(str, Chr$(0)) - 1&)
    Else
        TrimNULL = str
    End If
End Function

Public Function MakeLangID(ByVal usPrimaryLanguage As Integer, ByVal usSubLanguage As Long) As Long
    MakeLangID = usSubLanguage * 2 ^ 10 + usPrimaryLanguage
End Function

Public Function ReplaceStr(ByVal str As String, ByVal sReplace As String, Optional ByVal sReplaceWith As String, Optional fCompare As VbCompareMethod) As String
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    iLenOut = Len(sReplace)
    iLenIn = Len(sReplaceWith)
    If Len(str) > 0& Then
        If iLenOut > 0& Then
            Dim sOut As String
            i = InStr(1&, str, sReplace, fCompare)
            Do Until i = 0&
                If iLenIn > 0& Then
                    str = Left$(str, i - 1&) & sReplaceWith & Mid$(str, i + iLenOut)
                Else
                    str = Left$(str, i - 1&) & Mid$(str, i + iLenOut)
                End If
                i = InStr(i + iLenIn, str, sReplace, fCompare)
            Loop
        End If
    End If
    ReplaceStr = str
End Function



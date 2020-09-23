Attribute VB_Name = "modShowResource"

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As String, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetObjectAPI& Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any)

Public Const TEMP_FILE_NAME = "c:\tempfile.tmp"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public hDialog As Long

Public Function ShowAVI(ByVal ResName As String, pb As PictureBox) As Boolean
   Dim arr() As Byte
   Dim mciCmd As String
   Dim nFile As Integer
   Dim sReturn As String * 128
   Dim nWidth As Long, nHeight As Long
   Dim lStart As Long, lPos As Long
   If Dir(TEMP_FILE_NAME) <> "" Then
      Call mciSendString("close video", 0&, 0, 0)
      Kill TEMP_FILE_NAME
   End If
   arr = GetDataArray("AVI", ResName)
   nFile = FreeFile
   Open TEMP_FILE_NAME For Binary As #nFile
      Put #nFile, , arr
   Close #nFile
   mciCmd = "open " & TEMP_FILE_NAME & " Type avivideo Alias video parent " & pb.hwnd & " Style child"
   Call mciSendString(mciCmd, 0&, 0, 0)
   Call mciSendString("Where video destination", ByVal sReturn, Len(sReturn) - 1, 0)
   lStart = InStr(1, sReturn, " ")
   lPos = InStr(lStart + 1, sReturn, " ")
   lStart = InStr(lPos + 1, sReturn, " ")
   nWidth = Mid(sReturn, lPos, lStart - lPos) * Screen.TwipsPerPixelX
   nHeight = Mid(sReturn, lStart + 1) * Screen.TwipsPerPixelY
   pb.Move 0, 0
   If nWidth < picWidth Then pb.Width = picWidth Else pb.Width = nWidth
   If nHeight < picHeight Then pb.Height = picHeight Else pb.Height = nHeight
   Call mciSendString("put video window at " & (pb.Width - nWidth) \ (2 * Screen.TwipsPerPixelX) & " " & (pb.Height - nHeight) \ (2 * Screen.TwipsPerPixelY) & " " & nWidth \ Screen.TwipsPerPixelX & " " & nHeight \ Screen.TwipsPerPixelY, 0&, 0, 0)
   Call mciSendString("play video repeat", 0&, 0, 0)
   ShowAVI = True
End Function

Public Function ShowDialog(ByVal ResName As String, pb As PictureBox) As Boolean
   Dim rc As RECT, rcPic As RECT
   hDialog = CreateDialogParam(hModule, ResName, pb.hwnd, 0, 0)
   If hDialog Then
      If GetParent(hDialog) = pb.hwnd Then
         Call GetWindowRect(hDialog, rc)
         Call MoveWindow(hDialog, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, 1)
         pb.Move 0, 0, (rc.Right - rc.Left) * Screen.TwipsPerPixelX, (rc.Bottom - rc.Top + 24) * Screen.TwipsPerPixelY
      Else
         Call GetWindowRect(hDialog, rc)
         Call GetWindowRect(pb.hwnd, rcPic)
         Call MoveWindow(hDialog, rcPic.Left, rcPic.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 1)
      End If
      Call ShowWindow(hDialog, vbNormalFocus)
      ShowDialog = True
   End If
End Function

Public Function ShowPicture(Pic As StdPicture, pb As PictureBox) As Boolean
   If Pic Is Nothing Then Exit Function
   Dim nWidth As Long, nHeight As Long
   Dim sType As String
   nWidth = pb.ScaleX(Pic.Width, vbHimetric, vbTwips)
   nHeight = pb.ScaleY(Pic.Height, vbHimetric, vbTwips)
   If nWidth < picWidth Then nWidth = picWidth
   If nHeight < picHeight Then nHeight = picHeight
   pb.Move 0, 0, nWidth, nHeight
   pb.PaintPicture Pic, pb.Width \ 2 - pb.ScaleX(Pic.Width, vbHimetric, vbTwips) \ 2, pb.Height \ 2 - pb.ScaleY(Pic.Height, vbHimetric, vbTwips) \ 2
   pb.CurrentX = 0
   pb.CurrentY = 0
   pb.Print "Image info:" & vbNewLine
   Select Case Pic.Type
       Case 0: sType = "None"
       Case 1: sType = "Bitmap (*.bmp)"
            DeleteObject Pic.handle
       Case 2: sType = "Metafile (*.wmf)"
       Case 3: sType = "Icon/cursor (*.ico/*.cur)"
            DestroyIcon Pic.handle
       Case 4: sType = "Enh Metafile (*.emf)"
   End Select
   pb.Print "Type: " & sType
   pb.Print "Syze: " & CInt(pb.ScaleX(Pic.Width, vbHimetric, vbPixels)) & " x " & CInt(pb.ScaleY(Pic.Height, vbHimetric, vbPixels))
   ShowPicture = True
End Function

Public Function ShowText(ByVal sText As String, txt As TextBox) As Boolean
   If sText <> "" Then
      If Len(sText) > 65534 Then
         txt.Text = "Text too long to display it in text box." & vbNewLine & "Save it as file and view with notepad"
         Exit Function
      End If
      txt.Text = sText
      ShowText = True
   End If
End Function

Public Sub SaveData(ByVal sFileName As String, arrData As Variant)
   Dim nFile As Integer
   Dim arr() As Byte
   arr = arrData
   nFile = FreeFile
   Open sFileName For Binary As #nFile
      Put #nFile, , arr
   Close #nFile
End Sub

Public Sub SaveText(ByVal sFileName As String, sText As String)
   Dim nFile As Integer
   nFile = FreeFile
   Open sFileName For Binary As #nFile
      Put #nFile, , sText
   Close #nFile
End Sub




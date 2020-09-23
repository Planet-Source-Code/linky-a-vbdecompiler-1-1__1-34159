Attribute VB_Name = "modPosition"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Public Function PForm(ByVal frmA As Form, Optional bDevant As Boolean = True) As Long
If bDevant Then
PForm = SetWindowPos(frmA.hwnd, HWND_TOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
Else
PForm = SetWindowPos(frmA.hwnd, HWND_NOTOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
End If
End Function




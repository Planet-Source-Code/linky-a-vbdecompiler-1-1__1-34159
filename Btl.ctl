VERSION 5.00
Begin VB.UserControl BTL 
   AutoRedraw      =   -1  'True
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   DrawStyle       =   2  'Dot
   KeyPreview      =   -1  'True
   PropertyPages   =   "Btl.ctx":0000
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   ToolboxBitmap   =   "Btl.ctx":002C
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer TimerHover 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1215
      Top             =   0
   End
End
Attribute VB_Name = "BTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' FYI: The code used to draw the buttons was borrowed from a project posted on PSC.
'   Although programmer claims to be the original author, most of the code was found among 3 other sites,
'   MSDN, A1VBcode, VBcode, posted by different authors... in some cases the code was an exact match,
'   in others, slight changes like variable names.

' Regardless, let it be known, that I did not write that part of the code.
' My contribution to this project...
'   <> inclusion of rotated buttons
'   <> inclusion of segmented buttons (horizontal only). Maybe when I get time, I'll include vertical 90,270 degree buttons also
'   <> inclusion of text styles: Emboss, Engrave & Shadow
'   <> 2 angles for rotated buttons (+\- 90 degrees), default style only, not segmented style
'   <> use of button icons, with left/right justification & disabled color
'   <> Replaced Ampersand/Hotkey function with more robust code to find hotkey regardless of number of ampersands in caption
'   <> Added property comments to appear at bottom of property sheet
'   <> Added built-in override of the Timer Control.  See TimerHover_Timer() for details.
'   <> corrected following logical errors to original code
'       - Allowed button drawing to occur only once during control loading
'           (before it would draw several times before it was actually displayed)
'       - Key strokes fired mouse clicks (key strokes should fire key strokes, mouse fires mouse)
'       - Redrawing/repainting before mouse clicks & keystrokes sent to parent
'       - If user right-clicked on button, button could not be activated by keyboard unless focus was lost & regained
'       - Java Button focus rectangle was fine for 1 line captions, but looked bad on 2-line buttons
'       - personal pref: removed MouseOver event (really same as MouseMove)
'       - when selecting custom color property while ColorScheme <> Custom, program now automatically selects Custom for you
'       - when returning to non-Custom colorscheme, program automatically resets all colors to that default
'       - when selecting a system vs palette color for any color property, the program
'          would make that color black, cause system colors are < 0 and no
'          actual color is < 0. Added function to convert system color to real color

' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Changes in this revision of the program cannot be transferred directly... It is a complete rewrite
'   on how the control performs.

' This version reverts back to painting on demand or simply, when the status of a button changes. The previous version painted all buttons once,
'   but used more resources and had severe problems displaying in NT & Win2K.
' Frustrated by not being able to get it to work in NT forced me to try something else...
' I now use logical fonts created when needed. These logical fonts can be rotated assuming they are not System fonts. Lots of limitations,
'   but taken care of in extra programming and documented in those modules (WordWrapCaption, CreateDisplayFont, DetermineOS, DrawDisabledIcon)
' All but ONE capability has been carried over to this version... rotated icons when buttons are rotated. It was hard enough to get the disabled icon effect

' Note: When button attributes change via code, the button is forced to redraw, same as original project, although many options are read-only at runtime
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////

' The following was cut & pasted from original project
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%              <<< GONCHUKI SYSTEMS >>>              %
'%                                                    %
'%   CHAMELEON BUTTON - copyright ©2001 by gonchuki   %
'%                                                    %
'%  this custom control will emulate the most common  %
'%      command buttons that everyone knows.          %
'%                                                    %
'% it took me about two months to develop this control%
'%  and at this time i think it's completely bug free %
'%     ALL THE CODE WAS WRITTEN FROM SCRATCH!!!       %
'%                                                    %
'%   ever wanted to add cool buttons to your app???   %
'%          this is the BEST solution!!!              %
'%                                                    %
'%                                                    %
'%     e-mail: gonchuki@yahoo.es                      %
'%                                                    %
'%              Don't forget to vote!!!               %
'%                                                    %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' Now the API list
Private Const Version As String = "2.0"

' Drawing APIs
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hbr As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified bye TRIVERTEX structure) fore left and right edges. GDI interpolates
'the color frome top to bottom edge and fillse interior.
Private Const GRADIENT_FILL_RECT_V  As Long = &H1 'Inis mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified bye TRIVERTEX structure) fore top and bottom edges. GDI interpolates
'the color frome top to bottom edge and fillse interior.
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along wia list of array indexesat describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fillse interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Private Const GRADIENT_FILL_OP_FLAG As Long = &HFF

' Color APIs
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

' Word Width & Height API
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800

' Windows object selection/deletion
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' Window Shapes API
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

' Windows rectangle functions
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Miscellaneous APIs
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long

' following should be used if the Timer function is not used (See TimerHover_Timer() for details)
'Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long

' Font APIs
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

Public Event Click()            ' Button/Mouse events that will show up in attached programs
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOut()
Private WithEvents TextFont As StdFont 'current font
Attribute TextFont.VB_VarHelpID = -1

Private Type TRIVERTEX  ' used for Gradient fills
    x As Long
    y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort value
End Type
Private Type GRADIENT_RECT  ' used for graident fills
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type
Private Type OSVERSIONINFO          ' used to help identify operating system
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type TEXTMETRIC             ' used to get information for specific font
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type
Private Type LOGFONT                 ' used to define a font
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 33
End Type
Private Type RECT       ' Rectangle used in sizing windows and button captions
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI   ' Mouse X,Y coordinates used in MouseMove
        x As Long
        y As Long
End Type
Private Enum ButtonStatus   ' Status of active button
    [Normal Status] = 0
    [Disabled Status] = 1
    [Button Down] = 2
    [Mouse Over] = 4
    [Got Focus] = 3
End Enum
Private Type CaptionData    ' used to store information about the button caption
    cmdOffset As RECT
    cmdText As String
End Type
Private Type HotkeyData     ' used to store information about the hotkey (&)
    cmdHotKey As Integer
    cmdHotKeyXY As POINTAPI
    cmdHotKeyLen As Integer
End Type
Public Enum IconSizeDat     ' used to set icon size on button
    [8 x 8] = 1
    [16 x 16] = 2
    [24 x 24] = 3
    [32 x 32] = 4
End Enum
Public Enum GradientStyleDat
    [Left to Right] = GRADIENT_FILL_RECT_H
    [Top to Bottom] = GRADIENT_FILL_RECT_V
    [Right to Left] = 2
    [Bottom to Top] = 3
End Enum
Public Enum TextStyleDat    ' used for text styles
    [Plain Text] = 0
    [Embossed] = 1
    [Engraved] = 2
    [Shadowed] = 3
End Enum
Public Enum OrientationTypesDat ' Button orientation
    [Horizontal] = 0
    [Vertical 90] = 1
    [Vertical 270] = 2
End Enum
Public Enum ButtonTypes         ' Various button patterns
    [Windows 16-bit] = 1    'the old-fashioned Win16 button
    [Windows 32-bit] = 2    'the classic windows button
    [Windows XP] = 3        'the new brand XP button totally owner-drawn
    [Java metal] = 4        'there are also other styles but not so different from windows one
    [Netscape 6] = 5        'this is the button displayed in web-pages, it also appears in some java apps
    [Simple Flat] = 6       'the standard flat button seen on toolbars
    [Flat Highlight] = 7    'again the flat button but this one has no border until the mouse is over it
    [Mac] = 8               'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
End Enum
Public Enum GraphicModeDat  ' used to help ensure system can display vertical fonts
    [Default Mode] = -1
    [Non-NT and Win 2K] = 0
    [Windows NT] = 1
    [Other Mode] = 2
End Enum
Public Enum ColorTypes      ' Color schemes, only Custom allow color attributes
    [Use Windows] = 1       ' to change, except button text which can be changed
    [Custom] = 2            ' in all color schemes
    [Force Standard] = 3
    [Use Container] = 4
    [Custom Gradient] = 5
End Enum
Public Enum CaptionAlignment    ' Horizontal caption alignment
    [Left Justified] = 0
    [Center Justified] = 1
    [Right Justified] = 2
End Enum
Public Enum IconAlignment       ' Icon alignment to ends of button
    [Left Aligned] = 0
    [Right Aligned] = 1
End Enum
Public Enum ButtonStyleDat      ' Segmented button styles
    [Default Style] = 0
    [Left Segmented] = 1
    [Inner Segmented] = 2
    [Right Segmented] = 3
End Enum

'variables
Private Const SegIndent = 18                    ' depth of indent for segmented buttons
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes
Private BackC As Long 'back color
Private ForeC As Long 'fore color
Private ForeO As Long 'fore color when mouse is over
Private btnCaption As String     'current text
Private rgnNorm As Long
Private curStat As ButtonStatus
Private LastButton As Byte, LastKeyDown As Byte ' last mouse/keystroke
Private isEnabled As Boolean                    ' button enabled status
Private hasFocus As Boolean                     ' button focus status
Private bShown As Boolean
Private hMyIcon As StdPicture
' Color variables
Private cFace As Long, cLight As Long, cHighLight As Long
Private cShadow As Long, cDarkShadow As Long, cText As Long, cTextO As Long
Private cEmbossM As Long, cEmbossS As Long
Private isOver As Boolean                       ' mouse over status
Private bIcon2 As Boolean                       ' icon used or not
Private myIconSize As IconSizeDat
Private myCaptionAlign As CaptionAlignment      ' left, center, right justify
Private iconAlign As IconAlignment              ' icon left or right aligned
Private myOrientation As OrientationTypesDat    ' 180, 90 or 270 degree text
Private myTextStyle As TextStyleDat             ' shadow, emboss, engrave
Private hMyFont As Long, hPrevFont As Long      ' font handles
Private bWordWrap As Boolean
' following contains the x,y coords for text overlay on button,
'   length, height of text, and the text caption
Private CaptionInfo() As CaptionData
Private buttonBorder As RECT                    ' box coords around text
Private btnHotKey As HotkeyData                 ' hotkey information
Private iconXY As POINTAPI                      ' icon x,y coords on button
Private bShowFocus As Boolean
Private GStart As Long, GStop As Long, GradientStyle As GradientStyleDat ' Gradient Variables
Private myOptMode As Boolean, myOptValue As Boolean
Private btnStyle As ButtonStyleDat
' With rotated text and other graphics, they may not draw correctly on
' an NT system. The GraphicsMode API when set with parameter of 2 should
' correct the problem. But with ME, 2K and XP, not sure which, if any,
' setting may be required. So the user can opt to set the parameter between
'   0 to 2 to overcome system failing to print rotated text
Private Gmode As GraphicModeDat, GraphicsModeUsed As GraphicModeDat

Private Sub DrawButton(Optional bkgDisabled As Boolean = False)
' Function calls the drawing routine for the currently selected button type
Dim iLastStatus As Integer, bWasEnabled As Boolean

If bkgDisabled = True Then          ' Used when drawing a disabled icon
    iLastStatus = curStat                ' so save the current values and reset them at end of routine
    curStat = [Disabled Status]      ' we force the status to be Disabled so the background color could be drawn
    bWasEnabled = isEnabled       ' which is used for the background of the disabled button
    isEnabled = False
End If
' paint a rectangle where the button will be ensuring it is 'clean'
On Error Resume Next
If MyColorType = [Custom Gradient] Then
    DoGradientFill
Else
    If MyButtonType = [Flat Highlight] And (MyColorType = Custom Or MyColorType = [Custom Gradient]) And curStat <> [Mouse Over] Then
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, GetBkColor(UserControl.Parent.hdc)
    Else
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cFace
    End If
End If
If bWordWrap = True Then CreateDisplayFont bWordWrap
Select Case MyButtonType        ' Call routines to paint the buttons
    Case [Windows 16-bit]:   ' Windows 16-bit
        Win16button bkgDisabled
    Case [Windows 32-bit] 'Windows 32-bit
        Win32button bkgDisabled
    Case [Windows XP]  'Windows XP
        WinXPbutton bkgDisabled
    Case [Java metal]  'Java
        JavaButton bkgDisabled
    Case [Netscape 6]  'Netscape
        NetScapeButton bkgDisabled
    Case Mac 'Mac
        MacButton bkgDisabled
    Case Else 'Flat buttons
        FlatButton bkgDisabled
End Select
UserControl.Refresh
If bkgDisabled = True Then          ' if drawing a disabled button, return the original current values
    curStat = iLastStatus
    isEnabled = bWasEnabled
End If

' when in design mode show the Normal Status - disabled status always overrides all other status
If Ambient.UserMode = False Then curStat = [Normal Status]
End Sub

Private Sub Win16button(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
Dim bOptionBtn As Boolean
If myOptMode = True And myOptValue = True Then curStat = curStat * -1 - 1
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        DrawFrame cHighLight, cShadow, cHighLight, cShadow, 1
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If hasFocus Then DrawFocusR
    Else ' Disabled
        DrawFrame cHighLight, cShadow, cHighLight, cShadow, 1
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
        If bkgDisabled = False Then PrintText cShadow, cHighLight, -1
    End If
Case 2: 'Button Down
    DrawFrame cShadow, cHighLight, cShadow, cHighLight, 1
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
    DrawFocusR
    PrintText cText, -1, -1
Case Is < 0: ' Option Button
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cLight
    DrawFrame cShadow, cHighLight, cShadow, cHighLight, 0
    If isEnabled = True Then
        If hasFocus = True Or curStat = -3 Then DrawFocusR
        If curStat = [Mouse Over] * -1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else
        If bkgDisabled = False Then PrintText cShadow, cHighLight, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub Win32button(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
Dim bOptionBtn As Boolean
If myOptMode = True And myOptValue = True Then curStat = (curStat * -1 - 1)
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        DrawFrame cHighLight, cDarkShadow, cLight, cShadow, 1
        If hasFocus = True Then
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
            DrawFocusR
        End If
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else ' Disabled
        DrawFrame cHighLight, cDarkShadow, cLight, cShadow, 1
        If bkgDisabled = True Then Exit Sub
        If myOptMode = False Then
            PrintText cShadow, cHighLight, -1
        Else
            PrintText cShadow, -1, -1
        End If
    End If
Case 2: ' Button Down
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
    DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, cShadow, True
    DrawFocusR
    PrintText cText, -1, -1
Case Is < 0: ' Option Button
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cLight
    DrawFrame cDarkShadow, cHighLight, cShadow, cLight, 0
    If isEnabled = True Then
        If hasFocus = True Or curStat = -3 Then DrawFocusR
        If curStat = [Mouse Over] * -1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else
        If bkgDisabled = False Then PrintText cShadow, -1, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub WinXPbutton(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
Dim I As Long, stepXP1 As Single, XPface As Long
If myOptMode = True And myOptValue = True Then curStat = curStat * -1 - 1
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        stepXP1 = 25 / ScaleHeight
        XPface = ShiftColor(cFace, &H30, True)
        For I = 1 To ScaleHeight
            DrawLine 0, I, ScaleWidth, I, ShiftColor(XPface, -stepXP1 * I, True)
        Next
        If btnStyle = 0 Then
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
            mSetPixel 1, 1, &H7B4D10
            mSetPixel 1, ScaleHeight - 2, &H7B4D10
            mSetPixel ScaleWidth - 2, 1, &H7B4D10
            mSetPixel ScaleWidth - 2, ScaleHeight - 2, &H7B4D10
        End If
        If curStat = [Mouse Over] Then
            If btnStyle > 0 Then
                DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
                DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, &H6BCBFF, True
                DrawRectangle 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H31B2FF, True
            Else
                DrawRectangle 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H31B2FF, True
                DrawLine 1, 2, 1, ScaleHeight - 2, &H6BCBFF
                DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, &H96E7&
                DrawLine 2, 1, ScaleWidth - 2, 1, &HCEF3FF
                DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H6BCBFF
            End If
        ElseIf hasFocus Then
            If btnStyle = [Default Style] Then
                DrawRectangle 1, 2, ScaleWidth - 2, ScaleHeight - 2, &HE7AE8C, True
                DrawLine 3, 2, 3, ScaleHeight - 2, &HF0D1B5
                DrawLine 3, ScaleHeight - 2, ScaleWidth - 3, ScaleHeight - 2, &HEF826B
                DrawLine 4, 3, ScaleWidth - 4, 3, &HFFE7CE
                DrawLine ScaleWidth - 4, 3, ScaleWidth - 4, ScaleHeight - 3, &HF0D1B5
            Else
                DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
            End If
            DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, &HE7AE8C, True
        Else 'we do not draw the bevel always because the above code would repaint over it
            DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, -&H30, True)
            DrawLine 1, ScaleHeight - 3, ScaleWidth - 2, ScaleHeight - 3, ShiftColor(XPface, -&H20, True)
            DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, -&H24, True)
            DrawLine ScaleWidth - 3, 3, ScaleWidth - 3, ScaleHeight - 3, ShiftColor(XPface, -&H18, True)
            DrawLine 2, 1, ScaleWidth - 2, 1, ShiftColor(XPface, &H10, True)
            DrawLine 1, 2, ScaleWidth - 2, 2, ShiftColor(XPface, &HA, True)
            DrawLine 1, 2, 1, ScaleHeight - 2, ShiftColor(XPface, -&H5, True)
            DrawLine 2, 3, 2, ScaleHeight - 3, ShiftColor(XPface, -&HA, True)
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
        End If
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else 'Disabled or option button
        XPface = ShiftColor(cFace, &H30, True)
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, ShiftColor(XPface, -&H18, True)
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, ShiftColor(XPface, -&H54, True), True
        If btnStyle = 0 Then
            mSetPixel 1, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel 1, ScaleHeight - 2, ShiftColor(XPface, -&H48, True)
            mSetPixel ScaleWidth - 2, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, -&H48, True)
        End If
        If bkgDisabled = False Then PrintText ShiftColor(XPface, -&H68, True), -1, -1
    End If
Case 2: 'Button Down
    stepXP1 = 25 / ScaleHeight
    XPface = ShiftColor(cFace, &H30, True)
    XPface = ShiftColor(XPface, -32, True)
    For I = 1 To ScaleHeight
        DrawLine 0, ScaleHeight - I, ScaleWidth, ScaleHeight - I, ShiftColor(XPface, -stepXP1 * I, True)
    Next
    If btnStyle = 0 Then
        mSetPixel 1, 1, &H7B4D10
        mSetPixel 1, ScaleHeight - 2, &H7B4D10
        mSetPixel ScaleWidth - 2, 1, &H7B4D10
        mSetPixel ScaleWidth - 2, ScaleHeight - 2, &H7B4D10
    End If
    DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, &H10, True)
    DrawLine 1, ScaleHeight - 3, ScaleWidth - 2, ScaleHeight - 3, ShiftColor(XPface, &HA, True)
    DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, &H5, True)
    DrawLine ScaleWidth - 3, 3, ScaleWidth - 3, ScaleHeight - 3, XPface
    DrawLine 2, 1, ScaleWidth - 2, 1, ShiftColor(XPface, -&H20, True)
    DrawLine 1, 2, ScaleWidth - 2, 2, ShiftColor(XPface, -&H18, True)
    DrawLine 1, 2, 1, ScaleHeight - 2, ShiftColor(XPface, -&H20, True)
    DrawLine 2, 2, 2, ScaleHeight - 2, ShiftColor(XPface, -&H16, True)
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
    If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    If hasFocus = True Then DrawRectangle 1, 2, ScaleWidth - 2, ScaleHeight - 2, &HE7AE8C, True
Case Is < 0: ' Option Button
    stepXP1 = 25 / ScaleHeight
    XPface = ShiftColor(cFace, -&H16, True)
    XPface = ShiftColor(XPface, 32, True)
    For I = ScaleHeight To 1 Step -1
        DrawLine 0, ScaleHeight - I, ScaleWidth, ScaleHeight - I, ShiftColor(XPface, -stepXP1 * I, True)
    Next
    If btnStyle = 0 Then
        mSetPixel 1, 1, &H7B4D10
        mSetPixel 1, ScaleHeight - 2, &H7B4D10
        mSetPixel ScaleWidth - 2, 1, &H7B4D10
        mSetPixel ScaleWidth - 2, ScaleHeight - 2, &H7B4D10
    End If
    DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, &H10, True)
    DrawLine 1, ScaleHeight - 3, ScaleWidth - 2, ScaleHeight - 3, ShiftColor(XPface, &HA, True)
    DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, &H5, True)
    DrawLine ScaleWidth - 3, 3, ScaleWidth - 3, ScaleHeight - 3, XPface
    DrawLine 2, 1, ScaleWidth - 2, 1, ShiftColor(XPface, -&H20, True)
    DrawLine 1, 2, ScaleWidth - 2, 2, ShiftColor(XPface, -&H18, True)
    DrawLine 1, 2, 1, ScaleHeight - 2, ShiftColor(XPface, -&H20, True)
    DrawLine 2, 2, 2, ScaleHeight - 2, ShiftColor(XPface, -&H16, True)
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
    If isEnabled = True Then
        If curStat = [Mouse Over] * -1 - 1 Then
            If btnStyle > 0 Then
                DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
                DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, &H6BCBFF, True
                DrawRectangle 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H31B2FF, True
            Else
                DrawRectangle 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H31B2FF, True
                DrawLine 1, 2, 1, ScaleHeight - 2, &H6BCBFF
                DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, &H96E7&
                DrawLine 2, 1, ScaleWidth - 2, 1, &HCEF3FF
                DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, &H6BCBFF
            End If
        ElseIf hasFocus = True Then
            If btnStyle = [Default Style] Then
                DrawRectangle 1, 2, ScaleWidth - 2, ScaleHeight - 2, &HE7AE8C, True
                DrawLine 3, 2, 3, ScaleHeight - 2, &HF0D1B5
                DrawLine 3, ScaleHeight - 2, ScaleWidth - 3, ScaleHeight - 2, &HEF826B
                DrawLine 4, 3, ScaleWidth - 4, 3, &HFFE7CE
                DrawLine ScaleWidth - 4, 3, ScaleWidth - 4, ScaleHeight - 3, &HF0D1B5
            Else
                DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
            End If
            DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, &HE7AE8C, True
        Else
            DrawLine 2, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, -&H30, True)
            DrawLine 1, ScaleHeight - 3, ScaleWidth - 2, ScaleHeight - 3, ShiftColor(XPface, -&H20, True)
            DrawLine ScaleWidth - 2, 2, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(XPface, -&H24, True)
            DrawLine ScaleWidth - 3, 3, ScaleWidth - 3, ScaleHeight - 3, ShiftColor(XPface, -&H18, True)
            DrawLine 2, 1, ScaleWidth - 2, 1, ShiftColor(XPface, &H10, True)
            DrawLine 1, 2, ScaleWidth - 2, 2, ShiftColor(XPface, &HA, True)
            DrawLine 1, 2, 1, ScaleHeight - 2, ShiftColor(XPface, -&H5, True)
            DrawLine 2, 3, 2, ScaleHeight - 3, ShiftColor(XPface, -&HA, True)
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, &H733C00, True
        End If
        If curStat = [Mouse Over] * -1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else
        If bkgDisabled = False Then PrintText cHighLight, -1, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub MacButton(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
If myOptMode = True And myOptValue = True Then curStat = [Button Down]
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, cLight
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
        If btnStyle = 0 Then
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, ScaleHeight - 2, cDarkShadow
            mSetPixel ScaleWidth - 2, 1, cDarkShadow
            mSetPixel ScaleWidth - 2, ScaleHeight - 2, cDarkShadow
            mSetPixel 1, 2, cFace
            mSetPixel 2, 1, cFace
            mSetPixel 3, 3, cHighLight
            mSetPixel ScaleWidth - 4, ScaleHeight - 4, cFace
            mSetPixel ScaleWidth - 3, ScaleHeight - 3, cShadow
            mSetPixel 2, ScaleHeight - 2, cFace
            mSetPixel 2, ScaleHeight - 3, cLight
            mSetPixel ScaleWidth - 2, 2, cFace
            mSetPixel ScaleWidth - 3, 2, cLight
            DrawLine 3, 2, ScaleWidth - 3, 2, cHighLight
            DrawLine 2, 2, 2, ScaleHeight - 3, cHighLight
            DrawLine ScaleWidth - 3, 1, ScaleWidth - 3, ScaleHeight - 3, cFace
            DrawLine 1, ScaleHeight - 3, ScaleWidth - 3, ScaleHeight - 3, cFace
            DrawLine ScaleWidth - 2, 3, ScaleWidth - 2, ScaleHeight - 2, cShadow
            DrawLine 3, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, cShadow
        End If
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If hasFocus Then DrawFocusR
    Else 'Disabled
        DrawRectangle 1, 1, ScaleWidth - 2, ScaleHeight - 2, cLight
        DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
        If btnStyle = 0 Then
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, ScaleHeight - 2, cDarkShadow
            mSetPixel ScaleWidth - 2, 1, cDarkShadow
            mSetPixel ScaleWidth - 2, ScaleHeight - 2, cDarkShadow
            mSetPixel 1, 2, cFace
            mSetPixel 2, 1, cFace
            mSetPixel 3, 3, cHighLight
            mSetPixel ScaleWidth - 4, ScaleHeight - 4, cFace
            mSetPixel ScaleWidth - 3, ScaleHeight - 3, cShadow
            mSetPixel 2, ScaleHeight - 2, cFace
            mSetPixel 2, ScaleHeight - 3, cLight
            mSetPixel ScaleWidth - 2, 2, cFace
            mSetPixel ScaleWidth - 3, 2, cLight
        End If
        DrawLine 3, 2, ScaleWidth - 3, 2, cHighLight
        DrawLine 2, 2, 2, ScaleHeight - 3, cHighLight
        DrawLine ScaleWidth - 3, 1, ScaleWidth - 3, ScaleHeight - 3, cFace
        DrawLine 1, ScaleHeight - 3, ScaleWidth - 3, ScaleHeight - 3, cFace
        DrawLine ScaleWidth - 2, 3, ScaleWidth - 2, ScaleHeight - 2, cShadow
        DrawLine 3, ScaleHeight - 2, ScaleWidth - 2, ScaleHeight - 2, cShadow
        If btnStyle Then
            DrawRectangle 1, 1, ScaleWidth - 2, ScaleHeight - 2, cLight
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
        End If
        If bkgDisabled = False Then PrintText cShadow, cHighLight, -1
    End If
Case 2: 'Button Down
    If btnStyle = 0 Then
        mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
        mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
        mSetPixel 1, 1, cDarkShadow
        mSetPixel 1, ScaleHeight - 2, cDarkShadow
        mSetPixel ScaleWidth - 2, 1, cDarkShadow
        mSetPixel ScaleWidth - 2, ScaleHeight - 2, cDarkShadow
        mSetPixel ScaleWidth - 4, ScaleHeight - 4, cShadow
        mSetPixel ScaleWidth - 2, ScaleHeight - 3, ShiftColor(cShadow, -&H20)
        mSetPixel ScaleWidth - 3, ScaleHeight - 2, ShiftColor(cShadow, -&H20)
        mSetPixel 2, ScaleHeight - 2, ShiftColor(cShadow, -&H20)
        mSetPixel 2, ScaleHeight - 3, ShiftColor(cShadow, -&H10)
        mSetPixel 1, ScaleHeight - 3, ShiftColor(cShadow, -&H10)
        mSetPixel ScaleWidth - 2, 2, ShiftColor(cShadow, -&H20)
        mSetPixel ScaleWidth - 3, 2, ShiftColor(cShadow, -&H10)
        mSetPixel ScaleWidth - 3, 1, ShiftColor(cShadow, -&H10)
        DrawRectangle 1, 1, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(cShadow, -&H40), True
        DrawRectangle 2, 2, ScaleWidth - 4, ScaleHeight - 4, ShiftColor(cShadow, -&H20), True
    End If
    DrawLine ScaleWidth - 3, 1, ScaleWidth - 3, ScaleHeight - 3, cShadow
    DrawLine 1, ScaleHeight - 3, ScaleWidth - 2, ScaleHeight - 3, cShadow
    DrawLine ScaleWidth - 2, 3, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(cShadow, -&H10)
    DrawLine 2, ScaleHeight - 1, ScaleWidth - 2, ScaleHeight - 1, ShiftColor(cShadow, -&H10)
    DrawFocusR
    DrawRectangle 1, 1, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(cShadow, -&H10)
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cDarkShadow, True
    If isEnabled = False Then
        PrintText cLight, -1, -1
    Else
        If myOptMode = True And myOptValue = True Then
            If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cLight, -1, -1
            If hasFocus = True Then DrawFocusR
        Else
            PrintText cShadow, cHighLight, -1
        End If
    End If
End Select
End Sub

Private Sub FlatButton(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
If myOptMode = True And myOptValue = True Then curStat = curStat * -1 - 1
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If Not (MyButtonType = [Flat Highlight]) Then
            DrawFrame cHighLight, cShadow, 0, 0, 0, True
        ElseIf isOver Or curStat = [Mouse Over] Then
            DrawFrame cHighLight, cShadow, 0, 0, 0, True
        End If
        If hasFocus = True Then DrawFocusR
    Else 'Disabled
        If bkgDisabled = False Then PrintText cShadow, cHighLight, -1
        If MyButtonType = [Simple Flat] Then
            DrawFrame cHighLight, cShadow, 0, 0, 0, True
        Else
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cShadow, True
        End If
    End If
Case 2: 'Button Down
    PrintText cText, -1, -1
    DrawFocusR
    DrawFrame cShadow, cHighLight, 0, 0, 0, True
Case Is < 0: 'Option Button
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cLight
    DrawFrame cShadow, cHighLight, 0, 0, 0, True
    If isEnabled = True Then
        If curStat = [Mouse Over] * 1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If hasFocus = True Then DrawFocusR
    Else
        If bkgDisabled = False Then PrintText cShadow, cHighLight, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub JavaButton(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
If myOptMode = True And myOptValue = True Then curStat = curStat * -1 - 1
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cFace, &HC)
        DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, cHighLight, True
        DrawRectangle 0, 0, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cShadow, -&H1A), True
        If btnStyle Then
            mSetPixel 1, ScaleHeight - 2, ShiftColor(cShadow, &H1A)
            mSetPixel ScaleWidth - 2, 1, ShiftColor(cShadow, &H1A)
        End If
        If hasFocus Then DrawFocusR
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
    Else 'Disabled
        If myOptMode = False Then
            DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cShadow, True
        Else
            DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cFace, &HC)
            DrawRectangle 1, 1, ScaleWidth - 1, ScaleHeight - 1, cHighLight, True
            DrawRectangle 0, 0, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cShadow, -&H1A), True
            If btnStyle Then
                mSetPixel 1, ScaleHeight - 2, ShiftColor(cShadow, &H1A)
                mSetPixel ScaleWidth - 2, 1, ShiftColor(cShadow, &H1A)
            End If
        End If
        If bkgDisabled = False Then PrintText cShadow, -1, -1
    End If
Case 2: 'Button Down
    DrawRectangle 1, 1, ScaleWidth - 2, ScaleHeight - 2, ShiftColor(cShadow, &H10)
    PrintText cText, -1, -1
    DrawFocusR
    DrawRectangle 0, 0, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cShadow, -&H1A), True
Case Is < 0: 'Option Button
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, ShiftColor(cLight, &H2)
    'DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cHighLight, True
    DrawFrame cShadow, cHighLight, 0, 0, 0, True
    If isEnabled = True Then
        If curStat = [Mouse Over] * -1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If hasFocus = True Then DrawFocusR
    Else
        If bkgDisabled = False Then PrintText cShadow, -1, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub NetScapeButton(bkgDisabled As Boolean)
' mostly unmolested code from original Chameleon project
If myOptMode = True And myOptValue = True Then curStat = curStat * -1 - 1
Select Case curStat
Case 0, 1, 3, 4: 'Normal, focus & mouse over
    If isEnabled = True Then
        If curStat = [Mouse Over] Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, 0
        If hasFocus Then DrawFocusR
    Else 'Disabled
        If bkgDisabled = False Then PrintText cShadow, -1, -1
        DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, 0
    End If
Case 2: 'Button Down
    DrawRectangle 0, 0, ScaleWidth - 1, ScaleHeight - 1, ShiftColor(cShadow, -&H1A), True
    PrintText cText, -1, -1
    DrawFocusR
    DrawFrame cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), 0
Case Is < 0: ' Option Button
    DrawRectangle 0, 0, ScaleWidth, ScaleHeight, cLight
    DrawFrame cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), 0
    If isEnabled = True Then
        If curStat = [Mouse Over] * 1 - 1 Then PrintText cTextO, -1, -1 Else PrintText cText, -1, -1
        If hasFocus = True Then DrawFocusR
    Else
        If bkgDisabled = False Then PrintText cShadow, -1, -1
    End If
    curStat = Abs(curStat) - 1
End Select
End Sub

Private Sub DrawFocusR()

If bShowFocus = False Then Exit Sub     ' if property prevents display a focus rectangle then don't display one

' Otherwise display a focus rectangle on button, style dependent upon button style & button type
Dim rc3 As RECT, hColor As Long
If MyButtonType = [Java metal] Or btnStyle Then
    ' this routine draws a focus rectangle just around the caption, not around inside edge of entire button
    If MyButtonType = [Java metal] Then hColor = &HCC9999
    CopyRect rc3, buttonBorder
Else
    ' this routine draws focus rectangle around inside edge of entrire button
        rc3.Top = 4
        rc3.Bottom = ScaleHeight - 4
        rc3.Left = 4
        rc3.Right = ScaleWidth - 4
End If
Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, hColor)))
DrawFocusRect hdc, rc3
UserControl.ForeColor = cText
End Sub

Private Sub SetColors(Optional bReset As Boolean)
'this function sets the colors taken as a base to build
'all the other colors and styles.

If MyColorType = Custom Or MyColorType = [Custom Gradient] Then
    cFace = BackC
    cText = ForeC
    cTextO = ForeO
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
    cDarkShadow = ShiftColor(cFace, -&HC0)
ElseIf MyColorType = [Force Standard] Then
    cFace = &HC0C0C0
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
    cTextO = cText
    cEmbossM = &HC0C0C0
    cEmbossS = &HFFFFFF
    If bReset = True Then
        BackC = cFace
        ForeC = cText
        ForeO = cTextO
    End If
ElseIf MyColorType = [Use Container] Then
    cFace = GetBkColor(UserControl.Parent.hdc)
    cText = GetTextColor(UserControl.Parent.hdc)
    cTextO = cText
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F)
    cDarkShadow = ShiftColor(cFace, -&HC0)
    cEmbossM = &HC0C0C0
    cEmbossS = &HFFFFFF
    If bReset = True Then
        BackC = cFace
        ForeC = cText
        ForeO = cTextO
    End If
Else
'if MyColorType is 1 or has not been set then use windows colors
    cFace = GetSysColor(COLOR_BTNFACE)
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
    cTextO = cText
    cEmbossM = &HC0C0C0
    cEmbossS = &HFFFFFF
    If bReset = True Then
        BackC = cFace
        ForeC = cText
        ForeO = cTextO
    End If
End If
End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, ByVal ColLight As Long, ByVal ColShadow As Long, ByVal ExtraOffset As Integer, Optional ByVal Flat As Boolean = False)

'a very fast way to draw windows-like frames
Dim pt As POINTAPI
Dim frHe As Long, frWi As Long, frXtra As Long, polyOffset As RECT

frXtra = ExtraOffset
frHe = ScaleHeight - 1
frWi = ScaleWidth - 1
If btnStyle Then            ' with segmented buttons, we offset the left and/or right margins when drawing a border (to get that parallelogram effect)
    polyOffset.Left = Choose(btnStyle, 0, SegIndent, SegIndent)
    polyOffset.Right = Choose(btnStyle, SegIndent, SegIndent, 0)
    polyOffset.Bottom = 0
End If


    Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, ColHigh)))
    '=============================
    MoveToEx hdc, frXtra + polyOffset.Left, frHe - polyOffset.Bottom, pt ' bottom left of rectangle
    LineTo hdc, frXtra, frXtra               ' vertical line up to top
    LineTo hdc, frWi - polyOffset.Right, frXtra ' horizontal line to top right
    '=============================
    Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, ColDark)))
    '=============================
    LineTo hdc, frWi, frHe                     ' vertical line down to bottom
    LineTo hdc, frXtra - 1 + polyOffset.Left, frHe - polyOffset.Bottom ' horizontal line to far left
    MoveToEx hdc, frXtra + 1 + polyOffset.Left, frHe - polyOffset.Bottom, pt  ' move to bottom left
    If Flat Then Exit Sub
    '=============================
    Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, ColLight)))
    '=============================
    LineTo hdc, frXtra + 1, frXtra + 1     ' draw vertical line to top
    LineTo hdc, frWi - 1 - polyOffset.Right, frXtra + 1   ' horizontal line to top right
    '=============================
    Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, ColShadow)))
    '=============================
    LineTo hdc, frWi - 1, frHe - 1 - polyOffset.Bottom        ' vertical line to bottom right
    LineTo hdc, frXtra + polyOffset.Left, frHe - 1 - polyOffset.Bottom ' horizontal line to bottom left
        
End Sub

Private Sub mSetPixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
' change the color of just one pixel
    Call SetPixel(hdc, x, y, Color)
End Sub

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False, Optional OtherDC As Long)
'this is a custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRect As RECT
Dim hBrush As Long
Dim Ret As Long
Dim recPts(1 To 4) As POINTAPI

If OtherDC = 0 Then OtherDC = hdc
    
bRect.Left = x
bRect.Top = y
bRect.Right = Width
bRect.Bottom = Height
hBrush = CreateSolidBrush(Color)

If OnlyBorder = False Then
    Ret = FillRect(OtherDC, bRect, hBrush)
Else
    If btnStyle Then                ' we need to create a parallelogram by plotting the points
        Width = Width
        Height = Height - 1
        Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, Color)))
        recPts(1).x = x                                                                                             ' top x,y to start
        recPts(1).y = y
        recPts(2).x = Width - Choose(btnStyle, SegIndent, SegIndent, 0)                 ' next point is width,top
        recPts(2).y = y
        recPts(3).x = Width                                                                                     ' next point is width, bottom
        recPts(3).y = Height
        recPts(4).x = x + Choose(btnStyle, 0, SegIndent, SegIndent)                         ' last is left, bottom
        recPts(4).y = Height
        Ret = Polygon(hdc, recPts(1), 4)                                                                    ' the polygon function fills in the gap between 1st & last points
    Else                            ' non segmented buttons can use this API which only supports rectangles/squares
        Ret = FrameRect(OtherDC, bRect, hBrush)
    End If
End If
Ret = DeleteObject(hBrush)
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'a fast way to draw lines
Dim pt As POINTAPI

    Call DeleteObject(SelectObject(hdc, CreatePen(PS_SOLID, 1, Color)))
    MoveToEx hdc, X1, Y1, pt
    LineTo hdc, X2, Y2
End Sub

Private Sub SetAccessKeys()
'this is a TRUE access keys parser
'the basic rule is that if an ampersand is followed by another,
'  a single ampersand is drawn and this is not the access key.
'  So we continue searching for another possible access key.
Dim ampersandPos As Long, I As Integer, J As Integer, adjCaption As String

ampersandPos = 1    ' set a starting point
I = ampersandPos    ' set flag to non-zero
Do Until I = 0
    I = InStr(ampersandPos, btnCaption, "&")
    If I Then   ' did we find one?
        ' yep, let's see if it's really two
        J = InStr(ampersandPos, btnCaption, "&&")
        If J <> I Then Exit Do  ' nope just, one -- exit now
        I = I + 1   ' really two, increment starting point
    End If
    ampersandPos = I + 1    ' set new starting point
Loop
' reset previous hotkey information
btnHotKey.cmdHotKey = 0
btnHotKey.cmdHotKeyLen = 0
btnHotKey.cmdHotKeyXY.x = 0: btnHotKey.cmdHotKeyXY.y = 0
If I Then   ' hotkey found, let's store some info on it
    AccessKeys = LCase(Mid(btnCaption, I + 1, 1))   ' letter of the hotkey
    ' the WordWrap function removes double ampersands and the ampersand associated
    '   with the hotkey when processing, so we remove them here to calculate where
    '   the adjusted hotkey position will be when these are removed
    adjCaption = Replace(btnCaption, "&&", "&")
    btnHotKey.cmdHotKey = I - (Len(btnCaption) - Len(adjCaption))
End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False, Optional sRGBval As String) As Long
'this function will add or remove a certain color
'quantity and return the result, used for the WinXP buttons

Dim Red As Long, Blue As Long, Green As Long

If isXP = False Then
    Blue = ((Color \ &H10000) Mod &H100) + Value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value
    
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If
sRGBval = Format(Red, "000") & Format(Green, "000") & Format(Blue, "000")
ShiftColor = RGB(Red, Green, Blue)
End Function

Private Sub TextFont_FontChanged(ByVal PropertyName As String)
' When user changes the font, we need to destroy/create a new logical font & wordwrap the caption within the button
Set UserControl.Font = TextFont
If Ambient.UserMode = False Then PropertyChanged "FONT"
bWordWrap = True
RefreshButton
End Sub

Private Sub TimerHover_Timer()
' When the mouse is over a button, the MouseOver status is drawn on screen the first time
' We need to determine when the mouse left our control to repaint the
' correct status. To do this we call an API which tells us the X,Y coords
' of the mouse & if it is outside our button X,Y then we can repaint

' An optional method is to use the SetCapture API. This will report mouse coords even when outside the button, and then must be released by
' the ReleaseCapture API otherwise mouse actions will continue to be sent to this program. I choose to continue to use the Timer function for
' 2 reasons: 1) Mouse generally is over a button for a very short period of time and the timer function is ok. 2) I have experienced inconsistent
' results by using the SetCapture/ReleaseCapture method... primarily the ReleaseCapture can throw out the mouse down/click event meaning
' when a user leaves this button & clicks on another one, even though ReleaseCapture was activated, the destination button may not get the actual
' down event. As a user, this is annoying when you need to click the button twice!

' However, should you want to use the SetCapture/ReleaseCapture process. Then ....
' 1) Remove the ticks infront of the two API functions in the declarations section
' 2) Remove the ticks in front of the SetCapture statement in the MouseMove event
' 3) Remove the ticks infront of the ReleaseCapture event in both the MouseMove and MouseUp events
' 4) Remove any references of TimerHover.Enabled or TimerHover.Disabled
' 5) Delete the Timer control from within design mode
' 6) Delete this subroutine

Dim pt As POINTAPI
GetCursorPos pt
If UserControl.hWnd <> WindowFromPoint(pt.x, pt.y) Then
    TimerHover.Enabled = False
    isOver = False
    If curStat <> [Disabled Status] Then isEnabled = True
    If hasFocus = True Then curStat = [Got Focus] Else curStat = [Normal Status]
    RefreshButton
    RaiseEvent MouseOut
End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
' This event occurs when the Enter key is pressed or the accelerator key
'   for the button. Since the button isn't pressed we fake it
'   by painting the down status & calling the Click event which paints the
'   normal status. The delay function simply hesitates the program for a bit
If myOptMode = False Then
    curStat = [Button Down]
    RefreshButton
    DelayMe 0.15
    LastButton = 1
End If
If myOptMode = True And myOptValue = True Then
    RaiseEvent KeyDown(KeyAscii, 0)
    RaiseEvent KeyPress(KeyAscii)
    RaiseEvent KeyUp(KeyAscii, 0)
    Exit Sub
End If
Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
' When button is using the parent's colorscheme as its own and a property
'    within that colorscheme changes that could conflict with the current button
'   colorscheme, completely redraw the button to match the parent's colorscheme
If InStr("BackColorForeColorFontPalette", PropertyName) Then
    Select Case PropertyName
    Case "Font"
        Set TextFont = Parent.Font
        Set UserControl.Font = Parent.Font
    Case "ForeColor"
        SetColors
        RefreshButton
    Case Else
        If MyColorType = [Use Container] Then
            SetColors
            RefreshButton
        End If
    End Select
End If
End Sub

Private Sub UserControl_Click()
' Not triggered directly. Mouse & Key events trigger this through code
' Works kinda like this...
' When mouse is clicked its button is saved as LastButton
' However if the right button was clicked, the event below won't fire
' But because the LastButton value <> 1, immediately pressing the Enter
' key or accelerator key fails to fire the event
' So I added the isEnabled value which is set to false when the right
' button is clicked. This way, the right button won't fire the event, but
' the Enter key can cause it sets the LastButton value to 1 before calling this event
If isEnabled = False And LastButton <> 1 Then
    ' ensure button enable flag is true if button really is enabled truly is.
    If curStat <> [Disabled Status] Then isEnabled = True
    Exit Sub
End If
' this event will trigger the Focus status for the Down button, but will not
'    send a mouse_click to the parent cause of the -1 value below
If myOptMode = True Then
    If myOptValue = False Then
        Me.Value = True
        RaiseEvent Click
    End If
    Exit Sub
End If
Call UserControl_MouseUp(-1, 1, 1, 1)
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
' Function correctly repaints the button status when button double clicked
' The -1 prevents a mouseclick being sent to parent
If LastButton = 1 Then Call UserControl_MouseDown(-1, 1, 1, 1)
End Sub

Private Sub UserControl_EnterFocus()
' Reset some basic flags when button regains focus
hasFocus = True
LastButton = 1
LastKeyDown = 0
If isOver = False Then curStat = [Got Focus]
RefreshButton
End Sub

Private Sub UserControl_ExitFocus()
' Reset flags to enable hotkeys to trigger Click event
TimerHover.Enabled = False
hasFocus = False
LastButton = 1
LastKeyDown = 0
If isOver = True Then curStat = [Mouse Over] Else curStat = [Normal Status]
RefreshButton
End Sub

Private Sub UserControl_Initialize()
' Base values for button actions/drawing
LastButton = 1      ' allows hotkeys to activate button
bShown = False      ' prevents multiple redraws until control is fully displayed
End Sub

Public Property Let Alignment(ByVal newAlign As CaptionAlignment)
Attribute Alignment.VB_Description = "Alignment of caption within the button. Read-only at runtime"
' horizontal caption alignment
If newAlign < [Left Justified] Or newAlign > [Right Justified] Or newAlign = myCaptionAlign Then Exit Property
If Ambient.UserMode = False Then
    myCaptionAlign = newAlign
    PropertyChanged "ALIGN"
    WordWrapCaption
    RefreshButton
End If
End Property

Public Property Let ShowFocus(bFocusOn As Boolean)
Attribute ShowFocus.VB_Description = "When button has focus, display an inner rectangle or highlight button. Style dependent upon button type."
If Ambient.UserMode = False Then
    PropertyChanged "SHOWF"
    bShowFocus = bFocusOn
End If
End Property

Public Property Get ShowFocus() As Boolean
ShowFocus = bShowFocus
End Property

Public Property Get GraphicsMode() As GraphicModeDat
Attribute GraphicsMode.VB_Description = "Used to force vertical text to display."
    GraphicsMode = GraphicsModeUsed
End Property

Public Property Let GraphicsMode(iGraphicsMode As GraphicModeDat)
'If Ambient.UserMode = False Then
    If iGraphicsMode > -1 Then MsgBox "Caution: Using non-default may cause buttons to not display properly on other operating systems.", vbExclamation + vbOKOnly
    If Ambient.UserMode = False Then PropertyChanged "GMODE"
    Gmode = iGraphicsMode
    If Gmode < 0 Then GraphicsModeUsed = DetermineOS Else GraphicsModeUsed = Gmode
    RefreshButton
'End If
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Applicable only when OptionButton is set to true. Either True for selected or False for unselected."
Value = myOptValue
End Property

Public Property Let Value(newValue As Boolean)
' Sets a True/False value for option buttons if in OptionButton mode
myOptValue = newValue
If myOptMode = False Then Exit Property
If myOptValue = True Then
    UpdateOptionButtons
Else
    curStat = [Normal Status]
    RefreshButton
End If
End Property

Public Property Get OptionButton() As Boolean
Attribute OptionButton.VB_Description = "IF set to True then button acts as an option button, otherwise acts as a command button."
OptionButton = myOptMode
End Property

Public Property Let OptionButton(newValue As Boolean)
' Toggles button mode to either Command Button (false) or Option Button (true)
myOptMode = newValue
If myOptMode = False Then
    curStat = [Normal Status]
    RefreshButton
Else
    Me.Value = myOptValue
End If
End Property


Public Property Get Alignment() As CaptionAlignment
' horizontal caption alignment
Alignment = myCaptionAlign
End Property

Public Property Get IconAlignmentValue() As IconAlignment
Attribute IconAlignmentValue.VB_Description = "Icon left or right edge of caption. Read-only at runtime"
' left/right alignment of icon to edge of button
IconAlignmentValue = iconAlign
End Property

Public Property Let IconAlignmentValue(newAlignment As IconAlignment)
' left/right alignment of icon to edge of button
If newAlignment < [Left Aligned] Or newAlignment > [Right Aligned] Or newAlignment = iconAlign Then Exit Property
If Ambient.UserMode = False Then
    iconAlign = newAlignment
    PropertyChanged "ICONAlign"
    If bIcon2 = False Then Exit Property
    WordWrapCaption
    RefreshButton
End If
End Property

Public Property Get EmbossEngraveShadow() As OLE_COLOR
Attribute EmbossEngraveShadow.VB_Description = "The shadow color used when embossing, engraving or shadowing button caption."
' Backcolor of button
EmbossEngraveShadow = cEmbossS
End Property

Public Property Let EmbossEngraveShadow(ByVal theCol As OLE_COLOR)
' Backcolor of button
If theCol = cEmbossS Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
        PropertyChanged "BCOL"
        PropertyChanged "EMBOSSM"
    End If
End If
cEmbossS = ConvertFromSystemColor(theCol)
If Ambient.UserMode = False Then PropertyChanged "EMBOSSS"
Call SetColors
RefreshButton
End Property

Public Property Get EmbossEngraveMid() As OLE_COLOR
Attribute EmbossEngraveMid.VB_Description = "Middle color between fore color and EmbossEngrave Shadow color. Not used with Shadow caption style."
' Backcolor of button
EmbossEngraveMid = cEmbossM
End Property

Public Property Let EmbossEngraveMid(ByVal theCol As OLE_COLOR)
' Backcolor of button
If theCol = cEmbossM Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
        PropertyChanged "BCOL"
        PropertyChanged "EMBOSSS"
    End If
End If
cEmbossM = ConvertFromSystemColor(theCol)
If Ambient.UserMode = False Then PropertyChanged "EMBOSSM"
Call SetColors
RefreshButton
End Property

Public Property Get GradientStartColor() As ColorConstants
Attribute GradientStartColor.VB_Description = "N/A for Win95 & below. Color to begin gradient fill."
    ' Gradient "From" color - see GradientStopColor()
    GradientStartColor = GStart
End Property

Public Property Let GradientStartColor(ByVal theCol As ColorConstants)
If theCol = GStart Then Exit Property
If MyButtonType = [Windows XP] Or MyButtonType = Mac Then
    If Ambient.UserMode = True Then Exit Property
    MsgBox "This option is not available for Windowx XP or Machintosh style buttons", vbInformation + vbOKOnly
    Exit Property
End If
If MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = [Custom Gradient]
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
        PropertyChanged "BCOL"
        PropertyChanged "GStop"
    End If
End If
GStart = ConvertFromSystemColor(theCol)
If Ambient.UserMode = False Then PropertyChanged "GStart"
SetColors
DrawDisabledIcon
RefreshButton
End Property

Public Property Get GradientOrientation() As GradientStyleDat
    GradientOrientation = GradientStyle
End Property

Public Property Let GradientOrientation(newVal As GradientStyleDat)
    GradientStyle = newVal
    If MyColorType <> [Custom Gradient] Then Exit Property
    If Ambient.UserMode = False Then
        PropertyChanged "GStyle"
        RefreshButton
    End If
End Property

Public Property Get GradientStopColor() As ColorConstants
Attribute GradientStopColor.VB_Description = "N/A for Win95 & below. Color to end gradient fill."
    ' Gradient "To" color - see GradientStartColor()
    GradientStopColor = GStop
End Property

Public Property Let GradientStopColor(ByVal theCol As ColorConstants)
If theCol = GStop Then Exit Property
If MyButtonType = [Windows XP] Or MyButtonType = Mac Then
    If Ambient.UserMode = True Then Exit Property
    MsgBox "This option is not available for Windowx XP or Machintosh style buttons", vbInformation + vbOKOnly
    Exit Property
End If
If MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = [Custom Gradient]
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
        PropertyChanged "BCOL"
        PropertyChanged "GStart"
    End If
End If
GStop = ConvertFromSystemColor(theCol)
If Ambient.UserMode = False Then PropertyChanged "GStop"
SetColors
DrawDisabledIcon
RefreshButton
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Not applicable unless Custom ColorScheme is applied."
' Backcolor of button
BackColor = BackC
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
' Backcolor of button
If theCol = BackC Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
    End If
End If
BackC = ConvertFromSystemColor(theCol)
If Ambient.UserMode = False Then PropertyChanged "BCOL"
Call SetColors
DrawDisabledIcon
RefreshButton
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Color of button caption font."
' Font color of button text
ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
' Font color of button text
If theCol = ForeC Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "BCOL"
        PropertyChanged "FCOLO"
    End If
End If
    ForeC = ConvertFromSystemColor(theCol)
    If Ambient.UserMode = False Then PropertyChanged "FCOL"
    Call SetColors
    RefreshButton
End Property

Public Property Get MouseOver() As OLE_COLOR
Attribute MouseOver.VB_Description = "Font color when mouse hovers over a button. CustomColor Scheme must be applied. "
' Mouse Over caption color of button
MouseOver = ForeO
End Property

Public Property Let MouseOver(ByVal theCol As OLE_COLOR)
' Mouse Over caption color of button
If theCol = ForeO Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "BCOL"
        PropertyChanged "FCOL"
    End If
End If
    ForeO = ConvertFromSystemColor(theCol)
    Call SetColors
    If Ambient.UserMode = False Then PropertyChanged "FCOLO"
    RefreshButton
End Property

Public Property Set Icon(newIcon As StdPicture)
Attribute Icon.VB_Description = "Button icon."
On Error GoTo NoIcon
Set hMyIcon = newIcon
bIcon2 = (Not newIcon Is Nothing)
If Ambient.UserMode = False Then PropertyChanged "BTNICON"
RedrawIcon:
WordWrapCaption
DrawDisabledIcon
RefreshButton
Exit Property

NoIcon:
bIcon2 = False
Set hMyIcon = Nothing
Resume RedrawIcon
End Property

Public Property Get Icon() As StdPicture
Set Icon = hMyIcon
End Property

Public Property Get ButtonType() As ButtonTypes
Attribute ButtonType.VB_Description = "One of several button types. Read-only at runtime"
' Button type
ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)
' Button type
If newValue < [Windows 16-bit] Or newValue > Mac Or newValue = MyButtonType Then Exit Property
If (newValue = Mac Or newValue = [Windows XP]) And MyColorType = [Custom Gradient] Then MyColorType = Custom
If Ambient.UserMode = False Then
    MyButtonType = newValue
    'If ButtonType = [Java metal] Then UserControl.FontBold = True
    PropertyChanged "BTYPE"
    Call UserControl_Resize
End If
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Button Caption"
' Button caption
Caption = btnCaption
End Property

Public Property Get IconSize() As IconSizeDat
Attribute IconSize.VB_Description = "Icon size to display on button. Image list name and icon index must be provided."
IconSize = myIconSize
End Property

Public Property Let IconSize(NewSize As IconSizeDat)
If NewSize < [8 x 8] Or NewSize > [32 x 32] Or NewSize = myIconSize Then Exit Property
If Ambient.UserMode = True Then Exit Property
PropertyChanged "IconSize"
myIconSize = NewSize
If bIcon2 = False Then Exit Property
WordWrapCaption
DrawDisabledIcon
RefreshButton
End Property

Public Property Let Caption(ByVal newValue As String)
' Button caption
Dim bCreateFont As Boolean
If btnCaption = "" And Len(newValue) Then bCreateFont = True
btnCaption = newValue
Call SetAccessKeys
If Ambient.UserMode = False Then PropertyChanged "TX"
bWordWrap = True
RefreshButton
End Property

Public Property Let CaptionStyle(ByVal newStyle As TextStyleDat)
Attribute CaptionStyle.VB_Description = "Engraved, Embossed, Shadowed or Plain. Applicable when ColorScheme is Custom Colors."
If newStyle < [Plain Text] Or newStyle > Shadowed Or newStyle = myTextStyle Then Exit Property
If MyColorType <> Custom And MyColorType <> [Custom Gradient] Then
    If Ambient.UserMode = True Then Exit Property
    MyColorType = Custom
    If Ambient.UserMode = False Then
        PropertyChanged "COLTYPE"
        PropertyChanged "BCOL"
        PropertyChanged "FCOL"
        PropertyChanged "FCOLO"
        PropertyChanged "EMBOSSM"
        PropertyChanged "EMBOSSS"
    End If
End If
    myTextStyle = newStyle
    Call SetColors
    If Ambient.UserMode = False Then PropertyChanged "STYLE"
    RefreshButton
End Property

Public Property Get CaptionStyle() As TextStyleDat
CaptionStyle = myTextStyle
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/set value to determine whether button can interact with user-generated events."
' Enabled/Disabled button status
Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
' Enabled/Disabled button status
isEnabled = newValue
UserControl.Enabled = isEnabled
If Ambient.UserMode = False Then PropertyChanged "ENAB"
' when re-enabling button, need to recreate the font. The font gets destroyed
'   whenever a button is disabled. This is to save memory resources since the
'   button won't be redrawing while it is disabled
RefreshButton
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Button caption font"
' Font Attributes
Set Font = TextFont
End Property

Public Property Set Font(ByRef newFont As Font)
' Font Attributes
    Set TextFont = newFont
    Set UserControl.Font = TextFont
    If Ambient.UserMode = False Then PropertyChanged "FONT"
    bWordWrap = True
    RefreshButton
End Property

Public Property Get ColorScheme() As ColorTypes
Attribute ColorScheme.VB_Description = "Most colors are ignored unless Custom ColorScheme is applied. Read-only at runtime"
ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)
' Note: Any color scheme besides "Custom" prevents changing all colors of
'   the button with the exception of the text color
If newValue < [Use Windows] Or newValue > [Custom Gradient] Or newValue = MyColorType Then Exit Property
If DetermineOS(True) < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "This option is not available on the following operating systems." & _
            vbCrLf & vbCrLf & " - Win NT 4 and lower" & vbCrLf & _
            " - Windows 95 and earlier", vbInformation + vbOKOnly
    End If
    Exit Property
End If
If Ambient.UserMode = False Then
    MyColorType = newValue
    PropertyChanged "COLTYPE"
    PropertyChanged "BCOL"
    PropertyChanged "FCOL"
    PropertyChanged "FCOLO"
    PropertyChanged "EMBOSSM"
    PropertyChanged "EMBOSSS"
    If newValue <> [Custom Gradient] And newValue <> Custom Then myTextStyle = [Plain Text]
    Call SetColors(True)
    RefreshButton
End If
End Property

Public Property Get ButtonStyle() As ButtonStyleDat
Attribute ButtonStyle.VB_Description = "Default rectangular or segmeneted"
    ButtonStyle = btnStyle
End Property

Public Property Let ButtonStyle(newStyle As ButtonStyleDat)
If newStyle < [Default Style] Or newStyle > [Right Segmented] Or newStyle = btnStyle Then Exit Property
If Ambient.UserMode = False Then
    If newStyle Then
        myOrientation = Horizontal
        PropertyChanged "ORIENT"
    End If
    btnStyle = newStyle
    PropertyChanged "BSTYLE"
    Call UserControl_Resize
End If
End Property

Public Property Get CaptionOrientation() As OrientationTypesDat
Attribute CaptionOrientation.VB_Description = "Changes text between horizontal and veritcal display. Read-only at runtime"
' Button caption to be printed vertically or horizontally
CaptionOrientation = myOrientation
End Property

Public Property Let CaptionOrientation(ByVal newOrientation As OrientationTypesDat)
If newOrientation < Horizontal Or newOrientation > [Vertical 270] Or newOrientation = myOrientation Then Exit Property
If Ambient.UserMode = False Then
    myOrientation = newOrientation
    PropertyChanged "ORIENT"
    bWordWrap = True
    RefreshButton
End If
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Sets a custom mouse pointer."
MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
UserControl.MousePointer = newPointer
If Ambient.UserMode = False Then PropertyChanged "MPTR"
End Property

Public Property Get Container() As String
Dim objName As String, objIndex As Long, I As Integer, myContainer As Object
objName = Ambient.DisplayName
I = InStr(objName, "(")
If I Then
    objIndex = Val(Mid(objName, I + 1))
    objName = Left(objName, I - 1)
    If Parent.Controls(objName).Item(objIndex).Container.Name = Parent.Name Then
        Set myContainer = Parent
        Container = Parent.Name
    Else
        Set myContainer = Parent.Controls(objName).Item(objIndex).Container
        Container = myContainer.Name
    End If
Else
    If Parent.Controls(objName).Container.Name = Parent.Name Then
        Set myContainer = Parent
        Container = myContainer.Name
    Else
        Set myContainer = Parent.Controls(objName).Container
        Container = myContainer.Name
    End If
End If
On Error Resume Next
Container = Container & "(" & myContainer.Index & ")"
End Property


Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)
On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    If Ambient.UserMode = False Then PropertyChanged "MICON"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Button Window handle"
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_InitProperties()
    isEnabled = True
    btnCaption = Ambient.DisplayName
    If btnCaption = "" Then btnCaption = UserControl.Name
    Set UserControl.Font = Ambient.Font
    UserControl.Font.Name = "Times New Roman"
    Set TextFont = UserControl.Font
    MyButtonType = [Windows 32-bit]
    btnStyle = [Default Style]
    MyColorType = [Use Windows]
    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
    GStart = 0
    GStop = vbBlue
    GradientStyle = [Left to Right]
    myCaptionAlign = [Center Justified]
    myOrientation = Horizontal
    myIconSize = [16 x 16]
    iconAlign = [Left Aligned]
    Gmode = -1
    GraphicsModeUsed = DetermineOS
    bShowFocus = True
    SetColors
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
' Sends event to parent program and moves focus to parent's next control
' if the tab, right/left arrows were pressed

LastKeyDown = KeyCode
If KeyCode = 32 And (myOptMode = False Or myOptMode = True And myOptValue = False) Then 'spacebar pressed
    ' the -1 prevents the mouseclick from being passed to parent program
    Call UserControl_MouseDown(-1, 0, 0, 0)
    RaiseEvent KeyDown(KeyCode, Shift)
    Exit Sub
End If
If myOptMode = True And KeyCode = 13 Then Exit Sub
If (KeyCode = 39) Or (KeyCode = 40) Then 'right and down arrows
    SendKeys "{Tab}"
Else
    If (KeyCode = 37) Or (KeyCode = 38) Then
        SendKeys "+{Tab}"  'left and up arrows
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
' simply send the event to the parent program
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
' Let space equal an Enter key
If (KeyCode = LastKeyDown And KeyCode = 32) And (myOptMode = False Or myOptMode = True And myOptValue = False) Then 'spacebar pressed
    ' -1 value below prevents the MouseEvent from being passed to parent
    Call UserControl_MouseUp(-1, 0, 0, 0)
    RaiseEvent KeyUp(KeyCode, Shift)
    Call UserControl_Click
Else
    RaiseEvent KeyUp(KeyCode, Shift)
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' This event repaints button and sends event to parent
' The isEnabled flag prevents right click from firing the Click event
LastButton = Abs(Button)
If Button <> 2 Then
    curStat = [Button Down]
    If myOptMode = True Then
        If myOptValue = True Then Exit Sub
        Me.Value = True
        RaiseEvent Click
    Else
        RefreshButton
    End If
Else
    isEnabled = False
End If
If Button > -1 Then RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' This event repaints button and sends event to parent, and triggers
' a timer to help indicate when mouse leaves button area

curStat = [Mouse Over]
If Button < 2 Then
    If x < 0 Or y < 0 Or x > ScaleWidth Or y > ScaleHeight Then
        'we are outside the button
        If hasFocus = True Then curStat = [Got Focus] Else curStat = [Normal Status]
        ' ReleaseCapture    ' see TimerHover_Timer for details if this is to be used
    Else
        If Button = 1 Then curStat = [Button Down]
        ' SetCapture hWnd            ' see TimerHover_Timer for details if this is to be used
    End If
Else
End If
If isOver = False Then RefreshButton
isOver = True
TimerHover.Enabled = True
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' This event repaints button and sends event to parent, unless
'   button value = -1, a flag to prevent mouseevent being passed to parent
If Button <> 2 Then
    ' repaint current button status
    If hasFocus = True Then curStat = [Got Focus] Else curStat = [Normal Status]
    If isOver = True Then curStat = [Mouse Over]
    RefreshButton
End If
' ReleaseCapture    ' see TimerHover_Timer for details if this is to be used
If Button > -1 Then RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    MyButtonType = .ReadProperty("BTYPE", [Windows 32-bit])
    btnStyle = .ReadProperty("BSTYLE", [Default Style])
    myOptValue = .ReadProperty("OPTVAL", False)
    myOptMode = .ReadProperty("OPTMOD", False)
    btnCaption = .ReadProperty("TX", UserControl.Name)
    isEnabled = .ReadProperty("ENAB", True)
    Set TextFont = .ReadProperty("FONT", UserControl.Font)
    MyColorType = .ReadProperty("COLTYPE", 1)
    BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
    ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
    ForeO = .ReadProperty("FCOLO", GetSysColor(COLOR_BTNTEXT))
    GStart = .ReadProperty("GStart", 0)
    GStop = .ReadProperty("GStop", vbBlue)
    GradientStyle = .ReadProperty("GStyle", [Left to Right])
    cEmbossM = .ReadProperty("EMBOSSM", BackC)
    cEmbossS = .ReadProperty("EMBOSSS", GetSysColor(COLOR_BTNHIGHLIGHT))
    UserControl.MousePointer = .ReadProperty("MPTR", 0)
    Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
    myCaptionAlign = .ReadProperty("ALIGN", 1)
    iconAlign = .ReadProperty("ICONAlign", 0)
    myTextStyle = .ReadProperty("STYLE", 0)
    myOrientation = .ReadProperty("ORIENT", 0)
    myIconSize = .ReadProperty("IconSize", 2)
    Set hMyIcon = .ReadProperty("BTNICON", Nothing)
    Gmode = .ReadProperty("GMODE", DetermineOS)
    If Gmode < 0 Then GraphicsModeUsed = DetermineOS Else GraphicsModeUsed = Gmode
    bShowFocus = .ReadProperty("SHOWF", True)
End With
    ' Gradients aren't supported prior to WinNT. So if system is Win95/3.1, then
    ' force the custom type vs custom gradient
    If DetermineOS(True) < 0 And MyColorType = [Custom Gradient] Then MyColorType = Custom
    UserControl.Enabled = isEnabled
    Set UserControl.Font = TextFont
    Call SetAccessKeys
    bIcon2 = (Not hMyIcon Is Nothing)
    SetColors
End Sub

Private Sub UserControl_Resize()
' only draw the buttons when the control is done resizing and is shown.
' the bShown variable is set to true in the Show event
If bShown Then
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hWnd, rgnNorm, True
    WordWrapCaption
    DrawDisabledIcon
    RefreshButton
End If
End Sub

Private Sub UserControl_Show()
' Finished loading form, display buttons
bShown = True
If myOptMode = True And myOptValue = True Then UpdateOptionButtons
Call UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
' Closing usercontrol, unload following objects if needed
On Error Resume Next
DeleteObject rgnNorm
Set TextFont = Nothing
Set hMyIcon = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("BTNICON", hMyIcon)    'picIcon.Picture)
    Call .WriteProperty("BTYPE", MyButtonType)
    Call .WriteProperty("TX", btnCaption)
    Call .WriteProperty("ENAB", isEnabled)
    Call .WriteProperty("FONT", TextFont)
    Call .WriteProperty("COLTYPE", MyColorType)
    Call .WriteProperty("BCOL", BackC)
    Call .WriteProperty("FCOL", ForeC)
    Call .WriteProperty("FCOLO", ForeO)
    Call .WriteProperty("EMBOSSM", cEmbossM)
    Call .WriteProperty("EMBOSSS", cEmbossS)
    Call .WriteProperty("MPTR", UserControl.MousePointer)
    Call .WriteProperty("MICON", UserControl.MouseIcon)
    Call .WriteProperty("ALIGN", myCaptionAlign)
    Call .WriteProperty("ICONAlign", iconAlign)
    Call .WriteProperty("ORIENT", myOrientation)
    Call .WriteProperty("STYLE", myTextStyle)
    Call .WriteProperty("IconSize", myIconSize)
    Call .WriteProperty("SHOWF", bShowFocus)
    Call .WriteProperty("BSTYLE", btnStyle)
    Call .WriteProperty("OPTVAL", myOptValue)
    Call .WriteProperty("OPTMOD", myOptMode)
    Call .WriteProperty("GStart", GStart)
    Call .WriteProperty("GStop", GStop)
    Call .WriteProperty("GStyle", GradientStyle)
End With
End Sub

Public Sub RefreshButton()
Attribute RefreshButton.VB_Description = "Redraws button"
' public function to allow users to refresh the button. This will
' completely redraw all 5 buttons and display the current button status
bShown = True
DrawButton False
UserControl.Refresh
End Sub

Private Sub MakeRegion()
'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas

Dim rgn1 As Long, rgn2 As Long, ptTRI(0 To 3) As POINTAPI
DeleteObject rgnNorm

If btnStyle Then        ' segmented button style
        ptTRI(0).x = 0  ' top left portion
        ptTRI(0).y = 0
        ''''''''''''''''''''''''''''''''
        ptTRI(1).x = Choose(btnStyle, 0, SegIndent, SegIndent) ' bot left portion
        ptTRI(1).y = ScaleHeight
        ''''''''''''''''''''''''''''''
        ptTRI(2).x = ScaleWidth    ' bot right portion
        ptTRI(2).y = ScaleHeight
        '''''''''''''''''''''''''''''
        ptTRI(3).x = ScaleWidth - Choose(btnStyle, SegIndent, SegIndent, 0) ' top right portion
        ptTRI(3).y = 0
        rgnNorm = CreatePolygonRgn(ptTRI(0), 4, 2)
Else
    rgnNorm = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    
    Select Case MyButtonType
        Case [Windows 16-bit]  'Windows 16-bit
            rgn1 = CreateRectRgn(0, 0, 1, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth - 1, ScaleHeight - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        Case [Windows XP], Mac  'Windows XP and Mac
            rgn1 = CreateRectRgn(0, 0, 2, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, ScaleHeight, 2, ScaleHeight - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 2, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth - 2, ScaleHeight - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, 1, 1, 2)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(0, ScaleHeight - 1, 1, ScaleHeight - 2)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, 1, ScaleWidth - 1, 2)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, ScaleHeight - 1, ScaleWidth - 1, ScaleHeight - 2)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        Case [Java metal]   'Java
            rgn1 = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    End Select
    DeleteObject rgn1
    DeleteObject rgn2
End If
End Sub

Private Sub DelayMe(tTime As Single)
' Little sleeper program
Dim HDelay As Single
Refresh
HDelay = Timer
Do While Timer - HDelay < tTime And Timer >= HDelay
Loop
End Sub

Private Sub PrintText(ColorMain As Long, ColorMid As Long, ColorShadow As Long)
    
    On Error GoTo errHandler
    Dim textColor(1 To 3) As Long
    Dim textOffset(1 To 3) As Integer
    Dim iCaptionSeg As Integer, cLooper As Integer
    Dim tmpX As Long, tmpY As Long, pixelOffset As Integer
    Dim txtStyle As Integer, tmpColor(0 To 3) As Long
    
    If curStat = [Button Down] And MyButtonType <> [Java metal] And myOptMode = False Then pixelOffset = 1
    ' if shadowing, embossing or engraving & button is enabled, use custom colors
    If (myTextStyle And isEnabled) Or (MyColorType = [Custom Gradient] And isEnabled = False) Then
        If (myTextStyle And isEnabled) Then
            tmpColor(1) = cEmbossM
            tmpColor(2) = cEmbossS
            txtStyle = myTextStyle
            tmpColor(0) = cText
        Else
            tmpColor(3) = GetBkColor(UserControl.Parent.hdc)
            tmpColor(1) = ShiftColor(tmpColor(3), -&HC0)
            tmpColor(2) = ShiftColor(tmpColor(3), &H1F)
            If myOptMode = True And myOptValue = True Then
                tmpColor(0) = ShiftColor(tmpColor(3), &H2F)
                txtStyle = Embossed
            Else
                tmpColor(0) = 0
                txtStyle = Engraved
            End If
            
        End If
        textColor(1) = Choose(txtStyle, tmpColor(2), tmpColor(1), tmpColor(2))
        textColor(2) = Choose(txtStyle, tmpColor(1), tmpColor(2), tmpColor(2))
        textOffset(1) = Choose(txtStyle, -1, -1, 1)
        textOffset(2) = Choose(txtStyle, 1, 1, -1)
        textColor(3) = tmpColor(0)
    Else    ' otherwise, use colors passed by the individual button draw routines
        textColor(1) = ColorShadow
        textColor(2) = ColorMid
        textOffset(1) = -1
        textOffset(2) = 1
        textColor(3) = ColorMain
        If MyColorType = [Custom Gradient] And myOptMode And myOptValue = True Then textColor(3) = GStart
    End If
    CreateDisplayFont
    For iCaptionSeg = 0 To UBound(CaptionInfo)  ' loop thru each line of caption
        For cLooper = 1 To 3                    ' and print it with designated color
            If textColor(cLooper) <> -1 Then
                UserControl.ForeColor = textColor(cLooper)
                tmpX = CaptionInfo(iCaptionSeg).cmdOffset.Left + pixelOffset
                tmpY = CaptionInfo(iCaptionSeg).cmdOffset.Top + pixelOffset
                tmpX = tmpX + textOffset(cLooper)
                tmpY = tmpY + textOffset(cLooper)
                CurrentX = tmpX  'tmpx
                CurrentY = tmpY
                Print CaptionInfo(iCaptionSeg).cmdText          'inText
            End If
        Next
    Next
    With btnHotKey              ' if button has hot key, draw the underline
        If .cmdHotKey Then
            If myOrientation = Horizontal Then
                DrawLine .cmdHotKeyXY.x + pixelOffset, .cmdHotKeyXY.y + pixelOffset, .cmdHotKeyXY.x + .cmdHotKeyLen + pixelOffset, .cmdHotKeyXY.y + pixelOffset, ColorMain
            Else
                DrawLine .cmdHotKeyXY.x + pixelOffset, .cmdHotKeyXY.y + pixelOffset, .cmdHotKeyXY.x + pixelOffset, .cmdHotKeyXY.y + .cmdHotKeyLen + pixelOffset, ColorMain
            End If
        End If
    End With
    ' draw the icon, if used
    If bIcon2 = True Then
        If isEnabled = True Then
            PaintPicture hMyIcon, iconXY.x + pixelOffset, iconXY.y + pixelOffset, myIconSize * 8, myIconSize * 8
        Else
            BitBlt hdc, iconXY.x + pixelOffset, iconXY.y + pixelOffset, myIconSize * 8, myIconSize * 8, picIcon.hdc, 0, 0, vbSrcCopy
        End If
    End If
    SelectObject hdc, hPrevFont
    DeleteObject hMyFont
   
errHandler:
End Sub


Private Sub CreateDisplayFont(Optional bCaptionFormat As Boolean = False)
' Function creates a temporary font which could be rotated, if needed

    Dim newFont As String, mPrevFont As Long
    Dim myFont As LOGFONT, newTM As TEXTMETRIC
    Dim tmpX As Long, tmpY As Long, I As Integer
    Dim mresult, fontAttr As String
    Dim iCaptionSeg As Integer, cLooper As Integer, leftOffset As Integer
    Dim textColor(1 To 3) As Long, textOffset(1 To 3), topOffset As Integer
    
    ' For Windows NT to work the GraphicsModeUsed should be 2, 0 for Win98 & earlier, 0 or 1 for Win2K
    '   Not sure for ME or XP
    mresult = SetGraphicsMode(hdc, CLng(GraphicsModeUsed))
    
    ' Start creation of new font
    newFont = TextFont.Name
    If TextFont.Bold = True Then newFont = newFont & " Bold"
    If TextFont.Italic = True Then newFont = newFont & " Italic"
    If TextFont.Strikethrough = True Then myFont.lfStrikeOut = 1
    If TextFont.Underline = True Then myFont.lfUnderline = 1
    newFont = newFont & Chr$(0)
    myFont.lfFaceName = newFont
    myFont.lfEscapement = 0
    myFont.lfHeight = (Val(TextFont.Size) * -20) / Screen.TwipsPerPixelY
    
    hMyFont = CreateFontIndirect(myFont)    ' create the font
    hPrevFont = SelectObject(hdc, hMyFont)  ' load it into the DC
    ' note: the wordwrap function won't work with rotated fonts, so we don't
    ' rotate this font for that function & rotate it later if needed
    If myOrientation Then                   ' if rotated text, warn if using
        GetTextMetrics hdc, newTM           ' a non-true type font, cause these
        If (newTM.tmPitchAndFamily) < 4 Then   ' probably won't print rotated
            If Ambient.UserMode = False Then MsgBox "Non-True Type fonts may not draw vertically", vbInformation + vbOKOnly
        End If
    End If
    ' wordwrap the caption using this font
    If bCaptionFormat = True Then WordWrapCaption
    
    If myOrientation Then       ' if rotated font, then we need to recreate one
        If myOrientation = [Vertical 90] Then   ' set the rotation degree
            myFont.lfEscapement = 900
        Else
            myFont.lfEscapement = 2700
        End If
        SelectObject hdc, hPrevFont         ' we destroy the previos version
        DeleteObject hMyFont                ' and create a new one
        hMyFont = CreateFontIndirect(myFont)
        hPrevFont = SelectObject(hdc, hMyFont)
    End If
End Sub

Private Function WordWrapCaption()
' Function will wordwrap a caption within the boundaries of the button width/height.
' Note. The DrawText API is excellent if we were only playing with horizontal buttons.
'   - it will wordwrap for us, underline hotkeys and clip the text to the button size -- all in 2 calls
'   - however, per MSDN it will not process fonts that are rotated -- therefore we can't use the API as designed
' So, I still use it to test for Font height & width to wordwrap myself & identify the hotkey character's x,y coord
'   - The UserControl TextWidth & TextHeight functions can do the same, but since this was written & tested as
'       a form vs UserControl, I use the DrawText API. Besides, I'm sure the UserControl's function just wraps
'       around this API anyway.
'   - Also, since we are printing with 1 font, we don't have the capability of underlining the hotkey for the
'       button. So, need to track its exact x,y coordinates so we can draw a line under it after we print the text
'   - Last but certainly not least, each x,y coordinate of each line of the button is saved in a variable so
'       when we have to re-print text, this routine need not be run again unless the font or caption is changed
ReDim CaptionInfo(0)

Dim xOffset As Long, yOffset As Long, iconOffset As Integer, myRC As RECT
Dim inText As String, maxW As Integer, maxH As Integer, txtRC As RECT, testString As String, iMaxLines As Integer
Dim iLines As Integer, iSpace As Integer, iChar As Integer, iLastChar As Integer
Dim iHkeyLoc As Integer, bFoundHotKey As Boolean, mIconSize As Integer

inText = Replace(btnCaption, "&&", "&")
If btnHotKey.cmdHotKey > 0 Then
    inText = Left$(inText, btnHotKey.cmdHotKey - 1) & Mid$(inText, btnHotKey.cmdHotKey + 1)
End If
If bIcon2 = True Then               ' calculate true icon width/height
    mIconSize = myIconSize * 8
        If Len(btnCaption) = 0 Then ' if no caption, then center icon on button
            iconXY.x = (ScaleWidth - mIconSize) \ 2
            iconXY.y = (ScaleHeight - mIconSize) \ 2
            If btnStyle Then iconXY.x = iconXY.x - Choose(btnStyle + 1, 0, SegIndent, 0, -SegIndent) \ 2
            Exit Function
        End If
    mIconSize = mIconSize + 5
End If
iHkeyLoc = btnHotKey.cmdHotKey                   'location of accelerator key, if any
If myOrientation Then                              '
' 0 = horizontal
' 1 = 90
' 2 = 270
    maxW = ScaleHeight - 12 - mIconSize ' use height as max width of button
    maxH = ScaleWidth   ' use width as max height of button
Else
    maxW = ScaleWidth - mIconSize - 6 - Choose(btnStyle + 1, 6, SegIndent, SegIndent * 2 + 6, SegIndent)  ' max width of button
    maxH = ScaleHeight   ' max height of button
End If

' Get the max height of the caption, the DT_SINGLELINE prevents wordwrapping, otherwise height could be more than 1 line
DrawText hdc, inText, Len(inText), txtRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX

If txtRC.Right < maxW + 1 Or txtRC.Bottom > maxH Then
    ' if entire caption fits on 1 line of button or is too tall for button then no wordwrapping needed
    iLastChar = Len(inText)
Else
    iMaxLines = maxH \ txtRC.Bottom         ' calculate the max number of lines that will fit on the button
    If iMaxLines = 0 Then iMaxLines = 1     ' ensure at least one line will be printed, even if it is off the button
    inText = inText & " "                   ' add a trailing space to the caption, helps following routine
    Do
        iSpace = InStr(iChar + 1, inText, " ")      ' find the first space
        If iSpace = 0 Or iLines = iMaxLines Then    ' did we find a space, indicating a new word?
            If iLines = iMaxLines Then Exit Do      ' yep, has max lines already been met? If so, don't print it
            GoSub WriteText                         ' if not, print that line then
            Exit Do
        Else
            testString = RTrim$(Left$(inText, iSpace))  ' Store next word in the caption
            ' Now process that word to get the height & width in pixels ('cause form scale is pixels)
            DrawText hdc, testString, Len(testString), myRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
            If myRC.Right > maxW And iLastChar Then     ' is word wider than button width? & a previous word read?
                GoSub WriteText                         ' yep, write the previous part of the caption
                iChar = 0                               ' reset starting point to search string which is truncated with WriteText sub
            Else
                ' the word in the caption is > than the button width & it is the 1st word in the string.
                '   or the word is smaller than the button width
                iLastChar = iSpace          ' Set Left$ indicator to end of the word
                iChar = iSpace              ' set staring point to search string to next word
                If myRC.Right > maxW Then
                    ' if word fits in button width, then continue on to the next word & test that string
                    '   otherwise, word is larger than width of stirng & we will print part of it
                    GoSub DoPortion         ' Sub truncates word to closely fit in button width & prints it
                    iLastChar = 0           ' Set flag to indicate new line of text
                End If
            End If
        End If
    Loop
End If
GoSub WriteText                             ' Print last line of text processed

' Got all the caption lines and lengths to be printed, along with the hotkey information, if used
' Now we need to calculate the x,y positions on the button for subsequent printing & track widest line of text
txtRC.Right = 0
txtRC.Top = 0
For iLines = 0 To UBound(CaptionInfo)
    iconOffset = myIconSize * 8 + 5
    If myOrientation Then     ' vertical text alignment
        ' get left offset by subtracting button width from nr of text lines * their height
        buttonBorder.Left = (ScaleWidth - ((UBound(CaptionInfo) + 1) * CaptionInfo(iLines).cmdOffset.Bottom)) \ 2
        ' the xOffset & yOffset below depends on 90 or 270 degree rotation
        If myOrientation = 1 Then ' bottom to top
            If bIcon2 = False Or (iconAlign = [Right Aligned] And myCaptionAlign = [Left Justified]) Then iconOffset = 0
            If bIcon2 = False Or (iconAlign = [Left Aligned] And myCaptionAlign = [Right Justified]) Then iconOffset = 0
            xOffset = 0: yOffset = ScaleHeight
        Else                    ' top to bottom
            If bIcon2 = False Or (iconAlign = [Left Aligned] And myCaptionAlign = [Right Justified]) Then iconOffset = 0
            If bIcon2 = False Or (iconAlign = [Right Aligned] And myCaptionAlign = [Left Justified]) Then iconOffset = 0
            xOffset = -ScaleWidth: yOffset = 0
        End If
        ' Determine the x coordinate, depending on whether 90 or 270 degree rotation
        CaptionInfo(iLines).cmdOffset.Left = Abs(xOffset + iLines * txtRC.Bottom + buttonBorder.Left)
        ' Now determine the y coordinate, depending on rotation & text justification
        Select Case myCaptionAlign + 1
        Case 1: ' left justify is based off either top or bottom margin of button
            CaptionInfo(iLines).cmdOffset.Top = Abs(yOffset - 6 - iconOffset)
        Case 2: ' centered
            If myOrientation = 1 Then
                If iconAlign = [Left Aligned] Then iconOffset = iconOffset \ 2 Else iconOffset = -(iconOffset \ 2)
            Else
                If iconAlign = [Right Aligned] Then iconOffset = iconOffset \ 2 Else iconOffset = -(iconOffset \ 2)
            End If
            CaptionInfo(iLines).cmdOffset.Top = Abs(yOffset - (ScaleHeight - (CaptionInfo(iLines).cmdOffset.Right)) \ 2) - iconOffset
        Case 3: ' right justify
            CaptionInfo(iLines).cmdOffset.Top = ScaleHeight - Abs(yOffset - (6 + CaptionInfo(iLines).cmdOffset.Right + iconOffset))
        End Select
    Else                    ' horizontal text alignment, a bit easier
        If bIcon2 = False Or (iconAlign = [Right Aligned] And myCaptionAlign = [Left Justified]) Then iconOffset = 0
        If bIcon2 = False Or (iconAlign = [Left Aligned] And myCaptionAlign = [Right Justified]) Then iconOffset = 0
        buttonBorder.Top = (ScaleHeight - ((UBound(CaptionInfo) + 1) * CaptionInfo(iLines).cmdOffset.Bottom)) \ 2
        CaptionInfo(iLines).cmdOffset.Top = (iLines * CaptionInfo(iLines).cmdOffset.Bottom) + buttonBorder.Top
        Select Case myCaptionAlign + 1
        Case 1: ' left justify
            CaptionInfo(iLines).cmdOffset.Left = iconOffset + Choose(btnStyle + 1, 6, 6, SegIndent, SegIndent)
        Case 2: ' centered
            If bIcon2 = False Then
                iconOffset = Choose(btnStyle + 1, 0, -SegIndent, 0, SegIndent) \ 2
            Else
                If iconAlign = [Left Aligned] And bIcon2 = True Then iconOffset = (iconOffset - Choose(btnStyle + 1, 0, SegIndent, 0, 0)) \ 2
                If iconAlign = [Right Aligned] And bIcon2 = True Then iconOffset = -((iconOffset + Choose(btnStyle + 1, 0, SegIndent, 0, 0)) \ 2)
            End If
            CaptionInfo(iLines).cmdOffset.Left = (ScaleWidth - CaptionInfo(iLines).cmdOffset.Right) \ 2 + iconOffset
        Case 3: ' right justify
            CaptionInfo(iLines).cmdOffset.Left = ScaleWidth - (CaptionInfo(iLines).cmdOffset.Right + iconOffset + (Choose(btnStyle + 1, 6, SegIndent, SegIndent, 6)))
        End Select
    End If
Next iLines
' Used for Java buttons & placement of icons, we need to determine the rectangle size that will just fit just
'   around the text. Yes, know this could have been done in the same For Next above, but it is easier to read as separate loops

If myOrientation Then
    If myOrientation = 1 Then     ' set min/max flags for comparison below
        buttonBorder.Left = CaptionInfo(0).cmdOffset.Left
        buttonBorder.Right = CaptionInfo(UBound(CaptionInfo)).cmdOffset.Left + CaptionInfo(UBound(CaptionInfo)).cmdOffset.Bottom
        buttonBorder.Top = ScaleHeight
        buttonBorder.Bottom = 0
        For iLines = 0 To UBound(CaptionInfo)
            yOffset = Abs((CaptionInfo(iLines).cmdOffset.Top - CaptionInfo(iLines).cmdOffset.Right))
            If yOffset < buttonBorder.Top Then buttonBorder.Top = yOffset
            If CaptionInfo(iLines).cmdOffset.Right > buttonBorder.Bottom Then buttonBorder.Bottom = CaptionInfo(iLines).cmdOffset.Right
        Next
        buttonBorder.Bottom = buttonBorder.Top + buttonBorder.Bottom
    Else
        buttonBorder.Left = CaptionInfo(UBound(CaptionInfo)).cmdOffset.Left - CaptionInfo(0).cmdOffset.Bottom
        buttonBorder.Right = CaptionInfo(0).cmdOffset.Left
        buttonBorder.Top = ScaleHeight
        buttonBorder.Bottom = 0
        For iLines = 0 To UBound(CaptionInfo)
            If CaptionInfo(iLines).cmdOffset.Top < buttonBorder.Top Then buttonBorder.Top = CaptionInfo(iLines).cmdOffset.Top
            If CaptionInfo(iLines).cmdOffset.Right > buttonBorder.Bottom Then buttonBorder.Bottom = CaptionInfo(iLines).cmdOffset.Right
        Next
        buttonBorder.Bottom = buttonBorder.Bottom + buttonBorder.Top
    End If
Else
    buttonBorder.Top = CaptionInfo(0).cmdOffset.Top
    buttonBorder.Bottom = CaptionInfo(UBound(CaptionInfo)).cmdOffset.Top + CaptionInfo(0).cmdOffset.Bottom
    buttonBorder.Right = 0
    buttonBorder.Left = ScaleWidth
    For iLines = 0 To UBound(CaptionInfo)
        If CaptionInfo(iLines).cmdOffset.Left < buttonBorder.Left Then buttonBorder.Left = CaptionInfo(iLines).cmdOffset.Left
        If CaptionInfo(iLines).cmdOffset.Right > buttonBorder.Right Then buttonBorder.Right = CaptionInfo(iLines).cmdOffset.Right
    Next
    buttonBorder.Right = buttonBorder.Right + buttonBorder.Left
End If
    
' Almost done, need to calculate x,y coords to underline the hotkey if one was used
If bFoundHotKey Then                    ' was one used and found?
    iLines = btnHotKey.cmdHotKeyXY.x    ' let's get button caption line where hotkey was found
    ' we want to get the length of the string to that position
    DrawText hdc, Left$(CaptionInfo(iLines).cmdText, iHkeyLoc), iHkeyLoc, myRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
    If myOrientation Then
        If myOrientation = 1 Then    ' vertical bottom to top
            btnHotKey.cmdHotKeyXY.x = CaptionInfo(iLines).cmdOffset.Left + CaptionInfo(iLines).cmdOffset.Bottom
            btnHotKey.cmdHotKeyXY.y = CaptionInfo(iLines).cmdOffset.Top - myRC.Right
        Else                        ' vertical top to bottom
            btnHotKey.cmdHotKeyXY.x = CaptionInfo(iLines).cmdOffset.Left - CaptionInfo(iLines).cmdOffset.Bottom
            btnHotKey.cmdHotKeyXY.y = CaptionInfo(iLines).cmdOffset.Top + myRC.Right - btnHotKey.cmdHotKeyLen
        End If
    Else                            ' horizontal left to right
        btnHotKey.cmdHotKeyXY.x = myRC.Right + CaptionInfo(iLines).cmdOffset.Left - btnHotKey.cmdHotKeyLen
        btnHotKey.cmdHotKeyXY.y = CaptionInfo(iLines).cmdOffset.Top + CaptionInfo(iLines).cmdOffset.Bottom
    End If
End If

' Now to calculate the Icon x,y coords if one was used
If bIcon2 Then
    If myOrientation = Horizontal Then
        If iconAlign = [Left Aligned] Then
            iconXY.x = buttonBorder.Left - myIconSize * 8 - 5
        Else
            iconXY.x = buttonBorder.Right + 5
        End If
        iconXY.y = (buttonBorder.Bottom - buttonBorder.Top - myIconSize * 8) \ 2 + buttonBorder.Top
    Else
        iconXY.x = (buttonBorder.Right - buttonBorder.Left - myIconSize * 8) \ 2 + buttonBorder.Left
        If myOrientation = [Vertical 270] Then ' left aligned Icon
            If iconAlign = [Left Aligned] Then
                iconXY.y = buttonBorder.Top - myIconSize * 8 - 5
            Else
                iconXY.y = buttonBorder.Bottom + 5
            End If
        Else
            If iconAlign = [Left Aligned] Then
                iconXY.y = buttonBorder.Bottom + 5
            Else
                iconXY.y = buttonBorder.Top - myIconSize * 8 - 5
            End If
        End If
    End If
End If
InflateRect buttonBorder, 2, 2
OffsetRect buttonBorder, -1, -1
Exit Function


' Subroutine truncates a word that is too long to display on the button & adds an elipse to indicate to a
'   programmer/user that the button caption exceeds the width/height of the button
DoPortion:
' Get the actual width of the word
DrawText hdc, inText, Len(testString), myRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
' now calculate a percentage of the text that will fit (homemade clip function)
iChar = (Int(maxW / myRC.Right * Len(inText)) \ 2)
If iChar < 1 Then iChar = 1
testString = Left$(inText, iChar)
' remove last 3 characters (if applicable) so we can add the trailing elipse
If Len(testString) > 4 Then testString = Left$(testString, Len(testString) - 3)
testString = testString & "..."
iLastChar = Len(testString)   ' Set Left$ flag so following GOSUB will print it
inText = testString & Mid$(inText, iSpace + 1)  ' prepare rest of string for processing
iChar = 0   ' flag to indicate more text left in caption

' Subroutine stores location of the text and tracks location of hotkey, if used
WriteText:
If Len(inText) = 0 Then Return
ReDim Preserve CaptionInfo(0 To iLines)
testString = RTrim$(Left$(inText, iLastChar))   ' truncate the text to print on the button
If iHkeyLoc And bFoundHotKey = False Then       ' does a hot key exist?
    If iHkeyLoc > Len(testString) Then          ' if so, is its location in this portion of the caption?
        iHkeyLoc = iHkeyLoc - Len(testString) - 1   ' nope, offset the location to to compare with next portion of string
    Else                                            ' yep, lets get how long the underline should be
        ' do this by getting the width of the hotkey character as printed
        DrawText hdc, Mid$(inText, iHkeyLoc, 1), 1, txtRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
        btnHotKey.cmdHotKeyLen = txtRC.Right        ' store the length of the underline
        btnHotKey.cmdHotKeyXY.x = iLines            ' store relationship to line of button text for later
        bFoundHotKey = True
    End If
End If
' now to store size and caption of the button text, run API to calculate width of text
DrawText hdc, testString, Len(testString), myRC, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
' save the text & store lengths
CaptionInfo(iLines).cmdText = RTrim$(Left$(inText, iLastChar))
CaptionInfo(iLines).cmdOffset.Right = myRC.Right
CaptionInfo(iLines).cmdOffset.Bottom = myRC.Bottom
' increment number of text lines to be printed & truncate the button caption less the text already processed
iLines = iLines + 1
inText = Mid$(inText, iLastChar + 1)
Return
End Function

Private Function ConvertFromSystemColor(tColor As Long) As Long
'System colors in VB are &H80000000& to &H80000018& (-2147483648 to -2147483624).
'All other colors are &H00000000& to &H00FFFFFF& (0 to 16777215).

' Function used when programmer selects a system color for a property
' System colors will always print black unless converted.
If tColor < -1 Then  'If it's a System color...
    ConvertFromSystemColor = GetSysColor(tColor And &HFF&)
Else
    ConvertFromSystemColor = tColor
End If
End Function

Private Function DetermineOS(Optional bIsWin95 As Boolean = False) As Integer
' Determine OS. An API needs to run when trying to print rotated text. The function is activated whenever creating
'   a font in the CreateDisplayFont function.
' Known settings: Tested settings are shown below
'   According to MSDN, NT 3.51 only works on a setting of 2. Don't have the opportunity to test this.
' Don't know about ME or XP, hoping one of you will tell me which works with it.

' The following are the platform, major version & minor version of OS to date (acquired from MSDN)
Const isWin95 = "1.4.0"
Const isWin98 = "1.4.10"
Const isWinNT4 = "2.4.0"
Const isWinNT351 = "2.3.51"
Const isWin2K = "2.5.0"
Const isWinME = "1.4.90"
Const isWinXP = "2.5.1"

  Dim verinfo As OSVERSIONINFO, sVersion As String
  verinfo.dwOSVersionInfoSize = Len(verinfo)
  If (GetVersionEx(verinfo)) = 0 Then Exit Function         ' use default 0
  With verinfo
    sVersion = .dwPlatformId & "." & .dwMajorVersion & "." & .dwMinorVersion
  End With
  Select Case sVersion
  Case isWin98, isWin2K:            DetermineOS = 0     'tested
  Case isWinNT4:                    DetermineOS = 1     'tested
  Case isWinNT351:                  DetermineOS = 2     'untested to date
  Case isWin95, isWinXP, isWinME:   DetermineOS = 0     'untested to date
  End Select
  ' These systems do not suppor the gradient fill API call, per MSDN. So when
  ' called, return a negative value to indicate system doesn't meet requirements
  If bIsWin95 = True And (sVersion = isWin95 Or sVersion = isWinNT4 Or sVersion = isWinNT351) Then DetermineOS = -1
End Function

Private Sub DrawDisabledIcon()
' Function will create a disabled icon in the picture box to be used when Icon assigned and button is disabled

If bIcon2 = False Then Exit Sub     ' no icon, no function
Dim x As Long, y As Long, lColor As Long, dColor As Long, lDrawType As Long
Const DSS_DISABLED = &H20
Const DSS_NORMAL = &H0
Const DSS_BITMAP = &H4
Const DSS_ICON = &H3

    ' determine the pixel dimensions of the picture being used as an icon
    x = CLng(ScaleX(hMyIcon.Width, vbHimetric, vbPixels))
    y = CLng(ScaleY(hMyIcon.Height, vbHimetric, vbPixels))
    DrawRectangle 0, 0, x, y, cFace     ' create a blank area to transfer picture on, using the background color
    With picIcon                        ' resize the picture box to the eventual Icon Size
        .Height = myIconSize * 8
        .Width = myIconSize * 8
        .Left = Abs(.Width - myIconSize) \ 2
        Set .Picture = Nothing
        .Cls
        DrawRectangle 0, 0, .Width, .Height, cFace, , .hdc          ' create blank area in picture box to draw disabled icon
        ' call function to draw the disabled icon. When the icon is drawn it can't be resized simultaneously
        ' - it does draw it with transparency if exists in the icon
        ' - that's why the background color is important - you'll see later
        If hMyIcon.Type = 1 Then lDrawType = DSS_BITMAP Else lDrawType = DSS_ICON
        lDrawType = lDrawType Or DSS_DISABLED
        DrawState hdc, 0, 0, hMyIcon, 0, 0, 0, x, y, lDrawType
        ' now copy the icon into the picture box in the desired size
        ' - this function copies the background color along with the icon, stretching the icon to the desired size
        StretchBlt .hdc, 0, 0, myIconSize * 8, myIconSize * 8, hdc, 0, 0, x, y, vbSrcCopy
    End With
    DrawButton True     ' draw the disabled button, less the text so we know what it will look like in
                                   ' disabled mode. I.E. most buttons have different backcolors when disabled and may be multicolored
    ' now, loop through each of the pixels in the picture box and only process the pixels that
    ' are equal to the back color of the picturebox, leaving the disabled icon intact (it should not contain background colors)
    For x = 0 To myIconSize * 8 - 1
        For y = 0 To myIconSize * 8 - 1
            dColor = GetPixel(picIcon.hdc, x, y)
            If dColor = cFace Then  ' if a non-disabled icon pixel, continue...
                ' get color from the disabled button & transfer color to the disabled icon background
                lColor = GetPixel(hdc, x + iconXY.x + 1, y + iconXY.y + 1)
                If lColor > -1 And lColor <> cFace Then SetPixel picIcon.hdc, x, y, lColor
            End If
        Next
    Next
End Sub

Private Sub UpdateOptionButtons()
Dim myContainer As String, myHwnd As Long, lControls As Long, tgtObj As BTL
myContainer = Container
myHwnd = Me.hWnd
For lControls = 0 To ParentControls.Count - 2
    If TypeOf Parent.Controls(lControls) Is BTL Then
        Set tgtObj = Parent.Controls(lControls)
        If tgtObj.Container = myContainer Then
            If tgtObj.hWnd <> myHwnd Then
                If tgtObj.Value = True Then tgtObj.Value = False
            Else
                RefreshButton
            End If
        End If
    End If
Next
Set tgtObj = Nothing
End Sub

Private Function ConvertToRGB(HexLng As String) As String
'=======================================================================
'   'This will convert Hexidecimal color coding to RGB color coding
'       variables passed must either be numeric or in the format of &H########
'=======================================================================
' Inserted by LaVolpe
On Error GoTo Function_ConvertToRGB_General_ErrTrap_by_LaVolpe
If IsNumeric(Mid(HexLng, 2)) Then
    If Val(HexLng) < 0 Then HexLng = 255
    HexLng = BigDecToHex(HexLng)
End If
'For Convert Hexidecimal to RGB:  Converts Hexidecimal to RGB
On Error GoTo errorsub
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
Tmp = HexLng
If UCase(Left$(HexLng, 2)) = "&H" Then Tmp = Mid$(HexLng, 3)
Tmp = Right$("0000000" & Tmp, 8)
If IsNumeric(Hx & Tmp) Then
lo1 = CInt(Hx & Right$(Tmp, Two))       ' Red
hi1 = CLng(Hx & Mid$(Tmp, 5, Two))   ' Green
lo2 = CInt(Hx & Mid$(Tmp, 3, Two))     ' blue
hi2 = CLng(Hx & Left$(Tmp, Two))
If lo1 > 0 Then lo1 = lo1 + 1
If lo2 > 0 Then lo2 = lo2 + 1
If hi1 > 0 Then hi1 = hi1 + 1
ConvertToRGB = Format(lo1, "000") & Format(hi1, "000") & Format(lo2, "000")
End If
Exit Function

errorsub:  MsgBox Err.Description, vbExclamation, "Error"
Exit Function

Function_ConvertToRGB_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
'If MsgBox("Error " & Err.Number & " - Procedure [Function ConvertToRGB]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Private Function BigDecToHex(ByVal DecNum) As String
'=======================================================================
'   Used to convert any decimal value to hex equivalent
'=======================================================================
    ' This function is 100% accurate untill
    '     15,000,000,000,000,000 (1.5E+16)
    Dim NextHexDigit As Double
    Dim HexNum As String

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo BigDecToHex_General_ErrTrap

HexNum = ""
While DecNum <> 0
    NextHexDigit = DecNum - (Int(DecNum / 16) * 16)
    If NextHexDigit < 10 Then
        HexNum = Chr(Asc(NextHexDigit)) & HexNum
    Else
        HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
    End If
    DecNum = Int(DecNum / 16)
Wend

If HexNum = "" Then HexNum = "0"
BigDecToHex = HexNum
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
BigDecToHex_General_ErrTrap:
'MsgBox "Err: " & Err.Number & " - Procedure: BigDecToHex" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Private Sub DoGradientFill()
' Function extracted from AllAPI.net
' Basically sets up graident fills for any HDC object

    Dim vert(1) As TRIVERTEX, sRGB As String
    Dim gRect As GRADIENT_RECT, iOffset As Integer

    If GradientStyle > GRADIENT_FILL_RECT_V Then iOffset = 1
    'from GStart color
    ' We need to convert OLE colors to RGB
    'sRGB = ConvertToRGB(CStr(GStart))
    ShiftColor GStart, 0, , sRGB
    ' This function requires the RGB values to be negative-1 if non-zero
    With vert(0 + iOffset)
        .x = 0 + (UserControl.ScaleWidth * iOffset)
        .y = 0 + (UserControl.ScaleHeight * iOffset)
        .Red = -Val(Left(sRGB, 3))
        .Green = -Val(Mid(sRGB, 4, 3))
        .Blue = -Val(Right(sRGB, 3))
        If .Red < 0 Then .Red = .Red - 1
        If .Blue < 0 Then .Blue = .Blue - 1
        If .Green < 0 Then .Green = .Green - 1
        .Alpha = 0&
    End With

    'to GStop color
    'sRGB = ConvertToRGB(CStr(GStop))
    ShiftColor GStop, 0, , sRGB
    With vert(1 - iOffset)
        .x = UserControl.ScaleWidth - (UserControl.ScaleWidth * iOffset)
        .y = UserControl.ScaleHeight - (UserControl.ScaleHeight * iOffset)
        .Red = -Val(Left(sRGB, 3))
        .Green = -Val(Mid(sRGB, 4, 3))
        .Blue = -Val(Right(sRGB, 3))
        If .Red < 0 Then .Red = .Red - 1
        If .Blue < 0 Then .Blue = .Blue - 1
        If .Green < 0 Then .Green = .Green - 1
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect UserControl.hdc, vert(0), 2, gRect, 1, GradientStyle - (iOffset * 2)
End Sub

VERSION 5.00
Begin VB.UserControl ScrollBox 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   4530
   Begin VB.VScrollBar VScroll 
      Height          =   1575
      Left            =   2400
      Max             =   115
      SmallChange     =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      Max             =   80
      SmallChange     =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox pCorner 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2400
      ScaleHeight     =   375
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1440
      Width           =   315
   End
   Begin VB.PictureBox pView 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   0
      Width           =   2595
   End
   Begin VB.Image curMove 
      Height          =   480
      Left            =   2520
      Top             =   1320
      Width           =   480
   End
End
Attribute VB_Name = "ScrollBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private gScaleX As Single
Private gScaleY As Single
Private lPrevParent As Long
Private WithEvents pChild As PictureBox
Attribute pChild.VB_VarHelpID = -1
Const m_Def_Align = 0
Const m_def_BackColor = &H8000000C


Private m_Align As Integer
Private m_BackColor As OLE_COLOR

Event Resize()
Event Scroll()

Private Const WM_SIZE = &H5

Private Declare Function SetParent _
    Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long

Private Declare Function GetCursorPos _
    Lib "user32" ( _
    lpPoint As POINTAPI) As Long
    
Private Declare Function GetWindowRect _
    Lib "user32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT) As Long

Private Sub pView_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdatePos
End Sub

Private Sub UserControl_Resize()
    Dim loff As Integer
    Dim loffV As Integer
    Dim loffH As Integer
    Dim sV As Single
    Dim sH As Single
    
    On Error Resume Next
    
    loffV = 39
    loffH = 45
    
    Call VScroll.Move(UserControl.Width - VScroll.Width - loffV, 0, VScroll.Width, UserControl.Height - HScroll.Height - loffH)
    Call HScroll.Move(0, UserControl.Height - HScroll.Height - loffH, UserControl.Width - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(UserControl.Width - VScroll.Width - loffV, UserControl.Height - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(0, 0, Width - VScroll.Width, Height - HScroll.Height)
    
    HScroll.Min = 1
    VScroll.Min = 1
    
    sH = pChild.Width - pView.Width
    sV = pChild.Height - pView.Height
    
    If sV = 0 Then
        VScroll.Max = 1
        VScroll.Width = 0
        VScroll.Left = UserControl.Width
        loffV = 37
    ElseIf sV < 0 Then
        VScroll.Max = 1
        VScroll.Width = 0
        VScroll.Left = UserControl.Width
        loffV = 37
    Else
        VScroll.Max = sV
        VScroll.Width = 255
    End If
    

    If sH = 0 Then
        HScroll.Max = 1
        HScroll.Height = 0
        loffH = 25
    ElseIf sH < 0 Then
        HScroll.Max = 1
        HScroll.Visible = False
        HScroll.Height = 0
        loffH = 25
    Else
        HScroll.Max = sH
        HScroll.Visible = True
        HScroll.Height = 255
    End If
    
    Call VScroll.Move(UserControl.Width - VScroll.Width - loffV, 0, VScroll.Width, UserControl.Height - HScroll.Height - loffH)
    Call HScroll.Move(0, UserControl.Height - HScroll.Height - loffH, UserControl.Width - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(UserControl.Width - VScroll.Width - loffV, UserControl.Height - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(0, 0, Width - VScroll.Width, Height - HScroll.Height)
    
    HScroll.LargeChange = UserControl.Width
    VScroll.LargeChange = UserControl.Height
    
    RaiseEvent Resize
    
End Sub

Private Sub pChild_Resize()
    Call UserControl_Resize
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    pView.BackColor = m_BackColor
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)

End Sub

Private Sub UserControl_InitProperties()
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
    
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
    
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    pView.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
    
End Property


Private Sub VScroll_Change()
    UpdatePos
    
End Sub
   
Private Sub HScroll_Change()
    UpdatePos
    
End Sub

Sub UpdatePos()
    
    On Error Resume Next
    pChild.Move -HScroll.Value, -VScroll.Value
    pView.SetFocus
    RaiseEvent Scroll
    
End Sub

Public Sub Attatch(newChild As PictureBox)
    Set pChild = newChild
    lPrevParent = SetParent(newChild.hwnd, pView.hwnd)
    pChild.Move 0, 0
    pChild.MouseIcon = curMove.Picture
    pChild.MousePointer = 0
    UserControl_Resize
    UpdatePos
    
End Sub

Public Sub Detatch()
    SetParent pChild.hwnd, lPrevParent
    Set pChild = Nothing
    
End Sub

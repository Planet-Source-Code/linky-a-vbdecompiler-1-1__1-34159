VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIconsExtract 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extraire des Icones depuis un EXE,"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDL 
      Left            =   1680
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "icl"
      DialogTitle     =   "Open ICL Library"
      Filter          =   "Icon Libraries (*.icl;*.ni)|*.icl;*.ni;*.il|Icons (*.ico)|*.ico|Executables (*.exe;*.dll)|*.exe;*.dll|All files|*.*"
   End
   Begin MSComDlg.CommonDialog cdlSave 
      Left            =   1200
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btExtract 
      Caption         =   "Extraire..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   2280
   End
   Begin VB.CommandButton btOpen 
      Caption         =   "Selectionnez le fichier..."
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2280
   End
   Begin VB.CommandButton btClose 
      Caption         =   "Fermer"
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Width           =   2280
   End
   Begin VB.CheckBox chSel 
      Caption         =   "Les icônes séléctionné seulement"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2010
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2235
   End
   Begin VB.DirListBox Dir1 
      Height          =   5265
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2235
   End
   Begin MSComctlLib.ProgressBar PROG 
      Height          =   210
      Left            =   2280
      TabIndex        =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lvICONS 
      Height          =   6975
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   12303
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList IML 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   14805982
      _Version        =   393216
   End
   Begin VB.Label lbLib 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please, select an icon library."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   3885
   End
   Begin VB.Label lbStat 
      Caption         =   "Lecture."
      Height          =   240
      Left            =   2280
      TabIndex        =   8
      Top             =   7440
      Width           =   5550
   End
   Begin VB.Menu mnSelect 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnSaveAs 
         Caption         =   "Sauver sous..."
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnExtract 
         Caption         =   "Extraire les icônes séléctionnées"
      End
      Begin VB.Menu mnExtractAll 
         Caption         =   "Extraire touts les icônes"
      End
   End
End
Attribute VB_Name = "frmIconsExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Dim sLibraryName As String
    
Dim oIcon As cFileIcon
Dim IconName() As String
Dim TransparentColor As Long

Dim IconCount
Dim hModule

Dim Iconh&
Dim x&

Private Type tDeviceImage
   iSizeX As Long
   iSizeY As Long
   cDepth As Long
   cPal As cPalette
End Type

Private m_tDeviceImage As tDeviceImage

Private Sub btClose_Click()
    Set oIcon = Nothing
    Unload Me
End Sub

Private Sub btExtract_Click()
    ExtractIcons chSel.Value = 1
End Sub

Private Sub btOpen_Click()
    On Error GoTo 100
    CDL.CancelError = True
    CDL.ShowOpen
    sLibraryName = CDL.FileName
    lbLib = sLibraryName
    ReadLibrary
10
    Exit Sub
100
    Resume 10
End Sub

Private Sub Form_Load()
  cLanguage.SetLanguageInForm Me
    Set oIcon = New cFileIcon
    Drive1.Drive = Left(App.Path, 1)
    Dir1.Path = App.Path
    CDL.initDir = App.Path
    hModule = Me.hwnd
    Pic.AutoRedraw = True
    With m_tDeviceImage
      .iSizeX = 32
      .iSizeY = 32
      .cDepth = 256
      Set .cPal = New cPalette
        .cPal.CreateWebSafe
    End With
    TransparentColor = IML.MaskColor
    PForm Me, True
    Me.Show

End Sub

Private Sub GetIconNames()
    Dim s As String, FN As Long
    Dim x1 As Long, i As Long
    Dim Cnt As Long
    Dim Z As Long
    ReDim IconName(1 To IconCount)
    If Right(sLibraryName, 3) = "icl" Then
        Cnt = 0
        FN = FreeFile
        Open sLibraryName For Binary As #FN
        s = Space(LOF(FN))
        Get #FN, , s
        Close #FN
        x1 = InStr(1, s, "ICL", vbBinaryCompare)
        If x1 = 0 Then GoTo PutNumberedNames
        x1 = x1 + 3
        Do
                                                
                                                 
            Z = Asc(Mid(s, x1, 1))
            If Z = 0 Then Exit Do
            Cnt = Cnt + 1
            IconName(Cnt) = Mid(s, x1 + 1, Z)
            x1 = x1 + Z + 1
        Loop
        s = ""
    ElseIf Right(sLibraryName, 3) = "ico" Then
        
        For i = Len(sLibraryName) - 5 To 1 Step -1
            If Mid(sLibraryName, i, 1) = "\" Then
                IconName(1) = Mid(sLibraryName, i + 1, Len(sLibraryName) - i - 4)
                Exit For
            End If
        Next i
    Else
PutNumberedNames:
        
        For i = 1 To IconCount
            IconName(i) = "Icon" + Format(i, "0000")
        Next i
    End If
End Sub

Sub ReadLibrary()
    Dim i As Long
    Dim sAPILibName As String
    
    On Error GoTo 100
    sAPILibName = sLibraryName + Chr$(0)
    IconCount = ExtractIcon(hModule, sAPILibName, -1)
    
    lvICONS.Icons = Nothing
    IML.ListImages.Clear
    lvICONS.ListItems.Clear
    
    If IconCount > 0 Then
        lbStat = "Loading library..."
        lbInfo.Caption = "This file contains " + CStr(IconCount) + " icon/s."
        PROG.Max = IconCount
        PROG.Visible = True
        lvICONS.Visible = False
        GetIconNames
        For i = 1 To IconCount
            Set Pic.Picture = LoadPicture("")
            Iconh = ExtractIcon(hModule, sAPILibName, i - 1)
            x& = DrawIcon(Pic.hdc, 0, 0, Iconh)
            IML.ListImages.Add i, , Pic.Image
            lvICONS.Icons = IML
            lvICONS.ListItems.Add , , IconName(i), i
            PROG.Value = i
            DoEvents
        Next i
        lvICONS.Visible = True
        PROG.Visible = False
        lbStat = "Ready."
    Else
        lbInfo = "This file does not contain icons."
    End If
    btExtract.Enabled = IconCount > 0
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

Sub ExtractIcons(bSelectedOnly As Boolean)
    Dim i, j
    On Error GoTo 100
    Pic.Visible = True
    PROG.Visible = True
    lbStat = "Extraction des icônes..."
    For j = 1 To IconCount
        If bSelectedOnly Then
            If lvICONS.ListItems(j).Selected = False Then GoTo SkipIcon
        End If
        SaveIconToFile j, ToPath(Dir1.Path) + IconName(j) + ".ico"
        PROG.Value = j
        DoEvents
SkipIcon:
    Next j
    Pic.Visible = False
    PROG.Visible = False
    lbStat = "Ready."
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub mnSaveAs_Click()
    Dim sSaveName As String
    Dim IconIndex As Long
    IconIndex = lvICONS.SelectedItem.Index
    On Error GoTo 100
    With cdlSave
        .CancelError = True
        .DialogTitle = "Sauver les icones sous..."
        .Filter = "*.ico"
        .DefaultExt = "ico"
        .FileName = IconName(IconIndex)
        .initDir = Dir1.Path
        .ShowSave
        sSaveName = .FileName
    End With
    
    On Error GoTo 200
    SaveIconToFile IconIndex, sSaveName
10
    Exit Sub
100
    Resume 10
200
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub
Private Sub mnExtract_Click()
    ExtractIcons True
End Sub

Private Sub mnExtractAll_Click()
    ExtractIcons False
End Sub


Private Sub lvICONS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnSelect
    End If
End Sub

Private Sub lvICONS_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    mnSaveAs.Enabled = False
    mnExtract.Enabled = False
    For i = 1 To lvICONS.ListItems.Count
        If lvICONS.ListItems(i).Selected Then
            mnSaveAs.Enabled = True
            mnExtract.Enabled = True
            Exit For
        End If
    Next i
    
End Sub
Private Sub SaveIconToFile(ByVal Index As Long, ByVal SaveName As String)
    Set Pic.Picture = IML.ListImages(Index).Picture
    oIcon.AddImage m_tDeviceImage.iSizeX, m_tDeviceImage.iSizeY, m_tDeviceImage.cDepth
    m_tDeviceImage.cPal.SetPaletteToIcon oIcon, 1
    oIcon.SetIconFromBitmap Pic.hdc, 1, 0, 0, True, TransparentColor
    oIcon.SaveIcon SaveName
    oIcon.RemoveImage 1
End Sub

Function ToPath(ByVal s As String) As String
    If Right(s, 1) <> "\" Then s = s & "\"
    ToPath = s
End Function

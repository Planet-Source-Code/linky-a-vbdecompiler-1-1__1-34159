VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VisualBasic Decompiler"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9570
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmExt 
      Height          =   4815
      Left            =   2520
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VBDecompiler.BTL cmdExtIco 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Icônes"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":08CA
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VBDecompiler.BTL BTL2 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "Resources"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":08E6
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Extraire..."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5175
      End
   End
   Begin VBDecompiler.BTL BTL1 
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "A propos..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":0902
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Frame frmInfo 
      Height          =   1455
      Left            =   2520
      TabIndex        =   9
      Top             =   5760
      Width           =   5415
      Begin VB.Label lblInfo 
         Caption         =   $"frmMain.frx":091E
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5175
      End
   End
   Begin VBDecompiler.XpB XpB2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Extraire les resources"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   12746315
   End
   Begin VBDecompiler.XpB XpB1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Mettre à jour..."
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648384
   End
   Begin VBDecompiler.XpB XpB 
      Height          =   615
      Left            =   8160
      TabIndex        =   7
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Afficher le tableau"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   720
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VBDecompiler.XpB Generator 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Generer des Forms"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VBDecompiler.PBar PBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   238
      Appearance      =   0
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      FillColor       =   11643476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionStyle    =   2
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBDecompiler.XpB Decompiler 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Decompiler..."
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9570
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neoware, ® 1995-2002"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8130
         TabIndex        =   11
         Top             =   30
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   120
         Picture         =   "frmMain.frx":09C2
         Top             =   0
         Width           =   2460
      End
   End
   Begin VBDecompiler.XpB editer 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Editer..."
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VBDecompiler.XpB exporter 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Exporter..."
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VBDecompiler.BTL BTL3 
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Langages..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMain.frx":73B8
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   8040
      X2              =   8040
      Y1              =   960
      Y2              =   7200
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   2400
      X2              =   2400
      Y1              =   960
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   960
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   960
      Y2              =   7200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTL1_Click()
frmAbout.Show

End Sub




Private Sub BTL3_Click()
frmLangage.Show

End Sub

Private Sub cmdExtIco_Click()
frmIconsExtract.Show
End Sub

Private Sub Decompiler_Click()
On Error Resume Next
cd1.DialogTitle = "Selectionnez le fichier à décompiler"
cd1.CancelError = True
cd1.Filter = "Executables (*.exe) |*.exe|DLL (*.dll)|*.dll"
cd1.ShowOpen
If cd1.FileName <> "" Then
frmTree.List1.Clear
frmTree.List2.Clear
XpB.Enabled = True
frmTrait.Show
frmTrait.Visible = True
End If
End Sub

Private Sub Decompiler_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInfo.Caption = "Decompilez un fichier exe compilé avec Visual Basic 5 ou 6. Vous pouvez generer des forms avec les controls et les sources. Vous pouvez extraire divers messages, strings,...  Il est possible de modifier certains parametres des controls et les sauver dans l'EXe même."
End Sub


Private Sub editer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInfo.Caption = "Editer, pour editer les fichier EXE en Hexadécimal, suivre les procédure exécutée en assembleur."
End Sub

Private Sub exporter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInfo.Caption = "Cliquez ici pour exporter les modifications faites."
End Sub

Private Sub Form_Load()
XpB.Enabled = False

  cLanguage.SetLanguageInForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmGenerator
Unload frmTree
Unload frmTrait
Unload frmGenerated
Unload frmSplash
End
End Sub

Private Sub Generator_Click()
frmGenerator.Show
End Sub

Private Sub Generator_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInfo.Caption = "Ici, vous pourrez generer les forms, avoir un aperçu direct"
End Sub

Private Sub XpB_Click()
frmTree.Show
End Sub

Private Sub XpB1_Click()
frmDownload.Show
End Sub

Private Sub XpB2_Click()
If frmExt.Visible = True Then frmExt.Visible = False Else frmExt.Visible = True
End Sub

Private Sub XpB2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInfo.Caption = "Vous pouvez extraire les resources de n'importe quel EXE (icons, bitmaps, avi, curseurs, strings, dialogs, menus,...)"
End Sub

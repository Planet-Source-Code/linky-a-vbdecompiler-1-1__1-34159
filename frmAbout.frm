VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C7A989&
   BorderStyle     =   0  'None
   Caption         =   "A Propos de VisualBasic Decompiler"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "frmAbout"
   ScaleHeight     =   4695
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   6225
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin VB.Image Image1 
         Height          =   825
         Left            =   0
         Picture         =   "frmAbout.frx":0000
         Top             =   0
         Width           =   2460
      End
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
         Left            =   4770
         TabIndex        =   3
         Top             =   30
         Width           =   1350
      End
   End
   Begin VBDecompiler.BTL BTL1 
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Fermer"
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
      MICON           =   "frmAbout.frx":69F6
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   3
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VBDecompiler.BTL BTL2 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Enregistrement"
      ENAB            =   0   'False
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
      MICON           =   "frmAbout.frx":6A12
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   1
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Verion :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   0
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BTL1_Click()
Unload Me
End Sub

Private Sub BTL2_Click()
'RIEN  POUR L'INSTANT
End Sub

Private Sub Form_Load()
PForm frmAbout, True
lblVersion.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
cLanguage.SetLanguageInForm Me
End Sub

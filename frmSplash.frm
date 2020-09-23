VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VisualBasic Decompiler"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   LinkTopic       =   "frmSplash"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   120
      Top             =   3000
   End
   Begin VB.Image Image1 
      Height          =   3690
      Left            =   0
      Picture         =   "frmSplash.frx":71E32
      Top             =   0
      Width           =   9465
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
Load frmMain
PForm frmSplash, True
cLanguage.SetLanguageInForm Me
End Sub

Private Sub Image1_Click()
frmMain.Show
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Timer1_Timer()
frmMain.Show
Timer1.Enabled = False
Unload Me
End Sub

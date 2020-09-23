VERSION 5.00
Begin VB.Form frmLangage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Choissisez un langage..."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Afficher les langages disponibles"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox lstLangPacks 
      Height          =   2985
      ItemData        =   "frmLangage.frx":0000
      Left            =   120
      List            =   "frmLangage.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLangage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MLang As String
Private Sub Command1_Click()
   If lstLangPacks.ListCount = 0 Or lstLangPacks.ListIndex = -1 Then Exit Sub
  cLanguage.LoadLanguagePack App.Path & "\" & lstLangPacks.List(lstLangPacks.ListIndex)
  MLang = lstLangPacks.List(lstLangPacks.ListIndex)
  cLanguage.SetLanguageInForm Me
    cLanguage.SetLanguageInForm frmMain
  Unload Me

  Load frmMain
  frmMain.Show
End Sub

Private Sub Command2_Click()
  lstLangPacks.Clear
    Dim sTmp As String, sTmpArray() As String, i As Integer
  sTmp = cLanguage.EnumLanguagePacks(App.Path & "\", "*.lpk")
  sTmpArray = Split(sTmp, "|")
  For i = 0 To UBound(sTmpArray)
    If sTmpArray(i) <> "" Then lstLangPacks.AddItem sTmpArray(i)
  Next
End Sub


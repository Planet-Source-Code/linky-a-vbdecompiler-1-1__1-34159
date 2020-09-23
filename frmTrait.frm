VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrait 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "frmTraitement"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   900
      Width           =   5175
      Begin MSComctlLib.ProgressBar pbar1 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Timer Timr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   360
   End
   Begin VBDecompiler.XpB cmdcancel 
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Annuler"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Traitement en cours... Veuillez patienter."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmTrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Timer1.Enabled = False
Unload Me
Annulert
End Sub


Private Sub Form_Load()
cLanguage.SetLanguageInForm Me
Me.Show
Me.Left = (Screen.Width / 2) - (frmTrait.Width / 2)
Timr.Enabled = True
PForm Me, True
End Sub

Private Sub Timer1_Timer()
frmTree.Show
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Timr_Timer()
On Error Resume Next
Timr.Enabled = False
pbar1.Value = 5
   Dim FreeF As Variant
    Dim BByte As Variant
    Dim FChr As String
    Dim CountDown As Variant
    Dim SpaceByte As String
    Dim CountDownAgain As Variant
    Dim SecondCount As String
    Dim LastCount As Variant
    FreeF = FreeFile
    Open frmMain.cd1.FileName For Binary Access Read As #FreeF
    BByte = 3200
    FChr$ = Chr(0) + Chr(255) + Chr(1)


    For CountDown = 1 To LOF(FreeF) Step BByte
        SpaceByte$ = Space(BByte)
        Get #FreeF, CountDown, SpaceByte$
NouveauEssai:


        If InStr(1, SpaceByte$, FChr$, 1) Then
            CountDownAgain = InStr(1, SpaceByte$, FChr$, 1)
            SecondCount$ = Mid(SpaceByte$, CountDownAgain + 6, 22)


            If InStr(1, SecondCount$, Chr(46) + Chr(102) + Chr(114) + Chr(109), 1) Then
                LastCount = InStr(SecondCount$, Chr(0))
                SecondCount$ = Mid(SecondCount$, 1, LastCount - 1)
                frmTree.lstForm.AddItem SecondCount$
            End If
            SpaceByte$ = Mid(SpaceByte$, CountDownAgain + 4)
            GoTo NouveauEssai
        End If
    Next CountDown
    Close #FreeF
    
pbar1.Value = 20


Dim a$, b$, c$, d$
Open frmMain.cd1.FileName For Binary As #1
a$ = String(4, " ")
For i = 1 To FileLen(frmMain.cd1.FileName)
Get #1, i, a$
If a$ = "VB5!" Then GoTo suite
Next i


 MsgBox "Le fichier n'est pas un fichier VB valide"
Exit Sub

suite:
pbar1.Value = 30
c$ = String(1, " ")
d$ = String(1, " ")
e$ = String(1, " ")
e$ = String(12, " ")
For j = 1 To FileLen(frmMain.cd1.FileName)
b$ = String(5, " ")
If j = 4736 Then
coucou = ""
End If


Get #1, j, b$
Get #1, j, f$


Get #1, j + 5, c$
On Error Resume Next
Get #1, j - 2, d$
Get #1, j + 2, e$


If b$ = "Label" And Asc(d$) = 6 Then 'Cas d'un Label normal
frmTree.List2.AddItem ("Label" & c$)
frmTree.List1.AddItem (j)

End If





If frmTree.Check3.Value = 1 Then
b$ = String(3, " ")
Get #1, j, b$
If b$ = "lbl" And Asc(d$) = 3 Or b$ = "lbl" And Asc(d$) = 4 Or b$ = "lbl" And Asc(d$) = 5 Then  'Cas d'un Label "LBLx"

nombre = ""
i = 0
c$ = String(1, " ")
Do

i = i + 1
Get #1, j + i + 2, c$
nombre = nombre + c$

Loop Until Asc(c$) = 0
pbar1.Value = 40
frmTree.List2.AddItem (b$ & nombre)
frmTree.List1.AddItem (j)
End If



pbar1.Value = 60


End If
If frmTree.Check2.Value = 1 Then
Get #1, j, c$
Get #1, j + 1, d$
verif = Hex(Asc(c$)) & Hex(Asc(d$))
pbar1.Value = 68
If verif = "30" And Asc(e$) > 65 And Asc(e$) < 91 Or verif = "30" And Asc(e$) > 97 And Asc(e$) < 123 Or verif = "30" And Asc(e$) > 47 And Asc(e$) < 58 Then
pbar1.Value = 76
abc = Hex(Asc(e$))
abc2 = Hex(Asc(c$))
i = 0
newItem = ""
Do
Get #1, j + 2 + i, e$
i = i + 1
newItem = newItem + e$
Loop Until Asc(e$) = 0
pbar1.Value = 84
Get #1, j + 3 + i, e$
frmTree.List2.AddItem (newItem)
frmTree.List1.AddItem (j)
pbar1.Value = 93
End If
End If

Next j

Close #1
pbar1.Value = 95
End Sub


VERSION 5.00
Begin VB.Form frmTree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decompilation Terminée"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9825
   Icon            =   "frmTree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame_frm 
      Caption         =   "Form - Propriété "
      Height          =   6855
      Left            =   3000
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.ListBox lstForm 
      BackColor       =   &H00C27E4B&
      ForeColor       =   &H00FFFFFF&
      Height          =   6180
      IntegralHeight  =   0   'False
      ItemData        =   "frmTree.frx":0442
      Left            =   0
      List            =   "frmTree.frx":0449
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdDependances 
      Caption         =   "Dependances"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   34
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forms"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   35
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Frames"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame frame_lab 
      Caption         =   "Label - Propriété"
      Height          =   6840
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C27E4B&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   6465
         TabIndex        =   31
         Top             =   3120
         Width           =   6495
         Begin VB.CommandButton Command8 
            BackColor       =   &H00C27E4B&
            Caption         =   "Aperçu"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton Command7 
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   11.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6180
            TabIndex        =   32
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Text            =   "FORM"
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3240
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   2040
         Width           =   3855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3345
         ScaleWidth      =   6465
         TabIndex        =   15
         Top             =   3360
         Width           =   6495
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   360
            Width           =   45
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visible = FALSE"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         ScaleHeight     =   225
         ScaleWidth      =   1425
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         ScaleHeight     =   225
         ScaleWidth      =   1425
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Verifier les eventuelle Labelles cachées!"
         Height          =   435
         Left            =   5040
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Checker des ""lblx"""
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Text            =   "Text6"
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "LEFT"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "TOP"
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "WIDTH"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "HEIGHT"
         Height          =   255
         Left            =   2640
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Caption"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Back color:"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Fore color:"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Tag"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPicBox 
      Caption         =   "PicBox"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdTBox 
      Caption         =   "TextBox"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdFrames 
      Caption         =   "Frames"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdtimer 
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Command"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdLabel 
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "frmTree.frx":0456
      Left            =   7560
      List            =   "frmTree.frx":0458
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstFile 
      BackColor       =   &H00C27E4B&
      ForeColor       =   &H00FFFFFF&
      Height          =   6180
      IntegralHeight  =   0   'False
      ItemData        =   "frmTree.frx":045A
      Left            =   0
      List            =   "frmTree.frx":045C
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00C27E4B&
      ForeColor       =   &H00FFFFFF&
      Height          =   6180
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Frame frame_dep 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dependances - informations"
      Height          =   6855
      Left            =   3000
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   5370
         TabIndex        =   41
         Text            =   "*.dll, *.ocx, *.exe"
         Top             =   180
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   39
         Top             =   6480
         Width           =   6420
      End
      Begin VBDecompiler.BTL cmdExtractDep 
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Extraire"
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
         BCOL            =   12640511
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmTree.frx":045E
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
      Begin VBDecompiler.BTL cmdSaveListDep 
         Height          =   375
         Left            =   1920
         TabIndex        =   42
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Sauver la liste"
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
         BCOL            =   12640511
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmTree.frx":047A
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
      Begin VB.Label lblExt 
         BackStyle       =   0  'Transparent
         Caption         =   "Extentions:"
         Height          =   255
         Left            =   4395
         TabIndex        =   44
         Top             =   255
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Fichier :"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C27E4B&
         Height          =   315
         Left            =   110
         Top             =   6470
         Width           =   6450
      End
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim sOldOpenPath As String
Dim sOldSavePath As String


Private Sub cmdDependances_Click()
lstFile.Visible = True
List2.Visible = False
frame_dep.Visible = True
frame_lab.Visible = False
frame_frm.Visible = False
lstForm.Visible = False
End Sub

Private Sub cmdExtractDep_Click()
Dim i As Integer
Dim X As Integer
Dim Z As Integer
Dim iFind As Long
Dim sExt As String
Dim iLen As Integer
Dim sFile As String
Dim iFree As Integer
Dim sFound As String
Dim sQuery As String
Dim bValid As Boolean
Dim iTerminator(1) As Long

    
    iFree = FreeFile
    Open frmMain.cd1.FileName For Binary Access Read As #iFree
        sFile = Space(LOF(iFree))
        Get #iFree, , sFile
    Close #iFree

    
    sFile = LCase(sFile)
    txtExt = LCase(txtExt)
    
    
        Do
            DoEvents
            iFind = InStr(iFind + 1, sFile, ".")
            If iFind = 0 Then Exit Do
            iTerminator(0) = InStrRev(sFile, Chr(0), iFind)
           
            iTerminator(1) = iFind + 4
            If iTerminator(0) And Mid$(sFile, _
            iTerminator(1), 1) = Chr$(0) Then
                If iTerminator(1) - iTerminator(0) - 1 < 20 _
                And iTerminator(1) - iTerminator(0) - 1 > 5 Then
                                                                  
                    bValid = True
                    
                    sFound = Mid$(sFile, iTerminator(0) + 1, _
                    iTerminator(1) - iTerminator(0) - 1)
                    
                                        
                    If bValid Then
                        lstFile.AddItem sFound
                    End If
                    
                    Debug.Print sFound
                    
                    End If
            End If
        Loop
    
    
    If lstFile.ListCount > 0 Then lstFile.ListIndex = 0
        
    Exit Sub
ExitSearch:
    MsgBox "Erreur !!!!!!!!!!!!!!"
End Sub

Private Sub cmdLabel_Click()
MsgBox "Hexman : Suite a plusieurs modifications, j'ai eu qq pb avec les label, command. Donc tu ne pourras qu'avoir que certaines fonction"
frame_dep.Visible = False
frame_lab.Visible = True
lstFile.Visible = False
List2.Visible = True
lstForm.Visible = False
End Sub

Private Sub Command2_Click()
lstForm.Visible = True
lstFile.Visible = False
List2.Visible = False
frame_frm.Visible = True
frame_dep.Visible = False
frame_lab.Visible = False
End Sub

Private Sub Command7_Click()
If Picture1.Visible = True Then Picture1.Visible = False Else Picture1.Visible = True
End Sub

Private Sub Command8_Click()
Label7.Caption = Text5.text
Label7.Left = Text1.text
Label7.Top = Text2.text
Label7.Width = Text3.text
Label7.Height = Text4.text
Label7.AutoSize = True

End Sub


Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
        Clipboard.Clear
        Clipboard.SetText lstFile.List(lstFile.ListIndex)

    End If
End Sub

Private Sub List2_Click()
Dim d$
Dim k As Integer
Dim l As Inet
Dim a As String
Dim b As String
Dim g$
Dim b1$
Dim b2$
Dim b3$
Dim b4$

On Error Resume Next
Open frmMain.cd1.FileName For Binary As #2
d$ = String(1, " ")
k = List2.ListIndex
l = 0

a = List1.List(k + 1)
b = List1.List(k)

Do

l = l + 1
Get #2, List1.List(k) + l + Len(List2.List(k)) + 4, d$
If Asc(d$) <> 0 Then 'And Asc(d$) <> 4 And Asc(d$) <> 5 Then

Else: GoTo l2
End If
'Next l
Loop Until Asc(d$) = 0
l2:


Text5.Tag = List1.List(k) + Len(List2.List(k)) + 4

'====================================Backcolor=======================
Do
Get #2, List1.List(k) + l + 10, d$
l = l + 1
Loop Until Asc(d$) = 3 Or Asc(d$) = 5 Or Asc(d$) = 4

If Asc(d$) = 3 Then
g$ = String(4, " ")
Get #2, List1.List(k) + l + Len(List2.List(k)) + 4, g$
b1 = Hex(Asc(Mid$(g$, 4, 1)))
b2 = Hex(Asc(Mid$(g$, 3, 1)))
b3 = Hex(Asc(Mid$(g$, 2, 1)))
b4 = Hex(Asc(Mid$(g$, 1, 1)))

If Len(b1) = 1 Then b1 = "0" & b1
If Len(b2) = 1 Then b2 = "0" & b2
If Len(b3) = 1 Then b3 = "0" & b3
If Len(b4) = 1 Then b4 = "0" & b4
If b1 = "00" Then
Picture2.BackColor = RGB(Asc(Mid$(g$, 1, 1)), Asc(Mid$(g$, 2, 1)), Asc(Mid$(g$, 3, 1)))
Picture2.Tag = "+" & List1.List(k) + l + Len(List2.List(k)) + 4


End If
If b1 <> "00" Then
Dim newcolor As String
newcolor = "&H" & b1 & b2 & b3 & b4 & "&"
Picture2.BackColor = newcolor
Picture2.Tag = "+" & List1.List(k) + l + Len(List2.List(k)) + 4
End If
Else: Picture2.BackColor = RGB(192, 192, 192)

End If

'=================================================FIn backcolor===================

'======================================Forecolor================================
Do
'Get #2, List1.List(k) + l + 10, d$
Get #2, List1.List(k) + l, d$
l = l + 1
Loop Until Asc(d$) = 4 Or Asc(d$) = 5
Dim hnnhn As String
hnnhn = Asc(d$)



If Asc(d$) = 4 Then
g$ = String(4, " ")
'Get #2, List1.List(k) + l + 10, g$
Get #2, List1.List(k) + l, g$
b1 = Hex(Asc(Mid$(g$, 4, 1)))
b2 = Hex(Asc(Mid$(g$, 3, 1)))
b3 = Hex(Asc(Mid$(g$, 2, 1)))
b4 = Hex(Asc(Mid$(g$, 1, 1)))

If Len(b1) = 1 Then b1 = "0" & b1
If Len(b2) = 1 Then b2 = "0" & b2
If Len(b3) = 1 Then b3 = "0" & b3
If Len(b4) = 1 Then b4 = "0" & b4
If b1 = "00" Then
Picture3.BackColor = RGB(Asc(Mid$(g$, 1, 1)), Asc(Mid$(g$, 2, 1)), Asc(Mid$(g$, 3, 1)))
Picture3.Tag = "+" & List1.List(k) + l ' + 10


End If
If b1 <> "00" Then
GoTo suitefore
'On Error GoTo suitefore
Dim newcolor3 As String
newcolor3 = "&H" & b1 & b2 & b3 & b4 & "&"
Me.Picture3.BackColor = newcolor3
Me.Picture3.Tag = "+" & List1.List(k) + l + Len(List2.List(k)) + 4
End If
suitefore:
Else: Picture3.BackColor = RGB(192, 192, 192)


End If
Dim e$
Dim c As String
Dim deci As String
Dim H$

'"=============================Fin forecolor========================

e$ = String(1, " ")
c = List1.List(List2.ListIndex) 'Len(Text5.Text) + 11 + List1.List(List2.ListIndex)

Do
Get #2, c, d$
Get #2, c - 1, e$
deci = Asc(d$)
c = c + 1


If EOF(2) Then
MsgBox "Ceci n'est pas un Label valide!"
List1.RemoveItem (List2.ListIndex)
List2.RemoveItem (List2.ListIndex)
Close #2
Exit Sub
End If


Dim f$

Loop Until deci = 5 'And Asc(e$) <> 1
e$ = String(1, " ")
f$ = String(1, " ")

Get #2, c + 1, f$
Get #2, c, e$

GoTo suite
erreur:
MsgBox " Ceci n'est pas un Controle Valide!"
List1.RemoveItem (List2.ListIndex)
List2.RemoveItem (List2.ListIndex)
Close #2
Exit Sub
suite:




On Error GoTo erreur
Text1.text = Asc(f$) * 256 + Asc(e$)
Text1.Tag = c
Get #2, c + 3, f$
Get #2, c + 2, e$
Text2.text = Asc(f$) * 256 + Asc(e$)
Text2.Tag = c + 2
Get #2, c + 5, f$
Get #2, c + 4, e$
Text3.text = Asc(f$) * 256 + Asc(e$)
Text3.Tag = c + 4
Get #2, c + 7, f$
Get #2, c + 6, e$
Text4.text = Asc(f$) * 256 + Asc(e$)
Text4.Tag = c + 6



Get #2, c + 8, e$
H = Hex(Asc(e$))
If H = "A" Then             'Visible =False
'Check1.Visible = True
Check1.Enabled = True
Check1.Value = 1
Else
Check1.Enabled = False
Check1.Value = 0
If H = "B" Then Check1.Enabled = True
End If


Dim i As Integer

g$ = String(1, " ")
Do
i = i + 1
Get #2, c + 7 + i, g$
H = Hex(Asc(g$))
Loop Until H = "12"

If H = "12" Then                'TAG

g$ = String(2, " ")

Dim indic As String
indic = ""
'Do
'i = i + 1
Get #2, c + 7 + i + 3, g$
'loop until
Dim text As String
text = Hex(Asc(g$))
If Hex(Asc(g$)) = "1D" Then
Dim u$
Dim texttag As String
    u$ = String(1, " ")
    texttag = ""

Get #2, (c + 7 + i + 4), u$
Dim ncaracteres As String
ncaracteres = Asc(u$)
Dim debut As String
debut = c + 7 + i + 6
Dim j As Integer
For j = 0 To ncaracteres * 2 - 1 Step 2
Get #2, debut + j, u$
texttag = texttag + u$

Next j

Text6.text = texttag
End If

Else: Text6.text = ""
End If

'Next k
Close #2


End Sub

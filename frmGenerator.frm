VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGenerator 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Generateur de Forms"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8550
   LinkTopic       =   "Generator"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Generer"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sauvegarder"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox txtForm 
      Height          =   6465
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   11404
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmGenerator.frx":0000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C27E4B&
      Height          =   6495
      Left            =   120
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Composant As Boolean
Dim FormName As String

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
Dim Object As String
txtForm.text = "VERSION 5.00"
If Composant = True Then
txtForm.text = txtForm.text & vbCrLf & Object
End If
txtForm = txtForm.text & vbCrLf & "Begin VB.MDIForm " & FormName
txtForm.text = txtForm.text & vbCrLf & ""
   
End Sub

Private Sub Form_Load()
cLanguage.SetLanguageInForm Me
FormName = "FrmMain"
Composant = True
PForm frmGenerator, True
Object = """Object = " & """{C9FF5F4F-78AB-4799-A8B8-EA9191E3BBA7}#1.0#0" & "cPopMenu.ocx"
End Sub


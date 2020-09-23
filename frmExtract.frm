VERSION 5.00
Begin VB.Form frmExtract 
   Caption         =   "Resource Extractor"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cLanguage.SetLanguageInForm Me
End Sub

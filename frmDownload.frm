VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " - Mettre à jour VisualBasic Decompiler"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VBDecompiler.XpB cmdDownload 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Telecharger"
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
   Begin VB.Frame Frame3 
      Caption         =   "Statut des fichiers..."
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   4815
      Begin VB.Label lblRemaining 
         Caption         =   "1 fichier restant"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         Caption         =   "0 fichiers téléchargés"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informations..."
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   810
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label StatusLabel 
         Caption         =   "-"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label EstimatedTimeLeft 
         Caption         =   "Temps restant :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label SourceLabel 
         Caption         =   "-"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   4530
      End
      Begin VB.Label TimeLabel 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   1125
         Width           =   3045
      End
      Begin VB.Label DownloadTo 
         Caption         =   "Téléchargé de :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label ToLabel 
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   3075
      End
      Begin VB.Label TransferRate 
         Caption         =   "Vitesse :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label RateLabel 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   3045
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VBDecompiler.XpB cmdCancel 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Annuler..."
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
   Begin VBDecompiler.XpB cmdQuit 
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Quitter"
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
   Begin VBDecompiler.XpB cmdClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Fermer"
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
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CancelSearch As Boolean

Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean
Const CHUNK_SIZE As Long = 1024
Const ROLLBACK As Long = 4096
                                
                                
                                
Dim bData() As Byte
Dim blnResume As Boolean
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim lngX
Dim sglLastTime As Single
Dim sglRate As Single
Dim sglTime As Single
Dim strFile As String
Dim strHeader As String
Dim strHost As String
On Local Error GoTo InternetErrorHandler
CancelSearch = False
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty


StartDownload:

If blnResume Then
    StatusLabel = "Reprise du téléchargement..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Recherche des informations du fichier..."
End If
DoEvents
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    While .StillExecuting
        DoEvents
            If CancelSearch Then GoTo ExitDownload
    Wend

    StatusLabel = "Enregistrement :"
    SourceLabel = FitText(SourceLabel, strHost & " de " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)

    strHeader = .GetHeader
End With
Select Case Mid(strHeader, 10, 3)
    Case "200"
            If blnResume Then
        Kill strDestination
        If MsgBox("Impossible de reprendre le telechargement du serveur." & _
                 vbCr & vbCr & _
                "Voulez-vous continuer tout de même ?", _
                vbExclamation + vbYesNo, _
                "Reprise impossible") = vbYes Then
                blnResume = False
        Else
            
                    CancelSearch = True
                    GoTo ExitDownload
                End If
        End If
            
    Case "206"
    Case "204"
        GoTo ExitDownload
    Case "401"
        GoTo ExitDownload
    Case "404"
        MsgBox "Le fichier n'as pas été trouvé"
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf
        MsgBox "Impossible d'établir une connection"
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "Erreur"
        CancelSearch = True
        GoTo ExitDownload
End Select

If blnResume = False Then
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        MsgBox "Espace insuffisant sur le disque"
        GoTo ExitDownload
    End If
End If

With ProgressBar
    .Value = 0
    .Max = lngFileLength
End With
DoEvents
If blnResume = False Then lngBytesReceived = 0
On Local Error GoTo FileErrorHandler
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If
intFile = FreeFile
Open strDestination For Binary Access Write As #intFile
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do

    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar.Value = lngBytesReceived
    Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " de " & strFile & " téléchargés"
Loop While UBound(bData, 1) > 0
Close #intFile

ExitDownload:
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Téléchargement términé"
    DownloadFile = True
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        If CancelSearch = False Then
            If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to resume downloading the file." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to cancel downloading the file.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
        
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If

    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:

Inet1.Cancel
Exit Function

InternetErrorHandler:
    
    If Err.Number = 9 Then Resume Next
    MsgBox "Erreur: " & Err.Description & " à été provoquée.", _
           vbCritical, _
           "Erreur"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "Cannot write file to disk." & _
           vbCr & vbCr & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Erreur"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function

Private Sub cmdCancel_Click()
StatusLabel = "Annulation en cours..."
CancelSearch = True
End Sub

Private Sub cmdDownload_Click()
On Error Resume Next
Dim OldVersion As Integer

cmdQuit.Visible = False
cmdCancel.Enabled = True
cmdDownload.Enabled = False

StatusLabel.Caption = "..."
FileCopy App.Path & "\VBDecompiler.exe", App.Path & "\VBDecompiler2.exe"
OldVersion = DLLVersion
DownloadFile "http://www.dordevic.yucom.be/boris-site/vb/VBDecompiler.exe", App.Path & "\VBDecompiler.exe"
lblRemaining.Caption = "1 fichier restant"
lblTotal.Caption = "1 fichier téléchargé"

DownloadFile "http://www.dordevic.yucom.be/boris-site/vb/vbd.ocx", App.Path & "\VBD.ocx"
lblRemaining.Caption = "O fichiers restants"
lblTotal.Caption = "2 fichiers téléchargés"

StatusLabel.Caption = "Ajout de la nouvelle version..."

finish:
cmdCancel.Enabled = False
Unload MeApp.Path & "\VBDecompiler.exe"
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
inetDownload.Cancel
Unload Me
End Sub

Private Sub cmdQuit_Click()
On Error Resume Next
inetDownload.Cancel
End Sub

Private Sub Form_Load()
cLanguage.SetLanguageInForm Me
cmdClose.Caption = "Fermer"
cmdDownload.Enabled = True
PForm frmDownload, True
End Sub

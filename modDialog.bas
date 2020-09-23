Attribute VB_Name = "modDialog"

Option Explicit

Public Const OFN_SHOWHELP = &H10
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ALLOWMULTISELECT = &H200

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public dFileName As String
Public dFileTitle As String

Public Function Dialog(initFilter As String, initDialogTitle As String, currentForm As Form, Optional sExtention As String, Optional initDir As String = "C:\", Optional initOpenORSave As Boolean = False)
    Dim FileDialog As OPENFILENAME
    Dim FileName As String
    Dim Checkit As Boolean
    With FileDialog
        .lStructSize = Len(FileDialog)
        .hwndOwner = currentForm.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Replace(initFilter, "|", Chr$(0))
        
        .lpstrFile = FileName & Space$(254 - Len(FileName)) ' Text in the filename box
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = initDir ' The starting directory
        .lpstrTitle = initDialogTitle
        .flags = 0
        
        If initOpenORSave Then
            Checkit = GetOpenFileName(FileDialog) ' Open Dialog
        Else
            Checkit = GetSaveFileName(FileDialog) ' Save Dialog
        End If
        
        If Checkit Then
            dFileName = Trim$(.lpstrFile) ' File Name
            dFileTitle = Trim$(.lpstrFileTitle) ' FileTitle
            sExtention = LCase(sExtention)
            dFileName = LCase(Left(dFileName, InStr(dFileName, Chr(0)) - 1))
            dFileTitle = LCase(Left(dFileTitle, InStr(dFileTitle, Chr(0)) - 1))
            If initOpenORSave = False Then
                If Not Mid(Right(dFileName, 4), 1, 1) Like "." Then
                    dFileName = dFileName & sExtention
                    dFileTitle = dFileTitle & sExtention
                End If
            End If
        Else
            dFileName = ""
            dFileTitle = ""
        End If
    End With
End Function


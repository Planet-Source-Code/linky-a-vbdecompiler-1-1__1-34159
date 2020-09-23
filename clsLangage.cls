VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLangage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sCurrentFile As String
Private Type ObjectProperties
  Name As String
  Caption As String
  ToolTip As String
End Type
Private Type FormProperties
  Name As String
  Caption As String
  ObjectCount As Integer
  ObjProp() As ObjectProperties
End Type

Private FormProp() As FormProperties
Private iFormCount As Integer
Sub LoadLanguagePack(sFile As String)
  Dim sLine As String, iPos As Integer, sTmp As String
  Dim sFormName As String, sTmp2 As String
  Dim bFormFound As Boolean
  sCurrentFile = sFile
  iFormCount = 0
  Open sFile For Input As #1
    Do
      Input #1, sLine
      If Left$(sLine, 1) = ";" Or sLine = "" Then GoTo Jump

      If Left$(sLine, 1) = "[" And Right$(sLine, 5) = ".End]" Then
        bFormFound = False: GoTo Jump
      End If
      If Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
        bFormFound = True
        sFormName = Mid$(sLine, 2, Len(sLine) - 2)
        iFormCount = iFormCount + 1
        ReDim Preserve FormProp(iFormCount)
        FormProp(iFormCount).Name = sFormName: GoTo Jump
      End If
      
      If Left$(sLine, 7) = "Caption" Then
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
        FormProp(iFormCount).Caption = Left$(sTmp, Len(sTmp) - 1)
        GoTo Jump
      End If
            iPos = InStr(sLine, ".Caption")
      
      If iPos > 0 And bFormFound Then
        FormProp(iFormCount).ObjectCount = FormProp(iFormCount).ObjectCount + 1
        ReDim Preserve FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount)
        sTmp = Left$(sLine, iPos - 1)
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Name = sTmp
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
      
        If Right$(sTmp, 1) <> Chr$(34) Then
          Do While Right$(sTmp, 1) <> Chr$(34)
            Input #1, sTmp2
            sTmp = sTmp & ", " & sTmp2
          Loop
        End If
        If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
      
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Caption = sTmp
        GoTo Jump
      End If
            iPos = InStr(sLine, ".ToolTip")
      If iPos > 0 And bFormFound Then
        sTmp = Left$(sLine, iPos - 1)
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).Name = sTmp
        sTmp = Mid$(sLine, InStr(sLine, "=") + 2)
      
        If Right$(sTmp, 1) <> Chr$(34) Then
          Do While Right$(sTmp, 1) <> Chr$(34)
            Input #1, sTmp2
            sTmp = sTmp & ", " & sTmp2
          Loop
        End If
        If Right$(sTmp, 1) = Chr$(34) Then sTmp = Left$(sTmp, Len(sTmp) - 1)
      
        FormProp(iFormCount).ObjProp(FormProp(iFormCount).ObjectCount).ToolTip = sTmp
        GoTo Jump
      End If
      
Jump:


    Loop Until EOF(1)

  Close #1

End Sub


Sub SetLanguageInForm(frmForm As Form)
  
  On Local Error Resume Next
  Dim i As Integer, j As Integer
  Dim iForm As Integer
  

  For i = 1 To iFormCount
    If FormProp(i).Name = frmForm.Name Then
      iForm = i
      Exit For
    End If
  Next i

  
  frmForm.Caption = FormProp(iForm).Caption
    For j = 1 To FormProp(iForm).ObjectCount
    frmForm.Controls(FormProp(iForm).ObjProp(j).Name).Caption = FormProp(iForm).ObjProp(j).Caption
    frmForm.Controls(FormProp(iForm).ObjProp(j).Name).ToolTipText = FormProp(iForm).ObjProp(j).ToolTip
  Next j
  
End Sub

Function EnumLanguagePacks(sFolder As String, sExtension As String) As String

  If Not DirExists(sFolder) Then
    MsgBox sFolder & " doesn't exist.", vbCritical
    Exit Function
  End If
  
  Dim sTmp As String

  If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
  sTmp = Dir$(sFolder & sExtension)
  If sTmp <> "" Then
    EnumLanguagePacks = sTmp
    sTmp = Dir$
    While Len(sTmp) > 0
      EnumLanguagePacks = EnumLanguagePacks & "|" & sTmp
      DoEvents
      sTmp = Dir$
    Wend
  End If

End Function


Private Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    If Right$(strDirName, 1) <> "\" Then strDirName = strDirName & "\"
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function


Sub CarregaPacotedeLinguagem(sArquivo As String)
  
  LoadLanguagePack sArquivo
  
End Sub

Sub SetaLinguagemnoForm(frmForm As Form)

  SetLanguageInForm frmForm

End Sub

Function EnumeraPacotesdeLinguagem(sDiretorio As String, sExtensao As String) As String
  
  EnumLanguagePacks sDiretorio, sExtensao
  
End Function


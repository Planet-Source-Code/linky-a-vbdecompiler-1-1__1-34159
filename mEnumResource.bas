Attribute VB_Name = "mEnumResource"
Private Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNamesById Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Dim tv As TreeView, nd As Node

Function ResTypesCallBack(ByVal hMod As Long, ByVal ResType As Long, ByVal lParam As Long) As Long
    Dim nd As Node
    If (ResType And &HFFFF0000) = 0 Then
        Set nd = tv.Nodes.Add(tv.Nodes.Item(1), tvwChild, "#" & CStr(ResType), ResTypeName(ResType), 2, 3)
        tv.Nodes.Add nd, tvwChild, , "Dummy"
    Else
        Set nd = tv.Nodes.Add(tv.Nodes.Item(1), tvwChild, , StrFromPtrA(ResType), 2, 3)
        tv.Nodes.Add nd, tvwChild, , "Dummy"
    End If
    Set nd = Nothing
    ResTypesCallBack = True
End Function

Function ResNamesCallBack(ByVal hMod As Long, ByVal ResType As Long, ByVal ResId As Long, ByVal lParam As Long) As Long
    If (ResId And &HFFFF0000) <> 0 Then
       tv.Nodes.Add nd, tvwChild, , StrFromPtrA(ResId), 4, 4
    Else
       tv.Nodes.Add nd, tvwChild, , CStr(ResId), 4, 4
    End If
    ResNamesCallBack = True
End Function

Public Function FillResTypes(ByVal tvw As TreeView, ByVal sFileName As String, ByVal sLibName As String)
   Dim ret As Long
   Set tv = tvw
   tv.Nodes.Clear
   tv.Nodes.Add , , sFileName, sLibName, 1, 1
   Call InitResource(sFileName)
   If hModule Then ret = EnumResourceTypes(hModule, AddressOf ResTypesCallBack, 0)
   tv.Refresh
   tv.Nodes.Item(1).Expanded = True
   Set tv = Nothing
End Function

Public Function FillResNames(ByVal tvw As TreeView, ByVal nod As Node)
   Dim ret As Long
   Set tv = tvw
   Set nd = nod
   If nd.key = "" Then
      ret = EnumResourceNames(hModule, nd.Text, AddressOf ResNamesCallBack, 0)
   Else
      ret = EnumResourceNamesById(hModule, CLng(Mid(nd.key, 2)), AddressOf ResNamesCallBack, 0)
   End If
   Set tv = Nothing
   Set nd = Nothing
End Function


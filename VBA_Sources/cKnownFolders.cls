VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKnownFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public KnownFolderItem As cKnownFolderItem

Private Sub Class_Initialize()
    Set KnownFolderItem = New cKnownFolderItem
    Set FInFolderColl = New Collection
End Sub                                            ' cKnownFolders.Class_Initialize

'---------------------------------------------------------------------------------------
' Method : Function InRuleFolderColl
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function InRuleFolderColl(inFolderName As String) As Boolean
Dim zErr As cErr
Const zKey As String = "cKnownFolders.InRuleFolderColl"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim isOK As Long
    For i = 1 To FInFolderColl.Count
        Set KnownFolderItem = FInFolderColl.Item(i)
        If StrComp(KnownFolderItem.FInFolderstring, inFolderName, vbTextCompare) = 0 Then
            isOK = KnownFolderItem.FolderOK        ' check Property
            If isOK = -1 Then
                InRuleFolderColl = True            ' found existing entry
                ' Else ' name exists, but Folder not
            End If
            GoTo ProcReturn
        End If
    Next i
    KnownFolderItem.FInFolderstring = inFolderName
    FInFolderColl.Add KnownFolderItem, inFolderName
    InRuleFolderColl = False                       ' it is a new entry

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cKnownFolders.InRuleFolderColl


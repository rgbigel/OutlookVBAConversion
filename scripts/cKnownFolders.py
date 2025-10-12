# Converted from cKnownFolders.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cKnownFolders"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public KnownFolderItem As cKnownFolderItem

# Private Sub Class_Initialize()
# Set KnownFolderItem = New cKnownFolderItem
# Set FInFolderColl = New Collection

# '---------------------------------------------------------------------------------------
# ' Method : Function InRuleFolderColl
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def inrulefoldercoll():
    # Dim zErr As cErr
    # Const zKey As String = "cKnownFolders.InRuleFolderColl"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim isOK As Long
    # Set KnownFolderItem = FInFolderColl.Item(i)
    if StrComp(KnownFolderItem.FInFolderstring, inFolderName, vbTextCompare) = 0 Then:
    # isOK = KnownFolderItem.FolderOK        ' check Property
    if isOK = -1 Then:
    # InRuleFolderColl = True            ' found existing entry
    # ' Else ' name exists, but Folder not
    # GoTo ProcReturn
    # KnownFolderItem.FInFolderstring = inFolderName
    # FInFolderColl.Add KnownFolderItem, inFolderName
    # InRuleFolderColl = False                       ' it is a new entry

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKnownFolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FInFolderstring As String                   ' default value of class <self> ****
Attribute FInFolderstring.VB_VarUserMemId = 0
Attribute FInFolderstring.VB_VarDescription = "FolderName"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute FInFolderstring.VB_VarUserMemId = 0
' Attribute FInFolderstring.VB_VarDescription = "FolderName"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Private Type WhatIsFolderType
    FInNormalFolder As Long                        ' 0 = false, -1 = not a normal Folder
    FInSearchFolder As Long                        ' -1 = it is a normal Folder
    ' 0 = undefined so far
' 1 = definitely is search Folder
End Type                                           ' WhatIsFolderType

Private FInFolderType As WhatIsFolderType

Public FInFolderWarning As Long                    ' -1 warning needed, (0 = all OK)
' -2 =  with warning already issued
Private FInFolderOK As Long                        ' -1 = True, Folder does exist
' 0 = False, Folder does not exist or we dont know
' 1 Folder definitely does not exist

Private Sub Class_Initialize()
    FInFolderWarning = 0                           ' -1 warning needed, (0 = all OK)
    ' -2 =  with warning already issued
    FInFolderstring = vbNullString
    With FInFolderType
        .FInNormalFolder = -1                      ' 0 = false, -1 = not a normal Folder
        .FInSearchFolder = 0                       ' -1 = it is a normal Folder
        ' 0 = undefined so far
        ' 1 = definitely is search Folder
    End With                                       ' FInFolderType
    
    FInFolderOK = 0                                ' -1 = True, Folder does exist
    ' 0 = False, Folder does not exist or we dont know
    ' 1 Folder definitely does not exist
End Sub                                            ' cKnownFolderItem.Class_Initialize

Public Property Let FolderIsNormal(normalFolderState As Long)
    
    If FInFolderType.FInSearchFolder >= 0 Then
        FInFolderType.FInNormalFolder = 0
    ElseIf normalFolderState = -1 Then
        FInFolderType.FInNormalFolder = normalFolderState
        If normalFolderState = -1 And FInFolderWarning <> -2 Then
            FInFolderWarning = -1                  ' -1 warning needed,
        End If
    Else
        DoVerify False, " no such state"
    End If
    If (FInFolderType.FInNormalFolder = 0 And FInFolderType.FInSearchFolder = 0) Then
        FInFolderOK = 0
    End If

End Property                                       ' cKnownFolderItem.FolderIsNormal Let

Public Property Get FolderIsNormal() As Long
    
    FolderIsNormal = FInFolderType.FInNormalFolder

End Property                                       ' cKnownFolderItem.FolderIsNormal Get

Public Property Let FolderIsSearch(searchFolderState As Long)
    
    If searchFolderState = 1 Then
        FInFolderType.FInSearchFolder = 0
    Else
        FInFolderType.FInSearchFolder = searchFolderState
    End If

End Property                                       ' cKnownFolderItem.FolderIsSearch Let

Public Property Get FolderIsSearch() As Long
   
    FolderIsSearch = FInFolderType.FInSearchFolder

End Property                                       ' cKnownFolderItem.FolderIsSearch Get

Public Function FolderOK() As Long
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cKnownFolderItem.FolderOK"
    Call DoCall(zKey, tFunction, eQzMode)

    If FInFolderOK = 0 Then                        ' we know nothing, so find out
        If (FInFolderWarning < -1 _
         Or FInFolderType.FInNormalFolder <> 0 _
         Or FInFolderType.FInSearchFolder <> 0) Then
            DoVerify False, " state error"
        Else
            FInFolderOK = 1                        ' definitely not there
        End If
    Else
        FInFolderOK = 1                            ' definitely not there
        If FInFolderWarning <> -2 Then
            FInFolderWarning = -1
        End If
    End If
    If FInFolderOK = -1 _
    And FInFolderType.FInSearchFolder _
    And FInFolderType.FInNormalFolder Then
        DoVerify False, " state error"
    End If
    FolderOK = FInFolderOK

zExit:
    Call DoExit(zKey)

End Function                                       ' cKnownFolderItem.FolderOK



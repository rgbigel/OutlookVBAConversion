VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Key As String                               ' default value of class <self> ****
Attribute Key.VB_VarUserMemId = 0
Attribute Key.VB_VarDescription = "Store DisplayName = Key of Account"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute Key.VB_VarUserMemId = 0
' Attribute Key.VB_VarDescription = "Store DisplayName = Key of Account"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************
Public aAccount As Account
Public aAcStore As Store
Public aAcType As eAccountType                     ' see AccountTypeNames(me.aType)

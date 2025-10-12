# Converted from cNumbItem.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cNumbItem"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public NuIndex As Long
# Public Key As String                               ' default value of class <self> ****
# Public Alias As String
# Public Subfields As String

# Public ValueOfItem As Object

# '---------------------------------------------------------------------------------------
# ' Method : Function Exists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def exists():

    # Const zKey As String = "cNumbItem.Exists"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim aitem As Object

    for aitem in acoll:
    if IsMissing(aAlias) Then:
    # keyCheck:
    if aitem.Key = sKey Then:
    # Set Exists = aitem
    # GoTo FuncExit
    else:
    if aitem.Alias = aAlias Then:
    if LenB(sKey) > 0 Then:
    # GoTo keyCheck
    # Set Exists = aitem
    # GoTo FuncExit

    # FuncExit:
    # Set aitem = Nothing

    # zExit:
    # Call DoExit(zKey)


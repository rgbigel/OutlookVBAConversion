# Converted from cFilterCriterium.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cFilterCriterium"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public CritPropName As String
# Attribute CritPropName.VB_VarUserMemId = 0
# Attribute CritPropName.VB_VarDescription = "Display Criteria Name"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute CritPropName.VB_VarUserMemId = 0
# ' Attribute CritPropName.VB_VarDescription = "Display Criteria Name"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public CritIndex As Long
# Public CritType As Long
# Public PropertyIdent As String

# Public Operator As String
# Public Comparator As String
# Public ValueSeperator As String
# Public oBracket As String
# Public cBracket As String
# Public BracketOpenCount As Long
# Public adFormattedValue As String
# Public ValueIsTimeType As Boolean
# Public AttrRawValue As Variant

# Property Get Value() As String                     ' Default Property of Class <self>    ****
# Value = CritPropName
# End Property                                       ' cFilterCriterium.Value Get

# '---------------------------------------------------------------------------------------
# ' Method : Sub addTo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub addTo(PropColl As Collection)

# Const zKey As String = "cFilterCriterium.addTo"
# Call DoCall(zKey, tSub, eQzMode)

if PropColl Is Nothing Then:
# Set PropColl = New Collection
# Set SQLpropC = PropColl
# SQLpropC.Add Me
# Me.CritIndex = SQLpropC.Count

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function Clone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function Clone(TargetType As Long)

# Set Clone = New cFilterCriterium
# With Clone
# .CritType = TargetType                     ' not copied!
# ' cut off first and last chars [ ]
# .PropertyIdent = Mid(PropertyIdent, 2, Len(PropertyIdent) - 2)
# .CritIndex = CritIndex
# .Operator = Operator
# .Comparator = Comparator
# .ValueSeperator = ValueSeperator
# .oBracket = oBracket
# .cBracket = cBracket
# .BracketOpenCount = BracketOpenCount
# .adFormattedValue = adFormattedValue
# .ValueIsTimeType = ValueIsTimeType
# .AttrRawValue = AttrRawValue
# ' now we adjust some
# .CritPropName = Clone.ModifyCriteria(CritPropName, TargetType)
# End With                                       ' Clone


# '---------------------------------------------------------------------------------------
# ' Method : Function ModifyCriteria
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def modifycriteria():

    # Const zKey As String = "cFilterCriterium.ModifyCriteria"
    # Call DoCall(zKey, tFunction, eQzMode)

    if TargetType = 1 Then:
    # ModifyCriteria = PropName
    elif TargetType = 2 Then:
    # Operator = UCase(Operator)
    # Me.oBracket = vbNullString
    # Me.cBracket = vbNullString
    if ValueIsTimeType Then:
    # adFormattedValue = Replace(adFormattedValue, ValueSeperator, vbNullString)
    # CritPropName = PropName
    match LCase(CritPropName):
        case "subject":
    # PropName = "betreff"
        case "sendername":
    # PropName = "von"
        case "senton":
    # PropName = "gesendet"
        case "received":
    # PropName = "erhalten"
        case _:
    # DoVerify False, " not implemented yet"
    # PropName = vbNullString            ' mark this as ignore
    else:
    # DoVerify False, " not a valid TargetType"

    # zExit:
    # Call DoExit(zKey)


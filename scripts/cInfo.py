# Converted from cInfo.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cInfo"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# ' this describes most significant Object vars, inspecting Variants/Objects of any type

# Public iAssignmentMode As Long                     ' 0/-99 - no value to assign/not initialized
# ' 1 - normal Assign
# ' 2 - using set
# Public iClass As Long                              ' Class of original ItemProperty/Variable/Value to check, -99 if none
# Public iType As VbVarType                          ' Type of original ItemProperty/Variable/Value
# Public iTypeName As String                         ' TypeName of original
# Public iScalarType As VbVarType                    ' Type corresponding to TypeName
# Public iIsArray As Boolean                         ' value in ivalue is Array with Count
# Public iArraySize As Long                          ' used by getInfo and TypeDecode - Procs
# Public iValue As Variant                           ' This scalar, variant, Object
# Public DecodedStringValue As String                ' if iValue can be converted to String, this is it
# Public DecodeMessage As String                     ' if something must be said to Decoding...

# Public iDepth As Long                              ' how many times we successfully set iValue = ivalue.Value
# Public iUp As cInfo
# Public iDown As cInfo

# '---------------------------------------------------------------------------------------
# ' Method : Class_Initialize
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Defaults values for use in getInfo and TypeDecode Procs
# '---------------------------------------------------------------------------------------
# Private Sub Class_Initialize()

# iAssignmentMode = inv                          ' undefined

# iClass = inv                                   ' Class of original ItemProperty/Variable/Value
# iType = inv
# iTypeName = vbNullString                       ' TypeName of original
# iScalarType = inv
# iIsArray = False
# iArraySize = inv                               ' used by getInfo and TypeDecode - Procs. -2 means: is fresh.
# iValue = vbNullString                          ' value converted (or # CantDoThat) to a string

# iDepth = 0                                     ' how many times we successfully assign iValue = ivalue.Value
# Set iUp = Nothing
# Set iDown = Nothing
# Set iValue = Nothing                           ' This scalar, variant, Object

# '---------------------------------------------------------------------------------------
# ' Method : Function DrillDown
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def drilldown():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.DrillDown"
    # Call DoCall(zKey, "Sub", eQzMode)

    # iScalarType = IsScalar(TypeName(aValue))
    if iScalarType > 0 Then                        ' this does not always match iType:
    # iAssignmentMode = 1                        ' scalar
    elif iScalarType < 0 Then:
    # iAssignmentMode = 0                        ' do not assign
    else:
    # iAssignmentMode = 2                        ' not scalar, object

    # Set DrillDown = New cInfo
    # Set iDown = DrillDown

    # With DrillDown
    # Set .iUp = Me
    # .iDepth = Me.iDepth + 1
    match iAssignmentMode:
        case 0                                 ' do not assign:
        case 1:
    # .iType = VarType(aValue)
    # .iTypeName = TypeName(aValue)
    # .iValue = aValue                   ' scalar
        case 2:
    # .iType = VarType(aValue)
    # .iTypeName = TypeName(aValue)
    # Set .iValue = aValue               ' not scalar, object
        case _:
    # DoVerify False, "invalid Assignment mode " & iAssignmentMode
    # End With                                       ' DrillDown

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function Top
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def top():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.Top"
    # Call DoCall(zKey, "Function", eQzMode)

    # Set Top = Me
    # While Not Top.iUp Is Nothing
    # Set Top = Top.iUp
    # Wend                                           ' Top -> iUp

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function Find
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def find():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.Find"
    # Call DoCall(zKey, "Function", eQzMode)

    if Direction = 0 Then                          ' start search at top:
    # Set Find = Me.Top
    # Direction = 1
    else:
    # Set Find = Me                              ' start here

    # While Find.iTypeName <> aAttrName
    if Direction = 1 Then                      ' follow up / else down:
    # Find = Me.iDown
    if Find Is Nothing Then:
    # GoTo zExit
    else:
    # Find = Me.iUp
    if Find Is Nothing Then:
    # GoTo zExit
    # Wend                                           ' Find.iTypeName <> aAttrName

    # zExit:
    # Call DoExit(zKey)
    # ' returns "Nothing" if not found

# '---------------------------------------------------------------------------------------
# ' Method : Function IsArrayProperty
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isarrayproperty():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.IsArrayProperty"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Const toDepth As Long = 2
    # Dim haveDepth As Long
    # Dim PropObj As Object

    if DebugMode Then:
    if vArraySize >= 0 Then:
    # DoVerify False, " ??? not expected, who determined this value???"
    # vArraySize = -10000                            ' is not an array property

    # Set PropObj = PropValue
    # GoDeep:
    # With PropObj
    match .Class:
        case olActions, olAttachments, olUserProperties, olLinks, _:
    # olRecipients, olConflicts         ' not defined: , olReplyRecipients, olMemberCount
    # aBugTxt = "value depth=" & haveDepth
    # Call Try
    # vArraySize = .Count
    if Not Catch Then:
    # IsArrayProperty = True
    # Set PropValue = PropObj
    # GoTo FuncExit
        case _:
    # haveDepth = haveDepth + 1
    # aBugTxt = "Get Object value, depth=" & haveDepth
    # Call Try(allowAll)                    ' Try anything, autocatch, Err.Clear
    # Set PropObj = PropValue.Value
    if Catch Then:
    # GoTo FuncExit
    if haveDepth < toDepth Then:
    # GoTo GoDeep
    # End With                                       ' PropValue.value

    # FuncExit:
    # Set PropObj = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function ShowInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showinfo():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.ShowInfo"

    if LenB(iTypeName) = 0 Then:
    # ShowInfo = "This cInfo has no values"
    else:
    # ShowInfo = iDepth & " TypeName " & iTypeName & "(" & iType _
    # & "), ScalarType=" & iScalarType & ", Class="
    if iClass = inv Then:
    # ShowInfo = ShowInfo & "None, "
    else:
    # ShowInfo = ShowInfo & iClass & ", "
    if iIsArray Then:
    # ShowInfo = ShowInfo & "is array with " & iArraySize & " elements, "
    else:
    if iArraySize > 0 Then:
    # ShowInfo = ShowInfo & "array element nr. " & iArraySize & ", "
    else:
    # ShowInfo = ShowInfo & "no array, "
    if iAssignmentMode <= 0 Then:
    # ShowInfo = ShowInfo & "# not decoded or not decodable"
    elif iAssignmentMode = 1 Then:
    # ShowInfo = ShowInfo & "AssignmentMode=1, value: '" & DecodedStringValue & "'"
    elif isEmpty(iValue) Then:
    # ShowInfo = ShowInfo & "AssignmentMode=2, # Empty"
    elif iValue Is Nothing Then:
    # ShowInfo = ShowInfo & "AssignmentMode=2, # Nothing"
    else:
    # ShowInfo = ShowInfo & "AssignmentMode=2, # Object"


# '---------------------------------------------------------------------------------------
# ' Method : Function ShowType
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showtype():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cInfo.ShowType"

    # Dim tInfo As cInfo
    # Dim aLead As String
    # Dim ArrayElements As Boolean

    # Set tInfo = Me
    # Do
    # ShowType = ShowType & vbCrLf & aLead & tInfo.ShowInfo
    # aLead = aLead & withLeads
    if tInfo.iIsArray Then:
    # ArrayElements = True
    if ArrayElements Then:
    if tInfo.iArraySize = 0 Then:
    # aLead = Mid(aLead, Len(withLeads) + 1)
    # ArrayElements = False
    # Set tInfo = tInfo.iDown
    # Loop Until tInfo Is Nothing

    # ShowType = Mid(ShowType, 3)                    ' no initial vbCrLf


# '---------------------------------------------------------------------------------------
# ' Method : Sub PrintType
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def printtype():
    print(Debug.Print ShowType)


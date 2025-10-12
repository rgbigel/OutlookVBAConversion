# Converted from cAttrDsc.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cAttrDsc"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public adKey As String                             ' unique Key corresponding to adName
# Attribute adKey.VB_VarUserMemId = 0
# Attribute adKey.VB_VarDescription = "Display Key"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute adKey.VB_VarUserMemId = 0
# ' Attribute adKey.VB_VarDescription = "Display Key"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public adItem As Object                            ' Item containing this attribute or derived exceptionobject
# ' or adItem of "short AppointmentItem"
# Public adItmDsc As cItmDsc                         ' pointer to cItmDsc of current property -> cObjDsc indirectly
# Public adName As String                            ' Name of Attribute:
# Public adNr As Long                                ' 0 or Index>0 in the order of the decoded properties ??? not any more ***
# '    ==> attributes in aID(1:).idAttrDict
# Public adShowValue As String                       ' value showing or to Show (short)
# Public adFormattedValue As String                  ' value formatted (long)
# Public adDecodedValue As String                    ' current decoded value, shown before user/function changed it

# Public adDictIndex As Long                         ' Index in the property cDictionary aID(apindex) (use Let/Get adDictIndex).odItemDict
# Public adtrueIndex As Long                         ' this is the Index >=0 in the item properties, if exists there
# ' which occurs in the adItem of the decoded item (if not: inv)

# Public adPropClass As OlObjectClass                ' object class represeted with advValue
# Public adPropType As OlUserPropertyType            ' class=99 only: defines the type of the itempropertie's Value
# Public adItemProp As ItemProperty                  ' the original ItemProperty for this attribute. Nothing for PropClass<>99

# Public adInfo As cInfo                             ' all info about current attribute (and value)
# Public advValue As Variant                         ' raw value it the item property's value if any
# Public adHasValue As Boolean                       ' Property has a .Value(.value ... or other obj. content)
# Public adNotDecodable As Boolean                   ' Property is not decodable by rule
# Public adHasValueNow As Boolean                    ' .Value has been decoded
# Public adOrigValDecodingOK As Boolean              ' values have been determined and decoded at least once

# Public adKillMsg As String                         ' message indicating what function did after changing AttrValue

# Public adisUserAttr As Boolean                     ' results in Key = adName & "#W2"
# Public adBaseadMod As Boolean                      ' results in Key = adName & "#B"
# Public adRules As cAllNameRules                    ' rules defining the attribute functionality of the class
# '   and user-applied options
# Public adRuleIsModified As Boolean                 ' modifications to class Rule have been done
# Public adRuleBits As String                        ' string form of cRule bits

# Public adisSpecialAttr As Boolean                  ' type of adItemProp
# '       if true it is special,
# '       else    it is a normal ItemProperty
# ' then there is no ItemProperties in Item
# Public adisSel As Boolean                          ' this Property is selected by Rules OR special requests
# Public adLastCMode As Long                         ' value for Source of adictClone defines behavior

# Private Sub Class_Initialize()
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "cAttrDsc.Class_Initialize"
# Call DoCall(zKey, "Sub", eQzMode)


# ' using targetindex (there is no local copy)
# DoVerify PropertyNameX <> vbNullString
# adKey = MakeAttributeKey(PropertyNameX)        ' construct / treat special cases
# adName = PropertyNameX
# adDictIndex = inv                              ' mark as not inited for this property
# adNr = inv
# Set Me.adInfo = New cInfo

if ExceptionProcessing Then:
# DoVerify False, "is this real ???"
# adInfo.iType = vbNull
else:
# Set adItem = aID(targetIndex).idObjItem    ' owning item
# Set aDecProp(targetIndex) = Me

if aID(targetIndex) Is Nothing Then:
# Set aID(targetIndex) = New cItmDsc
if aID(targetIndex).idAttrDict Is Nothing Then:
# DoVerify False, "design check aID(targetIndex).idAttrDict Is Nothing ???"
# Call InitializeValues

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub mainInits
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub mainInits()
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "cAttrDsc.mainInits"
# Call DoCall(zKey, "Sub", eQzMode)


# Dim aObjDscClass As OlObjectClass
# Dim aDi As cAttrDsc
# Dim trapme As Long

if DoVerify(adItmDsc Is aID(aPindex), "adItmDsc IsNot aID(aPindex) ???") Then:
# Set adItmDsc = aID(aPindex)
# Set adItemProp = aProp                         ' owning item-Property

if DebugMode Then:
if workingOnNonspecifiedItem Then:
# DoVerify False, "Design stop **"

# trapme = adItmDsc.idAttrDict.Count
if adItmDsc.idAttrDict.Exists(Me.adKey) Then:
# DoVerify trapme = adItmDsc.idAttrDict.Count, "??? how that: got a bug! ???"
# aBugVer = Not isEmpty(adItmDsc.idAttrDict.Item(Me.adKey))
# aBugTxt = "bug escape: If Exists(" & adKey & ") should never change its count ???"
if DoVerify Then:
# Set aDi = Me
# Set adItmDsc.idAttrDict.Item(Me.adKey) = aDi ' replace by non-empty cAttrDsc (Me)
else:
# Set aDi = Me
else:
# adItmDsc.idAttrDict.Add Me.adKey, Me
# adDictIndex = adItmDsc.idAttrDict.Count - 1
if adDictIndex = inv Then:
# adDictIndex = adItmDsc.idAttrDict.Count - 1

# aObjDscClass = adItmDsc.idObjItem.Class
if aObjDscClass <= olException _:
# And aObjDscClass >= olRecurrencePattern Then '   Exception, Exceptions, RecurrencePattern
# DoVerify Not isSpecialName
# DoVerify adisSpecialAttr = False, "imposs after New???"
elif (adDictIndex > TotalPropertyCount _:
# And TotalPropertyCount > 0) _
# Or BaseAndSpecifiedDiffer _
# And Not workingOnNonspecifiedItem Then
if isSpecialName Then:
# DoVerify False, "Test support for isSpecialName for " & PropertyNameX
# adisSpecialAttr = True                 ' in a RecurrencePattern or Exception(s)
else:
# DoVerify Not isSpecialName
# DoVerify adisSpecialAttr = False, "isSpecialName imposs after New???"
# adBaseadMod = BaseAndSpecifiedDiffer
if DoVerify(adName = PropertyNameX, "design check Not aID(aPindex).idAttrDict Is Nothing ???") Then:
# adName = PropertyNameX
# adRuleIsModified = True                        ' brand new here, no Rules yet

# Set aDi = Nothing

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub InitializeValues
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub InitializeValues()
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "cAttrDsc.InitializeValues"
# Call DoCall(zKey, "Sub", eQzMode)

# Dim lowPx As Long

# Set adItmDsc = aID(aPindex)
# adtrueIndex = inv                              ' mark this as not linked to list item properties

match aCloneMode:
    case DummyTarget:
# Stop: Stop                             ' ????
# adLastCMode = withNewValues
# Set adItmDsc.idObjItem = aID(aPindex)  ' self link ???
if aPindex > 2 Then:
# lowPx = aPindex - 2
else:
# lowPx = aPindex
    case withNewValues:
# Call mainInits
    case FullCopy:
# Call mainInits

# adLastCMode = aCloneMode
# adBaseadMod = BaseAndSpecifiedDiffer
# adRuleIsModified = True                        ' brand new here, no Rules yet

# aBugVer = Not adItmDsc.idObjDsc Is Nothing
# DoVerify aBugVer, "** Not adItmDsc.idObjDsc Is Nothing ???"
# Set aTD = Me

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub GetSpecialAttributeValues
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getspecialattributevalues():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cAttrDsc.GetSpecialAttributeValues"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cAttrDsc")

    # Dim iClass As OlObjectClass

    # iClass = advValue.Class                        ' objekte und bekannte Flle
    if IsOneOf(Array(olConflicts, _:
    # olLinks, _
    # olActions, _
    # olAttachments, _
    # olRecipients, _
    # olUserProperties), iClass) > 0 Then ' these are the special Attributes

    # aBugTxt = "get count of attribute " & adName
    # Call Try
    # adInfo.iArraySize = advValue.Count         ' Count kann fehlen!
    if CatchNC Then:
    # adInfo.iArraySize = inv
    if DebugMode Then:
    # adInfo.iArraySize = adInfo.iArraySize
    if ErrorCaught = 287 Then:
    # DoVerify False, " user did not grant access to the attribute"
    # Call ErrReset(0)

    # Call PrepDecodeProp
    if DebugMode Then:
    # DoVerify aTD.adOrigValDecodingOK
    else:
    if DebugMode And Not aTD.adOrigValDecodingOK Then:
    print(Debug.Print "class not implemented", iClass, _)
    # aTD.adName, _
    # TypeName(aTD.advValue.Value), _
    # aTD.GetScalarValue
    # DoVerify False

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetScalarValue
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: (Attempt to) get value of Attribute-Property and try to get as Scalar
# '---------------------------------------------------------------------------------------
def getscalarvalue():
    # Dim zErr As cErr
    # Const zKey As String = "cAttrDsc.GetScalarValue"

    # Dim Significant As Long

    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tFunction, ExplainS:="cAttrDsc(" & adItemProp.Name & ")")

    # advValue = "# not decoded " & adRuleBits
    # adHasValueNow = False
    # DoVerify Not adItmDsc Is Nothing, "design check adItmDsc IsNot Nothing ??? "
    # DoVerify Not adItmDsc.idObjDsc Is Nothing, "design check adItmDsc.idObjDsc IsNot Nothing ??? "
    # DoVerify aObjDsc Is adItmDsc.idObjDsc, "design check aObjDsc Is adItmDsc.idObjDsc ??? "
    if aObjDsc.objClassKey = vbNullString Then:
    # aObjDsc.objClassKey = adItmDsc.idObjItem.Class
    # DoVerify D_TC.Exists(adItmDsc.idObjDsc.objClassKey), "design check D_TC.Exists(adItmDsc.idObjDsc.objClassKey) ???"
    if DoVerify(aObjDsc Is D_TC(adItmDsc.idObjDsc.objClassKey), "design check aObjDsc Is D_TC(adItmDsc.idObjDsc.objClassKey) ???") Then:
    # Set aObjDsc = D_TC.Item(aObjDsc.objClassKey)
    # Set adItmDsc.idObjDsc = aObjDsc

    if adRules Is Nothing Then:
    # aObjDsc.objClsRules.RuleIsSpecific = False
    # aObjDsc.objClsRules.clsObligMatches.RuleMatches = False
    # adNotDecodable = InStr(aObjDsc.objClsRules.clsNotDecodable.CleanMatchesString, adKey) > 0
    if iRules Is Nothing Then:
    # Set adRules = aObjDsc.objClsRules      ' use class locally
    else:
    # Set adRules = iRules
    elif adRules.clsNotDecodable.RuleMatches Then:
    # adNotDecodable = True                      ' we just can't
    # adHasValueNow = False
    elif adRules.clsNeverCompare.RuleMatches _:
    # And SkipDontCompare Then                ' can not be NotDecodable at the same time
    # advValue = advValue & " SkipRequest"
    # adHasValueNow = False
    # adNotDecodable = True                      ' do not want to

    # Significant = InStr(adRules.clsObligMatches.CleanMatchesString, adKey)
    # Significant = InStr(adRules.clsSimilarities.CleanMatchesString, adKey) + Significant
    if Significant = 0 Then:
    # adRules.clsObligMatches.RuleMatches = False
    if quickChecksOnly Then:
    # advValue = "# not decoded (not oblig.)"
    # adHasValueNow = False
    # adNotDecodable = True                  ' do not want to
    else:
    # adRules.clsObligMatches.RuleMatches = True
    # adRules.RuleIsSpecific = True
    # Call AppendTo(SelectedAttributes, adKey, b)

    if adOrigValDecodingOK _:
    # And adPropType > 0 _
    # And adInfo.iType < 10000 Then               ' known / defined attribute
    if DebugMode Then                          ' do some plausi checks:
    # DoVerify adPropClass = VarType(adItemProp)
    if adPropClass = olItemProperty Then:
    # DoVerify adInfo.iType = VarType(adItemProp.Value)

    # Call getInfo(adInfo, adItemProp, _
    # Not adNotDecodable)               ' define and go deep until Scalar *** Hier
    if adInfo.iAssignmentMode = 1 Then:
    # adDecodedValue = adInfo.DecodedStringValue
    # advValue = adDecodedValue
    # adShowValue = adDecodedValue
    # adKillMsg = vbNullString
    # adHasValue = True
    # adOrigValDecodingOK = True
    if Left(adInfo.DecodedStringValue, 1) = "#" Then:
    # adHasValueNow = False
    else:
    # adHasValueNow = True

    # FuncExit:
    # GetScalarValue = adInfo.iType
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr, Me.adName)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function isUserProperty
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isuserproperty():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cAttrDsc.isUserProperty"
    # Call DoCall(zKey, tFunction, eQzMode)

    # isUserProperty = adItemProp.isUserProperty
    # GoTo zExit
    if isUserProperty Then:
    # GoTo zExit
    if Parent Is Nothing Then:
    # DoVerify False
    if adInfo.iType > 0 Then:
    # isUserProperty = Parent.adisUserAttr
    else:
    # isUserProperty = False
    if Parent.adisUserAttr <> isUserProperty Then:
    # DoVerify False

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function adictClone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def adictclone():
    # Dim zErr As cErr
    # Const zKey As String = "cAttrDsc.adictClone"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cAttrDsc")

    # Dim nCloneAttrDsc As cAttrDsc

    if aPindex <> targetIndex Then:
    # targetIndex = aPindex
    # DoVerify Me.adName <> vbNullString
    # PropertyNameX = Me.adName
    # adtrueIndex = Me.adtrueIndex
    # Me.adLastCMode = aCloneMode                    ' parameter for New AttrDsc
    # Set nCloneAttrDsc = New cAttrDsc               ' not using ProvideAttrDsc!
    # Set aProp = Nothing
    # With nCloneAttrDsc
    if aID(targetIndex) Is Nothing Then:
    # DoVerify False, "Design check: aID(targetIndex) Is Nothing ???"
    # Set aID(targetIndex) = nCloneAttrDsc
    # Set aID(targetIndex).idObjDsc = New cObjDsc
    # Set aID(targetIndex).idObjItem.objSeqInImportant = New Collection

    if aCloneMode = FullCopy Then:
    # DoVerify False, "** Design check: cloning of aOD and/or aID needed ???"
    # Set aOD(targetIndex) = aOD(sourceIndex)
    # DoVerify aOD(targetIndex).objNameExt = aOD(sourceIndex).objNameExt, _
    # "Object Name extension mismatch"
    # Set aID(targetIndex) = aID(sourceIndex)
    else:
    # Set .adItem = Nothing
    # Set .adItmDsc.idAttrDict = Nothing
    # .adDictIndex = Me.adDictIndex
    # Set .adRules = Me.adRules
    # .adRuleIsModified = Me.adRuleIsModified
    # .adRuleBits = Me.adRuleBits
    # .adisSpecialAttr = Me.adisSpecialAttr

    match aCloneMode:
        case DummyTarget                       ' invalidate skipped attributes:
    # .adtrueIndex = inv
    # .adPropClass = inv
    # .adPropType = inv
    # .adInfo.iType = 10002
    # Set .advValue = Nothing
    # .adInfo.iArraySize = inv
    # .adInfo.iAssignmentMode = 1
    # .adOrigValDecodingOK = True        ' do not decode, its a dummy
    # .adDecodedValue = "# Property not decoded (DummyTarget)"
    # .adHasValue = False
    # .adHasValueNow = True
    # .adShowValue = vbNullString
    # .adFormattedValue = vbNullString
    # .adKillMsg = "# dummy Copy"
    # Set .adItemProp = Nothing
    # .adisSel = False

        case FullCopy:
    # .adtrueIndex = Me.adtrueIndex
    # .adPropClass = Me.adPropClass
    # .adPropType = Me.adPropType
    # .adInfo.iType = Me.adInfo.iType
    # .adInfo.iAssignmentMode = Me.adInfo.iAssignmentMode
    if Me.adHasValueNow And Me.adOrigValDecodingOK Then:
    if Me.adInfo.iAssignmentMode = 1 Or Not Me.adHasValue Then:
    # .advValue = Me.advValue
    else:
    # Set .advValue = Me.advValue
    # .adInfo.iArraySize = Me.adInfo.iArraySize
    # .adOrigValDecodingOK = Me.adOrigValDecodingOK
    # .adDecodedValue = Me.adDecodedValue
    # .adHasValue = Me.adHasValue
    # .adHasValueNow = Me.adHasValueNow
    # .adHasValueNow = False
    # .adShowValue = Me.adShowValue
    # .adFormattedValue = Me.adFormattedValue
    # .adKillMsg = Me.adKillMsg
    # Set .adItemProp = Me.adItemProp
    # .adNr = Me.adNr
    # .adisSel = Me.adisSel
        case withNewValues                     ' Value must be determined from (new) itemproperty:
    # .adtrueIndex = Me.adtrueIndex
    # .adPropClass = Me.adPropClass
    # .adPropType = Me.adPropType
    # .adInfo.iType = 10003
    # .adInfo.iAssignmentMode = 0
    # Set .advValue = Nothing
    # .adInfo.iArraySize = inv           ' to be corrected when we get the value
    # .adOrigValDecodingOK = False
    # .adHasValue = False
    # .adHasValueNow = False
    # .adDecodedValue = vbNullString
    # .adShowValue = vbNullString
    # .adFormattedValue = vbNullString
    # .adKillMsg = vbNullString
    # ' Set .adItemProp = Nothing                 ' corrected on the cObjDsc-Lvlel (no change here)
    # ' .attrPos = Me.attrPos                       not yet: remain zero until we add it to aID(1).idAttrDict
    # .adisSel = Me.adisSel
        case _:
    # DoVerify False
    # End With                                       ' nCloneAttrDsc
    # Set adictClone = nCloneAttrDsc
    # CloningMode = False

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PrintColl
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def printcoll():

    # Dim i As Long
    # Dim j As Long

    # Dim Line As String
    # Dim aE1 As String
    # Dim aE2 As String
    # Dim match As String

    if i = 0 Then:
    # Line = LString(Collections(j).Item(i), width0)
    elif i = 1 Then:
    # aE1 = LString(Collections(j).Item(i), width)
    elif i = 2 Then:
    # aE2 = LString(Collections(j).Item(i), width)
    if aE1 = aE2 Then:
    # match = " = "
    else:
    # match = " ! "
    # Line = Line & aE1 & match & aE2
    else:
    # Line = Line & LString(Collections(j).Item(i), width)
    print(Debug.Print Line)

    # ProcRet:


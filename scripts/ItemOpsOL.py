# Converted from ItemOpsOL.py

# Attribute VB_Name = "ItemOpsOL"
# Option Explicit

# ' all item properties decoded (and converted to string, if possible). Some are skipped.

# '---------------------------------------------------------------------------------------
# ' Method : Sub CheckPhoneNumber
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkphonenumber():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.CheckPhoneNumber"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim lName As String
    # lName = LCase(PropertyNameX)
    # NormalizedPhoneNumber = vbNullString
    if InStr(lName, "radio") > 0 Then:
    # ItsAPhoneNumber = False
    elif InStr(lName, "telex") <= 1 _:
    # And InStr(lName, "extension") = 0 _
    # And InStr(lName, "phone") = 0 _
    # And InStr(lName, "fax") = 0 Then
    # ItsAPhoneNumber = False
    elif InStr(lName, "phone") > 0 _:
    # Or InStr(lName, "fax") > 0 Then
    # ItsAPhoneNumber = True ' fax or phone IS a telco number
    else:
    # ItsAPhoneNumber = True ' anything with number in it is a telco number??
    # DoVerify False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PhoneNumberNormalize
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def phonenumbernormalize():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.PhoneNumberNormalize"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if Len(oFoneNr) < 5 And Left(oFoneNr, 1) <> "1" Then ' extension number:
    if InStr(oFoneNr, "*") = 0 Then:
    # NormalizedPhoneNumber = "*" & oFoneNr
    else:
    # NormalizedPhoneNumber = oFoneNr
    else:
    # NormalizedPhoneNumber = NormalizeTelefonNumber(oFoneNr)

    if oFoneNr <> NormalizedPhoneNumber Then:
    if SelectOnlyOne And px = 1 And Not aDecProp(1) Is Nothing Then    ' save original for Display (Excel):
    # aCloneMode = FullCopy
    # Set aDecProp(2) = aDecProp(1).adictClone        ' save original TelNumber value in number 2
    # MatchPoints(px) = MatchPoints(px) + 1               ' ratedelta = 1
    # PhoneNumberNormalized = True
    # aStringValue = NormalizedPhoneNumber                ' new value in px
    if CurIterationSwitches.SaveItemRequested Then:
    # WorkItemMod(px) = True
    # Message = fiMain(px) & b _
    # & aTD.adName _
    # & " gendert in " _
    # & NormalizedPhoneNumber _
    # & " (war " & oFoneNr & ")"
    # Call LogEvent(Message, eLall)
    # oFoneNr = NormalizedPhoneNumber
    elif Not saveItemNotAllowed Then:
    # ' not changing WorkItemMod(px) !!!
    # Message = fiMain(px) & b _
    # & aTD.adName _
    # & " wird verglichen als " _
    # & NormalizedPhoneNumber _
    # & " (war " & oFoneNr & ")"
    # Call LogEvent(Message, eLall)
    # oFoneNr = NormalizedPhoneNumber

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsPropertyOK
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def ispropertyok():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.IsPropertyOK"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim iProp As ItemProperty
    # Dim ThisPropisSelected As Boolean

    if apropTrueIndex >= 0 Then                             ' we know where it is:
    # GoTo useTrueIndex
    if LenB(CritPropName) > 0 _:
    # And UCase(CritPropName) <> "OR" _
    # And UCase(CritPropName) <> "NOT" _
    # And UCase(CritPropName) <> "A" Then
    # ' GetAttrDsc when finding...
    # useTrueIndex:
    # ThisPropisSelected = (apropTrueIndex < 0)
    if ThisPropisSelected Then                          ' we know where:
    if LenB(CritPropName) = 0 Then:
    # DoVerify False, "CritPropName empty, dead"
    # PropertyNameX = CritPropName                    ' we know a name but no index
    else:
    if LenB(CritPropName) = 0 Then                  ' check if we can use existing aTD:
    if aTD Is Nothing Then:
    # DoVerify False, "No aTD and no CritPropName: makes no sense. Must quit"
    if aTD.adName <> CritPropName Then:
    # DoVerify False, "aTD Name mismatches CritPropName"
    # aBugVer = iProp Is aTD.adItemProp
    if DoVerify(aBugVer, "design check iProp Is aTD.ADItemProp ???") Then:
    # Set iProp = aTD.adItemProp
    else:
    if aTD Is Nothing Then:
    if aProp Is Nothing Then:
    # PropertyNameX = vbNullString                  ' FindProperty in position aPropTrueIndex
    elif aProp.Name = PropertyNameX Then:
    # Set iProp = aProp
    # GoTo why
    else:
    if aTD.adName <> PropertyNameX Then:
    # DoVerify False, " inconsistent; very bad"
    # PropertyNameX = vbNullString                  ' FindProperty in position aPropTrueIndex
    # Set aTD = Nothing
    # ' FindProperty will also NewAttrDsc (if possible and needed)
    # ' and it will get the found ItemProperty into (new) aTD
    # ' +   evaluate Prop. Value into ADDecodedValue, determining the Value Type
    # ' +   if possible, determine the trivial value (without formatting)
    if apropTrueIndex <> aTD.adtrueIndex Then:
    # why:
    # apropTrueIndex = FindProperty(apropTrueIndex, _
    # PropertyNameX, iProp, _
    # Item)                       ' iProp is OUT, aPropTrueIndex checked or OUT
    if apropTrueIndex < 0 Then                          ' not found anywhere:
    # GoTo ProcReturn                                 ' property is not defined in thisItem,
    if aTD Is Nothing Then:
    # apropTrueIndex = -1
    # DoVerify False, " may be correct ???"
    # GoTo why

    # AttributeIndex = aTD.adtrueIndex
    # IsPropertyOK = True                                 ' make a new one in known position

    # ' we want all if we dont want OnlyMostImportantProperties
    if aTD.adNr > 0 Then:
    # ' format (e.g. Phone Numbers), specials for some array cases (e.g. MemberCount->Members, Photos)
    # Call PrepDecodeProp
    # Call logDecodedProperty(aStringValue, String(4, b))
    # Call StackAttribute                             ' add this to aID(aPindex).idAttrDict [dictionary]

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr, CStr(IsPropertyOK))

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function EvaluateSpecialRequirements
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def evaluatespecialrequirements():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.EvaluateSpecialRequirements"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim originalItemClass As OlObjectClass
    # Dim BaseItem As Object

    # workingOnNonspecifiedItem = False
    # originalItemClass = oItem.Class
    # BaseAndSpecifiedDiffer = False  ' normally, BaseItem[=standarditem] == aID(2..3
    # ' but may be moved to 3..4 for exc/occ
    # EvaluateSpecialRequirements = aPindex   ' normally, no change
    # DoVerify aID(aPindex).idObjItem Is oItem, "standard item class index 1..2 must match ** Remove 2 if no hit"
    # Set aID(aPindex).idObjItem = oItem   ' standard item class index 1..2

    if originalItemClass = olAppointment And aPindex < 3 Then:
    # ' could be olApptMaster, olApptOccurrence or olApptException:
    if oItem.ItemProperties.Count < minPropCountForFullItem Then ' short item:
    match oItem.RecurrenceState:
        case olApptOccurrence:
    # SpecialObjectNameAddition = "#O"
    # GoTo NonStandardItem
        case olApptException:
    # SpecialObjectNameAddition = "#E"
    # NonStandardItem:
    if aID(aPindex + 2) Is Nothing Then:
    # Set aID(aPindex + 2).idObjDsc = New cObjDsc
    # Set aID(aPindex + 2).idObjDsc.objSeqInImportant = New Collection
    # Call aID(aPindex + 2).SetDscValues(oItem, _
    # withValues:=True, aRules:=sRules, _
    # SD:=SpecialObjectNameAddition)
    # DoVerify aOD(aPindex + 2).objNameExt = SpecialObjectNameAddition, _
    # "creation of Name/Key/Extension not logical ???"
    # ' *** note: this aID now runs under aPindex, not aPindex+2
    # ' *** because we use the same idAttrDict for adding further attrs
    # Set aID(aPindex + 2).idObjItem = oItem  ' put specified item into +2 pos
    # ' and standard item (its parent) +0 pos
    # Set BaseItem = oItem.Parent             ' parent done first, specified item second
    # DoVerify BaseItem.RecurrenceState = olApptMaster, "oItem Parent not plausible"
    # Set aID(aPindex).idObjItem = BaseItem   ' standard item class index 1..2
    # aOD(aPindex).objMaxAttrCount = 0        ' this is a new item to decode
    # workingOnNonspecifiedItem = True        ' using parent ^= specified
    # BaseAndSpecifiedDiffer = True           ' not the specified item yet
    # EvaluateSpecialRequirements = aPindex + 2
        case olApptMaster:
    # ' no need to go to parent, specifiedItem == aID(apindex)
        case _:
    # DoVerify False, "RecurrenceState not expected/defined"

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitAttributeSetup
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Create Maintanance information for base Item Description / Extension
# '---------------------------------------------------------------------------------------
def initattributesetup():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.InitAttributeSetup"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim BasePx As Long

    # BasePx = baseID.idPindex
    if BasePx = 0 Then:
    # BasePx = 1
    # baseID.idPindex = BasePx
    if ExtensionID Is Nothing Then:
    # Set ExtensionID = baseID
    # Call N_ClearAppErr
    # Set rP(BasePx) = Nothing                       ' set to recurrencePattern when IsRecurring in IsPropertyOK
    if aID(BasePx).idAttrDict.Count = 0 Then:
    # AttributeIndex = 0
    # AttributeUndef(BasePx) = 0
    if aID(BasePx).idAttrDict.Count < AttributeIndex Then:
    # AttributeIndex = aID(BasePx).idAttrDict.Count

    if AttributeIndex > 0 Then:
    if AttributeIndex = ExtensionID.idObjDsc.objMaxAttrCount Then:
    # GoTo FuncExit                           ' all done

    # ' normal case presets:
    # MatchPoints(aPindex) = 0
    # ' we have a new aID when we are in SelectAndFind (apindex=2)
    if baseID.idObjItem.Class <> ExtensionID.idObjDsc.objItemClass Then:
    # DoVerify False, "design test only ???"
    # Call EvaluateSpecialRequirements(baseID.idObjItem)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' Important Properties (by Rules) first in sequence in aID(aPindex).idAttrDict, ??? design change not completed !!!
# ' then rest in PropertyIndex sequence (unless StopAfterMostRelevant),

# '---------------------------------------------------------------------------------------
# ' Method : GetMiAttrNr
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Most Important Attribute Number for PropertyNameX, added to SeqInImportant
# ' Note:    requires ATD and aObjDsc
# '---------------------------------------------------------------------------------------
def getmiattrnr():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "ItemOpsOL.GetMiAttrNr"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim i As Long
    # Dim testADNr As Boolean

    # ' note: MostImportantAttributes is usually identical with SelectedAttributes, cut can be longer. Never shorter.
    if InStr(MostImportantAttributes, PropertyNameX) > 0 Then:
    if aTD.adisSel _:
    # And aTD.adRules.clsObligMatches.RuleMatches _
    # And aTD.adRules.clsSimilarities.RuleMatches Then
    # testADNr = True
    if MostImportantProperties(i) = PropertyNameX Then:
    # DoVerify aTD.adDictIndex = apropTrueIndex + 1, _
    # "design check aTD.ADDictIndex = apropTrueIndex + 1 ???"
    if aTD.adNr >= 0 Then:
    if DoVerify(aTD.adNr = i, _:
    # "change in position in MostImportantAttributes ??? " & PropertyNameX) Then
    # aTD.adNr = i
    else:
    # GoTo zExit                          ' all as done befor
    else:
    # DoVerify aTD.adDictIndex = apropTrueIndex + 1, "design check ???"
    # aObjDsc.objSeqInImportant.Add aTD.adDictIndex
    if aTD.adNr <> aTD.adDictIndex Then:
    # DoVerify Not testADNr, "why do we assign a new .ADNR"
    # aTD.adNr = aTD.adDictIndex
    # aTD.adisSel = True
    # Call LogEvent("added objSeqInImportant(" _
    # & aObjDsc.objSeqInImportant.Count & ") for Property=" & aTD.adKey _
    # & ", mostImportantProperty#=" & i & " in Object Class " & aObjDsc.objItemClassName _
    # & " [Dict Index " & aTD.adDictIndex & "]", eLSome)
    # GoTo zExit
    # DoVerify aTD.adisSel Or aTD.adNr > 0, "look into this ???" ' note: ADNr=0 always Key=Application, do not select
    # aTD.adisSel = False
    # aTD.adNr = -1
    # zExit:
    # Call DoExit(zKey)

# ' from Attributes of Item px -> aPindex, formatting for output
# ' MAKE SURE IsRecurring is part of Important Properties if recurrence pattern needed
def getitemattrdscs():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetItemAttrDscs"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim somevaluesNotAvailable As Collection
    # Dim iPos As Long ' position in ItemProperties if we know it
    # Dim tPos As Long '              "                certain
    # Dim itmNo As Long
    # Dim aIDa As cObjDsc

    # aPindex = px
    # AttributeIndex = 0

    # Set somevaluesNotAvailable = New Collection
    if LenB(TrueCritList) = 0 Then:
    # SkipDontCompare = False

    if isEmpty(MostImportantProperties) Then:
    # MostImportantProperties = Array(sRules.clsObligMatches.CleanMatches(0))
    # MostImportantAttributes = sRules.clsObligMatches.CleanMatches(0)

    # Set aProps = Item.ItemProperties
    # TotalPropertyCount = aProps.Count

    # '  NOTE: we will not work on aID(apindex).odItemDict here, because
    # '        we also have to process the specified items (#B)
    # '        so we use Item.itemproperties instead

    # iPos = 1    ' start position for loop: index in odItemDict
    # doSpecialAttributes:
    # OneDiff = vbNullString
    # PropertyNameX = Item.ItemProperties.Item(tPos).Name
    if aID(aPindex) Is Nothing Then:
    # Set aTD = Nothing
    else:
    # Set aTD = aID(aPindex).GetAttrDsc4Prop(tPos)    ' setting AllPublic.aProp, reusing if present
    if aTD Is Nothing Then                      ' not in list of SelectedAttributes: new:
    # apropTrueIndex = tPos                   ' we know
    # aBugTxt = "Design check Not aID(aPindex).idAttrDict Is Nothing ???, Property tPos=" & tPos
    # DoVerify Not aID(aPindex).idAttrDict Is Nothing, aBugTxt
    # With aID(aPindex).idAttrDict
    if .Exists(PropertyNameX) Then:
    # Set aTD = .Item(PropertyNameX)
    if aTD.adtrueIndex < 0 Then:
    # aTD.adtrueIndex = apropTrueIndex
    # DoVerify aProp.Name = aTD.adName, _
    # "Property Name mismatch for aTD true index=" & aTD.adtrueIndex
    else:
    # Set aTD = ProvideAttrDsc
    if Not aTD.adRules.RuleInstanceValid Then:
    # Call SplitDescriptor(aTD)
    # End With ' aID(aPindex).idAttrDict

    if isSpecialName Then:
    # apropTrueIndex = -apropTrueIndex    ' must be correct for special attributes. never on -1!!!

    # aTD.adtrueIndex = apropTrueIndex
    # DoVerify PropertyNameX = aTD.adName, "Property Name messed up ???"

    # ' rules for property name and its isSpecialName (suffix #?) are always same
    if Not aTD.adRules Is iRules Then:
    # Set aTD.adRules = iRules
    # Call SplitDescriptor(aTD)               ' using PropertyNameX inside
    # Call GetMiAttrNr

    # PhoneNumberNormalized = False
    # IgString = vbNullString
    if aTD.adNr > 0 Or AllProps Then            ' SkipDontCompare is NOT considered for aTD setup:
    # ' format (e.g. Phone Numbers), specials for some array cases (e.g. MemberCount->Members, Photos)
    # DoVerify Not iRules Is Nothing, "iRules can't be missing ???"
    # Call PrepDecodeProp
    # Call StackAttribute                             ' add this to aID(aPindex).idAttrDict [dictionary]
    else:
    # aStringValue = "## value not evaluated"

    # aTD.adDecodedValue = aStringValue
    # AttributeIndex = aTD.adtrueIndex
    if Not displayInExcel Then:
    # Call logDecodedProperty(aStringValue, String(4, b))
    elif DebugLogging Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab & "Progress:", _)
    # itmNo, AttributeIndex, PropertyNameX, aStringValue, IgString
    # iPos = iPos + 1

    if workingOnNonspecifiedItem Then               ' more to do on specifiedItem ???    ' *** Hier:
    # workingOnNonspecifiedItem = False
    # ' we do NOT add these special attributes to aID(aPindex + 2)
    # Set aIDa = aID(aPindex)
    # PropertyNameX = vbNullString
    # IgString = vbNullString
    # logDecodedProperty vbNullString, " ++ Starting on special attributes +++"
    # ' but we get the properties from this idObjItem
    # Set Item = aID(aPindex + 2).idObjItem
    # isSpecialName = True                        ' use name suffix in dictionary
    # ' and simply continue adding to Dictionary, starting with tPos
    # ' from previous loop exit, iPos == (old) itemproperties.count
    # iPos = tPos                                 ' used to access attributes for apindex + 2
    # GoTo doSpecialAttributes                    ' of the specified item

    # asFarAsWeWanted:
    if aPindex < 3 Then                             ' Processing Recurrence and Exceptions, if any:
    if Not (rP(aPindex) Is Nothing Or ExceptionProcessing) Then:
    # ExceptionProcessing = True
    # IgString = vbNullString
    if Not rP(aPindex) Is Nothing Then      ' recurrence Pattern has no ItemProperties:
    # tPos = rPTrueIndex                  ' base of indirect attributes, < 0
    # Call RpStackAndLog(aPindex, rP(aPindex))
    # ' so we must process the properties we know
    # ' also the Exceptions thereof
    # Set rP(aPindex) = Nothing
    # ExceptionProcessing = False
    # ' maxProperties (ever)
    # iPos = aID(aPindex).idAttrDict.Count
    if aOD(0).objMaxAttrCount < iPos - 1 Then:
    # aOD(0).objMaxAttrCount = iPos + 1
    if aOD(0).objMinAttrCount = 0 Then:
    # aOD(0).objMinAttrCount = iPos - 1
    else:
    if aOD(0).objMinAttrCount > iPos Then:
    # aOD(0).objMinAttrCount = iPos - 1
    if CurIterationSwitches.SaveItemRequested And Not workingOnNonspecifiedItem Then:
    if aID(aPindex).idObjItem.Class = olContact Then:
    # MPEchanged = False
    # Call NameCheck(aID(aPindex).idObjItem)
    if MPEchanged Then:
    # WorkItemMod(aPindex) = True
    if somevaluesNotAvailable.Count = 0 Then:
    if DebugLogging And Not ShutUpMode Then:
    print(Debug.Print "no Attributes specified are missing")
    else:
    # Call LogEvent("missing " & somevaluesNotAvailable.Count _
    # & " specified attributes in this item")
    # Call LogEvent(iPos & vbTab & somevaluesNotAvailable.Item(iPos))

    if DebugMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
    # & "ended GetItemAttrDscs with " _
    # & AttributeIndex & " properties, subject:" _
    # & vbCrLf, Quote(aID(aPindex).idObjItem.Subject)

    if SelectOnlyOne Or FindMatchingItems Then:
    if LenB(SelectedAttributes) = 0 Then:
    # Call RulesToExcel(aPindex, Not FindMatchingItems)
    elif aPindex = 2 Then:
    # Call AppendMissingProperties(0, 1, 2)
    # Call AppendMissingProperties(0, 2, 1)
    # Call RulesToExcel(aPindex, Not FindMatchingItems)
    # OnlyMostImportantProperties = quickChecksOnly   ' restore user choice

    # FuncExit:
    # Set aIDa = Nothing
    # Set somevaluesNotAvailable = Nothing
    # Set aID(aPindex).idObjItem = Item

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function NewAttrDsc
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def newattrdsc():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.NewAttrDsc"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Set NewAttrDsc = ProvideAttrDsc                 ' uses aProp to find or make cAttrDsc
    if lpropTrueIndex > -1 Then:
    # NewAttrDsc.adtrueIndex = lpropTrueIndex
    if aTD Is Nothing Then:
    # DoVerify False, "design check ???"
    # Set aTD = NewAttrDsc
    # Set iRules = Nothing
    if iRules Is Nothing Then:
    # DoVerify False, "design check ???"
    # Call CreateIRule(PropertyNameX)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AppendMissingProperties
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def appendmissingproperties():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.AppendMissingProperties"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim startIndex As Long
    # Dim nDecProp As cAttrDsc
    # ' are we missing anything ?
    if AttributeIndex = 0 Then  ' select base count:
    # AttributeIndex = aID(base).idAttrDict.Count
    if aID(Copy).idAttrDict Is Nothing Then:
    # Stop ' ???
    # Set aID(Copy).idAttrDict = New Dictionary
    # startIndex = aID(Copy).idAttrDict.Count
    if startIndex < AttributeIndex Then:
    # i = startIndex
    # While i < aID(base).idAttrDict.Count - 1 ' additional attrs w/o value
    # i = i + 1
    # aCloneMode = DummyTarget
    # Set nDecProp = aID(base).idAttrDict.Item(i).adictClone  'frher ??? (Copy, base)
    # aID(Copy).idAttrDict.Add nDecProp.adKey, nDecProp
    # Wend

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetAobj
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getaobj():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetAobj"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if aPindex = px Then:
    # aPindex = px                                        ' this Aobj is new here, GetAobj?
    # aItmIndex = knownItemIndex
    # WorkItemMod(px) = False
    if knownItemIndex >= 0 Then:
    # targetIndex = px
    if sortedItems(px) Is Nothing Then                  ' sorted items have highest prio:
    if SelectedItems.Count < knownItemIndex Then    ' try selected items next:
    # DoVerify False, " caller uses index 2, but there is only one"
    if SelectedObjects Is Nothing Then:
    # aItmIndex = -1
    else:
    # Call GetSelectedItems(SelectedObjects)
    # GoTo useSelected
    else:
    # useSelected:
    if aItmIndex = 0 And SelectedItems.Count > 0 Then:
    # Set GetAobj = SelectedItems.Item(1)
    # aItmIndex = 1
    else:
    if aItmIndex = 0 Then:
    # Set GetAobj = SelectedItems.Item(px)
    # aItmIndex = px
    else:
    # Set GetAobj = SelectedItems.Item(knownItemIndex)
    # aItmIndex = knownItemIndex
    # Set WorkItem(aItmIndex) = GetAobj
    if GetAobj Is Nothing Then:
    # DoVerify False, "this path should not be taken: why use ActiveExplorer ???"
    if ActiveExplorerItem(aPindex) Is Nothing Then:
    # ' make sure we have some item (GetAobj)
    # ' to determine the item type and Folder
    if ActiveExplorerItem(1) Is Nothing Then:
    if ActiveExplorer Is Nothing Then:
    # DoVerify False
    else:
    # Set Folder(px) = ActiveExplorer.CurrentFolder
    if Folder(px).Items.Count = 0 Then:
    # DoVerify False
    # Set GetAobj = Folder(px).Items(1)
    else:
    # Set GetAobj = ActiveExplorerItem(aPindex)
    # Set WorkItem(aItmIndex) = GetAobj
    else:
    if aItmIndex > sortedItems(px).Count Then:
    # aItmIndex = -1
    # GoTo ProcReturn
    if aItmIndex = 0 Then:
    # figure_as_selection:
    if LenB(ActiveExplorerItem(px)) = 0 Then:
    if ActiveExplorer.Selection.Count >= px Then:
    # Set ActiveExplorerItem(px) = _
    # ActiveExplorer.Selection.Item(px)
    else:
    # aItmIndex = -1
    # GoTo ProcReturn
    # Set GetAobj = ActiveExplorerItem(px)
    # aItmIndex = px
    else:
    # Set GetAobj = sortedItems(px).Item(knownItemIndex)
    # Set WorkItem(aItmIndex) = GetAobj
    else:
    if GetAobj Is Nothing Then                          ' we use existing value::
    # Set GetAobj = aID(px).idObjItem
    if GetAobj Is Nothing Then                          ' use selected item(1):
    # aItmIndex = 0
    # GoTo useSelected
    # Set WorkItem(aItmIndex) = GetAobj

    # DoVerify Not GetAobj Is Nothing, "no Object determined"
    if knownItemIndex >= 0 Then                             ' special case, do not set aID for -1:
    if aID(px) Is Nothing Then:
    # Set aID(px) = New cItmDsc
    # GoTo DefineIt
    if Not aID(px).idObjItem Is GetAobj Then            ' Different Item needs own object description:
    # DefineIt:
    # Call DefObjDescriptors(GetAobj, px, withValues:=True)
    # knownItemIndex = aItmIndex

    # Set ActItemObject = GetAobj                             ' set as global value

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function DecodeObjectClass
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def decodeobjectclass():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.DecodeObjectClass"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim DescriptorIndex As Long

    # DescriptorIndex = EvaluateSpecialRequirements(ActItemObject)
    # ' ActItemObject is now always the standard object class on aIndex
    # '   aIndex+2-> may address special objects (SpecialRequirements for exception or occurrence items)

    if ActItemObject Is Nothing Then:
    # DecodeObjectClass = "-"
    # GoTo ProcReturn

    # Call DefObjDescriptors(ActItemObject, aPindex, withValues:=getValues)
    # DecodeObjectClass = aOD(aPindex).objItemClassName

    # fiMain(aPindex) = GetMainObjectIdentification

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetMainObjectIdentification
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getmainobjectidentification():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetMainObjectIdentification"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if aPindex < 3 Then ' MainObjectIdentification only for Standard Objects:
    if LenB(MainObjectIdentification) = 0 Then     ' try existing:
    # MainObjectIdentification = aOD(aPindex).objDefaultIdent
    if LenB(MainObjectIdentification) = 0 Then     ' try global default:
    # MainObjectIdentification = MostImportantProperties(0)
    if StrComp(MainObjectIdentification, "none", vbTextCompare) = 0 Then:
    # MainObjectIdentification = vbNullString   ' never valid
    if LenB(MainObjectIdentification) = 0 Then     ' try general rules:
    # Call aID(aPindex).UpdItmClsDetails(ActItemObject)
    # aBugVer = LenB(MainObjectIdentification) > 0
    if DoVerify(aBugVer, "no main identification") Then:
    # aOD(aPindex).objDefaultIdent = MainObjectIdentification
    # GetMainObjectIdentification = getPropertyValue(MainObjectIdentification)

    # fiMain(aPindex) = ReFormat(GetMainObjectIdentification, vbCrLf, "|", b)
    if LenB(GetMainObjectIdentification) = 0 Then:
    # fiMain(aPindex) = "## " & aOD(aPindex).objDefaultIdent & " is empty"
    else:
    # fiMain(aPindex) = ReFormat(GetMainObjectIdentification, vbCrLf, "|", b)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function getPropertyValue
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getpropertyvalue():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.getPropertyValue"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim thisProp As ItemProperty

    if LenB(PropName) > 0 Then                      ' only such value can exist:
    # Set thisProp = LookUpAttrName(PropName)     ' look up in aID(apindex) -> aTD
    if thisProp Is Nothing Then:
    # getPropertyValue = vbNullString
    else:
    # Call Try(allowNew)                          ' Try anything, autocatch, Err.Clear
    # getPropertyValue = thisProp.Value
    # Call ErrReset(0)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PrepDecodeProp
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def prepdecodeprop():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.PrepDecodeProp"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim j As Long
    # Dim i As Long
    # Dim iClass As OlObjectClass
    # Dim vValue As Variant
    # Dim rateDelta As Long
    # Dim aLink As Variant
    # Dim lObjItem As Object
    # Dim LinkPath1 As String
    # Dim linkPath2 As String
    # Dim attachmentFile As String
    # Dim thisAttachment As Attachment
    # Dim Item As Object

    # rateDelta = 0
    if aTD Is Nothing Then:
    # DoVerify False, "** aTD not available???"
    # aStringValue = "# aTD not available *"
    # GoTo ProcReturn
    # Set Item = aTD.adItmDsc.idObjItem

    # ProcessMore:
    # With aTD
    # aStringValue = .adDecodedValue
    # aBugVer = iRules Is .adRules
    if DoVerify(aBugVer, "design check iRules Is aTD.adRules ???") Then:
    # Set iRules = .adRules
    if iRules.clsNotDecodable.RuleMatches Then:
    # aStringValue = "# not decodable *"
    # .adOrigValDecodingOK = True                 ' Skipped
    # GoTo loggit
    elif iRules.clsNeverCompare.RuleMatches _:
    # Or SkipDontCompare _
    # Or Not AllProps Then
    # aStringValue = "# skipped on request *"
    # .adFormattedValue = aStringValue
    # .adOrigValDecodingOK = True                 ' Skipped
    # GoTo loggit
    else:
    # vValue = aStringValue
    if .adInfo.iAssignmentMode = 1 Or .adInfo.iArraySize < 0 Then  ' kein array:
    if Not .adOrigValDecodingOK Then:
    # MatchPoints(aPindex) = MatchPoints(aPindex) - 1 ' penalize
    if LenB(aStringValue) > 0 Then:
    # rateDelta = 2
    # MatchPoints(aPindex) = MatchPoints(aPindex) + rateDelta ' rate
    if rP(aPindex) Is Nothing Then:
    if .adName = "IsRecurring" Then:
    if .adDecodedValue Then:
    # Set rP(aPindex) = Item.GetRecurrencePattern
    # Call LogEvent("Recurring item with " _
    # & rP(aPindex).Exceptions.Count _
    # & " Exceptions")
    else:
    # Set rP(aPindex) = Nothing
    elif .adName = "MemberCount" Then:
    # DoVerify .adInfo.iArraySize = CInt(.adDecodedValue)
    # GoTo ArrayCase                  ' MemberCount announces array of ADInfo.iArraySize elements
    else:
    # Call CheckPhoneNumber           ' fone and fax numbers ?
    if ItsAPhoneNumber Then:
    # Call PhoneNumberNormalize(.adDecodedValue, aPindex)
    else:
    # aStringValue = ReFormat(aStringValue, vbCrLf, "|", b)
    else:
    # NormalizedPhoneNumber = vbNullString
    # ItsAPhoneNumber = False
    else:
    # ArrayCase:
    # DoVerify False, "needs redesign"
    # aStringValue = .adDecodedValue          ' we do not Show the rest here
    # rateDelta = 1
    # Set vValue = .adInfo.iValue
    if vValue Is Nothing Then:
    # GoTo loggit
    # iClass = vValue.Class
    # MatchPoints(aPindex) = MatchPoints(aPindex) + rateDelta ' rate
    # Call N_ClearAppErr
    # Call Try                            ' Try anything, autocatch
    if iClass = olActions Then:
    # aStringValue = aStringValue & ", " _
    # & LString(j & ": " & vValue.Item(j).Name _
    # & "= " & vValue.Item(j).Enabled, 30)
    elif iClass = -1 Then:
    if vValue.Name = "MemberCount" Then:
    # aStringValue = aStringValue & vbCrLf _
    # & LString("Member " & j _
    # & "= " & aID(aPindex).idObjItem.GetMember(j).Name, 30)
    elif iClass = olAttachments Then:
    # Set thisAttachment = vValue.Item(j)
    if aID(aPindex).idObjItem.Class = olContact _:
    # And vValue.Item(j).FileName = "ContactPicture.jpg" Then
    # aStringValue = aStringValue _
    # & vbCrLf & j & " (ContactPicture), " _
    # & "Size: " & thisAttachment.Size
    # Call addContactPic(aID(aPindex), thisAttachment)
    elif SaveAttachments Then:
    # attachmentFile = aPfad & DateId & "(" & j & ") " _
    # & thisAttachment.FileName
    # aStringValue = aStringValue _
    # & vbCrLf & j & ", Size: " & thisAttachment.Size _
    # & vbTab & "-> " & Quote(attachmentFile)
    # aBugTxt = "Save attachment" & attachmentFile
    # Call Try
    # vValue.Item(j).SaveAsFile attachmentFile
    if Not Catch Then:
    # Call LogEvent("saved attachment " & attachmentFile)
    else:
    # Call LogEvent("userrequest: attachment " & j & " not saved as file " _
    # & thisAttachment.FileName)
    elif iClass = olLinks Then:
    # Call ErrReset(0)
    if VarType(vValue.Item(i)) = vbString Then:
    # GoTo missedLink
    # Call Try(allowAll)
    # Set aLink = vValue.Item(i).Item
    # LinkPath1 = aLink.Parent.FullFolderPath
    # Set lObjItem = _
    # aNameSpace.GetItemFromID(aLink.EntryID)
    if Catch Then:
    # GoTo missedLink
    # linkPath2 = lObjItem.Parent.FullFolderPath
    if Catch Then:
    # missedLink:
    # aStringValue = aStringValue _
    # & vbCrLf _
    # & "       link " & i _
    # & " invalid, not found for " _
    # & vValue.Value(j)
    if DebugMode Or DebugLogging Then:
    # AttributeUndef(aPindex) = AttributeIndex
    # Call N_ClearAppErr
    elif LinkPath1 <> linkPath2 Then:
    # aStringValue = aStringValue _
    # & vbCrLf _
    # & "       link " & i & " points to different Folder: " _
    # & vbCrLf _
    # & "            " & lObjItem.Parent.FullFolderPath _
    # & " instead of " & vbCrLf _
    # & "            " & aID(aPindex).idObjItem.Parent.FullFolderPath _
    # & vbCrLf _
    # & "       for  " & vValue.Value(j)
    if DebugMode Or DebugLogging Then:
    # AttributeUndef(aPindex) = AttributeIndex
    else:
    # aStringValue = aStringValue & vbCrLf _
    # & "       found " & LString(j & ": " _
    # & Quote(vValue.Item(j)), 30) & b
    elif iClass = olRecipients Then:
    # aStringValue = aStringValue & ", " _
    # & Left("Recipient" _
    # & j & "= " & Quote(vValue.Item(j).Name), 30) & b
    else:
    # aStringValue = aStringValue & ", " _
    # & LString(j & ": " & Quote(vValue.Item(j)), 30) & b

    if Catch Then:
    # DoVerify False, "kann Property nicht auswerten"
    # Call frmErrStatus.fBeginTermination(True)
    # aStringValue = Replace(aStringValue, Chr(160), b)
    # loggit:
    # Call FormatAttrForDisplay
    if .adNr < 0 Then                   ' actually not required, but AllProps is set:
    # .adNr = .adDictIndex
    # Call AppendTo(MostImportantAttributes, .adKey, b)
    if StringMod Then               ' was it appended / if not, no need to split again:
    # MostImportantProperties = split(MostImportantAttributes)
    if LenB(.adKillMsg) > 0 Then:
    # aStringValue = .adKillMsg
    else:
    # aStringValue = .adDecodedValue
    # End With ' aTD

    # FuncExit:
    # Set Item = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FormatAttrForDisplay
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def formatattrfordisplay():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.FormatAttrForDisplay"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim TestBody As Variant

    # ' all sorts of output prettying:
    # ' Special Formatting for *body* (if not empty string)
    if PropertyNameX = "Body" _:
    # And LenB(aStringValue) > 0 Then
    # aBugTxt = "get body format"
    # Call Try(438)
    # TestBody = aID(aPindex).idObjItem.BodyFormat
    if Catch Then                               ' e.g. Contacts never have HTMLbody:
    # GoTo noHTMLbody

    if LenB(aStringValue) = 0 And TestBody <> olFormatHTML Then     ' somebody lied about it:
    # DoVerify False, "aStringValue="" impossible???"
    # aID(aPindex).idObjItem.BodyFormat = olFormatHTML ' causes conversion for HTML and RTF Bodies
    # ' GoTo reEvaluate   ' this should have fixed our problem
    # ' note: no loop because BodyFormat is now HTML
    # ' no idea if it works for olFormatRTF
    # noHTMLbody:
    # killStringMsg = vbNullString
    # StringsRemoved = vbNullString
    if LenB(Trim(aStringValue)) = 0 Then:
    # Exit For
    # aStringValue = RemoveWord(aStringValue, killWords.Item(i), ": ""|")
    if LenB(StringsRemoved) > 0 Then:
    # killStringMsg = Trim(Append(killStringMsg, _
    # StringsRemoved, vbCrLf))
    if DebugLogging Then:
    print(Debug.Print killStringMsg)

    if PropertyNameX <> aTD.adName Then:
    if LenB(aTD.adName) = 0 Then                        ' original item does not have:
    # aTD.adName = PropertyNameX                      ' IsRecurring or Exceptions or its sub-properties
    else:
    # DoVerify False, "design check ???"

    if aTD.adOrigValDecodingOK Then:
    # aTD.adFormattedValue = aTD.adDecodedValue
    else:
    # aTD.adFormattedValue = aStringValue

    # ' more formatting stuff
    if aPindex = 2 Then                                     ' shorten if necessary for display:
    if aID(aPindex).idAttrDict.Exists(aTD.adKey) Then   ' not yet in aID(aPindex) - dictionary:
    # ' will be added in this Sub below, take pattern from adecprop C(1)
    # Set aDecProp(1) = GetAttrDsc(aDecProp(2).adKey, Get_aTD:=False, FromIndex:=1)
    # ' If debugMode And aDecProp(1).attrPos = 0 Then
    # ' Debug.Assert False ' this is no fix!!!
    # aDecProp(2).adNr = aDecProp(1).adNr
    else:
    # Stop ' ???
    # Set aDecProp(1) = aID(1).idAttrDict(1).Items(aDecProp(2).adNr)
    # ' get corresponding item to side 1 (no Err, first time we sync)
    if aDecProp(1) Is Nothing Then:
    # DoVerify False, " aID(2).attrPos??? Clone??? ***"
    else:
    if aDecProp(2) Is Nothing Then:
    # ' get corresponding item to side 1 (no Err, first time we sync)
    # aCloneMode = withNewValues                  ' use ADItmDsc, rules etc, but not values
    # Set aDecProp(2) = aID(1).idAttrDict.Item(aDecProp(1).adKey).Clone()
    elif aDecProp(1).adNr <> aDecProp(2).adNr Then:
    # ' get any missing attributes from other side
    # '   (no Err, first time we sync)
    if DebugMode Then DoVerify False:
    # Call AppendMissingProperties(AttributeIndex, 1, 2)
    # Call AppendMissingProperties(AttributeIndex, 2, 1)
    # Dim p1 As String
    # Dim p2 As String

    # Call FirstDiff(aDecProp(1).adFormattedValue, _
    # aDecProp(2).adFormattedValue, _
    # p1, _
    # p2, _
    # 80, 30, "...", OneDiff_qualifier)
    if Left(aDecProp(1).adFormattedValue, 1) <> "{" Then ' can't shorten arrays:
    # aDecProp(1).adShowValue = p1                    ' parameters can not be byref and byval at the same time
    # aDecProp(2).adShowValue = p2                    ' :( so we must put p1/2 in between
    else:
    # aDecProp(1).adShowValue = aDecProp(1).adFormattedValue
    # aDecProp(2).adShowValue = aDecProp(2).adFormattedValue
    # aTD.adKillMsg = OneDiff_qualifier
    # OneDiff_qualifier = vbNullString
    if displayInExcel And (xDeferExcel Or xUseExcel) Then:
    # Call put2IntoExcel(aPindex, AttributeIndex + 1)
    else:
    # Set aDecProp(1) = aTD                               ' aPindex = 1 allways in this Proc
    # aTD.adShowValue = aTD.advValue
    # fm:

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : FindAttributeByName
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Find position in aID(aPindex).idAttrDict of attribute with CritPropName
# '---------------------------------------------------------------------------------------
def findattributebyname():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.FindAttributeByName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim logicalIndex As Long
    # Dim lTrueIndex As Long
    # Dim lpx As Long
    # Dim nAttrDsc As cAttrDsc

    # lpx = aPindex
    if aID(2) Is Nothing Then:
    # lpx = 1

    # ' most likely, the sequence in aID(1).idAttrDict is same as in aID(2).idAttrDict
    # ' idAttrDict (2) may not be OK, so check idAttrDict(1) first
    # Set nAttrDsc = aID(1).idAttrDict.Item(logicalIndex)
    # aBugVer = LCase(CritPropName) <> LCase(nAttrDsc.adName)
    if DoVerify(aBugVer, "? Order is off?" _:
    # & vbCrLf & logicalIndex & b & aID(1).idAttrDict.Item(logicalIndex).adName) Then

    # YleadsXby = logicalIndex - AttributeIndex   ' count out-of-order items
    # GoTo GotIt

    # ' did not find it aID(1).idAttrDict going forward, so look in reverse
    # Set nAttrDsc = aID(1).idAttrDict.Item(logicalIndex)
    if LCase(CritPropName) = LCase(nAttrDsc.adName) Then:
    # YleadsXby = AttributeIndex - logicalIndex   ' count out-of-order items in aID(1).idAttrDict
    # GotIt:
    # FindAttributeByName = logicalIndex
    # lTrueIndex = nAttrDsc.adDictIndex          ' position in dictionary
    # Set aTD = aID(1).idAttrDict.Item(Abs(lTrueIndex) - 1)
    if lpx = 2 And aID(lpx).idAttrDict.Count < logicalIndex Then:
    # lpx = 1
    if aID(lpx).idAttrDict Is Nothing Then      ' nothing there to clone:
    # Set aDecProp(lpx) = Nothing
    else:
    # Set aDecProp(lpx) = aID(lpx).idAttrDict.Item(logicalIndex)
    if lpx <> aPindex Then:
    # DoVerify False, "check this design ???"
    # ' If aDecProp C(aPindex) Is Nothing Then   ' will set up aDecProp C(px=2) by cloning
    # '     Set aDecProp C(aPindex) = New Collection
    # ' End If
    # aCloneMode = FullCopy               ' clone with old values ???
    # aPindex = aPindex
    # Set aTD = aDecProp(lpx).adictClone
    # ' aDecProp C(aPindex).Add aTD        ' fill (aPindex) with values from the other(lpx) side
    # ' aTD.ADNr = aDecProp C(aPindex).Count
    # Call Get_iRules(aTD)
    # GoTo ProcReturn
    # FindAttributeByName = 0                             ' nothing found at all

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetAttributeByName
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setattributebyname():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.SetAttributeByName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim propertyIndex As Long
    # Dim msg As String
    # aPindex = px
    if FindAttributeByName(1, CritPropName) > 0 Then:
    if VarType(PropValue) = vbString Then:
    # msg = aTD.adShowValue
    # msg = " value changed from " & Quote(msg) & "  to " & Quote(PropValue) & b
    # aTD.adShowValue = PropValue
    if andPropertyToo Then:
    try:
        # propertyIndex = aTD.adtrueIndex
        # Ci.ItemProperties.Item(propertyIndex).Value = PropValue
        # msg = msg & " (item property, too)"
        else:
        # DoVerify False
        # errenc:
        # msg = " could not be assigned to the item's property: " & Err.Description
        else:
        # msg = " not found, no value change attempted."
        print(Debug.Print "Attribute " & CritPropName & msg)

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# ' Find in ItemProperties and return Property's true index in ItemProperties
# '  will also deliver the curItemProp and aDictIndex (or Nothing/0 if not exists yet)
# '  '  '    ' i is known true Property index or -1 if unknown
# '  '  '        in which case we use ADName to find it
def findproperty():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.FindProperty"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim j As Long
    # Dim aDictItem As cAttrDsc
    # Dim withoutAdminData As Boolean

    # FindProperty = -1           ' no success so far
    if aObject Is Nothing Then  ' should not happen, try to deduce from other parms:
    # DoVerify aTD Is Nothing And curItemProp Is Nothing And i < 0, "impossible to do that", True
    # check_aTD:
    if aTD Is Nothing Then:
    if aID(aPindex).idAttrDict Is Nothing Then:
    # provideNew:
    # Set aTD = ProvideAttrDsc
    else:
    # aBugVer = aID(aPindex).idAttrDict.Exists(adName)
    if DoVerify(aBugVer, "not in dictionary or aTD, ADName=" & adName) Then:
    # GoTo buildAttrDsc
    else:
    if isEmpty(aID(aPindex).idAttrDict.Item(adName)) Then:
    # buildAttrDsc:
    # withoutAdminData = True
    # GoTo provideNew
    # Set aTD = aID(aPindex).idAttrDict.Item(adName)
    # aBugVer = aTD.adItem Is Nothing
    if DoVerify(aBugVer, "aID(" & aPindex & ") will not give us a correct aTD") Then:
    # GoTo FuncExit
    else:
    # Set aObject = aTD.adItem
    # GoTo got_aTD
    else:
    # got_aTD:
    if aTD.adName = adName Then:
    if aObject Is aTD.adItmDsc.idObjItem Then:
    # GoTo got_object
    else:
    # GoTo get_Object
    else:
    # aBugVer = LenB(adName) = 0
    if DoVerify(aBugVer, "ADName is empty: improbable") Then:
    # aBugVer = aID(aPindex).idAttrDict.Exists(aTD.adKey)
    if DoVerify(aBugVer, _:
    # "aTD is not matching by key") _
    # Then
    # aBugVer = aTD.adKey = aTD.adName
    if DoVerify(aBugVer, "wrong key/Name combination") Then:
    # Set aTD = aID(aPindex).idAttrDict.Item(aTD.adKey)
    else:
    # GoTo FuncExit
    else:
    # adName = aTD.adName
    # get_Object:
    # Set aObject = aTD.adItmDsc.idObjItem
    if aObject Is Nothing Then:
    # aBugVer = curItemProp Is Nothing
    if DoVerify(aBugVer, "impossible if aObject and curItemProp are Nothing") Then:
    # GoTo FuncExit
    # aBugVer = curItemProp.Name = adName
    if DoVerify(aBugVer, "curItemProp=" & curItemProp.Name _:
    # & " mismatches ADName=" & adName) Then
    # GoTo FuncExit
    else:
    # Set aObject = curItemProp.Parent.Parent
    # GoTo got_object
    else:
    # GoTo got_aTD
    else:
    # GoTo check_aTD

    # got_object:                                 ' also have aTD when we get here
    # Set aProps = aObject.ItemProperties
    if aTD.adtrueIndex > -1 Then:
    if aProps.Item(i) Is aTD.adItemProp Then:
    if curItemProp Is Nothing Then:
    # Set curItemProp = aTD.adItemProp
    else:
    # DoVerify curItemProp Is aTD.adItemProp, "murks ???"
    # GoTo FuncExit
    else:
    # DoVerify False, "aTD is not matching ???"
    # GoTo FuncExit

    # FindProperty = aTD.adtrueIndex
    # GoTo findVerify
    else:
    # aTD.adtrueIndex = apropTrueIndex

    if i < 0 Then ' Position unknown, use name only:
    # findNew:
    # DoVerify False, "this should not really be neccessary: unknown property position i<0 ???"
    # Set iRules = Nothing                             ' must determine again
    # Set curItemProp = Nothing
    # Set curItemProp = aProps.Item(j)
    if curItemProp.Name = adName Then:
    # FindProperty = j                        ' the Prop TrueIndex
    # apropTrueIndex = j
    # Set aProp = curItemProp
    if withoutAdminData Then:
    # GoTo FuncExit

    if Not aID(aPindex).idAttrDict Is Nothing Then:
    # With aID(aPindex).idAttrDict
    if .Exists(adName) Then:
    # Set aTD = .Item(adName)
    if aTD.adtrueIndex < 0 Then:
    # aTD.adtrueIndex = j     ' now we have true index
    # DoVerify aProp.Name = aTD.adName, _
    # "Property Name mismatch for aTD true index=" & aTD.adtrueIndex
    # End With ' aID(aPindex).idAttrDict

    if aTD Is Nothing Then                  ' try to get in established attributes:
    # Set aTD = GetAttrDsc(adName)       ' sets iRules if Rules already defined
    # GoTo findVerify                         ' aTD, iRules undefined, curItemProp is OK
    # Set aTD = Nothing
    # FindProperty = -1
    # Set curItemProp = Nothing
    # GoTo findVerify
    else:
    # Set curItemProp = aProps.Item(i)
    if LenB(PropertyNameX) > 0 And curItemProp.Name <> PropertyNameX Then:
    # DoVerify False, "curItemProp.Name <> PropertyNameX ???"
    # GoTo findNew
    # PropertyNameX = curItemProp.Name
    # adName = vbNullString   ' do not use to find
    # FindProperty = i
    # j = i                                           ' this would be the position if we loop

    # ' try to find in already established Attributes
    if LenB(adName) > 0 Then    ' use name::
    if aTD Is Nothing Then:
    # Set aTD = GetAttrDsc(adName)               ' sets iRules if Rules already defined
    else:
    # Set curItemProp = aTD.adItemProp
    # Set aProp = curItemProp
    # GoTo ProcReturn                             ' all is well
    else:
    if FindProperty >= 0 Then:
    # GoTo findVerify
    if aTD Is Nothing Then:
    # GoTo findNew        ' its not in attributes, get from Properties
    else:
    # With aTD
    if .adName = adName Then:
    # j = .adDictIndex
    if j > 0 Then:
    if aID(aPindex).idAttrDict.Exists(adName) Then:
    # DoVerify False, "code needed ???"
    # GoTo findVerify
    else:
    # Set curItemProp = Nothing
    # FindProperty = Abs(.adtrueIndex)
    # GoTo funex
    # End With ' aTD

    # ' found in ItemProperties
    # findVerify:
    if FindProperty < 0 Then                        ' if invalid: not in itemProperties either:
    if PropertyNameX <> "Links" Then            ' Links can always be missing:
    # DoVerify False, Quote(adName) & " not in itemProperties of " & TypeName(aObject)
    # Set aTD = Nothing
    # Set iRules = Nothing
    else:
    # Set aProp = curItemProp
    if apropTrueIndex < 0 Then:
    # apropTrueIndex = j
    if Not aID(aPindex).idAttrDict.Exists(PropertyNameX) Then:
    # GoTo WrongATD
    if Not aID(aPindex).idAttrDict.Item(PropertyNameX) Is aTD Then:
    # GoTo WrongATD
    if aTD Is Nothing Then  ' has no AttrDsc yet, make one:
    # WrongATD:
    # apropTrueIndex = FindProperty
    # Set aTD = NewAttrDsc(apropTrueIndex)   ' into aTD
    elif aID(aPindex).idAttrDict.Exists(PropertyNameX) Then:
    # DoVerify aID(aPindex).idAttrDict.Item(PropertyNameX) Is aTD Or aObjDsc.objMaxAttrCount = 0, _
    # "aTD mismatches idAttrDict for " & PropertyNameX & " of " & aID(aPindex).idObjDsc.objItemClassName
    if aID(aPindex).idAttrDict.Item(PropertyNameX).adtrueIndex <> apropTrueIndex Then:
    # aID(aPindex).idAttrDict.Item(PropertyNameX).adtrueIndex = apropTrueIndex
    elif aTD.adName <> aID(aPindex).idAttrDict.Item(j + 1).Item.adName Then:
    # DoVerify False, "it wasn't the right one *** Impossible, remove ???"
    # Set aTD = aID(aPindex).idAttrDict.Item(aTD.adName).Item
    # '***ElseIf aID(aPindex).odAttArray(j) Is Nothing Then
    else:
    # DoVerify False, "should not be reached ???"
    # DoVerify aProps(j).Name = aTD.adName, _
    # "error PropTrueIndex: aProps(" & j & ") = " _
    # & aProps(j).Name & " <> aTD.ADName " & aTD.adName
    # ' we have to set the odItemDict.item
    # '*** Set aID(aPindex).odAttArray(j) = aTD
    if LenB(adName) > 0 Then:
    # DoVerify aTD.adName = adName, "aTD.ADName <> ADName, this is extremely fishy"
    else:
    # adName = aTD.adName
    # funex:
    # With aTD
    # j = aID(aPindex).idAttrDict.Count - 1
    if j < .adNr Then ' no valid aID(aPindex).idAttrDict - entry:
    # ' aDecProp C(aPindex).Add aTD
    # j = j ' ??? aTD.ADNr = aDecProp C(aPindex).Count
    else:
    if .adNr > 0 And j >= .adNr Then:
    # ' If aDecProp C(aPindex).Item(.ADNr).ADName = ADName Then
    # '    Set aDecProp(aPindex) = aDecProp C(aPindex).Item(.ADNr)
    # ' Else
    # '    Set aDecProp(aPindex) = Nothing
    # ' End If
    else:
    if Not aTD Is aDecProp(aPindex) Then:
    # DoVerify False, "look into this ???"
    # Set aDecProp(aPindex) = Nothing
    # aBugVer = .adItemProp Is aProp
    if DoVerify(aBugVer, "design check aTD.ADItemProp Is aProp ???") Then:
    # Set .adItemProp = aProp  ' forces consistency
    # Set curItemProp = .adItemProp
    # aBugVer = iRules Is .adRules And Not .adRules Is Nothing
    if DoVerify(aBugVer, "design check iRules Is .adRules And Not .adRules Is Nothing ???") Then:
    # Set iRules = .adRules
    # End With ' aTD

    # FuncExit:
    # Set aDictItem = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FormatPhoneNumber
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def formatphonenumber():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.FormatPhoneNumber"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    if ItsAPhoneNumber Then:
    # FormatPhoneNumber = aNumber
    # usermodified:
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "00", "+")
    # FormatPhoneNumber = Replace(FormatPhoneNumber, ")", b)
    # i = InStr(FormatPhoneNumber, "+")
    if i = 0 Then:
    if InStr(FormatPhoneNumber, "(0") = 1 Then:
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "(0", "+49 ")
    elif InStr(FormatPhoneNumber, "0") = 1 Then:
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "0", "+49 ")
    if InStr(i + 1, FormatPhoneNumber, "+") > 0 _:
    # Or i > 1 Then
    # With frmStrEdit
    # .Caption = fiMain(1)

    # .StringModifierCancelLabel.Caption = "alt:"
    # .StringModifierCancelValue.Text = FormatPhoneNumber
    # .StringModifierExpectation = "Format der Telefonnummer bitte korrigieren"
    # .StringToConfirm = FormatPhoneNumber
    # .Show
    if .StringModifierRsp <> 0 Then:
    # FormatPhoneNumber = .StringToConfirm
    # GoTo usermodified
    # End With ' frmStrEdit
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "(0", b)
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "/", b)
    # FormatPhoneNumber = Replace(FormatPhoneNumber, "-", b)
    if Len(FormatPhoneNumber) < 5 Then:
    if InStr(FormatPhoneNumber, "*") = 0 Then:
    # FormatPhoneNumber = "*" & FormatPhoneNumber
    else:
    # FormatPhoneNumber = aNumber

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' get iRules and aTd:
# ' Get_aTD=False: wont get aTd, iRules etc., just check and use
def getattrdsc():
    # Optional Get_aTD As Boolean = True, _
    # Optional FromIndex As Long = 0) As cAttrDsc
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetAttrDsc"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim aKey As String

    if Mid(PropName, 2, 2) = "==" Then          ' ignore headline for group of attrs:
    # GoTo FuncExit

    # aKey = GetAttrKey(PropName, Not Get_aTD, FromIndex)

    if aTD Is Nothing Then                      ' there is a aKey defined in odItemDict:
    if aID(FromIndex).idAttrDict.Exists(aKey) Then:
    if isEmpty(aID(FromIndex).idAttrDict.Item(aKey)) Then:
    # Set aTD = New cAttrDsc
    if Not iRules Is Nothing Then:
    if Not iRules.RuleObjDsc Is Nothing And Not aObjDsc Is Nothing Then:
    if iRules.RuleObjDsc.objClassKey <> aObjDsc.objClassKey Then:
    # Set iRules = Nothing        ' not defined yet, no atd and no iStuff
    # iRuleBits = "(void)"
    # GoTo FuncExit
    else:
    # aBugVer = InStr(aTD.adKey, PropName) > 0
    # aBugTxt = "PropertyName is at least part of adKey ???"
    # DoVerify
    if aTD.adRules Is Nothing Or aTD.adRuleIsModified Then:
    if isSpecialName Then               ' raw rules only; why ??? ***:
    # Set GetAttrDsc = aID(FromIndex).idAttrDict.Item(PropName).Item
    # Set aTD.adRules = GetAttrDsc.adRules
    else:
    # Call CreateIRule(PropName)

    # Call Get_iRules(aTD)

    # Set GetAttrDsc = aTD                        ' deliver aTD (see Get_aTD)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function LookUpAttrName
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def lookupattrname():

    # Const zKey As String = "ItemOpsOL.LookUpAttrName"
    # Call DoCall(zKey, tFunction, eQzMode)

    # ' PropName is some MainIdent, but may not exist (yet)
    # With aID(aPindex).idAttrDict
    if .Exists(PropName) Then:
    # Set aTD = .Item(PropName)
    # Set LookUpAttrName = aTD.adItemProp
    else:
    # DoVerify False
    # Set LookUpAttrName = Nothing
    # End With ' aID(aPindex).idAttrDict

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub GetSelectedItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getselecteditems():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetSelectedItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim itm As Object
    # Dim vItm As Variant
    # Dim i As Long

    # ' SelectedItems collection does not have to be empty, but ...
    if SelectedItems Is Nothing Then:
    # Set SelectedItems = New Collection
    for vitm in s:
    if Not isEmpty(vItm) Then:
    if Not vItm Is Nothing Then:
    # Set itm = vItm  ' convert to object
    if ItemDateFilter(itm, logLvl) = vbNo Then:
    # GoTo nextOne
    # SelectedItems.Add itm
    # i = i + 1
    # ' determine parent Folder of item (makes sense only for first 2 items)
    if i < 3 Then:
    if Folder(i) Is Nothing Then:
    # Set Folder(i) = getParentFolder(itm)
    # ' Set topFolder = getDefaultFolderType(s) ??? *** makes no sense at this time

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub BestObjProps
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get best Object, its Object Description, Object details with time, and Rules
# '          if NewObjectItem is not Nothing, it is used as aItmObject
# '          else default it from curFolder, ActiveExplorer, SelectedItems, SortedItems
# '          if no previous class Description exists, build one
# '---------------------------------------------------------------------------------------
def bestobjprops():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.BestObjProps"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim px As Long
    # Dim msg As String
    # Dim aItemClass As OlObjectClass
    # Dim aItemClassName As String
    # Dim aClassKey As String
    # Dim Reused As String

    # Reused = "New"

    if Not Item Is Nothing Then            ' need to look for best Item?:
    if aObjDsc.objItemClass = Item.Class Then:
    # Set ActItemObject = Item
    # aItemClass = ActItemObject.Class
    # aClassKey = CStr(aItemClass)
    # GoTo gotOne

    if curFolder Is Nothing Then:
    # Set curFolder = ActiveExplorer.CurrentFolder
    if curFolder.Parent Is Nothing Then:
    # eOnlySelectedFolder = True   ' search Folder
    elif curFolder.Parent.Class = olNamespace Then:
    # eOnlySelectedFolder = False
    # GoTo noclass
    else:
    # eOnlySelectedFolder = True

    # DoVerify aPindex < 3, "BestObjProps can only work for non-extended items"
    # px = aPindex
    if px <= 0 Then:
    # px = 1
    # GoTo takefirst
    else:
    if px <= 0 Then:
    # px = 1
    # takefirst:
    if SelectedItems Is Nothing Then:
    # Set ActItemObject = curFolder.Items(1)
    # aClassKey = CStr(ActItemObject.Class)
    # GoTo gotActItem
    elif SelectedItems.Count = 0 Then:
    if curFolder.Items.Count = 0 Then:
    # DoVerify False, "no items in " & curFolder.FolderPath _
    # & " or selected items. Will continue guessing"
    # GoTo guess
    else:
    # Set ActItemObject = curFolder.Items(1)
    # GoTo gotActItem
    elif SelectedItems.Count >= px Then:
    # Set ActItemObject = SelectedItems.Item(px)
    # gotActItem:
    if Not ActItemObject Is Nothing Then:
    # aClassKey = CStr(ActItemObject.Class)
    # Set Item = ActItemObject
    # GoTo gotOne
    else:
    # GoTo noclass
    else:
    # noclass:
    # DoVerify False, "are you trying to invalidate ObjDesc?"
    # aItemClass = -1
    # aItemClassName = vbNullString

    if aItemClass = 0 Then                                  ' guess/try better default:
    try:
        if SelectedItems Is Nothing Then:
        if sortedItems(px) Is Nothing Then:
        # GoTo guess
        else:
        if sortedItems(px).Count > 0 Then:
        # Set ActItemObject = sortedItems(px).Item(1)
        # aItemClass = ActItemObject.Class
        # aClassKey = CStr(aItemClass)            ' top sorted determines aItemClass
        # GoTo gotOne
        else:
        if SelectedItems.Count > 0 Then:
        # Set ActItemObject = SelectedItems.Item(1)
        # aItemClass = ActItemObject.Class
        # aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
        # GoTo gotOne

        if curFolder.Items.Count > 0 Then:
        # aItemClass = curFolder.Items(1).Class
        elif Not aID(px) Is Nothing Then:
        if Not aID(px).idObjItem Is Nothing Then:
        # DoVerify aObjDsc.objItemClass = aID(px).idObjItem.Class, "Class Change!"
        # Set ActItemObject = aID(px).idObjItem
        # aItemClass = ActItemObject.Class
        # aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
        else:
        # guess:
        if curFolder Is Nothing Then:
        # Set curFolder = olApp.ActiveExplorer.CurrentFolder
        if curFolder.Items.Count = 0 Then:
        # Set curFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
        if curFolder.Items.Count = 0 Then:
        # DoVerify False
        # aItemClass = curFolder.Items.Item(1).Class
        # aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
        # ' if we ProcCall into a new kind of Folder,
        # ' forget previous (Additional) Rule.clsObligMatches.aRuleString
        # gotOne:
        if px <= 0 Then:
        # px = 1
        # aItmIndex = WorkIndex(px)
        # Call GetITMClsModel(ActItemObject, px)
        # Call aItmDsc.SetDscValues(Item, withValues:=withValues, aRules:=sRules)

        if DftItemClass <> aObjDsc.objItemClass Then:
        # ExtendedAttributeList = vbNullString
        if sRules Is Nothing Then:
        # GoTo getRule
        else:
        # getRule:
        if sRules Is Nothing Then:
        if UserRule Is Nothing Then:
        if D_TC.Exists(aClassKey) Then:
        # Set sRules = D_TC.Item(aClassKey).objClsRules
        else:
        if UserRule.ARName = aObjDsc.objItemClassName Then:
        # Set sRules = UserRule
        # msg = "Re-using UserRule without changes, class " _
        # & aObjDsc.objItemClassName _
        # & " Type " & aObjDsc.objTypeName
        # Reused = "Reused"
        # GoTo FuncExit
        else:
        if LenB(aItemClassName) = 0 Then:
        # aItemClassName = aObjDsc.objItemClassName
        else:
        # DoVerify aItemClassName = aObjDsc.objItemClassName, "class name in aObjDsc messed up"
        if sRules.ARName = aObjDsc.objTypeName Then:
        # msg = "Re-using sRules without changes, class " _
        # & aObjDsc.objItemClassName _
        # & " Type " & aObjDsc.objTypeName
        # Reused = "Reused"
        # GoTo FuncExit    ' same class as before: optimize redundant work

        # ' this is the default for all of them, but deltas are ok
        # ' derive from Folder defaultitem class
        # DftItemClass = aObjDsc.objItemClass

        # ' sRules and DftItemTypeName and objItemClassName determined

        # ' Call aItmDsc.UpdItmClsDetails(ActItemObject)
        # msg = "BestObjProps says: " & Reused & " sRules for class " _
        # & aObjDsc.objItemClassName _
        # & " Type " & aObjDsc.objTypeName _
        # & " ItemObject: " & ActItemObject.Subject _
        # & ", Folder: " & curFolder.FolderPath

        # FuncExit:
        # Call LogEvent(msg, eLmin)

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub logDecodedProperty
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def logdecodedproperty():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "ItemOpsOL.logDecodedProperty"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim showThisDetail As String

    # showThisDetail = aItmIndex & ", Index=" & RString(AttributeIndex, 3) & diffStr _
    # & IgString & PropertyNameX _
    # & "="
    if Len(p1) > 256 Then:
    # showThisDetail = showThisDetail & Quote(Replace(Left(p1, 256), vbCrLf, vbCrLf & "        ")) _
    # & vbCrLf & "        ... (cut at 256 of " & Len(p1) & ")"
    else:
    # showThisDetail = showThisDetail & Quote(Replace(p1, vbCrLf, vbCrLf & "        "))
    # AllDetails = AllDetails & showThisDetail & vbCrLf
    # Call LogEvent("DBg: Item=" & showThisDetail, eLnothing)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub DefObjDescriptors
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Define or use existing Object Descriptor for MapiObject
# '---------------------------------------------------------------------------------------
def defobjdescriptors():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.DefObjDescriptors"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' keep consistent with Sub BestObjProps *** *** ???
    # ' ====================================================
    # ' objitem As Object                 current objects
    # ' aID(1 To 2) As cObjDsc            corresponding set of properties
    # ' aDecProp c(1 To 2) As Collection   of Class cAttrDsc
    # ' D_TC As Dictionary                of all Object Descriptors

    # Dim aClassKey As String
    # Dim SD As String
    # Dim specialTypName As String

    # aPindex = px                            ' 1..2, modified for subtypes of appointment types: 3..4
    # targetIndex = aPindex
    # sourceIndex = targetIndex - 1           ' using this for cloning

    # aBugTxt = "Get object Descriptor from Class"
    # Call Try
    # aClassKey = CStr(Item.Class)            ' basic name without subtype
    if Catch Then:
    # Set Item = Nothing
    # GoTo ProcReturn
    # Call GetITMClsModel(Item, aPindex)      ' may change aObjDsc

    if px < 3 Then                          ' make or Clone item subtypes:
    if Not aObjDsc Is aOD(px) Then:
    # Set aOD(px) = aObjDsc
    else:
    # Stop ' ??? !!! do some work from this to next ???
    # ' this code is needed to decode the subtypes of appointment types
    # ' all if-parts below should occur only if px>2, else is normal
    if Item.Class = olRecurrencePattern Then:
    # Set Item = GetAobj(px, -1)
    # specialTypName = DecodeObjectClass(getValues:=withValues)
    # TotalPropertyCount = 0          ' determine new situation, attribs for
    # ' recurrences +count of exceptions
    elif Item.Class = olException Then:
    # Set Item = GetAobj(px, -1).Exception.Item(1)
    # specialTypName = DecodeObjectClass(getValues:=withValues)
    else:
    # TotalPropertyCount = Item.ItemProperties.Count
    if LenB(specialTypName) = 0 Then:
    # SD = vbNullString
    else:
    # SD = GetObjectTypeExtension(Item)
    if aObjDsc Is Nothing Then          ' we have not had this object type before::
    # aOD(0).objMaxAttrCount = 0      ' clear previous class data
    # aOD(0).objDumpMade = -1
    # aID(0).idAttrCount = -1

    # Set aObjDsc = D_TC.Item(aClassKey)
    # Call aObjDsc.ODescClone(aClassKey & SD, aItmDsc)

    # DoVerify Not aObjDsc Is Nothing, "aObjDsc is Nothing???"

    # With aObjDsc
    # aBugVer = aOD(px).objItemClass = Item.Class
    if DoVerify(aBugVer, _:
    # "** aOD(px).objItemClass <> Item.Class") Then
    # aOD(px).objItemClass = Item.Class ' ????????? not a good idea!
    # Set aID(px).idObjItem = Item
    # aBugVer = AllPublic.SortMatches = .objSortMatches
    if DoVerify(aBugVer, _:
    # "** should have been set when creating aObjDsc ???") Then
    # .objSortMatches = SortMatches
    # End With ' aObjDsc

    if withAttributeSetup Then              ' (in/out: aID(px)) create dynamically or re-use:
    # Call SetupAttribs(Item, px, withValues)
    # Set aObjDsc = aOD(px)
    else:
    # aID(px).idEntryId = vbNullString

    # With aObjDsc
    # ' at this time, iRules is Nothing, or iRules.RuleInstanceValid  is always false

    # ' Superrelevant default, also sort rules from here, not changing for instance
    if aID(aPindex) Is Nothing Then     ' observe: aPindex may have changed from px:
    # DoVerify False, _
    # "aID( " & aPindex & ") must never be Nothing: remove ifpart if no hit"
    else:
    if LenB(aOD(aPindex).objDefaultIdent) = 0 Then:
    if sRules Is Nothing Then:
    # MainObjectIdentification = dftRule.clsObligMatches.CleanMatches(0)
    else:
    # .objDftMatches = Trim(sRules.clsObligMatches.aRuleString)
    if isEmpty(sRules.clsObligMatches.CleanMatches) Then:
    # MainObjectIdentification = dftRule.clsObligMatches.CleanMatches(0)
    else:
    # MainObjectIdentification = sRules.clsObligMatches.CleanMatches(0)
    else:
    # MainObjectIdentification = aOD(aPindex).objDefaultIdent

    if LenB(DontCompareListDefault) = 0 Then:
    # DontCompareListDefault = Trim(dftRule.clsNeverCompare.aRuleString)
    # ' very likely done already: just to make sure, change DontCompareListDefault
    # ' changes to Class-Specific Rules (for new classes only)
    # Call aID(aPindex).UpdItmClsDetails(Item)
    # .objDefaultIdent = MainObjectIdentification
    # End With 'aObjDsc

    # PropertyNameX = MainObjectIdentification
    if aID(aPindex).idAttrDict Is Nothing Then ' we have Not withValues:
    # GoTo ProcReturn
    # Set aTD = GetAttrDsc(PropertyNameX)    ' sets up iRules or uses existing
    if aTD Is Nothing Then:
    # GoTo ProcReturn

    # ' we can't do CheckAllRulesInList because only one iRules defined here
    if iRules Is Nothing Then:
    # Set iRules = aTD.adRules
    # Call iRules.CheckAllRules(PropertyNameX, "MainID: ")

    # 'On Error GoTo 0
    if BaseAndSpecifiedDiffer Then:
    if px < 3 Then                      ' =working on parent,:
    # DoVerify False, " we may never need this *** ???"
    # Call DefObjDescriptors(Item, px + 2, withValues:=False)
    # TotalPropertyCount = aID(aPindex).idAttrDict.Count - 1
    # aPindex = aPindex - 2           ' back to standard item
    # SpecialObjectNameAddition = vbNullString
    # Set aTD = GetAttrDsc(PropertyNameX) ' sets up iRules
    else:
    # DoVerify Item.ItemProperties.Count = TotalPropertyCount

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetObjectTypeExtension
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getobjecttypeextension():

    # Const zKey As String = "ItemOpsOL.GetObjectTypeExtension"
    # Call DoCall(zKey, tFunction, eQzMode)

    if Item.RecurrenceState = olApptException Then:
    # GetObjectTypeExtension = "#E"
    # aOD(aPindex).objNameExt = GetObjectTypeExtension
    elif Item.RecurrenceState = olApptOccurrence Then:
    # GetObjectTypeExtension = "#O"
    # aOD(aPindex).objNameExt = GetObjectTypeExtension
    else:
    # GetObjectTypeExtension = "#B" ' search for properties in "short" AppointmentItem
    # ' NO:: aID(aPindex).objNameExt = GetObjectTypeExtension

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub NameCheck
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def namecheck():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.NameCheck"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim ImpossibleNames As Variant
    # Dim AdelName As Variant
    # Dim hilf As String
    # Dim i As Long
    # ImpossibleNames = split("AG KG GmBH DE der v. Trier Trier-Land", b)
    # AdelName = split("von_der von_dem van_den van_der van_dem van " _
    # & "von vom vor zu zum zur an")
    if UCase(Ci.Lastname) = UCase(ImpossibleNames(i)) Then:
    if InStr(UCase(Ci.FullName), ". " & UCase(ImpossibleNames(i))) > 0 _:
    # Then
    # Call SetAttributeByName(Ci, 2, "FileAs", _
    # Replace(Ci.FullName, ". ", "."), True)
    # Call SetAttributeByName(Ci, 2, "FullName", Ci.FileAs, True)
    # Call SetAttributeByName(Ci, 2, "LastName", "###", True)
    elif LenB(Trim(Ci.FullName)) > 0 Then:
    # Call SetAttributeByName(Ci, 2, "FileAs", Ci.FullName, True)
    # Call SetAttributeByName(Ci, 2, "LastName", Ci.CompanyName, True)
    else:
    # Ci.FileAs = Ci.Firstname & b & Ci.Lastname
    # Ci.Lastname = Ci.Firstname & b & Ci.Lastname
    # Ci.Firstname = vbNullString
    # Ci.CompanyName = Ci.FileAs
    # GoTo DidMod
    elif InStr(UCase(Ci.FullName), ". " & UCase(ImpossibleNames(i))) > 0 Then:
    # Ci.FileAs = Ci.Firstname & ", " & Ci.Lastname
    # Ci.Lastname = Ci.FileAs
    # Ci.Firstname = "###"
    # Ci.CompanyName = Ci.FileAs
    # GoTo DidMod
    elif InStr(UCase(Ci.Lastname), b & UCase(ImpossibleNames(i))) > 0 _:
    # And InStr(UCase(Ci.FullName), UCase(Ci.Lastname)) > 0 _
    # Then
    # Ci.Lastname = "###"
    # Ci.Firstname = "###"
    # Ci.CompanyName = Ci.FileAs
    # GoTo DidMod
    elif Ci.Firstname = Ci.FullName _:
    # Or Ci.Lastname = Ci.FullName Then
    # Ci.Firstname = "###"
    # Ci.Lastname = "###"
    # Ci.CompanyName = Ci.FileAs
    # GoTo DidMod
    elif InStr(Ci.Firstname, ",") > 0 Then:
    # Ci.FileAs = Ci.FullName
    # hilf = Trunc(1, Ci.Firstname, ",")
    # Ci.Firstname = Ci.Lastname
    # Ci.Lastname = hilf
    # GoTo DidMod
    # hilf = Replace(AdelName(i), "_", b)
    if InStr(UCase(Ci.Lastname), UCase(hilf)) = 1 Then:
    # Ci.Lastname = Mid(Ci.Lastname, Len(hilf) + 1)
    # Ci.FileAs = Ci.Lastname & ", " & Ci.Firstname _
    # & b & Ci.MiddleName & b & hilf
    # Ci.MiddleName = hilf
    # GoTo DidMod
    # GoTo ProcReturn
    # DidMod:
    # MPEchanged = True  ' and loop end

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub RpStackAndLog
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def rpstackandlog():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.RpStackAndLog"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' RecurrencePattern has no Itemproperties, do the ones we know of
    # Dim eXA As Long
    # Dim Except As Outlook.Exception
    # Dim rTypeshowValue As String
    # Dim aIDa As cItmDsc

    # Set aIDa = aID(px)                      ' add  additional properties

    match arP.RecurrenceType:
        case olRecursDaily:
    # rTypeshowValue = "Daily"
        case olRecursWeekly:
    # rTypeshowValue = "Weekly"
        case olRecursMonthly:
    # rTypeshowValue = "Monthly"
        case olRecursMonthNth:
    # rTypeshowValue = "every N Months"
        case olRecursYearly:
    # rTypeshowValue = "Yearly"
        case olRecursYearNth:
    # rTypeshowValue = "every N Years"
        case _:
    # rTypeshowValue = "# unknown " & arP.RecurrenceType

    # Call Try(allowAll)                   ' Try anything, autocatch, Err.Clear
    # Call StackPropertyAndLog(px, "==========", "Start of Recurrence Pattern")   ' 1
    # Call StackPropertyAndLog(px, "RecurrenceType", rTypeshowValue)              ' 2
    # aTD.adDecodedValue = arP.RecurrenceType  ' raw value restored
    # Call StackPropertyAndLog(px, "PatternStartDate", arP.PatternStartDate)      ' 3
    # Call StackPropertyAndLog(px, "PatternEndDate", arP.PatternEndDate)          ' 4
    # Call StackPropertyAndLog(px, "StartTime", arP.starttime)                    ' 5
    # Call StackPropertyAndLog(px, "Interval", arP.Interval)                      ' 6
    # Call StackPropertyAndLog(px, "Regenerate", arP.Regenerate)                  ' 7
    # Call StackPropertyAndLog(px, "NrOfExceptions", arP.Exceptions.Count)        ' 8
    # Call StackPropertyAndLog(px, "DayOfWeekMask", arP.DayOfWeekMask)            ' 9
    # Call StackPropertyAndLog(px, "DayOfMonth", arP.DayOfMonth)                  ' 10
    # Call StackPropertyAndLog(px, "MonthOfYear", arP.MonthOfYear)                ' 11
    # Call StackPropertyAndLog(px, "Instance", arP.Instance)                      ' 12
    # eXA = 0
    if arP.Exceptions.Count > 0 Then:
    # ExceptionProcessing = True
    for except in arp:
    # eXA = eXA + 1
    # Call StackPropertyAndLog(px, "==========" & eXA, _
    # "Recurrence Exception " & eXA)              ' ex 1
    # Call StackPropertyAndLog(px, "ExDeleted" & eXA, Except.Deleted)         ' ex 2
    # Call StackPropertyAndLog(px, "ExOriginalDate" & eXA, _
    # Except.OriginalDate)                        ' ex 3
    # Set Except = Nothing                                                ' IMPORTANT

    # FuncExit:
    # Catch
    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetPropertyList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setpropertylist():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.SetPropertyList"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim TimeIn As Variant

    try:
        if Not ExceptionProcessing Then             ' aTD is set correctly:
        # Set aTD = aID(aPindex).idAttrDict.Item(aPropName).Item
        if aTD.adName <> aPropName Then:
        # DoVerify False
        else:
        # GoTo reUse
        # newone:
        # reUse:
        if W Is Nothing Then:
        # GoTo gds
        if aOD(aPindex).objDumpMade < 1 Then ' omit when done:
        # gds:
        if DebugLogging Then:
        # TimeIn = Timer
        print(Debug.Print Format(TimeIn, "0#####.00"), _)
        # "Finding DescriptorStrings for " & PropertyNameX
        # Call SplitDescriptor(aTD)
        if DebugLogging Then:
        print(Debug.Print , Timer - TimeIn, "finished Finding DescriptorStrings for " _)
        # & PropertyNameX

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetupAttribs
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setupattribs():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.SetupAttribs"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim SD As String
    # Dim newPos As Long
    # Dim baseID As cItmDsc
    # Dim otherID As cItmDsc
    # Dim nMax As Long
    # Dim nMin As Long
    # Dim aClass As OlObjectClass
    # Dim aTypeName As String
    # Dim aClassKey As String
    # Dim BasePx As Long
    # Dim otherPx As Long
    # Dim thisObjDsc As cObjDsc
    # Dim thisItmDsc As cItmDsc
    # Dim Reused As String
    # Dim tracing As Boolean

    # tracing = True                      ' just checking new design ???
    # aPindex = px
    # BasePx = px
    if BasePx > 2 Then:
    # DoVerify False, "Design verify need for SD"
    # BasePx = BasePx - 2

    # aClassKey = CStr(Item.Class)
    # aTypeName = TypeName(Item)

    if aItmDsc Is Nothing Then:
    # GoTo NotGoodForReuse

    if aItmDsc.idEntryId = Item.EntryID Then:
    # Set aID(aPindex) = aItmDsc
    # DoVerify aBugVer, "corrected aID: using aItmDsc ???"
    # Reused = "EntryId <> aItmDsc.idEntryId"
    # aItmDsc.idEntryId = Item.EntryID
    # GoTo NotGoodForReuse
    else:
    # Reused = "EntryId failed in aItmDsc"
    # GoTo NotGoodForReuse
    if aItmDsc.idObjDsc Is Nothing Then:
    # Reused = Reused & "no idObjDsc"
    # GoTo NotGoodForReuse
    elif aID(aPindex).idEntryId <> aItmDsc.idEntryId Then:
    if LenB(aID(aPindex).idEntryId) = 0 Then:
    if LenB(aItmDsc.idEntryId) > 0 Then:
    # DoVerify False, "Design check ??? is reuse still possible?"
    # Reused = "aItmDsc reused "
    # Set aID(aPindex) = aItmDsc
    else:
    # Reused = "aItmDsc has no idEntryId, no reuse possible"
    # GoTo NotGoodForReuse
    else:
    # GoTo NotGoodForReuse
    elif aItmDsc.idObjDsc.objClassKey <> aID(aPindex).idObjDsc.objClassKey Then:
    # Reused = "Class changed from " & aItmDsc.idObjDsc.objClassKey & " to " & aID(aPindex).idObjDsc.objClassKey
    # GoTo NotGoodForReuse
    if aID(aPindex) Is Nothing Then:
    # DoVerify Not aID(aPindex) Is Nothing, "design check aID(aPindex) IsNot Nothing ???"
    # Set aID(aPindex) = aItmDsc
    # Reused = "using aItmDsc"
    else:
    if aID(aPindex).idEntryId <> Item.EntryID Then:
    # aBugVer = LenB(aID(aPindex).idEntryId) > 0
    if aBugVer Then:
    print(Debug.Print " in aID:    " & aID(aPindex).idEntryId)
    print(Debug.Print " aItmDsc:   " & aItmDsc.idEntryId)
    print(Debug.Print " Item:      " & Item.EntryID)
    # DoVerify aBugVer, "design check: corrected because aID(aPindex).idEntryId empty ???"
    # Set aID(aPindex) = aItmDsc
    # Set thisItmDsc = aID(aPindex)

    # Set thisObjDsc = findObjDsc(aClassKey, Reused)
    if thisObjDsc Is Nothing Then:
    # GoTo NotGoodForReuse
    # Set thisItmDsc.idObjDsc = aObjDsc

    # Set baseID = aID(BasePx)
    if BasePx = 1 Then:
    # otherPx = 2
    elif BasePx = 2 Then:
    # otherPx = 1
    elif isEmpty(thisItmDsc) Then                 ' 3 or 4 not decoded, but needed now:
    # GoTo NotGoodForReuse
    if aPindex = 3 Then:
    # Set thisItmDsc = aID(4)
    elif aPindex = 4 Then:
    # Set thisItmDsc = aID(3)

    # Set otherID = aID(otherPx)                      ' Item class can be decoded before,

    if baseID Is Nothing Then:
    # GoTo DontHaveBase
    if otherID Is Nothing Then:
    # DoVerify BasePx = 1, "why is otherId=Nothing for BasePx = 1 ???"
    # otherPx = 0

    if baseID.idObjItem Is Nothing Then             ' re-Use other Dictionary with new values:
    # DoVerify aPindex = 1, "design check ??? baseID.idObjItem (apindex=1?) is new?"
    # Set thisItmDsc.idObjDsc = thisObjDsc        ' Class defaults are "constants"
    # Set thisItmDsc.idObjItem = Item
    if Not otherID Is Nothing Then:
    # Set baseID.idAttrDict = otherID.IDictClone(True)    ' clone and get new values
    if DebugLogging Then:
    # DoVerify baseID.idObjItem Is Item, "Testing only during design ???"
    # Set baseID.idObjItem = Item             ' clone must point to Item to SetupAttribs for
    # Reused = " cloned ID " & otherID.idPindex
    # Set thisObjDsc = thisItmDsc.idObjDsc
    if otherPx <> BasePx Then:
    # Reused = Reused & ", alias for Item " & aPindex
    # GoTo FuncExit

    if baseID.idObjItem Is Item Then                ' but must also be same item:
    # aBugVer = thisObjDsc Is thisItmDsc.idObjDsc
    if DoVerify(aBugVer, "Design check thisObjDsc Is thisItmDsc.idObjDsc ???") Then:
    # Set thisObjDsc = thisItmDsc.idObjDsc
    if otherPx = 0 Then:
    # Reused = Reused & "|New Item, Rule " & aPindex & " is using default Class Rule"
    else:
    # Reused = Reused & "|Reuse Item " & otherPx
    if otherPx <> BasePx Then:
    # Reused = Reused & ", alias for Item " & aPindex
    if thisItmDsc.idAttrDict Is Nothing Then:
    # aBugVer = thisItmDsc.idAttrCount = 0
    # DoVerify aBugVer, "Design check thisItmDsc.idAttrCount = 0 ???"
    # Set thisItmDsc.idAttrDict = New Dictionary
    # Reused = Reused & "|new Dictionary"
    # GoTo NotGoodForReuse
    # GoTo FuncExit
    else:
    if baseID.idObjDsc.objClassKey = aClassKey Then:
    if thisItmDsc Is Nothing Then:
    # Set thisItmDsc = New cItmDsc        ' clone the baseId
    if baseID.idAttrCount < Item.ItemProperties.Count Then:
    # DoVerify False, "check what happens if property count has changed. The new one is larger"
    if baseID.idAttrCount > Item.ItemProperties.Count Then:
    # DoVerify False, "check what happens if property count has changed. The new one is smaller"
    # thisItmDsc.idAttrCount = baseID.idAttrCount
    # Set thisItmDsc.idObjDsc = baseID.idObjDsc
    # thisItmDsc.idEntryId = Item.EntryID ' and init thisItmDsc
    # Set thisItmDsc.idObjItem = Item
    # thisItmDsc.idPindex = aPindex
    if withValues Then:
    # Set thisItmDsc.idAttrDict = baseID.IDictClone(False)    ' re-Use Dictionary without values
    # Reused = " cloned ID w/o values" & baseID.idPindex
    # GoTo UseID
    else:
    # Set thisItmDsc.idAttrDict = Nothing
    # thisItmDsc.idAttrCount = inv
    # DoVerify True, "attribute setup results in Nothing-Dictionary because withValues=False ???"
    # GoTo noValues
    # Set thisItmDsc.idAttrDict = baseID.IDictClone(False)    ' re-Use Dictionary without values
    # Reused = " cloned ID w/o values" & baseID.idPindex
    # GoTo UseID
    else:
    # Set thisItmDsc = Nothing
    # Reused = Reused & "|base Key <> aClassKey, new thisItmDsc"
    # GoTo NotGoodForReuse

    # DontHaveBase:
    if isEmpty(thisItmDsc) Then:
    # GoTo NotGoodForReuse

    if Not thisItmDsc.idObjItem Is Item Then        ' check for re-usability:
    # Reused = Reused & "|new Item"
    # GoTo NotGoodForReuse

    if thisItmDsc.idAttrDict Is Nothing Then        ' entirely new object description:
    # Reused = Reused & "|no AttrDict"
    # GoTo nDsc
    if thisItmDsc.idAttrDict.Count < 2 Then         ' entirely new object description:
    # nDsc:
    # Reused = Reused & "|getting new rules"
    # Set sRules = Nothing                        ' new rules will be determined
    # GoTo NotGoodForReuse                        ' no point in cloning the original

    if withValues Then:
    # aBugVer = thisItmDsc.idAttrDict.Count >= thisItmDsc.idObjItem.ItemProperties.Count
    if DoVerify(aBugVer, _:
    # "design check thisItmDsc.idAttrDict.Count >= thisItmDsc.idObjItem.ItemProperties.Count ???") Then
    # GoTo NotGoodForReuse
    # aBugVer = thisItmDsc.idObjDsc Is thisObjDsc
    if DoVerify(aBugVer, "design check thisItmDsc.idObjDsc Is thisObjDsc check ???") Then:
    # GoTo NotGoodForReuse

    # GoTo FuncExit                                   ' for this Item Attributes have been setup with values

    # NotGoodForReuse:
    if aObjDsc Is Nothing Then:
    # DoVerify False, "design check, aObjDsc Is Nothing ??? (Called via GetItmClsModel)"
    # Set thisObjDsc = New cObjDsc
    # thisObjDsc.objClassKey = aClassKey
    # Set thisObjDsc.objClsRules = sRules         ' set if usable
    # D_TC.Add aClassKey, thisObjDsc              ' not checking for collision because it will cause deadly error anyway
    # Set aObjDsc = thisObjDsc
    else:
    # aBugVer = aClassKey = CStr(Item.Class) & aObjDsc.objNameExt
    # DoVerify aBugVer, "design check Item.Class = aClassKey"
    # Set thisObjDsc = aObjDsc

    if isEmpty(thisItmDsc) Or thisItmDsc Is Nothing Then:
    # NokItmDsc:
    # AllProps = True                             ' must decode all properties if new
    # Set thisItmDsc = New cItmDsc                ' uses cObjDsc
    # Set thisItmDsc.idObjItem = Item
    # Reusing:
    # Set thisItmDsc.idObjDsc = aObjDsc
    # Set aID(aPindex) = thisItmDsc               ' without any content
    # thisItmDsc.idPindex = aPindex
    # thisItmDsc.idEntryId = Item.EntryID
    else:
    # aBugVer = aClassKey = CStr(Item.Class) & aObjDsc.objNameExt
    if DoVerify(aBugVer, "design check Item.Class = aClassKey") Then GoTo NokItmDsc:

    # aBugVer = thisItmDsc.idObjDsc Is aObjDsc
    if DoVerify(aBugVer, "idObjDsc Is aObjDsc ???") Then GoTo NokItmDsc:

    # aBugVer = thisItmDsc.idObjItem Is Item
    if DoVerify(aBugVer, "idObjItem Is Item ???") Then:
    if aItmDsc.idAttrCount > 0 Then         ' try reusing dictionary (=structure) without values:
    # Set thisItmDsc.idAttrDict = aItmDsc.IDictClone(False)
    if thisItmDsc.idAttrDict.Count < 2 Then:
    # GoTo NokItmDsc
    else:
    # Reused = Reused & "|cloned AttrDict"
    # GoTo Reusing

    # aBugVer = aID(aPindex) Is thisItmDsc
    if DoVerify(aBugVer, "aID(aPindex) Is thisItmDsc ???") Then GoTo NokItmDsc:

    # aBugVer = thisItmDsc.idPindex = aPindex
    if DoVerify(aBugVer, "idPindex = aPindex ???") Then GoTo NokItmDsc:

    # aBugVer = thisItmDsc.idEntryId = Item.EntryID
    if DoVerify(aBugVer, "idEntryId = Item.EntryID ???") Then GoTo NokItmDsc:

    # aBugVer = aID(aPindex) Is thisItmDsc
    if DoVerify(aBugVer, "design check ???, could fail if index =3 or 4") Then:
    # Set aID(aPindex) = thisItmDsc                   ' without any values

    if baseID Is Nothing Then:
    # Set baseID = aID(BasePx)                        ' may be Nothing if new class
    else:
    # aBugVer = aID(aPindex) Is baseID
    # DoVerify aBugVer, "design check aID(aPindex) Is baseID ??? could fail if index =3 or 4"
    if Not BasePx = aPindex Then:
    # SD = thisObjDsc.objNameExt                      ' use non-default Class key, else: leave empty
    # aClassKey = CStr(Item.Class) & SD

    if thisObjDsc.objClsRules Is Nothing Then:
    # Reused = Reused & "|defaulting rules"
    # Set sRules = dftRule.AllRulesClone(ClassRules, thisObjDsc, withMatchBits:=False)
    if aOD(aPindex) Is Nothing Then:
    # Set aOD(aPindex) = aObjDsc
    if aOD(aPindex).objClsRules Is Nothing Then:
    # Call SetCriteria                                ' sRules evaluated for Class

    if thisItmDsc.idRules Is Nothing Then:
    # Set iRules = New cAllNameRules                  ' sets all specific rules to non-nothings
    # Call iRules.AllRulesCopy(InstanceRule, sRules, withMatchBits:=False)
    # Reused = Reused & "|copied Class=" & aObjDsc.objClassKey & " sRules as iRules"
    # Set thisItmDsc.idRules = iRules
    else:
    # Reused = Reused & "|reusing sRules"
    # Set iRules = thisItmDsc.idRules

    # aBugVer = thisItmDsc.idObjDsc Is thisObjDsc
    if DoVerify(aBugVer, "design check thisItmDsc.idObjDsc Is thisObjDsc next assignment is needed ???") Then:
    # Set thisItmDsc.idObjDsc = thisObjDsc            ' Link as parent
    # aBugVer = thisItmDsc.idObjItem Is Item
    if DoVerify(aBugVer, "design check thisItmDsc.idObjItem Is Item next assignment is needed ???") Then:
    # Set thisItmDsc.idObjItem = Item

    # thisItmDsc.idEntryId = thisItmDsc.idObjItem.EntryID
    # thisItmDsc.idTimeValue = 0                          ' this indicates that we did not set aID with values yet

    # UseID:
    # With thisItmDsc
    # DoVerify .idEntryId = .idObjItem.EntryID, " design check ???"
    if Not withValues Then:
    # GoTo noValues
    if .idAttrDict Is Nothing Then:
    # aBugVer = aPindex > 0
    if DoVerify(aBugVer, "for aPindex=0 there is no Dictionary or object item") Then:
    # GoTo FuncExit
    else:
    # Set .idAttrDict = New Dictionary        ' using item(0) as TypeClassName
    # .idAttrDict.Add aClassKey, thisItmDsc   ' the new parent: cItmDsc, NOT a cAttrDsc!!!!
    # DoVerify .idEntryId = .idObjItem.EntryID, "design check .idEntryId = .idObjItem.EntryID ???"
    else:
    if .idAttrDict.Count = 0 Then:
    # .idAttrDict.Add aClassKey, thisItmDsc       ' the new parent: cItmDsc, NOT a cAttrDsc!!!!
    if .idAttrDict.Count < 2 Then                       ' WithValues=False is ignored ???:
    # ' reset cloned value of clsNeverCompare to default
    if iRules.clsNeverCompare.aRuleString <> DontCompareListDefault Then:
    # iRules.clsNeverCompare.ChangeTo = DontCompareListDefault ' reset iRules, unusual case

    if TotalPropertyCount = 0 Then:
    # aBugTxt = "Get Class of Item"
    # Call Try
    # aClass = Item.Class
    # Catch
    match Item.Class:
        case olRecurrencePattern:
    # DoVerify False, "design check for olRecurrencePattern ??? "
    # Call AttrExtend("==========")
    # Call AttrExtend("RecurrenceType")
    # Call AttrExtend("DayOfMonth")
    # Call AttrExtend("PatternStartDate")
    # Call AttrExtend("PatternEndDate")
    # Call AttrExtend("StartTime")
    # Call AttrExtend("Interval")
    # Call AttrExtend("Regenerate")
    # Call AttrExtend("NrOfExceptions")
    # Call AttrExtend("DayOfWeekMask")
    # Call AttrExtend("MonthOfYear")
        case olException:
    # DoVerify False, "design check for olException ??? "
    # Call AttrExtend("----------")
    # Call AttrExtend("ExDeleted")
    # Call AttrExtend("ExOriginalDate")
        case _:
    # GoTo fillAuto
    else:
    # fillAuto:
    # nMax = thisObjDsc.objMaxAttrCount
    # nMin = thisObjDsc.objMinAttrCount
    if aPindex > 2 Then                 ' copy to base: make room:
    # DoVerify False, "old code, check if it still makes sense ???"
    # Set baseID = aID(aPindex - 2)
    # newPos = baseID.idAttrCount + 1
    # nMin = Max(thisObjDsc.objMinAttrCount, Item.ItemProperties.Count)
    # nMax = Max(nMin, thisObjDsc.objMaxAttrCount) ' max allowed index +1
    elif thisObjDsc.objMaxAttrCount <> Item.ItemProperties.Count Then:
    # nMin = Max(thisObjDsc.objMinAttrCount, Item.ItemProperties.Count)
    # nMax = Max(nMin, thisObjDsc.objMaxAttrCount) ' max allowed index +1

    if Not ShutUpMode Then:
    # Call LogEvent(Format(Timer, "0#####.00") & vbTab _
    # & aPindex & ". item, starting InitAttributeSetup on " _
    # & thisObjDsc.objTypeName & "(Class " & Item.Class & "), with " _
    # & thisItmDsc.idObjItem.ItemProperties.Count _
    # & " standard properties." _
    # & vbCrLf & vbTab & vbTab & vbTab _
    # & "Subject: " & Quote(thisItmDsc.idObjItem.Subject), eLall)

    # Call InitAttributeSetup(baseID, thisItmDsc)         ' few standard things, all Exception presets

    if withValues Or thisItmDsc.idTimeValue = 0 _:
    # Or (nMax <> thisObjDsc.objMaxAttrCount _
    # Or nMin <> thisObjDsc.objMinAttrCount) Then
    # Call GetItemAttrDscs(Item, aPindex)             ' loops Props, generates AttrDsc in .odAddrDict , may get values
    else:
    # aBugTxt = "design check ??? Check if Previous Call was skipped correctly"
    # DoVerify thisItmDsc.idTimeValue <> 0
    # ' Else: no need to set up Attribute Dictionary because .idAttrDict.Count < 2

    # TotalPropertyCount = .idAttrDict.Count - 1
    # ' ??? remove and check indentation, DoVerify TotalPropertyCount = .idAttrDict.Count - 1, _
    # "** mismatch .odItemDict.Count <> .idAttrDict.Count-1"
    # End With ' thisItmDsc
    # noValues:
    # Set iRules = Nothing ' no Attribute has been selected intentionally

    # FuncExit:
    # Set aItmDsc = baseID
    # Set aObjDsc = thisObjDsc
    if aOD(aPindex) Is Nothing Then:
    # aBugTxt = "** setting up new class ??? " & aObjDsc.objClassKey
    # DoVerify False, aBugTxt
    # Set aOD(aPindex) = aObjDsc
    else:
    if aObjDsc.objClassKey <> aOD(aPindex).objClassKey Then:
    # ' aBugVer = aObjDsc Is aOD(aPindex)
    # ' aBugTxt = "** design check aObjDsc Is aOD(aPindex) ??? OR just a class change " _
    # & aObjDsc.objClassKey & "/" & aOD(aPindex).objClassKey
    # ' If DoVerify(aBugVer, aBugTxt) Then
    # Set aOD(aPindex) = aObjDsc
    # ' End If
    if LenB(MainObjectIdentification) = 0 Then:
    # MainObjectIdentification = sRules.clsObligMatches.CleanMatches(0)
    # aObjDsc.objDefaultIdent = MainObjectIdentification
    # Set aID(aPindex) = thisItmDsc                       ' does not have to be = aITMDsc ???
    if aItmDsc Is Nothing Then:
    # Set aItmDsc = thisItmDsc
    # Set aOD(aPindex) = thisObjDsc

    # newPos = Item.ItemProperties.Count
    # aItmDsc.idAttrCount = newPos
    if aOD(aPindex).objMaxAttrCount < newPos Then:
    # aOD(aPindex).objMaxAttrCount = newPos
    if aOD(aPindex).objMinAttrCount = 0 Then:
    # aOD(aPindex).objMinAttrCount = newPos

    # Set baseID = Nothing
    # Set thisObjDsc = Nothing
    # Set thisItmDsc = Nothing
    # Call aItmDsc.UpdItmTime

    # ProcReturn:
    if tracing Then:
    print(Debug.Print Replace(Reused, "|", vbCrLf)    ' this line during design check Only ???)
    else:
    # Reused = vbNullString

    # Call ProcExit(zErr, Reused)
    # Recursive = False

    # ProcRet:
# '---------------------------------------------------------------------------------------
# ' Method : findObjDsc
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Find ObjDsc via key in D_TC or return Nothing
# '---------------------------------------------------------------------------------------
def findobjdsc():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "ItemOpsOL.findObjDsc"
    # #If MoreDiagnostics Then
    # Call DoCall(zKey, "Function", eQzMode)
    # #End If

    if D_TC.Exists(aClassKey) Then:
    # Set findObjDsc = D_TC.Item(aClassKey)       ' reuse object type description
    if Not sRules Is findObjDsc.objClsRules Then:
    # DoVerify False, "Design check ??? sRules Is'nt findObjDsc.objClsRules"
    # Set sRules = findObjDsc.objClsRules      ' may reuse Rules if same class (no clone)
    else:
    # Set findObjDsc = Nothing
    # msg = msg & "|not a previously known class"

    # zExit:
    # Call DoExit(zKey)

# ' Make sure     aProp properly set ! or aValueType is a non-object Variant
# ' NOTE: this will NOT decode the attribute unless it is scalar
# '---------------------------------------------------------------------------------------
# ' Method : ProvideAttrDsc
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Provide New cAttrDsc for any variant object; retrieve value if scalar
# '          does not provide iRules, uses aOD(0).objClsRules=dftRule as classRules
# '---------------------------------------------------------------------------------------
def provideattrdsc():
    # ' (Optional aValueType As Long = 0, Optional Name As String = vbNullString, Optional ByRef Value As Variant = Nothing) As cAttrDsc
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.ProvideAttrDsc"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim nAttrDsc As cAttrDsc

    # aBugVer = Not aProp Is Nothing And Not isEmpty(aProp)
    # DoVerify aBugVer, "aProp must be set", True
    # Set aProps = aProp.Parent
    # aBugVer = PropertyNameX = aProp.Name
    if DoVerify(aBugVer, "** change of PropertyNameX: " & PropertyNameX & " to " & aProp.Name & " ???") Then:
    # PropertyNameX = aProp.Name          ' PropertyNameX used for cAttrDsc_Initialize
    # aCloneMode = withNewValues              ' Model, Rules and optionally Attribute value(s)
    # Set nAttrDsc = New cAttrDsc
    # Call nAttrDsc.GetScalarValue            ' no parms: called without any info outside aProp/dftRule'

    # Set ProvideAttrDsc = nAttrDsc

    # FuncExit:
    # Set nAttrDsc = Nothing

    # ProcReturn:
    # Call ProcExit(zErr, ProvideAttrDsc.adShowValue)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AttrExtend
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Extend the Attributes of Class with Extension Properties
# '---------------------------------------------------------------------------------------
def attrextend():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.AttrExtend"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # PropertyNameX = Name
    # Call SetPropertyList(Name)
    # Call StackAttribute

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : MakeAttributeKey
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose:     ADKey is checked vs. Attrname, or it's computed from ADName
# '---------------------------------------------------------------------------------------
def makeattributekey():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "ItemOpsOL.MakeAttributeKey"
    # Call DoCall(zKey, "Function", eQzMode)

    # Dim subtype As String
    # Dim i As Long

    if aID(aPindex) Is Nothing Then:
    if aPindex > 1 Then:
    # Set aID(aPindex) = aID(1)
    # subtype = aOD(aPindex).objNameExt

    # MakeAttributeKey = PropertyNameX
    # i = InStr(adName, "#")                                  ' expecting #W2, #B, #O, #R
    if i > 0 Then    ' if called using key, correct this (out!):
    # subtype = Trim(Mid(adName, i, 2))
    if subtype = "#W2" Then:
    # isUserProperty = True
    # ' remove all indications of key usage
    # PropertyNameX = Trim(Replace(adName, subtype, vbNullString))
    else:
    # PropertyNameX = adName
    # DoVerify LenB(PropertyNameX) > 0

    if InStr(MakeAttributeKey, PropertyNameX) = 0 Then      ' key mismatch or empty key:
    if isUserProperty Then:
    # subtype = "#W2"
    elif isSpecialName _:
    # Or BaseAndSpecifiedDiffer And Not workingOnNonspecifiedItem Then
    # subtype = "#B"                                  ' search for properties in "short" AppointmentItem
    # isSpecialName = True
    # isUserProperty = False
    else:
    if aPindex > 2 Then                                 ' just to make sure:
    if aID(aPindex).idObjItem.RecurrenceState = olApptOccurrence Then:
    # DoVerify subtype = "#O"
    elif aID(aPindex).idObjItem.RecurrenceState = olApptException Then:
    # DoVerify subtype = "#R"
    # MakeAttributeKey = PropertyNameX & subtype              ' that's it folks

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetAttrKey
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getattrkey():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.GetAttrKey"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim trueindex As Long

    if FromIndex = 0 Then:
    # FromIndex = aPindex

    # isUserProperty = False
    # aBugVer = LenB(adName) > 0
    if DoVerify(Message:="LenB(ADName) > 0") Then:
    # Set aTD = Nothing
    # GetAttrKey = vbNullString
    # GoTo ProcReturn
    # GetAttrKey = MakeAttributeKey(adName)   ' check or make it

    if noget Then                           ' do not get (updated data for) aTd:
    if Not aTD Is Nothing Then:
    if aTD.adKey <> GetAttrKey Then ' but its the wrong one:
    # Set aTD = Nothing           ' so mark as invalid
    else:
    if aID(FromIndex).idAttrDict Is Nothing Then:
    # Set aTD = Nothing
    # GoTo ProcReturn
    # With aID(FromIndex).idAttrDict
    if .Exists(GetAttrKey) Then     ' determines dDitem. BUG: creates empty Item:
    if isEmpty(.Item(GetAttrKey)) Then:
    # Set aTD = Nothing
    else:
    # Set aTD = .Item(GetAttrKey)
    # trueindex = aTD.adtrueIndex
    if aTD.adisUserAttr Then:
    # GoTo isUattr
    else:
    if Not aTD Is Nothing Then:
    if aTD.isUserProperty Then:
    # isUattr:                                    ' this is guessing!
    # Call AppendTo(GetAttrKey, "#W2")
    # isUserProperty = True
    # Stop ' trueindex = aID(FromIndex).ItemPropFind(GetAttrKey)
    # End With ' aID(fromindex).odItemDict

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub StackAttribute
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stackattribute():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.StackAttribute"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim last1Index As String
    # Dim lastchar As String
    # Dim i As Long
    # Dim aKey As String
    # Dim nAttrDsc As cAttrDsc

    if aTD.adNr <= 0 Then                                   ' check if it is in aID(aPindex).idAttrDict already, true=yes:
    if aPindex > UBound(aID) Then                       ' ??? whot???:
    # GoTo ProcReturn                                 ' we don't need you
    if aID(aPindex) Is Nothing Then:
    # Set aID(aPindex).idAttrDict = New Dictionary
    # AttributeIndex = aID(aPindex).idAttrDict.Count - 1

    # ' find out if we have this decprop ( and ignore separator ' ========== )
    if InStr(aTD.adName, "=") = 0 Then:
    # Set nAttrDsc = aID(aPindex).idAttrDict.Item(aTD.adKey)
    # DoVerify nAttrDsc.adName = aTD.adName, "Error in ADName of aID(aPindex).idAttrDict at Position=" & AttributeIndex
    # DoVerify nAttrDsc.adKey = aTD.adKey, "Error in ADKey of aID(aPindex).idAttrDict at Position=" & AttributeIndex
    # ' something in this position present: check plausi: very bad if assert fails
    # aBugTxt = "Error in ADtrueIndex of aID(aPindex).idAttrDict at Position=" & AttributeIndex
    # DoVerify nAttrDsc.adtrueIndex = aTD.adtrueIndex
    # aKey = aTD.adKey

    if aPindex = 2 And Not aID(1).idAttrDict Is Nothing Then:
    # ' may have to synchronize names (with first object)
    if aID(1).idAttrDict.Count >= AttributeIndex Then:
    if aID(1).idAttrDict.Item(AttributeIndex).adName = PropertyNameX Then:
    # i = 0                                        ' OK, all is well, need no sync of names
    else:
    # ' If DebugMode Then
    # DoVerify False, "design check aID(1).idAttrDict.Count >= AttributeIndex ??? "
    # ' End If
    # i = FindAttributeByName(AttributeIndex, PropertyNameX) ' find in aID(1).idAttrDict
    if i > 0 Then:
    # Stop ' ???
    # AttributeIndex = i
    # ' If Not aTD.ADNr _
    # = aDecProp C(1).Item(i).ADNr _
    # Or Left(aTD.ADFormattedValue, 1) <> "*" Then
    # ' If Left(aTD.ADFormattedValue, 1) = "*" Then
    # ' pArr(1) = aDecProp(2).PropValue
    # ' Else                                ' could be a double name entry!
    # ' i = 1000 ' add at end, because > AttributeIndex always
    # ' End If
    # ' End If
    if i < aID(2).idAttrDict.Count Then:
    if aID(2).idAttrDict.Item(i).adFormattedValue _:
    # <> aTD.adFormattedValue Then
    if aID(2).idAttrDict.Item(i).adFormattedValue = Chr(0) Then:
    # aID(2).idAttrDict.Item(i).adFormattedValue _
    # = aTD.adFormattedValue
    else:
    # DoVerify False
    # pArr(2) = aTD.adFormattedValue
    else:
    # Stop ' ???
    # ' aDecProp C(2).Add aTD                ' add at end!
    # GoTo ProcReturn
    # '   ########

    # lastchar = Chr(0)                           ' inserting #2 into 1 !
    # ' If DebugMode Then
    # DoVerify False, " here we could mess up badly ???***"
    # ' check correct unique attributeindex and aID(2).idAttrDict
    # ' End If
    # cMissingPropertiesAdded = cMissingPropertiesAdded + 1
    # DoVerify False, " not tested! replace adecprop(1) with aTd ??? ***"
    # aCloneMode = FullCopy
    # Set nAttrDsc = New cAttrDsc
    # Set aDecProp(1) = nAttrDsc
    # Set nAttrDsc.adItem = aID(1).idObjItem
    # last1Index = aID(1).idAttrDict.Item(AttributeIndex - 1).adKey
    # lastchar = Right(last1Index, 1)
    if lastchar >= "a" Then:
    # lastchar = Chr(Asc(lastchar) + 1)
    # last1Index = Left(last1Index, Len(last1Index) - 1)
    else:
    # lastchar = "a"
    # nAttrDsc.adNr = last1Index & lastchar
    # nAttrDsc.adName = PropertyNameX
    # aDecProp(1).adFormattedValue = "***Missing Property*** key=" _
    # & last1Index & lastchar
    # Addit_Text = True
    # pArr(1) = PropertyNameX
    # pArr(2) = nAttrDsc.adFormattedValue
    elif aID(1).idAttrDict.Count < AttributeIndex Then:
    # ' have to synchronize names (with second object)
    # aCloneMode = FullCopy
    # aID(1).idAttrDict.Add aDecProp(2).adKey, aDecProp(2).adictClone        ' add at end!
    # GoTo ProcReturn

    # FuncExit:
    # Set nAttrDsc = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' used for Extended properties:
def stackpropertyandlog():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.StackPropertyAndLog"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # PropertyNameX = aName
    # aStringValue = AVal

    # Set aTD = Nothing                   ' always new AttrDsc
    # Call CreateIRule(aName)             ' using PropertyNameX inside, does SplitDescriptor to set Rules
    # aTD.adFormattedValue = AVal
    # aTD.adShowValue = AVal
    # aTD.adInfo.iAssignmentMode = 1
    # Call StackAttribute
    if InStr(aName, "=") > 0 Then:
    # aTD.adDecodedValue = vbNullString         ' Show sepline in excel
    else:
    # aTD.adDecodedValue = AVal
    # aTD.adOrigValDecodingOK = True

    # Call SetPropertyList(aName)

    if displayInExcel Then:
    # pArr(1) = PropertyNameX
    # pArr(1 + px) = aStringValue
    # Call addLine(O, AttributeIndex, pArr)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetITMClsModel
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get or Make the Object descriptor for an ObjectClass
# '          determine ClassName and TypeName from item's object class,
# '          determine ItemClassProperties (like MailLike, TimeType, has ReceivedTime, ...)
# '          determine parts of default Rule
# '---------------------------------------------------------------------------------------
def getitmclsmodel():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "ItemOpsOL.GetITMClsModel"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction, ExplainS:="GetITMClsModel")

    # Dim aItmClass As OlObjectClass
    # Dim aItmTypeName As String
    # Dim RecommendedMatches As String
    # Dim Reused As String
    # Dim InitialSetup As Boolean
    # Dim aClassKey As String
    # Dim inD_TC As Boolean

    # aItmTypeName = TypeName(Item)
    # aItmClass = Item.Class                              ' using this to get the ClassName
    # aClassKey = CStr(aItmClass)
    # Set aItmDsc = New cItmDsc
    # Set aItmDsc.idObjItem = Item

    if px = 0 Then:
    # px = 1
    if targetIndex = 0 Then:
    # targetIndex = px
    else:
    # DoVerify targetIndex = px, "targetIndex<>aPindex is not OK ???"

    if D_TC.Count = 1 Then:
    # InitialSetup = True
    if D_TC.Exists(aClassKey) Then:
    # inD_TC = True
    # Set aObjDsc = D_TC.Item(aClassKey)

    if aObjDsc.objClassKey = CStr(Item.Class) Then      ' correct item class ???:
    # Set aItmDsc.idObjDsc = aObjDsc

    if DoVerify(ItemValid(aItmDsc.idObjItem), "design check item gone/invalid ???") Then:
    # GoTo FuncExit
    if Not aOD(px) Is aObjDsc Then:
    # Set aOD(px) = aObjDsc
    # Reused = " Reused"
    # GoTo FuncExit
    else:
    # aBugTxt = "Design check: Wrong Item Class ???"
    # DoVerify False
    # Set aItmDsc = Nothing
    # Set aID(px) = Nothing
    # Set aDecProp(px) = Nothing
    # Set aID((px + 2)).idAttrDict = Nothing
    # Set aID((px)).idAttrDict = Nothing              ' reset all previously decoded values

    else:
    # DontCompareListDefault = Trim(dftRule.clsNeverCompare.aRuleString)
    # ' Creating new Class description model
    # Reused = "New"
    # Set aObjDsc = New cObjDsc


    # With aObjDsc
    # RecommendedMatches = "Subject"                      ' Superrelevant default, override OK
    # .objItemClass = aItmClass
    # .objItemClassName = aItmTypeName                    ' sometimes not correct for class, set in selected case
    # .objHasReceivedTime = False                         ' these are the default with EXCEPTIONS below
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = False
    # .objHasSentOnBehalfOf = False
    # Set .objSeqInImportant = New Collection

    # ' stops on unverified for classes 47, 49, 53-56, certain for 56, 57: .objHasHtmlBodyFlag = False
    match .objItemClass:
    # '----- very common ones are first ----
        case olMail                     ' 43:
    # .objIsMailLike = True
    # .objItemType = OlItemType.olMailItem
    # .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
    # .objHasHtmlBodyFlag = True
    # .objHasSenderName = True
    # .objHasSentOnBehalfOf = True
    # .objTimeType = "SentOn"

    # Call AppendTo(DontCompareListDefault, _
    # " HTMLBody " _
    # & "Responserequested ConferenceServerAllowExternal " _
    # & "SendUsingAccount SentOn Recipients " _
    # & "ReceivedByName CreationTime ", b)
    if LenB(aTimeFilter) = 0 Then:
    # RecommendedMatches = Append("Subject SenderName SentOn ", aTimeFilter, b)
        case olMeeting                  ' 1:
    # .objItemType = OlItemType.olAppointmentItem
    # ' create like this, but modify .Status = olMeeting
    # ' other .Status are
    # ' olNonMeeting (0)
    # ' olMeetingReceived (3)
    # ' olMeetingCancelled (5)
    # ' olMeetingReceivedAndCancelled (7)
    # .objHasSenderName = False
    # .objTimeType = "SentOn"

    # DoVerify False
    # ' *Time must be missing!
    # DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
    # Call AppendTo(DontCompareListDefault, _
    # "Ordinal CreationTime ConversationIndex Size SenderName" _
    # & "SentOn SentOnBehalfOfName", b)
    # RecommendedMatches = "Start End IsRecurring Exceptions"
    # ' no compare of Subject: time conflicts!
        case olAppointment          ' 26:
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasSenderName = True
    # .objTimeType = "Start"

    # Call AppendTo(DontCompareListDefault, _
    # "ConversationTopic", b)
    # RecommendedMatches = "Subject ! Start End IsRecurring Exceptions"
        case olContact              ' 40:
    # .objItemType = OlItemType.olContactItem
    # .objTimeType = "LastModificationTime"

    # Call AppendTo(DontCompareListDefault, _
    # " ConversationTopic LastFirst* Subject Initials", b)
    # RecommendedMatches = _
    # "%FullName ! FileAs LastName FirstName MiddleName MobileTelephoneNumber " _
    # & "HomeTelephoneNumber CompanyName BusinessTelephoneNumber BusinessFaxNumber " _
    # & "OtherTelephoneNumber Email1Address Email2Address Email3Address " _
    # & "WebPage HomeAddress BusinessAddress CompanyName Birthday User2 " _
    # & "HasPicture Links Attachments"
    # '----- uncommon ones  ----
        case olNote                     ' 44:
    # .objItemType = OlItemType.olNoteItem
    # .objTimeType = "LastModificationTime"
    # RecommendedMatches = "Subject Body"
        case olPost                     ' 45:
    # .objItemType = OlItemType.olPostItem
    # .objHasSenderName = True
    # .objTimeType = "LastModificationTime"
    # DoVerify False, " we never saw this. Recipient missing. .objHasSenderName is unverified"
    # ' RecommendedMatches = ???
        case olTask                     ' 48:
    # .objItemType = OlItemType.olTaskItem
    # .objHasSenderName = True
    # .objTimeType = "LastModificationTime"
    # DoVerify False, "watch further"
    # RecommendedMatches = "ConversationTopic Subject StartDate ! DueDate Body"
        case olTaskRequest              ' 49:
    # .objItemClassName = "TaskRequest"
    # .objItemType = OlItemType.olTaskItem
    # .objHasHtmlBodyFlag = False
    # .objIsMailLike = True
    # ' Not .objHasSenderName
    # ' Not .objHasSentOnBehalfOf
    # .objTimeType = "CreationTime"
        case olTaskRequestUpdate        ' 50:
    # .objItemClassName = "TaskRequestUpdate"
    # .objItemType = OlItemType.olTaskItem
    # .objIsMailLike = True
    # .objHasHtmlBodyFlag = True
    # .objHasSenderName = True
    # .objHasSentOnBehalfOf = True
    # .objTimeType = "SentOn"
        case olTaskRequestAccept        ' 51:
    # .objItemClassName = "TaskRequestAccept"
    # .objItemType = OlItemType.olTaskItem
    # .objIsMailLike = True
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
        case olTaskRequestDecline       ' 52:
    # .objItemClassName = "TaskRequestDecline"
    # .objItemType = OlItemType.olTaskItem
    # .objIsMailLike = True
    # .objHasHtmlBodyFlag = True
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
        case olMeetingRequest           ' 53:
    # .objItemClassName = "MeetingRequest"
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objHasSentOnBehalfOf = False
    # RecommendedMatches = "Start End IsRecurring Exceptions"
    # MeetingStuff:
    # .objIsMailLike = True
    # .objTimeType = "SentOn"
    # ' All *Time must be missing from Item Compares!
    # DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
    # Call AppendTo(DontCompareListDefault, _
    # "CreationTime ConversationIndex Size SenderName" _
    # & "SentOn", b)
        case olMeetingCancellation      ' 54:
    # .objItemClassName = "MeetingCancellation"
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objHasSentOnBehalfOf = True
    # .objTimeType = "SentOn"
    # GoTo MeetingStuff
        case olMeetingResponseNegative  ' 55:
    # .objItemClassName = "MeetingResponseNegative"
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
    # DoVerify False, "watch further"
    # GoTo MeetingStuff
        case olMeetingResponsePositive  ' 56:
    # .objItemClassName = "MeetingResponsePositive"
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
    # DoVerify False, "watch MeetingResponsePositive further"
    # GoTo MeetingStuff
        case olMeetingResponseTentative ' 57:
    # .objItemClassName = "MeetingResponseTentative"
    # .objItemType = OlItemType.olAppointmentItem
    # .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
    # .objHasHtmlBodyFlag = False
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
    # DoVerify False, "watch further"
    # RecommendedMatches = "Start End IsRecurring Exceptions"

    # gotNameIsMailLike:
    # .objIsMailLike = True
    # .objTimeType = "SentOn"
    # ' All *Time must be missing from Item Compares!
    # DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
    # Call AppendTo(DontCompareListDefault, _
    # "Ordinal CreationTime ConversationIndex Size SenderName" _
    # & "SentOn SentOnBehalfOfName", b)
    # ' no compare of Subject: time conflicts!
        case olSharing              ' 104:
    # .objItemType = OlItemType.olTaskItem
    # .objHasSenderName = True
    # .objTimeType = "SentOn"
    # RecommendedMatches = "Subject ! Body"
    # DoVerify False, "watch further"
    # GoTo gotNameIsMailLike
    # '----- unlikely ones  ----
        case olDocument             ' 41:
    # .objItemType = 41       ' OlItemType.olDocument is not defined!
    # .objTimeType = "SentOn"
    # DoVerify False, "we never saw this. Office Document or whatever"
    # ' RecommendedMatches = ???
        case olJournal              ' 42:
    # .objItemType = OlItemType.olJournalItem
    # .objTimeType = "SentOn"
    # DoVerify False, "we never saw this. Deprecated"
    # ' RecommendedMatches = ???
        case olReport               ' 46:
    # .objIsMailLike = True
    # .objItemType = 46       ' OlItemType.olReport is not defined!
    # .objHasReceivedTime = False                 ' this is one of the few EXCEPTIONS
    # .objHasHtmlBodyFlag = False
    # .objHasSentOnBehalfOf = False
    # .objHasSenderName = False
    # .objTimeType = "LastModificationTime"
        case olRemote               ' 47:
    # .objItemType = 47           ' OlItemType.olRemote is not defined!
    # .objTimeType = "SentOn"
    # DoVerify False, " we never saw this ."
    # ' Like mail, but no BillingInformation, Body, Categories, Companies, and Mileage
    # ' RecommendedMatches = ???
        case olDistributionList     ' 69:
    # .objItemType = OlItemType.olDistributionListItem
    # .objTimeType = "LastModificationTime"
    # Call AppendTo(RecommendedMatches, "! MemberCount", b)   ' ! means: do not sort
        case _:
    # Stop ' ???
    # Call LogEvent(.objItemClassName & "(" & .objItemClass _
    # & ") not expected with item " & aItmIndex _
    # & " in Folder " & curFolderPath & " is: ", eLall)
    # & .objItemClass & ", Try as mail?", vbYesNoCancel)
    if rsp = vbCancel Then:
    # End
    elif rsp = vbYes Then:
    # GoTo gotNameIsMailLike
    else:


    # MainObjectIdentification = vbNullString               ' Undefined
    # .objItemClassName = vbNullString

    # .objTypeName = Remove(.objItemClassName, "Item")
    # .objDftMatches = RecommendedMatches
    # .objClassKey = aClassKey

    # Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, False)
    # sRules.clsNeverCompare.ChangeTo = DontCompareListDefault
    # sRules.clsObligMatches.ChangeTo = RecommendedMatches
    # TrueCritList = sRules.clsObligMatches.CleanMatches(0)
    # MostImportantAttributes = Append(sRules.clsObligMatches.CleanMatchesString, sRules.clsSimilarities.CleanMatchesString, b)
    # MostImportantProperties = split(MostImportantAttributes)
    # Set .objClsRules = sRules                    ' this is only the starting point for sRules
    # DoVerify sRules.ARName = .objTypeName, " used to be an assignment, needed?"

    # Reused = "* New Object Class defined: Decoding values required for "
    # Call LogEvent(Reused & .objItemClassName & "(" & .objItemClass & "), " _
    # & "Maillike = " & .objIsMailLike & ", TimeType = " & .objTimeType _
    # & ", RecommendedMatches=" & RecommendedMatches, eLall)

    # End With ' aObjDsc

    if Not inD_TC Then:
    # D_TC.Add aClassKey, aObjDsc
    # inD_TC = True

    # FuncExit:
    # Set GetITMClsModel = aItmDsc
    # Set aID(aPindex) = aItmDsc
    if aOD(aPindex) Is Nothing Then:
    # Set aOD(aPindex) = aObjDsc

    # ProcReturn:
    # Call ProcExit(zErr, aObjDsc.objTypeName & Reused)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ReGet
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def reget():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.ReGet"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)


    # ' de-reference the old item, then open it again
    # ' otherwise Outlook might still have tidbits
    # ' left from the original message

    # Dim oldClass As OlObjectClass
    # Dim Retry As Long

    # E_Active.Permit = "*"
    # oID = oItem.EntryID
    if ErrorCaught <> 0 Then:
    # GoTo badItem
    # oldClass = oItem.Class
    # tryAgain:
    if Not oItem.Saved Then:
    # aBugTxt = "save original item"
    # Call Try("%da die Nachricht gendert wurde")
    # oItem.Save
    if Catch(DoMessage:=False) Then:
    # LogicTrace = LogicTrace _
    # & "trying save and close modified item " _
    # & oID & vbCrLf

    # Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    # oItem.Close olDiscard
    # Set oItem = Nothing
    # aBugTxt = "Get Item from EntryID=" & oID
    # Call Try
    # Set ReGet = aNameSpace.GetItemFromID(oID)
    if Catch Then:
    # badItem:
    # LogicTrace = LogicTrace & "Could not get Item from EntryID=" & oID & vbCrLf
    else:
    # Set aItmDsc.idObjItem = ReGet       ' changed by ReGet
    # nID = ReGet.EntryID
    # aBugVer = ReGet.Class = oldClass
    if DoVerify(aBugVer, "design check " _:
    # & "ReGet.Class = oldClass Class change on ReGet ???") Then
    # Call GetITMClsModel(ReGet, aPindex)
    # Call aItmDsc.UpdItmClsDetails(ReGet)
    if LenB(aNewCat) > 0 _:
    # And ReGet.Categories <> aNewCat Then
    # LogicTrace = LogicTrace & vbCrLf _
    # & " reget did set wanted Categories " & Quote(aNewCat) _
    # & b & oID & vbCrLf
    if Retry < 1 Then:
    print(Debug.Print "Retrying the save operation after setting new Categories")
    # Retry = Retry + 1
    # ReGet.Categories = aNewCat
    # Set oItem = ReGet
    # GoTo tryAgain

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CreateRawItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createrawitem():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.CreateRawItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim rawItem As Object
    # Dim retrycount As Long
    if DestFolder Is Nothing Then:
    # DoVerify False
    # ' create work item (uses default folder)
    # Set rawItem = olApp.CreateItem(aObjDsc.objItemType)
    if DebugMode Then:
    if rawItem.Class <> oItem.Class Then:
    print(Debug.Print "Class Change: ", rawItem.Class, , aObjDsc.objTypeName, oItem.Class)
    print(Debug.Print "TypeName Change: ", TypeName(rawItem), aObjDsc.objTypeName)
    # DoVerify False
    # rawItem.Subject = "Delete this Raw Item"
    # rawItem.Categories = "RAW ITEM"
    # ' use Close because Save would cause an open inspector defeating Move below
    # aBugTxt = "save raw item"
    # Call Try(0)
    # rawItem.Close olSave
    # Catch

    # ' default folder may not be our target: move it
    if rawItem.Parent.FolderPath <> DestFolder.FolderPath Then:
    # retrythis:
    # aBugTxt = "move raw item, retry #" & retrycount
    # Call Try
    # Set CreateRawItem = rawItem.Move(DestFolder)
    if Catch Then:
    # aBugTxt = "delete raw item, retry #" & retrycount
    # Call Try
    # rawItem.Delete
    # Catch
    if CreateRawItem.Parent.FolderPath <> DestFolder.FolderPath Then:
    if retrycount > 0 Then DoVerify False:
    if DebugMode Then DoVerify False:
    # Set rawItem = CreateRawItem
    # retrycount = retrycount + 1
    # GoTo retrythis
    else:
    # Set CreateRawItem = rawItem
    # DoVerify CreateRawItem.Parent.FolderPath = DestFolder.FolderPath
    if DebugMode Then:
    print(Debug.Print "RAW item (no content) moved to " _)
    # & Quote(DestFolder.FolderPath)
    # Set rawItem = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : CopyToWithRDO
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Copies Item to the DestFolder returning copied Item
# '---------------------------------------------------------------------------------------
def copytowithrdo():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.CopyToWithRDO"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim sItem As Object
    # Dim aTypeName As String

    # Dim rItem As Object ' Redemption.Safe-X-Item
    # Dim newItem As Object
    # Dim oID As String
    # Dim nID As String
    # Dim rDest As RDOFolder
    # Dim aClassKey As String
    # Dim fullTypeName As String

    # Set sItem = Nothing
    # Set rItem = Nothing
    # Set newItem = Nothing
    # Set rDest = Nothing
    if DestFolder Is Nothing Then:
    # DoVerify False
    if oItem Is Nothing Then:
    print(Debug.Print "Can not copy a Nothing")
    # DoVerify False
    # GoTo FuncExit

    # oID = oItem.EntryID
    # aTypeName = TypeName(oItem)
    if aOD(aPindex).objClsRules Is Nothing Then:
    # aClassKey = CStr(oItem.Class)
    # DoVerify D_TC.Exists(aClassKey), "Unknown Item Class " & oItem.Class & " can not be copied"
    # DoVerify oItem Is D_TC.Item(aClassKey).idObjItem, "not the same item"
    # Set aID(aPindex) = D_TC.Item(aClassKey)

    if Not aObjDsc.IsSame(aID(aPindex).idObjDsc, showdiffs:=DebugMode) Then:
    # Call LogEvent("mail-derived item type? " & aObjDsc.objTypeName & b _
    # & Quote1(oItem.Subject), eLSome)

    # aTypeName = aObjDsc.objTypeName
    # fullTypeName = TypeName(oItem)
    # aBugTxt = "Redemption.Safe" & fullTypeName
    # Call Try
    # Set sItem = CreateObject("Redemption.Safe" & fullTypeName)
    # Catch
    # aBugTxt = "RDO Session GetFolderFromPath" & Quote(DestFolder.FolderPath, Bracket)
    # Call Try
    # Set rDest = aRDOSession.GetFolderFromPath(DestFolder.FolderPath)
    # Catch
    # aBugTxt = "RDO Session GetRDOObjectFromOutlookObject(oItem)"
    # Call Try
    # Set rItem = aRDOSession.GetRDOObjectFromOutlookObject(oItem)
    # Catch

    # aBugTxt = "RDO item.add with MessageClass=" & oItem.MessageClass
    # Call Try
    # Set newItem = rDest.Items.Add(oItem.MessageClass)
    # aBugTxt = "close any open item displays"
    # Call Try
    # oItem.Close olDiscard
    # Catch
    # aBugTxt = "RDO rItem CopyTo newitem"
    # Call Try
    # Call rItem.CopyTo(newItem)                      ' CopyTo with rdoFolder does not work
    # Catch

    # aBugTxt = "RDO old Item Save"
    # Call Try(testAll)
    if rItem.modified Then:
    # ' *** attempts to get the MAPI Item in the DestFolder (it is there and OK so far)
    # rItem.Save
    # Catch
    if newItem.modified Then:
    # aBugTxt = "RDO new Item Save"
    # Call Try(testAll)
    # newItem.Save                                    ' rdo!
    # Catch

    print(' Debug.Print "EntryID original            : " & oID)
    print(' Debug.Print "EntryID after CopyTo(rItem ): " & rItem.EntryID)
    print(' Debug.Print "EntryID after CopyTo(newitem): " & newItem.EntryID)
    print(' Debug.Print "Folder(rItem ) = " & rItem.Parent.FolderPath, " Subject=" & rItem.Subject)
    print(' Debug.Print "Folder(newitem) = " & newItem.Parent.FolderPath, " Subject=" & newItem.Subject)
    print(' Debug.Print "rItem.Parent.FolderPath <> rDest.FolderPath          is ", rItem.Parent.FolderPath <> rDest.FolderPath)
    print(' Debug.Print "rItem.Parent.FolderPath <> newitem.Parent.FolderPath is ", rItem.Parent.FolderPath <> newItem.Parent.FolderPath)
    # ' Debug.Assert False

    # nID = newItem.EntryID               ' obtaining non-Rdo from Rdo Object
    # aBugTxt = "get the newItem from its EntryID (per NameSpace)"
    # Call Try
    # Set newItem = aNameSpace.GetItemFromID(nID)
    # Catch

    if Not newItem.Saved Then:
    # aBugTxt = "save newItem"
    # Call Try
    # newItem.Save                    ' RDO-Object has no save flag, get from non-RDO version
    # Catch

    # nID = newItem.EntryID

    # aBugVer = newItem.Class = oItem.Class
    # DoVerify aBugVer, "changing item Class: " & aObjDsc.objTypeName & b _
    # & Quote1(oItem.Subject)
    # aBugVer = aItmDsc.idObjItem.Class = oItem.Class
    # DoVerify aBugVer, "changing aItmDsc.IdObjItem or its item Class: " _
    # & aItmDsc.idObjItem.Class = oItem.Class _
    # & vbCrLf & String(5, b) & Quote1(oItem.Subject) _
    # & vbCrLf & String(5, b) & Quote1(aItmDsc.idObjItem.Subject)
    # 'de-reference the new item, then open it again
    # 'otherwise Outlook might still have tidbits
    # 'left from the original message
    # Set newItem = ReGet(newItem, nID)
    if newItem Is Nothing Then:
    # GoTo FuncExit
    # Set CopyToWithRDO = newItem
    if Not NewObjDsc Is aItmDsc.idObjDsc Then:
    # Set NewObjDsc = aItmDsc.idObjDsc
    # aItmDsc.idEntryId = nID
    # Call aItmDsc.UpdItmClsDetails(newItem)
    # aBugVer = aItmDsc.idObjItem Is newItem
    # DoVerify aBugVer, "aItmDsc.idObjItem Is newItem ???"

    # FuncExit:
    # Set newItem = Nothing   ' same on exit!
    # Set sItem = Nothing
    # Set rItem = Nothing
    # Set rDest = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CopyToWithSafeItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copytowithsafeitem():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.CopyToWithSafeItem"
    # DoVerify False, "*** CopyToWithSafeItem: function is not used ???"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # ' uses objItemType and aObjDsc.objTypeName (Global)

    # Dim sItem As Object                                     ' Redemption.Safe-X-Item
    # Dim newItem As Object
    # Dim aTypeName As String

    if DestFolder Is Nothing Then:
    # DoVerify False
    if oItem Is Nothing Then:
    # LogicTrace = LogicTrace & "Can not copy a Nothing" & vbCrLf
    # DoVerify False
    # GoTo cleanup
    # aTypeName = TypeName(oItem)
    # DoVerify aObjDsc.IsSame(aID(aPindex), showdiffs:=DebugMode), _
    # "** Safe Item has differrent object description"

    if DebugMode Then:
    print(Debug.Print " will create " & "Redemption.Safe" & aTypeName)
    # Set sItem = CreateObject("Redemption.Safe" & aTypeName)
    if DebugMode Then:
    print(Debug.Print " Redemption.Safe" & aTypeName _)
    # & " has ItemClass=? (no item)" & sItem.Item Is Nothing
    # sItem.Item = newItem
    if DebugMode Then:
    print(Debug.Print " Redemption.Safe" & aTypeName _)
    # & " now has ItemClass=sItem.Item.class ", _
    # " TypeName = " & TypeName(sItem.Item)
    # ' create work item (uses default folder? -- causes moveto)
    # Set newItem = CreateRawItem(DestFolder, oItem)
    if DebugMode Then:
    print(Debug.Print " created outlook rawitem for " & aTypeName _)
    # & " ItemClass=Item.class ", _
    # " TypeName = " & TypeName(newItem)
    # ' copy the item we want to copy into the new item
    # sItem.Item = newItem
    # aBugTxt = "save item using Redemption"
    # Call Try
    # Call sItem.CopyTo(newItem)
    if Catch Then:
    # LogicTrace = LogicTrace _
    # & "Redemption SafeItem.CopyTo failed" & vbCrLf
    # GoTo cleanup

    if sItem.Item.Subject <> oItem.Subject Then:
    # DoVerify False
    # Set newItem = sItem.Item
    # newItem.Save
    if Not newItem.Saved Then:
    # LogicTrace = LogicTrace _
    # & "Could not save item after copy with Redemption" _
    # & vbCrLf
    # GoTo cleanup
    # Call aItmDsc.UpdItmClsDetails(newItem)

    # 'de-reference the new item, then open it again
    # 'otherwise Outlook might still have tidbits
    # 'left from the original message
    # Set newItem = ReGet(newItem, newItem.EntryID)
    if newItem Is Nothing Then:
    # GoTo cleanup
    # Set CopyToWithSafeItem = newItem
    # cleanup:
    # Set newItem = Nothing
    # Set sItem = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CopyToWithRedemption
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copytowithredemption():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.CopyToWithRedemption"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim copiedItem As Object
    # Dim MovedItemO As Object
    # Dim ErrorInFunction As Boolean
    # Dim oClass As OlObjectClass
    # Dim aClassKey As String
    # Dim CopyIsSameAsOriginal As Boolean

    if toThisFolder Is Nothing Then:
    # DoVerify False
    if LogicTrace = "*" Then:
    # LogicTrace = vbNullString
    if oItem Is Nothing Then:
    print(Debug.Print "can't copy nothing-Object")
    # DoVerify False
    # oClass = oItem.Class
    # aClassKey = CStr(oClass)
    # Set aItmDsc = GetITMClsModel(oItem, aPindex)
    # Call aItmDsc.UpdItmClsDetails(oItem)            ' setting the global objItemType/Name
    # DoVerify Not aObjDsc Is Nothing, "Unknown object type " & TypeName(oItem)

    if trySave Then:
    if TrySaveItem(oItem) Then                  ' save original, try repair if save failed:
    if oItem Is Nothing Then:
    print(Debug.Print "TrySave ruined original Item")
    # DoVerify False
    # ErrorInFunction = True
    # LogicTrace = "Trying to repair failed Save of original item" & vbCrLf
    # Set copiedItem = CopyToWithRDO(oItem, toThisFolder, NewObjDsc)
    # '                            =============
    if copiedItem Is Nothing Then:
    # LogicTrace = LogicTrace & "Repairing Save of original item failed" & vbCrLf
    # ErrorInFunction = True
    else:
    # ErrorInFunction = False
    # aBugTxt = "delete original item " & Quote(oItem.Parent.FolderPath)
    # Call Try
    # oItem.Delete
    if CatchNC Then:
    # DoVerify False, " what is the err.number???"
    # LogicTrace = LogicTrace & "Could not delete original item in " _
    # & Quote(oItem.Parent.FolderPath) & vbCrLf
    if Catch Then:
    # LogicTrace = LogicTrace & "Original item gone, " _
    # & E_AppErr.Description & vbCrLf
    # ErrorInFunction = True
    # GoTo HardProblem
    else:
    # LogicTrace = LogicTrace & "Deleted original item in " _
    # & Quote(oItem.Parent.FolderPath) & vbCrLf

    # Set oItem = copiedItem
    if oItem.Saved Then:
    # LogicTrace = LogicTrace & "item Saved by repair!" & vbCrLf
    # ErrorInFunction = False
    else:
    # LogicTrace = LogicTrace & "item still not Saved" & vbCrLf
    # ErrorInFunction = True
    # DoVerify False
    else:
    # LogicTrace = LogicTrace & "item not saved before copy!" & vbCrLf

    # ' Copy original to copiedItem
    # Call N_ErrClear

    # Set copiedItem = CopyToWithRDO(oItem, toThisFolder, NewObjDsc)
    # '                    =============
    if copiedItem Is Nothing Then:
    # LogicTrace = LogicTrace & "CopyToWithRDO failed" & vbCrLf
    # ErrorInFunction = True
    # GoTo FuncExit
    else:
    # ErrorInFunction = False

    # ' check if the items are unchanged
    if copiedItem.EntryID <> oItem.EntryID Then:
    if copiedItem.Parent.FolderPath = toThisFolder.FolderPath Then:
    if aObjDsc.objHasReceivedTime Then:
    if oItem.SentOn = copiedItem.SentOn Then:
    # CopyIsSameAsOriginal = True
    else:
    # LogicTrace = LogicTrace & "copied item has a modified SentOn-Time" & vbCrLf
    if oItem.ReceivedTime = copiedItem.ReceivedTime Then:
    # CopyIsSameAsOriginal = True
    else:
    # LogicTrace = LogicTrace & "copied item has a different Received-Time" & vbCrLf
    # CopyIsSameAsOriginal = False
    if CopyIsSameAsOriginal Then:
    # LogicTrace = LogicTrace & "SafeItem has been copied to Folder " _
    # & toThisFolder.FolderPath
    # Set MovedItemO = copiedItem
    # GoTo NoNeedToMove
    else:
    # GoTo mustMove

    if CopyIsSameAsOriginal Then:
    # LogicTrace = LogicTrace & "SafeItem already is in Folder " _
    # & Quote(toThisFolder.FolderPath) & vbCrLf
    # Set MovedItemO = copiedItem
    # GoTo NoNeedToMove
    else:
    if copiedItem.Parent.FolderPath <> toThisFolder.FolderPath Then:
    # mustMove:
    # Set MovedItemO = copiedItem.Move(toThisFolder)
    else:
    # LogicTrace = LogicTrace & "SafeItem did not need to be moved to Folder " _
    # & Quote(toThisFolder.FolderPath) & vbCrLf
    # Set MovedItemO = copiedItem

    if Catch(AddMsg:="Move to new Folder failed") Then:
    # LogicTrace = LogicTrace & "Move to new Folder " _
    # & Quote(toThisFolder.FolderPath) & " failed" & vbCrLf
    # ErrorInFunction = True
    else:
    # Set copiedItem = Nothing

    # ' must save to have EntryID
    if TrySaveItem(MovedItemO) Then:
    print(Debug.Print "Moved Item could not be saved, new EntryID uncertain")
    # ErrorInFunction = True

    # NoNeedToMove:
    if DebugMode Then:
    # Call checkDates(oItem, MovedItemO)
    if DebugLogging Then:
    # Call ShowIdentifers(oItem)
    print(Debug.Print "Target:")
    # Call ShowIdentifers(MovedItemO)
    if ErrorInFunction Then:
    # LogicTrace = LogicTrace & "The operation has failed after all" & vbCrLf
    else:
    if CopyIsSameAsOriginal Then:
    # Call LogEvent("    : CopyToWithRedemption found " & LogicTrace, eLall)
    else:
    # Call LogEvent("    > CopyToWithRedemption successfull into " _
    # & MovedItemO.Parent.FolderPath, eLall)
    # Set CopyToWithRedemption = MovedItemO
    # HardProblem:
    if DebugMode Or ErrorInFunction Then:
    if LenB(LogicTrace) > 0 Then:
    print(Debug.Print LogicTrace)
    # DoVerify Not ErrorInFunction

    # Call N_ErrClear

    # FuncExit:
    # Set copiedItem = Nothing
    # Set MovedItemO = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function TrySaveItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trysaveitem():
    # Dim zErr As cErr
    # Const zKey As String = "ItemOpsOL.TrySaveItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim RetryWaits As Long
    # Const maxRetries As Long = 4
    # Dim trueStart As Variant
    # Dim TotalTime As Variant

    # ' ErrTrySaveItem = False: assume all works fine
    # trueStart = 0
    # TotalTime = 0
    if oItem.Saved Then:
    # TrySaveItem = False
    if DebugMode Then:
    print(Debug.Print "Saved already prior to TrySaveItem " & Quote(oItem.Parent.FolderPath))
    # GoTo ProcReturn

    # aBugTxt = "save modified item in " & Quote(oItem.Parent.FolderPath)
    # Call Try("*da die Nachricht gendert wurde.")
    # oItem.Save
    if Not Catch Then:
    # GoTo FuncExit
    # Set oItem = ReGet(oItem)
    if oItem.Saved Then:
    # GoTo FuncExit
    else:
    # aBugTxt = "save after ReGet in " & Quote(oItem.Parent.FolderPath)
    # Call Try
    # oItem.Save
    if Not Catch Then:
    # TrySaveItem = False
    print(Debug.Print "Saved without problems ")
    # GoTo FuncExit

    # Retry:
    if RetryWaits < maxRetries Then                         ' try to force with waits/repeats:
    # Wait 2 ^ RetryWaits, trueStart:=trueStart, _
    # TotalTime:=TotalTime, _
    # Retries:=RetryWaits
    if oItem.Saved Then:
    # Call LogEvent("     * Trysave may not have saved all data, " _
    # & "     * but Item was saved after " _
    # & RetryWaits _
    # & " Attempts (" & TotalTime & ") sec ", eLall)
    # TrySaveItem = False
    else:
    # oItem.Save
    if Not oItem.Saved Then:
    # GoTo Retry
    # TrySaveItem = True
    if DebugMode Or RetryWaits >= maxRetries Then:
    print(Debug.Print "retried " & RetryWaits & " times , wait time=" & CInt(TotalTime))
    if TrySaveItem Then:
    # Call LogEvent("TrySave failed")
    if DebugMode Then:
    # DoVerify False

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub checkDates
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkdates():
    # Optional mailTypeCheck As Boolean = True)
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "ItemOpsOL.checkDates"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    # Dim WhatEver As String

    try:

        if mailTypeCheck Then:
        if IsMailLike(item1) Then:
        # aBugVer = item1.Subject = item2.Subject
        if DoVerify(aBugVer, "don't compare items with distinct subject:") Then:
        print(Debug.Print "1: " & Quote(item1.Subject))
        print(Debug.Print "2: " & Quote(item2.Subject))
        # GoTo ProcReturn
        if IsMailLike(item2) Then:
        if ShowTimes(item1, item2, "EntryID") Then:
        print(Debug.Print "===> time values must be identical")
        else:
        # Call ShowTimes(item1, item2, "ReceivedTime")
        # Call ShowTimes(item1, item2, "SentOn")
        # Call ShowTimes(item1, item2, "CreationTime")
        else:
        # WhatEver = TypeName(item2)
        # GoTo noda
        else:
        # WhatEver = TypeName(item1)
        # GoTo noda
        else:
        # noda:
        print(Debug.Print "can't determine Date/Time values for type " _)
        # & WhatEver & " (non-mail type item)"
        # DoVerify False

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function ShowTimes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showtimes():
    # Const zKey As String = "ItemOpsOL.ShowTimes"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim x As Date
    # Dim y As Date

    match WhatEver:
        case "EntryID":
    # ShowTimes = item1.EntryID = item2.EntryID
    print(Debug.Print "Entry IDs match: " & CStr(ShowTimes) & " Subject: " & Quote(item1.Subject))
    # GoTo ProcRet
        case "ReceivedTime":
    # x = item1.ReceivedTime
    # y = item2.ReceivedTime
        case "SentOn":
    # x = item1.SentOn
    # y = item2.SentOn
        case "CreationTime":
    # x = item1.CreationTime
    # y = item2.CreationTime
        case _:
    # DoVerify False, " not implemented"
    print(Debug.Print WhatEver & "1 : " & x & " in " & Quote(item1.Parent.FullFolderPath))
    print(Debug.Print WhatEver & "2 : " & y & " in " & Quote(item2.Parent.FullFolderPath))
    print(Debug.Print WhatEver & "s match: " & CStr(x = y))

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowIdentifers
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showidentifers():

    # Const zKey As String = "ItemOpsOL.ShowIdentifers"
    # Call DoCall(zKey, tSub, eQzMode)

    # Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    print(Debug.Print "Saved=" & Item.Saved, Item.Subject, Item.CreationTime)
    print(Debug.Print Item.Parent.FolderPath, Item.EntryID)
    # Catch

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function ItemValid
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Determine if Item is (still) valid
# '---------------------------------------------------------------------------------------
def itemvalid():
    # '''' Proc Must ONLY CALL Z_Type PROCS                        ' May be Silent
    # Const zKey As String = "ItemOpsOL.ItemValid"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction, ExplainS:="ItemOpsOL")

    # Dim rClass As Long

    if testItem Is Nothing Then:
    # GoTo ProcReturn
    if isEmpty(testItem) Then:
    # GoTo ProcReturn
    # aBugTxt = "Item Class not available"
    # Call Try
    # rClass = testItem.Class
    # Catch
    if rClass = 0 Then:
    # GoTo ProcReturn
    if Not D_TC.Exists(CStr(testItem.Class)) Then:
    # Call LogEvent("Item Class has not been defined", eLall)
    # GoTo ProcReturn
    # aBugTxt = "Get EntryID for Item"
    # Call Try
    # ItemValid = testItem.EntryID <> vbNullString                      ' throws error &H8004010A when gone
    if Catch Then:
    # ItemValid = False
    else:
    # ItemValid = True
    # ProcReturn:
    # Call ProcExit(zErr)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ContactFixer
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def contactfixer():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "ItemOpsOL.ContactFixer"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="ContactFixer")

    # Call ContactFixItem(ActiveExplorer.Selection.Item(1))

    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ContactFixItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def contactfixitem():
    # Const zKey As String = "ItemOpsOL.ContactFixItem"
    # Static zErr As New cErr

    # Dim oneContact As ContactItem
    # Dim bestSaveAs As String

    if oneItem.Class <> olContact Then:
    if DebugLogging Then:
    # Call LogEvent("item is not a Contact, skipped " & oneItem.Subject, eLall)
    # GoTo skipExit

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="ContactFix")

    # Set oneContact = oneItem            ' or aRDOSession.GetRDOObjectFromOutlookObject(oneItem)
    # With oneContact
    if LenB(.Email1Address) > 0 Then:
    # bestSaveAs = .Email1Address
    elif LenB(.Email2Address) > 0 Then:
    # bestSaveAs = .Email2Address
    elif LenB(.Email3Address) > 0 Then:
    # bestSaveAs = .Email3Address
    elif InStr(.Body, "@") > 0 Then:
    # bestSaveAs = GetWordContaining(oneItem.Body, "@")
    if bestSaveAs <> .Email1Address Then    ' modify email addresses to avoid redundancies:
    # .Email1Address = bestSaveAs
    if bestSaveAs = .Email2Address Then:
    # .Email2Address = vbNullString
    if bestSaveAs = .Email3Address Then:
    # .Email3Address = vbNullString
    elif .Email2Address = .Email1Address Then:
    # .Email2Address = vbNullString
    elif .Email3Address = .Email2Address Then:
    # .Email3Address = vbNullString
    if LenB(bestSaveAs) = 0 Then:
    if LenB(.Home2TelephoneNumber) > 0 Then:
    # bestSaveAs = NormalizeTelefonNumber(.Home2TelephoneNumber, Reassign:=True)
    if bestSaveAs <> .Home2TelephoneNumber Then:
    # .Home2TelephoneNumber = bestSaveAs
    elif LenB(.HomeTelephoneNumber) > 0 Then:
    # bestSaveAs = NormalizeTelefonNumber(.HomeTelephoneNumber, Reassign:=True)
    elif LenB(.MobileTelephoneNumber) > 0 Then:
    # bestSaveAs = NormalizeTelefonNumber(.MobileTelephoneNumber, Reassign:=True)
    elif LenB(.BusinessTelephoneNumber) > 0 Then:
    # bestSaveAs = NormalizeTelefonNumber(.BusinessTelephoneNumber, Reassign:=True)
    elif LenB(.Business2TelephoneNumber) > 0 Then:
    # bestSaveAs = NormalizeTelefonNumber(.Business2TelephoneNumber, Reassign:=True)

    if LenB(oneContact.CompanyName) > 0 Then:
    if LenB(.CompanyAndFullName) > 0 Then:
    # .FileAs = .CompanyAndFullName
    else:
    # .FileAs = .CompanyName
    # bestSaveAs = .FileAs
    elif LenB(.FullName) > 0 Then:
    # .FileAs = .FullName
    # bestSaveAs = .FileAs
    elif LenB(.LastNameAndFirstName) > 0 Then:
    # .FileAs = .LastNameAndFirstName
    # bestSaveAs = .FileAs
    elif LenB(bestSaveAs) > 0 Then:
    # bestSaveAs = Trim(bestSaveAs)
    # .FileAs = bestSaveAs

    if Not .Saved Then:
    # .Save
    # LF_ItmChgCount = LF_ItmChgCount + 1
    # Call LogEvent("Contact " & bestSaveAs & " in " & .Parent.FolderPath & " corrected", eLall)
    # End With ' oneContact

    # Set oneContact = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # skipExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ChangeBirthdaySubject
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Ersetze "Geburtstag von" durch "*". Nur in Appointment-Ordnern
# '---------------------------------------------------------------------------------------
def changebirthdaysubject():
    # Const zKey As String = "ItemOpsOL.ChangeBirthdaySubject"
    # Static zErr As New cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSubEP, ExplainS:="ItemOpsOL")

    # Dim StrBuffer As String
    # Dim LenBuffer As Long
    # Dim Counter As Long
    # Dim aItems As Items
    # Dim aCalItem As Object
    # Dim i As Long

    # Counter = 0

    if thisFolder Is Nothing And LenB(fName) < 2 Then:
    # Set thisFolder = aNameSpace.GetDefaultFolder(olFolderCalendar)
    else:
    if LenB(fName) = 0 Or fName = "#" Then:
    # Set thisFolder = ActiveExplorer.CurrentFolder
    else:
    # Set thisFolder = GetFolderByName(fName)
    # fName = thisFolder.FolderPath

    if LenB(fName) > 0 Then:
    if thisFolder.FolderPath <> thisFolder.FolderPath Then:
    # GoTo ProcReturn

    if thisFolder.DefaultItemType <> olAppointmentItem Then:
    # GoTo ProcReturn

    # Set aItems = thisFolder.Items

    # Set aCalItem = thisFolder.Items(i)
    # StrBuffer = aCalItem.Subject
    if InStr(StrBuffer, "Geburtstag von ") Then:
    # ' aCalItem.Display
    # LenBuffer = Len(StrBuffer)
    # StrBuffer = Right(StrBuffer, (LenBuffer - Len("Geburtstag von ")))
    # StrBuffer = "*" + StrBuffer
    # aCalItem.Subject = StrBuffer
    # aCalItem.Save
    # 'aCalItem.Close 0
    # Counter = Counter + 1

    print('Fertig!')

    # ProcReturn:
    # Call ProcExit(zErr)


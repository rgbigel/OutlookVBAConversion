Attribute VB_Name = "ItemOpsOL"
Option Explicit

' all item properties decoded (and converted to string, if possible). Some are skipped.
                                        
'---------------------------------------------------------------------------------------
' Method : Sub CheckPhoneNumber
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CheckPhoneNumber()
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.CheckPhoneNumber"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim lName As String
    lName = LCase(PropertyNameX)
    NormalizedPhoneNumber = vbNullString
    If InStr(lName, "radio") > 0 Then
        ItsAPhoneNumber = False
    ElseIf InStr(lName, "telex") <= 1 _
        And InStr(lName, "extension") = 0 _
        And InStr(lName, "phone") = 0 _
        And InStr(lName, "fax") = 0 Then
        ItsAPhoneNumber = False
    ElseIf InStr(lName, "phone") > 0 _
    Or InStr(lName, "fax") > 0 Then
        ItsAPhoneNumber = True ' fax or phone IS a telco number
    Else
        ItsAPhoneNumber = True ' anything with number in it is a telco number??
        DoVerify False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.CheckPhoneNumber

'---------------------------------------------------------------------------------------
' Method : Sub PhoneNumberNormalize
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PhoneNumberNormalize(oFoneNr As String, px As Long)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.PhoneNumberNormalize"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If Len(oFoneNr) < 5 And Left(oFoneNr, 1) <> "1" Then ' extension number
        If InStr(oFoneNr, "*") = 0 Then
            NormalizedPhoneNumber = "*" & oFoneNr
        Else
            NormalizedPhoneNumber = oFoneNr
        End If
    Else
        NormalizedPhoneNumber = NormalizeTelefonNumber(oFoneNr)
    End If
    
    If oFoneNr <> NormalizedPhoneNumber Then
        If SelectOnlyOne And px = 1 And Not aDecProp(1) Is Nothing Then    ' save original for Display (Excel)
            aCloneMode = FullCopy
            Set aDecProp(2) = aDecProp(1).adictClone        ' save original TelNumber value in number 2
        End If
        MatchPoints(px) = MatchPoints(px) + 1               ' ratedelta = 1
        PhoneNumberNormalized = True
        aStringValue = NormalizedPhoneNumber                ' new value in px
        If CurIterationSwitches.SaveItemRequested Then
            WorkItemMod(px) = True
            Message = fiMain(px) & b _
                & aTD.adName _
                & " geändert in " _
                & NormalizedPhoneNumber _
                & " (war " & oFoneNr & ")"
            Call LogEvent(Message, eLall)
            oFoneNr = NormalizedPhoneNumber
        ElseIf Not saveItemNotAllowed Then
            ' not changing WorkItemMod(px) !!!
            Message = fiMain(px) & b _
                & aTD.adName _
                & " wird verglichen als " _
                & NormalizedPhoneNumber _
                & " (war " & oFoneNr & ")"
            Call LogEvent(Message, eLall)
            oFoneNr = NormalizedPhoneNumber
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.PhoneNumberNormalize

'---------------------------------------------------------------------------------------
' Method : Function IsPropertyOK
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsPropertyOK(Item As Object, CritPropName As String) As Boolean ' false if missing property
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.IsPropertyOK"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim iProp As ItemProperty
Dim ThisPropisSelected As Boolean

    If apropTrueIndex >= 0 Then                             ' we know where it is
        GoTo useTrueIndex
    End If
    If LenB(CritPropName) > 0 _
    And UCase(CritPropName) <> "OR" _
    And UCase(CritPropName) <> "NOT" _
    And UCase(CritPropName) <> "A" Then
        ' GetAttrDsc when finding...
useTrueIndex:
        ThisPropisSelected = (apropTrueIndex < 0)
        If ThisPropisSelected Then                          ' we know where
            If LenB(CritPropName) = 0 Then
                DoVerify False, "CritPropName empty, dead"
            End If                                          ' we know absolutely nothing to find property
            PropertyNameX = CritPropName                    ' we know a name but no index
        Else
            If LenB(CritPropName) = 0 Then                  ' check if we can use existing aTD
                If aTD Is Nothing Then
                    DoVerify False, "No aTD and no CritPropName: makes no sense. Must quit"
                End If
                If aTD.adName <> CritPropName Then
                    DoVerify False, "aTD Name mismatches CritPropName"
                End If
                aBugVer = iProp Is aTD.adItemProp
                If DoVerify(aBugVer, "design check iProp Is aTD.ADItemProp ???") Then
                    Set iProp = aTD.adItemProp
                End If
            Else
                If aTD Is Nothing Then
                    If aProp Is Nothing Then
                        PropertyNameX = vbNullString                  ' FindProperty in position aPropTrueIndex
                    ElseIf aProp.Name = PropertyNameX Then
                        Set iProp = aProp
                        GoTo why
                    End If
                Else
                    If aTD.adName <> PropertyNameX Then
                        DoVerify False, " inconsistent; very bad"
                        PropertyNameX = vbNullString                  ' FindProperty in position aPropTrueIndex
                        Set aTD = Nothing
                    End If
                End If
            End If
        End If
        ' FindProperty will also NewAttrDsc (if possible and needed)
        ' and it will get the found ItemProperty into (new) aTD
        ' +   evaluate Prop. Value into ADDecodedValue, determining the Value Type
        ' +   if possible, determine the trivial value (without formatting)
        If apropTrueIndex <> aTD.adtrueIndex Then
why:
            apropTrueIndex = FindProperty(apropTrueIndex, _
                                PropertyNameX, iProp, _
                                Item)                       ' iProp is OUT, aPropTrueIndex checked or OUT
        End If
        If apropTrueIndex < 0 Then                          ' not found anywhere
            GoTo ProcReturn                                 ' property is not defined in thisItem,
        End If                                              ' IsPropertyOK returning False
        If aTD Is Nothing Then
            apropTrueIndex = -1
            DoVerify False, " may be correct ???"
            GoTo why
        End If
        
        AttributeIndex = aTD.adtrueIndex
        IsPropertyOK = True                                 ' make a new one in known position
        
        ' we want all if we dont want OnlyMostImportantProperties
        If aTD.adNr > 0 Then
            ' format (e.g. Phone Numbers), specials for some array cases (e.g. MemberCount->Members, Photos)
            Call PrepDecodeProp
            Call logDecodedProperty(aStringValue, String(4, b))
            Call StackAttribute                             ' add this to aID(aPindex).idAttrDict [dictionary]
        End If
    End If  ' its a usable critProp name

FuncExit:

ProcReturn:
    Call ProcExit(zErr, CStr(IsPropertyOK))

pExit:
End Function ' ItemOpsOL.IsPropertyOK

'---------------------------------------------------------------------------------------
' Method : Function EvaluateSpecialRequirements
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function EvaluateSpecialRequirements(oItem As Object) As Long
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.EvaluateSpecialRequirements"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim originalItemClass As OlObjectClass
Dim BaseItem As Object

    workingOnNonspecifiedItem = False
    originalItemClass = oItem.Class
    BaseAndSpecifiedDiffer = False  ' normally, BaseItem[=standarditem] == aID(2..3
                                    ' but may be moved to 3..4 for exc/occ
    EvaluateSpecialRequirements = aPindex   ' normally, no change
    DoVerify aID(aPindex).idObjItem Is oItem, "standard item class index 1..2 must match ** Remove 2 if no hit"
    Set aID(aPindex).idObjItem = oItem   ' standard item class index 1..2
    
    If originalItemClass = olAppointment And aPindex < 3 Then
        ' could be olApptMaster, olApptOccurrence or olApptException:
        If oItem.ItemProperties.Count < minPropCountForFullItem Then ' short item
            Select Case oItem.RecurrenceState
            Case olApptOccurrence
                SpecialObjectNameAddition = "#O"
                GoTo NonStandardItem
            Case olApptException
                SpecialObjectNameAddition = "#E"
NonStandardItem:
                If aID(aPindex + 2) Is Nothing Then
                    Set aID(aPindex + 2).idObjDsc = New cObjDsc
                    Set aID(aPindex + 2).idObjDsc.objSeqInImportant = New Collection
                    Call aID(aPindex + 2).SetDscValues(oItem, _
                        withValues:=True, aRules:=sRules, _
                        SD:=SpecialObjectNameAddition)
                    DoVerify aOD(aPindex + 2).objNameExt = SpecialObjectNameAddition, _
                        "creation of Name/Key/Extension not logical ???"
                    ' *** note: this aID now runs under aPindex, not aPindex+2
                    ' *** because we use the same idAttrDict for adding further attrs
                End If
                Set aID(aPindex + 2).idObjItem = oItem  ' put specified item into +2 pos
                                                        ' and standard item (its parent) +0 pos
                Set BaseItem = oItem.Parent             ' parent done first, specified item second
                DoVerify BaseItem.RecurrenceState = olApptMaster, "oItem Parent not plausible"
                Set aID(aPindex).idObjItem = BaseItem   ' standard item class index 1..2
                aOD(aPindex).objMaxAttrCount = 0        ' this is a new item to decode
                workingOnNonspecifiedItem = True        ' using parent ^= specified
                BaseAndSpecifiedDiffer = True           ' not the specified item yet
                EvaluateSpecialRequirements = aPindex + 2
            Case olApptMaster
                ' no need to go to parent, specifiedItem == aID(apindex)
            Case Else
                DoVerify False, "RecurrenceState not expected/defined"
            End Select
        End If
    End If  ' class appointment, may be standard or extended by nonstandard

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.EvaluateSpecialRequirements

'---------------------------------------------------------------------------------------
' Method : Sub InitAttributeSetup
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Create Maintanance information for base Item Description / Extension
'---------------------------------------------------------------------------------------
Sub InitAttributeSetup(baseID As cItmDsc, Optional ExtensionID As cItmDsc)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.InitAttributeSetup"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    
Dim BasePx As Long
    
    BasePx = baseID.idPindex
    If BasePx = 0 Then
        BasePx = 1
        baseID.idPindex = BasePx
    End If
    If ExtensionID Is Nothing Then
        Set ExtensionID = baseID
    End If
    Call N_ClearAppErr
    Set rP(BasePx) = Nothing                       ' set to recurrencePattern when IsRecurring in IsPropertyOK
    If aID(BasePx).idAttrDict.Count = 0 Then
        AttributeIndex = 0
        AttributeUndef(BasePx) = 0
    End If
    If aID(BasePx).idAttrDict.Count < AttributeIndex Then
        AttributeIndex = aID(BasePx).idAttrDict.Count
    End If
    
    If AttributeIndex > 0 Then
        If AttributeIndex = ExtensionID.idObjDsc.objMaxAttrCount Then
            GoTo FuncExit                           ' all done
        End If
    End If
    
    ' normal case presets:
    MatchPoints(aPindex) = 0
    ' we have a new aID when we are in SelectAndFind (apindex=2)
    If baseID.idObjItem.Class <> ExtensionID.idObjDsc.objItemClass Then
        DoVerify False, "design test only ???"
        Call EvaluateSpecialRequirements(baseID.idObjItem)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.InitAttributeSetup

' Important Properties (by Rules) first in sequence in aID(aPindex).idAttrDict, ??? design change not completed !!!
' then rest in PropertyIndex sequence (unless StopAfterMostRelevant),

'---------------------------------------------------------------------------------------
' Method : GetMiAttrNr
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Most Important Attribute Number for PropertyNameX, added to SeqInImportant
' Note:    requires ATD and aObjDsc
'---------------------------------------------------------------------------------------
Sub GetMiAttrNr()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "ItemOpsOL.GetMiAttrNr"
    Call DoCall(zKey, tSub, eQzMode)

Dim i As Long
Dim testADNr As Boolean

    ' note: MostImportantAttributes is usually identical with SelectedAttributes, cut can be longer. Never shorter.
    If InStr(MostImportantAttributes, PropertyNameX) > 0 Then
        If aTD.adisSel _
        And aTD.adRules.clsObligMatches.RuleMatches _
        And aTD.adRules.clsSimilarities.RuleMatches Then
            testADNr = True
        End If
        For i = LBound(MostImportantProperties) To UBound(MostImportantProperties)
            If MostImportantProperties(i) = PropertyNameX Then
                DoVerify aTD.adDictIndex = apropTrueIndex + 1, _
                    "design check aTD.ADDictIndex = apropTrueIndex + 1 ???"
                If aTD.adNr >= 0 Then
                    If DoVerify(aTD.adNr = i, _
                        "change in position in MostImportantAttributes ??? " & PropertyNameX) Then
                        aTD.adNr = i
                    Else
                        GoTo zExit                          ' all as done befor
                    End If
                Else
                    DoVerify aTD.adDictIndex = apropTrueIndex + 1, "design check ???"
                    aObjDsc.objSeqInImportant.Add aTD.adDictIndex
                    If aTD.adNr <> aTD.adDictIndex Then
                        DoVerify Not testADNr, "why do we assign a new .ADNR"
                        aTD.adNr = aTD.adDictIndex
                    End If
                    aTD.adisSel = True
                    Call LogEvent("added objSeqInImportant(" _
                        & aObjDsc.objSeqInImportant.Count & ") for Property=" & aTD.adKey _
                        & ", mostImportantProperty#=" & i & " in Object Class " & aObjDsc.objItemClassName _
                        & " [Dict Index " & aTD.adDictIndex & "]", eLSome)
                End If
                GoTo zExit
            End If
        Next i
        DoVerify aTD.adisSel Or aTD.adNr > 0, "look into this ???" ' note: ADNr=0 always Key=Application, do not select
    End If
    aTD.adisSel = False
    aTD.adNr = -1
zExit:
    Call DoExit(zKey)
End Sub ' ItemOpsOL.GetMiAttrNr

' from Attributes of Item px -> aPindex, formatting for output
' MAKE SURE IsRecurring is part of Important Properties if recurrence pattern needed
Sub GetItemAttrDscs(Item As Object, px As Long)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetItemAttrDscs"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim somevaluesNotAvailable As Collection
Dim iPos As Long ' position in ItemProperties if we know it
Dim tPos As Long '              "                certain
Dim itmNo As Long
Dim aIDa As cObjDsc

    aPindex = px
    AttributeIndex = 0
    
    Set somevaluesNotAvailable = New Collection
    If LenB(TrueCritList) = 0 Then
        SkipDontCompare = False
    End If
    
    If isEmpty(MostImportantProperties) Then
        MostImportantProperties = Array(sRules.clsObligMatches.CleanMatches(0))
        MostImportantAttributes = sRules.clsObligMatches.CleanMatches(0)
    End If
    
    Set aProps = Item.ItemProperties
    TotalPropertyCount = aProps.Count

'  NOTE: we will not work on aID(apindex).odItemDict here, because
'        we also have to process the specified items (#B)
'        so we use Item.itemproperties instead
    
    iPos = 1    ' start position for loop: index in odItemDict
doSpecialAttributes:
    For tPos = 0 To Item.ItemProperties.Count - 1    ' tPos = 0 start of (true) properties
        OneDiff = vbNullString
        PropertyNameX = Item.ItemProperties.Item(tPos).Name
        If aID(aPindex) Is Nothing Then
            Set aTD = Nothing
        Else
            Set aTD = aID(aPindex).GetAttrDsc4Prop(tPos)    ' setting AllPublic.aProp, reusing if present
        End If
        If aTD Is Nothing Then                      ' not in list of SelectedAttributes: new
            apropTrueIndex = tPos                   ' we know
            aBugTxt = "Design check Not aID(aPindex).idAttrDict Is Nothing ???, Property tPos=" & tPos
            DoVerify Not aID(aPindex).idAttrDict Is Nothing, aBugTxt
            With aID(aPindex).idAttrDict
                If .Exists(PropertyNameX) Then
                    Set aTD = .Item(PropertyNameX)
                    If aTD.adtrueIndex < 0 Then
                        aTD.adtrueIndex = apropTrueIndex
                    End If
                    DoVerify aProp.Name = aTD.adName, _
                        "Property Name mismatch for aTD true index=" & aTD.adtrueIndex
                Else
                    Set aTD = ProvideAttrDsc
                    If Not aTD.adRules.RuleInstanceValid Then
                        Call SplitDescriptor(aTD)
                    End If
                End If
            End With ' aID(aPindex).idAttrDict
        
            If isSpecialName Then
                apropTrueIndex = -apropTrueIndex    ' must be correct for special attributes. never on -1!!!
            End If
        End If
        
        aTD.adtrueIndex = apropTrueIndex
        DoVerify PropertyNameX = aTD.adName, "Property Name messed up ???"
                
        ' rules for property name and its isSpecialName (suffix #?) are always same
        If Not aTD.adRules Is iRules Then
            Set aTD.adRules = iRules
            Call SplitDescriptor(aTD)               ' using PropertyNameX inside
            Call GetMiAttrNr
        End If
        
        PhoneNumberNormalized = False
        IgString = vbNullString
        If aTD.adNr > 0 Or AllProps Then            ' SkipDontCompare is NOT considered for aTD setup
            ' format (e.g. Phone Numbers), specials for some array cases (e.g. MemberCount->Members, Photos)
            DoVerify Not iRules Is Nothing, "iRules can't be missing ???"
            Call PrepDecodeProp
            Call StackAttribute                             ' add this to aID(aPindex).idAttrDict [dictionary]
        Else
            aStringValue = "## value not evaluated"
        End If
            
        aTD.adDecodedValue = aStringValue
        AttributeIndex = aTD.adtrueIndex
        If Not displayInExcel Then
            Call logDecodedProperty(aStringValue, String(4, b))
        ElseIf DebugLogging Then
            Debug.Print Format(Timer, "0#####.00") & vbTab & "Progress:", _
                itmNo, AttributeIndex, PropertyNameX, aStringValue, IgString
        End If
nextInLoop:
        iPos = iPos + 1
    Next tPos  ' loop for properties
    
    If workingOnNonspecifiedItem Then               ' more to do on specifiedItem ???    ' *** Hier
        workingOnNonspecifiedItem = False
        ' we do NOT add these special attributes to aID(aPindex + 2)
        Set aIDa = aID(aPindex)
        PropertyNameX = vbNullString
        IgString = vbNullString
        logDecodedProperty vbNullString, " ++ Starting on special attributes +++"
        ' but we get the properties from this idObjItem
        Set Item = aID(aPindex + 2).idObjItem
        isSpecialName = True                        ' use name suffix in dictionary
        ' and simply continue adding to Dictionary, starting with tPos
        ' from previous loop exit, iPos == (old) itemproperties.count
        iPos = tPos                                 ' used to access attributes for apindex + 2
        GoTo doSpecialAttributes                    ' of the specified item
    End If
    
asFarAsWeWanted:
    If aPindex < 3 Then                             ' Processing Recurrence and Exceptions, if any
        If Not (rP(aPindex) Is Nothing Or ExceptionProcessing) Then
            ExceptionProcessing = True
            IgString = vbNullString
            If Not rP(aPindex) Is Nothing Then      ' recurrence Pattern has no ItemProperties
                tPos = rPTrueIndex                  ' base of indirect attributes, < 0
                Call RpStackAndLog(aPindex, rP(aPindex))
                                                    ' so we must process the properties we know
                                                    ' also the Exceptions thereof
            End If
            Set rP(aPindex) = Nothing
            ExceptionProcessing = False
        End If
    End If
    ' maxProperties (ever)
    iPos = aID(aPindex).idAttrDict.Count
    If aOD(0).objMaxAttrCount < iPos - 1 Then
        aOD(0).objMaxAttrCount = iPos + 1
        If aOD(0).objMinAttrCount = 0 Then
            aOD(0).objMinAttrCount = iPos - 1
        Else
            If aOD(0).objMinAttrCount > iPos Then
                aOD(0).objMinAttrCount = iPos - 1
            End If
        End If
    End If
    If CurIterationSwitches.SaveItemRequested And Not workingOnNonspecifiedItem Then
        If aID(aPindex).idObjItem.Class = olContact Then
            MPEchanged = False
            Call NameCheck(aID(aPindex).idObjItem)
            If MPEchanged Then
               WorkItemMod(aPindex) = True
            End If
        End If
    End If
    If somevaluesNotAvailable.Count = 0 Then
        If DebugLogging And Not ShutUpMode Then
            Debug.Print "no Attributes specified are missing"
        End If
    Else
        Call LogEvent("missing " & somevaluesNotAvailable.Count _
                    & " specified attributes in this item")
        For iPos = 1 To somevaluesNotAvailable.Count
            Call LogEvent(iPos & vbTab & somevaluesNotAvailable.Item(iPos))
        Next iPos
    End If
    
    If DebugMode Then
        Debug.Print Format(Timer, "0#####.00") & vbTab _
            & "ended GetItemAttrDscs with " _
            & AttributeIndex & " properties, subject:" _
            & vbCrLf, Quote(aID(aPindex).idObjItem.Subject)
    End If
            
    If SelectOnlyOne Or FindMatchingItems Then
        If LenB(SelectedAttributes) = 0 Then
            Call RulesToExcel(aPindex, Not FindMatchingItems)
        End If
    ElseIf aPindex = 2 Then
        Call AppendMissingProperties(0, 1, 2)
        Call AppendMissingProperties(0, 2, 1)
        Call RulesToExcel(aPindex, Not FindMatchingItems)
    End If
    OnlyMostImportantProperties = quickChecksOnly   ' restore user choice

FuncExit:
    Set aIDa = Nothing
    Set somevaluesNotAvailable = Nothing
    Set aID(aPindex).idObjItem = Item
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.GetItemAttrDscs

'---------------------------------------------------------------------------------------
' Method : Function NewAttrDsc
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function NewAttrDsc(lpropTrueIndex As Long) As cAttrDsc
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.NewAttrDsc"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    Set NewAttrDsc = ProvideAttrDsc                 ' uses aProp to find or make cAttrDsc
    If lpropTrueIndex > -1 Then
        NewAttrDsc.adtrueIndex = lpropTrueIndex
    End If
    If aTD Is Nothing Then
        DoVerify False, "design check ???"
        Set aTD = NewAttrDsc
        Set iRules = Nothing
    End If
    If iRules Is Nothing Then
        DoVerify False, "design check ???"
        Call CreateIRule(PropertyNameX)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.NewAttrDsc

'---------------------------------------------------------------------------------------
' Method : Sub AppendMissingProperties
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AppendMissingProperties(AttributeIndex As Long, base As Long, Copy As Long)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.AppendMissingProperties"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim startIndex As Long
Dim nDecProp As cAttrDsc
    ' are we missing anything ?
    If AttributeIndex = 0 Then  ' select base count
        AttributeIndex = aID(base).idAttrDict.Count
    End If
    If aID(Copy).idAttrDict Is Nothing Then
        Stop ' ???
        Set aID(Copy).idAttrDict = New Dictionary
    End If
    startIndex = aID(Copy).idAttrDict.Count
    If startIndex < AttributeIndex Then
        i = startIndex
        While i < aID(base).idAttrDict.Count - 1 ' additional attrs w/o value
            i = i + 1
            aCloneMode = DummyTarget
            Set nDecProp = aID(base).idAttrDict.Item(i).adictClone  'früher ??? (Copy, base)
            aID(Copy).idAttrDict.Add nDecProp.adKey, nDecProp
        Wend
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.AppendMissingProperties

'---------------------------------------------------------------------------------------
' Method : Function GetAobj
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetAobj(px As Long, Optional ByVal knownItemIndex As Long = 0) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetAobj"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If aPindex = px Then
        aPindex = px                                        ' this Aobj is new here, GetAobj?
    End If
    aItmIndex = knownItemIndex
    WorkItemMod(px) = False
    If knownItemIndex >= 0 Then
        targetIndex = px
        If sortedItems(px) Is Nothing Then                  ' sorted items have highest prio
            If SelectedItems.Count < knownItemIndex Then    ' try selected items next
                DoVerify False, " caller uses index 2, but there is only one"
                If SelectedObjects Is Nothing Then
                    aItmIndex = -1
                Else
                    Call GetSelectedItems(SelectedObjects)
                    GoTo useSelected
                End If
            Else
useSelected:
                If aItmIndex = 0 And SelectedItems.Count > 0 Then
                    Set GetAobj = SelectedItems.Item(1)
                    aItmIndex = 1
                Else
                    If aItmIndex = 0 Then
                        Set GetAobj = SelectedItems.Item(px)
                        aItmIndex = px
                    Else
                        Set GetAobj = SelectedItems.Item(knownItemIndex)
                        aItmIndex = knownItemIndex
                    End If
                End If
                Set WorkItem(aItmIndex) = GetAobj
            End If
            If GetAobj Is Nothing Then
                DoVerify False, "this path should not be taken: why use ActiveExplorer ???"
                If ActiveExplorerItem(aPindex) Is Nothing Then
                    ' make sure we have some item (GetAobj)
                    ' to determine the item type and Folder
                    If ActiveExplorerItem(1) Is Nothing Then
                        If ActiveExplorer Is Nothing Then
                            DoVerify False
                        Else
                            Set Folder(px) = ActiveExplorer.CurrentFolder
                            If Folder(px).Items.Count = 0 Then
                                DoVerify False
                            End If
                            Set GetAobj = Folder(px).Items(1)
                        End If
                    End If
                Else
                    Set GetAobj = ActiveExplorerItem(aPindex)
                End If
                Set WorkItem(aItmIndex) = GetAobj
            End If
        Else                                ' try sorted items
            If aItmIndex > sortedItems(px).Count Then
                aItmIndex = -1
                GoTo ProcReturn
            End If
            If aItmIndex = 0 Then
figure_as_selection:
                If LenB(ActiveExplorerItem(px)) = 0 Then
                    If ActiveExplorer.Selection.Count >= px Then
                        Set ActiveExplorerItem(px) = _
                            ActiveExplorer.Selection.Item(px)
                    Else
                        aItmIndex = -1
                        GoTo ProcReturn
                    End If
                End If
                Set GetAobj = ActiveExplorerItem(px)
                aItmIndex = px
            Else
                Set GetAobj = sortedItems(px).Item(knownItemIndex)
            End If
            Set WorkItem(aItmIndex) = GetAobj
        End If
    Else                                                    ' knownItemIndex=0
        If GetAobj Is Nothing Then                          ' we use existing value:
            Set GetAobj = aID(px).idObjItem
        End If
        If GetAobj Is Nothing Then                          ' use selected item(1)
            aItmIndex = 0
            GoTo useSelected
        End If
        Set WorkItem(aItmIndex) = GetAobj
    End If                                                  ' knownItemIndex>=0
    
    DoVerify Not GetAobj Is Nothing, "no Object determined"
    If knownItemIndex >= 0 Then                             ' special case, do not set aID for -1
        If aID(px) Is Nothing Then
            Set aID(px) = New cItmDsc
            GoTo DefineIt
        End If
        If Not aID(px).idObjItem Is GetAobj Then            ' Different Item needs own object description
DefineIt:
            Call DefObjDescriptors(GetAobj, px, withValues:=True)
        End If                                              ' sets aID, Attributes, Values
    End If
    knownItemIndex = aItmIndex
    
    Set ActItemObject = GetAobj                             ' set as global value

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.GetAobj

'---------------------------------------------------------------------------------------
' Method : Function DecodeObjectClass
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DecodeObjectClass(getValues As Boolean) As String
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.DecodeObjectClass"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim DescriptorIndex As Long

    DescriptorIndex = EvaluateSpecialRequirements(ActItemObject)
    ' ActItemObject is now always the standard object class on aIndex
    '   aIndex+2-> may address special objects (SpecialRequirements for exception or occurrence items)
    
    If ActItemObject Is Nothing Then
        DecodeObjectClass = "-"
        GoTo ProcReturn
    End If
    
    Call DefObjDescriptors(ActItemObject, aPindex, withValues:=getValues)
    DecodeObjectClass = aOD(aPindex).objItemClassName
    
    fiMain(aPindex) = GetMainObjectIdentification

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.DecodeObjectClass

'---------------------------------------------------------------------------------------
' Method : Function GetMainObjectIdentification
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetMainObjectIdentification() As String
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetMainObjectIdentification"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If aPindex < 3 Then ' MainObjectIdentification only for Standard Objects
        If LenB(MainObjectIdentification) = 0 Then     ' try existing
            MainObjectIdentification = aOD(aPindex).objDefaultIdent
        End If
        If LenB(MainObjectIdentification) = 0 Then     ' try global default
            MainObjectIdentification = MostImportantProperties(0)
        End If
        If StrComp(MainObjectIdentification, "none", vbTextCompare) = 0 Then
            MainObjectIdentification = vbNullString   ' never valid
        End If
        If LenB(MainObjectIdentification) = 0 Then     ' try general rules
            Call aID(aPindex).UpdItmClsDetails(ActItemObject)
        End If
        aBugVer = LenB(MainObjectIdentification) > 0
        If DoVerify(aBugVer, "no main identification") Then
            aOD(aPindex).objDefaultIdent = MainObjectIdentification
            GetMainObjectIdentification = getPropertyValue(MainObjectIdentification)
        End If
            
        fiMain(aPindex) = ReFormat(GetMainObjectIdentification, vbCrLf, "|", b)
        If LenB(GetMainObjectIdentification) = 0 Then
            fiMain(aPindex) = "## " & aOD(aPindex).objDefaultIdent & " is empty"
        Else
            fiMain(aPindex) = ReFormat(GetMainObjectIdentification, vbCrLf, "|", b)
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.GetMainObjectIdentification

'---------------------------------------------------------------------------------------
' Method : Function getPropertyValue
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getPropertyValue(PropName As String) As String
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.getPropertyValue"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim thisProp As ItemProperty

    If LenB(PropName) > 0 Then                      ' only such value can exist
        Set thisProp = LookUpAttrName(PropName)     ' look up in aID(apindex) -> aTD
        If thisProp Is Nothing Then
            getPropertyValue = vbNullString
        Else
            Call Try(allowNew)                          ' Try anything, autocatch, Err.Clear
            getPropertyValue = thisProp.Value
            Call ErrReset(0)
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.getPropertyValue

'---------------------------------------------------------------------------------------
' Method : Sub PrepDecodeProp
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PrepDecodeProp()
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.PrepDecodeProp"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim j As Long
Dim i As Long
Dim iClass As OlObjectClass
Dim vValue As Variant
Dim rateDelta As Long
Dim aLink As Variant
Dim lObjItem As Object
Dim LinkPath1 As String
Dim linkPath2 As String
Dim attachmentFile As String
Dim thisAttachment As Attachment
Dim Item As Object
    
    rateDelta = 0
    If aTD Is Nothing Then
        DoVerify False, "** aTD not available???"
        aStringValue = "# aTD not available *"
        GoTo ProcReturn
    End If
    Set Item = aTD.adItmDsc.idObjItem
    
ProcessMore:
    With aTD
        aStringValue = .adDecodedValue
        aBugVer = iRules Is .adRules
        If DoVerify(aBugVer, "design check iRules Is aTD.adRules ???") Then
            Set iRules = .adRules
        End If
        If iRules.clsNotDecodable.RuleMatches Then
            aStringValue = "# not decodable *"
            .adOrigValDecodingOK = True                 ' Skipped
            GoTo loggit
        ElseIf iRules.clsNeverCompare.RuleMatches _
        Or SkipDontCompare _
        Or Not AllProps Then
            aStringValue = "# skipped on request *"
            .adFormattedValue = aStringValue
            .adOrigValDecodingOK = True                 ' Skipped
            GoTo loggit
        Else
            vValue = aStringValue
            If .adInfo.iAssignmentMode = 1 Or .adInfo.iArraySize < 0 Then  ' kein array
                If Not .adOrigValDecodingOK Then
                    MatchPoints(aPindex) = MatchPoints(aPindex) - 1 ' penalize
                End If                                  ' rate quality of the data
                If LenB(aStringValue) > 0 Then
                    rateDelta = 2
                    MatchPoints(aPindex) = MatchPoints(aPindex) + rateDelta ' rate
                    If rP(aPindex) Is Nothing Then
                        If .adName = "IsRecurring" Then
                            If .adDecodedValue Then
                                Set rP(aPindex) = Item.GetRecurrencePattern
                                Call LogEvent("Recurring item with " _
                                        & rP(aPindex).Exceptions.Count _
                                        & " Exceptions")
                            Else
                                Set rP(aPindex) = Nothing
                            End If
                        End If
                    ElseIf .adName = "MemberCount" Then
                        DoVerify .adInfo.iArraySize = CInt(.adDecodedValue)
                        GoTo ArrayCase                  ' MemberCount announces array of ADInfo.iArraySize elements
                    Else
                        Call CheckPhoneNumber           ' fone and fax numbers ?
                        If ItsAPhoneNumber Then
                            Call PhoneNumberNormalize(.adDecodedValue, aPindex)
                        Else                            ' reformat: replace line changes and remove multiple blanks
                            aStringValue = ReFormat(aStringValue, vbCrLf, "|", b)
                        End If
                    End If
                Else
                    NormalizedPhoneNumber = vbNullString
                    ItsAPhoneNumber = False
                End If
            Else
ArrayCase:
                DoVerify False, "needs redesign"
                aStringValue = .adDecodedValue          ' we do not Show the rest here
                rateDelta = 1
                Set vValue = .adInfo.iValue
                If vValue Is Nothing Then
                    GoTo loggit
                End If
                iClass = vValue.Class
                For j = 1 To vValue.Count
                    MatchPoints(aPindex) = MatchPoints(aPindex) + rateDelta ' rate
                    Call N_ClearAppErr
                    Call Try                            ' Try anything, autocatch
                    If iClass = olActions Then
                        aStringValue = aStringValue & ", " _
                            & LString(j & ": " & vValue.Item(j).Name _
                            & "= " & vValue.Item(j).Enabled, 30)
                    ElseIf iClass = -1 Then
                        If vValue.Name = "MemberCount" Then
                            aStringValue = aStringValue & vbCrLf _
                                & LString("Member " & j _
                                & "= " & aID(aPindex).idObjItem.GetMember(j).Name, 30)
                        End If
                    ElseIf iClass = olAttachments Then
                        Set thisAttachment = vValue.Item(j)
                        If aID(aPindex).idObjItem.Class = olContact _
                        And vValue.Item(j).FileName = "ContactPicture.jpg" Then
                            aStringValue = aStringValue _
                                    & vbCrLf & j & " (ContactPicture), " _
                                    & "Size: " & thisAttachment.Size
                            Call addContactPic(aID(aPindex), thisAttachment)
                        ElseIf SaveAttachments Then
                            attachmentFile = aPfad & DateId & "(" & j & ") " _
                                    & thisAttachment.FileName
                            aStringValue = aStringValue _
                                    & vbCrLf & j & ", Size: " & thisAttachment.Size _
                                        & vbTab & "-> " & Quote(attachmentFile)
                            aBugTxt = "Save attachment" & attachmentFile
                            Call Try
                            vValue.Item(j).SaveAsFile attachmentFile
                            If Not Catch Then
                                Call LogEvent("saved attachment " & attachmentFile)
                            End If
                        Else
                            Call LogEvent("userrequest: attachment " & j & " not saved as file " _
                                    & thisAttachment.FileName)
                        End If
                    ElseIf iClass = olLinks Then
                        Call ErrReset(0)
                        For i = 1 To vValue.Count
                            If VarType(vValue.Item(i)) = vbString Then
                                GoTo missedLink
                            End If
                            Call Try(allowAll)
                            Set aLink = vValue.Item(i).Item
                            LinkPath1 = aLink.Parent.FullFolderPath
                            Set lObjItem = _
                                aNameSpace.GetItemFromID(aLink.EntryID)
                            If Catch Then
                                GoTo missedLink
                            End If
                            linkPath2 = lObjItem.Parent.FullFolderPath
                            If Catch Then
missedLink:
                                aStringValue = aStringValue _
                                    & vbCrLf _
                                    & "       link " & i _
                                    & " invalid, not found for " _
                                    & vValue.Value(j)
                                If DebugMode Or DebugLogging Then
                                    MsgBox aStringValue
                                End If
                                AttributeUndef(aPindex) = AttributeIndex
                                Call N_ClearAppErr
                            ElseIf LinkPath1 <> linkPath2 Then
                                aStringValue = aStringValue _
                                    & vbCrLf _
                                    & "       link " & i & " points to different Folder: " _
                                    & vbCrLf _
                                    & "            " & lObjItem.Parent.FullFolderPath _
                                    & " instead of " & vbCrLf _
                                    & "            " & aID(aPindex).idObjItem.Parent.FullFolderPath _
                                    & vbCrLf _
                                    & "       for  " & vValue.Value(j)
                                If DebugMode Or DebugLogging Then
                                    MsgBox aStringValue
                                End If
                                AttributeUndef(aPindex) = AttributeIndex
                            Else
                                aStringValue = aStringValue & vbCrLf _
                                    & "       found " & LString(j & ": " _
                                    & Quote(vValue.Item(j)), 30) & b
                            End If
                        Next i
                    ElseIf iClass = olRecipients Then
                        aStringValue = aStringValue & ", " _
                            & Left("Recipient" _
                            & j & "= " & Quote(vValue.Item(j).Name), 30) & b
                    Else
                        aStringValue = aStringValue & ", " _
                            & LString(j & ": " & Quote(vValue.Item(j)), 30) & b
                    End If
                    
                    If Catch Then
                        DoVerify False, "kann Property nicht auswerten"
                        Call frmErrStatus.fBeginTermination(True)
                    End If
                Next j
                aStringValue = Replace(aStringValue, Chr(160), b)
            End If
        End If
loggit:
        Call FormatAttrForDisplay
        If .adNr < 0 Then                   ' actually not required, but AllProps is set
            .adNr = .adDictIndex
            Call AppendTo(MostImportantAttributes, .adKey, b)
            If StringMod Then               ' was it appended / if not, no need to split again
                MostImportantProperties = split(MostImportantAttributes)
            End If
        End If
        If LenB(.adKillMsg) > 0 Then
            aStringValue = .adKillMsg
        Else
            aStringValue = .adDecodedValue
        End If
    End With ' aTD
    
FuncExit:
    Set Item = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.PrepDecodeProp

'---------------------------------------------------------------------------------------
' Method : Sub FormatAttrForDisplay
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FormatAttrForDisplay()
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.FormatAttrForDisplay"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim TestBody As Variant

    ' all sorts of output prettying:
    ' Special Formatting for *body* (if not empty string)
    If PropertyNameX = "Body" _
    And LenB(aStringValue) > 0 Then
        aBugTxt = "get body format"
        Call Try(438)
        TestBody = aID(aPindex).idObjItem.BodyFormat
        If Catch Then                               ' e.g. Contacts never have HTMLbody
            GoTo noHTMLbody
        End If
        
        If LenB(aStringValue) = 0 And TestBody <> olFormatHTML Then     ' somebody lied about it
            DoVerify False, "aStringValue="" impossible???"
            aID(aPindex).idObjItem.BodyFormat = olFormatHTML ' causes conversion for HTML and RTF Bodies
            ' GoTo reEvaluate   ' this should have fixed our problem
                                ' note: no loop because BodyFormat is now HTML
                                ' no idea if it works for olFormatRTF
        End If
noHTMLbody:
        killStringMsg = vbNullString
        For i = 1 To killWords.Count
            StringsRemoved = vbNullString
            If LenB(Trim(aStringValue)) = 0 Then
                Exit For
            End If
            aStringValue = RemoveWord(aStringValue, killWords.Item(i), ": ""|")
            If LenB(StringsRemoved) > 0 Then
                killStringMsg = Trim(Append(killStringMsg, _
                                            StringsRemoved, vbCrLf))
            End If
        Next i
        If DebugLogging Then
            Debug.Print killStringMsg
        End If
    End If
    
    If PropertyNameX <> aTD.adName Then
        If LenB(aTD.adName) = 0 Then                        ' original item does not have
            aTD.adName = PropertyNameX                      ' IsRecurring or Exceptions or its sub-properties
        Else
            DoVerify False, "design check ???"
        End If
    End If
    
    If aTD.adOrigValDecodingOK Then
        aTD.adFormattedValue = aTD.adDecodedValue
    Else
        aTD.adFormattedValue = aStringValue
    End If
        
    ' more formatting stuff
    If aPindex = 2 Then                                     ' shorten if necessary for display
        If aID(aPindex).idAttrDict.Exists(aTD.adKey) Then   ' not yet in aID(aPindex) - dictionary
            ' will be added in this Sub below, take pattern from adecprop C(1)
            Set aDecProp(1) = GetAttrDsc(aDecProp(2).adKey, Get_aTD:=False, FromIndex:=1)
            ' If debugMode And aDecProp(1).attrPos = 0 Then
                ' Debug.Assert False ' this is no fix!!!
            aDecProp(2).adNr = aDecProp(1).adNr
        Else
            Stop ' ???
            Set aDecProp(1) = aID(1).idAttrDict(1).Items(aDecProp(2).adNr)
        End If
        ' get corresponding item to side 1 (no Err, first time we sync)
        If aDecProp(1) Is Nothing Then
            DoVerify False, " aID(2).attrPos??? Clone??? ***"
        Else
            If aDecProp(2) Is Nothing Then
                ' get corresponding item to side 1 (no Err, first time we sync)
                aCloneMode = withNewValues                  ' use ADItmDsc, rules etc, but not values
                Set aDecProp(2) = aID(1).idAttrDict.Item(aDecProp(1).adKey).Clone()
            ElseIf aDecProp(1).adNr <> aDecProp(2).adNr Then
                ' get any missing attributes from other side
                '   (no Err, first time we sync)
                If DebugMode Then DoVerify False
                Call AppendMissingProperties(AttributeIndex, 1, 2)
                Call AppendMissingProperties(AttributeIndex, 2, 1)
            End If
        End If
Dim p1 As String
Dim p2 As String
        
        Call FirstDiff(aDecProp(1).adFormattedValue, _
                       aDecProp(2).adFormattedValue, _
                         p1, _
                         p2, _
                       80, 30, "...", OneDiff_qualifier)
        If Left(aDecProp(1).adFormattedValue, 1) <> "{" Then ' can't shorten arrays
            aDecProp(1).adShowValue = p1                    ' parameters can not be byref and byval at the same time
            aDecProp(2).adShowValue = p2                    ' :( so we must put p1/2 in between
        Else
            aDecProp(1).adShowValue = aDecProp(1).adFormattedValue
            aDecProp(2).adShowValue = aDecProp(2).adFormattedValue
        End If
        aTD.adKillMsg = OneDiff_qualifier
        OneDiff_qualifier = vbNullString
        If displayInExcel And (xDeferExcel Or xUseExcel) Then
           Call put2IntoExcel(aPindex, AttributeIndex + 1)
        End If
    Else
        Set aDecProp(1) = aTD                               ' aPindex = 1 allways in this Proc
        aTD.adShowValue = aTD.advValue
    End If
fm:

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.FormatAttrForDisplay

'---------------------------------------------------------------------------------------
' Method : FindAttributeByName
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Find position in aID(aPindex).idAttrDict of attribute with CritPropName
'---------------------------------------------------------------------------------------
Function FindAttributeByName(AttributeIndex As Long, CritPropName As String) As Long
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.FindAttributeByName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim logicalIndex As Long
Dim lTrueIndex As Long
Dim lpx As Long
Dim nAttrDsc As cAttrDsc

    lpx = aPindex
    If aID(2) Is Nothing Then
        lpx = 1
    End If
    
    ' most likely, the sequence in aID(1).idAttrDict is same as in aID(2).idAttrDict
    ' idAttrDict (2) may not be OK, so check idAttrDict(1) first
    For logicalIndex = AttributeIndex To aID(1).idAttrDict.Count
        Set nAttrDsc = aID(1).idAttrDict.Item(logicalIndex)
        aBugVer = LCase(CritPropName) <> LCase(nAttrDsc.adName)
        If DoVerify(aBugVer, "? Order is off?" _
            & vbCrLf & logicalIndex & b & aID(1).idAttrDict.Item(logicalIndex).adName) Then
            
            YleadsXby = logicalIndex - AttributeIndex   ' count out-of-order items
            GoTo GotIt
        End If
    Next logicalIndex
    
    ' did not find it aID(1).idAttrDict going forward, so look in reverse
    For logicalIndex = AttributeIndex - 1 To 1 Step -1
        Set nAttrDsc = aID(1).idAttrDict.Item(logicalIndex)
        If LCase(CritPropName) = LCase(nAttrDsc.adName) Then
            YleadsXby = AttributeIndex - logicalIndex   ' count out-of-order items in aID(1).idAttrDict
GotIt:
            FindAttributeByName = logicalIndex
            lTrueIndex = nAttrDsc.adDictIndex          ' position in dictionary
            Set aTD = aID(1).idAttrDict.Item(Abs(lTrueIndex) - 1)
            If lpx = 2 And aID(lpx).idAttrDict.Count < logicalIndex Then
                lpx = 1
            End If
            If aID(lpx).idAttrDict Is Nothing Then      ' nothing there to clone
                Set aDecProp(lpx) = Nothing
            Else                                        ' can item on lpx side be cloned?
                Set aDecProp(lpx) = aID(lpx).idAttrDict.Item(logicalIndex)
                If lpx <> aPindex Then
                    DoVerify False, "check this design ???"
                   ' If aDecProp C(aPindex) Is Nothing Then   ' will set up aDecProp C(px=2) by cloning
                   '     Set aDecProp C(aPindex) = New Collection
                   ' End If
                    aCloneMode = FullCopy               ' clone with old values ???
                    aPindex = aPindex
                    Set aTD = aDecProp(lpx).adictClone
                   ' aDecProp C(aPindex).Add aTD        ' fill (aPindex) with values from the other(lpx) side
                   ' aTD.ADNr = aDecProp C(aPindex).Count
                End If
            End If
            Call Get_iRules(aTD)
            GoTo ProcReturn
        End If
    Next logicalIndex
    FindAttributeByName = 0                             ' nothing found at all

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.FindAttributeByName

'---------------------------------------------------------------------------------------
' Method : Sub SetAttributeByName
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetAttributeByName(Ci As Variant, px As Long, CritPropName As String, PropValue As Variant, Optional andPropertyToo As Boolean)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.SetAttributeByName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim propertyIndex As Long
Dim msg As String
    aPindex = px
    If FindAttributeByName(1, CritPropName) > 0 Then
        If VarType(PropValue) = vbString Then
            msg = aTD.adShowValue
            msg = " value changed from " & Quote(msg) & "  to " & Quote(PropValue) & b
            aTD.adShowValue = PropValue
            If andPropertyToo Then
                On Error GoTo errenc
                propertyIndex = aTD.adtrueIndex
                Ci.ItemProperties.Item(propertyIndex).Value = PropValue
                msg = msg & " (item property, too)"
            End If
        Else
            DoVerify False
errenc:
            msg = " could not be assigned to the item's property: " & Err.Description
        End If
    Else
        msg = " not found, no value change attempted."
    End If
    Debug.Print "Attribute " & CritPropName & msg

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.SetAttributeByName

' Find in ItemProperties and return Property's true index in ItemProperties
'  will also deliver the curItemProp and aDictIndex (or Nothing/0 if not exists yet)
'  '  '    ' i is known true Property index or -1 if unknown
'  '  '        in which case we use ADName to find it
Function FindProperty(ByVal i As Long, ByVal adName As String, curItemProp As ItemProperty, aObject As Object) As Long
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.FindProperty"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim j As Long
Dim aDictItem As cAttrDsc
Dim withoutAdminData As Boolean

    FindProperty = -1           ' no success so far
    If aObject Is Nothing Then  ' should not happen, try to deduce from other parms
        DoVerify aTD Is Nothing And curItemProp Is Nothing And i < 0, "impossible to do that", True
check_aTD:
        If aTD Is Nothing Then
            If aID(aPindex).idAttrDict Is Nothing Then
provideNew:
                Set aTD = ProvideAttrDsc
            Else
                aBugVer = aID(aPindex).idAttrDict.Exists(adName)
                If DoVerify(aBugVer, "not in dictionary or aTD, ADName=" & adName) Then
                    GoTo buildAttrDsc
                Else    ' verify checks OK
                    If isEmpty(aID(aPindex).idAttrDict.Item(adName)) Then
buildAttrDsc:
                        withoutAdminData = True
                        GoTo provideNew
                    End If
                    Set aTD = aID(aPindex).idAttrDict.Item(adName)
                    aBugVer = aTD.adItem Is Nothing
                    If DoVerify(aBugVer, "aID(" & aPindex & ") will not give us a correct aTD") Then
                        GoTo FuncExit
                    Else    ' verify checks OK
                        Set aObject = aTD.adItem
                        GoTo got_aTD
                    End If
                End If
            End If
        Else
got_aTD:
            If aTD.adName = adName Then
                If aObject Is aTD.adItmDsc.idObjItem Then
                    GoTo got_object
                Else
                    GoTo get_Object
                End If
            Else
                aBugVer = LenB(adName) = 0
                If DoVerify(aBugVer, "ADName is empty: improbable") Then
                    aBugVer = aID(aPindex).idAttrDict.Exists(aTD.adKey)
                    If DoVerify(aBugVer, _
                        "aTD is not matching by key") _
                    Then
                        aBugVer = aTD.adKey = aTD.adName
                        If DoVerify(aBugVer, "wrong key/Name combination") Then
                            Set aTD = aID(aPindex).idAttrDict.Item(aTD.adKey)
                        End If
                    Else
                        GoTo FuncExit
                    End If
                Else    ' verify checks OK
                    adName = aTD.adName
                End If
            End If
get_Object:
            Set aObject = aTD.adItmDsc.idObjItem
            If aObject Is Nothing Then
                aBugVer = curItemProp Is Nothing
                If DoVerify(aBugVer, "impossible if aObject and curItemProp are Nothing") Then
                    GoTo FuncExit
                End If
                aBugVer = curItemProp.Name = adName
                If DoVerify(aBugVer, "curItemProp=" & curItemProp.Name _
                        & " mismatches ADName=" & adName) Then
                    GoTo FuncExit
                Else    ' verify checks OK
                    Set aObject = curItemProp.Parent.Parent
                    GoTo got_object
                End If
            Else
                GoTo got_aTD
            End If
        End If
    Else                                    ' have aObject
        GoTo check_aTD
    End If
    
got_object:                                 ' also have aTD when we get here
    Set aProps = aObject.ItemProperties
    If aTD.adtrueIndex > -1 Then
        If aProps.Item(i) Is aTD.adItemProp Then
            If curItemProp Is Nothing Then
                Set curItemProp = aTD.adItemProp
            Else
                DoVerify curItemProp Is aTD.adItemProp, "murks ???"
                GoTo FuncExit
            End If
        Else
            DoVerify False, "aTD is not matching ???"
            GoTo FuncExit
        End If
        
        FindProperty = aTD.adtrueIndex
        GoTo findVerify
    Else
        aTD.adtrueIndex = apropTrueIndex
    End If
    
    If i < 0 Then ' Position unknown, use name only
findNew:
        DoVerify False, "this should not really be neccessary: unknown property position i<0 ???"
        Set iRules = Nothing                             ' must determine again
        Set curItemProp = Nothing
        For j = 0 To aProps.Count - 1
            Set curItemProp = aProps.Item(j)
            If curItemProp.Name = adName Then
                FindProperty = j                        ' the Prop TrueIndex
                apropTrueIndex = j
                Set aProp = curItemProp
                If withoutAdminData Then
                    GoTo FuncExit
                End If
                
                If Not aID(aPindex).idAttrDict Is Nothing Then
                With aID(aPindex).idAttrDict
                    If .Exists(adName) Then
                        Set aTD = .Item(adName)
                        If aTD.adtrueIndex < 0 Then
                            aTD.adtrueIndex = j     ' now we have true index
                        End If
                        DoVerify aProp.Name = aTD.adName, _
                            "Property Name mismatch for aTD true index=" & aTD.adtrueIndex
                    End If
                End With ' aID(aPindex).idAttrDict
                End If
                
                If aTD Is Nothing Then                  ' try to get in established attributes
                    Set aTD = GetAttrDsc(adName)       ' sets iRules if Rules already defined
                End If
                GoTo findVerify                         ' aTD, iRules undefined, curItemProp is OK
            End If
        Next j
        Set aTD = Nothing
        FindProperty = -1
        Set curItemProp = Nothing
        GoTo findVerify
    Else
        Set curItemProp = aProps.Item(i)
        If LenB(PropertyNameX) > 0 And curItemProp.Name <> PropertyNameX Then
            DoVerify False, "curItemProp.Name <> PropertyNameX ???"
            GoTo findNew
        End If
        PropertyNameX = curItemProp.Name
        adName = vbNullString   ' do not use to find
        FindProperty = i
        j = i                                           ' this would be the position if we loop
    End If
    
    ' try to find in already established Attributes
    If LenB(adName) > 0 Then    ' use name:
        If aTD Is Nothing Then
            Set aTD = GetAttrDsc(adName)               ' sets iRules if Rules already defined
        Else
            Set curItemProp = aTD.adItemProp
            Set aProp = curItemProp
            GoTo ProcReturn                             ' all is well
        End If
    Else
        If FindProperty >= 0 Then
            GoTo findVerify
        End If
    End If
    If aTD Is Nothing Then
        GoTo findNew        ' its not in attributes, get from Properties
    Else ' found it in Attributes!
        With aTD
            If .adName = adName Then
                j = .adDictIndex
                If j > 0 Then
                    If aID(aPindex).idAttrDict.Exists(adName) Then
                        DoVerify False, "code needed ???"
                        GoTo findVerify
                    End If
                Else    ' indirect attribute like RecursionPattern or Exceptions
                    Set curItemProp = Nothing
                End If
                FindProperty = Abs(.adtrueIndex)
                GoTo funex
            End If
        End With ' aTD
    End If
    
    ' found in ItemProperties
findVerify:
    If FindProperty < 0 Then                        ' if invalid: not in itemProperties either
        If PropertyNameX <> "Links" Then            ' Links can always be missing
            DoVerify False, Quote(adName) & " not in itemProperties of " & TypeName(aObject)
        End If
        Set aTD = Nothing
        Set iRules = Nothing
    Else    ' we did find it in the itemProperties as it should be
        Set aProp = curItemProp
        If apropTrueIndex < 0 Then
            apropTrueIndex = j
        End If
        If Not aID(aPindex).idAttrDict.Exists(PropertyNameX) Then
            GoTo WrongATD
        End If
        If Not aID(aPindex).idAttrDict.Item(PropertyNameX) Is aTD Then
            GoTo WrongATD
        End If
        If aTD Is Nothing Then  ' has no AttrDsc yet, make one
WrongATD:
            apropTrueIndex = FindProperty
            Set aTD = NewAttrDsc(apropTrueIndex)   ' into aTD
        ElseIf aID(aPindex).idAttrDict.Exists(PropertyNameX) Then
            DoVerify aID(aPindex).idAttrDict.Item(PropertyNameX) Is aTD Or aObjDsc.objMaxAttrCount = 0, _
                    "aTD mismatches idAttrDict for " & PropertyNameX & " of " & aID(aPindex).idObjDsc.objItemClassName
            If aID(aPindex).idAttrDict.Item(PropertyNameX).adtrueIndex <> apropTrueIndex Then
                aID(aPindex).idAttrDict.Item(PropertyNameX).adtrueIndex = apropTrueIndex
            End If
        ElseIf aTD.adName <> aID(aPindex).idAttrDict.Item(j + 1).Item.adName Then
            DoVerify False, "it wasn't the right one *** Impossible, remove ???"
            Set aTD = aID(aPindex).idAttrDict.Item(aTD.adName).Item
        '***ElseIf aID(aPindex).odAttArray(j) Is Nothing Then
        Else
            DoVerify False, "should not be reached ???"
            DoVerify aProps(j).Name = aTD.adName, _
                "error PropTrueIndex: aProps(" & j & ") = " _
                & aProps(j).Name & " <> aTD.ADName " & aTD.adName
            ' we have to set the odItemDict.item
            '*** Set aID(aPindex).odAttArray(j) = aTD
        End If
        If LenB(adName) > 0 Then
            DoVerify aTD.adName = adName, "aTD.ADName <> ADName, this is extremely fishy"
        Else
            adName = aTD.adName
        End If
funex:
        With aTD
            j = aID(aPindex).idAttrDict.Count - 1
            If j < .adNr Then ' no valid aID(aPindex).idAttrDict - entry
                ' aDecProp C(aPindex).Add aTD
                j = j ' ??? aTD.ADNr = aDecProp C(aPindex).Count
            Else
                If .adNr > 0 And j >= .adNr Then
                    ' If aDecProp C(aPindex).Item(.ADNr).ADName = ADName Then
                    '    Set aDecProp(aPindex) = aDecProp C(aPindex).Item(.ADNr)
                    ' Else
                    '    Set aDecProp(aPindex) = Nothing
                    ' End If
                Else    ' no decoded property = attribute yet
                    If Not aTD Is aDecProp(aPindex) Then
                        DoVerify False, "look into this ???"
                        Set aDecProp(aPindex) = Nothing
                    End If
                End If
            End If
            aBugVer = .adItemProp Is aProp
            If DoVerify(aBugVer, "design check aTD.ADItemProp Is aProp ???") Then
                Set .adItemProp = aProp  ' forces consistency
            End If
            Set curItemProp = .adItemProp
            aBugVer = iRules Is .adRules And Not .adRules Is Nothing
            If DoVerify(aBugVer, "design check iRules Is .adRules And Not .adRules Is Nothing ???") Then
                Set iRules = .adRules
            End If
        End With ' aTD
    End If

FuncExit:
    Set aDictItem = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.FindProperty

'---------------------------------------------------------------------------------------
' Method : Function FormatPhoneNumber
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FormatPhoneNumber(aNumber As String) As String
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.FormatPhoneNumber"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
    If ItsAPhoneNumber Then
        FormatPhoneNumber = aNumber
usermodified:
        FormatPhoneNumber = Replace(FormatPhoneNumber, "00", "+")
        FormatPhoneNumber = Replace(FormatPhoneNumber, ")", b)
        i = InStr(FormatPhoneNumber, "+")
        If i = 0 Then
            If InStr(FormatPhoneNumber, "(0") = 1 Then
                FormatPhoneNumber = Replace(FormatPhoneNumber, "(0", "+49 ")
            ElseIf InStr(FormatPhoneNumber, "0") = 1 Then
                FormatPhoneNumber = Replace(FormatPhoneNumber, "0", "+49 ")
            End If
            If InStr(i + 1, FormatPhoneNumber, "+") > 0 _
            Or i > 1 Then
                With frmStrEdit
                    .Caption = fiMain(1)
                
                    .StringModifierCancelLabel.Caption = "alt:"
                    .StringModifierCancelValue.Text = FormatPhoneNumber
                    .StringModifierExpectation = "Format der Telefonnummer bitte korrigieren"
                    .StringToConfirm = FormatPhoneNumber
                    .Show
                    If .StringModifierRsp <> 0 Then
                        FormatPhoneNumber = .StringToConfirm
                        GoTo usermodified
                    End If
                End With ' frmStrEdit
            End If
        End If
        FormatPhoneNumber = Replace(FormatPhoneNumber, "(0", b)
        FormatPhoneNumber = Replace(FormatPhoneNumber, "/", b)
        FormatPhoneNumber = Replace(FormatPhoneNumber, "-", b)
        If Len(FormatPhoneNumber) < 5 Then
            If InStr(FormatPhoneNumber, "*") = 0 Then
                FormatPhoneNumber = "*" & FormatPhoneNumber
            End If
        End If
    Else
        FormatPhoneNumber = aNumber
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.FormatPhoneNumber

' get iRules and aTd:
    ' Get_aTD=False: wont get aTd, iRules etc., just check and use
Function GetAttrDsc(ByVal PropName As String, _
                                Optional Get_aTD As Boolean = True, _
                                Optional FromIndex As Long = 0) As cAttrDsc
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetAttrDsc"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim aKey As String
    
    If Mid(PropName, 2, 2) = "==" Then          ' ignore headline for group of attrs
        GoTo FuncExit
    End If

    aKey = GetAttrKey(PropName, Not Get_aTD, FromIndex)
    
    If aTD Is Nothing Then                      ' there is a aKey defined in odItemDict
        If aID(FromIndex).idAttrDict.Exists(aKey) Then
            If isEmpty(aID(FromIndex).idAttrDict.Item(aKey)) Then
                Set aTD = New cAttrDsc
            End If
        End If
        If Not iRules Is Nothing Then
            If Not iRules.RuleObjDsc Is Nothing And Not aObjDsc Is Nothing Then
                If iRules.RuleObjDsc.objClassKey <> aObjDsc.objClassKey Then
                    Set iRules = Nothing        ' not defined yet, no atd and no iStuff
                    iRuleBits = "(void)"
                End If
            End If
        End If
        GoTo FuncExit
    Else                                        ' aTd has been set to that
        aBugVer = InStr(aTD.adKey, PropName) > 0
        aBugTxt = "PropertyName is at least part of adKey ???"
        DoVerify
        If aTD.adRules Is Nothing Or aTD.adRuleIsModified Then
            If isSpecialName Then               ' raw rules only; why ??? ***
                Set GetAttrDsc = aID(FromIndex).idAttrDict.Item(PropName).Item
                Set aTD.adRules = GetAttrDsc.adRules
            Else                                ' we have no aID to copy rules
                Call CreateIRule(PropName)
            End If
        End If
            
        Call Get_iRules(aTD)
    End If
    
    Set GetAttrDsc = aTD                        ' deliver aTD (see Get_aTD)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.GetAttrDsc

'---------------------------------------------------------------------------------------
' Method : Function LookUpAttrName
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function LookUpAttrName(PropName As String) As ItemProperty

Const zKey As String = "ItemOpsOL.LookUpAttrName"
    Call DoCall(zKey, tFunction, eQzMode)

    ' PropName is some MainIdent, but may not exist (yet)
    With aID(aPindex).idAttrDict
        If .Exists(PropName) Then
            Set aTD = .Item(PropName)
            Set LookUpAttrName = aTD.adItemProp
        Else
            DoVerify False
            Set LookUpAttrName = Nothing
        End If
    End With ' aID(aPindex).idAttrDict

zExit:
    Call DoExit(zKey)

End Function ' ItemOpsOL.LookUpAttrName

'---------------------------------------------------------------------------------------
' Method : Sub GetSelectedItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetSelectedItems(S As Variant, Optional logLvl As eLogLevel = eLSome) ' outvalue is global:On Error GoTo ErrHandler
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetSelectedItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim itm As Object
Dim vItm As Variant
Dim i As Long

' SelectedItems collection does not have to be empty, but ...
    If SelectedItems Is Nothing Then
        Set SelectedItems = New Collection
    End If
    For Each vItm In S
        If Not isEmpty(vItm) Then
            If Not vItm Is Nothing Then
                Set itm = vItm  ' convert to object
                If ItemDateFilter(itm, logLvl) = vbNo Then
                    GoTo nextOne
                End If
                SelectedItems.Add itm
                i = i + 1
                    ' determine parent Folder of item (makes sense only for first 2 items)
                If i < 3 Then
                    If Folder(i) Is Nothing Then
                        Set Folder(i) = getParentFolder(itm)
                    End If
                End If
            End If
        End If
nextOne:
    Next vItm
    ' Set topFolder = getDefaultFolderType(s) ??? *** makes no sense at this time

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.GetSelectedItems

'---------------------------------------------------------------------------------------
' Method : Sub BestObjProps
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get best Object, its Object Description, Object details with time, and Rules
'          if NewObjectItem is not Nothing, it is used as aItmObject
'          else default it from curFolder, ActiveExplorer, SelectedItems, SortedItems
'          if no previous class Description exists, build one
'---------------------------------------------------------------------------------------
Sub BestObjProps(curFolder As Folder, Optional Item As Object, Optional withValues As Boolean = True)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.BestObjProps"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim px As Long
Dim msg As String
Dim aItemClass As OlObjectClass
Dim aItemClassName As String
Dim aClassKey As String
Dim Reused As String

    Reused = "New"

    If Not Item Is Nothing Then            ' need to look for best Item?
        If aObjDsc.objItemClass = Item.Class Then
            Set ActItemObject = Item
            aItemClass = ActItemObject.Class
            aClassKey = CStr(aItemClass)
            GoTo gotOne
        End If
    End If

    If curFolder Is Nothing Then
        Set curFolder = ActiveExplorer.CurrentFolder
        If curFolder.Parent Is Nothing Then
            eOnlySelectedFolder = True   ' search Folder
        ElseIf curFolder.Parent.Class = olNamespace Then
            eOnlySelectedFolder = False
            GoTo noclass
        Else
            eOnlySelectedFolder = True
        End If
    End If
    
    DoVerify aPindex < 3, "BestObjProps can only work for non-extended items"
    px = aPindex
    If px <= 0 Then
        px = 1
        GoTo takefirst
    Else
        If px <= 0 Then
            px = 1
        End If
takefirst:
        If SelectedItems Is Nothing Then
            Set ActItemObject = curFolder.Items(1)
            aClassKey = CStr(ActItemObject.Class)
            GoTo gotActItem
        ElseIf SelectedItems.Count = 0 Then
            If curFolder.Items.Count = 0 Then
                DoVerify False, "no items in " & curFolder.FolderPath _
                    & " or selected items. Will continue guessing"
                GoTo guess
            Else
                Set ActItemObject = curFolder.Items(1)
                GoTo gotActItem
            End If
        ElseIf SelectedItems.Count >= px Then
            Set ActItemObject = SelectedItems.Item(px)
gotActItem:
            If Not ActItemObject Is Nothing Then
                aClassKey = CStr(ActItemObject.Class)
                Set Item = ActItemObject
                GoTo gotOne
            Else
                GoTo noclass
            End If
        Else
noclass:
            DoVerify False, "are you trying to invalidate ObjDesc?"
            aItemClass = -1
            aItemClassName = vbNullString
        End If
    End If
    
    If aItemClass = 0 Then                                  ' guess/try better default
        On Error GoTo guess
        If SelectedItems Is Nothing Then
            If sortedItems(px) Is Nothing Then
                GoTo guess
            Else
                If sortedItems(px).Count > 0 Then
                    Set ActItemObject = sortedItems(px).Item(1)
                    aItemClass = ActItemObject.Class
                    aClassKey = CStr(aItemClass)            ' top sorted determines aItemClass
                    GoTo gotOne
                End If
            End If
        Else
            If SelectedItems.Count > 0 Then
                Set ActItemObject = SelectedItems.Item(1)
                aItemClass = ActItemObject.Class
                aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
                GoTo gotOne
            End If
        End If
        
        If curFolder.Items.Count > 0 Then
            aItemClass = curFolder.Items(1).Class
        ElseIf Not aID(px) Is Nothing Then
            If Not aID(px).idObjItem Is Nothing Then
                DoVerify aObjDsc.objItemClass = aID(px).idObjItem.Class, "Class Change!"
                Set ActItemObject = aID(px).idObjItem
                aItemClass = ActItemObject.Class
                aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
            End If
        Else
guess:
            If curFolder Is Nothing Then
                Set curFolder = olApp.ActiveExplorer.CurrentFolder
            End If
            If curFolder.Items.Count = 0 Then
                Set curFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
                If curFolder.Items.Count = 0 Then
                    DoVerify False
                End If
            End If
                aItemClass = curFolder.Items.Item(1).Class
                aClassKey = CStr(aItemClass)                ' first selected determines aItemClass
        End If
    End If  ' guess/try better default
    ' if we ProcCall into a new kind of Folder,
    ' forget previous (Additional) Rule.clsObligMatches.aRuleString
gotOne:
    If px <= 0 Then
        px = 1
    End If
    aItmIndex = WorkIndex(px)
    Call GetITMClsModel(ActItemObject, px)
    Call aItmDsc.SetDscValues(Item, withValues:=withValues, aRules:=sRules)

    If DftItemClass <> aObjDsc.objItemClass Then
        ExtendedAttributeList = vbNullString
        If sRules Is Nothing Then
            GoTo getRule
        End If
    Else
getRule:
        If sRules Is Nothing Then
            If UserRule Is Nothing Then
                If D_TC.Exists(aClassKey) Then
                    Set sRules = D_TC.Item(aClassKey).objClsRules
                End If
            Else
                If UserRule.ARName = aObjDsc.objItemClassName Then
                    Set sRules = UserRule
                    msg = "Re-using UserRule without changes, class " _
                            & aObjDsc.objItemClassName _
                            & " Type " & aObjDsc.objTypeName
                    Reused = "Reused"
                    GoTo FuncExit
                End If
            End If
        Else
            If LenB(aItemClassName) = 0 Then
                aItemClassName = aObjDsc.objItemClassName
            Else
                DoVerify aItemClassName = aObjDsc.objItemClassName, "class name in aObjDsc messed up"
            End If
            If sRules.ARName = aObjDsc.objTypeName Then
                msg = "Re-using sRules without changes, class " _
                            & aObjDsc.objItemClassName _
                            & " Type " & aObjDsc.objTypeName
                Reused = "Reused"
                GoTo FuncExit    ' same class as before: optimize redundant work
            End If
        End If
    End If
    
    ' this is the default for all of them, but deltas are ok
    ' derive from Folder defaultitem class
    DftItemClass = aObjDsc.objItemClass
    
    ' sRules and DftItemTypeName and objItemClassName determined
    
    ' Call aItmDsc.UpdItmClsDetails(ActItemObject)
    msg = "BestObjProps says: " & Reused & " sRules for class " _
                            & aObjDsc.objItemClassName _
                            & " Type " & aObjDsc.objTypeName _
                            & " ItemObject: " & ActItemObject.Subject _
                            & ", Folder: " & curFolder.FolderPath

FuncExit:
    Call LogEvent(msg, eLmin)
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.BestObjProps

'---------------------------------------------------------------------------------------
' Method : Sub logDecodedProperty
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub logDecodedProperty(p1 As String, diffStr As String)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "ItemOpsOL.logDecodedProperty"
    Call DoCall(zKey, tSub, eQzMode)

Dim showThisDetail As String

    showThisDetail = aItmIndex & ", Index=" & RString(AttributeIndex, 3) & diffStr _
                    & IgString & PropertyNameX _
                    & "="
    If Len(p1) > 256 Then
        showThisDetail = showThisDetail & Quote(Replace(Left(p1, 256), vbCrLf, vbCrLf & "        ")) _
                    & vbCrLf & "        ... (cut at 256 of " & Len(p1) & ")"
    Else
        showThisDetail = showThisDetail & Quote(Replace(p1, vbCrLf, vbCrLf & "        "))
    End If
    AllDetails = AllDetails & showThisDetail & vbCrLf
    Call LogEvent("DBg: Item=" & showThisDetail, eLnothing)

zExit:
    Call DoExit(zKey)

End Sub ' ItemOpsOL.logDecodedProperty

'---------------------------------------------------------------------------------------
' Method : Sub DefObjDescriptors
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Define or use existing Object Descriptor for MapiObject
'---------------------------------------------------------------------------------------
Sub DefObjDescriptors(Item As Object, px As Long, withValues As Boolean, Optional withAttributeSetup As Boolean = True)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.DefObjDescriptors"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

' keep consistent with Sub BestObjProps *** *** ???
' ====================================================
' objitem As Object                 current objects
' aID(1 To 2) As cObjDsc            corresponding set of properties
' aDecProp c(1 To 2) As Collection   of Class cAttrDsc
' D_TC As Dictionary                of all Object Descriptors

Dim aClassKey As String
Dim SD As String
Dim specialTypName As String

    aPindex = px                            ' 1..2, modified for subtypes of appointment types: 3..4
    targetIndex = aPindex
    sourceIndex = targetIndex - 1           ' using this for cloning
    
    aBugTxt = "Get object Descriptor from Class"
    Call Try
    aClassKey = CStr(Item.Class)            ' basic name without subtype
    If Catch Then
        Set Item = Nothing
        GoTo ProcReturn
    End If
    Call GetITMClsModel(Item, aPindex)      ' may change aObjDsc
    
    If px < 3 Then                          ' make or Clone item subtypes
        If Not aObjDsc Is aOD(px) Then
            Set aOD(px) = aObjDsc
        End If
    Else                                    ' make or Clone item subtypes
        Stop ' ??? !!! do some work from this to next ???
        ' this code is needed to decode the subtypes of appointment types
        ' all if-parts below should occur only if px>2, else is normal
        If Item.Class = olRecurrencePattern Then
            Set Item = GetAobj(px, -1)
            specialTypName = DecodeObjectClass(getValues:=withValues)
            TotalPropertyCount = 0          ' determine new situation, attribs for
                                            ' recurrences +count of exceptions
        ElseIf Item.Class = olException Then
            Set Item = GetAobj(px, -1).Exception.Item(1)
            specialTypName = DecodeObjectClass(getValues:=withValues)
        Else
            TotalPropertyCount = Item.ItemProperties.Count
        End If
        If LenB(specialTypName) = 0 Then
            SD = vbNullString
        Else
            SD = GetObjectTypeExtension(Item)
        End If
        If aObjDsc Is Nothing Then          ' we have not had this object type before:
            aOD(0).objMaxAttrCount = 0      ' clear previous class data
            aOD(0).objDumpMade = -1
            aID(0).idAttrCount = -1
            
            Set aObjDsc = D_TC.Item(aClassKey)
            Call aObjDsc.ODescClone(aClassKey & SD, aItmDsc)
        End If
    End If                                  ' else part of clone or make subtype
    
    DoVerify Not aObjDsc Is Nothing, "aObjDsc is Nothing???"
    
    With aObjDsc
        aBugVer = aOD(px).objItemClass = Item.Class
        If DoVerify(aBugVer, _
                "** aOD(px).objItemClass <> Item.Class") Then
            aOD(px).objItemClass = Item.Class ' ????????? not a good idea!
        End If
        Set aID(px).idObjItem = Item
        aBugVer = AllPublic.SortMatches = .objSortMatches
        If DoVerify(aBugVer, _
            "** should have been set when creating aObjDsc ???") Then
                .objSortMatches = SortMatches
        End If                              ' fix on failed assertion
    End With ' aObjDsc
        
    If withAttributeSetup Then              ' (in/out: aID(px)) create dynamically or re-use
        Call SetupAttribs(Item, px, withValues)
        Set aObjDsc = aOD(px)
    Else
        aID(px).idEntryId = vbNullString
    End If
    
    With aObjDsc
        ' at this time, iRules is Nothing, or iRules.RuleInstanceValid  is always false
        
        ' Superrelevant default, also sort rules from here, not changing for instance
        If aID(aPindex) Is Nothing Then     ' observe: aPindex may have changed from px
            DoVerify False, _
                "aID( " & aPindex & ") must never be Nothing: remove ifpart if no hit"
        Else
            If LenB(aOD(aPindex).objDefaultIdent) = 0 Then
                If sRules Is Nothing Then
                    MainObjectIdentification = dftRule.clsObligMatches.CleanMatches(0)
                Else
                    .objDftMatches = Trim(sRules.clsObligMatches.aRuleString)
                    If isEmpty(sRules.clsObligMatches.CleanMatches) Then
                        MainObjectIdentification = dftRule.clsObligMatches.CleanMatches(0)
                    Else
                        MainObjectIdentification = sRules.clsObligMatches.CleanMatches(0)
                    End If
                End If
            Else
                MainObjectIdentification = aOD(aPindex).objDefaultIdent
            End If
        End If
        
        If LenB(DontCompareListDefault) = 0 Then
            DontCompareListDefault = Trim(dftRule.clsNeverCompare.aRuleString)
            ' very likely done already: just to make sure, change DontCompareListDefault
            ' changes to Class-Specific Rules (for new classes only)
            Call aID(aPindex).UpdItmClsDetails(Item)
        End If
        .objDefaultIdent = MainObjectIdentification
    End With 'aObjDsc
    
    PropertyNameX = MainObjectIdentification
    If aID(aPindex).idAttrDict Is Nothing Then ' we have Not withValues
        GoTo ProcReturn
    End If
    Set aTD = GetAttrDsc(PropertyNameX)    ' sets up iRules or uses existing
    If aTD Is Nothing Then
        GoTo ProcReturn
    End If
    
    ' we can't do CheckAllRulesInList because only one iRules defined here
    If iRules Is Nothing Then
        Set iRules = aTD.adRules
        Call iRules.CheckAllRules(PropertyNameX, "MainID: ")
    End If
    
    'On Error GoTo 0
    If BaseAndSpecifiedDiffer Then
        If px < 3 Then                      ' =working on parent, then move on to exception/occurrence:
            DoVerify False, " we may never need this *** ???"
            Call DefObjDescriptors(Item, px + 2, withValues:=False)
            TotalPropertyCount = aID(aPindex).idAttrDict.Count - 1
            aPindex = aPindex - 2           ' back to standard item
            SpecialObjectNameAddition = vbNullString
            Set aTD = GetAttrDsc(PropertyNameX) ' sets up iRules
        Else
            DoVerify Item.ItemProperties.Count = TotalPropertyCount
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.DefObjDescriptors

'---------------------------------------------------------------------------------------
' Method : Function GetObjectTypeExtension
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetObjectTypeExtension(Item As Object) As String

Const zKey As String = "ItemOpsOL.GetObjectTypeExtension"
    Call DoCall(zKey, tFunction, eQzMode)

    If Item.RecurrenceState = olApptException Then
        GetObjectTypeExtension = "#E"
        aOD(aPindex).objNameExt = GetObjectTypeExtension
    ElseIf Item.RecurrenceState = olApptOccurrence Then
        GetObjectTypeExtension = "#O"
        aOD(aPindex).objNameExt = GetObjectTypeExtension
    Else
        GetObjectTypeExtension = "#B" ' search for properties in "short" AppointmentItem
        ' NO:: aID(aPindex).objNameExt = GetObjectTypeExtension
    End If

zExit:
    Call DoExit(zKey)

End Function ' ItemOpsOL.GetObjectTypeExtension

'---------------------------------------------------------------------------------------
' Method : Sub NameCheck
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub NameCheck(Ci As ContactItem)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.NameCheck"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim ImpossibleNames As Variant
Dim AdelName As Variant
Dim hilf As String
Dim i As Long
    ImpossibleNames = split("AG KG GmBH DE der v. Trier Trier-Land", b)
    AdelName = split("von_der von_dem van_den van_der van_dem van " _
                & "von vom vor zu zum zur an")
    For i = LBound(ImpossibleNames) To UBound(ImpossibleNames)
        If UCase(Ci.Lastname) = UCase(ImpossibleNames(i)) Then
            If InStr(UCase(Ci.FullName), ". " & UCase(ImpossibleNames(i))) > 0 _
            Then
                Call SetAttributeByName(Ci, 2, "FileAs", _
                                        Replace(Ci.FullName, ". ", "."), True)
                Call SetAttributeByName(Ci, 2, "FullName", Ci.FileAs, True)
                Call SetAttributeByName(Ci, 2, "LastName", "###", True)
            ElseIf LenB(Trim(Ci.FullName)) > 0 Then
                Call SetAttributeByName(Ci, 2, "FileAs", Ci.FullName, True)
                Call SetAttributeByName(Ci, 2, "LastName", Ci.CompanyName, True)
            Else
                Ci.FileAs = Ci.Firstname & b & Ci.Lastname
                Ci.Lastname = Ci.Firstname & b & Ci.Lastname
                Ci.Firstname = vbNullString
                Ci.CompanyName = Ci.FileAs
            End If
            GoTo DidMod
        ElseIf InStr(UCase(Ci.FullName), ". " & UCase(ImpossibleNames(i))) > 0 Then
            Ci.FileAs = Ci.Firstname & ", " & Ci.Lastname
            Ci.Lastname = Ci.FileAs
            Ci.Firstname = "###"
            Ci.CompanyName = Ci.FileAs
            GoTo DidMod
        ElseIf InStr(UCase(Ci.Lastname), b & UCase(ImpossibleNames(i))) > 0 _
            And InStr(UCase(Ci.FullName), UCase(Ci.Lastname)) > 0 _
        Then
            Ci.Lastname = "###"
            Ci.Firstname = "###"
            Ci.CompanyName = Ci.FileAs
            GoTo DidMod
        ElseIf Ci.Firstname = Ci.FullName _
            Or Ci.Lastname = Ci.FullName Then
            Ci.Firstname = "###"
            Ci.Lastname = "###"
            Ci.CompanyName = Ci.FileAs
            GoTo DidMod
        ElseIf InStr(Ci.Firstname, ",") > 0 Then
            Ci.FileAs = Ci.FullName
            hilf = Trunc(1, Ci.Firstname, ",")
            Ci.Firstname = Ci.Lastname
            Ci.Lastname = hilf
            GoTo DidMod
        End If
    Next i
    For i = LBound(AdelName) To UBound(AdelName)
        hilf = Replace(AdelName(i), "_", b)
        If InStr(UCase(Ci.Lastname), UCase(hilf)) = 1 Then
            Ci.Lastname = Mid(Ci.Lastname, Len(hilf) + 1)
            Ci.FileAs = Ci.Lastname & ", " & Ci.Firstname _
                        & b & Ci.MiddleName & b & hilf
            Ci.MiddleName = hilf
            GoTo DidMod
        End If
    Next i
    GoTo ProcReturn
DidMod:
    MPEchanged = True  ' and loop end

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.NameCheck

'---------------------------------------------------------------------------------------
' Method : Sub RpStackAndLog
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub RpStackAndLog(px As Long, arP As RecurrencePattern)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.RpStackAndLog"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

' RecurrencePattern has no Itemproperties, do the ones we know of
Dim eXA As Long
Dim Except As Outlook.Exception
Dim rTypeshowValue As String
Dim aIDa As cItmDsc

    Set aIDa = aID(px)                      ' add  additional properties

    Select Case arP.RecurrenceType
    Case olRecursDaily
        rTypeshowValue = "Daily"
    Case olRecursWeekly
        rTypeshowValue = "Weekly"
    Case olRecursMonthly
        rTypeshowValue = "Monthly"
    Case olRecursMonthNth
        rTypeshowValue = "every N Months"
    Case olRecursYearly
        rTypeshowValue = "Yearly"
    Case olRecursYearNth
        rTypeshowValue = "every N Years"
    Case Else
        rTypeshowValue = "# unknown " & arP.RecurrenceType
    End Select
    
    Call Try(allowAll)                   ' Try anything, autocatch, Err.Clear
    Call StackPropertyAndLog(px, "==========", "Start of Recurrence Pattern")   ' 1
    Call StackPropertyAndLog(px, "RecurrenceType", rTypeshowValue)              ' 2
    aTD.adDecodedValue = arP.RecurrenceType  ' raw value restored
    Call StackPropertyAndLog(px, "PatternStartDate", arP.PatternStartDate)      ' 3
    Call StackPropertyAndLog(px, "PatternEndDate", arP.PatternEndDate)          ' 4
    Call StackPropertyAndLog(px, "StartTime", arP.starttime)                    ' 5
    Call StackPropertyAndLog(px, "Interval", arP.Interval)                      ' 6
    Call StackPropertyAndLog(px, "Regenerate", arP.Regenerate)                  ' 7
    Call StackPropertyAndLog(px, "NrOfExceptions", arP.Exceptions.Count)        ' 8
    Call StackPropertyAndLog(px, "DayOfWeekMask", arP.DayOfWeekMask)            ' 9
    Call StackPropertyAndLog(px, "DayOfMonth", arP.DayOfMonth)                  ' 10
    Call StackPropertyAndLog(px, "MonthOfYear", arP.MonthOfYear)                ' 11
    Call StackPropertyAndLog(px, "Instance", arP.Instance)                      ' 12
    eXA = 0
    If arP.Exceptions.Count > 0 Then
        ExceptionProcessing = True
    End If
    For Each Except In arP.Exceptions
        eXA = eXA + 1
        Call StackPropertyAndLog(px, "==========" & eXA, _
                                    "Recurrence Exception " & eXA)              ' ex 1
        Call StackPropertyAndLog(px, "ExDeleted" & eXA, Except.Deleted)         ' ex 2
        Call StackPropertyAndLog(px, "ExOriginalDate" & eXA, _
                                    Except.OriginalDate)                        ' ex 3
        Set Except = Nothing                                                ' IMPORTANT
    Next Except

FuncExit:
    Catch
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.RpStackAndLog

'---------------------------------------------------------------------------------------
' Method : Sub SetPropertyList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetPropertyList(aPropName As String)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.SetPropertyList"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim TimeIn As Variant
    
    On Error GoTo newone
    If Not ExceptionProcessing Then             ' aTD is set correctly
        Set aTD = aID(aPindex).idAttrDict.Item(aPropName).Item
    End If
    If aTD.adName <> aPropName Then
        DoVerify False
    Else
        GoTo reUse
newone:
    End If
reUse:
    If W Is Nothing Then
        GoTo gds
    End If
    If aOD(aPindex).objDumpMade < 1 Then ' omit when done
gds:
        If DebugLogging Then
            TimeIn = Timer
            Debug.Print Format(TimeIn, "0#####.00"), _
                        "Finding DescriptorStrings for " & PropertyNameX
        End If
        Call SplitDescriptor(aTD)
        If DebugLogging Then
            Debug.Print , Timer - TimeIn, "finished Finding DescriptorStrings for " _
                    & PropertyNameX
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.SetPropertyList

'---------------------------------------------------------------------------------------
' Method : Sub SetupAttribs
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetupAttribs(Item As Object, px As Long, withValues As Boolean)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.SetupAttribs"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                          & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcRet
    End If
    Recursive = True
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim SD As String
Dim newPos As Long
Dim baseID As cItmDsc
Dim otherID As cItmDsc
Dim nMax As Long
Dim nMin As Long
Dim aClass As OlObjectClass
Dim aTypeName As String
Dim aClassKey As String
Dim BasePx As Long
Dim otherPx As Long
Dim thisObjDsc As cObjDsc
Dim thisItmDsc As cItmDsc
Dim Reused As String
Dim tracing As Boolean

    tracing = True                      ' just checking new design ???
    aPindex = px
    BasePx = px
    If BasePx > 2 Then
        DoVerify False, "Design verify need for SD"
        BasePx = BasePx - 2
    End If
    
    aClassKey = CStr(Item.Class)
    aTypeName = TypeName(Item)
    
    If aItmDsc Is Nothing Then
        GoTo NotGoodForReuse
    End If
    
    If aItmDsc.idEntryId = Item.EntryID Then
        Set aID(aPindex) = aItmDsc
        DoVerify aBugVer, "corrected aID: using aItmDsc ???"
        Reused = "EntryId <> aItmDsc.idEntryId"
        aItmDsc.idEntryId = Item.EntryID
        GoTo NotGoodForReuse
    Else
        Reused = "EntryId failed in aItmDsc"
        GoTo NotGoodForReuse
    End If
    If aItmDsc.idObjDsc Is Nothing Then
        Reused = Reused & "no idObjDsc"
        GoTo NotGoodForReuse
    ElseIf aID(aPindex).idEntryId <> aItmDsc.idEntryId Then
        If LenB(aID(aPindex).idEntryId) = 0 Then
            If LenB(aItmDsc.idEntryId) > 0 Then
                DoVerify False, "Design check ??? is reuse still possible?"
                Reused = "aItmDsc reused "
                Set aID(aPindex) = aItmDsc
            Else
                Reused = "aItmDsc has no idEntryId, no reuse possible"
                GoTo NotGoodForReuse
            End If
        Else
            GoTo NotGoodForReuse
        End If
    ElseIf aItmDsc.idObjDsc.objClassKey <> aID(aPindex).idObjDsc.objClassKey Then
        Reused = "Class changed from " & aItmDsc.idObjDsc.objClassKey & " to " & aID(aPindex).idObjDsc.objClassKey
        GoTo NotGoodForReuse
    End If
    If aID(aPindex) Is Nothing Then
        DoVerify Not aID(aPindex) Is Nothing, "design check aID(aPindex) IsNot Nothing ???"
        Set aID(aPindex) = aItmDsc
        Reused = "using aItmDsc"
    Else
        If aID(aPindex).idEntryId <> Item.EntryID Then
            aBugVer = LenB(aID(aPindex).idEntryId) > 0
            If aBugVer Then
                Debug.Print " in aID:    " & aID(aPindex).idEntryId
                Debug.Print " aItmDsc:   " & aItmDsc.idEntryId
                Debug.Print " Item:      " & Item.EntryID
                DoVerify aBugVer, "design check: corrected because aID(aPindex).idEntryId empty ???"
                Set aID(aPindex) = aItmDsc
           End If
         End If
    End If
    Set thisItmDsc = aID(aPindex)
    
    Set thisObjDsc = findObjDsc(aClassKey, Reused)
    If thisObjDsc Is Nothing Then
        GoTo NotGoodForReuse
    End If
    Set thisItmDsc.idObjDsc = aObjDsc
       
    Set baseID = aID(BasePx)
    If BasePx = 1 Then
        otherPx = 2
    ElseIf BasePx = 2 Then
        otherPx = 1
    ElseIf isEmpty(thisItmDsc) Then                 ' 3 or 4 not decoded, but needed now
        GoTo NotGoodForReuse
    End If
    If aPindex = 3 Then
        Set thisItmDsc = aID(4)
    ElseIf aPindex = 4 Then
        Set thisItmDsc = aID(3)
    End If
    
    Set otherID = aID(otherPx)                      ' Item class can be decoded before,
    
    If baseID Is Nothing Then
        GoTo DontHaveBase
    End If
    If otherID Is Nothing Then
        DoVerify BasePx = 1, "why is otherId=Nothing for BasePx = 1 ???"
        otherPx = 0
    End If
    
    If baseID.idObjItem Is Nothing Then             ' re-Use other Dictionary with new values
        DoVerify aPindex = 1, "design check ??? baseID.idObjItem (apindex=1?) is new?"
        Set thisItmDsc.idObjDsc = thisObjDsc        ' Class defaults are "constants"
        Set thisItmDsc.idObjItem = Item
        If Not otherID Is Nothing Then
            Set baseID.idAttrDict = otherID.IDictClone(True)    ' clone and get new values
            If DebugLogging Then
                DoVerify baseID.idObjItem Is Item, "Testing only during design ???"
            End If
            Set baseID.idObjItem = Item             ' clone must point to Item to SetupAttribs for
            Reused = " cloned ID " & otherID.idPindex
            Set thisObjDsc = thisItmDsc.idObjDsc
            If otherPx <> BasePx Then
                Reused = Reused & ", alias for Item " & aPindex
            End If
            GoTo FuncExit
        End If
    End If
    
    If baseID.idObjItem Is Item Then                ' but must also be same item
        aBugVer = thisObjDsc Is thisItmDsc.idObjDsc
        If DoVerify(aBugVer, "Design check thisObjDsc Is thisItmDsc.idObjDsc ???") Then
            Set thisObjDsc = thisItmDsc.idObjDsc
        End If
        If otherPx = 0 Then
            Reused = Reused & "|New Item, Rule " & aPindex & " is using default Class Rule"
        Else
            Reused = Reused & "|Reuse Item " & otherPx
            If otherPx <> BasePx Then
                Reused = Reused & ", alias for Item " & aPindex
            End If
        End If
        If thisItmDsc.idAttrDict Is Nothing Then
            aBugVer = thisItmDsc.idAttrCount = 0
            DoVerify aBugVer, "Design check thisItmDsc.idAttrCount = 0 ???"
            Set thisItmDsc.idAttrDict = New Dictionary
            Reused = Reused & "|new Dictionary"
            GoTo NotGoodForReuse
        End If
        GoTo FuncExit
    Else                                            ' not the same item, but same class?
        If baseID.idObjDsc.objClassKey = aClassKey Then
            If thisItmDsc Is Nothing Then
                Set thisItmDsc = New cItmDsc        ' clone the baseId
                If baseID.idAttrCount < Item.ItemProperties.Count Then
                    DoVerify False, "check what happens if property count has changed. The new one is larger"
                End If
                If baseID.idAttrCount > Item.ItemProperties.Count Then
                    DoVerify False, "check what happens if property count has changed. The new one is smaller"
                End If
                thisItmDsc.idAttrCount = baseID.idAttrCount
                Set thisItmDsc.idObjDsc = baseID.idObjDsc
                thisItmDsc.idEntryId = Item.EntryID ' and init thisItmDsc
                Set thisItmDsc.idObjItem = Item
                thisItmDsc.idPindex = aPindex
            End If
            If withValues Then
                Set thisItmDsc.idAttrDict = baseID.IDictClone(False)    ' re-Use Dictionary without values
                Reused = " cloned ID w/o values" & baseID.idPindex
                GoTo UseID
            Else
                Set thisItmDsc.idAttrDict = Nothing
                thisItmDsc.idAttrCount = inv
                DoVerify True, "attribute setup results in Nothing-Dictionary because withValues=False ???"
                GoTo noValues
            End If
            Set thisItmDsc.idAttrDict = baseID.IDictClone(False)    ' re-Use Dictionary without values
            Reused = " cloned ID w/o values" & baseID.idPindex
            GoTo UseID
        Else                                        ' neither class nor item is the same
            Set thisItmDsc = Nothing
            Reused = Reused & "|base Key <> aClassKey, new thisItmDsc"
            GoTo NotGoodForReuse
        End If
    End If
    
DontHaveBase:
    If isEmpty(thisItmDsc) Then
        GoTo NotGoodForReuse
    End If
    
    If Not thisItmDsc.idObjItem Is Item Then        ' check for re-usability
        Reused = Reused & "|new Item"
        GoTo NotGoodForReuse
    End If
    
    If thisItmDsc.idAttrDict Is Nothing Then        ' entirely new object description
        Reused = Reused & "|no AttrDict"
        GoTo nDsc
    End If
    If thisItmDsc.idAttrDict.Count < 2 Then         ' entirely new object description
nDsc:
        Reused = Reused & "|getting new rules"
        Set sRules = Nothing                        ' new rules will be determined
        GoTo NotGoodForReuse                        ' no point in cloning the original
    End If
    
    If withValues Then
        aBugVer = thisItmDsc.idAttrDict.Count >= thisItmDsc.idObjItem.ItemProperties.Count
        If DoVerify(aBugVer, _
            "design check thisItmDsc.idAttrDict.Count >= thisItmDsc.idObjItem.ItemProperties.Count ???") Then
            GoTo NotGoodForReuse
        End If
        aBugVer = thisItmDsc.idObjDsc Is thisObjDsc
        If DoVerify(aBugVer, "design check thisItmDsc.idObjDsc Is thisObjDsc check ???") Then
            GoTo NotGoodForReuse
        End If
    End If
    
    GoTo FuncExit                                   ' for this Item Attributes have been setup with values
    
NotGoodForReuse:
    If aObjDsc Is Nothing Then
        DoVerify False, "design check, aObjDsc Is Nothing ??? (Called via GetItmClsModel)"
        Set thisObjDsc = New cObjDsc
        thisObjDsc.objClassKey = aClassKey
        Set thisObjDsc.objClsRules = sRules         ' set if usable
        D_TC.Add aClassKey, thisObjDsc              ' not checking for collision because it will cause deadly error anyway
        Set aObjDsc = thisObjDsc
    Else
        aBugVer = aClassKey = CStr(Item.Class) & aObjDsc.objNameExt
        DoVerify aBugVer, "design check Item.Class = aClassKey"
        Set thisObjDsc = aObjDsc
    End If
    
    If isEmpty(thisItmDsc) Or thisItmDsc Is Nothing Then
NokItmDsc:
        AllProps = True                             ' must decode all properties if new
        Set thisItmDsc = New cItmDsc                ' uses cObjDsc
        Set thisItmDsc.idObjItem = Item
Reusing:
        Set thisItmDsc.idObjDsc = aObjDsc
        Set aID(aPindex) = thisItmDsc               ' without any content
        thisItmDsc.idPindex = aPindex
        thisItmDsc.idEntryId = Item.EntryID
    Else                                            ' check for correct thisItmDsc
        aBugVer = aClassKey = CStr(Item.Class) & aObjDsc.objNameExt
        If DoVerify(aBugVer, "design check Item.Class = aClassKey") Then GoTo NokItmDsc
        
        aBugVer = thisItmDsc.idObjDsc Is aObjDsc
        If DoVerify(aBugVer, "idObjDsc Is aObjDsc ???") Then GoTo NokItmDsc
        
        aBugVer = thisItmDsc.idObjItem Is Item
        If DoVerify(aBugVer, "idObjItem Is Item ???") Then
            If aItmDsc.idAttrCount > 0 Then         ' try reusing dictionary (=structure) without values
                Set thisItmDsc.idAttrDict = aItmDsc.IDictClone(False)
            End If
            If thisItmDsc.idAttrDict.Count < 2 Then
                GoTo NokItmDsc
            Else
                 Reused = Reused & "|cloned AttrDict"
                GoTo Reusing
            End If
        End If
        
        aBugVer = aID(aPindex) Is thisItmDsc
        If DoVerify(aBugVer, "aID(aPindex) Is thisItmDsc ???") Then GoTo NokItmDsc
        
        aBugVer = thisItmDsc.idPindex = aPindex
        If DoVerify(aBugVer, "idPindex = aPindex ???") Then GoTo NokItmDsc
        
        aBugVer = thisItmDsc.idEntryId = Item.EntryID
        If DoVerify(aBugVer, "idEntryId = Item.EntryID ???") Then GoTo NokItmDsc
    End If
    
    aBugVer = aID(aPindex) Is thisItmDsc
    If DoVerify(aBugVer, "design check ???, could fail if index =3 or 4") Then
        Set aID(aPindex) = thisItmDsc                   ' without any values
    End If
    
    If baseID Is Nothing Then
        Set baseID = aID(BasePx)                        ' may be Nothing if new class
    Else
        aBugVer = aID(aPindex) Is baseID
        DoVerify aBugVer, "design check aID(aPindex) Is baseID ??? could fail if index =3 or 4"
    End If
    If Not BasePx = aPindex Then
        SD = thisObjDsc.objNameExt                      ' use non-default Class key, else: leave empty
        aClassKey = CStr(Item.Class) & SD
    End If
    
    If thisObjDsc.objClsRules Is Nothing Then
        Reused = Reused & "|defaulting rules"
        Set sRules = dftRule.AllRulesClone(ClassRules, thisObjDsc, withMatchBits:=False)
    End If
    If aOD(aPindex) Is Nothing Then
        Set aOD(aPindex) = aObjDsc
    End If
    If aOD(aPindex).objClsRules Is Nothing Then
        Call SetCriteria                                ' sRules evaluated for Class
    End If
       
    If thisItmDsc.idRules Is Nothing Then
        Set iRules = New cAllNameRules                  ' sets all specific rules to non-nothings
        Call iRules.AllRulesCopy(InstanceRule, sRules, withMatchBits:=False)
        Reused = Reused & "|copied Class=" & aObjDsc.objClassKey & " sRules as iRules"
        Set thisItmDsc.idRules = iRules
    Else
        Reused = Reused & "|reusing sRules"
        Set iRules = thisItmDsc.idRules
    End If
     
    aBugVer = thisItmDsc.idObjDsc Is thisObjDsc
    If DoVerify(aBugVer, "design check thisItmDsc.idObjDsc Is thisObjDsc next assignment is needed ???") Then
        Set thisItmDsc.idObjDsc = thisObjDsc            ' Link as parent
    End If
    aBugVer = thisItmDsc.idObjItem Is Item
    If DoVerify(aBugVer, "design check thisItmDsc.idObjItem Is Item next assignment is needed ???") Then
        Set thisItmDsc.idObjItem = Item
    End If
    
    thisItmDsc.idEntryId = thisItmDsc.idObjItem.EntryID
    thisItmDsc.idTimeValue = 0                          ' this indicates that we did not set aID with values yet
    
UseID:
    With thisItmDsc
        DoVerify .idEntryId = .idObjItem.EntryID, " design check ???"
        If Not withValues Then
            GoTo noValues
        End If
        If .idAttrDict Is Nothing Then
            aBugVer = aPindex > 0
            If DoVerify(aBugVer, "for aPindex=0 there is no Dictionary or object item") Then
                GoTo FuncExit
            Else    ' verify checks OK
                Set .idAttrDict = New Dictionary        ' using item(0) as TypeClassName
                .idAttrDict.Add aClassKey, thisItmDsc   ' the new parent: cItmDsc, NOT a cAttrDsc!!!!
                DoVerify .idEntryId = .idObjItem.EntryID, "design check .idEntryId = .idObjItem.EntryID ???"
            End If
        Else
            If .idAttrDict.Count = 0 Then
                .idAttrDict.Add aClassKey, thisItmDsc       ' the new parent: cItmDsc, NOT a cAttrDsc!!!!
            End If
        End If
        If .idAttrDict.Count < 2 Then                       ' WithValues=False is ignored ???
            ' reset cloned value of clsNeverCompare to default
            If iRules.clsNeverCompare.aRuleString <> DontCompareListDefault Then
                iRules.clsNeverCompare.ChangeTo = DontCompareListDefault ' reset iRules, unusual case
            End If
            
            If TotalPropertyCount = 0 Then
                aBugTxt = "Get Class of Item"
                Call Try
                aClass = Item.Class
                Catch
                Select Case Item.Class
                Case olRecurrencePattern
                    DoVerify False, "design check for olRecurrencePattern ??? "
                    Call AttrExtend("==========")
                    Call AttrExtend("RecurrenceType")
                    Call AttrExtend("DayOfMonth")
                    Call AttrExtend("PatternStartDate")
                    Call AttrExtend("PatternEndDate")
                    Call AttrExtend("StartTime")
                    Call AttrExtend("Interval")
                    Call AttrExtend("Regenerate")
                    Call AttrExtend("NrOfExceptions")
                    Call AttrExtend("DayOfWeekMask")
                    Call AttrExtend("MonthOfYear")
                Case olException
                    DoVerify False, "design check for olException ??? "
                    Call AttrExtend("----------")
                    Call AttrExtend("ExDeleted")
                    Call AttrExtend("ExOriginalDate")
                Case Else
                    GoTo fillAuto
                End Select
            Else    ' fill propertyattributes automatically
fillAuto:
                nMax = thisObjDsc.objMaxAttrCount
                nMin = thisObjDsc.objMinAttrCount
                If aPindex > 2 Then                 ' copy to base: make room
                    DoVerify False, "old code, check if it still makes sense ???"
                    Set baseID = aID(aPindex - 2)
                    newPos = baseID.idAttrCount + 1
                    nMin = Max(thisObjDsc.objMinAttrCount, Item.ItemProperties.Count)
                    nMax = Max(nMin, thisObjDsc.objMaxAttrCount) ' max allowed index +1
                ElseIf thisObjDsc.objMaxAttrCount <> Item.ItemProperties.Count Then
                    nMin = Max(thisObjDsc.objMinAttrCount, Item.ItemProperties.Count)
                    nMax = Max(nMin, thisObjDsc.objMaxAttrCount) ' max allowed index +1
                End If
    
                If Not ShutUpMode Then
                    Call LogEvent(Format(Timer, "0#####.00") & vbTab _
                        & aPindex & ". item, starting InitAttributeSetup on " _
                        & thisObjDsc.objTypeName & "(Class " & Item.Class & "), with " _
                        & thisItmDsc.idObjItem.ItemProperties.Count _
                        & " standard properties." _
                        & vbCrLf & vbTab & vbTab & vbTab _
                        & "Subject: " & Quote(thisItmDsc.idObjItem.Subject), eLall)
                End If
                
                Call InitAttributeSetup(baseID, thisItmDsc)         ' few standard things, all Exception presets
                
                If withValues Or thisItmDsc.idTimeValue = 0 _
                        Or (nMax <> thisObjDsc.objMaxAttrCount _
                        Or nMin <> thisObjDsc.objMinAttrCount) Then
                    Call GetItemAttrDscs(Item, aPindex)             ' loops Props, generates AttrDsc in .odAddrDict , may get values
                Else
                    aBugTxt = "design check ??? Check if Previous Call was skipped correctly"
                    DoVerify thisItmDsc.idTimeValue <> 0
                End If
            End If
      ' Else: no need to set up Attribute Dictionary because .idAttrDict.Count < 2
        End If
        
        TotalPropertyCount = .idAttrDict.Count - 1
        ' ??? remove and check indentation, DoVerify TotalPropertyCount = .idAttrDict.Count - 1, _
                            "** mismatch .odItemDict.Count <> .idAttrDict.Count-1"
    End With ' thisItmDsc
noValues:
    Set iRules = Nothing ' no Attribute has been selected intentionally

FuncExit:
    Set aItmDsc = baseID
    Set aObjDsc = thisObjDsc
    If aOD(aPindex) Is Nothing Then
        aBugTxt = "** setting up new class ??? " & aObjDsc.objClassKey
        DoVerify False, aBugTxt
        Set aOD(aPindex) = aObjDsc
    Else
        If aObjDsc.objClassKey <> aOD(aPindex).objClassKey Then
            ' aBugVer = aObjDsc Is aOD(aPindex)
            ' aBugTxt = "** design check aObjDsc Is aOD(aPindex) ??? OR just a class change " _
                        & aObjDsc.objClassKey & "/" & aOD(aPindex).objClassKey
            ' If DoVerify(aBugVer, aBugTxt) Then
                Set aOD(aPindex) = aObjDsc
            ' End If
        End If
    End If
    If LenB(MainObjectIdentification) = 0 Then
        MainObjectIdentification = sRules.clsObligMatches.CleanMatches(0)
    End If
    aObjDsc.objDefaultIdent = MainObjectIdentification
    Set aID(aPindex) = thisItmDsc                       ' does not have to be = aITMDsc ???
    If aItmDsc Is Nothing Then
        Set aItmDsc = thisItmDsc
    End If
    Set aOD(aPindex) = thisObjDsc
    
    newPos = Item.ItemProperties.Count
    aItmDsc.idAttrCount = newPos
    If aOD(aPindex).objMaxAttrCount < newPos Then
        aOD(aPindex).objMaxAttrCount = newPos
        If aOD(aPindex).objMinAttrCount = 0 Then
            aOD(aPindex).objMinAttrCount = newPos
        End If
    End If
    
    Set baseID = Nothing
    Set thisObjDsc = Nothing
    Set thisItmDsc = Nothing
    Call aItmDsc.UpdItmTime
    
ProcReturn:
    If tracing Then
        Debug.Print Replace(Reused, "|", vbCrLf)    ' this line during design check Only ???
    Else
        Reused = vbNullString
    End If
    
    Call ProcExit(zErr, Reused)
    Recursive = False

ProcRet:
End Sub ' ItemOpsOL.SetupAttribs
'---------------------------------------------------------------------------------------
' Method : findObjDsc
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Find ObjDsc via key in D_TC or return Nothing
'---------------------------------------------------------------------------------------
Function findObjDsc(aClassKey As String, msg As String) As cObjDsc
    '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "ItemOpsOL.findObjDsc"
#If MoreDiagnostics Then
    Call DoCall(zKey, "Function", eQzMode)
#End If

   If D_TC.Exists(aClassKey) Then
        Set findObjDsc = D_TC.Item(aClassKey)       ' reuse object type description
        If Not sRules Is findObjDsc.objClsRules Then
            DoVerify False, "Design check ??? sRules Is'nt findObjDsc.objClsRules"
            Set sRules = findObjDsc.objClsRules      ' may reuse Rules if same class (no clone)
        End If
    Else
        Set findObjDsc = Nothing
        msg = msg & "|not a previously known class"
    End If

zExit:
    Call DoExit(zKey)
End Function ' ItemOpsOL.findObjDsc

' Make sure     aProp properly set ! or aValueType is a non-object Variant
' NOTE: this will NOT decode the attribute unless it is scalar
'---------------------------------------------------------------------------------------
' Method : ProvideAttrDsc
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Provide New cAttrDsc for any variant object; retrieve value if scalar
'          does not provide iRules, uses aOD(0).objClsRules=dftRule as classRules
'---------------------------------------------------------------------------------------
Function ProvideAttrDsc() As cAttrDsc
    ' (Optional aValueType As Long = 0, Optional Name As String = vbNullString, Optional ByRef Value As Variant = Nothing) As cAttrDsc
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.ProvideAttrDsc"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim nAttrDsc As cAttrDsc

    aBugVer = Not aProp Is Nothing And Not isEmpty(aProp)
    DoVerify aBugVer, "aProp must be set", True
    Set aProps = aProp.Parent
    aBugVer = PropertyNameX = aProp.Name
    If DoVerify(aBugVer, "** change of PropertyNameX: " & PropertyNameX & " to " & aProp.Name & " ???") Then
        PropertyNameX = aProp.Name          ' PropertyNameX used for cAttrDsc_Initialize
    End If
    aCloneMode = withNewValues              ' Model, Rules and optionally Attribute value(s)
    Set nAttrDsc = New cAttrDsc
    Call nAttrDsc.GetScalarValue            ' no parms: called without any info outside aProp/dftRule'
    
    Set ProvideAttrDsc = nAttrDsc

FuncExit:
    Set nAttrDsc = Nothing

ProcReturn:
    Call ProcExit(zErr, ProvideAttrDsc.adShowValue)

pExit:
End Function ' ItemOpsOL.ProvideAttrDsc

'---------------------------------------------------------------------------------------
' Method : Sub AttrExtend
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Extend the Attributes of Class with Extension Properties
'---------------------------------------------------------------------------------------
Sub AttrExtend(Name As String)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.AttrExtend"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    PropertyNameX = Name
    Call SetPropertyList(Name)
    Call StackAttribute

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.AttrExtend

'---------------------------------------------------------------------------------------
' Method : MakeAttributeKey
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose:     ADKey is checked vs. Attrname, or it's computed from ADName
'---------------------------------------------------------------------------------------
Function MakeAttributeKey(adName As String) As String
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "ItemOpsOL.MakeAttributeKey"
    Call DoCall(zKey, "Function", eQzMode)

Dim subtype As String
Dim i As Long

    If aID(aPindex) Is Nothing Then
        If aPindex > 1 Then
            Set aID(aPindex) = aID(1)
            subtype = aOD(aPindex).objNameExt
        End If
    End If
    
    MakeAttributeKey = PropertyNameX
    i = InStr(adName, "#")                                  ' expecting #W2, #B, #O, #R
    If i > 0 Then    ' if called using key, correct this (out!)
        subtype = Trim(Mid(adName, i, 2))
        If subtype = "#W2" Then
            isUserProperty = True
        End If
        ' remove all indications of key usage
        PropertyNameX = Trim(Replace(adName, subtype, vbNullString))
    Else
        PropertyNameX = adName
    End If
    DoVerify LenB(PropertyNameX) > 0
    
    If InStr(MakeAttributeKey, PropertyNameX) = 0 Then      ' key mismatch or empty key
        If isUserProperty Then
            subtype = "#W2"
        ElseIf isSpecialName _
            Or BaseAndSpecifiedDiffer And Not workingOnNonspecifiedItem Then
            subtype = "#B"                                  ' search for properties in "short" AppointmentItem
            isSpecialName = True
            isUserProperty = False
        End If
    Else
        If aPindex > 2 Then                                 ' just to make sure
            If aID(aPindex).idObjItem.RecurrenceState = olApptOccurrence Then
                DoVerify subtype = "#O"
            ElseIf aID(aPindex).idObjItem.RecurrenceState = olApptException Then
                DoVerify subtype = "#R"
            End If
        End If
    End If
    MakeAttributeKey = PropertyNameX & subtype              ' that's it folks

FuncExit:

zExit:
    Call DoExit(zKey)

End Function ' ItemOpsOL.MakeAttributeKey

'---------------------------------------------------------------------------------------
' Method : Function GetAttrKey
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetAttrKey(adName As String, Optional noget As Boolean = False, Optional FromIndex As Long = 0) As String
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.GetAttrKey"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim trueindex As Long

    If FromIndex = 0 Then
        FromIndex = aPindex
    End If
    
    isUserProperty = False
    aBugVer = LenB(adName) > 0
    If DoVerify(Message:="LenB(ADName) > 0") Then
        Set aTD = Nothing
        GetAttrKey = vbNullString
        GoTo ProcReturn
    End If
    GetAttrKey = MakeAttributeKey(adName)   ' check or make it
    
    If noget Then                           ' do not get (updated data for) aTd
        If Not aTD Is Nothing Then
            If aTD.adKey <> GetAttrKey Then ' but its the wrong one
                Set aTD = Nothing           ' so mark as invalid
            End If
        End If
    Else
        If aID(FromIndex).idAttrDict Is Nothing Then
            Set aTD = Nothing
            GoTo ProcReturn
        End If
        With aID(FromIndex).idAttrDict
            If .Exists(GetAttrKey) Then     ' determines dDitem. BUG: creates empty Item
                If isEmpty(.Item(GetAttrKey)) Then
                    Set aTD = Nothing
                Else
                    Set aTD = .Item(GetAttrKey)
                    trueindex = aTD.adtrueIndex
                    If aTD.adisUserAttr Then
                        GoTo isUattr
                    End If
                End If
            Else
                If Not aTD Is Nothing Then
                    If aTD.isUserProperty Then
isUattr:                                    ' this is guessing!
                       Call AppendTo(GetAttrKey, "#W2")
                       isUserProperty = True
                    End If
                End If
                Stop ' trueindex = aID(FromIndex).ItemPropFind(GetAttrKey)
            End If
        End With ' aID(fromindex).odItemDict
    End If  ' noget is false, had to get ATD

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.GetAttrKey

'---------------------------------------------------------------------------------------
' Method : Sub StackAttribute
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub StackAttribute()
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.StackAttribute"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim last1Index As String
Dim lastchar As String
Dim i As Long
Dim aKey As String
Dim nAttrDsc As cAttrDsc

    If aTD.adNr <= 0 Then                                   ' check if it is in aID(aPindex).idAttrDict already, true=yes
        If aPindex > UBound(aID) Then                       ' ??? whot???
            GoTo ProcReturn                                 ' we don't need you
        End If
        If aID(aPindex) Is Nothing Then
            Set aID(aPindex).idAttrDict = New Dictionary
        End If
        AttributeIndex = aID(aPindex).idAttrDict.Count - 1
    End If
    
    ' find out if we have this decprop ( and ignore separator ' ========== )
    If InStr(aTD.adName, "=") = 0 Then
        Set nAttrDsc = aID(aPindex).idAttrDict.Item(aTD.adKey)
        DoVerify nAttrDsc.adName = aTD.adName, "Error in ADName of aID(aPindex).idAttrDict at Position=" & AttributeIndex
        DoVerify nAttrDsc.adKey = aTD.adKey, "Error in ADKey of aID(aPindex).idAttrDict at Position=" & AttributeIndex
        ' something in this position present: check plausi: very bad if assert fails
        aBugTxt = "Error in ADtrueIndex of aID(aPindex).idAttrDict at Position=" & AttributeIndex
        DoVerify nAttrDsc.adtrueIndex = aTD.adtrueIndex
        aKey = aTD.adKey
    End If
    
    If aPindex = 2 And Not aID(1).idAttrDict Is Nothing Then
        ' may have to synchronize names (with first object)
        If aID(1).idAttrDict.Count >= AttributeIndex Then
            If aID(1).idAttrDict.Item(AttributeIndex).adName = PropertyNameX Then
               i = 0                                        ' OK, all is well, need no sync of names
            Else
                ' If DebugMode Then
                    DoVerify False, "design check aID(1).idAttrDict.Count >= AttributeIndex ??? "
                ' End If
                i = FindAttributeByName(AttributeIndex, PropertyNameX) ' find in aID(1).idAttrDict
                If i > 0 Then
                    Stop ' ???
                    AttributeIndex = i
                    ' If Not aTD.ADNr _
                        = aDecProp C(1).Item(i).ADNr _
                    Or Left(aTD.ADFormattedValue, 1) <> "*" Then
                        ' If Left(aTD.ADFormattedValue, 1) = "*" Then
                            ' pArr(1) = aDecProp(2).PropValue
                        ' Else                                ' could be a double name entry!
                           ' i = 1000 ' add at end, because > AttributeIndex always
                        ' End If
                    ' End If
                    If i < aID(2).idAttrDict.Count Then
                        If aID(2).idAttrDict.Item(i).adFormattedValue _
                        <> aTD.adFormattedValue Then
                            If aID(2).idAttrDict.Item(i).adFormattedValue = Chr(0) Then
                                aID(2).idAttrDict.Item(i).adFormattedValue _
                                    = aTD.adFormattedValue
                            Else
                                DoVerify False
                            End If
                        End If
                        pArr(2) = aTD.adFormattedValue
                    Else
                        Stop ' ???
                        ' aDecProp C(2).Add aTD                ' add at end!
                    End If
                    GoTo ProcReturn
                '   ########
                End If
                
                lastchar = Chr(0)                           ' inserting #2 into 1 !
                ' If DebugMode Then
                    DoVerify False, " here we could mess up badly ???***"
                    ' check correct unique attributeindex and aID(2).idAttrDict
                ' End If
                cMissingPropertiesAdded = cMissingPropertiesAdded + 1
                DoVerify False, " not tested! replace adecprop(1) with aTd ??? ***"
                aCloneMode = FullCopy
                Set nAttrDsc = New cAttrDsc
                Set aDecProp(1) = nAttrDsc
                Set nAttrDsc.adItem = aID(1).idObjItem
                last1Index = aID(1).idAttrDict.Item(AttributeIndex - 1).adKey
                lastchar = Right(last1Index, 1)
                If lastchar >= "a" Then
                    lastchar = Chr(Asc(lastchar) + 1)
                    last1Index = Left(last1Index, Len(last1Index) - 1)
                Else
                    lastchar = "a"
                End If
                nAttrDsc.adNr = last1Index & lastchar
                nAttrDsc.adName = PropertyNameX
                aDecProp(1).adFormattedValue = "***Missing Property*** key=" _
                                            & last1Index & lastchar
                Addit_Text = True
                pArr(1) = PropertyNameX
                pArr(2) = nAttrDsc.adFormattedValue
            End If
        ElseIf aID(1).idAttrDict.Count < AttributeIndex Then
            ' have to synchronize names (with second object)
            aCloneMode = FullCopy
            aID(1).idAttrDict.Add aDecProp(2).adKey, aDecProp(2).adictClone        ' add at end!
            GoTo ProcReturn
        End If
    End If

FuncExit:
    Set nAttrDsc = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.StackAttribute

' used for Extended properties:
Sub StackPropertyAndLog(px As Long, aName As String, AVal As String)
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.StackPropertyAndLog"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    PropertyNameX = aName
    aStringValue = AVal

    Set aTD = Nothing                   ' always new AttrDsc
    Call CreateIRule(aName)             ' using PropertyNameX inside, does SplitDescriptor to set Rules
    aTD.adFormattedValue = AVal
    aTD.adShowValue = AVal
    aTD.adInfo.iAssignmentMode = 1
    Call StackAttribute
    If InStr(aName, "=") > 0 Then
        aTD.adDecodedValue = vbNullString         ' Show sepline in excel
    Else
        aTD.adDecodedValue = AVal
    End If
    aTD.adOrigValDecodingOK = True
    
    Call SetPropertyList(aName)
    
    If displayInExcel Then
        pArr(1) = PropertyNameX
        pArr(1 + px) = aStringValue
        Call addLine(O, AttributeIndex, pArr)
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' ItemOpsOL.StackPropertyAndLog

'---------------------------------------------------------------------------------------
' Method : Function GetITMClsModel
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get or Make the Object descriptor for an ObjectClass
'          determine ClassName and TypeName from item's object class,
'          determine ItemClassProperties (like MailLike, TimeType, has ReceivedTime, ...)
'          determine parts of default Rule
'---------------------------------------------------------------------------------------
Function GetITMClsModel(Item As Object, px As Long) As cItmDsc
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "ItemOpsOL.GetITMClsModel"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction, ExplainS:="GetITMClsModel")

Dim aItmClass As OlObjectClass
Dim aItmTypeName As String
Dim RecommendedMatches As String
Dim Reused As String
Dim InitialSetup As Boolean
Dim aClassKey As String
Dim inD_TC As Boolean

    aItmTypeName = TypeName(Item)
    aItmClass = Item.Class                              ' using this to get the ClassName
    aClassKey = CStr(aItmClass)
    Set aItmDsc = New cItmDsc
    Set aItmDsc.idObjItem = Item
    
    If px = 0 Then
        px = 1
    End If
    If targetIndex = 0 Then
        targetIndex = px
    Else
        DoVerify targetIndex = px, "targetIndex<>aPindex is not OK ???"
    End If
    
    If D_TC.Count = 1 Then
        InitialSetup = True
    End If
    If D_TC.Exists(aClassKey) Then
        inD_TC = True
        Set aObjDsc = D_TC.Item(aClassKey)
        
        If aObjDsc.objClassKey = CStr(Item.Class) Then      ' correct item class ???
            Set aItmDsc.idObjDsc = aObjDsc
            
            If DoVerify(ItemValid(aItmDsc.idObjItem), "design check item gone/invalid ???") Then
                GoTo FuncExit
            End If
            If Not aOD(px) Is aObjDsc Then
                Set aOD(px) = aObjDsc
            End If
            Reused = " Reused"
            GoTo FuncExit
        Else
            aBugTxt = "Design check: Wrong Item Class ???"
            DoVerify False
            Set aItmDsc = Nothing
            Set aID(px) = Nothing
            Set aDecProp(px) = Nothing
            Set aID((px + 2)).idAttrDict = Nothing
            Set aID((px)).idAttrDict = Nothing              ' reset all previously decoded values
        End If
        
    Else                                                    ' when new, reset DontCompareListDefault
        DontCompareListDefault = Trim(dftRule.clsNeverCompare.aRuleString)
        ' Creating new Class description model
        Reused = "New"
        Set aObjDsc = New cObjDsc
    End If
    
    
    With aObjDsc
        RecommendedMatches = "Subject"                      ' Superrelevant default, override OK
        .objItemClass = aItmClass
        .objItemClassName = aItmTypeName                    ' sometimes not correct for class, set in selected case
        .objHasReceivedTime = False                         ' these are the default with EXCEPTIONS below
        .objHasHtmlBodyFlag = False
        .objHasSenderName = False
        .objHasSentOnBehalfOf = False
        Set .objSeqInImportant = New Collection
        
' stops on unverified for classes 47, 49, 53-56, certain for 56, 57: .objHasHtmlBodyFlag = False
        Select Case .objItemClass
            '----- very common ones are first ----
            Case olMail                     ' 43
                .objIsMailLike = True
                .objItemType = OlItemType.olMailItem
                .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
                .objHasHtmlBodyFlag = True
                .objHasSenderName = True
                .objHasSentOnBehalfOf = True
                .objTimeType = "SentOn"
                
                Call AppendTo(DontCompareListDefault, _
                    " HTMLBody " _
                        & "Responserequested ConferenceServerAllowExternal " _
                        & "SendUsingAccount SentOn Recipients " _
                        & "ReceivedByName CreationTime ", b)
                If LenB(aTimeFilter) = 0 Then
                    RecommendedMatches = Append("Subject SenderName SentOn ", aTimeFilter, b)
                End If
            Case olMeeting                  ' 1
                .objItemType = OlItemType.olAppointmentItem
                        ' create like this, but modify .Status = olMeeting
                        ' other .Status are
                        ' olNonMeeting (0)
                        ' olMeetingReceived (3)
                        ' olMeetingCancelled (5)
                        ' olMeetingReceivedAndCancelled (7)
                .objHasSenderName = False
                .objTimeType = "SentOn"

                DoVerify False
                ' *Time must be missing!
                DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
                Call AppendTo(DontCompareListDefault, _
                            "Ordinal CreationTime ConversationIndex Size SenderName" _
                            & "SentOn SentOnBehalfOfName", b)
                RecommendedMatches = "Start End IsRecurring Exceptions"
                                                            ' no compare of Subject: time conflicts!
            Case olAppointment          ' 26
                .objItemType = OlItemType.olAppointmentItem
                .objHasSenderName = True
                .objTimeType = "Start"

                Call AppendTo(DontCompareListDefault, _
                    "ConversationTopic", b)
                RecommendedMatches = "Subject ! Start End IsRecurring Exceptions"
            Case olContact              ' 40
                .objItemType = OlItemType.olContactItem
                .objTimeType = "LastModificationTime"

                Call AppendTo(DontCompareListDefault, _
                    " ConversationTopic LastFirst* Subject Initials", b)
                RecommendedMatches = _
                    "%FullName ! FileAs LastName FirstName MiddleName MobileTelephoneNumber " _
                    & "HomeTelephoneNumber CompanyName BusinessTelephoneNumber BusinessFaxNumber " _
                    & "OtherTelephoneNumber Email1Address Email2Address Email3Address " _
                    & "WebPage HomeAddress BusinessAddress CompanyName Birthday User2 " _
                    & "HasPicture Links Attachments"
            '----- uncommon ones  ----
            Case olNote                     ' 44
                .objItemType = OlItemType.olNoteItem
                .objTimeType = "LastModificationTime"
                RecommendedMatches = "Subject Body"
            Case olPost                     ' 45
                .objItemType = OlItemType.olPostItem
                .objHasSenderName = True
                .objTimeType = "LastModificationTime"
                DoVerify False, " we never saw this. Recipient missing. .objHasSenderName is unverified"
                                                            ' RecommendedMatches = ???
            Case olTask                     ' 48
                .objItemType = OlItemType.olTaskItem
                .objHasSenderName = True
                .objTimeType = "LastModificationTime"
                DoVerify False, "watch further"
                RecommendedMatches = "ConversationTopic Subject StartDate ! DueDate Body"
            Case olTaskRequest              ' 49
                .objItemClassName = "TaskRequest"
                .objItemType = OlItemType.olTaskItem
                .objHasHtmlBodyFlag = False
                .objIsMailLike = True
                                                            ' Not .objHasSenderName
                                                            ' Not .objHasSentOnBehalfOf
                .objTimeType = "CreationTime"
            Case olTaskRequestUpdate        ' 50
                .objItemClassName = "TaskRequestUpdate"
                .objItemType = OlItemType.olTaskItem
                .objIsMailLike = True
                .objHasHtmlBodyFlag = True
                .objHasSenderName = True
                .objHasSentOnBehalfOf = True
                .objTimeType = "SentOn"
            Case olTaskRequestAccept        ' 51
                .objItemClassName = "TaskRequestAccept"
                .objItemType = OlItemType.olTaskItem
                .objIsMailLike = True
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objTimeType = "SentOn"
            Case olTaskRequestDecline       ' 52
                .objItemClassName = "TaskRequestDecline"
                .objItemType = OlItemType.olTaskItem
                .objIsMailLike = True
                .objHasHtmlBodyFlag = True
                .objHasSenderName = True
                .objTimeType = "SentOn"
            Case olMeetingRequest           ' 53
                .objItemClassName = "MeetingRequest"
                .objItemType = OlItemType.olAppointmentItem
                .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objHasSentOnBehalfOf = False
                RecommendedMatches = "Start End IsRecurring Exceptions"
MeetingStuff:
                .objIsMailLike = True
                .objTimeType = "SentOn"
                ' All *Time must be missing from Item Compares!
                DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
                Call AppendTo(DontCompareListDefault, _
                            "CreationTime ConversationIndex Size SenderName" _
                            & "SentOn", b)
            Case olMeetingCancellation      ' 54
                .objItemClassName = "MeetingCancellation"
                .objItemType = OlItemType.olAppointmentItem
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objHasSentOnBehalfOf = True
                .objTimeType = "SentOn"
                GoTo MeetingStuff
            Case olMeetingResponseNegative  ' 55
                .objItemClassName = "MeetingResponseNegative"
                .objItemType = OlItemType.olAppointmentItem
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objTimeType = "SentOn"
                DoVerify False, "watch further"
                GoTo MeetingStuff
            Case olMeetingResponsePositive  ' 56
                .objItemClassName = "MeetingResponsePositive"
                .objItemType = OlItemType.olAppointmentItem
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objTimeType = "SentOn"
                DoVerify False, "watch MeetingResponsePositive further"
                GoTo MeetingStuff
            Case olMeetingResponseTentative ' 57
                .objItemClassName = "MeetingResponseTentative"
                .objItemType = OlItemType.olAppointmentItem
                .objHasReceivedTime = True                  ' this is one of the few EXCEPTIONS
                .objHasHtmlBodyFlag = False
                .objHasSenderName = True
                .objTimeType = "SentOn"
                DoVerify False, "watch further"
                RecommendedMatches = "Start End IsRecurring Exceptions"

gotNameIsMailLike:
                .objIsMailLike = True
                .objTimeType = "SentOn"
                ' All *Time must be missing from Item Compares!
                DontCompareListDefault = Remove(DontCompareListDefault, "*Time*", b)
                Call AppendTo(DontCompareListDefault, _
                            "Ordinal CreationTime ConversationIndex Size SenderName" _
                            & "SentOn SentOnBehalfOfName", b)
                                                            ' no compare of Subject: time conflicts!
            Case olSharing              ' 104
                .objItemType = OlItemType.olTaskItem
                .objHasSenderName = True
                .objTimeType = "SentOn"
                RecommendedMatches = "Subject ! Body"
                DoVerify False, "watch further"
                GoTo gotNameIsMailLike
            '----- unlikely ones  ----
            Case olDocument             ' 41
                .objItemType = 41       ' OlItemType.olDocument is not defined!
                .objTimeType = "SentOn"
                DoVerify False, "we never saw this. Office Document or whatever"
                ' RecommendedMatches = ???
            Case olJournal              ' 42
                .objItemType = OlItemType.olJournalItem
                .objTimeType = "SentOn"
                DoVerify False, "we never saw this. Deprecated"
                ' RecommendedMatches = ???
            Case olReport               ' 46
                .objIsMailLike = True
                .objItemType = 46       ' OlItemType.olReport is not defined!
                .objHasReceivedTime = False                 ' this is one of the few EXCEPTIONS
                .objHasHtmlBodyFlag = False
                .objHasSentOnBehalfOf = False
                .objHasSenderName = False
                .objTimeType = "LastModificationTime"
            Case olRemote               ' 47
                .objItemType = 47           ' OlItemType.olRemote is not defined!
                .objTimeType = "SentOn"
                DoVerify False, " we never saw this ."
                ' Like mail, but no BillingInformation, Body, Categories, Companies, and Mileage
                ' RecommendedMatches = ???
            Case olDistributionList     ' 69
                .objItemType = OlItemType.olDistributionListItem
                .objTimeType = "LastModificationTime"
                Call AppendTo(RecommendedMatches, "! MemberCount", b)   ' ! means: do not sort
            Case Else
                Stop ' ???
                Call LogEvent(.objItemClassName & "(" & .objItemClass _
                    & ") not expected with item " & aItmIndex _
                    & " in Folder " & curFolderPath & " is: ", eLall)
                rsp = MsgBox(.objItemClassName & vbCrLf & "Problem with unknown Objectclass " _
                    & .objItemClass & ", Try as mail?", vbYesNoCancel)
                If rsp = vbCancel Then
                    End
                ElseIf rsp = vbYes Then
                    GoTo gotNameIsMailLike
                Else
                    
                End If
                
                MainObjectIdentification = vbNullString               ' Undefined
                .objItemClassName = vbNullString
        End Select        '   Case .objItemClass
        
        .objTypeName = Remove(.objItemClassName, "Item")
        .objDftMatches = RecommendedMatches
        .objClassKey = aClassKey
        
        Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, False)
        sRules.clsNeverCompare.ChangeTo = DontCompareListDefault
        sRules.clsObligMatches.ChangeTo = RecommendedMatches
        TrueCritList = sRules.clsObligMatches.CleanMatches(0)
        MostImportantAttributes = Append(sRules.clsObligMatches.CleanMatchesString, sRules.clsSimilarities.CleanMatchesString, b)
        MostImportantProperties = split(MostImportantAttributes)
        Set .objClsRules = sRules                    ' this is only the starting point for sRules
        DoVerify sRules.ARName = .objTypeName, " used to be an assignment, needed?"
        
        Reused = "* New Object Class defined: Decoding values required for "
        Call LogEvent(Reused & .objItemClassName & "(" & .objItemClass & "), " _
                & "Maillike = " & .objIsMailLike & ", TimeType = " & .objTimeType _
                & ", RecommendedMatches=" & RecommendedMatches, eLall)
    
    End With ' aObjDsc

    If Not inD_TC Then
        D_TC.Add aClassKey, aObjDsc
        inD_TC = True
    End If
    
FuncExit:
    Set GetITMClsModel = aItmDsc
    Set aID(aPindex) = aItmDsc
    If aOD(aPindex) Is Nothing Then
        Set aOD(aPindex) = aObjDsc
    End If
    
ProcReturn:
    Call ProcExit(zErr, aObjDsc.objTypeName & Reused)

pExit:
End Function ' ItemOpsOL.GetITMClsModel

'---------------------------------------------------------------------------------------
' Method : Function ReGet
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ReGet(ByRef oItem As Object, Optional ByRef oID As String, Optional ByRef nID As String) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.ReGet"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

   
   ' de-reference the old item, then open it again
    ' otherwise Outlook might still have tidbits
    ' left from the original message

Dim oldClass As OlObjectClass
Dim Retry As Long

    E_Active.Permit = "*"
    oID = oItem.EntryID
    If ErrorCaught <> 0 Then
        GoTo badItem
    End If
    oldClass = oItem.Class
tryAgain:
    If Not oItem.Saved Then
        aBugTxt = "save original item"
        Call Try("%da die Nachricht geändert wurde")
        oItem.Save
        If Catch(DoMessage:=False) Then
            LogicTrace = LogicTrace _
                & "trying save and close modified item " _
                & oID & vbCrLf
        End If
    End If
    
    Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    oItem.Close olDiscard
    Set oItem = Nothing
    aBugTxt = "Get Item from EntryID=" & oID
    Call Try
    Set ReGet = aNameSpace.GetItemFromID(oID)
    If Catch Then
badItem:
        LogicTrace = LogicTrace & "Could not get Item from EntryID=" & oID & vbCrLf
    Else
        Set aItmDsc.idObjItem = ReGet       ' changed by ReGet
        nID = ReGet.EntryID
        aBugVer = ReGet.Class = oldClass
        If DoVerify(aBugVer, "design check " _
            & "ReGet.Class = oldClass Class change on ReGet ???") Then
            Call GetITMClsModel(ReGet, aPindex)
            Call aItmDsc.UpdItmClsDetails(ReGet)
        End If
        If LenB(aNewCat) > 0 _
        And ReGet.Categories <> aNewCat Then
            LogicTrace = LogicTrace & vbCrLf _
                & " reget did set wanted Categories " & Quote(aNewCat) _
                & b & oID & vbCrLf
            If Retry < 1 Then
                Debug.Print "Retrying the save operation after setting new Categories"
                Retry = Retry + 1
                ReGet.Categories = aNewCat
                Set oItem = ReGet
                GoTo tryAgain
            End If
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.ReGet

'---------------------------------------------------------------------------------------
' Method : Function CreateRawItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CreateRawItem(DestFolder As Folder, oItem As Object) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.CreateRawItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim rawItem As Object
Dim retrycount As Long
    If DestFolder Is Nothing Then
        DoVerify False
    End If
    ' create work item (uses default folder)
    Set rawItem = olApp.CreateItem(aObjDsc.objItemType)
    If DebugMode Then
        If rawItem.Class <> oItem.Class Then
            Debug.Print "Class Change: ", rawItem.Class, , aObjDsc.objTypeName, oItem.Class
            Debug.Print "TypeName Change: ", TypeName(rawItem), aObjDsc.objTypeName
            DoVerify False
        End If
    End If
    rawItem.Subject = "Delete this Raw Item"
    rawItem.Categories = "RAW ITEM"
    ' use Close because Save would cause an open inspector defeating Move below
    aBugTxt = "save raw item"
    Call Try(0)
    rawItem.Close olSave
    Catch
    
    ' default folder may not be our target: move it
    If rawItem.Parent.FolderPath <> DestFolder.FolderPath Then
retrythis:
        aBugTxt = "move raw item, retry #" & retrycount
        Call Try
        Set CreateRawItem = rawItem.Move(DestFolder)
        If Catch Then
            aBugTxt = "delete raw item, retry #" & retrycount
            Call Try
            rawItem.Delete
            Catch
        End If
        If CreateRawItem.Parent.FolderPath <> DestFolder.FolderPath Then
            If retrycount > 0 Then DoVerify False
            If DebugMode Then DoVerify False
            Set rawItem = CreateRawItem
            retrycount = retrycount + 1
            GoTo retrythis
        End If
    Else
        Set CreateRawItem = rawItem
    End If
    DoVerify CreateRawItem.Parent.FolderPath = DestFolder.FolderPath
    If DebugMode Then
        Debug.Print "RAW item (no content) moved to " _
            & Quote(DestFolder.FolderPath)
    End If
    Set rawItem = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.CreateRawItem

'---------------------------------------------------------------------------------------
' Method : CopyToWithRDO
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Copies Item to the DestFolder returning copied Item
'---------------------------------------------------------------------------------------
Function CopyToWithRDO(oItem As Object, DestFolder As Folder, NewObjDsc As cObjDsc) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.CopyToWithRDO"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim sItem As Object
Dim aTypeName As String

Dim rItem As Object ' Redemption.Safe-X-Item
Dim newItem As Object
Dim oID As String
Dim nID As String
Dim rDest As RDOFolder
Dim aClassKey As String
Dim fullTypeName As String

    Set sItem = Nothing
    Set rItem = Nothing
    Set newItem = Nothing
    Set rDest = Nothing
    If DestFolder Is Nothing Then
        DoVerify False
    End If
    If oItem Is Nothing Then
        Debug.Print "Can not copy a Nothing"
        DoVerify False
        GoTo FuncExit
    End If
    
    oID = oItem.EntryID
    aTypeName = TypeName(oItem)
    If aOD(aPindex).objClsRules Is Nothing Then
        aClassKey = CStr(oItem.Class)
        DoVerify D_TC.Exists(aClassKey), "Unknown Item Class " & oItem.Class & " can not be copied"
        DoVerify oItem Is D_TC.Item(aClassKey).idObjItem, "not the same item"
        Set aID(aPindex) = D_TC.Item(aClassKey)
    End If
    
    If Not aObjDsc.IsSame(aID(aPindex).idObjDsc, showdiffs:=DebugMode) Then
        Call LogEvent("mail-derived item type? " & aObjDsc.objTypeName & b _
            & Quote1(oItem.Subject), eLSome)
    End If

    aTypeName = aObjDsc.objTypeName
    fullTypeName = TypeName(oItem)
    aBugTxt = "Redemption.Safe" & fullTypeName
    Call Try
    Set sItem = CreateObject("Redemption.Safe" & fullTypeName)
    Catch
    aBugTxt = "RDO Session GetFolderFromPath" & Quote(DestFolder.FolderPath, Bracket)
    Call Try
    Set rDest = aRDOSession.GetFolderFromPath(DestFolder.FolderPath)
    Catch
    aBugTxt = "RDO Session GetRDOObjectFromOutlookObject(oItem)"
    Call Try
    Set rItem = aRDOSession.GetRDOObjectFromOutlookObject(oItem)
    Catch
    
    aBugTxt = "RDO item.add with MessageClass=" & oItem.MessageClass
    Call Try
    Set newItem = rDest.Items.Add(oItem.MessageClass)
    aBugTxt = "close any open item displays"
    Call Try
    oItem.Close olDiscard
    Catch
    aBugTxt = "RDO rItem CopyTo newitem"
    Call Try
    Call rItem.CopyTo(newItem)                      ' CopyTo with rdoFolder does not work
    Catch
    
    aBugTxt = "RDO old Item Save"
    Call Try(testAll)
    If rItem.modified Then
        ' *** attempts to get the MAPI Item in the DestFolder (it is there and OK so far)
        rItem.Save
        Catch
    End If
    If newItem.modified Then
        aBugTxt = "RDO new Item Save"
        Call Try(testAll)
        newItem.Save                                    ' rdo!
        Catch
    End If
    
' Debug.Print "EntryID original            : " & oID
' Debug.Print "EntryID after CopyTo(rItem ): " & rItem.EntryID
' Debug.Print "EntryID after CopyTo(newitem): " & newItem.EntryID
' Debug.Print "Folder(rItem ) = " & rItem.Parent.FolderPath, " Subject=" & rItem.Subject
' Debug.Print "Folder(newitem) = " & newItem.Parent.FolderPath, " Subject=" & newItem.Subject
' Debug.Print "rItem.Parent.FolderPath <> rDest.FolderPath          is ", rItem.Parent.FolderPath <> rDest.FolderPath
' Debug.Print "rItem.Parent.FolderPath <> newitem.Parent.FolderPath is ", rItem.Parent.FolderPath <> newItem.Parent.FolderPath
' Debug.Assert False
    
    nID = newItem.EntryID               ' obtaining non-Rdo from Rdo Object
    aBugTxt = "get the newItem from its EntryID (per NameSpace)"
    Call Try
    Set newItem = aNameSpace.GetItemFromID(nID)
    Catch
    
    If Not newItem.Saved Then
        aBugTxt = "save newItem"
        Call Try
        newItem.Save                    ' RDO-Object has no save flag, get from non-RDO version
        Catch
    End If
    
    nID = newItem.EntryID
    
    aBugVer = newItem.Class = oItem.Class
    DoVerify aBugVer, "changing item Class: " & aObjDsc.objTypeName & b _
                    & Quote1(oItem.Subject)
    aBugVer = aItmDsc.idObjItem.Class = oItem.Class
    DoVerify aBugVer, "changing aItmDsc.IdObjItem or its item Class: " _
                    & aItmDsc.idObjItem.Class = oItem.Class _
                    & vbCrLf & String(5, b) & Quote1(oItem.Subject) _
                    & vbCrLf & String(5, b) & Quote1(aItmDsc.idObjItem.Subject)
    'de-reference the new item, then open it again
    'otherwise Outlook might still have tidbits
    'left from the original message
    Set newItem = ReGet(newItem, nID)
    If newItem Is Nothing Then
        GoTo FuncExit
    End If
    Set CopyToWithRDO = newItem
    If Not NewObjDsc Is aItmDsc.idObjDsc Then
        Set NewObjDsc = aItmDsc.idObjDsc
    End If
    aItmDsc.idEntryId = nID
    Call aItmDsc.UpdItmClsDetails(newItem)
    aBugVer = aItmDsc.idObjItem Is newItem
    DoVerify aBugVer, "aItmDsc.idObjItem Is newItem ???"
    
FuncExit:
    Set newItem = Nothing   ' same on exit!
    Set sItem = Nothing
    Set rItem = Nothing
    Set rDest = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.CopyToWithRDO

'---------------------------------------------------------------------------------------
' Method : Function CopyToWithSafeItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CopyToWithSafeItem(oItem As Object, DestFolder As Folder, myItemClassDsc As cObjDsc) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.CopyToWithSafeItem"
    DoVerify False, "*** CopyToWithSafeItem: function is not used ???"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    ' uses objItemType and aObjDsc.objTypeName (Global)

Dim sItem As Object                                     ' Redemption.Safe-X-Item
Dim newItem As Object
Dim aTypeName As String

    If DestFolder Is Nothing Then
        DoVerify False
    End If
    If oItem Is Nothing Then
        LogicTrace = LogicTrace & "Can not copy a Nothing" & vbCrLf
        DoVerify False
        GoTo cleanup
    End If
    aTypeName = TypeName(oItem)
    DoVerify aObjDsc.IsSame(aID(aPindex), showdiffs:=DebugMode), _
            "** Safe Item has differrent object description"
    
    If DebugMode Then
        Debug.Print " will create " & "Redemption.Safe" & aTypeName
    End If
    Set sItem = CreateObject("Redemption.Safe" & aTypeName)
    If DebugMode Then
        Debug.Print " Redemption.Safe" & aTypeName _
                & " has ItemClass=? (no item)" & sItem.Item Is Nothing
    End If
    sItem.Item = newItem
    If DebugMode Then
        Debug.Print " Redemption.Safe" & aTypeName _
                & " now has ItemClass=sItem.Item.class ", _
                " TypeName = " & TypeName(sItem.Item)
    End If
    ' create work item (uses default folder? -- causes moveto)
    Set newItem = CreateRawItem(DestFolder, oItem)
    If DebugMode Then
        Debug.Print " created outlook rawitem for " & aTypeName _
                & " ItemClass=Item.class ", _
                " TypeName = " & TypeName(newItem)
    End If
    ' copy the item we want to copy into the new item
    sItem.Item = newItem
    aBugTxt = "save item using Redemption"
    Call Try
    Call sItem.CopyTo(newItem)
    If Catch Then
        LogicTrace = LogicTrace _
            & "Redemption SafeItem.CopyTo failed" & vbCrLf
        GoTo cleanup
    End If
    
    If sItem.Item.Subject <> oItem.Subject Then
        DoVerify False
    End If
    Set newItem = sItem.Item
    newItem.Save
    If Not newItem.Saved Then
        LogicTrace = LogicTrace _
            & "Could not save item after copy with Redemption" _
            & vbCrLf
        GoTo cleanup
    End If
    Call aItmDsc.UpdItmClsDetails(newItem)
    
    'de-reference the new item, then open it again
    'otherwise Outlook might still have tidbits
    'left from the original message
    Set newItem = ReGet(newItem, newItem.EntryID)
    If newItem Is Nothing Then
        GoTo cleanup
    End If
    Set CopyToWithSafeItem = newItem
cleanup:
    Set newItem = Nothing
    Set sItem = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.CopyToWithSafeItem

'---------------------------------------------------------------------------------------
' Method : Function CopyToWithRedemption
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CopyToWithRedemption(oItem As Object, toThisFolder As Folder, trySave As Boolean, NewObjDsc As cObjDsc) As Object
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.CopyToWithRedemption"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim copiedItem As Object
Dim MovedItemO As Object
Dim ErrorInFunction As Boolean
Dim oClass As OlObjectClass
Dim aClassKey As String
Dim CopyIsSameAsOriginal As Boolean

    If toThisFolder Is Nothing Then
        DoVerify False
    End If
    If LogicTrace = "*" Then
        LogicTrace = vbNullString
    End If
    If oItem Is Nothing Then
        Debug.Print "can't copy nothing-Object"
        DoVerify False
    End If
    oClass = oItem.Class
    aClassKey = CStr(oClass)
    Set aItmDsc = GetITMClsModel(oItem, aPindex)
    Call aItmDsc.UpdItmClsDetails(oItem)            ' setting the global objItemType/Name
    DoVerify Not aObjDsc Is Nothing, "Unknown object type " & TypeName(oItem)
    
    If trySave Then
        If TrySaveItem(oItem) Then                  ' save original, try repair if save failed
            If oItem Is Nothing Then
                Debug.Print "TrySave ruined original Item"
                DoVerify False
            End If
ErrorInFunction = True
            LogicTrace = "Trying to repair failed Save of original item" & vbCrLf
            Set copiedItem = CopyToWithRDO(oItem, toThisFolder, NewObjDsc)
'                            =============
            If copiedItem Is Nothing Then
                LogicTrace = LogicTrace & "Repairing Save of original item failed" & vbCrLf
ErrorInFunction = True
            Else
ErrorInFunction = False
                aBugTxt = "delete original item " & Quote(oItem.Parent.FolderPath)
                Call Try
                oItem.Delete
                If CatchNC Then
                    DoVerify False, " what is the err.number???"
                    LogicTrace = LogicTrace & "Could not delete original item in " _
                                & Quote(oItem.Parent.FolderPath) & vbCrLf
                    If Catch Then
                        LogicTrace = LogicTrace & "Original item gone, " _
                                            & E_AppErr.Description & vbCrLf
ErrorInFunction = True
                        GoTo HardProblem
                    End If
                Else
                    LogicTrace = LogicTrace & "Deleted original item in " _
                                            & Quote(oItem.Parent.FolderPath) & vbCrLf
                End If
                
                Set oItem = copiedItem
            End If
            If oItem.Saved Then
                LogicTrace = LogicTrace & "item Saved by repair!" & vbCrLf
ErrorInFunction = False
            Else
                LogicTrace = LogicTrace & "item still not Saved" & vbCrLf
ErrorInFunction = True
                DoVerify False
            End If
        End If  ' ErrTrySave
    Else
        LogicTrace = LogicTrace & "item not saved before copy!" & vbCrLf
    End If
    
   ' Copy original to copiedItem
    Call N_ErrClear
    
    Set copiedItem = CopyToWithRDO(oItem, toThisFolder, NewObjDsc)
'                    =============
    If copiedItem Is Nothing Then
        LogicTrace = LogicTrace & "CopyToWithRDO failed" & vbCrLf
ErrorInFunction = True
        GoTo FuncExit
    Else
ErrorInFunction = False
    End If
    
    ' check if the items are unchanged
    If copiedItem.EntryID <> oItem.EntryID Then
        If copiedItem.Parent.FolderPath = toThisFolder.FolderPath Then
            If aObjDsc.objHasReceivedTime Then
                If oItem.SentOn = copiedItem.SentOn Then
                    CopyIsSameAsOriginal = True
                Else
                    LogicTrace = LogicTrace & "copied item has a modified SentOn-Time" & vbCrLf
                End If
                If oItem.ReceivedTime = copiedItem.ReceivedTime Then
                    CopyIsSameAsOriginal = True
                Else
                    LogicTrace = LogicTrace & "copied item has a different Received-Time" & vbCrLf
                    CopyIsSameAsOriginal = False
                End If
            End If
            If CopyIsSameAsOriginal Then
                LogicTrace = LogicTrace & "SafeItem has been copied to Folder " _
                                        & toThisFolder.FolderPath
                Set MovedItemO = copiedItem
                GoTo NoNeedToMove
            End If
        Else
            GoTo mustMove
        End If
    End If
    
    If CopyIsSameAsOriginal Then
        LogicTrace = LogicTrace & "SafeItem already is in Folder " _
                                & Quote(toThisFolder.FolderPath) & vbCrLf
        Set MovedItemO = copiedItem
        GoTo NoNeedToMove
    Else
        If copiedItem.Parent.FolderPath <> toThisFolder.FolderPath Then
mustMove:
            Set MovedItemO = copiedItem.Move(toThisFolder)
        Else
            LogicTrace = LogicTrace & "SafeItem did not need to be moved to Folder " _
                                    & Quote(toThisFolder.FolderPath) & vbCrLf
            Set MovedItemO = copiedItem
        End If
        
        If Catch(AddMsg:="Move to new Folder failed") Then
            LogicTrace = LogicTrace & "Move to new Folder " _
                                    & Quote(toThisFolder.FolderPath) & " failed" & vbCrLf
ErrorInFunction = True
        Else
            Set copiedItem = Nothing
        End If
    End If
    
    ' must save to have EntryID
    If TrySaveItem(MovedItemO) Then
        Debug.Print "Moved Item could not be saved, new EntryID uncertain"
ErrorInFunction = True
    End If
    
NoNeedToMove:
    If DebugMode Then
        Call checkDates(oItem, MovedItemO)
    End If
    If DebugLogging Then
        Call ShowIdentifers(oItem)
        Debug.Print "Target:"
        Call ShowIdentifers(MovedItemO)
    End If
If ErrorInFunction Then
        LogicTrace = LogicTrace & "The operation has failed after all" & vbCrLf
    Else
        If CopyIsSameAsOriginal Then
            Call LogEvent("    : CopyToWithRedemption found " & LogicTrace, eLall)
        Else
            Call LogEvent("    > CopyToWithRedemption successfull into " _
                          & MovedItemO.Parent.FolderPath, eLall)
        End If
        Set CopyToWithRedemption = MovedItemO
    End If
HardProblem:
If DebugMode Or ErrorInFunction Then
        If LenB(LogicTrace) > 0 Then
            Debug.Print LogicTrace
        End If
        DoVerify Not ErrorInFunction
    End If
    
    Call N_ErrClear

FuncExit:
    Set copiedItem = Nothing
    Set MovedItemO = Nothing
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.CopyToWithRedemption

'---------------------------------------------------------------------------------------
' Method : Function TrySaveItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TrySaveItem(oItem As Object) As Boolean    ' true when failed
Dim zErr As cErr
Const zKey As String = "ItemOpsOL.TrySaveItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim RetryWaits As Long
Const maxRetries As Long = 4
Dim trueStart As Variant
Dim TotalTime As Variant

    ' ErrTrySaveItem = False: assume all works fine
    trueStart = 0
    TotalTime = 0
    If oItem.Saved Then
        TrySaveItem = False
        If DebugMode Then
            Debug.Print "Saved already prior to TrySaveItem " & Quote(oItem.Parent.FolderPath)
        End If
        GoTo ProcReturn
    End If
    
    aBugTxt = "save modified item in " & Quote(oItem.Parent.FolderPath)
    Call Try("*da die Nachricht geändert wurde.")
    oItem.Save
    If Not Catch Then
        GoTo FuncExit
    End If
    Set oItem = ReGet(oItem)
    If oItem.Saved Then
        GoTo FuncExit
    Else
        aBugTxt = "save after ReGet in " & Quote(oItem.Parent.FolderPath)
        Call Try
        oItem.Save
        If Not Catch Then
            TrySaveItem = False
            Debug.Print "Saved without problems "
            GoTo FuncExit
        End If
    End If

Retry:
    If RetryWaits < maxRetries Then                         ' try to force with waits/repeats
        Wait 2 ^ RetryWaits, trueStart:=trueStart, _
                    TotalTime:=TotalTime, _
                    Retries:=RetryWaits
        If oItem.Saved Then
            Call LogEvent("     * Trysave may not have saved all data, " _
                & "     * but Item was saved after " _
                & RetryWaits _
                & " Attempts (" & TotalTime & ") sec ", eLall)
            TrySaveItem = False
        Else
            oItem.Save
            If Not oItem.Saved Then
                GoTo Retry
            End If
            TrySaveItem = True
        End If
    End If
    If DebugMode Or RetryWaits >= maxRetries Then
        Debug.Print "retried " & RetryWaits & " times , wait time=" & CInt(TotalTime)
        If TrySaveItem Then
            Call LogEvent("TrySave failed")
        End If
        If DebugMode Then
            DoVerify False
        End If
    End If

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' ItemOpsOL.TrySaveItem

'---------------------------------------------------------------------------------------
' Method : Sub checkDates
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub checkDates(item1 As Object, item2 As Object, _
                Optional mailTypeCheck As Boolean = True)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "ItemOpsOL.checkDates"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

Dim WhatEver As String

    On Error GoTo noda
  
    If mailTypeCheck Then
        If IsMailLike(item1) Then
            aBugVer = item1.Subject = item2.Subject
            If DoVerify(aBugVer, "don't compare items with distinct subject:") Then
                Debug.Print "1: " & Quote(item1.Subject)
                Debug.Print "2: " & Quote(item2.Subject)
                GoTo ProcReturn
            End If
            If IsMailLike(item2) Then
                If ShowTimes(item1, item2, "EntryID") Then
                    Debug.Print "===> time values must be identical"
                Else
                    Call ShowTimes(item1, item2, "ReceivedTime")
                    Call ShowTimes(item1, item2, "SentOn")
                    Call ShowTimes(item1, item2, "CreationTime")
                End If
            Else
                WhatEver = TypeName(item2)
                GoTo noda
            End If
        Else
            WhatEver = TypeName(item1)
            GoTo noda
        End If
    Else
noda:
        Debug.Print "can't determine Date/Time values for type " _
            & WhatEver & " (non-mail type item)"
        DoVerify False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub ' ItemOpsOL.checkDates

'---------------------------------------------------------------------------------------
' Method : Function ShowTimes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ShowTimes(item1, item2, WhatEver) As Boolean
Const zKey As String = "ItemOpsOL.ShowTimes"
    Call DoCall(zKey, tFunction, eQzMode)

Dim x As Date
Dim y As Date

    Select Case WhatEver
        Case "EntryID"
            ShowTimes = item1.EntryID = item2.EntryID
            Debug.Print "Entry IDs match: " & CStr(ShowTimes) & " Subject: " & Quote(item1.Subject)
            GoTo ProcRet
        Case "ReceivedTime"
            x = item1.ReceivedTime
            y = item2.ReceivedTime
        Case "SentOn"
            x = item1.SentOn
            y = item2.SentOn
        Case "CreationTime"
            x = item1.CreationTime
            y = item2.CreationTime
        Case Else
            DoVerify False, " not implemented"
    End Select
    Debug.Print WhatEver & "1 : " & x & " in " & Quote(item1.Parent.FullFolderPath)
    Debug.Print WhatEver & "2 : " & y & " in " & Quote(item2.Parent.FullFolderPath)
    Debug.Print WhatEver & "s match: " & CStr(x = y)

ProcRet:
End Function ' ItemOpsOL.ShowTimes

'---------------------------------------------------------------------------------------
' Method : Sub ShowIdentifers
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowIdentifers(Item)

Const zKey As String = "ItemOpsOL.ShowIdentifers"
    Call DoCall(zKey, tSub, eQzMode)
    
    Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    Debug.Print "Saved=" & Item.Saved, Item.Subject, Item.CreationTime
    Debug.Print Item.Parent.FolderPath, Item.EntryID
    Catch

zExit:
    Call DoExit(zKey)

End Sub ' ItemOpsOL.ShowIdentifers

'---------------------------------------------------------------------------------------
' Method : Function ItemValid
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Determine if Item is (still) valid
'---------------------------------------------------------------------------------------
Function ItemValid(testItem As Object) As Boolean
    '''' Proc Must ONLY CALL Z_Type PROCS                        ' May be Silent
Const zKey As String = "ItemOpsOL.ItemValid"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction, ExplainS:="ItemOpsOL")
    
Dim rClass As Long

    If testItem Is Nothing Then
        GoTo ProcReturn
    End If
    If isEmpty(testItem) Then
        GoTo ProcReturn
    End If
    aBugTxt = "Item Class not available"
    Call Try
    rClass = testItem.Class
    Catch
    If rClass = 0 Then
        GoTo ProcReturn
    End If
    If Not D_TC.Exists(CStr(testItem.Class)) Then
        Call LogEvent("Item Class has not been defined", eLall)
        GoTo ProcReturn
    End If
    aBugTxt = "Get EntryID for Item"
    Call Try
    ItemValid = testItem.EntryID <> vbNullString                      ' throws error &H8004010A when gone
    If Catch Then
        ItemValid = False
    Else
        ItemValid = True
    End If
ProcReturn:
    Call ProcExit(zErr)

ProcRet:
End Function ' ItemOpsOL.ItemValid

'---------------------------------------------------------------------------------------
' Method : Sub ContactFixer
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ContactFixer()
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "ItemOpsOL.ContactFixer"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="ContactFixer")
    
    Call ContactFixItem(ActiveExplorer.Selection.Item(1))

    Call ProcExit(zErr)

End Sub ' ItemOpsOL.ContactFixer

'---------------------------------------------------------------------------------------
' Method : Sub ContactFixItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ContactFixItem(oneItem As Object)
Const zKey As String = "ItemOpsOL.ContactFixItem"
Static zErr As New cErr

Dim oneContact As ContactItem
Dim bestSaveAs As String
    
    If oneItem.Class <> olContact Then
        If DebugLogging Then
            Call LogEvent("item is not a Contact, skipped " & oneItem.Subject, eLall)
        End If
        GoTo skipExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="ContactFix")
        
    Set oneContact = oneItem            ' or aRDOSession.GetRDOObjectFromOutlookObject(oneItem)
    With oneContact
        If LenB(.Email1Address) > 0 Then
            bestSaveAs = .Email1Address
        ElseIf LenB(.Email2Address) > 0 Then
            bestSaveAs = .Email2Address
        ElseIf LenB(.Email3Address) > 0 Then
            bestSaveAs = .Email3Address
        ElseIf InStr(.Body, "@") > 0 Then
            bestSaveAs = GetWordContaining(oneItem.Body, "@")
        End If
        If bestSaveAs <> .Email1Address Then    ' modify email addresses to avoid redundancies
            .Email1Address = bestSaveAs
        End If
        If bestSaveAs = .Email2Address Then
            .Email2Address = vbNullString
        End If
        If bestSaveAs = .Email3Address Then
            .Email3Address = vbNullString
        ElseIf .Email2Address = .Email1Address Then
            .Email2Address = vbNullString
        ElseIf .Email3Address = .Email2Address Then
            .Email3Address = vbNullString
        End If
        If LenB(bestSaveAs) = 0 Then
            If LenB(.Home2TelephoneNumber) > 0 Then
                bestSaveAs = NormalizeTelefonNumber(.Home2TelephoneNumber, Reassign:=True)
                If bestSaveAs <> .Home2TelephoneNumber Then
                    .Home2TelephoneNumber = bestSaveAs
                End If
            ElseIf LenB(.HomeTelephoneNumber) > 0 Then
                bestSaveAs = NormalizeTelefonNumber(.HomeTelephoneNumber, Reassign:=True)
             ElseIf LenB(.MobileTelephoneNumber) > 0 Then
                bestSaveAs = NormalizeTelefonNumber(.MobileTelephoneNumber, Reassign:=True)
            ElseIf LenB(.BusinessTelephoneNumber) > 0 Then
                bestSaveAs = NormalizeTelefonNumber(.BusinessTelephoneNumber, Reassign:=True)
            ElseIf LenB(.Business2TelephoneNumber) > 0 Then
                bestSaveAs = NormalizeTelefonNumber(.Business2TelephoneNumber, Reassign:=True)
            End If
        End If
            
        If LenB(oneContact.CompanyName) > 0 Then
            If LenB(.CompanyAndFullName) > 0 Then
                .FileAs = .CompanyAndFullName
            Else
                .FileAs = .CompanyName
            End If
            bestSaveAs = .FileAs
        ElseIf LenB(.FullName) > 0 Then
            .FileAs = .FullName
            bestSaveAs = .FileAs
        ElseIf LenB(.LastNameAndFirstName) > 0 Then
            .FileAs = .LastNameAndFirstName
            bestSaveAs = .FileAs
        ElseIf LenB(bestSaveAs) > 0 Then
            bestSaveAs = Trim(bestSaveAs)
            .FileAs = bestSaveAs
        End If
    
        If Not .Saved Then
            .Save
            LF_ItmChgCount = LF_ItmChgCount + 1
            Call LogEvent("Contact " & bestSaveAs & " in " & .Parent.FolderPath & " corrected", eLall)
        End If
    End With ' oneContact
    
    Set oneContact = Nothing
    
ProcReturn:
    Call ProcExit(zErr)

skipExit:
End Sub ' ItemOpsOL.ContactFixItem

'---------------------------------------------------------------------------------------
' Method : Sub ChangeBirthdaySubject
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Ersetze "Geburtstag von" durch "*". Nur in Appointment-Ordnern
'---------------------------------------------------------------------------------------
Sub ChangeBirthdaySubject(Optional fName As String, Optional thisFolder As Folder)
Const zKey As String = "ItemOpsOL.ChangeBirthdaySubject"
Static zErr As New cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSubEP, ExplainS:="ItemOpsOL")

Dim StrBuffer As String
Dim LenBuffer As Long
Dim Counter As Long
Dim aItems As Items
Dim aCalItem As Object
Dim i As Long

    Counter = 0

    If thisFolder Is Nothing And LenB(fName) < 2 Then
        Set thisFolder = aNameSpace.GetDefaultFolder(olFolderCalendar)
    Else
        If LenB(fName) = 0 Or fName = "#" Then
            Set thisFolder = ActiveExplorer.CurrentFolder
        Else
            Set thisFolder = GetFolderByName(fName)
            fName = thisFolder.FolderPath
        End If
    End If
    
    If LenB(fName) > 0 Then
        If thisFolder.FolderPath <> thisFolder.FolderPath Then
            Call LogEvent("specified folder name " & fName & " mismatches Folder Path=" & thisFolder.FolderPath, eLall, withMsgBox:=True)
            GoTo ProcReturn
        End If
    End If
    
    If thisFolder.DefaultItemType <> olAppointmentItem Then
        Call LogEvent("Folder " & thisFolder.FolderPath & " may not contain AppointmentItems", eLall, withMsgBox:=True)
        GoTo ProcReturn
    End If
    
    Set aItems = thisFolder.Items

    For i = thisFolder.Items.Count To 1 Step -1
        Set aCalItem = thisFolder.Items(i)
        StrBuffer = aCalItem.Subject
        If InStr(StrBuffer, "Geburtstag von ") Then
            ' aCalItem.Display
            LenBuffer = Len(StrBuffer)
            StrBuffer = Right(StrBuffer, (LenBuffer - Len("Geburtstag von ")))
            StrBuffer = "*" + StrBuffer
            aCalItem.Subject = StrBuffer
            aCalItem.Save
            'aCalItem.Close 0
            Counter = Counter + 1
        End If
    Next i
    
    MsgBox "Fertig!" & vbCrLf & Counter & " Geburtstagseinträge geändert.", vbInformation, "Geburtstage angepasst "

ProcReturn:
    Call ProcExit(zErr)
End Sub ' ItemOpsOL.ChangeBirthdaySubject



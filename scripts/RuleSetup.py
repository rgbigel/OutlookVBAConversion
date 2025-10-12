# Converted from RuleSetup.py

# Attribute VB_Name = "RuleSetup"
# Option Explicit

# Public CategoryString As String
# Public RuleTable As Collection
# Public staticRuleTable As Boolean                ' if true, uses stored tabs,
# ' else get them from Excel
# Public UseExcelRuleTable As Boolean              ' Dynamischer Excel-modus wenn True

# Public Const CategoryKeepList As String = "Unbekannt; "
# Public CategoryDroplist As String
# Public RulesExplained As String

# Dim modCell As Long
# Dim MatchMode As String

# '---------------------------------------------------------------------------------------
# ' Method : Function BestRule
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bestrule():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.BestRule"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if withIrule And Not iRules Is Nothing Then:
    # Set BestRule = iRules
    elif Not sRules Is Nothing Then:
    # Set BestRule = sRules
    # Set iRules = Nothing
    else:
    # Set BestRule = dftRule
    # Set aID(1).idAttrDict = Nothing
    # Set aID(2).idAttrDict = Nothing
    # Set sRules = Nothing
    # Set aTD = Nothing
    # Set sDictionary = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CreateIRule
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createirule():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.CreateIRule"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aDictItem As cAttrDsc

    if aTD Is Nothing Then:
    # GoTo make_new
    else:
    # newPropKey = aTD.adKey
    if Not aTD.adRules Is Nothing Then:
    if aTD.adKey = newPropKey Then:
    if aTD.adRules.RuleInstanceValid Then:
    # Call Get_iRules(aTD)
    # GoTo FuncExit                ' it is there already
    else:
    # GoTo use_old
    else:
    # GoTo make_new
    elif aTD.adKey <> newPropKey Then      ' wrong trail here, define again:
    # DoVerify False, "does this make sense in this version ***???"
    # Set aTD = Nothing
    # Call GetAttrKey(newPropKey, noget:=False)
    if aTD Is Nothing Then:
    # GoTo make_new
    elif aTD.adNr = 0 Then:
    # AttributeIndex = -1                  ' no aID(1).idAttrDict yet
    else:
    # aTD.adNr = AttributeIndex
    # GoTo use_old
    if aTD.adisUserAttr <> isUserProperty Then   ' Misspelled Ucase/Lcase (!!!):
    # DoVerify False
    # GoTo make_new
    # GoTo use_old

    # make_new:
    # ' reached for some specials, like Seperator lines
    # aCloneMode = FullCopy
    # Set aTD = New cAttrDsc

    if aID(aPindex).idAttrDict Is Nothing Then   ' somone kills it???*** (Termination of temporary instance after .addItem).idAttrDict:
    # DoVerify False, " CRAP:  but there is an easy fix"
    # Set aID(aPindex).idAttrDict = aID(aPindex).idAttrDict
    if aID(aPindex).idAttrDict Is Nothing Then DoVerify False, "double check, Debug.Assert False if not fixed":
    if aTD Is Nothing Then:
    # With aID(aPindex).idAttrDict
    if .Exists(newPropKey) Then          ' fix aTd:
    # Set aDictItem = .Item(newPropKey)
    # Set aTD = aDictItem
    # DoVerify PropertyNameX = aTD.adName
    # GoTo use_old
    else:
    # DoVerify False
    # End With                                 ' aID(aPindex).idAttrDict

    # use_old:
    # Set iRules = sRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=False)
    # iRules.ARName = newPropKey                   ' !!! not PropertyNameX!!!
    # Set aTD.adRules = iRules
    # Call SplitDescriptor(aTD)                    ' determine Rules for this Attribute
    # aTD.adRuleIsModified = False                 ' straight from sRules, but now RuleIsSpecific and RuleInstanceValid

    # FuncExit:
    # Set aDictItem = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub RulesToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def rulestoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.adRulesToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if withAttributes Then:
    if xlApp Is Nothing Then                 ' no: put it there:
    # Call XlgetApp
    # Call AttrDscs2Excel
    if xlApp Is Nothing Then:
    # GoTo FuncExit
    if SelectOnlyOne Then:
    if Not aID(2) Is Nothing Then:
    if Not aID(2).idAttrDict Is Nothing Then:
    # DoVerify False, "??? why two operands for this?"
    if xDeferExcel Then:
    # displayInExcel = True
    if xlApp Is Nothing Then             ' no: put it there:
    # Call XlgetApp
    # Set O = xlWBInit(xlA, TemplateFile, _
    # cOE_SheetName, sHdl, showWorkbook:=DebugMode)
    elif O Is Nothing Then:
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl, showWorkbook:=DebugMode)
    # aOD(px).objDumpMade = 0
    # xlApp.ScreenUpdating = False
    # Call StckedAttrs2Xcel(O)
    # xlApp.ScreenUpdating = True
    # ' are we done with all existing properties?
    if aID(px).idAttrDict.Count >= TotalPropertyCount Then:
    # AllPropsDecoded = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Get_iRules
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def get_irules():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.Get_iRules"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if iRules Is Nothing Then:
    if xTD.adRules.RuleInstanceValid Then:
    # Set iRules = xTD.adRules
    # GoTo couldMatch
    # Set iRules = xTD.adRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=True)
    else:
    # couldMatch:
    if iRules.ARName <> xTD.adRules.ARName Then ' wrong one::
    if xTD.adRules.RuleType = "InstanceRule" Then:
    # Set iRules = xTD.adRules
    else:
    # Set iRules = xTD.adRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=True)
    # iRules.RuleIsSpecific = True
    # iRules.RuleInstanceValid = True
    # Call GetRuleBits(xTD)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : GetRuleBits
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: set the iRule Bits including AllPublic.iRuleBits
# '---------------------------------------------------------------------------------------
def getrulebits():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "RuleSetup.Get_iRules"
    # #If MoreDiagnostics Then
    # Call DoCall(zKey, "Sub", eQzMode)
    # #End If

    if iRules.RuleInstanceValid Then:
    if LenB(xTD.adRuleBits) = 0 Or Left(xTD.adRuleBits, 1) = "(" Then:
    # iRuleBits = xTD.adName & _
    # " E:" & Left(iRules.RuleIsSpecific, 1) & _
    # " M:" & Left(iRules.clsObligMatches.RuleMatches, 1) & _
    # " D:" & Left(iRules.clsNeverCompare.RuleMatches, 1) & _
    # " N:" & Left(iRules.clsNotDecodable.RuleMatches, 1) & _
    # " S:" & Left(iRules.clsSimilarities.RuleMatches, 1)
    # xTD.adRuleBits = iRuleBits
    else:
    # iRuleBits = xTD.adRuleBits
    else:
    # iRuleBits = "(not val.)"
    # xTD.adRuleBits = iRuleBits

    # zExit:
    # Call DoExit(zKey)

# '---------------------------------------------------------------------------------------
# ' Method : Function IsAMandatory
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isamandatory():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.IsAMandatory"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # ' compare full word Match, wildcards or ":" will not work
    # i = InStr(b & Trim(sRules.clsObligMatches.aRuleString) & b _
    # & ExtendedAttributeList & b, _
    # b & PropertyNameX & b)
    # IsAMandatory = (i > 0)
    if DebugLogging Then:
    print(Debug.Print "MandatoryAttribute", PropertyNameX, IsAMandatory)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SplitDescriptor
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub SplitDescriptor(xTD As cAttrDsc)
# Dim zErr As cErr
# Const zKey As String = "RuleSetup.SplitDescriptor"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

# Dim LookupName As String

if xTD Is Nothing Then:
# GoTo ProcReturn                          ' no can do at this point
if xTD.adRules Is Nothing Then:
# GoTo ProcReturn                          ' Seperator Lines ========= etc
if xTD.adRules.RuleInstanceValid Then        ' if yes, did split already:
if xTD.adName <> PropertyNameX Then      ' but for wrong one...:
# LookupName = xTD.adName
# GoTo getit                           ' not existing, may need to create it
# Call Get_iRules(xTD)                     ' use results of previous splitDescriptor
# GoTo ProcReturn                          ' already consistent
else:
if xTD.adName <> PropertyNameX Then      ' valid, but the wrong one:
if LenB(PropertyNameX) = 0 Then:
if DebugMode Then:
# DoVerify False, "check if aID(aPindex).odItemDict ok? --> fixed next"
# LookupName = xTD.adName
# getit:
# Set xTD = GetAttrDsc(LookupName, Get_aTD:=False)
if xTD Is Nothing Then:
# GoTo ProcReturn                  ' no can do at this point
# PropertyNameX = LookupName
else:
# Call Get_iRules(xTD)

# With iRules
if (.clsObligMatches.bConsistent _:
# And .clsSimilarities.bConsistent _
# And .clsNotDecodable.bConsistent _
# And .clsNeverCompare.bConsistent) Then
# Call iRules.CheckAllRules(PropertyNameX, "->") ' re-use value settings
else:
# Call iRules.CheckAllRules(PropertyNameX, vbNullString) ' original value setting

# ' check logic and set result of logic as message
if .ARName <> vbNullString And .clsNotDecodable.RuleMatches Then ' unMatchable because undecodable:
# IgString = "  --non-decodable prop."

if .RuleIsSpecific Then                  ' some rule explicitly defined:
if .clsNotDecodable.RuleMatches Then:
# .clsNeverCompare.RuleMatches = True
# .clsSimilarities.RuleMatches = False
# .clsObligMatches.RuleMatches = False
if LenB(Trim(IgString)) = 0 Then:
# IgString = "  (will not decode or compare) "
elif .clsObligMatches.RuleMatches Then:
# .clsNeverCompare.RuleMatches = False
# .clsSimilarities.RuleMatches = False
elif .clsNeverCompare.RuleMatches Then:
# ' although not compared, first decode it
# IgString = "   --dontcompare: "

if .clsSimilarities.RuleMatches Then:
if .clsNeverCompare.RuleMatches Then:
# .clsNeverCompare.RuleMatches = False
# IgString = "  ignoring don't compare, check if similar: "

if LenB(TrueCritList) > 0 Then       ' we do not want any re-ordered attributes:
if .clsObligMatches.RuleMatches _:
# Or .clsSimilarities.RuleMatches Then
# Call AppendTo(SelectedAttributes, PropertyNameX, b)
# End With                                     ' iRules

# FuncExit:
# xTD.adRules.RuleInstanceValid = True

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SplitMandatories
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def splitmandatories():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.SplitMandatories"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcReturn
    # Recursive = True                             ' restored by    Recursive = False ProcReturn:

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' also creates sort criteria and sets TrueCritList - string
    # Dim i As Long
    # Dim j As Long
    # Dim Ci As Long
    # Dim PropName As Variant
    # Dim NoSkipSortAttrs As Boolean

    # 'On Error GoTo 0
    # NoSkipSortAttrs = True

    if sRules Is Nothing Then:
    if MandatoryWorkRule Is Nothing Then:
    # Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, False)
    # Set MandatoryWorkRule = sRules.clsObligMatches
    else:
    # Set sRules = MandatoryWorkRule.PropAllRules
    if Not sRules Is Nothing Then:
    # GoTo GotIt
    else:
    # GotIt:
    if LenB(Trim(sRules.clsObligMatches.aRuleString)) = 0 Then:
    # j = 0
    else:
    # Set MandatoryWorkRule = sRules.clsObligMatches
    if Not sRules.RuleInstanceValid Then:
    # GoTo noRebuildFromRule
    # j = InStr(MatchRequest, sRules.clsObligMatches.CleanMatches(i))
    if j = 0 Then                    ' invalid list, (some) Matchrequests in sRules irregular:
    # Exit For
    if j > 0 Then                            ' no override wanted:
    # MatchRequest = Trim(sRules.clsObligMatches.aRuleString)
    # noRebuildFromRule:
    if ModObligMatches(MatchRequest) Then        ' Will always set sRules:
    if MandatoryWorkRule Is Nothing Then:
    # Set MandatoryWorkRule = sRules.clsObligMatches
    # SortMatches = vbNullString
    # ExtendedAttributeList = vbNullString
    # j = 0
    # i = 0
    # With MandatoryWorkRule
    # .CleanMatchesString = vbNullString
    # Ci = UBound(.CleanMatches)
    if LenB(PropName) > 0 _:
    # And LCase(PropName) <> "or" _
    # And LCase(PropName) <> "and" _
    # And LCase(PropName) <> "not" Then ' remove empty or operators
    if i <> j Then:
    # MandatoryWorkRule.CleanMatches(j) = PropName ' pull down to correct position
    # Ci = Ci - 1              ' cut end of array
    # .CleanMatchesString = .CleanMatchesString & b & PropName
    # j = j + 1                    ' copy into this position when true CritPropName NEXT time
    # ' find out if we want to Debug.Assert False building objSortMatches
    if InStr(.MatchesList(i), "|") > 0 Then:
    # NoSkipSortAttrs = False  ' stays this way for this call
    if NoSkipSortAttrs Then:
    if InStr(ExtendedAttributeList & b, .CleanMatches(i) & b) = 0 Then:
    # SortMatches = SortMatches & "[" & PropName & "] "
    else:
    # ExtendedAttributeList = ExtendedAttributeList & b & PropName
    else:
    # j = j                        ' no step of target array
    # i = i + 1
    # ' correction of element count in TrueImportantProperties
    if j < 1 Then:
    # Erase TrueImportantProperties
    elif Ci < UBound(.CleanMatches) Then:
    # ' a vbNullString Property was removed
    # ReDim Preserve TrueImportantProperties(j - 1)
    # .CleanMatchesString = Trim(.CleanMatchesString)
    # ExtendedAttributeList = Trim(ExtendedAttributeList)
    # TrueCritList = Trim(.aRuleString)
    # MostImportantProperties = .CleanMatches
    # MostImportantAttributes = .CleanMatchesString
    # MainObjectIdentification = .CleanMatches(0)
    # End With                                 ' MandatoryWorkRule
    if MandatoryWorkRule Is Nothing Then:
    # MostImportantProperties = Array(vbNullString)
    # MostImportantAttributes = vbNullString
    else:
    # MostImportantProperties = MandatoryWorkRule.CleanMatches
    # MostImportantAttributes = MandatoryWorkRule.CleanMatchesString
    if Not isEmpty(MandatoryWorkRule.CleanMatches) Then:
    # MainObjectIdentification = MandatoryWorkRule.CleanMatches(0)
    # SortMatches = Trim(SortMatches)
    # Call AppendTo(SortMatches, "[LastModificationTime]", b)
    # sRules.RuleInstanceValid = True

    # FuncExit:
    # Recursive = False
    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub StckedAttrs2Xcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stckedattrs2xcel():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.StckedAttrs2Xcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim px As Long
    # Dim DidIndex As Long
    # Dim i As Long
    # Dim FilterOff As Boolean

    if aOD(0).objDumpMade >= 2 Then:
    # GoTo ProcReturn

    # aOD(0).objDumpMade = 0
    if xlApp.ActiveSheet.Name <> aTab.xlTName Then:
    # Call XlopenObjAttrSheet(xlA)
    # aTab.xlTSheet.Activate

    # noquickclear:
    # 'On Error GoTo 0
    # aTab.xHdl = sHdl                             ' sets first line as headline ( property let )
    if Not (DebugMode Or DebugLogging) Then:
    # xlApp.Visible = False
    # xlApp.Cursor = xlWait
    if SelectOnlyOne Then:
    # FilterOff = True
    if WorkItemMod(1) Then:
    # i = 2
    else:
    # i = 1
    else:
    # i = 2

    if Not aID(px).idAttrDict Is Nothing Then:
    if aID(px).idAttrDict.Count > 0 Then:
    # Call StckedAttrLoop(px)
    # DidIndex = px
    if aID(px) Is Nothing Then:
    else:
    # aOD(px).objDumpMade = px
    # px = DidIndex
    # aOD(0).objDumpMade = px
    if FilterOff Then:
    # aTab.xlTSheet.Range("$A$1").AutoFilter
    else:
    # aTab.xlTSheet.Range("$A$1:$H$1").AutoFilter _
    # Field:=6, Criteria1:="="
    # aTab.xlTSheet.Cells(2, 1).Select

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub StckedAttrLoop
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stckedattrloop():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.StckedAttrLoop"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim N As Long
    # aPindex = px
    if xlApp Is Nothing Then:
    # GoTo ProcReturn
    # N = 1
    # AttributeIndex = i
    # Set aTD = aID(px).idAttrDict.Items(i)
    if OnlyMostImportantProperties Then:
    if InStr(SelectedAttributes & b, aTD.adKey & b) = 0 Then:
    # GoTo nextInLoop
    # N = N + 1
    # Call put2IntoExcel(px, N)
    if N = 1 Or N Mod 10 = 0 Then:
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab & px, _)
    # "inserting attribute # " & i _
    # & " into Sheet " & W.xlTName

    print(Debug.Print Format(Timer, "0#####.00") & vbTab & px, _)
    # "last cAttrDsc = " & N & " into Sheet " & W.xlTName

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ChkCatLogic
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub ChkCatLogic()
# Dim zErr As cErr
# Const zKey As String = "RuleSetup.ChkCatLogic"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

# Dim temp As Boolean
# With CurIterationSwitches
# temp = .ReProcessDontAsk _
# Or (.ReprocessLOGGEDItems And Not eOnlySelectedItems)
if temp <> .ReProcessDontAsk Then:
# .ReProcessDontAsk = temp
# End With                                     ' CuriterationSwitches

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CategorizeItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def categorizeitem():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.CategorizeItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim pos As Long
    # Dim wLen As Long
    # Dim ca As Variant
    # Dim aAttributeValue As String
    # Dim Matched As Boolean
    # Dim Prefix As String
    # Dim categoryKeepV As Variant
    # Dim CategoryCandidates As String
    # Dim WordStart As Boolean
    # Dim WordEnd As Boolean

    # CategoryCandidates = vbNullString
    # aAttributeValue = curItem.Categories
    if Not CurIterationSwitches.ResetCategories Then:
    if isEmpty(categoryKeepV) Then           ' never dropped:
    # categoryKeepV = split(CategoryKeepList, "; ")
    if Not isEmpty(categoryKeepV) Then:
    # Prefix = Trim(categoryKeepV(i))
    if Prefix <> vbNullString And InStr(aAttributeValue, Prefix) > 0 Then:
    # Call AppendTo(CategoryCandidates, Prefix, ";")
    # Prefix = "     "
    # CategorizeItem = vbNullString

    # aAttributeValue = vbNullString
    # Matched = False
    # Set ca = RuleTable.Item(i)
    # WordStart = False
    # WordEnd = False

    # With ca
    if LenB(.typ) = 0 Then:
    # Exit For                         ' invalid, terminate loop
    # '           .typ                       ' col  1
    # '           .checkitem                 '      2
    # '           .category                  '      3
    # '           .final                     '      4
    # '           .exact                     '      5
    # '           .never                     '      6
    # '           .pattern                   '      7

    # wLen = Len(.checkitem)
    # MatchMode = vbNullString                       ' 4 states possible:
    if .final Then:
    # MatchMode = "="
    if .never Then:
    # MatchMode = "-"
    if .Exact Then:
    # MatchMode = "*"
    if Not .pattern Then:
    # MatchMode = "="

    match LCase(.typ):
        case "subject":
    # aAttributeValue = UCase(curItem.Subject)
        case "sender":
    if aObjDsc.objHasSenderName Then:
    # aAttributeValue = UCase(curItem.SenderEmailAddress)
    elif aObjDsc.objHasSentOnBehalfOf Then:
    # aAttributeValue = UCase(curItem.SenderName & b & curItem.SentOnBehalfOfName)
    else:
    # aAttributeValue = "# anonymous " & ca.typ
        case "body":
    if Not aObjDsc.objHasHtmlBodyFlag Then:
    # GoTo noHTML                  ' some Classes do not have that, so body is always text
    if curItem.BodyFormat <> olFormatHTML Then:
    # noHTML:
    # Call ErrReset(0)
    if Len(curItem.Body) > 300000 Then:
    # Debug.Assert False
    # aAttributeValue = UCase(curItem.Body)
    else:
    if Len(curItem.HTMLBody) > 500000 Then:
    # 'Debug.Assert False      #### Switch this on?
    # Call Try                     ' Try anything, autocatch, Err.Clear
    # aAttributeValue = UCase(curItem.HTMLBody)
    if Catch Then:
    # GoTo noHTML
    # Call ErrReset(0)
        case _:
    # DoVerify False, " not implemented as a rule"

    # pos = InStr(aAttributeValue, UCase(.checkitem))

    if pos > 0 Then                      ' if found at all:
    if pos = 1 Or .pattern Then:
    # WordStart = True
    else:
    if InStr(WordSep, Mid(aAttributeValue, pos - 1, 1)) > 0 Then:
    # WordStart = True         ' WordSep preceedes it
    elif Asc(Mid(aAttributeValue, pos - 1, 1)) < Asc(b) Then:
    # WordStart = True         ' match on nonprintable start seperators
    if wLen <= Len(aAttributeValue) Then:
    if InStr(WordSep, Mid(aAttributeValue, pos + wLen, 1)) > 0 Then:
    # WordEnd = True           ' WordSep follows it
    elif Asc(Mid(aAttributeValue, pos + wLen, 1)) < Asc(b) Then:
    # WordEnd = True           ' match on nonprintable end seperators

    if WordStart And WordEnd Then:
    if CurIterationSwitches.CategoryConfirmation Then:
    # RulesExplained = Append(RulesExplained, TypeName(curItem) & _
    # " chosen Category: " & .category & _
    # " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
    # Call AppendTo(CategoryCandidates, .category, "; ")
    else:
    match (MatchMode):
        case "-"                 ' if .never then:
    # RulesExplained = Append(RulesExplained, TypeName(curItem) & _
    # " can not be Category " & .category & _
    # " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
    # Call AppendTo(CategoryDroplist, .category, "; ")
    # CategoryCandidates = StringRemove(CategoryCandidates, _
    # CategoryDroplist, "; ")
        case vbNullString                  ' Category is not uniquely set (may have several):
    # RulesExplained = Append(RulesExplained, TypeName(curItem) & _
    # " fits Category " & .category & _
    # " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
    # Call AppendTo(CategoryCandidates, .category, "; ")
        case "="                 ' Category is final (multiple) + LOGGED always):
    # RulesExplained = Append(RulesExplained, TypeName(curItem) _
    # & " given final categories " & .category _
    # & " due to tag '" & .typ & "' " & .checkitem)
    if LenB(CategoryCandidates) > 0 Then:
    # RulesExplained = RulesExplained & ", previous categories " _
    # & Quote(CategoryCandidates) & " kept" & vbCrLf
    else:
    # RulesExplained = RulesExplained & vbCrLf
    # Call AppendTo(CategoryCandidates, .category, "; ")
    # GoTo FunExit
        case Else                ' if .exact; should be same as "*", final unique cat.:
    # RulesExplained = Append(RulesExplained, TypeName(curItem) _
    # & " given unique categories " & .category _
    # & " due to tag '" & .typ & "' " & .checkitem)
    if LenB(CategoryCandidates) > 0 Then:
    # RulesExplained = RulesExplained & ", previous categories " _
    # & StringRemove(CategoryCandidates, .category, "; ") _
    # & " dropped" & vbCrLf
    else:
    # RulesExplained = RulesExplained & vbCrLf
    # CategoryCandidates = .checkitem
    # GoTo FunExit
    else:
    # RulesExplained = RulesExplained & Prefix & TypeName(curItem) _
    # & " has no word-Match in " & .typ & " for " _
    # & .checkitem & ": " & Quote(Mid(aAttributeValue, pos, wLen + 2)) _
    # & vbCrLf
    else:
    # RulesExplained = RulesExplained & Prefix & TypeName(curItem) _
    # & " has no Match in " & .typ & " for " _
    # & .checkitem & ": " & Quote(Mid(aAttributeValue, 1, wLen + 2) & "... ") _
    # & vbCrLf
    # End With                                 ' ca

    # FunExit:
    if CurIterationSwitches.CategoryConfirmation Then:
    # Call LogEvent("Category Candidates are: " & CategoryCandidates)
    else:
    if LenB(CategoryCandidates) = 0 Then:
    # Call LogEvent(Prefix & TypeName(curItem) & _
    # " not assigned specific Category, no tags Matched", eLmin)
    else:
    # Call LogEvent(Prefix & TypeName(curItem) & _
    # " was assigned to the following categories: " _
    # & CategoryCandidates)
    # CategorizeItem = CategoryCandidates
    if ShowFunctionValues Then:
    # Call LogEvent(Prefix & RulesExplained)

    # FuncExit:
    # RulesExplained = vbNullString
    # Call ErrReset(4)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function DetectCategory
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def detectcategory():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.DetectCategory"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ShortName As String
    # Dim XlwasOpenedHere As Boolean
    # Dim editmode As Boolean
    # Dim NewCategory As String
    # Dim tempCategory As String
    # Dim SourceFolder As Folder
    # Dim SourceFolderPath As String

    # RulesExplained = vbNullString
    if RuleTable Is Nothing Then:
    # Set RuleTable = New Collection
    # Set SourceFolder = curItem.Parent
    # SourceFolderPath = SourceFolder.FullFolderPath

    # ReEstablishCategoryRules:
    if CurIterationSwitches.ResetCategories Then:
    # NewCategory = vbNullString
    else:
    # NewCategory = OldCategory
    if RuleTable.Count = 0 Then:
    if staticRuleTable And Not UseExcelRuleTable Then:
    # Call InitRuleTable
    else:
    # GoTo getOrEdittags
    else:
    if UseExcelRuleTable Then:
    # getOrEdittags:
    if OpenRuleTable(XlwasOpenedHere) Then:
    # Call Xl2RuleTable(editmode)      ' edit or not depends on Z_AppEntry
    # Call GotRuleTabXl(XlwasOpenedHere)
    else:
    # ' can use the existing static rule table

    # OldCategory = curItem.Categories
    # MatchMode = vbNullString

    if aObjDsc.objIsMailLike Then:
    # ShortName = Left(TargetFolder.Name, 4)
    if ShortName = "Junk" _:
    # Or ShortName = "Spam" _
    # Or ShortName = "Uner" _
    # Then
    # NewCategory = "Junk"
    # Call LogEvent("Mail assigned OldCategory " & NewCategory _
    # & " due to Folder name " & TargetFolder.FullFolderPath, eLmin)
    else:
    if CurIterationSwitches.ResetCategories _:
    # Or InStr(OldCategory, LOGGED) = 0 Then
    # tempCategory = CategorizeItem(curItem)
    if CurIterationSwitches.ResetCategories Then:
    # NewCategory = tempCategory
    elif LenB(tempCategory) > 0 Then ' keep the existing value if no Rule:
    # NewCategory = tempCategory
    # Call AppendTo(NewCategory, curItem.Categories, "; ")
    else:
    # NewCategory = curItem.Categories
    # OKOK:
    if InStr(1, NewCategory, Unbekannt, vbTextCompare) > 0 Then:
    # GoTo unk
    elif Not CurIterationSwitches.ResetCategories Then ' reset Cat ==> may correct source folder:
    if InStr(1, SourceFolderPath, Unbekannt, vbTextCompare) > 0 Then:
    # GoTo unk1
    else:
    if InStr(curItem.Parent.FullFolderPath, "SMS") > 0 Then:
    # Call AppendTo(NewCategory, "SMS", "; ")
    if Not aObjDsc.objHasSenderName Then:
    # GoTo unk1
    if IsUnkContact(curItem.SenderEmailAddress) Then:
    # unk1:
    # Call AppendTo(NewCategory, Unbekannt, "; ")
    # unk:
    if FolderUnknown Is Nothing Then:
    # Set TargetFolder = SourceFolder  ' not moving anything
    else:
    # Set TargetFolder = FolderUnknown
    else:
    if InStr(1, SourceFolderPath, "BACKUP", vbTextCompare) > 0 Then:
    if CurIterationSwitches.ResetCategories Then:
    # Set TargetFolder = FolderInbox ' no longer unknown sender
    else:
    # Set TargetFolder = SourceFolder ' remain in Folder (do not move)
    else:
    if TargetFolder Is Nothing Then:
    # Set TargetFolder = FolderInbox
    # NewCategory = Append(NewCategory, LOGGED, "; ", ToFront:=True)
    if CurIterationSwitches.ReProcessDontAsk Then ' user interaction has been turned off:
    # NewCategory = NewCategory
    # GoTo FunExit
    if CurIterationSwitches.CategoryConfirmation Then:
    if LenB(RulesExplained) = 0 Then:
    # RulesExplained = "Keine Regel gefunden => keine besondere Kategorie"
    try:
        # frmStrEdit.Caption = curItem.Subject
        # frmStrEdit.chSaveItemRequested = CurIterationSwitches.SaveItemRequested
        # frmStrEdit.StringModifierCancelLabel.Caption = "alt:"
        # frmStrEdit.StringModifierCancelValue.Text = OldCategory
        # frmStrEdit.StringModifierExpectation = _
        # "Die folgenden Kategorien sind u.W2. fr dieses " _
        # & TypeName(curItem) _
        # & " geeignet. Bitte prfen und ggf. korrigieren."
        # frmStrEdit.StringToConfirm = NewCategory
        # frmStrEdit.Explanations = RulesExplained
        # CurIterationSwitches.CategoryConfirmation = False ' once assumed remember globally
        # frmStrEdit.Show
        # rsp = frmStrEdit.StringModifierRsp
        match rsp:
            case vbOK                                ' user closed this and/or did not answer...:
        if frmStrEdit.CategoryConfirmation And Not frmStrEdit.ReProcessDontAsk Then:
        # RulesExplained = "<=== Categorie-Regeln werden erneut angewendet ===>"
        # GoTo ReEstablishCategoryRules
        # NewCategory = frmStrEdit.StringToConfirm
            case vbYes                               ' want to edit the rules:
        # editmode = True
        # frmMaintenance.someAction = 1
        # GoTo getOrEdittags
            case vbNo:
        # NewCategory = OldCategory            ' not changing anything and no save
            case Else                                ' no, cancel, retry: user closed this and/or did not answer...:
        # Call TerminateRun
        # End
        # FunExit:
        # Set frmStrEdit = Nothing
        if NewCategory <> OldCategory Then:
        # DetectCategory = NewCategory
        # MailModified = True

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# ' Achtung Generierter Code, nach frmMaintenance ersetzen
def initruletable():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.InitRuleTable"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Set RuleTable = New Collection
    # Call SetRuleTabFlds("Subject", "empfie", "Newsflash", True, False, False, True)
    # Call SetRuleTabFlds("Subject", "statusreport", "Log", False, True, False, True)
    # Call SetRuleTabFlds("Body", "status report", "Log", False, True, False, True)
    # Call SetRuleTabFlds("Sender", "Stern", "Stern", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "Save.TV", "SaveTV", True, False, False, False)
    # Call SetRuleTabFlds("Sender", "redaktion@JungeWelt", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "support", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "infomail", "Newsflash", False, True, False, True)
    # Call SetRuleTabFlds("Sender", "jaxx", "Lotto", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "MS", "MS", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "bounce", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "news", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "@facebookmail.com", "Newsflash", True, False, False, True)
    # Call SetRuleTabFlds("Sender", "one4all", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Sender", "wissenschaft-online", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "Sparaktion", "Newsflash", True, False, False, True)
    # Call SetRuleTabFlds("Sender", "eilmeldung", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Sender", "oxygen", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Body", "newsletter regist", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Body", "newsletter", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Sender", "orders", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "payment@paypal", "Rechnung", False, True, False, False)
    # Call SetRuleTabFlds("Body", "MS", "MS", True, False, False, False)
    # Call SetRuleTabFlds("Sender", "Lockergnome", "Lockergome", False, True, False, False)
    # Call SetRuleTabFlds("Body", "Lockergnome", "Lockergome", False, False, False, True)
    # Call SetRuleTabFlds("Body", "Fritz!Box", "FritzBox", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Redmond", "MS", True, False, False, False)
    # Call SetRuleTabFlds("Body", "codeproject", "MS", True, False, False, False)
    # Call SetRuleTabFlds("Body", "SmartTools", "SmartTools", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Order Confirmation", "MailOrder", False, True, False, False)
    # Call SetRuleTabFlds("Body", "Ausgangsbesttigung", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Rcksendung", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Rechnung", "Rechnung", False, False, False, False)
    # Call SetRuleTabFlds("Body", "payment", "Rechnung", True, False, False, True)
    # Call SetRuleTabFlds("Sender", "anja.weber", "newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Body", "eur statt", "newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Bestellung", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Body", "key", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "licen", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Anmeldung", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Zugang", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "lizen", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Auftrag", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Body", "account", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "passwor", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Anmeld", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Abmeld", "Saved Mail", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "pearl", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "digitalo", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "valentins", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "Oxygen3", "Junk", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "kundenservice", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "vmwareteam", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "SmartTools", "SmartTools", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Angebot", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Kauf", "MailOrder", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "supp0rt", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "Bercht", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Sender", "Igel-Soft", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Newsgroup", "Groups", False, False, False, False)
    # Call SetRuleTabFlds("Body", "web.de informiert", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Body", "News", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Body", "event", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Angebote", "Newsflash", False, False, False, True)
    # Call SetRuleTabFlds("Body", "Sparen", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Sparbrief", "Junk", True, False, False, False)
    # Call SetRuleTabFlds("Body", "web.de gmbh", "Newsflash", False, True, False, False)
    # Call SetRuleTabFlds("Body", "n e w s", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Gutschein", "Newsflash", False, False, False, False)
    # Call SetRuleTabFlds("Body", "linke zeitung", "Newsflash", False, True, False, True)
    # Call SetRuleTabFlds("Body", "cxtreme", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Body", "conrad", "Newsflash", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Lockergnome", "Lockergnome", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Windows Fanatics", "Lockergnome", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Stern", "Stern", True, False, False, False)
    # Call SetRuleTabFlds("Body", "Jubilum", "Junk", False, True, False, False)
    # Call SetRuleTabFlds("Body", "MSN Groups", "Groups", False, False, False, False)
    # Call SetRuleTabFlds("Body", "bercht", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Marc", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Saskia", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Body", "Franzi", "Rolf", False, False, False, False)
    # Call SetRuleTabFlds("Body", "regist", "Saved Mail", False, False, False, False)
    # staticRuleTable = True
    # UseExcelRuleTable = False                    ' execute sub InitRuleTable

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function OpenRuleTable
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def openruletable():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.OpenRuleTable"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ntag(1 To 7) As Variant
    # Dim aHdl As String

    # aHdl = "Typ------------ CheckItem----------- Category---------- Final--- Exact--- Never--- Pattern-"
    # modCell = UBound(ntag) + 1                   ' this cell will Show if we changed anything
    # Set E = xlWBInit(xlA, TemplateFile, "RuleCategoriesTable", aHdl, showWorkbook:=DebugMode)
    # XlwasOpenedHere = XlOpenedHere And Not (xUseExcel Or xDeferExcel)
    if DebugMode Then:
    # Call DisplayExcel(E, relevant_only:=False, _
    # EnableEvents:=True, xlY:=xlA)
    # E.xlTSheet.Activate
    # Call GetLine(1, ntag)
    if InStr(aHdl, ntag(1)) <> 1 Then:
    # OpenRuleTable = False                    ' incorrect Headline, do not read this RuleCategoriesTable
    else:
    # OpenRuleTable = True
    if LCase(E.xlTSheet.Cells(1, modCell)) = "modified" Then:
    # UseExcelRuleTable = True                 ' and use them from now on because inline code can not be changed here

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Xl2RuleTable
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xl2ruletable():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.Xl2RuleTable"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim ntag(1 To 7) As Variant
    # Dim i As Long
    # Dim nFile As Long
    # Set RuleTable = New Collection
    # ' = am anfang des CheckItem heit: finale Zuordnung, nicht weitersuchen (ab hier)
    # ' *                                finale Zuordnung wie "=", exakt diese eine Kategorie
    # ' -                                diese Kategorie sicher nicht setzen
    # ' sonstige                         diese Kategorie zuordnen (bis evtl. Widerruf)
    # ' = am Ende: Wildcard-modus AUS (Wichtig!)
    # '                Typ      CheckItem     Category   (ordered, high priority first)
    if LCase(W.xlTSheet.Cells(1, clickColumn)) = "modified" Then:
    # UseExcelRuleTable = True                 ' and use them from now on because inline code can not be changed here
    else:
    # W.xlTSheet.Cells(1, clickColumn).Value = vbNullString
    if editmode Then:
    # nFile = -2
    # Call openFile(nFile, genCodePath, "InitRuleTable", ".bas", "Output")
    # Print #nFile, " & quote( Achtung Generierter Code, nach frmMaintenance ersetzen"
    # Print #nFile, "Sub InitRuleTable()"
    # Print #nFile, "    Set RuleTable = New Collection"
    # i = 2                                        ' skip headline
    # While i > 1
    # Call GetLine(i, ntag)
    if LenB(ntag(1)) = 0 Then:
    # GoTo LoopEnd                         ' end of File
    # Call SetRuleTabFlds(ntag(1), ntag(2), ntag(3), _
    # CBool(ntag(4)), CBool(ntag(5)), _
    # CBool(ntag(6)), CBool(ntag(7)))
    if editmode Then:
    # Print #nFile, "    Call SetRuleTabFlds(" _
    # & Quote(ntag(1)) & ", " _
    # & Quote(ntag(2)) & ", " _
    # & Quote(ntag(3)) & ", " _
    # & DeBoolToEn(ntag(4)) & ", " _
    # & DeBoolToEn(ntag(5)) & ", " _
    # & DeBoolToEn(ntag(6)) & ", " _
    # & DeBoolToEn(ntag(7)) & ")"
    # i = i + 1
    # Wend
    # LoopEnd:
    if editmode Then:
    # Print #nFile, "StaticRuleTable =  True"
    # Print #nFile, "    UseExcelRuleTable = False ' execute sub InitRuleTable"
    print(Debug.Print "En" & "d Su" & "b      ' Xl2RuleTable" & vbCrLf)
    # Print #nFile, " & quote( Ende des Generierten Codes"
    # Close #nFile
    # Call LogEvent("Created new file " & Quote(genCodePath & "\InitRuleTable.bas") _
    # & " for " & RuleTable.Count & " Rules", eLall)
    # UseExcelRuleTable = True                     ' until we execute sub InitRuleTable

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetRuleTabFlds
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose: Trivial, Inline for setting and adding RuleTab via parameter list.
# '---------------------------------------------------------------------------------------
def setruletabflds():

    # Dim xTag As cRuleCat

    # Set xTag = New cRuleCat
    # With xTag
    # .typ = typ
    # .checkitem = checkitem
    # .category = category
    # .final = CBool(final)
    # .Exact = CBool(Exact)
    # .never = CBool(never)
    # .pattern = CBool(pattern)
    # End With                                     ' xTag
    # RuleTable.Add xTag


# '---------------------------------------------------------------------------------------
# ' Method : Sub GotRuleTabXl
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def gotruletabxl():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.GotRuleTabXl"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if RuleTable.Count < 1 Then:
    # DoVerify False, " Excel Rules table is empty???"
    # staticRuleTable = True
    # Call InitRuleTable                       ' use the static ones
    else:
    # Call LogEvent("RuleSetup is using Excel tags", _
    # eLmin)
    if XlwasOpenedHere Then:
    # Call xlEndApp

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub setItmCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setitmcats():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.setItmCats"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim sa As Variant
    # Dim sO As Variant
    # Dim sX As Variant
    # Dim category As String
    # Dim msg As String
    # Dim sT As String

    # oldCatList = curItem.Categories
    # sa = split(addCatList, ";")
    # category = vbNullString

    if MatchMode = "*" _:
    # Or LenB(oldCatList) = 0 _
    # Or CurIterationSwitches.ResetCategories _
    # Then                                      ' override Category by new value(sX) (forget old)
    # msg = "     " & TypeName(curItem) & " categories "
    else:
    # sO = split(oldCatList, ";")
    for sx in so:
    # sX = Trim(sX)
    # sT = sX & "; "
    if InStr(dropCatList, sT) = 0 Then:
    if InStr(category & "; ", sT) = 0 And LenB(sX) > 0 Then:
    # Call AppendTo(category, Trim(sX), "; ")
    # msg = "     " & TypeName(curItem) & vbTab _
    # & "assigned categories: "
    for sx in sa:
    if InStr(category, sX) = 0 And LenB(sX) > 0 Then:
    # Call AppendTo(category, Trim(sX), "; ")

    # category = Trim(LOGGED & "; " _
    # & StringRemove(category, dropCatList, "; "))
    # category = Replace(category, "; ;", ";")
    # category = Replace(category, ";;", ";")
    # category = RCut(category, 1)

    # curItem.UnRead = False                       ' should normally cause sA change
    if oldCatList <> category Then:
    # curItem.Categories = category            ' source-side modification
    # category = "changed from " & Quote(oldCatList) & " to "
    else:
    # category = "not changed from "

    # ' always save original (which was very likely changed)
    # ' because item.saved is not reliable when Categories change (IMAP has no categories)
    # curItem.Save
    if T_DC.DCerrNum = 0 Then:
    # MailModified = False                     ' so far, so good
    # Call LogEvent(msg & category & Quote(curItem.Categories) & " and saved", eLall)
    else:
    # Call LogEvent(msg & category & Quote(curItem.Categories) & " NOT saved", eLall)

    # aNewCat = curItem.Categories

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ModRuleTab
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def modruletab():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.ModRuleTab"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim modified As Boolean
    # Dim AR As Range

    # Call Xl2RuleTable(editmode:=True)
    # modified = x.xlTSheet.Cells(1, changeCounter).Value > 0
    # Call Xl2RuleTable(editmode:=True)
    if modified Then:
    print(Debug.Print " & quote( Regeln wurden gendert, bitte die folgenden " _)
    # & "Zeilen in Rule-Wizard ersetzen"
    print(Debug.Print " & quote( *** Anfang des aus Excel generierten Codes")
    print(Debug.Print " & quote( *** Ende des generierten Codes")
    # x.xlTSheet.Cells(1, changeCounter).Clear
    else:
    # Call LogEvent("Es wurden keine nderungen in Excel durchgefhrt", eLall)
    # UserDecisionEffective = True
    # ' clear changes done in this session
    # x.xlTLastLine = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Row - 1
    # x.xlTLastCol = ActiveSheet.UsedRange.columns.Count + ActiveSheet.UsedRange.Column - 1
    # Set AR = ActiveSheet.Range(Cells(2, clickColumn - 1), Cells(x.xlTLastLine, x.xlTLastCol))
    # Set E = x
    # AR.Clear
    # 'Range(Cells(2, clickColumn - 1), _
    # A.xlTSheet.Cells.SpecialCells(xlCellTypeLastCell)).Clear ??? *** old code to remove

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EditRulesTable
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def editrulestable():
    # Dim zErr As cErr
    # Const zKey As String = "RuleSetup.EditRulesTable"
    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # Dim XlwasOpenedHere As Boolean
    # frmMaintenance.Hide
    if OpenRuleTable(XlwasOpenedHere) Then:
    # DoVerify False
    # Call ExcelEditSession(1)
    # Call GotRuleTabXl(XlwasOpenedHere)
    else:
    # UseExcelRuleTable = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:


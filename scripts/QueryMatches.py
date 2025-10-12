# Converted from QueryMatches.py

# Attribute VB_Name = "QueryMatches"
# Option Explicit

# Private DidFindInits As Boolean ' used to init only once for Entry Points

# '---------------------------------------------------------------------------------------
# ' Method : Sub WasEmailProcessed
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def wasemailprocessed():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "QueryMatches.WasEmailProcessed"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="QueryMatches")

    # Call FindEntryInit("User initiated Event")
    # eOnlySelectedItems = True
    # Call FirstPrepare                           ' Folder by default (no user interaction)
    # Call SelectAndFind

    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SelectAndFind
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def selectandfind():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "QueryMatches.SelectAndFind"
    # Call DoCall(zKey, "Sub", eQzMode)

    # xReportExcel = True  ' maybe other values?
    # Call MatchingItems(MatchMode:=1)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub FindEntryInit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findentryinit():
    # ' and call ReturnEP at Exit there???
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "QueryMatches.FindEntryInit"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:=MyExplanation)

    # IsEntryPoint = True                     ' common for several Entry Points
    # Call SetEventMode(force:=True)

    # xUseExcel = False                       ' Defaults for QueryMatches only
    # xDeferExcel = False
    # xReportExcel = False
    # quickChecksOnly = True
    # SelectOnlyOne = True
    # SelectMulti = False

    # ActionID = 0
    # ActionTitle(0) = "dynamic action SelectAndFind"

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub FirstPrepare
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def firstprepare():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.FirstPrepare"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if ChosenTargetFolder Is Nothing Then:
    # Set ChosenTargetFolder = GetFolderByName("Erhalten", _
    # beginInFolder:=FolderBackup, _
    # noSearchFolders:=True)
    # Set Folder(2) = ChosenTargetFolder
    if Folder(1) Is Nothing Then:
    # Set Folder(1) = ChosenTargetFolder
    # Set LF_CurLoopFld = Folder(1)
    # Call Initialize_UI     ' displays options dialogue
    match rsp:
        case vbCancel:
    # Call LogEvent("=======> Stopped before processing any items . Time: " _
    # & Now(), eLnothing)
    if TerminateRun Then:
    # GoTo ProcReturn
    # End ' abort
    # GoTo ProcReturn
        case Else   ' loop Candidates:
    if topFolder Is Nothing Then:
    # Set topFolder = LookupFolders.Item(LF_DoneFldrCount)
    # Call FindTrashFolder
    # Call InitFindSelect
    # DidFindInits = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CheckItemProcessed
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkitemprocessed():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "QueryMatches.CheckItemProcessed"
    # Call DoCall(zKey, "Sub", eQzMode)

    # DidFindInits = True
    # Set ActItemObject = oneItem
    # Call SelectAndFind

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub InitFindModel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initfindmodel():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.InitFindModel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # TrueCritList = vbNullString                       ' just make consistent, do not change
    # eOnlySelectedFolder = False
    # FindMatchingItems = True
    # Call getCriteriaList

    # IsComparemode = True                    ' as opposed to delete/doubles

    # aPindex = 1
    # Call GetITMClsModel(Item, aPindex)
    # Call aItmDsc.SetDscValues(Item, withValues:=False)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitFindSelect
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initfindselect():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.InitFindSelect"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call SelectAndCompare(DontDecode:=True)  ' only one, no decode, aID(1).idObjItem
    # Call InitFindModel(ActiveExplorerItem(1))   ' use this for New cObjDsc

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function MatchingItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def matchingitems():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.MatchingItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # '                      = -1: just get the count of items that match
    # '                      = 0:  delete duplicates except the newest
    # '                      = 1:  let user decide what to delete
    # '                      = 2:  Interactive Answer
    # Dim MandatoryWorkRule As cNameRule
    # Dim MyMsg As String
    if DidFindInits Then:
    # DidFindInits = False ' Preselection valid only once
    else:
    # Call InitFindSelect
    # Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)
    # Call find_Corresponding(ActItemObject, _
    # CritList:=MandatoryWorkRule, _
    # howmany:=MatchingItems, _
    # eliminateID:=False)

    # Matches = MatchingItems
    if Matches = 0 Then:
    # MyMsg = "Es wurden keine bereinstimmende Items"
    elif Matches = 1 Then:
    # MyMsg = "Es wurde nur ein bereinstimmendes Item"
    else:
    # MyMsg = "Es wurden " & Matches & " bereinstimmende Items"
    # MyMsg = MyMsg & " in " & Quote(Folder(2).FolderPath) & " gefunden"
    # MyMsg = MyMsg & vbCrLf & "   Kriterien: " _
    # & Replace(MandatoryWorkRule.CritRestrictString, " And ", _
    # vbCrLf & vbTab & "And ")
    # ' nb: kleines "and" wird nicht ersetzt

    if MatchMode = -1 Then:
    # Call CleanUpRun(False)
    # GoTo ProcReturn

    if DebugMode Then:
    print(Debug.Print MyMsg)
    if MatchMode = 0 And MatchingItems < 2 Then ' too few to remove duplicates:
    # GoTo Finish ' with termination E

    if MatchMode > 0 Then:
    if Matches = 0 Then:
    # Folder(1).Display
    # GoTo Finish
    else:
    # Folder(2).Display   ' prepare for FilterDisplay
    # ActiveExplorer.ClearSelection
    # ActiveExplorer.AddToSelection ActItemObject
    if Matches = 1 Then:
    # Set ActItemObject = RestrictedItemCollection.Item(1)
    # & vbCrLf & "   Yes: das Item wird dargestellt" _
    # & vbCrLf & "   No:  das Item wird selektiert" _
    # & vbCrLf & "   Cancel: nichts tun, weiter ausfhren" _
    # , vbYesNoCancel + vbDefaultButton2)
    if rsp = vbYes Then:
    # ActItemObject.Display
    # GoTo Finish
    elif rsp = vbNo Then:

    # Call FilterDisplay(MandatoryWorkRule.CritFilterString)
    # GoTo Finish
    else:
    # GoTo Finish
    else:
    # Call FilterDisplay(MandatoryWorkRule.CritFilterString)
    # GoTo Finish

    if xReportExcel Then   ' decode for display in excel not wanted:
    # aPindex = 1
    # Call startReportToExcel
    # Call ReportMatchItems(MatchingItems)
    # Call endReportToExcel
    else:
    if MatchMode = 0 Then:
    # DoVerify False
    # dcCount = Matches
    # Call QueryAboutDelete(CStr(dcCount) & " Items matching selection item")

    elif MatchMode = 1 Then:
    # Call PutSelectedItemDataIntoList
    else:
    # ActiveExplorer.ClearSearch
    # GoTo Finish

    # Call DoTheDeletes
    # Finish:
    # Call CleanUpRun

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FilterDisplay
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def filterdisplay():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.FilterDisplay"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Filter As String
    # Filter = Restrictions
    if InStr(Filter, "[") > 0 Then:
    # Filter = Replace(Filter, "] = ", ":")
    # Filter = Replace(Filter, "] > ", ":>")
    # Filter = Replace(Filter, "] < ", ":<")
    # Filter = Replace(Filter, "] >= ", ":>=")
    # Filter = Replace(Filter, "] <= ", ":<=")
    # Filter = Replace(Filter, "'", Q, 1, -1)
    # Filter = Replace(Filter, "[", vbNullString)

    # ' Filter=replace(filter,a,b)  ' for each word not following syntax or translated
    # Filter = Replace(Filter, " and ", " A ", 1, -1, vbTextCompare)
    # Filter = Replace(Filter, " or ", " OR ", 1, -1, vbTextCompare)
    # Filter = Replace(Filter, "subject:", "betreff:", 1, -1, vbTextCompare)
    # Filter = Replace(Filter, "sendername:", "von:", 1, -1, vbTextCompare)
    # Filter = Replace(Filter, "senton:", "gesendet:", 1, -1, vbTextCompare)
    # Filter = Replace(Filter, "received:", "erhalten:", 1, -1, vbTextCompare)
    if DebugMode Then:
    print(Debug.Print Restrictions)
    print(Debug.Print Filter)
    # ActiveExplorer.Search Filter, olSearchScopeCurrentFolder

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PutSelectedItemDataIntoList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def putselecteditemdataintolist():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.PutSelectedItemDataIntoList"
    # Dim ReShowFrmErrStatus As Boolean

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Call AddItemToList(CStr(Matches), "Gefundene bereinstimmende Items zu " _
    # & RestrictCriteriaString, vbNullString, vbNullString)
    # Call AddItemToList(CStr(i), RestrictedItemCollection(i).Subject, vbNullString, vbNullString)
    if DateSkipCount > 0 Then:
    # ListContent(ListCount).MatchData = "vor dem"
    # ListContent(ListCount).DiffsRecognized = CStr(CutOffDate)

    if ListContent.Count > 0 Then:
    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
    # ReShowFrmErrStatus = True
    # Set FRM = New frmDeltaList
    # Call ShowOrHideForm(FRM, ShowIt:=True)
    # Set FRM = Nothing
    # endsub:
    # Set ListContent = Nothing

    # FuncExit:
    if ReShowFrmErrStatus Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindSchema
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findschema():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.FindSchema"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ResCell As Range
    # Dim ResRow As Long
    # Dim LastPos As Long
    # Dim ResType As String
    # Dim CaseMatch As Boolean

    if LenB(adName) = 0 Then:
    # GoTo FunExit
    if S Is Nothing Then:
    # Call OpenAllSchemata
    else:
    # S.xlTSheet.Select
    # With S.xlTSheet
    # CaseMatch = True
    # .Range("A1").Select
    # trymore:
    # Set ResCell = .Cells.Find(what:=adName, _
    # After:=xlApp.ActiveCell, LookIn:=xlFormulas, _
    # LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    # MatchCase:=CaseMatch, SearchFormat:=False)
    if ResCell Is Nothing Then:
    if CaseMatch Then:
    # CaseMatch = False
    # GoTo trymore
    # NotThere:
    # FindSchema = "*" & adName & " noch nicht in " _
    # & MapiItemType & " gefunden*"
    print(Debug.Print FindSchema)
    # GoTo FunExit

    if ResCell.Row < LastPos Then:
    if ResRow > ResCell.Row Then:
    # GoTo NotThere

    # GoTo incLast

    # ResRow = ResCell.Row
    # .Range("A" & ResRow).Select
    # ResType = LCase(xlApp.Selection.Value)

    if ResType <> LCase(MapiItemType) Then:
    # CaseMatch = True
    # incLast:
    if LastPos > ResRow Then:
    # LastPos = LastPos + 1
    else:
    # LastPos = ResRow + 1
    # .Range("A" & LastPos).Select
    # ResRow = LastPos
    # GoTo trymore
    # ' got a match!
    # End With ' S.xlTSheet
    # FindSchema = xlApp.Selection.Cells(1, 4).Text
    # FunExit:

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ModObligMatches
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def modobligmatches():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.ModObligMatches"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if sRules Is Nothing Then:
    # Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, withMatchBits:=False)
    # Set iRules = Nothing
    # sRules.clsObligMatches.ChangeTo = NewObligAttribs
    # ' in ChangeTo, .clsObligMatches.RuleMatches signals change
    # ModObligMatches = sRules.clsObligMatches.RuleMatches
    if ModObligMatches Then  ' re-check consistency before use:
    # sRules.RuleInstanceValid = False ' never is before we check
    if Not aTD Is Nothing Then:
    # ' Call aTD.adRules.AllRulesCopy(InstanceRule, sRules, withMatchBits:=False)
    # Call Get_iRules(aTD)
    # SelectedAttributes = vbNullString ' append Similarities later
    else:
    if Not aTD Is Nothing Then:
    # Call Get_iRules(aTD)
    # ' the following True... are shortcuts only for speed
    # TrueCritList = Trim(iRules.clsObligMatches.aRuleString)
    # TrueImportantProperties = iRules.clsObligMatches.MatchesList

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GenCrit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def gencrit():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.GenCrit"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim j As Long
    # Dim tDictIndex As Long
    # Dim gItemProp As ItemProperty
    # Dim sT As String
    # Dim PorOP As String ' CritPropName or operator
    # Dim nextOperator As String
    # Dim CritRestrictElement As cFilterCriterium
    # Dim CritFilterElement As cFilterCriterium
    # Dim UseTimeCompare  As Boolean  ' Time without seconds only, >= instead of =
    # Dim TimeString As String
    # Dim eTimeString As String
    # Dim RawValue As Variant
    # Dim PropertyIdent As String
    # Dim TimeEndAdder As Long
    # Dim ValSep As String
    # Dim tempstr As String

    # ' #########################
    # j = LBound(Criteria.CleanMatches)
    # i = LBound(Criteria.MatchesList)
    # TimeEndAdder = 1 ' normally, one minute
    # Set SQLpropC = Nothing  ' new criteria

    # startloop:
    # While i <= UBound(Criteria.MatchesList) And j <= UBound(Criteria.CleanMatches)
    # Set CritRestrictElement = New cFilterCriterium
    # CritRestrictElement.CritType = 1

    # With CritRestrictElement
    # UseTimeCompare = False
    # .ValueIsTimeType = UseTimeCompare
    if SQLpropC Is Nothing Then:
    # Set SQLpropC = New Collection
    # .CritIndex = j + 1
    # .CritPropName = Criteria.CleanMatches(j)
    # PropertyNameX = .CritPropName
    # ' and formatted value
    # .adFormattedValue = aID(1).idAttrDict.Item(.CritIndex).adFormattedValue

    # ' get itemProperty for this attribute
    # tDictIndex = CInt(aID(1).idAttrDict.Item(.CritIndex).adtrueIndex)
    # Set gItemProp = aID(1).GetAttrDsc4Prop(tDictIndex).adItemProp
    # DoVerify aID(1).idAttrDict.Item(.CritIndex).adName = gItemProp.Name
    # DoVerify Criteria.CleanMatches(j) = gItemProp.Name
    # ' get original raw value as variant if possible
    # RawValue = vbNullString
    # aBugTxt = "Get Item Property Value"
    # Call Try
    # RawValue = gItemProp.Value
    # Catch

    # PorOP = Criteria.MatchesList(i)
    # sT = Left(PorOP, 1)
    if sT = "!" Then:
    # ' ! means value to stop building CriteriaString used in FI,
    # ' keeping most relevant
    # GoTo EndWhile

    if UCase(PorOP) = "OR" Then:
    # GoTo swallowit
    elif UCase(PorOP) = "NOT" Then:
    # GoTo swallowit
    elif UCase(PorOP) = "TOETIME" Then:
    # i = i + 1   ' skip this, use next as parameter
    # TimeEndAdder = Criteria.MatchesList(i)
    # GoTo swallowit
    elif UCase(PorOP) = "A" Then:
    # swallowit:
    # GoTo incI
    else:
    # buildIt:
    if LenB(PorOP) = 0 Then   ' was an operator only, not just prefix:
    # PorOP = Criteria.MatchesList(i)
    # GoTo SkipOpOnly
    # While Right(PorOP, 1) = ")"
    # .cBracket = .cBracket & ")"
    # PorOP = Mid(PorOP, 1, Len(PorOP) - 1)
    # .BracketOpenCount = .BracketOpenCount - 1
    # Wend
    # While Left(PorOP, 1) = "("
    # .oBracket = .oBracket & "("
    # PorOP = Mid(PorOP, 2)
    # .BracketOpenCount = .BracketOpenCount + 1
    # Wend

    # sT = Left(PorOP, 1)
    match porop:
        case "+":
    # .Comparator = " = "
    # .Operator = " And "
    # PorOP = Mid(PorOP, 2)
    # GoTo buildIt
        case "|":
    # .Comparator = " = "
    # .Operator = " Or "
    # PorOP = Mid(PorOP, 2)
    # GoTo buildIt
        case "-":
    # .Comparator = " <> "  ' And Not ?? operator = " And Not "
    # PorOP = Mid(PorOP, 2)
    # GoTo buildIt
        case "%":
    # DoVerify isSQL, "GenCrit Criteria might not work"
    # .Comparator = " Like "
    # isSQL = True
    # PorOP = Mid(PorOP, 2)
    # GoTo buildIt
        case "~":
    # .Comparator = vbNullString
    # PorOP = Mid(PorOP, 2)
    # StringMod = False
    # tempstr = Append(Trim(sRules.clsSimilarities.aRuleString), PorOP)
    if StringMod Then:
    # sRules.clsSimilarities.ChangeTo = _
    # tempstr _
    # ' ??? clean this assignment
        case _:
    if LenB(.Comparator) = 0 Then:
    # .Comparator = " = "
    if .CritIndex = 1 Then:
    if isSQL Then:
    # .Operator = "@SQL="
    else:
    # .Operator = vbNullString
    else:
    # .Operator = nextOperator
    # .BracketOpenCount = SQLpropC.Item(.CritIndex - 1).BracketOpenCount
    # DoVerify .BracketOpenCount >= 0, " error: more closing brackets than open ones!"
    # ' special for date/time
    if IsDate(.adFormattedValue) Then:
    if InStr(.adFormattedValue, ":") > 0 Then ' contains time:
    # UseTimeCompare = True
    # .ValueIsTimeType = UseTimeCompare

    if LenB(.Comparator) > 0 Then   ' any valid condition?:
    if isSQL Then:
    # .ValueSeperator = "'%"
    else:
    # .ValueSeperator = "'"
    # PropertyIdent = fixPropertyname(PorOP, isSQL)
    if .ValueIsTimeType Then:
    # ' NO seconds, check endtime, conditionally UTC base
    # TimeString = StandardTime(gItemProp, (RawValue), UTCisUsed)
    # TimeString = Format(TimeString, "dd.mm.yyyy hh:nn")
    # ' add endtime TimeEndAdder (+= 1 minute as default)
    # eTimeString = DateAdd("n", TimeEndAdder, TimeString)
    # eTimeString = Format(eTimeString, "dd.mm.yyyy hh:nn")
    if TimeEndAdder < 0 Then:
    # Call Swap(TimeString, eTimeString)
    # .adFormattedValue = TimeString
    # Call AppendTo(.oBracket, "(", always:=True, ToFront:=True)
    # .Comparator = " >= "
    # .AttrRawValue = RawValue
    # .PropertyIdent = PropertyIdent
    # ValSep = .ValueSeperator    ' save for next CritRestrictElement
    # .addTo SQLpropC   ' add start time to collection
    # Set CritFilterElement = CritRestrictElement.Clone(2)
    # CritFilterElement.addTo SQLpropC
    # Set CritRestrictElement = New cFilterCriterium
    # End With ' CritRestrictElement
    # With CritRestrictElement  ' start over because of change if UseTimeCompare
    if UseTimeCompare Then:
    # .ValueIsTimeType = True
    # .ValueSeperator = ValSep    ' restore from prev CritRestrictElement
    # ' create End-Time aspect
    # .CritPropName = Criteria.CleanMatches(j)
    # .CritIndex = SQLpropC.Count + 1
    # .adFormattedValue = eTimeString
    # .Comparator = " <= "
    # Call AppendTo(.cBracket, ")", always:=True, ToFront:=False)
    # ' operator now and
    # .Operator = " and "
    # .AttrRawValue = RawValue
    # .PropertyIdent = PropertyIdent
    # .CritType = 1
    # .addTo SQLpropC   ' add it to collection
    # SkipOpOnly:
    # j = j + 1
    # incI:
    # i = i + 1
    # End With ' CritRestrictElement
    # Set CritFilterElement = CritRestrictElement.Clone(2)
    # CritFilterElement.addTo SQLpropC
    # Wend    ' criteria loop
    # EndWhile:
    # DoVerify CritRestrictElement.BracketOpenCount = 0, " error: more opening brackets than terms!"

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' build Match Criteria as string (e.g. for .Restrict) from criteria
def buildfindcriteria():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.BuildFindCriteria"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim CritRestrictElement As cFilterCriterium
    # Dim RestrictCompareTerm As String
    if Criteria Is Nothing Then ' specific call to use RestrictCriteriaString:
    # BuildFindCriteria = RestrictCriteriaString ' without rule applied
    # ' make sure we use intended sRules.clsObligMatches
    if Not sRules Is Nothing Then    ' ! really???:
    # DoVerify False
    # sRules.clsObligMatches.CritRestrictString = RestrictCriteriaString
    else:
    # Call GenCrit(Criteria)
    # RestrictCriteriaString = vbNullString
    # FilterCriteriaString = vbNullString

    # Set CritRestrictElement = SQLpropC.Item(i)
    if LenB(CritRestrictElement.Comparator) > 0 Then:
    if LenB(CritRestrictElement.PropertyIdent) > 0 _:
    # And Mid(CritRestrictElement.PropertyIdent, 2, 1) <> "*" Then
    # RestrictCompareTerm = CritRestrictElement.Operator _
    # & CritRestrictElement.oBracket _
    # & CritRestrictElement.PropertyIdent _
    # & CritRestrictElement.Comparator _
    # & QuoteWithDoubleQ(CritRestrictElement.adFormattedValue, _
    # CritRestrictElement.ValueSeperator) _
    # & CritRestrictElement.cBracket ' no term yet
    if isEmpty(CritRestrictElement.AttrRawValue) Then:
    if DebugMode Then DoVerify False, " value missing in object":
    else:
    if CritRestrictElement.CritType = 1 Then:
    # RestrictCriteriaString = RestrictCriteriaString _
    # & RestrictCompareTerm
    else:
    # FilterCriteriaString = FilterCriteriaString _
    # & RestrictCompareTerm
    # Criteria.CritRestrictString = RestrictCriteriaString
    # Criteria.CritFilterString = FilterCriteriaString
    # BuildFindCriteria = RestrictCriteriaString

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' gets RestrictedItems and adds entries into RestrictedItemCollection
def getrestricteditems():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.GetRestrictedItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim limCount As Long
    # Dim A As String
    # Dim pushFolder As Folder
    # Dim ReviseTextForm As Object
    # Dim MandatoryWorkRule As cNameRule
    # Dim WeRevisedRestrictCriteriaString As Boolean
    # Dim LocalRestrictString As String
    # Dim RuleString As String

    if Criteria Is Nothing Then                 ' use RestrictCriteriaString as is:
    # Call BestObjProps(FolderToSeekIn, matchToItemO, withValues:=False)
    # Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)
    # ' will not set up new RestrictCriteriaString unless in error
    # DoVerify IsMailLike(matchToItemO), _
    # "design expects only Mail-like Items, but is " _
    # & MandatoryWorkRule.PropAllRules.ARName
    # RuleString = matchToItemO.Subject
    if InStr(RuleString, "'") > 0 Then      ' this could contain Quotes, so double these:
    # RuleString = DoubleInternalQuotes(RuleString, "'")
    if InStr(RuleString, Q) > 0 Then        ' or even Double quotes ":
    # RuleString = DoubleInternalQuotes(RuleString, Q)

    # LocalRestrictString = "[" & MandatoryWorkRule.aRuleString & "] = '" _
    # & RuleString & "' And " & RestrictCriteriaString
    # MandatoryWorkRule.MatchOn = "Working=" & Quote(LocalRestrictString)
    else:
    # Set MandatoryWorkRule = Criteria    ' use RestrictCriteriaString from Rule
    # RestrictCriteriaString = MandatoryWorkRule.aRuleString
    # LocalRestrictString = RestrictCriteriaString
    # Message = vbNullString
    if RestrictedItemCollection.Count > 0 Then:
    # Set RestrictedItemCollection = New Collection
    # aBugTxt = "Restrict folder " & FolderToSeekIn.FolderPath _
    # & " Using " & LocalRestrictString
    # Call Try(allowNew)
    # Set RestrictedItems = FolderToSeekIn.Items.Restrict(RestrictCriteriaString)
    if Catch Then:
    # DoVerify False, "Unable to continue"

    if TestCriteriaEditing <= vbCancel Then ' undefined or OK: get default again:
    if TestCriteriaEditing <> vbOK And TestCriteriaEditing <> vbCancel Then:
    if DebugLogging And TestCriteriaEditing <> vbOK _:
    # Or TestCriteriaEditing = vbCancel Then
    # TestCriteriaEditing = vbYes
    else:
    # TestCriteriaEditing = vbNo

    # tryAgain:
    # While Err.Number <> 0 Or TestCriteriaEditing = vbYes
    # Set ReviseTextForm = New frmLongText    ' init asks if edit wanted=VbYes / cancel stops
    if ErrorCaught <> 0 Then:
    # Message = "Suche passende items in " & Quote(FolderToSeekIn.FolderPath) _
    # & " ergab Fehler " & Err.Description
    else:
    # Message = "Suche passende items in " & Quote(FolderToSeekIn.FolderPath) _
    # & " ergab " & RestrictedItems.Count & " Ergebnisse"
    # Call ErrReset(0)
    # Call ReviseTextForm.UserMsg(Message)
    # Call ReviseTextForm.SetText(RestrictCriteriaString, " And ")
    if Not ReviseTextForm.Visible Then:
    # ReviseTextForm.Show
    # RestrictCriteriaString = ReviseTextForm.EditedText()
    # Call ErrReset(4)
    if TestCriteriaEditing = vbCancel Then  ' that indicates we did not change now and exit the loop:
    # limCount = 0
    elif TestCriteriaEditing = vbOK Then  ' that indicates we are happy with the selection criteria:
    else:
    # ' start over with new RestrictCriteriaString: so delete previous
    # Set RestrictedItemCollection = New Collection
    # ErrReset
    # Set RestrictedItems = FolderToSeekIn.Items.Restrict(LocalRestrictString)   ' replace name appendage
    # WeRevisedRestrictCriteriaString = True
    # Wend
    # Set ReviseTextForm = Nothing
    # findCount = Add2RestrictedCol(matchToItemO, eliminateID)

    # limCount = RestrictedItemCollection.Count
    if eliminateID Then:
    # A = " andere passenden Items"
    # findCount = limCount - 1
    else:
    # A = " passende Items"
    # findCount = limCount
    if findCount < 0 Then   ' ==> we did elimiateID the self search already:
    # A = "kein weiteres passendes Item in " & Quote(matchToItemO.Parent.FolderPath)
    # GoTo CheckConditions
    elif limCount < 1 Or (eliminateID And limCount < 2) Then:
    # Message = A & " nicht gefunden"
    else:
    # Message = findCount & A & " gefunden"
    if findCount > 20 Then:
    # Call LogEvent(Message)
    # & RestrictCriteriaString & vbCrLf & vbCrLf _
    # & " (Nein: nur 20 items)", vbYesNoCancel, matchToItemO.Subject)
    if rsp = vbCancel Then:
    # GoTo cleanup
    elif rsp = vbNo Then:
    print(Debug.Print "verarbeite nur 20 der " & findCount & " items")
    # limCount = 20
    else:
    print(Debug.Print "alle " & limCount & " passenden items werden verwendet!")

    # CheckConditions:
    # Message = Message & " in " & Quote(FolderToSeekIn.FolderPath) _
    # & " Kriterien:" & vbCrLf & vbTab _
    # & Replace(LocalRestrictString, " And ", vbCrLf & vbTab & "And ")
    if ErgebnisseAlsListe Then:
    print(Debug.Print Message)
    # Call RestrictedItemsShow(limCount)

    if findCount > 0 And (limCount = 0 _:
    # Or (Not eliminateID And findCount < 1)) _
    # And (DebugLogging Or ActionID = 0) Then
    # Call LogEvent(Message)
    if ActionID = 0 Then:
    # & "Kriterien korrekt?              (Nein: erlaube nderung)", _
    # vbYesNo, "Gesuchtes Objekt nicht vorhanden")
    if TestCriteriaEditing = vbYes Then:
    # GoTo cleanup
    # Message = Message & vbCrLf _
    # & "Ja: Direktes Verndern der verwendeten Regeln " & vbCrLf _
    # & "Nein: Direktes Verndern der generierten Abfrage " & vbCrLf _
    # & "Abbruch: Problem ignorieren " & vbCrLf _
    # & "     ggf. Zeilenumbruch bei '|' beachten!"
    if rsp <> vbNo Then:
    # Call ReviseTextForm.UserMsg(Message)
    # Call frmLongText.UserMsg(Message)
    # TrueCritList = frmLongText.TextEdit(TrueCritList)
    # frmLongText.Show
    if rsp = vbCancel Then  ' use default again next time:
    # GoTo cleanup
    # Call SplitMandatories(TrueCritList, MandatoryWorkRule)
    # RestrictCriteriaString = BuildFindCriteria(MandatoryWorkRule)
    # GoTo testCriteria
    elif TestCriteriaEditing = vbYes Then  ' allow user to edit:
    # RestrictCriteriaString = frmLongText.TextEdit(RestrictCriteriaString)
    # frmLongText.Show
    if TestCriteriaEditing = vbCancel Then:
    # GoTo cleanup
    # testCriteria:
    if LenB(RestrictCriteriaString) = 0 Then   ' restore the original RuleString:
    # TrueCritList = MandatoryWorkRule.aRuleString
    # Set pushFolder = LF_CurLoopFld      ' save
    # Set LF_CurLoopFld = FolderToSeekIn
    # Call Initialize_UI
    if DebugMode Then   ' not expecting Debug.Assert False::
    # DoVerify LF_CurLoopFld.FolderPath = pushFolder.FolderPath
    # Set LF_CurLoopFld = pushFolder      ' restore
    # Call BuildFindCriteria(Nothing)  ' special call preserving RestrictCriteriaString
    # Set RestrictedItemCollection = New Collection
    # WeRevisedRestrictCriteriaString = True
    # GoTo tryAgain
    else:
    # Call LogEvent(Message)
    # cleanup:
    # TestCriteriaEditing = vbOK  ' default used again next time

    # FuncExit:
    # Set ReviseTextForm = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function Add2RestrictedCol
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def add2restrictedcol():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "QueryMatches.Add2RestrictedCol"
    # Call DoCall(zKey, "Function", eQzMode)

    # Dim i As Long
    # Dim aFileSpec As Object

    # Add2RestrictedCol = 0
    # Set aFileSpec = RestrictedItems.Item(i)

    if eliminateID = True And matchToItemO.EntryID = aFileSpec.EntryID Then ' do not look for identical ones:
    print(Debug.Print vbNullString, "= " & i, aFileSpec.EntryID)
    if eliminateID Then:
    # Call LogEvent(vbTab & "= " & i & vbTab _
    # & "nicht bercksichtigt weil identisch mit Suchobjekt")
    else:
    if matchToItemO.Subject = aFileSpec.Subject Then ' double check!:
    # Add2RestrictedCol = Add2RestrictedCol + 1
    # RestrictedItemCollection.Add aFileSpec
    else:
    # ' or the subjects are very similar.
    if DebugLogging Then DoVerify False:

    # FuncExit:
    # Set RestrictedItems = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub find_Corresponding
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def find_corresponding():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.find_Corresponding"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim NondefaultMatches As Variant
    if Folder(2) Is Nothing Then:
    # Call PickAFolder(2, "in diesem Ordner hnliche suchen", "Whle Ordner", "OK", "Cancel")
    # Stop ' ???
    # Set aID(1).idAttrDict = New Dictionary
    # AttributeUndef(1) = 0
    # Set aID(2).idAttrDict = New Dictionary
    # AttributeUndef(2) = 0
    # Set matchToItemO = GetAobj(1, -1)
    # objTypName = DecodeObjectClass(getValues:=False)
    # NondefaultMatches = TrueImportantProperties
    # sRules.ARName = aObjDsc.objTypeName         ' do NOT reinitialize ..ImportantProperties
    # Call SetCriteria                            ' using TrueImportantProperties

    if quickChecksOnly Then:
    # AttributeIndex = -2                     ' check and decode all mandatory properties
    else:
    # AttributeIndex = -1                     ' check all ( most important ONLY would be -2 )
    # MaxPropertyCount = Max(MaxPropertyCount, aID(1).idAttrDict.Count)

    if quickChecksOnly Then:
    # AttributeIndex = -2  ' next time, do all ???
    else:
    # AttributeIndex = -1
    # Call BuildFindCriteria(CritList)
    # Call GetRestrictedItems(matchToItemO, Folder(2), CritList, _
    # howmany, eliminateID:=eliminateID)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' delivers duplicates in RestrictedItemCollection or RestrictedItems
def finduniqueemailitems():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.findUniqueEmailItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim findUniqueEmailItem As MailItem
    # Dim tS As String
    # Dim tEnd As String
    # Dim absMtDiff As Double
    # Dim oDate As Date
    # Dim timeFilter As String

    match matchToItemO.Class:
        case olMail, olMeetingRequest, olMeetingResponseTentative, olMeetingResponseNegative, olMeetingResponsePositive:
        case olReport, olMeetingCancellation, olTaskRequest:
    # howmany = 0     ' these "mails" have no Sent on or Received Time, so that's ok
    # Set findUniqueEmailItem = Nothing   ' it's not of class MailItem
    # GoTo ProcReturn
        case Else      ' Other cases not expected yet:
    # DoVerify False

    if WhichTime = "SentOn" Then:
    # oDate = matchToItemO.SentOn
    if oDate = BadDate Then:
    # oDate = matchToItemO.ReceivedTime
    # WhichTime = "ReceivedTime"
    if LenB(aTimeFilter) = 0 Then:
    # aTimeFilter = WhichTime
    elif DebugMode Then:
    if WhichTime <> aTimeFilter Then:
    if DebugLogging Then:
    print(Debug.Print "Unusual comparison: WhichTime <> aTimeFilter " _)
    # & WhichTime & " <> " & aTimeFilter
    if maxTimeDiff < 0 Then:
    # tS = Format(DateAdd("n", maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
    # tEnd = Format(DateAdd("n", -maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
    # absMtDiff = -maxTimeDiff
    else:
    # tS = Format(oDate, "dd.mm.yyyy hh:mm")  ' start +0
    # tEnd = Format(DateAdd("n", maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
    # absMtDiff = maxTimeDiff
    # ' using >= to avoid a bug (will not find exact match)
    # RestrictCriteriaString = "[SenderName] = " & Quote1(matchToItemO.SenderName)
    # timeFilter = " And [" & aTimeFilter & "] >= " & Quote1(tS)
    if absMtDiff >= 1# Then  ' range must be at least 1 Minutes (Seconds not available):
    # timeFilter = timeFilter & " And [" & aTimeFilter & "] < " & Quote1(tEnd)
    # Call LogEvent("  -- locating similar items in " & Quote(inFolder.FolderPath) _
    # & vbCrLf & "      for " & Quote(matchToItemO.Subject) _
    # & ", original item size=" & matchToItemO.Size _
    # & vbCrLf & LString("       0", 4) _
    # & matchToItemO.EntryID & vbCrLf _
    # & "      at around " & tEnd)
    # RestrictCriteriaString = RestrictCriteriaString & timeFilter
    # ' Criteria = Nothing means: RestrictCriteriaString is built already
    # ' Call BuildFindCriteria(inFolder, Nothing)  ' so we assume this was already done on caller side
    # Call GetRestrictedItems(matchToItemO, _
    # FolderToSeekIn:=inFolder, _
    # Criteria:=Nothing, _
    # findCount:=howmany, _
    # eliminateID:=GetFirstOnly)  ' nothing forces use of RestrictCriteriaString
    # howmany = FindNonUnique(matchToItemO, oDate, maxTimeDiff)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindNonUnique
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findnonunique():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.FindNonUnique"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim tS As String
    # Dim sameID As Boolean
    # Dim sufficientlySameButDifferentId As Boolean
    # Dim aDiffTimeInMinutes As Double
    # Dim afilespecObject As Object
    # Dim aMtDiff As Double
    # Dim sign As String
    # Dim eDate As Date
    # Dim dDate As Date

    # FindNonUnique = RestrictedItemCollection.Count
    if FindNonUnique <= 0 Then:
    # GoTo ProcReturn
    if FindNonUnique = 1 Then   ' Only the original was found:
    if DebugMode Then:
    # Call checkDates(matchToItemO, RestrictedItemCollection(1))
    # GoTo ProcReturn
    if mtDiff < 0 Then:
    # sign = Chr(177)     ' +-
    # aMtDiff = -mtDiff
    else:
    # sign = "+"
    # aMtDiff = mtDiff
    if DebugMode Then:
    print(Debug.Print "        we found " & FindNonUnique & " items which match")
    # tS = vbNullString
    if matchToItemO.Class <> RestrictedItemCollection.Item(i).Class Then:
    # RestrictedItemCollection.Remove i
    # FindNonUnique = RestrictedItemCollection.Count
    if i > FindNonUnique Then:
    # Exit For
    # Set afilespecObject = RestrictedItemCollection.Item(i)
    if matchToItemO.Subject <> afilespecObject.Subject Then:
    # tS = " ??? item " & i _
    # & " has a different Subject " & Quote(afilespecObject.Subject)
    # GoTo someSecondsOff
    if aTimeFilter = "SentOn" Then:
    # eDate = afilespecObject.SentOn
    else:
    # eDate = afilespecObject.ReceivedTime
    # dDate = DateDiff("n", eDate, oDate)
    # sameID = matchToItemO.EntryID = afilespecObject.EntryID
    # aDiffTimeInMinutes = Abs(dDate)
    if sameID Then:
    if aDiffTimeInMinutes = 0 Then:
    # tS = ", Exact time "
    # GoTo someSecondsOff
    else:
    # DoVerify False
    # GoTo littleDiff
    if aDiffTimeInMinutes = 0 Then:
    # sufficientlySameButDifferentId = True
    elif aDiffTimeInMinutes > 0 Then:
    # littleDiff:
    # tS = ", time "
    if aDiffTimeInMinutes < 3 Then:
    # sufficientlySameButDifferentId = True   ' ignore diff in IDs when times are very similar
    # tS = tS & "delta< "
    elif aDiffTimeInMinutes > aMtDiff Then:
    # tS = tS & "off (by "
    else:
    # tS = tS & "delta> "
    # sameID = sufficientlySameButDifferentId   ' ignore diff in IDs when times are  similar
    # tS = tS & Format(aDiffTimeInMinutes, "@0") & ":00" _
    # & " min, Tolerance = " _
    # & sign & Format(aMtDiff, "0#") & ":00"
    # ' check the item we know is there, because it is identical by ID
    if matchToItemO.Size <> afilespecObject.Size Then:
    try:
        if matchToItemO.BodyFormat <> afilespecObject.BodyFormat _:
        # And afilespecObject.BodyFormat <> olFormatHTML Then
        if matchToItemO.Body <> afilespecObject.Body Then:
        # tS = tS & " Size differs by " _
        # & afilespecObject.Size - matchToItemO.Size _
        # & ", size=" & afilespecObject.Size
        # noBodyFormat:
        # sameID = False  ' can'tS be same ID ?!
        # sufficientlySameButDifferentId = False
        else:
        # tS = tS & ", BodyFormat in target had been changed to HTML"

        if sameID Or sufficientlySameButDifferentId Then:
        # someSecondsOff:
        # FindNonUnique = FindNonUnique - 1   ' correct for true elements
        if DebugMode Or DebugLogging Then:
        # Call checkDates(RestrictedItemCollection(1), afilespecObject)
        if sameID And i > 1 _:
        # And RestrictedItemCollection(1).EntryID <> afilespecObject.EntryID Then
        if RestrictedItemCollection(1).CreationTime <= afilespecObject.CreationTime Then:
        # ' move the best=oldest match to position 1
        # ' so we never delete the best match
        # ' note: now RestrictedItems are not conform regarding index
        # ' unfortunately, swapping items in a collection does not work. So we try to delete the younger one instead.
        # ' Call Swap(RestrictedItemCollection.Item(i), RestrictedItemCollection.Item(1), asObject:=True)
        # RestrictedItemCollection.Item(1).Delete         ' remove from folder
        # RestrictedItemCollection.Remove 1               ' remove from collection (shifting down)
        # Call LogEvent("Removed RestrictedItemCollection positions 1 because it is younger than that in pos=" & i)
        else:
        # RestrictedItemCollection.Item(i).Delete         ' remove from folder
        # RestrictedItemCollection.Remove i               ' remove from collection (shifting down)
        # Call LogEvent("Removed RestrictedItemCollection pos=" & i & " 1 because it is younger than that in pos=1")
        # i = i - 1                                           ' do not advance loop
        if FindNonUnique <= 0 Then:
        # Exit For

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function fixPropertyname
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def fixpropertyname():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.fixPropertyname"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim msg As String
    if sql Then:
    match PropName:
        case "Categories":
    # fixPropertyname = "urn:schemas-microsoft-com:office:office#Keywords"
        case "LastName":
    # fixPropertyname = "urn:schemas:contacts:sn"
        case "FullName":
    # fixPropertyname = "urn:schemas:contacts:cn"
        case "FirstName":
    # fixPropertyname = "urn:schemas:contacts:givenName"
        case "MiddleName":
    # fixPropertyname = "urn:schemas:contacts:middlename"
        case "FileAs":
    # fixPropertyname = "urn:schemas:contacts:fileas"
        case _:
    # msg = "Kenne das urn schema (noch) nicht fr " & PropName

    # fixPropertyname = FindSchema(aOD(1).objItemClassName, PropName)
    if Left(fixPropertyname, 1) = "*" Then:
    # & "  Ja: Ignorieren" & vbCrLf _
    # & "  Nein: Editieren in Excel", _
    # vbYesNoCancel, "Abfrage/Vergleich ignorieren?")
    if rsp = vbCancel Then:
    print(Debug.Print msg)
    # DoVerify False
    # End
    elif rsp = vbNo Then:
    # fixPropertyname = FindSchema(aOD(1).objItemClassName, _
    # PropName)
    else:
    print(Debug.Print msg)
    print(Debug.Print "Abfrage/Vergleich von " _)
    # & PropName & " wird ignoriert"
    # fixPropertyname = vbNullString
    # GoTo FunExit
    # fixPropertyname = Chr(34) & fixPropertyname & Chr(34)
    else:
    # fixPropertyname = "[" & PropName & "]"
    # FunExit:

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub getCriteriaList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getcriterialist():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.getCriteriaList"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aItmClass As OlObjectClass
    # Dim DefaultCritListForClass As String

    # ' operators are
    # ' + A
    # ' - NOT
    # ' | OR
    # ' ~ relaxed, no Match required
    # ' % like
    # ' ! ignore this and all following Criteria for FIND
    # ' debugMode = True
    # DefaultCritListForClass = TrueCritList
    if SelectedItems Is Nothing Then:
    # useSorted:
    # aItmClass = sortedItems(1).Item(1).Class
    # 'SelectedItems.Add sortedItems(1) ???
    else:
    if SelectedItems.Count = 0 Then:
    # GoTo useSorted
    # aItmClass = SelectedItems.Item(1).Class
    # ' Now, we always have item in "sortedItems"
    if aItmClass = olContactItem Or aItmClass = olContact Then:
    # ' "%LastName |%FirstName"
    if LenB(DefaultCritListForClass) = 0 _:
    # And LenB(Trim(sRules.clsObligMatches.aRuleString)) = 0 Then
    # DefaultCritListForClass = "LastName %FirstName" ' this is like a parameter, used for test only
    else:
    # DefaultCritListForClass = Trim(sRules.clsObligMatches.aRuleString)
    elif aItmClass = olMail Then:
    # DefaultCritListForClass = "Subject SenderName SentOn "
    elif aItmClass = olAppointment Then:
    # DefaultCritListForClass = "Subject Start End IsRecurring Exceptions "
    else:
    if DebugMode Then DoVerify False, " must define value":

    if LenB(TrueCritList) = 0 Then     ' no change, just make consistent:
    # DefaultCritListForClass = Trim(sRules.clsObligMatches.aRuleString)
    # TrueCritList = DefaultCritListForClass
    else:
    # Call SplitMandatories(DefaultCritListForClass)

    if InStr(DefaultCritListForClass, "%") > 0 Then:
    # isSQL = True
    else:
    # isSQL = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub OpenAllSchemata
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def openallschemata():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.OpenAllSchemata"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aHdl As String

    # aHdl = "ItemType- ADName---------------  AttrNameOhneLeer------------- AttributeAccessString------------------------------------------"
    # Set S = xlWBInit(xlA, TemplateFile, "AllSchemata", _
    # aHdl, showWorkbook:=DebugMode, mustClear:=False)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CleanUpRun
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def cleanuprun():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.CleanUpRun"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if Not xlApp Is Nothing Then:
    # Call ClearWorkSheet(xlA, O)
    # StopRecursionNonLogged = False
    if Terminate Then:
    if DebugMode Then DoVerify False:
    if TerminateRun Then:
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ReportMatchItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def reportmatchitems():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.ReportMatchItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim same As Boolean
    # Dim matching As Long
    # Dim Item As Object

    # aPindex = 1
    # Call ItemReportToExcel(1)
    # aPindex = 2
    # targetIndex = aPindex
    # sourceIndex = 0                                 ' do not copy object item properties, only description

    # workingOnNonspecifiedItem = False
    # BaseAndSpecifiedDiffer = False
    # Set Item = RestrictedItemCollection.Item(i)
    if Item.Class <> aID(1).idObjItem.Class Then:
    print(Debug.Print "!!!! target Folder contains items of other type than source")
    # DoVerify False, " target Folder contains items of other type"
    # GoTo skipitem
    # Set aOD(2) = GetITMClsModel(Item, aPindex).idObjDsc
    # Call aItmDsc.SetDscValues(Item, withValues:=True, aRules:=sRules)

    # aOD(0).objDumpMade = 1
    # Set aID(2).idAttrDict = New Dictionary      ' object/item has not been decoded
    # WorkIndex(1) = 1                            ' our base item ???
    # WorkIndex(2) = i
    # SelectOnlyOne = False                       ' takes 2 to compare
    # fiMain(2) = fiMain(1)                       ' by find/findnext
    # same = ItemIdentity()
    if same Then:
    # matching = matching + 1
    # Call ItemReportToExcel(i)
    if DebugMode Then DoVerify False, " testing phase only":
    # skipitem:
    # Set aID(2).idAttrDict = New Dictionary
    # aOD(0).objDumpMade = 1

    if matching = 0 Then:
    if DebugMode Then:
    print(Debug.Print "Found no complete match in " _)
    # & reportedMatches - 1 & " relevant matches"
    # Call CleanUpRun                             ' with termination E
    else:
    if DebugMode Then:
    print(Debug.Print "Found " & matching & " relevant matches")

    # FindMatchingItems = False

    # FuncExit:
    # Set Item = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SetCriteria
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setcriteria():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.SetCriteria"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # ' Out or confirmed values are:
    # ' Globals in AllPublic: TrueImportantProperties (array of strings)
    # '                       Similarities       (array of cleanMatches, these are appended to:)
    # '                       SelectedAttributes (in/out: must be vbNullString if we want to set new (incl. Similarities)
    # '                       TrueCritList       (out as cleaned string)
    # '                       MandatoryMatches   (== _Rulestring, out, criteria with OPs, not clean,
    # '                                           Similarities are not included here)
    # Dim i As Long
    # Dim sT As String
    # Dim sr As String
    # Dim sc As String

    if sRules Is Nothing Then:
    if Not aObjDsc Is Nothing Then:
    # Set sRules = aObjDsc.objClsRules
    if sRules Is Nothing Then:
    # DoVerify False, "remove block if no hit ???"
    # Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, withMatchBits:=False)
    # Set iRules = Nothing
    # sr = Trim(sRules.clsObligMatches.aRuleString)
    # sT = Trim(sRules.clsObligMatches.CleanMatchesString)
    # sc = Trim(sRules.clsSimilarities.CleanMatchesString)
    if LenB(sr) > 0 _:
    # And TrueCritList = sr _
    # And SelectedAttributes = sT & b & sc Then   ' alles da, NOP
    # GoTo ProcReturn                         ' there is no change for sRules

    # SetCriteria = True                          ' we are changing sRule for this class
    if isEmpty(TrueImportantProperties) Then:
    # SelectedAttributes = sRules.clsObligMatches.CleanMatchesString
    else:
    # ' reconstruct from TrueImportantProperties:
    # ' rebuild _RuleString, but controlled:
    # ' from TrueImportantProperties A Similarities (avoiding doubles)
    # TrueCritList = vbNullString                       '  rebuilt WITH special chars from both
    # SelectedAttributes = vbNullString                 '         ONLY from TrueImportantProperties
    # ' built SelectedAttributes, without special chars
    # To UBound(TrueImportantProperties)
    # sr = TrueImportantProperties(i)
    # sc = sr
    # shorten:
    # sT = Left(sc, 1)
    if LenB(sT) > 0 Then:
    if sT < "A" Or sT > "z" Then    ' allow ASCII letters only in front:
    # sc = Mid(sc, 2)
    # GoTo shorten

    # Call AppendTo(TrueCritList, _
    # sr, b, _
    # always:=False, ToFront:=False) ' special chars Not removed
    # ' add only if new in critlist, always if unique in TrueItemProperties
    # ' special chars were removed in sr
    # Call AppendTo(SelectedAttributes, sc, b, _
    # always:=False, ToFront:=False) ' special chars are removed
    # sRules.clsObligMatches.ChangeTo = TrueCritList
    # ' use what's clean already and append Similarities
    if LenB(TrueCritList) > 0 Then                              ' some criteria wanted::
    if LenB(sRules.clsSimilarities.CleanMatchesString) > 0 Then:
    # ' include the similarities and default matches (without operators) into SelectedAttributes
    # Call AppendTo(SelectedAttributes, _
    # sRules.clsSimilarities.CleanMatchesString _
    # & b & RemoveChars(aObjDsc.objDftMatches, "*!%-+|:^()"), _
    # b, False, False)
    # While InStr(SelectedAttributes, B2) > 0                     ' reduce double blanks
    # SelectedAttributes = Replace(SelectedAttributes, B2, b)
    # Wend
    # sRules.clsSimilarities.CleanMatches = split(SelectedAttributes)
    # TotalPropertyCount = UBound(sRules.clsSimilarities.CleanMatches) + 1
    # TrueCritList = Trim(TrueCritList)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub startReportToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def startreporttoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.startReportToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim j As Long
    # Dim col As Long
    # Dim aText As String
    # Dim sColList As String
    # Dim skipAttr As Boolean

    # sColList = vbNullString
    if Not xReportExcel Then:
    # GoTo ProcReturn
    if aPindex = 1 Then:
    # cHdl = vbNullString     ' put1intoExcel must match this:
    # cHdl = "Item"                    ' 1
    # With aID(aPindex)
    # j = Min(.idAttrDict.Count, UBound(pArr))
    # col = 2
    # skipAttr = False
    # aText = .idAttrDict.Keys(i)
    if Left(aText, 2) = "==" Then  ' skip seperators ... and DontCompare:
    # skipAttr = True
    else:
    if InStr(sRules.clsNeverCompare.CleanMatchesString, aText) > 0 Then:
    # skipAttr = True
    if InStr(sRules.clsSimilarities.CleanMatchesString, aText) > 0 Then:
    # skipAttr = False
    if InStr(sRules.clsObligMatches.CleanMatchesString, aText) > 0 Then:
    # skipAttr = False
    if skipAttr Then:
    if DebugMode Then:
    # DoVerify False
    else:
    # cHdl = cHdl & b & aText
    # sColList = sColList & b & CStr(col)
    # col = col + 1

    # Set O = xlWBInit(O.xlTWBook, TemplateFile, "Report", _
    # cHdl, showWorkbook:=DebugMode, mustClear:=True)
    # End With ' aID(aPindex)

    # xlApp.ScreenUpdating = False
    # O.xlTSheet.EnableCalculation = False  ' for attribute rules do not calculate
    # xlApp.Cursor = xlWait
    # O.xlTHeadline = cHdl
    # O.xlTHead = split(cHdl)
    # xColList = split(Trim(sColList))
    if Not aID(0) Is Nothing Then:
    # aOD(0).objDumpMade = 0                          ' no dumps made so far

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub endReportToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def endreporttoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.endReportToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' Note: leaving W again (if we can)
    # xlApp.ScreenUpdating = True
    # xlApp.Cursor = xlDefault
    if Not O Is Nothing Then:
    # Set x = O
    if x Is Nothing Then:
    # Set W.xlTSheet = Nothing
    else:
    # Set W.xlTSheet = x.xlTSheet

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ItemReportToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def itemreporttoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.ItemReportToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim aText As String
    # Dim Finished As Boolean
    # Dim maxCols As Long
    # Dim col As Long

    # 'On Error GoTo 0
    # With aID(aPindex).idAttrDict
    # pArr(1) = Line
    # i = LBound(xColList)
    # maxCols = UBound(xColList)
    # maxCols = Min(UBound(pArr), maxCols)
    # Finished = maxCols = 0
    # Do Until Finished  ' get the cAttrDsc items
    if i > maxCols Then:
    # Finished = True
    # GoTo skipitem
    # col = CLng(xColList(i))
    # Set aTD = .Item(col - 1).Item
    # i = i + 1
    # aText = aTD.adFormattedValue
    # pArr(col) = aText
    if DebugLogging Or i Mod 10 = 0 Or col >= maxCols Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
    # & i & b & x.xlTHead(col - 1) & " = " & Quote(aText) _
    # & " into Excel Sheet " & W.xlTName & " column " & col
    # skipitem:
    # Loop
    # Call addLine(x, Line - 1, pArr)
    # End With ' aID(aPindex)
    # Call DisplayExcel(W, EnableEvents:=False, unconditionallyShow:=True)
    # Set aTD = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' derived from: http://www.slipstick.com/developer/instant-search-messages-selected-contact/
# ' Search For Messages From Email or Contact
# ' Creates an instant search for all messages to or from the sender, including messages you sent.
# ' Two versions: 1 - Uses selected Contact. 2 - Uses selected Message
# ' Searches for messages from all three email addresses on a contact, if additional addresses exist.

# '---------------------------------------------------------------------------------------
# ' Method : Sub SearchAssociatedItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def searchassociateditems():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.SearchAssociatedItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # IsEntryPoint = True

    # Call SearchByItem(Nothing)      ' Select item from ActiveExplorer

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SearchByItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def searchbyitem():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.SearchByItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim strFilter As String
    # Dim oFolder As Folder
    # Dim Operator As String
    # Dim noOfEmails As Long
    if oBaseItem Is Nothing Then:
    # Set oBaseItem = ActiveExplorer.Selection.Item(1)

    # ' get result window as (new) ActiveExplorer
    if oBaseItem.Parent Is Nothing Then:
    # Set oFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
    # Set olApp.ActiveExplorer.CurrentFolder = oFolder
    else:
    # Set oFolder = getParentFolder(oBaseItem)
    if oFolder Is Nothing Then:
    # DoVerify False, " item has no parent"
    else:
    # oFolder.Display ' providing ActiveExplorer of correct type

    if IsMailLike(oBaseItem) Then:
    # noOfEmails = noOfEmails + 1
    if InStr(LCase(oFolder.FolderPath), "sen") > 0 Then:
    # strFilter = "an:" & oBaseItem.SenderEmailAddress
    else:
    # strFilter = "von:" & oBaseItem.SenderEmailAddress
    elif oBaseItem.Class = olContact Then:
    if LenB(oBaseItem.Email1Address) > 0 Then:
    # strFilter = Chr(34) & oBaseItem.Email1Address & Chr(34)
    # Operator = " OR "
    # noOfEmails = noOfEmails + 1
    if LenB(oBaseItem.Email2Address) > 0 Then:
    # strFilter = strFilter & Operator & Chr(34) & oBaseItem.Email2Address & Chr(34)
    # Operator = " OR "
    # noOfEmails = noOfEmails + 1

    if LenB(oBaseItem.Email3Address) > 0 Then:
    # strFilter = strFilter & Operator & Chr(34) & oBaseItem.Email3Address & Chr(34)
    # ' Operator = " OR "
    # noOfEmails = noOfEmails + 1
    if SenderOrRecipient = 3 Then       ' both:
    # strFilter = "von:(" & strFilter & ") OR an:(" & strFilter & ")"
    elif SenderOrRecipient = 2 Then   ' only sender:
    # strFilter = "an:" & strFilter
    elif SenderOrRecipient = 1 Then   ' only recipient:
    # strFilter = "von:" & strFilter
    else:
    # DoVerify False, " not implemented"
    if noOfEmails = 0 Then:
    print('Keine Email Adresse vorhanden fr ')
    # GoTo ProcReturn
    # ' filter result window
    # olApp.ActiveExplorer.Search strFilter, olSearchScopeAllFolders

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub reportOnSelectedEmail
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def reportonselectedemail():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.reportOnSelectedEmail"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim msg As String

    # IsEntryPoint = True

    # msg = IsSelectedEmailKnown(3)   ' all 3 are reported

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' get Contacts for email address
def isselectedemailknown():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.IsSelectedEmailKnown"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim strFilter As String
    # Dim oFolder As Folder
    # Dim noOfEmails As Long
    # Dim i As Long
    # Dim oBaseItem As Object
    # Set Contacts = New Collection
    # Set oBaseItem = ActiveExplorer.Selection.Item(1)    ' Expecting mail object
    if Not IsMailLike(oBaseItem) Then:
    # IsSelectedEmailKnown = "*** Das selektierte Item ist keine Mail"
    # GoTo ProcReturn       ' hat also keine
    # Set oFolder = getParentFolder(oBaseItem)
    if oFolder Is Nothing Then:
    # DoVerify False, " item has no parent"
    # strFilter = oBaseItem.SenderEmailAddress

    if IsEmailKnown(strFilter, Contacts) Then:
    # noOfEmails = Contacts.Count
    # IsSelectedEmailKnown = "Es wurden " & noOfEmails _
    # & " Kontakteintrge fr " & strFilter & " gefunden " & vbCrLf
    # IsSelectedEmailKnown = IsSelectedEmailKnown _
    # & "Email " & noOfEmails & " ist Email" & Contacts(i).Alias _
    # & " von " & Quote(Contacts(i).ValueOfItem.Subject)
    if TestOnly < 3 Then    ' Only check for one email:
    # GoTo ProcReturn
    else:
    # IsSelectedEmailKnown = "*** Kein Kontakt vorhanden fr " & strFilter
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsEmailKnown
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isemailknown():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.IsEmailKnown"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # IsEmailKnown = IsEmailInAddressLists(MailAddress, Contacts)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsEmailInAddressLists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isemailinaddresslists():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.IsEmailInAddressLists"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim myaddresslists As AddressLists
    # Dim oneAddressList As AddressList
    # Set myaddresslists = aNameSpace.AddressLists
    for oneaddresslist in myaddresslists:
    if IsEmailInThisAddressList(oneAddressList, MailAddress, Contacts) Then:
    # IsEmailInAddressLists = True
    # Exit For

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsEmailInThisAddressList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isemailinthisaddresslist():
    # Dim zErr As cErr
    # Const zKey As String = "QueryMatches.IsEmailInThisAddressList"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim myAddressEntries As AddressEntries
    # Dim ContactEntry As cNumbItem
    # Dim thisContact As ContactItem
    # Dim thisAddressEntry As AddressEntry
    # Dim i As Long
    if oneAddressList.AddressListType <= olOutlookAddressList Then:
    # Set myAddressEntries = oneAddressList.AddressEntries
    # Set thisAddressEntry = myAddressEntries.Item(i)
    # Set thisContact = thisAddressEntry.GetContact
    if thisContact.Email1Address = MailAddress Then:
    # Set ContactEntry = New cNumbItem
    # ContactEntry.Key = "1"
    # Set ContactEntry.ValueOfItem = thisContact
    # ContactEntry.Alias = MailAddress
    # Contacts.Add ContactEntry
    if thisContact.Email2Address = MailAddress Then:
    # Set ContactEntry = New cNumbItem
    # ContactEntry.Key = "2"
    # ContactEntry.ValueOfItem = thisContact
    # ContactEntry.Alias = MailAddress
    # Contacts.Add ContactEntry
    if thisContact.Email3Address = MailAddress Then:
    # Set ContactEntry = New cNumbItem
    # ContactEntry.Key = "3"
    # ContactEntry.ValueOfItem = thisContact
    # ContactEntry.Alias = MailAddress
    # Contacts.Add ContactEntry
    else:
    # DoVerify False, " not implemented"

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

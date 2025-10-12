# Converted from CompareOps.py

# Attribute VB_Name = "CompareOps"
# Option Explicit

# Dim CompareState As Boolean
# Dim RrX As String, RrY As String
# Dim LocalPropertyCount As Long
# Dim undecodedPropertyCount As Long
# Dim MatchIndicator As String
# Dim currPropMisMatch As Boolean
# Dim MisMatchIgnore As Boolean
# Dim UserAnswer As String
# Dim xQuickchecksonly As Boolean
# Dim xTDeferExcel As Boolean
# Dim xTuseExcel As Boolean
# Dim doingTheRest As Boolean
# Dim AbortSignal As Long
# Dim sTime As Variant

# '---------------------------------------------------------------------------------------
# ' Method : Sub C1SI
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def c1si():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "CompareOps.C1SI"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")

    # xUseExcel = False
    # xDeferExcel = True
    # displayInExcel = True
    # SelectOnlyOne = True
    # Call SelectAndCompare
    # StopRecursionNonLogged = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub C2SI
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def c2si():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "CompareOps.C2SI"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")

    # IsEntryPoint = True

    # xUseExcel = False
    # xDeferExcel = False
    # SelectOnlyOne = False
    # FindMatchingItems = True                       ' do not dump item model attributes
    # UI_DontUse_Sel = True
    # Call SelectAndCompare
    # StopRecursionNonLogged = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Compare2SelectedItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def compare2selecteditems():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.Compare2SelectedItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # Dim ExcelDefer As Boolean
    # ' Process Folder items
    # Restart:
    # AllProps = False
    # AllPropsDecoded = InitsForDecoding(forItem1, forItem2)
    # ' Use main Object Identification as used by sort!
    # Matchcode = GetObjectAttributes(True, True)

    # Message = vbCrLf & "Prfung von " & Folder(1).FolderPath & ": " & WorkIndex(1)
    if Not forItem1 And forItem2 Then:
    # fiMain(2) = fiMain(1)
    else:
    # Message = Message & " und " & Folder(2).FolderPath & ": " & WorkIndex(2)
    # Message = Message & ", " & MainObjectIdentification & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1)
    if fiMain(1) <> fiMain(2) Then:
    # Message = Message & vbCrLf & WorkIndex(2) & ": " & fiMain(2)
    # Call LogEvent(String(60, "_") & Message)

    # ExcelDefer = xDeferExcel                       ' save here cause localy changed
    if Matchcode = Passt_Synch Then:
    if ItemIdentity(AllPropsDecoded) Then:
    # ' ==================================================================
    # ListContent(ListCount).Compares = "="
    if UserDecisionRequest And Not UserDecisionEffective And (xDeferExcel Or Not xUseExcel) Then:
    else:
    # rsp = vbNo
    else:
    # ListContent(ListCount).Compares = "<>"
    # ListContent(ListCount).MatchData = MatchData
    if IsComparemode And rsp <> vbCancel Then:
    if rsp = vbOK Then:
    if xlA Is Nothing Then:
    # Call ShowDetails
    if UserDecisionRequest Then        ' Show decode and Show in excel:
    # xDeferExcel = False
    # displayInExcel = True
    # quickChecksOnly = False
    # OnlyMostImportantProperties = False
    # IsComparemode = True
    # ' ?                    doingTheRest = True
    # GoTo Restart
    else:
    # Call DisplayExcel(O)
    # Set ListContent = Nothing
    # Set sortedItems(1) = Nothing
    # Set sortedItems(2) = Nothing
    # xDeferExcel = ExcelDefer                       ' restore detailed analysis

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CompareAllStackedAttributes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def compareallstackedattributes():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.CompareAllStackedAttributes"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # '       Textual results of Comparison also generated here.
    # '       note that aDecProp C may be filled up to MostRelevant only in some cases
    if Not doingTheRest Then:
    # relevantIndex = 1                          ' start here
    else:
    if DebugMode And relevantIndex <> 1 Then   ' interesting ???:
    # DoVerify False
    # LocalPropertyCount = aID(1).idAttrDict.Count - 1
    # stpcnt = aID(2).idAttrDict.Count - 1
    if relevantIndex = 1 Then:
    # Call initializeComparison
    # undecodedPropertyCount = 0
    # CompareState = True
    # LocalPropertyCount = MaxPropertyCount
    else:
    # LocalPropertyCount = Max(LocalPropertyCount, stpcnt)
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab & "comparing first property")

    # OneDiff = vbNullString
    # OneDiff_qualifier = vbNullString
    # currPropMisMatch = False
    # MisMatchIgnore = False
    # ' get display values into Rrx, RrY
    if synchedNames(RrX, RrY) Then             ' synchs attribute names and more...:
    # AbortSignal = 1                        ' did not work: no way to success here
    if DebugMode Then DoVerify False:
    # GoTo ProcReturn

    # ' done in synchedNames: Set aTd = GetAttrDsc(PropertyNameX)
    # '                                 also sets correct iRules
    # pArr(1) = PropertyNameX
    # pArr(2) = RrX
    # pArr(3) = RrY
    if PropertyNameX = "HTMLBody" Then DoVerify False:

    if InStr(PropertyNameX, "EntryID") > 0 Then ' should NOT match:
    if RrX = RrY _:
    # And LenB(RrX) > 0 Then              ' Vergleich nicht sinnvoll fr EntryID
    # Message = PropertyNameX & " identisch, vermutlich gleiches Objekt"
    if PropertyNameX = "EntryID" Then:
    # Message = Message & vbCrLf _
    # & "<> Vergleichen der Objekte gibt keinen Sinn: " _
    # & "beendet wegen Fehler: " _
    # & vbCrLf & Err.Description
    # CompareState = False
    # Call logMatchInfo
    # DoItAnyway:
    if aTD Is Nothing Then:
    if aDecProp(1) Is Nothing Or aDecProp(2) Is Nothing Then:
    if DebugMode Then DoVerify False:
    # GoTo skipCompare
    # Set iRules = sRules
    if DebugMode Then DoVerify False:

    if iRules.clsObligMatches.RuleMatches _:
    # Or (aDecProp(1).adOrigValDecodingOK _
    # And aDecProp(2).adOrigValDecodingOK) Then

    if Left(RrX, 1) = "#" Then             ' like unequal, but always same:
    if iRules.clsNeverCompare.RuleMatches Then:
    # IgString = " #. "
    # GoTo iggit
    else:
    # GoTo genMessage
    if RrX <> RrY Then:
    # currPropMisMatch = True
    # IgString = vbNullString
    # MatchIndicator = "<>  "
    # AttributeIndex = aTD.adtrueIndex
    # logDecodedProperty aDecProp(1).adShowValue, MatchIndicator
    # logDecodedProperty aDecProp(2).adShowValue, MatchIndicator

    if ShowEmptyAttributes _:
    # Or LenB(RrX) > 0 _
    # Or LenB(RrY) > 0 Then
    if LenB(killStringMsg) > 0 Then:
    # MatchData = MatchData & vbCrLf & killStringMsg
    # genMessage:
    if Not iRules.RuleIsSpecific Then  ' just a mismatch:
    # MatchIndicator = "...  "
    # GoTo iggit

    if iRules.clsSimilarities.RuleMatches Then:
    # MatchIndicator = "~~  "

    # UserDecisionRequest = UserDecisionRequest _
    # Or AcceptCloseMatches
    # OneDiff = "~~~ Bedingte Abweichung bei Eigenschaft " _
    # & iRuleBits & vbCrLf
    # MisMatchIgnore = True
    # SimilarityCount = SimilarityCount + 1
    # GoTo Logshort
    elif iRules.clsNeverCompare.RuleMatches Then:
    # MatchIndicator = "__  "
    # iggit:
    # IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
    if iRules.clsNotDecodable.RuleMatches Then:
    # undecodedPropertyCount = undecodedPropertyCount + 1
    # OneDiff = " kein Vergleichsergebnis, nicht dekodiert " _
    # & iRuleBits & vbCrLf
    else:
    # OneDiff = "... Abweichung ignoriert fr " _
    # & iRuleBits & vbCrLf
    # MisMatchIgnore = True
    # GoTo Logshort
    elif iRules.clsObligMatches.RuleMatches Then:
    # CompareState = False
    if Left(RrX, 1) = "#" Then:
    # OneDiff = "undefinierte Abweichung bei Eigenschaft " _
    # & iRuleBits & vbCrLf
    else:
    # cMisMatchesFound = cMisMatchesFound + 1
    # OneDiff = cMisMatchesFound _
    # & ". relevante Abweichung bei Eigenschaft " _
    # & iRuleBits & vbCrLf
    # SuperRelevantMisMatch = True
    # MisMatchIgnore = False
    # doLog:
    # DiffsRecognized = DiffsRecognized & vbCrLf & "#### " _
    # & OneDiff
    # GoTo setDiff
    # Logshort:
    # DiffsRecognized = DiffsRecognized & vbCrLf & OneDiff & _
    # vbCrLf & aDecProp(2).adKillMsg
    # setDiff:
    # OneDiff_qualifier = "    " _
    # & WorkIndex(1) & ": " & aDecProp(1).adShowValue _
    # & vbCrLf & "    " _
    # & WorkIndex(2) & ": " & aDecProp(2).adShowValue
    else:
    if iRules.clsSimilarities.RuleMatches Then:
    # OneDiff = OneDiff & "~~~ Abweichung, hnlichkeit prfen " _
    # & iRuleBits & vbCrLf
    elif iRules.clsNeverCompare.RuleMatches Then:
    # OneDiff = OneDiff & "??? von Vergleich ausgeschlossen " _
    # & iRuleBits & vbCrLf
    elif iRules.clsNotDecodable.RuleMatches Then:
    # OneDiff = OneDiff & "nicht dekodiert " _
    # & iRuleBits & vbCrLf
    else:
    # OneDiff = OneDiff & " Wertabweichung bei " _
    # & iRuleBits & vbCrLf
    # cMisMatchesFound = cMisMatchesFound + 1
    # MisMatchIgnore = False
    # CompareState = False
    # MatchIndicator = "##  "
    # DiffsRecognized = DiffsRecognized & vbCrLf & "#### " _
    # & OneDiff
    # GoTo Logshort
    # ' ignored misMatch for whatever reason
    # MisMatchIgnore = True
    # MatchIndicator = "--  "
    # IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
    # GoTo Logshort                  ' what shall we do with onediff???
    # Call logDiffInfo(MatchIndicator & OneDiff & OneDiff_qualifier)
    if MisMatchIgnore Then:
    # DiffsIgnored = DiffsIgnored & vbCrLf _
    # & OneDiff & OneDiff_qualifier
    else:
    # DiffsRecognized = DiffsRecognized & vbCrLf _
    # & OneDiff_qualifier
    else:
    # MatchIndicator = "    "
    # Matches = Matches + 1
    if (ShowEmptyAttributes Or LenB(RrX) > 0) _:
    # And MinimalLogging < eLmin Then
    # MatchData = MatchData & vbCrLf _
    # & MatchIndicator & WorkIndex(1) & b _
    # & PropertyNameX & "=" & Quote(Left(RrX, 80))
    else:
    if RrX <> RrY Then:
    # currPropMisMatch = True
    # MatchIndicator = "..  "
    # OneDiff = "... Abweichung nicht gewertet zu *" & iRuleBits & vbCrLf
    # MisMatchIgnore = True
    # IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
    # GoTo Logshort

    if LenB(OneDiff) > 0 Then:
    # AllItemDiffs = AllItemDiffs & OneDiff & OneDiff_qualifier & vbCrLf
    if displayInExcel Then                 ' And (xDeferExcel Or xUseExcel) ???:
    if O Is Nothing Then               ' not defined in excel ==> add at end:
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
    # sHdl, False, DebugMode) ' no: put it there
    # pArr(4) = MatchIndicator
    # pArr(5) = Trunc(1, OneDiff, vbCrLf)
    # Call addLine(O, AttributeIndex, pArr)
    if Not aCell Is Nothing Then:
    # Set aCell = O.xlTSheet.Cells(AttributeIndex + 1, ValidCol)
    if DebugMode Or DebugLogging Then:
    # aCell.Select
    if MisMatchIgnore Then:
    # aCell.Interior.pattern = xlSolid
    # aCell.Interior.PatternColorIndex = xlAutomatic
    # aCell.Interior.ThemeColor = xlThemeColorAccent6
    else:
    # aCell.Font.Color = -16776961 ' ROT
    # O.xlTSheet.Cells(AttributeIndex + 1, 1).Font.Color = -16776961

    if currPropMisMatch And Not MisMatchIgnore Then:
    if SuperRelevantMisMatch Then          ' some prop misMatch that counts:
    if AcceptCloseMatches Then         ' check approximate Matches:
    if cMisMatchesFound > MaxMisMatchesForCandidates _:
    # And quickChecksOnly Then
    # Exit For                   ' no point to compare more
    else:
    # Exit For

    if AttributeIndex Mod 10 = 0 Then:
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") _)
    # & vbTab & "comparing property # " & AttributeIndex
    # skipCompare:

    # ' hurra, alle vergleiche sind gemacht. Zeig das Ergebnis.
    print(Debug.Print Format(Timer, "0#####.00") & vbTab & " Match=" & CompareState _)
    # & vbTab & "compared last property, # " & AttributeIndex - 1
    # AllPropsDecoded = (MaxPropertyCount >= TotalPropertyCount)

    # Call GenerateSummary

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CreateStatisticOutput
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createstatisticoutput():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.CreateStatisticOutput"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Message = "    " & Matches & " bereinstimmungen / " _
    # & cMisMatchesFound & " relevante Abweichungen" _
    # & " / " & SimilarityCount & " hnlichkeiten"

    # Message = Message & vbCrLf & "    es wurden " _
    # & IgnoredPropertyComparisons _
    # & " Eigenschaften ignoriert"

    if AttributeUndef(1) > 0 Then:
    # displayInExcel = True
    # Message = Message & vbCrLf & "  ## Attribut " _
    # & AttributeUndef(1) + 1 _
    # & " inkonsistent in item " & WorkIndex(1)
    if AttributeUndef(2) > 0 Then:
    # displayInExcel = True
    # Message = Message & vbCrLf & "  ## Attribut " _
    # & AttributeUndef(2) + 1 _
    # & " inkonsistent in item " & WorkIndex(2)

    if saveItemNotAllowed Then:
    # Message = Message & vbCrLf & "    " & YleadsXby _
    # & " / " & NotDecodedProperties _
    # & " Eigenschaften sind nicht in beiden Items enthalten" _
    # & " / nicht decodiert"
    # Call logMatchInfo
    # CreateStatisticOutput = Message

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayWithExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def displaywithexcel():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.DisplayWithExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if displayInExcel Then:
    # ' xUseExcel = True obsolete ***???***
    if O Is Nothing Then                       ' defined in excel ?:
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
    # sHdl, DebugMode)      ' no: put it there
    # Call StckedAttrs2Xcel(O)
    if WorkIndex(1) > 0 And LenB(Statistics) = 0 Then:
    # Call CompareAllStackedAttributes       ' not necessarily all that exist in Items
    # Statistics = CreateStatisticOutput

    if LenB(Statistics) > 0 Then:
    # PutStatisticOutputToExcel (Statistics)
    # Call ExcelEditSession(0)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayWithoutExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def displaywithoutexcel():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.DisplayWithoutExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim textmod As String
    if Not xlApp Is Nothing Then:
    if Not (displayInExcel Or O Is Nothing) Then:
    if Not UserDecisionRequest Then:
    if Not AllPropsDecoded Then:
    # textmod = "(Partiellen) "
    if vbYes = MsgBox(textmod _:
    # & "Vergleich in Excel anzeigen?", vbYesNo) Then
    # displayInExcel = True
    else:
    # displayInExcel = False
    if O Is Nothing Then:
    # UserDecisionRequest = True
    else:
    if UserDecisionEffective Then:
    # GoTo ProcReturn
    # Set O = Nothing            ' assume No as an answer next time

    if Not displayInExcel Then:
    # showResults:
    if UserDecisionRequest Or Not (xDeferExcel Or xUseExcel) Then:
    # CompareState = AskUserAndInterpretAnswer(oMessage)
    else:
    # displayInExcel = True                  ' displayWithExcel always called AFTER this !!!
    # mustDecodeRest = Not AllPropsDecoded   ' if still incomplete, do rest now
    # CheckWithUser:
    if rsp = vbRetry Then                      ' we ALWAYS need full comparison now!:
    # UserDecisionEffective = False
    # displayInExcel = True
    # mustDecodeRest = Not AllPropsDecoded   ' if still incomplete, do rest now
    # quickChecksOnly = False                ' do a complete decode; value is restored on exit
    if displayInExcel And xlApp Is Nothing Then:
    # Call XlgetApp

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' Evaluation ExplainS
# ' 0: PerformChangeOpsForMapiItems
# ' 1: ModRuleTab
def exceleditsession():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.ExcelEditSession"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # rsp = vbNo
    # Call Try(allowAll)                                ' Try anything, autocatch, Err.Clear
    # xlApp.EnableEvents = False
    # Catch
    # With x.xlTSheet
    # .Cells(1, clickColumn) = "Hier Klicken:"
    # .Cells(1, moreColumn) = vbNullString
    # .Cells(1, clickColumn).Interior.ColorIndex = 35 ' hell GRN
    # .Cells(1, moreColumn).Interior.ColorIndex = 35
    # .Cells(1, promptColumn).Interior.ColorIndex = 0
    # .Cells(1, promptColumn).Font.Color = -16776961 ' ROT
    # .Cells(1, promptColumn) = "Return -> Outlook"
    # .Cells(1, changeCounter) = 0
    # Call DisplayExcel(x, EnableEvents:=False, unconditionallyShow:=True)

    # xlApp.EnableEvents = False
    # Catch
    # .Cells(1, WatchingChanges) = True          ' set for editing
    # .Cells(1, WatchingChanges).Interior.ColorIndex = 35
    if .Cells(1, flagColumn) = "Flag" Then:
    # .Cells(1, flagColumn).Interior.ColorIndex = 35
    # .Activate
    # Catch

    # xlApp.EnableEvents = True
    # sTime = 0
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
    # & "Beginning Excel Edit Session"

    # Call DisplayWindowInFront(xLMainWindowHdl, 1)
    # xlApp.Cursor = xlDefault
    # stillEditing:
    if xlApp.ActiveWindow Is Nothing Then:
    # Call ClearWorkSheet(xlA, x)

    # rsp = vbYes                            ' user closed the window, so it is like edit aborted
    if AllPropsDecoded Then:
    # Call LogEvent("Excel Window wurde geschlossen, nderungen " _
    # & "werden nicht aus Excel bernommen.", eLall)
    # rsp = vbCancel
    else:
    # Call LogEvent("Excel geschlossen (kein Edit), " _
    # & "es erfolgt eine neue, vollstndige Darstellung aller Attribute", eLall)
    # rsp = vbRetry                      ' do a retry with all
    # GoTo canMod
    elif Not xlApp.Visible Then:
    # xlApp.EnableEvents = False
    # GoTo showResults
    else:
    # showResults:
    # xlApp.Cursor = xlDefault
    # Catch

    if DebugMode Then:
    if MsgBox("waiting here for end of edit in Excel, Click into cell(1, " _:
    # & clickColumn & ")", vbOKCancel) = vbCancel Then
    # DoVerify False
    # excelFinished:
    if xlApp.ActiveSheet Is Nothing Then:
    print(Debug.Print "user has closed excel sheet or application")
    # Set x = Nothing
    # Set x = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
    # sHdl, showWorkbook:=DebugMode) ' no: put it there
    # rsp = vbRetry
    # quickChecksOnly = False
    # mustDecodeRest = True
    # GoTo canMod
    if InStr(.Cells(1, promptColumn).Text, "Go Outlook") > 0 Then:
    # displayInExcel = False
    # xlApp.Visible = False
    # xlApp.EnableEvents = False
    # x.xlTSheet.EnableCalculation = False
    # x.xlTSheet.EnableFormatConditionsCalculation = False
    # Catch
    # Call LogEvent(Format(Timer, "0#####.00") & " Edit in Excel fertig", eLall)
    # Call DisplayWindowInFront(FWP_xLW_Hdl, 2)
    # rsp = vbYes                        ' Excel Return is like Answer = yes
    match EvaluationMode:
        case 0::
    # Call PerformChangeOpsForMapiItems
        case 1::
    # Call ModRuleTab
    # GoTo ProcReturn
        case _:
    # DoVerify False, " not imp"
    if Folder(2) Is Nothing Then:
    if Folder(1).Parent Is Nothing Then:
    # Message = vbNullString
    else:
    # Message = "Item ist in " & Folder(1).FolderPath
    elif Folder(1).Parent Is Nothing Or Folder(2).Parent Is Nothing Then:
    # GoTo setFmsg
    elif Folder(1).Parent.Name = Folder(2).Parent.Name Then:
    # Message = vbNullString
    else:
    # setFmsg:
    # Message = "Item 2 ist in " & Folder(2).FolderPath
    else:
    if Wait(0.5, sTime, "Warte auf Ende in Excel. Klicken Sie auf 'Hier Klicken' ") Then:
    # DoVerify False
    # GoTo excelFinished
    # Call ShowStatusUpdate
    # GoTo stillEditing
    # canMod:
    if rsp = vbYes Then:
    # Call SaveItemsIfChanged(True)
    # AllPropsDecoded = True                 ' do not get more attributes
    elif rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    elif rsp = vbNo Then:
    # GoTo showResults
    # '               vbRetry is also possible, user closed excel ==> decode all
    # End With                                       ' X.xlTSheet

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ExcelShowItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def excelshowitem():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "CompareOps.ExcelShowItem"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")

    # IsEntryPoint = True

    # ActionTitle(0) = "Attribute in Excel zeigen fr selektiertes Element in"
    # AttributeUndef(1) = 0
    # AttributeUndef(2) = 0
    # aPindex = 1

    # xUseExcel = True
    # xDeferExcel = False
    # displayInExcel = True
    # SelectOnlyOne = True
    # IsComparemode = True
    # eOnlySelectedFolder = False
    # Set LF_CurLoopFld = Nothing
    # UI_DontUse_Sel = True                          ' no special selection/filter parameters to be used
    # UI_DontUseDel = True                           ' no deletion rules to be used
    # E_Active.EventBlock = True
    # Call SelectAndCompare                          ' converts and displays in Excel sheet O (Objekteigenschaften)
    # E_Active.EventBlock = False
    if T_DC.TermRQ Then:
    # Call TerminateRun
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GenerateSummary
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def generatesummary():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.GenerateSummary"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if quickChecksOnly Then                        ' !! idObjItem is used to access oTi !!:
    # NotDecodedProperties = _
    # Max(aID(1).idObjItem.ItemProperties.Count, _
    # aID(2).idObjItem.ItemProperties.Count) _
    # - LocalPropertyCount
    else:
    # NotDecodedProperties = LocalPropertyCount - (AttributeIndex - 1)
    # NotDecodedProperties = NotDecodedProperties + undecodedPropertyCount

    if quickChecksOnly Or Not (MisMatchIgnore Or IsComparemode) Then:
    if CompareState Then:
    if Not quickChecksOnly Then:
    # ' we must continue at least until the cmisMatchesfound-limit is exceeded
    # GoTo keepChecking

    if NotDecodedProperties > 0 Then:
    # saveItemNotAllowed = True
    if Not quickChecksOnly _:
    # Or mustDecodeRest Then
    # ' we must continue at least until the cmisMatchesfound-limit is exceeded
    # GoTo keepChecking
    if displayInExcel Then:
    # pArr(1) = "*** es wurden nicht alle Merkmale verglichen"
    # Call addLine(O, aID(2).idAttrDict.Count + 2, pArr)
    # AllItemDiffs = AllItemDiffs & vbCrLf & pArr(1)
    # IgnoredPropertyComparisons = IgnoredPropertyComparisons _
    # + NotDecodedProperties
    else:
    # keepChecking:
    # doingTheRest = True


    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GetCompareItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getcompareitems():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.GetCompareItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # rsp = 0                                        ' not a valid value at all!!
    # LocalPropertyCount = 0
    # currPropMisMatch = False
    # YleadsXby = 0

    # sHdl = "CritPropName---------------" _
    # & " Objekt-" & WorkIndex(1) & "----------------------" _
    # & " Objekt-" & WorkIndex(2) & "----------------------" _
    # & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    # Call InitsForPropertyDecoding(doingTheRest)

    if mustDecodeRest And Not quickChecksOnly Then:
    # Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
    # ContinueAfterMostRelevant:=mustDecodeRest, _
    # onlyItemNo:=3)
    else:
    if isEmpty(MostImportantProperties) Then:
    # GoTo allOfThem
    # Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
    # ContinueAfterMostRelevant:=mustDecodeRest, _
    # onlyItemNo:=1)
    # allOfThem:
    if displayInExcel _:
    # And Not xlApp _
    # And Not O Is Nothing Then               ' unless we defer
    # Call StckedAttrs2Xcel(O)
    # Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
    # ContinueAfterMostRelevant:=mustDecodeRest, _
    # onlyItemNo:=2)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetObjectAttributes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getobjectattributes():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.GetObjectAttributes"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim MainCompare As Long

    # aCloneMode = withNewValues

    if forItem1 Then:
    if SelectOnlyOne And forItem2 Then:
    # DoVerify False, " inconsistent ???"
    # Call GetAobj(1, WorkIndex(1))
    # objTypName = DecodeObjectClass(getValues:=True)
    # DoVerify WorkIndex(1) = aItmIndex
    # Call AddItemToList((0), fiMain(1), (Passt_Deleted), vbNullString)
    if forItem2 Then:
    if SelectOnlyOne And forItem1 Then:
    # DoVerify False, "SelectOnlyOne inconsistent with: decode for item 1 when decoding item 2"
    # Call GetAobj(2, WorkIndex(2))
    # objTypName = DecodeObjectClass(getValues:=True)
    if objTypName = "not defined" Then:
    # WorkIndex(2) = 0
    # GoTo ProcReturn                        ' value 0
    # Call AddItemToList(vbNullString, fiMain(2), (Passt_Inserted), (0))

    if SelectOnlyOne Then:
    # MainCompare = True
    else:
    # MainCompare = StrComp(fiMain(1), fiMain(2), vbTextCompare)
    # GetObjectAttributes = Passt_Synch

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub IdentityCheck
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def identitycheck():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.IdentityCheck"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Statistics As String
    # Dim DontSkipActions As Boolean
    # Call GetCompareItems
    # ' FindMatchingItems = true ==> selected main properties by search already
    # CompareState = (fiMain(1) = fiMain(2)) Or FindMatchingItems
    # UserDecisionEffective = False
    if quickChecksOnly Then:
    if Not CompareState Then:
    # GoTo ProcReturn                        ' primary ids are different, we don't care about other details

    if Not AllPropsDecoded Then:
    if Not quickChecksOnly Then:
    # quickChecksOnly = False
    # Retry:
    if AllPropsDecoded Then:
    # Call initializeComparison          ' restart here *** for manual debug
    # Call initializeExcel
    # AllPropsDecoded = False                ' rules could have changed
    # mustDecodeRest = True
    # aPindex = 1
    if aOD(0).objMaxAttrCount <> aID(1).idAttrDict.Count - 1 Then:
    # aOD(0).objDumpMade = -1            ' dump AttributeDescriptors again (seldom)
    # aOD(1).objDumpMade = -1
    # aOD(2).objDumpMade = -1
    # Call GetCompareItems

    # Call CompareAllStackedAttributes               ' not necessarily all that exist in Items

    # DontSkipActions = True
    if CompareState Then                           ' we found identical one:
    match ActionID:
        case atOrdnerinhalteZusammenfhren:
    # DontSkipActions = False
    if DebugMode Or DebugLogging Then:
    print(Debug.Print "Fast Check completed, items Match with " _)
    # & Matches & " Attributes, no relevant misMatches"
        case _:
    if DontSkipActions And Not UserDecisionEffective And Not xReportExcel Then:
    # Statistics = CreateStatisticOutput
    # Call DisplayWithoutExcel(Statistics)
    if rsp = vbRetry Then:
    # GoTo Retry
    # Call DisplayWithExcel(Statistics)
    if rsp = vbRetry Then:
    # GoTo Retry
    # Call SaveItemsIfChanged                    ' sets SaveItemRequested if we saved 1 or 2
    # Call QueryAboutDelete(Statistics)
    if rsp = vbRetry Then:
    # GoTo Retry

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitObjectSelection
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initobjectselection():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.InitObjectSelection"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # WorkIndex(1) = 1
    if aPindex > 0 Then                            ' there is no (need no) aID(0).idAttrDict:
    # ' set dynamic headline
    if SelectOnlyOne Then:
    # sHdl = "CritPropName---------------" _
    # & " Objekt-----------------------" _
    # & " -----------------------------" _
    # & " Comp-- Info-- Flag-- ign.parts1"
    # WorkIndex(2) = inv                     ' undef
    else:
    # sHdl = "CritPropName---------------" _
    # & " Objekt-" & WorkIndex(1) & "----------------------" _
    # & " Objekt-" & WorkIndex(2) & "----------------------" _
    # & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    # WorkIndex(2) = 2

    # ' if we compare, close matches are ok if no quickchecks
    # AcceptCloseMatches = IsComparemode And Not quickChecksOnly
    # OnlyMostImportantProperties = quickChecksOnly  ' decode OnlyMostImportantProperties
    # MinimalLogging = 3
    # WantConfirmation = True
    # MatchMin = 1000
    # MaxMisMatchesForCandidates = MaxMisMatchesForCandidatesDefault
    # ListCount = 0
    # MaxPropertyCount = 0
    # Set sortedItems(1) = Nothing
    # Set sortedItems(2) = Nothing
    # Set Folder(2) = ChosenTargetFolder
    if LF_UsrRqAtionId <> atBearbeiteAllebereinstimmungenzueinerSuche Then:
    # Set SelectedItems = New Collection
    # Set Folder(1) = Nothing
    # Set ChosenTargetFolder = Nothing           ' default: let user find Folder next time
    # Set ListContent = Nothing

    # AllDetails = vbNullString
    # eOnlySelectedItems = True

    # bDefaultButton = "No"
    # rsp = vbOK

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function InitsForDecoding
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initsfordecoding():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.InitsForDecoding"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # DeletedItem = inv
    # DeleteIndex = -1
    if Not forItem1 And forItem2 Then:
    # InitsForDecoding = True
    else:
    # Set LF_CurLoopFld = Folder(1)
    # Call BestObjProps(LF_CurLoopFld, withValues:=True)
    # Call Initialize_UI
    if StopLoop Then:
    # GoTo ProcReturn
    # InitsForDecoding = False

    # Call FindTopFolder(LF_CurLoopFld)

    if xUseExcel Then:
    if xlApp Is Nothing Then:
    # Call XlgetApp
    # Call XlopenObjAttrSheet(xlA)
    elif O Is Nothing Then:
    # Call XlopenObjAttrSheet(xlA)
    else:
    if O.xlTSheet.Name <> cOE_SheetName Then:
    # Call XlopenObjAttrSheet(xlA)

    # FindTrashFolder

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ItemIdentity
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def itemidentity():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.ItemIdentity"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # AbortSignal = 0                                ' all is ok so far
    # AllPropsDecoded = DecodingFinished
    # xQuickchecksonly = quickChecksOnly             ' keep for restore at end of Function
    # xTDeferExcel = xDeferExcel
    # xTuseExcel = xUseExcel
    # ' ActionID = 0, 5, flagcolumn, normally lead to editable excel display
    # displayInExcel = (xUseExcel Or xDeferExcel) _
    # And (ActionID = 0 _
    # Or ActionID = atNormalreprsentationerzwingen _
    # Or ActionID = atOrdnerinhalteZusammenfhren)

    if Not quickChecksOnly Then                    ' mostimportant ones Match already:
    # mustDecodeRest = True                      ' we already did those, now the rest

    # Call IdentityCheck
    if Not (DebugMode Or DebugLogging) Then:
    if Not xlApp Is Nothing Then:
    if Not CheckExcelOK Then:
    # Call xlEndApp
    if Not xlC Is Nothing Then:
    # Call ClearWorkSheet(xlC, O)        ' it is NOT closed here!
    # rsp = vbCancel                             ' do not redisplay user questions

    # quickChecksOnly = xQuickchecksonly
    # xDeferExcel = xTDeferExcel
    # xUseExcel = xTuseExcel
    # ItemIdentity = CompareState

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub multiLinesToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def multilinestoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.multiLinesToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim MultiLine As Variant
    # Dim MsgLine As Variant
    # MultiLine = split(Message, vbCrLf)
    for msgline in multiline:
    # pArr(col) = MsgLine
    # LineNo = LineNo + 1
    # Call addLine(xw, LineNo, pArr)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PutStatisticOutputToExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def putstatisticoutputtoexcel():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.PutStatisticOutputToExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if Not displayInExcel Then:
    # GoTo ProcReturn
    # stpcnt = LocalPropertyCount + 1                ' excel Row number for messages
    if Not O Is Nothing Then:
    # Call multiLinesToExcel(Message, O, LineNo:=stpcnt, col:=1)
    if LenB(DiffsRecognized) > 0 Then:
    # Call multiLinesToExcel(DiffsRecognized, O, LineNo:=stpcnt, col:=1)
    if LenB(DiffsIgnored) > 0 Then:
    # Call multiLinesToExcel(DiffsIgnored, O, LineNo:=stpcnt, col:=1)
    # W.xlTSheet.Activate
    # W.xlTSheet.Cells(LocalPropertyCount + 3, 1).Select

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub QueryAboutDelete
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def queryaboutdelete():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.QueryAboutDelete"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' find out if we must query about deletes and if its reasonable
    if UserDecisionEffective Then:
    # GoTo ProcReturn                            ' we already asked user
    if ActionID = 0 _:
    # Or cMisMatchesFound + SimilarityCount < MaxMisMatchesForCandidates Then
    if IsComparemode _:
    # And NotDecodedProperties > 0 _
    # And Not UserDecisionRequest _
    # And Not SuperRelevantMisMatch Then      ' Automatic decision or ...
    if NotDecodedProperties > 0 Then:
    # UserAnswer = Quote(fiMain(1)) _
    # & "  sollte bei unvollstndigem Vergleich nicht gelscht werden " _
    # & vbCrLf & vbCrLf & AllItemDiffs
    if AllPropsDecoded Then:
    # CompareState = AskUserAndInterpretAnswer(oMessage)
    elif displayInExcel Then:
    # Call DisplayExcel(O, relevant_only:=True, _
    # unconditionallyShow:=True)
    # Call PerformChangeOpsForMapiItems
    else:
    # CompareState = AskUserAndInterpretAnswer(oMessage)
    # CompareState = (cMisMatchesFound = 0) _
    # Or ((cMisMatchesFound <= MaxMisMatchesForCandidates) _
    # And AcceptCloseMatches)
    else:
    if ActionID <> atOrdnerinhalteZusammenfhren Then:
    # ' merge Folders without delete (normally)
    # CompareState = AskUserAndInterpretAnswer(oMessage)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SelectAndCompare
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def selectandcompare():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.SelectAndCompare"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim action As String
    # Dim SelCount As Long

    if SelectedItems Is Nothing Then:
    # GoTo prep
    if Not (eOnlySelectedItems And SelectedItems.Count > 0) Then:
    # prep:
    # Call InitObjectSelection
    # SelCount = ActiveExplorer.Selection.Count
    else:
    # SelCount = SelectedItems.Count
    # GoTo matchGet
    if SelCount >= 1 Then:
    # Set ActiveExplorerItem(1) = GetDItem_P()
    if SelectMulti Then:
    # GoTo gothemAll

    if SelCount > 1 Then                           ' do we have enough to work on?:
    if SelectOnlyOne Then:
    # ActiveExplorer.ClearSelection          ' too much
    # GoTo ask1again
    else:
    # ' ====== AllPublic Global value used, set VVVVVVVVVVVVV before call
    if SelCount = 0 _:
    # Or (SelCount < 2 And Not SelectOnlyOne) _
    # Or (SelectOnlyOne And SelCount <> 1) Then
    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask1again:
    if SelectOnlyOne Then:
    # & action, _
    # "OK", "Cancel", "Auswahl Item fr hnlichkeitssuche")
    else:
    # & " oder zwei zu vergleichende Objekte" _
    # & action, _
    # "OK", "Cancel", "Auswahl von Objekten")

    match rsp:
        case vbOK:
    if ActiveExplorer.Selection.Count = 1 Then:
    # Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
    if SelectOnlyOne Then:
    # GoTo gothemAll
    elif ActiveExplorer.Selection.Count = 2 Then:
    # Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
    # Set ActiveExplorerItem(2) = ActiveExplorer.Selection.Item(2)
    # GoTo gothemAll
    else:
    # action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
    # GoTo ask1again
        case vbCancel:
    # Call LogEvent("=======> Stopped before processing any Folders . Time: " _
    # & Now())
    if TerminateRun Then:
    # GoTo ProcReturn
    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask2again:
    # & action, _
    # "OK", "Cancel", "Auswahl der zu vergleichenden Items")

    match rsp:
        case vbOK:
    if ActiveExplorer.Selection.Count = 1 Then:
    # Set ActiveExplorerItem(2) = ActiveExplorer.Selection.Item(1)
    else:
    # action = " (nur ein Item, bitte)"
    # GoTo ask2again
        case vbCancel:
    # Call LogEvent("=======> Stopped before processing any Folders . Time: " _
    # & Now())
    if TerminateRun Then:
    # GoTo ProcReturn

    # gothemAll:
    # Set SelectedObjects = ActiveExplorer.Selection
    # Call GetSelectedItems(ActiveExplorerItem)      ' Selection -> SelectedItems
    # matchGet:
    if Not sRules Is Nothing Then:
    if sRules.RuleObjDsc Is Nothing Then:
    # GoTo NoObjDsc
    if SelectedItems.Item(1).Class <> sRules.RuleObjDsc.objItemClass Then:
    # NoObjDsc:
    # Set sRules = Nothing                   ' Rules can not be used for differrent class
    # Set aTD = Nothing
    # Set sDictionary = Nothing
    # ' get all attributes; relevant properties are located
    # dcCount = 0
    # Set DeletionCandidates = New Dictionary
    # LListe = vbNullString

    if SelectOnlyOne Then:
    # aItmIndex = 1
    # AllPropsDecoded = InitsForDecoding(forItem2:=False)
    if DontDecode Then:
    # GoTo preponly                          ' no decode or get any attributes yet
    # ' Use main Object Identification as used by sort!
    # Stop                                       ' hier geht's daneben:
    # Matchcode = GetObjectAttributes(True, False)

    # Call LogEvent(String(60, "_") & vbCrLf & "Attribute von " _
    # & Folder(1).FolderPath & ": " & WorkIndex(1) _
    # & b & MainObjectIdentification _
    # & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))

    if displayInExcel Then:
    if xlApp Is Nothing Then:
    # Call XlgetApp
    # OisN:
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl)
    # Set O = x                          ' open default but don't Show it
    if O Is Nothing Then:
    # GoTo OisN
    # Call StckedAttrs2Xcel(O)

    # aOD(0).objMaxAttrCount = aOD(1).objMaxAttrCount
    if Not xlApp Is Nothing Then:
    # Call DisplayExcel(O, relevant_only:=True, _
    # EnableEvents:=False, _
    # unconditionallyShow:=True)
    # Call DisplayWindowInFront(xLMainWindowHdl, 1)
    if Not ShutUpMode Then:
    print(Debug.Print "done, one Selected Compare Item, Excel visible and not waiting")
    # GoTo preponly
    else:
    # aItmIndex = 0
    # Call Compare2SelectedItems
    # Call DoTheDeletes

    # FuncExit:
    if TerminateRun Then:
    # GoTo ProcReturn
    # preponly:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function setDefaultWindowName
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setdefaultwindowname():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.setDefaultWindowName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Static WindowNamePattern(1 To 2) As String
    if LenB(WindowNamePattern(wTx)) = 0 Then:
    match wTx:
        case 1                                 ' Excel:
    # WindowNamePattern(wTx) = RTail(TemplateFile, "\") _
    # & " - " & Replace(xlApp.Name, "Microsoft ", "*") ' geht bei [Schreibgeschtzt] nicht
        case 2                                 ' Outlook:
    # WindowNamePattern(wTx) = "* - Microsoft " & olApp.Name
        case _:
    # DoVerify False
    # setDefaultWindowName = WindowNamePattern(wTx)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowDetails
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showdetails():
    # Dim zErr As cErr
    # Const zKey As String = "CompareOps.ShowDetails"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ListContent(ListCount).MatchData = Message _
    # & vbCrLf & MainObjectIdentification _
    # & ": " & fiMain(1) & vbCrLf & AllDetails
    # ListContent(ListCount).DiffsRecognized = AllItemDiffs
    # rID = ListCount
    # frmCompareInfo.Show

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AddItemDataToOlderFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def additemdatatoolderfolder():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.AddItemDataToOlderFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call Pick2Folders
    # AcceptCloseMatches = True                      ' User may refine comparison result if closely related items
    # quickChecksOnly = True                         ' initially, only check most relevant
    # Call InitFolderOps(False)                      'sort both Folders, not only second
    # Call getCriteriaList

    # Call MergeItemsNewIntoOld
    if TerminateRun Then:
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AskForUserAction
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def askforuseraction():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.AskForUserAction"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim movedItem As Object
    # Dim targetItem As Object
    # Dim DelObjectItem As cDelObjectsEntry

    # Set targetItem = aID(1).idObjItem
    # Set DelObjectItem = New cDelObjectsEntry
    # DelObjectItem.DelObjPindex = 1
    # DelObjectItem.DelObjPos = WorkIndex(1)
    # DelObjectItem.DelObjInd = sortedItems(1).Count > 0 ' from sorted items?

    if Missing Then:
    # & Quote(Folder(2).FolderPath) & "  bernehmen?" _
    # & vbCrLf & "    " & Quote(fiMain(1)), vbYesNoCancel, _
    # "Behandlung eines Items in " & Quote(Folder(1).FolderPath))
    if rsp = vbYes Then:
    # ' if copy fails, it is general error so we do no ErrTry()
    # Set movedItem = aID(1).idObjItem.Copy
    if movedItem.Class = olAppointment Then:
    # movedItem.Subject = aID(1).idObjItem.Subject ' get rid of rename (seen on Appointment)
    else:
    # 'On Error GoTo 0
    if movedItem.Subject <> aID(1).idObjItem.Subject Then:
    # DoVerify False
    # aBugTxt = "move copied item to folder " & Folder(2).FolderPath
    # Call Try
    # Set movedItem = movedItem.Move(Folder(2))
    if Catch Then:
    # Message = "Item Copy " & aID(1).idObjItem.Subject _
    # & " NOT moved to " & Folder(2).FolderPath _
    # & " reason: " & E_AppErr.Description
    else:
    # Message = "Item Copy successfully moved into " _
    # & Folder(2).FolderPath _
    # & " as " & Quote(movedItem.Subject) & b
    # Call LogEvent(Message, eLall)

    elif rsp = vbNo Then:
    # Set targetItem = aID(1).idObjItem
    # delOption:
    # & Quote(targetItem.Parent.FolderPath) & "  lschen?" _
    # & vbCrLf & "    " & Quote(fiMain(1)) & "  -> " _
    # & Quote(TrashFolder.FolderPath), _
    # vbYesNoCancel, "Behandlung eines Items in " _
    # & Quote(targetItem.Parent.FolderPath))
    if rsp = vbYes Then:
    # Call TrashOrDeleteItem(DelObjectItem)
    elif rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    elif rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    if SomeMatch Or ActionID = atOrdnerinhalteZusammenfhren Then:
    if ActionID = atOrdnerinhalteZusammenfhren Then:
    # ' do not copy to target, identical item exists
    else:
    # DoVerify False, " strange or not implemented ???"
    else:
    # DelObjectItem.DelObjPindex = 2
    # DelObjectItem.DelObjPos = WorkIndex(2)
    # DelObjectItem.DelObjInd = sortedItems(2).Count > 0 ' from sorted items?
    # Set targetItem = aID(2).idObjItem
    # GoTo delOption

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CompareItemsIn2Folders
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def compareitemsin2folders():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.CompareItemsIn2Folders"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call Pick2Folders
    # Call InitFolderOps(False)                      'sort both Folders, not only second
    # Call CompareItemsNew2Old
    if TerminateRun Then:
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CompareItemsNew2Old
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def compareitemsnew2old():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.CompareItemsNew2Old"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # ' Process Folder items
    # DeletedItem = inv
    # DeleteIndex = -1
    # DateSkipCount = 0

    # WorkIndex(1) = 0
    # WorkIndex(2) = 0
    if 0 >= Ictr(1) Or 0 >= Ictr(2) Then           ' nothing to compare:
    # GoTo endsub
    # DateSkipCount = 0
    # Do                                             ' loop follows a synch as long as there's something left to do
    # WorkIndex(1) = WorkIndex(1) + 1
    if ItemDateFilter(sortedItems(1).Item(WorkIndex(1))) = vbNo Then:
    # GoTo nextOne
    # WorkIndex(2) = WorkIndex(2) + 1
    if ItemDateFilter(sortedItems(2).Item(WorkIndex(2))) = vbNo Then:
    # WorkIndex(1) = WorkIndex(1) - 1
    # GoTo nextOne
    # AllDetails = vbNullString

    # ' Use main Object Identification as used by sort!
    # Matchcode = quickCheck()

    if Matchcode = inv Then:
    # GoTo nextOne

    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & Folder(1).FolderPath & ": " & WorkIndex(1) & " und " _
    # & WorkIndex(2) & ", " & MainObjectIdentification _
    # & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))
    if fiMain(1) <> fiMain(2) Then:
    # Call LogEvent(WorkIndex(2) & ": " & fiMain(2))

    if Matchcode = Passt_Synch Then:
    if ItemIdentity() Then:
    # ' ==================================================================
    # ListContent(ListCount).Compares = "="
    else:
    # ListContent(ListCount).Compares = "<>"
    # ListContent(ListCount).MatchData = MatchData
    # ListContent(ListCount).DiffsRecognized = DiffsRecognized
    # Loop Until WorkIndex(1) >= Ictr(1) Or WorkIndex(2) >= Ictr(2)

    if DateSkipCount > 0 Then:
    # Call AddItemToList((DateSkipCount), "Eintrge bersprungen weil Datum", vbNullString, vbNullString)
    # ListContent(ListCount).MatchData = "vor dem"
    # ListContent(ListCount).DiffsRecognized = CStr(CutOffDate)

    if ListContent.Count > 0 Then:
    # frmDeltaList.Show
    # endsub:
    # Set ListContent = Nothing
    # Set sortedItems(1) = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MergeItemsNewIntoOld
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mergeitemsnewintoold():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.MergeItemsNewIntoOld"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim reportedMatches As Long
    # Dim ix1 As Long
    # Dim ix2 As Long
    # Dim foundItem As Object
    # Dim MandatoryWorkRule As cNameRule
    # DeletedItem = inv
    # DeleteIndex = -1
    # ListCount = 0
    # ' work rule is potentially changed interactively
    # Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)

    # Set foundItem = GetAobj(1, ix1)

    # Call find_Corresponding(foundItem, _
    # MandatoryWorkRule, _
    # reportedMatches, _
    # eliminateID:=True)

    if reportedMatches > 0 Then                ' found one or more:
    # Call LogEvent("   " & SelectedItems.Count _
    # & " corresponding item(s) found for " _
    # & RestrictCriteriaString)


    # sHdl = "CritPropName---------------" _
    # & " Quelle-" & ix1 & "----------------------" _
    # & " Ziel-" & ix2 & "------------------------" _
    # & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    # ' loop follows a synch as long as there's something left to do
    # Call GetAobj(aPindex, ix2)
    if objTypName <> DecodeObjectClass(True) Then:
    # GoTo nextOne
    # AllDetails = vbNullString
    # Call AddItemToList((ix1), fiMain(2), vbNullString, (ix2))
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & Folder(1).FolderPath & ": " & ix1 & " und " _
    # & Folder(2).FolderPath & ": " & ix2 & ", " _
    # & MainObjectIdentification _
    # & ": " & vbCrLf & ix1 & ": " & fiMain(1))
    if fiMain(1) <> fiMain(2) Then:
    # Call LogEvent(ix2 & ": " & fiMain(2))

    if ItemIdentity() Then:
    # ' ==================================================================
    # ListContent(ListCount).Compares = "="
    # Call AskForUserAction(False, True)
    else:
    # ListContent(ListCount).Compares = "<>"
    # Call AskForUserAction(False, False)
    # ListContent(ListCount).MatchData = MatchData
    # ListContent(ListCount).DiffsRecognized = DiffsRecognized
    else:
    # Call LogEvent("   No corresponding item found for " _
    # & RestrictCriteriaString)
    # Call AskForUserAction(True, False)
    if ActionID <> 6 Then:
    # frmDeltaList.Show
    # Set ListContent = Nothing

    # Set sortedItems(1) = Nothing
    # Set SelectedItems = New Collection

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Pick2Folders
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def pick2folders():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.Pick2Folders"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim action As String
    # Dim afolder As Folder
    # Call InitObjectSelection
    # eOnlySelectedItems = False
    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # bDefaultButton = "Go"
    # Call PickAFolder(1, _
    # "bitte whlen Sie den Ordner mit den neueren (=Quell) Objekten" _
    # & action, "Auswahl der zu vergleichenden Ordner", "OK", "Cancel")

    if topFolder Is Nothing Then:
    # Set topFolder = Folder(1)
    else:

    # curFolderPath = Folder(1).FolderPath

    # Set afolder = Folder(1)

    # Call LogEvent("==== User has selected Source " & curFolderPath, eLall)
    # Call PickAFolder(2, _
    # "bitte whlen Sie den Ordner mit den lteren (=Ziel) Objekten" _
    # & action, "Auswahl der zu vergleichenden Ordner", _
    # "OK", "Cancel")
    # Call LogEvent("==== Target Folder is called  " _
    # & Folder(2).FolderPath, eLall)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MakeSelection
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def makeselection():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.MakeSelection"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aFolderpath As String

    # doAskForSelection:
    match rsp:
        case vbOK:
    if ActiveExplorer.Selection.Count = 0 Then:
    # GoTo doAskForSelection
    if ActiveExplorer.CurrentFolder Is Nothing Then:
    # Set Folder(i) = aNameSpace.PickFolder
    # GoTo doAskForSelection
    else:
    # Set Folder(i) = ActiveExplorer.CurrentFolder
    # ' DoVerify False, " watchout: volatile result, but sometimes this works, sometimes not"
    # aFolderpath = Folder(i).FolderPath
    # ' *** bug: setting Folder sets/ruins Folder(1) although i=2, both are set to same Folder
    if i > 1 Then:
    # Set Folder(i) = GetFolderByName(aFolderpath, Folder(i)) ' get a non-volalite pointer
    # SkipNextInteraction = True
        case vbCancel:
    # SkipNextInteraction = False
    if But2 = "Cancel" Then:
    # Call LogEvent("=======> Keine Auswahl, Abbruch: " & Now(), eLnothing)
    # End
    # Call LogEvent("=======> Es wurden " & ActiveExplorer.Selection.Count _
    # & " Items aus dem Ordner " & Quote(aFolderpath) _
    # & " gewhlt. Time: " & Now(), eLnothing)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PickAFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def pickafolder():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.PickAFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aFolderpath As String
    # rsp = vbAbort
    if LF_UsrRqAtionId = atPostEingangsbearbeitungdurchfhren Then:
    if ActiveExplorer.CurrentFolder = SpecialSearchFolderName Then:
    # rsp = vbOK
    if rsp = vbAbort Then:
    # doAskForFolder:

    match rsp:
        case vbOK:
    # ' *** bug: this sets/ruins Folder(1) although i=2, both are set to same Folder
    if ActiveExplorer.CurrentFolder Is Nothing Then:
    # Set Folder(i) = aNameSpace.PickFolder
    else:
    # Set Folder(i) = ActiveExplorer.CurrentFolder
    if DebugMode Then:
    print(Debug.Print "watchout: volatile result, but sometimes this works, sometimes not")
    # aFolderpath = Folder(i).FolderPath
    # Set Folder(i) = GetFolderByName(aFolderpath, Folder(i), noSearchFolders:=True) ' get a non-volalite pointer
    if Folder(i) Is Nothing Then:
    # PickTopFolder = True
    # GoTo doAskForFolder
    else:
    # PickTopFolder = False
    # SkipNextInteraction = True
        case vbCancel:
    # SkipNextInteraction = False
    if But2 = "Cancel" Then:
    # Call LogEvent("=======> Kein Ordner gewhlt, Abbruch: " & Now(), eLnothing)
    # End
    # Call LogEvent("=======> Walking all Folders. Time: " & Now(), eLnothing)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitFolderOps
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initfolderops():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.InitFolderOps"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Filter As String
    if Folder(1).DefaultItemType <> Folder(2).DefaultItemType Then:
    # Folder(2).Name, _
    # "Folder Types do not Match: Choose primary type")
    if rsp = vbOK Then:
    # Matchcode = 1
    elif rsp = vbCancel Then:
    # Matchcode = 2
    else:
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    # Matchcode = 1                              ' the Folder types are identical, use only first to determine GetObjectAttributess

    # Call BestObjProps(Folder(Matchcode), withValues:=False)
    # Restart:
    if SortingOnlySecond Then:
    if getFolderFilter(Folder(2).Items.Item(1), CutOffDate, _:
    # Filter, ">=") Then
    # Set sortedItems(1) = Folder(2).Items.Restrict(Filter)
    else:
    # Set sortedItems(1) = Folder(2).Items
    else:
    if Folder(1).Items.Count = 0 Then:
    # Set sortedItems(1) = Folder(1).Items
    elif getFolderFilter(Folder(1).Items.Item(1), CutOffDate, _:
    # Filter, ">=") Then
    # Set sortedItems(1) = Folder(1).Items.Restrict(Filter)
    else:
    # Set sortedItems(1) = Folder(1).Items
    # aBugTxt = "Sort Matches"
    # Call Try                                       ' Try anything, autocatch
    # sortedItems(1).sort SortMatches, False
    if Catch Then:
    # Message = vbCrLf & E_AppErr.Description _
    # & vbCrLf & vbCrLf & "Bitte ndern Sie die [Sortierparameter]"
    # b1text = "Weiter"
    # b2text = vbNullString
    # Set FRM = New frmDelParms
    # FRM.Show
    # Set FRM = Nothing
    if rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo Restart

    # Ictr(1) = sortedItems(1).Count
    if Not SortingOnlySecond Then:
    # Set sortedItems(2) = Folder(2).Items
    # Call Try                                   ' Try anything, autocatch
    # sortedItems(2).sort SortMatches, False
    if Catch Then:
    # Message = vbCrLf & E_AppErr.Description & vbCrLf & vbCrLf _
    # & "Bitte ndern Sie die [Sortierparameter]"
    # b1text = "Weiter"
    # b2text = vbNullString
    # Set FRM = Nothing
    # Set FRM = New frmDelParms
    # FRM.Show
    # Set FRM = Nothing
    if rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo Restart
    else:
    # Call ErrReset(0)
    # Message = "Sortierung erfolgreich." & vbCrLf _
    # & "Anzahl der enthaltenen Objekte:" & vbCrLf _
    # & "Ordner 1=" & Folder(1).FolderPath _
    # & ", Items: " & Ictr(1) _
    # & " Ordner 2=" & Folder(2).FolderPath _
    # & ", Items: " & sortedItems(2).Count & vbCrLf _
    # & "Bitte besttigen Sie die Parameter fr die Vergleichsoperationen"
    # b1text = "Weiter"
    # b2text = vbNullString
    # b3text = "Cancel"
    # Set FRM = Nothing
    # Set FRM = New frmDelParms
    # FRM.Show
    if DebugMode Then:
    # DoVerify False
    # Set FRM = Nothing
    if rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    # Ictr(2) = sortedItems(2).Count

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DetailsToPrintFile
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def detailstoprintfile():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.DetailsToPrintFile"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    try:
        # Open FileName For Output As #2
        # Print #2, "Details zu: " _
        # & MainObjectIdentification & ": " & fiMain(1)
        # Print #2, vbCrLf & AllDetails
        # ficl:
        # Close #2

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function quickCheck
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def quickcheck():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.quickCheck"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i1 As Long
    # Dim i2 As Long
    # Dim MainCompare As Long
    # quickCheck = Passt_Synch

    # Synch:

    # Call GetAobj(1, i1)
    # objTypName = DecodeObjectClass(getValues:=False)

    # Call GetAobj(1, i2)
    # objTypName = DecodeObjectClass(getValues:=False)

    # MainCompare = StrComp(fiMain(1), fiMain(2), vbTextCompare)
    if MainCompare = 0 Then                ' Identische strings:
    # quickCheck = Passt_Synch
    if ListCount > 0 Then              ' schon ein Eintrag da?:
    if LenB(ListContent(ListCount).Index2) = 0 _:
    # And ListContent(ListCount).Index1 = i1 Then ' der linke Eintrag ist schon da: zusammenfassen
    # ListContent(ListCount).Index2 = i2 ' dann rechts mit diesem vergleichen
    # ListContent(ListCount).Compares = (Passt_Deleted) ' deleted, weil rechts i1 nicht gefunden wurde
    elif ListContent(ListCount).Index1 = vbNullString _:
    # And ListContent(ListCount).Index2 = i2 Then ' der rechte Eintrag ist schon da: zusammenfassen
    # ListContent(ListCount).Index1 = i1 ' dann rechts mit diesem vergleichen
    # ListContent(ListCount).Compares = (Passt_Deleted) ' deleted, weil rechts i2 nicht gefunden wurde
    else:
    if ListContent(ListCount).Index1 <> i1 _:
    # Or ListContent(ListCount).Index2 <> i2 Then ' Doublette mglich
    # Call AddItemToList((i1), fiMain(1), (Passt_Synch), (i2))
    else:
    # DoVerify False, " wie geht das denn ?"
    else:
    # Call AddItemToList((i1), fiMain(1), (Passt_Synch), (i2))
    # WorkIndex(1) = i1                  ' we continue here on next entry
    # WorkIndex(2) = i2
    # GoTo ProcReturn
    elif MainCompare < 0 Then            ' links ist das kleinere, als erstes zufgen:
    # Call AddItemToList((i1), fiMain(1), (Passt_Deleted), vbNullString)
    # WorkIndex(1) = i1
    # WorkIndex(2) = i2
    # ListContent(ListCount).DiffsRecognized = _
    # MainObjectIdentification _
    # & "(2): " & fiMain(2) & " kleiner als " & fiMain(1)
    # GoTo ni1                           ' do not inc i2
    else:
    # Call AddItemToList(vbNullString, fiMain(2), (Passt_Inserted), (i2))
    # WorkIndex(1) = i1
    # WorkIndex(2) = i2
    # ListContent(ListCount).DiffsRecognized = _
    # MainObjectIdentification _
    # & "(1): " & fiMain(1) & " grer als Rechts"
    # ni1:
    # quickCheck = inv
    # WorkIndex(1) = i1
    # WorkIndex(2) = i2

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SetFilter
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setfilter():
    # Dim zErr As cErr
    # Const zKey As String = "FolderOps.SetFilter"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if LenB(Trim(AttrValue)) > 0 Then:
    # SetFilter = adName & " = " & Quote(AttrValue)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:


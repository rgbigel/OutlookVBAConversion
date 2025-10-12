# Converted from DupeDeleter.py

# Attribute VB_Name = "DupeDeleter"
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Function AskUserAndInterpretAnswer
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def askuserandinterpretanswer():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.AskUserAndInterpretAnswer"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim UserAnswer As String
    # Dim passtGenau As Boolean
    # Dim passtEtwa As Boolean
    # Dim itemCount As Long
    # Dim isFromSortedCollection As Boolean

    # '   Message = oMessage ?????
    # passtGenau = cMisMatchesFound = 0
    # passtEtwa = (cMisMatchesFound < MaxMisMatchesForCandidates _
    # And AcceptCloseMatches _
    # And Not passtGenau _
    # And Not SuperRelevantMisMatch)
    if sortedItems(1) Is Nothing Then:
    if SelectedItems Is Nothing Then:
    if Not passtGenau Or MPEchanged Then:
    # AllItemDiffs = AllItemDiffs & vbCrLf & Mid(MPEItemDiffs, 3)
    # UserAnswer = Quote(fiMain(1)) & " hat folgende Abweichungen: " _
    # & vbCrLf & AllItemDiffs
    else:
    # UserAnswer = Quote(fiMain(1)) & " hat keine relevanten Abweichungen" _
    # & vbCrLf & AllItemDiffs
    if MPEchanged Or IsComparemode Then:
    # GoTo justShow
    else:
    # rsp = vbIgnore
    # GoTo rspIsSet
    else:
    # itemCount = SelectedItems.Count
    # isFromSortedCollection = False
    else:
    # itemCount = sortedItems(1).Count
    # isFromSortedCollection = True

    if AcceptCloseMatches Then:
    if cMisMatchesFound < MaxMisMatchesForCandidates Then:
    # UserDecisionRequest = True
    elif Not IsComparemode Then:
    # UserDecisionRequest = False

    if eOnlySelectedItems _:
    # And itemCount < 5 _
    # And Not IsComparemode Then
    # UserDecisionRequest = True
    if Not AllPropsDecoded Then:
    # UserAnswer = Quote(fiMain(1)) _
    # & " sollte bei unvollstndigem Vergleich nicht gelscht werden " _
    # & vbCrLf & AllItemDiffs
    # justShow:
    # AskUserAndInterpretAnswer = False
    # Diffs = AllItemDiffs
    if Not displayInExcel Or Not AllPropsDecoded Then:
    # GoTo askuser
    else:
    # Call DisplayExcel(O, _
    # relevant_only:=True, _
    # unconditionallyShow:=True)
    elif IsComparemode Or passtGenau _:
    # Or passtEtwa Or UserDecisionRequest Then
    # ' im Prinzip lschen sinnvoll bzw Benutzer kann alles...
    # LoeschbesttigungCaption = "Lschen von Objekten besttigen "
    # askuser:
    # Message = oMessage
    # bDefaultButton = "Go"  ' der vordere Button ist IMMER Default
    if MatchPoints(1) <= MatchPoints(2) Then:
    # rsp = vbYes         ' der Bessere oder der ltere wird gelscht
    # ' (Modification Date Sorted ascending)
    # b1text = WorkIndex(1)
    # b2text = WorkIndex(2)
    else:
    # b2text = WorkIndex(1)
    # b1text = WorkIndex(2)
    # rsp = vbNo
    if Not AllPropsDecoded Then:
    # rsp = vbCancel  ' default is not delete if compare is incomplete
    # bDefaultButton = "Cancel"
    # Diffs = UserAnswer
    # GoTo ShowLoeschbestaetigung

    if Not passtEtwa _:
    # And UserDecisionRequest And AcceptCloseMatches _
    # And cMisMatchesFound < MaxMisMatchesForCandidates Then
    # passtEtwa = True

    if WantConfirmation Or passtEtwa And Not passtGenau Then:
    # Diffs = fiMain(1) & vbCrLf & " +++ " & objTypName _
    # & " in " & curFolderPath _
    # & ", Item Nr's. " & b1text & " und " & b2text _
    # & " Lnge des Inhalts: " & Len(fiBody(1)) _
    # & vbCrLf & Message
    # ' Nutzerbesttigung notwendig
    if passtGenau Then ' ... weil explizit gewnscht:

    # Message = "==> Item " & b1text _
    # & " oder " _
    # & b2text _
    # & " wird " _
    # & killMsg
    # Message = Message & vbCrLf _
    # & "[Wenn Lschungen im Ordner " & curFolderPath _
    # & " ab jetzt nicht mehr besttigt werden sollen, Parameter ndern.]"
    else:
    # Message = "==> Item " & b1text _
    # & " oder " _
    # & b2text _
    # & " wird " _
    # & killMsg _
    # & ". " _
    # & b3text _
    # & " lscht nichts."
    # bDefaultButton = "Cancel"

    if xUseExcel And displayInExcel Then:
    if Not xlApp.Visible Then:
    # UserDecisionRequest = True  ' ask after excel display
    # Call DisplayExcel(O, _
    # relevant_only:=True, _
    # unconditionallyShow:=True)
    # ShowLoeschbestaetigung:
    # Diffs = Diffs & vbCrLf & vbCrLf & Mid(MatchData, 3)
    if InStr(Diffs, AllItemDiffs) = 0 Then:
    # Diffs = Diffs & vbCrLf & AllItemDiffs
    # DeleteNow = False
    # Set LBF = Nothing
    # Set LBF = New frmDelConfirm
    # Call Try
    # LBF.Show
    if Catch Then:
    # DoVerify False
    # Set LBF = Nothing
    # Call ErrReset(0)
    if DeleteNow Then:
    # DoTheDeletes
    # GoTo askuser
    if askforParams Then:
    # askforParams:
    # b1text = "Zurck"
    # b2text = "DoVerify"                        ' NOTE: button Name is "bDebugStop"
    # Message = "Bei " & b3text _
    # & " wird die Doublettensuche beendet, " _
    # & "ohne Lschungen " _
    # & " durchzufhren. Es liegen bisher Lsch-Vormerkungen fr " _
    # & dcCount & " Eintrge vor."
    # Set LBF = Nothing
    # Set LBF = New frmDelParms
    # LBF.Show
    # Set LBF = Nothing
    # askforParams = False
    if rsp = vbCancel Then:
    # Message = "Verarbeitung abgebrochen"
    # Call LogEvent(Message, eLall)
    if TerminateRun Then:
    # GoTo FuncExit
    if LenB(b3text) = 0 Then:
    # UserAnswer = "Lschantwort=Debug Debug.Assert False" & vbCrLf
    # AskUserAndInterpretAnswer = False
    # DeletedItem = WorkIndex(1)
    # DeleteIndex = 0
    # rsp = vbRetry
    # GoTo logandout
    else:
    if askforParams Then:
    # GoTo askforParams
    else:
    # GoTo askuser
    # bDefaultButton = "Nutzer whlte Button "
    elif rsp = vbIgnore Then     ' full Match and no confirmation:
    # DeleteIndex = 1
    # bDefaultButton = "Automatische Selektion, gelscht wird " _
    # & b1text _
    # & ", " & fiMain(DeleteIndex) & vbCrLf
    # rsp = vbYes ' delete default item, lesser Match or older
    # rspIsSet:
    if displayInExcel Then:
    # Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    if xlApp.Visible Then:
    # xlApp.EnableEvents = False
    # xlApp.Visible = False
    # Catch

    match rsp:
        case vbNo:
    if b2text = WorkIndex(2) Then:
    # DeleteIndex = 2
    elif b2text = WorkIndex(1) Then:
    # DeleteIndex = 1
    else:
    # DoVerify False
    # UserAnswer = bDefaultButton & DeleteIndex _
    # & ", Lschvormerkung zu " & b2text _
    # & b & fiMain(DeleteIndex) & vbCrLf
    # DeletedItem = b2text
    # AskUserAndInterpretAnswer = True
    # Call GenerateDeleteRequest(DeleteIndex, DeletedItem, isFromSortedCollection)
        case vbYes:
    if b1text = WorkIndex(2) Then:
    # DeleteIndex = 2
    elif b1text = WorkIndex(1) Then:
    # DeleteIndex = 1
    else:
    # DoVerify False
    # UserAnswer = bDefaultButton & DeleteIndex _
    # & ", Lschung von " & b1text _
    # & b & fiMain(DeleteIndex) & vbCrLf
    # DeletedItem = b1text
    # AskUserAndInterpretAnswer = True
    # Call GenerateDeleteRequest(DeleteIndex, DeletedItem, isFromSortedCollection)
        case vbRetry:
    # UserAnswer = "Retry: Items komplett in Excel anzeigen (keine Lschung)" & vbCrLf
    # AllPropsDecoded = False
    # AskUserAndInterpretAnswer = False
        case Else       ' cancel request:
    # UserAnswer = "Lschantwort=Cancel (keine Lschung)" & vbCrLf
    # AskUserAndInterpretAnswer = False
    # DeletedItem = -1
    # DeleteIndex = 0
    else:
    # AskUserAndInterpretAnswer = False
    # UserAnswer = "(" & fiMain(1) & ") sind nicht gleich " _
    # & vbCrLf & AllItemDiffs
    # logandout:
    # Call LogEvent("Items: " & WorkIndex(1) & "/" & WorkIndex(2) _
    # & b & UserAnswer, eLall)
    # UserDecisionEffective = True

    # FuncExit:
    # Set LBF = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GenerateDeleteRequest
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def generatedeleterequest():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.GenerateDeleteRequest"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim DelObjectEntry As cDelObjectsEntry

    # Set DelObjectEntry = New cDelObjectsEntry
    # DelObjectEntry.DelObjPos = delPx
    # DelObjectEntry.DelObjPindex = delItemIndex
    # DelObjectEntry.DelObjInd = doSort
    if DeletionCandidates Is Nothing Then:
    # Set DeletionCandidates = New Dictionary
    # DeletionCandidates.Add delItemIndex, WorkIndex(delPx)

    # LListe = LimitAppended(LListe, ", " & delItemIndex, 255, "... ")

    # dcCount = dcCount + 1

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub PerformChangeOpsForMapiItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def performchangeopsformapiitems():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.PerformChangeOpsForMapiItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim sourcecol As Long
    # Dim px As Long
    if O Is Nothing Then:
    # GoTo ProcReturn
    # ' if WithEditing only!!
    # With O.xlTSheet
    if Not O Is Nothing Then:
    for px in range(1, 3):
    if .Cells(1, px + 1).Text = "del" Then:
    if px = 1 Then:
    # delSource(px) = WorkIndex(1)
    else:
    # delSource(px) = WorkIndex(2)
    if sortedItems(px) Is Nothing Then:
    # Call GenerateDeleteRequest(px, delSource(px), False)
    else:
    # Call GenerateDeleteRequest(px, delSource(px), _
    # sortedItems(px).Count > 0)

    # Call LogEvent("Item " & delSource(px) _
    # & " in Excel zum Lschen vorgemerkt", eLall)
    # UserDecisionEffective = True
    else:
    # delSource(px) = 0

    if .Cells(1, changeCounter).Value = 0 Then:
    # Call LogEvent("Es wurden keine nderungen in Excel durchgefhrt", eLall)
    # UserDecisionEffective = True
    # GoTo ProcReturn

    if DebugMode Or DebugLogging Then:
    print(Debug.Print Format(Timer, "0#####.00") _)
    # & vbTab & "Performing changes/deletes from Excel to Outlook"
    # Err.Clear
    if Left(.Cells(i, ValidCol).Text, 3) <> "***" Then ' Value is editable:
    # px = 0    ' used as flag for no changes in this line
    if .Cells(i, ChangeCol3).Text = "<" Then:
    # sourcecol = ChangeCol2
    # px = 1
    elif .Cells(i, ChangeCol3).Text = ">" Then:
    # sourcecol = ChangeCol1
    # px = 2
    elif .Cells(i, ChangeCol3).Text = "?" Then ' Error previously:
    # GoTo nloop  ' will not edit this
    elif .Cells(i, ChangeCol3).Text = "!" Then ' success previously:
    # GoTo nloop  ' will not edit this
    # ' Indicator of change "<=>" is in cols 14/15
    elif InStr(.Cells(i, WatchingChanges).Text, ">") > 0 Then:
    # ' changecounter1--V V--changecounter2
    # ' if something changed in col 2/3, col 17/18 contain old values
    # ' col 7 / 8 contain raw values, e.g. used to compare empty sources
    # ' col 7, 17/ 8, 18 is updated on the excel side when selection changes
    # ' check if the original value differs at all
    if .Cells(i, ChangeCol1).Text <> .Cells(i, 17).Text _:
    # Or .Cells(i, ChangeCol1).Text <> .Cells(i, 7).Text Then
    # sourcecol = ChangeCol1
    # px = 1
    # ' Indicator of change is in cols 14/15
    elif InStr(.Cells(i, changeCounter).Text, ">") > 0 Then:
    if .Cells(i, ChangeCol2).Text <> .Cells(i, 18).Text _:
    # Or .Cells(i, ChangeCol2).Text <> .Cells(i, 8).Text Then
    # sourcecol = ChangeCol2
    # px = 2
    else:
    # GoTo nloop
    if ErrorCaught = 0 Then:
    # .Cells(i, ChangeCol3).Value = "!"
    else:
    # .Cells(i, ChangeCol3).Value = "?"
    # nloop:
    if delSource(px) = 0 Then:
    if px > 0 Then:
    # Call storeAttribute(i, sourcecol, px)

    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") _)
    # & vbTab & "Ended Performing changes/deletes from Excel to Outlook"
    # End With ' O.xlTSheet

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub CheckDoublesInFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkdoublesinfolder():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.CheckDoublesInFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim itemNo As Long
    # Dim sortFields As String
    # Dim sortOrder As String
    # Dim sCompRes As Long
    # Dim ShortName As String
    # Dim totalDeletes As Long
    # Dim deletedNow As Long
    # Dim tCompRes As String

    # IsComparemode = False
    # Call FindTrashFolder
    if curFolder = topFolder And FolderPathLevel = 0 Then:
    # itemNo = 0
    else:
    # ' *** curFolderPath = FullFolderPath(FolderPathLevel) & "\" & curFolder.Name
    # ' FolderPathLevel = FolderPathLevel + 1
    # ' FullFolderPath(FolderPathLevel) = curFolderPath
    # Set Folder(1) = curFolder
    # Set Folder(2) = curFolder
    if eOnlySelectedItems Then:
    # Fctr(1) = 0
    # Ictr(1) = SelectedItems.Count
    else:
    if curFolder.Parent Is Nothing Then:
    # Fctr(1) = 0 ' no parent => no Folders
    else:
    # Fctr(1) = curFolder.Folders.Count
    # Ictr(1) = curFolder.Items.Count
    # Ictr(2) = Ictr(1)
    # Fctr(2) = Fctr(1)
    # ShortName = Left(curFolder.Name, 4)
    if ShortName = "Dele" _:
    # Or ShortName = "Gel" _
    # Or ShortName = "Tras" _
    # Or ShortName = "Junk" _
    # Or ShortName = "Spam" _
    # Or ShortName = "Uner" _
    # Or ShortName = "Dupl" _
    # Or Fctr(1) + Ictr(1) = 0 _
    # Then ' skip unwanted stuff; could contain any item type!
    # skipThis:
    # Call LogEvent("<======> skipping Folder " & curFolder.FolderPath _
    # & " because it matches " & ShortName _
    # & " Time: " & Now())
    # GoTo ProcReturn

    # ' hamma schon: Call BestObjProps(curFolder)
    if AskEveryFolder _:
    # And Not (SkipNextInteraction Or eOnlySelectedFolder) Then
    # WantConfirmationThisFolder = WantConfirmation

    if eOnlySelectedItems Then:
    # Message = "Bitte besttigen Sie die Parameter der Doublettensuche" _
    # & "in den selektierten Items"
    else:
    # Message = "Wollen Sie Doubletten suchen im Ordner " _
    # & curFolderPath
    if curFolder.Folders.Count > 0 Then:
    # Message = Message & vbCrLf & "(beginnend mit seinen " _
    # & curFolder.Folders.Count & " enthaltenen Ordnern)"
    # b1text = "Ja"
    # b2text = "Nein"
    # Set LBF = Nothing
    # Set LBF = New frmDelParms
    # bDefaultButton = "Go"
    # LBF.Caption = "Parameter fr das Lschen von Doubletten in Ordner " _
    # & curFolderPath
    # LBF.Show
    # Set LBF = Nothing

    match rsp:
        case vbNo:
    # GoTo skipThis
        case vbCancel:
    # Call LogEvent("=======> Stopped before Entering Folder " & Quote(curFolder.FolderPath) & _
    # " containing " & Fctr(1) & " Folders and " _
    # & Ictr(1) & " Items. Time: " & Now())
    if TerminateRun Then:
    # GoTo ProcReturn
        case _:
    # SkipNextInteraction = False

    if WantConfirmation = True Then:
    # Message = " (Confirmation mode)"
    else:
    # Message = vbNullString
    if eOnlySelectedItems Then:
    # Call LogEvent("=======> Processing selected Items in " & curFolderPath _
    # & Message & " Time: " & Now(), eLnothing)
    else:
    # Call LogEvent("=======> Entering Folder " & curFolderPath _
    # & " containing " & Fctr(1) & " Sub-Folders and " _
    # & Ictr(1) & " Items." & Message & " Time: " & Now())
    # Message = vbNullString
    # dcCount = 0
    # Set DeletionCandidates = New Dictionary
    # LListe = vbNullString

    # ' recurse into FOlDERS
    # ' ====================
    if DebugLogging Then:
    # DoVerify False, " if we are in the call chain of loopFolders, this should never be reached"
    if curFolder.DefaultItemType = olContactItem Then:
    # Exit For                                                ' do not recurse ContactItems
    if curFolder.Folders(itemNo).Items.Count > 0 Then:
    # Call CheckDoublesInFolder(curFolder.Folders(itemNo))    ' ProcCall recursion

    # ' Process Folder ITEMS
    # ' ====================
    # Restart:
    # StopLoop = False
    # Call GetSortableItems(curFolder)
    # Set aTD = Nothing
    # ' ??? *** only needed if a change of class is possible
    # Call SplitMandatories(TrueCritList)
    if Not aTD Is Nothing Then                  ' items to work on need new rules:
    # DoVerify False, " ??? *** change of class is possible"
    # aTD.adRules.RuleInstanceValid = False
    # Call SplitDescriptor(aTD)   ' aTD is possibly changed now
    # sortFields = SortMatches

    # Err.Clear
    if Not sortedItems(1) Is Nothing Then:
    # aBugTxt = "sort using sortFields=" & sortFields
    # Call Try
    if sortedItems(1).Count > 1 Then:
    if curFolder.DefaultItemType = olContactItem Then:
    # sortedItems(1).sort sortFields, olAscending
    # sortOrder = "Ascending"
    else:
    # sortedItems(1).sort sortFields, False
    # sortOrder = "Descending"
    if Catch Then:
    # Message = vbCrLf & Err.Description & vbCrLf & vbCrLf _
    # & "Bitte ndern Sie die [Sortierparameter]: " _
    # & sortFields
    # b1text = "Weiter"
    # b2text = vbNullString
    # Set LBF = Nothing
    # Set LBF = New frmDelParms
    # LBF.Show
    # Set LBF = Nothing
    if rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo Restart

    # Ictr(1) = sortedItems(1).Count
    # Set sortedItems(2) = sortedItems(1)
    # Ictr(2) = sortedItems(2).Count

    # fiMain(1) = vbNullString
    # fiMain(2) = vbNullString  ' This will cause the first comparison result =0

    # DeletedItem = -1
    # DeleteIndex = -1

    # WorkIndex(1) = 1  ' if ascending, this is the smaller one
    # WorkIndex(2) = 1  ' if descending, the next one is the smaller
    # Set aID(1) = Nothing
    # Set aID(2) = Nothing
    # Set aObjDsc = Nothing
    # sCompRes = 0
    # tCompRes = " ? "

    # ' note this scheme only makes sense if we are working on only one
    # ' sorted collection, which are obviously in one Folder
    # ' beware: as deletes are probably done in batches, item idices
    # ' are not monotonic when the deletes are actually done

    # Do While (WorkIndex(1) < Ictr(1) And WorkIndex(2) < Ictr(2))
    # Call ShowStatusUpdate
    # AllDetails = vbNullString
    # AllPropsDecoded = False
    # Set aTD = Nothing
    if DeletedItem = -1 Then:
    # ' doing a very fast compare based on main identifications
    # sCompRes = StrComp(fiMain(1), fiMain(2), vbTextCompare)
    # ' -1 means first is bigger, +1 first is smaller, =0 if same
    if sCompRes = 0 Then    ' fiMain(1) = fiMain(2):
    # tCompRes = " = "
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & curFolderPath & ": " _
    # & WorkIndex(1) & tCompRes _
    # & WorkIndex(2) & ", " & MainObjectIdentification _
    # & "= " & vbCrLf & WorkIndex(1) & ": " & Quote(fiMain(1)))
    if sortedItems(1).Item(WorkIndex(1)).EntryID _:
    # = sortedItems(2).Item(WorkIndex(2)).EntryID Then
    # GoTo advance_Second
    elif sCompRes < 0 Then ' fiMain(1) > fiMain(2):
    # tCompRes = " > "
    if MinimalLogging < eLSome Then:
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & curFolderPath & ": " _
    # & WorkIndex(1) & tCompRes _
    # & WorkIndex(2) & ", " & MainObjectIdentification _
    # & ": " & vbCrLf & WorkIndex(1) _
    # & ": " & Quote(fiMain(1)) & b, eLSome)
    # Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
    if sortOrder = "Descending" Then:
    # ' desc sort: stepping first will (probably) make
    # ' fimain(1) smaller, so match is possible next time
    # GoTo advance_First ' make fiMain(1) decrease
    else:
    # GoTo advance_Second
    else:
    # tCompRes = " < "
    if MinimalLogging < eLSome Then:
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & curFolderPath & ": " _
    # & WorkIndex(1) & tCompRes _
    # & WorkIndex(2) & ", " & MainObjectIdentification _
    # & ": " & vbCrLf & WorkIndex(1) _
    # & ": " & Quote(fiMain(1)) & b, eLSome)
    # Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
    if sortOrder = "Descending" Then:
    # GoTo advance_Second ' make fiMain(2) decrease
    else:
    # GoTo advance_First  ' make fiMain(1) bigger
    elif DeletedItem = WorkIndex(1) Then:
    # ' we have plans to delete first item, skip it for further comparison
    # advance_First:
    # WorkIndex(1) = WorkIndex(2)
    # Call CopyAttributeStack(2)
    # fiMain(1) = fiMain(2)
    # Set aID(1) = aID(2)
    # Set aObjDsc = aID(2).idObjDsc
    # Set aID(2) = Nothing
    # ' as item 1 is now on "left" side already, advance "right" side
    # WorkIndex(2) = WorkIndex(2) + 1
    # AttributeUndef(2) = 0
    elif DeletedItem = WorkIndex(2) Then:
    # ' if item 2 is deleted, skip on to the next on "right" side
    # advance_Second:
    # WorkIndex(2) = WorkIndex(2) + 1
    # Set aID(2) = Nothing
    # fiMain(2) = vbNullString
    # AttributeUndef(2) = 0
    else:
    # tCompRes = " und "
    if MinimalLogging < eLSome Then:
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " _
    # & curFolderPath & ": " _
    # & WorkIndex(1) & tCompRes _
    # & WorkIndex(2) & ", " & MainObjectIdentification _
    # & ": " & vbCrLf & WorkIndex(1) _
    # & ": " & Quote(fiMain(1)) & b, eLSome)
    # Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
    # advance_Both:
    # ' this because muliplicates are possible, not only duplicates
    # WorkIndex(1) = WorkIndex(1) + 1
    # Set aID(1) = Nothing
    # Set aObjDsc = Nothing
    # AttributeUndef(1) = 0
    # WorkIndex(2) = WorkIndex(2) + 1
    # Set aID(2) = Nothing
    # AttributeUndef(2) = 0
    # ' never compare identical objects
    if WorkIndex(1) = WorkIndex(2) Then ' possible after deleting:
    # ' do same steps as advance_Second
    # WorkIndex(2) = WorkIndex(2) + 1
    # Set aID(2) = Nothing
    # AttributeUndef(2) = 0

    if Not xlApp Is Nothing Then:
    if Not O Is Nothing Then:
    # ' if excel is open, we could have objDumpMade > 0
    # aOD(0).objDumpMade = 0   ' which would be bad info

    # mustDecodeRest = False
    if WorkIndex(2) > sortedItems(2).Count Then:
    # GoTo leaveLoop  ' running out of bounds
    # ' Progress indicator
    if (Max(WorkIndex(1), WorkIndex(2)) - 1) Mod 25 = 1 _:
    # And sortedItems(1).Count > 2 Then
    # Call LogEvent("   (progress info:) processing items " _
    # & WorkIndex(1) & " and " & WorkIndex(2), eLall)

    # ' this determines fiMain()
    if aID(1) Is Nothing Then:
    # Call GetAobj(1, WorkIndex(1))
    # objTypName = DecodeObjectClass(getValues:=True)
    if aID(2) Is Nothing Then:
    # Call GetAobj(2, WorkIndex(2))
    # objTypName = DecodeObjectClass(getValues:=True)

    if objTypName = "-" Then:
    # DoVerify False
    # GoTo nextOne

    if fiMain(1) <> fiMain(2) Then:
    if Not IsComparemode Or quickChecksOnly Then:
    # GoTo nextOne

    # Call ItemIdentity   ' identical by test OR by user decision possible!!
    # '            ==================================================================
    # DeleteIndex = 0
    if StopLoop Then:
    # GoTo leaveLoop
    if dcCount Mod 25 = 0 Then:
    if dcCount > 0 Then:
    # deletedNow = DoTheDeletes   ' user can still decide yes/no, true deletes 0/25
    # totalDeletes = totalDeletes + deletedNow ' we actually deleted that many, max 25
    # WorkIndex(1) = WorkIndex(1) - deletedNow ' correct for deleted items
    # WorkIndex(2) = WorkIndex(2) - deletedNow
    # Set logItem = Nothing   ' start a new log entry (speed!)
    # Loop

    # leaveLoop:
    if Not xlApp Is Nothing Then:
    # Call ClearWorkSheet(xlA, O)   ' erase this one, new one will be started when needed
    # totalDeletes = totalDeletes + DoTheDeletes  ' last round of deletes
    if sortedItems(1) Is Nothing Then:
    # Message = "<======= Exiting Folder "
    else:
    # itemNo = sortedItems(1).Count ' this is whats left after deletes
    if eOnlySelectedItems Then:
    # Message = "<=== compared " & Ictr(1) & " selected items, deleted " _
    # & totalDeletes & " as duplicates"
    # Call UnSelectItems  ' ???*** undoes Marking as SEL (in ManagerName)
    else:
    # Message = "<======= Exiting Folder "

    # Call LogEvent(Message & Quote(curFolderPath) _
    # & " . Nr. of items removed: " & totalDeletes _
    # & " Time: " & Now(), eLall)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyAttributeStack
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyattributestack():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.CopyAttributeStack"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Attrib As cAttrDsc
    # Dim DestPX As Long
    if sourcePX = 1 Then:
    # DestPX = 2
    else:
    # DestPX = 1
    # Set aID(DestPX).idAttrDict = New Dictionary
    # AttributeUndef(DestPX) = AttributeUndef(sourcePX)
    for attrib in aid:
    # aID(DestPX).idAttrDict.Add Attrib.adKey, Attrib

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub DecodeAllPropertiesFor2Items
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def decodeallpropertiesfor2items():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.DecodeAllPropertiesFor2Items"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim checkPx As Long
    if onlyItemNo < 3 Then:
    # checkPx = onlyItemNo
    else:
    # checkPx = 1
    # With aID(checkPx)
    if .idAttrDict Is Nothing Then:
    # Set .idAttrDict = New Dictionary
    # End With ' aID(checkPx)
    if ContinueAfterMostRelevant Then:
    if isEmpty(MostImportantProperties) Then:
    # AttributeIndex = 0
    else:
    # AttributeIndex = UBound(MostImportantProperties) + 1
    else:
    if DebugMode = Not DebugMode Then               ' UNREACHABLE !!! Here ???:
    # Call initializeComparison                   ' restart here for manual debug
    # Call initializeExcel
    # ContinueAfterMostRelevant = False

    # doDecode:
    if onlyItemNo <> 2 Then:
    # Call SetupAttribs(aID(1).idObjItem, 1, True)
    if onlyItemNo = 1 Then:
    # Call RulesToExcel(1, True)
    else:
    if isEmpty(MostImportantProperties) Then:
    # AttributeIndex = 0
    elif ContinueAfterMostRelevant Then:
    # AttributeIndex = UBound(MostImportantProperties) + 1
    else:
    # AttributeIndex = 0
    # cMissingPropertiesAdded = 0
    # Call SetupAttribs(aID(2).idObjItem, 2, True)
    # stpcnt = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)

    if onlyItemNo <> 1 Then ' do some plausi checks:
    if aID(1).idObjItem.ItemProperties.Count <> aID(2).idObjItem.ItemProperties.Count Then:
    if aID(1).idObjItem.ItemProperties.Count > aID(2).idObjItem.ItemProperties.Count Then:
    if aID(2).idObjItem.Class = aID(2).idObjItem.Parent.Class Then:
    # Set aID(2).idObjItem = aID(2).idObjItem.Parent
    else:
    # Debug.Assert False
    # ' GoTo objecterror
    else:
    if aID(1).idObjItem.Class = aID(1).idObjItem.Parent.Class Then:
    # Set aID(1).idObjItem = aID(1).idObjItem.Parent
    else:
    # Debug.Assert False
    # ' GoTo objecterror
    if aID(1).idObjItem.ItemProperties.Count <> aID(2).idObjItem.ItemProperties.Count Then:
    # Debug.Assert False
    # ' GoTo objecterror

    # YleadsXby = 0 ' we can never have a misadjustment yet

    if onlyItemNo <> 1 Then:
    # MaxPropertyCount = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)
    else:
    # MaxPropertyCount = aID(1).idAttrDict.Count

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Function DoTheDeletes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def dothedeletes():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.DoTheDeletes"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim delItem As cDelObjectsEntry

    if dcCount > 0 Then:
    # Message = "Folgende " & dcCount & " Eintrge knnen " _
    # & killMsg & " werden: " & vbCrLf
    # Set delItem = DeletionCandidates.Items(i)
    # Message = Message & b & delItem.DelObjPindex
    # bDefaultButton = "Go"
    # b1text = Replace(killMsg, "o", "ie", 1, 1)  ' "verschoben" -> "verschieben"
    # b1text = Left(b1text, 10)
    # b2text = "Nicht lschen"
    # Set LBF = Nothing
    # Set LBF = New frmDelParms
    # LBF.Caption = "Besttigung vor dem Lschvorgang (" _
    # & TrashFolder.FullFolderPath & ")"
    # LBF.Show
    # Set LBF = Nothing
    if rsp = vbCancel Then:
    # Call LogEvent(Message, eLall)
    if TerminateRun Then:
    # GoTo ProcReturn
    if rsp = vbYes Then:
    # DoTheDeletes = dcCount
    # Set delItem = DeletionCandidates.Items(i)
    # Call ShowStatusUpdate
    # Call TrashOrDeleteItem(delItem)
    # Message = Replace(Message, "knnen", "wurden")
    # Message = Replace(Message, "werden", vbNullString)
    else:
    # Message = "Die Lschungen wurden nicht besttigt"
    else:
    # Message = "Es wurden keine Lschungen ausgewhlt"

    # Set DeletionCandidates = New Dictionary

    # LListe = vbNullString
    # dcCount = 0
    # Call LogEvent(Message, eLall)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FindTrashFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findtrashfolder():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.FindTrashFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim baseFolder As Folder
    # Set baseFolder = topFolder
    # ' curFolderPath = baseFolder.FolderPath ???***
    # ' FolderPathLevel = 0
    # ' FullFolderPath(FolderPathLevel) = vbNullString
    if getTrashFolder(baseFolder, vbNullString) Is Nothing Then:
    # killType = "Lschungen sind endgltig"
    # killMsg = "gelscht"
    else:
    # killType = "verschiebt in " & TrashFolderPath
    # killMsg = "verschoben in " & TrashFolderPath

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub GetSortableItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getsortableitems():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.GetSortableItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Filter As String
    # Dim i As Long
    # Dim sI As Object
    if eOnlySelectedItems Then:
    # DoVerify False, " debug this, can it work???"
    # Set sortedItems(1) = Nothing
    # Set sI = SelectedItems.Item(i)
    # sortedItems(1).Add sI    ' no date filtering here
    else:
    if getFolderFilter(curFolder.Items(1), CutOffDate, Filter, ">=") _:
    # Then
    # Set sortedItems(1) = curFolder.Items.Restrict(Filter)
    else:
    # Set sortedItems(1) = curFolder.Items
    # Set sI = sortedItems(1).Item(1)
    if sI.Class = olAppointment Then   ' works only if sorted by [Start]:
    # sortedItems(1).IncludeRecurrences = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Function getTrashFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def gettrashfolder():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.getTrashFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim curFolder As Folder
    # Dim DuplikateFehlt As Boolean
    # Dim MatchedFolder As Boolean
    # Dim goingUp As Boolean
    # Dim ShortName As String

    # DuplikateFehlt = False
    # goingUp = False
    if selectedFolder Is Nothing Then:
    # Set selectedFolder = Folder(1)
    if selectedFolder.Parent Is Nothing Then:
    # GoTo multiFoldertable   ' no parent, => no kid-Folders
    try:
        # Do While selectedFolder.Folders.Count = 0
        # goUpOneLevel:
        if selectedFolder.Parent Is Nothing Then:
        # Set selectedFolder = Folder(1)
        if selectedFolder.Parent Is Nothing Then:
        # multiFoldertable:
        # Set selectedFolder = Session.GetDefaultFolder(olFolderDeletedItems)
        if selectedFolder.Parent.Class = aNameSpace Then:
        # Exit Do
        # Set selectedFolder = selectedFolder.Parent
        # Loop

        # tryrelaxed:
        # Call ErrReset(0)
        for curfolder in selectedfolder:
        # ShortName = Left(curFolder.Name, 4)
        # MatchedFolder = (InStr(ShortName, "Dele") > 0 _
        # Or InStr(ShortName, "Gel") > 0 _
        # Or InStr(ShortName, "Tras") > 0)

        if (ShortName = "Dupl" And Not IsComparemode) _:
        # Or DuplikateFehlt And MatchedFolder Then
        # Set TrashFolder = curFolder
        # GoTo gotOne
        else:
        if ShortName = "Sync" Then:
        # GoTo skpToNextFolder
        if DuplikateFehlt And curFolder.Folders.Count > 0 Then:
        # Set getTrashFolder = getTrashFolder(curFolder, FullFolderPath & "\" & curFolder.Name)
        if getTrashFolder Is Nothing Then:
        # goingUp = True
        # GoTo goUpOneLevel ' so we go up one Folder level
        else:
        # GoTo gotOne

        # skpToNextFolder:

        # DuplikateFehlt = Not DuplikateFehlt
        if DuplikateFehlt Then:
        # GoTo tryrelaxed
        if Not goingUp Then:
        # goingUp = True
        # GoTo goUpOneLevel
        # gotOne:
        if Not TrashFolder Is Nothing Then:
        # TrashFolderPath = TrashFolder.FolderPath
        # Set getTrashFolder = TrashFolder

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Initialize_UI
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initialize_ui():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.Initialize_UI"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim otherFolder As Folder
    # Dim j As Long
    # WantConfirmationThisFolder = WantConfirmation
    # MaxMisMatchesForCandidates = MaxMisMatchesForCandidatesDefault
    if LenB(ActionTitle(UBound(ActionTitle))) = 0 Then:
    # Call SetStaticActionTitles
    # Message = vbNullString
    # killType = "Vergleichen, Lschung oder Verschiebung"    ' options available on Item
    # killMsg = "wird noch ermittelt"

    # Message = ActionTitle(ActionID)

    # ' Plausis for "parameters"

    if LF_CurLoopFld Is Nothing Then:
    if SelectedItems Is Nothing Then:
    if ActiveExplorer.Selection.Count = 0 Then:
    # DoVerify False
    # Set SelectedItems = New Collection
    # SelectedItems.Add ActiveExplorer.Selection.Item(j)
    # Set LF_CurLoopFld = getParentFolder(SelectedItems.Item(1))
    if Folder(1) Is Nothing Then:
    # Set Folder(1) = LF_CurLoopFld
    # Set otherFolder = Folder(1)
    if LF_CurLoopFld.FolderPath <> otherFolder.FolderPath Then:
    # DoVerify False
    # Set LF_CurLoopFld = otherFolder
    # curFolderPath = LF_CurLoopFld.FolderPath
    # eOnlySelectedFolder = True
    if eOnlySelectedFolder Then:
    if eOnlySelectedItems Then:
    # DoVerify False, "bad combi"
    # eOnlySelectedFolder = False
    # Set LF_CurLoopFld = Nothing
    # GoTo selOnly
    # Set Folder(1) = LF_CurLoopFld
    elif eOnlySelectedItems Then:
    # selOnly:
    if SelectedItems Is Nothing Then:
    # DoVerify False
    elif SelectedItems.Count = 0 Then:
    # DoVerify False
    if UI_DontUseDel Or Not UI_DontUse_Sel Then:
    # Call LogEvent("Using Standard Deletion Parameters")
    if UI_DontUse_Sel Then                                  ' implies UI_DontUseDel = True:
    # Call LogEvent("Using Standard Selection Parameters")
    else:
    # Call DisplayParameters(2)
    else:
    # Call DisplayParameters(1)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayParameters
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def displayparameters():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.DisplayParameters"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim CountOfItems As Long
    # Dim Pad As Long
    # Dim ReShowFrmErrStatus As Boolean

    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, False)
    # ReShowFrmErrStatus = True

    # Set LBF = Nothing
    # b3text = "Abbruch"
    match form:
        case 1: Set LBF = New frmDelParms:
        case 2: Set LBF = New frmSelParms:
        case _:
    # DoVerify False

    if eOnlySelectedFolder Then:
    # ' just process the entire Selected Folder
    # Message = Message & ": " & Quote(curFolderPath) & "  Anzahl Items: " _
    # & CountOfItems & vbCrLf
    # CountOfItems = LF_CurLoopFld.Items.Count
    elif eOnlySelectedItems Then:
    # CountOfItems = SelectedItems.Count
    # Message = Message & b & CountOfItems _
    # & " selektierte Items in Ordner "
    if DateSkipCount > 0 Then:
    # Message = Message & vbCrLf & "     " & DateSkipCount _
    # & " Items wegen Datumsfilter ausgeschlossen"
    # Pad = Len(Message)
    # Message = Message _
    # & Quote(LF_CurLoopFld.FolderPath)
    else:
    # CountOfItems = LF_CurLoopFld.Items.Count
    # Message = Message & " fr " & Quote(curFolderPath)
    # ' working on 2 different Folders?
    if Not Folder(2) Is Nothing Then    ' yes, 2 Folders, Show msg:
    if LF_CurLoopFld.FolderPath <> Folder(2).FolderPath Then:
    # Message = Message & vbCrLf & "und" _
    # & String(Pad + 10, b) & b _
    # & Quote(Folder(2).FolderPath) & "  (enthlt " _
    # & Folder(2).Items.Count & " Items)"

    if eOnlySelectedItems Then:
    # bDefaultButton = "Go"
    # b1text = bDefaultButton
    # b2text = vbNullString                         ' hidden button
    if form = 1 Then:
    # LBF.Frame2.Visible = False
    # killType = "Vergleich ohne geplante Aktionen"
    # LBF.LPWantConfirmationThisFolder.Visible = False
    # LBF.LPWantConfirmation.Visible = False
    # LBF.Controls("Go").Caption = "Go"
    # LBF.Controls("Go").Default = True
    # LBF.Controls("bDebugStop").Visible = True
    else:
    # bDefaultButton = "Go"
    # b1text = "Prfen"
    # b2text = "bergehen"
    if form = 1 Then:
    # LBF.LPWantConfirmationThisFolder.Visible = True
    # LBF.LPWantConfirmation.Visible = True
    if eOnlySelectedItems Then:
    # Message = SelectedItems.Count _
    # & " selektierte Items werden verglichen"
    # LBF.Frame2.Visible = False
    elif eOnlySelectedFolder Then:
    # ' no ops, message already set
    else:
    # Message = Message & " (Ordner " & LF_recursedFldInx _
    # & " von  " _
    # & LookupFolders.Count & " auf dieser Ebene)"
    # LBF.Frame2.Visible = True
    # LBF.Controls("bDebugStop").Visible = True

    if form = 1 Then:
    if eOnlySelectedItems Or eOnlySelectedFolder Then:
    # LBF.LPAskEveryFolder.Visible = False
    else:
    # LBF.LPAskEveryFolder.Visible = True

    if LF_CurLoopFld.Items.Count > 0 Then:
    # Call ShowOrHideForm(LBF, True)
    else:
    if Not LF_CurLoopFld Is Nothing Then:
    # Call LogEvent("         No items in Folder " _
    # & LF_CurLoopFld.FullFolderPath, eLall)
    if Not LBF Is Nothing Then:
    # Set LBF = Nothing

    # FuncExit:
    if ReShowFrmErrStatus Then:
    # Call ShowOrHideForm(frmErrStatus, True)

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub initializeComparison
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initializecomparison():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.initializeComparison"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Matches = 0
    # cMisMatchesFound = 0
    # MatchPoints(1) = 0
    # MatchPoints(2) = 0
    # IgnoredPropertyComparisons = 0
    # NotDecodedProperties = 0
    # SimilarityCount = 0
    # SuperRelevantMisMatch = False
    # DecodingStatusOK = False
    # mustDecodeRest = False
    # AllItemDiffs = vbNullString
    # DiffsIgnored = vbNullString
    # DiffsRecognized = vbNullString
    # MatchData = vbNullString
    # Message = vbNullString
    # fiBody(1) = vbNullString
    # fiBody(2) = vbNullString
    # OneDiff = vbNullString
    # AttributeIndex = 0
    if aID(1).idAttrDict.Count = 0 Then:
    # AttributeUndef(1) = 0
    # AllPropsDecoded = False
    if aID(2).idAttrDict.Count = 0 Then:
    # AttributeUndef(2) = 0
    # AllPropsDecoded = False
    # MaxPropertyCount = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)

    # Set killWords = Nothing
    # Set killWords = New Collection
    # killWords.Add "*@*"         ' do not compare email adresses in body etc.
    # killWords.Add "*aspx?*"     ' do not compare dynamic HTML in body

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub initializeExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initializeexcel():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.initializeExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if xlApp Is Nothing Then:
    # Call XlgetApp
    # GoTo OisN
    elif O Is Nothing Then:
    # OisN:
    if displayInExcel Then:
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl)
    # ' open default but don't Show it
    elif O.xlTabIsEmpty <> 1 Then     ' excel open, set workbook empty:
    # Call ClearWorkSheet(xlA, O)     ' previous workbook is no longer relevant if there is one
    # O.xHdl = sHdl

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitsForPropertyDecoding
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initsforpropertydecoding():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.InitsForPropertyDecoding"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # saveItemNotAllowed = False
    if aID(1).idAttrDict Is Nothing Then:
    # Set aID(1).idAttrDict = New Dictionary
    # AttributeUndef(1) = 0
    if aID(2).idAttrDict Is Nothing Then:
    # Set aID(2).idAttrDict = New Dictionary
    # AttributeUndef(2) = 0
    if aID(1).idAttrDict.Count > 0 Or _:
    # aID(2).idAttrDict.Count > 0 Then ' decoding has started before:
    if aID(2).idAttrDict.Count = 0 Then  ' *** adjust line len below:
    # mustDecodeRest = doingTheRest _
    # Or Not (AllPropsDecoded Or quickChecksOnly)
    elif aID(1).idAttrDict.Count = aID(2).idAttrDict.Count Then:
    # mustDecodeRest = doingTheRest _
    # Or Not (AllPropsDecoded Or quickChecksOnly)
    else:
    # mustDecodeRest = True
    # UserDecisionRequest = True
    # DoVerify Not (AllPropsDecoded And mustDecodeRest)
    if AllPropsDecoded And mustDecodeRest Then:
    # AllPropsDecoded = False
    # GoTo ProcReturn

    # mustDecodeRest = False    ' Kompletter Neuanfang, initially only most important
    # AllPropsDecoded = False
    # UserDecisionRequest = False
    if aID(1).idAttrDict.Count > 0 Then:
    # Set aID(1).idAttrDict = New Dictionary
    # AttributeUndef(1) = 0
    if aID(2).idAttrDict.Count > 0 Then:
    # Set aID(2).idAttrDict = New Dictionary
    # AttributeUndef(2) = 0

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub logDiffInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def logdiffinfo():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.logDiffInfo"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if Left(Text, 3) = "## " Then:
    # MatchData = MatchData & vbCrLf & Text
    # Call LogEvent(Text)
    else:
    # Call LogEvent(Text, eLmin)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub logMatchInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def logmatchinfo():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.logMatchInfo"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Message = "+++ Resumee: " & objTypName _
    # & " items " & WorkIndex(1) & "/" _
    # & WorkIndex(2) & vbCrLf & Message
    if fiMain(1) <> fiMain(2) Then:
    # Message = Message & vbCrLf _
    # & " +++ Objekte haben unterschiedliche Hauptidentifikationen"
    else:
    # Message = Message & vbCrLf & fiMain(1) & vbCrLf & " +++ Ende "
    # Call LogEvent(Message)
    # Call LogEvent(MatchData, eLnothing)
    # Call LogEvent(String(Len(Message), "="))

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub NoDupes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def nodupes():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "DupeDeleter.NoDupes"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="DupeDeleter")

    # ActionID = atDoppelteItemslschen

    # IsEntryPoint = True

    # ' set dynamic headline
    # sHdl = "CritPropName---------------" _
    # & " Objekt-" & WorkIndex(1) & "----------------------" _
    # & " Objekt-" & WorkIndex(2) & "----------------------" _
    # & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    # AcceptCloseMatches = True
    # quickChecksOnly = Not AcceptCloseMatches
    # AskEveryFolder = True
    # WantConfirmation = True
    # MatchMin = 1000
    # IsComparemode = False
    # SelectOnlyOne = False
    # eOnlySelectedItems = False
    # StopLoop = False
    # PickTopFolder = True
    # Set SelectedItems = New Collection

    # bDefaultButton = "Go"
    # ' look for the best suitable DftItemClass and its rules
    # Call BestObjProps(Folder(1), withValues:=False)
    if eOnlySelectedFolder Then:
    # ' Just do this one Folder / subFolders thereof
    # Set LF_CurLoopFld = Folder(1)
    # Call Initialize_UI              ' displays options dialogue
    # Call CheckOneFolder
    else:
    # Set LF_CurLoopFld = LookupFolders(FolderLoopIndex)
    # Call Initialize_UI          ' displays options dialogue
    # Call CheckOneFolder
    # done:
    if TerminateRun Then:
    # GoTo ProcReturn
    # StopRecursionNonLogged = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub CheckOneFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkonefolder():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.CheckOneFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    match rsp:
        case vbYes:
    if eOnlySelectedItems Then:
    # FolderLoopIndex = LookupFolders.Count    ' do not loop Folder list if Selection
    # DoVerify False, " this code used as model only"
    # GoTo likepicked
    elif PickTopFolder Then:
    # FolderLoopIndex = LookupFolders.Count    ' do not loop Folder list if Folder is picked
    if Folder(1) Is Nothing Then:
    # Call PickAFolder(1, _
    # "bitte whlen Sie den obersten Ordner fr die Doublettensuche ", _
    # "Auswahl des Hauptordners fr die Doublettensuche", _
    # "OK", "Cancel")
    # Set topFolder = Folder(1)
    # likepicked:
    # Call FindTrashFolder
    # Set ParentFolder = Nothing
    # Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
    # ' no parent if top Folder, parentFolder = nothing
    # Set ParentFolder = topFolder.Parent
    # Catch
    # curFolderPath = topFolder.FolderPath
    # FullFolderPath(FolderPathLevel) = "\\" _
    # & Trunc(3, curFolderPath, "\")
    # Call CheckDoublesInFolder(topFolder)        ' ########## Main Work here ##########*
    else:
    # Call CheckDoublesInFolder(topFolder)        ' ########## Main Work here ##########*
    # bDefaultButton = "Go"
        case vbCancel:
    # Call LogEvent("=======> Stopped before processing any Folders . Time: " _
    # & Now(), eLmin)
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo ProcReturn
        case Else   ' loop Candidates:
    # Set topFolder = LookupFolders.Item(FolderLoopIndex)
    # Call FindTrashFolder
    if WantConfirmation = True Then:
    # Call LogEvent("=======> Confirmation mode starts. Time: " _
    # & Now(), eLmin)
    else:
    # Call LogEvent("=======> Confirmation mode ends. Time: " _
    # & Now(), eLmin)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub SaveItemsIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def saveitemsifchanged():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.SaveItemsIfChanged"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim ItemSaved As Boolean
    # Dim Both As String
    if saveItemNotAllowed Then:
    # Message = "Items sollen aufgrund der Nutzerangaben nicht gespeichert werden"
    # Call LogEvent(Message, eLmin)
    else:
    if MustConfirm _:
    # And (WorkItemMod(1) Or WorkItemMod(2)) Then
    if WorkItemMod(1) And WorkItemMod(2) Then:
    # Both = "beiden "
    # LF_ItmChgCount = LF_ItmChgCount + 2
    else:
    # LF_ItmChgCount = LF_ItmChgCount + 1
    if fiMain(1) = fiMain(2) Then:
    # & " besttigen", vbYesNoCancel, ActionTitle(ActionID))
    else:
    # & "  and " & Quote(fiMain(2)) & b, vbYesNoCancel, _
    # ActionTitle(ActionID))
    if rsp = vbNo Then:
    # saveItemNotAllowed = True
    if CurIterationSwitches.SaveItemRequested And Not saveItemNotAllowed Then:
    for i in range(1, 3):
    if WorkItemMod(i) Then:
    # WorkItemMod(i) = False
    # Err.Clear
    # aBugTxt = "save item #" & i & b & fiMain(i)
    # Call Try
    # aID(i).idObjItem.Save
    if Catch Then:
    # Message = "Item changes NOT saved in " _
    # & Quote(aID(i).idObjItem.Parent.FullFolderPath) _
    # & " for " & Quote(fiMain(i)) & ": " _
    # & Err.Description
    # ' NO! this could cause delete query: ItemSaved = False
    else:
    # Message = "Item changes successfully saved in " _
    # & Quote(aID(i).idObjItem.Parent.FullFolderPath) _
    # & " for " & Quote(fiMain(i))
    # ItemSaved = True
    # Call LogEvent(Message, eLall)
    # CurIterationSwitches.SaveItemRequested = ItemSaved  ' report to the outside world: we did save

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub ScanItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def scanitem():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.ScanItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim Propertycount As Long
    # WorkIndex(1) = itemNo
    # OnlyMostImportantProperties = False

    # Set oneItem = GetAobj(1, WorkIndex(1))
    # objTypName = DecodeObjectClass(getValues:=True)
    # Call LogEvent(String(60, "_") & vbCrLf & "Prfung von " & objTypName & b _
    # & curFolderPath & ": " _
    # & WorkIndex(1) _
    # & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))
    # IgnoredPropertyComparisons = 0
    # NotDecodedProperties = 0
    # MatchData = vbNullString
    # Message = vbNullString
    # saveItemNotAllowed = False
    if xUseExcel Or xDeferExcel Or O Is Nothing Then:
    # Call XlgetApp
    # Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl, showWorkbook:=DebugMode)
    if O.xlTabIsEmpty > 1 Then  ' excel open, workbook empty:
    # Call ClearWorkSheet(xlA, O)   ' previous workbook is no longer relevant
    # ' *** Set aItemProps(1) = aID(1).idObjItem.ItemProperties ' done inside GetItemAttrDscs
    # Propertycount = aID(1).idObjItem.ItemProperties.Count
    if TotalPropertyCount = 0 Then:
    # TotalPropertyCount = Propertycount
    elif TotalPropertyCount <> Propertycount Then:
    # DoVerify False, " more analysis needed here."
    # TotalPropertyCount = Propertycount
    # aID(1).idAttrDict = New Dictionary
    # AttributeUndef(1) = 0
    # SelectOnlyOne = True ' not working with tuples
    # Call SetupAttribs(oneItem, 1, True)
    if AttributeIndex = 0 Then:
    # Message = "Skipping item " & WorkIndex(1) _
    # & ", can not process " _
    # & aOD(aPindex).objItemClassName _
    # & " (" & aID(1).idObjItem & ") ," _
    # & Message
    # Call LogEvent(Message, eLall)
    # GoTo ProcReturn ' no can do
    # NotDecodedProperties = Propertycount - AttributeIndex

    if NotDecodedProperties > 0 Then:
    # pArr(1) = "*** es wurden nicht alle Merkmale untersucht"
    # Call addLine(O, Propertycount + 3, pArr)
    # AllItemDiffs = AllItemDiffs & vbCrLf & pArr(1)
    # IgnoredPropertyComparisons = IgnoredPropertyComparisons _
    # + NotDecodedProperties
    # displayInExcel = xUseExcel Or xDeferExcel
    if displayInExcel Or (WorkItemMod(1) And CurIterationSwitches.SaveItemRequested) Then:
    # rsp = vbYes
    # UserDecisionRequest = False
    if displayInExcel Then:
    # Call DisplayWithExcel(vbNullString)
    else:
    # Call DisplayWithoutExcel(vbNullString)
    if rsp = vbCancel Then:
    # Call ClearWorkSheet(xlA, O)     ' erase this one, new one will be started when needed
    # xUseExcel = False
    # rsp = vbNo
    if rsp = vbYes And CurIterationSwitches.SaveItemRequested And Not saveItemNotAllowed Then:
    # Call PerformChangeOpsForMapiItems
    # Call Try                         ' Try anything, autocatch
    # aID(1).idObjItem.Save
    if Catch Then:
    # Message = "Item changes NOT saved for " _
    # & aID(1).idObjItem.Subject _
    # & "   " & Err.Description
    else:
    # Message = "Item changes successfully saved for " _
    # & aID(1).idObjItem.Subject
    else:
    # Message = "Item NOT saved (result of user choice)"
    else:
    # Message = "Item has no changes"
    # Call LogEvent(Message, eLall)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub storeAttribute
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def storeattribute():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.storeAttribute"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim NewAttrValue As String
    # Dim adName As String
    # Dim Properties As ItemProperties
    # Dim selectedProperty As ItemProperty
    # Dim oldValue As String
    # Dim oi As Long

    # With O.xlTSheet
    # NewAttrValue = .Cells(i, sourcecol).Text
    if Left(NewAttrValue, 1) = "#" Then:
    # GoTo ProcReturn
    if Left(NewAttrValue, 1) = " & quote(" Then:
    # NewAttrValue = Mid(NewAttrValue, 2)

    # adName = .Cells(i, 1).Text
    # aBugTxt = "access to list of item properties"
    # Call Try(testAll)
    # ' sinngem: aID(pxindex).idObjItem.ADName = NewAttrValue
    # Set Properties = aID(px).idObjItem.ItemProperties
    if Catch Then:
    # GoTo ProcReturn
    # aBugTxt = "get property " & adName
    # Call Try
    # Set selectedProperty = Properties.Item(adName)
    if Catch Then:
    # & vbCrLf & Err.Description)
    # .Cells(i, promptColumn).Value = "?NAC"
    # GoTo ProcReturn

    if selectedProperty.Name = "Attachments" Then    ' we do an array here!:
    if InStr(NewAttrValue, "ContactPicture") > 0 Then:
    if selectedProperty.Value.Count = 0 Then:
    if px = 1 Then:
    # oi = 2
    else:
    # oi = 1
    # DoVerify False, " incomplete here!"
    # NewAttrValue = cPfad & "Ascher, Christian (Home - 20140219).jpg"
    if LenB(NewAttrValue) = 0 Then:
    # NewAttrValue = InputBox("Type the Filename for the contact:")
    # aBugTxt = "add picture to contact"
    # Call Try                 ' Try anything, autocatch
    # aID(px).idObjItem.AddPicture NewAttrValue
    if Catch Then:
    # GoTo FuncExit
    # WorkItemMod(px) = True
    else:
    # oldValue = selectedProperty.Value    ' not .text, we are using the value from the item
    # aBugTxt = "assign new property value for " _
    # & selectedProperty.Name
    # Call Try                         ' Try anything, autocatch
    # selectedProperty.Value = NewAttrValue
    if Catch Then:
    # & vbCrLf & "   " & Err.Description)
    # .Cells(i, promptColumn).Value = "'(?ERR-Could not assign to item)"
    # GoTo FuncExit
    if selectedProperty.Name = MainObjectIdentification Then:
    # fiMain(px) = NewAttrValue & " [genderter Wert], war " & Quote(oldValue) & b
    # WorkItemMod(px) = True
    # CurIterationSwitches.SaveItemRequested = True
    # saveItemNotAllowed = False
    # rsp = vbYes
    # Call LogEvent("New value for " & adName & " = " _
    # & Quote(NewAttrValue) & " on item " & px, eLall)
    # .Cells(i, 15 + px).Value = Quote(oldValue)  ' from item, not from excel
    # .Cells(i, 15 + px).Interior.ColorIndex = xlColorIndexNone
    # .Cells(i, px + 1).Interior.ColorIndex = 35  ' light green
    # End With ' O.xlTSheet

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Function synchedNames
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def synchednames():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.synchedNames"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # didOneSynch:
    if AttributeIndex > aID(1).idAttrDict.Count Then:
    # Message = "Property " _
    # & Quote(aID(2).idAttrDict.Item(AttributeIndex + YleadsXby).adName) _
    # & "  fehlt im 1. Item"
    # GoTo PropCountError
    if AttributeIndex + YleadsXby > aID(2).idAttrDict.Count Then:
    # Message = "Property " & Quote(aID(1).idAttrDict.Item(AttributeIndex).adName) _
    # & "  fehlt im 2. Item"
    # GoTo PropCountError
    # Set aDecProp(1) = aID(1).idAttrDict.Item(AttributeIndex)
    # Set aDecProp(2) = aID(2).idAttrDict.Item(AttributeIndex + YleadsXby)
    # RrX = aDecProp(1).adFormattedValue        ' ItemPropertyValues(1, AttributeIndex)
    # RrY = aDecProp(2).adFormattedValue
    if Left(RrX, 1) = "# " Then:
    if aDecProp(1).adOrigValDecodingOK Then:
    # RrX = aDecProp(1).adDecodedValue
    if Left(RrY, 2) = "# " Then:
    if aDecProp(2).adOrigValDecodingOK Then:
    # RrY = aDecProp(2).adDecodedValue
    # PropertyNameX = aDecProp(1).adName        ' = PropertyNames(1, AttributeIndex)
    # PropertyNameY = aDecProp(2).adName        ' PropertyNames(2, AttributeIndex + YleadsXby)

    if LenB(PropertyNameX) = 0 Or LenB(PropertyNameY) = 0 Then:
    # Message = "Missing Decoded Property for entry " _
    # & AttributeIndex & " / " _
    # & AttributeIndex + YleadsXby
    if LenB(PropertyNameX) = 0 Then:
    # Message = Message & " Side 1"
    if LenB(PropertyNameY) = 0 Then:
    # Message = Message & " Side 2"
    # GoTo PropCountError

    if PropertyNameX = PropertyNameY Then:
    # i = AttributeIndex
    else:
    # ' first, look for innermost occurrence of PropertyNameY
    if PropertyNameY = aID(2).idAttrDict.Item(i).adName Then:
    # YleadsXby = i - AttributeIndex
    # ' AttributeIndex = i wre falsch!
    # GoTo didOneSynch
    # ' if not found here, try first occ. of PropertynameX on side 1
    if PropertyNameX = aID(1).idAttrDict.Item(i).adName Then:
    # YleadsXby = AttributeIndex - i
    # AttributeIndex = i
    # GoTo didOneSynch
    if i = 0 Then:
    # Message = "keine vergleichbaren Attribute?!"
    # PropCountError:
    if DebugMode Then:
    # DoVerify False
    # saveItemNotAllowed = True
    # Call logMatchInfo
    # synchedNames = True
    if i > 0 Then:
    # Call GetAttrDsc(aID(1).idAttrDict.Item(i).adKey)
    else:
    # Set aTD = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub TrashOrDeleteItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trashordeleteitem():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.TrashOrDeleteItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim delObj As Object
    # Dim GoneObj As Object
    # With DelObjectItem
    if .DelObjInd Then:
    # Set delObj = sortedItems(.DelObjPos).Item(.DelObjPindex)
    else:
    # Set delObj = SelectedItems(.DelObjPos)
    # End With ' DelObjectItem
    # 'On Error Resume Next
    if LenB(TrashFolderPath) = 0 Then:
    # hardDelete:
    # aBugTxt = "deleting Item"
    # Call Try(testAll)
    # delObj.Delete
    if CatchNC Then:
    # GoTo errHandling
    # Message = "Endgltige Lschung war erfolgreich"
    else:
    # Message = "Lschung in den Mlleimer versucht"
    # Set GoneObj = delObj.Move(TrashFolder)
    # errHandling:
    if E_Active.errNumber = -2147467259 Then:
    # & "Wollen Sie " & Quote(delObj) _
    # & "  endgltig lschen?" _
    # & vbCrLf & vbCrLf _
    # & "Wenn Sie die ganze Serie lschen wollen, " _
    # & "brechen Sie jetzt ab und whlen Sie die Eintrge " _
    # & "in der Listenansicht des Kalenders aus." _
    # , vbYesNo, "Fehler beim Verschieben nach " _
    # & TrashFolderPath)
    # GoTo doDecide
    elif Catch Then:
    # & "Wollen Sie " & Quote(delObj) & "  endgltig lschen?", vbOKCancel, _
    # "Fehler beim Verschieben nach " & TrashFolderPath)
    # doDecide:
    if rsp = vbOK Then:
    # Message = Replace(Message, "verschoben", _
    # "endgltig gelscht statt verschoben")
    # GoTo hardDelete
    else:
    # Message = "Endgltige Lschung war nicht erfolgreich"
    else:
    # Message = "Lschung erfolgte in " & GoneObj.Parent.FolderPath

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub UnSelectItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def unselectitems():
    # Dim zErr As cErr
    # Const zKey As String = "DupeDeleter.UnSelectItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim A As Object
    if eOnlySelectedItems Then:
    # Set A = sortedItems(1).Item(i)
    # A.managerName = vbNullString ' clear unused field
    # A.Save

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

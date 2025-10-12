# Converted from MailProcessing.py

# Attribute VB_Name = "MailProcessing"
# Option Explicit

# Dim ProblemHelp As String                        ' benenne was faul ist
# Dim ArchiveFolder As Outlook.Folder
# Dim ArchiveSubFolder As Outlook.Folder
# Dim ConfirmOperation As Boolean
# Dim Abbruch As Boolean
# Dim Filter As String
# Dim ItemCounter As Long
# Dim thisTopFolder As String

# ' Achtung, hier must Du vorgeben, wie es funktionieren soll!
# Const ArchivDateiName As String = "my arch"
# Const ArchivierungLschtOriginal As Boolean = True ' False geht vermutlich nicht... braucht Redemption.DLL
# ' anzahl Tage vor heute (-30) als CutOffDate
# Const MaximalAlter As Long = -30

# ' Das ArchiveByDate aufrufen, z.B in ThisOutlookSession.Application_Startup, dann gehts bei jedem Outlook-Start

# '---------------------------------------------------------------------------------------
# ' Method : Sub ArchiveByDate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def archivebydate():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.ArchiveByDate"
    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # Dim InBoxFolder As Outlook.Folder
    # Dim MainFolder As Outlook.Folder

    try:
        # ConfirmOperation = True                      ' fragt ob OK vor Archivierung
        # Abbruch = False
        # ItemCounter = 0
        # thisTopFolder = StdInboxFolder

        # ProblemHelp = "Datumsfehler"
        # CutOffDate = DateAdd("d", MaximalAlter, Now())
        # ' ============

        # ProblemHelp = "ArchiveFolder bestimmen"
        # ' der Name DEINES Archiv - .PSD knnte anders sein, anpassen !!!
        # Set ArchiveFolder = GetFolderByName(ArchivDateiName, aNameSpace, MaxDepth:=1)
        # ' ===============
        # ' Annahme: es gibt nur ein Konto, also nur einen Posteingang
        if thisTopFolder = StdInboxFolder Then:
        # ProblemHelp = "Inbox Ordner bestimmen"
        # Set InBoxFolder = olApp.Session.GetDefaultFolder(olFolderInbox)
        # Call doArchiveWork(InBoxFolder)
        # ' ==========================
        if Abbruch Then GoTo aBug:
        else:
        # ' Falls Annahme falsch, muss eine Schleife ber die InBoxen erfolgen:
        for mainfolder in anamespace:
        # Set InBoxFolder = GetFolderByName(thisTopFolder, MainFolder)
        if Not InBoxFolder Is Nothing Then:
        # ' Archiv nicht archivieren !
        if InStr(1, InBoxFolder.FolderPath, ArchivDateiName, vbTextCompare) = 0 Then:
        # Call doArchiveWork(InBoxFolder)
        # ' ==========================
        if Abbruch Then GoTo aBug:
        else:
        # DoVerify False, " debugphase only"

        # ' Feddisch
        # ProblemHelp = "Es wurden insgesamt " & CStr(ItemCounter) & " Objekte archiviert"
        print(Debug.Print ProblemHelp)
        # GoTo ProcReturn
        # aBug:
        print(Debug.Print ProblemHelp)
        print(Debug.Print Err.Description)
        # DoVerify False

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub doArchiveWork
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def doarchivework():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.doArchiveWork"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    try:
        # Set topFolder = InBoxFolder.Parent

        # ' auf gehts mit der Archivierung
        # Set ArchiveSubFolder = GetCorrespondingFolder(InBoxFolder, ArchiveFolder)
        # Call DoAllItemsIn(InBoxFolder, ArchiveSubFolder)
        if Abbruch Then GoTo dBug:
        # ' Unterordner
        # Call ArchiveSubFolders(InBoxFolder, ArchiveFolder) ' now process subfolders of InBoxFolder
        if Abbruch Then GoTo dBug:
        # GoTo ProcReturn
        # dBug:
        print(Debug.Print ProblemHelp)
        # Abbruch = True
        # DoVerify False

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoAllItemsIn
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def doallitemsin():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.DoAllItemsIn"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim ItemsHere As Items
    # Dim thisitem As Variant
    # Dim thisItemID As String
    # Dim thisArchiveFolder As Outlook.Folder

    # Set thisArchiveFolder = actArchiveFolder     ' protect against changes

    # ProblemHelp = "archivieren des Ordners '" & actFolder.FolderPath & "'" _
    # & " in den Ordner '" & thisArchiveFolder.FolderPath & "'"
    # Set ItemsHere = RestrictItemsByDate(actFolder, Filter, "<=") ' using CutOffDate
    if ItemsHere Is Nothing Then:
    # ProblemHelp = "Auswahl der items mit Filter " & Filter _
    # & " in '" & actFolder.FolderPath & "' fehlgeschlagen"
    print(Debug.Print ProblemHelp)
    else:
    if actFolder.Items.Count = ItemsHere.Count Then:
    # ProblemHelp = "Auswahl der items mit Filter " & Filter _
    # & " in '" & actFolder.FolderPath _
    # & "' umfasst alle " & ItemsHere.Count _
    # & " Objekte, ist das richtig???"
    if MsgBox(ProblemHelp, vbOKCancel) = vbCancel Then:
    # ProblemHelp = " Operation abgebrochen"
    # GoTo ProcReturn
    else:
    # ProblemHelp = "Auswahl der items mit Filter " & Filter _
    # & " in '" & actFolder.FolderPath & "'" & vbCrLf _
    # & " umfasst " & ItemsHere.Count _
    # & " zu archivierende Objekte, " _
    # & vbCrLf & "ist das plausibel?" _
    # & vbCrLf & "(Weiterhin besttigen: Ja, Nein: diesen Ordner auslassen, Cancel: Abbruch)"
    if rsp = vbCancel Then:
    # ProblemHelp = " Operation abgebrochen"
    # GoTo ProcReturn
    elif rsp = vbNo Then:
    # Abbruch = True
    # GoTo ProcReturn
    else:
    # ProblemHelp = "Archivierung umfasst " & ItemsHere.Count & " Objekte, Kriterien " _
    # & Filter & " aus '" & actFolder.FolderPath
    print(Debug.Print ProblemHelp)
    for thisitem in itemshere:
    if Not thisitem.Saved Then:
    if thisItemID = thisitem.EntryID Then:
    # ProblemHelp = " Duplicate Item " & thisItemID
    # GoTo SkipIt                      ' double entry...
    # thisItemID = thisitem.EntryID
    # aBugTxt = "Save Item " & thisItemID
    # Call Try(testAll)
    # thisitem.Save
    if Catch Then:
    # ProblemHelp = E_Active.Reasoning & ": " _
    # & E_Active.Description
    # aBugTxt = "Fetch Item "
    # Call Try                         ' Try anything, autocatch
    # Set thisitem = aNameSpace.GetItemFromID(thisItemID)
    if Catch Then:
    # Call LogEvent("Item existiert nicht mehr, " _
    # & "evtl. gelscht durch Regel, Virenchecker o..")
    # ProblemHelp = ProblemHelp & E_Active.Reasoning
    # GoTo SkipIt
    # ProblemHelp = vbNullString
    # aBugTxt = "Copy Item to " & thisArchiveFolder.FolderPath
    # Call Try
    # Call CopyItemTo(thisitem, thisArchiveFolder)
    # Catch
    # SkipIt:

    # ProblemHelp = "Archivierung beendet, " & ItemsHere.Count & " Objekte, Zielordner '" _
    # & actFolder.FolderPath & "'"
    print(Debug.Print ProblemHelp)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' Recursive Sub to DoAllItemsIn Subfolders
def archivesubfolders():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.ArchiveSubFolders"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim FolderIndex As Long
    # Dim loopFolder As Outlook.Folder
    # Dim actFolderPath As String
    # Dim EntryProblem As String

    if actFolder.Folders.Count = 0 Then:
    # GoTo ProcReturn
    # EntryProblem = ProblemHelp

    try:
        # actFolderPath = actFolder.FolderPath
        # ProblemHelp = " alle Unterordner von '" & actFolderPath & "' bearbeiten"
        # Set loopFolder = actFolder.Folders.Item(FolderIndex)

        # ArchiveSubFolder = GetCorrespondingFolder(loopFolder, actArchiveFolder)
        if Abbruch Then GoTo boo:

        # ' First, do the items here
        # Call DoAllItemsIn(loopFolder, ArchiveSubFolder)
        # ' then, recurse into local subfolders (if any)
        # Call ArchiveSubFolders(loopFolder, ArchiveSubFolder)

        # ProblemHelp = EntryProblem
        # GoTo ProcReturn
        # boo:
        print(Debug.Print ProblemHelp)
        # DoVerify False

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetCorrespondingFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getcorrespondingfolder():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.GetCorrespondingFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim actFolderPath As String
    # Dim ArcFolderPath As String
    # Dim ArcFolderName As String

    try:
        # actFolderPath = actFolder.FolderPath
        # ' parse this name and replace the front part to the archivefolder
        # ArcFolderName = RTail(actFolderPath, "\", Front:=ArcFolderPath)
        # ArcFolderPath = Replace(actFolderPath, ArcFolderPath, actArchiveFolder.FolderPath)

        # ProblemHelp = " korrespondierenden ArchiveFolder festlegen fr '" & actFolderPath & "'"
        # Set GetCorrespondingFolder = GetFolderByName(ArcFolderName, actArchiveFolder)
        if GetCorrespondingFolder Is Nothing Then:
        # ProblemHelp = " korrespondierenden ArchiveFolder existiert nicht: '" _
        # & ArcFolderPath & "' wird erstellt"
        # Set GetCorrespondingFolder = actArchiveFolder.Folders.Add(actFolder.Name)
        else:
        if GetCorrespondingFolder.FolderPath <> ArcFolderPath Then:
        # ProblemHelp = " korrespondierenden ArchiveFolder '" _
        # & GetCorrespondingFolder.FolderPath & "' bzw.'" & ArcFolderPath _
        # & "' passen nicht zu '" & actFolderPath & "'"
        # GoTo bug
        # GoTo ProcReturn
        # bug:
        # Abbruch = True

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyItemTo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyitemto():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.CopyItemTo"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim myCopiedItem As Object
    # Dim FilterName As String
    # Dim AttrValue As String
    # Dim itemProps As ItemProperties
    # Dim itemProp As ItemProperty

    try:
        # Set myCopiedItem = olApp.CreateItem(olMailItem)
        if Len(Filter) < 4 Then:
        else:
        # FilterName = Mid(Filter, 2, InStr(Filter, "]") - 2)
        # Set itemProps = myItem.ItemProperties
        # Set itemProp = itemProps(FilterName)
        # AttrValue = itemProp.Value
        if IsDate(AttrValue) Then:
        # AttrValue = Format(itemProp.Value, "dd.mm.yyyy hh:mm:ss")
        # ProblemHelp = "  Verschieben (" & FilterName & b & AttrValue & ") '" _
        # & myItem.Subject & "'"
        if ArchivierungLschtOriginal Then:
        # Set myCopiedItem = myItem
        else:
        # ' copy item in the same folder as original.
        # Set myCopiedItem = myItem.Copy           ' it will not work for sources in Exchange Active Sync ???
        # myItem.Delete                            ' delete Original ???
        # Set myItem = Nothing
        # Set aItmDsc.idObjItem = myCopiedItem

        # ' move this copy to TargetFolder (which also deletes Original if ArchivierungLschtOriginal
        # myCopiedItem.Move TargetFolder
        if Not myCopiedItem Is Nothing Then:
        if Not myCopiedItem.Saved Then:
        # ProblemHelp = "  Save Item"
        # Call Try                             ' Try anything, autocatch
        # myCopiedItem.Save
        if Catch Then:
        # ProblemHelp = E_Active.Reasoning & ": " & E_Active.Description
        print(Debug.Print ProblemHelp)
        # Set myCopiedItem = Nothing
        # GoTo fixed
        # ItemCounter = ItemCounter + 1
        # ProblemHelp = CStr(ItemCounter) & ProblemHelp _
        # & " nach '" & myCopiedItem.Parent.FolderPath _
        # & "' OK"
        print(Debug.Print ProblemHelp)
        # Set myCopiedItem = Nothing
        # GoTo fixed
        # nixGutt:
        print(Debug.Print ProblemHelp)
        print(Debug.Print Err.Description)
        # DoVerify False
        # fixed:
        # ProblemHelp = vbNullString

        # FuncExit:
        # Call ErrReset(0)

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function RestrictItemsByDate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def restrictitemsbydate():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.RestrictItemsByDate"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if getFolderFilter(curFolder.Items(1), CutOffDate, Filter, Comparator) Then:
    try:
        # Set RestrictItemsByDate = curFolder.Items.Restrict(Filter)
        else:
        # invalid:
        # Set RestrictItemsByDate = Nothing

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function getFolderFilter
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getfolderfilter():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.getFolderFilter"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim filterValue As String
    # Filter = vbNullString                                  ' we will not filter
    if IsDate(compareTo) Then:
    if compareTo < CDate(BadDate) And compareTo <> "00:00:00" Then:
    # filterValue = Quote1(Format(CStr(compareTo), "yyyy mm dd"))
    else:
    # getFolderFilter = False
    # GoTo ProcReturn
    else:
    # DoVerify False, " not implemented"

    match A.Class:
        case olMail:
    # Filter = "[SentOn]"
    # getFolderFilter = True
        case olAppointment:
    # Filter = "[ReceivedTime]"
    # getFolderFilter = True
        case olReport:
    # Filter = "[CreationTime]"
    # getFolderFilter = True
        case _:
    # DoVerify False, " class not implemented"
    # getFolderFilter = False
    # GoTo ProcReturn

    # aTimeFilter = Replace(Replace(Filter, "[", vbNullString), "]", vbNullString)
    # Filter = Filter & b & Comparator & b & filterValue & b

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function Sender2Contact
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def sender2contact():
    # Const zKey As String = "MailProcessing.Sender2Contact"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    # Dim EmailMatch As String
    # Dim ContactSearch As String
    # Dim i As Long

    # ContactSearch = vbNullString
    for i in range(1, 4):
    # EmailMatch = "[Email" & i & "Address] = '" & EmailAddr & "'"
    if i < 3 Then:
    # ContactSearch = ContactSearch & EmailMatch & " OR "
    else:
    # ContactSearch = ContactSearch & EmailMatch

    # aBugTxt = "Restrict matching contacts for " & ContactSearch
    # Call Try
    # Set Sender2Contact = ContactFolder.Items.Restrict(ContactSearch)
    if Catch Then:
    # GoTo FuncExit
    if DebugMode Then:
    if Sender2Contact.Count = 0 Then:
    print(Debug.Print "there are no Contacts matching email sender " _)
    # & Quote(EmailAddr)
    else:
    print(Debug.Print "there are " & Sender2Contact.Count _)
    # & " Contacts matching email sender " & Quote(EmailAddr)
    print(Debug.Print vbTab & i & vbTab & Sender2Contact.Item(i).Subject)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsUnkContact
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isunkcontact():
    # Const zKey As String = "MailProcessing.IsUnkContact"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    # Dim myMatchingContacts As Items

    # Set myMatchingContacts = Sender2Contact(EmailAddr)
    if myMatchingContacts.Count = 0 Then:
    # IsUnkContact = True
    # Set myMatchingContacts = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' Wenn dem entsprechenden Ordner ein Item hinzugefgt wird,
# ' wird eine der MailProcessing Subs ausgefhrt:
# ' Hier: Anhnge speichern, ggf. Kategorie setzen, Reminder eintragen
# ' Regeln werden in RuleWizard definiert und ausgewertet

# '---------------------------------------------------------------------------------------
# ' Method : Sub CollectItemsToLog
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Collect any (Mail-) items not in LOGGED state
# ' Note   : only up to DeferredLimit added at one time
# '          when all folders were done before, specificIndex :=1
# '---------------------------------------------------------------------------------------
def collectitemstolog():
    # Const zKey As String = "MailProcessing.CollectItemsToLog"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")

    # Const NotLoggedItems = "@SQL=NOT ""urn:schemas-microsoft-com:office:office#Keywords"" LIKE "

    # Dim afolder As Folder
    # Dim curObj As Object
    # Dim LogFolderIndex As Long
    # Dim initialFolderIndex As Long

    if DeferredLimit = 0 Then:
    # ' limit number of items currently not processed because we could have storage problems
    # DeferredLimit = maxDeferredLimit
    # DeferredLimitExceeded = False
    # RestrictCriteriaString = NotLoggedItems & Quote1(LOGGED)
    # ' check against bounds
    if specificIndex > LoggableFolders.Count Or specificIndex < 0 Then:
    # specificIndex = 0               ' we finished previous LoggableFolders, item 1 = items(0)
    # initialFolderIndex = -1         ' require do all LoggableFolders again

    # Set afolder = LoggableFolders.Items(LogFolderIndex)
    # curFolderPath = afolder.FolderPath
    # Set RestrictedItems = Nothing
    # aBugTxt = "Restrict folder # " & LogFolderIndex & b & curFolderPath
    # Call Try(allowNew)                          ' Try anything, autocatch
    # Set RestrictedItems = afolder.Items.Restrict(RestrictCriteriaString)
    # Catch
    # ItemsToDoCount = RestrictedItems.Count
    # Set curObj = RestrictedItems.GetFirst

    # gotCurObject:
    if curObj Is Nothing Then:
    # specificIndex = LogFolderIndex      ' resume with next folder, if any
    # Call LogEvent("* " & ItemsToDoCount _
    # & " Items collected in " & "(" & specificIndex & ") " _
    # & Quote(curFolderPath), eLall)
    # GoTo NextFolder
    # Call DeferredActionAdd(curObj, atPostEingangsbearbeitungdurchfhren, NoChecking:=True)
    if DeferredLimitExceeded Then:
    # specificIndex = LogFolderIndex       ' resume with next folder, if any
    # GoTo finishLater
    # Set curObj = RestrictedItems.GetNext
    # GoTo gotCurObject
    if initialFolderIndex = 0 Then          ' doing just one folder:
    # specificIndex = -1                  ' do all again next time
    # GoTo finishLater                    ' but do not loop further

    # finishLater:
    # StopRecursionNonLogged = True

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FldActions2Do
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose: do action on not LOGGED mail-like items, all Folders in LoggableFolders
# '---------------------------------------------------------------------------------------
def fldactions2do():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.FldActions2Do"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # ' List all items in the Inbox that do NOT have a flag:
    # Dim sActionId As Long
    # Dim skipIndex As Long
    # Dim doneCtr As Long
    # Dim doneTotal As Long

    # ' possibly removed Unbekannt nach Trash ber Regel dieses Namens
    if StopRecursionNonLogged Then:
    # GoTo ProcReturn                 ' === prohibited just now===
    if MailEventsViaRules Then:
    # Call ExecuteDefinedRule(inFolderName:="unerwnscht", RuleName:="unerwnscht")

    if ActionID <> atFindealleDeferredSuchordner Then:
    # sActionId = ActionID            ' save for interruptions
    # ActionID = 0                        ' no user choice
    # skipIndex = -1                      ' start with first LoggableFolders
    # ' and do all of them
    # LF_UsrRqAtionId = atFindealleDeferredSuchordner ' this action (only)

    # EventHappened = False
    # NoEventOnAddItem = True             ' defer all new mail events until we are done
    # ItemsToDoCount = 0
    # frmErrStatus.lblDeferredCount.Caption = "Deferred Count"
    # resumeCollect:
    # loopItmIndex = 0
    # Do
    # loopItmIndex = loopItmIndex + 1
    # DeferredLimitExceeded = False
    # Call CollectItemsToLog(skipIndex)   ' accumulate Deferred from all LoggableFolders
    # StopRecursionNonLogged = True
    # TotalDeferred = Deferred.Count
    # frmErrStatus.fDeferredCount = Deferred.Count & "/" & doneTotal
    # frmErrStatus.lblDeferredCount.Caption = "Doing"
    # Call DoAllDeferred
    # doneCtr = TotalDeferred
    # doneTotal = doneTotal + doneCtr
    if DeferredLimitExceeded Then:
    # Exit Do
    # skipIndex = skipIndex + 1
    if LoggableFolders.Count < skipIndex Then:
    # Exit Do
    # Loop
    # frmErrStatus.fDeferredCount = TotalDeferred & "/" & doneTotal
    if doneCtr > 0 And doneCtr Mod DeferredLimit = 0 Then:
    print(Debug.Print LString("  continuing after Exceeding Deferred Limit", OffObj))
    # skipIndex = -1              ' start with first LoggableFolders, for all
    # GoTo resumeCollect

    # ' do not finish operation here yet if StopRecursionNonLogged = False
    # LF_UsrRqAtionId = sActionId             ' completed
    # ActionID = sActionId
    # SkipedEventsCounter = 0
    # loopItmIndex = 0
    if doneTotal <> 0 Then:
    # frmErrStatus.lblDeferredCount.Caption = "Total done"
    # TotalDeferred = doneTotal

    # FuncExit:
    # Call LogEvent("* FldActions2Do processed " & doneTotal & " Items", eLall)
    # EventHappened = False

    # ProcReturn:
    # Call ProcExit(zErr)
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoAllDeferred
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def doalldeferred():
    # Const zKey As String = "MailProcessing.DoAllDeferred"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQArMode, CallType:=tSub, ExplainS:="MailProcessing")

    # Dim oItem As Object
    # Dim ID As String
    # Dim PrevId As String
    # Dim ThisSubject As String

    # TotalDeferred = 0
    # aPindex = 1                                  ' always!
    # quickChecksOnly = True
    if Deferred.Count > 0 Then:
    # Call N_ShowProgress(CallNr, zErr.atDsc, zErr.atKey, _
    # "#Deferred=" & Deferred.Count, _
    # "Limit#=" & DeferredLimit)
    else:
    # Call N_ShowProgress(CallNr, zErr.atDsc, zErr.atKey, _
    # "no deferred actions", vbNullString)
    # ItemsToDoCount = Deferred.Count

    # While Deferred.Count > 0                     ' Process all that were deferred first
    # TotalDeferred = TotalDeferred + 1
    # PrevId = ID
    # ID = Deferred.Item(1).aoObjID
    try:
        # Call Try(testAll)
        # Set oItem = aNameSpace.GetItemFromID(ID)
        if oItem Is Nothing Then:
        # GoTo gone
        # ThisSubject = oItem.Subject
        # curFolderPath = oItem.Parent.FolderPath
        try:
            # aBugVer = PrevId <> ID
            # DoVerify aBugVer, "TotalDeferred hacking on same item ??? ID=" & ID

            if aID(aPindex) Is Nothing Then:
            # makeAnother:
            # Set aID(aPindex) = New cItmDsc       ' uses NewObjectItem
            # Call aID(aPindex).SetDscValues(oItem, withValues:=False)
            else:
            if aOD(aPindex).objItemClass <> oItem.Class Then:
            # GoTo makeAnother

            # Call DefObjDescriptors(oItem, aPindex, _
            # withValues:=False, _
            # withAttributeSetup:=False)
            if oItem Is Nothing Then:
            # gone:
            # Call LogEvent("**** Item #" & TotalDeferred & " ID=" & ID & b _
            # & ThisSubject _
            # & " nicht mehr vorhanden, es verbleiben " _
            # & ItemsToDoCount, eLall)
            # ItemsToDoCount = ItemsToDoCount - 1
            elif aObjDsc.objIsMailLike Then:
            # Call LogEvent("==== Processing #" _
            # & Deferred.Count & "/" & ItemsToDoCount _
            # & " deferred Item " & TotalDeferred _
            # & " in '" & curFolderPath _
            # & "' ID=" & ID & vbCrLf & String(5, b) _
            # & aObjDsc.objTimeType & "=" _
            # & aItmDsc.idTimeValue & b & oItem.Subject, eLall)
            # ActionID = atPostEingangsbearbeitungdurchfhren

            # Call CopyToWithRDO(oItem, FolderAggregatedInbox, aObjDsc)
            # Call DoOneItm(oItem)
            # Call N_ClearAppErr
            else:
            # Call LogEvent("**** Skipping #" & TotalDeferred & " ID=" & ID _
            # & " of " & ItemsToDoCount _
            # & " deferred Items because it is not mail-like: " _
            # & aObjDsc.objTypeName)
            # DoVerify Not DebugMode, "** analyze mail-like Attribute "
            # ItemsToDoCount = ItemsToDoCount - 1
            if Deferred.Count > 0 Then:
            # Deferred.Remove 1
            if ItemsToDoCount > 0 Then:
            # ItemsToDoCount = ItemsToDoCount - 1
            # frmErrStatus.fDeferredCount = Deferred.Count & "/" & ItemsToDoCount
            # Wend
            # DoVerify Deferred.Count = 0, " all done"
            if TotalDeferred > 0 Then:
            # Call LogEvent(TotalDeferred & " neue " _
            # & Quote(SpecialSearchFolderName & b & NLoggedName) _
            # & "  Mail-Eingnge verarbeitet", eLall)
            else:
            # Call LogEvent(" keine " & Quote(SpecialSearchFolderName _
            # & b & NLoggedName) _
            # & "  Mail-Eingnge verarbeitet", eLall)
            # Set oItem = Nothing

            # FuncExit:
            # quickChecksOnly = False

            # ProcReturn:
            # Call ProcExit(zErr)

            # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IsMailLike
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def ismaillike():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "MailProcessing.IsMailLike"
    # Dim zErr As cErr

    # Set CurrentSessionEmail = Nothing
    # Set CurrentSessionReport = Nothing
    # Set CurrentSessionMeetRQ = Nothing
    # Set CurrentSessionTaskRQ = Nothing

    match (nlItem.Class):
        case olMail:
    # Set CurrentSessionEmail = nlItem
    # IsMailLike = True
        case olReport:
    # Set CurrentSessionReport = nlItem
    # IsMailLike = True
        case olMeeting:
    # Set CurrentSessionMeetRQ = nlItem
    # IsMailLike = True
        case olTaskRequest:
    # Set CurrentSessionTaskRQ = nlItem
    # IsMailLike = True
        case _:
    if (nlItem.Class >= olMeetingRequest _:
    # And nlItem.Class <= olMeetingResponseTentative) _
    # Then ' range class: 53-57
    # Set CurrentSessionMeetRQ = nlItem
    # IsMailLike = True
    else:
    # DoVerify False, "can't map to CurrentSession Class"
    # IsMailLike = False

    # ProcReturn:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoOneItm
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def dooneitm():
    # Const zKey As String = "MailProcessing.DoOneItm"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="DoOneItm")

    # Dim pFolder As Folder

    # Set aID(aPindex).idAttrDict = New Dictionary ' invalidate previous decoding results
    # Set aDecProp(aPindex) = Nothing
    if aObjDsc.objIsMailLike Then:
    # ' if the item is reported to be in the search Folder,
    # ' it has no pointer to parent.parent
    # ' so we need to find where it really is
    # Set pFolder = getParentFolder(nlItem)
    if pFolder Is Nothing Then:
    # Call LogEvent("??? a deferred item to process is no longer available", eLSome)
    else:
    # ItemInIMAPFolder = getAccountType(pFolder.FolderPath, aAccountTypeName) = olImap
    if InStr(nlItem.Parent.FullFolderPath, "Suchordner") > 0 Then:
    # Set nlItem = aNameSpace.GetItemFromID(nlItem.EntryID)
    if InStr(nlItem.Parent.FullFolderPath, "Suchordner") > 0 Then:
    # DoVerify False, " that should never happen"
    # Set ParentFolder = nlItem.Parent     ' now !never! = Nothing
    # Call DoMailLike(nlItem)
    else:
    if DebugMode Or DebugLogging Then:
    # Call LogEvent("---- no log categories will be assigned for object of type " _
    # & TypeName(nlItem), eLall)
    # DoVerify False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoMailLike
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def domaillike():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.DoMailLike"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim aSubject As String
    # Dim NameOrMsg As String
    # Dim SourceFolder As Folder
    # Dim SourceFolderPath As String
    # Dim saveXldefer As Boolean
    # Dim TargetFolder As Folder
    # Dim TargetFolderPath As String
    # Dim otherChanges As Boolean
    # Dim DontMove As Boolean
    # Dim IsInTarget As Boolean
    # Dim Multiple As Long
    # Dim PushEventMode As Boolean
    # Dim nItem As Object
    # Dim asNonLoopFolder As Boolean
    # Dim oID As String
    # Dim nID As String
    # Dim UnReadString As String
    # Dim oldItemCategories As String
    # Dim ReProcessCategories As Boolean
    # Dim DelAfterCopy As Boolean
    # Dim i As Long

    # aNewCat = vbNullString
    # aSubject = oItem.Subject

    if Not (LF_DontAskAgain Or oItem.Parent Is Nothing) Then:
    # asNonLoopFolder = NonLoopFolder(oItem.Parent.Name)
    if asNonLoopFolder Then:
    # Call LogEvent("<======> skipping item action " & Quote(ActionTitle(ActionID)) _
    # & vbCrLf & " because it is inFolder " & Quote(oItem.Parent.FolderPath) _
    # & vbCrLf & " loop item " & WorkIndex(1) _
    # & " Time: " & Now(), eLall)
    # GoTo ProcReturn                      ' this is an item we should not be looping

    # PushEventMode = NoEventOnAddItem

    # MailModified = False
    # saveXldefer = xDeferExcel
    # xDeferExcel = xUseExcel
    # Call N_ClearAppErr

    # Set oItem = ReGet(oItem, oID, nID)
    if oItem Is Nothing Then:
    # GoTo flushBrake
    # oldItemCategories = oItem.Categories
    if DebugMode Then:
    if oID = nID Then:
    print(Debug.Print "     the item re-get did not change EntryID=" & oID)
    print(Debug.Print "     Item categories: " & Quote(oldItemCategories))
    else:
    print(Debug.Print "     the item re-get changed EntryID=" & oID)
    print(Debug.Print "     Item categories: " & Quote(oldItemCategories))
    print(Debug.Print "                       to    EntryID=" & nID)
    print(Debug.Print "     Item categories: " & Quote(nItem.Categories))
    # DoVerify False

    # Set SourceFolder = getParentFolder(oItem)
    if SourceFolder Is Nothing Then:
    # DoVerify False
    # GoTo flushBrake
    for i in range(1, 6):
    # SourceFolderPath = SourceFolder.FolderPath
    if LenB(SourceFolderPath) > 0 Then:
    # Exit For
    # DoVerify LenB(SourceFolderPath) > 0, "error getting FolderPath, tries=" & i - 1

    if InStr(1, SourceFolderPath, "SUCHORDNER\", vbTextCompare) > 0 Then:
    # DoVerify False                           ' can't find true Folder for search Folder??? ***"
    # aBugTxt = "get body of item"
    # Call Try                                     ' Try anything, autocatch
    # NameOrMsg = oItem.Body                       ' check if item still exists
    if Catch Then:
    # Call LogEvent("Item kann nicht mehr gefunden werden")
    # GoTo flushBrake

    # Call LogEvent("==== " & Time() & b & TypeName(oItem) & b & oItem.Subject _
    # & " in " & SourceFolderPath _
    # & "    (" & Deferred.Count & " Unbearbeitete in Suchordner " _
    # & Quote(SpecialSearchFolderName) & ")", eLall)

    if InStr(1, oldItemCategories, LOGGED, vbTextCompare) = 0 Then:
    # Call LogEvent("---- NotLogged " & TypeName(oItem) & ": " _
    # & oItem.Subject, eLall)
    elif CurIterationSwitches.ReProcessDontAsk Then:
    # GoTo Auto
    elif CurIterationSwitches.ReprocessLOGGEDItems Then:
    if Not CurIterationSwitches.ReProcessDontAsk Then:
    # & " with Categories: " _
    # & Quote(oldItemCategories), vbYesNo, "Besttigung")
    if rsp <> vbNo Then:
    # Call LogEvent("---- This item IS reprocessed by user choice: " & aSubject _
    # & ", Categories: " & Quote(oldItemCategories), eLall)
    # ReProcessCategories = True
    else:
    # GoTo dontProcess
    else:
    # Auto:
    # Call LogEvent("---- Automatically re-processed: " & aSubject, eLall)
    # GoTo doprocess
    if eOnlySelectedItems Then:
    # ReProcessCategories = True
    # GoTo doprocess
    else:
    # dontProcess:
    # Call LogEvent("---- This item not reprocessed: " _
    # & aSubject _
    # & ", Categories: " & Quote(oldItemCategories), eLall)

    # GoTo Epilog

    # doprocess:
    # LF_ItmChgCount = LF_ItmChgCount + 1

    # ' Preconditions for Move to Target:
    # ' PosteINgang, ERHALTEN, INbox, SMS (cond.), GES,  UNB, UNK,  but NOT SENt
    if InStr(1, SourceFolderPath, "SEN", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderSent
    # GoTo Sent

    elif InStr(1, SourceFolderPath, "IN", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderInbox
    # GoTo Received

    elif InStr(1, SourceFolderPath, "SMS", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderSMS             ' Phone number not visible
    # GoTo Received

    elif InStr(1, SourceFolderPath, "UNB", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderUnknown
    # GoTo Received

    elif InStr(1, SourceFolderPath, "UNK", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderUnknown
    # GoTo Received

    elif InStr(1, SourceFolderPath, "ERHALTEN", vbTextCompare) > 0 Then:
    # Set TargetFolder = FolderInbox
    # GoTo Received
    if TargetFolder Is Nothing Then:
    # Set TargetFolder = FolderInbox
    if Not (LF_DontAskAgain Or isNonLoopFolder) Then:
    # DoVerify False
    # Received:
    if oItem.Class = olReport Then               ' Cases: may not have SenderName:
    # Call LogEvent("    -- " & oItem.Body, eLall)
    elif oItem.Class = olTaskRequest Then:
    # Call LogEvent("    -- Antwortstatus NewMail", eLall)
    else:
    # Call LogEvent("    -- from " & Quote(oItem.SenderName) _
    # & " (" & oItem.SenderEmailAddress & ")" _
    # & vbCrLf & "     -- received " & oItem.ReceivedTime _
    # & ", sent on " & oItem.SentOn _
    # & ", created on " & oItem.CreationTime, eLall)
    # GoTo DoChanges
    # Sent:
    if oItem.Class = olMail Then:
    # Call LogEvent("    -- Sent to " & oItem.To, eLall)
    elif oItem.Class = olReport Then:

    elif oItem.Class >= olMeetingRequest _:
    # And oItem.Class <= olMeetingResponseTentative Then
    # Call LogEvent("    Associated Meeting with " _
    # & oItem.Recipients.Count & " recipients", eLall)
    else:
    # aBugTxt = "get attachment count " & NameOrMsg
    # Call Try
    # NameOrMsg = TypeName(oItem)
    # Catch
    # NameOrMsg = NameOrMsg & " has " & oItem.Attachments.Count & " attachments"

    # DoChanges:
    # NameOrMsg = vbNullString
    # CategoryDroplist = LOGGED & "; Aktuell; "       ' always dropped
    # aNewCat = DetectCategory(TargetFolder, oItem, NameOrMsg) ' may change TargetFolder
    # '   (eg. if sender is unknown)!

    # Call setItmCats(oItem, aNewCat, CategoryDroplist)
    if aNewCat <> oldItemCategories Then:
    if DebugMode Then:
    print(Debug.Print "about to change the Item's Categories to " & Quote(aNewCat))
    if FolderUnknown Is Nothing Then:
    # 'Stop 'checkme
    else:
    if InStr(aNewCat, Unbekannt) = 0 Then:
    if TargetFolder.FolderPath = FolderUnknown.FolderPath Then:
    # Set TargetFolder = FolderInbox ' special for items with contact now known:
    # DelAfterCopy = True          ' copy to FolderInbox then remove from Unknown
    elif SourceFolder.FolderPath = FolderUnknown.FolderPath Then:
    # Set TargetFolder = FolderInbox ' special for items with contact now known:
    # DelAfterCopy = True          ' copy to FolderInbox then remove from Unknown
    # TargetFolderPath = TargetFolder.FullFolderPath

    if TargetFolderPath = SourceFolderPath Then:
    # DoVerify Not otherChanges, " OtherChanges remains False"
    # DontMove = True
    # IsInTarget = True
    else:
    # DontMove = False

    # Set nItem = CopyItm2Trg(oItem, _
    # TargetFolderPath, _
    # SourceFolderPath, _
    # TargetFolder)
    if DelAfterCopy Then:
    # oItem.Delete
    # Call LogEvent("     original item deleted from folder " & SourceFolderPath, eLall)
    # Set aItmDsc.idObjItem = Nothing
    # Set oItem = Nothing
    elif oItem.Categories <> nItem.Categories Then:
    # oItem.Categories = nItem.Categories
    # oItem.Save
    if RestrictedItemCollection.Count > 1 Then:
    # Multiple = RedceRestrColl(DontMove, _
    # otherChanges, _
    # RemoveEmailItems:=True)
    # Set nItem = RestrictedItemCollection(1)
    # Set RestrictedItemCollection = New Collection ' un-use the remaining Item
    # Call DoChng2Item(nItem, TargetFolder)
    # Call GenerateTaskReminder(nItem)             ' Erinnerung erstellen (ggf.)

    # flushBrake:
    # NameOrMsg = Err.Description
    if ErrorCaught = Hell Then:
    # GoTo ProcReturn
    if CatchNC(HandleErr:=-2147221233) Then:
    if nItem Is Nothing Then:
    # GoTo Epilog
    if InStr(NameOrMsg, "kann nicht gefunden werden") = 0 Then:
    # & Quote(nItem.Subject))
    # Call LogEvent("**** Fehler: " & NameOrMsg & b _
    # & TypeName(nItem) & ": " & Quote(nItem.Subject), eLall)

    # ' Original gets same changes as nItem
    if InStr(UCase(nItem.Parent.Name), "SEN") = 0 Then:
    # nItem.UnRead = True                      ' may or may not cause change
    # UnReadString = " UnRead"
    else:
    # nItem.UnRead = False                     ' usually causes a change
    # UnReadString = " Read"
    if oItem Is Nothing Then:
    # Set aItmDsc.idObjItem = nItem            ' aObjdDsc of oItem can not be cloned, replace
    # UnReadString = " original item moved to " & Quote(TargetFolderPath) & UnReadString
    else:
    if Not oItem.Saved Then                  ' should be saved before:
    # Call ForceSave(oItem, "(old) ")
    # Call LogEvent("     " & Quote(SourceFolderPath) & UnReadString _
    # & " + Categories set to " _
    # & Quote(nItem.Categories), eLall)
    if Not nItem.Saved Then                      ' max have been changed if categories changed:
    # nItem.Save
    if Not nItem.Saved Then                      ' catch problem with previous Save attempt:
    # Call ForceSave(nItem, "(new) ")
    # Epilog:
    # Call N_ClearAppErr
    # Set RestrictedItemCollection = New Collection
    # xDeferExcel = saveXldefer
    if Not PushEventMode Then:
    # Call RestEvn4Item
    # NoEventOnAddItem = PushEventMode
    # Set oItem = Nothing
    # Set nItem = Nothing

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function Replicate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def replicate():
    # Const zKey As String = "MailProcessing.Replicate"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    # Dim SourceFolder As Folder
    # Dim NewRef As Object
    # Dim oItemID As String
    # Dim nItemID As String

    # Call SetEventMode                            ' MASK OUT AddItemEvent inner processing
    # ' no test on NoEventOnAddItem, always replicate

    # ' copy and forward, do not work sometimes (???)
    if DebugLogging Then DoVerify False:
    # oItemID = Item.EntryID
    # Set Item = Nothing

    # aBugTxt = "get item from ID " & oItemID
    # Call Try                                     ' Try anything, autocatch
    # Set NewRef = aNameSpace.GetItemFromID(oItemID)
    if Catch Then:
    # Set Replicate = Nothing
    # GoTo skipThis

    # Set SourceFolder = getParentFolder(NewRef)
    # Set Replicate = CopyToWithRedemption(NewRef, SourceFolder, trySave:=False, NewObjDsc:=aObjDsc)
    # '               ====================

    if isEmpty(Replicate) Or Replicate Is Nothing Then:
    # Set Replicate = Nothing
    else:
    if Replicate.Subject <> Item.Subject Then:
    # Replicate.Subject = Item.Subject     ' no FWD: or WG:
    # Replicate.Body = Item.Body
    # Replicate.HTMLBody = Item.HTMLBody
    # Call ShowStatusUpdate
    if delOriginal And Not Replicate Is Nothing Then:
    # nItemID = Replicate.EntryID
    if oItemID = nItemID Then DoVerify False, "shit":
    # Set Item = Nothing
    # Set NewRef = Nothing

    # aBugTxt = "Replicate item"
    # Call Try("Die angegebene Nachricht kann nicht gefunden werden.")
    # Set NewRef = aNameSpace.GetItemFromID(oItemID)
    if Catch Then:
    # Call LogEvent("Delete original not needed because item already gone")
    # GoTo skipThis
    if Not NewRef.Saved Then:
    # DoVerify False, " shit"
    # Call ShowDbgStatus                   ' if pending error, set up frmErrStatus
    # GoTo skipThis
    # aBugTxt = "delete original item " & Quote(NewRef.Subject)
    # Call Try
    # NewRef.Delete
    # Catch

    # skipThis:
    # Set NewRef = Nothing
    # Call RestEvn4Item                            ' AddItemEvent was not triggered by MASK OUT

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CopyItm2Trg
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Function CopyItm2Trg(Item As Object, TargetFolderPath As String, SourceFolderPath As String, TargetFolder As Folder, Optional withDupeCheck As Double = -1#) As Object
# Const zKey As String = "MailProcessing.CopyItm2Trg"
# Dim zErr As cErr
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

# ' withDupeCheck = 0: dont check for dupes
# '               = +1: check exact criteria before copyTo
# '               else  check with criteria in time window

# Dim DontMoveItem As Boolean
# Dim otherChanges As Boolean
# Dim IsInTarget As Boolean
# Dim Multiple As Long
# Dim oldClass As OlObjectClass

if TargetFolder Is Nothing Then:
# DoVerify False
# LogicTrace = "*"                             ' start a new logic trace here

if Item Is Nothing Then:
# GoTo ProcReturn

# DontMoveItem = MailEventsViaRules            ' changed later if not in TargetFolder

if ItemInIMAPFolder Then:
# MailModified = False                     ' item.saved irrelevant for Imap Items
else:
# MailModified = Not Item.Saved
# oldClass = Item.Class

# Set CopyItm2Trg = Item
if TargetFolderPath = SourceFolderPath Then:
# DoVerify Not otherChanges, " OtherChanges remains False"
# DontMoveItem = True
# IsInTarget = True
else:
# DontMoveItem = False

# ' withDupeCheck = 1: do duplicateChecking before save attempt
# '              = -1: do duplicateChecking when save is needed (further below)
if withDupeCheck = 1# Then           ' check if we already have exactly same mail in TargetFolder:
# Call findUniqueEmailItems(CopyItm2Trg, TargetFolder, _
# GetFirstOnly:=False, _
# howmany:=Multiple, _
# maxTimeDiff:=withDupeCheck)

# ' potential reasons for saving item before changing anything
# LogicTrace = "Modified=" & MailModified _
# & " otherchanges=" & otherChanges _
# & " dontmoveitem=" & DontMoveItem _
# & " saved=" & CopyItm2Trg.Saved _
# & vbCrLf
if MailModified _:
# Or otherChanges _
# Or Not DontMoveItem _
# Or Not CopyItm2Trg.Saved Then
# Set CopyItm2Trg = CopyToWithRedemption(CopyItm2Trg, TargetFolder, True, CopiedObjDsc)
# '                 ====================
# MailModified = Not CopyItm2Trg.Saved
if MailModified Then             ' $$$ impossible:
# DoVerify False
# CopyItm2Trg.Save
# MailModified = Not CopyItm2Trg.Saved
# DoVerify Not MailModified
if CopyItm2Trg.Parent.FolderPath = TargetFolder.FolderPath Then:
# IsInTarget = True
# LogicTrace = LogicTrace & " IsInTarget=" & IsInTarget & vbCrLf

if withDupeCheck <> 0 Then:
if withDupeCheck <> 1# Then      ' check if we already have exactly same mail in TargetFolder:
# Call findUniqueEmailItems(CopyItm2Trg, TargetFolder, _
# GetFirstOnly:=False, _
# howmany:=Multiple, _
# maxTimeDiff:=withDupeCheck)
if RestrictedItemCollection.Count >= 1 Then ' in target or some multiple duplicates left:
if CopyItm2Trg.EntryID = RestrictedItemCollection(1).EntryID Then:
# ' no reordering
else:
# Set CopyItm2Trg = RestrictedItemCollection(1)
# DontMoveItem = True          ' it was successfully copied to target before
elif RestrictedItemCollection.Count < 1 Then ' not in target or some multiple duplicates left?!:
# DoVerify withDupeCheck = -1# Or Not DebugMode, " $$$ early debug only"

# LogicTrace = LogicTrace _
# & " dontmoveitem=" & DontMoveItem _
# & " IsInTarget=" & IsInTarget & vbCrLf
if Not (DontMoveItem Or IsInTarget) Then ' Here we do the CopyTo:
# Set CopyItm2Trg = CopyToWithRedemption(CopyItm2Trg, TargetFolder, trySave:=False, NewObjDsc:=CopiedObjDsc)
# '                         ====================

# CopyToFinished:
# doMarkDone:
if Not CopyItm2Trg.Saved Then        ' strange if not:
# Call ForceSave(CopyItm2Trg)
# MailModified = False
if CopyItm2Trg.Class <> oldClass Then:
# DoVerify False

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function RedceRestrColl
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def redcerestrcoll():
    # Const zKey As String = "MailProcessing.RedceRestrColl"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    # Dim i As Long
    # Dim pFolder As String

    if RestrictedItemCollection.Count > 0 Then   ' is a multiple:
    # DontMoveItem = True                      ' no need to move, it's there already
    if RestrictedItemCollection.Count > 1 Then ' too many multiple::
    # Call Try(allowAll)                      ' Try anything, autocatch, Err.Clear
    # pFolder = Quote(RestrictedItemCollection(i).Parent.FullFolderPath)
    # RestrictedItemCollection(i).Delete
    # Catch
    if RemoveEmailItems Then:
    # RestrictedItemCollection.Remove i ' remove email in target Folder
    # Call LogEvent("     " & i & ": sufficiently similar item found, deleted " _
    # & "from Collection and " & pFolder, eLall)
    else:
    # Call LogEvent("     " & i & ": sufficiently similar item found, deleted " _
    # & "from Collection", eLall)
    # DontMoveItem = True                      ' no need to move, it's there already
    else:
    # DontMoveItem = False                     ' move because rule did not move it
    # otherChanges = True
    # RedceRestrColl = RestrictedItemCollection.Count

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function DoChng2Item
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def dochng2item():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.DoChng2Item"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # ' Beginning changes on the Target
    # Dim iAttachCnt As Long
    # Dim j As Long
    # Dim k As Long
    # Dim JFirst As Long                           ' First Attachment corrected by deletes
    # Dim mailAttachments As Variant
    # Dim thisAttachment As Attachment
    # Dim Absender As String
    # Dim attachmentFile As String
    # Dim HeaderFirst As Boolean
    # Dim InsertText As String                     ' double byte character set for HTML
    # Dim bodyHTMLtext As String
    # Dim OriginalIsNotHTML As Boolean
    # Dim TargetShouldChangeToHTML As Boolean
    # Dim DoNotDeleteAttachment As Boolean         ' ??? currently not used
    # Static DontCareFileNames As Variant

    # Dim iClass As Long

    # ' init defaults
    # DoChng2Item = vbOK
    if isEmpty(DontCareFileNames) Then           ' One time only::
    # DontCareFileNames = split(IgnoredAttachmentNames, b)

    # ' find out if we need to/can change to HTML
    # aBugTxt = "Get Class of Item"
    # Call Try
    # iClass = Item.Class
    if Catch Then:
    # DoVerify False, "if no hit, remove permit and testing ???"
    if iClass = olMail Then:
    # aBugTxt = "Get folder path of Item"
    # Call Try
    if InStr(UCase(Item.Parent.FullFolderPath), "SMS") = 0 Then:
    # j = -1                               ' JUST for errors:
    # j = Len(Item.HTMLBody) > 0 Or Len(Item.Body) > 0
    # ' we have some sort of body if j < 0
    # OriginalIsNotHTML = (Item.BodyFormat <> olFormatHTML) And j < 0
    if OriginalIsNotHTML Then:
    # TargetShouldChangeToHTML = True  ' if allowed, do it
    else:
    # TargetShouldChangeToHTML = False ' no change needed here
    # Catch
    else:
    # OriginalIsNotHTML = True                 ' only Mails can be Html
    # TargetShouldChangeToHTML = False

    # ' do change to format
    if TargetShouldChangeToHTML _:
    # And SetAllToHTML Then
    # aBugTxt = "set body format to HTML"
    # Call Try
    # Item.BodyFormat = olFormatHTML
    # Catch
    # MailModified = True                      ' Relevant mod done

    if iClass = olMail Then:
    if InStr(1, Item.ReceivedByName, "Marc", vbTextCompare) > 0 Then:
    if Len(Item.HTMLBody) > 1000 Then:
    # Debug.Assert False
    # bodyHTMLtext = Item.HTMLBody

    # Absender = Item.SenderName
    if InStr(Absender, "@") > 0 Then:
    if InStr(Item.SenderEmailAddress, "@") > 0 Then:
    # Absender = Item.SenderEmailAddress
    # Absender = Mid(Absender, 1, InStr(Absender, "@") - 1)
    if InStr(Absender, " PH/DE") > 0 Then ' (Marc):
    # Absender = Replace(Absender, " PH/DE", vbNullString)
    if InStr(Absender, "/") > 0 Then     ' (Marc):
    # Absender = Replace(Absender, "/", vbNullString)

    # HeaderFirst = True
    # aBugTxt = "get mail attachments"
    # Call Try
    # Set mailAttachments = Item.Attachments
    # Catch

    if mailAttachments Is Nothing Then:
    # iAttachCnt = 0
    else:
    # iAttachCnt = mailAttachments.Count
    # Call LogEvent("     " & TypeName(Item) _
    # & " has " & iAttachCnt & " attachments", eLall)

    # JFirst = 1
    # attachmentFile = "     Anhang #" & JFirst & " wurde entfernt (Virus?) und " _
    # & "konnte nicht gespeichert werden."
    # ' attachment with virus may have been deleted by now...
    # aBugTxt = "get mail attachment " & JFirst
    # Call Try
    # Set thisAttachment = mailAttachments.Item(JFirst) ' always store first remaining attachment
    if Catch Then:
    # Call N_ClearAppErr
    # GoTo AttachmentDone
    if thisAttachment.Type = olEmbeddeditem Then:
    # Call LogEvent("     not trying to save Embeddet Attachment " & JFirst _
    # & b & Quote(thisAttachment.FileName), eLall)
    # JFirst = JFirst + 1
    # GoTo AttachmentDone

    # k = ArrayMatch(DontCareFileNames, thisAttachment.FileName)
    # DoNotDeleteAttachment = k > -1
    if Item.Class = olMail Then:
    # DoVerify thisAttachment.Type <> 0
    # Call Try
    # attachmentFile = thisAttachment.FileName
    if Catch(True, "unable to access file name for attachment " & JFirst & b & Item.Name) Then:
    # JFirst = JFirst + 1
    # GoTo AttachmentDone

    if InStr(UCase(Item.HTMLBody), UCase(attachmentFile)) > 0 Then:
    if Not OriginalIsNotHTML Then:
    if Not SaveAttachmentMode Then:
    # Call LogEvent("     Attachment no. " & JFirst & _
    # " NOT saved as attachment " & attachmentFile _
    # & " because it is part of HTML body")
    # JFirst = JFirst + 1
    # GoTo AttachmentDone
    if DoNotDeleteAttachment Then:
    # Call LogEvent("      irrelevant attachment name " & JFirst & ": " _
    # & thisAttachment.FileName, eLmin)
    # JFirst = JFirst + 1
    # GoTo AttachmentDone
    else:
    # attachmentFile = aPfad & TargetFolder.Name _
    # & "\" & DateId & b & _
    # ReFormat(Absender, ".\/?*", b, b) _
    # & " - " & thisAttachment.FileName

    # aBugTxt = "Save mail attachment " & JFirst & " to " & attachmentFile
    # Call Try
    # thisAttachment.SaveAsFile attachmentFile
    if Catch Then:
    # Call LogEvent("     Fehler beim Speichern des MailAttachments position " & j & _
    # Err.Description, eLall)
    # Call N_ClearAppErr
    # GoTo AttachmentDone
    # Call LogEvent("      saved attachment " & JFirst _
    # & " as " & attachmentFile, eLSome)

    if CopyOriginal And Not ItemInIMAPFolder And Not DoNotDeleteAttachment Then:
    if DelSavedAttachments Then:
    # aBugTxt = "delete mail attachment " & JFirst
    # Call Try
    # thisAttachment.Delete
    if Catch Then:
    # Call N_ClearAppErr
    # GoTo AttachmentDone
    else:
    # JFirst = JFirst + 1
    if HeaderFirst Then                      ' log announcement needed?:
    # HeaderFirst = False
    # InsertText = "<p>" & "Extrahierte Anhnge: " & iAttachCnt
    if Item.Class = olMail Then:
    if Item.BodyFormat = olFormatHTML Then:
    # InsertText = InsertText & "<br>" & j & ": " & "<A HREF=""" & _
    # attachmentFile & """>" & attachmentFile
    else:
    # InsertText = InsertText & "<br>" & j & ": " & attachmentFile
    else:
    # InsertText = InsertText & "<br>" & j & ": " & attachmentFile
    # AttachmentDone:

    if Item.Class = olMail Then:
    if LenB(InsertText) = 0 Then:
    if TargetShouldChangeToHTML Then:
    # Call LogEvent("     Email converted to HTML in " _
    # & TargetFolder.FullFolderPath, eLall)
    else:
    # Call LogEvent("     Email is HTML already in " _
    # & TargetFolder.FullFolderPath)
    else:
    # bodyHTMLtext = Replace(Item.HTMLBody, "</body>", InsertText & _
    # "</BODY>", 1, 1, vbTextCompare)
    # Call LogEvent("     Target item converted to HTML " _
    # & "and attachment references inserted ", eLnothing)
    if Item.HTMLBody <> bodyHTMLtext And LenB(bodyHTMLtext) > 0 Then:
    # Item.HTMLBody = bodyHTMLtext
    # MailModified = True

    if Catch Then:
    # Call LogEvent("**** Fehler: " & Err.Description & b & TypeName(Item) & ": " _
    # & Item.Subject)
    elif DoChng2Item = 0 Then:
    # DoChng2Item = vbOK

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub addContactPic
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def addcontactpic():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.addContactPic"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim ContactPicture As String
    # ContactPicture = cPfad & ReFormat(cItem.FileAs, ".\/?*", b, b) _
    # & " (" & cItem.Parent.Parent.Name & " - " _
    # & Replace(DateId, b, vbNullString) _
    # & ").jpg"
    # aBugTxt = "save contact's picture in " & ContactPicture
    # Call Try
    # attThisContact.SaveAsFile ContactPicture
    if Not Catch Then:
    # Call LogEvent("      Bild fr Kontakt gespeichert in " & ContactPicture)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ItmDispose
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Item is Saved, Deleted or MovedToTrash
# '---------------------------------------------------------------------------------------
def itmdispose():
    # Const zKey As String = "MailProcessing.ItmDispose"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")

    # Dim MovedItemO As Object
    # Dim Source As String

    # aItemO.UnRead = True                         ' Mark unread whereever it now goes...
    # Call ForceSave(aItemO)
    if toThisFolder.FolderPath = aItemO.Parent.FolderPath Then:
    # Call LogEvent("     " & msg & " was not moved to Folder " _
    # & Quote(aItemO.Parent.FolderPath) _
    # & "  because it is already there.", eLnothing)
    # GoTo FuncExit
    # Source = aItemO.Parent.FolderPath

    # aBugTxt = "Item.Move"                        ' Try anything, autocatch
    # Call Try
    # Set MovedItemO = aItemO.Move(toThisFolder)
    if CatchNC Then:
    if Left(Hex(ErrorCaught), 7) = "8004010" Then ' we have seen last hex as A-F:
    # DoVerify False, " maybe the SpecialSearchFolderName is gone???"

    if isEmpty(MovedItemO) Then:
    # Set MovedItemO = Nothing
    # DoVerify False, "design big change ???"
    # Call N_ClearAppErr
    # Call ErrReset(0)
    # GoTo doDel

    if Not MovedItemO Is Nothing Then:
    if CatchNC Then:
    if InStr(E_AppErr.Description, "sondern kopiert") > 0 Then:
    # Call LogEvent("     " & msg & " has been copied to " & Quote(toThisFolder) _
    # & " and no longer exists in " & Quote(aItemO.Parent.FolderPath) _
    # & " , " & vbCrLf & "but could now be a duplicate in " _
    # & Quote(Source))
    # Call ErrReset(4)
    # GoTo FuncExit
    else:
    # doDel:
    # aBugTxt = "delete item"          ' Try anything, autocatch
    # Call Try
    # aItemO.Delete                    ' if it is no longer there, that's OK:
    if Catch(DoMessage:=False) Then:
    # Call LogEvent("     " & msg & " has been copied to " _
    # & Quote(toThisFolder) _
    # & "  but no longer exists in " _
    # & Quote(aItemO.Parent.FolderPath) & b)
    # GoTo FuncExit
    # Set aItemO = MovedItemO
    # Set aItmDsc.idObjItem = aItemO
    # Call LogEvent("     " & msg & " has been moved to Folder " _
    # & Quote(aItemO.Parent.FolderPath) & b, eLmin)
    elif E_Active.errNumber = -2147219840 Then:
    # Set MovedItemO = aItemO
    # aBugTxt = "delete Item"
    # Call Try(-2147219840)
    # aItemO.UnRead = True
    if Catch(DoMessage:=False) Then:
    # Call LogEvent("     " & msg & " can not be moved or deleted from " _
    # & Quote(aItemO.Parent.FolderPath), eLmin)
    else:
    # Call LogEvent("     " & msg & " has been copied to " & Quote(toThisFolder) _
    # & "  and deleted from " & Quote(aItemO.Parent.FolderPath), eLmin)
    else:
    # Call LogEvent("     " & msg & " should be in " _
    # & Quote(toThisFolder.FolderPath))

    # FuncExit:
    # aBugTxt = "MailProcessing.ItmDispose failed"
    # Call Try
    # ItemInIMAPFolder = getAccountType(Source, aAccountTypeName) = olImap
    # Catch

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub RestEvn4Item
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def restevn4item():
    # Const zKey As String = "MailProcessing.RestEvn4Item"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")


    if StopRecursionNonLogged Then               ' no change in NoEventOnAddItem:
    if Not NoEventOnAddItem Then:
    if SkipedEventsCounter > 1 Then:
    # NoEventOnAddItem = True
    # SkipedEventsCounter = SkipedEventsCounter - 1
    else:
    # SkipedEventsCounter = SkipedEventsCounter - 1
    if SkipedEventsCounter > 1 Then          ' some elements waiting:
    if DebugMode Then:
    # & vbCrLf & "Reset? Cancel=Stop", vbYesNoCancel + vbDefaultButton2)
    if rsp = vbYes Then:
    # SkipedEventsCounter = -SkipedEventsCounter
    elif rsp = vbNo Then:
    elif rsp = vbCancel Then:
    # DoVerify False
    if SkipedEventsCounter <= 0 Then         ' no reason now ???:
    # SkipedEventsCounter = -1
    # NoEventOnAddItem = False
    # SkipedEventsCounter = 0

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SetEventMode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def seteventmode():
    # Const zKey As String = "MailProcessing.SetEventMode"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    if Not StopRecursionNonLogged Then:
    if Not NoEventOnAddItem _:
    # And SkipedEventsCounter Mod 9 = 2 Then
    if DebugMode Then:
    print('Check Unlogged items, Deferred >= ')

    # SkipedEventsCounter = SkipedEventsCounter + 1
    if Not NoEventOnAddItem Then:
    if SkipedEventsCounter > 1 Then:
    # NoEventOnAddItem = True              ' so we can ProcCall additem once
    if force Or SkipedEventsCounter > 2 Then     ' accept no more:
    # SetEventMode = True
    # StopRecursionNonLogged = True
    else:
    # SetEventMode = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoDeferred
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def dodeferred():
    # Dim zErr As cErr
    # Const zKey As String = "MailProcessing.DoDeferred"

    # '------------------- gated Entry -------------------------------------------------------


    if Deferred Is Nothing Then:
    # Set Deferred = New Collection            ' define a new one
    # GoTo pExit
    if Deferred.Count = 0 Then:
    # GoTo pExit

    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim collItem As Object
    # Dim ProcessItem As Object
    # Dim AO As cActionObject

    # 'On Error GoTo 0
    for collitem in deferred:
    # i = i + 1
    # Set AO = collItem
    # Set ProcessItem = aNameSpace.GetItemFromID(AO.aoObjID)
    if E_AppErr.errNumber <> 0 Then DoVerify False:
    if IsMailLike(ProcessItem) Then:
    # ActionID = AO.ActionID
    match AO.ActionID:
        case 3:
    if Not ShutUpMode Then:
    if aID(aPindex).idObjDsc.objHasReceivedTime Then:
    print(Debug.Print "Starting on Deferred Item " & i _)
    # & " (Received " _
    # & ProcessItem.ReceivedTime _
    # & " SentOn " & ProcessItem.SentOn & ") of " _
    # & Deferred.Count
    else:
    print(Debug.Print "Starting on Deferred Item " & i & " Created On " & aID(aPindex).idTimeValue)
    # Call DoMailLike(ProcessItem)
    # Call ShowStatusUpdate
        case _:
    # DoVerify False, " not implemented Action"
    else:
    if DebugMode Then DoVerify False, _:
    # "deferred processing only intended for Mail-Like Items"
    # Set Deferred = New Collection                ' define a new one
    if i > 0 Then:
    # Call LogEvent("==== Processed " & i _
    # & " previously un-processed items", eLmin)

    # FuncExit:
    # Set ProcessItem = Nothing
    # Set collItem = Nothing
    # Set AO = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Set2Logged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def set2logged():

    # Const zKey As String = "MailProcessing.Set2Logged"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim oCat As String
    try:
        # oCat = aMiObj.Categories
        if InStr(oCat, LOGGED) = 0 Then:
        # Call AppendTo(oCat, LOGGED, ";", ToFront:=True)
        # aMiObj.Categories = oCat
        # aMiObj.UnRead = False
        if Not aMiObj.Saved Then:
        # aMiObj.Save
        # GoTo FuncExit
        # bad:                DoVerify False

        # FuncExit:
        # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function CloseInsptrs
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Close open Inspector windows, dft: only type Email, else CloseAnything:=True
# ' Note   : Working backwards to prevent skipping any items
# '---------------------------------------------------------------------------------------
def closeinsptrs():
    # Const zKey As String = "MailProcessing.CloseInsptrs"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim lInsp As Outlook.Inspector
    # Dim lCount As Long
    # Dim i As Long
    # Dim lItem As Object

    # lCount = olApp.Inspectors.Count

    # Set lInsp = olApp.Inspectors(i)
    # Set lItem = lInsp.CurrentItem
    if CloseAnything Then:
    if DebugMode Then:
    print(Debug.Print "Item class: " & lInsp.CurrentItem.Class & b;)
    # GoTo DoAny
    else:
    if IsMailLike(lItem) Then:
    # DoAny:
    # lItem.Close olDiscard
    if DebugMode Then:
    print(Debug.Print "Item discarded, Subject: " & Quote(lItem.Subject))

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : RestrictedItemsShow
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show items in RestrictedItemCollection
# '---------------------------------------------------------------------------------------
def restricteditemsshow():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "MailProcessing.RestrictedItemsShow"
    # Call DoCall(zKey, "Sub", eQzMode)

    # Dim aObj As Object
    # Dim i As Long

    if findCount = 0 Then:
    # findCount = RestrictedItemCollection.Count
    else:
    # findCount = Min(findCount, RestrictedItemCollection.Count)

    # Set aObj = RestrictedItemCollection.Item(i)
    print(Debug.Print LString(i, 5) & LString(aObj.Subject, lKeyM) _)
    # & b & LString(aObj.SenderName, 30) & b & aObj.SentOn

    # FuncExit:
    # Set aObj = Nothing

    # zExit:
    # Call DoExit(zKey)



# Converted from File_FolderOps.py

# Attribute VB_Name = "File_FolderOps"
# Option Explicit

# Public OpenFileNames As Dictionary
# Public ClosedFileNames As Dictionary
# Public FileStates As Dictionary

# Public aOpenKey As Long
# Public aClosedKey As Long
# Public aFileState As String
# Public aFileSpec As String
# Public MoveMode As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Function FolderActions
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def folderactions():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.FolderActions"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ActionOnWhat As String
    # Dim ActionResult As String
    # Dim curFolder As Folder
    # Dim oneItem As Object

    if itemNo < 0 Then                  ' only one item within selection:
    if -itemNo > curObj.Count Then:
    # GoTo FuncExit               ' all has been done
    # Set oneItem = curObj(-itemNo)   ' OlApp.ActiveExplorer.Selection

    # aBugTxt = "access to parent of object"
    # Call Try("%Das Element wurde verschoben")
    # Set curFolder = oneItem.Parent
    if Catch Then:
    # GoTo FuncExit
    # WorkIndex(1) = -itemNo
    # Call InitFindModel(oneItem)     ' get Search criteria
    else:
    if itemNo > curObj.Items.Count Then:
    # itemNo = curObj.Items.Count
    if itemNo = 0 Then:
    # FolderActions = vbCancel
    # GoTo FuncExit
    else:
    # Set oneItem = curObj.Items(itemNo)
    # Set curFolder = curObj
    # WorkIndex(1) = itemNo

    if Not (LF_DontAskAgain _:
    # Or oneItem.Parent Is Nothing) Then
    if NonLoopFolder(oneItem.Parent.Name) Then:
    # GoTo FuncExit

    if eOnlySelectedItems Then:
    # FolderActions = vbOK  ' no filtering checks if already selected
    else:
    # FolderActions = ItemDateFilter(oneItem, eLall)
    if FolderActions = vbOK Or FolderActions = vbIgnore Then:
    # ActionOnWhat = "   processing item " & Quote(oneItem.Subject)
    else:
    # ActionOnWhat = "   skipping item " & Quote(oneItem.Subject)
    # GoTo FuncExit
    match ActionID:
    # ' collect these items for later
        case atDefaultAktion::
    if eOnlySelectedItems Then:
    # oneItem.Display True    ' modal display
    else:
    # Call LogEvent("---- " & TypeName(oneItem) & ": " _
    # & oneItem.Subject, eLnothing)
    # WorkIndex(1) = SelectedItems.Count + 1
    # SelectedItems.Add oneItem, CStr(WorkIndex(1))
    # FolderActions = vbOK
        case atKategoriederMailbestimmen::
    # Call LogEvent("---- " & TypeName(oneItem) & ": " _
    # & oneItem.Subject, eLnothing)
    # CurIterationSwitches.ResetCategories = True ' that's what we are here for
    if Not CurIterationSwitches.CategoryConfirmation Then   ' confirming changed unless we said "dontAsk":
    # CurIterationSwitches.CategoryConfirmation = Not CurIterationSwitches.ReProcessDontAsk
    # ActionResult = DetectCategory(curFolder, oneItem, curFolder.FullFolderPath)
    if CurIterationSwitches.SaveItemRequested And MailModified And LenB(ActionResult) > 0 Then:
    # LF_ItmChgCount = LF_ItmChgCount + 1
    # oneItem.Categories = ActionResult
    # Call LogEvent("Mail categories assigned: " _
    # & oneItem.Categories, eLnothing)
    # oneItem.Save
    else:
    # Call LogEvent("Mail categories not changed: " _
    # & oneItem.Categories, eLnothing)
        case atPostEingangsbearbeitungdurchfhren::
    # Call DeferredActionAdd(oneItem, curAction:=3)
        case atDoppelteItemslschen::
    if eOnlySelectedItems Then:
    # Call MatchingItems(MatchMode:=0)
    else:
    # Call CheckDoublesInFolder(curFolder)
        case atNormalreprsentationerzwingen::
    # Call ScanItem(itemNo, oneItem)
        case atOrdnerinhalteZusammenfhren::
        case atFindealleDeferredSuchordner::
        case atBearbeiteAllebereinstimmungenzueinerSuche::
    # Call CheckItemProcessed(oneItem)
        case atContactFixer::
    # Call ContactFixItem(oneItem)
        case _:
    print('Aktion ')

    # FuncExit:
    # Call N_ClearAppErr

    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Function ActivateDeferredFavorites
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def activatedeferredfavorites():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.ActivateDeferredFavorites"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim objPane As NavigationPane
    # Dim mailModule As Outlook.mailModule
    # Dim FavFolders As Outlook.NavigationFolders
    # Dim curFolder As Outlook.NavigationFolder
    # Dim i As Long

    # ' Get the NavigationPane object for the
    # ' currently displayed Explorer object.
    # 'On Error GoTo 0

    # Set objPane = olApp.ActiveExplorer.NavigationPane
    # Set mailModule = objPane.Modules.GetNavigationModule(olApp.OlNavigationModuleType.olModuleMail)
    # Set ActivateDeferredFavorites = mailModule.NavigationGroups.GetDefaultNavigationGroup(olApp.OlGroupType.olFavoriteFoldersGroup)
    if Not justGet Then:
    # Set FavFolders = ActivateDeferredFavorites.NavigationFolders
    # Set curFolder = FavFolders.Item(i)
    if InStr(1, curFolder.DisplayName, _:
    # SpecialSearchFolderName, vbTextCompare) > 0 Then
    if DebugMode Then:
    print(Debug.Print LString("+ " & curFolder.Folder.FolderPath, OffObj) _)
    # & "contains" & RString(curFolder.Folder.Items.Count, 8) _
    # & " Deferred Items "
    # FavNoLogCtr = FavNoLogCtr + 1
    else:
    if DebugMode Then:
    print(Debug.Print LString("- " & curFolder.Folder.FolderPath, OffObj) _)
    # & "shows   " & RString(curFolder.Folder.Items.Count, 8) _
    # & " Deferred Items "
    if FavNoLogCtr < FldCnt Then:
    # Call LogEvent("Found " & FavNoLogCtr _
    # & " regular Folders named " & Quote(SpecialSearchFolderName) _
    # & " within " _
    # & FldCnt & " Favorite (Navigation) Folders", eLall)
    if FavNoLogCtr <= 0 Then:
    # & vbCrLf & "  (Jedes Konto sollte die Posteingnge und Sendungen hierauf absuchen)" _
    # & vbCrLf & "Abbrechen des Laufs: OK, Cancel ignoriert dieses Problem", vbOKCancel)
    if rsp = vbOK Then:
    # End

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' collect all Folders of same type as aFolder on this level => lFolders()
def getfoldersoftype():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.getFoldersOfType"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim thisitemType As OlObjectClass
    # Dim aFld As Folder

    if afolder.Parent Is Nothing Then:
    # Set afolder = aNameSpace.GetDefaultFolder(olFolderInbox)
    if afolder.Parent.Class = olFolder Then:
    # Set topFolder = afolder.Parent
    # thisitemType = afolder.DefaultItemType
    else:
    # DoVerify False, " wattndattn?"
    # ' loop to get number of Folders of same type
    for afld in topfolder:
    if aFld.DefaultItemType = thisitemType Then:
    # FldCnt = FldCnt + 1
    # ReDim lFolders(1 To FldCnt)
    # FldCnt = 0
    for afld in topfolder:
    if aFld.DefaultItemType = thisitemType Then:
    # FldCnt = FldCnt + 1
    # Set lFolders(FldCnt) = aFld

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function getParentFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getparentfolder():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.getParentFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim iObj As Object
    # Dim lAccountDsc As cAccount
    # Dim DisplayName As String

    # ' if error: result is nothing
    # Set iObj = aItemO.Parent    ' usually  aItemO = SelectedItems.Item(1)
    if Catch Then:
    # GoTo FuncExit

    # While Not iObj Is Nothing
    if iObj.Class = olFolder Then:
    # Set getParentFolder = iObj
    if Catch Then:
    # GoTo FuncExit
    # DisplayName = getParentFolder.Store.DisplayName
    if Not D_AccountDscs.Exists(DisplayName) Then   ' folders/stores not having account are ok (e.g. backup):
    # GoTo FuncExit
    # Set lAccountDsc = D_AccountDscs.Item(DisplayName)
    if Catch Then:
    # GoTo FuncExit
    # ItemInIMAPFolder = lAccountDsc.aAcType = olImap
    # Catch
    # GoTo FuncExit
    else:
    # Set iObj = iObj.Parent
    # Wend

    # FuncExit:
    # Call ErrReset(4)
    # Set iObj = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : GetOrMakeNotLogged
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Search Folders for NotLogged Mail Items
# '---------------------------------------------------------------------------------------
def getormakenotlogged():

    # Const zKey As String = "File_FolderOps.GetOrMakeNotLogged"
    # Call DoCall(zKey, tSub, eQzMode)

    # Call GetAccountSearchFolders(SpecialSearchFolderName, _
    # "NOT Categories LIKE " & Quote1(LOGGED))

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : GetAccountSearchFolders
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Search (or make required) search folders for all Accounts
# '---------------------------------------------------------------------------------------
def getaccountsearchfolders():
    # Const zKey As String = "File_FolderOps.GetAccountSearchFolders"
    # Dim zErr As cErr

    # Dim aAccount As Account
    # Dim oStore As Outlook.Store
    # Dim oFolder As Outlook.Folder

    # Dim Scope As String
    # Dim P As Long

    # Dim objSearch As Search

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="File_FolderOps")

    if SessionAccounts Is Nothing Then:
    # Set SessionAccounts = olApp.Session.Accounts

    for aaccount in sessionaccounts:
    # Set oStore = aAccount.DeliveryStore
    # Set oSearchFolders = oStore.GetSearchFolders

    for ofolder in osearchfolders:
    # Scope = oFolder.FolderPath
    # P = InStr(3, Scope, "\")
    if P = 0 Then                       ' only in hierarchy:
    # GoTo NSF
    print(Debug.Print Scope, oFolder.Name)
    if oFolder.Name = FolderName Then:
    print(Debug.Print "ofolder matches folder name: ", Scope)
    # GoTo NAC                        ' it exists for this account, nice
    # NSF:
    # ' not existing yet: make it
    # Scope = "'" & Left(Scope, P) & StdInboxFolder & "',"
    print(Debug.Print "Making " & Scope, "Filter=" & Quote(Filter))
    # Set objSearch = olApp.AdvancedSearch(Scope:=Scope, _
    # Filter:=Filter, _
    # SearchSubFolders:=True, _
    # Tag:=FolderName)
    # oFolder.ShowItemCount = olShowTotalItemCount
    # Call objSearch.Save(FolderName)
    # GoTo FuncExit
    # NAC:

    # FuncExit:
    # Set aAccount = Nothing
    # Set oStore = Nothing
    # Set oFolder = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub GetSearchFolders
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get all / select some Search Folders in all Stores
# '---------------------------------------------------------------------------------------
def getsearchfolders():
    # Const zKey As String = "File_FolderOps.GetSearchFolders"
    # Dim zErr As cErr

    # Dim oStore As Outlook.Store
    # Dim oFolder As Outlook.Folder
    # Dim retrycount As Long
    # Dim i As Long

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="File_FolderOps")

    if UInxDeferred <= 0 Then:
    # UInxDeferred = 0
    # Retry:

    for ostore in ostores:
    # ' like Ungesehen, NotLogged etc. some always exist!
    # aBugTxt = "Get Search Folders in " & oStore.DisplayName
    # Call Try
    # Set oSearchFolders = oStore.GetSearchFolders
    if Not Catch Then:
    for ofolder in osearchfolders:
    # Call LogEvent("=======> Found search Folder " _
    # & oFolder.FullFolderPath)
    if oFolder.Name = SpecialSearchFolderName Then:
    # Call LogEvent("!!!=> Found SpecialSearchFolderName Folder " _
    # & oFolder.FullFolderPath, eLall)
    # UInxDeferred = UInxDeferred + 1
    # Set DeferredFolder(UInxDeferred) = oFolder
    # Catch
    # noSearchFolders:

    # Call LogEvent("Es wurden " & UInxDeferred & " Suchordner vom Typ " _
    # & Quote(SpecialSearchFolderName) & " gefunden.")
    # Set oFolder = DeferredFolder(i)

    print(Debug.Print LString("+ " & oFolder.Folder.FolderPath, OffObj) _)
    # & "contains" & RString(oFolder.Folder.Items.Count, 8) & " Deferred Items "
    # ' currently at least one SpecialSearchFolderName Folder rqd in Backup
    if UInxDeferred > 1 Then:
    if UInxDeferred < FldCnt And Not UInxDeferredIsValid Then:
    # Call ActivateDeferredFavorites
    if FavNoLogCtr > 0 Then:
    # UInxDeferred = 0
    # retrycount = retrycount + 1
    if retrycount < 4 Then:
    # GoTo Retry
    else:
    # retrycount = 0
    if DebugMode Then:
    # DoVerify False, " SpecialSearchFolderName Folders not visible or do not exist"
    else:
    # Call LogEvent("Es wird nicht mehr nach weiteren " _
    # & Quote(SpecialSearchFolderName) _
    # & " Ordnern gesucht", eLall)
    # UInxDeferredIsValid = True

    # FuncExit:
    # Set oStore = Nothing
    # Set oFolder = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ItemDateFilter
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def itemdatefilter():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.ItemDateFilter"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ItemDate As Date
    if CutOffDate = "00:00:00" Then:
    # ItemDateFilter = vbOK   ' No filter set, pass thru
    else:
    # ' IsMailLike aufgesplittet:
    match TypeName(oneItem):
        case "MailItem", "MeetingItem":
    # ItemDate = oneItem.SentOn
    # aTimeFilter = "SentOn"
    if ItemDate = BadDate Then:
    # ItemDate = oneItem.ReceivedTime
    # aTimeFilter = "ReceivedTime"
    # doFilter:
    if ItemDate >= CutOffDate Then:
    # ItemDateFilter = vbOK   ' do process this
    else:
    # DateSkipCount = DateSkipCount + 1
    # ItemDateFilter = vbNo   ' do ignore this
    # Call LogEvent(LimitAppended("     ", Quote(oneItem.Subject), 30, "... ") _
    # & " verfehlt Datumsauswahl: " _
    # & CStr(ItemDate) & "<" & CStr(CutOffDate), logLvl)
        case "AppointmentItem":
    # ItemDate = Format(oneItem.End, "dd.mm.yyyy")
    if ItemDate = BadDate Then:
    # ItemDateFilter = vbOK   ' no end: keep going
    # GoTo doFilter
        case _:
    # ItemDateFilter = vbIgnore   ' no specific filter defined

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function PostFolderActions
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def postfolderactions():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.PostFolderActions"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ActionOnWhat As String
    if Not LF_CurLoopFld Is Nothing Then:
    # ActionOnWhat = "   processing Folder " & LF_CurLoopFld
    # PostFolderActions = vbOK
    match ActionID:
        case atDefaultAktion                                ' 1:
    if eOnlySelectedItems Then:
    # Call ItemActions(SelectedItems)
        case atKategoriederMailbestimmen                    ' 2:
        case atPostEingangsbearbeitungdurchfhren           ' 3:
        case atDoppelteItemslschen                         ' 4:
        case atNormalreprsentationerzwingen                ' 5:
        case atOrdnerinhalteZusammenfhren                  ' 6:
        case atFindealleDeferredSuchordner                  ' 7:
        case atBearbeiteAllebereinstimmungenzueinerSuche   ' 8:
        case atContactFixer                                 ' 9:
        case _:
    print('Post-Folder-Action ')

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' this function runs before we ProcCall a Folder level
def beforefolderactions():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.BeforeFolderActions"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ActionOnWhat As String
    # BeforeFolderActions = vbOK
    if Not LF_CurLoopFld Is Nothing Then:
    # ActionOnWhat = "   processing Folder " & LF_CurLoopFld
    # Call BestObjProps(LF_CurLoopFld, withValues:=False)

    match ActionID:
        case atDefaultAktion                                                ' 1:
    # SelectMulti = False
    if aObjDsc.objItemClass = olAppointmentItem Then:
    # Call ItemActions(LF_CurLoopFld)
    else:
    print('Eine Default-Operation ist fr den Item-Typ ')
    # & aObjDsc.objItemClassName & " nicht definiert"
    # BeforeFolderActions = vbNo
    # GoTo ProcReturn
    # Call LogEvent("Bis Ordner " & Quote(LF_CurLoopFld.FolderPath) _
    # & "  wurden " & SelectedItems.Count & " Items gewhlt")
        case atKategoriederMailbestimmen                                    ' 2:
    # SelectMulti = False
        case atPostEingangsbearbeitungdurchfhren, _:
    # atNormalreprsentationerzwingen                                 ' 3, 5
    # SelectMulti = True
    if eOnlySelectedItems Then:
    # Call ItemActions(SelectedItems)
    else:
    # Call ItemActions(LF_CurLoopFld)
        case atDoppelteItemslschen                                         ' 4:
    # SelectMulti = True
    # Call Initialize_UI     ' displays options dialogue
    match rsp:
        case vbYes:
    if eOnlySelectedItems Then   ' set loop exit for Folders::
    # LF_DoneFldrCount = LookupFolders.Count    ' do not loop Folder list if Selection
    # GoTo likepicked
    elif PickTopFolder Then:
    # LF_DoneFldrCount = LookupFolders.Count    ' do not loop Folder list if Folder is picked
    # Call PickAFolder(1, _
    # "bitte besttigen oder whlen Sie " _
    # & "den obersten Ordner fr die Doublettensuche ", _
    # "Auswahl des Hauptordners fr die Doublettensuche", _
    # "OK", "Cancel")
    # Set topFolder = Folder(1)
    # likepicked:
    # Call FindTrashFolder
    # Set ParentFolder = Nothing

    # aBugTxt = "Get Parent folder of " _
    # & topFolder.FolderPath
    # Call Try
    # Set ParentFolder = topFolder.Parent
    # Catch

    # Set LF_CurLoopFld = topFolder
    # curFolderPath = LF_CurLoopFld.FolderPath
    # FullFolderPath(FolderPathLevel) = "\\" _
    # & Trunc(3, curFolderPath, "\")
    if BeforeItemActions() = vbOK Then:
    # ' Debug.Assert False
    else:
    if BeforeItemActions() = vbOK Then:
    # bDefaultButton = "Go"
        case vbCancel:
    # Call LogEvent("=======> Stopped before processing any Folders . Time: " _
    # & Now(), eLnothing)
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo ProcReturn
        case Else   ' loop Candidates:
    if topFolder Is Nothing Then:
    # Set topFolder = LookupFolders.Item(LF_DoneFldrCount)
    # Call FindTrashFolder
    if eOnlySelectedItems Then:
    # BeforeFolderActions = vbNo ' we are done
    # GoTo ProcReturn
        case atOrdnerinhalteZusammenfhren                                  ' 6:
    # SelectMulti = True
    # Call AddItemDataToOlderFolder
        case atFindealleDeferredSuchordner                               ' 7:
    # SelectMulti = False
    # Call FldActions2Do    ' if we have open items, do em now
    # BeforeFolderActions = vbNo
        case atBearbeiteAllebereinstimmungenzueinerSuche                   ' 8:
    # SelectMulti = True
    # ' Folder by default (no user interaction)
    # Set ChosenTargetFolder = GetFolderByName("Erhalten")
    # Call FirstPrepare
        case atContactFixer:
    # SelectMulti = False
        case _:
    # SelectMulti = False
    print('Pre-Folder-Action ')

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' this function runs before we process items in a single Folder
def beforeitemactions():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.BeforeItemActions"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ActionOnWhat As String

    # BeforeItemActions = vbOK
    match ActionID:
        case atDefaultAktion:
    # BeforeItemActions = SkipTop(LF_CurLoopFld)
        case atKategoriederMailbestimmen:
        case atPostEingangsbearbeitungdurchfhren:
    if eOnlySelectedItems Then:
    # Call ItemActions(SelectedItems)
    else:
    # Call ItemActions(LF_CurLoopFld)
        case atDoppelteItemslschen:
    if eOnlySelectedItems Then:
    # Call ItemActions(SelectedItems)
    else:
    # Call ItemActions(LF_CurLoopFld)
    # ' Call CheckDoublesInFolder(topFolder)    Main Work in "ItemActions"
        case atNormalreprsentationerzwingen:
    # Call BestObjProps(LF_CurLoopFld, withValues:=False)
        case atOrdnerinhalteZusammenfhren:
        case atFindealleDeferredSuchordner:
        case atContactFixer:
        case _:
    print('Pre-Item-Action ')
    if LF_CurLoopFld Is Nothing Then:
    # GoTo ProcReturn
    # ActionOnWhat = "   processing Folder " _
    # & LF_CurLoopFld & " is " & BeforeItemActions

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SetItemCategory
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setitemcategory():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.SetItemCategory"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ShortName As String
    if IsMailLike(curItem) Then:
    if curFolder <> topFolder Then:
    # ShortName = Left(curFolder.Name, 4)
    if ShortName = "Junk" Or ShortName = "Spam" Or ShortName = "Uner" Then:
    # category = "Junk"
    else:
    # curItem.Categories = category
    # LF_ItmChgCount = LF_ItmChgCount + 1
    # curItem.Save
    # SetItemCategory = vbOK

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SkipTop
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def skiptop():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.SkipTop"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if curFolder = topFolder Then:
    # SkipTop = vbNo
    else:
    # SkipTop = vbOK

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FindTopFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findtopfolder():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.FindTopFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if topFolder Is Nothing Then:
    if afolder.Parent Is Nothing Then:
    # ' we are in a search Folder=
    if afolder.Items.Count > 0 Then ' all items are in same Folder!!:
    # Call Try
    # Set topFolder = afolder.Items(1).Parent
    # Catch
    else:
    # topFolder = Nothing
    # GoTo ProcReturn
    else:
    # Set topFolder = afolder

    # Call ErrReset(0)
    # Do                                  ' loop until we reach the outermost folder
    if topFolder.Parent.Class = olFolder Then:
    # Set topFolder = topFolder.Parent
    else:
    # Exit Do                     ' topfolder.parent is Mapi
    # Loop

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyitems():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.CopyItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long, j As Long
    # Static S_Folders(2) As Folder

    # Dim action As String

    # Set LookupFolders = aNameSpace.Folders
    # bDefaultButton = "No"

    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask1again:
    # "OK", "Cancel", "Auswahl des Quell- und Zielordners")

    match rsp:
        case vbOK:
    # Set S_Folders(1) = ActiveExplorer.CurrentFolder
    if Not S_Folders(1) Is Nothing Then:
    # i = S_Folders(1).Items.Count
    if i < 1 Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    if ActiveExplorer.Selection.Count > 0 Then:
    # Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
    else:
    # Set ActiveExplorerItem(1) = S_Folders(1).Items(1)
    else:
    # action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
    # GoTo ask1again
        case vbCancel:
    # Call LogEvent("=======> Beendet ohne Quell-Auswahl . Time: " & Now(), eLmin)
    if TerminateRun Then:
    # GoTo ProcReturn
    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask2again:
    # "OK", "Cancel", "Auswahl des Quell- und Zielordners")

    match rsp:
        case vbOK:
    # Set S_Folders(2) = ActiveExplorer.CurrentFolder
    if S_Folders(2) Is Nothing Then:
    # action = " (ein Ziel muss angegeben werden!)"
    # GoTo ask2again
    if S_Folders(1) = S_Folders(2) Then:
    # DoVerify False, " select S_Folders now and then F5"
    # GoTo ask2again
        case vbCancel:
    # Call LogEvent("=======> Beendet ohne Zielauswahl. Time: " & Now(), eLmin)
    if TerminateRun Then:
    # GoTo ProcReturn

    if MoveMode Then:
    # Call LogEvent("All Items from " & S_Folders(1).FolderPath _
    # & " will be moved to " & S_Folders(2).FolderPath, eLmin)
    else:
    # Call LogEvent("All Items from " & S_Folders(1).FolderPath _
    # & " will be copied to " & S_Folders(2).FolderPath, eLmin)

    # gothemAll:
    # i = S_Folders(1).Items.Count
    if i < j Then:
    # GoTo gothemAll
    try:
        if MoveMode Then:
        # Set ActItemObject = S_Folders(1).Items.Item(j)
        # ActItemObject.Move S_Folders(2)
        if DebugLogging Or DebugMode Then:
        print(Debug.Print "Moved " & ActItemObject.Subject & " to " _)
        # & S_Folders(2).FolderPath
        else:
        # Set ActItemObject = S_Folders(1).Items.Item(j).Copy
        # ActItemObject.Move S_Folders(2)
        if DebugLogging Or DebugMode Then:
        print(Debug.Print "Copied " & ActItemObject.Subject & " to " _)
        # & S_Folders(2).FolderPath
        # GoTo OK
        # coudNotCopy:
        print(Debug.Print j, ActItemObject.Subject, Err.Description)
        if Not MoveMode Then:
        # ActItemObject.Delete        ' sonst wurde doublette erzeugt
        # Call N_ClearAppErr
        # Call ErrReset(0)
        # OK:
        # MoveMode = False

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyNotes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copynotes():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.CopyNotes"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim j As Long
    # Dim bodystring As String
    # Dim action As String

    # Set LookupFolders = aNameSpace.Folders
    # bDefaultButton = "No"

    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask1again:
    # "OK", "Cancel", _
    # "Auswahl des Quell- und Zielordners")

    match rsp:
        case vbOK:
    # Set Folder(1) = ActiveExplorer.CurrentFolder
    if Not Folder(1) Is Nothing Then:
    # i = Folder(1).Items.Count
    if i < 1 Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    if ActiveExplorer.Selection.Count > 0 Then:
    # Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
    else:
    # Set ActiveExplorerItem(1) = Folder(1).Items(1)
    else:
    # action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
    # GoTo ask1again
        case vbCancel:
    # Call LogEvent("=======> Beendet ohne Quell-Auswahl . Time: " _
    # & Now(), eLmin)
    # End
    # action = " und besttigen Sie die Auswahl mit 'OK'"
    # ask2again:
    # "OK", "Cancel", "Auswahl des Quell- und Zielordners")

    match rsp:
        case vbOK:
    # Set Folder(2) = ActiveExplorer.CurrentFolder
    if Folder(2) Is Nothing Then:
    # action = " (ein Ziel muss angegeben werden!)"
    # GoTo ask2again
        case vbCancel:
    if DebugMode Or MinimalLogging < 2 Then:
    # Call LogEvent("=======> Beendet ohne Zielauswahl. Time: " _
    # & Now(), eLmin)
    if TerminateRun Then:
    # GoTo ProcReturn
    # gothemAll:
    # Set ActItemObject = Folder(2).Items.Add(Folder(2).DefaultItemType)
    # bodystring = Folder(1).Items.Item(j).Body
    if InStr(bodystring, "Notizen\") = 1 Then:
    # bodystring = Mid(bodystring, 9)
    # ActItemObject.Body = bodystring
    # ActItemObject.Save

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MoveItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def moveitems():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.MoveItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # MoveMode = True
    # Call CopyItems

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : GetOrMakeOlFolder
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get an Outlook Folder by Name; if not found ask if it is to be created
# ' Result : True if Folder NOT found and not made successfully after request
# '---------------------------------------------------------------------------------------
# '        allowMissing=0                      ' optional without question
# '        allowMissing=1                      ' optional with question
# '        allowMissing=Else                   ' mandatory folder existance
def getormakeolfolder():
    # Dim zErr As cErr
    # Const zKey As String = "File_FolderOps.GetOrMakeOlFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # aBugTxt = "Ordner " & FolderName & " nicht gefunden."
    # Call Try(-2147221233)
    # Set useFolder = belowFolders.Item(FolderName)
    # Catch
    # '   Verweis auf Zielordners des Transports aus der Inbox
    if useFolder Is Nothing Then:
    match allowMissing:
        case 1                      ' optional with question:
    # Call ErrReset(4)
    if rsp = vbYes Then:
    # rsp = vbNo
    else:
    # rsp = vbYes
        case 0                      ' optional without question:
    # Call LogEvent(E_Active.Explanations & " Akzeptiert.", eLSome)
    # Call ErrReset(4)
    # rsp = vbNo
        case Else                   ' mandatory folder existance:

    if rsp = vbYes Then:
    # aBugTxt = "Make new Folder " & Quote(FolderName)
    # Call Try
    # Set useFolder = CreateFolderIfNotExists(FolderName, belowFolders.Item(1))
    # Catch
    # GetOrMakeOlFolder = useFolder Is Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub LogEvent
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# '---------------------------------------------------------------------------------------
def logevent():


    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Long

    if Recursive > 1 Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if Recursive > 3 Then:
    # GoTo refuse
    if StackDebug > 8 Then:
    # GoTo refuse
    if AppStartComplete Then:
    # refuse:
    print(Debug.Print String(OffCal, b) & "Ignored recursion to LogEvent " & vbCrLf & Text)
    # GoTo ProcRet
    # Recursive = Recursive + 1

    # Dim lLevel As eLogLevel
    # Dim LogThisToFile As Boolean
    # Dim SaveTry As Variant
    # Dim msg As String

    # SaveTry = E_Active.Permitted
    if fileNo <> 1 Then:
    print(Debug.Print "printing to logfile #" & fileNo & " not implemented")
    # GoTo FuncExit

    # With T_DC
    if DateId = vbNullString Then:
    # Call GetDateId(-1)                  ' just set global vars, value not needed
    else:
    # Call GetDateId(-2)                  ' long values remember globallly
    if DebugMode Then:
    # lLevel = ifLevelLess - 1            ' make logging more likely
    else:
    # lLevel = ifLevelLess

    if withMsgBox And lLevel <= eLSome Then:
    if rsp = vbCancel Then:
    # Call TerminateRun(withStop:=True)

    # LogThisToFile = StackDebug = 0 Or DebugMode _
    # Or DebugLogging Or StackDebug > 6 _
    # Or (lLevel <= eLall _
    # And lLevel <= eLnothing - MinimalLogging)

    if Not LogThisToFile Then:
    if LogImmediate Then:
    print(Debug.Print Text                    ' put into Direct Window)
    elif lLevel = 0 Or ((DebugMode Or DebugLogging Or StackDebug > 7)) Then:
    print(Debug.Print Text                    ' put into Direct Window)
    # GoTo FuncExit

    # OpenNextLogFile:
    # Err.Clear
    if LenB(.LogNameNext) = 0 Then:
    if LenB(.LogName) > 0 Then:
    # msg = "Old Log: " & .LogName & " closing. "
    elif LenB(.LogNamePrev) > 0 Then:
    # msg = "Old Log: " & .LogNamePrev & vbCrLf & "new Log: "
    else:
    # msg = "No old log, "
    # Call GetDateId(-2)              ' remember globally
    # .LogNameNext = lPfad & "Outlook VBA " & DateIdNB & ".log"
    # msg = msg & .LogNameNext
    elif InStr(.LogNameNext, Left(DateIdNB, 8)) = 0 Then:
    if LenB(.LogName) > 0 Then:
    # msg = "Old Log: " & .LogName & ", switch to new day: "
    elif LenB(.LogNamePrev) > 0 Then:
    # msg = "Old Log: " & .LogNamePrev & vbCrLf & "next Log: "
    else:
    # msg = "No old log "
    # .LogNameNext = lPfad & "Outlook VBA " & DateIdNB & ".log"
    # msg = msg & vbCrLf & "next log: " & .LogNameNext

    if .LogIsOpen Then:
    if .LogName <> .LogNameNext Then:
    # Call CloseLog(msg:=msg)                     ' changes .LogIsOpen := False
    # msg = vbNullString                                    ' Closelog did print
    # .LogName = .LogNameNext
    else:
    # .LogName = .LogNameNext

    if LenB(msg) > 0 Then:
    print(Debug.Print msg)
    # msg = vbNullString                                        ' print done
    if LenB(Text) > 0 Then:
    if LogImmediate _:
    # Or lLevel = 0 _
    # Or DebugMode _
    # Or DebugLogging _
    # Or StackDebug > 7 _
    # Then
    print(Debug.Print Text                    ' put into Direct Window)

    if .LogIsOpen Then:
    # GoTo Output

    try:
        # Open .LogName For Append As #1
        if Err.Number = 0 Then:
        # GoTo AllOk
        elif Err.Number = 55 Then                     ' already open:
        # E_Active.FoundBadErrorNr = 0
        # GoTo skipErrorTest
        elif Err.Number <> 0 Then:
        print(Debug.Print "Log " & .LogName & " did not open, Error " & Err.msg)
        # Debug.Assert False
        # .LogNameNext = vbNullString
        # GoTo OpenNextLogFile
        # skipErrorTest:
        print(Debug.Print vbCrLf & String(OffCal, b) & "Logfile reused: " & .LogName _)
        # & vbCrLf
        # .LogIsOpen = True
        else:
        # AllOk:
        # .LogIsOpen = True
        if .LogNamePrev <> .LogName Then:
        if LenB(.LogNamePrev) > 0 Then:
        # msg = "Previous Log Name: " & .LogNamePrev
        # Print #1, msg
        else:
        print(Debug.Print String(OffCal, b) & "Logfile re-opened: " & .LogName)

        # Output:
        if .LogIsOpen Then:
        try:
            # Print #1, Text
            if Err.Number = 0 Then:
            if DebugLogging And Not DebugMode Then:
            print(Debug.Print String(OffCal, b) & "Log appended:      " & .LogName)
            # GoTo LogFileOK
            else:
            # GoTo OpenNextLogFile

            if .DCerrNum = 52 Then:
            print(Debug.Print "LogFile Error: " & Text)
            if Catch Then:
            # .LogIsOpen = False
            # .LogName = vbNullString
            # .LogNameNext = vbNullString
            # GoTo OpenNextLogFile

            # LogFileOK:
            # .LogFileLen = .LogFileLen + Len(Text)
            if LimitLog > 0 Then:
            if .LogFileLen Mod 100 * LimitLog = 0 Then:
            print(Debug.Print " Log limit reached ~: " & .LogFileLen / LimitLog & " lines. Press F5 to continue")
            # ' Debug.Assert False     #### Switch this on?
            # ' Check max file size
            if .LogFileLen >= MaxCharsPerLogFile Then:
            # .LogNamePrev = .LogName                       ' force a new file name
            # Call CloseLog(msg:="Start new Log, Length exceeds " & MaxCharsPerLogFile & "<" & .LogFileLen)
            # GoTo OpenNextLogFile

            # FuncExit:
            if AppStartComplete Then:
            # Call ShowStatusUpdate
            # E_Active.Permit = SaveTry                        ' must not change
            # End With ' T_DC

            # Recursive = Recursive - 1

            # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CloseLog
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def closelog():
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean
    # Const zKey As String = "File_FolderOps.CloseLog"
    # Dim zErr As cErr

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # With T_DC
    if .LogIsOpen Then                              ' implies .logname<>"":
    if LenB(.LogName) = 0 Then:
    print(Debug.Print "Design check: LogName can't ever be NullString when closing ???")
    # Debug.Assert False
    else:
    # .LogNamePrev = .LogName
    if LenB(.LogNameNext) > 0 And InStr(.LogNameNext, Left(DateIdNB, 8)) = 0 Then:
    # msg = "Switching to new day, old Logfile " & .LogName
    # .LogNameNext = vbNullString               ' force new time stamp
    # KeepName = False
    if .LogNameNext = .LogName And Not KeepName Then:
    # .LogNameNext = vbNullString               ' force new time stamp
    # msg = "Next Logfile will get new timestamp"
    elif LenB(.LogNameNext) > 0 Or KeepName Then:
    if KeepName Then:
    # .LogNameNext = .LogName
    # msg = "Next Logfile will not change, to be re-opened: " & .LogName
    else:
    # msg = "Next selected Logfile: " & .LogNameNext
    else:
    if InStr(.LogName, Left(DateIdNB, 8)) = 0 Then:
    # msg = "Restarting for new day, old Logfile " & .LogName
    # KeepName = False
    else:
    # msg = "Logfile closed:    " & .LogName
    # .LogName = vbNullString

    # msg = String(OffCal, b) & msg
    try:
        if .LogIsOpen Then:
        # Print #1, msg
        print(Debug.Print msg)
        # Close #1
        # .LogName = vbNullString
        # .LogIsOpen = False
        # .LogFileLen = 0
        # End With ' T_DC

        # FuncExit:
        # Recursive = False

        # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowLog
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showlog():
    if OlSession Is Nothing Then:
    # Call N_CheckStartSession(False)
    # Call ShowLogWait(False)

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowLogWait
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showlogwait():

    # Dim DateNow As String
    # Dim daFilt As String
    # Dim nDigits As String
    # Dim NextFile As String

    # With T_DC
    if LenB(.LogName) = 0 Then:
    # DateNow = GetDateId(-2)
    # getDigits:
    # Call ClipBoard_SetData(DateNow)
    # nDigits = InputBox( _
    # "Optional: Geben Sie an, auf wieviel Stellen genau die Log-Namen mit " _
    # & vbCrLf & "   " & DateNow & "*.log bereinstimmen sollen (1-" & Len(DateNow) - 2 & ")" _
    # & vbCrLf & "oder geben sie den gewnschten Teil des Namens ein" _
    # & vbCrLf & " sie knnen mit Ctrl-v den Wert direkt einfgen", _
    # "Mehrfachdatei-ffnen")
    if IsNumeric(nDigits) Then:
    if CDbl(nDigits) > Len(DateNow) - 2 Then:
    # daFilt = nDigits & "*.log"
    else:
    # daFilt = Left(DateNow, nDigits) & "*.log"
    # GoTo DoRunEditor
    else:
    # daFilt = Left(DateNow, 10) & "*.log"
    # GoTo DoRunEditor
    # & " der Logdatei bereinstimmen sollen, also " _
    # & daFilt & vbCrLf & "(in " & lPfad & ")" _
    # & vbCrLf & "   Erneut versuchen==>Nein", _
    if rsp = vbCancel Then:
    print(Debug.Print "Unable to edit log file because name is not known")
    # DoVerify False
    elif rsp = vbNo Then:
    # GoTo getDigits
    elif rsp = vbYes Then:
    # DoRunEditor:
    # Call RunEditor(lPfad & "Outlook VBA " & daFilt)
    else:
    # Call CloseLog(KeepName:=True)
    # Call RunEditor(.LogNamePrev)
    # End With ' T_DC


# '---------------------------------------------------------------------------------------
# ' Method : Sub RunEditor
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def runeditor():

    # Const zKey As String = "File_FolderOps.RunEditor"
    # Call DoCall(zKey, tSub, eQzMode)

    # Call N_Suppress(Push, zKey)

    # ' Open and Close a File
    # Dim vPID As Variant

    # ' Launch file
    if LenB(FullFileName) = 0 Then:
    # GoTo zExit

    # vPID = Shell(Quote(EditProg) & b & Quote(FullFileName), vbNormalFocus)

    # & vbCrLf & " Press Yes to Terminate Editor" _
    # & vbCrLf & " Press No to ignore: continue running Editor and Macros," _
    # & vbCrLf & " Press Cancel to Debug and optionally Terminate" _
    # , vbYesNoCancel)
    match rsp:
        case vbYes:
    # ' Kill file
    # Call Shell("TaskKill /F /PID " & CStr(vPID), vbHide)
        case vbCancel:
    # DoVerify False
    # Call TerminateRun

    # ProcReturn:
    # Call N_Suppress(Pop, zKey)

    # zExit:
    # Call DoExit(zKey)



# '---------------------------------------------------------------------------------------
# ' Method : Function CreateOpenFileEntry
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createopenfileentry():
    # Const zKey As String = "File_FolderOps.CreateOpenFileEntry"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Call FileEntryExists(Key, PathAndName, makeOne, reqState)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function FileEntryExists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def fileentryexists():
    # Const zKey As String = "File_FolderOps.FileEntryExists"
    # Call DoCall(zKey, tFunction, eQzMode)

    if OpenFileNames Is Nothing Then:
    # Set OpenFileNames = New Dictionary
    if ClosedFileNames Is Nothing Then:
    # Set ClosedFileNames = New Dictionary
    if FileStates Is Nothing Then:
    # Set FileStates = New Dictionary

    # DoVerify Key <> inv, "*** invalid Key: " & Key
    # Key = Abs(Key)                                          ' file numbers must be >= 0
    # aFileSpec = CStr(inv)                                   ' undefined path and name

    if OpenFileNames.Exists(Key) Then:
    # aFileSpec = OpenFileNames.Item(Key)
    if aFileSpec <> PathAndName Then:
    # aOpenKey = inv
    # GoTo NoFit
    # aOpenKey = Key
    if ClosedFileNames.Exists(Key) Then:
    # aFileSpec = ClosedFileNames.Item(Key)
    if aFileSpec <> PathAndName Then:
    # aClosedKey = inv
    # GoTo NoFit

    # aBugVer = aOpenKey <> inv And aClosedKey <> inv
    # aBugTxt = "File Key=" & Key & " is both open and closed ??? " & aFileSpec & " aFileState=" & aFileState
    if DoVerify Then:
    if aFileState <> inv Then:
    # Call ClosedFileNames.Remove(Key)
    # Call CloseFile(Key)
    # Call OpenFileNames.Remove(Key)          ' incomplete fix
    # aFileState = 0
    # aFileSpec = inv
    if aOpenKey <> inv Then:
    # aClosedKey = inv
    else:
    # aOpenKey = inv

    # NoFit:
    if makeOne Then:
    # aBugVer = Mid(PathAndName, 2, 2) = ":\"
    # aBugTxt = "the PathName incorrect: '" & Quote(PathAndName)
    # DoVerify

    if aClosedKey <> inv Then:
    # FileEntryExists = False
    elif aOpenKey <> inv Then:
    # FileEntryExists = True
    else:
    # FileEntryExists = False
    if FileEntryExists Then:
    if FileStates.Exists(Key) Then:
    # aFileState = FileStates.Item(Key)
    else:
    # Call FileStates.Add(Key, "New")

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub openForAccess
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def openforaccess():

    # Const zKey As String = "File_FolderOps.openForAccess"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim modeAcc As Variant

    # modeAcc = split(OpenMode)
    if UBound(modeAcc) > 0 Then:
    print(Debug.Print OpenMode & " not implemented yet, too many situations to cover")
    # DoVerify False

    match LCase(modeAcc(0)):
        case "append":
    # Open FullPath For Append As #nFile
        case "binary":
    # Open FullPath For Binary As #nFile
        case "input":
    # Open FullPath For Input As #nFile
        case "output":
    # Open FullPath For Output As #nFile
        case "random":
    # Open FullPath For Random As #nFile
        case _:
    print(Debug.Print OpenMode & " is invalid file open mode")
    # DoVerify False

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub openFile
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def openfile():
    # fName As String, fExt As String, _
    # Optional OpenMode As String = "Append")

    # Const zKey As String = "File_FolderOps.openFile"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim FullPath As String

    # FullPath = fPath & "\" & fName & fExt
    if nFile = 0 Then:
    # nFile = OpenFileNames.Count + 1 ' next free slot
    # Call FileEntryExists(nFile, FullPath, makeOne:=True, reqState:=OpenMode)

    if aFileState = "New" Then:
    # GoTo noCl

    if aFileState <> OpenMode Then:
    if aFileState <> "Closed" Then:
    # Print #nFile, Quote(aFileSpec) & "File will be closed and reopened with mode=" & OpenMode
    # Err.Clear
    if Not CloseFile(nFile) Then:
    # GoTo noCl                       ' close did not work ...

    # noCl:
    # Print #nFile, Quote(aFileSpec) & "*** attempting File open with mode=" & OpenMode

    if aFileState <> OpenMode Then:
    # E_Active.Permit = "*"
    # Call openForAccess(FullPath, OpenMode, nFile)
    if ErrorCaught = 0 Then:
    if DebugMode Or DebugLogging Then:
    print(Debug.Print "Open " & FullPath _)
    # & " for " & OpenMode _
    # & " As #" & nFile & " successful"
    # aFileState = OpenMode
    else:
    print(Debug.Print "Open " & FullPath _)
    # & " for " & OpenMode _
    # & " As #" & nFile & " failed, Error " _
    # & Err.Number & ": " & Err.Description
    # Err.Clear
    # aFileState = "Undefined"
    else:
    if DebugMode Then:
    # DoVerify False, "*** file is open already and no reopen specified"

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetFileInDir
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getfileindir():

    # Const zKey As String = "File_FolderOps.GetFileInDir"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim thisFile As String
    # Dim thisSuffix As String
    # Dim aFileNumber As Long
    # Dim maxFileNumber As Long
    # Dim nextFileNumber As Long
    # Dim aFileDate As Date
    # Dim maxFileDate As Date
    # Dim fileCount As Long
    # Dim RealPath As String
    # Dim FoundFiles As Variant
    # Dim isMaxDate As Boolean
    # Dim j As Long
    # maxFileBefore = vbNullString
    # maxFileNumber = 0
    # maxFileDate = CDate(0)

    # FoundFiles = getDirFileList(MasterPath, fileCount, "\" & FileNamePrefix & CodeNamePattern, False)
    if fileCount = 0 Then:
    print('keine Dateien gefunden, die ')
    # End
    # RealPath = MasterPath
    # thisFile = FoundFiles(j)
    if InStr(thisFile, "\") > 0 Then:
    # thisFile = RTail(thisFile, "\", RealPath)
    # thisFile = Trunc(1, thisFile, ".") ' drop "extension" (left to right!)
    if Left(thisFile, Len(FileNamePrefix)) = FileNamePrefix Then:
    # thisSuffix = Mid(thisFile, Len(FileNamePrefix) + 1)
    if Not isMaxDate And IsNumeric(thisSuffix) Then ' use numeric only if we have no dates:
    # aFileNumber = CLng(thisSuffix)
    if aFileNumber > maxFileNumber Then:
    # maxFileNumber = aFileNumber
    # maxFileBefore = thisFile
    elif IsDate(thisSuffix) Then:
    # isMaxDate = True    ' no longer loook for numbers
    # aFileDate = CDate(thisSuffix)
    if aFileDate > maxFileDate Then:
    # maxFileBefore = thisFile
    # maxFileDate = aFileDate
    else:
    # DoVerify False
    if isMaxDate Then:
    # GetFileInDir = RealPath & "\" & Now()
    else:
    # GetFileInDir = RealPath & "\" & LPad(CStr(maxFileNumber + 1), 8)
    # ' add extension
    # GetFileInDir = GetFileInDir & Mid(CodeNamePattern, 2)


# '---------------------------------------------------------------------------------------
# ' Method : Function getMasterPath
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getmasterpath():

    # Const zKey As String = "File_FolderOps.getMasterPath"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim Count As Long
    # Dim thisFile As String
    # Dim FoundFiles As Variant

    # thisFile = RTail(olApp.RecentFiles(1).Path, "\", startOfPath)
    if thisFile = "Briefe" Then:
    # getMasterPath = startOfPath & "\Briefe"
    else:
    # FoundFiles = getDirFileList(startOfPath, Count, "\*.doc*", True)       ' doc or docx, only first file
    if Count = 0 Then:
    print('keinen Ordner ')
    # End
    else:
    # thisFile = FoundFiles(0)
    # thisFile = RTail(thisFile, "\", getMasterPath)


# '---------------------------------------------------------------------------------------
# ' Method : Function getDirFileList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getdirfilelist():

    # Const zKey As String = "File_FolderOps.getDirFileList"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim MyFile As String
    # ReDim DirListArray(1000) As String

    # MyFile = Dir$(dirName & fileType)
    # fileCount = 0
    # Do While MyFile <> vbNullString
    # DirListArray(fileCount) = MyFile
    # MyFile = Dir$
    # fileCount = fileCount + 1
    if firstFileOnly Or fileCount > 1000 - 1 Then:
    # Exit Do
    # Loop

    # fileCount = fileCount - 1
    # ' Reset the size of the array without losing its values by using Redim Preserve
    # ReDim Preserve DirListArray(fileCount)
    # Application.WordBasic.sortarray DirListArray()
    # getDirFileList = DirListArray


# '---------------------------------------------------------------------------------------
# ' Method : CloseFile
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Close file #nFile and note this in OpenFileNames
# '---------------------------------------------------------------------------------------
def closefile():
    # Const zKey As String = "File_FolderOps.CloseFile"
    # Call DoCall(zKey, tSub, eQzMode)

    # E_Active.Permit = "*"
    # Close #nFile
    if ErrorCaught <> 0 Then:
    # Err.Clear
    # CloseFile = False
    # CloseFile = True
    # ClosedFileNames.Add nFile, OpenFileNames.Item(nFile)
    # aFileState = "Closed"
    # Call OpenFileNames.Remove(nFile)
    # FileStates.Item(nFile) = aFileState

    # zExit:
    # Call DoExit(zKey)


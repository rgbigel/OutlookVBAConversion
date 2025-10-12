# Converted from HelpersOL.py

# Attribute VB_Name = "HelpersOL"
# Option Explicit

# ' other stuff
# Public AddType As String        ' task type added by eventroutines below

# ' mail-like objects, just to make readable
# Public CurrentSessionEmail As MailItem
# Public CurrentSessionReport As ReportItem
# Public CurrentSessionMeetRQ As MeetingItem
# Public CurrentSessionTaskRQ As TaskItem

# Public SpecialSearchComplete As Boolean
# Public forceOrdering As Boolean
# Public NeverDelete As Boolean
# Private DeleteAllOld As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Sub Z_olInits
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def z_olinits():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                        ' Standard procConst zKey As String = "HelpersOL.Z_olInits"
    # Const zKey As String = "HelpersOL.Z_olInits"

    # Dim sBlockEvents As Boolean
    # Dim i As Long

    # Call DoCall(zKey, "Sub", eQzMode)

    # sBlockEvents = E_Active.EventBlock        ' Save for exit state
    # E_Active.EventBlock = True                ' no events during init
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call BugEval
    # StopLoop = False

    if OlSession Is Nothing Then:
    # Stop ' *****imposs
    # Set OlSession = New cOutlookSession
    if dftRule Is Nothing Then:
    # Set dftRule = New cAllNameRules
    # ActionTitle(0) = "Undefined or dynamic action"
    # StopRecursionNonLogged = StopRecursionNonLogged Or EPCalled
    # EPCalled = IsEntryPoint

    # Set RestrictedItemCollection = New Collection
    # With dftRule                                 ' First-Time Inits if invalid:
    if Not dftRule.RuleInstanceValid Then:
    # Call getAccountDscriptors
    # ' contains all properties that can not be decoded
    if aPindex = 0 And aID(0) Is Nothing Then:
    # Set aOD(0) = New cObjDsc
    # Set aOD(0).objSeqInImportant = New Collection
    if D_TC Is Nothing Then:
    # Set D_TC = New Dictionary
    # D_TC.Add "0", aOD(0)
    # aOD(0).objItemType = inv         ' mark this as invalid
    # Set aOD(0).objClsRules = dftRule ' NO cloning here!
    # Set aOD(0).objClsRules.clsNeverCompare.PropAllRules = dftRule
    # Set aOD(0).objClsRules.clsObligMatches.PropAllRules = dftRule
    # Set aOD(0).objClsRules.clsNotDecodable.PropAllRules = dftRule
    # Set aOD(0).objClsRules.clsSimilarities.PropAllRules = dftRule
    if Left(ActionTitle(0), 5) = IgnoredHeader Then:
    print(Debug.Print ActionTitle(0))
    # DoVerify False
    # Call SetStaticActionTitles
    if LenB(ActionTitle(UBound(ActionTitle))) = 0 Then:
    # Call SetStaticActionTitles
    # Set LookupFolders = aNameSpace.Folders

    if LF_CurActionObj Is Nothing Then:
    # Call DefineLocalEnvironment
    # CurIterationSwitches.SaveItemRequested = True ' for mail processing, Action = 3

    # ' Obligatory Matches: the only case where we
    # ' access class Property "aRuleString" is to set as default
    # TrueCritList = "Subject"
    # .clsObligMatches.ChangeTo = TrueCritList
    # ' Not Decodable
    # .clsNotDecodable.ChangeTo = "ItemProperties " _
    # & "Session RTFBody FormDescription PermissionTemplateGuid " _
    # & "GetInspector PropertyAccessor SaveSentMessageFolder IsLatestVersion "
    # ' dont compare
    # DontCompareListDefault = "*ID: Organizer Ordinal *Time* *UTC Size " _
    # & "SenderName SentOn SentOnBehalfOfName " _
    # & "*Version *DisplayName: *Xml " _
    # & "CompanyAnd* CompanyLast* FullNameAnd* " _
    # & "LastNameAnd* Yomi* BusinessCard* " _
    # & "BillingInformation ConversationIndex Mileage Saved " _
    # & "UnRead VotingOptions VotingResponse Parent MessageClass " _
    # & "OutlookVersion OutlookInternalVersion " _
    # & "BodyFormat InternetCodepage Left Top Width Height " _
    # & "Organizer SendUsingAccount GlobalAppointmentID "
    # .clsNeverCompare.ChangeTo = DontCompareListDefault
    # ' Similarities
    # .clsSimilarities.ChangeTo = "Parent Categories"
    # ' mark this initialization done:
    # .RuleType = DefaultRule
    # .clsNeverCompare.RuleMatches = False
    # .clsSimilarities.RuleMatches = False
    # .clsObligMatches.RuleMatches = False
    # .clsNotDecodable.RuleMatches = False
    # .RuleInstanceValid = True            ' and Write-Only now?
    # End With                                     ' dftrule

    # ' Lists must have leading and trailing blank for each word  !
    # Set killWords = Nothing
    # Set killWords = New Collection
    # sourceIndex = 0
    # targetIndex = 0
    # Set sDictionary = Nothing
    # Set aProp = Nothing
    # apropTrueIndex = inv
    # workingOnNonspecifiedItem = False
    # BaseAndSpecifiedDiffer = False
    if TrashFolder Is Nothing Then:
    # Set TrashFolder = aNameSpace.GetDefaultFolder(olFolderDeletedItems)
    # TrashFolderPath = TrashFolder.FolderPath
    # pArr(i) = Chr(0)

    # ' Inits for RuleTable
    # staticRuleTable = True
    # UseExcelRuleTable = False                    ' execute sub InitRuleTable

    # Call InitEventTraps
    # Set Deferred = New Collection

    print(Debug.Print "* EntryPoint=" & IsEntryPoint & "* Testvar=" & Quote(Testvar))
    # E_Active.EventBlock = sBlockEvents        ' all first-time inits were done
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call BugEval

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ClearUnwantedCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clearunwantedcats():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.ClearUnwantedCats"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim inCat As category
    # Dim outCat As category
    # Dim i As Long
    # Dim k As Long
    # Dim anyMod As Boolean
    # Dim thisMod As String
    # ' loop to remove existing invalid category definitions (invalid = not in inStore)
    if i > outStore.Categories.Count Then:
    # Exit For
    # Set outCat = outStore.Categories.Item(i)
    if DeleteAllOld Then                     ' set this to force ordering like in inStore:
    # thisMod = outCat.Name
    # outStore.Categories.Remove i
    # thisMod = i & vbTab & "Category " & Quote(thisMod) & " deleted from Store " & outStore.DisplayName
    # anyMod = True
    else:
    # thisMod = outCat.Name
    # Set inCat = StoreCatGet(inStore, thisMod, k)
    if inCat Is Nothing And Not NeverDelete Then ' no wanted because no valid inCat:
    # outStore.Categories.Remove thisMod
    # thisMod = i & vbTab & "Category deleted " & Quote(thisMod) _
    # & " previously in pos. " & k & " of Store " & outStore.DisplayName
    # anyMod = True
    if anyMod And i <> k Then:
    # DeleteAllOld = True              ' force original order from here on
    # ' will not create full ordering!
    if DebugMode Then:
    print(Debug.Print thisMod)
    # thisMod = vbNullString
    if anyMod Then:
    else:
    if DebugMode Then:
    print(Debug.Print "no unwanted Categories had to be removed")

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CompareCatNames
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def comparecatnames():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.CompareCatNames"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim returnvalue As Long
    # ' 0 if identical
    # ' +1 if not identical
    # ' -1 if not in same sequence
    # Dim i As Long
    # Dim k As Long
    # Dim mini As Long
    # Dim maxi As Long
    # Dim inCatName As String
    # Dim outCatName As String
    # Dim inCat As category
    # Dim outCat As category
    # Dim bigCatStore As Store
    # Dim compMsg As String

    # 'On Error GoTo 0     ' allow no errors in this proc
    if inStore.Categories.Count > outStore.Categories.Count Then:
    # mini = outStore.Categories.Count
    # maxi = inStore.Categories.Count
    # Set bigCatStore = inStore
    # compMsg = " more items in InStore " & Quote(inStore.DisplayName) _
    # & " than in " & Quote(outStore.DisplayName)
    # returnvalue = 1
    elif inStore.Categories.Count < outStore.Categories.Count Then:
    # returnvalue = 1
    # mini = inStore.Categories.Count
    # maxi = outStore.Categories.Count
    # Set bigCatStore = outStore
    # compMsg = " fewer items in InStore " & Quote(inStore.DisplayName) _
    # & " than in " & Quote(outStore.DisplayName) & " will delete=" & outStoreDelete
    else:
    # mini = inStore.Categories.Count
    # maxi = mini

    if DebugMode And LenB(compMsg) > 0 Then:
    print(Debug.Print (maxi - mini) & compMsg)

    # k = 1                                        ' initally = i
    # i = 1
    # Do                                           ' find by items in outStore
    if i > outStore.Categories.Count Then:
    # outCatName = vbNullString
    else:
    # Set outCat = outStore.Categories.Item(i)
    # outCatName = outCat.Name
    if k > inStore.Categories.Count Then:
    # GoTo gotNone
    # Set inCat = inStore.Categories.Item(k)   ' comparing inStore
    # inCatName = inCat.Name
    if LenB(outCatName) = 0 Then:
    # compMsg = i & vbTab & inCatName & " occurs in Store " & Quote(inStore.DisplayName) & " in position " & k & " but does not exist in store " & Quote(outStore.DisplayName)
    # k = k + 1                            ' step k
    if returnvalue = 0 Then:
    # returnvalue = -1
    # GoTo NotInTarget
    if inCatName <> outCatName Then:
    # Set inCat = StoreCatGet(inStore, outCatName, k)
    if inCat Is Nothing Then:
    # gotNone:
    # compMsg = i & vbTab & outCatName & " does not occur in Store " & Quote(inStore.DisplayName)
    # returnvalue = 1
    if outStoreDelete Then:
    # outStore.Categories.Remove i
    # maxi = maxi - 1
    # ' no step in k, not present any longer
    # compMsg = compMsg & " and was deleted from Store " & Quote(outStore.DisplayName)
    else:
    # k = k + 1                    ' step k
    else:
    # compMsg = i & vbTab & outCatName & " occurs in Store " & Quote(inStore.DisplayName) & " in position " & k
    if returnvalue = 0 Then:
    # returnvalue = -1
    # k = k + 1                        ' step k
    else:
    # compMsg = i & vbTab & inCatName & " occurs in same position in " & Quote(outStore.DisplayName)
    # k = k + 1                            ' step k
    if DebugMode Then:
    # NotInTarget:
    print(Debug.Print compMsg)
    if k < 1 Then                            ' leave if we deleted the last one:
    # Exit Do
    # i = i + 1
    # Loop Until i > maxi

    if outStoreDelete Then:
    # GoTo fini

    # Set outCat = bigCatStore.Categories.Item(i)
    # outCatName = outCat.Name
    # compMsg = i & vbTab & outCatName & " occurs only in " & Quote(bigCatStore.DisplayName)
    if outStoreDelete And Not NeverDelete And bigCatStore.DisplayName = outStore.DisplayName Then:
    # outStore.Categories.Remove i
    # i = i - 1                            ' not present any longer
    # maxi = maxi - 1
    # compMsg = compMsg & ". It has been deleted"
    if DebugMode Then:
    print(Debug.Print compMsg)

    # returnvalue = 1
    if i < 1 Then                            ' leave if we deleted the last one:
    # Exit For

    # fini:
    # CompareCatNames = returnvalue

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyAllBackupCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyallbackupcats():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "HelpersOL.CopyAllBackupCats"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="HelpersOL")

    # Dim sourceStore As Store

    # IsEntryPoint = True
    # Set sourceStore = FolderBackup.Store         ' global for session, from Z_AppEntry
    # Call ShowCats(sourceStore)
    # forceOrdering = False
    # Call CopyAllCats(Array(OlHotMailHome, OlWEBmailHome, OlGooMailHome), sourceStore)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyAllCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyallcats():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.CopyAllCats"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim targetStore As Store
    # Dim TargetFolder As Folder
    # Dim i As Long
    # NeverDelete = Not forceOrdering
    # Set TargetFolder = targetStoreArray(i)
    # Set targetStore = TargetFolder.Store
    # Call CopyCats(targetStore, sourceStore)

    # FuncExit:

    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyAllHotmailCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copyallhotmailcats():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "HelpersOL.CopyAllHotmailCats"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="HelpersOL")

    # Dim sourceStore As Store

    # IsEntryPoint = True

    # Set sourceStore = OlHotMailHome.Store        ' global for session, from Z_AppEntry
    # Call ShowCats(sourceStore)
    # forceOrdering = True
    # Call CopyAllCats(Array(FolderBackup, OlWEBmailHome, OlGooMailHome), sourceStore)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CopyCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def copycats():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.CopyCats"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim inCat As category
    # Dim outCat As category
    # Dim j As Long
    # Dim k As Long
    # Dim anyMod As Boolean
    # Dim thisMod As String
    # Dim oIdent As String
    # anyMod = False
    # DeleteAllOld = forceOrdering
    if DebugMode Then:
    # Call ShowCats(outStore)

    if DeleteAllOld Then:
    # Call ClearUnwantedCats(outStore, inStore)
    else:
    # Call CompareCatNames(outStore, inStore, outStoreDelete:=True)

    # ' loop to add / alter new category definitions
    # k = 1                                        ' position in outstore
    # thisMod = vbNullString
    # Set inCat = inStore.Categories.Item(j)
    # oIdent = j & vbTab & inCat.Name & vbTab & " in " & Quote(outStore.DisplayName)
    if j > outStore.Categories.Count Then    ' cat obviously missing::
    # Set outCat = StoreCatGet(outStore, inCat.Name, k)
    if outCat Is Nothing Then:
    # outStore.Categories.Add inCat    ' add to end of list again
    # k = outStore.Categories.Count
    # thisMod = "Category added as number " & k
    else:
    # DoVerify False, " incorrect ordering in outStore"
    # DeleteAllOld = True              ' disorder???
    # thisMod = "Category already present in pos. " & k
    else:
    # Set outCat = StoreCatGet(outStore, inCat.Name, k)
    if outCat Is Nothing Then:
    # outStore.Categories.Add inCat    ' might change sequence?
    # k = outStore.Categories.Count
    # thisMod = "New Category added as number " & k
    else:
    if j <> k Then:
    # DeleteAllOld = True          ' disorder???
    if outCat.ShortcutKey <> inCat.ShortcutKey Then:
    # Call AppendTo(thisMod, vbTab _
    # & "corrected ShortcutKey mismatch " _
    # & outCat.ShortcutKey _
    # & " to " & inCat.ShortcutKey)
    # outCat.ShortcutKey = inCat.ShortcutKey
    if outCat.Color <> inCat.Color Then:
    # Call AppendTo(thisMod, vbTab _
    # & "corrected Color mismatch " _
    # & outCat.Color _
    # & " to " & inCat.Color)
    # outCat.Color = inCat.Color
    if LenB(thisMod) = 0 Then:
    if DebugMode Then:
    print(Debug.Print oIdent & vbTab & thisMod & vbTab _)
    # & "no change, ShortcutKey " & vbTab _
    # & outCat.ShortcutKey _
    # & ", Color " & outCat.Color _
    # & " as number " & k
    else:
    # anyMod = True
    if DebugMode Then:
    print(Debug.Print oIdent & thisMod)

    if DebugMode Then:
    if outStore.Categories.Count <> inStore.Categories.Count Then:
    # DoVerify False
    print(Debug.Print "There is a difference in the number of categories: Backup has " _)
    # & inStore.Categories.Count & ", target has " _
    # & outStore.Categories.Count
    if anyMod Then:
    # ' save? outStore ???

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : CreateSearchFolder
# ' Author : Mark Withall @markwithall.com
# ' Date   : 20211108@11_47
# ' Purpose: Create Search Folder
# '---------------------------------------------------------------------------------------
def createsearchfolder():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "HelpersOL.CreateSearchFolder"
    # Const zKey As String = "HelpersOL.CreateSearchFolder"

    # Call DoCall(zKey, "Function", eQzMode)

    # Dim CreateOk As Boolean
    # Dim objSearch As Search
    # Dim Warn As String
    # Dim sFolder As Folder

    if olApp Is Nothing Then:
    # Call Z_StartUp

    if LenB(ScopeFolders) > 0 Then:
    if LenB(Account) = 0 Then:
    # Account = Replace(ScopeFolders, "'", vbNullString)
    # Account = Trunc(3, Account, "\")     ' single Store only
    # Warn = "Account defaulted to " & Account
    else:
    # aBugTxt = "Error in ScopeFolcers, not matching account"
    # DoVerify InStr(LCase(ScopeFolders), "\\" & LCase(Account) & "\") > 0
    else:
    if LenB(Account) = 0 Then:
    # Account = olApp.Session.Stores.Item(1)
    # Warn = " in first Store"
    # ScopeFolders = "'\\" & Account & "\" & StdInboxFolder & "'"
    # Warn = "ScopeFolders defined as " & ScopeFolders & Warn

    if LenB(Warn) > 0 Then:
    print(Debug.Print Warn)

    if SearchFolderExists(Account, FolderName, sFolder) Then:
    if LenB(Warn) > 0 Then:
    # Call LogEvent("Would like to test of Scopes Match expectations, " _
    # & "but no access method found for 'Search' Objects")
    if sFolder.ShowItemCount <> olShowTotalItemCount Then:
    # Call LogEvent("'.ShowItemCount' indicates that SearchFolder " _
    # & FolderName & vbCrLf & "    in " & sFolder.FolderPath _
    # & " may not be synchronized." & vbCrLf & "    Re-Creating it.")
    # GoTo DefineAgain
    # Call LogEvent("SearchFolder " & Quote(sFolder.FolderPath) _
    # & " is up to date and usable")
    # CreateSearchFolder = False
    else:
    # DefineAgain:
    # aBugTxt = "set up Advanced search, folder '" & FolderName _
    # & "' in " & ScopeFolders
    # Call Try
    # Set objSearch = olApp.AdvancedSearch( _
    # Scope:=ScopeFolders, _
    # Filter:=Filter, _
    # SearchSubFolders:=True, _
    # Tag:=FolderName)
    # aBugTxt = "Save search folder '" & FolderName _
    # & "' in " & Account
    # Call Try                                 ' Try anything, autocatch
    # Call objSearch.Save(FolderName)
    # Catch

    # CreateSearchFolder = True
    # CreateOk = SearchFolderExists(Account, FolderName, sFolder)
    if Not CreateOk Then:
    # Call LogEvent("Missing search folder '" & FolderName & "' in " & Account)
    if Not sFolder Is Nothing Then:
    # sFolder.ShowItemCount = olShowTotalItemCount ' default for MY search folders
    # GoTo FuncExit
    # Call LogEvent("created search folder '" & FolderName & "' in " & Account _
    # & vbCrLf & "may need to wait for AdvancedSearchComplete-Event")

    # UInxDeferred = UInxDeferred + 1
    # Set DeferredFolder(UInxDeferred) = sFolder

    # FuncExit:
    # Set sFolder = Nothing
    # Set objSearch = Nothing

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DoMaintenance
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub DoMaintenance()
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "HelpersOL.DoMaintenance"
# Call DoCall(zKey, "Sub", eQzMode)

# frmMaintenance.Show
# MaintenanceAction = 0

# FuncExit:

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ForceSave
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def forcesave():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.ForceSave"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim defectiveCopy As Variant
    # Dim logStatus As Boolean
    # Dim mgPart As String
    # Dim waitRetries As Long

    if aItemO Is Nothing Then:
    print(Debug.Print "can't save a Nothing Object")
    # DoVerify False
    # GoTo ProcReturn
    if aItemO.Saved Then:
    # GoTo ProcReturn                          ' dont save if superfluous
    # logStatus = dontLog

    if TrySaveItem(aItemO) Then                  ' klappt nicht, wenn von Outlook inzwischen gendert wurde:
    # aBugTxt = "Replicate Item EntryID " & aItemO.EntryID
    # Call Try                                 ' Try anything, autocatch
    # Set defectiveCopy = Replicate(aItemO, delOriginal:=True)
    # Catch
    # Set aItemO = defectiveCopy
    if aItemO Is Nothing Then:
    # GoTo FunExit
    # dontLog = True
    if Not logStatus Then:
    if Left(aItemO.Parent.Name, 3) = "Log" Then:
    # Call LogEvent(MessagePrefix & "Conflicting Log copied to trashFolder " _
    # & TrashFolder.FullFolderPath)
    else:
    # Call LogEvent(MessagePrefix & "Mail copied to trashFolder " _
    # & TrashFolder.FullFolderPath, eLmin)
    # Set defectiveCopy = Nothing
    else:
    # waitTry:
    # dontLog = True
    # ' if log wanted and not saving the log itself, log save result
    if Not logStatus And Left(aItemO.Parent.Name, 3) <> "Log" Then:
    if aItemO.Saved Then:
    # mgPart = vbNullString
    else:
    # waitRetries = waitRetries + 1
    if waitRetries < 5 Then:
    # Call Sleep(100)              ' wait .1 seconds, maximally 4 times
    # GoTo waitTry
    elif waitRetries <= 5 Then:
    # aItemO.Save
    # Call Sleep(1000)             ' wait 1 seconds, one more time
    # GoTo waitTry
    # mgPart = "not "
    # Call LogEvent("     " & MessagePrefix & TypeName(aItemO) & " changes " & mgPart & "saved after " & waitRetries & " retries in " _
    # & aItemO.Parent.FullFolderPath, eLall)
    # FunExit:
    # dontLog = logStatus

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GenerateTaskReminder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def generatetaskreminder():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.GenerateTaskReminder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim newReminder As TaskItem
    # Dim i As Long
    # Dim j As Long
    # Dim dueDates As String
    # Dim DueDate As Date

    try:
        if Item.Class = olMail Then                  ' nur bei Markierten, gesendeten eMails:
        # ' nicht bei anderen Mail-like objekten
        if Item.FlagStatus = olFlagMarked Then:
        if DebugMode And DebugLogging Then DoVerify False, " ??? test phase":
        # aBugTxt = "add new reminder task"
        # Call Try                             ' Try anything, autocatch, Err.Clear
        # Set newReminder = FolderTasks.Items.Add(olTaskItem)
        if CatchNC Then:
        # GoTo dontDoIt
        if Item.Recipients.Count > 0 Then:
        # newReminder.Assign               ' Zugeordnet!
        # aBugTxt = "add reminder recipient " & i
        # Call Try                     ' Try anything, autocatch
        # newReminder.Recipients.Add (Item.Recipients(i))
        if Catch Then:
        # GoTo dontDoIt
        # Call ErrReset(0)

        # newReminder.Subject = Item.Subject
        # newReminder.Body = Item.Body

        # ' Prfen, ob ein Flligkeitsdatum drin steht, ggf. vergleichen
        # j = 1
        # i = InStr(j, Item.Body, " erled")
        if i > 0 Then:
        # j = InStr(i, Item.Body, " bis ")
        if j > i Then:
        # dueDates = Mid(Item.Body, j + 4)
        if FindFirstDate(dueDates, DueDate) Then:
        if Item.ReminderSet Then ' doppelte Angabe, ZWEI Wecker gesetzt? Prfe Sinn!:
        if DueDate > Item.ReminderTime Then ' frher wecken als Ergebnis erwartet? Seltsam:
        # uAnswer = _
        # & " liegt vor der Zeit, die dem Empfnger genannt ist." _
        # & " Wenn dies nicht OK ist, drcken sie 'Abbrechen'." _
        # & " Dann wird die in der eMail genannte Zeit als Erinnerungszeit verwendet.", _
        # vbOKCancel)
        if uAnswer = vbCancel Then:
        # newReminder.DueDate = DueDate ' war wohl ein Versehen, nimm das, was im Text steht
        else:
        # newReminder.DueDate = Item.ReminderTime
        else:
        # newReminder.DueDate = Item.ReminderTime
        else:
        # newReminder.DueDate = DueDate ' also nimm das Datum in der Mail
        else:
        # & "haben aber kein Datum angegeben. " _
        # & "Sie sollten die gleich angezeigte Aufgabe " _
        # & "und/oder die eMail ergnzen.")
        # Call LogEvent("Mail displayed to user ", eLnothing)
        # newReminder.Display                  ' save is user matter
        # Set newReminder = Nothing
        # dontDoIt:
        if Catch Then:
        # Call LogEvent("Fehler beim Erzeugen einer Aufgabe/Erinnerung", eLall)

        # FuncExit:
        # Call ErrReset(0)

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub getInfo
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Use/Get/Set Information in cInfo. Uses ActItemObject to which fInfo belongs
# '          if Assign, evaluate aValue of variant/object.
# '          if AssignmentMode <=0, determine AssignmentMode 1 or 2
# '          if AssignmentMode = 1, convert to String
# '          if AssignmentMode = 2, use set to assign
# '             and recursively DrillDown the value.
# ' as soon as AssignmentMode =1 is reached and Assign is True, assign aValue=cstr(.iValue)
# '---------------------------------------------------------------------------------------
def getinfo():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.getInfo"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim dInfo As cInfo
    # Dim vElement As Long
    # Dim vValue As Variant

    if fInfo Is Nothing Then:
    # Set fInfo = New cInfo                    ' Class_Initialize d
    elif fInfo.iType = -1 Then                 ' use this if new set of values needed:
    # Set fInfo = New cInfo                    ' Class_Initialize to clear

    # Set dInfo = fInfo                            ' save outermost for recursion
    # With fInfo
    if Assign Then:
    # .DecodedStringValue = vbNullString
    # .DecodeMessage = vbNullString
    if .iAssignmentMode <= 0 Then:
    # Call InspectType(aValue, dInfo)
    # GoTo assignIvalue
    elif .iIsArray Then                    ' arrays need complex analysis every time:
    # Call InspectType(aValue, dInfo)
    # GoTo assignIvalue
    else:
    # assignIvalue:
    if .iAssignmentMode = 0 Then:
    # Assign = False
    elif .iAssignmentMode = 1 Then:
    # .iValue = aValue                 ' scalar
    elif .iAssignmentMode = 2 Then:
    # Set .iValue = aValue             ' not scalar / object

    # ' Note: analyzing .iType may give different results to analyzing .iTypeName via IsScalar
    match .iType:
        case vbInteger, vbLong, vbSingle, vbDouble, vbString, vbBoolean, vbDecimal, 20:
    # ' all of these are scalars:    2 - 6, 11, 14, 20                       (20=LongLong)
    if .iScalarType < 0 Then             ' not an object having String Value at the same time:
    # GoTo noVName                     ' they may not have an aValue.Name
    if aValue.Name = "MemberCount" Then  ' except this one: is an array without count:
    # ' requires GetMember to access the Objects
    # ' which are of type Recipient, NOT Contact
    # .iArraySize = .iValue
    # .DecodeMessage = "# object array (Contact) "
    # .DecodedStringValue = "{} " & .iArraySize & " values"

    # aBugTxt = "decoding Members of DistributionList"
    # Set vValue = ActItemObject.GetMember(vElement)
    # Set dInfo = dInfo.DrillDown(vValue)
    # .DecodedStringValue = .DecodedStringValue & vbCrLf _
    # & dInfo.iValue.Name
    # GoTo FuncExit
    # noVName:
    if .iDepth >= MaxDepth Then          ' not going any deeper:
    # .DecodedStringValue = CStr(dInfo.iValue)
    # .DecodeMessage = "# non-scalar: " & dInfo.iTypeName
    # .iAssignmentMode = 1
    elif Assign Then:
    # .DecodedStringValue = CStr(.iValue)
    if .iScalarType <= 0 Then:
    # .DecodeMessage = "# non-scalar " & .iTypeName
    if .iClass = olItemProperty Then:
    # .DecodeMessage = .DecodeMessage & ", Name=" & .iValue.Name
    else:
    # .DecodeMessage = .iTypeName
    # GoTo FuncExit
        case vbDate                              '  7 Datumswert (Date):
    if Assign Then:
    # .DecodedStringValue = CStr(.iValue)
    if .iValue = BadDate Then        ' obviously must exist:
    # .DecodedStringValue = "## Datum nicht angegeben"
    # .DecodeMessage = .DecodedStringValue
    else:
    # .DecodeMessage = .iTypeName
    # GoTo FuncExit
        case vbObject                            '  9 Objekt:
    # Set .iValue = aValue
    # likeObject:
    if .iDepth > MaxDepth Then:
    # .DecodeMessage = "not decoded, depth " & MaxDepth
    # .DecodedStringValue = "## unvollstndige Decodierung"
    # GoTo nonValue
    if .iScalarType < 0 Then             ' not decodable:
    # .DecodeMessage = "not decodable by Rule: "
    # .DecodeMessage = .DecodeMessage & .iValue.Name
    # .DecodedStringValue = "## " & .iValue.Name & " nicht dekodierbar"
    # Call ErrReset(4)                 ' no problem if object has no name
    # GoTo RuleDetermined
    elif Not aTD.adRules.clsObligMatches.RuleMatches Then:
    # .DecodeMessage = "# not oblig.: " & .iValue.Name
    # .DecodedStringValue = "## " & .iValue.Name & " nicht oblig."
    # .iAssignmentMode = -2
    # Call ErrReset(4)                 ' no problem if object has no name
    # GoTo RuleDetermined
    if fInfo.iClass = olItemProperty Then '  all have a .Name:
    if DecodeSpecialProperties(fInfo, .iValue.Name) Then:
    if fInfo.iIsArray Then       ' qualified because fInfo may have changed:
    # fInfo.DecodedStringValue = dInfo.DecodedStringValue
    # Set dInfo = fInfo
    # Set dInfo = dInfo.DrillDown(fInfo.iValue.Item(vElement))
    # Call getInfo(dInfo, dInfo.iValue, Assign, MaxDepth:=0)
    if fInfo.iTypeName = "ActionsCount" Then ' Select Case? if more with iValue:
    if LenB(dInfo.DecodedStringValue) > 0 Then:
    if dInfo.iValue.Enabled Then:
    # dInfo.DecodedStringValue = "Aktiv:  " & dInfo.DecodedStringValue
    else:
    # dInfo.DecodedStringValue = "Passiv: " & dInfo.DecodedStringValue
    if LenB(dInfo.DecodedStringValue) > 0 Then:
    # fInfo.DecodedStringValue = fInfo.DecodedStringValue & vbCrLf _
    # & dInfo.iTypeName & "(" & vElement & ")=" & dInfo.DecodedStringValue
    # Set dInfo = fInfo            ' restore qualified state, this is the final result
    # dInfo.DecodeMessage = "# " & dInfo.iTypeName
    # GoTo FuncExit
    else:
    # .iClass = aValue.Class           ' .iAssignmentMode always=2 !
    # aBugTxt = "Get Class of aValue"
    if Catch Then:
    # .DecodeMessage = E_Active.Reasoning
    # GoTo FuncExit

    # aBugTxt = "Set value from Object " & .iTypeName
    # Set vValue = aValue.Value            ' Objects with Object value do not cause error
    if Catch Then:
    # .DecodeMessage = E_Active.Reasoning
    match .iClass:
        case olItemProperty              ' Properties always have a Name:
    # DoVerify False, _
    # "Problem with Set ItemProperty Value.Value to Variant: " _
    # & .iTypeName & b _
    # & .iTypeName & "->" & aValue.Name
    # .DecodedStringValue = "# unable to obtain value for " & aValue.Name
    # GoTo FuncExit
        case _:
    # .DecodedStringValue = "** Class not implemented: " & .iClass _
    # & " TypeName: " & .iTypeName
    # DoVerify False, CStr(.DecodedStringValue)
    # GoTo FuncExit

    if vValue Is Nothing Then:
    # .DecodeMessage = "Object with Null value"
    # GoTo nonValue

    # Set dInfo = .DrillDown(vValue)       ' note: .DecodedStringValue never changed on this level
    # .DecodeMessage = "get down Value, depth=" & .iDepth
    # Call getInfo(dInfo, vValue, Assign:=Assign)
    if dInfo.iAssignmentMode = 1 Then:
    # GoTo FuncExit                    ' it was fully decoded
    # dInfo.iTypeName = .iTypeName & " (" & dInfo.iTypeName ' leaves open bracket!
    if dInfo.iIsArray Then:
    # dInfo.iTypeName = dInfo.iTypeName & " Array)" ' close bracket (1)
    # GoTo nonValue
    if dInfo.iAssignmentMode = 0 Then:
    # dInfo.iTypeName = dInfo.iTypeName & " No String Rep.)" ' close bracket (2)
    # GoTo RuleDetermined

    # dInfo.iTypeName = .iTypeName & b & dInfo.iValue.Name & " Value="
    # With dInfo
    # .iClass = .iValue.Class
    # aBugTxt = "get string value of .iValue"
    # vValue = CStr(.iValue)           ' may be empty
    if Catch Then:
    # .DecodeMessage = E_Active.Reasoning
    # .iTypeName = .iTypeName & "None)" ' close bracket (3)
    # .DecodedStringValue = "#* " & .DecodeMessage
    # GoTo didSet                  ' did not work
    else:
    # .iAssignmentMode = 1         ' worked, change from 2->1
    # .iTypeName = .iTypeName & vValue & ")" ' close bracket (4)
    # .DecodeMessage = .iTypeName
    # .DecodedStringValue = vValue
    # GoTo FuncExit
    # End With                             ' dInfo
        case vbEmpty                             '  0 Empty (not initialized):
        case vbNull                              '  1 Null  (no valid data, Nothing):
        case vbError                             ' 10 Fehlerwert:
        case vbVariant                           ' 12 Variant (only arrays of Variant values):
    # DoVerify False, "arrays of variants should have been covered by getInfo ???"
    # .iClass = aValue.Value.Class
        case vbDataObject                        ' 13 Ein Datenzugriffsobjekt:
    # .iClass = aValue.Value.Class
        case _:
    if .iIsArray Then                    ' array needing value:
    # Set .iValue = aValue
    # GoTo likeObject
    # .iTypeName = "# Unknown type " & CStr(.iType)
    # .DecodeMessage = .iTypeName
    # DoVerify False

    if Not Assign Then:
    # GoTo nonValue

    # aBugTxt = "Assign scalar from " & .iTypeName
    # vValue = aValue.Value                    ' try non/-scalar types
    if Not Catch Then:
    # dInfo.iAssignmentMode = 1
    # dInfo.DecodedStringValue = CStr(vValue)
    # GoTo FuncExit
    # .DecodeMessage = E_Active.Reasoning
    # didSet:                                          ' scalar assign did not work, try last resort
    if dInfo.iType < vbInteger Then          ' Empty or Null: no further attempts:
    # .DecodeMessage = dInfo.iTypeName & " is Empty or has Null value"
    # GoTo nonValue

    # aBugTxt = "Assign object value from variant via Set " & .iTypeName
    # Set vValue = aValue.Value                ' try with variant object
    if Catch Then:
    # dInfo.DecodeMessage = "not obtained from variant " & .iTypeName
    # DoVerify False, "Problem assigning to Variant .iValue: " _
    # & .iTypeName & b _
    # & .iTypeName & "->" & aValue.Name
    # .DecodedStringValue = "# unable to obtain value for " & aValue.Name
    # .DecodeMessage = .DecodedStringValue
    # nonValue:
    # DoVerify .iAssignmentMode <> 1, "only for analysis in design ???"
    # Call AppendTo(testNonValueProperties, aTD.adName, sep:=b)
    # RuleDetermined:
    if DebugMode Then:
    print(Debug.Print ">>>> Non-scalar Value for Property " _)
    # & Quote(aTD.adName)
    else:
    # .iAssignmentMode = 1
    # .DecodeMessage = .iTypeName
    # End With                                     ' fInfo

    # FuncExit:
    # Set fInfo = dInfo                            ' this replaces fInfo with bottom Element!
    if fInfo.iAssignmentMode <> 1 And LenB(fInfo.DecodedStringValue) > 0 Then:
    if Left(fInfo.DecodedStringValue, 1) <> "#" Then:
    if DebugLogging And Left(fInfo.DecodeMessage, 1) <> "#" Then:
    # DoVerify False, "Complex value without Explanation"
    else:
    if DebugLogging Then:
    print(Debug.Print fInfo.DecodeMessage & " value=" & fInfo.DecodedStringValue)
    # fInfo.iAssignmentMode = 1        ' it is not necessary to DrillDown again
    # Set dInfo = Nothing
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr, fInfo.DecodeMessage)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : GetPropertyByNumber
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: for the item specified, get Prop = Itemproperties(trueindex)
# '---------------------------------------------------------------------------------------
def getpropertybynumber():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "HelpersOL.GetPropertyByNumber"
    # Call DoCall(zKey, "Function", eQzMode)
    # ' *** bug escape
    # ' not working: Set aProp = anyItem.ItemProperties(trueindex)
    # ' direct assignment would sometimes cause Error 450, so we do it step by step:
    # Dim temp As ItemProperties

    # Set temp = anyItem.ItemProperties
    if DoVerify(temp.Count >= trueindex, "the item has no Itemproperty at trueindex=" & trueindex) Then:
    # Set aProps = Nothing
    else:
    if Not aProps Is temp Then:
    # Set aProps = temp
    # Set GetPropertyByNumber = temp(trueindex)
    # Set temp = Nothing
    # ' *** End bug escape

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function isNotDecodable
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isnotdecodable():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "HelpersOL.isNotDecodable"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim aRules As cAllNameRules

    # isNotDecodable = InStr(testNonValueProperties & b, AttrName & b) > 0
    if isNotDecodable Then:
    # GoTo zExit

    if iRules Is Nothing Then:
    if aOD(aPindex) Is Nothing Then:
    # Set aRules = aOD(0).objClsRules
    else:
    # Set aRules = aOD(aPindex).objClsRules
    else:
    # Set aRules = iRules

    # isNotDecodable = InStr(aRules.clsNotDecodable.aRuleString, AttrName) > 0

    # FuncExit:
    # Set aRules = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : MakeNotLogged
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Build a search folder for NotLogged Items (in active/specified ScopeFolders)
# '---------------------------------------------------------------------------------------
def makenotlogged():

    # Const zKey As String = "TestMailAggregation.AddNotInternalSearchFolder"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim i As Long
    # Dim Filter As String
    # Dim oFolder As Folder
    # Dim ScopeFoldersA As Variant

    # Filter = "NOT Categories LIKE " & Quote1(LOGGED)
    if LenB(ScopeFolders) = 0 Then:
    if LenB(Account) = 0 Then:
    # Set oFolder = ActiveExplorer.CurrentFolder
    # ScopeFolders = "'" & oFolder.FolderPath & "'"
    # Account = Trunc(3, oFolder.FolderPath, "\")
    else:
    # ScopeFolders = "'\\" & Account & "\" & StdInboxFolder & "'"
    if TestEx Then:
    # oFolder = GetFolderByName(ScopeFolders) ' fails if not exists
    elif TestEx Then:
    # ScopeFoldersA = split(Replace(ScopeFolders, "'", vbNullString), ",")
    # ScopeFolders = vbNullString
    # oFolder = GetFolderByName(CStr(ScopeFoldersA(i)))
    # ScopeFolders = ScopeFolders & "'" & oFolder.FolderPath & "',"
    # ScopeFolders = Left(ScopeFolders, Len(ScopeFolders) - 1)

    # Call CreateSearchFolder(Account, ScopeFolders, Filter, _
    # SpecialSearchFolderName & b & NLoggedName, WithSubFolders:=True)

    # Set oFolder = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ModItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def moditem():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.ModItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim oldValue As Variant
    # Dim OtherValue As Variant
    # Dim i As Long
    # Dim N As Long
    # Dim iOther As Long
    # Dim nOther As Long
    # Dim nproperties As ItemProperties
    # Dim Removed_Comma As Boolean

    # oldValue = vbNullString
    # Set nproperties = aktItem.ItemProperties
    # aPindex = 1                                  ' find in aID(1).idAttrDict
    # i = FindAttributeByName(1, adName)
    # N = FindAttributeByName(1, nAttrName)
    # iOther = FindAttributeByName(1, adName)
    # nOther = FindAttributeByName(1, nAttrName)
    if i > 0 And N > 0 And nOther > 0 And iOther > 0 Then ' property exists:
    # ' Fixing iOther = aID(1).idAttrDict.item(i).Index = iOther
    if iOther <> aID(1).idAttrDict.Item(i).DictIndex Then:
    # aID(1).idAttrDict.Item(i).DictIndex = iOther
    if aID(1).idAttrDict.Item(N).DictIndex <> nOther Then:
    # Call ShowAttrs(nproperties, 1)
    if aID(1).idAttrDict.Item(i).adName <> nproperties.Item(iOther).Name Then:
    # Call ShowAttrs(nproperties, 1)
    if aID(1).idAttrDict.Item(N).adName <> nproperties.Item(nOther).Name Then:
    # Call ShowAttrs(nproperties, 1)
    # ' swap lastname and firstname, omitting ","
    # oldValue = nproperties.Item(iOther).Value ' current firstname here
    # OtherValue = nproperties.Item(nOther).Value ' current lastname
    if Right(oldValue, 1) = Right(processingOptions, 1) Then ' has a comma ending:
    # oldValue = Left(oldValue, Len(oldValue) - 1)
    # Removed_Comma = True
    if Right(OtherValue, 1) = Right(processingOptions, 1) Then ' has a comma ending:
    # OtherValue = Left(OtherValue, Len(OtherValue) - 1)
    # Removed_Comma = True
    if Removed_Comma Then:
    # nproperties.Item(iOther) = oldValue
    # nproperties.Item(i) = OtherValue

    # nproperties.Item("FullName") = OtherValue & b & oldValue

    # aktItem.Save
    else:
    # DoVerify False, " Prop not found by name"

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function SearchFolderExists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def searchfolderexists():

    # Const zKey As String = "HelpersOL.SearchFolderExists"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim oStore As Outlook.Store

    # Set oStore = olApp.Session.Stores.Item(StoreName)
    # Set oSearchFolders = oStore.GetSearchFolders

    for sfolder in osearchfolders:
    if sFolder.Name = FolderName Then:
    # SearchFolderExists = True
    # GoTo FuncExit

    # Set sFolder = Nothing                        ' no match

    # FuncExit:
    # Set oStore = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetDebugMode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setdebugmode():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
    # Const zKey As String = "HelpersOL.SetDebugMode"
    # Dim zErr As cErr

    # Dim WithLogDefault As Variant
    # Dim WithDebugDefault As Variant
    # Dim DebugString As String

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    if IsMissing(WithLog) Then:
    # WithLog = False
    # WithLogDefault = True
    if IsMissing(WithDebug) Then:
    # ' debugmode toggle
    # WithDebugDefault = True

    if WithDebugDefault Then:
    if DebugMode Then            ' toggle all to off:
    # Call SetDebugModes("off", noMsg:=noMsg)
    # Call SetErrLogging(LogAllErrors)
    else:
    # Call AppendTo(DebugString, "Debug", b)
    elif WithDebug Then:
    # Call AppendTo(DebugString, "Debug", b)

    if WithLog = "Ask" Then:
    if aNonModalForm Is Nothing Then:
    # ErrStatusFormUsable = True
    # Call N_Suppress(Push, zKey)
    # Call ShowDbgStatus("Choose debug options")
    # Call N_Suppress(Pop, zKey)
    elif WithLog Then:
    # Call AppendTo(DebugString, "Log", b)

    # Call SetDebugModes(DebugString, noMsg:=noMsg)
    # ' interactive setting of Debug options
    if ErrStatusFormUsable Then:
    # With aNonModalForm
    if LenB(Testvar) = 0 Then:
    # .ToggleDebug.Caption = "currently all off"
    else:
    # .ToggleDebug.Caption = Testvar
    if ErrorCaught <> 0 Then:
    # DoVerify DebugMode
    # GoTo pExit
    # .Show
    if DebugMode Then:
    # .ToggleDebug.BackColor = &H80FFFF
    else:
    # .ToggleDebug.BackColor = &H8000000F
    # End With                     ' aNonModalForm

    # pExit:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetDebugModes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setdebugmodes():

    # Const zKey As String = "HelpersOL.SetDebugModes"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim msg As String

    if volltest Then:
    # more = "Debug Log"
    elif InStr(LCase(more), "off") > 0 Then:
    # more = vbNullString

    # Call SetEnvironmentVariable("Test", Trim(more))
    # Call getDebugMode(True)
    if DebugMode Then:
    # msg = msg & vbCrLf & "debugmode ist AN"
    # MinimalLogging = 1
    else:
    # msg = msg & vbCrLf & "debugmode ist AUS"
    # MinimalLogging = 3
    if DebugLogging Then:
    # msg = msg & vbCrLf & "debuglogging ist AN"
    # MinimalLogging = 1
    else:
    # msg = msg & vbCrLf & "debuglogging ist AUS"
    if Not noMsg Then:
    print(Debug.Print msg)
    print(Debug.Print "Testvar=" & Quote(Testvar))

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetErrLogging
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def seterrlogging():

    # Const zKey As String = "HelpersOL.SetErrLogging"
    # Call DoCall(zKey, tSub, eQzMode)

    # Call getDebugMode
    # LogAllErrors = eLall
    # Testvar = Trim(Remove(Testvar, "ERR"))
    if eLall Then:
    # Testvar = "ERR " & Testvar
    # Call SetEnvironmentVariable("Test", Trim(Testvar))
    print(Debug.Print "Testvar=" & Quote(Testvar))

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetLogMode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setlogmode():

    # Const zKey As String = "HelpersOL.SetLogMode"
    # Call DoCall(zKey, tSub, eQzMode)

    if DebugLogging Then:
    # Call SetDebugModes("noLog", False, noMsg:=False)
    else:
    # Call SetDebugModes("Log", False, noMsg:=False)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetTrapMode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def settrapmode():

    # Const zKey As String = "HelpersOL.SetTrapMode"
    # Call DoCall(zKey, tSub, eQzMode)

    # Call getDebugMode
    if onoff And InStr(1, Testvar, "Trap", vbTextCompare) > 0 Then:
    # onoff = False
    # Testvar = Trim(Remove(Testvar, "Trap"))
    if onoff Then:
    # Testvar = "Trap " / Testvar
    # Call SetEnvironmentVariable("Test", Trim(Testvar))
    print(Debug.Print "Testvar=" & Quote(Testvar))

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowAttrs
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showattrs():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.ShowAttrs"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim hi As Long
    # DoVerify aID(StackIndex).idAttrDict.Count <= nproperties.Count, "more decoded properties than itemproperties"
    # hi = aID(StackIndex).idAttrDict.Item(i).DictIndex
    print(Debug.Print i, aID(StackIndex).idAttrDict.Item(i).adName = _)
    # nproperties.Item(hi).Name, _
    # aID(StackIndex).idAttrDict.Item(i).adName, _
    # hi, nproperties.Item(hi).Name

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowCats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showcats():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.ShowCats"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim targetCat As category
    # Dim MyStore As Store
    # Dim msg As String

    match MyStoreOrFolder.Class:
        case olFolder:
    # Set MyStore = MyStoreOrFolder.Store
    # msg = "Dumping Categories for Store " & Quote(MyStore.DisplayName) _
    # & " corresponding to folder " & Quote(MyStoreOrFolder.FolderPath)
        case olStore:
    # Set MyStore = MyStoreOrFolder
    # msg = "Dumping Categories for Store" & Quote(MyStoreOrFolder.DisplayName)
        case _:
    # DoVerify False, " does not make sense"
    if MyStore.Categories.Count = 0 Then:
    # msg = "There are no Categories in Store" & Quote(MyStoreOrFolder.DisplayName)
    if DebugMode Then:
    print(Debug.Print "=======================================================" & vbCrLf _)
    # & msg
    # Set targetCat = MyStore.Categories.Item(i)
    print(Debug.Print i & vbTab & "Category " & Quote(targetCat.Name) _)
    # & vbTab & " ShortcutKey " & targetCat.ShortcutKey _
    # & vbTab & " Color " & targetCat.Color

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub showDebugSettings
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showdebugsettings():

    # Const zKey As String = "HelpersOL.showDebugSettings"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim Test As String
    # Test = GetEnvironmentVar("Test")
    if Test <> Testvar Then:
    # Test = Testvar & "?" & Test
    print(Debug.Print "Testvar=" & Quote(Test), DebugMode, DebugLogging)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowSearchAttrs
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showsearchattrs():

    # Dim aSearches As RDOSearches
    # Dim aSearch As RDOSearch
    # Dim aStore As RDOStore
    # Dim StoreName As String
    # Dim i As Long

    # aRDOSession.MAPIOBJECT = olApp.Session.MAPIOBJECT
    for astore in ardosession:
    print(Debug.Print "-------------------------------------")
    # StoreName = LString(aStore.Name, lDbgM)
    if (aStore.StoreKind = skPstAnsi) Or (aStore.StoreKind = skPstUnicode) Then:
    print(Debug.Print StoreName & " - " & aStore.PstPath)
    elif (aStore.StoreKind = skIMAP4) Then:
    print(Debug.Print StoreName & " - " & aStore.OstPath)
    elif (aStore.StoreKind = skPrimaryExchangeMailbox) Or (aStore.StoreKind = skDelegateExchangeMailbox) Or (aStore.StoreKind = skPublicFolders) Then:
    print(Debug.Print StoreName & " - " & aStore.OstPath& & " - " & aStore.StoreAccount.CurrentUser.Name)
    else:
    print(Debug.Print StoreName & " - " & "unknown Store kind=" & aStore.StoreKind)
    # DoVerify False
    # Set aSearches = aStore.Searches
    for asearch in asearches:
    print(Debug.Print "-------------")
    print(Debug.Print aSearch.Name & ": ")
    print(Debug.Print String(10, b) & aSearch.SearchContainers.Item(i).Name)
    print(Debug.Print aSearch.SearchCriteria.AsSQL)

    # '    Set aSearches = aRDOSession.Stores.DefaultStore.Searches


# '---------------------------------------------------------------------------------------
# ' Method : Function StandardTime
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def standardtime():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "HelpersOL.StandardTime"
    # Call DoCall(zKey, "Function", eQzMode)

    # ' must use with string-formatted time, but also in UTC time zone
    if useUTC Then:

    # Dim oPA As Outlook.PropertyAccessor

    # Set oPA = itemProp.Item.PropertyAccessor
    # StandardTime = oPA.LocalTimeToUTC(tD)
    if DebugMode Then:
    print(Debug.Print tD & " -> " & StandardTime & " (UTC)")
    # Set oPA = Nothing
    else:
    # StandardTime = tD

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function StoreCatGet
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def storecatget():
    # Dim zErr As cErr
    # Const zKey As String = "HelpersOL.StoreCatGet"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim thisCat As category
    # Dim i As Long

    # aBugTxt = "Get Categories from Item"
    # Call Try(testAll)                               ' Try anything, autocatch
    # Set StoreCatGet = MyStore.Categories.Item(catName)
    # Catch
    if IsMissing(Position) Then                  ' no position index (out):
    # Set StoreCatGet = MyStore.Categories.Item(catName)
    else:
    # Set thisCat = MyStore.Categories.Item(i)
    if thisCat.Name = catName Then:
    # Position = i
    # Set StoreCatGet = thisCat
    # GoTo FuncExit
    # ' StoreCatGet is nothing when catName not found

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # ProcRet:

def alleemails():
    # Dim objFolder As Folder
    # Dim sFolder As Folder
    # Dim oItem As Object
    # Dim mytItems As Items
    # Dim tFolder As Folder
    # Dim strFolder As String
    # Dim strFilter As String
    # Dim myExplorers As Explorers
    # Dim myPlorer As Explorer
    # Stop ' set nlogged Hotmail
    # Set myExplorers = Application.Explorers
    print(Debug.Print "Count explorers:"; myExplorers.Count)
    # Set sFolder = Application.ActiveExplorer.CurrentFolder
    # Stop  ' switch folder to target = AggregatedInbox
    # Set tFolder = Application.ActiveExplorer.CurrentFolder
    # Set myPlorer = tFolder.GetExplorer(olFolderDisplayNormal)
    # myPlorer.SelectAllItems
    # Set mytItems = tFolder.Items
    # Set oItem = mytItems.GetFirst
    # While Not oItem Is Nothing
    # ' delete target emails
    # Set oItem = mytItems.GetNext
    # Wend
    # Set oItem = sFolder.Items.GetFirst
    # While Not oItem Is Nothing
    # Call CopyItemTo(oItem, tFolder)
    # Set oItem = sFolder.Items.GetNext
    # Wend

    # ' objFolder = myExplorers.Add(objFolder, olFolderDisplayNormal)
    # ' Set myPlorer = myExplorers.Item(2)
    # ' myPlorer.Display
    # ' myPlorer.Activate
    # Stop
    # Set objFolder = _
    # Application.ActiveExplorer.CurrentFolder

    if objFolder.DefaultItemType <> olMailItem Then:
    # Prompt:="Die Aktion kann im aktuellen " & _
    # "Ordner nicht ausgefhrt werden." & _
    # String(2, vbCrLf) & _
    # "Wechseln Sie bitte erst in einen " & _
    # "E-Mail-Ordner.", _
    # Buttons:=vbExclamation, _
    # Title:="Alle E-Mails anzeigen"
    # Exit Sub

    # strFolder = Chr(34) & "Posteingang" & Chr(34)
    # strFilter = "ordnerpfad:(" & strFolder & ")"

    # objFolder.GetExplorer.Search _
    # strFilter, _
    # olSearchScopeAllFolders

    # Set objFolder = Nothing

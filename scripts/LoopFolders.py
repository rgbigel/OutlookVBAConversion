# Converted from LoopFolders.py

# Attribute VB_Name = "LoopFolders"
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub DeferredActionAdd
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def deferredactionadd():
    # Const zKey As String = "LoopFolders.DeferredActionAdd"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="add Email: " & curObj.Subject)

    # Dim AO As cActionObject
    # Dim aSubject As String
    # Dim aPath As String
    # Dim aCat As String
    # Dim LogMsg As String
    # Dim objTimeType As String
    # Dim aTime As String

    if NoChecking Then:
    # LogMsg = " preselected from "
    else:
    if Not IsMailLike(curObj) Then:
    # Call LogEvent("---- Object is not Mail-Like, EntryID=" & curObj.EntryID _
    # & vbCrLf & "   Class=" & curObj.Class & "/" & TypeName(curObj) _
    # & vbCrLf & "   so it will not be processed", eLall)
    # GoTo ProcReturn
    # LogMsg = b & LOGGED & " is not specified as Categories in "

    # Set AO = New cActionObject
    # aBugTxt = "Set up deferred action"
    # Call Try
    # AO.aoObjID = curObj.EntryID
    if Catch Then:
    # GoTo SkipUntilLater
    # AO.ActionID = curAction
    # aPindex = 1

    # Call GetITMClsModel(curObj, aPindex)                    ' changes aID(aPindex)=aItmDsc, uses or makes aOD

    # Call aItmDsc.UpdItmClsDetails(curObj)

    if NoChecking Then                                      ' already checked and filtered:
    # aPath = curObj.Parent.FolderPath
    # aSubject = Quote(curObj.Subject)
    # aCat = curObj.Categories
    # GoTo TestLimit

    # aSubject = "Item without Subject Property: " & curObj.EntryID ' set default for none there
    # aBugTxt = "get Subject for EntryID=" & curObj.EntryID
    # Call Try
    # aSubject = Quote(curObj.Subject)
    # Catch
    # aBugTxt = "get folder path=Parent of " & aSubject ' it may not have Subject, we accept that
    # Call Try
    # aPath = curObj.Parent.FolderPath
    # Catch
    # aCat = vbNullString
    # aBugTxt = "get Categories of " & aSubject
    # Call Try
    # aCat = curObj.Categories
    # Catch

    if CurIterationSwitches.ReProcessDontAsk Or CurIterationSwitches.ReprocessLOGGEDItems Then:
    # GoTo TestLimit
    if InStr(aCat, LOGGED) = 0 Then                         ' filter not always active ?? so filter again:
    # TestLimit:
    if RestrictedItems Is Nothing Then:
    # DeferredLimitExceeded = False
    elif RestrictedItems.Count = 0 Then:
    # DeferredLimitExceeded = False
    else:
    if Deferred.Count >= DeferredLimit - 1 Then:
    # Call LogEvent(LString("* exceeding the maximum number (" _
    # & RString(DeferredLimit, 3) _
    # & ") of Deferred items on stack", OffObj) _
    # & "remaining items in '" & aPath & "'will be done next time")
    # DeferredLimitExceeded = True
    # Deferred.Add AO                                 ' add as deferred action object
    # EventHappened = True

    # DoVerify LenB(aPath) > 0

    # aTime = LString(" no time info", 30) & b
    # objTimeType = aOD(aPindex).objTimeType
    if LenB(objTimeType) > 0 Then:
    if InStr("SentOn Sent LastModificationTime CreationTime", objTimeType) > 0 Then:
    # aTime = LString(b & LString(objTimeType, 8) & b & aID(aPindex).idTimeValue, 30) & b

    # Call LogEvent(LString("added #" & CStr(Deferred.Count) _
    # & LogMsg & Quote(aPath), 60) & aTime _
    # & LString(aOD(aPindex).objTypeName, 15) _
    # & b & aSubject, eLdebug)
    else:
    # Call LogEvent(LString("did not add #" & CStr(Deferred.Count) _
    # & " action for " & LOGGED & " already done in " _
    # & Quote(aPath), OffAdI) & aSubject, eLmin)
    # SkipUntilLater:
    # Set AO = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DefineLocalEnvironment
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def definelocalenvironment():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.DefineLocalEnvironment"

    # '------------------- gated Entry -------------------------------------------------------
    if Not LF_CurActionObj Is Nothing Then          ':
    # GoTo pExit

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Set LF_CurActionObj = New cActionObject
    # Set CurIterationSwitches = New cIterationSwitches
    # Set LF_CurActionObj.IterationSettings = CurIterationSwitches

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindbeginInFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findbegininfolder():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.FindbeginInFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim curFolder As Folder
    # Dim RootFolderName As String

    # inFolder = Replace(inFolder, "/", "\")
    if olApp.ActiveExplorer Is Nothing Then:
    # Set curFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
    else:
    # Set curFolder = olApp.ActiveExplorer.CurrentFolder
    if beginInFolder Is Nothing Then:
    # Set beginInFolder = topFolder
    if beginInFolder.Class = olNamespace Then:
    # Set beginInFolder = topFolder
    if curFolder Is Nothing Then:
    # DoVerify False

    if topFolder Is Nothing Then:
    # Set topFolder = curFolder
    if Not topFolder Is Nothing Then:
    # ' -1 means first is bigger, +1 first is smaller, =0 if same
    if StrComp(inFolder, topFolder.FolderPath, vbTextCompare) = 0 Then ' exact match::
    # Set FindBeginInFolder = topFolder     ' ignoring NoSearchFolders, allowed always
    # GoTo ProcReturn                       ' no need to search
    # RootFolderName = TrimTail(curFolder.FolderPath, "\")
    if InStr(1, topFolder.FolderPath, "Suchordner", vbTextCompare) _:
    # > 0 And noSearchFolders Then
    # ' already is: Set FindbeginInFolder = Nothing
    # Call LogEvent(Quote(RootFolderName & "\" & inFolder) _
    # & " nicht gesucht: topFolder ist Suchordner")
    # GoTo ProcReturn                       ' no need to search, returning "Nothing"

    if Not curFolder Is Nothing Then            ' exact match::
    if StrComp(inFolder, curFolder.FolderPath, vbTextCompare) = 0 Then:
    # Set FindBeginInFolder = curFolder     ' ignoring NoSearchFolders
    # GoTo ProcReturn                       ' no need to search, returning "Nothing"

    if beginInFolder Is Nothing Then  ' we are looking everywhere:
    # ' start with topFolder parent object
    # Set FindBeginInFolder = FindBeginInFolder(inFolder, topFolder.Parent, noSearchFolders)
    else:
    if beginInFolder.Class = olFolder Then:
    # Set FindBeginInFolder = beginInFolder
    else:
    # Set FindBeginInFolder = Nothing

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FolderFromName
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get folder from namespace by name (Case Sensitive)
# '---------------------------------------------------------------------------------------
def folderfromname():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "LoopFolders.FolderFromName"
    # Call DoCall(zKey, "Function", eQzMode)

    # Dim FolderName As String
    # Dim LookInPath As String
    # Dim Rest As String
    # Dim i As Long
    # Dim j As Long
    # Dim beginInFolder As Object

    if Left(FolderPath, 2) = "\\" Then:
    # i = 3                               ' initially, skip "\\"
    # Set beginInFolder = olApp.GetNamespace("MAPI")  ' = aNameSpace
    # LookInPath = "MAPI Namespace"
    else:
    # DoVerify False, "FolderFromName requires fully qualified inFolder path"

    # FolderName = Trunc(i, FolderPath, "\", Rest, vbBinaryCompare, j)
    if InStr(FolderName, "Suchordner") > 0 Then:
    # i = j
    # FolderName = Rest
    # aBugTxt = "FolderName " & Quote(FolderName) & " in " & LookInPath
    # Call Try(allowNew)
    # Set FolderFromName = beginInFolder.Folders(FolderName)
    # Catch
    if FolderFromName Is Nothing Then:
    # GoTo FuncExit

    if j > 0 Then                           ' more parts delimited by \:
    # LookInPath = Left(Rest, j - 1)
    # i = j + 1                           ' start after \ next time
    # Set beginInFolder = FolderFromName
    # LookInPath = beginInFolder.FolderPath
    # GoTo NextLevel

    # FuncExit:
    # Set beginInFolder = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetFolderByName
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getfolderbyname():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.GetFolderByName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim Folder As Folder
    # Dim RootFolderName As String
    # Dim FolderName As String
    # Dim startFolder As Folder
    # Dim workFolderName As String
    # Dim nFolderName As String
    # Dim nRootFolderName As String
    # Dim i As Long

    # Dim myRecursionDepth As Long                             ' *** try to eliminate this static

    if Left(inFolder, 2) = "\\" Then:
    if CaseKnown Then:
    # Set GetFolderByName = FolderFromName(inFolder)
    if Not GetFolderByName Is Nothing Then:
    # GoTo ProcReturn

    if myRecursionDepth = 0 Then:
    # StringMod = False
    # myRecursionDepth = E_AppErr.atRecursionLvl

    # FolderName = inFolder
    if beginInFolder Is Nothing Then:
    if Left(inFolder, 2) = "\\" Then:
    # Set beginInFolder = aNameSpace
    # workFolderName = inFolder
    # i = 3                               ' initially, skip "\\"
    # innerFolder:
    # i = InStr(i, workFolderName, "\")
    if i = 0 Then:
    # nFolderName = workFolderName
    else:
    # nFolderName = Left(workFolderName, i - 1)
    for folder in begininfolder:
    if StrComp(Folder.Name, nFolderName, vbTextCompare) = 0 Then:
    # Set GetFolderByName = Folder
    if StrComp(GetFolderByName.Name, inFolder, vbTextCompare) = 0 Then:
    if i > 0 Then           ' more inner folders:
    # i = i + 1           ' skip to next "\"
    # Set beginInFolder = GetFolderByName
    # GoTo innerFolder
    else:
    # GoTo SearchExit ' with message
    # Set GetFolderByName = Nothing
    # GoTo SearchExit
    else:
    # Set beginInFolder = FindBeginInFolder(inFolder, beginInFolder, noSearchFolders)
    elif MaxDepth = 1 _:
    # And beginInFolder.Class = olNamespace Then ' all Folders in top level only
    # ' do all Folders in NameSpace
    for folder in begininfolder:
    if StrComp(Folder.Name, FolderName, vbTextCompare) = 0 Then:
    # Set GetFolderByName = Folder
    # GoTo SearchExit ' with message
    if beginInFolder Is Nothing Then  ' unable to find start object:
    # GoTo ProcReturn
    # myRecursionDepth = myRecursionDepth + 1
    if myRecursionDepth > MaxDepth And MaxDepth <> 0 Then:
    # GoTo SearchExit
    if beginInFolder.Class = olFolder Then:
    # Set startFolder = beginInFolder
    elif beginInFolder.Class = olNamespace Then ' all Folders from top level down:
    # ' do all Folders in NameSpace
    for folder in begininfolder:
    # Set GetFolderByName = GetFolderByName(inFolder, Folder, noSearchFolders)
    if Not GetFolderByName Is Nothing Then:
    # GoTo SearchExit ' with message
    # GoTo SearchExit
    else:
    # DoVerify False, " impossible?"
    # ' note: startFolder is a Folder now, can never be Nothing
    # ' temporarily remove initial \\ before we search subFolders (added again later)
    # RootFolderName = Replace(startFolder.FullFolderPath, "\\", vbNullString)
    # nFolderName = RTail(RootFolderName, "\", nRootFolderName)
    # nRootFolderName = "\\" & nRootFolderName
    if nRootFolderName = "\\" Then  ' not fully qualified FolderPath:
    # RootFolderName = startFolder.FullFolderPath ' fix this
    if RootFolderName Like FolderName Then:
    # workFolderName = RootFolderName
    else:
    if InStr(FolderName, "\\") = 1 Then:
    # workFolderName = FolderName
    else:
    # ' compose a fully qualified name
    # workFolderName = RootFolderName & "\" & FolderName
    else:
    # workFolderName = RootFolderName & "\" & nFolderName
    # Set Folder = startFolder
    if StrComp(Folder.FullFolderPath, workFolderName, vbTextCompare) = 0 Then:
    # Set GetFolderByName = Folder    ' explict and complete match, use it
    # GoTo SearchExit
    else:
    if noSearchFolders Then:
    if InStr(1, Folder.FolderPath, "Suchordner", vbTextCompare) _:
    # > 0 Then ' search Folders are undesired:
    # FolderName = RTail(Folder.FolderPath, "\", RootFolderName)
    # Call RTail(RootFolderName, "\", RootFolderName) ' remove SearchFolder part
    # ' compose a fully qualified name
    # workFolderName = RootFolderName & "\" & FolderName
    if startFolder Is Nothing Then:
    for folder in lookupfolders:
    # Set GetFolderByName = getSubFolderByName(workFolderName, Folder)
    if Not GetFolderByName Is Nothing Then:
    # GoTo SearchExit
    else:
    # Set GetFolderByName = getSubFolderByName(workFolderName, startFolder)
    # SearchExit:
    if DebugMode Then:
    if GetFolderByName Is Nothing Then:
    # Call LogEvent("In " & Quote(RootFolderName) _
    # & " den gesuchten Ordner " & Quote(inFolder) _
    # & " nicht gefunden", eLall)
    elif myRecursionDepth <= 2 Then:
    # myRecursionDepth = 2
    # Call LogEvent("Der gesuchte Ordner wurde gefunden: " _
    # & GetFolderByName.FolderPath, eLall)
    # RecursionExit:
    # myRecursionDepth = myRecursionDepth - 1

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function getDefaultFolderType
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getdefaultfoldertype():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.getDefaultFolderType"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # 'On Error GoTo 0
    if TypeName(curObj) = "Variant()" Then:
    # Set ActItemObject = curObj(1)
    elif TypeName(curObj) = "Collection" Then:
    if curObj.Count = 0 Then:
    # GoTo gotNone
    # Set ActItemObject = curObj.Item(1)
    elif curObj.Class = olSelection Then:
    if curObj.Count = 0 Then:
    # GoTo gotNone
    # Set ActItemObject = curObj.Item(1)
    elif curObj.Class = olFolder Then:
    if curObj.Items.Count = 0 Then:
    # GoTo gotNone
    # Set ActItemObject = curObj.Items(1)
    else:
    # gotNone:
    if DebugMode Then DoVerify False, " no item in curobj":
    # Set getDefaultFolderType = Nothing
    # GoTo ProcReturn
    if ActItemObject.Parent.Class = olAppointment Then:
    # Set getDefaultFolderType = ActItemObject.Parent.Parent
    else:
    # Set getDefaultFolderType = ActItemObject.Parent
    # Call BestObjProps(getDefaultFolderType, ActItemObject, withValues:=False) ' seek only if not given

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' InFolder must be a fully qualified FolderPath
def getsubfolderbyname():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.getSubFolderByName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim oFolder As Folder
    # Dim i As Long
    # Dim SubFolderName As String
    # Dim RootFolderName As String

    # 'On Error GoTo 0
    # SearchFolderNameResult = vbNullString
    # SubFolderName = RTail(inFolder, "\", RootFolderName)
    if Not startFolder Is Nothing Then:
    if InStr(1, QuoteWithDoubleQ(startFolder, "\"), _:
    # QuoteWithDoubleQ(SubFolderName, "\"), vbTextCompare) > 0 Then
    # Set getSubFolderByName = startFolder
    # SearchFolderNameResult = startFolder & " Matches InFolder " & inFolder
    # GoTo FuncExit
    if InStr(1, startFolder.FolderPath, inFolder, vbTextCompare) Then:
    # Set getSubFolderByName = startFolder
    # SearchFolderNameResult = startFolder.FolderPath _
    # & " Matches InFolder " & inFolder
    # GoTo FuncExit
    if InStr(1, inFolder, "Suchordner", vbTextCompare) > 0 Then:
    # isSearchFolder:
    # i = 0
    for astore in ostores:
    # Set oSearchFolders = aStore.GetSearchFolders
    for ofolder in osearchfolders:
    # i = i + 1
    # ' find first SearchFolder if 0, else the one after Sfstart
    if (Sfstart = 0 Or i > Sfstart) _:
    # And InStr(1, oFolder.FolderPath, inFolder, vbTextCompare) > 0 Then
    # Set getSubFolderByName = oFolder
    # Sfstart = i ' occurrence in matching search Folders
    # SearchFolderNameResult = "Store Search Folder " _
    # & oFolder.FolderPath & " Matches InFolder " & inFolder
    # GoTo FuncExit
    # Set getSubFolderByName = Nothing
    # GoTo FuncExit

    if startFolder Is Nothing Then:
    # DoVerify False, " ???"
    # ' could we be looking at a search Folder
    if startFolder.Parent Is Nothing Then   ' multiFolder-table:
    # Set getSubFolderByName = startFolder
    else:
    # Set oFolder = DeferredFolder(i)
    if oFolder Is Nothing Then:
    # Exit For
    else:
    if oFolder.FolderPath = startFolder.FolderPath Then:
    if oFolder.Items.Count > 0 Then:
    # Set topFolder = oFolder.Items.Item(1).Parent
    else:
    # DoVerify False, " what else can we do..."
    # Exit For
    # Set getSubFolderByName = topFolder
    else:
    if InStr(1, startFolder.FolderPath, "Suchordner", vbTextCompare) > 0 Then:
    # GoTo isSearchFolder ' this can never have subFolders, Folders property invalid
    else:
    # Set oSearchFolders = startFolder.Folders
    for ofolder in osearchfolders:
    if InStr(1, oFolder.FolderPath, inFolder, vbTextCompare) > 0 Then:
    # Set getSubFolderByName = oFolder
    # SearchFolderNameResult = "Search Folder " _
    # & oFolder.FolderPath & " Matches InFolder " & inFolder
    # GoTo FuncExit
    # ' not found yet: try recursion
    for ofolder in osearchfolders:
    # ' not on original level: recurse subFolders
    if oFolder.Folders.Count > 0 Then:
    # Set getSubFolderByName = getSubFolderByName(inFolder, oFolder)
    if Not getSubFolderByName Is Nothing Then:
    # SearchFolderNameResult = "Located " & oFolder.FolderPath _
    # & " walking SubFolders, InFolder " & inFolder
    # GoTo FuncExit

    # FuncExit:
    if LenB(SearchFolderNameResult) = 0 Then:
    # SearchFolderNameResult = "Found no suitable Folder for " & inFolder
    if DebugMode Then:
    print(Debug.Print SearchFolderNameResult)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# ' Process (all or selected) items in a Folder
def itemactions():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.ItemActions"
    # Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    # Dim itemNo As Long
    # Dim NoOfItems As Long

    # ' Process Folder ITEMS
    # ' ====================
    # NoOfItems = CountItemsIn(curObj, LF_ItemCount)
    if NoOfItems = 0 And Not curObj Is Nothing Then:
    # Call LogEvent("         No items left in " & curObj, eLmin)
    # Call ErrReset(0)
    # DateSkipCount = 0
    # NoOfItems = CountItemsIn(curObj, LF_ItemCount)
    if NoOfItems < itemNo Then  ' itemNo has to be corrected:
    # itemNo = NoOfItems      ' some items of Folder have gone (e.g. Processed, or done by rule)
    if NoOfItems < 1 Then:
    # Exit For

    # Call ShowStatusUpdate
    if DebugLogging Or LF_ItemCount Mod 100 = 0 Then:
    if eOnlySelectedItems Then:
    print(Debug.Print "Now processing item no. " & itemNo _)
    # & " of " & NoOfItems & " selected items"
    else:
    print(Debug.Print "Now processing item no. " & itemNo _)
    # & " of " & NoOfItems & " in Folder No. " _
    # & LF_DoneFldrCount + 1 _
    # & b & Quote(curObj.FolderPath)
    # Call ShowStatusUpdate
    # LF_ItemCount = LF_ItemCount + 1
    # Restart:
    if eOnlySelectedItems Then:
    # aResult = FolderActions(curObj, -itemNo)
    if itemNo <= curObj.Count Then:
    # curObj.Remove itemNo
    # itemNo = itemNo - 1
    else:
    # aResult = FolderActions(curObj, itemNo)
    if ActionID = atDoppelteItemslschen Then:
    # itemNo = NoOfItems ' this stops outer loop, we process all selected directly in action
    match aResult:
        case vbCancel:
    if TerminateRun Then:
    # GoTo ProcReturn
    # GoTo ProcReturn
        case 0      ' user just wanted NOTHING:
    # GoTo FuncExit
        case vbOK:
        case vbNo:
        case vbIgnore:
        case vbCancel:
    # GoTo ProcReturn
        case _:
    # DoVerify False
    if DeferredLimitExceeded Then:
    # Exit For

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub LoopToDoItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def looptodoitems():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "LoopFolders.LoopToDoItems"
    # Static zErr As New cErr

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & zKey & "Ignored, recursion from " _)
    # & P_Active.DbgId
    # GoTo ProcRet
    # Recursive = True                        ' restored by    Recursive = False ProcRet:

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="LoopToDoItems")

    # Call FldActions2Do                      ' (must) have (at least 1) open items

    # ProcReturn:
    # Call ProcExit(zErr)
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub LoopFoldersDialog
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def loopfoldersdialog():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "LoopFolders.LoopFoldersDialog"
    # Static zErr As New cErr
    # Dim ReShowFrmErrStatus As Boolean

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="LoopFolders")

    # Dim OldStatusOfNowNotLater As Boolean

    # IsEntryPoint = True
    # E_Active.EventBlock = False                            ' overriding application NoEvent Stop
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call BugEval

    # ' Statistics inits
    # LF_DontAskAgain = False
    # LF_ItemCount = 0
    # LF_DoneFldrCount = 0
    # LF_ItmChgCount = 0
    # ' State and Action Presets
    # Call DefineLocalEnvironment
    # CurIterationSwitches.ReprocessLOGGEDItems = False
    # CurIterationSwitches.CategoryConfirmation = False

    # AllPublic.eActFolderChoice = True

    # ' Processing Mode Now/Later
    if Not NoEventOnAddItem Then 'save for later:
    # ' maybe delayed processing present, now to be done?
    # OldStatusOfNowNotLater = NoEventOnAddItem
    # NoEventOnAddItem = True     ' let's not be interrupted
    if Not StopRecursionNonLogged Then:
    # Call DoDeferred

    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
    # ReShowFrmErrStatus = True
    # Set FRM = New frmFolderLoop
    # Call ShowOrHideForm(FRM, ShowIt:=True)
    # ActionID = LF_UsrRqAtionId
    # aPindex = 1
    # Call LoopRecursiveFolders(AllPublic.eActFolderChoice)

    # ' Processing Mode "Later"
    # Call DoDeferred
    # & LF_ItmChgCount)
    # Call LogEvent("==== Total number of items modified: " _
    # & LF_ItmChgCount, eLmin)
    # NoEventOnAddItem = OldStatusOfNowNotLater
    if TerminateRun Then:
    # GoTo FuncExit

    # FuncExit:
    # Set FRM = Nothing
    if ReShowFrmErrStatus Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub LoopRecursiveFolders
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def looprecursivefolders():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.LoopRecursiveFolders"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # AskEveryFolder = True
    # WantConfirmation = True

    if LF_UsrRqAtionId = atFindealleDeferredSuchordner Then    ' all-Folder Operations:
    # eOnlySelectedItems = False
    # eActFolderChoice = False
    # ChooseTopFolder = False
    # Call LogEvent("Checking all Search Folders for " _
    # & SpecialSearchFolderName)
    # Call Recurse(topFolder)
    # GoTo RecursionExit

    if LF_UsrRqAtionId = atOrdnerinhalteZusammenfhren Then    ' 2-Folder Operation:
    # AskEveryFolder = True
    # WantConfirmation = True
    # Call Recurse(Nothing)
    # Call AddItemDataToOlderFolder
    elif eAllFoldersOfType Then:
    if eOnlySelectedItems Then ' selection done, process all items:
    if SelectedItems.Count = 0 Then:
    # Call MakeSelection(1, _
    # "bitte selektierien Sie die gewnschten Objekte im mageblichen Ordner ", _
    # "Auswahl selektierter Objekte", "OK", "Cancel")
    else:
    # Call PickAFolder(1, "bitte besttigen oder whlen Sie den mageblichen Ordner ", _
    # "Auswahl des Ordners zu dem gleichartige bestehen", "OK", "Cancel")
    # Call getFoldersOfType(Folder(1))
    # ' all relevant items are collected from all Folders we found
    # Set SelectedItems = New Collection
    # Set topFolder = lFolders(i)
    # Call Recurse(topFolder)
    # eOnlySelectedItems = True    ' now work on them
    # Call Recurse(SelectedItems)
    # ' free space
    # Erase lFolders
    elif ChooseTopFolder And Not eOnlySelectedItems Then:
    # eOnlySelectedItems = False
    # Set SelectedItems = New Collection
    # Set topFolder = Nothing
    # Call PickAFolder(1, "bitte besttigen oder whlen Sie den obersten Ordner ", _
    # "Auswahl des Ordners auf der obersten Ebene", "OK", "Cancel")
    # Set topFolder = Folder(1)
    # LF_DoneFldrCount = 1
    # curFolderPath = topFolder.FolderPath
    # ' compute topmost Folder above the current one:
    if Left(curFolderPath, 2) = "\\" Then:
    # FullFolderPath(FolderPathLevel) = "\\" & Trunc(3, curFolderPath, "\")
    else:
    # FullFolderPath(FolderPathLevel) = curFolderPath
    # Call LogEvent("==== User has selected Folder " & curFolderPath, eLall)
    # Call Recurse(topFolder)
    else:
    if eOnlySelectedItems Then:
    # eOnlySelectedItems = True
    if ActiveExplorer.Selection.Count < 2 Then:
    # Call MakeSelection(1, _
    # "bitte selektierien Sie die gewnschten Objekte im mageblichen Ordner ", _
    # "Auswahl selektierter Objekte", "OK", "Cancel")
    else:
    # Call LogEvent("==== Es wurden " & ActiveExplorer.Selection.Count _
    # & " Objekte vorselektiert.", eLall)
    # Set SelectedItems = New Collection
    # Call GetSelectedItems(olApp.ActiveExplorer.Selection)
    # Set LF_CurLoopFld = Folder(1)
    if BeforeFolderActions() = vbNo Then:
    # GoTo ProcReturn    ' recursion makes no sense
    # Call Recurse(SelectedItems)
    # GoTo RecursionExit
    else:
    # Call LogEvent("==== User requested all Folders to be processed.", eLall)
    # Set topFolder = LookupFolders.Item(i)
    # curFolderPath = topFolder.FolderPath
    # FolderPathLevel = 0
    # FullFolderPath(FolderPathLevel) = vbNullString

    # Call LogEvent("==== now recursing " & Quote(topFolder.FullFolderPath))
    # Call Recurse(topFolder)
    if eOnlySelectedItems Then:
    # GoTo RecursionExit  ' do not ProcCall incorrect selection mode
    # RecursionExit:
    if Not xlApp Is Nothing Then:
    # Call EndAllWorkbooks

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function NonLoopFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def nonloopfolder():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.NonLoopFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim ShortName As String
    # ShortName = LCase(Left(FolderName, 4))
    if ShortName = "dele" Then GoTo asNonLoopFolder:
    if ShortName = "gel" Then GoTo asNonLoopFolder:
    if ShortName = "tras" Then GoTo asNonLoopFolder:
    if ShortName = "junk" Then GoTo asNonLoopFolder:
    if ShortName = "spam" Then GoTo asNonLoopFolder:
    if ShortName = "uner" Then GoTo asNonLoopFolder:
    if ShortName = "dupl" Then GoTo asNonLoopFolder:
    # NonLoopFolder = False
    # isNonLoopFolder = False
    # GoTo ProcReturn
    # asNonLoopFolder:
    # NonLoopFolder = True    ' user can revise this
    # isNonLoopFolder = True
    if ActionID > 0 Then:
    # & vbCrLf & "    ist eventuell sinnlos in " & Quote(FolderName) _
    # & vbCrLf & "Empfehlung: Lass den Quatsch   Ja!" _
    # & vbCrLf & "Trotzdem ausfhren:            Nein" _
    # & vbCrLf & "immer ausfhren:    Cancel", vbYesNoCancel)
    if rsp = vbNo Then:
    # NonLoopFolder = False
    # LF_DontAskAgain = False
    # GoTo ProcReturn
    elif rsp = vbYes Then:
    # Call LogEvent("<======> skipping item action " _
    # & Quote(ActionTitle(ActionID)) _
    # & " because it is inFolder " & Quote(FolderName) _
    # & " loop item " & WorkIndex(1) _
    # & " Time: " & Now(), eLmin)
    # LF_DontAskAgain = False
    # GoTo ProcReturn  ' pretend normal Folder, returning false
    elif rsp = vbCancel Then:
    # NonLoopFolder = False
    # LF_DontAskAgain = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Recurse
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def recurse():
    # Dim zErr As cErr
    # Const zKey As String = "LoopFolders.Recurse"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Dim iFolder As Folder
    # Dim subFolderCount As Long
    # Dim recursedFolder As Folder

    if curObj Is Nothing Then:
    # subFolderCount = 0
    # Set LF_CurLoopFld = Nothing
    elif eOnlySelectedItems Then:
    if SelectedItems.Count = 0 Then:
    # GoTo ItemActionsDone    ' finished on selection/Collection
    # Set LF_CurLoopFld = SelectedItems.Item(1).Parent
    if LF_CurLoopFld Is Nothing Then:
    # DoVerify False, " item without Folder as parent"
    if eAllFoldersOfType Then:
    # Set LF_CurLoopFld = getDefaultFolderType(curObj)
    # Set topFolder = LF_CurLoopFld
    # GoTo selOnly
    # Call ItemActions(SelectedItems)
    else:
    if Not curObj.Parent Is Nothing Then    ' Search Folder: no recursion:
    if curObj.Parent.Class = olNamespace Then   ' it is a top Folder:
    # ' top Folders have no DefaultItemType
    print('Rekursion sinnlos fr oberste Ordnerebenen')
    # GoTo ProcReturn
    # Set LF_CurLoopFld = curObj
    if LF_CurLoopFld.Parent Is Nothing Then:
    # subFolderCount = 0  ' no parent => no kid Folders
    else:
    # subFolderCount = LF_CurLoopFld.Folders.Count ' do recurse

    if BeforeFolderActions() = vbNo Then:
    # GoTo ProcReturn    ' recursion makes no sense
    # Set recursedFolder = LF_CurLoopFld  ' remember for recursion exit
    if eAllFoldersOfType Then:
    # GoTo selOnly    ' no recursion in this case
    # ' Prolog for recursion
    # ' ====================
    # LF_DoneFldrCount = LF_DoneFldrCount + 1
    # Set iFolder = recursedFolder.Folders(LF_recursedFldInx)
    if iFolder = topFolder And FolderPathLevel = 0 Then:
    # ' no op because this is the selected Folder, level will always be =1
    else:
    # curFolderPath = FullFolderPath(FolderPathLevel) _
    # & "\" & iFolder.Name
    # Set Folder(1) = iFolder
    if iFolder.DefaultItemType = olContactItem Then:
    if Not iFolder.ShowAsOutlookAB Then:
    # Call LogEvent("<======> skipping Folder " & iFolder.FolderPath & ": not in Addressbook")
    # GoTo noRecurse
    if NonLoopFolder(iFolder.Name) Then ' could contain any item type! ??? *** what about Entw ?:
    # skipThis:
    # Call LogEvent("<======> skipping Folder " & curFolderPath _
    # & " Time: " & Now(), eLmin)
    # GoTo noRecurse

    # ' Entry to recursion
    # ' ====================
    # FolderPathLevel = FolderPathLevel + 1
    # FullFolderPath(FolderPathLevel) = curFolderPath
    # Call Recurse(iFolder)

    # ' Epilog of recursion
    # ' ====================
    # FolderPathLevel = FolderPathLevel - 1
    # curFolderPath = FullFolderPath(FolderPathLevel)
    # noRecurse:
    # Set LF_CurLoopFld = recursedFolder  ' restore from recursion
    # selOnly:
    if eOnlySelectedItems Then ' only selection in one Folder, no recursion:
    if SelectedItems.Count = 0 Then:
    # Call LogEvent("Es wurden nichts (mehr) selektiert, Verarbeitung wird beendet.", _
    # eLall)
    if TerminateRun Then:
    # GoTo ProcReturn
    else:
    # Call LogEvent("     Bearbeitung von " & SelectedItems.Count _
    # & " Selektionen in " & Quote(LF_CurLoopFld.FullFolderPath) _
    # & " . Beginn: " & Now())
    elif subFolderCount > 0 Then:
    # Call LogEvent("=======> All Ordner unterhalb " & curFolderPath)
    # Call LogEvent("         aktuell: " _
    # & Quote(LF_CurLoopFld.FullFolderPath) _
    # & " mit den enthaltenen " & LF_CurLoopFld.Items.Count _
    # & " Items. Zeit: " _
    # & Now())
    else:
    if Not LF_CurLoopFld Is Nothing Then:
    # Call LogEvent("=======> keine Unterordner in " _
    # & Quote(LF_CurLoopFld.FullFolderPath) & ". Zeit: " & Now())
    # skipRecursion:
    if BeforeItemActions() = vbOK Then ' working on LF_CurLoopFld:
    if eOnlySelectedItems Then:
    # Call ItemActions(SelectedItems)
    else:
    # Call ItemActions(curObj)
    # ItemActionsDone:
    # Call PostFolderActions

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:



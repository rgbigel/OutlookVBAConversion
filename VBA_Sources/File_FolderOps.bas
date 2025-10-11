Attribute VB_Name = "File_FolderOps"
Option Explicit

Public OpenFileNames As Dictionary
Public ClosedFileNames As Dictionary
Public FileStates As Dictionary

Public aOpenKey As Long
Public aClosedKey As Long
Public aFileState As String
Public aFileSpec As String
Public MoveMode As Boolean

'---------------------------------------------------------------------------------------
' Method : Function FolderActions
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FolderActions(curObj As Object, itemNo As Long) As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.FolderActions"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ActionOnWhat As String
Dim ActionResult As String
Dim curFolder As Folder
Dim oneItem As Object

    If itemNo < 0 Then                  ' only one item within selection
        If -itemNo > curObj.Count Then
            GoTo FuncExit               ' all has been done
        End If
        Set oneItem = curObj(-itemNo)   ' OlApp.ActiveExplorer.Selection
        
        aBugTxt = "access to parent of object"
        Call Try("%Das Element wurde verschoben")
        Set curFolder = oneItem.Parent
        If Catch Then
            GoTo FuncExit
        End If
        WorkIndex(1) = -itemNo
        Call InitFindModel(oneItem)     ' get Search criteria
    Else
        If itemNo > curObj.Items.Count Then
            itemNo = curObj.Items.Count
        End If
        If itemNo = 0 Then
            FolderActions = vbCancel
            GoTo FuncExit
        Else
            Set oneItem = curObj.Items(itemNo)
        End If
        Set curFolder = curObj
        WorkIndex(1) = itemNo
    End If
    
    If Not (LF_DontAskAgain _
    Or oneItem.Parent Is Nothing) Then
        If NonLoopFolder(oneItem.Parent.Name) Then
            GoTo FuncExit
        End If
    End If
    
    If eOnlySelectedItems Then
        FolderActions = vbOK  ' no filtering checks if already selected
    Else
        FolderActions = ItemDateFilter(oneItem, eLall)
    End If
    If FolderActions = vbOK Or FolderActions = vbIgnore Then
        ActionOnWhat = "   processing item " & Quote(oneItem.Subject)
    Else
        ActionOnWhat = "   skipping item " & Quote(oneItem.Subject)
        GoTo FuncExit
    End If
    Select Case ActionID
    ' collect these items for later
    Case atDefaultAktion:
        If eOnlySelectedItems Then
            oneItem.Display True    ' modal display
        Else    ' must select now
            Call LogEvent("---- " & TypeName(oneItem) & ": " _
                    & oneItem.Subject, eLnothing)
            WorkIndex(1) = SelectedItems.Count + 1
            SelectedItems.Add oneItem, CStr(WorkIndex(1))
            FolderActions = vbOK
        End If
    Case atKategoriederMailbestimmen:
        Call LogEvent("---- " & TypeName(oneItem) & ": " _
                & oneItem.Subject, eLnothing)
        CurIterationSwitches.ResetCategories = True ' that's what we are here for
        If Not CurIterationSwitches.CategoryConfirmation Then   ' confirming changed unless we said "dontAsk"
            CurIterationSwitches.CategoryConfirmation = Not CurIterationSwitches.ReProcessDontAsk
        End If
        ActionResult = DetectCategory(curFolder, oneItem, curFolder.FullFolderPath)
        If CurIterationSwitches.SaveItemRequested And MailModified And LenB(ActionResult) > 0 Then
            LF_ItmChgCount = LF_ItmChgCount + 1
            oneItem.Categories = ActionResult
            Call LogEvent("Mail categories assigned: " _
                & oneItem.Categories, eLnothing)
            oneItem.Save
        Else
            Call LogEvent("Mail categories not changed: " _
                & oneItem.Categories, eLnothing)
        End If
    Case atPostEingangsbearbeitungdurchführen:
        Call DeferredActionAdd(oneItem, curAction:=3)
    Case atDoppelteItemslöschen:
        If eOnlySelectedItems Then
            Call MatchingItems(MatchMode:=0)
        Else
            Call CheckDoublesInFolder(curFolder)
        End If
    Case atNormalrepräsentationerzwingen:
        Call ScanItem(itemNo, oneItem)
    Case atOrdnerinhalteZusammenführen:
    Case atFindealleDeferredSuchordner:
    Case atBearbeiteAlleÜbereinstimmungenzueinerSuche:
        Call CheckItemProcessed(oneItem)
    Case atContactFixer:
        Call ContactFixItem(oneItem)
    Case Else
        MsgBox "Aktion " & ActionID & " nicht definiert"
    End Select

FuncExit:
    Call N_ClearAppErr

    Call ProcExit(zErr)
End Function ' File_FolderOps.FolderActions

'---------------------------------------------------------------------------------------
' Method : Function ActivateDeferredFavorites
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ActivateDeferredFavorites(Optional justGet As Boolean = False) As Outlook.NavigationGroup
Dim zErr As cErr
Const zKey As String = "File_FolderOps.ActivateDeferredFavorites"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim objPane As NavigationPane
Dim mailModule As Outlook.mailModule
Dim FavFolders As Outlook.NavigationFolders
Dim curFolder As Outlook.NavigationFolder
Dim i As Long

    ' Get the NavigationPane object for the
    ' currently displayed Explorer object.
    'On Error GoTo 0
    
    Set objPane = olApp.ActiveExplorer.NavigationPane
    Set mailModule = objPane.Modules.GetNavigationModule(olApp.OlNavigationModuleType.olModuleMail)
    Set ActivateDeferredFavorites = mailModule.NavigationGroups.GetDefaultNavigationGroup(olApp.OlGroupType.olFavoriteFoldersGroup)
    If Not justGet Then
        Set FavFolders = ActivateDeferredFavorites.NavigationFolders
        For i = 1 To FavFolders.Count
            Set curFolder = FavFolders.Item(i)
            If InStr(1, curFolder.DisplayName, _
                        SpecialSearchFolderName, vbTextCompare) > 0 Then
                If DebugMode Then
                    Debug.Print LString("+ " & curFolder.Folder.FolderPath, OffObj) _
                        & "contains" & RString(curFolder.Folder.Items.Count, 8) _
                        & " Deferred Items "
                End If
                FavNoLogCtr = FavNoLogCtr + 1
            Else
                If DebugMode Then
                    Debug.Print LString("- " & curFolder.Folder.FolderPath, OffObj) _
                        & "shows   " & RString(curFolder.Folder.Items.Count, 8) _
                        & " Deferred Items "
                End If
            End If
        Next i
        If FavNoLogCtr < FldCnt Then
            Call LogEvent("Found " & FavNoLogCtr _
                & " regular Folders named " & Quote(SpecialSearchFolderName) _
                & " within " _
                & FldCnt & " Favorite (Navigation) Folders", eLall)
            If FavNoLogCtr <= 0 Then
                rsp = MsgBox("Es sind keine NotLogged Suchordner in den Favoriten!" _
                            & vbCrLf & "  (Jedes Konto sollte die Posteingänge und Sendungen hierauf absuchen)" _
                            & vbCrLf & "Abbrechen des Laufs: OK, Cancel ignoriert dieses Problem", vbOKCancel)
                If rsp = vbOK Then
                    End
                End If
            End If
        End If
    End If  ' Not justGet

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.ActivateDeferredFavorites

' collect all Folders of same type as aFolder on this level => lFolders()
Sub getFoldersOfType(afolder As Folder)
Dim zErr As cErr
Const zKey As String = "File_FolderOps.getFoldersOfType"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim thisitemType As OlObjectClass
Dim aFld As Folder

    If afolder.Parent Is Nothing Then
        Set afolder = aNameSpace.GetDefaultFolder(olFolderInbox)
    End If
    If afolder.Parent.Class = olFolder Then
        Set topFolder = afolder.Parent
        thisitemType = afolder.DefaultItemType
    Else
        DoVerify False, " wattndattn?"
    End If
    ' loop to get number of Folders of same type
    For Each aFld In topFolder.Folders
        If aFld.DefaultItemType = thisitemType Then
            FldCnt = FldCnt + 1
        End If
    Next aFld
    ReDim lFolders(1 To FldCnt)
    FldCnt = 0
    For Each aFld In topFolder.Folders
        If aFld.DefaultItemType = thisitemType Then
            FldCnt = FldCnt + 1
            Set lFolders(FldCnt) = aFld
        End If
    Next aFld

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.getFoldersOfType

'---------------------------------------------------------------------------------------
' Method : Function getParentFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getParentFolder(aItemO As Object) As Folder
Dim zErr As cErr
Const zKey As String = "File_FolderOps.getParentFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim iObj As Object
Dim lAccountDsc As cAccount
Dim DisplayName As String

    ' if error: result is nothing
    Set iObj = aItemO.Parent    ' usually  aItemO = SelectedItems.Item(1)
    If Catch Then
        GoTo FuncExit
    End If
    
    While Not iObj Is Nothing
        If iObj.Class = olFolder Then
            Set getParentFolder = iObj
            If Catch Then
                GoTo FuncExit
            End If
            DisplayName = getParentFolder.Store.DisplayName
            If Not D_AccountDscs.Exists(DisplayName) Then   ' folders/stores not having account are ok (e.g. backup)
                GoTo FuncExit
            End If
            Set lAccountDsc = D_AccountDscs.Item(DisplayName)
            If Catch Then
                GoTo FuncExit
            End If
            ItemInIMAPFolder = lAccountDsc.aAcType = olImap
            Catch
            GoTo FuncExit
        Else
            Set iObj = iObj.Parent
        End If
    Wend

FuncExit:
    Call ErrReset(4)
    Set iObj = Nothing
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.getParentFolder

'---------------------------------------------------------------------------------------
' Method : GetOrMakeNotLogged
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Search Folders for NotLogged Mail Items
'---------------------------------------------------------------------------------------
Sub GetOrMakeNotLogged()

Const zKey As String = "File_FolderOps.GetOrMakeNotLogged"
    Call DoCall(zKey, tSub, eQzMode)

    Call GetAccountSearchFolders(SpecialSearchFolderName, _
                        "NOT Categories LIKE " & Quote1(LOGGED))

zExit:
    Call DoExit(zKey)

End Sub ' File_FolderOps.GetOrMakeNotLogged

'---------------------------------------------------------------------------------------
' Method : GetAccountSearchFolders
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Search (or make required) search folders for all Accounts
'---------------------------------------------------------------------------------------
Sub GetAccountSearchFolders(FolderName As String, Filter As String)
Const zKey As String = "File_FolderOps.GetAccountSearchFolders"
Dim zErr As cErr

Dim aAccount As Account
Dim oStore As Outlook.Store
Dim oFolder As Outlook.Folder

Dim Scope As String
Dim P As Long
    
Dim objSearch As Search

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="File_FolderOps")

    If SessionAccounts Is Nothing Then
        Set SessionAccounts = olApp.Session.Accounts
    End If
    
    For Each aAccount In SessionAccounts
        Set oStore = aAccount.DeliveryStore
        Set oSearchFolders = oStore.GetSearchFolders
        
        For Each oFolder In oSearchFolders
            Scope = oFolder.FolderPath
            P = InStr(3, Scope, "\")
            If P = 0 Then                       ' only in hierarchy
                GoTo NSF
            End If
            Debug.Print Scope, oFolder.Name
            If oFolder.Name = FolderName Then
                Debug.Print "ofolder matches folder name: ", Scope
                GoTo NAC                        ' it exists for this account, nice
            End If
NSF:
        Next    ' aStore
        ' not existing yet: make it
        Scope = "'" & Left(Scope, P) & StdInboxFolder & "',"
        Debug.Print "Making " & Scope, "Filter=" & Quote(Filter)
        Set objSearch = olApp.AdvancedSearch(Scope:=Scope, _
                            Filter:=Filter, _
                            SearchSubFolders:=True, _
                            Tag:=FolderName)
        oFolder.ShowItemCount = olShowTotalItemCount
        Call objSearch.Save(FolderName)
        GoTo FuncExit
NAC:
    Next aAccount
    
FuncExit:
    Set aAccount = Nothing
    Set oStore = Nothing
    Set oFolder = Nothing

ProcReturn:
    Call ProcExit(zErr)

End Sub ' File_FolderOps.GetAccountSearchFolders

'---------------------------------------------------------------------------------------
' Method : Sub GetSearchFolders
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get all / select some Search Folders in all Stores
'---------------------------------------------------------------------------------------
Sub GetSearchFolders()   ' must set UInxDeferred to 0 for ReNew
Const zKey As String = "File_FolderOps.GetSearchFolders"
Dim zErr As cErr

Dim oStore As Outlook.Store
Dim oFolder As Outlook.Folder
Dim retrycount As Long
Dim i As Long

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="File_FolderOps")

    If UInxDeferred <= 0 Then
        UInxDeferred = 0
Retry:
        
        For Each oStore In oStores
            ' like Ungesehen, NotLogged etc. some always exist!
            aBugTxt = "Get Search Folders in " & oStore.DisplayName
            Call Try
            Set oSearchFolders = oStore.GetSearchFolders
            If Not Catch Then
                For Each oFolder In oSearchFolders
                    Call LogEvent("=======> Found search Folder " _
                            & oFolder.FullFolderPath)
                    If oFolder.Name = SpecialSearchFolderName Then
                        Call LogEvent("!!!=> Found SpecialSearchFolderName Folder " _
                            & oFolder.FullFolderPath, eLall)
                        UInxDeferred = UInxDeferred + 1
                        Set DeferredFolder(UInxDeferred) = oFolder
                    End If
                    Catch
                Next oFolder
            End If
noSearchFolders:
        Next oStore
        
        Call LogEvent("Es wurden " & UInxDeferred & " Suchordner vom Typ " _
                 & Quote(SpecialSearchFolderName) & " gefunden.")
        For i = 1 To UInxDeferred
            Set oFolder = DeferredFolder(i)
    
            Debug.Print LString("+ " & oFolder.Folder.FolderPath, OffObj) _
                & "contains" & RString(oFolder.Folder.Items.Count, 8) & " Deferred Items "
        Next i
    End If
    ' currently at least one SpecialSearchFolderName Folder rqd in Backup
    If UInxDeferred > 1 Then
        If UInxDeferred < FldCnt And Not UInxDeferredIsValid Then
            Call ActivateDeferredFavorites
            If FavNoLogCtr > 0 Then
                UInxDeferred = 0
                retrycount = retrycount + 1
                If retrycount < 4 Then
                    GoTo Retry
                Else
                    retrycount = 0
                    If DebugMode Then
                        DoVerify False, " SpecialSearchFolderName Folders not visible or do not exist"
                    End If
                End If
            Else
                Call LogEvent("Es wird nicht mehr nach weiteren " _
                    & Quote(SpecialSearchFolderName) _
                    & " Ordnern gesucht", eLall)
            End If
        End If
        UInxDeferredIsValid = True
    End If

FuncExit:
    Set oStore = Nothing
    Set oFolder = Nothing
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.GetSearchFolders

'---------------------------------------------------------------------------------------
' Method : Function ItemDateFilter
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ItemDateFilter(oneItem As Object, Optional logLvl As eLogLevel = eLSome) As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.ItemDateFilter"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ItemDate As Date
    If CutOffDate = "00:00:00" Then
        ItemDateFilter = vbOK   ' No filter set, pass thru
    Else
        ' IsMailLike aufgesplittet:
        Select Case TypeName(oneItem)
        Case "MailItem", "MeetingItem"
            ItemDate = oneItem.SentOn
            aTimeFilter = "SentOn"
            If ItemDate = BadDate Then
                ItemDate = oneItem.ReceivedTime
                aTimeFilter = "ReceivedTime"
            End If
doFilter:
            If ItemDate >= CutOffDate Then
                ItemDateFilter = vbOK   ' do process this
            Else
                DateSkipCount = DateSkipCount + 1
                ItemDateFilter = vbNo   ' do ignore this
                Call LogEvent(LimitAppended("     ", Quote(oneItem.Subject), 30, "... ") _
                        & " verfehlt Datumsauswahl: " _
                        & CStr(ItemDate) & "<" & CStr(CutOffDate), logLvl)
            End If
        Case "AppointmentItem"
            ItemDate = Format(oneItem.End, "dd.mm.yyyy")
            If ItemDate = BadDate Then
                ItemDateFilter = vbOK   ' no end: keep going
            End If
            GoTo doFilter
        Case Else
            ItemDateFilter = vbIgnore   ' no specific filter defined
        End Select
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.ItemDateFilter

'---------------------------------------------------------------------------------------
' Method : Function PostFolderActions
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function PostFolderActions() As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.PostFolderActions"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ActionOnWhat As String
    If Not LF_CurLoopFld Is Nothing Then
        ActionOnWhat = "   processing Folder " & LF_CurLoopFld
    End If
    PostFolderActions = vbOK
    Select Case ActionID
        Case atDefaultAktion                                ' 1
            If eOnlySelectedItems Then
                Call ItemActions(SelectedItems)
            End If
        Case atKategoriederMailbestimmen                    ' 2
        Case atPostEingangsbearbeitungdurchführen           ' 3
        Case atDoppelteItemslöschen                         ' 4
        Case atNormalrepräsentationerzwingen                ' 5
        Case atOrdnerinhalteZusammenführen                  ' 6
        Case atFindealleDeferredSuchordner                  ' 7
        Case atBearbeiteAlleÜbereinstimmungenzueinerSuche   ' 8
        Case atContactFixer                                 ' 9
    Case Else
        MsgBox "Post-Folder-Action " & ActionID & " undefined"
    End Select

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.PostFolderActions

' this function runs before we ProcCall a Folder level
Function BeforeFolderActions() As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.BeforeFolderActions"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ActionOnWhat As String
    BeforeFolderActions = vbOK
    If Not LF_CurLoopFld Is Nothing Then
        ActionOnWhat = "   processing Folder " & LF_CurLoopFld
        Call BestObjProps(LF_CurLoopFld, withValues:=False)
    End If
    
    Select Case ActionID
        Case atDefaultAktion                                                ' 1
            SelectMulti = False
            If aObjDsc.objItemClass = olAppointmentItem Then
                Call ItemActions(LF_CurLoopFld)
            Else
                MsgBox "Eine Default-Operation ist für den Item-Typ " _
                        & aObjDsc.objItemClassName & " nicht definiert"
                BeforeFolderActions = vbNo
                GoTo ProcReturn
            End If
            Call LogEvent("Bis Ordner " & Quote(LF_CurLoopFld.FolderPath) _
                        & "  wurden " & SelectedItems.Count & " Items gewählt")
        Case atKategoriederMailbestimmen                                    ' 2
            SelectMulti = False
        Case atPostEingangsbearbeitungdurchführen, _
            atNormalrepräsentationerzwingen                                 ' 3, 5
            SelectMulti = True
            If eOnlySelectedItems Then
                Call ItemActions(SelectedItems)
            Else
                Call ItemActions(LF_CurLoopFld)
            End If
        Case atDoppelteItemslöschen                                         ' 4
            SelectMulti = True
            Call Initialize_UI     ' displays options dialogue
            Select Case rsp
            Case vbYes
                If eOnlySelectedItems Then   ' set loop exit for Folders:
                    LF_DoneFldrCount = LookupFolders.Count    ' do not loop Folder list if Selection
                    GoTo likepicked
                ElseIf PickTopFolder Then
                    LF_DoneFldrCount = LookupFolders.Count    ' do not loop Folder list if Folder is picked
                    Call PickAFolder(1, _
                        "bitte bestätigen oder wählen Sie " _
                        & "den obersten Ordner für die Doublettensuche ", _
                        "Auswahl des Hauptordners für die Doublettensuche", _
                        "OK", "Cancel")
                    Set topFolder = Folder(1)
likepicked:
                    Call FindTrashFolder
                    Set ParentFolder = Nothing
                    
                    aBugTxt = "Get Parent folder of " _
                                    & topFolder.FolderPath
                    Call Try
                    Set ParentFolder = topFolder.Parent
                    Catch
                    
                    Set LF_CurLoopFld = topFolder
                    curFolderPath = LF_CurLoopFld.FolderPath
                    FullFolderPath(FolderPathLevel) = "\\" _
                                    & Trunc(3, curFolderPath, "\")
                    If BeforeItemActions() = vbOK Then
                       ' Debug.Assert False
                    End If
                Else    ' loop Folders items , no (single) Folder was picked
                    If BeforeItemActions() = vbOK Then
                        bDefaultButton = "Go"
                    End If
                End If
            Case vbCancel
                Call LogEvent("=======> Stopped before processing any Folders . Time: " _
                    & Now(), eLnothing)
                If TerminateRun Then
                    GoTo ProcReturn
                End If
                GoTo ProcReturn
            Case Else   ' loop Candidates
                If topFolder Is Nothing Then
                    Set topFolder = LookupFolders.Item(LF_DoneFldrCount)
                End If
                Call FindTrashFolder
            End Select ' rsp (response from InitializeUserID)
            If eOnlySelectedItems Then
                BeforeFolderActions = vbNo ' we are done
                GoTo ProcReturn
            End If
        Case atOrdnerinhalteZusammenführen                                  ' 6
            SelectMulti = True
            Call AddItemDataToOlderFolder
        Case atFindealleDeferredSuchordner                               ' 7
            SelectMulti = False
            Call FldActions2Do    ' if we have open items, do em now
            BeforeFolderActions = vbNo
        Case atBearbeiteAlleÜbereinstimmungenzueinerSuche                   ' 8
            SelectMulti = True
            ' Folder by default (no user interaction)
            Set ChosenTargetFolder = GetFolderByName("Erhalten")
            Call FirstPrepare
        Case atContactFixer
            SelectMulti = False
        Case Else
            SelectMulti = False
            MsgBox "Pre-Folder-Action " & ActionID & " undefined"
    End Select ' ActionID

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.BeforeFolderActions

' this function runs before we process items in a single Folder
Function BeforeItemActions() As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.BeforeItemActions"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ActionOnWhat As String

    BeforeItemActions = vbOK
    Select Case ActionID
    Case atDefaultAktion
        BeforeItemActions = SkipTop(LF_CurLoopFld)
    Case atKategoriederMailbestimmen
    Case atPostEingangsbearbeitungdurchführen
        If eOnlySelectedItems Then
            Call ItemActions(SelectedItems)
        Else
            Call ItemActions(LF_CurLoopFld)
        End If
    Case atDoppelteItemslöschen
        If eOnlySelectedItems Then
            Call ItemActions(SelectedItems)
        Else
            Call ItemActions(LF_CurLoopFld)
        End If
        ' Call CheckDoublesInFolder(topFolder)    Main Work in "ItemActions"
    Case atNormalrepräsentationerzwingen
        Call BestObjProps(LF_CurLoopFld, withValues:=False)
    Case atOrdnerinhalteZusammenführen
    Case atFindealleDeferredSuchordner
    Case atContactFixer
    Case Else
        MsgBox "Pre-Item-Action " & ActionID & " undefined"
    End Select
    If LF_CurLoopFld Is Nothing Then
        GoTo ProcReturn
    End If
    ActionOnWhat = "   processing Folder " _
                    & LF_CurLoopFld & " is " & BeforeItemActions

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.BeforeItemActions

'---------------------------------------------------------------------------------------
' Method : Function SetItemCategory
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SetItemCategory(curFolder As Folder, curItem As Object, category As String) As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.SetItemCategory"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ShortName As String
    If IsMailLike(curItem) Then
        If curFolder <> topFolder Then
            ShortName = Left(curFolder.Name, 4)
            If ShortName = "Junk" Or ShortName = "Spam" Or ShortName = "Uner" Then
                category = "Junk"
            Else
                curItem.Categories = category
            End If
            LF_ItmChgCount = LF_ItmChgCount + 1
            curItem.Save
        End If
    End If
    SetItemCategory = vbOK

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.SetItemCategory

'---------------------------------------------------------------------------------------
' Method : Function SkipTop
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SkipTop(curFolder As Folder) As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "File_FolderOps.SkipTop"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If curFolder = topFolder Then
        SkipTop = vbNo
    Else
        SkipTop = vbOK
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.SkipTop

'---------------------------------------------------------------------------------------
' Method : Sub FindTopFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FindTopFolder(afolder As Folder)
Dim zErr As cErr
Const zKey As String = "File_FolderOps.FindTopFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    
    If topFolder Is Nothing Then
        If afolder.Parent Is Nothing Then
            ' we are in a search Folder=
            If afolder.Items.Count > 0 Then ' all items are in same Folder!!
                Call Try
                Set topFolder = afolder.Items(1).Parent
                Catch
            Else
                MsgBox ("In diesem Suchordner sind keine Daten")
                topFolder = Nothing
                GoTo ProcReturn
            End If
        Else
            Set topFolder = afolder
        End If
        
        Call ErrReset(0)
        Do                                  ' loop until we reach the outermost folder
            If topFolder.Parent.Class = olFolder Then
                Set topFolder = topFolder.Parent
            Else
                Exit Do                     ' topfolder.parent is Mapi
            End If
        Loop
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.FindTopFolder

'---------------------------------------------------------------------------------------
' Method : Sub CopyItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyItems()
Dim zErr As cErr
Const zKey As String = "File_FolderOps.CopyItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long, j As Long
Static S_Folders(2) As Folder

Dim action As String

    Set LookupFolders = aNameSpace.Folders
    bDefaultButton = "No"

    action = " und bestätigen Sie die Auswahl mit 'OK'"
ask1again:
    rsp = NonModalMsgBox("bitte wählen Sie den Quellordner" & action, _
                         "OK", "Cancel", "Auswahl des Quell- und Zielordners")
    
    Select Case rsp
    Case vbOK
        Set S_Folders(1) = ActiveExplorer.CurrentFolder
        If Not S_Folders(1) Is Nothing Then
            i = S_Folders(1).Items.Count
            If i < 1 Then
                MsgBox ("Der Quellordner ist leer")
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            Else
                If ActiveExplorer.Selection.Count > 0 Then
                    Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
                Else
                    Set ActiveExplorerItem(1) = S_Folders(1).Items(1)
                End If
            End If
        Else
            action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
            GoTo ask1again
        End If
    Case vbCancel
        Call LogEvent("=======> Beendet ohne Quell-Auswahl . Time: " & Now(), eLmin)
        If TerminateRun Then
            GoTo ProcReturn
        End If
    End Select
    action = " und bestätigen Sie die Auswahl mit 'OK'"
ask2again:
    rsp = NonModalMsgBox("bitte wählen Sie den Zielordner" & action, _
                         "OK", "Cancel", "Auswahl des Quell- und Zielordners")
    
    Select Case rsp
    Case vbOK
        Set S_Folders(2) = ActiveExplorer.CurrentFolder
        If S_Folders(2) Is Nothing Then
            action = " (ein Ziel muss angegeben werden!)"
            GoTo ask2again
        End If
        If S_Folders(1) = S_Folders(2) Then
            DoVerify False, " select S_Folders now and then F5"
            GoTo ask2again
        End If
    Case vbCancel
        Call LogEvent("=======> Beendet ohne Zielauswahl. Time: " & Now(), eLmin)
        If TerminateRun Then
            GoTo ProcReturn
        End If
    End Select
    
    If MoveMode Then
        Call LogEvent("All Items from " & S_Folders(1).FolderPath _
                & " will be moved to " & S_Folders(2).FolderPath, eLmin)
    Else
        Call LogEvent("All Items from " & S_Folders(1).FolderPath _
                & " will be copied to " & S_Folders(2).FolderPath, eLmin)
    End If
    
gothemAll:
    For j = 1 To i
            i = S_Folders(1).Items.Count
            If i < j Then
                GoTo gothemAll
            End If
        On Error GoTo coudNotCopy
        If MoveMode Then
            Set ActItemObject = S_Folders(1).Items.Item(j)
            ActItemObject.Move S_Folders(2)
            If DebugLogging Or DebugMode Then
                Debug.Print "Moved " & ActItemObject.Subject & " to " _
                            & S_Folders(2).FolderPath
            End If
        Else
            Set ActItemObject = S_Folders(1).Items.Item(j).Copy
            ActItemObject.Move S_Folders(2)
            If DebugLogging Or DebugMode Then
                Debug.Print "Copied " & ActItemObject.Subject & " to " _
                            & S_Folders(2).FolderPath
            End If
        End If
        GoTo OK
coudNotCopy:
        Debug.Print j, ActItemObject.Subject, Err.Description
        If Not MoveMode Then
            ActItemObject.Delete        ' sonst wurde doublette erzeugt
        End If
        Call N_ClearAppErr
        Call ErrReset(0)
OK:
    Next j
    MoveMode = False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.CopyItems

'---------------------------------------------------------------------------------------
' Method : Sub CopyNotes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyNotes()
Dim zErr As cErr
Const zKey As String = "File_FolderOps.CopyNotes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim j As Long
Dim bodystring As String
Dim action As String

    Set LookupFolders = aNameSpace.Folders
    bDefaultButton = "No"

    action = " und bestätigen Sie die Auswahl mit 'OK'"
ask1again:
    rsp = NonModalMsgBox("bitte wählen Sie den Quellordner" & action, _
                         "OK", "Cancel", _
                         "Auswahl des Quell- und Zielordners")
    
    Select Case rsp
    Case vbOK
        Set Folder(1) = ActiveExplorer.CurrentFolder
        If Not Folder(1) Is Nothing Then
            i = Folder(1).Items.Count
            If i < 1 Then
                MsgBox ("Der Quellordner ist leer")
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            Else
                If ActiveExplorer.Selection.Count > 0 Then
                    Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
                Else
                    Set ActiveExplorerItem(1) = Folder(1).Items(1)
                End If
            End If
        Else
            action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
            GoTo ask1again
        End If
    Case vbCancel
        Call LogEvent("=======> Beendet ohne Quell-Auswahl . Time: " _
                        & Now(), eLmin)
        End
    End Select
    action = " und bestätigen Sie die Auswahl mit 'OK'"
ask2again:
    rsp = NonModalMsgBox("bitte wählen Sie den Zielordner" & action, _
                         "OK", "Cancel", "Auswahl des Quell- und Zielordners")
    
    Select Case rsp
    Case vbOK
        Set Folder(2) = ActiveExplorer.CurrentFolder
        If Folder(2) Is Nothing Then
            action = " (ein Ziel muss angegeben werden!)"
            GoTo ask2again
        End If
    Case vbCancel
        If DebugMode Or MinimalLogging < 2 Then
            Call LogEvent("=======> Beendet ohne Zielauswahl. Time: " _
                        & Now(), eLmin)
        End If
        If TerminateRun Then
            GoTo ProcReturn
        End If
    End Select
gothemAll:
    For j = 1 To i
        Set ActItemObject = Folder(2).Items.Add(Folder(2).DefaultItemType)
        bodystring = Folder(1).Items.Item(j).Body
        If InStr(bodystring, "Notizen\") = 1 Then
            bodystring = Mid(bodystring, 9)
        End If
        ActItemObject.Body = bodystring
        ActItemObject.Save
    Next j

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.CopyNotes

'---------------------------------------------------------------------------------------
' Method : Sub MoveItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MoveItems()
Dim zErr As cErr
Const zKey As String = "File_FolderOps.MoveItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    MoveMode = True
    Call CopyItems

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' File_FolderOps.MoveItems

'---------------------------------------------------------------------------------------
' Method : GetOrMakeOlFolder
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Get an Outlook Folder by Name; if not found ask if it is to be created
' Result : True if Folder NOT found and not made successfully after request
'---------------------------------------------------------------------------------------
'        allowMissing=0                      ' optional without question
'        allowMissing=1                      ' optional with question
'        allowMissing=Else                   ' mandatory folder existance
Function GetOrMakeOlFolder(FolderName As String, useFolder As Folder, belowFolders As Folders, Optional allowMissing As Integer = -1) As Boolean
Dim zErr As cErr
Const zKey As String = "File_FolderOps.GetOrMakeOlFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    aBugTxt = "Ordner " & FolderName & " nicht gefunden."
    Call Try(-2147221233)
    Set useFolder = belowFolders.Item(FolderName)
    Catch
'   Verweis auf Zielordners des Transports aus der Inbox
    If useFolder Is Nothing Then
        Select Case allowMissing
        Case 1                      ' optional with question
            rsp = MsgBox(E_Active.Explanations & " nicht gefunden. Soll fehlen?", vbYesNo)
            Call ErrReset(4)
            If rsp = vbYes Then
                rsp = vbNo
            Else
                rsp = vbYes
            End If
        Case 0                      ' optional without question
            Call LogEvent(E_Active.Explanations & " Akzeptiert.", eLSome)
            Call ErrReset(4)
            rsp = vbNo
        Case Else                   ' mandatory folder existance
            rsp = MsgBox("Ordner " & FolderName & " nicht gefunden. Anlegen?", vbYesNo)
        End Select ' allowMissing
        
        If rsp = vbYes Then
            aBugTxt = "Make new Folder " & Quote(FolderName)
            Call Try
            Set useFolder = CreateFolderIfNotExists(FolderName, belowFolders.Item(1))
            Catch
        End If
    End If
    GetOrMakeOlFolder = useFolder Is Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' File_FolderOps.GetOrMakeOlFolder

'---------------------------------------------------------------------------------------
' Method : Sub LogEvent
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Output Text to Log depending on log level, optional with MsgBox
'---------------------------------------------------------------------------------------
Sub LogEvent(Text As String, Optional ByVal ifLevelLess As eLogLevel = eLSome, Optional withMsgBox As Boolean = False, Optional fileNo As Long = 1)


'------------------- gated Entry -------------------------------------------------------
Static Recursive As Long

    If Recursive > 1 Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If Recursive > 3 Then
            Call MsgBox(Recursive + 1 & " recursive calls to LogEvent")
            GoTo refuse
        End If
        If StackDebug > 8 Then
            GoTo refuse
        End If
        If AppStartComplete Then
refuse:
            Debug.Print String(OffCal, b) & "Ignored recursion to LogEvent " & vbCrLf & Text
            GoTo ProcRet
        End If
    End If
    Recursive = Recursive + 1

Dim rsp As VbMsgBoxResult
Dim lLevel As eLogLevel
Dim LogThisToFile As Boolean
Dim SaveTry As Variant
Dim msg As String

    SaveTry = E_Active.Permitted
    If fileNo <> 1 Then
        Debug.Print "printing to logfile #" & fileNo & " not implemented"
        GoTo FuncExit
    End If

    With T_DC
        If DateId = vbNullString Then
            Call GetDateId(-1)                  ' just set global vars, value not needed
        Else
            Call GetDateId(-2)                  ' long values remember globallly
        End If
        If DebugMode Then
            lLevel = ifLevelLess - 1            ' make logging more likely
        Else
            lLevel = ifLevelLess
        End If
         
        If withMsgBox And lLevel <= eLSome Then
            rsp = MsgBox(Text, vbOKCancel, "Event logging")
            If rsp = vbCancel Then
                Call TerminateRun(withStop:=True)
            End If
        End If
        
        LogThisToFile = StackDebug = 0 Or DebugMode _
                     Or DebugLogging Or StackDebug > 6 _
                     Or (lLevel <= eLall _
                     And lLevel <= eLnothing - MinimalLogging)
        
        If Not LogThisToFile Then
            If LogImmediate Then
                Debug.Print Text                    ' put into Direct Window
            ElseIf lLevel = 0 Or ((DebugMode Or DebugLogging Or StackDebug > 7)) Then
                Debug.Print Text                    ' put into Direct Window
            End If
            GoTo FuncExit
        End If
        
OpenNextLogFile:
        Err.Clear
        If LenB(.LogNameNext) = 0 Then
            If LenB(.LogName) > 0 Then
                msg = "Old Log: " & .LogName & " closing. "
            ElseIf LenB(.LogNamePrev) > 0 Then
                msg = "Old Log: " & .LogNamePrev & vbCrLf & "new Log: "
            Else
                msg = "No old log, "
            End If
            Call GetDateId(-2)              ' remember globally
            .LogNameNext = lPfad & "Outlook VBA " & DateIdNB & ".log"
            msg = msg & .LogNameNext
        ElseIf InStr(.LogNameNext, Left(DateIdNB, 8)) = 0 Then
            If LenB(.LogName) > 0 Then
                msg = "Old Log: " & .LogName & ", switch to new day: "
            ElseIf LenB(.LogNamePrev) > 0 Then
                msg = "Old Log: " & .LogNamePrev & vbCrLf & "next Log: "
            Else
                msg = "No old log "
            End If
            .LogNameNext = lPfad & "Outlook VBA " & DateIdNB & ".log"
            msg = msg & vbCrLf & "next log: " & .LogNameNext
        End If
            
        If .LogIsOpen Then
            If .LogName <> .LogNameNext Then
                Call CloseLog(msg:=msg)                     ' changes .LogIsOpen := False
                msg = vbNullString                                    ' Closelog did print
                .LogName = .LogNameNext
            End If
        Else
            .LogName = .LogNameNext
        End If
        
        If LenB(msg) > 0 Then
            Debug.Print msg
            msg = vbNullString                                        ' print done
        End If
        If LenB(Text) > 0 Then
            If LogImmediate _
            Or lLevel = 0 _
            Or DebugMode _
            Or DebugLogging _
            Or StackDebug > 7 _
            Then
                Debug.Print Text                    ' put into Direct Window
            End If
        End If
        
        If .LogIsOpen Then
            GoTo Output
        End If
        
        On Error Resume Next
        Open .LogName For Append As #1
        If Err.Number = 0 Then
           GoTo AllOk
        ElseIf Err.Number = 55 Then                     ' already open
            E_Active.FoundBadErrorNr = 0
            GoTo skipErrorTest
        ElseIf Err.Number <> 0 Then
            Debug.Print "Log " & .LogName & " did not open, Error " & Err.msg
            Debug.Assert False
            .LogNameNext = vbNullString
            GoTo OpenNextLogFile
skipErrorTest:
            Debug.Print vbCrLf & String(OffCal, b) & "Logfile reused: " & .LogName _
                        & vbCrLf
            .LogIsOpen = True
        Else
AllOk:
            .LogIsOpen = True
            If .LogNamePrev <> .LogName Then
                If LenB(.LogNamePrev) > 0 Then
                    msg = "Previous Log Name: " & .LogNamePrev
                    Print #1, msg
                End If
            Else
                Debug.Print String(OffCal, b) & "Logfile re-opened: " & .LogName
            End If
        End If
        
Output:
        If .LogIsOpen Then
            On Error Resume Next                ' often getting (52)
            Print #1, Text
            If Err.Number = 0 Then
                If DebugLogging And Not DebugMode Then
                    Debug.Print String(OffCal, b) & "Log appended:      " & .LogName
                End If
                GoTo LogFileOK
            End If
        Else
            GoTo OpenNextLogFile
        End If
        
        If .DCerrNum = 52 Then
            Debug.Print "LogFile Error: " & Text
        End If
        If Catch Then
            .LogIsOpen = False
            .LogName = vbNullString
            .LogNameNext = vbNullString
            GoTo OpenNextLogFile
        End If
        
LogFileOK:
        .LogFileLen = .LogFileLen + Len(Text)
        If LimitLog > 0 Then
            If .LogFileLen Mod 100 * LimitLog = 0 Then
                Debug.Print "§§§ Log limit reached ~: " & .LogFileLen / LimitLog & " lines. Press F5 to continue"
                ' Debug.Assert False     #### Switch this on?
            End If
        End If
        ' Check max file size
        If .LogFileLen >= MaxCharsPerLogFile Then
            .LogNamePrev = .LogName                       ' force a new file name
            Call CloseLog(msg:="Start new Log, Length exceeds " & MaxCharsPerLogFile & "<" & .LogFileLen)
            GoTo OpenNextLogFile
        End If
        
FuncExit:
        If AppStartComplete Then
            Call ShowStatusUpdate
        End If
        E_Active.Permit = SaveTry                        ' must not change
    End With ' T_DC
    
    Recursive = Recursive - 1

ProcRet:
End Sub ' File_FolderOps.LogEvent

'---------------------------------------------------------------------------------------
' Method : Sub CloseLog
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CloseLog(Optional KeepName As Boolean, Optional msg As String)
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean
Const zKey As String = "File_FolderOps.CloseLog"
Dim zErr As cErr

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                          & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True

    With T_DC
        If .LogIsOpen Then                              ' implies .logname<>""
            If LenB(.LogName) = 0 Then
                Debug.Print "Design check: LogName can't ever be NullString when closing ???"
                Debug.Assert False
            Else
                .LogNamePrev = .LogName
            End If
            If LenB(.LogNameNext) > 0 And InStr(.LogNameNext, Left(DateIdNB, 8)) = 0 Then
                msg = "Switching to new day, old Logfile " & .LogName
                .LogNameNext = vbNullString               ' force new time stamp
                KeepName = False
            End If
            If .LogNameNext = .LogName And Not KeepName Then
                .LogNameNext = vbNullString               ' force new time stamp
                msg = "Next Logfile will get new timestamp"
            ElseIf LenB(.LogNameNext) > 0 Or KeepName Then
                If KeepName Then
                    .LogNameNext = .LogName
                    msg = "Next Logfile will not change, to be re-opened: " & .LogName
                Else
                    msg = "Next selected Logfile: " & .LogNameNext
                End If
            Else
                If InStr(.LogName, Left(DateIdNB, 8)) = 0 Then
                    msg = "Restarting for new day, old Logfile " & .LogName
                    KeepName = False
                Else
                    msg = "Logfile closed:    " & .LogName
                End If
                .LogName = vbNullString
            End If
            
            msg = String(OffCal, b) & msg
            On Error Resume Next
            If .LogIsOpen Then
                Print #1, msg
                Debug.Print msg
            End If
            Close #1
            .LogName = vbNullString
            .LogIsOpen = False
            .LogFileLen = 0
        End If
    End With ' T_DC

FuncExit:
    Recursive = False

ProcRet:
End Sub ' File_FolderOps.CloseLog

'---------------------------------------------------------------------------------------
' Method : Sub ShowLog
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowLog()
    If OlSession Is Nothing Then
        Call N_CheckStartSession(False)
    End If
    Call ShowLogWait(False)
End Sub ' File_FolderOps.ShowLog

'---------------------------------------------------------------------------------------
' Method : Sub ShowLogWait
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowLogWait(KeepName As Boolean)

Dim DateNow As String
Dim daFilt As String
Dim nDigits As String
Dim NextFile As String
   
    With T_DC
        If LenB(.LogName) = 0 Then
            DateNow = GetDateId(-2)
getDigits:
            Call ClipBoard_SetData(DateNow)
            nDigits = InputBox( _
                "Optional: Geben Sie an, auf wieviel Stellen genau die Log-Namen mit " _
                & vbCrLf & "   " & DateNow & "*.log übereinstimmen sollen (1-" & Len(DateNow) - 2 & ")" _
                & vbCrLf & "oder geben sie den gewünschten Teil des Namens ein" _
                & vbCrLf & " sie können mit Ctrl-v den Wert direkt einfügen", _
                "Mehrfachdatei-öffnen")
            If IsNumeric(nDigits) Then
                If CDbl(nDigits) > Len(DateNow) - 2 Then
                    daFilt = nDigits & "*.log"
                Else
                    daFilt = Left(DateNow, nDigits) & "*.log"
                    GoTo DoRunEditor
                End If
            Else
                daFilt = Left(DateNow, 10) & "*.log"
                GoTo DoRunEditor
            End If
            rsp = MsgBox("Bestätigen Sie, dass " & nDigits _
                    & " der Logdatei übereinstimmen sollen, also " _
                    & daFilt & vbCrLf & "(in " & lPfad & ")" _
                & vbCrLf & "   Erneut versuchen==>Nein", _
                vbMsgBoxSetForeground + vbYesNoCancel)
            If rsp = vbCancel Then
                Debug.Print "Unable to edit log file because name is not known"
                DoVerify False
            ElseIf rsp = vbNo Then
                GoTo getDigits
            ElseIf rsp = vbYes Then
DoRunEditor:
                Call RunEditor(lPfad & "Outlook VBA " & daFilt)
            End If
        Else
            Call CloseLog(KeepName:=True)
            Call RunEditor(.LogNamePrev)
        End If
    End With ' T_DC

End Sub ' File_FolderOps.ShowLogWait

'---------------------------------------------------------------------------------------
' Method : Sub RunEditor
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub RunEditor(FullFileName As String)

Const zKey As String = "File_FolderOps.RunEditor"
    Call DoCall(zKey, tSub, eQzMode)

    Call N_Suppress(Push, zKey)

' Open and Close a File
Dim vPID As Variant
Dim rsp As VbMsgBoxResult

    ' Launch file
    If LenB(FullFileName) = 0 Then
        Call MsgBox("No file name specified. Exiting")
        GoTo zExit
    End If
    
    vPID = Shell(Quote(EditProg) & b & Quote(FullFileName), vbNormalFocus)
    
    rsp = MsgBox("Editor showing '" & FullFileName & "'" _
                & vbCrLf & " Press Yes to Terminate Editor" _
                & vbCrLf & " Press No to ignore: continue running Editor and Macros," _
                & vbCrLf & " Press Cancel to Debug and optionally Terminate" _
                , vbYesNoCancel)
    Select Case rsp
    Case vbYes
        ' Kill file
        Call Shell("TaskKill /F /PID " & CStr(vPID), vbHide)
    Case vbCancel
        DoVerify False
        Call TerminateRun
    End Select

ProcReturn:
    Call N_Suppress(Pop, zKey)

zExit:
    Call DoExit(zKey)

End Sub ' File_FolderOps.RunEditor


'---------------------------------------------------------------------------------------
' Method : Function CreateOpenFileEntry
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CreateOpenFileEntry(Key As Long, PathAndName As String, reqState As String, Optional ByVal makeOne As Boolean = False) As Long
Const zKey As String = "File_FolderOps.CreateOpenFileEntry"
    Call DoCall(zKey, tFunction, eQzMode)
    
    Call FileEntryExists(Key, PathAndName, makeOne, reqState)
  
zExit:
    Call DoExit(zKey)
  
End Function ' File_FolderOps.CreateOpenFileEntry

'---------------------------------------------------------------------------------------
' Method : Function FileEntryExists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FileEntryExists(Key As Long, PathAndName As String, Optional makeOne As Boolean = False, Optional reqState As String) As Long
Const zKey As String = "File_FolderOps.FileEntryExists"
    Call DoCall(zKey, tFunction, eQzMode)
    
    If OpenFileNames Is Nothing Then
        Set OpenFileNames = New Dictionary
    End If
    If ClosedFileNames Is Nothing Then
        Set ClosedFileNames = New Dictionary
    End If
    If FileStates Is Nothing Then
        Set FileStates = New Dictionary
    End If
    
    DoVerify Key <> inv, "*** invalid Key: " & Key
    Key = Abs(Key)                                          ' file numbers must be >= 0
    aFileSpec = CStr(inv)                                   ' undefined path and name
    
    If OpenFileNames.Exists(Key) Then
        aFileSpec = OpenFileNames.Item(Key)
        If aFileSpec <> PathAndName Then
            aOpenKey = inv
            GoTo NoFit
        End If
        aOpenKey = Key
    End If
    If ClosedFileNames.Exists(Key) Then
        aFileSpec = ClosedFileNames.Item(Key)
        If aFileSpec <> PathAndName Then
            aClosedKey = inv
            GoTo NoFit
        End If
    End If
    
    aBugVer = aOpenKey <> inv And aClosedKey <> inv
    aBugTxt = "File Key=" & Key & " is both open and closed ??? " & aFileSpec & " aFileState=" & aFileState
    If DoVerify Then
        If aFileState <> inv Then
            Call ClosedFileNames.Remove(Key)
            Call CloseFile(Key)
            Call OpenFileNames.Remove(Key)          ' incomplete fix
            aFileState = 0
            aFileSpec = inv
        End If
    End If
    If aOpenKey <> inv Then
        aClosedKey = inv
    Else
        aOpenKey = inv
    End If
    
NoFit:
    If makeOne Then
        aBugVer = Mid(PathAndName, 2, 2) = ":\"
        aBugTxt = "the PathName incorrect: '" & Quote(PathAndName)
        DoVerify
        
        If aClosedKey <> inv Then
            FileEntryExists = False
        ElseIf aOpenKey <> inv Then
            FileEntryExists = True
        Else
            FileEntryExists = False
        End If
    End If
    If FileEntryExists Then
        If FileStates.Exists(Key) Then
            aFileState = FileStates.Item(Key)
        Else
            Call FileStates.Add(Key, "New")
        End If
    End If

zExit:
    Call DoExit(zKey)

End Function ' File_FolderOps.FileEntryExists

'---------------------------------------------------------------------------------------
' Method : Sub openForAccess
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub openForAccess(FullPath, OpenMode As String, nFile As Long)

Const zKey As String = "File_FolderOps.openForAccess"
    Call DoCall(zKey, tSub, eQzMode)

Dim modeAcc As Variant

    modeAcc = split(OpenMode)
    If UBound(modeAcc) > 0 Then
        Debug.Print OpenMode & " not implemented yet, too many situations to cover"
        DoVerify False
    End If
    
    Select Case LCase(modeAcc(0))
    Case "append"
        Open FullPath For Append As #nFile
    Case "binary"
        Open FullPath For Binary As #nFile
    Case "input"
        Open FullPath For Input As #nFile
    Case "output"
        Open FullPath For Output As #nFile
    Case "random"
        Open FullPath For Random As #nFile
    Case Else
        Debug.Print OpenMode & " is invalid file open mode"
        DoVerify False
    End Select
    
zExit:
    Call DoExit(zKey)

End Sub ' File_FolderOps.openForAccess

'---------------------------------------------------------------------------------------
' Method : Sub openFile
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub openFile(nFile As Long, fPath As String, _
                    fName As String, fExt As String, _
                    Optional OpenMode As String = "Append")

Const zKey As String = "File_FolderOps.openFile"
    Call DoCall(zKey, tSub, eQzMode)

Dim FullPath As String
    
    FullPath = fPath & "\" & fName & fExt
    If nFile = 0 Then
        nFile = OpenFileNames.Count + 1 ' next free slot
    End If
    Call FileEntryExists(nFile, FullPath, makeOne:=True, reqState:=OpenMode)
    
    If aFileState = "New" Then
        GoTo noCl
    End If
        
    If aFileState <> OpenMode Then
        If aFileState <> "Closed" Then
            Print #nFile, Quote(aFileSpec) & "File will be closed and reopened with mode=" & OpenMode
            Err.Clear
            If Not CloseFile(nFile) Then
                GoTo noCl                       ' close did not work ...
            End If
        End If
    End If

noCl:
    Print #nFile, Quote(aFileSpec) & "*** attempting File open with mode=" & OpenMode

    If aFileState <> OpenMode Then
        E_Active.Permit = "*"
        Call openForAccess(FullPath, OpenMode, nFile)
        If ErrorCaught = 0 Then
            If DebugMode Or DebugLogging Then
                Debug.Print "Open " & FullPath _
                    & " for " & OpenMode _
                    & " As #" & nFile & " successful"
            End If
            aFileState = OpenMode
        Else
            Debug.Print "Open " & FullPath _
                & " for " & OpenMode _
                & " As #" & nFile & " failed, Error " _
                & Err.Number & ": " & Err.Description
            Err.Clear
            aFileState = "Undefined"
        End If
    Else
        If DebugMode Then
            DoVerify False, "*** file is open already and no reopen specified"
        End If
    End If

zExit:
    Call DoExit(zKey)

End Sub ' File_FolderOps.openFile

'---------------------------------------------------------------------------------------
' Method : Function GetFileInDir
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetFileInDir(FileNamePrefix As String, maxFileBefore As String, Optional MasterPath As String = "GeneratedCode", Optional CodeNamePattern As String = "*.bas") As String

Const zKey As String = "File_FolderOps.GetFileInDir"
    Call DoCall(zKey, tFunction, eQzMode)

Dim thisFile As String
Dim thisSuffix As String
Dim aFileNumber As Long
Dim maxFileNumber As Long
Dim nextFileNumber As Long
Dim aFileDate As Date
Dim maxFileDate As Date
Dim fileCount As Long
Dim RealPath As String
Dim FoundFiles As Variant
Dim isMaxDate As Boolean
Dim j As Long
    maxFileBefore = vbNullString
    maxFileNumber = 0
    maxFileDate = CDate(0)
   
    FoundFiles = getDirFileList(MasterPath, fileCount, "\" & FileNamePrefix & CodeNamePattern, False)
    If fileCount = 0 Then
        MsgBox "keine Dateien gefunden, die " & Quote(CodeNamePattern) & " heissen. (Ordner: " & Quote(MasterPath) & " )"
        End
    End If
    RealPath = MasterPath
    For j = 0 To fileCount  ' not sorted, so we have to do all of them
        thisFile = FoundFiles(j)
        If InStr(thisFile, "\") > 0 Then
            thisFile = RTail(thisFile, "\", RealPath)
        End If
        thisFile = Trunc(1, thisFile, ".") ' drop "extension" (left to right!)
        If Left(thisFile, Len(FileNamePrefix)) = FileNamePrefix Then
            thisSuffix = Mid(thisFile, Len(FileNamePrefix) + 1)
            If Not isMaxDate And IsNumeric(thisSuffix) Then ' use numeric only if we have no dates
                aFileNumber = CLng(thisSuffix)
                If aFileNumber > maxFileNumber Then
                    maxFileNumber = aFileNumber
                    nextFileNumber = maxFileNumber + 1
                    maxFileBefore = thisFile
                End If
            ElseIf IsDate(thisSuffix) Then
                isMaxDate = True    ' no longer loook for numbers
                aFileDate = CDate(thisSuffix)
                If aFileDate > maxFileDate Then
                    maxFileBefore = thisFile
                    maxFileDate = aFileDate
                End If
            Else
                DoVerify False
            End If
        End If
    Next j
    If isMaxDate Then
        GetFileInDir = RealPath & "\" & Now()
    Else
        GetFileInDir = RealPath & "\" & LPad(CStr(maxFileNumber + 1), 8)
    End If
    ' add extension
    GetFileInDir = GetFileInDir & Mid(CodeNamePattern, 2)
  
End Function ' File_FolderOps.GetFileInDir

'---------------------------------------------------------------------------------------
' Method : Function getMasterPath
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getMasterPath(startOfPath As String) As String

Const zKey As String = "File_FolderOps.getMasterPath"
    Call DoCall(zKey, tFunction, eQzMode)

Dim Count As Long
Dim thisFile As String
Dim FoundFiles As Variant

    thisFile = RTail(olApp.RecentFiles(1).Path, "\", startOfPath)
    If thisFile = "Briefe" Then
        getMasterPath = startOfPath & "\Briefe"
    Else
        FoundFiles = getDirFileList(startOfPath, Count, "\*.doc*", True)       ' doc or docx, only first file
        If Count = 0 Then
            MsgBox "keinen Ordner ""Briefe"" gefunden (in Ordner " & startOfPath & ")"
            End
        Else
            thisFile = FoundFiles(0)
            thisFile = RTail(thisFile, "\", getMasterPath)
        End If
    End If
  
End Function ' File_FolderOps.getMasterPath

'---------------------------------------------------------------------------------------
' Method : Function getDirFileList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getDirFileList(dirName As String, fileCount As Long, fileType As String, Optional firstFileOnly As Boolean) As Variant

Const zKey As String = "File_FolderOps.getDirFileList"
    Call DoCall(zKey, tFunction, eQzMode)

Dim MyFile As String
ReDim DirListArray(1000) As String

    MyFile = Dir$(dirName & fileType)
    fileCount = 0
    Do While MyFile <> vbNullString
        DirListArray(fileCount) = MyFile
        MyFile = Dir$
        fileCount = fileCount + 1
        If firstFileOnly Or fileCount > 1000 - 1 Then
            Exit Do
        End If
    Loop
    
    fileCount = fileCount - 1
    ' Reset the size of the array without losing its values by using Redim Preserve
    ReDim Preserve DirListArray(fileCount)
    Application.WordBasic.sortarray DirListArray()
    getDirFileList = DirListArray
  
End Function ' File_FolderOps.getDirFileList

'---------------------------------------------------------------------------------------
' Method : CloseFile
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Close file #nFile and note this in OpenFileNames
'---------------------------------------------------------------------------------------
Function CloseFile(nFile As Long) As Boolean
Const zKey As String = "File_FolderOps.CloseFile"
    Call DoCall(zKey, tSub, eQzMode)

    E_Active.Permit = "*"
    Close #nFile
    If ErrorCaught <> 0 Then
        Err.Clear
        CloseFile = False
    End If
    CloseFile = True
    ClosedFileNames.Add nFile, OpenFileNames.Item(nFile)
    aFileState = "Closed"
    Call OpenFileNames.Remove(nFile)
    FileStates.Item(nFile) = aFileState

zExit:
    Call DoExit(zKey)

End Function ' File_FolderOps.CloseFile


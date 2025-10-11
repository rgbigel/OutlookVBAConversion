Attribute VB_Name = "HelpersOL"
Option Explicit

' other stuff
Public AddType As String        ' task type added by eventroutines below

' mail-like objects, just to make readable
Public CurrentSessionEmail As MailItem
Public CurrentSessionReport As ReportItem
Public CurrentSessionMeetRQ As MeetingItem
Public CurrentSessionTaskRQ As TaskItem

Public SpecialSearchComplete As Boolean
Public forceOrdering As Boolean
Public NeverDelete As Boolean
Private DeleteAllOld As Boolean

'---------------------------------------------------------------------------------------
' Method : Sub Z_olInits
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Z_olInits()
'--- Proc MAY ONLY CALL Z_Type PROCS                        ' Standard procConst zKey As String = "HelpersOL.Z_olInits"
Const zKey As String = "HelpersOL.Z_olInits"

Dim sBlockEvents As Boolean
Dim i As Long

    Call DoCall(zKey, "Sub", eQzMode)

    sBlockEvents = E_Active.EventBlock        ' Save for exit state
    E_Active.EventBlock = True                ' no events during init
    If ErrStatusFormUsable Then
        frmErrStatus.fNoEvents = E_Active.EventBlock
        Call BugEval
    End If
    StopLoop = False
    
    If OlSession Is Nothing Then
        Stop ' *****imposs
        Set OlSession = New cOutlookSession
    End If
    If dftRule Is Nothing Then
        Set dftRule = New cAllNameRules
    End If
    ActionTitle(0) = "Undefined or dynamic action"
    StopRecursionNonLogged = StopRecursionNonLogged Or EPCalled
    EPCalled = IsEntryPoint
        
    Set RestrictedItemCollection = New Collection
    With dftRule                                 ' First-Time Inits if invalid:
        If Not dftRule.RuleInstanceValid Then
            Call getAccountDscriptors
            ' contains all properties that can not be decoded
            If aPindex = 0 And aID(0) Is Nothing Then
                Set aOD(0) = New cObjDsc
                Set aOD(0).objSeqInImportant = New Collection
                If D_TC Is Nothing Then
                    Set D_TC = New Dictionary
                End If
                D_TC.Add "0", aOD(0)
                aOD(0).objItemType = inv         ' mark this as invalid
                Set aOD(0).objClsRules = dftRule ' NO cloning here!
                Set aOD(0).objClsRules.clsNeverCompare.PropAllRules = dftRule
                Set aOD(0).objClsRules.clsObligMatches.PropAllRules = dftRule
                Set aOD(0).objClsRules.clsNotDecodable.PropAllRules = dftRule
                Set aOD(0).objClsRules.clsSimilarities.PropAllRules = dftRule
            End If
            If Left(ActionTitle(0), 5) = IgnoredHeader Then
                Debug.Print ActionTitle(0)
                DoVerify False
                Call SetStaticActionTitles
            End If
            If LenB(ActionTitle(UBound(ActionTitle))) = 0 Then
                Call SetStaticActionTitles
            End If
            Set LookupFolders = aNameSpace.Folders
            
            If LF_CurActionObj Is Nothing Then
                Call DefineLocalEnvironment
            End If
            CurIterationSwitches.SaveItemRequested = True ' for mail processing, Action = 3
            
            ' Obligatory Matches: the only case where we
            ' access class Property "aRuleString" is to set as default
            TrueCritList = "Subject"
            .clsObligMatches.ChangeTo = TrueCritList
            ' Not Decodable
            .clsNotDecodable.ChangeTo = "ItemProperties " _
                                      & "Session RTFBody FormDescription PermissionTemplateGuid " _
                                      & "GetInspector PropertyAccessor SaveSentMessageFolder IsLatestVersion "
            ' dont compare
            DontCompareListDefault = "*ID: Organizer Ordinal *Time* *UTC Size " _
                                   & "SenderName SentOn SentOnBehalfOfName " _
                                   & "*Version *DisplayName: *Xml " _
                                   & "CompanyAnd* CompanyLast* FullNameAnd* " _
                                   & "LastNameAnd* Yomi* BusinessCard* " _
                                   & "BillingInformation ConversationIndex Mileage Saved " _
                                   & "UnRead VotingOptions VotingResponse Parent MessageClass " _
                                   & "OutlookVersion OutlookInternalVersion " _
                                   & "BodyFormat InternetCodepage Left Top Width Height " _
                                   & "Organizer SendUsingAccount GlobalAppointmentID "
            .clsNeverCompare.ChangeTo = DontCompareListDefault
            ' Similarities
            .clsSimilarities.ChangeTo = "Parent Categories"
            ' mark this initialization done:
            .RuleType = DefaultRule
            .clsNeverCompare.RuleMatches = False
            .clsSimilarities.RuleMatches = False
            .clsObligMatches.RuleMatches = False
            .clsNotDecodable.RuleMatches = False
            .RuleInstanceValid = True            ' and Write-Only now?
        End If
    End With                                     ' dftrule
    
    ' Lists must have leading and trailing blank for each word  !
    Set killWords = Nothing
    Set killWords = New Collection
    sourceIndex = 0
    targetIndex = 0
    Set sDictionary = Nothing
    Set aProp = Nothing
    apropTrueIndex = inv
    workingOnNonspecifiedItem = False
    BaseAndSpecifiedDiffer = False
    If TrashFolder Is Nothing Then
        Set TrashFolder = aNameSpace.GetDefaultFolder(olFolderDeletedItems)
        TrashFolderPath = TrashFolder.FolderPath
    End If
    For i = 1 To UBound(pArr)
        pArr(i) = Chr(0)
    Next i
    
    ' Inits for RuleTable
staticRuleTable = True
    UseExcelRuleTable = False                    ' execute sub InitRuleTable
    
    Call InitEventTraps
    Set Deferred = New Collection
    
    Debug.Print "* EntryPoint=" & IsEntryPoint & "* Testvar=" & Quote(Testvar)
    E_Active.EventBlock = sBlockEvents        ' all first-time inits were done
    If ErrStatusFormUsable Then
        frmErrStatus.fNoEvents = E_Active.EventBlock
        Call BugEval
    End If

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                          ' HelpersOL.Z_olInits

'---------------------------------------------------------------------------------------
' Method : Sub ClearUnwantedCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClearUnwantedCats(outStore As Store, inStore As Store)
Dim zErr As cErr
Const zKey As String = "HelpersOL.ClearUnwantedCats"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim inCat As category
Dim outCat As category
Dim i As Long
Dim k As Long
Dim anyMod As Boolean
Dim thisMod As String
    ' loop to remove existing invalid category definitions (invalid = not in inStore)
    For i = outStore.Categories.Count To 1 Step -1
        If i > outStore.Categories.Count Then
            Exit For
        End If
        Set outCat = outStore.Categories.Item(i)
        If DeleteAllOld Then                     ' set this to force ordering like in inStore
            thisMod = outCat.Name
            outStore.Categories.Remove i
            thisMod = i & vbTab & "Category " & Quote(thisMod) & " deleted from Store " & outStore.DisplayName
            anyMod = True
        Else
            thisMod = outCat.Name
            Set inCat = StoreCatGet(inStore, thisMod, k)
            If inCat Is Nothing And Not NeverDelete Then ' no wanted because no valid inCat
                outStore.Categories.Remove thisMod
                thisMod = i & vbTab & "Category deleted " & Quote(thisMod) _
      & " previously in pos. " & k & " of Store " & outStore.DisplayName
                anyMod = True
            End If
            If anyMod And i <> k Then
                DeleteAllOld = True              ' force original order from here on
                ' will not create full ordering!
            End If
        End If
        If DebugMode Then
            Debug.Print thisMod
            thisMod = vbNullString
        End If
    Next i
    If anyMod Then
    Else
        If DebugMode Then
            Debug.Print "no unwanted Categories had to be removed"
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.ClearUnwantedCats

'---------------------------------------------------------------------------------------
' Method : Function CompareCatNames
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CompareCatNames(outStore As Store, inStore As Store, Optional outStoreDelete As Boolean = False) As Long
Dim zErr As cErr
Const zKey As String = "HelpersOL.CompareCatNames"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim returnvalue As Long
    ' 0 if identical
    ' +1 if not identical
    ' -1 if not in same sequence
Dim i As Long
Dim k As Long
Dim mini As Long
Dim maxi As Long
Dim inCatName As String
Dim outCatName As String
Dim inCat As category
Dim outCat As category
Dim bigCatStore As Store
Dim compMsg As String

    'On Error GoTo 0     ' allow no errors in this proc
    If inStore.Categories.Count > outStore.Categories.Count Then
        mini = outStore.Categories.Count
        maxi = inStore.Categories.Count
        Set bigCatStore = inStore
        compMsg = " more items in InStore " & Quote(inStore.DisplayName) _
      & " than in " & Quote(outStore.DisplayName)
        returnvalue = 1
    ElseIf inStore.Categories.Count < outStore.Categories.Count Then
        returnvalue = 1
        mini = inStore.Categories.Count
        maxi = outStore.Categories.Count
        Set bigCatStore = outStore
        compMsg = " fewer items in InStore " & Quote(inStore.DisplayName) _
      & " than in " & Quote(outStore.DisplayName) & " will delete=" & outStoreDelete
    Else
        mini = inStore.Categories.Count
        maxi = mini
    End If
    
    If DebugMode And LenB(compMsg) > 0 Then
        Debug.Print (maxi - mini) & compMsg
    End If
    
    k = 1                                        ' initally = i
    i = 1
    Do                                           ' find by items in outStore
        If i > outStore.Categories.Count Then
            outCatName = vbNullString
        Else
            Set outCat = outStore.Categories.Item(i)
            outCatName = outCat.Name
        End If
        If k > inStore.Categories.Count Then
            GoTo gotNone
        End If
        Set inCat = inStore.Categories.Item(k)   ' comparing inStore
        inCatName = inCat.Name
        If LenB(outCatName) = 0 Then
            compMsg = i & vbTab & inCatName & " occurs in Store " & Quote(inStore.DisplayName) & " in position " & k & " but does not exist in store " & Quote(outStore.DisplayName)
            k = k + 1                            ' step k
            If returnvalue = 0 Then
                returnvalue = -1
            End If
            GoTo NotInTarget
        End If
        If inCatName <> outCatName Then
            Set inCat = StoreCatGet(inStore, outCatName, k)
            If inCat Is Nothing Then
gotNone:
                compMsg = i & vbTab & outCatName & " does not occur in Store " & Quote(inStore.DisplayName)
                returnvalue = 1
                If outStoreDelete Then
                    outStore.Categories.Remove i
                    maxi = maxi - 1
                    ' no step in k, not present any longer
                    compMsg = compMsg & " and was deleted from Store " & Quote(outStore.DisplayName)
                Else
                    k = k + 1                    ' step k
                End If
            Else
                compMsg = i & vbTab & outCatName & " occurs in Store " & Quote(inStore.DisplayName) & " in position " & k
                If returnvalue = 0 Then
                    returnvalue = -1
                End If
                k = k + 1                        ' step k
            End If
        Else
            compMsg = i & vbTab & inCatName & " occurs in same position in " & Quote(outStore.DisplayName)
            k = k + 1                            ' step k
        End If
        If DebugMode Then
NotInTarget:
            Debug.Print compMsg
        End If
        If k < 1 Then                            ' leave if we deleted the last one
            Exit Do
        End If
        i = i + 1
    Loop Until i > maxi
    
    If outStoreDelete Then
        GoTo fini
    End If
    
    For i = mini + 1 To maxi
        Set outCat = bigCatStore.Categories.Item(i)
        outCatName = outCat.Name
        compMsg = i & vbTab & outCatName & " occurs only in " & Quote(bigCatStore.DisplayName)
        If outStoreDelete And Not NeverDelete And bigCatStore.DisplayName = outStore.DisplayName Then
            outStore.Categories.Remove i
            i = i - 1                            ' not present any longer
            maxi = maxi - 1
            compMsg = compMsg & ". It has been deleted"
        End If
        If DebugMode Then
            Debug.Print compMsg
        End If
        
        returnvalue = 1
        If i < 1 Then                            ' leave if we deleted the last one
            Exit For
        End If
    Next i

fini:
    CompareCatNames = returnvalue

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' HelpersOL.CompareCatNames

'---------------------------------------------------------------------------------------
' Method : Sub CopyAllBackupCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyAllBackupCats()                          ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "HelpersOL.CopyAllBackupCats"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="HelpersOL")

Dim sourceStore As Store

    IsEntryPoint = True
    Set sourceStore = FolderBackup.Store         ' global for session, from Z_AppEntry
    Call ShowCats(sourceStore)
    forceOrdering = False
    Call CopyAllCats(Array(OlHotMailHome, OlWEBmailHome, OlGooMailHome), sourceStore)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.CopyAllBackupCats

'---------------------------------------------------------------------------------------
' Method : Sub CopyAllCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyAllCats(targetStoreArray As Variant, sourceStore As Store)
Dim zErr As cErr
Const zKey As String = "HelpersOL.CopyAllCats"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim targetStore As Store
Dim TargetFolder As Folder
Dim i As Long
    For i = 0 To UBound(targetStoreArray)
        NeverDelete = Not forceOrdering
        Set TargetFolder = targetStoreArray(i)
        Set targetStore = TargetFolder.Store
        Call CopyCats(targetStore, sourceStore)
    Next i

FuncExit:

    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.CopyAllCats

'---------------------------------------------------------------------------------------
' Method : Sub CopyAllHotmailCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyAllHotmailCats()                         ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "HelpersOL.CopyAllHotmailCats"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="HelpersOL")

Dim sourceStore As Store

    IsEntryPoint = True
    
    Set sourceStore = OlHotMailHome.Store        ' global for session, from Z_AppEntry
    Call ShowCats(sourceStore)
    forceOrdering = True
    Call CopyAllCats(Array(FolderBackup, OlWEBmailHome, OlGooMailHome), sourceStore)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.CopyAllHotmailCats

'---------------------------------------------------------------------------------------
' Method : Sub CopyCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyCats(outStore As Store, inStore As Store)
Dim zErr As cErr
Const zKey As String = "HelpersOL.CopyCats"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim inCat As category
Dim outCat As category
Dim j As Long
Dim k As Long
Dim anyMod As Boolean
Dim thisMod As String
Dim oIdent As String
    anyMod = False
    DeleteAllOld = forceOrdering
    If DebugMode Then
        Call ShowCats(outStore)
    End If
        
    If DeleteAllOld Then
        Call ClearUnwantedCats(outStore, inStore)
    Else
        Call CompareCatNames(outStore, inStore, outStoreDelete:=True)
    End If
            
    ' loop to add / alter new category definitions
    k = 1                                        ' position in outstore
    For j = 1 To inStore.Categories.Count
        thisMod = vbNullString
        Set inCat = inStore.Categories.Item(j)
        oIdent = j & vbTab & inCat.Name & vbTab & " in " & Quote(outStore.DisplayName)
        If j > outStore.Categories.Count Then    ' cat obviously missing:
            Set outCat = StoreCatGet(outStore, inCat.Name, k)
            If outCat Is Nothing Then
                outStore.Categories.Add inCat    ' add to end of list again
                k = outStore.Categories.Count
                thisMod = "Category added as number " & k
            Else
                DoVerify False, " incorrect ordering in outStore"
                DeleteAllOld = True              ' disorder???
                thisMod = "Category already present in pos. " & k
            End If
        Else
            Set outCat = StoreCatGet(outStore, inCat.Name, k)
            If outCat Is Nothing Then
                outStore.Categories.Add inCat    ' might change sequence?
                k = outStore.Categories.Count
                thisMod = "New Category added as number " & k
            Else
                If j <> k Then
                    DeleteAllOld = True          ' disorder???
                End If
                If outCat.ShortcutKey <> inCat.ShortcutKey Then
                    Call AppendTo(thisMod, vbTab _
                                        & "corrected ShortcutKey mismatch " _
                                        & outCat.ShortcutKey _
                                        & " to " & inCat.ShortcutKey)
                    outCat.ShortcutKey = inCat.ShortcutKey
                End If
                If outCat.Color <> inCat.Color Then
                    Call AppendTo(thisMod, vbTab _
                                        & "corrected Color mismatch " _
                                        & outCat.Color _
                                        & " to " & inCat.Color)
                    outCat.Color = inCat.Color
                End If
            End If
        End If
        If LenB(thisMod) = 0 Then
            If DebugMode Then
                Debug.Print oIdent & vbTab & thisMod & vbTab _
                          & "no change, ShortcutKey " & vbTab _
                          & outCat.ShortcutKey _
                          & ", Color " & outCat.Color _
                          & " as number " & k
            End If
        Else
            anyMod = True
            If DebugMode Then
                Debug.Print oIdent & thisMod
            End If
        End If
    Next j
        
    If DebugMode Then
        If outStore.Categories.Count <> inStore.Categories.Count Then
            DoVerify False
            Debug.Print "There is a difference in the number of categories: Backup has " _
                      & inStore.Categories.Count & ", target has " _
                      & outStore.Categories.Count
        End If
    End If
    If anyMod Then
        ' save? outStore ???
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.CopyCats

'---------------------------------------------------------------------------------------
' Method : CreateSearchFolder
' Author : Mark Withall @markwithall.com
' Date   : 20211108@11_47
' Purpose: Create Search Folder
'---------------------------------------------------------------------------------------
Function CreateSearchFolder(Account As String, ScopeFolders As String, Filter As String, FolderName As String, Optional WithSubFolders As Boolean) As Boolean
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "HelpersOL.CreateSearchFolder"
Const zKey As String = "HelpersOL.CreateSearchFolder"

    Call DoCall(zKey, "Function", eQzMode)

Dim CreateOk As Boolean
Dim objSearch As Search
Dim Warn As String
Dim sFolder As Folder

    If olApp Is Nothing Then
        Call Z_StartUp
    End If
    
    If LenB(ScopeFolders) > 0 Then
        If LenB(Account) = 0 Then
            Account = Replace(ScopeFolders, "'", vbNullString)
            Account = Trunc(3, Account, "\")     ' single Store only
            Warn = "Account defaulted to " & Account
        Else
            aBugTxt = "Error in ScopeFolcers, not matching account"
            DoVerify InStr(LCase(ScopeFolders), "\\" & LCase(Account) & "\") > 0
        End If
    Else
        If LenB(Account) = 0 Then
            Account = olApp.Session.Stores.Item(1)
            Warn = " in first Store"
        End If
        ScopeFolders = "'\\" & Account & "\" & StdInboxFolder & "'"
        Warn = "ScopeFolders defined as " & ScopeFolders & Warn
    End If
    
    If LenB(Warn) > 0 Then
        Debug.Print Warn
    End If
    
    If SearchFolderExists(Account, FolderName, sFolder) Then
        If LenB(Warn) > 0 Then
            Call LogEvent("Would like to test of Scopes Match expectations, " _
                        & "but no access method found for 'Search' Objects")
        End If
        If sFolder.ShowItemCount <> olShowTotalItemCount Then
            Call LogEvent("'.ShowItemCount' indicates that SearchFolder " _
                        & FolderName & vbCrLf & "    in " & sFolder.FolderPath _
                        & " may not be synchronized." & vbCrLf & "    Re-Creating it.")
            GoTo DefineAgain
        End If
        Call LogEvent("SearchFolder " & Quote(sFolder.FolderPath) _
                    & " is up to date and usable")
        CreateSearchFolder = False
    Else
DefineAgain:
        aBugTxt = "set up Advanced search, folder '" & FolderName _
                & "' in " & ScopeFolders
        Call Try
        Set objSearch = olApp.AdvancedSearch( _
                        Scope:=ScopeFolders, _
                        Filter:=Filter, _
                        SearchSubFolders:=True, _
                        Tag:=FolderName)
        aBugTxt = "Save search folder '" & FolderName _
                & "' in " & Account
        Call Try                                 ' Try anything, autocatch
        Call objSearch.Save(FolderName)
        Catch
        
        CreateSearchFolder = True
        CreateOk = SearchFolderExists(Account, FolderName, sFolder)
        If Not CreateOk Then
            Call LogEvent("Missing search folder '" & FolderName & "' in " & Account)
            If Not sFolder Is Nothing Then
                sFolder.ShowItemCount = olShowTotalItemCount ' default for MY search folders
            End If
            GoTo FuncExit
        End If
        Call LogEvent("created search folder '" & FolderName & "' in " & Account _
                    & vbCrLf & "may need to wait for AdvancedSearchComplete-Event")
    End If
    
    UInxDeferred = UInxDeferred + 1
    Set DeferredFolder(UInxDeferred) = sFolder
    
FuncExit:
    Set sFolder = Nothing
    Set objSearch = Nothing

zExit:
    Call DoExit(zKey)
ProcRet:
End Function                                     ' HelpersOL.CreateSearchFolder

'---------------------------------------------------------------------------------------
' Method : Sub DoMaintenance
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub DoMaintenance()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "HelpersOL.DoMaintenance"
    Call DoCall(zKey, "Sub", eQzMode)

    frmMaintenance.Show
    MaintenanceAction = 0

FuncExit:

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.DoMaintenance

'---------------------------------------------------------------------------------------
' Method : Sub ForceSave
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ForceSave(aItemO As Object, Optional MessagePrefix As String = vbNullString)
Dim zErr As cErr
Const zKey As String = "HelpersOL.ForceSave"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim defectiveCopy As Variant
Dim logStatus As Boolean
Dim mgPart As String
Dim waitRetries As Long
    
    If aItemO Is Nothing Then
        Debug.Print "can't save a Nothing Object"
        DoVerify False
        GoTo ProcReturn
    End If
    If aItemO.Saved Then
        GoTo ProcReturn                          ' dont save if superfluous
    End If
    logStatus = dontLog
    
    If TrySaveItem(aItemO) Then                  ' klappt nicht, wenn von Outlook inzwischen geändert wurde
        aBugTxt = "Replicate Item EntryID " & aItemO.EntryID
        Call Try                                 ' Try anything, autocatch
        Set defectiveCopy = Replicate(aItemO, delOriginal:=True)
        Catch
        Set aItemO = defectiveCopy
        If aItemO Is Nothing Then
            GoTo FunExit
        End If
        dontLog = True
        If Not logStatus Then
            If Left(aItemO.Parent.Name, 3) = "Log" Then
                Call LogEvent(MessagePrefix & "Conflicting Log copied to trashFolder " _
                            & TrashFolder.FullFolderPath)
            Else
                Call LogEvent(MessagePrefix & "Mail copied to trashFolder " _
                            & TrashFolder.FullFolderPath, eLmin)
            End If
        End If
        Set defectiveCopy = Nothing
    Else
waitTry:
        dontLog = True
        ' if log wanted and not saving the log itself, log save result
        If Not logStatus And Left(aItemO.Parent.Name, 3) <> "Log" Then
            If aItemO.Saved Then
                mgPart = vbNullString
            Else
                waitRetries = waitRetries + 1
                If waitRetries < 5 Then
                    Call Sleep(100)              ' wait .1 seconds, maximally 4 times
                    GoTo waitTry
                ElseIf waitRetries <= 5 Then
                    aItemO.Save
                    Call Sleep(1000)             ' wait 1 seconds, one more time
                    GoTo waitTry
                End If
                mgPart = "not "
            End If
            Call LogEvent("     " & MessagePrefix & TypeName(aItemO) & " changes " & mgPart & "saved after " & waitRetries & " retries in " _
                        & aItemO.Parent.FullFolderPath, eLall)
        End If
    End If
FunExit:
    dontLog = logStatus

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.ForceSave

'---------------------------------------------------------------------------------------
' Method : Sub GenerateTaskReminder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GenerateTaskReminder(ByVal Item As Object)
Dim zErr As cErr
Const zKey As String = "HelpersOL.GenerateTaskReminder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim newReminder As TaskItem
Dim i As Long
Dim j As Long
Dim dueDates As String
Dim DueDate As Date
Dim uAnswer As VbMsgBoxResult

    On Error GoTo dontDoIt
    If Item.Class = olMail Then                  ' nur bei Markierten, gesendeten eMails
        ' nicht bei anderen Mail-like objekten
        If Item.FlagStatus = olFlagMarked Then
            If DebugMode And DebugLogging Then DoVerify False, " ??? test phase"
            aBugTxt = "add new reminder task"
            Call Try                             ' Try anything, autocatch, Err.Clear
            Set newReminder = FolderTasks.Items.Add(olTaskItem)
            If CatchNC Then
                GoTo dontDoIt
            End If
            If Item.Recipients.Count > 0 Then
                newReminder.Assign               ' Zugeordnet!
                For i = 1 To Item.Recipients.Count
                    aBugTxt = "add reminder recipient " & i
                    Call Try                     ' Try anything, autocatch
                    newReminder.Recipients.Add (Item.Recipients(i))
                    If Catch Then
                        GoTo dontDoIt
                    End If
                Next i
            End If
            Call ErrReset(0)
            
            newReminder.Subject = Item.Subject
            newReminder.Body = Item.Body
        
            ' Prüfen, ob ein Fälligkeitsdatum drin steht, ggf. vergleichen
            j = 1
            i = InStr(j, Item.Body, " erled")
            If i > 0 Then
                j = InStr(i, Item.Body, " bis ")
                If j > i Then
                    dueDates = Mid(Item.Body, j + 4)
                    If FindFirstDate(dueDates, DueDate) Then
                        If Item.ReminderSet Then ' doppelte Angabe, ZWEI Wecker gesetzt? Prüfe Sinn!
                            If DueDate > Item.ReminderTime Then ' früher wecken als Ergebnis erwartet? Seltsam
                                uAnswer = _
                                        MsgBox("Die eigene Erinnerungszeit " _
                                             & " liegt vor der Zeit, die dem Empfänger genannt ist." _
                                             & " Wenn dies nicht OK ist, drücken sie 'Abbrechen'." _
                                             & " Dann wird die in der eMail genannte Zeit als Erinnerungszeit verwendet.", _
                                               vbOKCancel)
                                If uAnswer = vbCancel Then
                                    newReminder.DueDate = DueDate ' war wohl ein Versehen, nimm das, was im Text steht
                                Else             ' das OK sagt, ich will vorab geweckt werden
                                    newReminder.DueDate = Item.ReminderTime
                                End If
                            Else                 ' Karenzzeit Angabe, also nimm das Datum wenn die Überprüfung stattfinden muss
                                newReminder.DueDate = Item.ReminderTime
                            End If
                        Else                     ' einzige Angabe ist in der Mail...
                            newReminder.DueDate = DueDate ' also nimm das Datum in der Mail
                        End If
                    Else                         ' in der eMail steht kein Datum,  und es wurde auch sonst keines festgesetzt ?!
                        MsgBox ("Sie wollen eine Aufgabe überprüfen, " _
                              & "haben aber kein Datum angegeben. " _
                              & "Sie sollten die gleich angezeigte Aufgabe " _
                              & "und/oder die eMail ergänzen.")
                    End If
                End If
            End If
            Call LogEvent("Mail displayed to user ", eLnothing)
            newReminder.Display                  ' save is user matter
        End If                                   ' marked status
        Set newReminder = Nothing
    End If                                       ' nur eMails
dontDoIt:
    If Catch Then
        Call LogEvent("Fehler beim Erzeugen einer Aufgabe/Erinnerung", eLall)
    End If

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.GenerateTaskReminder

'---------------------------------------------------------------------------------------
' Method : Sub getInfo
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Use/Get/Set Information in cInfo. Uses ActItemObject to which fInfo belongs
'          if Assign, evaluate aValue of variant/object.
'          if AssignmentMode <=0, determine AssignmentMode 1 or 2
'          if AssignmentMode = 1, convert to String
'          if AssignmentMode = 2, use set to assign
'             and recursively DrillDown the value.
' as soon as AssignmentMode =1 is reached and Assign is True, assign aValue=cstr(.iValue)
'---------------------------------------------------------------------------------------
Sub getInfo(fInfo As cInfo, aValue As Variant, Optional Assign As Boolean = True, Optional MaxDepth As Long = 1)
Dim zErr As cErr
Const zKey As String = "HelpersOL.getInfo"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim dInfo As cInfo
Dim vElement As Long
Dim vValue As Variant

    If fInfo Is Nothing Then
        Set fInfo = New cInfo                    ' Class_Initialize d
    ElseIf fInfo.iType = -1 Then                 ' use this if new set of values needed
        Set fInfo = New cInfo                    ' Class_Initialize to clear
    End If                                       ' .iAssignmentMode not changed for all others, -99 after new
    
    Set dInfo = fInfo                            ' save outermost for recursion
    With fInfo
        If Assign Then
            .DecodedStringValue = vbNullString
            .DecodeMessage = vbNullString
        End If
        If .iAssignmentMode <= 0 Then
            Call InspectType(aValue, dInfo)
            GoTo assignIvalue
        ElseIf .iIsArray Then                    ' arrays need complex analysis every time
            Call InspectType(aValue, dInfo)
            GoTo assignIvalue
        Else                                     ' use value known from last InspectType
assignIvalue:
            If .iAssignmentMode = 0 Then
                Assign = False
            ElseIf .iAssignmentMode = 1 Then
                .iValue = aValue                 ' scalar
            ElseIf .iAssignmentMode = 2 Then
                Set .iValue = aValue             ' not scalar / object
            End If
        End If
        
        ' Note: analyzing .iType may give different results to analyzing .iTypeName via IsScalar
        Select Case .iType
        Case vbInteger, vbLong, vbSingle, vbDouble, vbString, vbBoolean, vbDecimal, 20
            ' all of these are scalars:    2 - 6, 11, 14, 20                       (20=LongLong)
            If .iScalarType < 0 Then             ' not an object having String Value at the same time
                GoTo noVName                     ' they may not have an aValue.Name
            End If
            If aValue.Name = "MemberCount" Then  ' except this one: is an array without count
                ' requires GetMember to access the Objects
                ' which are of type Recipient, NOT Contact
                .iArraySize = .iValue
                .DecodeMessage = "# object array (Contact) "
                .DecodedStringValue = "{} " & .iArraySize & " values"
                    
                aBugTxt = "decoding Members of DistributionList"
                For vElement = 1 To .iArraySize
                    Set vValue = ActItemObject.GetMember(vElement)
                    Set dInfo = dInfo.DrillDown(vValue)
                    .DecodedStringValue = .DecodedStringValue & vbCrLf _
                                        & dInfo.iValue.Name
                Next vElement
                GoTo FuncExit
            End If
noVName:
            If .iDepth >= MaxDepth Then          ' not going any deeper
                .DecodedStringValue = CStr(dInfo.iValue)
                .DecodeMessage = "# non-scalar: " & dInfo.iTypeName
                .iAssignmentMode = 1
            ElseIf Assign Then
                .DecodedStringValue = CStr(.iValue)
                If .iScalarType <= 0 Then
                    .DecodeMessage = "# non-scalar " & .iTypeName
                    If .iClass = olItemProperty Then
                        .DecodeMessage = .DecodeMessage & ", Name=" & .iValue.Name
                    End If
                Else
                    .DecodeMessage = .iTypeName
                End If
            End If
            GoTo FuncExit
        Case vbDate                              '  7 Datumswert (Date)
            If Assign Then
                .DecodedStringValue = CStr(.iValue)
                If .iValue = BadDate Then        ' obviously must exist
                    .DecodedStringValue = "## Datum nicht angegeben"
                    .DecodeMessage = .DecodedStringValue
                Else
                    .DecodeMessage = .iTypeName
                End If
            End If
            GoTo FuncExit
        Case vbObject                            '  9 Objekt
            Set .iValue = aValue
likeObject:
            If .iDepth > MaxDepth Then
                .DecodeMessage = "not decoded, depth " & MaxDepth
                .DecodedStringValue = "## unvollständige Decodierung"
                GoTo nonValue
            End If
            If .iScalarType < 0 Then             ' not decodable
                .DecodeMessage = "not decodable by Rule: "
                .DecodeMessage = .DecodeMessage & .iValue.Name
                .DecodedStringValue = "## " & .iValue.Name & " nicht dekodierbar"
                Call ErrReset(4)                 ' no problem if object has no name
                GoTo RuleDetermined
            ElseIf Not aTD.adRules.clsObligMatches.RuleMatches Then
                .DecodeMessage = "# not oblig.: " & .iValue.Name
                .DecodedStringValue = "## " & .iValue.Name & " nicht oblig."
                .iAssignmentMode = -2
                Call ErrReset(4)                 ' no problem if object has no name
                GoTo RuleDetermined
            End If
            If fInfo.iClass = olItemProperty Then '  all have a .Name
                If DecodeSpecialProperties(fInfo, .iValue.Name) Then
                    If fInfo.iIsArray Then       ' qualified because fInfo may have changed
                        fInfo.DecodedStringValue = dInfo.DecodedStringValue
                        Set dInfo = fInfo
                        For vElement = 1 To fInfo.iArraySize
                            Set dInfo = dInfo.DrillDown(fInfo.iValue.Item(vElement))
                            Call getInfo(dInfo, dInfo.iValue, Assign, MaxDepth:=0)
                            If fInfo.iTypeName = "ActionsCount" Then ' Select Case? if more with iValue
                                If LenB(dInfo.DecodedStringValue) > 0 Then
                                    If dInfo.iValue.Enabled Then
                                        dInfo.DecodedStringValue = "Aktiv:  " & dInfo.DecodedStringValue
                                    Else
                                        dInfo.DecodedStringValue = "Passiv: " & dInfo.DecodedStringValue
                                    End If
                                End If
                            End If
                            If LenB(dInfo.DecodedStringValue) > 0 Then
                                fInfo.DecodedStringValue = fInfo.DecodedStringValue & vbCrLf _
                                                         & dInfo.iTypeName & "(" & vElement & ")=" & dInfo.DecodedStringValue
                            End If
                        Next vElement
                    End If
                    Set dInfo = fInfo            ' restore qualified state, this is the final result
                    dInfo.DecodeMessage = "# " & dInfo.iTypeName
                    GoTo FuncExit
                End If
            Else                                 ' (user-)Property: may or may not have class
                .iClass = aValue.Class           ' .iAssignmentMode always=2 !
                aBugTxt = "Get Class of aValue"
                If Catch Then
                    .DecodeMessage = E_Active.Reasoning
                    GoTo FuncExit
                End If
            End If
                
            aBugTxt = "Set value from Object " & .iTypeName
            Set vValue = aValue.Value            ' Objects with Object value do not cause error
            If Catch Then
                .DecodeMessage = E_Active.Reasoning
                Select Case .iClass
                Case olItemProperty              ' Properties always have a Name
                    DoVerify False, _
                             "Problem with Set ItemProperty Value.Value to Variant: " _
                           & .iTypeName & b _
                           & .iTypeName & "->" & aValue.Name
                    .DecodedStringValue = "# unable to obtain value for " & aValue.Name
                    GoTo FuncExit
                Case Else
                    .DecodedStringValue = "** Class not implemented: " & .iClass _
                                        & " TypeName: " & .iTypeName
                    DoVerify False, CStr(.DecodedStringValue)
                    GoTo FuncExit
                End Select                       ' .iClass
            End If
                                
            If vValue Is Nothing Then
                .DecodeMessage = "Object with Null value"
                GoTo nonValue
            End If
                
            Set dInfo = .DrillDown(vValue)       ' note: .DecodedStringValue never changed on this level
            .DecodeMessage = "get down Value, depth=" & .iDepth
            Call getInfo(dInfo, vValue, Assign:=Assign)
            If dInfo.iAssignmentMode = 1 Then
                GoTo FuncExit                    ' it was fully decoded
            End If
            dInfo.iTypeName = .iTypeName & " (" & dInfo.iTypeName ' leaves open bracket!
            If dInfo.iIsArray Then
                dInfo.iTypeName = dInfo.iTypeName & " Array)" ' close bracket (1)
                GoTo nonValue
            End If
            If dInfo.iAssignmentMode = 0 Then
                dInfo.iTypeName = dInfo.iTypeName & " No String Rep.)" ' close bracket (2)
                GoTo RuleDetermined
            End If
                
            dInfo.iTypeName = .iTypeName & b & dInfo.iValue.Name & " Value="
            With dInfo
                .iClass = .iValue.Class
                aBugTxt = "get string value of .iValue"
                vValue = CStr(.iValue)           ' may be empty
                If Catch Then
                    .DecodeMessage = E_Active.Reasoning
                    .iTypeName = .iTypeName & "None)" ' close bracket (3)
                    .DecodedStringValue = "#* " & .DecodeMessage
                    GoTo didSet                  ' did not work
                Else
                    .iAssignmentMode = 1         ' worked, change from 2->1
                    .iTypeName = .iTypeName & vValue & ")" ' close bracket (4)
                    .DecodeMessage = .iTypeName
                    .DecodedStringValue = vValue
                    GoTo FuncExit
                End If
            End With                             ' dInfo
        Case vbEmpty                             '  0 Empty (not initialized)
        Case vbNull                              '  1 Null  (no valid data, Nothing)
        Case vbError                             ' 10 Fehlerwert
        Case vbVariant                           ' 12 Variant (only arrays of Variant values)
            DoVerify False, "arrays of variants should have been covered by getInfo ???"
            .iClass = aValue.Value.Class
        Case vbDataObject                        ' 13 Ein Datenzugriffsobjekt
            .iClass = aValue.Value.Class
        Case Else
            If .iIsArray Then                    ' array needing value
                Set .iValue = aValue
                GoTo likeObject
            End If
            .iTypeName = "# Unknown type " & CStr(.iType)
            .DecodeMessage = .iTypeName
            DoVerify False
        End Select                               ' .iType
        
        If Not Assign Then
            GoTo nonValue
        End If
            
        aBugTxt = "Assign scalar from " & .iTypeName
        vValue = aValue.Value                    ' try non/-scalar types
        If Not Catch Then
            dInfo.iAssignmentMode = 1
            dInfo.DecodedStringValue = CStr(vValue)
            GoTo FuncExit
        End If
        .DecodeMessage = E_Active.Reasoning
didSet:                                          ' scalar assign did not work, try last resort
        If dInfo.iType < vbInteger Then          ' Empty or Null: no further attempts
            .DecodeMessage = dInfo.iTypeName & " is Empty or has Null value"
            GoTo nonValue
        End If
        
        aBugTxt = "Assign object value from variant via Set " & .iTypeName
        Set vValue = aValue.Value                ' try with variant object
        If Catch Then
            dInfo.DecodeMessage = "not obtained from variant " & .iTypeName
            DoVerify False, "Problem assigning to Variant .iValue: " _
                         & .iTypeName & b _
                         & .iTypeName & "->" & aValue.Name
            .DecodedStringValue = "# unable to obtain value for " & aValue.Name
            .DecodeMessage = .DecodedStringValue
nonValue:
            DoVerify .iAssignmentMode <> 1, "only for analysis in design ???"
            Call AppendTo(testNonValueProperties, aTD.adName, sep:=b)
RuleDetermined:
            If DebugMode Then
                Debug.Print ">>>> Non-scalar Value for Property " _
                          & Quote(aTD.adName)
            End If
        Else
            .iAssignmentMode = 1
            .DecodeMessage = .iTypeName
        End If
    End With                                     ' fInfo

FuncExit:
    Set fInfo = dInfo                            ' this replaces fInfo with bottom Element!
    If fInfo.iAssignmentMode <> 1 And LenB(fInfo.DecodedStringValue) > 0 Then
        If Left(fInfo.DecodedStringValue, 1) <> "#" Then
            If DebugLogging And Left(fInfo.DecodeMessage, 1) <> "#" Then
                DoVerify False, "Complex value without Explanation"
            Else
                If DebugLogging Then
                    Debug.Print fInfo.DecodeMessage & " value=" & fInfo.DecodedStringValue
                End If
                fInfo.iAssignmentMode = 1        ' it is not necessary to DrillDown again
            End If
        End If
    End If
    Set dInfo = Nothing
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr, fInfo.DecodeMessage)

pExit:
End Sub                                          ' HelpersOL.getInfo

'---------------------------------------------------------------------------------------
' Method : GetPropertyByNumber
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: for the item specified, get Prop = Itemproperties(trueindex)
'---------------------------------------------------------------------------------------
Function GetPropertyByNumber(trueindex As Long, anyItem As Object) As ItemProperty
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "HelpersOL.GetPropertyByNumber"
    Call DoCall(zKey, "Function", eQzMode)
    ' *** bug escape
    ' not working: Set aProp = anyItem.ItemProperties(trueindex)
    ' direct assignment would sometimes cause Error 450, so we do it step by step:
Dim temp As ItemProperties

    Set temp = anyItem.ItemProperties
    If DoVerify(temp.Count >= trueindex, "the item has no Itemproperty at trueindex=" & trueindex) Then
        Set aProps = Nothing
    Else
        If Not aProps Is temp Then
            Set aProps = temp
        End If
    End If
    Set GetPropertyByNumber = temp(trueindex)
    Set temp = Nothing
    ' *** End bug escape

zExit:
    Call DoExit(zKey)
ProcRet:
End Function                                     ' HelpersOL.GetPropertyByNumber

'---------------------------------------------------------------------------------------
' Method : Function isNotDecodable
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function isNotDecodable(AttrName As String, Optional iRules As cAllNameRules) As Boolean
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "HelpersOL.isNotDecodable"
    Call DoCall(zKey, tFunction, eQzMode)

Dim aRules As cAllNameRules

    isNotDecodable = InStr(testNonValueProperties & b, AttrName & b) > 0
    If isNotDecodable Then
        GoTo zExit
    End If

    If iRules Is Nothing Then
        If aOD(aPindex) Is Nothing Then
            Set aRules = aOD(0).objClsRules
        Else
            Set aRules = aOD(aPindex).objClsRules
        End If
    Else
        Set aRules = iRules
    End If
    
    isNotDecodable = InStr(aRules.clsNotDecodable.aRuleString, AttrName) > 0
    
FuncExit:
    Set aRules = Nothing
    
zExit:
    Call DoExit(zKey)

End Function                                     ' HelpersOL.isNotDecodable

'---------------------------------------------------------------------------------------
' Method : MakeNotLogged
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Build a search folder for NotLogged Items (in active/specified ScopeFolders)
'---------------------------------------------------------------------------------------
Sub MakeNotLogged(Account As String, ScopeFolders As String, Optional TestEx As Boolean)

Const zKey As String = "TestMailAggregation.AddNotInternalSearchFolder"
    Call DoCall(zKey, tSub, eQzMode)

Dim i As Long
Dim Filter As String
Dim oFolder As Folder
Dim ScopeFoldersA As Variant

    Filter = "NOT Categories LIKE " & Quote1(LOGGED)
    If LenB(ScopeFolders) = 0 Then
        If LenB(Account) = 0 Then
            Set oFolder = ActiveExplorer.CurrentFolder
            ScopeFolders = "'" & oFolder.FolderPath & "'"
            Account = Trunc(3, oFolder.FolderPath, "\")
        Else
            ScopeFolders = "'\\" & Account & "\" & StdInboxFolder & "'"
            If TestEx Then
                oFolder = GetFolderByName(ScopeFolders) ' fails if not exists
            End If
        End If
    ElseIf TestEx Then
        ScopeFoldersA = split(Replace(ScopeFolders, "'", vbNullString), ",")
        ScopeFolders = vbNullString
        For i = 0 To UBound(ScopeFoldersA)
            oFolder = GetFolderByName(CStr(ScopeFoldersA(i)))
            ScopeFolders = ScopeFolders & "'" & oFolder.FolderPath & "',"
        Next i
        ScopeFolders = Left(ScopeFolders, Len(ScopeFolders) - 1)
    End If
    
    Call CreateSearchFolder(Account, ScopeFolders, Filter, _
            SpecialSearchFolderName & b & NLoggedName, WithSubFolders:=True)
    
    Set oFolder = Nothing

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.MakeNotLogged

'---------------------------------------------------------------------------------------
' Method : Sub ModItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ModItem(aktItem As Object, adName As String, nAttrName As String, processingOptions As Variant)
Dim zErr As cErr
Const zKey As String = "HelpersOL.ModItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim oldValue As Variant
Dim OtherValue As Variant
Dim i As Long
Dim N As Long
Dim iOther As Long
Dim nOther As Long
Dim nproperties As ItemProperties
Dim Removed_Comma As Boolean

    oldValue = vbNullString
    Set nproperties = aktItem.ItemProperties
    aPindex = 1                                  ' find in aID(1).idAttrDict
    i = FindAttributeByName(1, adName)
    N = FindAttributeByName(1, nAttrName)
    iOther = FindAttributeByName(1, adName)
    nOther = FindAttributeByName(1, nAttrName)
    If i > 0 And N > 0 And nOther > 0 And iOther > 0 Then ' property exists
        ' Fixing iOther = aID(1).idAttrDict.item(i).Index = iOther
        If iOther <> aID(1).idAttrDict.Item(i).DictIndex Then
            aID(1).idAttrDict.Item(i).DictIndex = iOther
        End If
        If aID(1).idAttrDict.Item(N).DictIndex <> nOther Then
            Call ShowAttrs(nproperties, 1)
        End If
        If aID(1).idAttrDict.Item(i).adName <> nproperties.Item(iOther).Name Then
            Call ShowAttrs(nproperties, 1)
        End If
        If aID(1).idAttrDict.Item(N).adName <> nproperties.Item(nOther).Name Then
            Call ShowAttrs(nproperties, 1)
        End If
        ' swap lastname and firstname, omitting ","
        oldValue = nproperties.Item(iOther).Value ' current firstname here
        OtherValue = nproperties.Item(nOther).Value ' current lastname
        If Right(oldValue, 1) = Right(processingOptions, 1) Then ' has a comma ending
            oldValue = Left(oldValue, Len(oldValue) - 1)
            Removed_Comma = True
        End If
        If Right(OtherValue, 1) = Right(processingOptions, 1) Then ' has a comma ending
            OtherValue = Left(OtherValue, Len(OtherValue) - 1)
            Removed_Comma = True
        End If
        If Removed_Comma Then
            nproperties.Item(iOther) = oldValue
            nproperties.Item(i) = OtherValue
            
            nproperties.Item("FullName") = OtherValue & b & oldValue
                                                
            aktItem.Save
        End If
    Else
        DoVerify False, " Prop not found by name"
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.ModItem

'---------------------------------------------------------------------------------------
' Method : Function SearchFolderExists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SearchFolderExists(StoreName As String, FolderName As String, Optional sFolder As Folder) As Boolean

Const zKey As String = "HelpersOL.SearchFolderExists"
    Call DoCall(zKey, tFunction, eQzMode)
     
Dim oStore As Outlook.Store

    Set oStore = olApp.Session.Stores.Item(StoreName)
    Set oSearchFolders = oStore.GetSearchFolders
    
    For Each sFolder In oSearchFolders
        If sFolder.Name = FolderName Then
            SearchFolderExists = True
            GoTo FuncExit
        End If
    Next                                         ' aStore
    
    Set sFolder = Nothing                        ' no match
    
FuncExit:
    Set oStore = Nothing

zExit:
    Call DoExit(zKey)

End Function                                     ' HelpersOL.SearchFolderExists

'---------------------------------------------------------------------------------------
' Method : Sub SetDebugMode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetDebugMode(Optional WithLog As Variant, Optional WithDebug As Variant, Optional noMsg As Boolean = False)
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "HelpersOL.SetDebugMode"
Dim zErr As cErr

Dim WithLogDefault As Variant
Dim WithDebugDefault As Variant
Dim DebugString As String

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)
    
    If IsMissing(WithLog) Then
        WithLog = False
            WithLogDefault = True
End If
If IsMissing(WithDebug) Then
    ' debugmode toggle
    WithDebugDefault = True
    End If

    If WithDebugDefault Then
        If DebugMode Then            ' toggle all to off
            Call SetDebugModes("off", noMsg:=noMsg)
            Call SetErrLogging(LogAllErrors)
        Else                         ' toggle on
            Call AppendTo(DebugString, "Debug", b)
        End If
    ElseIf WithDebug Then
        Call AppendTo(DebugString, "Debug", b)
    End If

    If WithLog = "Ask" Then
        If aNonModalForm Is Nothing Then
            ErrStatusFormUsable = True
            Call N_Suppress(Push, zKey)
            Call ShowDbgStatus("Choose debug options")
            Call N_Suppress(Pop, zKey)
        End If
    ElseIf WithLog Then
        Call AppendTo(DebugString, "Log", b)
    End If

    Call SetDebugModes(DebugString, noMsg:=noMsg)
    ' interactive setting of Debug options
    If ErrStatusFormUsable Then
        With aNonModalForm
            If LenB(Testvar) = 0 Then
                .ToggleDebug.Caption = "currently all off"
            Else
                .ToggleDebug.Caption = Testvar
            End If
            If ErrorCaught <> 0 Then
                DoVerify DebugMode
                GoTo pExit
            End If
            .Show
            If DebugMode Then
                .ToggleDebug.BackColor = &H80FFFF
            Else
                .ToggleDebug.BackColor = &H8000000F
            End If
        End With                     ' aNonModalForm
    End If

pExit:
    Call ProcExit(zErr)

End Sub                              ' HelpersOL.SetDebugMode

'---------------------------------------------------------------------------------------
' Method : Sub SetDebugModes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetDebugModes(Optional more As String, Optional volltest As Boolean = True, Optional noMsg As Boolean = True)

Const zKey As String = "HelpersOL.SetDebugModes"
    Call DoCall(zKey, tSub, eQzMode)

Dim msg As String

    If volltest Then
        more = "Debug Log"
    ElseIf InStr(LCase(more), "off") > 0 Then
        more = vbNullString
    End If
    
    Call SetEnvironmentVariable("Test", Trim(more))
    Call getDebugMode(True)
    If DebugMode Then
        msg = msg & vbCrLf & "debugmode ist AN"
        MinimalLogging = 1
    Else
        msg = msg & vbCrLf & "debugmode ist AUS"
        MinimalLogging = 3
    End If
    If DebugLogging Then
        msg = msg & vbCrLf & "debuglogging ist AN"
        MinimalLogging = 1
    Else
        msg = msg & vbCrLf & "debuglogging ist AUS"
    End If
    If Not noMsg Then
        MsgBox msg
    End If
    Debug.Print msg
    Debug.Print "Testvar=" & Quote(Testvar)

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.SetDebugModes

'---------------------------------------------------------------------------------------
' Method : Sub SetErrLogging
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetErrLogging(Optional eLall As Boolean = True)

Const zKey As String = "HelpersOL.SetErrLogging"
    Call DoCall(zKey, tSub, eQzMode)

    Call getDebugMode
    LogAllErrors = eLall
    Testvar = Trim(Remove(Testvar, "ERR"))
    If eLall Then
        Testvar = "ERR " & Testvar
    End If
    Call SetEnvironmentVariable("Test", Trim(Testvar))
    Debug.Print "Testvar=" & Quote(Testvar)

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.SetErrLogging

'---------------------------------------------------------------------------------------
' Method : Sub SetLogMode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetLogMode()

Const zKey As String = "HelpersOL.SetLogMode"
    Call DoCall(zKey, tSub, eQzMode)

    If DebugLogging Then
        Call SetDebugModes("noLog", False, noMsg:=False)
    Else
        Call SetDebugModes("Log", False, noMsg:=False)
    End If

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.SetLogMode

'---------------------------------------------------------------------------------------
' Method : Sub SetTrapMode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetTrapMode(Optional onoff As Boolean = True)

Const zKey As String = "HelpersOL.SetTrapMode"
    Call DoCall(zKey, tSub, eQzMode)
    
    Call getDebugMode
    If onoff And InStr(1, Testvar, "Trap", vbTextCompare) > 0 Then
        onoff = False
    End If
    Testvar = Trim(Remove(Testvar, "Trap"))
    If onoff Then
        Testvar = "Trap " / Testvar
    End If
    Call SetEnvironmentVariable("Test", Trim(Testvar))
    Debug.Print "Testvar=" & Quote(Testvar)

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.SetTrapMode

'---------------------------------------------------------------------------------------
' Method : Sub ShowAttrs
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowAttrs(nproperties As ItemProperties, StackIndex As Long)
Dim zErr As cErr
Const zKey As String = "HelpersOL.ShowAttrs"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim hi As Long
    DoVerify aID(StackIndex).idAttrDict.Count <= nproperties.Count, "more decoded properties than itemproperties"
    For i = 1 To aID(StackIndex).idAttrDict.Count
        hi = aID(StackIndex).idAttrDict.Item(i).DictIndex
        Debug.Print i, aID(StackIndex).idAttrDict.Item(i).adName = _
                                                                 nproperties.Item(hi).Name, _
                                                                 aID(StackIndex).idAttrDict.Item(i).adName, _
                                                                 hi, nproperties.Item(hi).Name
    Next i

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.ShowAttrs

'---------------------------------------------------------------------------------------
' Method : Sub ShowCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowCats(MyStoreOrFolder As Variant)
Dim zErr As cErr
Const zKey As String = "HelpersOL.ShowCats"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim targetCat As category
Dim MyStore As Store
Dim msg As String

    Select Case MyStoreOrFolder.Class
    Case olFolder
        Set MyStore = MyStoreOrFolder.Store
        msg = "Dumping Categories for Store " & Quote(MyStore.DisplayName) _
      & " corresponding to folder " & Quote(MyStoreOrFolder.FolderPath)
    Case olStore
        Set MyStore = MyStoreOrFolder
        msg = "Dumping Categories for Store" & Quote(MyStoreOrFolder.DisplayName)
    Case Else
        DoVerify False, " does not make sense"
    End Select
    If MyStore.Categories.Count = 0 Then
        msg = "There are no Categories in Store" & Quote(MyStoreOrFolder.DisplayName)
    End If
    If DebugMode Then
        Debug.Print "=======================================================" & vbCrLf _
                  & msg
    End If
    For i = 1 To MyStore.Categories.Count
        Set targetCat = MyStore.Categories.Item(i)
        Debug.Print i & vbTab & "Category " & Quote(targetCat.Name) _
      & vbTab & " ShortcutKey " & targetCat.ShortcutKey _
      & vbTab & " Color " & targetCat.Color
    Next i

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' HelpersOL.ShowCats

'---------------------------------------------------------------------------------------
' Method : Sub showDebugSettings
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub showDebugSettings()

Const zKey As String = "HelpersOL.showDebugSettings"
    Call DoCall(zKey, tSub, eQzMode)

Dim Test As String
    Test = GetEnvironmentVar("Test")
    If Test <> Testvar Then
        Test = Testvar & "?" & Test
    End If
    Debug.Print "Testvar=" & Quote(Test), DebugMode, DebugLogging

zExit:
    Call DoExit(zKey)

End Sub                                          ' HelpersOL.showDebugSettings

'---------------------------------------------------------------------------------------
' Method : Sub ShowSearchAttrs
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowSearchAttrs()

Dim aSearches As RDOSearches
Dim aSearch As RDOSearch
Dim aStore As RDOStore
Dim StoreName As String
Dim i As Long
    
    aRDOSession.MAPIOBJECT = olApp.Session.MAPIOBJECT
    For Each aStore In aRDOSession.Stores
        Debug.Print "-------------------------------------"
        StoreName = LString(aStore.Name, lDbgM)
        If (aStore.StoreKind = skPstAnsi) Or (aStore.StoreKind = skPstUnicode) Then
            Debug.Print StoreName & " - " & aStore.PstPath
        ElseIf (aStore.StoreKind = skIMAP4) Then
            Debug.Print StoreName & " - " & aStore.OstPath
        ElseIf (aStore.StoreKind = skPrimaryExchangeMailbox) Or (aStore.StoreKind = skDelegateExchangeMailbox) Or (aStore.StoreKind = skPublicFolders) Then
            Debug.Print StoreName & " - " & aStore.OstPath& & " - " & aStore.StoreAccount.CurrentUser.Name
        Else
            Debug.Print StoreName & " - " & "unknown Store kind=" & aStore.StoreKind
            DoVerify False
        End If
        Set aSearches = aStore.Searches
        For Each aSearch In aSearches
            Debug.Print "-------------"
            Debug.Print aSearch.Name & ": "
            For i = 1 To aSearch.SearchContainers.Count
                Debug.Print String(10, b) & aSearch.SearchContainers.Item(i).Name
            Next i
            Debug.Print aSearch.SearchCriteria.AsSQL
        Next
    Next
    
    '    Set aSearches = aRDOSession.Stores.DefaultStore.Searches

End Sub                                          ' HelpersOL.ShowSearchAttrs

'---------------------------------------------------------------------------------------
' Method : Function StandardTime
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function StandardTime(itemProp As ItemProperty, tD As Date, Optional useUTC As Boolean = False) As String
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "HelpersOL.StandardTime"
    Call DoCall(zKey, "Function", eQzMode)

    ' must use with string-formatted time, but also in UTC time zone
    If useUTC Then
    
Dim oPA As Outlook.PropertyAccessor

        Set oPA = itemProp.Item.PropertyAccessor
        StandardTime = oPA.LocalTimeToUTC(tD)
        If DebugMode Then
            Debug.Print tD & " -> " & StandardTime & " (UTC)"
        End If
        Set oPA = Nothing
    Else
        StandardTime = tD
    End If

zExit:
    Call DoExit(zKey)

End Function                                     ' HelpersOL.StandardTime

'---------------------------------------------------------------------------------------
' Method : Function StoreCatGet
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function StoreCatGet(MyStore As Store, catName As String, Optional Position) As category
Dim zErr As cErr
Const zKey As String = "HelpersOL.StoreCatGet"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim thisCat As category
Dim i As Long
    
    aBugTxt = "Get Categories from Item"
    Call Try(testAll)                               ' Try anything, autocatch
    Set StoreCatGet = MyStore.Categories.Item(catName)
    Catch
    If IsMissing(Position) Then                  ' no position index (out)
        Set StoreCatGet = MyStore.Categories.Item(catName)
    Else
        For i = 1 To MyStore.Categories.Count
            Set thisCat = MyStore.Categories.Item(i)
            If thisCat.Name = catName Then
                Position = i
                Set StoreCatGet = thisCat
                GoTo FuncExit
            End If
        Next i
        ' StoreCatGet is nothing when catName not found
    End If

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

ProcRet:
End Function                                     ' HelpersOL.StoreCatGet

Sub AlleEmails()    ' ??under test
Dim objFolder As Folder
Dim sFolder As Folder
Dim oItem As Object
Dim mytItems As Items
Dim tFolder As Folder
Dim strFolder As String
Dim strFilter As String
Dim myExplorers As Explorers
Dim myPlorer As Explorer
    Stop ' set nlogged Hotmail
  Set myExplorers = Application.Explorers
    Debug.Print "Count explorers:"; myExplorers.Count
  Set sFolder = Application.ActiveExplorer.CurrentFolder
    Stop  ' switch folder to target = AggregatedInbox
  Set tFolder = Application.ActiveExplorer.CurrentFolder
  Set myPlorer = tFolder.GetExplorer(olFolderDisplayNormal)
    myPlorer.SelectAllItems
  Set mytItems = tFolder.Items
  Set oItem = mytItems.GetFirst
    While Not oItem Is Nothing
        ' delete target emails
    Set oItem = mytItems.GetNext
    Wend
  Set oItem = sFolder.Items.GetFirst
    While Not oItem Is Nothing
        Call CopyItemTo(oItem, tFolder)
    Set oItem = sFolder.Items.GetNext
    Wend

   ' objFolder = myExplorers.Add(objFolder, olFolderDisplayNormal)
   ' Set myPlorer = myExplorers.Item(2)
   ' myPlorer.Display
   ' myPlorer.Activate
Stop
  Set objFolder = _
    Application.ActiveExplorer.CurrentFolder
    
  If objFolder.DefaultItemType <> olMailItem Then
    MsgBox _
      Prompt:="Die Aktion kann im aktuellen " & _
        "Ordner nicht ausgeführt werden." & _
        String(2, vbCrLf) & _
        "Wechseln Sie bitte erst in einen " & _
        "E-Mail-Ordner.", _
      Buttons:=vbExclamation, _
      Title:="Alle E-Mails anzeigen"
    Exit Sub
  End If

  strFolder = Chr(34) & "Posteingang" & Chr(34)
  strFilter = "ordnerpfad:(" & strFolder & ")"

  objFolder.GetExplorer.Search _
    strFilter, _
    olSearchScopeAllFolders

  Set objFolder = Nothing
End Sub


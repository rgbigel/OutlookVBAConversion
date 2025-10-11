Attribute VB_Name = "MailProcessing"
Option Explicit

Dim ProblemHelp As String                        ' benenne was faul ist
Dim ArchiveFolder As Outlook.Folder
Dim ArchiveSubFolder As Outlook.Folder
Dim ConfirmOperation As Boolean
Dim Abbruch As Boolean
Dim Filter As String
Dim ItemCounter As Long
Dim thisTopFolder As String

' Achtung, hier must Du vorgeben, wie es funktionieren soll!
Const ArchivDateiName As String = "my arch"
Const ArchivierungLöschtOriginal As Boolean = True ' False geht vermutlich nicht... braucht Redemption.DLL
' anzahl Tage vor heute (-30) als CutOffDate
Const MaximalAlter As Long = -30

' Das ArchiveByDate aufrufen, z.B in ThisOutlookSession.Application_Startup, dann gehts bei jedem Outlook-Start

'---------------------------------------------------------------------------------------
' Method : Sub ArchiveByDate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ArchiveByDate()
Dim zErr As cErr
Const zKey As String = "MailProcessing.ArchiveByDate"
    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

Dim InBoxFolder As Outlook.Folder
Dim MainFolder As Outlook.Folder

    On Error GoTo aBug
    ConfirmOperation = True                      ' fragt ob OK vor Archivierung
    Abbruch = False
    ItemCounter = 0
    thisTopFolder = StdInboxFolder
    
    ProblemHelp = "Datumsfehler"
    CutOffDate = DateAdd("d", MaximalAlter, Now())
    ' ============
    
    ProblemHelp = "ArchiveFolder bestimmen"
    ' der Name DEINES Archiv - .PSD könnte anders sein, anpassen !!!
    Set ArchiveFolder = GetFolderByName(ArchivDateiName, aNameSpace, MaxDepth:=1)
    ' ===============
    ' Annahme: es gibt nur ein Konto, also nur einen Posteingang
    If thisTopFolder = StdInboxFolder Then
        ProblemHelp = "Inbox Ordner bestimmen"
        Set InBoxFolder = olApp.Session.GetDefaultFolder(olFolderInbox)
        Call doArchiveWork(InBoxFolder)
        ' ==========================
        If Abbruch Then GoTo aBug
    Else
        ' Falls Annahme falsch, muss eine Schleife über die InBoxen erfolgen:
        For Each MainFolder In aNameSpace.Folders
            Set InBoxFolder = GetFolderByName(thisTopFolder, MainFolder)
            If Not InBoxFolder Is Nothing Then
                ' Archiv nicht archivieren !
                If InStr(1, InBoxFolder.FolderPath, ArchivDateiName, vbTextCompare) = 0 Then
                    Call doArchiveWork(InBoxFolder)
                    ' ==========================
                    If Abbruch Then GoTo aBug
                Else
                    DoVerify False, " debugphase only"
                End If
            End If
        Next MainFolder
    End If

    ' Feddisch
    ProblemHelp = "Es wurden insgesamt " & CStr(ItemCounter) & " Objekte archiviert"
    Debug.Print ProblemHelp
    GoTo ProcReturn
aBug:
    Debug.Print ProblemHelp
    Debug.Print Err.Description
    DoVerify False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.ArchiveByDate

'---------------------------------------------------------------------------------------
' Method : Sub doArchiveWork
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub doArchiveWork(InBoxFolder As Outlook.Folder)
Dim zErr As cErr
Const zKey As String = "MailProcessing.doArchiveWork"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    On Error GoTo dBug
    Set topFolder = InBoxFolder.Parent
    
    ' auf gehts mit der Archivierung
    Set ArchiveSubFolder = GetCorrespondingFolder(InBoxFolder, ArchiveFolder)
    Call DoAllItemsIn(InBoxFolder, ArchiveSubFolder)
    If Abbruch Then GoTo dBug
    ' Unterordner
    Call ArchiveSubFolders(InBoxFolder, ArchiveFolder) ' now process subfolders of InBoxFolder
    If Abbruch Then GoTo dBug
    GoTo ProcReturn
dBug:
    Debug.Print ProblemHelp
    Abbruch = True
    DoVerify False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.doArchiveWork

'---------------------------------------------------------------------------------------
' Method : Sub DoAllItemsIn
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DoAllItemsIn(actFolder As Outlook.Folder, actArchiveFolder As Outlook.Folder)
Dim zErr As cErr
Const zKey As String = "MailProcessing.DoAllItemsIn"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim ItemsHere As Items
Dim thisitem As Variant
Dim thisItemID As String
Dim thisArchiveFolder As Outlook.Folder

    Set thisArchiveFolder = actArchiveFolder     ' protect against changes
    
    ProblemHelp = "archivieren des Ordners '" & actFolder.FolderPath & "'" _
                & " in den Ordner '" & thisArchiveFolder.FolderPath & "'"
    Set ItemsHere = RestrictItemsByDate(actFolder, Filter, "<=") ' using CutOffDate
    If ItemsHere Is Nothing Then
        ProblemHelp = "Auswahl der items mit Filter " & Filter _
                    & " in '" & actFolder.FolderPath & "' fehlgeschlagen"
        Debug.Print ProblemHelp
    Else
        If actFolder.Items.Count = ItemsHere.Count Then
            ProblemHelp = "Auswahl der items mit Filter " & Filter _
                        & " in '" & actFolder.FolderPath _
                        & "' umfasst alle " & ItemsHere.Count _
                        & " Objekte, ist das richtig???"
            If MsgBox(ProblemHelp, vbOKCancel) = vbCancel Then
                ProblemHelp = " Operation abgebrochen"
                GoTo ProcReturn
            End If
        Else
            ProblemHelp = "Auswahl der items mit Filter " & Filter _
                        & " in '" & actFolder.FolderPath & "'" & vbCrLf _
                        & " umfasst " & ItemsHere.Count _
                        & " zu archivierende Objekte, " _
                        & vbCrLf & "ist das plausibel?" _
                        & vbCrLf & "(Weiterhin bestätigen: Ja, Nein: diesen Ordner auslassen, Cancel: Abbruch)"
            rsp = MsgBox(ProblemHelp, vbYesNoCancel)
            If rsp = vbCancel Then
                ProblemHelp = " Operation abgebrochen"
                GoTo ProcReturn
            ElseIf rsp = vbNo Then
                Abbruch = True
                GoTo ProcReturn
            Else
                ProblemHelp = "Archivierung umfasst " & ItemsHere.Count & " Objekte, Kriterien " _
                            & Filter & " aus '" & actFolder.FolderPath
                Debug.Print ProblemHelp
            End If
        End If
    End If
    For Each thisitem In ItemsHere
        If Not thisitem.Saved Then
            If thisItemID = thisitem.EntryID Then
                ProblemHelp = " Duplicate Item " & thisItemID
                GoTo SkipIt                      ' double entry...
            End If
            thisItemID = thisitem.EntryID
            aBugTxt = "Save Item " & thisItemID
            Call Try(testAll)
            thisitem.Save
            If Catch Then
                ProblemHelp = E_Active.Reasoning & ": " _
                            & E_Active.Description
                aBugTxt = "Fetch Item "
                Call Try                         ' Try anything, autocatch
                Set thisitem = aNameSpace.GetItemFromID(thisItemID)
                If Catch Then
                    Call LogEvent("Item existiert nicht mehr, " _
                                & "evtl. gelöscht durch Regel, Virenchecker o.ä.")
                    ProblemHelp = ProblemHelp & E_Active.Reasoning
                    GoTo SkipIt
                End If
            End If
        End If
        ProblemHelp = vbNullString
        aBugTxt = "Copy Item to " & thisArchiveFolder.FolderPath
        Call Try
        Call CopyItemTo(thisitem, thisArchiveFolder)
        Catch
SkipIt:
    Next thisitem
    
    ProblemHelp = "Archivierung beendet, " & ItemsHere.Count & " Objekte, Zielordner '" _
                & actFolder.FolderPath & "'"
    Debug.Print ProblemHelp

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.DoAllItemsIn

' Recursive Sub to DoAllItemsIn Subfolders
Sub ArchiveSubFolders(actFolder As Outlook.Folder, actArchiveFolder As Outlook.Folder)
Dim zErr As cErr
Const zKey As String = "MailProcessing.ArchiveSubFolders"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim FolderIndex As Long
Dim loopFolder As Outlook.Folder
Dim actFolderPath As String
Dim EntryProblem As String

    If actFolder.Folders.Count = 0 Then
        GoTo ProcReturn
    End If
    EntryProblem = ProblemHelp
    
    On Error GoTo boo
    actFolderPath = actFolder.FolderPath
    For FolderIndex = 1 To actFolder.Folders.Count
        ProblemHelp = " alle Unterordner von '" & actFolderPath & "' bearbeiten"
        Set loopFolder = actFolder.Folders.Item(FolderIndex)
        
        ArchiveSubFolder = GetCorrespondingFolder(loopFolder, actArchiveFolder)
        If Abbruch Then GoTo boo
        
        ' First, do the items here
        Call DoAllItemsIn(loopFolder, ArchiveSubFolder)
        ' then, recurse into local subfolders (if any)
        Call ArchiveSubFolders(loopFolder, ArchiveSubFolder)
        
    Next FolderIndex
    ProblemHelp = EntryProblem
    GoTo ProcReturn
boo:
    Debug.Print ProblemHelp
    DoVerify False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.ArchiveSubFolders

'---------------------------------------------------------------------------------------
' Method : Function GetCorrespondingFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetCorrespondingFolder(actFolder As Outlook.Folder, actArchiveFolder As Outlook.Folder) As Outlook.Folder
Dim zErr As cErr
Const zKey As String = "MailProcessing.GetCorrespondingFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim actFolderPath As String
Dim ArcFolderPath As String
Dim ArcFolderName As String
    
    On Error GoTo bug
    actFolderPath = actFolder.FolderPath
    ' parse this name and replace the front part to the archivefolder
    ArcFolderName = RTail(actFolderPath, "\", Front:=ArcFolderPath)
    ArcFolderPath = Replace(actFolderPath, ArcFolderPath, actArchiveFolder.FolderPath)
    
    ProblemHelp = " korrespondierenden ArchiveFolder festlegen für '" & actFolderPath & "'"
    Set GetCorrespondingFolder = GetFolderByName(ArcFolderName, actArchiveFolder)
    If GetCorrespondingFolder Is Nothing Then
        ProblemHelp = " korrespondierenden ArchiveFolder existiert nicht: '" _
                    & ArcFolderPath & "' wird erstellt"
        Set GetCorrespondingFolder = actArchiveFolder.Folders.Add(actFolder.Name)
    Else
        If GetCorrespondingFolder.FolderPath <> ArcFolderPath Then
            ProblemHelp = " korrespondierenden ArchiveFolder '" _
                        & GetCorrespondingFolder.FolderPath & "' bzw.'" & ArcFolderPath _
                        & "' passen nicht zu '" & actFolderPath & "'"
            GoTo bug
        End If
    End If
    GoTo ProcReturn
bug:
    Abbruch = True

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.GetCorrespondingFolder

'---------------------------------------------------------------------------------------
' Method : Sub CopyItemTo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyItemTo(ByVal myItem As Object, TargetFolder As Outlook.Folder)
Dim zErr As cErr
Const zKey As String = "MailProcessing.CopyItemTo"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim myCopiedItem As Object
Dim FilterName As String
Dim AttrValue As String
Dim itemProps As ItemProperties
Dim itemProp As ItemProperty

    On Error GoTo nixGutt
    Set myCopiedItem = olApp.CreateItem(olMailItem)
    If Len(Filter) < 4 Then
    Else
        FilterName = Mid(Filter, 2, InStr(Filter, "]") - 2)
        Set itemProps = myItem.ItemProperties
        Set itemProp = itemProps(FilterName)
        AttrValue = itemProp.Value
        If IsDate(AttrValue) Then
            AttrValue = Format(itemProp.Value, "dd.mm.yyyy hh:mm:ss")
        End If
        ProblemHelp = "  Verschieben (" & FilterName & b & AttrValue & ") '" _
                    & myItem.Subject & "'"
        If ArchivierungLöschtOriginal Then
            Set myCopiedItem = myItem
        Else
            ' copy item in the same folder as original.
            Set myCopiedItem = myItem.Copy           ' it will not work for sources in Exchange Active Sync ???
            myItem.Delete                            ' delete Original ???
            Set myItem = Nothing
            Set aItmDsc.idObjItem = myCopiedItem
        End If
        
        ' move this copy to TargetFolder (which also deletes Original if ArchivierungLöschtOriginal
        myCopiedItem.Move TargetFolder
    End If
    If Not myCopiedItem Is Nothing Then
        If Not myCopiedItem.Saved Then
            ProblemHelp = "  Save Item"
            Call Try                             ' Try anything, autocatch
            myCopiedItem.Save
            If Catch Then
                ProblemHelp = E_Active.Reasoning & ": " & E_Active.Description
                Debug.Print ProblemHelp
                Set myCopiedItem = Nothing
                GoTo fixed
            End If
        End If
    End If
    ItemCounter = ItemCounter + 1
    ProblemHelp = CStr(ItemCounter) & ProblemHelp _
                                  & " nach '" & myCopiedItem.Parent.FolderPath _
                                  & "' OK"
    Debug.Print ProblemHelp
    Set myCopiedItem = Nothing
    GoTo fixed
nixGutt:
    Debug.Print ProblemHelp
    Debug.Print Err.Description
    DoVerify False
fixed:
    ProblemHelp = vbNullString

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.CopyItemTo

'---------------------------------------------------------------------------------------
' Method : Function RestrictItemsByDate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RestrictItemsByDate(ByRef curFolder As Folder, Filter As String, Optional Comparator As String = "<=") As Items
Dim zErr As cErr
Const zKey As String = "MailProcessing.RestrictItemsByDate"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If getFolderFilter(curFolder.Items(1), CutOffDate, Filter, Comparator) Then
        On Error GoTo invalid
        Set RestrictItemsByDate = curFolder.Items.Restrict(Filter)
    Else
invalid:
        Set RestrictItemsByDate = Nothing
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.RestrictItemsByDate

'---------------------------------------------------------------------------------------
' Method : Function getFolderFilter
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getFolderFilter(A As Object, compareTo As Variant, Filter As String, Comparator As String) As Boolean
Dim zErr As cErr
Const zKey As String = "MailProcessing.getFolderFilter"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim filterValue As String
    Filter = vbNullString                                  ' we will not filter
    If IsDate(compareTo) Then
        If compareTo < CDate(BadDate) And compareTo <> "00:00:00" Then
            filterValue = Quote1(Format(CStr(compareTo), "yyyy mm dd"))
        Else
            getFolderFilter = False
            GoTo ProcReturn
        End If
    Else
        DoVerify False, " not implemented"
    End If
    
    Select Case A.Class
    Case olMail
        Filter = "[SentOn]"
        getFolderFilter = True
    Case olAppointment
        Filter = "[ReceivedTime]"
        getFolderFilter = True
    Case olReport
        Filter = "[CreationTime]"
        getFolderFilter = True
    Case Else
        DoVerify False, " class not implemented"
        getFolderFilter = False
        GoTo ProcReturn
    End Select
    
    aTimeFilter = Replace(Replace(Filter, "[", vbNullString), "]", vbNullString)
    Filter = Filter & b & Comparator & b & filterValue & b

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.getFolderFilter

'---------------------------------------------------------------------------------------
' Method : Function Sender2Contact
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Sender2Contact(EmailAddr As String) As Items
Const zKey As String = "MailProcessing.Sender2Contact"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

Dim EmailMatch As String
Dim ContactSearch As String
Dim i As Long

    ContactSearch = vbNullString
    For i = 1 To 3
        EmailMatch = "[Email" & i & "Address] = '" & EmailAddr & "'"
        If i < 3 Then
            ContactSearch = ContactSearch & EmailMatch & " OR "
        Else
            ContactSearch = ContactSearch & EmailMatch
        End If
    Next i
    
    aBugTxt = "Restrict matching contacts for " & ContactSearch
    Call Try
    Set Sender2Contact = ContactFolder.Items.Restrict(ContactSearch)
    If Catch Then
        GoTo FuncExit
    End If
    If DebugMode Then
        If Sender2Contact.Count = 0 Then
            Debug.Print "there are no Contacts matching email sender " _
                      & Quote(EmailAddr)
        Else
            Debug.Print "there are " & Sender2Contact.Count _
                      & " Contacts matching email sender " & Quote(EmailAddr)
            For i = 1 To Sender2Contact.Count
                Debug.Print vbTab & i & vbTab & Sender2Contact.Item(i).Subject
            Next i
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.Sender2Contact

'---------------------------------------------------------------------------------------
' Method : Function IsUnkContact
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsUnkContact(EmailAddr As String) As Boolean
Const zKey As String = "MailProcessing.IsUnkContact"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

Dim myMatchingContacts As Items

    Set myMatchingContacts = Sender2Contact(EmailAddr)
    If myMatchingContacts.Count = 0 Then
        IsUnkContact = True
    End If
    Set myMatchingContacts = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.IsUnkContact

' Wenn dem entsprechenden Ordner ein Item hinzugefügt wird,
' wird eine der MailProcessing Subs ausgeführt:
' Hier: Anhänge speichern, ggf. Kategorie setzen, Reminder eintragen
' Regeln werden in RuleWizard definiert und ausgewertet

'---------------------------------------------------------------------------------------
' Method : Sub CollectItemsToLog
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Collect any (Mail-) items not in LOGGED state
' Note   : only up to DeferredLimit added at one time
'          when all folders were done before, specificIndex :=1
'---------------------------------------------------------------------------------------
Sub CollectItemsToLog(specificIndex As Long)
Const zKey As String = "MailProcessing.CollectItemsToLog"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")

Const NotLoggedItems = "@SQL=NOT ""urn:schemas-microsoft-com:office:office#Keywords"" LIKE "

Dim afolder As Folder
Dim curObj As Object
Dim LogFolderIndex As Long
Dim initialFolderIndex As Long
    
    If DeferredLimit = 0 Then
        ' limit number of items currently not processed because we could have storage problems
        DeferredLimit = maxDeferredLimit
        DeferredLimitExceeded = False
    End If
    RestrictCriteriaString = NotLoggedItems & Quote1(LOGGED)
    ' check against bounds
    If specificIndex > LoggableFolders.Count Or specificIndex < 0 Then
        specificIndex = 0               ' we finished previous LoggableFolders, item 1 = items(0)
        initialFolderIndex = -1         ' require do all LoggableFolders again
    End If      ' else we stay with this folder initialFolderIndex = 0
    
    For LogFolderIndex = specificIndex To LoggableFolders.Count - 1
        Set afolder = LoggableFolders.Items(LogFolderIndex)
        curFolderPath = afolder.FolderPath
        Set RestrictedItems = Nothing
        aBugTxt = "Restrict folder # " & LogFolderIndex & b & curFolderPath
        Call Try(allowNew)                          ' Try anything, autocatch
        Set RestrictedItems = afolder.Items.Restrict(RestrictCriteriaString)
        Catch
        ItemsToDoCount = RestrictedItems.Count
        Set curObj = RestrictedItems.GetFirst

gotCurObject:
        If curObj Is Nothing Then
            specificIndex = LogFolderIndex      ' resume with next folder, if any
            Call LogEvent("* " & ItemsToDoCount _
                & " Items collected in " & "(" & specificIndex & ") " _
                & Quote(curFolderPath), eLall)
            GoTo NextFolder
        End If
        Call DeferredActionAdd(curObj, atPostEingangsbearbeitungdurchführen, NoChecking:=True)
        If DeferredLimitExceeded Then
            specificIndex = LogFolderIndex       ' resume with next folder, if any
            GoTo finishLater
        End If
        Set curObj = RestrictedItems.GetNext
        GoTo gotCurObject
NextFolder:
        If initialFolderIndex = 0 Then          ' doing just one folder
            specificIndex = -1                  ' do all again next time
            GoTo finishLater                    ' but do not loop further
        End If
    Next LogFolderIndex
    
finishLater:
    StopRecursionNonLogged = True

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.CollectItemsToLog

'---------------------------------------------------------------------------------------
' Method : Sub FldActions2Do
' Author : rgbig
' Date   : 20211108@11_47
' Purpose: do action on not LOGGED mail-like items, all Folders in LoggableFolders
'---------------------------------------------------------------------------------------
Sub FldActions2Do()
Dim zErr As cErr
Const zKey As String = "MailProcessing.FldActions2Do"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                        & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcRet
    End If
    Recursive = True

    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

    ' List all items in the Inbox that do NOT have a flag:
Dim sActionId As Long
Dim skipIndex As Long
Dim doneCtr As Long
Dim doneTotal As Long

    ' possibly removed Unbekannt nach Trash über Regel dieses Namens
    If StopRecursionNonLogged Then
        GoTo ProcReturn                 ' === prohibited just now===
    End If
    If MailEventsViaRules Then
        Call ExecuteDefinedRule(inFolderName:="unerwünscht", RuleName:="unerwünscht")
    End If
    
    If ActionID <> atFindealleDeferredSuchordner Then
        sActionId = ActionID            ' save for interruptions
    End If
    ActionID = 0                        ' no user choice
    skipIndex = -1                      ' start with first LoggableFolders
                                        ' and do all of them
    LF_UsrRqAtionId = atFindealleDeferredSuchordner ' this action (only)
    
    EventHappened = False
    NoEventOnAddItem = True             ' defer all new mail events until we are done
    ItemsToDoCount = 0
    frmErrStatus.lblDeferredCount.Caption = "Deferred Count"
resumeCollect:
    loopItmIndex = 0
    Do
        loopItmIndex = loopItmIndex + 1
        DeferredLimitExceeded = False
        Call CollectItemsToLog(skipIndex)   ' accumulate Deferred from all LoggableFolders
        StopRecursionNonLogged = True
        TotalDeferred = Deferred.Count
        frmErrStatus.fDeferredCount = Deferred.Count & "/" & doneTotal
        frmErrStatus.lblDeferredCount.Caption = "Doing"
        Call DoAllDeferred
        doneCtr = TotalDeferred
        doneTotal = doneTotal + doneCtr
        If DeferredLimitExceeded Then
            Exit Do
        End If
        skipIndex = skipIndex + 1
        If LoggableFolders.Count < skipIndex Then
            Exit Do
        End If
    Loop
    frmErrStatus.fDeferredCount = TotalDeferred & "/" & doneTotal
    If doneCtr > 0 And doneCtr Mod DeferredLimit = 0 Then
        Debug.Print LString("  continuing after Exceeding Deferred Limit", OffObj)
        skipIndex = -1              ' start with first LoggableFolders, for all
        GoTo resumeCollect
    End If
    
    ' do not finish operation here yet if StopRecursionNonLogged = False
    LF_UsrRqAtionId = sActionId             ' completed
    ActionID = sActionId
    SkipedEventsCounter = 0
    loopItmIndex = 0
    If doneTotal <> 0 Then
        frmErrStatus.lblDeferredCount.Caption = "Total done"
        TotalDeferred = doneTotal
    End If

FuncExit:
    Call LogEvent("* FldActions2Do processed " & doneTotal & " Items", eLall)
    EventHappened = False

ProcReturn:
    Call ProcExit(zErr)
    Recursive = False

ProcRet:
End Sub                                     ' MailProcessing.FldActions2Do

'---------------------------------------------------------------------------------------
' Method : Sub DoAllDeferred
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DoAllDeferred()
Const zKey As String = "MailProcessing.DoAllDeferred"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQArMode, CallType:=tSub, ExplainS:="MailProcessing")

Dim oItem As Object
Dim ID As String
Dim PrevId As String
Dim ThisSubject As String

    TotalDeferred = 0
    aPindex = 1                                  ' always!
    quickChecksOnly = True
    If Deferred.Count > 0 Then
        Call N_ShowProgress(CallNr, zErr.atDsc, zErr.atKey, _
                            "#Deferred=" & Deferred.Count, _
                            "Limit#=" & DeferredLimit)
    Else
        Call N_ShowProgress(CallNr, zErr.atDsc, zErr.atKey, _
                            "no deferred actions", vbNullString)
    End If
    ItemsToDoCount = Deferred.Count
    
    While Deferred.Count > 0                     ' Process all that were deferred first
        TotalDeferred = TotalDeferred + 1
        PrevId = ID
        ID = Deferred.Item(1).aoObjID
        On Error GoTo gone
        Call Try(testAll)
        Set oItem = aNameSpace.GetItemFromID(ID)
        If oItem Is Nothing Then
            GoTo gone
        End If
        ThisSubject = oItem.Subject
        curFolderPath = oItem.Parent.FolderPath
        On Error GoTo 0
        aBugVer = PrevId <> ID
        DoVerify aBugVer, "TotalDeferred hacking on same item ??? ID=" & ID
        
        If aID(aPindex) Is Nothing Then
makeAnother:
            Set aID(aPindex) = New cItmDsc       ' uses NewObjectItem
            Call aID(aPindex).SetDscValues(oItem, withValues:=False)
        Else
            If aOD(aPindex).objItemClass <> oItem.Class Then
                GoTo makeAnother
            End If
        End If
        
        Call DefObjDescriptors(oItem, aPindex, _
                               withValues:=False, _
                               withAttributeSetup:=False)
        If oItem Is Nothing Then
gone:
            Call LogEvent("**** Item #" & TotalDeferred & " ID=" & ID & b _
                        & ThisSubject _
                        & " nicht mehr vorhanden, es verbleiben " _
                        & ItemsToDoCount, eLall)
            ItemsToDoCount = ItemsToDoCount - 1
        ElseIf aObjDsc.objIsMailLike Then
            Call LogEvent("==== Processing #" _
                        & Deferred.Count & "/" & ItemsToDoCount _
                        & " deferred Item " & TotalDeferred _
                        & " in '" & curFolderPath _
                        & "' ID=" & ID & vbCrLf & String(5, b) _
                        & aObjDsc.objTimeType & "=" _
                        & aItmDsc.idTimeValue & b & oItem.Subject, eLall)
            ActionID = atPostEingangsbearbeitungdurchführen
            
            Call CopyToWithRDO(oItem, FolderAggregatedInbox, aObjDsc)
            Call DoOneItm(oItem)
            Call N_ClearAppErr
        Else
            Call LogEvent("**** Skipping #" & TotalDeferred & " ID=" & ID _
                        & " of " & ItemsToDoCount _
                        & " deferred Items because it is not mail-like: " _
                        & aObjDsc.objTypeName)
            DoVerify Not DebugMode, "** analyze mail-like Attribute "
            ItemsToDoCount = ItemsToDoCount - 1
        End If
        If Deferred.Count > 0 Then
            Deferred.Remove 1
            If ItemsToDoCount > 0 Then
                ItemsToDoCount = ItemsToDoCount - 1
            End If
            frmErrStatus.fDeferredCount = Deferred.Count & "/" & ItemsToDoCount
        End If
    Wend
    DoVerify Deferred.Count = 0, " all done"
    If TotalDeferred > 0 Then
        Call LogEvent(TotalDeferred & " neue " _
                    & Quote(SpecialSearchFolderName & b & NLoggedName) _
                    & "  Mail-Eingänge verarbeitet", eLall)
    Else
        Call LogEvent(" keine " & Quote(SpecialSearchFolderName _
                    & b & NLoggedName) _
                    & "  Mail-Eingänge verarbeitet", eLall)
    End If
    Set oItem = Nothing

FuncExit:
    quickChecksOnly = False

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.DoAllDeferred

'---------------------------------------------------------------------------------------
' Method : Function IsMailLike
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsMailLike(nlItem As Object) As Boolean
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "MailProcessing.IsMailLike"
Dim zErr As cErr

    Set CurrentSessionEmail = Nothing
    Set CurrentSessionReport = Nothing
    Set CurrentSessionMeetRQ = Nothing
    Set CurrentSessionTaskRQ = Nothing

    Select Case (nlItem.Class)
    Case olMail
        Set CurrentSessionEmail = nlItem
        IsMailLike = True
    Case olReport
        Set CurrentSessionReport = nlItem
        IsMailLike = True
    Case olMeeting
        Set CurrentSessionMeetRQ = nlItem
        IsMailLike = True
    Case olTaskRequest
        Set CurrentSessionTaskRQ = nlItem
        IsMailLike = True
    Case Else
        If (nlItem.Class >= olMeetingRequest _
            And nlItem.Class <= olMeetingResponseTentative) _
        Then ' range class: 53-57
            Set CurrentSessionMeetRQ = nlItem
            IsMailLike = True
        Else
            DoVerify False, "can't map to CurrentSession Class"
            IsMailLike = False
        End If
    End Select
 
ProcReturn:
End Function                                     ' MailProcessing.IsMailLike

'---------------------------------------------------------------------------------------
' Method : Sub DoOneItm
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DoOneItm(nlItem As Object)
Const zKey As String = "MailProcessing.DoOneItm"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="DoOneItm")

Dim pFolder As Folder

    Set aID(aPindex).idAttrDict = New Dictionary ' invalidate previous decoding results
    Set aDecProp(aPindex) = Nothing
    If aObjDsc.objIsMailLike Then
        ' if the item is reported to be in the search Folder,
        ' it has no pointer to parent.parent
        ' so we need to find where it really is
        Set pFolder = getParentFolder(nlItem)
        If pFolder Is Nothing Then
            Call LogEvent("??? a deferred item to process is no longer available", eLSome)
        Else
            ItemInIMAPFolder = getAccountType(pFolder.FolderPath, aAccountTypeName) = olImap
            If InStr(nlItem.Parent.FullFolderPath, "Suchordner") > 0 Then
                Set nlItem = aNameSpace.GetItemFromID(nlItem.EntryID)
                If InStr(nlItem.Parent.FullFolderPath, "Suchordner") > 0 Then
                    DoVerify False, " that should never happen"
                End If
            End If
            Set ParentFolder = nlItem.Parent     ' now !never! = Nothing
            Call DoMailLike(nlItem)
        End If
    Else
        If DebugMode Or DebugLogging Then
            Call LogEvent("---- no log categories will be assigned for object of type " _
                        & TypeName(nlItem), eLall)
            DoVerify False
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.DoOneItm

'---------------------------------------------------------------------------------------
' Method : Sub DoMailLike
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DoMailLike(oItem As Object)
Dim zErr As cErr
Const zKey As String = "MailProcessing.DoMailLike"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aSubject As String
Dim rsp As VbMsgBoxResult
Dim NameOrMsg As String
Dim SourceFolder As Folder
Dim SourceFolderPath As String
Dim saveXldefer As Boolean
Dim TargetFolder As Folder
Dim TargetFolderPath As String
Dim otherChanges As Boolean
Dim DontMove As Boolean
Dim IsInTarget As Boolean
Dim Multiple As Long
Dim PushEventMode As Boolean
Dim nItem As Object
Dim asNonLoopFolder As Boolean
Dim oID As String
Dim nID As String
Dim UnReadString As String
Dim oldItemCategories As String
Dim ReProcessCategories As Boolean
Dim DelAfterCopy As Boolean
Dim i As Long
    
    aNewCat = vbNullString
    aSubject = oItem.Subject
    
    If Not (LF_DontAskAgain Or oItem.Parent Is Nothing) Then
        asNonLoopFolder = NonLoopFolder(oItem.Parent.Name)
        If asNonLoopFolder Then
            Call LogEvent("<======> skipping item action " & Quote(ActionTitle(ActionID)) _
                        & vbCrLf & " because it is inFolder " & Quote(oItem.Parent.FolderPath) _
                        & vbCrLf & " loop item " & WorkIndex(1) _
                        & " Time: " & Now(), eLall)
            GoTo ProcReturn                      ' this is an item we should not be looping
        End If
    End If
    
    PushEventMode = NoEventOnAddItem
    
    MailModified = False
    saveXldefer = xDeferExcel
    xDeferExcel = xUseExcel
    Call N_ClearAppErr
    
    Set oItem = ReGet(oItem, oID, nID)
    If oItem Is Nothing Then
        GoTo flushBrake
    End If
    oldItemCategories = oItem.Categories
    If DebugMode Then
        If oID = nID Then
            Debug.Print "     the item re-get did not change EntryID=" & oID
            Debug.Print "     Item categories: " & Quote(oldItemCategories)
        Else
            Debug.Print "     the item re-get changed EntryID=" & oID
            Debug.Print "     Item categories: " & Quote(oldItemCategories)
            Debug.Print "                       to    EntryID=" & nID
            Debug.Print "     Item categories: " & Quote(nItem.Categories)
            DoVerify False
        End If
    End If
    
    Set SourceFolder = getParentFolder(oItem)
    If SourceFolder Is Nothing Then
        DoVerify False
        GoTo flushBrake
    End If
    For i = 1 To 5                               ' intermittent bug escape
        SourceFolderPath = SourceFolder.FolderPath
        If LenB(SourceFolderPath) > 0 Then
            Exit For
        End If
    Next i
    DoVerify LenB(SourceFolderPath) > 0, "error getting FolderPath, tries=" & i - 1
        
    If InStr(1, SourceFolderPath, "SUCHORDNER\", vbTextCompare) > 0 Then
        DoVerify False                           ' can't find true Folder for search Folder??? ***"
    End If
    aBugTxt = "get body of item"
    Call Try                                     ' Try anything, autocatch
    NameOrMsg = oItem.Body                       ' check if item still exists
    If Catch Then
        Call LogEvent("Item kann nicht mehr gefunden werden")
        GoTo flushBrake
    End If
    
    Call LogEvent("==== " & Time() & b & TypeName(oItem) & b & oItem.Subject _
                & " in " & SourceFolderPath _
                & "    (" & Deferred.Count & " Unbearbeitete in Suchordner " _
                & Quote(SpecialSearchFolderName) & ")", eLall)
            
    If InStr(1, oldItemCategories, LOGGED, vbTextCompare) = 0 Then
        Call LogEvent("---- NotLogged " & TypeName(oItem) & ": " _
                    & oItem.Subject, eLall)
    ElseIf CurIterationSwitches.ReProcessDontAsk Then
        GoTo Auto
    ElseIf CurIterationSwitches.ReprocessLOGGEDItems Then
        If Not CurIterationSwitches.ReProcessDontAsk Then
            rsp = MsgBox("re-process this item? " & vbCrLf & aSubject _
                       & " with Categories: " _
                       & Quote(oldItemCategories), vbYesNo, "Bestätigung")
            If rsp <> vbNo Then
                Call LogEvent("---- This item IS reprocessed by user choice: " & aSubject _
                            & ", Categories: " & Quote(oldItemCategories), eLall)
                ReProcessCategories = True
            Else
                GoTo dontProcess
            End If
        Else
Auto:
            Call LogEvent("---- Automatically re-processed: " & aSubject, eLall)
            GoTo doprocess
        End If
        If eOnlySelectedItems Then
            ReProcessCategories = True
            GoTo doprocess
        End If
    Else
dontProcess:
        Call LogEvent("---- This item not reprocessed: " _
                    & aSubject _
                    & ", Categories: " & Quote(oldItemCategories), eLall)
                    
        GoTo Epilog
    End If
    
doprocess:
    LF_ItmChgCount = LF_ItmChgCount + 1
    
    ' Preconditions for Move to Target:
    ' PosteINgang, ERHALTEN, INbox, SMS (cond.), GES,  UNB, UNK,  but NOT SENt
    If InStr(1, SourceFolderPath, "SEN", vbTextCompare) > 0 Then
        Set TargetFolder = FolderSent
        GoTo Sent
    
    ElseIf InStr(1, SourceFolderPath, "IN", vbTextCompare) > 0 Then
        Set TargetFolder = FolderInbox
        GoTo Received
    
    ElseIf InStr(1, SourceFolderPath, "SMS", vbTextCompare) > 0 Then
        Set TargetFolder = FolderSMS             ' Phone number not visible
        GoTo Received
    
    ElseIf InStr(1, SourceFolderPath, "UNB", vbTextCompare) > 0 Then
        Set TargetFolder = FolderUnknown
        GoTo Received
    
    ElseIf InStr(1, SourceFolderPath, "UNK", vbTextCompare) > 0 Then
        Set TargetFolder = FolderUnknown
        GoTo Received
    
    ElseIf InStr(1, SourceFolderPath, "ERHALTEN", vbTextCompare) > 0 Then
        Set TargetFolder = FolderInbox
        GoTo Received
    End If
    If TargetFolder Is Nothing Then
        Set TargetFolder = FolderInbox
    End If
    If Not (LF_DontAskAgain Or isNonLoopFolder) Then
        DoVerify False
    End If
Received:
    If oItem.Class = olReport Then               ' Cases: may not have SenderName
        Call LogEvent("    -- " & oItem.Body, eLall)
    ElseIf oItem.Class = olTaskRequest Then
        Call LogEvent("    -- Antwortstatus NewMail", eLall)
    Else                                         ' should have SenderName
        Call LogEvent("    -- from " & Quote(oItem.SenderName) _
                    & " (" & oItem.SenderEmailAddress & ")" _
                    & vbCrLf & "     -- received " & oItem.ReceivedTime _
                    & ", sent on " & oItem.SentOn _
                    & ", created on " & oItem.CreationTime, eLall)
    End If
    GoTo DoChanges
Sent:
    If oItem.Class = olMail Then
        Call LogEvent("    -- Sent to " & oItem.To, eLall)
    ElseIf oItem.Class = olReport Then
        
    ElseIf oItem.Class >= olMeetingRequest _
        And oItem.Class <= olMeetingResponseTentative Then
            Call LogEvent("    Associated Meeting with " _
                    & oItem.Recipients.Count & " recipients", eLall)
    Else
        aBugTxt = "get attachment count " & NameOrMsg
        Call Try
        NameOrMsg = TypeName(oItem)
        Catch
        NameOrMsg = NameOrMsg & " has " & oItem.Attachments.Count & " attachments"
    End If
    
DoChanges:
    NameOrMsg = vbNullString
    CategoryDroplist = LOGGED & "; Aktuell; "       ' always dropped
    aNewCat = DetectCategory(TargetFolder, oItem, NameOrMsg) ' may change TargetFolder
    '   (eg. if sender is unknown)!
    
    Call setItmCats(oItem, aNewCat, CategoryDroplist)
    If aNewCat <> oldItemCategories Then
        If DebugMode Then
            Debug.Print "about to change the Item's Categories to " & Quote(aNewCat)
        End If
        If FolderUnknown Is Nothing Then
            'Stop 'checkme
        Else
            If InStr(aNewCat, Unbekannt) = 0 Then
                If TargetFolder.FolderPath = FolderUnknown.FolderPath Then
                    Set TargetFolder = FolderInbox ' special for items with contact now known:
                    DelAfterCopy = True          ' copy to FolderInbox then remove from Unknown
                ElseIf SourceFolder.FolderPath = FolderUnknown.FolderPath Then
                    Set TargetFolder = FolderInbox ' special for items with contact now known:
                    DelAfterCopy = True          ' copy to FolderInbox then remove from Unknown
                End If
            End If
        End If
    End If
    TargetFolderPath = TargetFolder.FullFolderPath
    
    If TargetFolderPath = SourceFolderPath Then
        DoVerify Not otherChanges, " OtherChanges remains False"
        DontMove = True
        IsInTarget = True
    Else
        DontMove = False
    End If
    
    Set nItem = CopyItm2Trg(oItem, _
                            TargetFolderPath, _
                            SourceFolderPath, _
                            TargetFolder)
    If DelAfterCopy Then
        oItem.Delete
        Call LogEvent("     original item deleted from folder " & SourceFolderPath, eLall)
        Set aItmDsc.idObjItem = Nothing
        Set oItem = Nothing
    ElseIf oItem.Categories <> nItem.Categories Then
        oItem.Categories = nItem.Categories
        oItem.Save
    End If
    If RestrictedItemCollection.Count > 1 Then
        Multiple = RedceRestrColl(DontMove, _
                                  otherChanges, _
                                  RemoveEmailItems:=True)
        Set nItem = RestrictedItemCollection(1)
    End If
    Set RestrictedItemCollection = New Collection ' un-use the remaining Item
    Call DoChng2Item(nItem, TargetFolder)
    Call GenerateTaskReminder(nItem)             ' Erinnerung erstellen (ggf.)
    
flushBrake:
    NameOrMsg = Err.Description
    If ErrorCaught = Hell Then
        GoTo ProcReturn
    End If
    If CatchNC(HandleErr:=-2147221233) Then
        If nItem Is Nothing Then
            GoTo Epilog
        End If
        If InStr(NameOrMsg, "kann nicht gefunden werden") = 0 Then
            rsp = MsgBox(NameOrMsg, vbOKOnly, TypeName(nItem) & ": " _
                       & Quote(nItem.Subject))
        End If
        Call LogEvent("**** Fehler: " & NameOrMsg & b _
                    & TypeName(nItem) & ": " & Quote(nItem.Subject), eLall)
    End If
    
    ' Original gets same changes as nItem
    If InStr(UCase(nItem.Parent.Name), "SEN") = 0 Then
        nItem.UnRead = True                      ' may or may not cause change
        UnReadString = " UnRead"
    Else
        nItem.UnRead = False                     ' usually causes a change
        UnReadString = " Read"
    End If
    If oItem Is Nothing Then
        Set aItmDsc.idObjItem = nItem            ' aObjdDsc of oItem can not be cloned, replace
        UnReadString = " original item moved to " & Quote(TargetFolderPath) & UnReadString
    Else
        If Not oItem.Saved Then                  ' should be saved before
            Call ForceSave(oItem, "(old) ")
        End If
    End If
    Call LogEvent("     " & Quote(SourceFolderPath) & UnReadString _
                & " + Categories set to " _
                & Quote(nItem.Categories), eLall)
    If Not nItem.Saved Then                      ' max have been changed if categories changed
        nItem.Save
    End If
    If Not nItem.Saved Then                      ' catch problem with previous Save attempt
        Call ForceSave(nItem, "(new) ")
    End If
Epilog:
    Call N_ClearAppErr
    Set RestrictedItemCollection = New Collection
    xDeferExcel = saveXldefer
    If Not PushEventMode Then
        Call RestEvn4Item
    End If
    NoEventOnAddItem = PushEventMode
    Set oItem = Nothing
    Set nItem = Nothing

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.DoMailLike

'---------------------------------------------------------------------------------------
' Method : Function Replicate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Replicate(Item As Object, Optional delOriginal = False) As Object
Const zKey As String = "MailProcessing.Replicate"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

Dim SourceFolder As Folder
Dim NewRef As Object
Dim oItemID As String
Dim nItemID As String

    Call SetEventMode                            ' MASK OUT AddItemEvent inner processing
    ' no test on NoEventOnAddItem, always replicate
    
    ' copy and forward, do not work sometimes (???)
    If DebugLogging Then DoVerify False
    oItemID = Item.EntryID
    Set Item = Nothing
    
    aBugTxt = "get item from ID " & oItemID
    Call Try                                     ' Try anything, autocatch
    Set NewRef = aNameSpace.GetItemFromID(oItemID)
    If Catch Then
        Set Replicate = Nothing
        GoTo skipThis
    End If
    
    Set SourceFolder = getParentFolder(NewRef)
    Set Replicate = CopyToWithRedemption(NewRef, SourceFolder, trySave:=False, NewObjDsc:=aObjDsc)
    '               ====================
    
    If isEmpty(Replicate) Or Replicate Is Nothing Then
        Set Replicate = Nothing
    Else
        If Replicate.Subject <> Item.Subject Then
            Replicate.Subject = Item.Subject     ' no FWD: or WG:
            Replicate.Body = Item.Body
            Replicate.HTMLBody = Item.HTMLBody
        End If
        Call ShowStatusUpdate
    End If
    If delOriginal And Not Replicate Is Nothing Then
        nItemID = Replicate.EntryID
        If oItemID = nItemID Then DoVerify False, "shit"
        Set Item = Nothing
        Set NewRef = Nothing
        
        aBugTxt = "Replicate item"
        Call Try("Die angegebene Nachricht kann nicht gefunden werden.")
        Set NewRef = aNameSpace.GetItemFromID(oItemID)
        If Catch Then
            Call LogEvent("Delete original not needed because item already gone")
            GoTo skipThis
        End If
        If Not NewRef.Saved Then
            DoVerify False, " shit"
            Call ShowDbgStatus                   ' if pending error, set up frmErrStatus
            GoTo skipThis
        End If
        aBugTxt = "delete original item " & Quote(NewRef.Subject)
        Call Try
        NewRef.Delete
        Catch
    End If
    
skipThis:
    Set NewRef = Nothing
    Call RestEvn4Item                            ' AddItemEvent was not triggered by MASK OUT

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.Replicate

'---------------------------------------------------------------------------------------
' Method : Function CopyItm2Trg
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Function CopyItm2Trg(Item As Object, TargetFolderPath As String, SourceFolderPath As String, TargetFolder As Folder, Optional withDupeCheck As Double = -1#) As Object
Const zKey As String = "MailProcessing.CopyItm2Trg"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    ' withDupeCheck = 0: dont check for dupes
    '               = +1: check exact criteria before copyTo
    '               else  check with criteria in time window
 
Dim DontMoveItem As Boolean
Dim otherChanges As Boolean
Dim IsInTarget As Boolean
Dim Multiple As Long
Dim oldClass As OlObjectClass

    If TargetFolder Is Nothing Then
        DoVerify False
    End If
    LogicTrace = "*"                             ' start a new logic trace here
   
    If Item Is Nothing Then
        GoTo ProcReturn
    End If
    
    DontMoveItem = MailEventsViaRules            ' changed later if not in TargetFolder
    
    If ItemInIMAPFolder Then
        MailModified = False                     ' item.saved irrelevant for Imap Items
    Else
        MailModified = Not Item.Saved
    End If
    oldClass = Item.Class
    
    Set CopyItm2Trg = Item
    If TargetFolderPath = SourceFolderPath Then
        DoVerify Not otherChanges, " OtherChanges remains False"
        DontMoveItem = True
        IsInTarget = True
    Else
        DontMoveItem = False
    End If
    
    ' withDupeCheck = 1: do duplicateChecking before save attempt
        '              = -1: do duplicateChecking when save is needed (further below)
            If withDupeCheck = 1# Then           ' check if we already have exactly same mail in TargetFolder
                Call findUniqueEmailItems(CopyItm2Trg, TargetFolder, _
                                          GetFirstOnly:=False, _
                                          howmany:=Multiple, _
                                          maxTimeDiff:=withDupeCheck)
            End If
    
            ' potential reasons for saving item before changing anything
            LogicTrace = "Modified=" & MailModified _
                       & " otherchanges=" & otherChanges _
                       & " dontmoveitem=" & DontMoveItem _
                       & " saved=" & CopyItm2Trg.Saved _
                       & vbCrLf
            If MailModified _
            Or otherChanges _
            Or Not DontMoveItem _
            Or Not CopyItm2Trg.Saved Then
                Set CopyItm2Trg = CopyToWithRedemption(CopyItm2Trg, TargetFolder, True, CopiedObjDsc)
                '                 ====================
                MailModified = Not CopyItm2Trg.Saved
                If MailModified Then             ' $$$ impossible
                    DoVerify False
                    CopyItm2Trg.Save
                    MailModified = Not CopyItm2Trg.Saved
                    DoVerify Not MailModified
                End If
                If CopyItm2Trg.Parent.FolderPath = TargetFolder.FolderPath Then
                    IsInTarget = True
                End If
            End If
            LogicTrace = LogicTrace & " IsInTarget=" & IsInTarget & vbCrLf
        
            If withDupeCheck <> 0 Then
                If withDupeCheck <> 1# Then      ' check if we already have exactly same mail in TargetFolder
                    Call findUniqueEmailItems(CopyItm2Trg, TargetFolder, _
                                              GetFirstOnly:=False, _
                                              howmany:=Multiple, _
                                              maxTimeDiff:=withDupeCheck)
                End If
                If RestrictedItemCollection.Count >= 1 Then ' in target or some multiple duplicates left
                    If CopyItm2Trg.EntryID = RestrictedItemCollection(1).EntryID Then
                        ' no reordering
                    Else
                        Set CopyItm2Trg = RestrictedItemCollection(1)
                    End If
                    DontMoveItem = True          ' it was successfully copied to target before
                ElseIf RestrictedItemCollection.Count < 1 Then ' not in target or some multiple duplicates left?!
                    DoVerify withDupeCheck = -1# Or Not DebugMode, " $$$ early debug only"
                End If
            End If
    
            LogicTrace = LogicTrace _
                       & " dontmoveitem=" & DontMoveItem _
                       & " IsInTarget=" & IsInTarget & vbCrLf
            If Not (DontMoveItem Or IsInTarget) Then ' Here we do the CopyTo
                Set CopyItm2Trg = CopyToWithRedemption(CopyItm2Trg, TargetFolder, trySave:=False, NewObjDsc:=CopiedObjDsc)
                '                         ====================
            End If
    
CopyToFinished:
doMarkDone:
            If Not CopyItm2Trg.Saved Then        ' strange if not
                Call ForceSave(CopyItm2Trg)
            End If
            MailModified = False
            If CopyItm2Trg.Class <> oldClass Then
                DoVerify False
            End If

FuncExit:

ProcReturn:
            Call ProcExit(zErr)

pExit:
    End Function                             ' MailProcessing.CopyItm2Trg

'---------------------------------------------------------------------------------------
' Method : Function RedceRestrColl
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RedceRestrColl(DontMoveItem As Boolean, otherChanges As Boolean, Optional RemoveEmailItems As Boolean = False) As Long
Const zKey As String = "MailProcessing.RedceRestrColl"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

Dim i As Long
Dim pFolder As String

    If RestrictedItemCollection.Count > 0 Then   ' is a multiple
        DontMoveItem = True                      ' no need to move, it's there already
        If RestrictedItemCollection.Count > 1 Then ' too many multiple:
            Call Try(allowAll)                      ' Try anything, autocatch, Err.Clear
            For i = RestrictedItemCollection.Count To 2 Step -1 ' delete all except first already existing
                pFolder = Quote(RestrictedItemCollection(i).Parent.FullFolderPath)
                RestrictedItemCollection(i).Delete
                Catch
                If RemoveEmailItems Then
                    RestrictedItemCollection.Remove i ' remove email in target Folder
                    Call LogEvent("     " & i & ": sufficiently similar item found, deleted " _
                                & "from Collection and " & pFolder, eLall)
                Else
                    Call LogEvent("     " & i & ": sufficiently similar item found, deleted " _
                                & "from Collection", eLall)
                End If
            Next i
        End If
        DontMoveItem = True                      ' no need to move, it's there already
    Else
        DontMoveItem = False                     ' move because rule did not move it
        otherChanges = True
    End If
    RedceRestrColl = RestrictedItemCollection.Count

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.RedceRestrColl

'---------------------------------------------------------------------------------------
' Method : Function DoChng2Item
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DoChng2Item(Item As Object, TargetFolder As Folder) As VbMsgBoxResult
Dim zErr As cErr
Const zKey As String = "MailProcessing.DoChng2Item"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)
   
    ' Beginning changes on the Target
Dim iAttachCnt As Long
Dim j As Long
Dim k As Long
Dim JFirst As Long                           ' First Attachment corrected by deletes
Dim mailAttachments As Variant
Dim thisAttachment As Attachment
Dim Absender As String
Dim attachmentFile As String
Dim HeaderFirst As Boolean
Dim InsertText As String                     ' double byte character set for HTML
Dim bodyHTMLtext As String
Dim OriginalIsNotHTML As Boolean
Dim TargetShouldChangeToHTML As Boolean
Dim DoNotDeleteAttachment As Boolean         ' ??? currently not used
Static DontCareFileNames As Variant
    
Dim iClass As Long

    ' init defaults
    DoChng2Item = vbOK
    If isEmpty(DontCareFileNames) Then           ' One time only:
        DontCareFileNames = split(IgnoredAttachmentNames, b)
    End If
       
    ' find out if we need to/can change to HTML
    aBugTxt = "Get Class of Item"
    Call Try
    iClass = Item.Class
    If Catch Then
        DoVerify False, "if no hit, remove permit and testing ???"
    End If
    If iClass = olMail Then
        aBugTxt = "Get folder path of Item"
        Call Try
        If InStr(UCase(Item.Parent.FullFolderPath), "SMS") = 0 Then
            j = -1                               ' JUST for errors:
            j = Len(Item.HTMLBody) > 0 Or Len(Item.Body) > 0
            ' we have some sort of body if j < 0
            OriginalIsNotHTML = (Item.BodyFormat <> olFormatHTML) And j < 0
            If OriginalIsNotHTML Then
                TargetShouldChangeToHTML = True  ' if allowed, do it
            Else
                TargetShouldChangeToHTML = False ' no change needed here
            End If
        End If
        Catch
    Else
        OriginalIsNotHTML = True                 ' only Mails can be Html
        TargetShouldChangeToHTML = False
    End If
    
    ' do change to format
    If TargetShouldChangeToHTML _
    And SetAllToHTML Then
        aBugTxt = "set body format to HTML"
        Call Try
        Item.BodyFormat = olFormatHTML
        Catch
        MailModified = True                      ' Relevant mod done
    End If
    
    If iClass = olMail Then
        If InStr(1, Item.ReceivedByName, "Marc", vbTextCompare) > 0 Then
            If Len(Item.HTMLBody) > 1000 Then
                Debug.Assert False
            End If
            bodyHTMLtext = Item.HTMLBody
            
            Absender = Item.SenderName
            If InStr(Absender, "@") > 0 Then
                If InStr(Item.SenderEmailAddress, "@") > 0 Then
                    Absender = Item.SenderEmailAddress
                End If
                Absender = Mid(Absender, 1, InStr(Absender, "@") - 1)
            End If
            If InStr(Absender, " PH/DE") > 0 Then ' (Marc)
                Absender = Replace(Absender, " PH/DE", vbNullString)
            End If
            If InStr(Absender, "/") > 0 Then     ' (Marc)
                Absender = Replace(Absender, "/", vbNullString)
            End If
        End If                                   ' special for Marc
    End If
    
    HeaderFirst = True
    aBugTxt = "get mail attachments"
    Call Try
    Set mailAttachments = Item.Attachments
    Catch
    
    If mailAttachments Is Nothing Then
        iAttachCnt = 0
    Else
        iAttachCnt = mailAttachments.Count
    End If
    Call LogEvent("     " & TypeName(Item) _
                & " has " & iAttachCnt & " attachments", eLall)
    
    JFirst = 1
    For j = 1 To iAttachCnt
        attachmentFile = "     Anhang #" & JFirst & " wurde entfernt (Virus?) und " _
                       & "konnte nicht gespeichert werden."
        ' attachment with virus may have been deleted by now...
        aBugTxt = "get mail attachment " & JFirst
        Call Try
        Set thisAttachment = mailAttachments.Item(JFirst) ' always store first remaining attachment
        If Catch Then
            Call N_ClearAppErr
            GoTo AttachmentDone
        End If
        If thisAttachment.Type = olEmbeddeditem Then
            Call LogEvent("     not trying to save Embeddet Attachment " & JFirst _
                        & b & Quote(thisAttachment.FileName), eLall)
            JFirst = JFirst + 1
            GoTo AttachmentDone
        End If
        
        k = ArrayMatch(DontCareFileNames, thisAttachment.FileName)
        DoNotDeleteAttachment = k > -1
        If Item.Class = olMail Then
            DoVerify thisAttachment.Type <> 0
            Call Try
            attachmentFile = thisAttachment.FileName
            If Catch(True, "unable to access file name for attachment " & JFirst & b & Item.Name) Then
                JFirst = JFirst + 1
                GoTo AttachmentDone
            End If
            
            If InStr(UCase(Item.HTMLBody), UCase(attachmentFile)) > 0 Then
                If Not OriginalIsNotHTML Then
                    If Not SaveAttachmentMode Then
                        Call LogEvent("     Attachment no. " & JFirst & _
                                      " NOT saved as attachment " & attachmentFile _
                                    & " because it is part of HTML body")
                        JFirst = JFirst + 1
                        GoTo AttachmentDone
                    End If
                End If
            End If
        End If
        If DoNotDeleteAttachment Then
            Call LogEvent("      irrelevant attachment name " & JFirst & ": " _
                        & thisAttachment.FileName, eLmin)
            JFirst = JFirst + 1
            GoTo AttachmentDone
        Else
            attachmentFile = aPfad & TargetFolder.Name _
                           & "\" & DateId & b & _
                             ReFormat(Absender, ".\/?*", b, b) _
                           & " - " & thisAttachment.FileName
            
            aBugTxt = "Save mail attachment " & JFirst & " to " & attachmentFile
            Call Try
            thisAttachment.SaveAsFile attachmentFile
            If Catch Then
                Call LogEvent("     Fehler beim Speichern des MailAttachments position " & j & _
                              Err.Description, eLall)
                Call N_ClearAppErr
                GoTo AttachmentDone
            End If
            Call LogEvent("      saved attachment " & JFirst _
                        & " as " & attachmentFile, eLSome)
        End If
                
        If CopyOriginal And Not ItemInIMAPFolder And Not DoNotDeleteAttachment Then
            If DelSavedAttachments Then
                aBugTxt = "delete mail attachment " & JFirst
                Call Try
                thisAttachment.Delete
                If Catch Then
                    Call N_ClearAppErr
                    GoTo AttachmentDone
                End If
            End If
        Else
            JFirst = JFirst + 1
        End If
        If HeaderFirst Then                      ' log announcement needed?
            HeaderFirst = False
            InsertText = "<p>" & "Extrahierte Anhänge: " & iAttachCnt
        End If
        If Item.Class = olMail Then
            If Item.BodyFormat = olFormatHTML Then
                InsertText = InsertText & "<br>" & j & ": " & "<A HREF=""" & _
                             attachmentFile & """>" & attachmentFile
            Else
                InsertText = InsertText & "<br>" & j & ": " & attachmentFile
            End If
        Else
            InsertText = InsertText & "<br>" & j & ": " & attachmentFile
        End If
AttachmentDone:
    Next j                                       ' looping attachments
    
    If Item.Class = olMail Then
        If LenB(InsertText) = 0 Then
            If TargetShouldChangeToHTML Then
                Call LogEvent("     Email converted to HTML in " _
                            & TargetFolder.FullFolderPath, eLall)
            Else
                Call LogEvent("     Email is HTML already in " _
                            & TargetFolder.FullFolderPath)
            End If
        Else
            bodyHTMLtext = Replace(Item.HTMLBody, "</body>", InsertText & _
                                                            "</BODY>", 1, 1, vbTextCompare)
            Call LogEvent("     Target item converted to HTML " _
                        & "and attachment references inserted ", eLnothing)
        End If
        If Item.HTMLBody <> bodyHTMLtext And LenB(bodyHTMLtext) > 0 Then
            Item.HTMLBody = bodyHTMLtext
            MailModified = True
        End If
    End If
    
    If Catch Then
        ' rsp = msgbox(Err.Description, vbOKOnly, TypeName(item) & ": " & item.subject)
        Call LogEvent("**** Fehler: " & Err.Description & b & TypeName(Item) & ": " _
                    & Item.Subject)
    ElseIf DoChng2Item = 0 Then
        DoChng2Item = vbOK
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.DoChng2Item

'---------------------------------------------------------------------------------------
' Method : Sub addContactPic
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub addContactPic(cItem As Object, attThisContact As Attachment)
Dim zErr As cErr
Const zKey As String = "MailProcessing.addContactPic"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim ContactPicture As String
    ContactPicture = cPfad & ReFormat(cItem.FileAs, ".\/?*", b, b) _
      & " (" & cItem.Parent.Parent.Name & " - " _
      & Replace(DateId, b, vbNullString) _
      & ").jpg"
    aBugTxt = "save contact's picture in " & ContactPicture
    Call Try
    attThisContact.SaveAsFile ContactPicture
    If Not Catch Then
        Call LogEvent("      Bild für Kontakt gespeichert in " & ContactPicture)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.addContactPic

'---------------------------------------------------------------------------------------
' Method : Sub ItmDispose
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Item is Saved, Deleted or MovedToTrash
'---------------------------------------------------------------------------------------
Sub ItmDispose(aItemO As Object, toThisFolder As Folder, msg As String)
Const zKey As String = "MailProcessing.ItmDispose"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")

Dim MovedItemO As Object
Dim Source As String

    aItemO.UnRead = True                         ' Mark unread whereever it now goes...
    Call ForceSave(aItemO)
    If toThisFolder.FolderPath = aItemO.Parent.FolderPath Then
        Call LogEvent("     " & msg & " was not moved to Folder " _
                    & Quote(aItemO.Parent.FolderPath) _
                    & "  because it is already there.", eLnothing)
        GoTo FuncExit
    End If
    Source = aItemO.Parent.FolderPath
    
    aBugTxt = "Item.Move"                        ' Try anything, autocatch
    Call Try
    Set MovedItemO = aItemO.Move(toThisFolder)
    If CatchNC Then
        If Left(Hex(ErrorCaught), 7) = "8004010" Then ' we have seen last hex as A-F
            DoVerify False, " maybe the SpecialSearchFolderName is gone???"
        End If
    End If
    
    If isEmpty(MovedItemO) Then
        Set MovedItemO = Nothing
        DoVerify False, "design big change ???"
        Call N_ClearAppErr
        Call ErrReset(0)
        GoTo doDel
    End If
    
    If Not MovedItemO Is Nothing Then
        If CatchNC Then
            If InStr(E_AppErr.Description, "sondern kopiert") > 0 Then
                Call LogEvent("     " & msg & " has been copied to " & Quote(toThisFolder) _
                            & " and no longer exists in " & Quote(aItemO.Parent.FolderPath) _
                            & " , " & vbCrLf & "but could now be a duplicate in " _
                            & Quote(Source))
                Call ErrReset(4)
                GoTo FuncExit
            Else
doDel:
                aBugTxt = "delete item"          ' Try anything, autocatch
                Call Try
                aItemO.Delete                    ' if it is no longer there, that's OK:
                If Catch(DoMessage:=False) Then
                    Call LogEvent("     " & msg & " has been copied to " _
                                & Quote(toThisFolder) _
                                & "  but no longer exists in " _
                                & Quote(aItemO.Parent.FolderPath) & b)
                    GoTo FuncExit
                End If
            End If
        End If
        Set aItemO = MovedItemO
        Set aItmDsc.idObjItem = aItemO
        Call LogEvent("     " & msg & " has been moved to Folder " _
                    & Quote(aItemO.Parent.FolderPath) & b, eLmin)
    ElseIf E_Active.errNumber = -2147219840 Then
        Set MovedItemO = aItemO
        aBugTxt = "delete Item"
        Call Try(-2147219840)
        aItemO.UnRead = True
        If Catch(DoMessage:=False) Then
            Call LogEvent("     " & msg & " can not be moved or deleted from " _
                        & Quote(aItemO.Parent.FolderPath), eLmin)
        Else
            Call LogEvent("     " & msg & " has been copied to " & Quote(toThisFolder) _
                        & "  and deleted from " & Quote(aItemO.Parent.FolderPath), eLmin)
        End If
    Else
        Call LogEvent("     " & msg & " should be in " _
                    & Quote(toThisFolder.FolderPath))
    End If

FuncExit:
    aBugTxt = "MailProcessing.ItmDispose failed"
    Call Try
    ItemInIMAPFolder = getAccountType(Source, aAccountTypeName) = olImap
    Catch

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.ItmDispose

'---------------------------------------------------------------------------------------
' Method : Sub RestEvn4Item
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub RestEvn4Item()
Const zKey As String = "MailProcessing.RestEvn4Item"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="MailProcessing")

Dim rsp As VbMsgBoxResult

    If StopRecursionNonLogged Then               ' no change in NoEventOnAddItem
        If Not NoEventOnAddItem Then
            If SkipedEventsCounter > 1 Then
                NoEventOnAddItem = True
            End If
        End If
        SkipedEventsCounter = SkipedEventsCounter - 1
    Else
        SkipedEventsCounter = SkipedEventsCounter - 1
        If SkipedEventsCounter > 1 Then          ' some elements waiting
            If DebugMode Then
                rsp = MsgBox("Check Unlogged items, Deferred >= " & SkipedEventsCounter _
                           & vbCrLf & "Reset? Cancel=Stop", vbYesNoCancel + vbDefaultButton2)
                If rsp = vbYes Then
                    SkipedEventsCounter = -SkipedEventsCounter
                ElseIf rsp = vbNo Then
                ElseIf rsp = vbCancel Then
                    DoVerify False
                End If
            End If
        End If
        If SkipedEventsCounter <= 0 Then         ' no reason now ???
            SkipedEventsCounter = -1
            NoEventOnAddItem = False
            SkipedEventsCounter = 0
        End If
    End If

FuncExit:
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.RestEvn4Item

'---------------------------------------------------------------------------------------
' Method : Function SetEventMode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SetEventMode(Optional force As Boolean) As Boolean
Const zKey As String = "MailProcessing.SetEventMode"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="MailProcessing")

    If Not StopRecursionNonLogged Then
        If Not NoEventOnAddItem _
        And SkipedEventsCounter Mod 9 = 2 Then
            If DebugMode Then
                MsgBox "Check Unlogged items, Deferred >= " & SkipedEventsCounter
            End If
        End If
    End If
    
    SkipedEventsCounter = SkipedEventsCounter + 1
    If Not NoEventOnAddItem Then
        If SkipedEventsCounter > 1 Then
            NoEventOnAddItem = True              ' so we can ProcCall additem once
        End If
    End If
    If force Or SkipedEventsCounter > 2 Then     ' accept no more
        SetEventMode = True
        StopRecursionNonLogged = True
    Else
        SetEventMode = False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' MailProcessing.SetEventMode

'---------------------------------------------------------------------------------------
' Method : Sub DoDeferred
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DoDeferred()
Dim zErr As cErr
Const zKey As String = "MailProcessing.DoDeferred"

'------------------- gated Entry -------------------------------------------------------
    

    If Deferred Is Nothing Then
        Set Deferred = New Collection            ' define a new one
        GoTo pExit
    End If
    If Deferred.Count = 0 Then
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim collItem As Object
Dim ProcessItem As Object
Dim AO As cActionObject
    
    'On Error GoTo 0
    For Each collItem In Deferred
        i = i + 1
        Set AO = collItem
        Set ProcessItem = aNameSpace.GetItemFromID(AO.aoObjID)
        If E_AppErr.errNumber <> 0 Then DoVerify False
        If IsMailLike(ProcessItem) Then
            ActionID = AO.ActionID
            Select Case AO.ActionID
            Case 3
                If Not ShutUpMode Then
                    If aID(aPindex).idObjDsc.objHasReceivedTime Then
                        Debug.Print "Starting on Deferred Item " & i _
                                  & " (Received " _
                                  & ProcessItem.ReceivedTime _
                                  & " SentOn " & ProcessItem.SentOn & ") of " _
                                  & Deferred.Count
                    Else
                        Debug.Print "Starting on Deferred Item " & i & " Created On " & aID(aPindex).idTimeValue
                    End If
                End If
                Call DoMailLike(ProcessItem)
                Call ShowStatusUpdate
            Case Else
                DoVerify False, " not implemented Action"
            End Select
        Else
            If DebugMode Then DoVerify False, _
               "deferred processing only intended for Mail-Like Items"
        End If
    Next collItem
    Set Deferred = New Collection                ' define a new one
    If i > 0 Then
        Call LogEvent("==== Processed " & i _
                    & " previously un-processed items", eLmin)
    End If
    
FuncExit:
    Set ProcessItem = Nothing
    Set collItem = Nothing
    Set AO = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' MailProcessing.DoDeferred

'---------------------------------------------------------------------------------------
' Method : Sub Set2Logged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Set2Logged(aMiObj As MailItem)

Const zKey As String = "MailProcessing.Set2Logged"
    Call DoCall(zKey, tSub, eQzMode)

Dim oCat As String
    On Error GoTo bad
    oCat = aMiObj.Categories
    If InStr(oCat, LOGGED) = 0 Then
        Call AppendTo(oCat, LOGGED, ";", ToFront:=True)
        aMiObj.Categories = oCat
    End If
    aMiObj.UnRead = False
    If Not aMiObj.Saved Then
        aMiObj.Save
    End If
    GoTo FuncExit
bad:                DoVerify False

FuncExit:
    Call DoExit(zKey)

End Sub                                          ' MailProcessing.Set2Logged

'---------------------------------------------------------------------------------------
' Method : Function CloseInsptrs
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Close open Inspector windows, dft: only type Email, else CloseAnything:=True
' Note   : Working backwards to prevent skipping any items
'---------------------------------------------------------------------------------------
Function CloseInsptrs(Optional CloseAnything As Boolean)
Const zKey As String = "MailProcessing.CloseInsptrs"
    Call DoCall(zKey, tFunction, eQzMode)

Dim lInsp As Outlook.Inspector
Dim lCount As Long
Dim i As Long
Dim lItem As Object
 
    lCount = olApp.Inspectors.Count
 
    For i = lCount To 1 Step -1
        Set lInsp = olApp.Inspectors(i)
        Set lItem = lInsp.CurrentItem
        If CloseAnything Then
            If DebugMode Then
                Debug.Print "Item class: " & lInsp.CurrentItem.Class & b;
            End If
            GoTo DoAny
        Else
            If IsMailLike(lItem) Then
DoAny:
                lItem.Close olDiscard
                If DebugMode Then
                    Debug.Print "Item discarded, Subject: " & Quote(lItem.Subject)
                End If
            End If
        End If
    Next i

FuncExit:

zExit:
    Call DoExit(zKey)

End Function                                     ' MailProcessing.CloseInsptrs

'---------------------------------------------------------------------------------------
' Method : RestrictedItemsShow
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Show items in RestrictedItemCollection
'---------------------------------------------------------------------------------------
Sub RestrictedItemsShow(Optional findCount As Long)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "MailProcessing.RestrictedItemsShow"
    Call DoCall(zKey, "Sub", eQzMode)

Dim aObj As Object
Dim i As Long

    If findCount = 0 Then
        findCount = RestrictedItemCollection.Count
    Else
        findCount = Min(findCount, RestrictedItemCollection.Count)
    End If
    
    For i = 1 To findCount
        Set aObj = RestrictedItemCollection.Item(i)
        Debug.Print LString(i, 5) & LString(aObj.Subject, lKeyM) _
      & b & LString(aObj.SenderName, 30) & b & aObj.SentOn
    Next i
    
FuncExit:
    Set aObj = Nothing

zExit:
    Call DoExit(zKey)

End Sub                                          ' MailProcessing.RestrictedItemsShow



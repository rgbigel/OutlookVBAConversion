Attribute VB_Name = "DupeDeleter"
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Function AskUserAndInterpretAnswer
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function AskUserAndInterpretAnswer(ByVal oMessage As String)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.AskUserAndInterpretAnswer"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim UserAnswer As String
Dim passtGenau As Boolean
Dim passtEtwa As Boolean
Dim itemCount As Long
Dim isFromSortedCollection As Boolean

'   Message = oMessage ?????
    passtGenau = cMisMatchesFound = 0
    passtEtwa = (cMisMatchesFound < MaxMisMatchesForCandidates _
                    And AcceptCloseMatches _
                    And Not passtGenau _
                    And Not SuperRelevantMisMatch)
    If sortedItems(1) Is Nothing Then
        If SelectedItems Is Nothing Then
            If Not passtGenau Or MPEchanged Then
                AllItemDiffs = AllItemDiffs & vbCrLf & Mid(MPEItemDiffs, 3)
                UserAnswer = Quote(fiMain(1)) & " hat folgende Abweichungen: " _
                            & vbCrLf & AllItemDiffs
            Else
                UserAnswer = Quote(fiMain(1)) & " hat keine relevanten Abweichungen" _
                            & vbCrLf & AllItemDiffs
            End If
            If MPEchanged Or IsComparemode Then
                GoTo justShow
            Else
                rsp = vbIgnore
                GoTo rspIsSet
            End If
        Else
            itemCount = SelectedItems.Count
            isFromSortedCollection = False
        End If
    Else
        itemCount = sortedItems(1).Count
        isFromSortedCollection = True
    End If
    
    If AcceptCloseMatches Then
        If cMisMatchesFound < MaxMisMatchesForCandidates Then
            UserDecisionRequest = True
        ElseIf Not IsComparemode Then
            UserDecisionRequest = False
        End If
    End If
    
    If eOnlySelectedItems _
    And itemCount < 5 _
    And Not IsComparemode Then
        UserDecisionRequest = True
    End If
    If Not AllPropsDecoded Then
        UserAnswer = Quote(fiMain(1)) _
                & " sollte bei unvollständigem Vergleich nicht gelöscht werden " _
                & vbCrLf & AllItemDiffs
justShow:
        AskUserAndInterpretAnswer = False
        Diffs = AllItemDiffs
        If Not displayInExcel Or Not AllPropsDecoded Then
            GoTo askuser
        Else
            Call DisplayExcel(O, _
                            relevant_only:=True, _
                            unconditionallyShow:=True)
            MsgBox UserAnswer, vbOKOnly
        End If
    ElseIf IsComparemode Or passtGenau _
        Or passtEtwa Or UserDecisionRequest Then
        ' im Prinzip löschen sinnvoll bzw Benutzer kann alles...
        LoeschbestätigungCaption = "Löschen von Objekten bestätigen "
askuser:
        Message = oMessage
        bDefaultButton = "Go"  ' der vordere Button ist IMMER Default
        If MatchPoints(1) <= MatchPoints(2) Then
            rsp = vbYes         ' der Bessere oder der Ältere wird gelöscht
                                ' (Modification Date Sorted ascending)
            b1text = WorkIndex(1)
            b2text = WorkIndex(2)
        Else
            b2text = WorkIndex(1)
            b1text = WorkIndex(2)
            rsp = vbNo
        End If
        If Not AllPropsDecoded Then
            rsp = vbCancel  ' default is not delete if compare is incomplete
            bDefaultButton = "Cancel"
            Diffs = UserAnswer
            GoTo ShowLoeschbestaetigung
        End If
        
        If Not passtEtwa _
        And UserDecisionRequest And AcceptCloseMatches _
        And cMisMatchesFound < MaxMisMatchesForCandidates Then
           passtEtwa = True
        End If
        
        If WantConfirmation Or passtEtwa And Not passtGenau Then
              Diffs = fiMain(1) & vbCrLf & " +++ " & objTypName _
                        & " in " & curFolderPath _
                        & ", Item Nr's. " & b1text & " und " & b2text _
                        & " Länge des Inhalts: " & Len(fiBody(1)) _
                        & vbCrLf & Message
            ' Nutzerbestätigung notwendig
            If passtGenau Then ' ... weil explizit gewünscht
                        
                Message = "==> Item " & b1text _
                                       & " oder " _
                                      & b2text _
                                       & " wird " _
                                      & killMsg
                Message = Message & vbCrLf _
                    & "[Wenn Löschungen im Ordner " & curFolderPath _
                    & " ab jetzt nicht mehr bestätigt werden sollen, Parameter ändern.]"
            Else                        ' ... weil Abweichungen vorhanden
                Message = "==> Item " & b1text _
                                       & " oder " _
                                      & b2text _
                                       & " wird " _
                                      & killMsg _
                                      & ". " _
                                      & b3text _
                                      & " löscht nichts."
                bDefaultButton = "Cancel"
            End If
            
            If xUseExcel And displayInExcel Then
                If Not xlApp.Visible Then
                    UserDecisionRequest = True  ' ask after excel display
                End If
                Call DisplayExcel(O, _
                                    relevant_only:=True, _
                                    unconditionallyShow:=True)
            End If
ShowLoeschbestaetigung:
            Diffs = Diffs & vbCrLf & vbCrLf & Mid(MatchData, 3)
            If InStr(Diffs, AllItemDiffs) = 0 Then
                Diffs = Diffs & vbCrLf & AllItemDiffs
            End If
            DeleteNow = False
            Set LBF = Nothing
            Set LBF = New frmDelConfirm
            Call Try
            LBF.Show
            If Catch Then
                DoVerify False
            End If
            Set LBF = Nothing
            Call ErrReset(0)
            If DeleteNow Then
                DoTheDeletes
                GoTo askuser
            End If
            If askforParams Then
askforParams:
                b1text = "Zurück"
                b2text = "DoVerify"                        ' NOTE: button Name is "bDebugStop"
                Message = "Bei " & b3text _
                        & " wird die Doublettensuche beendet, " _
                        & "ohne Löschungen " _
                        & " durchzuführen. Es liegen bisher Lösch-Vormerkungen für " _
                        & dcCount & " Einträge vor."
                Set LBF = Nothing
                Set LBF = New frmDelParms
                LBF.Show
                Set LBF = Nothing
                askforParams = False
                If rsp = vbCancel Then
                    Message = "Verarbeitung abgebrochen"
                    Call LogEvent(Message, eLall)
                    If TerminateRun Then
                        GoTo FuncExit
                    End If
                End If
                If LenB(b3text) = 0 Then
                    UserAnswer = "Löschantwort=Debug Debug.Assert False" & vbCrLf
                    AskUserAndInterpretAnswer = False
                    DeletedItem = WorkIndex(1)
                    DeleteIndex = 0
                    rsp = vbRetry
                    GoTo logandout
                Else
                    If askforParams Then
                        GoTo askforParams
                    Else
                        GoTo askuser
                    End If
                End If
            End If
            bDefaultButton = "Nutzer wählte Button "
        ElseIf rsp = vbIgnore Then     ' full Match and no confirmation
            DeleteIndex = 1
            bDefaultButton = "Automatische Selektion, gelöscht wird " _
                        & b1text _
                        & ", " & fiMain(DeleteIndex) & vbCrLf
            rsp = vbYes ' delete default item, lesser Match or older
        End If
rspIsSet:
        If displayInExcel Then
            Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
            If xlApp.Visible Then
                xlApp.EnableEvents = False
                xlApp.Visible = False
            End If
            Catch
        End If
        
        Select Case rsp
        Case vbNo
            If b2text = WorkIndex(2) Then
                DeleteIndex = 2
            ElseIf b2text = WorkIndex(1) Then
                DeleteIndex = 1
            Else
                DoVerify False
            End If
            UserAnswer = bDefaultButton & DeleteIndex _
                        & ", Löschvormerkung zu " & b2text _
                        & b & fiMain(DeleteIndex) & vbCrLf
            DeletedItem = b2text
            AskUserAndInterpretAnswer = True
            Call GenerateDeleteRequest(DeleteIndex, DeletedItem, isFromSortedCollection)
        Case vbYes
            If b1text = WorkIndex(2) Then
                DeleteIndex = 2
            ElseIf b1text = WorkIndex(1) Then
                DeleteIndex = 1
            Else
                DoVerify False
            End If
            UserAnswer = bDefaultButton & DeleteIndex _
                        & ", Löschung von " & b1text _
                        & b & fiMain(DeleteIndex) & vbCrLf
            DeletedItem = b1text
            AskUserAndInterpretAnswer = True
            Call GenerateDeleteRequest(DeleteIndex, DeletedItem, isFromSortedCollection)
        Case vbRetry
            UserAnswer = "Retry: Items komplett in Excel anzeigen (keine Löschung)" & vbCrLf
            AllPropsDecoded = False
            AskUserAndInterpretAnswer = False
        Case Else       ' cancel request
            UserAnswer = "Löschantwort=Cancel (keine Löschung)" & vbCrLf
            AskUserAndInterpretAnswer = False
            DeletedItem = -1
            DeleteIndex = 0
        End Select
    Else ' Passt überhaupt nicht
        AskUserAndInterpretAnswer = False
        UserAnswer = "(" & fiMain(1) & ") sind nicht gleich " _
                        & vbCrLf & AllItemDiffs
    End If
logandout:
    Call LogEvent("Items: " & WorkIndex(1) & "/" & WorkIndex(2) _
                & b & UserAnswer, eLall)
    UserDecisionEffective = True

FuncExit:
    Set LBF = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' DupeDeleter.AskUserAndInterpretAnswer

'---------------------------------------------------------------------------------------
' Method : Sub GenerateDeleteRequest
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GenerateDeleteRequest(delPx As Long, delItemIndex As Long, doSort As Boolean)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.GenerateDeleteRequest"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim DelObjectEntry As cDelObjectsEntry

    Set DelObjectEntry = New cDelObjectsEntry
    DelObjectEntry.DelObjPos = delPx
    DelObjectEntry.DelObjPindex = delItemIndex
    DelObjectEntry.DelObjInd = doSort
    If DeletionCandidates Is Nothing Then
        Set DeletionCandidates = New Dictionary
    End If
    DeletionCandidates.Add delItemIndex, WorkIndex(delPx)
    
    LöListe = LimitAppended(LöListe, ", " & delItemIndex, 255, "... ")
    
    dcCount = dcCount + 1

FuncExit:
    
ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.GenerateDeleteRequest

'---------------------------------------------------------------------------------------
' Method : Sub PerformChangeOpsForMapiItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PerformChangeOpsForMapiItems()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.PerformChangeOpsForMapiItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim sourcecol As Long
Dim px As Long
    If O Is Nothing Then
        GoTo ProcReturn
    End If
' if WithEditing only!!
    With O.xlTSheet
        If Not O Is Nothing Then
            For px = 1 To 2
                If .Cells(1, px + 1).Text = "del" Then
                    If px = 1 Then
                        delSource(px) = WorkIndex(1)
                    Else
                        delSource(px) = WorkIndex(2)
                    End If
                    If sortedItems(px) Is Nothing Then
                        Call GenerateDeleteRequest(px, delSource(px), False)
                    Else
                        Call GenerateDeleteRequest(px, delSource(px), _
                                                    sortedItems(px).Count > 0)
                    End If
                    
                    Call LogEvent("Item " & delSource(px) _
                            & " in Excel zum Löschen vorgemerkt", eLall)
                    UserDecisionEffective = True
                Else
                    delSource(px) = 0
                End If
            Next px
            
            If .Cells(1, changeCounter).Value = 0 Then
                Call LogEvent("Es wurden keine Änderungen in Excel durchgeführt", eLall)
                UserDecisionEffective = True
                GoTo ProcReturn
            End If
            
            If DebugMode Or DebugLogging Then
                Debug.Print Format(Timer, "0#####.00") _
                    & vbTab & "Performing changes/deletes from Excel to Outlook"
            End If
            For i = 2 To aID(2).idAttrDict.Count + 1 ' i numbers rows
                Err.Clear
                If Left(.Cells(i, ValidCol).Text, 3) <> "***" Then ' Value is editable
                    px = 0    ' used as flag for no changes in this line
                    If .Cells(i, ChangeCol3).Text = "<" Then
                        sourcecol = ChangeCol2
                        px = 1
                    ElseIf .Cells(i, ChangeCol3).Text = ">" Then
                        sourcecol = ChangeCol1
                        px = 2
                    ElseIf .Cells(i, ChangeCol3).Text = "?" Then ' Error previously
                        GoTo nloop  ' will not edit this
                    ElseIf .Cells(i, ChangeCol3).Text = "!" Then ' success previously
                        GoTo nloop  ' will not edit this
                        ' Indicator of change "<=>" is in cols 14/15
                    ElseIf InStr(.Cells(i, WatchingChanges).Text, ">") > 0 Then
                                    ' changecounter1--V V--changecounter2
                        ' if something changed in col 2/3, col 17/18 contain old values
                        ' col 7 / 8 contain raw values, e.g. used to compare empty sources
                        ' col 7, 17/ 8, 18 is updated on the excel side when selection changes
                        ' check if the original value differs at all
                        If .Cells(i, ChangeCol1).Text <> .Cells(i, 17).Text _
                        Or .Cells(i, ChangeCol1).Text <> .Cells(i, 7).Text Then
                            sourcecol = ChangeCol1
                            px = 1
                        End If
                        ' Indicator of change is in cols 14/15
                    ElseIf InStr(.Cells(i, changeCounter).Text, ">") > 0 Then
                        If .Cells(i, ChangeCol2).Text <> .Cells(i, 18).Text _
                        Or .Cells(i, ChangeCol2).Text <> .Cells(i, 8).Text Then
                            sourcecol = ChangeCol2
                            px = 2
                        End If
                    Else
                        GoTo nloop
                    End If
                    If ErrorCaught = 0 Then
                        .Cells(i, ChangeCol3).Value = "!"
                    Else
                        .Cells(i, ChangeCol3).Value = "?"
                    End If
                End If
nloop:
                If delSource(px) = 0 Then
                    If px > 0 Then
                        Call storeAttribute(i, sourcecol, px)
                    End If
                End If
    
            Next i
            If Not ShutUpMode Then
                Debug.Print Format(Timer, "0#####.00") _
                    & vbTab & "Ended Performing changes/deletes from Excel to Outlook"
            End If
        End If
    End With ' O.xlTSheet

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.PerformChangeOpsForMapiItems

'---------------------------------------------------------------------------------------
' Method : Sub CheckDoublesInFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CheckDoublesInFolder(ByRef curFolder As Folder)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.CheckDoublesInFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim itemNo As Long
Dim sortFields As String
Dim sortOrder As String
Dim sCompRes As Long
Dim ShortName As String
Dim totalDeletes As Long
Dim deletedNow As Long
Dim tCompRes As String
   
    IsComparemode = False
    Call FindTrashFolder
    If curFolder = topFolder And FolderPathLevel = 0 Then
        itemNo = 0
    Else    ' ***???***
        ' *** curFolderPath = FullFolderPath(FolderPathLevel) & "\" & curFolder.Name
    End If
    ' FolderPathLevel = FolderPathLevel + 1
    ' FullFolderPath(FolderPathLevel) = curFolderPath
    Set Folder(1) = curFolder
    Set Folder(2) = curFolder
    If eOnlySelectedItems Then
        Fctr(1) = 0
        Ictr(1) = SelectedItems.Count
    Else
        If curFolder.Parent Is Nothing Then
            Fctr(1) = 0 ' no parent => no Folders
        Else
            Fctr(1) = curFolder.Folders.Count
        End If
        Ictr(1) = curFolder.Items.Count
    End If
    Ictr(2) = Ictr(1)
    Fctr(2) = Fctr(1)
    ShortName = Left(curFolder.Name, 4)
    If ShortName = "Dele" _
    Or ShortName = "Gelö" _
    Or ShortName = "Tras" _
    Or ShortName = "Junk" _
    Or ShortName = "Spam" _
    Or ShortName = "Uner" _
    Or ShortName = "Dupl" _
    Or Fctr(1) + Ictr(1) = 0 _
    Then ' skip unwanted stuff; could contain any item type!
skipThis:
        Call LogEvent("<======> skipping Folder " & curFolder.FolderPath _
            & " because it matches " & ShortName _
            & " Time: " & Now())
        GoTo ProcReturn
    End If
    
    ' hamma schon: Call BestObjProps(curFolder)
    If AskEveryFolder _
    And Not (SkipNextInteraction Or eOnlySelectedFolder) Then
        WantConfirmationThisFolder = WantConfirmation
        
        If eOnlySelectedItems Then
            Message = "Bitte bestätigen Sie die Parameter der Doublettensuche" _
                & "in den selektierten Items"
        Else
            Message = "Wollen Sie Doubletten suchen im Ordner " _
                        & curFolderPath
            If curFolder.Folders.Count > 0 Then
                Message = Message & vbCrLf & "(beginnend mit seinen " _
                    & curFolder.Folders.Count & " enthaltenen Ordnern)"
            End If
        End If
        b1text = "Ja"
        b2text = "Nein"
        Set LBF = Nothing
        Set LBF = New frmDelParms
        bDefaultButton = "Go"
        LBF.Caption = "Parameter für das Löschen von Doubletten in Ordner " _
                        & curFolderPath
        LBF.Show
        Set LBF = Nothing
        
        Select Case rsp
        Case vbNo
            GoTo skipThis
        Case vbCancel
            Call LogEvent("=======> Stopped before Entering Folder " & Quote(curFolder.FolderPath) & _
                    " containing " & Fctr(1) & " Folders and " _
                    & Ictr(1) & " Items. Time: " & Now())
            If TerminateRun Then
                GoTo ProcReturn
            End If
        Case Else
        End Select
    End If
    SkipNextInteraction = False
    
    If WantConfirmation = True Then
        Message = " (Confirmation mode)"
    Else: WantConfirmation = False
        Message = vbNullString
    End If
    If eOnlySelectedItems Then
        Call LogEvent("=======> Processing selected Items in " & curFolderPath _
            & Message & " Time: " & Now(), eLnothing)
    Else
        Call LogEvent("=======> Entering Folder " & curFolderPath _
            & " containing " & Fctr(1) & " Sub-Folders and " _
            & Ictr(1) & " Items." & Message & " Time: " & Now())
    End If
    Message = vbNullString
    dcCount = 0
    Set DeletionCandidates = New Dictionary
    LöListe = vbNullString
    
    ' recurse into FOlDERS
    ' ====================
    For itemNo = 1 To Fctr(1)   ' ??? ***
        If DebugLogging Then
            DoVerify False, " if we are in the call chain of loopFolders, this should never be reached"
        End If
        If curFolder.DefaultItemType = olContactItem Then
            Exit For                                                ' do not recurse ContactItems
        End If
        If curFolder.Folders(itemNo).Items.Count > 0 Then
            Call CheckDoublesInFolder(curFolder.Folders(itemNo))    ' ProcCall recursion
        End If
    Next itemNo
    
    ' Process Folder ITEMS
    ' ====================
Restart:
    StopLoop = False
    Call GetSortableItems(curFolder)
    Set aTD = Nothing
    ' ??? *** only needed if a change of class is possible
    Call SplitMandatories(TrueCritList)
    If Not aTD Is Nothing Then                  ' items to work on need new rules
        DoVerify False, " ??? *** change of class is possible"
        aTD.adRules.RuleInstanceValid = False
        Call SplitDescriptor(aTD)   ' aTD is possibly changed now
    End If
    sortFields = SortMatches
    
    Err.Clear
    If Not sortedItems(1) Is Nothing Then
        aBugTxt = "sort using sortFields=" & sortFields
        Call Try
        If sortedItems(1).Count > 1 Then
            If curFolder.DefaultItemType = olContactItem Then
                sortedItems(1).sort sortFields, olAscending
                sortOrder = "Ascending"
            Else
                sortedItems(1).sort sortFields, False
                sortOrder = "Descending"
            End If
        End If
        If Catch Then
            Message = vbCrLf & Err.Description & vbCrLf & vbCrLf _
                        & "Bitte ändern Sie die [Sortierparameter]: " _
                        & sortFields
            b1text = "Weiter"
            b2text = vbNullString
            Set LBF = Nothing
            Set LBF = New frmDelParms
            LBF.Show
            Set LBF = Nothing
            If rsp = vbCancel Then
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            End If
            GoTo Restart
        End If
        
        Ictr(1) = sortedItems(1).Count
        Set sortedItems(2) = sortedItems(1)
        Ictr(2) = sortedItems(2).Count
    End If  ' sorting was possible
    
    fiMain(1) = vbNullString
    fiMain(2) = vbNullString  ' This will cause the first comparison result =0
    
    DeletedItem = -1
    DeleteIndex = -1
        
    WorkIndex(1) = 1  ' if ascending, this is the smaller one
    WorkIndex(2) = 1  ' if descending, the next one is the smaller
    Set aID(1) = Nothing
    Set aID(2) = Nothing
    Set aObjDsc = Nothing
    sCompRes = 0
    tCompRes = " ? "
    
    ' note this scheme only makes sense if we are working on only one
    ' sorted collection, which are obviously in one Folder
    ' beware: as deletes are probably done in batches, item idices
    ' are not monotonic when the deletes are actually done
        
    Do While (WorkIndex(1) < Ictr(1) And WorkIndex(2) < Ictr(2))
        Call ShowStatusUpdate
        AllDetails = vbNullString
        AllPropsDecoded = False
        Set aTD = Nothing
        If DeletedItem = -1 Then
            ' doing a very fast compare based on main identifications
            sCompRes = StrComp(fiMain(1), fiMain(2), vbTextCompare)
            ' -1 means first is bigger, +1 first is smaller, =0 if same
            If sCompRes = 0 Then    ' fiMain(1) = fiMain(2)
                tCompRes = " = "
                Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                    & curFolderPath & ": " _
                    & WorkIndex(1) & tCompRes _
                    & WorkIndex(2) & ", " & MainObjectIdentification _
                    & "= " & vbCrLf & WorkIndex(1) & ": " & Quote(fiMain(1)))
                If sortedItems(1).Item(WorkIndex(1)).EntryID _
                = sortedItems(2).Item(WorkIndex(2)).EntryID Then
                    GoTo advance_Second
                End If
            ElseIf sCompRes < 0 Then ' fiMain(1) > fiMain(2)
                tCompRes = " > "
                If MinimalLogging < eLSome Then
                    Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                        & curFolderPath & ": " _
                        & WorkIndex(1) & tCompRes _
                        & WorkIndex(2) & ", " & MainObjectIdentification _
                        & ": " & vbCrLf & WorkIndex(1) _
                        & ": " & Quote(fiMain(1)) & b, eLSome)
                    Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
                End If
                If sortOrder = "Descending" Then
                    ' desc sort: stepping first will (probably) make
                    ' fimain(1) smaller, so match is possible next time
                    GoTo advance_First ' make fiMain(1) decrease
                Else    ' ascending sort: make fiMain(1) bigger
                    GoTo advance_Second
                End If
            Else ' fiMain(1) < fiMain(2)
                tCompRes = " < "
                If MinimalLogging < eLSome Then
                    Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                        & curFolderPath & ": " _
                        & WorkIndex(1) & tCompRes _
                        & WorkIndex(2) & ", " & MainObjectIdentification _
                        & ": " & vbCrLf & WorkIndex(1) _
                        & ": " & Quote(fiMain(1)) & b, eLSome)
                    Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
                End If
                If sortOrder = "Descending" Then
                    GoTo advance_Second ' make fiMain(2) decrease
                Else
                    GoTo advance_First  ' make fiMain(1) bigger
                End If
            End If
        ElseIf DeletedItem = WorkIndex(1) Then
            ' we have plans to delete first item, skip it for further comparison
advance_First:
            WorkIndex(1) = WorkIndex(2)
            Call CopyAttributeStack(2)
            fiMain(1) = fiMain(2)
            Set aID(1) = aID(2)
            Set aObjDsc = aID(2).idObjDsc
            Set aID(2) = Nothing
            ' as item 1 is now on "left" side already, advance "right" side
            WorkIndex(2) = WorkIndex(2) + 1
            AttributeUndef(2) = 0
        ElseIf DeletedItem = WorkIndex(2) Then
            ' if item 2 is deleted, skip on to the next on "right" side
advance_Second:
            WorkIndex(2) = WorkIndex(2) + 1
            Set aID(2) = Nothing
            fiMain(2) = vbNullString
            AttributeUndef(2) = 0
        Else
            tCompRes = " und "
            If MinimalLogging < eLSome Then
                Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                    & curFolderPath & ": " _
                    & WorkIndex(1) & tCompRes _
                    & WorkIndex(2) & ", " & MainObjectIdentification _
                    & ": " & vbCrLf & WorkIndex(1) _
                    & ": " & Quote(fiMain(1)) & b, eLSome)
                Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
            End If
advance_Both:
            ' this because muliplicates are possible, not only duplicates
            WorkIndex(1) = WorkIndex(1) + 1
            Set aID(1) = Nothing
            Set aObjDsc = Nothing
            AttributeUndef(1) = 0
            WorkIndex(2) = WorkIndex(2) + 1
            Set aID(2) = Nothing
            AttributeUndef(2) = 0
        End If
        ' never compare identical objects
        If WorkIndex(1) = WorkIndex(2) Then ' possible after deleting
            ' do same steps as advance_Second
            WorkIndex(2) = WorkIndex(2) + 1
            Set aID(2) = Nothing
            AttributeUndef(2) = 0
        End If
        
        If Not xlApp Is Nothing Then
            If Not O Is Nothing Then
                ' if excel is open, we could have objDumpMade > 0
                aOD(0).objDumpMade = 0   ' which would be bad info
            End If
        End If
        
        mustDecodeRest = False
        If WorkIndex(2) > sortedItems(2).Count Then
            GoTo leaveLoop  ' running out of bounds
        End If
        ' Progress indicator
        If (Max(WorkIndex(1), WorkIndex(2)) - 1) Mod 25 = 1 _
        And sortedItems(1).Count > 2 Then
            Call LogEvent("   (progress info:) processing items " _
                & WorkIndex(1) & " and " & WorkIndex(2), eLall)
        End If
        
        ' this determines fiMain()
        If aID(1) Is Nothing Then
            Call GetAobj(1, WorkIndex(1))
            objTypName = DecodeObjectClass(getValues:=True)
        End If
        If aID(2) Is Nothing Then
            Call GetAobj(2, WorkIndex(2))
            objTypName = DecodeObjectClass(getValues:=True)
        End If  ' if not, item same as in last loop pass
        
        If objTypName = "-" Then
            DoVerify False
            GoTo nextOne
        End If
                    
        If fiMain(1) <> fiMain(2) Then
            If Not IsComparemode Or quickChecksOnly Then
                GoTo nextOne
            End If
        End If
        
        Call ItemIdentity   ' identical by test OR by user decision possible!!
'            ==================================================================
        DeleteIndex = 0
nextOne:
        If StopLoop Then
            GoTo leaveLoop
        End If
        If dcCount Mod 25 = 0 Then
            If dcCount > 0 Then
                deletedNow = DoTheDeletes   ' user can still decide yes/no, true deletes 0/25
                totalDeletes = totalDeletes + deletedNow ' we actually deleted that many, max 25
                WorkIndex(1) = WorkIndex(1) - deletedNow ' correct for deleted items
                WorkIndex(2) = WorkIndex(2) - deletedNow
                Set logItem = Nothing   ' start a new log entry (speed!)
            End If
        End If
    Loop
    
leaveLoop:
    If Not xlApp Is Nothing Then
        Call ClearWorkSheet(xlA, O)   ' erase this one, new one will be started when needed
    End If
    totalDeletes = totalDeletes + DoTheDeletes  ' last round of deletes
    If sortedItems(1) Is Nothing Then
        Message = "<======= Exiting Folder "
    Else
        itemNo = sortedItems(1).Count ' this is whats left after deletes
        If eOnlySelectedItems Then
            Message = "<=== compared " & Ictr(1) & " selected items, deleted " _
                    & totalDeletes & " as duplicates"
            Call UnSelectItems  ' ???*** undoes Marking as SEL (in ManagerName)
        Else
            Message = "<======= Exiting Folder "
        End If
    End If
    
    Call LogEvent(Message & Quote(curFolderPath) _
        & " . Nr. of items removed: " & totalDeletes _
        & " Time: " & Now(), eLall)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.CheckDoublesInFolder

'---------------------------------------------------------------------------------------
' Method : Sub CopyAttributeStack
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CopyAttributeStack(sourcePX As Long)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.CopyAttributeStack"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Attrib As cAttrDsc
Dim DestPX As Long
    If sourcePX = 1 Then
        DestPX = 2
    Else
        DestPX = 1
    End If
    Set aID(DestPX).idAttrDict = New Dictionary
    AttributeUndef(DestPX) = AttributeUndef(sourcePX)
    For Each Attrib In aID(sourcePX).idAttrDict
        aID(DestPX).idAttrDict.Add Attrib.adKey, Attrib
    Next Attrib

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.CopyAttributeStack

'---------------------------------------------------------------------------------------
' Method : Sub DecodeAllPropertiesFor2Items
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DecodeAllPropertiesFor2Items(ByVal StopAfterMostRelevant As Boolean, ByVal ContinueAfterMostRelevant As Boolean, Optional onlyItemNo As Long)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.DecodeAllPropertiesFor2Items"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim checkPx As Long
    If onlyItemNo < 3 Then
        checkPx = onlyItemNo
    Else
        checkPx = 1
    End If
    With aID(checkPx)
        If .idAttrDict Is Nothing Then
            Set .idAttrDict = New Dictionary
        End If
    End With ' aID(checkPx)
    If ContinueAfterMostRelevant Then
        If isEmpty(MostImportantProperties) Then
            AttributeIndex = 0
        Else
            AttributeIndex = UBound(MostImportantProperties) + 1
        End If
    Else
        If DebugMode = Not DebugMode Then               ' UNREACHABLE !!! Here ???
            Call initializeComparison                   ' restart here for manual debug
            Call initializeExcel
        End If
        ContinueAfterMostRelevant = False
    End If

doDecode:
    If onlyItemNo <> 2 Then
        Call SetupAttribs(aID(1).idObjItem, 1, True)
    End If
    If onlyItemNo = 1 Then
        Call RulesToExcel(1, True)
    Else
        If isEmpty(MostImportantProperties) Then
            AttributeIndex = 0
        ElseIf ContinueAfterMostRelevant Then
            AttributeIndex = UBound(MostImportantProperties) + 1
        Else
            AttributeIndex = 0
        End If
        cMissingPropertiesAdded = 0
        Call SetupAttribs(aID(2).idObjItem, 2, True)
        stpcnt = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)
    End If
   
    If onlyItemNo <> 1 Then ' do some plausi checks
        If aID(1).idObjItem.ItemProperties.Count <> aID(2).idObjItem.ItemProperties.Count Then
            If aID(1).idObjItem.ItemProperties.Count > aID(2).idObjItem.ItemProperties.Count Then
                If aID(2).idObjItem.Class = aID(2).idObjItem.Parent.Class Then
                    Set aID(2).idObjItem = aID(2).idObjItem.Parent
                Else
                    Debug.Assert False
                   ' GoTo objecterror
                End If
            Else
                If aID(1).idObjItem.Class = aID(1).idObjItem.Parent.Class Then
                    Set aID(1).idObjItem = aID(1).idObjItem.Parent
                Else
                    Debug.Assert False
                   ' GoTo objecterror
                End If
            End If
            If aID(1).idObjItem.ItemProperties.Count <> aID(2).idObjItem.ItemProperties.Count Then
                    Debug.Assert False
                  ' GoTo objecterror
            End If
        End If
    End If
    
    YleadsXby = 0 ' we can never have a misadjustment yet
    
    If onlyItemNo <> 1 Then
        MaxPropertyCount = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)
    Else
        MaxPropertyCount = aID(1).idAttrDict.Count
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.DecodeAllPropertiesFor2Items

'---------------------------------------------------------------------------------------
' Method : Function DoTheDeletes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DoTheDeletes() As Long
Dim zErr As cErr
Const zKey As String = "DupeDeleter.DoTheDeletes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim delItem As cDelObjectsEntry

    If dcCount > 0 Then
        Message = "Folgende " & dcCount & " Einträge können " _
                    & killMsg & " werden: " & vbCrLf
        For i = 0 To dcCount - 1
            Set delItem = DeletionCandidates.Items(i)
            Message = Message & b & delItem.DelObjPindex
        Next i
        bDefaultButton = "Go"
        b1text = Replace(killMsg, "o", "ie", 1, 1)  ' "verschoben" -> "verschieben"
        b1text = Left(b1text, 10)
        b2text = "Nicht löschen"
        Set LBF = Nothing
        Set LBF = New frmDelParms
        LBF.Caption = "Bestätigung vor dem Löschvorgang (" _
                        & TrashFolder.FullFolderPath & ")"
        LBF.Show
        Set LBF = Nothing
        If rsp = vbCancel Then
            Call LogEvent(Message, eLall)
            If TerminateRun Then
                GoTo ProcReturn
            End If
        End If
        If rsp = vbYes Then
            DoTheDeletes = dcCount
            For i = 0 To DeletionCandidates.Count - 1
                Set delItem = DeletionCandidates.Items(i)
                Call ShowStatusUpdate
                Call TrashOrDeleteItem(delItem)
            Next i
            Message = Replace(Message, "können", "wurden")
            Message = Replace(Message, "werden", vbNullString)
        Else
            Message = "Die Löschungen wurden nicht bestätigt"
       End If
    Else
        Message = "Es wurden keine Löschungen ausgewählt"
    End If
    
    Set DeletionCandidates = New Dictionary
    
    LöListe = vbNullString
    dcCount = 0
    Call LogEvent(Message, eLall)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' DupeDeleter.DoTheDeletes

'---------------------------------------------------------------------------------------
' Method : Sub FindTrashFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FindTrashFolder() ' in selected TopFolder
Dim zErr As cErr
Const zKey As String = "DupeDeleter.FindTrashFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim baseFolder As Folder
    Set baseFolder = topFolder
    ' curFolderPath = baseFolder.FolderPath ???***
    ' FolderPathLevel = 0
    ' FullFolderPath(FolderPathLevel) = vbNullString
    If getTrashFolder(baseFolder, vbNullString) Is Nothing Then
        killType = "Löschungen sind endgültig"
        killMsg = "gelöscht"
    Else
        killType = "verschiebt in " & TrashFolderPath
        killMsg = "verschoben in " & TrashFolderPath
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.FindTrashFolder

'---------------------------------------------------------------------------------------
' Method : Sub GetSortableItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetSortableItems(ByRef curFolder As Folder)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.GetSortableItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Filter As String
Dim i As Long
Dim sI As Object
    If eOnlySelectedItems Then
        DoVerify False, " debug this, can it work???"
        Set sortedItems(1) = Nothing
        For i = 1 To SelectedItems.Count
            Set sI = SelectedItems.Item(i)
            sortedItems(1).Add sI    ' no date filtering here
        Next i
    Else
        If getFolderFilter(curFolder.Items(1), CutOffDate, Filter, ">=") _
        Then
            Set sortedItems(1) = curFolder.Items.Restrict(Filter)
        Else
            Set sortedItems(1) = curFolder.Items
        End If
        Set sI = sortedItems(1).Item(1)
    End If
    If sI.Class = olAppointment Then   ' works only if sorted by [Start]
        sortedItems(1).IncludeRecurrences = True
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.GetSortableItems

'---------------------------------------------------------------------------------------
' Method : Function getTrashFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getTrashFolder(ByVal selectedFolder As Folder, FullFolderPath As String) As Folder
Dim zErr As cErr
Const zKey As String = "DupeDeleter.getTrashFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim curFolder As Folder
Dim DuplikateFehlt As Boolean
Dim MatchedFolder As Boolean
Dim goingUp As Boolean
Dim ShortName As String

    DuplikateFehlt = False
    goingUp = False
    If selectedFolder Is Nothing Then
        Set selectedFolder = Folder(1)
    End If
    If selectedFolder.Parent Is Nothing Then
        GoTo multiFoldertable   ' no parent, => no kid-Folders
    End If
    On Error GoTo goUpOneLevel
    Do While selectedFolder.Folders.Count = 0
goUpOneLevel:
        If selectedFolder.Parent Is Nothing Then
            Set selectedFolder = Folder(1)
            If selectedFolder.Parent Is Nothing Then
multiFoldertable:
                Set selectedFolder = Session.GetDefaultFolder(olFolderDeletedItems)
            End If
        End If
        If selectedFolder.Parent.Class = aNameSpace Then
            Exit Do
        End If
        Set selectedFolder = selectedFolder.Parent
    Loop
    
tryrelaxed:
    Call ErrReset(0)
    For Each curFolder In selectedFolder.Folders
        ShortName = Left(curFolder.Name, 4)
        MatchedFolder = (InStr(ShortName, "Dele") > 0 _
                      Or InStr(ShortName, "Gelö") > 0 _
                      Or InStr(ShortName, "Tras") > 0)
            
        If (ShortName = "Dupl" And Not IsComparemode) _
        Or DuplikateFehlt And MatchedFolder Then
            Set TrashFolder = curFolder
            GoTo gotOne
        Else
            If ShortName = "Sync" Then
                GoTo skpToNextFolder
            End If
            If DuplikateFehlt And curFolder.Folders.Count > 0 Then
                Set getTrashFolder = getTrashFolder(curFolder, FullFolderPath & "\" & curFolder.Name)
                If getTrashFolder Is Nothing Then
                    goingUp = True
                    GoTo goUpOneLevel ' so we go up one Folder level
                Else
                    GoTo gotOne
                End If
            End If
        End If
    
skpToNextFolder:
    Next curFolder
    
    DuplikateFehlt = Not DuplikateFehlt
    If DuplikateFehlt Then
        GoTo tryrelaxed
    End If
    If Not goingUp Then
        goingUp = True
        GoTo goUpOneLevel
    End If
gotOne:
    If Not TrashFolder Is Nothing Then
        TrashFolderPath = TrashFolder.FolderPath
    End If
    Set getTrashFolder = TrashFolder

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' DupeDeleter.getTrashFolder

'---------------------------------------------------------------------------------------
' Method : Sub Initialize_UI
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Initialize_UI()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.Initialize_UI"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim otherFolder As Folder
Dim j As Long
    WantConfirmationThisFolder = WantConfirmation
    MaxMisMatchesForCandidates = MaxMisMatchesForCandidatesDefault
    If LenB(ActionTitle(UBound(ActionTitle))) = 0 Then
        Call SetStaticActionTitles
    End If
    Message = vbNullString
    killType = "Vergleichen, Löschung oder Verschiebung"    ' options available on Item
    killMsg = "wird noch ermittelt"
    
    Message = ActionTitle(ActionID)
    
    ' Plausis for "parameters"
    
    If LF_CurLoopFld Is Nothing Then
        If SelectedItems Is Nothing Then
            If ActiveExplorer.Selection.Count = 0 Then
                DoVerify False
            End If
            Set SelectedItems = New Collection
            For j = 1 To ActiveExplorer.Selection.Count
                SelectedItems.Add ActiveExplorer.Selection.Item(j)
            Next j
        End If
        Set LF_CurLoopFld = getParentFolder(SelectedItems.Item(1))
        If Folder(1) Is Nothing Then
            Set Folder(1) = LF_CurLoopFld
        End If
        Set otherFolder = Folder(1)
        If LF_CurLoopFld.FolderPath <> otherFolder.FolderPath Then
            DoVerify False
            Set LF_CurLoopFld = otherFolder
        End If
        curFolderPath = LF_CurLoopFld.FolderPath
        eOnlySelectedFolder = True
    End If
    If eOnlySelectedFolder Then
        If eOnlySelectedItems Then
            DoVerify False, "bad combi"
            eOnlySelectedFolder = False
            Set LF_CurLoopFld = Nothing
            GoTo selOnly
        End If
        Set Folder(1) = LF_CurLoopFld
    ElseIf eOnlySelectedItems Then
selOnly:
        If SelectedItems Is Nothing Then
            DoVerify False
        ElseIf SelectedItems.Count = 0 Then
            DoVerify False
        End If
    End If
    If UI_DontUseDel Or Not UI_DontUse_Sel Then
        Call LogEvent("Using Standard Deletion Parameters")
        If UI_DontUse_Sel Then                                  ' implies UI_DontUseDel = True
            Call LogEvent("Using Standard Selection Parameters")
        Else
            Call DisplayParameters(2)
        End If
    Else
        Call DisplayParameters(1)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.Initialize_UI

'---------------------------------------------------------------------------------------
' Method : Sub DisplayParameters
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DisplayParameters(form As Long)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.DisplayParameters"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim CountOfItems As Long
Dim Pad As Long
Dim ReShowFrmErrStatus As Boolean

    If frmErrStatus.Visible Then
        Call ShowOrHideForm(frmErrStatus, False)
        ReShowFrmErrStatus = True
    End If
    
    Set LBF = Nothing
    b3text = "Abbruch"
    Select Case form
    Case 1: Set LBF = New frmDelParms
    Case 2: Set LBF = New frmSelParms
    Case Else
        DoVerify False
    End Select
    
    If eOnlySelectedFolder Then
        ' just process the entire Selected Folder
        Message = Message & ": " & Quote(curFolderPath) & "  Anzahl Items: " _
                & CountOfItems & vbCrLf
        CountOfItems = LF_CurLoopFld.Items.Count
    ElseIf eOnlySelectedItems Then
        CountOfItems = SelectedItems.Count
        Message = Message & b & CountOfItems _
            & " selektierte Items in Ordner "
        If DateSkipCount > 0 Then
            Message = Message & vbCrLf & "     " & DateSkipCount _
                    & " Items wegen Datumsfilter ausgeschlossen"
        End If
        Pad = Len(Message)
        Message = Message _
            & Quote(LF_CurLoopFld.FolderPath)
    Else
        CountOfItems = LF_CurLoopFld.Items.Count
        Message = Message & " für " & Quote(curFolderPath)
    End If
       ' working on 2 different Folders?
    If Not Folder(2) Is Nothing Then    ' yes, 2 Folders, Show msg
        If LF_CurLoopFld.FolderPath <> Folder(2).FolderPath Then
            Message = Message & vbCrLf & "und" _
                & String(Pad + 10, b) & b _
                & Quote(Folder(2).FolderPath) & "  (enthält " _
                & Folder(2).Items.Count & " Items)"
        End If
    End If
    
    If eOnlySelectedItems Then
        bDefaultButton = "Go"
        b1text = bDefaultButton
        b2text = vbNullString                         ' hidden button
        If form = 1 Then
            LBF.Frame2.Visible = False
            killType = "Vergleich ohne geplante Aktionen"
            LBF.LPWantConfirmationThisFolder.Visible = False
            LBF.LPWantConfirmation.Visible = False
        End If
        LBF.Controls("Go").Caption = "Go"
        LBF.Controls("Go").Default = True
        LBF.Controls("bDebugStop").Visible = True
    Else            ' Folders loop or selected items only
        bDefaultButton = "Go"
        b1text = "Prüfen"
        b2text = "Übergehen"
        If form = 1 Then
            LBF.LPWantConfirmationThisFolder.Visible = True
            LBF.LPWantConfirmation.Visible = True
        End If
        If eOnlySelectedItems Then
            Message = SelectedItems.Count _
                & " selektierte Items werden verglichen"
            LBF.Frame2.Visible = False
        ElseIf eOnlySelectedFolder Then
            ' no ops, message already set
        Else
            Message = Message & " (Ordner " & LF_recursedFldInx _
                    & " von  " _
                    & LookupFolders.Count & " auf dieser Ebene)"
            LBF.Frame2.Visible = True
        End If
        LBF.Controls("bDebugStop").Visible = True
    End If
    
    If form = 1 Then
        If eOnlySelectedItems Or eOnlySelectedFolder Then
            LBF.LPAskEveryFolder.Visible = False
        Else
            LBF.LPAskEveryFolder.Visible = True
        End If
    End If
    
    If LF_CurLoopFld.Items.Count > 0 Then
        Call ShowOrHideForm(LBF, True)
    Else
        If Not LF_CurLoopFld Is Nothing Then
            Call LogEvent("         No items in Folder " _
                & LF_CurLoopFld.FullFolderPath, eLall)
        End If
    End If
    If Not LBF Is Nothing Then
        Set LBF = Nothing
    End If

FuncExit:
    If ReShowFrmErrStatus Then
        Call ShowOrHideForm(frmErrStatus, True)
    End If

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.DisplayParameters

'---------------------------------------------------------------------------------------
' Method : Sub initializeComparison
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub initializeComparison(Optional DecodingStatusOK As Boolean)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.initializeComparison"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Matches = 0
    cMisMatchesFound = 0
    MatchPoints(1) = 0
    MatchPoints(2) = 0
    IgnoredPropertyComparisons = 0
    NotDecodedProperties = 0
    SimilarityCount = 0
    SuperRelevantMisMatch = False
    DecodingStatusOK = False
    mustDecodeRest = False
    AllItemDiffs = vbNullString
    DiffsIgnored = vbNullString
    DiffsRecognized = vbNullString
    MatchData = vbNullString
    Message = vbNullString
    fiBody(1) = vbNullString
    fiBody(2) = vbNullString
    OneDiff = vbNullString
    AttributeIndex = 0
    If aID(1).idAttrDict.Count = 0 Then
        AttributeUndef(1) = 0
        AllPropsDecoded = False
    End If
    If aID(2).idAttrDict.Count = 0 Then
        AttributeUndef(2) = 0
        AllPropsDecoded = False
    End If
    MaxPropertyCount = Max(aID(1).idAttrDict.Count, aID(2).idAttrDict.Count)
    
    Set killWords = Nothing
    Set killWords = New Collection
    killWords.Add "*@*"         ' do not compare email adresses in body etc.
    killWords.Add "*aspx?*"     ' do not compare dynamic HTML in body

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.initializeComparison

'---------------------------------------------------------------------------------------
' Method : Sub initializeExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub initializeExcel()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.initializeExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If xlApp Is Nothing Then
        Call XlgetApp
        GoTo OisN
    ElseIf O Is Nothing Then
OisN:
        If displayInExcel Then
            Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl)
            ' open default but don't Show it
        End If
    ElseIf O.xlTabIsEmpty <> 1 Then     ' excel open, set workbook empty
        Call ClearWorkSheet(xlA, O)     ' previous workbook is no longer relevant if there is one
        O.xHdl = sHdl
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.initializeExcel

'---------------------------------------------------------------------------------------
' Method : Sub InitsForPropertyDecoding
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitsForPropertyDecoding(doingTheRest As Boolean)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.InitsForPropertyDecoding"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    saveItemNotAllowed = False
    If aID(1).idAttrDict Is Nothing Then
        Set aID(1).idAttrDict = New Dictionary
        AttributeUndef(1) = 0
    End If
    If aID(2).idAttrDict Is Nothing Then
        Set aID(2).idAttrDict = New Dictionary
        AttributeUndef(2) = 0
    End If
    If aID(1).idAttrDict.Count > 0 Or _
       aID(2).idAttrDict.Count > 0 Then ' decoding has started before:
        If aID(2).idAttrDict.Count = 0 Then  ' *** adjust line len below
            mustDecodeRest = doingTheRest _
                           Or Not (AllPropsDecoded Or quickChecksOnly)
        ElseIf aID(1).idAttrDict.Count = aID(2).idAttrDict.Count Then
            mustDecodeRest = doingTheRest _
                           Or Not (AllPropsDecoded Or quickChecksOnly)
        Else    ' properties of both items are not all specified
            mustDecodeRest = True
            UserDecisionRequest = True
        End If
        DoVerify Not (AllPropsDecoded And mustDecodeRest)
        If AllPropsDecoded And mustDecodeRest Then
            AllPropsDecoded = False
        End If
        GoTo ProcReturn
    End If
   
    mustDecodeRest = False    ' Kompletter Neuanfang, initially only most important
    AllPropsDecoded = False
    UserDecisionRequest = False
    If aID(1).idAttrDict.Count > 0 Then
        Set aID(1).idAttrDict = New Dictionary
        AttributeUndef(1) = 0
    End If
    If aID(2).idAttrDict.Count > 0 Then
        Set aID(2).idAttrDict = New Dictionary
        AttributeUndef(2) = 0
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.InitsForPropertyDecoding

'---------------------------------------------------------------------------------------
' Method : Sub logDiffInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub logDiffInfo(Text As String)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.logDiffInfo"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If Left(Text, 3) = "## " Then
            MatchData = MatchData & vbCrLf & Text
        Call LogEvent(Text)
    Else
        Call LogEvent(Text, eLmin)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.logDiffInfo

'---------------------------------------------------------------------------------------
' Method : Sub logMatchInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub logMatchInfo()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.logMatchInfo"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Message = "+++ Resumee: " & objTypName _
                & " items " & WorkIndex(1) & "/" _
                & WorkIndex(2) & vbCrLf & Message
    If fiMain(1) <> fiMain(2) Then
        Message = Message & vbCrLf _
        & " +++ Objekte haben unterschiedliche Hauptidentifikationen"
    Else
        Message = Message & vbCrLf & fiMain(1) & vbCrLf & " +++ Ende "
    End If
    Call LogEvent(Message)
    Call LogEvent(MatchData, eLnothing)
    Call LogEvent(String(Len(Message), "="))

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.logMatchInfo

'---------------------------------------------------------------------------------------
' Method : Sub NoDupes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub NoDupes()    ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "DupeDeleter.NoDupes"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="DupeDeleter")
    
    ActionID = atDoppelteItemslöschen
    
    IsEntryPoint = True
        
    ' set dynamic headline
    sHdl = "CritPropName---------------" _
            & " Objekt-" & WorkIndex(1) & "----------------------" _
            & " Objekt-" & WorkIndex(2) & "----------------------" _
            & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    AcceptCloseMatches = True
    quickChecksOnly = Not AcceptCloseMatches
    AskEveryFolder = True
    WantConfirmation = True
    MatchMin = 1000
    IsComparemode = False
    SelectOnlyOne = False
    eOnlySelectedItems = False
    StopLoop = False
    PickTopFolder = True
    Set SelectedItems = New Collection
    
    bDefaultButton = "Go"
    ' look for the best suitable DftItemClass and its rules
    Call BestObjProps(Folder(1), withValues:=False)
    If eOnlySelectedFolder Then
        ' Just do this one Folder / subFolders thereof
        Set LF_CurLoopFld = Folder(1)
        Call Initialize_UI              ' displays options dialogue
        Call CheckOneFolder
    Else                                ' loop lookup Folders
        For FolderLoopIndex = 1 To LookupFolders.Count
            Set LF_CurLoopFld = LookupFolders(FolderLoopIndex)
            Call Initialize_UI          ' displays options dialogue
            Call CheckOneFolder
        Next
    End If
done:
    If TerminateRun Then
        GoTo ProcReturn
    End If
    StopRecursionNonLogged = False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.NoDupes

'---------------------------------------------------------------------------------------
' Method : Sub CheckOneFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CheckOneFolder()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.CheckOneFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Select Case rsp
    Case vbYes
        If eOnlySelectedItems Then
            FolderLoopIndex = LookupFolders.Count    ' do not loop Folder list if Selection
            DoVerify False, " this code used as model only"
            GoTo likepicked
        ElseIf PickTopFolder Then
            FolderLoopIndex = LookupFolders.Count    ' do not loop Folder list if Folder is picked
            If Folder(1) Is Nothing Then
                Call PickAFolder(1, _
                    "bitte wählen Sie den obersten Ordner für die Doublettensuche ", _
                    "Auswahl des Hauptordners für die Doublettensuche", _
                    "OK", "Cancel")
            End If
            Set topFolder = Folder(1)
likepicked:
            Call FindTrashFolder
            Set ParentFolder = Nothing
            Call Try(allowNew)                   ' Try anything, autocatch, Err.Clear
            ' no parent if top Folder, parentFolder = nothing
            Set ParentFolder = topFolder.Parent
            Catch
            curFolderPath = topFolder.FolderPath
            FullFolderPath(FolderPathLevel) = "\\" _
                            & Trunc(3, curFolderPath, "\")
            Call CheckDoublesInFolder(topFolder)        ' ########## Main Work here ##########*
        Else    ' loop Folders items , no (single) Folder was picked
            Call CheckDoublesInFolder(topFolder)        ' ########## Main Work here ##########*
            bDefaultButton = "Go"
        End If
    Case vbCancel
        Call LogEvent("=======> Stopped before processing any Folders . Time: " _
            & Now(), eLmin)
        If TerminateRun Then
            GoTo ProcReturn
        End If
        GoTo ProcReturn
    Case Else   ' loop Candidates
        Set topFolder = LookupFolders.Item(FolderLoopIndex)
        Call FindTrashFolder
    End Select
    If WantConfirmation = True Then
        Call LogEvent("=======> Confirmation mode starts. Time: " _
                & Now(), eLmin)
    Else: WantConfirmation = False
        Call LogEvent("=======> Confirmation mode ends. Time: " _
                & Now(), eLmin)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.CheckOneFolder

'---------------------------------------------------------------------------------------
' Method : Sub SaveItemsIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SaveItemsIfChanged(Optional MustConfirm As Boolean)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.SaveItemsIfChanged"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim ItemSaved As Boolean
Dim Both As String
    If saveItemNotAllowed Then
        Message = "Items sollen aufgrund der Nutzerangaben nicht gespeichert werden"
        Call LogEvent(Message, eLmin)
    Else
        If MustConfirm _
        And (WorkItemMod(1) Or WorkItemMod(2)) Then
            If WorkItemMod(1) And WorkItemMod(2) Then
                Both = "beiden "
                LF_ItmChgCount = LF_ItmChgCount + 2
            Else
                LF_ItmChgCount = LF_ItmChgCount + 1
            End If
            If fiMain(1) = fiMain(2) Then
                rsp = MsgBox("Änderungen an " & Both & "Items " & Quote(fiMain(1)) _
                    & " bestätigen", vbYesNoCancel, ActionTitle(ActionID))
            Else
                rsp = MsgBox("confirm " & Both & "changes to " & Quote(fiMain(1)) _
                    & "  and " & Quote(fiMain(2)) & b, vbYesNoCancel, _
                        ActionTitle(ActionID))
                If rsp = vbNo Then
                    saveItemNotAllowed = True
                End If
            End If
        End If
       If CurIterationSwitches.SaveItemRequested And Not saveItemNotAllowed Then
            For i = 1 To 2
                If WorkItemMod(i) Then
                    WorkItemMod(i) = False
                    Err.Clear
                    aBugTxt = "save item #" & i & b & fiMain(i)
                    Call Try
                    aID(i).idObjItem.Save
                    If Catch Then
                        Message = "Item changes NOT saved in " _
                            & Quote(aID(i).idObjItem.Parent.FullFolderPath) _
                            & " for " & Quote(fiMain(i)) & ": " _
                            & Err.Description
                        ' NO! this could cause delete query: ItemSaved = False
                    Else
                        Message = "Item changes successfully saved in " _
                        & Quote(aID(i).idObjItem.Parent.FullFolderPath) _
                        & " for " & Quote(fiMain(i))
                        ItemSaved = True
                    End If
                    Call LogEvent(Message, eLall)
                End If
            Next i
        End If
    End If
    CurIterationSwitches.SaveItemRequested = ItemSaved  ' report to the outside world: we did save

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.SaveItemsIfChanged

'---------------------------------------------------------------------------------------
' Method : Sub ScanItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ScanItem(itemNo As Long, oneItem As Object) ' like nodupes without compare
Dim zErr As cErr
Const zKey As String = "DupeDeleter.ScanItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Propertycount As Long
    WorkIndex(1) = itemNo
    OnlyMostImportantProperties = False
    
    Set oneItem = GetAobj(1, WorkIndex(1))
    objTypName = DecodeObjectClass(getValues:=True)
    Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " & objTypName & b _
        & curFolderPath & ": " _
        & WorkIndex(1) _
        & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))
    IgnoredPropertyComparisons = 0
    NotDecodedProperties = 0
    MatchData = vbNullString
    Message = vbNullString
    saveItemNotAllowed = False
    If xUseExcel Or xDeferExcel Or O Is Nothing Then
        Call XlgetApp
        Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl, showWorkbook:=DebugMode)
    End If
    If O.xlTabIsEmpty > 1 Then  ' excel open, workbook empty
        Call ClearWorkSheet(xlA, O)   ' previous workbook is no longer relevant
    End If
    ' *** Set aItemProps(1) = aID(1).idObjItem.ItemProperties ' done inside GetItemAttrDscs
    Propertycount = aID(1).idObjItem.ItemProperties.Count
    If TotalPropertyCount = 0 Then
        TotalPropertyCount = Propertycount
    ElseIf TotalPropertyCount <> Propertycount Then
        DoVerify False, " more analysis needed here."
        TotalPropertyCount = Propertycount
    End If
    aID(1).idAttrDict = New Dictionary
    AttributeUndef(1) = 0
    SelectOnlyOne = True ' not working with tuples
    Call SetupAttribs(oneItem, 1, True)
    If AttributeIndex = 0 Then
        Message = "Skipping item " & WorkIndex(1) _
                    & ", can not process " _
                    & aOD(aPindex).objItemClassName _
                    & " (" & aID(1).idObjItem & ") ," _
                    & Message
        Call LogEvent(Message, eLall)
        GoTo ProcReturn ' no can do
    End If
    NotDecodedProperties = Propertycount - AttributeIndex
    
    If NotDecodedProperties > 0 Then
        pArr(1) = "*** es wurden nicht alle Merkmale untersucht"
        Call addLine(O, Propertycount + 3, pArr)
        AllItemDiffs = AllItemDiffs & vbCrLf & pArr(1)
        IgnoredPropertyComparisons = IgnoredPropertyComparisons _
                            + NotDecodedProperties
    End If
    displayInExcel = xUseExcel Or xDeferExcel
    If displayInExcel Or (WorkItemMod(1) And CurIterationSwitches.SaveItemRequested) Then
        rsp = vbYes
        UserDecisionRequest = False
        If displayInExcel Then
            Call DisplayWithExcel(vbNullString)
        Else
            Call DisplayWithoutExcel(vbNullString)
        End If
        If rsp = vbCancel Then
            Call ClearWorkSheet(xlA, O)     ' erase this one, new one will be started when needed
            xUseExcel = False
            rsp = vbNo
        End If
        If rsp = vbYes And CurIterationSwitches.SaveItemRequested And Not saveItemNotAllowed Then
            Call PerformChangeOpsForMapiItems
            Call Try                         ' Try anything, autocatch
            aID(1).idObjItem.Save
            If Catch Then
                Message = "Item changes NOT saved for " _
                    & aID(1).idObjItem.Subject _
                    & "   " & Err.Description
            Else
                Message = "Item changes successfully saved for " _
                    & aID(1).idObjItem.Subject
            End If
        Else
                Message = "Item NOT saved (result of user choice)"
        End If
    Else
                Message = "Item has no changes"
    End If
    Call LogEvent(Message, eLall)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.ScanItem

'---------------------------------------------------------------------------------------
' Method : Sub storeAttribute
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub storeAttribute(i As Long, sourcecol As Long, px As Long)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.storeAttribute"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim NewAttrValue As String
Dim adName As String
Dim Properties As ItemProperties
Dim selectedProperty As ItemProperty
Dim oldValue As String
Dim oi As Long

    With O.xlTSheet
        NewAttrValue = .Cells(i, sourcecol).Text
        If Left(NewAttrValue, 1) = "#" Then
            GoTo ProcReturn
        End If
        If Left(NewAttrValue, 1) = " & quote(" Then
            NewAttrValue = Mid(NewAttrValue, 2)
        End If
        
        adName = .Cells(i, 1).Text
        aBugTxt = "access to list of item properties"
        Call Try(testAll)
        ' sinngemäß: aID(pxindex).idObjItem.ADName = NewAttrValue
        Set Properties = aID(px).idObjItem.ItemProperties
        If Catch Then
            GoTo ProcReturn
        End If
        aBugTxt = "get property " & adName
        Call Try
        Set selectedProperty = Properties.Item(adName)
        If Catch Then
            MsgBox ("Attribute Name " & adName & " not accessible" _
                    & vbCrLf & Err.Description)
            .Cells(i, promptColumn).Value = "?NAC"
            GoTo ProcReturn
        End If
        
        If selectedProperty.Name = "Attachments" Then    ' we do an array here!
            If InStr(NewAttrValue, "ContactPicture") > 0 Then
                If selectedProperty.Value.Count = 0 Then
                    If px = 1 Then
                        oi = 2
                    Else
                        oi = 1
                    End If
                    DoVerify False, " incomplete here!"
                    NewAttrValue = cPfad & "Ascher, Christian (Home - 20140219).jpg"
                    If LenB(NewAttrValue) = 0 Then
                        NewAttrValue = InputBox("Type the Filename for the contact:")
                    End If
                    aBugTxt = "add picture to contact"
                    Call Try                 ' Try anything, autocatch
                    aID(px).idObjItem.AddPicture NewAttrValue
                    If Catch Then
                        GoTo FuncExit
                    End If
                    WorkItemMod(px) = True
                End If
            End If
        Else
            oldValue = selectedProperty.Value    ' not .text, we are using the value from the item
            aBugTxt = "assign new property value for " _
                        & selectedProperty.Name
            Call Try                         ' Try anything, autocatch
            selectedProperty.Value = NewAttrValue
            If Catch Then
                MsgBox ("Assignment to " & adName & " failed with message:" _
                        & vbCrLf & "   " & Err.Description)
                .Cells(i, promptColumn).Value = "'(?ERR-Could not assign to item)"
                GoTo FuncExit
            End If
            If selectedProperty.Name = MainObjectIdentification Then
                fiMain(px) = NewAttrValue & " [geänderter Wert], war " & Quote(oldValue) & b
            End If
        End If
        WorkItemMod(px) = True
        CurIterationSwitches.SaveItemRequested = True
        saveItemNotAllowed = False
        rsp = vbYes
        Call LogEvent("New value for " & adName & " = " _
                & Quote(NewAttrValue) & " on item " & px, eLall)
        .Cells(i, 15 + px).Value = Quote(oldValue)  ' from item, not from excel
        .Cells(i, 15 + px).Interior.ColorIndex = xlColorIndexNone
        .Cells(i, px + 1).Interior.ColorIndex = 35  ' light green
    End With ' O.xlTSheet

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.storeAttribute

'---------------------------------------------------------------------------------------
' Method : Function synchedNames
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function synchedNames(RrX As String, RrY As String) As Boolean
Dim zErr As cErr
Const zKey As String = "DupeDeleter.synchedNames"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
didOneSynch:
    If AttributeIndex > aID(1).idAttrDict.Count Then
        Message = "Property " _
            & Quote(aID(2).idAttrDict.Item(AttributeIndex + YleadsXby).adName) _
            & "  fehlt im 1. Item"
        GoTo PropCountError
    End If
    If AttributeIndex + YleadsXby > aID(2).idAttrDict.Count Then
        Message = "Property " & Quote(aID(1).idAttrDict.Item(AttributeIndex).adName) _
            & "  fehlt im 2. Item"
        GoTo PropCountError
    End If
    Set aDecProp(1) = aID(1).idAttrDict.Item(AttributeIndex)
    Set aDecProp(2) = aID(2).idAttrDict.Item(AttributeIndex + YleadsXby)
    RrX = aDecProp(1).adFormattedValue        ' ItemPropertyValues(1, AttributeIndex)
    RrY = aDecProp(2).adFormattedValue
    If Left(RrX, 1) = "# " Then
        If aDecProp(1).adOrigValDecodingOK Then
            RrX = aDecProp(1).adDecodedValue
        End If
    End If
    If Left(RrY, 2) = "# " Then
        If aDecProp(2).adOrigValDecodingOK Then
            RrY = aDecProp(2).adDecodedValue
        End If
    End If
    PropertyNameX = aDecProp(1).adName        ' = PropertyNames(1, AttributeIndex)
    PropertyNameY = aDecProp(2).adName        ' PropertyNames(2, AttributeIndex + YleadsXby)
    
    If LenB(PropertyNameX) = 0 Or LenB(PropertyNameY) = 0 Then
        Message = "Missing Decoded Property for entry " _
                    & AttributeIndex & " / " _
                    & AttributeIndex + YleadsXby
        If LenB(PropertyNameX) = 0 Then
            Message = Message & " Side 1"
        End If
        If LenB(PropertyNameY) = 0 Then
            Message = Message & " Side 2"
        End If
        GoTo PropCountError
    End If
    
    If PropertyNameX = PropertyNameY Then
        i = AttributeIndex
    Else
        ' first, look for innermost occurrence of PropertyNameY
        For i = AttributeIndex + 1 To aID(2).idAttrDict.Count
            If PropertyNameY = aID(2).idAttrDict.Item(i).adName Then
                YleadsXby = i - AttributeIndex
                ' AttributeIndex = i wäre falsch!
                GoTo didOneSynch
            End If
        Next i
        ' if not found here, try first occ. of PropertynameX on side 1
        For i = AttributeIndex - 1 To 1 Step -1
            If PropertyNameX = aID(1).idAttrDict.Item(i).adName Then
                YleadsXby = AttributeIndex - i
                AttributeIndex = i
                GoTo didOneSynch
            End If
        Next i
        If i = 0 Then
            Message = "keine vergleichbaren Attribute?!"
        End If
PropCountError:
        If DebugMode Then
            DoVerify False
        End If
        saveItemNotAllowed = True
        Call logMatchInfo
        synchedNames = True
    End If
    If i > 0 Then
        Call GetAttrDsc(aID(1).idAttrDict.Item(i).adKey)
    Else
        Set aTD = Nothing
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' DupeDeleter.synchedNames

'---------------------------------------------------------------------------------------
' Method : Sub TrashOrDeleteItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub TrashOrDeleteItem(DelObjectItem As cDelObjectsEntry)
Dim zErr As cErr
Const zKey As String = "DupeDeleter.TrashOrDeleteItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim delObj As Object
Dim GoneObj As Object
    With DelObjectItem
        If .DelObjInd Then
            Set delObj = sortedItems(.DelObjPos).Item(.DelObjPindex)
        Else
            Set delObj = SelectedItems(.DelObjPos)
        End If
    End With ' DelObjectItem
    'On Error Resume Next
    If LenB(TrashFolderPath) = 0 Then
hardDelete:
        aBugTxt = "deleting Item"
        Call Try(testAll)
        delObj.Delete
        If CatchNC Then
            GoTo errHandling
        End If
        Message = "Endgültige Löschung war erfolgreich"
    Else
        Message = "Löschung in den Mülleimer versucht"
        Set GoneObj = delObj.Move(TrashFolder)
errHandling:
        If E_Active.errNumber = -2147467259 Then
            rsp = MsgBox(Err.Description & vbCrLf & vbCrLf _
                & "Wollen Sie " & Quote(delObj) _
                & "  endgültig löschen?" _
                & vbCrLf & vbCrLf _
                & "Wenn Sie die ganze Serie löschen wollen, " _
                & "brechen Sie jetzt ab und wählen Sie die Einträge " _
                & "in der Listenansicht des Kalenders aus." _
                , vbYesNo, "Fehler beim Verschieben nach " _
                & TrashFolderPath)
            GoTo doDecide
        ElseIf Catch Then
            rsp = MsgBox(Err.Description & vbCrLf & vbCrLf _
                & "Wollen Sie " & Quote(delObj) & "  endgültig löschen?", vbOKCancel, _
                "Fehler beim Verschieben nach " & TrashFolderPath)
doDecide:
            If rsp = vbOK Then
                Message = Replace(Message, "verschoben", _
                    "endgültig gelöscht statt verschoben")
                GoTo hardDelete
            Else
                Message = "Endgültige Löschung war nicht erfolgreich"
            End If
        Else
            Message = "Löschung erfolgte in " & GoneObj.Parent.FolderPath
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.TrashOrDeleteItem

'---------------------------------------------------------------------------------------
' Method : Sub UnSelectItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub UnSelectItems()
Dim zErr As cErr
Const zKey As String = "DupeDeleter.UnSelectItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim A As Object
    If eOnlySelectedItems Then
        For i = sortedItems(1).Count To 1 Step -1
            Set A = sortedItems(1).Item(i)
            A.managerName = vbNullString ' clear unused field
            A.Save
        Next
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)
End Sub ' DupeDeleter.UnSelectItems


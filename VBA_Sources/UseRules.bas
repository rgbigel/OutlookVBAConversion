Attribute VB_Name = "UseRules"
Option Explicit

Const txtTitel As String = "EMail-Notiz:"
Const WantedFolders As String = _
"\Gese+ \Post* -\Gelö+ -\Junk* -\Not+ -Auf+ -\Un+ -onta* -alend* \\+"

Public LogFolder As Folder
' Public SearchFoldersFolder As Folders
Public SearchObject As Search

Public FilterMain As cRuleFilter
Public FilterPaths As cRuleFilter
Public FilterRules As cRuleFilter

Dim objMail As MailItem
Dim objNote As NoteItem
Dim aI As Inspector

Dim objNotesFolder As Folder
Dim objMailNotesFolder As Folder
Dim objArchiveFolder As Folder
Dim objArchiveNotesFolder As Folder
Dim objArchiveMailNotesFolder As Folder

Dim strEntryID As String
Dim strSubject As String
Dim itemClass As Long
Dim NotesName As String
Dim Relinked As Long

'---------------------------------------------------------------------------------------
' Method : Sub AddNewNote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AddNewNote()
Dim zErr As cErr
Const zKey As String = "UseRules.AddNewNote"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call GetMailNotesFolder
    If GetMatchingItems() Then
        If Not objNote Is Nothing Then
            MsgBox ("Notiz schon vorhanden")
            GoTo ProcReturn
        End If
        If LenB(objMail.VotingOptions) = 0 Then
            Call AddNote
        Else
            Call eMailNoteOps
        End If
    Else
        If objMail Is Nothing Then
            MsgBox ("keine eMail gewählt")
        Else
            Call AddNote
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.AddNewNote

'---------------------------------------------------------------------------------------
' Method : Sub AddNote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AddNote()
Dim zErr As cErr
Const zKey As String = "UseRules.AddNote"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    objMail.FlagIcon = olBlueFlagIcon
    objMail.FlagRequest = "Zur Nachverfolgung"
    Set objNote = objMailNotesFolder.Items.Add
    With objNote
        .Categories = strEntryID
        .Body = "Notiz zu: " & objMail.Subject & vbCrLf
        .width = 500
        .height = 250
        .Color = olBlue
        .Save
        .Display
    End With                                     ' objNote
    objMail.VotingOptions = objNote.EntryID
    Call setItmCats(objMail, "anotiert", LOGGED)
    objMail.Save

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.AddNote

'---------------------------------------------------------------------------------------
' Method : Sub ClearEmailNote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClearEmailNote()
Dim zErr As cErr
Const zKey As String = "UseRules.ClearEmailNote"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call GetMailNotesFolder
    GetMatchingItems
  
    If objNote Is Nothing Then
        MsgBox "Keine zu löschenden eMail-Notizen zu " & vbCrLf & B2 _
             & Quote(strSubject) & vbCrLf & "gefunden ...", _
               vbOKOnly + vbInformation, txtTitel
    Else
        objMail.FlagIcon = olNone
        objMail.FlagRequest = vbNullString
        objMail.VotingOptions = vbNullString
        Call setItmCats(objMail, vbNullString, LOGGED & "; anotiert")
        objMail.Save
        objNote.Delete
        Set aItmDsc.idObjItem = Nothing
        MsgBox "eMail-Notiz gelöscht zu" & vbCrLf & b & Quote(strSubject), _
        vbOKOnly + vbInformation, txtTitel
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.ClearEmailNote

'---------------------------------------------------------------------------------------
' Method : Function CountItemsIn
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CountItemsIn(curObj As Object, LoopProgress As Long) As Long
Dim zErr As cErr
Const zKey As String = "UseRules.CountItemsIn"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If curObj Is Nothing Then
        GoTo ProcReturn
    End If
    If TypeName(curObj) = "Selection" _
    Or TypeName(curObj) = "Collection" Then
        CountItemsIn = curObj.Count
        LoopProgress = 0
        eOnlySelectedItems = True
    Else
        CountItemsIn = curObj.Items.Count
        eOnlySelectedItems = False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.CountItemsIn

'---------------------------------------------------------------------------------------
' Method : Sub CreateRules
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CreateRules()                                ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "UseRules.CreateRules"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="UseRules")

Dim oStore As Outlook.Store
Dim oRule As Outlook.Rule
Dim myfolder As Outlook.Folder
Dim colRules As Outlook.Rules
Dim oCondition As Outlook.RuleCondition
Dim oRuleAction As Outlook.RuleAction
Dim i As Long
Dim Name As String
Dim pName As String
    If LookupFolders Is Nothing Then
        DoVerify False
        IsEntryPoint = True
        EPCalled = True
    End If
    
    Set myfolder = ActiveInspector.Parent
    Set oStore = myfolder.Store
    Set colRules = oStore.GetRules
    Name = "unerwünscht (nur per Makro)"
    If RuleExists(colRules, Name, precise:=False, RuleIndex:=i) Then
        DoVerify False                           ' can, "W.xlTSheet create existing rule"
        colRules.Remove i
    Else
        Set oRule = colRules.Create(Name, olRuleReceive)
        Select Case Name
        Case "unerwünscht (nur per Makro)"
            pName = Trunc((Name), 1, b)
            'Specify the condition in a xxRuleCondition object
            Set oCondition = oRule.Conditions.MessageHeader
            oCondition.Enabled = True
            'Specify the action in a MoveOrCopyRuleAction object
            Set oRuleAction = oRule.Actions.MoveToFolder
            ' Action is to move the message to the target Folder
            oRuleAction.Enabled = True
            oRuleAction.Folder = GetFolderByName(pName)
        Case Else
            DoVerify False                       ' don, "W.xlTSheet know rule yet"
        End Select
    End If
    If EPCalled Then
        Call TerminateApp
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.CreateRules

'---------------------------------------------------------------------------------------
' Method : Sub eMailNoteOps
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub eMailNoteOps()
Dim zErr As cErr
Const zKey As String = "UseRules.eMailNoteOps"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If GetMatchingItems() Then                   ' beschafft die strEntryId der geöffneten eMail
    
        If objNote Is Nothing Then               ' neue Notiz erstellen zur selektierten email
            Call AddNewNote
        End If
    ElseIf objMail Is Nothing Then
        rsp = MsgBox("Keine passende, aktuelle eMail zur " _
                   & vbCrLf & B2 & Quote(strSubject) & vbCrLf _
                   & "gefunden. Neu anlegen=Ja, in allen eMail-Ordnern suchen (dauert länger)=Nein?", _
                     vbYesNoCancel, txtTitel)
        If rsp = vbNo Then
            If Not FindPartnerMail(aNameSpace.Folders, WantedFolders) Then
                rsp = MsgBox("Keine passende eMail zur " & vbCrLf & B2 _
                           & Quote(strSubject) & vbCrLf & _
                             "gefunden (in keinem der eMail-Ordner) ...", _
                             vbOKOnly + vbInformation, txtTitel)
            End If
        End If
        If rsp = vbYes Then
            Call AddNewNote
        End If
    ElseIf objNote Is Nothing Then
        If LenB(objMail.VotingOptions) = 0 Then
            Call AddNewNote
        Else
            If FindPartnerNote(objMailNotesFolder) Then
                DoVerify False
            Else
                rsp = MsgBox("Notiz zu " & vbCrLf & B2 & Quote(strSubject) & vbCrLf _
                           & "nicht am erwarteten Ort. Neu anlegen=Ja" & vbCrLf _
                           & "in allen Notizen-Ordnern suchen (dauert länger)=Nein?", _
                             vbYesNoCancel, txtTitel)
                If rsp = vbNo Then
                    DoVerify False
                End If
                If rsp = vbYes Then
                    Call AddNewNote
                End If
            End If
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.eMailNoteOps

'---------------------------------------------------------------------------------------
' Method : Sub EMailNotes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub EMailNotes()
Dim zErr As cErr
Const zKey As String = "UseRules.EMailNotes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call GetMailNotesFolder
    Call eMailNoteOps

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.EMailNotes

'---------------------------------------------------------------------------------------
' Method : Sub EMailNotesCleanup
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub EMailNotesCleanup()
Dim zErr As cErr
Const zKey As String = "UseRules.EMailNotesCleanup"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
 
Dim cnt&
 
    GetMailNotesFolder
  
    
    For Each objNote In objMailNotesFolder.Items
        strEntryID = objNote.Categories
        aBugTxt = "find note for EntryID=" & strEntryID
        Call Try
        Set objMail = aNameSpace.GetItemFromID(strEntryID)
        If Catch Then
            If notFoundInArchive() Then
                objNote.Delete
                If Not Catch Then
                    cnt = cnt + 1
                End If
            Else
                DoVerify False, "*** Note Item is not in archive"
            End If
        End If
    Next objNote
  
    Beep
    If cnt > 0 Then
        MsgBox "Es wurde(n) " & CStr(cnt) & _
                                          " nicht mehr benötigte eMail-Notize(n) " & _
                                          "gelöscht...", vbOKOnly + vbInformation, txtTitel
    Else
        MsgBox "Keine zu löschenden eMail-Notizen " & _
               "gefunden...", vbOKOnly + vbInformation, txtTitel
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.EMailNotesCleanup

'---------------------------------------------------------------------------------------
' Method : Sub ExecuteDefinedRule
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ExecuteDefinedRule(inFolderName As String, RuleName As String, Optional inFolder As Folder)
Dim zErr As cErr
Const zKey As String = "UseRules.ExecuteDefinedRule"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="rule=" & RuleName)

Static oStore As Outlook.Store
Static oRule As Outlook.Rule
Dim myfolder As Outlook.Folder

Dim i As Long
Dim j As Long
Dim colRules As Outlook.Rules
Dim FullFolderPath As String
Static fin As cKnownFolders
Dim reUse As Boolean
Dim translatedFolderName As Variant
Dim msg As String
    
    If LookupFolders Is Nothing Then
        DoVerify False
        IsEntryPoint = True
    End If
    
    If inFolder Is Nothing Then
        If fin Is Nothing Then
            Set fin = New cKnownFolders
        End If
    Else
        Set myfolder = inFolder
        Set oStore = inFolder.Store
        GoTo folderPassed
    End If
    
    reUse = fin.InRuleFolderColl(inFolderName)
    With fin.KnownFolderItem
        If reUse Then                            ' Folder to apply rules is reusable
            If .FolderOK < 0 Then
                If .FInFolderWarning = -1 Then
                    .FInFolderWarning = -2       ' warn only once
                    Call LogEvent("Folder " & .FInFolderstring _
                                & " previously marked as bad, kann Regel " _
                                & RuleName & " nicht ausführen")
                End If
                GoTo FuncExit
            ElseIf .FolderOK = 1 Then
                If .FInFolderWarning = -1 Then
                    .FInFolderWarning = -2       ' warn only once
                    Call LogEvent("Skipping rule for bad Folder: " & .FInFolderstring, eLmin)
                End If
            Else
                Call LogEvent("Using previously found Folder " _
                            & .FInFolderstring & " again", eLmin)
            End If
        Else
            .FInFolderstring = inFolderName
            .FolderIsNormal = 0
            If LenB(inFolderName) = 0 Then       ' try to use the current Folder
                Set myfolder = ActiveInspector.Parent
                FullFolderPath = myfolder.FolderPath
            Else                                 ' use this Folder, but no search Folders (they do not work rules)
                If inFolderName = "unerwünscht" Then
                    For Each translatedFolderName _
                        In Array("unerwünscht", "Junk-E-Mail", "Spam", "Junk")
                        Set myfolder = GetFolderByName( _
                                       CStr(translatedFolderName), noSearchFolders:=True)
                        If Not myfolder Is Nothing Then
                            inFolderName = translatedFolderName
                            Exit For
                        End If
                    Next translatedFolderName
                Else
                    Set myfolder = GetFolderByName(inFolderName, noSearchFolders:=True)
                End If
            End If
            If myfolder Is Nothing Then
                Call LogEvent("Keinen normalen Ordner namens " & Quote(inFolderName) _
                            & " gefunden, kann Regel " _
                            & RuleName & " nicht ausführen")
                inFolderName = .FInFolderstring
                .FolderIsNormal = -1             ' not a normal Folder
                GoTo FuncExit
            End If
            Set oStore = myfolder.Store
        End If                                   ' Folder and associated store to apply rules has been found
folderPassed:
        If myfolder Is Nothing Then
            Call LogEvent("Ordner " & Quote(inFolderName) _
                        & " nicht vorhanden, kann Regel " _
                        & RuleName & " nicht ausführen")
            .FolderIsNormal = -1                 ' not a normal Folder
            GoTo FuncExit
        End If
        
        aBugTxt = "get item count in folder " & Quote(myfolder.FolderPath)
        Call Try
        If myfolder.Items.Count = 0 Then         ' no items to process: skip
            If Catch(DoMessage:=False) Then
                Call LogEvent("Fehler bei Ordner " & Quote(myfolder.FolderPath) _
                            & vbCrLf & "Regel " & RuleName _
                            & " kann nicht ausgeführt werden ")
                Call ErrReset(4)
                .FolderIsNormal = -1             ' not a normal Folder
                GoTo FuncExit
            End If
            Call LogEvent("Die Regel " & RuleName _
                        & " wird nicht ausgeführt weil keine Items in: " _
                        & Quote(myfolder.FolderPath), eLmin)
            GoTo FuncExit
        End If
        
        Set colRules = oStore.GetRules
        If colRules.Count = 0 Then
            Call LogEvent("Keine Regeln definiert, daher Regel " _
                        & Quote(RuleName) & " nicht gefunden", eLall)
        Else
            aBugTxt = "Find Rule " & RuleName
            Call Try
            Set oRule = colRules.Item(RuleName)  ' try find exact name
            If Not Catch Then
                GoTo ProcReturn
            End If
            ' search by similarity
            For i = 1 To colRules.Count
                Set oRule = colRules.Item(i)
                If InStr(1, oRule.Name, RuleName, vbTextCompare) > 0 Then
                    GoTo DoExec
                End If
            Next i
        
            Call LogEvent("Keine Regel mit passendem Namen zu " _
                        & RuleName & " gefunden." _
                        & vbCrLf & "Es sind insgesamt " & colRules.Count _
                        & " Regeln vorhanden.", eLall)
        End If
        GoTo FuncExit
DoExec:
        If oRule.Enabled Then
            If DebugLogging Then
                Debug.Print i, oRule.Name, oRule.RuleType
                For j = 1 To oRule.Actions.Count
                    If oRule.Actions.Item(j).Enabled Then
                        Debug.Print "Action", j, oRule.Actions.Item(j).ActionType
                    End If
                Next j
            End If
        Else
            If DebugLogging Then
                Debug.Print oRule.Name & " is not enabled"
            End If
            GoTo FuncExit
        End If
        
        msg = Quote(oRule.Name) & " wird ausgeführt in: " _
                              & myfolder.FullFolderPath
        Call LogEvent("Execute Rule " & msg, eLmin)
        
        aBugTxt = "Execute Rule"
        Call Try
        oRule.Execute DebugMode, myfolder
        Catch
        
        If Not fin Is Nothing Then
            .FolderIsNormal = 0                  ' good normal Folder
            .FolderIsSearch = -1
        End If
    End With                                     ' FIn.KnownFolderItem

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.ExecuteDefinedRule

'---------------------------------------------------------------------------------------
' Method : Function ExplorerSelectedItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ExplorerSelectedItem() As Long
Dim zErr As cErr
Const zKey As String = "UseRules.ExplorerSelectedItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim exp As Explorer
    Set exp = olApp.ActiveExplorer
    If exp.Selection.Count = 1 Then
        ExplorerSelectedItem = exp.Selection(1).Class
        If ExplorerSelectedItem = olMail Then
            Set objMail = exp.Selection(1)
        ElseIf exp.Selection(1).Class = olNote Then
            Set objNote = exp.Selection(1)
        Else
            ExplorerSelectedItem = 0
        End If
    Else
        ExplorerSelectedItem = 0
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.ExplorerSelectedItem

'---------------------------------------------------------------------------------------
' Method : Function FindPartnerMail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindPartnerMail(topFolder As Folders, Optional Filterlist As String) As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.FindPartnerMail"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim objFolder As Folder
    'On Error GoTo 0
    For Each objFolder In topFolder
        ' Folderpath starts with \\
        If ListChecker(Filterlist, objFolder.FolderPath) Then
            If FindPartnerMailInFolder(objFolder, Filterlist) Then
                FindPartnerMail = True
                GoTo ProcReturn
            End If
        End If
    Next objFolder

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.FindPartnerMail

'---------------------------------------------------------------------------------------
' Method : Function FindPartnerMailInFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindPartnerMailInFolder(objAnyMailFolder As Folder, Optional Filterlist As String) As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.FindPartnerMailInFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim objMapi As Object
Dim objFolder As Folder
Dim FpM As Boolean

    For Each objMapi In objAnyMailFolder.Items
        Call ErrReset(4)
        If DebugMode Then
            Call LogEvent(objMapi.Subject & b & objMapi.Class, eLall)
        End If
        If objMapi.Class = olMail Then
            Set objMail = Nothing
            Set objMail = objMapi
            If objMail.VotingOptions = objNote.EntryID Then
                FindPartnerMailInFolder = True
                GoTo ProcReturn
            End If
            If Len(objMail.Subject) > 4 _
            And InStr(objNote.Body, objMail.Subject) > 0 Then
                objMail.Display
                rsp = MsgBox("passt diese eMail?", vbYesNo, objNote.Subject)
                objMail.Close (olPromptForSave)
                If rsp = vbYes Then
                    FindPartnerMailInFolder = True
                    GoTo ProcReturn
                End If
            End If
            
        Else                                     ' objects other than mail found: quit Folder
            GoTo noMailFolder
        End If
    Next objMapi
    
noMailFolder:
    For Each objMapi In objAnyMailFolder.Folders
        Call LogEvent(CStr(objMapi) & b & Filterlist & b _
                    & objMapi.FolderPath, eLnothing + DebugMode)
        If ListChecker(Filterlist, objMapi.FolderPath) Then
            Call ErrReset(4)
            Set objFolder = objMapi
            FpM = FindPartnerMailInFolder(objFolder) Or FpM
            FindPartnerMailInFolder = FpM
        End If
    Next objMapi

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.FindPartnerMailInFolder

'---------------------------------------------------------------------------------------
' Method : Function FindPartnerNote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindPartnerNote(objAnyNotesFolder As Folder) As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.FindPartnerNote"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim objT As MailItem
    If objAnyNotesFolder Is Nothing Then
        GoTo ProcReturn                          ' returning false
    End If
    For Each objNote In objAnyNotesFolder.Items
        If objNote.Categories = strEntryID Then
            FindPartnerNote = True
            ' check reverse mapping; only OK if unique Note !
            If objMail.VotingOptions <> objNote.EntryID Then
                objMail.Display
                On Error GoTo linkerr
                Set objT = aNameSpace.GetItemFromID(objNote.Categories)
                objT.Display
                Set objMail = objT
relink:
                objMail.VotingOptions = objNote.EntryID
                objMail.Save
                Relinked = Relinked + 1
            End If
            Exit For
linkerr:
            rsp = MsgBox("Notiz zu " & vbCrLf & B2 & Quote(strSubject) & vbCrLf _
                       & "verknüpft sich nicht zur eMail. Reparieren=Ja" & vbCrLf _
                       & "Aufhören=Nein, Alternative suchen=Cancel", _
                         vbYesNoCancel, txtTitel)
            If rsp = vbNo Then
                DoVerify False
            ElseIf rsp = vbYes Then
                GoTo relink
            Else
                FindPartnerNote = False
                Exit For
            End If
            FindPartnerNote = False
            GoTo ProcReturn
        End If
    Next objNote

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.FindPartnerNote

'---------------------------------------------------------------------------------------
' Method : Sub getAccountDscriptors
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub getAccountDscriptors()
Dim zErr As cErr
Const zKey As String = "UseRules.getAccountDscriptors"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim lAccountDsc As cAccount
Dim lAccount As Account

    Set D_AccountDscs = New Dictionary
    Set oStores = olApp.Session.Stores
    Set SessionAccounts = olApp.Session.Accounts
    
    With SessionAccounts
        For i = 1 To .Count
            Set lAccount = .Item(i)
            Set lAccountDsc = New cAccount
            With lAccountDsc
                Set .aAcStore = lAccount.DeliveryStore
                E_Active.Permit = 0
                .Key = .aAcStore.DisplayName
                Call ErrReset(0)
                .aAcType = lAccount.AccountType
                D_AccountDscs.Add .Key, lAccountDsc
            End With                             ' laccountdsc
        Next i
    End With                                     ' ...Accounts
    aAccountNumber = 0                           ' no current one

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.getAccountDscriptors

'---------------------------------------------------------------------------------------
' Method : Function getAccountType
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getAccountType(ByVal FolderPath As String, Optional withType As Variant) As OlAccountType
Dim zErr As cErr
Const zKey As String = "UseRules.getAccountType"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim AccountName As String
Dim aAccount As Account
    If LenB(FolderPath) = 0 Then
        GoTo stmailAccount
    End If
    If Left(FolderPath, 2) = "\\" Then
        i = 3                                    ' if fully qualified, cut \\
    Else
        i = 1
    End If
    AccountName = Trunc(i, FolderPath, "\")
    With SessionAccounts
        ' do a fast check without loop if it matches
        If aAccountNumber > 0 And aAccountNumber <= .Count Then
            Set aAccount = .Item(aAccountNumber)
            aBugTxt = "find delivery store " & AccountName
            Call Try(allowAll)
            Catch False                          ' could not find delivery store if not registered as email?
            If aAccount.DeliveryStore.DisplayName = AccountName Then
                getAccountType = aAccount.AccountType
                If Not IsMissing(withType) Then
                    withType = getAccountTypeName(aAccountNumber)
                    End If
                    GoTo FuncExit                ' fast check successful
                End If
            End If
            Call ErrReset(0)
        
            ' loop through accounts because it did not match
            For aAccountNumber = 1 To .Count
                Set aAccount = .Item(aAccountNumber)
                aBugTxt = "find delivery store " & aAccount.DeliveryStore.DisplayName
                Call Try(allowAll)
                Catch
                If aAccount.DeliveryStore.DisplayName = AccountName Then
                    getAccountType = aAccount.AccountType
                    If Not IsMissing(withType) Then
                        withType = getAccountTypeName(aAccountNumber)
                        End If
                        GoTo FuncExit
                    End If
                Next aAccountNumber
            End With                             ' ...Accounts
            Call ErrReset(0)

stmailAccount:                                   ' Debug.Assert False    ' Folder without account
            If Not IsMissing(withType) Then
                withType = "stmailAccount"
                End If
                aAccountNumber = 0

FuncExit:
                Call N_ErrClear
                aAccountType = getAccountType

ProcReturn:
                Call ProcExit(zErr)
  
        End Function                         ' UseRules.getAccountType

            ' DISCARD, NOT USED
Function getAccountTypeName(aCC As Variant) As String
Dim zErr As cErr
Const zKey As String = "UseRules.getAccountTypeName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim acc1 As OlAccountType

    ' convert acc to a numeric
    Select Case VarType(aCC)
    Case vbString
        acc1 = getAccountType(aCC) + 1
    Case vbInteger, vbLong:
        acc1 = aCC + 1
    Case Else
        DoVerify False, " not impl"
    End Select
    getAccountTypeName = AccountTypeNames(acc1)
    
FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)
  
End Function                                     ' UseRules.getAccountTypeName

'---------------------------------------------------------------------------------------
' Method : Function GetDItem_P
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetDItem_P(Optional OnlyType As String = vbNullString) As Object
Dim zErr As cErr
Const zKey As String = "UseRules.GetDItem_P"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    ' note that this requires the IsExplorer and IsInspector functions found elsewhere
Dim obj As Object
   
    Select Case True
    Case (TypeName(olApp.ActiveWindow) = "Explorer")
        Set obj = ActiveExplorer.Selection.Item(1)
    Case (TypeName(olApp.ActiveWindow) = "Inspector")
        Set obj = ActiveInspector.CurrentItem
    End Select
    If LenB(OnlyType) > 0 Then
        If InStr(OnlyType & b, TypeName(obj) & b) > 0 Then
            Set GetDItem_P = obj
        End If
    Else
        Set GetDItem_P = obj                     ' any type possible
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.GetDItem_P

'---------------------------------------------------------------------------------------
' Method : Sub GetMailNotesFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetMailNotesFolder()
Dim zErr As cErr
Const zKey As String = "UseRules.GetMailNotesFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim aFolderName As String

    Set objNotesFolder = aNameSpace.GetDefaultFolder(olFolderNotes)
    NotesName = objNotesFolder.FullFolderPath
    CurIterationSwitches.ResetCategories = False
    If objArchiveFolder Is Nothing Then
        For i = 1 To aNameSpace.Folders.Count
            Set objArchiveFolder = aNameSpace.Folders(i)
            If Left(objArchiveFolder.Name, 6) = "Archiv" Then
                Exit For
            End If
        Next i
        If i > aNameSpace.Folders.Count Then     ' loop exhausted...
            Set objArchiveFolder = Nothing       ' ... nothing found
            Set objArchiveMailNotesFolder = Nothing
        End If
    End If
    
    If objArchiveFolder Is Nothing Then
        GoTo ProcReturn
    End If
    
    aFolderName = objArchiveFolder.Name & "\" & NotesName
    aBugTxt = "get note folder " & Quote(aFolderName)
    Call Try
    Set objArchiveNotesFolder = objArchiveFolder.Folders(NotesName)
    If Catch Then
        Beep
        rsp = MsgBox("Ordner " & Quote1(NotesName) & " im Ordner " _
                   & Quote1(objArchiveFolder.Name) & " nicht gefunden...", _
                     vbOKCancel, txtTitel)
        If rsp = vbCancel Then
            GoTo BadExit
        Else
            aBugTxt = "add note folder " & Quote1(aFolderName)
            Call Try
            Set objArchiveNotesFolder = objArchiveFolder.Folders.Add(NotesName)
            If Catch Then
                GoTo BadExit
            End If
        End If
    End If
        
    aFolderName = objArchiveMailNotesFolder.Name & "\" & "eMail-" & NotesName
    aBugTxt = "get note folder " & Quote(aFolderName)
    Call Try
    Set objArchiveMailNotesFolder = _
                                  objArchiveNotesFolder.Folders("eMail-" & NotesName)
    If Catch Then
        Beep
        rsp = MsgBox("Ordner " & Quote1("eMail-" & NotesName) & " im Ordner " _
                   & Quote1(objArchiveNotesFolder.Name) & " nicht gefunden...", _
                     vbOKCancel, txtTitel)
        If rsp = vbCancel Then
            GoTo BadExit
        Else
            aBugTxt = "add note folder " & Quote(aFolderName)
            Call Try
            Set objArchiveMailNotesFolder = objNotesFolder.Folders.Add("eMail-" _
                                                                     & NotesName)
            If Catch Then
                GoTo BadExit
            End If
        End If
    End If
    
BadExit:
    Set objArchiveFolder = Nothing
    Set objArchiveMailNotesFolder = Nothing
    Set objMail = Nothing
    If EPCalled Then
        Call TerminateApp(True)
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.GetMailNotesFolder

'---------------------------------------------------------------------------------------
' Method : Function GetMatchingItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetMatchingItems() As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.GetMatchingItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim itemClass As Long

    Set objMail = Nothing
    Set objNote = Nothing
    strSubject = vbNullString
    
    Set aI = olApp.ActiveInspector               ' auf geöffnete Items (aktiv!) prüfen
    If aI Is Nothing Then                        ' es ist kein item geöffnet
        itemClass = ExplorerSelectedItem         ' prüfe, ob explorer-Markierung vhd.
    Else                                         ' ein Item ist geöffnet
        itemClass = aI.CurrentItem.Class
        If itemClass = olMail Then
            Set objMail = aI.CurrentItem
        ElseIf itemClass = olNote Then
            Set objNote = aI.CurrentItem
        Else                                     ' Rettungsversuch, weil kein geeignetes Objekt gewählt
            itemClass = ExplorerSelectedItem     ' versuchen wir, ob explorer-Markierung vhd.
        End If
    End If
    
    ' Entweder objMail oder objNote (oder keines, wenn falscher Typ) ist <> nothing
    
    GetMatchingItems = True                      ' Annahme wir haben was geeignetes:
    
    If itemClass = olMail Then
        strEntryID = objMail.EntryID
        strSubject = objMail.Subject
        If FindPartnerNote(objMailNotesFolder) Then
            objNote.Display
        Else
            GoTo notfound
        End If
    ElseIf itemClass = olNote Then
        strEntryID = objNote.Categories
        
        aBugTxt = "GetItemFromID " & strEntryID
        Call Try
        Set objMail = aNameSpace.GetItemFromID(strEntryID)
        Catch
        If objMail Is Nothing Then
            strSubject = Trunc(1, objNote.Body, vbCrLf)
            GoTo notfound
        Else
            GetMatchingItems = True
            objMail.Display
        End If
    Else                                         ' dann sind einer oder beide objMail und objNote = nothing
notfound:
        GetMatchingItems = False                 ' weder-noch: dann geht das nicht!
        GoTo ProcReturn
    End If
    
    If LenB(strSubject) = 0 Then                 ' email ohne betreff...
        strSubject = "*** ohne Betreff ***"
    End If

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.GetMatchingItems

'---------------------------------------------------------------------------------------
' Method : Function ListChecker
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ListChecker(PermittedItems As String, testItem As String, Optional SplitString As String) As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.ListChecker"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Static arr As Variant
Static PermString As String
Static cLen(100) As Long
Static NotInStr(100) As Boolean              ' string unerwünscht
Static FullRightPartMatch(100) As Boolean
Static MustBeLongerThan(100) As Long
Dim i&
Dim arri As String
Dim star As Long
Dim lItem&
Dim pItem&
Dim Item As String

    If LenB(SplitString) = 0 Then SplitString = b
    If PermString <> PermittedItems Then
        arr = split(PermittedItems, SplitString)
        PermString = PermittedItems
        For i = 0 To UBound(arr)
            arri = arr(i)
            
            star = InStr(arri, "+")
            If star > 0 And star = Len(arri) Then
                arri = Left(arri, star - 1)
                arr(i) = arri
                MustBeLongerThan(i) = star
                FullRightPartMatch(i) = False
            Else
                MustBeLongerThan(i) = 0
                FullRightPartMatch(i) = True
            End If
            
            star = InStr(arri, "-")
            If star = 1 Then
                arri = Mid(arri, 2)
                arr(i) = arri
                NotInStr(i) = True
                MustBeLongerThan(i) = MustBeLongerThan(i) - 1
            Else
                NotInStr(i) = False
            End If
            
            star = InStr(arri, "*")
            If star = 0 Then
                cLen(i) = Len(testItem)
            Else
                arr(i) = Left(arri, star - 1)
                cLen(i) = star - 1
                MustBeLongerThan(i) = MustBeLongerThan(i) - 1
                FullRightPartMatch(i) = False
            End If
        Next i
    End If
    
    For i = 0 To UBound(arr)
        Item = Left(testItem, cLen(i))
        lItem = Len(testItem)
        pItem = InStr(testItem, arr(i))
        If FullRightPartMatch(i) Then            ' hat keinen * oder +: muss genau passen
            If pItem + Len(arr(i)) <> lItem Then
                Call LogEvent(Item & b & " ist nicht genau = " & b & arr(i), eLnothing)
                pItem = 0
            End If
        End If
        If pItem > 0 Then
            If NotInStr(i) Then
                Call LogEvent(testItem & b & " verworfen wegen -" & b & arr(i), eLnothing)
                ListChecker = False
                GoTo ProcReturn
            Else
                Call LogEvent(testItem & b & " passt zu " & b & arr(i), eLnothing)
                ListChecker = True
            End If
            If MustBeLongerThan(i) > lItem And Len(testItem) < pItem + lItem + MustBeLongerThan(i) Then
                Call LogEvent(testItem & b & " verworfen wegen " & b & arr(i) & "+", eLnothing)
                ListChecker = False
            End If
        End If
    Next i

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.ListChecker

'---------------------------------------------------------------------------------------
' Method : Sub MailNotesFix
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MailNotesFix()
Dim zErr As cErr
Const zKey As String = "UseRules.MailNotesFix"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

 
Dim cnt&, i&
Dim objT As MailItem
  
    GetMailNotesFolder
    Relinked = 0
  
    For Each objNote In objMailNotesFolder.Items
        i = i + 1
        strEntryID = objNote.Categories
        Set objMail = Nothing
        aBugTxt = "find note EntryID=" & strEntryID
        Call Try
        Set objMail = aNameSpace.GetItemFromID(strEntryID)
        If Catch Then
            If FindPartnerMail(aNameSpace.Folders, WantedFolders) Then ' suche...
                cnt = cnt + 1
            End If
        Else
            If objMail.VotingOptions <> objNote.EntryID Then
                objMail.Display
                objT = aNameSpace.GetItemFromID(objNote.EntryID)
                If objT Is Nothing Then
                    objMail.Close (olPromptForSave)
                Else
                    objT.Display
                    objMail.VotingOptions = objNote.EntryID
                    objMail.Close (olSave)
                    Relinked = Relinked + 1
                    objT.Close (olDiscard)
                End If
            End If
        End If
        Catch
    Next objNote
  
    Beep
    MsgBox "Stats des Fix: cnt open notes / i Notes / relinked = " _
         & cnt & "/" & i & "/" & Relinked

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.MailNotesFix

'---------------------------------------------------------------------------------------
' Method : Function notFoundInArchive
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function notFoundInArchive() As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.notFoundInArchive"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If objArchiveFolder Is Nothing Then
        notFoundInArchive = True
    Else
        If FindPartnerNote(objArchiveMailNotesFolder) Then
            notFoundInArchive = False
        Else
            notFoundInArchive = True
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.notFoundInArchive

'---------------------------------------------------------------------------------------
' Method : Function RuleExists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RuleExists(colRules As Outlook.Rules, Name As String, precise As Boolean, ByRef RuleIndex As Long, Optional ByRef foundRule As Rule) As Boolean
Dim zErr As cErr
Const zKey As String = "UseRules.RuleExists"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
    For i = 1 To colRules.Count
        If precise Then
            If colRules.Item(i).Name = Name Then
                RuleExists = True
                foundRule = colRules.Item(i)
                RuleIndex = i
                GoTo ProcReturn
            End If
        Else
            If InStr(1, colRules.Item(i).Name, Name, vbTextCompare) = 1 Then
                RuleExists = True
                foundRule = colRules.Item(i)
                RuleIndex = i
                GoTo ProcReturn
            End If
        End If
    Next i
    RuleIndex = -1

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' UseRules.RuleExists

'---------------------------------------------------------------------------------------
' Method : Sub RunMissedRules
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub RunMissedRules()                             ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "UseRules.RunMissedRules"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="UseRules")

Dim RuleName As String
Dim rRule As String

Dim oStore As Outlook.Store
Dim MainFolder As Outlook.Folder
Dim myfolder As Outlook.Folder
Dim SubFolder As Outlook.Folder

Dim i As Long
Dim j As Long
Dim k As Long
Dim S As Long
Dim colRules As Outlook.Rules
Dim foundRule As Rule
    If LookupFolders Is Nothing Then
        DoVerify False
        IsEntryPoint = True
        EPCalled = True
    End If
    
    If FilterMain Is Nothing Then
        Set FilterMain = New cRuleFilter
        Call FilterMain.setRuleFilter("*My Arch*|*Backup*")
    End If
    If FilterPaths Is Nothing Then
        Set FilterPaths = New cRuleFilter
        Call FilterPaths.setRuleFilter( _
             "*Entwü*|*Drafts*|*Gelös*|*Archi*|*Gesen*|*Junk*|*Papie*|*Posta*|*Sucho*|*Synch*")
    End If
    If FilterRules Is Nothing Then
        Set FilterRules = New cRuleFilter
        Call FilterRules.setRuleFilter(vbNullString)
    End If
    
    For j = 1 To LookupFolders.Count
        Set MainFolder = LookupFolders.Item(j)
        If FilterMain.RuleFilter(MainFolder.Name) Then
            If DebugMode Then
                Call LogEvent("Es werden keine Regeln angewendet im Bereich " _
                            & MainFolder.FolderPath, eLall)
            End If
            GoTo SkipMainFolder
        End If
        If DebugMode Then
            Call LogEvent("Regeln werden angewendet im Bereich " _
                        & MainFolder.FolderPath, eLall)
        End If
        For k = 1 To MainFolder.Folders.Count
            Set myfolder = MainFolder.Folders.Item(k)
            S = -1                               ' my Folder first, this will determine the rules.
            ' Applying rules with subFolders in Execute would not allow name filtering
doSubFolders:
            If S > 0 Then
                S = S + 1
                If myfolder.Folders.Count <= S Then
                    Set SubFolder = myfolder.Folders.Item(S)
                End If
            Else
                Set SubFolder = myfolder
                S = 0
            End If
            
            If FilterPaths.RuleFilter(Left(SubFolder.Name, 5)) Then
                If DebugMode Then
                    Call LogEvent("Es werden keine Regeln angewendet im Ordner " _
                                & SubFolder.FolderPath, eLall)
                End If
                GoTo SkipThisFolder
            End If
            If SubFolder.Items.Count = 0 Then    ' no items to process: skip
                If DebugMode Then
                    Call LogEvent("Die Regel " & RuleName _
                                & " wird nicht ausgeführt weil keine Items in: " _
                                & SubFolder.FolderPath, eLall)
                End If
            Else
                If S = 0 Then                    ' Rules are defined for MyFolder, not subFolder (which inherit those)
                    Set oStore = myfolder.Store
                    Set colRules = oStore.GetRules
                    If DebugMode And colRules.Count = 0 Then
                        Call LogEvent("Zum Ordner " & myfolder.FullFolderPath _
                                    & " gibt es keine Regeln", eLall)
                    End If
                End If
                
                For i = 1 To colRules.Count
                    Set foundRule = colRules.Item(i)
                    RuleName = foundRule.Name
                    ' filter rules RuleName
                    rRule = Trim(Trunc(1, RuleName, "("))
                    rRule = RTail(rRule, b)
                    If FilterRules.RuleFilter(rRule) Then
                        If DebugMode Then
                            Call LogEvent("Die Regel " & Quote(RuleName) _
                                        & " wird nicht berücksichtigt für: " _
                                        & myfolder.FullFolderPath, eLall)
                        End If
                    Else
                        'On Error GoTo 0
                        If DebugMode Then
                            Call LogEvent("Die Regel " & Quote(RuleName) _
                                        & " wird ausgeführt in: " _
                                        & SubFolder.FullFolderPath, eLall)
                        End If
                        aBugVer = "Execute Rule " & Quote(RuleName) _
      & " in Folder: " _
      & Quote(SubFolder.FullFolderPath)
                        Call Try                 ' Try anything, autocatch
                        foundRule.Execute True, SubFolder
                        Catch
                    End If
                Next i
            End If
            If myfolder.Folders.Count > S Then
                GoTo doSubFolders
            End If
SkipThisFolder:
        Next k
SkipMainFolder:
    Next j
    If EPCalled Then
        Call TerminateApp(True)                  ' ! ! ! Entry Point must call TerminateApp ! ! !
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' UseRules.RunMissedRules


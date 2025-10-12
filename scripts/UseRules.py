# Converted from UseRules.py

# Attribute VB_Name = "UseRules"
# Option Explicit

# Const txtTitel As String = "EMail-Notiz:"
# Const WantedFolders As String = _
# "\Gese+ \Post* -\Gel+ -\Junk* -\Not+ -Auf+ -\Un+ -onta* -alend* \\+"

# Public LogFolder As Folder
# ' Public SearchFoldersFolder As Folders
# Public SearchObject As Search

# Public FilterMain As cRuleFilter
# Public FilterPaths As cRuleFilter
# Public FilterRules As cRuleFilter

# Dim objMail As MailItem
# Dim objNote As NoteItem
# Dim aI As Inspector

# Dim objNotesFolder As Folder
# Dim objMailNotesFolder As Folder
# Dim objArchiveFolder As Folder
# Dim objArchiveNotesFolder As Folder
# Dim objArchiveMailNotesFolder As Folder

# Dim strEntryID As String
# Dim strSubject As String
# Dim itemClass As Long
# Dim NotesName As String
# Dim Relinked As Long

# '---------------------------------------------------------------------------------------
# ' Method : Sub AddNewNote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def addnewnote():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.AddNewNote"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call GetMailNotesFolder
    if GetMatchingItems() Then:
    if Not objNote Is Nothing Then:
    # GoTo ProcReturn
    if LenB(objMail.VotingOptions) = 0 Then:
    # Call AddNote
    else:
    # Call eMailNoteOps
    else:
    if objMail Is Nothing Then:
    else:
    # Call AddNote

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AddNote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def addnote():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.AddNote"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # objMail.FlagIcon = olBlueFlagIcon
    # objMail.FlagRequest = "Zur Nachverfolgung"
    # Set objNote = objMailNotesFolder.Items.Add
    # With objNote
    # .Categories = strEntryID
    # .Body = "Notiz zu: " & objMail.Subject & vbCrLf
    # .width = 500
    # .height = 250
    # .Color = olBlue
    # .Save
    # .Display
    # End With                                     ' objNote
    # objMail.VotingOptions = objNote.EntryID
    # Call setItmCats(objMail, "anotiert", LOGGED)
    # objMail.Save

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ClearEmailNote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clearemailnote():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.ClearEmailNote"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call GetMailNotesFolder
    # GetMatchingItems

    if objNote Is Nothing Then:
    print('Keine zu lschenden eMail-Notizen zu ')
    # & Quote(strSubject) & vbCrLf & "gefunden ...", _
    # vbOKOnly + vbInformation, txtTitel
    else:
    # objMail.FlagIcon = olNone
    # objMail.FlagRequest = vbNullString
    # objMail.VotingOptions = vbNullString
    # Call setItmCats(objMail, vbNullString, LOGGED & "; anotiert")
    # objMail.Save
    # objNote.Delete
    # Set aItmDsc.idObjItem = Nothing
    print('eMail-Notiz gelscht zu')
    # vbOKOnly + vbInformation, txtTitel

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function CountItemsIn
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def countitemsin():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.CountItemsIn"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if curObj Is Nothing Then:
    # GoTo ProcReturn
    if TypeName(curObj) = "Selection" _:
    # Or TypeName(curObj) = "Collection" Then
    # CountItemsIn = curObj.Count
    # LoopProgress = 0
    # eOnlySelectedItems = True
    else:
    # CountItemsIn = curObj.Items.Count
    # eOnlySelectedItems = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub CreateRules
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createrules():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "UseRules.CreateRules"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="UseRules")

    # Dim oStore As Outlook.Store
    # Dim oRule As Outlook.Rule
    # Dim myfolder As Outlook.Folder
    # Dim colRules As Outlook.Rules
    # Dim oCondition As Outlook.RuleCondition
    # Dim oRuleAction As Outlook.RuleAction
    # Dim i As Long
    # Dim Name As String
    # Dim pName As String
    if LookupFolders Is Nothing Then:
    # DoVerify False
    # IsEntryPoint = True
    # EPCalled = True

    # Set myfolder = ActiveInspector.Parent
    # Set oStore = myfolder.Store
    # Set colRules = oStore.GetRules
    # Name = "unerwnscht (nur per Makro)"
    if RuleExists(colRules, Name, precise:=False, RuleIndex:=i) Then:
    # DoVerify False                           ' can, "W.xlTSheet create existing rule"
    # colRules.Remove i
    else:
    # Set oRule = colRules.Create(Name, olRuleReceive)
    match Name:
        case "unerwnscht (nur per Makro)":
    # pName = Trunc((Name), 1, b)
    # 'Specify the condition in a xxRuleCondition object
    # Set oCondition = oRule.Conditions.MessageHeader
    # oCondition.Enabled = True
    # 'Specify the action in a MoveOrCopyRuleAction object
    # Set oRuleAction = oRule.Actions.MoveToFolder
    # ' Action is to move the message to the target Folder
    # oRuleAction.Enabled = True
    # oRuleAction.Folder = GetFolderByName(pName)
        case _:
    # DoVerify False                       ' don, "W.xlTSheet know rule yet"
    if EPCalled Then:
    # Call TerminateApp

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub eMailNoteOps
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def emailnoteops():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.eMailNoteOps"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if GetMatchingItems() Then                   ' beschafft die strEntryId der geffneten eMail:

    if objNote Is Nothing Then               ' neue Notiz erstellen zur selektierten email:
    # Call AddNewNote
    elif objMail Is Nothing Then:
    # & vbCrLf & B2 & Quote(strSubject) & vbCrLf _
    # & "gefunden. Neu anlegen=Ja, in allen eMail-Ordnern suchen (dauert lnger)=Nein?", _
    # vbYesNoCancel, txtTitel)
    if rsp = vbNo Then:
    if Not FindPartnerMail(aNameSpace.Folders, WantedFolders) Then:
    # & Quote(strSubject) & vbCrLf & _
    # "gefunden (in keinem der eMail-Ordner) ...", _
    # vbOKOnly + vbInformation, txtTitel)
    if rsp = vbYes Then:
    # Call AddNewNote
    elif objNote Is Nothing Then:
    if LenB(objMail.VotingOptions) = 0 Then:
    # Call AddNewNote
    else:
    if FindPartnerNote(objMailNotesFolder) Then:
    # DoVerify False
    else:
    # & "nicht am erwarteten Ort. Neu anlegen=Ja" & vbCrLf _
    # & "in allen Notizen-Ordnern suchen (dauert lnger)=Nein?", _
    # vbYesNoCancel, txtTitel)
    if rsp = vbNo Then:
    # DoVerify False
    if rsp = vbYes Then:
    # Call AddNewNote

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EMailNotes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def emailnotes():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.EMailNotes"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Call GetMailNotesFolder
    # Call eMailNoteOps

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EMailNotesCleanup
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def emailnotescleanup():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.EMailNotesCleanup"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim cnt&

    # GetMailNotesFolder


    for objnote in objmailnotesfolder:
    # strEntryID = objNote.Categories
    # aBugTxt = "find note for EntryID=" & strEntryID
    # Call Try
    # Set objMail = aNameSpace.GetItemFromID(strEntryID)
    if Catch Then:
    if notFoundInArchive() Then:
    # objNote.Delete
    if Not Catch Then:
    # cnt = cnt + 1
    else:
    # DoVerify False, "*** Note Item is not in archive"

    # Beep
    if cnt > 0 Then:
    print('Es wurde(n) ')
    # " nicht mehr bentigte eMail-Notize(n) " & _
    # "gelscht...", vbOKOnly + vbInformation, txtTitel
    else:
    print('Keine zu lschenden eMail-Notizen ')
    # "gefunden...", vbOKOnly + vbInformation, txtTitel

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ExecuteDefinedRule
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def executedefinedrule():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.ExecuteDefinedRule"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="rule=" & RuleName)

    # Static oStore As Outlook.Store
    # Static oRule As Outlook.Rule
    # Dim myfolder As Outlook.Folder

    # Dim i As Long
    # Dim j As Long
    # Dim colRules As Outlook.Rules
    # Dim FullFolderPath As String
    # Static fin As cKnownFolders
    # Dim reUse As Boolean
    # Dim translatedFolderName As Variant
    # Dim msg As String

    if LookupFolders Is Nothing Then:
    # DoVerify False
    # IsEntryPoint = True

    if inFolder Is Nothing Then:
    if fin Is Nothing Then:
    # Set fin = New cKnownFolders
    else:
    # Set myfolder = inFolder
    # Set oStore = inFolder.Store
    # GoTo folderPassed

    # reUse = fin.InRuleFolderColl(inFolderName)
    # With fin.KnownFolderItem
    if reUse Then                            ' Folder to apply rules is reusable:
    if .FolderOK < 0 Then:
    if .FInFolderWarning = -1 Then:
    # .FInFolderWarning = -2       ' warn only once
    # Call LogEvent("Folder " & .FInFolderstring _
    # & " previously marked as bad, kann Regel " _
    # & RuleName & " nicht ausfhren")
    # GoTo FuncExit
    elif .FolderOK = 1 Then:
    if .FInFolderWarning = -1 Then:
    # .FInFolderWarning = -2       ' warn only once
    # Call LogEvent("Skipping rule for bad Folder: " & .FInFolderstring, eLmin)
    else:
    # Call LogEvent("Using previously found Folder " _
    # & .FInFolderstring & " again", eLmin)
    else:
    # .FInFolderstring = inFolderName
    # .FolderIsNormal = 0
    if LenB(inFolderName) = 0 Then       ' try to use the current Folder:
    # Set myfolder = ActiveInspector.Parent
    # FullFolderPath = myfolder.FolderPath
    else:
    if inFolderName = "unerwnscht" Then:
    # In Array("unerwnscht", "Junk-E-Mail", "Spam", "Junk")
    # Set myfolder = GetFolderByName( _
    # CStr(translatedFolderName), noSearchFolders:=True)
    if Not myfolder Is Nothing Then:
    # inFolderName = translatedFolderName
    # Exit For
    else:
    # Set myfolder = GetFolderByName(inFolderName, noSearchFolders:=True)
    if myfolder Is Nothing Then:
    # Call LogEvent("Keinen normalen Ordner namens " & Quote(inFolderName) _
    # & " gefunden, kann Regel " _
    # & RuleName & " nicht ausfhren")
    # inFolderName = .FInFolderstring
    # .FolderIsNormal = -1             ' not a normal Folder
    # GoTo FuncExit
    # Set oStore = myfolder.Store
    # folderPassed:
    if myfolder Is Nothing Then:
    # Call LogEvent("Ordner " & Quote(inFolderName) _
    # & " nicht vorhanden, kann Regel " _
    # & RuleName & " nicht ausfhren")
    # .FolderIsNormal = -1                 ' not a normal Folder
    # GoTo FuncExit

    # aBugTxt = "get item count in folder " & Quote(myfolder.FolderPath)
    # Call Try
    if myfolder.Items.Count = 0 Then         ' no items to process: skip:
    if Catch(DoMessage:=False) Then:
    # Call LogEvent("Fehler bei Ordner " & Quote(myfolder.FolderPath) _
    # & vbCrLf & "Regel " & RuleName _
    # & " kann nicht ausgefhrt werden ")
    # Call ErrReset(4)
    # .FolderIsNormal = -1             ' not a normal Folder
    # GoTo FuncExit
    # Call LogEvent("Die Regel " & RuleName _
    # & " wird nicht ausgefhrt weil keine Items in: " _
    # & Quote(myfolder.FolderPath), eLmin)
    # GoTo FuncExit

    # Set colRules = oStore.GetRules
    if colRules.Count = 0 Then:
    # Call LogEvent("Keine Regeln definiert, daher Regel " _
    # & Quote(RuleName) & " nicht gefunden", eLall)
    else:
    # aBugTxt = "Find Rule " & RuleName
    # Call Try
    # Set oRule = colRules.Item(RuleName)  ' try find exact name
    if Not Catch Then:
    # GoTo ProcReturn
    # ' search by similarity
    # Set oRule = colRules.Item(i)
    if InStr(1, oRule.Name, RuleName, vbTextCompare) > 0 Then:
    # GoTo DoExec

    # Call LogEvent("Keine Regel mit passendem Namen zu " _
    # & RuleName & " gefunden." _
    # & vbCrLf & "Es sind insgesamt " & colRules.Count _
    # & " Regeln vorhanden.", eLall)
    # GoTo FuncExit
    # DoExec:
    if oRule.Enabled Then:
    if DebugLogging Then:
    print(Debug.Print i, oRule.Name, oRule.RuleType)
    if oRule.Actions.Item(j).Enabled Then:
    print(Debug.Print "Action", j, oRule.Actions.Item(j).ActionType)
    else:
    if DebugLogging Then:
    print(Debug.Print oRule.Name & " is not enabled")
    # GoTo FuncExit

    # msg = Quote(oRule.Name) & " wird ausgefhrt in: " _
    # & myfolder.FullFolderPath
    # Call LogEvent("Execute Rule " & msg, eLmin)

    # aBugTxt = "Execute Rule"
    # Call Try
    # oRule.Execute DebugMode, myfolder
    # Catch

    if Not fin Is Nothing Then:
    # .FolderIsNormal = 0                  ' good normal Folder
    # .FolderIsSearch = -1
    # End With                                     ' FIn.KnownFolderItem

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ExplorerSelectedItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def explorerselecteditem():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.ExplorerSelectedItem"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim exp As Explorer
    # Set exp = olApp.ActiveExplorer
    if exp.Selection.Count = 1 Then:
    # ExplorerSelectedItem = exp.Selection(1).Class
    if ExplorerSelectedItem = olMail Then:
    # Set objMail = exp.Selection(1)
    elif exp.Selection(1).Class = olNote Then:
    # Set objNote = exp.Selection(1)
    else:
    # ExplorerSelectedItem = 0
    else:
    # ExplorerSelectedItem = 0

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindPartnerMail
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findpartnermail():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.FindPartnerMail"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim objFolder As Folder
    # 'On Error GoTo 0
    for objfolder in topfolder:
    # ' Folderpath starts with \\
    if ListChecker(Filterlist, objFolder.FolderPath) Then:
    if FindPartnerMailInFolder(objFolder, Filterlist) Then:
    # FindPartnerMail = True
    # GoTo ProcReturn

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindPartnerMailInFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findpartnermailinfolder():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.FindPartnerMailInFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim objMapi As Object
    # Dim objFolder As Folder
    # Dim FpM As Boolean

    for objmapi in objanymailfolder:
    # Call ErrReset(4)
    if DebugMode Then:
    # Call LogEvent(objMapi.Subject & b & objMapi.Class, eLall)
    if objMapi.Class = olMail Then:
    # Set objMail = Nothing
    # Set objMail = objMapi
    if objMail.VotingOptions = objNote.EntryID Then:
    # FindPartnerMailInFolder = True
    # GoTo ProcReturn
    if Len(objMail.Subject) > 4 _:
    # And InStr(objNote.Body, objMail.Subject) > 0 Then
    # objMail.Display
    # objMail.Close (olPromptForSave)
    if rsp = vbYes Then:
    # FindPartnerMailInFolder = True
    # GoTo ProcReturn

    else:
    # GoTo noMailFolder

    # noMailFolder:
    for objmapi in objanymailfolder:
    # Call LogEvent(CStr(objMapi) & b & Filterlist & b _
    # & objMapi.FolderPath, eLnothing + DebugMode)
    if ListChecker(Filterlist, objMapi.FolderPath) Then:
    # Call ErrReset(4)
    # Set objFolder = objMapi
    # FpM = FindPartnerMailInFolder(objFolder) Or FpM
    # FindPartnerMailInFolder = FpM

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindPartnerNote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findpartnernote():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.FindPartnerNote"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim objT As MailItem
    if objAnyNotesFolder Is Nothing Then:
    # GoTo ProcReturn                          ' returning false
    for objnote in objanynotesfolder:
    if objNote.Categories = strEntryID Then:
    # FindPartnerNote = True
    # ' check reverse mapping; only OK if unique Note !
    if objMail.VotingOptions <> objNote.EntryID Then:
    # objMail.Display
    try:
        # Set objT = aNameSpace.GetItemFromID(objNote.Categories)
        # objT.Display
        # Set objMail = objT
        # relink:
        # objMail.VotingOptions = objNote.EntryID
        # objMail.Save
        # Relinked = Relinked + 1
        # Exit For
        # linkerr:
        # & "verknpft sich nicht zur eMail. Reparieren=Ja" & vbCrLf _
        # & "Aufhren=Nein, Alternative suchen=Cancel", _
        # vbYesNoCancel, txtTitel)
        if rsp = vbNo Then:
        # DoVerify False
        elif rsp = vbYes Then:
        # GoTo relink
        else:
        # FindPartnerNote = False
        # Exit For
        # FindPartnerNote = False
        # GoTo ProcReturn

        # FuncExit:

        # ProcReturn:
        # Call ProcExit(zErr)

        # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub getAccountDscriptors
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getaccountdscriptors():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.getAccountDscriptors"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim lAccountDsc As cAccount
    # Dim lAccount As Account

    # Set D_AccountDscs = New Dictionary
    # Set oStores = olApp.Session.Stores
    # Set SessionAccounts = olApp.Session.Accounts

    # With SessionAccounts
    # Set lAccount = .Item(i)
    # Set lAccountDsc = New cAccount
    # With lAccountDsc
    # Set .aAcStore = lAccount.DeliveryStore
    # E_Active.Permit = 0
    # .Key = .aAcStore.DisplayName
    # Call ErrReset(0)
    # .aAcType = lAccount.AccountType
    # D_AccountDscs.Add .Key, lAccountDsc
    # End With                             ' laccountdsc
    # End With                                     ' ...Accounts
    # aAccountNumber = 0                           ' no current one

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function getAccountType
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getaccounttype():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.getAccountType"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim AccountName As String
    # Dim aAccount As Account
    if LenB(FolderPath) = 0 Then:
    # GoTo stmailAccount
    if Left(FolderPath, 2) = "\\" Then:
    # i = 3                                    ' if fully qualified, cut \\
    else:
    # i = 1
    # AccountName = Trunc(i, FolderPath, "\")
    # With SessionAccounts
    # ' do a fast check without loop if it matches
    if aAccountNumber > 0 And aAccountNumber <= .Count Then:
    # Set aAccount = .Item(aAccountNumber)
    # aBugTxt = "find delivery store " & AccountName
    # Call Try(allowAll)
    # Catch False                          ' could not find delivery store if not registered as email?
    if aAccount.DeliveryStore.DisplayName = AccountName Then:
    # getAccountType = aAccount.AccountType
    if Not IsMissing(withType) Then:
    # withType = getAccountTypeName(aAccountNumber)
    # GoTo FuncExit                ' fast check successful
    # Call ErrReset(0)

    # ' loop through accounts because it did not match
    # Set aAccount = .Item(aAccountNumber)
    # aBugTxt = "find delivery store " & aAccount.DeliveryStore.DisplayName
    # Call Try(allowAll)
    # Catch
    if aAccount.DeliveryStore.DisplayName = AccountName Then:
    # getAccountType = aAccount.AccountType
    if Not IsMissing(withType) Then:
    # withType = getAccountTypeName(aAccountNumber)
    # GoTo FuncExit
    # End With                             ' ...Accounts
    # Call ErrReset(0)

    # stmailAccount:                                   ' Debug.Assert False    ' Folder without account
    if Not IsMissing(withType) Then:
    # withType = "stmailAccount"
    # aAccountNumber = 0

    # FuncExit:
    # Call N_ErrClear
    # aAccountType = getAccountType

    # ProcReturn:
    # Call ProcExit(zErr)


# ' DISCARD, NOT USED
def getaccounttypename():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.getAccountTypeName"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim acc1 As OlAccountType

    # ' convert acc to a numeric
    match VarType(aCC):
        case vbString:
    # acc1 = getAccountType(aCC) + 1
        case vbInteger, vbLong::
    # acc1 = aCC + 1
        case _:
    # DoVerify False, " not impl"
    # getAccountTypeName = AccountTypeNames(acc1)

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetDItem_P
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getditem_p():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.GetDItem_P"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # ' note that this requires the IsExplorer and IsInspector functions found elsewhere
    # Dim obj As Object

    match True:
        case (TypeName(olApp.ActiveWindow) = "Explorer"):
    # Set obj = ActiveExplorer.Selection.Item(1)
        case (TypeName(olApp.ActiveWindow) = "Inspector"):
    # Set obj = ActiveInspector.CurrentItem
    if LenB(OnlyType) > 0 Then:
    if InStr(OnlyType & b, TypeName(obj) & b) > 0 Then:
    # Set GetDItem_P = obj
    else:
    # Set GetDItem_P = obj                     ' any type possible

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub GetMailNotesFolder
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getmailnotesfolder():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.GetMailNotesFolder"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim aFolderName As String

    # Set objNotesFolder = aNameSpace.GetDefaultFolder(olFolderNotes)
    # NotesName = objNotesFolder.FullFolderPath
    # CurIterationSwitches.ResetCategories = False
    if objArchiveFolder Is Nothing Then:
    # Set objArchiveFolder = aNameSpace.Folders(i)
    if Left(objArchiveFolder.Name, 6) = "Archiv" Then:
    # Exit For
    if i > aNameSpace.Folders.Count Then     ' loop exhausted...:
    # Set objArchiveFolder = Nothing       ' ... nothing found
    # Set objArchiveMailNotesFolder = Nothing

    if objArchiveFolder Is Nothing Then:
    # GoTo ProcReturn

    # aFolderName = objArchiveFolder.Name & "\" & NotesName
    # aBugTxt = "get note folder " & Quote(aFolderName)
    # Call Try
    # Set objArchiveNotesFolder = objArchiveFolder.Folders(NotesName)
    if Catch Then:
    # Beep
    # & Quote1(objArchiveFolder.Name) & " nicht gefunden...", _
    # vbOKCancel, txtTitel)
    if rsp = vbCancel Then:
    # GoTo BadExit
    else:
    # aBugTxt = "add note folder " & Quote1(aFolderName)
    # Call Try
    # Set objArchiveNotesFolder = objArchiveFolder.Folders.Add(NotesName)
    if Catch Then:
    # GoTo BadExit

    # aFolderName = objArchiveMailNotesFolder.Name & "\" & "eMail-" & NotesName
    # aBugTxt = "get note folder " & Quote(aFolderName)
    # Call Try
    # Set objArchiveMailNotesFolder = _
    # objArchiveNotesFolder.Folders("eMail-" & NotesName)
    if Catch Then:
    # Beep
    # & Quote1(objArchiveNotesFolder.Name) & " nicht gefunden...", _
    # vbOKCancel, txtTitel)
    if rsp = vbCancel Then:
    # GoTo BadExit
    else:
    # aBugTxt = "add note folder " & Quote(aFolderName)
    # Call Try
    # Set objArchiveMailNotesFolder = objNotesFolder.Folders.Add("eMail-" _
    # & NotesName)
    if Catch Then:
    # GoTo BadExit

    # BadExit:
    # Set objArchiveFolder = Nothing
    # Set objArchiveMailNotesFolder = Nothing
    # Set objMail = Nothing
    if EPCalled Then:
    # Call TerminateApp(True)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetMatchingItems
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getmatchingitems():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.GetMatchingItems"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim itemClass As Long

    # Set objMail = Nothing
    # Set objNote = Nothing
    # strSubject = vbNullString

    # Set aI = olApp.ActiveInspector               ' auf geffnete Items (aktiv!) prfen
    if aI Is Nothing Then                        ' es ist kein item geffnet:
    # itemClass = ExplorerSelectedItem         ' prfe, ob explorer-Markierung vhd.
    else:
    # itemClass = aI.CurrentItem.Class
    if itemClass = olMail Then:
    # Set objMail = aI.CurrentItem
    elif itemClass = olNote Then:
    # Set objNote = aI.CurrentItem
    else:
    # itemClass = ExplorerSelectedItem     ' versuchen wir, ob explorer-Markierung vhd.

    # ' Entweder objMail oder objNote (oder keines, wenn falscher Typ) ist <> nothing

    # GetMatchingItems = True                      ' Annahme wir haben was geeignetes:

    if itemClass = olMail Then:
    # strEntryID = objMail.EntryID
    # strSubject = objMail.Subject
    if FindPartnerNote(objMailNotesFolder) Then:
    # objNote.Display
    else:
    # GoTo notfound
    elif itemClass = olNote Then:
    # strEntryID = objNote.Categories

    # aBugTxt = "GetItemFromID " & strEntryID
    # Call Try
    # Set objMail = aNameSpace.GetItemFromID(strEntryID)
    # Catch
    if objMail Is Nothing Then:
    # strSubject = Trunc(1, objNote.Body, vbCrLf)
    # GoTo notfound
    else:
    # GetMatchingItems = True
    # objMail.Display
    else:
    # notfound:
    # GetMatchingItems = False                 ' weder-noch: dann geht das nicht!
    # GoTo ProcReturn

    if LenB(strSubject) = 0 Then                 ' email ohne betreff...:
    # strSubject = "*** ohne Betreff ***"

    # FuncExit:
    # Call ErrReset(0)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function ListChecker
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def listchecker():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.ListChecker"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Static arr As Variant
    # Static PermString As String
    # Static cLen(100) As Long
    # Static NotInStr(100) As Boolean              ' string unerwnscht
    # Static FullRightPartMatch(100) As Boolean
    # Static MustBeLongerThan(100) As Long
    # Dim i&
    # Dim arri As String
    # Dim star As Long
    # Dim lItem&
    # Dim pItem&
    # Dim Item As String

    if LenB(SplitString) = 0 Then SplitString = b:
    if PermString <> PermittedItems Then:
    # arr = split(PermittedItems, SplitString)
    # PermString = PermittedItems
    # arri = arr(i)

    # star = InStr(arri, "+")
    if star > 0 And star = Len(arri) Then:
    # arri = Left(arri, star - 1)
    # arr(i) = arri
    # MustBeLongerThan(i) = star
    # FullRightPartMatch(i) = False
    else:
    # MustBeLongerThan(i) = 0
    # FullRightPartMatch(i) = True

    # star = InStr(arri, "-")
    if star = 1 Then:
    # arri = Mid(arri, 2)
    # arr(i) = arri
    # NotInStr(i) = True
    # MustBeLongerThan(i) = MustBeLongerThan(i) - 1
    else:
    # NotInStr(i) = False

    # star = InStr(arri, "*")
    if star = 0 Then:
    # cLen(i) = Len(testItem)
    else:
    # arr(i) = Left(arri, star - 1)
    # cLen(i) = star - 1
    # MustBeLongerThan(i) = MustBeLongerThan(i) - 1
    # FullRightPartMatch(i) = False

    # Item = Left(testItem, cLen(i))
    # lItem = Len(testItem)
    # pItem = InStr(testItem, arr(i))
    if FullRightPartMatch(i) Then            ' hat keinen * oder +: muss genau passen:
    if pItem + Len(arr(i)) <> lItem Then:
    # Call LogEvent(Item & b & " ist nicht genau = " & b & arr(i), eLnothing)
    # pItem = 0
    if pItem > 0 Then:
    if NotInStr(i) Then:
    # Call LogEvent(testItem & b & " verworfen wegen -" & b & arr(i), eLnothing)
    # ListChecker = False
    # GoTo ProcReturn
    else:
    # Call LogEvent(testItem & b & " passt zu " & b & arr(i), eLnothing)
    # ListChecker = True
    if MustBeLongerThan(i) > lItem And Len(testItem) < pItem + lItem + MustBeLongerThan(i) Then:
    # Call LogEvent(testItem & b & " verworfen wegen " & b & arr(i) & "+", eLnothing)
    # ListChecker = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MailNotesFix
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mailnotesfix():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.MailNotesFix"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")


    # Dim cnt&, i&
    # Dim objT As MailItem

    # GetMailNotesFolder
    # Relinked = 0

    for objnote in objmailnotesfolder:
    # i = i + 1
    # strEntryID = objNote.Categories
    # Set objMail = Nothing
    # aBugTxt = "find note EntryID=" & strEntryID
    # Call Try
    # Set objMail = aNameSpace.GetItemFromID(strEntryID)
    if Catch Then:
    if FindPartnerMail(aNameSpace.Folders, WantedFolders) Then ' suche...:
    # cnt = cnt + 1
    else:
    if objMail.VotingOptions <> objNote.EntryID Then:
    # objMail.Display
    # objT = aNameSpace.GetItemFromID(objNote.EntryID)
    if objT Is Nothing Then:
    # objMail.Close (olPromptForSave)
    else:
    # objT.Display
    # objMail.VotingOptions = objNote.EntryID
    # objMail.Close (olSave)
    # Relinked = Relinked + 1
    # objT.Close (olDiscard)
    # Catch

    # Beep
    print('Stats des Fix: cnt open notes / i Notes / relinked = ')
    # & cnt & "/" & i & "/" & Relinked

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function notFoundInArchive
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def notfoundinarchive():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.notFoundInArchive"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    if objArchiveFolder Is Nothing Then:
    # notFoundInArchive = True
    else:
    if FindPartnerNote(objArchiveMailNotesFolder) Then:
    # notFoundInArchive = False
    else:
    # notFoundInArchive = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function RuleExists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def ruleexists():
    # Dim zErr As cErr
    # Const zKey As String = "UseRules.RuleExists"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    if precise Then:
    if colRules.Item(i).Name = Name Then:
    # RuleExists = True
    # foundRule = colRules.Item(i)
    # RuleIndex = i
    # GoTo ProcReturn
    else:
    if InStr(1, colRules.Item(i).Name, Name, vbTextCompare) = 1 Then:
    # RuleExists = True
    # foundRule = colRules.Item(i)
    # RuleIndex = i
    # GoTo ProcReturn
    # RuleIndex = -1

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub RunMissedRules
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def runmissedrules():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "UseRules.RunMissedRules"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="UseRules")

    # Dim RuleName As String
    # Dim rRule As String

    # Dim oStore As Outlook.Store
    # Dim MainFolder As Outlook.Folder
    # Dim myfolder As Outlook.Folder
    # Dim SubFolder As Outlook.Folder

    # Dim i As Long
    # Dim j As Long
    # Dim k As Long
    # Dim S As Long
    # Dim colRules As Outlook.Rules
    # Dim foundRule As Rule
    if LookupFolders Is Nothing Then:
    # DoVerify False
    # IsEntryPoint = True
    # EPCalled = True

    if FilterMain Is Nothing Then:
    # Set FilterMain = New cRuleFilter
    # Call FilterMain.setRuleFilter("*My Arch*|*Backup*")
    if FilterPaths Is Nothing Then:
    # Set FilterPaths = New cRuleFilter
    # Call FilterPaths.setRuleFilter( _
    # "*Entw*|*Drafts*|*Gels*|*Archi*|*Gesen*|*Junk*|*Papie*|*Posta*|*Sucho*|*Synch*")
    if FilterRules Is Nothing Then:
    # Set FilterRules = New cRuleFilter
    # Call FilterRules.setRuleFilter(vbNullString)

    # Set MainFolder = LookupFolders.Item(j)
    if FilterMain.RuleFilter(MainFolder.Name) Then:
    if DebugMode Then:
    # Call LogEvent("Es werden keine Regeln angewendet im Bereich " _
    # & MainFolder.FolderPath, eLall)
    # GoTo SkipMainFolder
    if DebugMode Then:
    # Call LogEvent("Regeln werden angewendet im Bereich " _
    # & MainFolder.FolderPath, eLall)
    # Set myfolder = MainFolder.Folders.Item(k)
    # S = -1                               ' my Folder first, this will determine the rules.
    # ' Applying rules with subFolders in Execute would not allow name filtering
    # doSubFolders:
    if S > 0 Then:
    # S = S + 1
    if myfolder.Folders.Count <= S Then:
    # Set SubFolder = myfolder.Folders.Item(S)
    else:
    # Set SubFolder = myfolder
    # S = 0

    if FilterPaths.RuleFilter(Left(SubFolder.Name, 5)) Then:
    if DebugMode Then:
    # Call LogEvent("Es werden keine Regeln angewendet im Ordner " _
    # & SubFolder.FolderPath, eLall)
    # GoTo SkipThisFolder
    if SubFolder.Items.Count = 0 Then    ' no items to process: skip:
    if DebugMode Then:
    # Call LogEvent("Die Regel " & RuleName _
    # & " wird nicht ausgefhrt weil keine Items in: " _
    # & SubFolder.FolderPath, eLall)
    else:
    if S = 0 Then                    ' Rules are defined for MyFolder, not subFolder (which inherit those):
    # Set oStore = myfolder.Store
    # Set colRules = oStore.GetRules
    if DebugMode And colRules.Count = 0 Then:
    # Call LogEvent("Zum Ordner " & myfolder.FullFolderPath _
    # & " gibt es keine Regeln", eLall)

    # Set foundRule = colRules.Item(i)
    # RuleName = foundRule.Name
    # ' filter rules RuleName
    # rRule = Trim(Trunc(1, RuleName, "("))
    # rRule = RTail(rRule, b)
    if FilterRules.RuleFilter(rRule) Then:
    if DebugMode Then:
    # Call LogEvent("Die Regel " & Quote(RuleName) _
    # & " wird nicht bercksichtigt fr: " _
    # & myfolder.FullFolderPath, eLall)
    else:
    # 'On Error GoTo 0
    if DebugMode Then:
    # Call LogEvent("Die Regel " & Quote(RuleName) _
    # & " wird ausgefhrt in: " _
    # & SubFolder.FullFolderPath, eLall)
    # aBugVer = "Execute Rule " & Quote(RuleName) _
    # & " in Folder: " _
    # & Quote(SubFolder.FullFolderPath)
    # Call Try                 ' Try anything, autocatch
    # foundRule.Execute True, SubFolder
    # Catch
    if myfolder.Folders.Count > S Then:
    # GoTo doSubFolders
    # SkipThisFolder:
    # SkipMainFolder:
    if EPCalled Then:
    # Call TerminateApp(True)                  ' ! ! ! Entry Point must call TerminateApp ! ! !

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

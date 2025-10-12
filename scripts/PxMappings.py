# Converted from PxMappings.py

# Attribute VB_Name = "PxMappings"
# Option Explicit
# ' ALL Columns from MPE, in that ORDER
# '    1        2          3        4        5       6             7     8
# ' Name    Nachname    Vorname   Mobil   Zuhause Geschftlich    Fax Andere
# '    9       10         11       12       13      14            15
# ' e-mail  2. e-mail   3. e-mail Web Adresse Adresse (geschftlich)  Firma
# '  16          17       18
# ' Info    Geburtstag  Konto

# Public MPEColumnNames
# Public MPEPropertyNames
# Public MPEItems As Collection
# Public AddressSubfields
# Public MPEchanged As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Sub carefullAssign
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def carefullassign():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "PxMappings.carefullAssign"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)
    # ' do not forget to Call N_ReturnProc(zErr)

    # Dim oldValue As Variant
    # oldValue = ciString
    if MPEstring = "###" Then:
    # MPEstring = vbNullString
    # GoTo setit
    if (ciString = vbNullString Or InStr(ciString, "###") = 1) _:
    # And LenB(MPEstring) > 0 _
    # And ciString <> MPEstring Then
    # setit:
    # ciString = MPEstring
    # MPEItemDiffs = MPEItemDiffs & vbCrLf _
    # & "changed " & ColName & "=" _
    # & Quote(oldValue) & "  to " _
    # & Quote(MPEstring) & vbCrLf
    # MPEchanged = True
    # WorkItemMod(2) = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ColumnStructure
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def columnstructure():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.ColumnStructure"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long, P As Long, j As Long
    # Dim oneColumnElement As cNumbItem
    # Dim onePropertyElement As cNumbItem

    # Set MPEItems = New Collection
    # Set oneColumnElement = New cNumbItem
    # Set onePropertyElement = New cNumbItem
    # With oneColumnElement
    # .NuIndex = i
    # .Key = MPEColumnNames(i)
    # .Alias = MPEPropertyNames(i)
    # .Subfields = vbNullString
    # P = InStr(MPEPropertyNames(i), "Address")
    if P > 0 Then:
    # .Subfields = AddressSubfields(j)
    # MPEItems.Add oneColumnElement
    # End With                                 ' oneColumnElement

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindContact
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findcontact():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.FindContact"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    # Dim retrycount As Long
    # With whatWeLookFor
    # FindContact = 0
    # i = InStr(.FullName, ",")
    if i > 0 Then:
    # FindContact = FindContact + 1
    # .Lastname = Trim(Left(.FullName, i - 1))
    # .Firstname = Trim(Replace(Mid(.FullName, i + 1), ",", vbNullString))
    else:
    # FindContact = FindContact + 2
    # i = InStrRev(.FullName, b)
    if i = 0 Then:
    # .Firstname = vbNullString
    # .Lastname = .FullName
    else:
    # .Firstname = Trim(Left(.FullName, i - 1))
    # .Lastname = Trim(Mid(.FullName, i + 1))
    # Retry:
    # Set ciItem = MainFolderContacts.Items.Find("[FileAs] = " & Quote1(.FullName) & b)
    if ciItem Is Nothing Then:
    # FindContact = FindContact + 4
    # Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.FullName) & b)
    if ciItem Is Nothing Then:
    # Set ciItem = MainFolderContacts.Items.Find("[Fullname] = " & Quote1(.FullName) & b)
    if ciItem Is Nothing Then:
    # Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.Lastname) & b)
    if Not ciItem Is Nothing Then:
    # ciItem.Firstname = Trim(Replace(ciItem.Firstname, ",", vbNullString))
    # ciItem.Lastname = Trim(Replace(ciItem.Lastname, ",", vbNullString))
    if ciItem.Firstname <> .Firstname And .Firstname <> .Lastname Then:
    if Trunc(1, ciItem.Firstname, b) = .Firstname Then:
    # .Firstname = ciItem.Firstname ' more complete than MPE
    else:
    # Message = "kein Eintrag fr " & .Lastname & ", " & .Firstname & " in " & MainFolderContacts.FolderPath
    if retrycount < 1 Then:
    # retrycount = retrycount + 1
    # GoTo Retry
    else:
    # GoTo lastResort
    else:
    # lastResort:
    # FindContact = FindContact + 8
    if .Firstname = .Lastname Then:
    # ciItem.Firstname = vbNullString
    # .Lastname = vbNullString               'tricky, after swap the .Firstname is vbNullString
    # Call Swap(.Lastname, .Firstname)
    # Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.Lastname) & b)
    if Not ciItem Is Nothing Then:
    if ciItem.Firstname <> .Firstname Then:
    # Message = "kein Eintrag fr " & .Firstname & ", " & .Lastname & " in " & MainFolderContacts.FolderPath
    # GoTo skipitem
    if ciItem Is Nothing Then:
    # FindContact = FindContact + 16
    # Set ciItem = MainFolderContacts.Items.Find("[FullName] = " & Quote(.FullName) & b)
    if ciItem Is Nothing Then:
    # Message = "kein Eintrag fr " & Quote(.FullName) & " in " & Quote(MainFolderContacts.FolderPath)
    # skipitem:
    # FindContact = 0
    else:
    # Message = "item found"
    # ciItem.Firstname = Trim(Replace(ciItem.Firstname, ",", vbNullString))
    # End With                                     ' WhatWeLookFor

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub InitMPEColumnNames
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initmpecolumnnames():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.InitMPEColumnNames"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # MPEColumnNames = split(", Name, Nachname, Vorname, Mobil, Zuhause, " _
    # & "Geschftlich, Fax, Andere, " _
    # & "e-mail, 2. e-mail, 3. e-mail, Web, Adresse, Adresse (geschftlich), " _
    # & "Firma, Info, Geburtstag, Konto", ", ")
    # MPEPropertyNames = split(", FileAs, LastName, FirstName, MobileTelephoneNumber, " _
    # & "HomeTelephoneNumber, BusinessTelephoneNumber, BusinessFaxNumber, " _
    # & "OtherTelephoneNumber, Email1Address, Email2Address, Email3Address, " _
    # & "WebPage, HomeAddress, BusinessAddress, CompanyName, " _
    # & "Body, Birthday, User2", ", ")
    # AddressSubfields = split(", City, Street, Country, PostalCode", ", ")
    # Call ColumnStructure

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MapMPEtoContact
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mapmpetocontact():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.MapMPEtoContact"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim MPEcol As Long
    # Dim MPEcolName As String
    # MPEItemDiffs = vbNullString
    # With MPEitem
    # MPEcolName = MPEColumnNames(MPEcol)
    match MPEcol:
        case 1:
    # Call carefullAssign(MPEcolName, Ci.FullName, _
    # .FullName)
        case 2:
    # Call carefullAssign(MPEcolName, Ci.Lastname, _
    # .Lastname)
        case 3:
    # Call carefullAssign(MPEcolName, Ci.Firstname, _
    # .Firstname)
        case 4:
    # Call carefullAssign(MPEcolName, Ci.MobileTelephoneNumber, _
    # .MobileTelephoneNumber)
        case 5:
    # Call carefullAssign(MPEcolName, Ci.HomeTelephoneNumber, _
    # .HomeTelephoneNumber)
        case 6:
    # Call carefullAssign(MPEcolName, Ci.BusinessTelephoneNumber, _
    # .BusinessTelephoneNumber)
        case 7:
    # Call carefullAssign(MPEcolName, Ci.BusinessFaxNumber, _
    # .BusinessFaxNumber)
        case 8:
    # Call carefullAssign(MPEcolName, Ci.OtherTelephoneNumber, _
    # .OtherTelephoneNumber)
        case 9:
    # Call carefullAssign(MPEcolName, Ci.Email1Address, _
    # .Email1Address)
        case 10:
    # Call carefullAssign(MPEcolName, Ci.Email2Address, _
    # .Email2Address)
        case 11:
    # Call carefullAssign(MPEcolName, Ci.Email3Address, _
    # .Email3Address)
        case 12:
    # Call carefullAssign(MPEcolName, Ci.WebPage, _
    # .WebPage)
        case 13:
    # Call carefullAssign(MPEcolName, Ci.HomeAddress, _
    # .HomeAddress)
        case 14:
    # Call carefullAssign(MPEcolName, Ci.BusinessAddress, _
    # .BusinessAddress)
        case 15:
    # Call carefullAssign(MPEcolName, Ci.CompanyName, _
    # .CompanyName)
        case 16:
    # Call carefullAssign(MPEcolName, Ci.Body, _
    # .Body)
        case 17:
    # Call carefullAssign(MPEcolName, Ci.Birthday, _
    # .Birthday)
        case 18:
    # Call carefullAssign(MPEcolName, Ci.User2, _
    # .User2)
        case _:
    # DoVerify False
    # Call NameCheck(Ci)
    if LenB(.Firstname) = 0 Then:
    # Call carefullAssign("FileAs", Ci.FileAs, .Lastname)
    else:
    # Call carefullAssign("FileAs", Ci.FileAs, .Lastname _
    # & ", " & .Firstname)
    # End With                                     ' MPEitem

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MPEdecode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mpedecode():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.MPEdecode"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim TableValue As String
    # Dim MPEcol As Long
    # Set MPEitem = New cMPEObject
    # With MPEitem
    # TableValue = aWS.Cells(MPEline, MPEcol)
    match MPEcol:
        case 1:
    # .FullName = TableValue
        case 2:
    # .Lastname = TableValue
        case 3:
    # .Firstname = TableValue
        case 4:
    # .MobileTelephoneNumber = TableValue
        case 5:
    # .HomeTelephoneNumber = TableValue
        case 6:
    # .BusinessTelephoneNumber = TableValue
        case 7:
    # .BusinessFaxNumber = TableValue
        case 8:
    # .OtherTelephoneNumber = TableValue
        case 9:
    # .Email1Address = TableValue
        case 10:
    # .Email2Address = TableValue
        case 11:
    # .Email3Address = TableValue
        case 12:
    # .WebPage = TableValue
        case 13:
    # .HomeAddress = TableValue
        case 14:
    # .BusinessAddress = TableValue
        case 15:
    # .CompanyName = TableValue
        case 16:
    # .Body = TableValue
        case 17:
    # .Birthday = TableValue
        case 18:
    # .User2 = TableValue
        case _:
    # DoVerify False
    # End With                                     ' MPEitem

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MPEinit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mpeinit():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.MPEinit"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim i As Long
    # Dim xw As cXLTab

    # Set xw = New cXLTab
    # Call InitMPEColumnNames
    if xlApp Is Nothing Then:
    # GoTo startOver
    if xlC Is Nothing Then:
    # startOver:
    # Call XlgetApp
    # wrongfile:
    # Call DisplayExcel(xw, unconditionallyShow:=True, xlY:=W.xlTSheet)
    if rsp = vbCancel Then:
    if TerminateRun Then:
    # GoTo ProcReturn
    elif rsp = vbNo Then:
    # GoTo ProcReturn
    else:
    # Set xlC = xlApp.ActiveWorkbook
    if xlC Is Nothing Then:
    # GoTo wrongfile
    # i = 1
    if DoVerify(xlC.Worksheets.Count > 0, "No worksheets in Workbook " & xlC.FullName) Then:
    # GoTo FuncExit

    # Set W = xlC.Worksheets(i)
    if verifyMPEheader(W) Then:
    # GoTo gotsheet
    if i < xlC.Worksheets.Count Then:
    # i = i + 1
    # GoTo nextsheet
    else:
    # GoTo wrongfile
    # gotsheet:

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MPEmap
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def mpemap():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "PxMappings.MPEmap"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="PxMappings")

    # Dim MPEline As Long
    # Dim MPEitem As cMPEObject
    # Dim Matchrule As Long

    # IsEntryPoint = True

    # bDefaultButton = "Go"
    # Set MainFolderContacts = ActiveExplorer.CurrentFolder
    # While MainFolderContacts.DefaultItemType <> olContactItem
    # Call PickAFolder(1, "bitte whlen Sie den Kontakte-Ordner ", _
    # "Auswahl des Orders fr den Kontakte-Abgleich " _
    # & "mit MPE/Excel", _
    # "OK", "Cancel")
    # Set MainFolderContacts = Folder(1)
    # Wend

    # Call MPEinit
    # Call DisplayExcel(xlC, unconditionallyShow:=True)

    # Set aID(1).idObjItem = MainFolderContacts.Items(1)
    # Call BestObjProps(MainFolderContacts, aID(1).idObjItem, withValues:=True)

    # Call MPEdecode(W, MPEitem, MPEline)
    if LenB(MPEitem.FullName) = 0 Then:
    # GoTo skipitem
    # seekitem:
    # MPEchanged = False
    # Matchrule = FindContact(aID(1).idObjItem, MPEitem)
    if Matchrule = 0 Then:
    # "Kontakt neu anlegen?", vbYesNoCancel)
    if rsp = vbNo Then:
    # GoTo skipitem
    elif rsp = vbCancel Then:
    # GoTo stopall
    # ' we got a yes
    # ' we add a new item
    # Set aID(2).idObjItem = MainFolderContacts.Items.Add
    # Call MapMPEtoContact(aID(2).idObjItem, MPEitem)
    # MPEchanged = True
    else:
    # Set aID(2).idObjItem = aID(1).idObjItem ' 1 is the original contact item, 2 will contain the change

    # Call GetAobj(1)
    # objTypName = DecodeObjectClass(getValues:=False)
    # Call GetAobj(2)
    # objTypName = DecodeObjectClass(getValues:=False)
    # AllItemDiffs = vbNullString
    # Call DecodeAllPropertiesFor2Items(False, 1)

    # AllItemDiffs = vbNullString
    # Call MapMPEtoContact(aID(1).idObjItem, MPEitem) ' preset with changes
    # mustDecodeRest = False
    # Call DecodeAllPropertiesFor2Items(2, True)

    if MPEchanged Then:
    if Not ItemIdentity(True) Or WorkItemMod(2) Then:
    if Not aID(2).idObjItem.Saved Then:
    # aID(2).idObjItem.Save
    # Call LogEvent("Contact saved: " & aID(2).idObjItem.FileAs & " in " _
    # & MainFolderContacts.FolderPath, eLall)
    # skipitem:
    # O.xlTabIsEmpty = 1                       ' start a new excel workbook
    # stopall:
    # Call ClearWorkSheet(xlC, O)
    # StopRecursionNonLogged = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function verifyMPEheader
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def verifympeheader():
    # Dim zErr As cErr
    # Const zKey As String = "PxMappings.verifyMPEheader"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    # Dim i As Long
    if LCase(aSheet.Cells(1, i)) <> LCase(MPEColumnNames(i)) Then:
    # verifyMPEheader = False
    # GoTo ProcReturn
    # verifyMPEheader = True

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:


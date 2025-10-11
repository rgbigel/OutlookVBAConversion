Attribute VB_Name = "PxMappings"
Option Explicit
' ALL Columns from MPE, in that ORDER
'    1        2          3        4        5       6             7     8
' Name    Nachname    Vorname   Mobil   Zuhause Geschäftlich    Fax Andere
'    9       10         11       12       13      14            15
' e-mail  2. e-mail   3. e-mail Web Adresse Adresse (geschäftlich)  Firma
'  16          17       18
' Info    Geburtstag  Konto

Public MPEColumnNames
Public MPEPropertyNames
Public MPEItems As Collection
Public AddressSubfields
Public MPEchanged As Boolean

'---------------------------------------------------------------------------------------
' Method : Sub carefullAssign
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub carefullAssign(ColName As String, ByRef ciString As Variant, MPEstring As String)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "PxMappings.carefullAssign"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)
    ' do not forget to Call N_ReturnProc(zErr)

Dim oldValue As Variant
    oldValue = ciString
    If MPEstring = "###" Then
        MPEstring = vbNullString
        GoTo setit
    End If
    If (ciString = vbNullString Or InStr(ciString, "###") = 1) _
    And LenB(MPEstring) > 0 _
    And ciString <> MPEstring Then
setit:
        ciString = MPEstring
        MPEItemDiffs = MPEItemDiffs & vbCrLf _
                        & "changed " & ColName & "=" _
                        & Quote(oldValue) & "  to " _
                        & Quote(MPEstring) & vbCrLf
        MPEchanged = True
        WorkItemMod(2) = True
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' PxMappings.carefullAssign

'---------------------------------------------------------------------------------------
' Method : Sub ColumnStructure
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ColumnStructure()
Dim zErr As cErr
Const zKey As String = "PxMappings.ColumnStructure"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long, P As Long, j As Long
Dim oneColumnElement As cNumbItem
Dim onePropertyElement As cNumbItem

    Set MPEItems = New Collection
    For i = 1 To UBound(MPEColumnNames)
        Set oneColumnElement = New cNumbItem
        Set onePropertyElement = New cNumbItem
        With oneColumnElement
            .NuIndex = i
            .Key = MPEColumnNames(i)
            .Alias = MPEPropertyNames(i)
            .Subfields = vbNullString
            P = InStr(MPEPropertyNames(i), "Address")
            If P > 0 Then
                .Subfields = AddressSubfields(j)
            End If
            MPEItems.Add oneColumnElement
        End With                                 ' oneColumnElement
    Next i

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.ColumnStructure

'---------------------------------------------------------------------------------------
' Method : Function FindContact
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindContact(ByRef ciItem As ContactItem, whatWeLookFor As cMPEObject) As Long
Dim zErr As cErr
Const zKey As String = "PxMappings.FindContact"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim retrycount As Long
    With whatWeLookFor
        FindContact = 0
        i = InStr(.FullName, ",")
        If i > 0 Then
            FindContact = FindContact + 1
            .Lastname = Trim(Left(.FullName, i - 1))
            .Firstname = Trim(Replace(Mid(.FullName, i + 1), ",", vbNullString))
        Else
            FindContact = FindContact + 2
            i = InStrRev(.FullName, b)
            If i = 0 Then
                .Firstname = vbNullString
                .Lastname = .FullName
            Else
                .Firstname = Trim(Left(.FullName, i - 1))
                .Lastname = Trim(Mid(.FullName, i + 1))
            End If
        End If
Retry:
        Set ciItem = MainFolderContacts.Items.Find("[FileAs] = " & Quote1(.FullName) & b)
        If ciItem Is Nothing Then
            FindContact = FindContact + 4
            Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.FullName) & b)
            If ciItem Is Nothing Then
                Set ciItem = MainFolderContacts.Items.Find("[Fullname] = " & Quote1(.FullName) & b)
            End If
        End If
        If ciItem Is Nothing Then
            Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.Lastname) & b)
            If Not ciItem Is Nothing Then
                ciItem.Firstname = Trim(Replace(ciItem.Firstname, ",", vbNullString))
                ciItem.Lastname = Trim(Replace(ciItem.Lastname, ",", vbNullString))
                If ciItem.Firstname <> .Firstname And .Firstname <> .Lastname Then
                    If Trunc(1, ciItem.Firstname, b) = .Firstname Then
                        .Firstname = ciItem.Firstname ' more complete than MPE
                    Else
                        Message = "kein Eintrag für " & .Lastname & ", " & .Firstname & " in " & MainFolderContacts.FolderPath
                        If retrycount < 1 Then
                            retrycount = retrycount + 1
                            GoTo Retry
                        Else
                            GoTo lastResort
                        End If
                    End If
                End If
            Else
lastResort:
                FindContact = FindContact + 8
                If .Firstname = .Lastname Then
                    ciItem.Firstname = vbNullString
                    .Lastname = vbNullString               'tricky, after swap the .Firstname is vbNullString
                End If
                Call Swap(.Lastname, .Firstname)
                Set ciItem = MainFolderContacts.Items.Find("[LastName] = " & Quote1(.Lastname) & b)
                If Not ciItem Is Nothing Then
                    If ciItem.Firstname <> .Firstname Then
                        Message = "kein Eintrag für " & .Firstname & ", " & .Lastname & " in " & MainFolderContacts.FolderPath
                        GoTo skipitem
                    End If
                End If
            End If
        End If
        If ciItem Is Nothing Then
            FindContact = FindContact + 16
            Set ciItem = MainFolderContacts.Items.Find("[FullName] = " & Quote(.FullName) & b)
        End If
        If ciItem Is Nothing Then
            Message = "kein Eintrag für " & Quote(.FullName) & " in " & Quote(MainFolderContacts.FolderPath)
skipitem:
            FindContact = 0
        Else
            Message = "item found"
            ciItem.Firstname = Trim(Replace(ciItem.Firstname, ",", vbNullString))
        End If
    End With                                     ' WhatWeLookFor

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' PxMappings.FindContact

'---------------------------------------------------------------------------------------
' Method : Sub InitMPEColumnNames
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitMPEColumnNames()
Dim zErr As cErr
Const zKey As String = "PxMappings.InitMPEColumnNames"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    MPEColumnNames = split(", Name, Nachname, Vorname, Mobil, Zuhause, " _
                         & "Geschäftlich, Fax, Andere, " _
                         & "e-mail, 2. e-mail, 3. e-mail, Web, Adresse, Adresse (geschäftlich), " _
                         & "Firma, Info, Geburtstag, Konto", ", ")
    MPEPropertyNames = split(", FileAs, LastName, FirstName, MobileTelephoneNumber, " _
                           & "HomeTelephoneNumber, BusinessTelephoneNumber, BusinessFaxNumber, " _
                           & "OtherTelephoneNumber, Email1Address, Email2Address, Email3Address, " _
                           & "WebPage, HomeAddress, BusinessAddress, CompanyName, " _
                           & "Body, Birthday, User2", ", ")
    AddressSubfields = split(", City, Street, Country, PostalCode", ", ")
    Call ColumnStructure

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.InitMPEColumnNames

'---------------------------------------------------------------------------------------
' Method : Sub MapMPEtoContact
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MapMPEtoContact(Ci As ContactItem, MPEitem As cMPEObject)
Dim zErr As cErr
Const zKey As String = "PxMappings.MapMPEtoContact"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim MPEcol As Long
Dim MPEcolName As String
    MPEItemDiffs = vbNullString
    With MPEitem
        For MPEcol = 1 To UBound(MPEColumnNames)
            MPEcolName = MPEColumnNames(MPEcol)
            Select Case MPEcol
            Case 1
                Call carefullAssign(MPEcolName, Ci.FullName, _
                                    .FullName)
            Case 2
                Call carefullAssign(MPEcolName, Ci.Lastname, _
                                    .Lastname)
            Case 3
                Call carefullAssign(MPEcolName, Ci.Firstname, _
                                    .Firstname)
            Case 4
                Call carefullAssign(MPEcolName, Ci.MobileTelephoneNumber, _
                                    .MobileTelephoneNumber)
            Case 5
                Call carefullAssign(MPEcolName, Ci.HomeTelephoneNumber, _
                                    .HomeTelephoneNumber)
            Case 6
                Call carefullAssign(MPEcolName, Ci.BusinessTelephoneNumber, _
                                    .BusinessTelephoneNumber)
            Case 7
                Call carefullAssign(MPEcolName, Ci.BusinessFaxNumber, _
                                    .BusinessFaxNumber)
            Case 8
                Call carefullAssign(MPEcolName, Ci.OtherTelephoneNumber, _
                                    .OtherTelephoneNumber)
            Case 9
                Call carefullAssign(MPEcolName, Ci.Email1Address, _
                                    .Email1Address)
            Case 10
                Call carefullAssign(MPEcolName, Ci.Email2Address, _
                                    .Email2Address)
            Case 11
                Call carefullAssign(MPEcolName, Ci.Email3Address, _
                                    .Email3Address)
            Case 12
                Call carefullAssign(MPEcolName, Ci.WebPage, _
                                    .WebPage)
            Case 13
                Call carefullAssign(MPEcolName, Ci.HomeAddress, _
                                    .HomeAddress)
            Case 14
                Call carefullAssign(MPEcolName, Ci.BusinessAddress, _
                                    .BusinessAddress)
            Case 15
                Call carefullAssign(MPEcolName, Ci.CompanyName, _
                                    .CompanyName)
            Case 16
                Call carefullAssign(MPEcolName, Ci.Body, _
                                    .Body)
            Case 17
                Call carefullAssign(MPEcolName, Ci.Birthday, _
                                    .Birthday)
            Case 18
                Call carefullAssign(MPEcolName, Ci.User2, _
                                    .User2)
            Case Else
                DoVerify False
            End Select
        Next MPEcol
        Call NameCheck(Ci)
        If LenB(.Firstname) = 0 Then
            Call carefullAssign("FileAs", Ci.FileAs, .Lastname)
        Else
            Call carefullAssign("FileAs", Ci.FileAs, .Lastname _
                                                  & ", " & .Firstname)
        End If
    End With                                     ' MPEitem

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.MapMPEtoContact

'---------------------------------------------------------------------------------------
' Method : Sub MPEdecode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MPEdecode(aWS As Worksheet, MPEitem As cMPEObject, MPEline As Long)
Dim zErr As cErr
Const zKey As String = "PxMappings.MPEdecode"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim TableValue As String
Dim MPEcol As Long
    Set MPEitem = New cMPEObject
    With MPEitem
        For MPEcol = 1 To UBound(MPEColumnNames)
            TableValue = aWS.Cells(MPEline, MPEcol)
            Select Case MPEcol
            Case 1
                .FullName = TableValue
            Case 2
                .Lastname = TableValue
            Case 3
                .Firstname = TableValue
            Case 4
                .MobileTelephoneNumber = TableValue
            Case 5
                .HomeTelephoneNumber = TableValue
            Case 6
                .BusinessTelephoneNumber = TableValue
            Case 7
                .BusinessFaxNumber = TableValue
            Case 8
                .OtherTelephoneNumber = TableValue
            Case 9
                .Email1Address = TableValue
            Case 10
                .Email2Address = TableValue
            Case 11
                .Email3Address = TableValue
            Case 12
                .WebPage = TableValue
            Case 13
                .HomeAddress = TableValue
            Case 14
                .BusinessAddress = TableValue
            Case 15
                .CompanyName = TableValue
            Case 16
                .Body = TableValue
            Case 17
                .Birthday = TableValue
            Case 18
                .User2 = TableValue
            Case Else
                DoVerify False
            End Select
        Next MPEcol
    End With                                     ' MPEitem

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.MPEdecode

'---------------------------------------------------------------------------------------
' Method : Sub MPEinit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MPEinit()
Dim zErr As cErr
Const zKey As String = "PxMappings.MPEinit"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim xw As cXLTab

    Set xw = New cXLTab
    Call InitMPEColumnNames
    If xlApp Is Nothing Then
        GoTo startOver
    End If
    If xlC Is Nothing Then
startOver:
        Call XlgetApp
wrongfile:
        Call DisplayExcel(xw, unconditionallyShow:=True, xlY:=W.xlTSheet)
        rsp = MsgBox("Bitte die MPE-Datei in Excel öffnen", vbYesNoCancel)
        If rsp = vbCancel Then
            If TerminateRun Then
                GoTo ProcReturn
            End If
        ElseIf rsp = vbNo Then
            GoTo ProcReturn
        Else
            Set xlC = xlApp.ActiveWorkbook
            If xlC Is Nothing Then
                GoTo wrongfile
            End If
            i = 1
nextsheet:
            If DoVerify(xlC.Worksheets.Count > 0, "No worksheets in Workbook " & xlC.FullName) Then
                GoTo FuncExit
            End If
            
            Set W = xlC.Worksheets(i)
            If verifyMPEheader(W) Then
                GoTo gotsheet
            End If
            If i < xlC.Worksheets.Count Then
                i = i + 1
                GoTo nextsheet
            Else
                GoTo wrongfile
            End If
        End If
    End If
gotsheet:

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.MPEinit

'---------------------------------------------------------------------------------------
' Method : Sub MPEmap
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MPEmap()                                     ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "PxMappings.MPEmap"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="PxMappings")

Dim MPEline As Long
Dim MPEitem As cMPEObject
Dim Matchrule As Long

    IsEntryPoint = True
    
    bDefaultButton = "Go"
    Set MainFolderContacts = ActiveExplorer.CurrentFolder
    While MainFolderContacts.DefaultItemType <> olContactItem
        Call PickAFolder(1, "bitte wählen Sie den Kontakte-Ordner ", _
                         "Auswahl des Orders für den Kontakte-Abgleich " _
                       & "mit MPE/Excel", _
                         "OK", "Cancel")
        Set MainFolderContacts = Folder(1)
    Wend
    
    Call MPEinit
    Call DisplayExcel(xlC, unconditionallyShow:=True)
    
    Set aID(1).idObjItem = MainFolderContacts.Items(1)
    Call BestObjProps(MainFolderContacts, aID(1).idObjItem, withValues:=True)
    
    For MPEline = 2 To xlC.UsedRange.columns.Count + xlC.UsedRange.Column - 1
        Call MPEdecode(W, MPEitem, MPEline)
        If LenB(MPEitem.FullName) = 0 Then
            GoTo skipitem
        End If
seekitem:
        MPEchanged = False
        Matchrule = FindContact(aID(1).idObjItem, MPEitem)
        If Matchrule = 0 Then
            rsp = MsgBox(Message & vbCrLf & _
                         "Kontakt neu anlegen?", vbYesNoCancel)
            If rsp = vbNo Then
                GoTo skipitem
            ElseIf rsp = vbCancel Then
                GoTo stopall
            End If
            ' we got a yes
            ' we add a new item
            Set aID(2).idObjItem = MainFolderContacts.Items.Add
            Call MapMPEtoContact(aID(2).idObjItem, MPEitem)
            MPEchanged = True
        Else
            Set aID(2).idObjItem = aID(1).idObjItem ' 1 is the original contact item, 2 will contain the change
        End If
        
        Call GetAobj(1)
        objTypName = DecodeObjectClass(getValues:=False)
        Call GetAobj(2)
        objTypName = DecodeObjectClass(getValues:=False)
        AllItemDiffs = vbNullString
        Call DecodeAllPropertiesFor2Items(False, 1)
        
        AllItemDiffs = vbNullString
        Call MapMPEtoContact(aID(1).idObjItem, MPEitem) ' preset with changes
        mustDecodeRest = False
        Call DecodeAllPropertiesFor2Items(2, True)
        
        If MPEchanged Then
            If Not ItemIdentity(True) Or WorkItemMod(2) Then
                If Not aID(2).idObjItem.Saved Then
                    aID(2).idObjItem.Save
                    Call LogEvent("Contact saved: " & aID(2).idObjItem.FileAs & " in " _
                                & MainFolderContacts.FolderPath, eLall)
                End If
            End If
        End If
skipitem:
        O.xlTabIsEmpty = 1                       ' start a new excel workbook
    Next MPEline
stopall:
    Call ClearWorkSheet(xlC, O)
    StopRecursionNonLogged = False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' PxMappings.MPEmap

'---------------------------------------------------------------------------------------
' Method : Function verifyMPEheader
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function verifyMPEheader(aSheet As Excel.Worksheet) As Boolean
Dim zErr As cErr
Const zKey As String = "PxMappings.verifyMPEheader"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
    For i = 1 To UBound(MPEColumnNames)
        If LCase(aSheet.Cells(1, i)) <> LCase(MPEColumnNames(i)) Then
            verifyMPEheader = False
            GoTo ProcReturn
        End If
    Next i
    verifyMPEheader = True

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' PxMappings.verifyMPEheader



# Converted from O_Goodies.py

# Attribute VB_Name = "O_Goodies"
# Option Explicit

# Public Const GH = &H42
# Public Const CF_TEXT = 1
# Public Const MAXSIZE = 4096

# Public DataObject As MSForms.DataObject
# Public ClipBoardIsEmpty As cTriState

# Type PictureDescription
# height As Single
# width As Single
# End Type

# Dim arr                                          ' Variant for splitting texts

# Public ChangeAssignReverse As Boolean            ' request functions ...AssignIfChanged to reverse source and target
# Public ModThisTo As Variant                      ' new value in AssignIfChanged
# Public DecodedValue As Variant                   ' set by Function if the value can be determined as String
# Public AssignmentMode As Long                    ' HasValue Assignment mode:
# ' 0=Imposs., 1 = Set, 2 = direct Scalar,
# ' 3 = Object Default (Scalar) or result of 4 below
# ' 4 = ItemProperty value evaluation

# '---------------------------------------------------------------------------------------
# ' Method : Sub AddItemToList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def additemtolist():


    # Dim LI As cListItm

    # Set LI = New cListItm
    # LI.Index1 = i1
    # LI.MainId = ID
    # LI.Compares = comp
    # LI.Index2 = i2
    if ListContent Is Nothing Then:
    # Set ListContent = New Collection
    # ListCount = 0
    else:
    # ListCount = ListContent.Count
    # ListContent.Add LI
    # ListCount = ListContent.Count

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : Sub AddNumbItemToCollection
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def addnumbitemtocollection():

    # Const zKey As String = "O_Goodies.AddNumbItemToCollection"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)

    # Dim ItemProperty As cNumbItem
    # Set ItemProperty = New cNumbItem
    # ItemProperty.NuIndex = myColl.Count + 1
    # ItemProperty.Key = Name
    # Set ItemProperty.ValueOfItem = val

    # aBugTxt = "Add item to collection " & modifier & Name
    # Call Try
    # myColl.Add ItemProperty, Key:=modifier & Name
    # Catch

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function Append
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def append():


    # Dim L As Long
    # Dim listx As Variant
    # Dim aListx As Variant

    # Static myRecursionDepth As Long

    if myRecursionDepth = 0 Then:
    # StringMod = False
    # myRecursionDepth = myRecursionDepth + 1

    # ' append to front will not split at sep, because that would result in wrong ordering
    if LenB(sep) > 0 And InStr(S2, sep) > 0 And Not ToFront Then:
    # listx = split(S2, sep)
    for alistx in listx:
    if Not isEmpty(aListx) And LenB(aListx) > 0 Then:
    # myRecursionDepth = myRecursionDepth + 1 ' next level of recursion
    # S1 = Append(S1, (aListx), sep, always, ToFront)
    # myRecursionDepth = myRecursionDepth - 1
    # DoVerify myRecursionDepth > 0
    # Append = S1
    # GoTo FuncExit

    if Not always Then                           ' test if append creates a double entry:
    # L = InStr(1, sep & S1 & sep, sep & S2 & sep, vbTextCompare)
    if L > 0 Then                            ' string kommt vor:
    # ' schon enthalten: return original ohne Zufgen
    # Append = S1
    # GoTo FuncExit
    if LenB(S1) = 0 Then:
    # Append = S2
    else:
    if ToFront Then:
    # Append = S2 & sep & S1
    else:
    # Append = S1 & sep & S2

    # FuncExit:
    if Append <> S1 Then:
    # StringMod = True
    # myRecursionDepth = myRecursionDepth - 1


# '---------------------------------------------------------------------------------------
# ' Method : AppendTo
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: If string(s) to be inserted are new to first string, or always is specified:
# '               add it to front or back of first string.
# '          calls Append: String compare is via vbTextCompare (ignore case of chars)
# '---------------------------------------------------------------------------------------
def appendto():


    # Dim Result As String

    # StringMod = False
    # Result = Append(S1, S2, sep, always, ToFront)
    if StringMod Then:
    # S1 = Result


# '---------------------------------------------------------------------------------------
# ' Method : Function ArrayMatch
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def arraymatch():


    # ' check if the entries in testarray(String) occur anywhere in matchString
    # ' no WildCard match: string must occur at the end
    # Dim i As Long
    # Dim k As Long
    # Dim testword As String

    # testword = Replace(testArray(k), WildCard, vbNullString) ' remove WildCard (at end)
    if LenB(Trim(testword)) > 0 Then:
    # i = InStr(1, matchString, testword, CaseSensitive)
    if i > 0 Then                        ' it is somewhere:
    if Len(testword) = Len(testArray(k)) Then ' no WildCard in testarray(k):
    if i + Len(testword) = Len(matchString) Then ' at the end:
    # ArrayMatch = k           ' match condition OK
    # GoTo FuncExit
    else:
    # ArrayMatch = k               ' match condition OK, not at end
    # GoTo FuncExit
    # ' if we drop out, no match
    # ArrayMatch = LBound(testArray) - 1           ' not in array

    # FuncExit:


# '---------------------------------------------------------------------------------------
# ' Method : Sub ArrayOrderMaxMin
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def arrayordermaxmin():

    # Const zKey As String = "O_Goodies.ArrayOrderMaxMin"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim j As Long
    # Dim MaxOfIndex As Long
    # Dim MinOfIndex As Long

    # 'On Error GoTo 0

    # ' assume valid defaults
    # MaxOfIndex = LBound(AR)
    # MinOfIndex = MaxOfIndex
    # ' find a bigger one
    if AR(MaxOfIndex) < AR(j) Then:
    # MaxOfIndex = Max(j, MaxOfIndex)
    if AR(MinOfIndex) > AR(j) Then:
    # MinOfIndex = Min(j, MinOfIndex)
    if Not IsMissing(aPosMax) Then:
    # aPosMax = MaxOfIndex
    if Not IsMissing(aPosMin) Then:
    # aPosMin = MinOfIndex

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Methods: (type)AssignIfChanged
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Simple and more general Assignment with Check if any change done
# ' NOTE   : ByRef for Targets does not work for values WITHIN Objects/Classes/Types
# '          In that case, use this model:
# '        If AssignIfChanged(\(.*?), (.*?)\) Then
# '            If ChangeAssignReverse Then
# '                $2 = ModThisTo
# '            Else
# '                $1 = ModThisTo
# '            End If
# '        end if
# '---------------------------------------------------------------------------------------
def assignifchanged():


    if CStr(Source) <> CStr(Target) Then:
    if ChangeAssignReverse Then:
    # ModThisTo = Target
    else:
    # ModThisTo = Source
    # AssignIfChanged = True


# '---------------------------------------------------------------------------------------
# ' Method : CatStrings
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Variable number of Strings are concatenated with sep between (but not for last)
# '---------------------------------------------------------------------------------------
def catstrings():
    # Dim i As Long
    # Dim npart As String

    if LenB(sep) = 0 Then:
    # DoVerify False                           ' likely it won, "W.xlTSheet work in split later"

    # npart = Cats(i)
    if InStr(npart, sep) > 0 Then:
    # DoVerify False                       ' likely it won, "W.xlTSheet work in split later"
    # CatStrings = CatStrings & npart & sep
    # CatStrings = CatStrings & Cats(UBound(Cats))

# '---------------------------------------------------------------------------------------
# ' Method : Function CharsToHexString
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def charstohexstring():
    # Dim i As Long

    # CharsToHexString = CharsToHexString & Right("0" & Hex(Asc(Mid(S, i, 1))), 2)


# '---------------------------------------------------------------------------------------
# ' Method : Function CheckForAreaCodes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkforareacodes():
    # Const zKey As String = "O_Goodies.CheckForAreaCodes"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim ThisNumber As String
    # Dim minlen As Long
    # Dim addI As Long
    # ThisNumber = Trim(Replace(Replace(aString, b, vbNullString), "-", vbNullString))
    if RecognizeCode(ThisNumber, aCode, "711") Then:
    elif RecognizeCode(ThisNumber, aCode, "650") Then:
    # minlen = 4
    elif RecognizeCode(ThisNumber, aCode, "800") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "651") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "461") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "65") Then:
    # minlen = 4
    elif RecognizeCode(ThisNumber, aCode, "261") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "23") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "15") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "16") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "17") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "18") Then:
    # minlen = 3
    elif RecognizeCode(ThisNumber, aCode, "40") Then:
    elif RecognizeCode(ThisNumber, aCode, "30") Then:
    elif RecognizeCode(ThisNumber, aCode, "89") Then:
    elif RecognizeCode(ThisNumber, aCode, "228") Then:
    elif RecognizeCode(ThisNumber, aCode, "6131") Then:
    # minlen = 4
    else:
    # ' vorwahl und nummer sind nicht getrennt: keine Normalisierung mglich
    if Len(ThisNumber) > 4 Then:
    # aCode = Left(ThisNumber, 4)
    # ThisNumber = Mid(ThisNumber, 5)
    # addI = minlen - Len(aCode)
    if addI > 0 Then:
    # aCode = aCode & Left(ThisNumber, addI)
    # ThisNumber = Mid(ThisNumber, addI + 1)
    # CheckForAreaCodes = ThisNumber

    # FuncExit:
    # zErr.atFuncResult = CStr(CheckForAreaCodes)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function CheckForCountryCodes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkforcountrycodes():
    # Dim ThisNumber As String
    # Dim plusBug As String

    # ThisNumber = Tail(Trim(Replace(aString, "-", b)), _
    # Sat:="+", Front:=plusBug)
    if Left(aString, 1) = "+" And Left(ThisNumber, 1) <> "+" Then:
    # ThisNumber = "+" & ThisNumber

    if RecognizeCode(ThisNumber, aCode, "+49") Then:
    elif RecognizeCode(ThisNumber, aCode, "+48") Then:
    elif RecognizeCode(ThisNumber, aCode, "+43") Then:
    elif RecognizeCode(ThisNumber, aCode, "+34") Then:
    elif RecognizeCode(ThisNumber, aCode, "+33") Then:
    elif RecognizeCode(ThisNumber, aCode, "+1") Then:
    elif RecognizeCode(ThisNumber, aCode, "+352") Then:
    # CheckForCountryCodes = ThisNumber


# '---------------------------------------------------------------------------------------
# ' Method : Function CheckSimilarityIn
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checksimilarityin():
    # Dim i As Long
    # Dim chW As String

    if sim = "*" Then:
    # CheckSimilarityIn = True
    else:
    # chW = ch(i)
    # CheckSimilarityIn = IsSimilar(chW, sim)
    if CheckSimilarityIn Then:
    # Exit For

# '---------------------------------------------------------------------------------------
# ' Method : Sub ClearClipboard
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clearclipboard():


    # OpenClipboard (0&)
    # EmptyClipboard
    # CloseClipboard
    # ClipBoardIsEmpty = TristateTrue


# '---------------------------------------------------------------------------------------
# ' Method : Sub ClipBoard_SetData
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clipboard_setdata():

    # Const zKey As String = "O_Goodies.ClipBoard_SetData"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongLong
    # Dim hClipMemory As LongLong, x As LongLong

    # ' Allocate moveable global memory.
    # '-------------------------------------------
    # hGlobalMemory = GlobalAlloc(GH, Len(MyString) + 1)

    # ' Lock the block to get a far pointer
    # ' to this memory.
    # lpGlobalMemory = GlobalLock(hGlobalMemory)

    # ' Copy the string to this global memory.
    # lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    # ' Unlock the memory.
    if GlobalUnlock(hGlobalMemory) <> 0 Then:
    print('Could not unlock memory location. Copy aborted.')
    # GoTo OutOfHere2

    # ' Open the Clipboard to copy data to.
    if OpenClipboard(0&) = 0 Then:
    print('Could not open the Clipboard. Copy aborted.')
    # GoTo zExit

    # ' Clear the Clipboard.
    # x = EmptyClipboard()

    # ' Copy the data to the Clipboard.
    # hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

    # OutOfHere2:

    if CloseClipboard() = 0 Then:
    print('Could not close Clipboard.')

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : ColumnPrint
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Print using columns forcing length
# '---------------------------------------------------------------------------------------
def columnprint():
    # Static aCol As Variant
    # Dim ValArr As Variant
    # Dim i As Long
    # Dim cL As Long
    # Dim lne As String

    if IsMissing(columns) Then:
    if TypeName(aCol) <> "String()" Then Stop:
    else:
    # aCol = split(columns, sep)
    if Not IsMissing(values) Then:
    if LenB(values) > 0 Then:
    # ValArr = split(values, sep)
    if UBound(ValArr) > UBound(aCol) Then Stop:

    # cL = aCol(i)
    if cL < 0 Then:
    # lne = lne & RString(ValArr(i), -cL)
    else:
    # lne = lne & LString(ValArr(i), cL)
    print(Debug.Print lne)

# '---------------------------------------------------------------------------------------
# ' Method : Sub Combo_Define_DatumsBedingungen
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def combo_define_datumsbedingungen():


    # Combo.addItem "keine Datumsbeschrnkung"
    # Combo.addItem "heute"
    # Combo.addItem "ab gestern"
    # Combo.addItem "letzte Woche"
    # Combo.addItem "letzte 30 Tage"
    # Combo.BoundColumn = 0
    # Combo.ListIndex = 0


# '---------------------------------------------------------------------------------------
# ' Method : Function CompareNumString
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def comparenumstring():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
    # Const zKey As String = "O_Goodies.CompareNumString"
    # Dim zErr As cErr

    # Dim i As Long
    # Dim minlen As Long
    # Dim aChar As String
    # Dim bchar As String
    # Dim cComp As Long
    # Dim numComp As Boolean
    # Dim l1 As Long
    # Dim l2 As Long
    # Dim maxlen As Long
    # Dim aLong As Long
    # Dim bLong As Long

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)

    if Len(str1) = Len(str2) Then:
    # GoTo simple
    # CompareNumString = 0                         ' equal as default
    # l1 = Len(str1)
    # l2 = Len(str2)
    # minlen = Min(l1, l2)
    # maxlen = Max(l1, l2)
    # aChar = Mid(str1, i, 1)
    # bchar = Mid(str2, i, 1)
    if InStr("0123456789.", aChar) > 0 _:
    # And InStr("0123456789.", bchar) > 0 Then
    # numComp = True
    else:
    # numComp = False

    if numComp Then:
    if Not (aChar = vbNullString Or aChar = ".") Then:
    # aLong = aLong * 10 + CLng(aChar)
    if Not (bchar = vbNullString Or bchar = ".") Then:
    # bLong = bLong * 10 + CLng(bchar)
    if aLong = bLong Then:
    # cComp = 0
    elif aLong < bLong Then:
    # cComp = -1
    else:
    # cComp = 1
    else:
    # aLong = 0
    # bLong = 0
    # cComp = StrComp(aChar, bchar, comparemode)

    if cComp = 0 Then                        ' same:
    # ' same up to now
    else:
    if numComp Then:
    if aChar = bchar Then:
    # cComp = 0
    else:
    # GoTo NE
    else:
    # ' cComp <> 0
    # NE:
    # CompareNumString = cComp
    if numComp Then:
    if i = minlen Then:
    if l1 < maxlen Or l2 < maxlen Then:
    # i = i + 1
    # GoTo nextdigit
    else:
    # GoTo funxit
    if i < Max(Len(str1), Len(str2)) + 1 Then:
    # simple:
    # CompareNumString = StrComp(str1, str2, comparemode)
    # funxit:

    # FuncExit:
    # zErr.atFuncResult = CStr(CompareNumString)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function CountStringOccurrences
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def countstringoccurrences():


    # Dim L As Long
    # Dim T As String

    if LenB(S) = 0 Or LenB(P) = 0 Then:
    # GoTo ProcRet
    # L = Len(S)
    # T = Replace(T, P, vbNullString)                        ' remove all occurrences of p in t
    # CountStringOccurrences = (L - Len(T)) / Len(P) ' use number of occurrences removed

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function CreateFolderIfNotExists
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def createfolderifnotexists():
    # Dim zErr As cErr
    # Const zKey As String = "O_Goodies.CreateFolderIfNotExists"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:=strFolderName)

    # Dim ParentFolder As Folder
    # Dim ParentFolders As Folders

    if underWhat.Class = olFolder Then:
    # Set ParentFolder = underWhat
    # Set ParentFolders = underWhat.Folders
    else:
    # Set ParentFolders = underWhat

    # aBugTxt = "Get or use Folder " & Quote1(strFolderName)
    # Call Try(testAll)                               ' Try anything, autocatch
    # Set CreateFolderIfNotExists = ParentFolders(strFolderName)
    if Not CatchNC Then:
    # Call LogEvent("Folder " & CreateFolderIfNotExists.Name & " exists")
    # Call ErrReset(4)
    # GoTo FuncExit                            ' no problem, file exists

    if StrComp(Hex$(E_AppErr.errNumber), "8004010F", vbTextCompare) = 0 Then:
    # ' ....wegen der Lesbarkeit vergleiche ich hier:
    # ' .....If Hex$(Err.Number) = "8004010F" Then
    # '  der "echte" Fehlercode ist  -2147221233 == &H8004010F
    # ' ..so (in Hex) kann ich den wiedererkennen. Bedeutet: Object nicht gefunden
    # '  wenn kein Type als Parameter mitgegeben, dann als Type "Mails" setzen
    if DefaultItemType = 0 Then:
    # DefaultItemType = olFolderInbox
    # aBugTxt = "Create Folder of type " & DefaultItemType & " and add to ParentFolders"
    # Set CreateFolderIfNotExists = _
    # ParentFolders.Add(strFolderName, DefaultItemType) ' , olMail)
    else:
    # ' was immer sonst passiert sein mag.. ich habs nicht abgefangen..
    # Call TerminateApp                        ' hier also Crash & Burn...
    # ' falls das Ordner-Neuanlegen in die Grtze geht...
    if Catch Then:
    # DoVerify False, "even in debugmode!"

    # FuncExit:
    # zErr.atFuncResult = CStr(CreateFolderIfNotExists)
    # Set ParentFolder = Nothing
    # Set ParentFolders = Nothing

    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Function cvExcelVal
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def cvexcelval():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.cvExcelVal"

    # cvExcelVal = val
    if LenB(CStr(val)) = 0 Then:
    # val = Chr(0)

    if val = Chr(0) _:
    # Or IsDate(val) _
    # Or IsNumeric(val) Then
    # ' exit function
    else:
    # cvExcelVal = CStr(val)
    if Left(cvExcelVal, 1) <> "'" Then       ' Do not double that:
    # cvExcelVal = "'" & CStr(val)

    # FuncExit:


# ' Englische bools aus Deutsch
def debooltoen():
    match LCase(b):
        case "true", "wahr":
    # DeBoolToEn = "True"
        case "false", "falsch":
    # DeBoolToEn = "False"
        case _:
    print(Debug.Print b & " ist kein bool'scher Wert")
    if DebugMode Then DoVerify False:

# '---------------------------------------------------------------------------------------
# ' Method : DecodeSpecialProperties
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Decodes Properties with unusual Values
# '---------------------------------------------------------------------------------------
def decodespecialproperties():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.DecodeSpecialProperties"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim dInfo As cInfo

    # With hInfo
    match PropName:
        case "Parent":
    # ' special parent olFolder: only care about folder path
    if .iValue.Value.Class = olFolder Then:
    # Set dInfo = .DrillDown(.iValue.Value)
    # dInfo.DecodedStringValue = dInfo.iValue.FolderPath
    # dInfo.iAssignmentMode = 1
    # dInfo.iType = vbString
    # dInfo.iTypeName = "ParentFolder"
    # Set hInfo = dInfo
    # DecodeSpecialProperties = True

        case "Actions", "Attachments", "UserProperties", "Recipients", _:
    # "ReplyRecipients", "Links", "Conflicts" ' all (known) Properties with Count
    # Set dInfo = .DrillDown(.iValue.Value)
    # dInfo.iTypeName = PropName & "Count"
    # dInfo.iArraySize = dInfo.iValue.Count
    # dInfo.iAssignmentMode = 1
    # dInfo.iIsArray = True
    # .DecodedStringValue = "{} " & dInfo.iArraySize & " values"
    # Set hInfo = dInfo
    # DecodeSpecialProperties = True
        case "Nothing":
    # Set dInfo = .DrillDown(Nothing)
    # dInfo.iTypeName = PropName
    # dInfo.iAssignmentMode = 2
    # dInfo.iType = vbNull
    # DecodeSpecialProperties = True
        case _:
    # DecodeSpecialProperties = False
    # End With                                     ' hInfo

    # FuncExit:
    # Set dInfo = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function DoubleInternalQuotes
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def doubleinternalquotes():
    # Dim Q As String

    # Q = Left(DQ, 1)                              ' use ONLY on nonquoted strings s
    # DoubleInternalQuotes = Replace(S, Q, Q & Q)  ' double internal quotes

# '---------------------------------------------------------------------------------------
# ' Method : Sub DumpAllPushDicts
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def dumpallpushdicts():

    # Const zKey As String = "O_Goodies.DumpAllPushDicts"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True
    # Dim aDictVar As Variant
    # Dim aValue As Variant
    # Dim aVobject As Variant
    # Dim aStackDict As Dictionary
    # Dim TupleNameValue As String
    # Dim i As Long

    # aBugTxt = "Dumping D_PushDict"
    # Call Try(allowAll)                              ' Try anything, autocatch, Err.Clear
    for adictvar in d_pushdict:
    # Set aStackDict = D_PushDict.Item(aDictVar)
    print(Debug.Print)
    print(Debug.Print "Dictionary: " & aDictVar)
    # i = 0
    for avalue in astackdict:
    # i = i + 1
    if VarType(aValue) = vbObject Then:
    if aStackDict.Exists(aValue.Key) Then:
    # Set aVobject = aStackDict.Item(aValue.Key)
    if isEmpty(aVobject) Then:
    # aVobject = "Empty Key"
    else:
    # aVobject = "has no key"
    else:
    if aStackDict.Exists(aValue) Then:
    if VarType(aValue) = vbObject Then:
    # Set aVobject = aValue
    else:
    if VarType(aStackDict.Item(aValue)) = vbObject Then:
    # Set aVobject = aStackDict.Item(aValue)
    else:
    # aVobject = CVar(aStackDict.Item(aValue))
    if isEmpty(aVobject) Then:
    # aVobject = "Empty"
    else:
    if VarType(aVobject) = vbObject Then:
    if TypeName(aVobject) = "cTuple" Then:
    # TupleNameValue = ", TupleNameValue=" & aVobject.TupleNameValue
    else:
    # aVobject = "TypeName '" & TypeName(aVobject) _
    # & "' not supported"
    else:
    # aVobject = "Trivial=" & CStr(aValue)
    else:
    # aVobject = "no item value"
    print(Debug.Print i, "StackObject_TypeName=" & TypeName(aValue);)
    print(Debug.Print ", Key=" & aValue.Key;)
    print(Debug.Print ", objKey=" & aVobject;)
    print(Debug.Print TupleNameValue & ", ValueStr=" & CStr(aValue))
    print(Debug.Print)
    # Catch

    # FuncExit:
    # Call ErrReset(0)
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EnvironmentPrintout
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def environmentprintout():
    # Dim UmgZF As String
    # Dim Indx As Long
    # Dim PfadLnge As String

    # Indx = 1                                     ' Index mit 1 initialisieren.
    # Do
    # UmgZF = Environ(Indx)                    ' Umgebungsvariable
    # PfadLnge = PfadLnge & vbCrLf & UmgZF
    # Indx = Indx + 1                          ' Kein PATH-Eintrag,
    # Loop Until UmgZF = vbNullString

    print(Debug.Print PfadLnge)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ErrReset
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: ReSet Error information / this does NOT clear information not related to Err
# '---------------------------------------------------------------------------------------
def errreset():
    # Const zKey As String = "O_GoodiesCall ErrReset"

    # Dim ErrOnEntry As Long
    # ErrOnEntry = Err.Number

    # StackDebug = Abs(StackDebugOverride)
    match upToMode:
        case 0                                       ' IgnoreUnhandledError is NOT changed:
    # aBugTxt = vbNullString
    # Call Try(Empty)                          ' Includes:  ErrSnoCatch, ZErrSnoCatch
    if MayChangeErr And ErrOnEntry <> 0 Then:
    # GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
    # GoTo FuncExit
        case 1:
    # MayChangeErr = True                      ' always clears Err, ErrorCaught is NOT cleared
    # IgnoreUnhandledError = False
    # Call Try(Empty)                          ' Includes:  ErrSnoCatch, ZErrSnoCatch
    # GoTo ShowChange
        case 2:
    # MayChangeErr = True                      ' always clears Err, ErrorCaught is NOT cleared
    # IgnoreUnhandledError = False
    # Call T_DC.N_ClearTermination             ' Includes:  ErrSnoCatch, ZErrSnoCatch
    if ErrOnEntry <> 0 Then:
    # GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
    # GoTo FuncExit
        case 3:
    # aBugTxt = vbNullString
    # IgnoreUnhandledError = False
    # E_AppErr.Permit = Empty                  ' EXcludes:  ErrSnoCatch, ZErrSnoCatch
    if ZErrSnoCatch Then                    ' ErrorCaught is NOT cleared:
    # GoTo FuncExit
    # GoTo ShowChange
        case 4                                       ' keep the Try, but forget all signs of an Err:
    # With E_Active
    # .ErrSnoCatch = False                 ' ErrSnoCatch, ZErrSnoCatch cleared
    # .errNumber = 0                       ' ErrorCaught is NOT cleared
    # .FoundBadErrorNr = 0
    # .Description = vbNullString
    # .Reasoning = vbNullString
    # .Explanations = vbNullString
    # ErrorCaught = 0                      ' Clearing this because error was accepted
    # T_DC.DCerrSource = .atKey
    # ErrDisplayModify = True
    # ZErrSnoCatch = False                ' what the hell is ZErrSnoCatch ???
    # GoTo ShowChange
    # End With                                 ' E_Active
    if ErrStatusFormUsable Then:
    # frmErrStatus.fErrNumber = 0
    if MayChangeErr Then:
    if ErrOnEntry <> 0 Then:
    # Err.Clear
    # ErrDisplayModify = True
    elif ErrOnEntry <> 0 Then:
    # GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
    # GoTo FuncExit
        case Else                                    ' Resets everything, including ErrorCaught:
    # aBugTxt = vbNullString
    # IgnoreUnhandledError = False
    # E_AppErr.Permit = Empty
    # MayChangeErr = True                      ' always clears Err
    # Err.Clear
    # E_Active.ErrSnoCatch = False
    # ErrorCaught = 0                          ' Clearing this, too
    # Call T_DC.N_ClearTermination
    # GoTo ShowChange

    # ShowChange:
    if ErrDisplayModify Or Not (ZErrSnoCatch Or SuppressStatusFormUpdate) Then:
    # Call ShowStatusUpdate

    # FuncExit:
    # aBugVer = True
    # ' NO: aBugTxt = vbNullString, wait for Catch
    # ZErrSnoCatch = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindFirstChar
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findfirstchar():
    # Dim cPos As Long
    if InStr(CharSet, Mid(x, cPos, 1)) > 0 Then:
    # FindFirstChar = cPos
    # GoTo FuncExit

    # FuncExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FindFirstDate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findfirstdate():
    # Const zKey As String = "O_Goodies.FindFirstDate"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction)

    # Const dateChars As String = "./"
    # Const digits As String = "0123456789"

    # Dim cPos As Long
    # cPos = FindFirstChar(x, digits)
    if cPos > 0 Then:
    # x = Mid(x, cPos)
    else:
    # FindFirstDate = False
    # GoTo FuncExit
    # ' suche erstes Zeichen, dass nicht im Datum vorkommen kann
    if InStr(digits & dateChars, Mid(x, cPos, 1)) = 0 Then:
    # Exit For
    if cPos < 11 Then:
    # x = Left(x, cPos - 1)
    if IsDate(x) Then:
    # fdate = CDate(x)
    # FindFirstDate = True
    else:
    # FindFirstDate = False
    else:
    # FindFirstDate = False

    # FuncExit:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function FindValueInArray
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def findvalueinarray():
    # Dim ExplainS As String
    # Const MyId As String = "FindValueInArray"

    # Dim j As Long

    if canPrint Then:
    # Call getInfo(aInfo, val)
    if aInfo.iAssignmentMode = 1 Then:
    # ExplainS = aInfo.iValue
    else:
    # ExplainS = "# value not printable"

    if j <= UBound(AR) Then:
    if val = AR(j) Then:
    # FindValueInArray = j
    if canPrint Then:
    # E_AppErr.atFuncResult = MyId & " found " & ExplainS & " at pos " & j
    # GoTo ProcRet
    # FindValueInArray = LBound(AR) - 1            ' val not found
    if canPrint Then:
    # E_AppErr.atFuncResult = MyId & " did not find " & ExplainS

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub FirstDiff
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def firstdiff():
    # Const zKey As String = "O_Goodies.FirstDiff"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")

    # Dim i As Long, j As Long, c1 As String, c2 As String, lmin As Long
    # Dim k As Long
    # Dim lmax As Long
    # Dim tlMax As Long
    # Dim mL As Long
    # Dim i1s As String, i2s As String
    # Dim Prefix As String

    # FirstDiff = vbNullString
    # p1 = vbNullString
    # p2 = vbNullString
    # i1s = Right("    " & WorkIndex(1), k)
    # i2s = Right("    " & WorkIndex(2), k)
    # lmax = Max(Len(S1), Len(S2))
    if targetlen = 0 Then:
    # targetlen = lmax                         ' unlimited length, never cut
    # lmin = Min(contextlen, targetlen)            ' must Show this much
    # tlMax = Max(contextlen, targetlen)           ' if we need 2 lines, use max possible
    # Prefix = "     "
    if S1 = S2 Then                              ' no first diff at all:
    if lmax > lmin Then:
    # p1 = Left(S1, lmin - Len(dots)) & dots
    # p2 = Left(S2, lmin - Len(dots)) & dots ' can not Show all of it
    else:
    # p1 = Left(S1, lmin)
    # p2 = Left(S2, lmin)
    # FirstDiff = WorkIndex(1) & "==" & WorkIndex(2) & ": " & Quote(p1)
    # GoTo FuncExit

    # mL = Len(S1) + Len(S2)
    if 2 * (mL + k + 2) < contextlen Then        ' fits one line:
    if Len(S1) > lmin And Right(S1, Len(dots)) <> dots Then:
    # p1 = Left(S1, lmin - Len(dots)) & dots
    else:
    # p1 = S1
    if Len(S2) > lmin And Right(S2, Len(dots)) <> dots Then:
    # p2 = Left(S2, lmin - Len(dots)) & dots
    else:
    # p2 = S2
    if mL < targetlen Then:
    # c1 = " / "
    else:
    # c1 = vbCrLf
    # FirstDiff = i1s & ": " & Quote(p1) & b & Quote(c1 & i2s) & ": " & Quote(p2)
    # GoTo FuncExit
    else:
    # p1 = vbNullString
    # p2 = vbNullString
    # ' find difference
    # i = 1
    # j = 1
    # Do While i <= Len(S1)
    # c1 = vbNullString
    # c2 = vbNullString
    if j > Len(S2) Then:
    # Exit Do
    # c1 = Mid(S1, i, 1)
    # c2 = Mid(S2, j, 1)
    if c1 <> c2 Then:
    if LenB(Ignore) > 0 And InStr(c1, Ignore) > 0 Then:
    # i = i + 1                        ' no relevance if ignore character
    elif Ignore <> vbNullString And InStr(c2, Ignore) > 0 Then:
    # j = j + 1
    else:
    # Exit Do                          ' relevant mismatch
    else:
    # i = i + 1
    # j = j + 1
    # Loop                                         ' all characters in S1 and S2 until mismatch

    if i + j > mL Then:
    # DoVerify False, " how can they be different but reach the end of the loop???"
    # FirstDiff = vbNullString
    # GoTo FuncExit

    if i < lmin - k Then                         ' we need not cut beginning if we fit in:
    # i = 1                                    ' start from beginning
    # FirstDiff = i1s & ": " & Left(S1, lmin) & vbCrLf
    # FirstDiff = FirstDiff & i2s & ": " & Left(S2, lmin)
    else:
    # FirstDiff = String(k, b) & "erster Unterschied an Position " & i & B2
    # k = Len(i1s) + 3
    if i < 6 Or i < lmin Then                ' dont strip front:
    if i < Len(FirstDiff) + k Then:
    # c1 = vbCrLf & String(i + k, b) & "V" ' V is down arrow
    else:
    # c1 = String(i - Len(FirstDiff) + k, b) & "V"
    # FirstDiff = FirstDiff & c1 & vbCrLf & i1s & ": " & Mid(S1, 1, lmin)
    # FirstDiff = FirstDiff & vbCrLf & i2s & ": " & Mid(S2, 1, lmin)
    else:
    if i < Len(FirstDiff) Then:
    # c1 = vbCrLf & String(i + k, b) & "V"
    else:
    # c1 = String(i - Len(FirstDiff) + k, b) & "V"
    # FirstDiff = FirstDiff & vbCrLf & i1s & ": " & dots & Mid(S1, i, lmin - Len(dots))
    # ' append second stripped value
    # FirstDiff = FirstDiff & vbCrLf & i2s & ": " & dots & Mid(S2, j, lmin - Len(dots))

    if i >= lmin Then                            ' cut front:
    # '       will we need to cut at all?
    # j = i + lmin - lmax
    if j > 0 Then                            ' true if both fit already:
    # i = lmax - j                         ' start relative to end of shorter
    # ' i will still be > 1, so
    # ' always Show dots at start
    # p1 = dots
    # p2 = dots
    # tlMax = tlMax - Len(dots)

    if i + tlMax > Len(S1) Or Right(S1, Len(dots)) = dots Then:
    # p1 = p1 & Mid(S1, i, tlMax)              ' can't Show more than total length, no double dots
    else:
    # p1 = p1 & Mid(S1, i, tlMax - Len(dots)) & dots
    if i + tlMax > Len(S2) Or Right(S2, Len(dots)) = dots Then:
    # p2 = p2 & Mid(S2, i, tlMax)              ' can't Show more than total length
    else:
    # p2 = p2 & Mid(S2, i, tlMax - Len(dots)) & dots

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function FormatRight
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def formatright():

    # ' simple replacement (for strings only!)     : Right(String(length, b) & value, length)
    # ' replacement (for int or long) use string of: Format(Value, String(length, "#") & "0")
    # ' replacement (for single or double  use two steps   ( Example is for length=8, frac=4 )
    # '    FormatRight = Format(Value, String(length - frac - 1, "#") & "0" & "." & String(frac, "0"))
    # '    FormatRight = Right(String(length, b) & FormatRight, length)
    # ' ex: tRT = Format(time, padNum3 & "." & String(4, "0"))
    # ' ex: tRT = Right(String(tRT, b) & tRT, 8)

    if VarType(Value) = vbString Then:
    # FormatRight = Right(String(length, b) & Value, length)
    # GoTo FuncExit
    elif Not IsNumeric(Value) Then:
    # DoVerify False
    # GoTo FuncExit                            ' returning vbNullString
    if VarType(Value) = vbDouble Or VarType(Value) = vbSingle Then:
    # FormatRight = Format(Value, String(length - frac - 1, "#") _
    # & "0" & "." & String(frac, "0"))
    else:
    # FormatRight = Format(Value, String(length, "#") & "0")
    # FormatRight = Right(String(length, b) & FormatRight, length)

    # FuncExit:


# ' for loop functions requireing an array
def genarr():
    if isArray(S) Then:
    # GenArr = S
    else:
    if LenB(splitter) = 0 Then:
    # GenArr = Array(S)
    else:
    # GenArr = split(S, splitter)

# '---------------------------------------------------------------------------------------
# ' Method : Function GetClipboardString
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getclipboardstring():

    # Const zKey As String = "O_Goodies.GetClipboardString"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction)

    # Dim GTO As Variant

    # ClipBoardIsEmpty = TriStateMixed             ' error in Extras/Verweise
    # Set DataObject = New MSForms.DataObject
    # Call Try(testAll)
    # DataObject.GetFromClipboard
    # GTO = DataObject.GetText(1)
    if GTO = False Then:
    if Catch Then:
    # ClipBoardIsEmpty = TristateTrue      ' = True
    # GetClipboardString = vbNullString
    else:
    # ClipBoardIsEmpty = TriStateFalse         ' = False
    # GetClipboardString = GTO

    # Call ErrReset(0)
    # Set DataObject = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)


# ' get  actDate    as the current date in normal format (always)
# ' or   DateId     as yyyy dd mm or yyyy dd mm_hh mm    (+1, 0)
# ' and  DateIdNB   as yyyyddmm or yyyyddmmhhmm          (+2, <=0)
def getdateid():
    # actDate = Now()
    match withTime:
        case 2, -1:
    # GetDateId = Format(actDate, "yyyymmdd_hhmm")
        case 1, 0:
    # GetDateId = Format(actDate, "yyyymmdd")
        case Else                                    ' e.g. -1 will set global variables, +1 or +2 will not:
    # GetDateId = Format(actDate, "yyyymmdd_hhmm")

    if withTime <= 0 Then                        ' set DateIdNB and DateId, else leave unchanged:
    # DateIdNB = GetDateId                     ' always no Blanks
    if Len(GetDateId) > 8 Then               ' with time and blanks:
    # DateId = Format(actDate, "yyyy mm dd hh mm")
    else:
    # DateId = Format(actDate, "yyyy mm dd")

# '---------------------------------------------------------------------------------------
# ' Method : Function GetEnvironmentVar
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getenvironmentvar():

    # Const zKey As String = "O_Goodies.GetEnvironmentVar"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction, _
    # ExplainS:="O_Goodies")

    # GetEnvironmentVar = String(255, 0)
    # Call GetEnvironmentVariable(Name, GetEnvironmentVar, Len(GetEnvironmentVar))
    # GetEnvironmentVar = TrimNul(GetEnvironmentVar)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub GetLine
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getline():


    # Dim j As Long
    # Dim aSheet As Worksheet

    # Set aSheet = O
    # Set aCell = aSheet.Cells(i, j)
    if DebugMode Then:
    # aCell.Select
    if VarType(aCell.Value) = vbString Then:
    if Left(aCell.Value, 1) = "'" Then:
    # varr(j) = Trim(Mid(aCell.Value, 2))
    else:
    # varr(j) = Trim(aCell)
    else:
    # varr(j) = aCell


# '---------------------------------------------------------------------------------------
# ' Method : Function GetPart
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getpart():
    # Dim Braces As Variant

    # sPos = InStr(Source, Fsep)
    if sPos > 0 Then:
    if Bsep = vbNullString Then:
    # Bsep = Fsep
    # Braces = split(Brace)
    # GetPart = Mid(Source, sPos + Len(Fsep))
    # Source = Left(Source, sPos - 1)
    # ePos = InStr(GetPart, Bsep)
    if ePos > 0 Then:
    # Source = Source & Mid(GetPart, ePos + Len(Bsep))
    # GetPart = Left(GetPart, ePos - 1)
    if UBound(Braces) >= 0 Then:
    # GetPart = Braces(0) & GetPart & Braces(UBound(Braces))

# ' find word surrounding pos in source, delimited by chars in split, giving word position
def getthisword():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
    # Const zKey As String = "O_Goodies.GetThisWord"
    # Dim zErr As cErr

    # Dim N As Long, j As Long, k As Long

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)

    # wPos = pos                                   ' word starts here if nothing is before it
    if wPos = 1 Then GoTo getTail:
    # ' find beginning of the word
    if Mid(Source, j, 1) = Mid(split, k, 1) Then:
    # ' we found a delimiter up front
    # wPos = j + 1
    # GoTo getTail

    # getTail:
    if Mid(Source, N, 1) = Mid(split, k, 1) Then:
    # ' we found a delimiter at back
    # GoTo gotTail

    # gotTail:
    # GetThisWord = Mid(Source, wPos, N - wPos)

    # FuncExit:
    # zErr.atFuncResult = CStr(GetThisWord)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetVtypeInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getvtypeinfo():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.GetVtypeInfo"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim vType As VbVarType

    # ' the sequence of the cases should be according to likelyhood to minimize branches
    match vTypeName:
        case "ItemProperty":
    # vType = olItemProperty
        case "String"                                ' 8:
    # vType = vbString
        case "Variant"                               ' 12:
    # vType = vbVariant
        case "Boolean"                               ' 11:
    # vType = vbBoolean
        case vbNullString                                      '  Name does not exist:
    # vTypeName = "Null"
    # vType = vbNull
        case "Null", "Nothing"                       '  Simulated, such a Name does not really exist:
    # vType = vbVariant
        case "Long"                                  ' 2:
    # vType = vbInteger
        case "Date"                                  ' 7:
    # vType = vbDate
        case "Object"                                ' 9:
    # vType = vbObject
        case "Long"                                  ' 3:
    # vType = vbLong
        case "Double"                                ' 5:
    # vType = vbDouble
        case "Single"                                ' 4:
    # vType = vbSingle
        case "Empty"                                 ' 0:
    # vType = vbEmpty
        case "Byte"                                  ' 17:
    # vType = vbByte
        case "LongLong"                              ' 20:
    # vType = 20&                              ' LongLong, only on 64 bit systems
        case "Null"                                  ' 1:
    # vType = vbNull
        case "Currency"                              ' 6:
    # vType = vbCurrency
        case "Error"                                 ' 10:
    # vType = vbError
        case "DataObject"                            ' 13:
    # vType = vbDataObject
        case "Decimal"                               ' 14:
    # vType = vbDecimal
        case _:
    # vType = 50                               ' this should mean we are using a (User-) Class Object
    # GetVtypeInfo = vType

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function GetWordContaining
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getwordcontaining():
    # Const zKey As String = "O_Goodies.GetWordContaining"
    # Static zErr As New cErr

    # Dim HookPos As Long
    # Dim aInstance As Long

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # HookPos = 1
    # HookPos = InStr(HookPos, Source, Hook)
    if HookPos = 0 Then:
    # notfound:
    # lWord = 0
    # wPos = 0
    # GoTo ProcReturn
    # aInstance = aInstance + 1
    if Instance > aInstance Then:
    # HookPos = InStr(HookPos + 1, Source, Hook)
    if HookPos = 0 Then:
    # GoTo notfound
    if aInstance < Instance Then:
    # GoTo nextinstance

    # GetWordContaining = GetThisWord(Source, HookPos, lWord, split, wPos)

    # ProcReturn:
    # Call ProcExit(zErr)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function Hex8
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def hex8():
    # Hex8 = Right("00000000" & Hex(aNum), 8)

# '---------------------------------------------------------------------------------------
# ' Method : Function HexN
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def hexn():
    # HexN = Right(String(N, "0") & Hex(O), N)
    if L > 0 Then:
    # HexN = Left(HexN, L)

# '---------------------------------------------------------------------------------------
# ' Method : Sub InspectType
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Detect Type Name / Type Id etc. into cInfo-Object
# '          if no previous AssignmentMode, get AssignmentMode 1 or 2
# '---------------------------------------------------------------------------------------
def inspecttype():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.InspectType"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim vArraySize As Long
    # Dim vIsArray As Boolean
    # Dim vAssignmentMode As Long
    # Dim vType As VbVarType
    # Dim vTypeId As String
    # Dim vScalarType As Long

    # Dim lb As Long
    # Dim ub As Long

    # With fInfo
    # vType = VarType(aVar)
    # vTypeId = TypeName(aVar)
    # vIsArray = False
    # vArraySize = -99
    # vScalarType = -99
    # vAssignmentMode = -99

    if vType >= vbArray Then:
    # vIsArray = True
    # vAssignmentMode = 0
    # vTypeId = TypeName(aVar)
    # vTypeId = "Array of " & vTypeId
    # .iType = vType - vbArray
    if .iType = vbVariant Then:
    # ABounds:
    # ub = UBound(aVar)
    # lb = LBound(aVar)
    # vArraySize = ub - lb + 1
    if vArraySize >= 0 Then:
    # vTypeId = Left(vTypeId, Len(vTypeId) - 1) _
    # & lb & " to " & ub & ")"
    # vAssignmentMode = 1
    elif .iType = vbByte Then:
    # vAssignmentMode = 2
    else:
    # DoVerify False, "array of WHAT???"
    # vArraySize = UBound(aVar.Value) - LBound(aVar.Value) + 1
    else:
    if vTypeId = "ItemProperty" Then:
    # .iClass = olItemProperty         ' here, always work on value
    if InStr(dftRule.clsNotDecodable.aRuleString & b, aVar.Name & b) = 0 Then:
    # DoVerify T_DC.DCAllowedMatch = testAll, "Expected to have ** from Caller"
    # vScalarType = IsScalar(TypeName(aVar.Value))
    if ErrorCaught <> 0 Then:
    # .DecodedStringValue = "# No value for variable Type " & aVar.Name & vbCrLf & "# " & Err.Description
    # Call ErrReset(4)
    else:
    # vScalarType = IsScalar(vTypeId) - .iDepth ' -.iDepth to indicate (User-) (not Item-) Property
    if vScalarType > 0 Then              ' is a scalar:
    # vAssignmentMode = 1
    elif vScalarType < 0 Then:
    # vAssignmentMode = 0              ' not decodable or decode not wanted
    else:
    # vAssignmentMode = 2              ' some sort of object

    # .iArraySize = vArraySize
    # .iIsArray = vIsArray
    # .iType = vType
    # .iTypeName = vTypeId
    # .iScalarType = vScalarType
    if .iAssignmentMode <= 0 Then            ' gate deciding if we want change:
    # .iAssignmentMode = vAssignmentMode
    else:
    # DoVerify .iAssignmentMode = vAssignmentMode, _
    # "Attention: .iAssignmentMode will not be changed, is=" _
    # & .iAssignmentMode & " rejecting new=" & vAssignmentMode
    # End With                                     ' fInfo

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function IntAssignIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def intassignifchanged():


    if Source <> Target Then                     ' only valid for Integer not WITHIN Objects:
    if ChangeAssignReverse Then:
    # IntAssignIfChanged = Target
    else:
    # IntAssignIfChanged = Source
    # modified = True


# '---------------------------------------------------------------------------------------
# ' Method : Function IsBetween
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isbetween():
    if x >= L And x <= W2 Then:
    # IsBetween = True

# '---------------------------------------------------------------------------------------
# ' Method : Function IsOneOf
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isoneof():
    # Dim x As Variant
    # Dim i As Long
    # i = 0
    if VarType(AR) >= vbArray Then:
    for x in ar:
    # i = i + 1
    if x = Testvar Then:
    # IsOneOf = i
    # GoTo ProcRet
    else:
    # DoVerify False, "only designed for arrays ar=array(a,b,c...)"

    # ProcRet:

# ' find a term in s matching one of array vSo
# ' if ignore specified, term is delimited
# ' if op >= len(s) lookup is in reverse
# ' Note: default comptype is vbTextCompare (not vbbinarycompare)
def isoneofpos():
    # Const zKey As String = "O_Goodies.IsOneOfPos"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim v As Variant
    # Dim splitignore As Variant
    # Dim vS As Variant
    # Dim FI As Variant
    # Dim P As Long
    # Dim k As Long
    # Dim reverseIndex As Boolean
    # Dim ls As String

    # ' because instr / instrrev appear to ignore the compare method, we do it in Lcase
    if compType = vbTextCompare Then:
    # ls = LCase(S)
    else:
    # ls = S
    if Op >= Len(ls) Then                        ' find in reverse:
    # reverseIndex = True
    # k = 1                                    ' not used here
    else:
    # k = Op                                   ' original start for search
    if Not IsMissing(Ignore) Then:
    if isArray(Ignore) Then:
    # splitignore = Ignore
    if isArray(vSo) Then:
    # vS = vSo
    else:
    if VarType(vSo) = vbString Then:
    if LenB(vSo) = 0 Then:
    # DoVerify False, " parm invalid"
    # vS = split(vSo, sep)
    else:
    # DoVerify False, " not implemented"
    for v in vs:
    if compType = vbTextCompare Then:
    # v = LCase(v)
    if reverseIndex Then                     ' find in reverse:
    # P = InStrRev(ls & sep, v & sep)      ' comptype not working in Instr/Rev
    else:
    if P = 0 Then:
    # P = k
    # P = InStr(P, ls & sep, v & sep)
    if P = 0 And Not IsMissing(Ignore) Then:
    if Not isArray(splitignore) Then:
    # splitignore = split(Ignore, b)
    for fi in splitignore:
    # P = k
    if reverseIndex Then             ' find in reverse:
    # P = InStrRev(ls, v)
    else:
    # P = InStr(P, ls, v)
    if P > 0 Then:
    if StrComp(FI, Mid(P, ls + Len(v), Len(FI)), compType) Then:
    # Exit For
    if P > 0 Then:
    if P > 1 Then:
    if LenB(sep) > 0 Then            ' sonst kein test auf Wort:
    if Mid(ls, P - 1, 1) <> sep Then:
    # P = P + 1
    # k = P
    if P <= Len(ls) Then:
    # GoTo nextOne
    # GoTo loopex
    # k = Op
    # loopex:
    # IsOneOfPos = P
    # Op = Len(v) + Len(sep)                       ' end of recognized word

    # FuncExit:
    # zErr.atFuncResult = CStr(IsOneOfPos)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : IsScalar
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Find if aTypeName is Scalar and optionally convert to aType
# '---------------------------------------------------------------------------------------
def isscalar():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.IsScalar"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim i As Long
    # Dim aStr As String
    # Dim aChar As String

    if DebugLogging Then:
    if aTypeName = "RTFBody" Then:
    # aStr = vbNullString
    print(Debug.Print "------------ RTFBody ------------")
    # aChar = Chr(aTD.adItemProp.Value(i))
    if aChar = vbCr Then:
    # aStr = aStr & vbCrLf
    elif aChar <> vbLf Then:
    # aStr = aStr & aChar
    print(Debug.Print aStr)
    print(Debug.Print "-------- End RTFBody ------------")
    if isNotDecodable(aTypeName) Then:
    # IsScalar = -1                            ' no decode possible or wanted
    else:
    # i = InStr(ScalarTypeNames, aTypeName & b)
    if i > 0 Then:
    # IsScalar = True
    # IsScalar = dSType.Item(i)            ' always >0
    else:
    # IsScalar = 0                         ' non Scalar Vartype

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function IsSimilar
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def issimilar():
    # Dim la As Long
    # Dim lb As Long
    # Dim le As Long

    # la = Len(A)
    # lb = Len(b)
    if la < lb Then:
    # le = la
    else:
    # le = lb
    if le = 0 Then:
    if la = lb Then:
    # IsSimilar = True
    else:
    # IsSimilar = False
    else:
    if ignoreCase Then:
    # IsSimilar = Left(LCase(A), le) = Left(LCase(b), le)
    else:
    # IsSimilar = Left(A, le) = Left(b, le)

# '---------------------------------------------------------------------------------------
# ' Method : Function IsUcase
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def isucase():
    # IsUcase = (b <> LCase(b))

# '---------------------------------------------------------------------------------------
# ' Method : Function LastTrail
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def lasttrail():
    # arr = split(S, at)
    # LastTrail = arr(UBound(arr))

# '---------------------------------------------------------------------------------------
# ' Method : LimitAppended
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Abschneiden des Anfangs wenn maxlnge berschritten
# '---------------------------------------------------------------------------------------
def limitappended():
    # Dim i As Long, ldiff As Long
    # LimitAppended = Sta & aWhat
    # ldiff = Len(LimitAppended) - maxl - Len(px)
    if ldiff >= 0 Then                           ' kein Platz mehr im String: ersten Trenner finden:
    # i = InStr(Sta, ", ")
    # LimitAppended = px & Mid(LimitAppended, i + 2)

# '---------------------------------------------------------------------------------------
# ' Method : Function LongAssignIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def longassignifchanged():


    if Source <> Target Then                     ' only valid for Long not WITHIN Objects:
    if ChangeAssignReverse Then:
    # LongAssignIfChanged = Target
    else:
    # LongAssignIfChanged = Source
    # modified = True


# '---------------------------------------------------------------------------------------
# ' Method : Function LPad
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def lpad():
    # LPad = Right(String(L, FrontChar) & CLng(N), L)

# '---------------------------------------------------------------------------------------
# ' Method : Function LRString
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Construct a string with left and right justified strings
# '---------------------------------------------------------------------------------------
def lrstring():
    # Dim P As Long
    # Dim px As String
    # Dim pa As Long
    # Dim alPart As String

    # alPart = lPart
    # pa = Abs(Indent)
    # P = TLen - Len(rPart) - pa - Len(alPart) - 1
    if P < 1 Then:
    if rCutL > 0 Then:
    # px = " ." & Mid(rPart, rCutL - P - 1)
    else:
    if InStr(rPart, ".") > 0 Then:
    # px = " ." & Tail(rPart, ".")
    else:
    # px = rPart
    # P = P + 2                        ' did not use " ." prefix gives 2 more chars
    # alPart = Mid(alPart, 4 - P)
    # P = TLen - Len(px) - Indent - Len(alPart) - 1
    if P < 1 Then:
    # P = 1
    else:
    # px = rPart
    # LRString = String(pa, b) & alPart & String(P, b) & px


# '---------------------------------------------------------------------------------------
# ' Method : LString
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Left pad to fixed length
# '          needed only because there is no Function, only Statement LSet
# '---------------------------------------------------------------------------------------
def lstring():
    # LString = String(L, b)
    # LSet LString = x

# '---------------------------------------------------------------------------------------
# ' Method : Function Max
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def max():

    if A > b Then:
    # Max = A
    else:
    # Max = b

# '---------------------------------------------------------------------------------------
# ' Method : Function Min
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def min():
    if A < b Then:
    # Min = A
    else:
    # Min = b

# '---------------------------------------------------------------------------------------
# ' Method : Function NextNumberInString
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def nextnumberinstring():
    # Dim ltr As String
    # ltr = Mid(str, start, 1)
    if ltr >= "0" And ltr <= "9" Then:
    # Exit For

# '---------------------------------------------------------------------------------------
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def nonmodalmsgbox():
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim MsgFrm As frmMessage
    # Dim sTime As Variant
    # Dim sBlockEvents As Boolean
    # Dim ReShowFrmErrStatus As Boolean
    # Dim MsgHdl As cFindWindowParms
    # Dim TotalTime As Single
    # Dim saveErrEx As Boolean
    # Dim WaitCycles As Long

    # sBlockEvents = E_Active.EventBlock
    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
    # ReShowFrmErrStatus = True
    # saveErrEx = ErrExActive
    # Call ErrEx.Disable
    if ErrStatusFormUsable Then:
    # frmErrStatus.fUseErrExOn = vbNullString
    # Set MsgFrm = New frmMessage
    if LenB(Title) > 0 Then:
    # MsgFrm.Caption = Title
    # MsgFrm.Message = Prompt
    # MsgFrm.B1.Caption = B1Label
    if LenB(B2Label) = 0 Then:
    # MsgFrm.B2.Visible = False
    else:
    # MsgFrm.B2.Caption = B2Label
    # MsgFrm.Show vbModeless
    try:
        # E_Active.EventBlock = False
        if ErrStatusFormUsable Then:
        # frmErrStatus.fNoEvents = E_Active.EventBlock
        # Call BugEval
        # doMyEvents                               ' allow interaction, delay and wait
        # Call WindowSetForeground(MsgFrm.Caption, MsgHdl)
        if Wait(0.2, trueStart:=sTime, TotalTime:=TotalTime, Retries:=WaitCycles) _:
        # Or WaitCycles > 300 Then                 ' true in debug mode
        # Exit Do
        # MsgFrm.ResponseWaits = MsgFrm.ResponseWaits + 1
        # Loop

        # MsgFrm.Hide
        # & MsgFrm.ResponseWaits & " cycles", eLall)
        # Set MsgFrm = Nothing
        # GoTo FuncExit
        # BadExit:

        # FuncExit:
        # Set MsgHdl = Nothing
        if ReShowFrmErrStatus Then:
        # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
        if saveErrEx And LenB(UseErrExOn) > 0 Then:
        # Call ErrEx.Enable(UseErrExOn)
        # E_Active.EventBlock = sBlockEvents
        if ErrStatusFormUsable Then:
        # frmErrStatus.fUseErrExOn = UseErrExOn
        # Call BugEval
        # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function NormalizeTelefonNumber
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def normalizetelefonnumber():
    # Const zKey As String = "O_Goodies.NormalizeTelefonNumber"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim ThisNumber As String, countryCode As String, areaCode As String
    # Dim localNumber As String, extensionNumber As String
    # Dim i As Long
    # Dim startswith3stars As Boolean
    # Dim Result As String

    # Result = original
    if Left(Result, 3) = "***" Then:
    # DoVerify False, " wo kommt den so was vor?"
    # startswith3stars = True
    # Result = Mid(Result, 4)                  ' could be CarTelephoneNumber
    else:
    # i = InStr(Result, "*")
    if i > 0 Then:
    # NormalizeTelefonNumber = Mid(Result, i) ' numbers like *123# have no area code
    # Result = vbNullString                          ' no attempt to change something else,
    # GoTo FuncExit
    # ThisNumber = Trim(Replace(Replace(Result, ")", b), "/", b))
    # areaCode = vbNullString
    # countryCode = vbNullString
    # localNumber = vbNullString
    if LenB(Result) = 0 Then:
    # GoTo FuncExit

    # ' numbers without country codes or with nondefault codings thereof
    if Left(ThisNumber, 2) = "(+" Or Left(ThisNumber, 2) = "00" Then:
    # ThisNumber = "+" & Trim(Mid(ThisNumber, 3))

    if Left(ThisNumber, 1) = "(" Then:
    # ThisNumber = Trim(Mid(ThisNumber, 2))    ' drop this char and check some countries in brackets
    # ThisNumber = CheckForAreaCodes(ThisNumber, areaCode)
    if Left(ThisNumber, 1) = "0" Then:
    # GoTo localCoding
    # ' this may be unsafe !?
    # ' Debug.Assert False                                ' manual inspection of thisnumber is better


    if CountStringOccurrences(ThisNumber, "+") > 1 Then:
    # ThisNumber = "+" & Replace(LastTrail(ThisNumber, "+"), b, vbNullString)

    if Left(ThisNumber, 1) = "+" Then            ' standard coding:
    # localNumber = Trim(Tail(ThisNumber, b, countryCode))
    # normalCoding:
    if LenB(countryCode) = 0 Then            ' kein Erfolg mit b als Trennzeichen:
    # localNumber = Trim(Tail(localNumber, "-", countryCode))
    if LenB(countryCode) = 0 Then:
    # localNumber = CheckForCountryCodes(localNumber, countryCode)
    else:
    if countryCode = "+1" Then:
    # localNumber = Replace(localNumber, "(", vbNullString)
    # areaCode = Left(localNumber, 3)
    # localNumber = Trim(Mid(localNumber, 4))
    # localNumber = Replace(localNumber, b, vbNullString, 1, 1)
    else:
    # areaCode = CheckForCountryCodes(countryCode, countryCode)
    if areaCode = countryCode Then:
    # areaCode = vbNullString                ' unregistered county is not an area code
    if Left(localNumber, 1) = "(" Then:
    # localNumber = Trim(Mid(localNumber, 2))
    if Left(localNumber, 1) = "0" Then:
    # localNumber = Trim(Mid(localNumber, 2))
    if LenB(areaCode) = 0 Then:
    # localNumber = CheckForAreaCodes(localNumber, areaCode)
    if LenB(areaCode) = 0 And Len(countryCode) > 4 Then ' could be inside of country code:
    # areaCode = CheckForCountryCodes(countryCode, countryCode)
    if LenB(areaCode) = 0 Then:
    # localNumber = Trim(Tail(localNumber, b, areaCode))
    if LenB(areaCode) = 0 Then:
    # localNumber = Trim(Tail(localNumber, "-", areaCode))
    if Len(areaCode) > 4 Then:
    # localNumber = CheckForAreaCodes(localNumber, areaCode) _
    # & "-" & localNumber

    if LenB(areaCode) = 0 Then               ' no sep from Area to local: test some Matches:
    # localNumber = CheckForAreaCodes(localNumber, areaCode)
    if LenB(areaCode) = 0 Then               ' could be inside of country code:
    # areaCode = CheckForCountryCodes(countryCode, countryCode)
    if LenB(areaCode) = 0 Then:
    if Left(localNumber, 1) = "1" Or Left(localNumber, 1) = "9" Then:
    if Left(localNumber, 2) = "18" Then:
    # areaCode = Left(localNumber, 4)
    # localNumber = Mid(localNumber, 5)
    else:
    # areaCode = Left(localNumber, 3)
    # localNumber = Mid(localNumber, 4)
    # i = InStr(areaCode, "-")
    if i > 1 Then:
    # localNumber = Mid(areaCode, i + 1) & b & localNumber
    # areaCode = Left(areaCode, i - 1)
    # extensionNumber = Trim(Tail(localNumber, "-", localNumber))
    elif Left(ThisNumber, 1) = "(" Or Left(ThisNumber, 1) = "0" Then:
    # ThisNumber = Mid(ThisNumber, 2)
    # localCoding:
    # localNumber = ThisNumber
    # countryCode = "+49"                      ' default country
    # GoTo normalCoding
    # localNumber = Trim(Tail(areaCode, "-", extensionNumber))
    else:
    if LenB(areaCode) = 0 Then:
    # areaCode = "6501"                    ' default area
    # countryCode = "+49"
    # localNumber = Trim(Tail(ThisNumber, "-", extensionNumber))

    if LenB(localNumber) = 0 Then:
    # localNumber = extensionNumber
    # extensionNumber = vbNullString

    if LenB(countryCode) = 0 Then:
    if Left(areaCode, 1) = "+" Then:
    # countryCode = Left(areaCode, 3)
    # areaCode = Mid(areaCode, 4)
    elif countryCode <> "+1" _:
    # And Len(countryCode) < 3 _
    # And Len(areaCode) > 2 Then
    # countryCode = countryCode & Left(areaCode, 2)
    # areaCode = Mid(areaCode, 3)
    # CondensedPhoneNumber = Replace(countryCode, "+", "00") _
    # & areaCode & localNumber & extensionNumber
    if LenB(countryCode) > 0 Then countryCode = countryCode & b:
    if LenB(areaCode) > 0 Then areaCode = "(" & areaCode & ") ":
    if LenB(extensionNumber) > 0 Then extensionNumber = "-" & extensionNumber:

    # NormalizeTelefonNumber = countryCode & areaCode & localNumber & extensionNumber
    if startswith3stars Then:
    # NormalizeTelefonNumber = "***" & NormalizeTelefonNumber

    # FuncExit:
    # zErr.atFuncResult = CStr(NormalizeTelefonNumber)
    if NormalizeTelefonNumber <> original _:
    # And Reassign Then
    # original = NormalizeTelefonNumber

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function ObjAssignIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def objassignifchanged():


    if Source <> Target Then                     ' only valid for Object not WITHIN Objects:
    if ChangeAssignReverse Then:
    # Set ObjAssignIfChanged = Target
    else:
    # Set ObjAssignIfChanged = Source
    # modified = True

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function Pad
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Pad (or cut) to toLen with PadChar
# '---------------------------------------------------------------------------------------
def pad():
    if Len(PadChar) = 0 Then:
    # PadChar = b
    elif Len(PadChar) > 1 Then:
    # PadChar = Left(PadChar, 1)
    if toLen < Len(str) Then:
    # Pad = LString(str, toLen)
    else:
    # Pad = str & String(toLen - Len(str), PadChar)

# '---------------------------------------------------------------------------------------
# ' Method : PartialMatch
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Matches stringlist built as [*]xxxx[*] versus sItem Partial by length of xxxx
# '---------------------------------------------------------------------------------------
def partialmatch():
    # Const zKey As String = "O_Goodies.PartialMatch"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim f As Long
    # Dim EI As Long
    # Dim aToken As Variant
    # Dim sTrunc As String

    # ' maybe use Vartype? IsMissing???
    # ' if ErrArray is specified as array, just use it. If not: split ListToken
    for atoken in errarray:
    if LenB(aToken) = 0 Then:
    # GoTo SkipIt
    if StrComp(sItem, aToken, compType) = 0 Then ' Match complete identity:
    # PartialMatch = True
    # GoTo FuncExit
    # sTrunc = sItem
    # ' remove and remember [*]xxxx[*]
    # EI = InStrRev(aToken, "*")               ' end of interesting part xxxx
    # f = InStr(aToken, "*")
    if EI + f = 0 Then:
    # GoTo SkipIt
    elif f = EI Then:
    if EI > 1 Then                       ' "*" only allowed at start or end of aToken:
    # f = 0                            ' it only is at the end, none at start
    else:
    if EI = 2 Then:
    # DoVerify False, " invalid match pattern (too short)"
    if f > 1 Then:
    # DoVerify False, " " * " not at start x*xx* invalid match pattern"

    if EI > 1 Then                           ' [*]xxxx*...:
    # aToken = Left(aToken, EI - 1)        ' xxxx or *xxxx
    # sTrunc = Mid(sTrunc, f + 1, EI - 1)  ' restrict compare len
    if StrComp(sTrunc, aToken, compType) = 0 Then ' Match left part identity:
    # PartialMatch = True              ' xxxx
    # GoTo FuncExit

    if f > 0 Then                            ' *xxxx           start of interesting part:
    # aToken = Mid(aToken, 2)              ' now is xxxx
    # sTrunc = Left(sTrunc, Len(aToken))
    if StrComp(sTrunc, aToken, compType) = 0 Then ' Match right part identity:
    # PartialMatch = True              ' xxxx
    # GoTo FuncExit
    # SkipIt:

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function PickDateFromString
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def pickdatefromstring():
    # Const zKey As String = "O_Goodies.PickDateFromString"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim ePos As Long
    # Dim aPos As Long
    # Dim tPos As Long
    # Dim MDT As Date
    # Dim tpart As String
    # Dim ePart As String

    # ePos = 1                                     ' eliminate weekday
    # aPos = IsOneOfPos(aDate, split("tag day Mittwoch", b), vbNullString, ePos, " ,")
    if aPos > 0 Then:
    # aDate = TrimRemove(Mid(aDate, aPos + ePos), " ,") ' trim blank or ,

    # aDate = Translate(aDate, " um / at / ab ", b, "/") ' irrelevante worte
    # ePos = 1                                     ' seperate date from time
    # aPos = IsOneOfPos(aDate, split("am on den dem im", b), b, ePos)
    if aPos > 0 Then:
    # tpart = Trim(Mid(aDate, aPos, ePos + 10))
    # tPos = IsOneOfPos(aDate, split("um at ab", b), b, ePos)
    if tPos > 0 Then                         ' redundant time specification:
    # tpart = Trunc(1, Mid(aDate, tPos + ePos), b)
    if tPos - 1 > aPos Then:
    # tpart = Mid(aDate, aPos, tPos - aPos - 1)
    # DoVerify False, " why this ???"
    # aDate = Replace(Mid(aDate, tPos + ePos), ".", ":")
    # aDate = tpart & b & Trunc(1, aDate, b)
    if IsDate(aDate) Then:
    # tpart = aDate
    else:
    # ePos = 1
    # aPos = IsOneOfPos(aDate, split("Uhr o'clock"), vbNullString, ePos, " ,")
    if aPos > 0 Then:
    # tpart = Trim(Left(aDate, aPos - 1))
    else:
    # tpart = Trim(Replace(aDate, ", ", b))

    # ePos = InStrRev(tpart, b)                    ' date + time: 12.2. 18:00, seperate
    if ePos > 4 Then                             ' both date and time present in tpart:
    # ePart = Mid(tpart, ePos + 1)
    if InStr(ePart, ".") > 0 Then            ' correct times 20.00 -> 20:00:
    # tpart = Trim(Left(tpart, ePos - 1))  ' take apart
    for apos in range(0, 6):
    # ePart = Replace(ePart, "." & aPos, ":" & aPos)
    # tpart = tpart & b & ePart            ' put back together
    else:
    # ePart = vbNullString                               ' probably no time specified

    if Len(tpart) < 3 Then                       ' no minutes specified, add :00:
    # tpart = tpart & ":00"
    if IsDate(tpart) Then:
    # MDT = CDate(tpart)
    # NoDate = False
    else:
    # daterr:
    if Not NoDate Then                       ' locally report error:
    # NoDate = True
    # PickDateFromString = MDT

    # FuncExit:
    # Call ProcExit(zErr)

    # zErr.atFuncResult = CStr(PickDateFromString)


# '---------------------------------------------------------------------------------------
# ' Method : Function Quote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def quote():
    # Dim la As String
    # Dim le As String
    if LenB(lr) = 0 Then:
    # la = Q
    # le = Q
    else:
    # la = Left(lr, 1)
    # le = Right(lr, 1)
    # Quote = la & S & le                          ' " or ( ) at beginning and end

# '---------------------------------------------------------------------------------------
# ' Method : Function Quote1
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def quote1():
    # Quote1 = QuoteWithDoubleQ(S, "'")            ' quote with ' at beginning and end

# '---------------------------------------------------------------------------------------
# ' Method : Function QuoteWithDoubleQ
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def quotewithdoubleq():
    # Dim DiQ As String, Ss As String

    # Ss = CStr(S)
    # DiQ = DoubleInternalQuotes(Ss, DQ)           ' double quotation marks inside
    # QuoteWithDoubleQ = DQ & DiQ & StrReverse(DQ) ' quote with DQ at beginning and end

# '---------------------------------------------------------------------------------------
# ' Method : RandSequence
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Create array of random integers without duplicates
# '---------------------------------------------------------------------------------------
def randsequence():

    # Const zKey As String = "O_Goodies.RandSequence"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim i As Long
    # Dim RS() As Long
    # Dim iTemp As Long
    # Dim iZ As Long

    # ReDim RS(1 To anz) As Long
    # RS(i) = i
    # Randomize Timer
    # iZ = Int((i * Rnd) + 1)
    # iTemp = RS(iZ)
    # RS(iZ) = RS(i)
    # RS(i) = iTemp
    # RandSequence = RS

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function RCut
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def rcut():
    # RCut = Left(S, Len(S) - i)

# '---------------------------------------------------------------------------------------
# ' Method : Function RecognizeCode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def recognizecode():
    # Const zKey As String = "O_Goodies.RecognizeCode"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

    # Dim L As Long
    # Dim h As String
    # Dim N As String
    # Dim P As Long
    # L = Len(pattern)
    # N = Number
    if LenB(Code) = 0 Then:
    # P = InStr(Number, pattern)
    if (Left(Number, 1) = "0" And P = 2) Or P = 1 Then:
    # Code = pattern
    # Number = Trim(Mid(Number, L + P))
    # RecognizeCode = True
    else:
    if Left(Code, L) = pattern Then:
    # h = Code
    # Code = pattern
    # Number = Trim(Mid(h, L + 1))
    # RecognizeCode = True

    # FuncExit:
    # 'zErr.atFuncResult = CStr(RecognizeCode)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function reduceDouble
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def reducedouble():
    # Dim j As Long
    # Dim FI As String
    # Dim c2 As String
    # Dim i As Long

    # c2 = C & C
    # FI = Left(N, 2)
    # i = InStr(j, N, c2)
    if i > 0 Then:
    # j = i
    else:
    # Exit For
    if FI = C And FI = Mid(N, j, Len(C)) Then:
    # N = Mid(N, 1, j - 1) & Mid(N, j + Len(C))
    # FI = Mid(N, j, Len(C))
    # reduceDouble = N

# '---------------------------------------------------------------------------------------
# ' Method : Function ReFormat
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def reformat():
    # Dim i As Long
    # Dim M As String
    # Dim N As String
    # N = S
    # M = Mid(replaceChars, i, 1)
    # N = Replace(N, M, replaceWith)
    # N = TrimTail(N, replaceWith)
    # N = TrimFront(N, replaceWith)
    # N = reduceDouble(N, replaceWith)
    # ReFormat = reduceDouble(N, nonRepeating)

# '---------------------------------------------------------------------------------------
# ' Method : Function Remove
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Remove / Replace string using (a list of) characters to remove. Optional replacement by ReplaceWith
# '---------------------------------------------------------------------------------------
# ' if sepToRemove specified, and it is a required seperator (= leading and trailing)
# '       then RemoveThis is replaced.
# '       If ReplaceWith="", replacing by sepToRemove else by ReplaceWith
# ' if it is not a required seperator (= not leading), always ReplaceWith is used.
def remove():
    # Dim A As Long
    # Dim ls As Long
    # Dim remList As Variant
    # Static oTaS As String
    # Static Recursion As Long                     ' special gated proc

    if oTaS <> TargetASource Then:
    # oTaS = TargetASource
    # Recursion = 0

    # Recursion = Recursion + 1
    if Recursion = 1 Then:
    # StringMod = False

    # Remove = TargetASource
    if IsMissing(compType) Then:
    # compType = -1

    if IsMissing(sepToRemove) Then               ' Remove all occurrences ignoring sepToRemove as delimiter:
    if LenB(Remove) > 0 And LenB(RemoveThis) > 0 Then ' not required seperator::
    # Remove = Replace(TargetASource, RemoveThis, replaceWith, 1, -1, compType)
    # GoTo FuncExit
    # ' it is required seperator:
    if LenB(Remove) = 0 Or LenB(RemoveThis) = 0 Or LenB(sepToRemove) = 0 Then:
    # GoTo FuncExit                            ' nothing can be removed

    if replaceWith = vbNullString Then:
    # replaceWith = sepToRemove

    if InStr(RemoveThis, sepToRemove) > 0 Then   ' list of stuff to remove:
    # remList = split(RemoveThis, sepToRemove)
    # Remove = Replace(Remove, remList(A) & sepToRemove, replaceWith, 1, -1, compType)
    # GoTo FuncExit

    # A = InStr(1, TargetASource, RemoveThis, compType)
    if A = 0 Then                                ' it is not contained at all:
    # GoTo FuncExit

    # ls = Len(sepToRemove)
    if A > ls Then:
    if Mid(TargetASource, A - ls, ls) <> sepToRemove Then:
    # GoTo FuncExit                        ' does not contain sepToRemove in front of RemoveThis: there's a wrong seperator
    else:
    if A > 1 Then:
    if A <= ls Then:
    # GoTo FuncExit                    ' does not contain sepToRemove only partial match of sepToRemove,
    else:
    # GoTo FuncExit
    # Remove = Replace(sepToRemove & Remove & sepToRemove, _
    # sepToRemove & RemoveThis & sepToRemove, _
    # replaceWith, 1, -1, compType)

    # ' Start/End could have several sepToRemoves, ( Trimming )
    # While Right(Remove, ls) = sepToRemove
    # Remove = Left(Remove, Len(Remove) - ls)
    # Wend

    # While Left(Remove, ls) = sepToRemove
    # Remove = Right(Remove, Len(Remove) - ls)
    # Wend

    # FuncExit:
    if Remove <> TargetASource Then:
    # StringMod = True
    # Recursion = Recursion - 1
    if Recursion = 0 Then:
    # oTaS = vbNullString

# '---------------------------------------------------------------------------------------
# ' Method : Function RemoveChars
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def removechars():
    # Dim i As Long

    # RemoveChars = orig
    # RemoveChars = Replace(RemoveChars, Mid(NoNoChars, i, 1), vbNullString)

# '---------------------------------------------------------------------------------------
# ' Method : Function RemoveDoubleBlanks
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def removedoubleblanks():
    # RemoveDoubleBlanks = S
    # While InStr(RemoveDoubleBlanks, B2) > 0
    # RemoveDoubleBlanks = Replace(RemoveDoubleBlanks, B2, b)
    # Wend

# '---------------------------------------------------------------------------------------
# ' Method : Function RemoveWord
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def removeword():
    # Dim i As Long, lsplit As String, actpattern As String
    # Dim N As Long, j As Long, Word As String
    # Dim skiPropTailWild As Boolean

    # RemoveWord = Source

    if LenB(split) = 0 Then:
    # lsplit = b
    else:
    # lsplit = split

    # actpattern = pattern
    if Left(actpattern, 1) = "*" Then:
    # actpattern = Mid(actpattern, 2)
    # skiPropTailWild = True
    if Right(actpattern, 1) = "*" Then:
    # actpattern = Left(actpattern, Len(actpattern) - 1)
    # skiPropTailWild = True

    # i = InStr(N, RemoveWord, actpattern)
    if i = 0 Then:
    # GoTo ProcRet                         ' no pattern Match
    else:
    # Word = GetThisWord(RemoveWord, i, Len(actpattern), lsplit, j)
    if skiPropTailWild Or j = i Then:
    if skiPropTailWild Or i + Len(actpattern) = j + Len(Word) Then:
    # RemoveWord = Left(RemoveWord, j - 1) & Mid(RemoveWord, j + Len(Word) + 1)
    # StringsRemoved = StringsRemoved & "entfernt: " & Word & " at pos. " & i & vbCrLf
    # N = i + Len(actpattern)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Repeat
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Repeat a string n times with sep between (but not for last)
# '---------------------------------------------------------------------------------------
def repeat():
    # Dim i As Long

    # Repeat = Repeat & str & sep
    # Repeat = Repeat & str

# '---------------------------------------------------------------------------------------
# ' Method : Function ReplaceAll
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def replaceall():
    # Dim ach As String

    # ReplaceAll = S
    # While ach <> ReplaceAll
    # ach = ReplaceAll
    if compType = -1 Then:
    # ReplaceAll = Replace(ReplaceAll, f, DS, 1, -1)
    else:
    # ReplaceAll = Replace(ReplaceAll, f, DS, 1, -1, compType)
    # Wend

# '---------------------------------------------------------------------------------------
# ' Method : RString
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Righ path to fixed length
# '          needed only because there is no Function, only Statement RSet
# '---------------------------------------------------------------------------------------
def rstring():
    # RString = String(L, b)
    # RSet RString = x

# '---------------------------------------------------------------------------------------
# ' Method : RTail
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Schneide String hinter Sat ab, Suche von hinten. Optional Front Teil vor Sat
# '---------------------------------------------------------------------------------------
def rtail():
    # Dim Bpos As Long

    # Bpos = InStrRev(S, Sat, -1, compType)
    if Bpos > 0 Then:
    # RTail = Mid(S, Bpos + Len(Sat))
    # Front = Left(S, Bpos - 1)
    else:
    # RTail = S
    # Front = vbNullString

# '---------------------------------------------------------------------------------------
# ' Method : Function RTrimC
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def rtrimc():
    # Dim k As Long
    # Dim lt As Long
    # lt = Len(DQ)
    # k = Len(S) - lt + 1
    # While k > 0 And Mid(S, k, lt) = DQ
    # k = k - lt
    # Wend
    # RTrimC = Mid(S, 1, k)

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetGlobal
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setglobal():

    # Const zKey As String = "O_Goodies.SetGlobal"
    # Const MyId As String = "SetGlobal"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="O_Goodies")

    # Dim SetX As String
    # Dim ShellRet As Long

    # SetX = "SETX " & VarName & b & Quote(VarValue) ' do NOT use Quote here ???
    if LogAllErrors Then:
    # ShellRet = Shell("CMD /K echo Ausfhren von " & SetX & " & " & SetX, vbNormalFocus)
    else:
    # ShellRet = Shell("CMD /C " & SetX, vbMinimizedFocus)

    # FuncExit:
    if VarName = "Test" Then                     ' check for selftest:
    # TestTail = b & Trim(Mid(Testvar, (InStr(Testvar, "|") + 1)))
    # aDebugState = InStr(TestTail, b & MyId) > 0

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowAsc
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showasc():
    # Dim i As Long

    print(Debug.Print i, Asc(Mid(S, i, 1)))

# '---------------------------------------------------------------------------------------
# ' Method : Function SimpleAsciiText
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def simpleasciitext():
    # Dim work As String
    # Dim C As String
    # Dim i As Long

    # firstcontrolcharPos = 0                      ' none found so far
    # work = aText
    # C = Mid(work, i, 1)
    if Asc(C) < 32 Then                      ' replace all control chars by blanks:
    if firstcontrolcharPos = 0 Then:
    # firstcontrolcharPos = i
    # Mid(work, i, 1) = b
    if stopimmediately Then              ' stop after replacing first control char:
    # Exit For
    # SimpleAsciiText = work

# '---------------------------------------------------------------------------------------
# ' Method : Function SplitAtWordsWhenCaseChanges
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def splitatwordswhencasechanges():
    # Dim i As Long
    # Dim L As Long
    # Dim caseIsUp As Boolean
    # Dim ncaseIsUp As Boolean
    # Dim WordStarted As Boolean                   ' Cap followed by LowerCase is not a new word
    # Dim work As String
    # Dim A As String                              ' char at pos i
    # Dim N As String                              ' char after  i

    # modified = False
    # L = Len(S)
    if L < 2 Then:
    # SplitAtWordsWhenCaseChanges = S
    # GoTo ProcRet

    # A = Left(S, 1)
    # caseIsUp = IsUcase(A)
    # WordStarted = caseIsUp
    # N = Mid(S, i + 1, 1)
    # N = Mid(S, i + 1, 1)
    # ncaseIsUp = IsUcase(N)
    if N = b Or A = b Then                   ' already seperate word:
    # WordStarted = True
    if ncaseIsUp = caseIsUp Or WordStarted Or WordStarted <> caseIsUp Then:
    # work = work & A
    # WordStarted = False
    else:
    # work = work & A & b
    # WordStarted = ncaseIsUp
    # modified = True
    # caseIsUp = ncaseIsUp
    # A = N
    # SplitAtWordsWhenCaseChanges = work & A

    # ProcRet:

# ' split string into non-blank words as variant array of strings
def splitnbwords():
    # Dim j As Long, A As Variant, S As Variant
    # j = InStr(aString, B2)                       ' find double blanks
    # While j > 0
    # aString = Replace(aString, B2, b)
    # j = InStr(aString, B2)                   ' find double blanks
    # Wend
    if LenB(Trim(aString) = 0) Then:
    # SplitNBWords = vbNullString
    # GoTo FuncExit
    # S = split(aString, b)
    # j = 0
    for a in s:
    # A = Trim(A)
    if LenB(A) > 0 Then:
    # S(j) = A
    # j = j + 1
    # ReDim Preserve S(j - 1)
    # SplitNBWords = S

    # FuncExit:


# '---------------------------------------------------------------------------------------
# ' Method : Function StrAssignIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def strassignifchanged():
    if Source <> Target Then                     ' only valid for String not WITHIN Objects:
    if ChangeAssignReverse Then:
    # StrAssignIfChanged = Target
    else:
    # StrAssignIfChanged = Source
    # modified = True


# '---------------------------------------------------------------------------------------
# ' Method : Sub StrBetween
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Find the string between two delimiting strings
# '---------------------------------------------------------------------------------------
def strbetween():
    # Dim A As Long
    # Dim P As Long

    # inS = vbNullString
    # A = InStr(Source, startS)
    if A > 0 Then:
    # P = InStr(A + Len(startS), Source, endS)
    if P > 0 Then:
    if KeepDelims Then:
    # StartPos = A
    # inS = Mid(Source, StartPos, P + Len(endS) - 1)
    else:
    # StartPos = A + Len(startS)
    # inS = Mid(Source, StartPos, P - Len(endS) - 1)


# '---------------------------------------------------------------------------------------
# ' Method : Sub StringDiff
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stringdiff():

    # Dim i As Long, lmax As Long, lmin As Long
    # Dim uline As String
    # Dim allsame As Boolean

    # allsame = True
    # S1 = Trim(S1)
    # S2 = Trim(S2)
    # lmin = Min(Len(S1), Len(S2))
    # lmax = Max(Len(S1), Len(S2))
    if lmax = 0 Then:
    print(Debug.Print "both S1 and S2 are null strings")
    else:
    if lmin = 0 Then:
    if Len(S1) = 0 Then:
    print(Debug.Print "S1 is null string")
    print(Debug.Print "S2=" & Quote(S2))
    elif Len(S2) = 0 Then:
    print(Debug.Print "S2 is null string")
    print(Debug.Print "S1=" & Quote(S1))
    print(Debug.Print "S2 is null string")
    else:
    if i > lmax Then:
    # allsame = False
    # uline = uline & "-"
    else:
    if Mid(S1, i, 1) = Mid(S2, i, 1) Then:
    # uline = uline & b
    else:
    # allsame = False
    # uline = uline & "|"
    if allsame Then:
    print(Debug.Print "===" & Quote(S1))
    else:
    print(Debug.Print "^^^ " & uline & String(lmax - lmin, "-"))
    print(Debug.Print "S1=" & Quote(S1))
    print(Debug.Print "S2=" & Quote(S2))

    # ProcRet:

# ' Insert a text at specified position and return the position at the end of this insertion
def stringinsert():
    # Dim L As Long
    # Dim paddedInsert As String
    # L = Len(inToString)
    # paddedInsert = addSep & InsertText & addSep
    if atpos >= L Then:
    # Call AppendTo(inToString, InsertText, addSep)
    else:
    # inToString = Left(inToString, atpos) & paddedInsert & Mid(inToString, atpos + 1)
    # atpos = InStr(inToString, paddedInsert)

    # StringInsert = atpos + Len(paddedInsert)     ' end of inserted text at this position


# '---------------------------------------------------------------------------------------
# ' Method : Function StringRemove
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stringremove():
    # Dim D As Variant
    # Dim S As Variant
    # Const MyId As String = "StringRemove"

    # StringRemove = Trim(inputString)
    if LenB(StringRemove) = 0 Then:
    # GoTo FuncExit
    if Right(StringRemove, 1) <> Left(sep, 1) Then:
    if Len(sep) = 2 And Right(sep, 1) = b Then:
    # StringRemove = StringRemove & Left(sep, 1)
    # D = split(droplist, Left(sep, 1))
    for s in d:
    # S = Trim(S)
    if LenB(S) > 0 Then:
    # StringRemove = Replace(StringRemove, S & sep, vbNullString)
    if Len(sep) = 2 And Right(sep, 1) = b Then:
    # StringRemove = Replace(StringRemove, S & Left(sep, 1), b)
    # StringRemove = Trim(StringRemove)

    # FuncExit:
    if (DebugMode Or aDebugState) And ShowFunctionValues Then:
    # Call ShowFunctionValue(MyId, CStr(StringRemove), False)

# '---------------------------------------------------------------------------------------
# ' Method : Sub StrReplBetween
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Replace a string between two delimiting strings
# '---------------------------------------------------------------------------------------
def strreplbetween():

    # Call StrBetween(Source, startS, endS, RemovePart, KeepDelims:=False)
    if LenB(RemovePart) > 0 Then:
    if withStartS Then:
    # RemovePart = startS & RemovePart
    if withEndS Then:
    # RemovePart = RemovePart & endS
    # res = Replace(Source, RemovePart, withWhatS, Count:=1)
    else:
    # res = Source


# '---------------------------------------------------------------------------------------
# ' Method : Sub Swap
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def swap():
    # Dim C As Variant

    if asObject Then:
    # GoTo TreatAsObject

    if VarType(A) = vbObject Then:
    # TreatAsObject:
    # Set C = A
    # Set A = b
    # Set b = C
    else:
    # C = A
    # A = b
    # b = C
    # Set C = Nothing

# '---------------------------------------------------------------------------------------
# ' Method : Tail
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Schneide String hinter Sat ab, Suche von vorn. Optional Front Teil vor Sat
# '---------------------------------------------------------------------------------------
def tail():
    # Dim i As Long, lFront As String

    # lFront = S
    # i = InStr(S, Sat)
    if i > 0 Then:
    # Tail = Mid(S, i + Len(Sat))
    # Front = Left(S, i - 1)
    else:
    # Tail = S
    # Front = vbNullString                               ' do not duplicate

    if DebugMode Then:
    if ShowFunctionValues Then:
    if LenB(S) > 0 Then:
    print(Debug.Print "Tail='" & Tail & "'", "Front='" & Front & "'", "Sat='" & Sat & "'")

    # FuncExit:
    if (DebugMode Or aDebugState) And ShowFunctionValues Then:
    print(Debug.Print "Tail='" & CStr(Tail) & "'", "Front='" & Front & "'", "Sat='" & Sat & "'")

# '---------------------------------------------------------------------------------------
# ' Method : Function TextEdit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def textedit():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
    # Const zKey As String = "O_Goodies.TextEdit"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction) ' or drop for ) ' Z_Type

    # Call frmLongText.TextEdit(Text, ReplCrLf)

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : TimerNow
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get the present time using time and exact timer
# '---------------------------------------------------------------------------------------
def timernow():
    # Dim nw As Date
    # Dim TT As Single
    # Dim nwl As Single

    # Const spd As Long = 86400

    # TT = Timer
    # nw = Time
    # nwl = CSng(nw) * spd

    # TimerNow = nw & "," & Mid(TT - nwl, 3)


# '---------------------------------------------------------------------------------------
# ' Method : Function Trail
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trail():
    # Dim Bpos As Long

    # Bpos = InStrRev(S, at)
    if Bpos > 0 Then:
    # Trail = Mid(S, Bpos + Len(at))
    else:
    # Trail = S

# ' what / tothis can be single string or array of strings
# ' if single string, Array is generated using splitter
# ' splitter="" means what is no array, but atothis is
# ' splitter="$$1$$x" selects x only for what
# ' splitter="$$2$$x" selects x only for atothis
def translate():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
    # Const zKey As String = "O_Goodies.Translate"
    # Dim zErr As cErr

    # Dim aWhat As Variant
    # Dim aToThis As Variant
    # Dim i As Long
    # Dim uT As Long
    # Dim ms1 As String
    # Dim ms2 As String
    # Dim A As String
    # Dim DS As String

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)

    # ms1 = Replace(splitter, "$$1$$", vbNullString)
    # ms2 = RTail(splitter, "$$2$$")
    # ms1 = Trunc(1, ms1, "$$")
    # aWhat = GenArr(what, ms1)
    # aToThis = GenArr(tothis, ms2)
    # uT = UBound(aToThis)
    # Translate = S
    # A = aWhat(i)
    if i > uT Then                           ' fewer replacements than replacers:
    # DS = aToThis(uT)
    else:
    # DS = aToThis(i)
    # Translate = ReplaceAll(Translate, A, DS, compType)

    # FuncExit:
    # zErr.atFuncResult = CStr(Translate)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function TrimFront
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trimfront():
    # Dim j As Long

    # TrimFront = N
    # j = InStrRev(N, C)
    if j > Len(C) Then                           ' if c used like a string quote, remove that:
    if Right(N, Len(C)) = Left(N, Len(C)) Then ' front == tail:
    # TrimFront = Mid(N, j + Len(C), Len(N) - 2 * Len(C)) ' remove c from front and tail completely

    # j = InStr(TrimFront, C)                      ' cut off the part in front of c
    # While j > 0                                  '
    # TrimFront = Mid(TrimFront, j + Len(C))
    # j = InStr(TrimFront, C)                  ' repeat if it contains several c
    # Wend

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function TrimNul
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trimnul():
    # '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS

    # Dim iPos As Long                             ' Nul markiert String-Ende

    # iPos = InStr(Item, vbNullChar)
    # TrimNul = IIf(iPos > 0, Left$(Item, iPos - 1), Item)


# '---------------------------------------------------------------------------------------
# ' Method : Function TrimRemove
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trimremove():
    # Dim thisTrimArray As Variant
    # Dim trimalso As Boolean
    # Dim L As Variant

    if isArray(Trimlist) Then:
    # thisTrimArray = Trimlist
    else:
    # thisTrimArray = split(Trimlist, b)
    # TrimRemove = N
    for l in thistrimarray:
    if LenB(L) = 0 Then:
    # trimalso = True
    else:
    # While Left(TrimRemove, Len(L)) = L
    # TrimRemove = Mid(TrimRemove, Len(L) + 1)
    # Wend
    # While Right(TrimRemove, Len(L)) = L
    # TrimRemove = Mid(TrimRemove, 1, Len(TrimRemove) - Len(L))
    # Wend
    if trimalso Then:
    # TrimRemove = Trim(TrimRemove)

# ' remove all chars following last c
def trimtail():

    # Dim j As Long
    # j = InStrRev(N, C)
    if j > 0 Then:
    if j > 1 Then:
    # TrimTail = Left(N, j - 1)
    else:
    # TrimTail = vbNullString
    else:
    # TrimTail = N

    # FuncExit:
    if (DebugMode Or aDebugState) And ShowFunctionValues Then:
    # Call ShowFunctionValue("TrimTail", TrimTail, False)

# '---------------------------------------------------------------------------------------
# ' Method : Function Trunc
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def trunc():
    if StartPos <= 0 Then:
    # StartPos = 1
    # Bpos = InStr(StartPos, inS, bS, compType)
    if Bpos > 0 Then:
    # Trunc = Mid(inS, StartPos, Bpos - StartPos)
    # Tail = Mid(inS, Bpos + Len(bS))
    else:
    # Trunc = Mid(inS, StartPos)
    # Tail = vbNullString

# '---------------------------------------------------------------------------------------
# ' Method : Function UnQuote
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def unquote():
    # UnQuote = Mid(S, 2, Len(S) - 1)

# '---------------------------------------------------------------------------------------
# ' Method : Function VarAssignIfChanged
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def varassignifchanged():


    if Source <> Target Then                     ' only valid for Variant not WITHIN Objects:
    if ChangeAssignReverse Then:
    # VarAssignIfChanged = Target
    else:
    # VarAssignIfChanged = Source
    # modified = True


# '---------------------------------------------------------------------------------------
# ' Method : Verify
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: delivers first position of valid char (Valid defined by letters in OK)
# '          if negate:=True,delivers first position of invalid char
# '---------------------------------------------------------------------------------------
def verify():
    # Dim i As Long
    # Dim j As Long

    # j = InStr(OK, Mid(S, i, 1))
    if j = 0 Then                            ' not in ok:
    if Not negate Then:
    # Exit For
    else:
    if negate Then:
    # Exit For
    if i <= Len(S) Then:
    # Verify = i                               ' else = 0

# '---------------------------------------------------------------------------------------
# ' Method : Function Wait
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def wait():

    # Const zKey As String = "O_Goodies.Wait"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim starttime As Variant
    # Dim LastInDate As Date
    # Dim PrevEntrySec As Double
    # Dim LastInSec As Double
    # Static thisLong As Variant

    # LastInDate = E_Active.atLastInDate        ' remember state before wait
    # PrevEntrySec = E_Active.atPrevEntrySec
    # LastInSec = E_Active.atLastInSec

    if IsMissing(trueStart) Or isEmpty(trueStart) Then:
    # trueStart = 0
    if LenB(Title) = 0 Then:
    # Title = "Waited for dialog to complete for "

    # starttime = Timer
    if trueStart = 0 Or Retries = 0 Then:
    # trueStart = starttime
    # thisLong = 0
    # Retries = Retries + 1
    # Call Sleep(sec)                              ' wait here with modal box open
    # Call ShowStatusUpdate
    # thisLong = thisLong + Timer - trueStart
    # TotalTime = thisLong
    if DebugOutput And DebugMode And (Retries) Mod 5 = 0 Then:
    print(Debug.Print Title & thisLong & " sec")
    if DebugMode Or TotalTime > 60# Then:
    # ' Don'W.xlTSheet wait,  because sleep causes problems in debugmode
    print(Debug.Print "Please continue debug mode, can't wait for timer in debug mode" & vbCrLf _)
    # & "==> assuming wait has completed, so either left button was pressed or Continue/Step (F5/F8)"
    if Not aNonModalForm Is Nothing Then:
    # ErrStatusFormUsable = True
    # Call BugEval
    # Wait = True                              ' End of wait
    # ' Debug.Assert False on caller's side recommended

    # FuncExit:

    # E_Active.atLastInDate = LastInDate        ' restore state before Wait
    # E_Active.atPrevEntrySec = PrevEntrySec
    # E_Active.atLastInSec = LastInSec
    # Recursive = False

    # zExit:
    # Call DoExit(zKey, "Waited " & TotalTime)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Z_GetApplication
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Use Named Application or open if not running yet. Returns Used or Opened
# '---------------------------------------------------------------------------------------
def z_getapplication():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "O_Goodies.Z_GetApplication"
    # Call DoCall(zKey, "Sub", eQzMode)

    # OpenedHere = False
    # startOver:
    if Not ShutUpMode Then:
    print(Debug.Print LString(String(OffCal, b) & zKey _)
    # & " Started get Application " & AppName & ", Time: ", OffTim) & Format(Timer, " #####.00") _


    # Call ErrEx.Disable                           ' Must disable, it would set Exception Options
    if ErrStatusFormUsable Then:
    # frmErrStatus.fErrAppl = vbNullString
    # Call BugEval

    # aBugTxt = "Get Excel Object"                 ' Extras\Options\General\Stop on not handled Errors
    # Call Try(testAll)
    # Set Z_GetApplication = GetObject(, AppName & ".Application")
    if Z_GetApplication Is Nothing Then          ' returns error 429 if application not running:
    if CatchNC Then:
    print(Debug.Print LString(String(OffCal, b) & AppName _)
    # & " is not found running at ", OffTim) _
    # & Format(Timer, " #####.00")
    if Err.Number <> 429 Then:
    print(Debug.Print "Error " _)
    # & Err.Number & vbCrLf & Err.Description
    # Call ErrReset(4)
    else:
    if Not ShutUpMode Then:
    print(Debug.Print LString(String(OffCal, b) & zKey _)
    # & " found Application " & AppName _
    # & " already running, Time: ", OffTim) _
    # & Format(Timer, " #####.00")
    # Z_GetApplication = False
    # GoTo FuncExit

    # aBugTxt = "CreateObject(AppName" & Quote(".Application", Bracket)
    # Call Try
    # Set Z_GetApplication = CreateObject(AppName & ".Application")
    if Catch Then:
    print(Debug.Print LString(String(OffCal, b) & AppName _)
    # & " Application not available, Time: ", OffTim) _
    # & Format(Timer, " #####.00") _
    # & vbCrLf & " Error " & Err.Number _
    # & vbCrLf & Err.Description
    else:
    print(Debug.Print LString(String(OffCal, b) & AppName _)
    # & " Application successfully launched, Time: ", OffTim) _
    # & Format(Timer, " #####.00")
    # OpenedHere = True

    # FuncExit:
    if LenB(UseErrExOn) > 0 Then:
    # Call ErrEx.Enable(UseErrExOn)
    if ErrStatusFormUsable Then:
    # frmErrStatus.fErrAppl = AppName
    # Call BugEval

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function SetOffline
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setoffline():
    # Const zKey As String = "O_Goodies.SetOffline"
    # Static zErr As New cErr

    if aRDOSession Is Nothing Then:
    # withProlog = True
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")

    # SetOffline = SetOnline(OnlineMode, withProlog, RepeatLimit)

    # ProcReturn:
    if withProlog Then:
    # Call ProcExit(zErr)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetOnline
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Change Online Status,  ToggleMode is only method working
# '---------------------------------------------------------------------------------------
def setonline():
    # Const zKey As String = "O_Goodies.SetOnline"
    # Static zErr As New cErr

    # Dim oriStatus As OlExchangeConnectionMode
    # Dim oriMode As String
    # Dim actStatus As OlExchangeConnectionMode
    # Dim actMode As String
    # Dim reqMode As String

    # Dim Retries As Long
    # Dim ModeWanted As String
    # Dim ModeNow As String

    # Dim debugthis As Boolean
    # debugthis = DebugMode And LogZProcs

    if aRDOSession Is Nothing Then:
    # withProlog = True
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")

    if Not aRDOSession.LoggedOn Then:
    # Call aRDOSession.Logon(aRDOSession.Profiles(1))

    # Call getOnlineStatus(oriStatus, oriMode, ModeNow)
    # actMode = oriMode
    # actStatus = oriStatus

    # reqMode = ExStatusToText(RequestStatus, ModeWanted)

    if ErrStatusFormUsable Then:
    if debugthis Then:
    if (actOnlineStatus <> ModeNow And ModeWanted <> ModeNow) Then:
    # DoVerify InStr(frmErrStatus.fOnline.Caption, ModeNow) > 0, _
    # "*** ErrStatus displays online status=" _
    # & frmErrStatus.fOnline.Caption & ", actually=" & ModeNow _
    # & ", wanting=" & ModeWanted
    # frmErrStatus.fOnline.Caption = ModeNow
    # actOnlineStatus = ModeNow               ' globally correct(ed) state

    # Do While InStr(ModeWanted, ModeNow) = 0 And Retries < RepeatLimit
    # Retries = Retries + 1
    # ' toggle mode because desired state not reached
    # Call olApp.ActiveExplorer.CommandBars.ExecuteMso("ToggleOnline")
    # DoEvents
    # Call Sleep(200)

    # Call getOnlineStatus(actStatus, actMode, ModeNow)
    if debugthis And InStr(ModeWanted, ModeNow) = 0 Then:
    print(Debug.Print "shit, Wanting " & ModeWanted & "<>" & ModeNow, Retries)
    # Debug.Assert True
    else:
    # Exit Do
    # Loop
    # DoVerify Retries < RepeatLimit, "** Unable to reach connection mode " & ModeWanted

    # SetOnline = oriMode <> actMode              ' indicate changed mode if not bad, else false
    if debugthis Then:
    if actStatus = RequestStatus Then       ' no change necessary:
    # ModeWanted = "exact " & ModeNow
    elif ModeWanted = ModeNow Then:
    # ModeWanted = "Equivalent " & ModeWanted
    # Call LogEvent("original Connection State=" & oriStatus _
    # & "(" & oriMode _
    # & ")" & vbCrLf & " current State=" & actStatus _
    # & "(" & actMode & ")" & vbCrLf & " Wanted: " _
    # & ModeWanted & "(" & RequestStatus & ") " _
    # & ", tried " & Retries & " times")

    # actOnlineStatus = ModeNow
    if ErrStatusFormUsable Then:
    if frmErrStatus.fOnline.Caption <> ModeNow Then:
    # frmErrStatus.fOnline.Caption = ModeNow

    # ProcReturn:
    if withProlog Then:
    # Call ProcExit(zErr, CStr(SetOnline))

    # ProcRet:

def getonlinestatus():
    # isStatus = aRDOSession.ExchangeConnectionMode
    # isMode = ExStatusToText(isStatus, isNow)

# '---------------------------------------------------------------------------------------
# ' Method : ExStatusToText
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Convert
# '---------------------------------------------------------------------------------------
def exstatustotext():
    # ' ----  Trivial Proc ----

    if aStatus >= olCachedConnectedHeaders Then:
    # aState = "Online"
    else:
    # aState = "Offline"
    # ExStatusToText = ExModeNames(aStatus / 100)



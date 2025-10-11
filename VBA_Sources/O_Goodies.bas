Attribute VB_Name = "O_Goodies"
Option Explicit

Public Const GH = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Public DataObject As MSForms.DataObject
Public ClipBoardIsEmpty As cTriState

Type PictureDescription
    height As Single
    width As Single
End Type

Dim arr                                          ' Variant for splitting texts

Public ChangeAssignReverse As Boolean            ' request functions ...AssignIfChanged to reverse source and target
Public ModThisTo As Variant                      ' new value in AssignIfChanged
Public DecodedValue As Variant                   ' set by Function if the value can be determined as String
Public AssignmentMode As Long                    ' HasValue Assignment mode:
' 0=Imposs., 1 = Set, 2 = direct Scalar,
' 3 = Object Default (Scalar) or result of 4 below
' 4 = ItemProperty value evaluation

'---------------------------------------------------------------------------------------
' Method : Sub AddItemToList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AddItemToList(i1 As String, ID As String, comp As String, i2 As String)


Dim LI As cListItm

    Set LI = New cListItm
    LI.Index1 = i1
    LI.MainId = ID
    LI.Compares = comp
    LI.Index2 = i2
    If ListContent Is Nothing Then
        Set ListContent = New Collection
        ListCount = 0
    Else
        ListCount = ListContent.Count
    End If
    ListContent.Add LI
    ListCount = ListContent.Count

zExit:

End Sub                                          ' O_Goodies.AddItemToList

'---------------------------------------------------------------------------------------
' Method : Sub AddNumbItemToCollection
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AddNumbItemToCollection(ByRef myColl As Collection, val As Object, Name As String, Optional modifier As String)

Const zKey As String = "O_Goodies.AddNumbItemToCollection"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)

Dim ItemProperty As cNumbItem
    Set ItemProperty = New cNumbItem
    ItemProperty.NuIndex = myColl.Count + 1
    ItemProperty.Key = Name
    Set ItemProperty.ValueOfItem = val
    
    aBugTxt = "Add item to collection " & modifier & Name
    Call Try
    myColl.Add ItemProperty, Key:=modifier & Name
    Catch

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' O_Goodies.AddNumbItemToCollection

'---------------------------------------------------------------------------------------
' Method : Function Append
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Append(ByVal S1 As String, ByVal S2 As String, Optional ByVal sep As String, Optional always As Boolean, Optional ToFront As Boolean) As String


Dim L As Long
Dim listx As Variant
Dim aListx As Variant

Static myRecursionDepth As Long

    If myRecursionDepth = 0 Then
        StringMod = False
    End If
    myRecursionDepth = myRecursionDepth + 1
    
    ' append to front will not split at sep, because that would result in wrong ordering
    If LenB(sep) > 0 And InStr(S2, sep) > 0 And Not ToFront Then
        listx = split(S2, sep)
        For Each aListx In listx
            If Not isEmpty(aListx) And LenB(aListx) > 0 Then
                myRecursionDepth = myRecursionDepth + 1 ' next level of recursion
                S1 = Append(S1, (aListx), sep, always, ToFront)
                myRecursionDepth = myRecursionDepth - 1
                DoVerify myRecursionDepth > 0
            End If
        Next aListx
        Append = S1
        GoTo FuncExit
    End If
    
    If Not always Then                           ' test if append creates a double entry
        L = InStr(1, sep & S1 & sep, sep & S2 & sep, vbTextCompare)
        If L > 0 Then                            ' string kommt vor
            ' schon enthalten: return original ohne Zufügen
            Append = S1
            GoTo FuncExit
        End If
    End If
    If LenB(S1) = 0 Then
        Append = S2
    Else
        If ToFront Then
            Append = S2 & sep & S1
        Else
            Append = S1 & sep & S2
        End If
    End If
    
FuncExit:
    If Append <> S1 Then
        StringMod = True
    End If
    myRecursionDepth = myRecursionDepth - 1

End Function                                     ' O_Goodies.Append

'---------------------------------------------------------------------------------------
' Method : AppendTo
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: If string(s) to be inserted are new to first string, or always is specified:
'               add it to front or back of first string.
'          calls Append: String compare is via vbTextCompare (ignore case of chars)
'---------------------------------------------------------------------------------------
Sub AppendTo(S1 As String, ByVal S2 As String, Optional ByVal sep As String, Optional always As Boolean, Optional ToFront As Boolean)


Dim Result As String

    StringMod = False
    Result = Append(S1, S2, sep, always, ToFront)
    If StringMod Then
        S1 = Result
    End If

End Sub                                          ' O_Goodies.AppendTo

'---------------------------------------------------------------------------------------
' Method : Function ArrayMatch
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ArrayMatch(testArray As Variant, matchString As String, Optional CaseSensitive As VbCompareMethod = vbTextCompare, Optional WildCard As String = "*") As Long


' check if the entries in testarray(String) occur anywhere in matchString
' no WildCard match: string must occur at the end
Dim i As Long
Dim k As Long
Dim testword As String

    For k = LBound(testArray) To UBound(testArray)
        testword = Replace(testArray(k), WildCard, vbNullString) ' remove WildCard (at end)
        If LenB(Trim(testword)) > 0 Then
            i = InStr(1, matchString, testword, CaseSensitive)
            If i > 0 Then                        ' it is somewhere
                If Len(testword) = Len(testArray(k)) Then ' no WildCard in testarray(k)
                    If i + Len(testword) = Len(matchString) Then ' at the end
                        ArrayMatch = k           ' match condition OK
                        GoTo FuncExit
                    End If
                Else
                    ArrayMatch = k               ' match condition OK, not at end
                    GoTo FuncExit
                End If
            End If
        End If
    Next k
    ' if we drop out, no match
    ArrayMatch = LBound(testArray) - 1           ' not in array

FuncExit:

End Function                                     ' O_Goodies.ArrayMatch

'---------------------------------------------------------------------------------------
' Method : Sub ArrayOrderMaxMin
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ArrayOrderMaxMin(AR As Variant, Optional aPosMax As Variant, Optional aPosMin As Variant)

Const zKey As String = "O_Goodies.ArrayOrderMaxMin"
    Call DoCall(zKey, tSub, eQzMode)

Dim j As Long
Dim MaxOfIndex As Long
Dim MinOfIndex As Long

    'On Error GoTo 0
    
    ' assume valid defaults
    MaxOfIndex = LBound(AR)
    MinOfIndex = MaxOfIndex
    ' find a bigger one
    For j = LBound(AR) To UBound(AR)
        If AR(MaxOfIndex) < AR(j) Then
            MaxOfIndex = Max(j, MaxOfIndex)
        End If
        If AR(MinOfIndex) > AR(j) Then
            MinOfIndex = Min(j, MinOfIndex)
        End If
    Next j
    If Not IsMissing(aPosMax) Then
        aPosMax = MaxOfIndex
    End If
    If Not IsMissing(aPosMin) Then
        aPosMin = MinOfIndex
    End If

FuncExit:

zExit:
    Call DoExit(zKey)

End Sub                                          ' O_Goodies.ArrayOrderMaxMin

'---------------------------------------------------------------------------------------
' Methods: (type)AssignIfChanged
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Simple and more general Assignment with Check if any change done
' NOTE   : ByRef for Targets does not work for values WITHIN Objects/Classes/Types
'          In that case, use this model:
'        If AssignIfChanged(\(.*?), (.*?)\) Then
'            If ChangeAssignReverse Then
'                $2 = ModThisTo
'            Else
'                $1 = ModThisTo
'            End If
'        end if
'---------------------------------------------------------------------------------------
Function AssignIfChanged(ByVal Target As Variant, ByVal Source As Variant) As Boolean


    If CStr(Source) <> CStr(Target) Then
        If ChangeAssignReverse Then
            ModThisTo = Target
        Else
            ModThisTo = Source
        End If
        AssignIfChanged = True
    End If

End Function                                     ' O_Goodies.BoolAssignIfChanged

'---------------------------------------------------------------------------------------
' Method : CatStrings
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Variable number of Strings are concatenated with sep between (but not for last)
'---------------------------------------------------------------------------------------
Function CatStrings(sep As String, ParamArray Cats()) As String
Dim i As Long
Dim npart As String

    If LenB(sep) = 0 Then
        DoVerify False                           ' likely it won, "W.xlTSheet work in split later"
    End If
    
    For i = LBound(Cats) To UBound(Cats) - 1
        npart = Cats(i)
        If InStr(npart, sep) > 0 Then
            DoVerify False                       ' likely it won, "W.xlTSheet work in split later"
        End If
        CatStrings = CatStrings & npart & sep
    Next i
    CatStrings = CatStrings & Cats(UBound(Cats))
End Function                                     ' O_Goodies.CatStrings

'---------------------------------------------------------------------------------------
' Method : Function CharsToHexString
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CharsToHexString(S As String) As String
Dim i As Long

    For i = 1 To Len(S)
        CharsToHexString = CharsToHexString & Right("0" & Hex(Asc(Mid(S, i, 1))), 2)
    Next i

End Function                                     ' O_Goodies.CharsToHexString

'---------------------------------------------------------------------------------------
' Method : Function CheckForAreaCodes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CheckForAreaCodes(aString As String, aCode As String) As String
Const zKey As String = "O_Goodies.CheckForAreaCodes"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim ThisNumber As String
Dim minlen As Long
Dim addI As Long
    ThisNumber = Trim(Replace(Replace(aString, b, vbNullString), "-", vbNullString))
    If RecognizeCode(ThisNumber, aCode, "711") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "650") Then
        minlen = 4
    ElseIf RecognizeCode(ThisNumber, aCode, "800") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "651") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "461") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "65") Then
        minlen = 4
    ElseIf RecognizeCode(ThisNumber, aCode, "261") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "23") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "15") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "16") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "17") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "18") Then
        minlen = 3
    ElseIf RecognizeCode(ThisNumber, aCode, "40") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "30") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "89") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "228") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "6131") Then
        minlen = 4
    Else
        ' vorwahl und nummer sind nicht getrennt: keine Normalisierung möglich
        If Len(ThisNumber) > 4 Then
            aCode = Left(ThisNumber, 4)
            ThisNumber = Mid(ThisNumber, 5)
        End If
    End If
    addI = minlen - Len(aCode)
    If addI > 0 Then
        aCode = aCode & Left(ThisNumber, addI)
        ThisNumber = Mid(ThisNumber, addI + 1)
    End If
    CheckForAreaCodes = ThisNumber

FuncExit:
    zErr.atFuncResult = CStr(CheckForAreaCodes)

ProcReturn:
    Call ProcExit(zErr)
  
End Function                                     ' O_Goodies.CheckForAreaCodes

'---------------------------------------------------------------------------------------
' Method : Function CheckForCountryCodes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CheckForCountryCodes(aString As String, aCode As String) As String
Dim ThisNumber As String
Dim plusBug As String

    ThisNumber = Tail(Trim(Replace(aString, "-", b)), _
                      Sat:="+", Front:=plusBug)
    If Left(aString, 1) = "+" And Left(ThisNumber, 1) <> "+" Then
        ThisNumber = "+" & ThisNumber
    End If
    
    If RecognizeCode(ThisNumber, aCode, "+49") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+48") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+43") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+34") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+33") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+1") Then
    ElseIf RecognizeCode(ThisNumber, aCode, "+352") Then
    End If
    CheckForCountryCodes = ThisNumber
  
End Function                                     ' O_Goodies.CheckForCountryCodes

'---------------------------------------------------------------------------------------
' Method : Function CheckSimilarityIn
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CheckSimilarityIn(ch As Variant, sim As String) As Boolean
Dim i As Long
Dim chW As String

    For i = 0 To UBound(ch)
        If sim = "*" Then
            CheckSimilarityIn = True
        Else
            chW = ch(i)
            CheckSimilarityIn = IsSimilar(chW, sim)
            If CheckSimilarityIn Then
                Exit For
            End If
        End If
    Next i
End Function                                     ' O_Goodies.CheckSimilarityIn

'---------------------------------------------------------------------------------------
' Method : Sub ClearClipboard
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClearClipboard()


    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
    ClipBoardIsEmpty = TristateTrue

End Sub                                          ' O_Goodies.ClearClipboard

'---------------------------------------------------------------------------------------
' Method : Sub ClipBoard_SetData
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClipBoard_SetData(MyString As String)

Const zKey As String = "O_Goodies.ClipBoard_SetData"
    Call DoCall(zKey, tSub, eQzMode)
 
Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongLong
Dim hClipMemory As LongLong, x As LongLong

    ' Allocate moveable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GH, Len(MyString) + 1)
 
    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)
 
    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)
 
    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If
 
    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        GoTo zExit
    End If
 
    ' Clear the Clipboard.
    x = EmptyClipboard()
 
    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
 
OutOfHere2:
 
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If

zExit:
    Call DoExit(zKey)

End Sub                                          ' O_Goodies.ClipBoard_SetData

'---------------------------------------------------------------------------------------
' Method : ColumnPrint
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Print using columns forcing length
'---------------------------------------------------------------------------------------
Sub ColumnPrint(Optional values As Variant, Optional columns As Variant, Optional sep As String = "§")
Static aCol As Variant
Dim ValArr As Variant
Dim i As Long
Dim cL As Long
Dim lne As String
    
    If IsMissing(columns) Then
        If TypeName(aCol) <> "String()" Then Stop
    Else
        aCol = split(columns, sep)
    End If
    If Not IsMissing(values) Then
        If LenB(values) > 0 Then
            ValArr = split(values, sep)
            If UBound(ValArr) > UBound(aCol) Then Stop
            
            For i = 0 To UBound(ValArr)
                cL = aCol(i)
                If cL < 0 Then
                    lne = lne & RString(ValArr(i), -cL)
                Else
                    lne = lne & LString(ValArr(i), cL)
                End If
            Next i
            Debug.Print lne
        End If
    End If
End Sub                                          ' O_Goodies.ColumnPrint

'---------------------------------------------------------------------------------------
' Method : Sub Combo_Define_DatumsBedingungen
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Combo_Define_DatumsBedingungen(Combo As ComboBox)


    Combo.addItem "keine Datumsbeschränkung"
    Combo.addItem "heute"
    Combo.addItem "ab gestern"
    Combo.addItem "letzte Woche"
    Combo.addItem "letzte 30 Tage"
    Combo.BoundColumn = 0
    Combo.ListIndex = 0

End Sub                                          ' O_Goodies.Combo_Define_DatumsBedingungen

'---------------------------------------------------------------------------------------
' Method : Function CompareNumString
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CompareNumString(str1 As String, str2 As String, Optional comparemode As VbCompareMethod) As Long
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "O_Goodies.CompareNumString"
Dim zErr As cErr

Dim i As Long
Dim minlen As Long
Dim aChar As String
Dim bchar As String
Dim cComp As Long
Dim numComp As Boolean
Dim l1 As Long
Dim l2 As Long
Dim maxlen As Long
Dim aLong As Long
Dim bLong As Long

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)

    If Len(str1) = Len(str2) Then
        GoTo simple
    End If
    CompareNumString = 0                         ' equal as default
    l1 = Len(str1)
    l2 = Len(str2)
    minlen = Min(l1, l2)
    maxlen = Max(l1, l2)
    For i = 1 To minlen
nextdigit:
        aChar = Mid(str1, i, 1)
        bchar = Mid(str2, i, 1)
        If InStr("0123456789.", aChar) > 0 _
        And InStr("0123456789.", bchar) > 0 Then
            numComp = True
        Else
            numComp = False
        End If
        
        If numComp Then
            If Not (aChar = vbNullString Or aChar = ".") Then
                aLong = aLong * 10 + CLng(aChar)
            End If
            If Not (bchar = vbNullString Or bchar = ".") Then
                bLong = bLong * 10 + CLng(bchar)
            End If
            If aLong = bLong Then
                cComp = 0
            ElseIf aLong < bLong Then
                cComp = -1
            Else
                cComp = 1
            End If
        Else
            aLong = 0
            bLong = 0
            cComp = StrComp(aChar, bchar, comparemode)
        End If
        
        If cComp = 0 Then                        ' same
            ' same up to now
        Else
            If numComp Then
                If aChar = bchar Then
                    cComp = 0
                Else
                    GoTo NE
                End If
            Else
                ' cComp <> 0
NE:
                CompareNumString = cComp
                If numComp Then
                    If i = minlen Then
                        If l1 < maxlen Or l2 < maxlen Then
                            i = i + 1
                            GoTo nextdigit
                        End If
                    End If
                Else
                    GoTo funxit
                End If
            End If
        End If
    Next i
    If i < Max(Len(str1), Len(str2)) + 1 Then
simple:
        CompareNumString = StrComp(str1, str2, comparemode)
    End If
funxit:

FuncExit:
    zErr.atFuncResult = CStr(CompareNumString)

ProcReturn:
    Call ProcExit(zErr)

End Function                                     ' O_Goodies.CompareNumString

'---------------------------------------------------------------------------------------
' Method : Function CountStringOccurrences
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CountStringOccurrences(S As String, P As String) As Long


Dim L As Long
Dim T As String

    If LenB(S) = 0 Or LenB(P) = 0 Then
        GoTo ProcRet
    End If
    L = Len(S)
    T = Replace(T, P, vbNullString)                        ' remove all occurrences of p in t
    CountStringOccurrences = (L - Len(T)) / Len(P) ' use number of occurrences removed

ProcRet:
End Function                                     ' O_Goodies.CountStringOccurrences

'---------------------------------------------------------------------------------------
' Method : Function CreateFolderIfNotExists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CreateFolderIfNotExists(strFolderName As String, underWhat As Object, Optional DefaultItemType As Long) As Folder
Dim zErr As cErr
Const zKey As String = "O_Goodies.CreateFolderIfNotExists"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:=strFolderName)
    
Dim ParentFolder As Folder
Dim ParentFolders As Folders
    
    If underWhat.Class = olFolder Then
        Set ParentFolder = underWhat
        Set ParentFolders = underWhat.Folders
    Else
        Set ParentFolders = underWhat
    End If
    
    aBugTxt = "Get or use Folder " & Quote1(strFolderName)
    Call Try(testAll)                               ' Try anything, autocatch
    Set CreateFolderIfNotExists = ParentFolders(strFolderName)
    If Not CatchNC Then
        Call LogEvent("Folder " & CreateFolderIfNotExists.Name & " exists")
        Call ErrReset(4)
        GoTo FuncExit                            ' no problem, file exists
    End If
    
    If StrComp(Hex$(E_AppErr.errNumber), "8004010F", vbTextCompare) = 0 Then
        ' ....wegen der Lesbarkeit vergleiche ich hier:
        ' .....If Hex$(Err.Number) = "8004010F" Then
        '  der "echte" Fehlercode ist  -2147221233 == &H8004010F
        ' ..so (in Hex) kann ich den wiedererkennen. Bedeutet: Object nicht gefunden
        '  wenn kein Type als Parameter mitgegeben, dann als Type "Mails" setzen
        If DefaultItemType = 0 Then
            DefaultItemType = olFolderInbox
        End If
        aBugTxt = "Create Folder of type " & DefaultItemType & " and add to ParentFolders"
        Set CreateFolderIfNotExists = _
                                    ParentFolders.Add(strFolderName, DefaultItemType) ' , olMail)
    Else
        ' was immer sonst passiert sein mag.. ich habs nicht abgefangen..
        Call TerminateApp                        ' hier also Crash & Burn...
    End If
    ' falls das Ordner-Neuanlegen in die Grütze geht...
    If Catch Then
        DoVerify False, "even in debugmode!"
    End If

FuncExit:
    zErr.atFuncResult = CStr(CreateFolderIfNotExists)
    Set ParentFolder = Nothing
    Set ParentFolders = Nothing
    
    Call ProcExit(zErr)
End Function                                     ' O_Goodies.CreateFolderIfNotExists

'---------------------------------------------------------------------------------------
' Method : Function cvExcelVal
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function cvExcelVal(val As Variant) As Variant
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.cvExcelVal"

    cvExcelVal = val
    If LenB(CStr(val)) = 0 Then
        val = Chr(0)
    End If
        
    If val = Chr(0) _
    Or IsDate(val) _
    Or IsNumeric(val) Then
        ' exit function
    Else                                         ' assume it is a string
        cvExcelVal = CStr(val)
        If Left(cvExcelVal, 1) <> "'" Then       ' Do not double that
            cvExcelVal = "'" & CStr(val)
        End If
    End If

FuncExit:

End Function                                     ' O_Goodies.cvExcelVal

' Englische bools aus Deutsch
Function DeBoolToEn(b As Variant) As String
    Select Case LCase(b)
    Case "true", "wahr"
        DeBoolToEn = "True"
    Case "false", "falsch"
        DeBoolToEn = "False"
    Case Else
        Debug.Print b & " ist kein bool'scher Wert"
        If DebugMode Then DoVerify False
    End Select
End Function                                     ' O_Goodies.DeBoolToEn

'---------------------------------------------------------------------------------------
' Method : DecodeSpecialProperties
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Decodes Properties with unusual Values
'---------------------------------------------------------------------------------------
Function DecodeSpecialProperties(hInfo As cInfo, PropName As String) As Boolean
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.DecodeSpecialProperties"
    Call DoCall(zKey, tFunction, eQzMode)

Dim dInfo As cInfo

    With hInfo
        Select Case PropName
        Case "Parent"
            ' special parent olFolder: only care about folder path
            If .iValue.Value.Class = olFolder Then
                Set dInfo = .DrillDown(.iValue.Value)
                dInfo.DecodedStringValue = dInfo.iValue.FolderPath
                dInfo.iAssignmentMode = 1
                dInfo.iType = vbString
                dInfo.iTypeName = "ParentFolder"
                Set hInfo = dInfo
                DecodeSpecialProperties = True
            End If
        
        Case "Actions", "Attachments", "UserProperties", "Recipients", _
             "ReplyRecipients", "Links", "Conflicts" ' all (known) Properties with Count
            Set dInfo = .DrillDown(.iValue.Value)
            dInfo.iTypeName = PropName & "Count"
            dInfo.iArraySize = dInfo.iValue.Count
            dInfo.iAssignmentMode = 1
            dInfo.iIsArray = True
            .DecodedStringValue = "{} " & dInfo.iArraySize & " values"
            Set hInfo = dInfo
            DecodeSpecialProperties = True
        Case "Nothing"
            Set dInfo = .DrillDown(Nothing)
            dInfo.iTypeName = PropName
            dInfo.iAssignmentMode = 2
            dInfo.iType = vbNull
            DecodeSpecialProperties = True
        Case Else
            DecodeSpecialProperties = False
        End Select                               ' .ivalue.name
    End With                                     ' hInfo
    
FuncExit:
    Set dInfo = Nothing

zExit:
    Call DoExit(zKey)

End Function                                     ' O_Goodies.DecodeSpecialProperties

'---------------------------------------------------------------------------------------
' Method : Function DoubleInternalQuotes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DoubleInternalQuotes(S As String, DQ As String) As String
Dim Q As String

    Q = Left(DQ, 1)                              ' use ONLY on nonquoted strings s
    DoubleInternalQuotes = Replace(S, Q, Q & Q)  ' double internal quotes
End Function                                     ' O_Goodies.DoubleInternalQuotes

'---------------------------------------------------------------------------------------
' Method : Sub DumpAllPushDicts
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DumpAllPushDicts()

Const zKey As String = "O_Goodies.DumpAllPushDicts"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                    & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True
Dim aDictVar As Variant
Dim aValue As Variant
Dim aVobject As Variant
Dim aStackDict As Dictionary
Dim TupleNameValue As String
Dim i As Long

    aBugTxt = "Dumping D_PushDict"
    Call Try(allowAll)                              ' Try anything, autocatch, Err.Clear
    For Each aDictVar In D_PushDict
        Set aStackDict = D_PushDict.Item(aDictVar)
        Debug.Print
        Debug.Print "Dictionary: " & aDictVar
        i = 0
        For Each aValue In aStackDict
            i = i + 1
            If VarType(aValue) = vbObject Then
                If aStackDict.Exists(aValue.Key) Then
                    Set aVobject = aStackDict.Item(aValue.Key)
                    If isEmpty(aVobject) Then
                        aVobject = "Empty Key"
                    End If
                Else
                    aVobject = "has no key"
                End If
            Else
                If aStackDict.Exists(aValue) Then
                    If VarType(aValue) = vbObject Then
                        Set aVobject = aValue
                    Else
                        If VarType(aStackDict.Item(aValue)) = vbObject Then
                            Set aVobject = aStackDict.Item(aValue)
                        Else
                            aVobject = CVar(aStackDict.Item(aValue))
                        End If
                    End If
                    If isEmpty(aVobject) Then
                        aVobject = "Empty"
                    Else
                        If VarType(aVobject) = vbObject Then
                            If TypeName(aVobject) = "cTuple" Then
                                TupleNameValue = ", TupleNameValue=" & aVobject.TupleNameValue
                            Else
                                aVobject = "TypeName '" & TypeName(aVobject) _
      & "' not supported"
                            End If
                        Else
                            aVobject = "Trivial=" & CStr(aValue)
                        End If
                    End If
                Else
                    aVobject = "no item value"
                End If
            End If
            Debug.Print i, "StackObject_TypeName=" & TypeName(aValue);
            Debug.Print ", Key=" & aValue.Key;
            Debug.Print ", objKey=" & aVobject;
            Debug.Print TupleNameValue & ", ValueStr=" & CStr(aValue)
            Debug.Print
            Catch
        Next aValue
    Next aDictVar

FuncExit:
    Call ErrReset(0)
    Recursive = False

ProcRet:
End Sub                                          ' O_Goodies.DumpAllPushDicts

'---------------------------------------------------------------------------------------
' Method : Sub EnvironmentPrintout
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub EnvironmentPrintout()
Dim UmgZF As String
Dim Indx As Long
Dim PfadLänge As String

    Indx = 1                                     ' Index mit 1 initialisieren.
    Do
        UmgZF = Environ(Indx)                    ' Umgebungsvariable
        PfadLänge = PfadLänge & vbCrLf & UmgZF
        Indx = Indx + 1                          ' Kein PATH-Eintrag,
    Loop Until UmgZF = vbNullString

    Debug.Print PfadLänge

    MsgBox Indx - 1 & " Umgebungsvariablen:" & PfadLänge
End Sub                                          ' O_Goodies.EnvironmentPrintout

'---------------------------------------------------------------------------------------
' Method : Sub ErrReset
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: ReSet Error information / this does NOT clear information not related to Err
'---------------------------------------------------------------------------------------
Sub ErrReset(Optional upToMode As Long = 0)
Const zKey As String = "O_GoodiesCall ErrReset"
    
Dim ErrOnEntry As Long
    ErrOnEntry = Err.Number
    
    StackDebug = Abs(StackDebugOverride)
    Select Case upToMode
    Case 0                                       ' IgnoreUnhandledError is NOT changed
        aBugTxt = vbNullString
        Call Try(Empty)                          ' Includes:  ErrSnoCatch, Z§ErrSnoCatch
        If MayChangeErr And ErrOnEntry <> 0 Then
            GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
        End If
        GoTo FuncExit
    Case 1
        MayChangeErr = True                      ' always clears Err, ErrorCaught is NOT cleared
        IgnoreUnhandledError = False
        Call Try(Empty)                          ' Includes:  ErrSnoCatch, Z§ErrSnoCatch
        GoTo ShowChange
    Case 2
        MayChangeErr = True                      ' always clears Err, ErrorCaught is NOT cleared
        IgnoreUnhandledError = False
        Call T_DC.N_ClearTermination             ' Includes:  ErrSnoCatch, Z§ErrSnoCatch
        If ErrOnEntry <> 0 Then
            GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
        End If
        GoTo FuncExit
    Case 3
        aBugTxt = vbNullString
        IgnoreUnhandledError = False
        E_AppErr.Permit = Empty                  ' EXcludes:  ErrSnoCatch, Z§ErrSnoCatch
        If Z§ErrSnoCatch Then                    ' ErrorCaught is NOT cleared
            GoTo FuncExit
        End If
        GoTo ShowChange
    Case 4                                       ' keep the Try, but forget all signs of an Err
        With E_Active
            .ErrSnoCatch = False                 ' ErrSnoCatch, Z§ErrSnoCatch cleared
            .errNumber = 0                       ' ErrorCaught is NOT cleared
            .FoundBadErrorNr = 0
            .Description = vbNullString
            .Reasoning = vbNullString
            .Explanations = vbNullString
            ErrorCaught = 0                      ' Clearing this because error was accepted
            T_DC.DCerrSource = .atKey
            ErrDisplayModify = True
            Z§ErrSnoCatch = False                ' what the hell is Z§ErrSnoCatch ???
            GoTo ShowChange
        End With                                 ' E_Active
        If ErrStatusFormUsable Then
            frmErrStatus.fErrNumber = 0
        End If
        If MayChangeErr Then
            If ErrOnEntry <> 0 Then
                Err.Clear
                ErrDisplayModify = True
            End If
        ElseIf ErrOnEntry <> 0 Then
            GoTo ShowChange                      ' ErrorCaught is NOT cleared, must show
        End If
        GoTo FuncExit
    Case Else                                    ' Resets everything, including ErrorCaught
        aBugTxt = vbNullString
        IgnoreUnhandledError = False
        E_AppErr.Permit = Empty
        MayChangeErr = True                      ' always clears Err
        Err.Clear
        E_Active.ErrSnoCatch = False
        ErrorCaught = 0                          ' Clearing this, too
        Call T_DC.N_ClearTermination
        GoTo ShowChange
    End Select                                   ' upToMode
        
ShowChange:
    If ErrDisplayModify Or Not (Z§ErrSnoCatch Or SuppressStatusFormUpdate) Then
        Call ShowStatusUpdate
    End If
    
FuncExit:
    aBugVer = True
    ' NO: aBugTxt = vbNullString, wait for Catch
    Z§ErrSnoCatch = False
    
ProcRet:
End Sub                                          ' O_GoodiesCall ErrReset

'---------------------------------------------------------------------------------------
' Method : Function FindFirstChar
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindFirstChar(ByRef x As String, CharSet As String) As Long
Dim cPos As Long
    For cPos = 1 To Len(x)
        If InStr(CharSet, Mid(x, cPos, 1)) > 0 Then
            FindFirstChar = cPos
            GoTo FuncExit
        End If
    Next cPos

FuncExit:
End Function                                     ' O_Goodies.FindFirstChar

'---------------------------------------------------------------------------------------
' Method : Function FindFirstDate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindFirstDate(ByVal x As String, fdate As Date) As Boolean
Const zKey As String = "O_Goodies.FindFirstDate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction)

Const dateChars As String = "./"
Const digits As String = "0123456789"

Dim cPos As Long
    cPos = FindFirstChar(x, digits)
    If cPos > 0 Then
        x = Mid(x, cPos)
    Else
        FindFirstDate = False
        GoTo FuncExit
    End If
    For cPos = 1 To Min(Len(x), 10)
        ' suche erstes Zeichen, dass nicht im Datum vorkommen kann
        If InStr(digits & dateChars, Mid(x, cPos, 1)) = 0 Then
            Exit For
        End If
    Next cPos
    If cPos < 11 Then
        x = Left(x, cPos - 1)
        If IsDate(x) Then
            fdate = CDate(x)
            FindFirstDate = True
        Else
            FindFirstDate = False
        End If
    Else
        FindFirstDate = False
    End If
  
FuncExit:
    Call ProcExit(zErr)

End Function                                     ' O_Goodies.FindFirstDate

'---------------------------------------------------------------------------------------
' Method : Function FindValueInArray
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindValueInArray(AR As Variant, val As Variant, Optional canPrint As Boolean = True) As Long ' returns its position if found
Dim ExplainS As String
Const MyId As String = "FindValueInArray"

Dim j As Long

    If canPrint Then
        Call getInfo(aInfo, val)
        If aInfo.iAssignmentMode = 1 Then
            ExplainS = aInfo.iValue
        Else
            ExplainS = "# value not printable"
        End If
    End If

    For j = LBound(AR) To UBound(AR)
        If j <= UBound(AR) Then
            If val = AR(j) Then
                FindValueInArray = j
                If canPrint Then
                    E_AppErr.atFuncResult = MyId & " found " & ExplainS & " at pos " & j
                End If
                GoTo ProcRet
            End If
        End If
    Next j
    FindValueInArray = LBound(AR) - 1            ' val not found
    If canPrint Then
        E_AppErr.atFuncResult = MyId & " did not find " & ExplainS
    End If

ProcRet:
End Function                                     ' O_Goodies.FindValueInArray

'---------------------------------------------------------------------------------------
' Method : Sub FirstDiff
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FirstDiff(ByVal S1 As String, ByVal S2 As String, p1 As String, p2 As String, contextlen As Long, targetlen As Long, dots As String, FirstDiff As String, Optional Ignore As String)
Const zKey As String = "O_Goodies.FirstDiff"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")

Dim i As Long, j As Long, c1 As String, c2 As String, lmin As Long
Dim k As Long
Dim lmax As Long
Dim tlMax As Long
Dim mL As Long
Dim i1s As String, i2s As String
Dim Prefix As String

    FirstDiff = vbNullString
    p1 = vbNullString
    p2 = vbNullString
    i1s = Right("    " & WorkIndex(1), k)
    i2s = Right("    " & WorkIndex(2), k)
    lmax = Max(Len(S1), Len(S2))
    If targetlen = 0 Then
        targetlen = lmax                         ' unlimited length, never cut
    End If
    lmin = Min(contextlen, targetlen)            ' must Show this much
    tlMax = Max(contextlen, targetlen)           ' if we need 2 lines, use max possible
    Prefix = "     "
    If S1 = S2 Then                              ' no first diff at all
        If lmax > lmin Then
            p1 = Left(S1, lmin - Len(dots)) & dots
            p2 = Left(S2, lmin - Len(dots)) & dots ' can not Show all of it
        Else
            p1 = Left(S1, lmin)
            p2 = Left(S2, lmin)
        End If
        FirstDiff = WorkIndex(1) & "==" & WorkIndex(2) & ": " & Quote(p1)
        GoTo FuncExit
    End If
    
    mL = Len(S1) + Len(S2)
    If 2 * (mL + k + 2) < contextlen Then        ' fits one line
        If Len(S1) > lmin And Right(S1, Len(dots)) <> dots Then
            p1 = Left(S1, lmin - Len(dots)) & dots
        Else
            p1 = S1
        End If
        If Len(S2) > lmin And Right(S2, Len(dots)) <> dots Then
            p2 = Left(S2, lmin - Len(dots)) & dots
        Else
            p2 = S2
        End If
        If mL < targetlen Then
            c1 = " / "
        Else
            c1 = vbCrLf
        End If
        FirstDiff = i1s & ": " & Quote(p1) & b & Quote(c1 & i2s) & ": " & Quote(p2)
        GoTo FuncExit
    Else
        p1 = vbNullString
        p2 = vbNullString
    End If
    ' find difference
    i = 1
    j = 1
    Do While i <= Len(S1)
        c1 = vbNullString
        c2 = vbNullString
        If j > Len(S2) Then
            Exit Do
        End If
        c1 = Mid(S1, i, 1)
        c2 = Mid(S2, j, 1)
        If c1 <> c2 Then
            If LenB(Ignore) > 0 And InStr(c1, Ignore) > 0 Then
                i = i + 1                        ' no relevance if ignore character
            ElseIf Ignore <> vbNullString And InStr(c2, Ignore) > 0 Then
                j = j + 1
            Else
                Exit Do                          ' relevant mismatch
            End If
        Else
            i = i + 1
            j = j + 1
        End If
    Loop                                         ' all characters in S1 and S2 until mismatch
    
    If i + j > mL Then
        DoVerify False, " how can they be different but reach the end of the loop???"
        FirstDiff = vbNullString
        GoTo FuncExit
    End If
    
    If i < lmin - k Then                         ' we need not cut beginning if we fit in
        i = 1                                    ' start from beginning
        FirstDiff = i1s & ": " & Left(S1, lmin) & vbCrLf
        FirstDiff = FirstDiff & i2s & ": " & Left(S2, lmin)
    Else
        FirstDiff = String(k, b) & "erster Unterschied an Position " & i & B2
        k = Len(i1s) + 3
        If i < 6 Or i < lmin Then                ' dont strip front
            If i < Len(FirstDiff) + k Then
                c1 = vbCrLf & String(i + k, b) & "V" ' V is down arrow
            Else
                c1 = String(i - Len(FirstDiff) + k, b) & "V"
            End If
            FirstDiff = FirstDiff & c1 & vbCrLf & i1s & ": " & Mid(S1, 1, lmin)
            FirstDiff = FirstDiff & vbCrLf & i2s & ": " & Mid(S2, 1, lmin)
        Else                                     ' strip front chars
            If i < Len(FirstDiff) Then
                c1 = vbCrLf & String(i + k, b) & "V"
            Else
                c1 = String(i - Len(FirstDiff) + k, b) & "V"
            End If
            FirstDiff = FirstDiff & vbCrLf & i1s & ": " & dots & Mid(S1, i, lmin - Len(dots))
            ' append second stripped value
            FirstDiff = FirstDiff & vbCrLf & i2s & ": " & dots & Mid(S2, j, lmin - Len(dots))
        End If
    End If
    
    If i >= lmin Then                            ' cut front
        '       will we need to cut at all?
        j = i + lmin - lmax
        If j > 0 Then                            ' true if both fit already
            i = lmax - j                         ' start relative to end of shorter
            ' i will still be > 1, so
            ' always Show dots at start
        End If
        p1 = dots
        p2 = dots
        tlMax = tlMax - Len(dots)
    End If
    
    If i + tlMax > Len(S1) Or Right(S1, Len(dots)) = dots Then
        p1 = p1 & Mid(S1, i, tlMax)              ' can't Show more than total length, no double dots
    Else
        p1 = p1 & Mid(S1, i, tlMax - Len(dots)) & dots
    End If
    If i + tlMax > Len(S2) Or Right(S2, Len(dots)) = dots Then
        p2 = p2 & Mid(S2, i, tlMax)              ' can't Show more than total length
    Else
        p2 = p2 & Mid(S2, i, tlMax - Len(dots)) & dots
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' O_Goodies.FirstDiff

'---------------------------------------------------------------------------------------
' Method : Function FormatRight
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FormatRight(Value As Variant, length As Long, Optional frac As Long = 2) As String

' simple replacement (for strings only!)     : Right(String(length, b) & value, length)
' replacement (for int or long) use string of: Format(Value, String(length, "#") & "0")
' replacement (for single or double  use two steps   ( Example is for length=8, frac=4 )
'    FormatRight = Format(Value, String(length - frac - 1, "#") & "0" & "." & String(frac, "0"))
'    FormatRight = Right(String(length, b) & FormatRight, length)
' ex: tRT = Format(time, padNum3 & "." & String(4, "0"))
' ex: tRT = Right(String(tRT, b) & tRT, 8)

    If VarType(Value) = vbString Then
        FormatRight = Right(String(length, b) & Value, length)
        GoTo FuncExit
    ElseIf Not IsNumeric(Value) Then
        DoVerify False
        GoTo FuncExit                            ' returning vbNullString
    End If
    If VarType(Value) = vbDouble Or VarType(Value) = vbSingle Then
        FormatRight = Format(Value, String(length - frac - 1, "#") _
                           & "0" & "." & String(frac, "0"))
    Else
        FormatRight = Format(Value, String(length, "#") & "0")
    End If
    FormatRight = Right(String(length, b) & FormatRight, length)
    
FuncExit:

End Function                                     ' O_Goodies.FormatRight

' for loop functions requireing an array
Function GenArr(S As Variant, Optional splitter As String = b) As Variant
    If isArray(S) Then
        GenArr = S
    Else
        If LenB(splitter) = 0 Then
            GenArr = Array(S)
        Else
            GenArr = split(S, splitter)
        End If
    End If
End Function                                     ' O_Goodies.GenArr

'---------------------------------------------------------------------------------------
' Method : Function GetClipboardString
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetClipboardString() As String          ' also sets ClipBoardIsEmpty

Const zKey As String = "O_Goodies.GetClipboardString"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction)

Dim GTO As Variant
    
    ClipBoardIsEmpty = TriStateMixed             ' error in Extras/Verweise
    Set DataObject = New MSForms.DataObject
    Call Try(testAll)
    DataObject.GetFromClipboard
    GTO = DataObject.GetText(1)
    If GTO = False Then
        If Catch Then
            ClipBoardIsEmpty = TristateTrue      ' = True
        End If
        GetClipboardString = vbNullString
    Else
        ClipBoardIsEmpty = TriStateFalse         ' = False
        GetClipboardString = GTO
    End If
    
    Call ErrReset(0)
    Set DataObject = Nothing

ProcReturn:
    Call ProcExit(zErr)

End Function                                     ' O_Goodies.GetClipboardString

' get  actDate    as the current date in normal format (always)
' or   DateId     as yyyy dd mm or yyyy dd mm_hh mm    (+1, 0)
' and  DateIdNB   as yyyyddmm or yyyyddmmhhmm          (+2, <=0)
Function GetDateId(Optional withTime As Long = 0) As String
    actDate = Now()
    Select Case withTime
    Case 2, -1
        GetDateId = Format(actDate, "yyyymmdd_hhmm")
    Case 1, 0
        GetDateId = Format(actDate, "yyyymmdd")
    Case Else                                    ' e.g. -1 will set global variables, +1 or +2 will not
        GetDateId = Format(actDate, "yyyymmdd_hhmm")
    End Select                                   ' withTime
    
    If withTime <= 0 Then                        ' set DateIdNB and DateId, else leave unchanged
        DateIdNB = GetDateId                     ' always no Blanks
        If Len(GetDateId) > 8 Then               ' with time and blanks
            DateId = Format(actDate, "yyyy mm dd hh mm")
        Else
            DateId = Format(actDate, "yyyy mm dd")
        End If
    End If
End Function                                     ' O_Goodies.GetDateId

'---------------------------------------------------------------------------------------
' Method : Function GetEnvironmentVar
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetEnvironmentVar(Name As String) As String

Const zKey As String = "O_Goodies.GetEnvironmentVar"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tFunction, _
                  ExplainS:="O_Goodies")
    
    GetEnvironmentVar = String(255, 0)
    Call GetEnvironmentVariable(Name, GetEnvironmentVar, Len(GetEnvironmentVar))
    GetEnvironmentVar = TrimNul(GetEnvironmentVar)

ProcReturn:
    Call ProcExit(zErr)

End Function                                     ' O_Goodies.GetEnvironmentVar

'---------------------------------------------------------------------------------------
' Method : Sub GetLine
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetLine(i As Long, varr As Variant)


Dim j As Long
Dim aSheet As Worksheet

    Set aSheet = O
    For j = 1 To UBound(varr)                    ' not using lbound = 0 as after split
        Set aCell = aSheet.Cells(i, j)
        If DebugMode Then
            aCell.Select
        End If
        If VarType(aCell.Value) = vbString Then
            If Left(aCell.Value, 1) = "'" Then
                varr(j) = Trim(Mid(aCell.Value, 2))
            Else
                varr(j) = Trim(aCell)
            End If
        Else
            varr(j) = aCell
        End If
    Next j

End Sub                                          ' O_Goodies.GetLine

'---------------------------------------------------------------------------------------
' Method : Function GetPart
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetPart(Source As String, Optional Fsep As String = "%", Optional Bsep As String, Optional sPos As Long, Optional ePos As String, Optional Brace As String = "( )") As String
Dim Braces As Variant

    sPos = InStr(Source, Fsep)
    If sPos > 0 Then
        If Bsep = vbNullString Then
            Bsep = Fsep
        End If
        Braces = split(Brace)
        GetPart = Mid(Source, sPos + Len(Fsep))
        Source = Left(Source, sPos - 1)
        ePos = InStr(GetPart, Bsep)
        If ePos > 0 Then
            Source = Source & Mid(GetPart, ePos + Len(Bsep))
            GetPart = Left(GetPart, ePos - 1)
        End If
        If UBound(Braces) >= 0 Then
            GetPart = Braces(0) & GetPart & Braces(UBound(Braces))
        End If
    End If
End Function                                     ' O_Goodies.GetPart

' find word surrounding pos in source, delimited by chars in split, giving word position
Function GetThisWord(Source As String, pos As Long, lWord As Long, ByVal split As String, wPos As Long) As String
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "O_Goodies.GetThisWord"
Dim zErr As cErr

Dim N As Long, j As Long, k As Long

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)

    wPos = pos                                   ' word starts here if nothing is before it
    If wPos = 1 Then GoTo getTail
    ' find beginning of the word
    For j = wPos - 1 To 1 Step -1
        For k = 1 To Len(split)
            If Mid(Source, j, 1) = Mid(split, k, 1) Then
                ' we found a delimiter up front
                wPos = j + 1
                GoTo getTail
            End If
        Next k
    Next j
    
getTail:
    For N = pos + lWord To Len(Source)
        For k = 1 To Len(split)
            If Mid(Source, N, 1) = Mid(split, k, 1) Then
                ' we found a delimiter at back
                GoTo gotTail
            End If
        Next k
    Next N
        
gotTail:
    GetThisWord = Mid(Source, wPos, N - wPos)

FuncExit:
    zErr.atFuncResult = CStr(GetThisWord)

ProcReturn:
    Call ProcExit(zErr)
  
End Function                                     ' O_Goodies.GetThisWord

'---------------------------------------------------------------------------------------
' Method : Function GetVtypeInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetVtypeInfo(vTypeName As String) As VbVarType
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.GetVtypeInfo"
    Call DoCall(zKey, tFunction, eQzMode)

Dim vType As VbVarType

    ' the sequence of the cases should be according to likelyhood to minimize branches
    Select Case vTypeName
    Case "ItemProperty"
        vType = olItemProperty
    Case "String"                                ' 8
        vType = vbString
    Case "Variant"                               ' 12
        vType = vbVariant
    Case "Boolean"                               ' 11
        vType = vbBoolean
    Case vbNullString                                      '  Name does not exist
        vTypeName = "Null"
        vType = vbNull
    Case "Null", "Nothing"                       '  Simulated, such a Name does not really exist
        vType = vbVariant
    Case "Long"                                  ' 2
        vType = vbInteger
    Case "Date"                                  ' 7
        vType = vbDate
    Case "Object"                                ' 9
        vType = vbObject
    Case "Long"                                  ' 3
        vType = vbLong
    Case "Double"                                ' 5
        vType = vbDouble
    Case "Single"                                ' 4
        vType = vbSingle
    Case "Empty"                                 ' 0
        vType = vbEmpty
    Case "Byte"                                  ' 17
        vType = vbByte
    Case "LongLong"                              ' 20
        vType = 20&                              ' LongLong, only on 64 bit systems
    Case "Null"                                  ' 1
        vType = vbNull
    Case "Currency"                              ' 6
        vType = vbCurrency
    Case "Error"                                 ' 10
        vType = vbError
    Case "DataObject"                            ' 13
        vType = vbDataObject
    Case "Decimal"                               ' 14
        vType = vbDecimal
    Case Else
        vType = 50                               ' this should mean we are using a (User-) Class Object
    End Select
    GetVtypeInfo = vType

zExit:
    Call DoExit(zKey)

End Function                                     ' O_Goodies.GetVtypeInfo

'---------------------------------------------------------------------------------------
' Method : Function GetWordContaining
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetWordContaining(Source As String, Hook As String, Optional lWord As Long, Optional split As String = " ", Optional wPos As Long, Optional Instance As Long) As String
Const zKey As String = "O_Goodies.GetWordContaining"
Static zErr As New cErr

Dim HookPos As Long
Dim aInstance As Long
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")
    
    HookPos = 1
    HookPos = InStr(HookPos, Source, Hook)
    If HookPos = 0 Then
notfound:
        lWord = 0
        wPos = 0
        GoTo ProcReturn
    End If
nextinstance:                                    ' get the Instance of hook
    aInstance = aInstance + 1
    If Instance > aInstance Then
        HookPos = InStr(HookPos + 1, Source, Hook)
        If HookPos = 0 Then
            GoTo notfound
        End If
        If aInstance < Instance Then
            GoTo nextinstance
        End If
    End If
    
    GetWordContaining = GetThisWord(Source, HookPos, lWord, split, wPos)
    
ProcReturn:
    Call ProcExit(zErr)

ProcRet:
End Function                                     ' O_Goodies.GetWordContaining

'---------------------------------------------------------------------------------------
' Method : Function Hex8
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Hex8(aNum As Variant) As String
    Hex8 = Right("00000000" & Hex(aNum), 8)
End Function                                     ' O_Goodies.Hex8

'---------------------------------------------------------------------------------------
' Method : Function HexN
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function HexN(O As Variant, N As Long, Optional L As Long) As String
    HexN = Right(String(N, "0") & Hex(O), N)
    If L > 0 Then
        HexN = Left(HexN, L)
    End If
End Function                                     ' O_Goodies.HexN

'---------------------------------------------------------------------------------------
' Method : Sub InspectType
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Detect Type Name / Type Id etc. into cInfo-Object
'          if no previous AssignmentMode, get AssignmentMode 1 or 2
'---------------------------------------------------------------------------------------
Sub InspectType(ByRef aVar As Variant, fInfo As cInfo)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.InspectType"
    Call DoCall(zKey, tSub, eQzMode)

Dim vArraySize As Long
Dim vIsArray As Boolean
Dim vAssignmentMode As Long
Dim vType As VbVarType
Dim vTypeId As String
Dim vScalarType As Long

Dim lb As Long
Dim ub As Long
    
    With fInfo
        vType = VarType(aVar)
        vTypeId = TypeName(aVar)
        vIsArray = False
        vArraySize = -99
        vScalarType = -99
        vAssignmentMode = -99
        
        If vType >= vbArray Then
            vIsArray = True
            vAssignmentMode = 0
            vTypeId = TypeName(aVar)
            vTypeId = "Array of " & vTypeId
            .iType = vType - vbArray
            If .iType = vbVariant Then
ABounds:
                ub = UBound(aVar)
                lb = LBound(aVar)
                vArraySize = ub - lb + 1
                If vArraySize >= 0 Then
                    vTypeId = Left(vTypeId, Len(vTypeId) - 1) _
      & lb & " to " & ub & ")"
                    vAssignmentMode = 1
                End If
            ElseIf .iType = vbByte Then
                vAssignmentMode = 2
            Else
                DoVerify False, "array of WHAT???"
                vArraySize = UBound(aVar.Value) - LBound(aVar.Value) + 1
            End If
        Else
            If vTypeId = "ItemProperty" Then
                .iClass = olItemProperty         ' here, always work on value
                If InStr(dftRule.clsNotDecodable.aRuleString & b, aVar.Name & b) = 0 Then
                    DoVerify T_DC.DCAllowedMatch = testAll, "Expected to have ** from Caller"
                    vScalarType = IsScalar(TypeName(aVar.Value))
                    If ErrorCaught <> 0 Then
                        .DecodedStringValue = "# No value for variable Type " & aVar.Name & vbCrLf & "# " & Err.Description
                    End If
                    Call ErrReset(4)
                End If
            Else
                vScalarType = IsScalar(vTypeId) - .iDepth ' -.iDepth to indicate (User-) (not Item-) Property
            End If
            If vScalarType > 0 Then              ' is a scalar
                vAssignmentMode = 1
            ElseIf vScalarType < 0 Then
                vAssignmentMode = 0              ' not decodable or decode not wanted
            Else
                vAssignmentMode = 2              ' some sort of object
            End If
        End If
    
        .iArraySize = vArraySize
        .iIsArray = vIsArray
        .iType = vType
        .iTypeName = vTypeId
        .iScalarType = vScalarType
        If .iAssignmentMode <= 0 Then            ' gate deciding if we want change
            .iAssignmentMode = vAssignmentMode
        Else
            DoVerify .iAssignmentMode = vAssignmentMode, _
                     "Attention: .iAssignmentMode will not be changed, is=" _
                   & .iAssignmentMode & " rejecting new=" & vAssignmentMode
        End If
    End With                                     ' fInfo

zExit:
    Call DoExit(zKey)

End Sub                                          ' O_Goodies.InspectType

'---------------------------------------------------------------------------------------
' Method : Function IntAssignIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IntAssignIfChanged(ByRef Target As Long, Source As Long, modified As Boolean) As Integer


    If Source <> Target Then                     ' only valid for Integer not WITHIN Objects
        If ChangeAssignReverse Then
            IntAssignIfChanged = Target
        Else
            IntAssignIfChanged = Source
        End If
        modified = True
    End If

End Function                                     ' O_Goodies.IntAssignIfChanged

'---------------------------------------------------------------------------------------
' Method : Function IsBetween
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsBetween(x, W2, L) As Boolean
    If x >= L And x <= W2 Then
        IsBetween = True
    End If
End Function                                     ' O_Goodies.IsBetween

'---------------------------------------------------------------------------------------
' Method : Function IsOneOf
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsOneOf(AR As Variant, Testvar As Variant) As Long
Dim x As Variant
Dim i As Long
    i = 0
    If VarType(AR) >= vbArray Then
        For Each x In AR
            i = i + 1
            If x = Testvar Then
                IsOneOf = i
                GoTo ProcRet
            End If
        Next x
    Else
        DoVerify False, "only designed for arrays ar=array(a,b,c...)"
    End If
    
ProcRet:
End Function                                     ' O_Goodies.IsOneOf

' find a term in s matching one of array vSo
' if ignore specified, term is delimited
' if op >= len(s) lookup is in reverse
' Note: default comptype is vbTextCompare (not vbbinarycompare)
Function IsOneOfPos(S As String, vSo As Variant, Optional sep As String = b, Optional Op As Long = 1, Optional Ignore As Variant, Optional ByVal compType As VbCompareMethod = vbTextCompare) As Long
Const zKey As String = "O_Goodies.IsOneOfPos"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim v As Variant
Dim splitignore As Variant
Dim vS As Variant
Dim FI As Variant
Dim P As Long
Dim k As Long
Dim reverseIndex As Boolean
Dim ls As String

    ' because instr / instrrev appear to ignore the compare method, we do it in Lcase
    If compType = vbTextCompare Then
        ls = LCase(S)
    Else
        ls = S
    End If
    If Op >= Len(ls) Then                        ' find in reverse
        reverseIndex = True
        k = 1                                    ' not used here
    Else
        k = Op                                   ' original start for search
    End If
    If Not IsMissing(Ignore) Then
        If isArray(Ignore) Then
            splitignore = Ignore
        End If
    End If
    If isArray(vSo) Then
        vS = vSo
    Else
        If VarType(vSo) = vbString Then
            If LenB(vSo) = 0 Then
                DoVerify False, " parm invalid"
            End If
            vS = split(vSo, sep)
        Else
            DoVerify False, " not implemented"
        End If
    End If
    For Each v In vS
        If compType = vbTextCompare Then
            v = LCase(v)
        End If
nextOne:
        If reverseIndex Then                     ' find in reverse
            P = InStrRev(ls & sep, v & sep)      ' comptype not working in Instr/Rev
        Else
            If P = 0 Then
                P = k
            End If
            P = InStr(P, ls & sep, v & sep)
        End If
        If P = 0 And Not IsMissing(Ignore) Then
            If Not isArray(splitignore) Then
                splitignore = split(Ignore, b)
            End If
            For Each FI In splitignore
                P = k
                If reverseIndex Then             ' find in reverse
                    P = InStrRev(ls, v)
                Else
                    P = InStr(P, ls, v)
                End If
                If P > 0 Then
                    If StrComp(FI, Mid(P, ls + Len(v), Len(FI)), compType) Then
                        Exit For
                    End If
                End If
            Next
        End If
        If P > 0 Then
            If P > 1 Then
                If LenB(sep) > 0 Then            ' sonst kein test auf Wort
                    If Mid(ls, P - 1, 1) <> sep Then
                        P = P + 1
                        k = P
                        If P <= Len(ls) Then
                            GoTo nextOne
                        End If
                    End If
                End If
            End If
            GoTo loopex
        End If
        k = Op
    Next v
loopex:
    IsOneOfPos = P
    Op = Len(v) + Len(sep)                       ' end of recognized word

FuncExit:
    zErr.atFuncResult = CStr(IsOneOfPos)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' O_Goodies.IsOneOfPos

'---------------------------------------------------------------------------------------
' Method : IsScalar
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Find if aTypeName is Scalar and optionally convert to aType
'---------------------------------------------------------------------------------------
Function IsScalar(aTypeName As String) As VbVarType
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.IsScalar"
    Call DoCall(zKey, tFunction, eQzMode)
    
Dim i As Long
Dim aStr As String
Dim aChar As String

    If DebugLogging Then
        If aTypeName = "RTFBody" Then
            aStr = vbNullString
            Debug.Print "------------ RTFBody ------------"
            For i = 1 To UBound(aTD.adItemProp.Value)
                aChar = Chr(aTD.adItemProp.Value(i))
                If aChar = vbCr Then
                    aStr = aStr & vbCrLf
                ElseIf aChar <> vbLf Then
                    aStr = aStr & aChar
                End If
            Next i
            Debug.Print aStr
            Debug.Print "-------- End RTFBody ------------"
        End If
    End If
    If isNotDecodable(aTypeName) Then
        IsScalar = -1                            ' no decode possible or wanted
    Else
        i = InStr(ScalarTypeNames, aTypeName & b)
        If i > 0 Then
            IsScalar = True
            IsScalar = dSType.Item(i)            ' always >0
        Else
            IsScalar = 0                         ' non Scalar Vartype
        End If
    End If
    
zExit:
    Call DoExit(zKey)

End Function                                     ' O_Goodies.IsScalar

'---------------------------------------------------------------------------------------
' Method : Function IsSimilar
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsSimilar(A As String, b As String, Optional ignoreCase As Boolean)
Dim la As Long
Dim lb As Long
Dim le As Long

    la = Len(A)
    lb = Len(b)
    If la < lb Then
        le = la
    Else
        le = lb
    End If
    If le = 0 Then
        If la = lb Then
            IsSimilar = True
        Else
            IsSimilar = False
        End If
    Else
        If ignoreCase Then
            IsSimilar = Left(LCase(A), le) = Left(LCase(b), le)
        Else
            IsSimilar = Left(A, le) = Left(b, le)
        End If
    End If
End Function                                     ' O_Goodies.IsSimilar

'---------------------------------------------------------------------------------------
' Method : Function IsUcase
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsUcase(b As String) As Boolean
    IsUcase = (b <> LCase(b))
End Function                                     ' O_Goodies.IsUcase

'---------------------------------------------------------------------------------------
' Method : Function LastTrail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function LastTrail(S As String, at As String) As String
    arr = split(S, at)
    LastTrail = arr(UBound(arr))
End Function                                     ' O_Goodies.LastTrail

'---------------------------------------------------------------------------------------
' Method : LimitAppended
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Abschneiden des Anfangs wenn maxlänge überschritten
'---------------------------------------------------------------------------------------
Function LimitAppended(Sta As String, aWhat As String, maxl As Long, px As String) As String
Dim i As Long, ldiff As Long
    LimitAppended = Sta & aWhat
    ldiff = Len(LimitAppended) - maxl - Len(px)
    If ldiff >= 0 Then                           ' kein Platz mehr im String: ersten Trenner finden
        i = InStr(Sta, ", ")
        LimitAppended = px & Mid(LimitAppended, i + 2)
    End If
End Function                                     ' O_Goodies.LimitAppended

'---------------------------------------------------------------------------------------
' Method : Function LongAssignIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function LongAssignIfChanged(Target As Long, Source As Long, modified As Boolean) As Long


    If Source <> Target Then                     ' only valid for Long not WITHIN Objects
        If ChangeAssignReverse Then
            LongAssignIfChanged = Target
        Else
            LongAssignIfChanged = Source
        End If
        modified = True
    End If

End Function                                     ' O_Goodies.LongAssignIfChanged

'---------------------------------------------------------------------------------------
' Method : Function LPad
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function LPad(N As Long, L As Long, Optional FrontChar As String = "0") As String
    LPad = Right(String(L, FrontChar) & CLng(N), L)
End Function                                     ' O_Goodies.LPad

'---------------------------------------------------------------------------------------
' Method : Function LRString
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Construct a string with left and right justified strings
'---------------------------------------------------------------------------------------
Function LRString(lPart As String, rPart As String, Indent As Long, TLen As Long, Optional rCutL As Long)
Dim P As Long
Dim px As String
Dim pa As Long
Dim alPart As String

    alPart = lPart
    pa = Abs(Indent)
    P = TLen - Len(rPart) - pa - Len(alPart) - 1
    If P < 1 Then
        If rCutL > 0 Then
            px = " ." & Mid(rPart, rCutL - P - 1)
        Else
            If InStr(rPart, ".") > 0 Then
                px = " ." & Tail(rPart, ".")
            Else
                px = rPart
                P = P + 2                        ' did not use " ." prefix gives 2 more chars
            End If
        End If
        alPart = Mid(alPart, 4 - P)
        P = TLen - Len(px) - Indent - Len(alPart) - 1
        If P < 1 Then
            P = 1
        End If
    Else
        px = rPart
    End If
    LRString = String(pa, b) & alPart & String(P, b) & px

End Function                                     ' O_Goodies.LRString

'---------------------------------------------------------------------------------------
' Method : LString
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Left pad to fixed length
'          needed only because there is no Function, only Statement LSet
'---------------------------------------------------------------------------------------
Function LString(x As Variant, L As Long) As String
    LString = String(L, b)
    LSet LString = x
End Function                                     ' O_Goodies.LString

'---------------------------------------------------------------------------------------
' Method : Function Max
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Max(A, b) As Variant

    If A > b Then
        Max = A
    Else
        Max = b
    End If
End Function                                     ' O_Goodies.Max

'---------------------------------------------------------------------------------------
' Method : Function Min
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Min(A, b) As Variant
    If A < b Then
        Min = A
    Else
        Min = b
    End If
End Function                                     ' O_Goodies.Min

'---------------------------------------------------------------------------------------
' Method : Function NextNumberInString
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function NextNumberInString(str As String, ByRef start As Long) As String
Dim ltr As String
    For start = start To Len(str)
        ltr = Mid(str, start, 1)
        If ltr >= "0" And ltr <= "9" Then
            NextNumberInString = ltr
            Exit For
        End If
    Next start
End Function                                     ' O_Goodies.NextNumberInString

'---------------------------------------------------------------------------------------
' Method : Function NonModalMsgBox
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function NonModalMsgBox(Prompt As String, B1Label As String, Optional B2Label As String, Optional Title As String) As Long
Const zKey As String = "O_Goodies.NonModalMsgBox"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim MsgFrm As frmMessage
Dim sTime As Variant
Dim sBlockEvents As Boolean
Dim ReShowFrmErrStatus As Boolean
Dim MsgHdl As cFindWindowParms
Dim TotalTime As Single
Dim saveErrEx As Boolean
Dim WaitCycles As Long

    sBlockEvents = E_Active.EventBlock
    If frmErrStatus.Visible Then
        Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
        ReShowFrmErrStatus = True
    End If
    saveErrEx = ErrExActive
    Call ErrEx.Disable
    If ErrStatusFormUsable Then
        frmErrStatus.fUseErrExOn = vbNullString
    End If
    Set MsgFrm = New frmMessage
    MsgFrm.SetMsgBoxResponse (0)
    If LenB(Title) > 0 Then
        MsgFrm.Caption = Title
    End If
    MsgFrm.Message = Prompt
    MsgFrm.B1.Caption = B1Label
    If LenB(B2Label) = 0 Then
        MsgFrm.B2.Visible = False
    Else
        MsgFrm.B2.Caption = B2Label
    End If
    MsgFrm.Show vbModeless
    On Error GoTo BadExit
    E_Active.EventBlock = False
    If ErrStatusFormUsable Then
        frmErrStatus.fNoEvents = E_Active.EventBlock
        Call BugEval
    End If
    NonModalMsgBox = 0
    Do While NonModalMsgBox = 0
        doMyEvents                               ' allow interaction, delay and wait
        Call WindowSetForeground(MsgFrm.Caption, MsgHdl)
        If Wait(0.2, trueStart:=sTime, TotalTime:=TotalTime, Retries:=WaitCycles) _
        Or WaitCycles > 300 Then                 ' true in debug mode
            NonModalMsgBox = 1                   ' ? Check this in debugmode!
            DoVerify False, "NonModalMsgBox wait end forced", ShowMsgBox:=True
            Exit Do
        End If
        MsgFrm.ResponseWaits = MsgFrm.ResponseWaits + 1
        NonModalMsgBox = MsgFrm.GetMsgBoxResponse
    Loop
    
    MsgFrm.Hide
    Call LogEvent("NonModalMsgBox waited " _
                & MsgFrm.ResponseWaits & " cycles", eLall)
    Set MsgFrm = Nothing
    GoTo FuncExit
BadExit:
    NonModalMsgBox = 0

FuncExit:
    Set MsgHdl = Nothing
    If ReShowFrmErrStatus Then
        Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
    End If
    If saveErrEx And LenB(UseErrExOn) > 0 Then
        Call ErrEx.Enable(UseErrExOn)
    End If
    E_Active.EventBlock = sBlockEvents
    If ErrStatusFormUsable Then
        frmErrStatus.fUseErrExOn = UseErrExOn
        Call BugEval
    End If
    Call ProcExit(zErr)
    
    zErr.atFuncResult = CStr(NonModalMsgBox)
End Function                                     ' O_Goodies.NonModalMsgBox

'---------------------------------------------------------------------------------------
' Method : Function NormalizeTelefonNumber
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function NormalizeTelefonNumber(original As String, Optional Reassign As Boolean) As String
Const zKey As String = "O_Goodies.NormalizeTelefonNumber"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim ThisNumber As String, countryCode As String, areaCode As String
Dim localNumber As String, extensionNumber As String
Dim i As Long
Dim startswith3stars As Boolean
Dim Result As String

    Result = original
    If Left(Result, 3) = "***" Then
        DoVerify False, " wo kommt den so was vor?"
        startswith3stars = True
        Result = Mid(Result, 4)                  ' could be CarTelephoneNumber
    Else
        i = InStr(Result, "*")
        If i > 0 Then
            NormalizeTelefonNumber = Mid(Result, i) ' numbers like *123# have no area code
            Result = vbNullString                          ' no attempt to change something else,
            GoTo FuncExit
        End If
    End If
    ThisNumber = Trim(Replace(Replace(Result, ")", b), "/", b))
    areaCode = vbNullString
    countryCode = vbNullString
    localNumber = vbNullString
    If LenB(Result) = 0 Then
        GoTo FuncExit
    End If
    
    ' numbers without country codes or with nondefault codings thereof
    If Left(ThisNumber, 2) = "(+" Or Left(ThisNumber, 2) = "00" Then
        ThisNumber = "+" & Trim(Mid(ThisNumber, 3))
    End If
    
    If Left(ThisNumber, 1) = "(" Then
        ThisNumber = Trim(Mid(ThisNumber, 2))    ' drop this char and check some countries in brackets
        ThisNumber = CheckForAreaCodes(ThisNumber, areaCode)
        If Left(ThisNumber, 1) = "0" Then
            GoTo localCoding
        End If
        ' this may be unsafe !?
        ' Debug.Assert False                                ' manual inspection of thisnumber is better
        
    End If
    
    If CountStringOccurrences(ThisNumber, "+") > 1 Then
        ThisNumber = "+" & Replace(LastTrail(ThisNumber, "+"), b, vbNullString)
    End If
    
    If Left(ThisNumber, 1) = "+" Then            ' standard coding
        localNumber = Trim(Tail(ThisNumber, b, countryCode))
normalCoding:
        If LenB(countryCode) = 0 Then            ' kein Erfolg mit b als Trennzeichen
            localNumber = Trim(Tail(localNumber, "-", countryCode))
        End If
        If LenB(countryCode) = 0 Then
            localNumber = CheckForCountryCodes(localNumber, countryCode)
        Else
            If countryCode = "+1" Then
                localNumber = Replace(localNumber, "(", vbNullString)
                areaCode = Left(localNumber, 3)
                localNumber = Trim(Mid(localNumber, 4))
                localNumber = Replace(localNumber, b, vbNullString, 1, 1)
            Else
                areaCode = CheckForCountryCodes(countryCode, countryCode)
                If areaCode = countryCode Then
                    areaCode = vbNullString                ' unregistered county is not an area code
                End If
            End If
        End If
        If Left(localNumber, 1) = "(" Then
            localNumber = Trim(Mid(localNumber, 2))
        End If
        If Left(localNumber, 1) = "0" Then
            localNumber = Trim(Mid(localNumber, 2))
        End If
        If LenB(areaCode) = 0 Then
            localNumber = CheckForAreaCodes(localNumber, areaCode)
        End If
        If LenB(areaCode) = 0 And Len(countryCode) > 4 Then ' could be inside of country code
            areaCode = CheckForCountryCodes(countryCode, countryCode)
        End If
        If LenB(areaCode) = 0 Then
            localNumber = Trim(Tail(localNumber, b, areaCode))
        End If
        If LenB(areaCode) = 0 Then
            localNumber = Trim(Tail(localNumber, "-", areaCode))
        End If
        If Len(areaCode) > 4 Then
            localNumber = CheckForAreaCodes(localNumber, areaCode) _
      & "-" & localNumber
        End If
        
        If LenB(areaCode) = 0 Then               ' no sep from Area to local: test some Matches
            localNumber = CheckForAreaCodes(localNumber, areaCode)
        End If
        If LenB(areaCode) = 0 Then               ' could be inside of country code
            areaCode = CheckForCountryCodes(countryCode, countryCode)
        End If
        If LenB(areaCode) = 0 Then
            If Left(localNumber, 1) = "1" Or Left(localNumber, 1) = "9" Then
                If Left(localNumber, 2) = "18" Then
                    areaCode = Left(localNumber, 4)
                    localNumber = Mid(localNumber, 5)
                Else
                    areaCode = Left(localNumber, 3)
                    localNumber = Mid(localNumber, 4)
                End If
            End If
        End If
        i = InStr(areaCode, "-")
        If i > 1 Then
            localNumber = Mid(areaCode, i + 1) & b & localNumber
            areaCode = Left(areaCode, i - 1)
        End If
        extensionNumber = Trim(Tail(localNumber, "-", localNumber))
    ElseIf Left(ThisNumber, 1) = "(" Or Left(ThisNumber, 1) = "0" Then
        ThisNumber = Mid(ThisNumber, 2)
localCoding:
        localNumber = ThisNumber
        countryCode = "+49"                      ' default country
        GoTo normalCoding
        localNumber = Trim(Tail(areaCode, "-", extensionNumber))
    Else
        If LenB(areaCode) = 0 Then
            areaCode = "6501"                    ' default area
        End If
        countryCode = "+49"
        localNumber = Trim(Tail(ThisNumber, "-", extensionNumber))
    End If
        
    If LenB(localNumber) = 0 Then
        localNumber = extensionNumber
        extensionNumber = vbNullString
    End If
    
    If LenB(countryCode) = 0 Then
        If Left(areaCode, 1) = "+" Then
            countryCode = Left(areaCode, 3)
            areaCode = Mid(areaCode, 4)
        End If
    ElseIf countryCode <> "+1" _
        And Len(countryCode) < 3 _
        And Len(areaCode) > 2 Then
        countryCode = countryCode & Left(areaCode, 2)
        areaCode = Mid(areaCode, 3)
    End If
    CondensedPhoneNumber = Replace(countryCode, "+", "00") _
      & areaCode & localNumber & extensionNumber
    If LenB(countryCode) > 0 Then countryCode = countryCode & b
    If LenB(areaCode) > 0 Then areaCode = "(" & areaCode & ") "
    If LenB(extensionNumber) > 0 Then extensionNumber = "-" & extensionNumber
    
    NormalizeTelefonNumber = countryCode & areaCode & localNumber & extensionNumber
    If startswith3stars Then
        NormalizeTelefonNumber = "***" & NormalizeTelefonNumber
    End If

FuncExit:
    zErr.atFuncResult = CStr(NormalizeTelefonNumber)
    If NormalizeTelefonNumber <> original _
    And Reassign Then
        original = NormalizeTelefonNumber
    End If

ProcReturn:
    Call ProcExit(zErr)

End Function                                     ' O_Goodies.NormalizeTelefonNumber

'---------------------------------------------------------------------------------------
' Method : Function ObjAssignIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ObjAssignIfChanged(Target As Object, Source As Object, modified As Boolean) As Object


    If Source <> Target Then                     ' only valid for Object not WITHIN Objects
        If ChangeAssignReverse Then
            Set ObjAssignIfChanged = Target
        Else
            Set ObjAssignIfChanged = Source
        End If
        modified = True
    End If
    
ProcRet:
End Function                                     ' O_Goodies.AssignIfChanged

'---------------------------------------------------------------------------------------
' Method : Function Pad
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Pad (or cut) to toLen with PadChar
'---------------------------------------------------------------------------------------
Function Pad(str As String, toLen As Long, PadChar As String)
    If Len(PadChar) = 0 Then
        PadChar = b
    ElseIf Len(PadChar) > 1 Then
        PadChar = Left(PadChar, 1)
    End If
    If toLen < Len(str) Then
        Pad = LString(str, toLen)
    Else
        Pad = str & String(toLen - Len(str), PadChar)
    End If
End Function                                     ' O_Goodies.Pad

'---------------------------------------------------------------------------------------
' Method : PartialMatch
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Matches stringlist built as [*]xxxx[*] versus sItem Partial by length of xxxx
'---------------------------------------------------------------------------------------
Function PartialMatch(ErrArray As Variant, sItem As String, Optional ByVal compType As VbCompareMethod = 1) As Boolean
Const zKey As String = "O_Goodies.PartialMatch"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim f As Long
Dim EI As Long
Dim aToken As Variant
Dim sTrunc As String

    ' maybe use Vartype? IsMissing???
    ' if ErrArray is specified as array, just use it. If not: split ListToken
    For Each aToken In ErrArray
        If LenB(aToken) = 0 Then
            GoTo SkipIt
        End If
        If StrComp(sItem, aToken, compType) = 0 Then ' Match complete identity
            PartialMatch = True
            GoTo FuncExit
        End If
        sTrunc = sItem
        ' remove and remember [*]xxxx[*]
        EI = InStrRev(aToken, "*")               ' end of interesting part xxxx
        f = InStr(aToken, "*")
        If EI + f = 0 Then
            GoTo SkipIt
        ElseIf f = EI Then
            If EI > 1 Then                       ' "*" only allowed at start or end of aToken
                f = 0                            ' it only is at the end, none at start
            Else
                If EI = 2 Then
                    DoVerify False, " invalid match pattern (too short)"
                End If
            End If
            If f > 1 Then
                DoVerify False, " " * " not at start x*xx* invalid match pattern"
            End If
        End If
        
        If EI > 1 Then                           ' [*]xxxx*...
            aToken = Left(aToken, EI - 1)        ' xxxx or *xxxx
            sTrunc = Mid(sTrunc, f + 1, EI - 1)  ' restrict compare len
            If StrComp(sTrunc, aToken, compType) = 0 Then ' Match left part identity
                PartialMatch = True              ' xxxx
                GoTo FuncExit
            End If
        End If
        
        If f > 0 Then                            ' *xxxx           start of interesting part
            aToken = Mid(aToken, 2)              ' now is xxxx
            sTrunc = Left(sTrunc, Len(aToken))
            If StrComp(sTrunc, aToken, compType) = 0 Then ' Match right part identity
                PartialMatch = True              ' xxxx
                GoTo FuncExit
            End If
        End If
SkipIt:
    Next aToken
    
FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' O_Goodies.PartialMatch

'---------------------------------------------------------------------------------------
' Method : Function PickDateFromString
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function PickDateFromString(ByVal aDate As String, NoDate As Boolean) As Date
Const zKey As String = "O_Goodies.PickDateFromString"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim ePos As Long
Dim aPos As Long
Dim tPos As Long
Dim MDT As Date
Dim tpart As String
Dim ePart As String

    ePos = 1                                     ' eliminate weekday
    aPos = IsOneOfPos(aDate, split("tag day Mittwoch", b), vbNullString, ePos, " ,")
    If aPos > 0 Then
        aDate = TrimRemove(Mid(aDate, aPos + ePos), " ,") ' trim blank or ,
    End If
    
    aDate = Translate(aDate, " um / at / ab ", b, "/") ' irrelevante worte
    ePos = 1                                     ' seperate date from time
    aPos = IsOneOfPos(aDate, split("am on den dem im", b), b, ePos)
    If aPos > 0 Then
        tpart = Trim(Mid(aDate, aPos, ePos + 10))
        tPos = IsOneOfPos(aDate, split("um at ab", b), b, ePos)
        If tPos > 0 Then                         ' redundant time specification
            tpart = Trunc(1, Mid(aDate, tPos + ePos), b)
            If tPos - 1 > aPos Then
                tpart = Mid(aDate, aPos, tPos - aPos - 1)
            End If
            DoVerify False, " why this ???"
            aDate = Replace(Mid(aDate, tPos + ePos), ".", ":")
            aDate = tpart & b & Trunc(1, aDate, b)
            If IsDate(aDate) Then
                tpart = aDate
            End If
        End If
    Else
        ePos = 1
        aPos = IsOneOfPos(aDate, split("Uhr o'clock"), vbNullString, ePos, " ,")
        If aPos > 0 Then
            tpart = Trim(Left(aDate, aPos - 1))
        Else
            tpart = Trim(Replace(aDate, ", ", b))
        End If
    End If
    
    ePos = InStrRev(tpart, b)                    ' date + time: 12.2. 18:00, seperate
    If ePos > 4 Then                             ' both date and time present in tpart
        ePart = Mid(tpart, ePos + 1)
        If InStr(ePart, ".") > 0 Then            ' correct times 20.00 -> 20:00
            tpart = Trim(Left(tpart, ePos - 1))  ' take apart
            For aPos = 0 To 5                    ' replace odd spelling, max 5 times
                ePart = Replace(ePart, "." & aPos, ":" & aPos)
            Next aPos
            tpart = tpart & b & ePart            ' put back together
        End If
    Else
        ePart = vbNullString                               ' probably no time specified
    End If
        
    If Len(tpart) < 3 Then                       ' no minutes specified, add :00
        tpart = tpart & ":00"
    End If
    If IsDate(tpart) Then
        MDT = CDate(tpart)
        NoDate = False
    Else
daterr:
        If Not NoDate Then                       ' locally report error
            NoDate = True
            rsp = MsgBox("kein Datumshinweis gefunden", vbExclamation)
        End If
    End If
    PickDateFromString = MDT

FuncExit:
    Call ProcExit(zErr)
    
    zErr.atFuncResult = CStr(PickDateFromString)

End Function                                     ' O_Goodies.PickDateFromString

'---------------------------------------------------------------------------------------
' Method : Function Quote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Quote(S As Variant, Optional lr As String) As String
Dim la As String
Dim le As String
    If LenB(lr) = 0 Then
        la = Q
        le = Q
    Else
        la = Left(lr, 1)
        le = Right(lr, 1)
    End If
    Quote = la & S & le                          ' " or ( ) at beginning and end
End Function                                     ' O_Goodies.Quote

'---------------------------------------------------------------------------------------
' Method : Function Quote1
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Quote1(S As String)                     ' like Quote, but single quote used
    Quote1 = QuoteWithDoubleQ(S, "'")            ' quote with ' at beginning and end
End Function                                     ' O_Goodies.Quote1

'---------------------------------------------------------------------------------------
' Method : Function QuoteWithDoubleQ
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function QuoteWithDoubleQ(S As Variant, Optional DQ As String = Q) As String
Dim DiQ As String, Ss As String

    Ss = CStr(S)
    DiQ = DoubleInternalQuotes(Ss, DQ)           ' double quotation marks inside
    QuoteWithDoubleQ = DQ & DiQ & StrReverse(DQ) ' quote with DQ at beginning and end
End Function                                     ' O_Goodies.QuoteWithDoubleQ

'---------------------------------------------------------------------------------------
' Method : RandSequence
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Create array of random integers without duplicates
'---------------------------------------------------------------------------------------
Function RandSequence(anz As Long) As Variant

Const zKey As String = "O_Goodies.RandSequence"
    Call DoCall(zKey, tFunction, eQzMode)

Dim i As Long
Dim RS() As Long
Dim iTemp As Long
Dim iZ As Long

    ReDim RS(1 To anz) As Long
    For i = 1 To anz
        RS(i) = i
    Next i
    For i = anz To 1 Step -1
        Randomize Timer
        iZ = Int((i * Rnd) + 1)
        iTemp = RS(iZ)
        RS(iZ) = RS(i)
        RS(i) = iTemp
    Next i
    RandSequence = RS

zExit:
    Call DoExit(zKey)

End Function                                     ' O_Goodies.RandSequence

'---------------------------------------------------------------------------------------
' Method : Function RCut
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RCut(S As String, Optional i As Long = 1) As String
    RCut = Left(S, Len(S) - i)
End Function                                     ' O_Goodies.RCut

'---------------------------------------------------------------------------------------
' Method : Function RecognizeCode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RecognizeCode(Number As String, Code As String, pattern As String) As Boolean
Const zKey As String = "O_Goodies.RecognizeCode"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="O_Goodies")

Dim L As Long
Dim h As String
Dim N As String
Dim P As Long
    L = Len(pattern)
    N = Number
    If LenB(Code) = 0 Then
        P = InStr(Number, pattern)
        If (Left(Number, 1) = "0" And P = 2) Or P = 1 Then
            Code = pattern
            Number = Trim(Mid(Number, L + P))
            RecognizeCode = True
        End If
    Else
        If Left(Code, L) = pattern Then
            h = Code
            Code = pattern
            Number = Trim(Mid(h, L + 1))
            RecognizeCode = True
        End If
    End If

FuncExit:
    'zErr.atFuncResult = CStr(RecognizeCode)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' O_Goodies.RecognizeCode

'---------------------------------------------------------------------------------------
' Method : Function reduceDouble
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function reduceDouble(N As String, C As String) As String
Dim j As Long
Dim FI As String
Dim c2 As String
Dim i As Long

    c2 = C & C
    FI = Left(N, 2)
    For j = 3 To Len(N)
        i = InStr(j, N, c2)
        If i > 0 Then
            j = i
        Else
            Exit For
        End If
        If FI = C And FI = Mid(N, j, Len(C)) Then
            N = Mid(N, 1, j - 1) & Mid(N, j + Len(C))
        End If
        FI = Mid(N, j, Len(C))
    Next j
    reduceDouble = N
End Function                                     ' O_Goodies.reduceDouble

'---------------------------------------------------------------------------------------
' Method : Function ReFormat
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ReFormat(S As String, replaceChars As String, replaceWith As String, nonRepeating As String) As String
Dim i As Long
Dim M As String
Dim N As String
    N = S
    For i = 1 To Len(replaceChars)               ' replacing with list of chars from replaceWith for all in replaceChars
        M = Mid(replaceChars, i, 1)
        N = Replace(N, M, replaceWith)
    Next i
    N = TrimTail(N, replaceWith)
    N = TrimFront(N, replaceWith)
    N = reduceDouble(N, replaceWith)
    ReFormat = reduceDouble(N, nonRepeating)
End Function                                     ' O_Goodies.ReFormat

'---------------------------------------------------------------------------------------
' Method : Function Remove
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Remove / Replace string using (a list of) characters to remove. Optional replacement by ReplaceWith
'---------------------------------------------------------------------------------------
' if sepToRemove specified, and it is a required seperator (= leading and trailing)
'       then RemoveThis is replaced.
'       If ReplaceWith="", replacing by sepToRemove else by ReplaceWith
' if it is not a required seperator (= not leading), always ReplaceWith is used.
Function Remove(ByVal TargetASource As String, ByVal RemoveThis As String, Optional ByVal sepToRemove, Optional compType As VbCompareMethod, Optional replaceWith As String) As String ' -1 = vbUseCompareOption (literal is undef.)
Dim A As Long
Dim ls As Long
Dim remList As Variant
Static oTaS As String
Static Recursion As Long                     ' special gated proc

    If oTaS <> TargetASource Then
        oTaS = TargetASource
        Recursion = 0
    End If
    
    Recursion = Recursion + 1
    If Recursion = 1 Then
        StringMod = False
    End If
    
    Remove = TargetASource
    If IsMissing(compType) Then
        compType = -1
    End If
    
    If IsMissing(sepToRemove) Then               ' Remove all occurrences ignoring sepToRemove as delimiter
        If LenB(Remove) > 0 And LenB(RemoveThis) > 0 Then ' not required seperator:
            Remove = Replace(TargetASource, RemoveThis, replaceWith, 1, -1, compType)
        End If
        GoTo FuncExit
    End If
    ' it is required seperator:
    If LenB(Remove) = 0 Or LenB(RemoveThis) = 0 Or LenB(sepToRemove) = 0 Then
        GoTo FuncExit                            ' nothing can be removed
    End If
    
    If replaceWith = vbNullString Then
        replaceWith = sepToRemove
    End If
    
    If InStr(RemoveThis, sepToRemove) > 0 Then   ' list of stuff to remove
        remList = split(RemoveThis, sepToRemove)
        For A = 0 To UBound(remList)
            Remove = Replace(Remove, remList(A) & sepToRemove, replaceWith, 1, -1, compType)
        Next A
        GoTo FuncExit
    End If
    
    A = InStr(1, TargetASource, RemoveThis, compType)
    If A = 0 Then                                ' it is not contained at all
        GoTo FuncExit
    End If
    
    ls = Len(sepToRemove)
    If A > ls Then
        If Mid(TargetASource, A - ls, ls) <> sepToRemove Then
            GoTo FuncExit                        ' does not contain sepToRemove in front of RemoveThis: there's a wrong seperator
        End If
    Else
        If A > 1 Then
            If A <= ls Then
                GoTo FuncExit                    ' does not contain sepToRemove only partial match of sepToRemove,
            End If                               ' not starting as first
        Else
            GoTo FuncExit
        End If
    End If
    Remove = Replace(sepToRemove & Remove & sepToRemove, _
                     sepToRemove & RemoveThis & sepToRemove, _
                     replaceWith, 1, -1, compType)
    
    ' Start/End could have several sepToRemoves, ( Trimming )
    While Right(Remove, ls) = sepToRemove
        Remove = Left(Remove, Len(Remove) - ls)
    Wend
    
    While Left(Remove, ls) = sepToRemove
        Remove = Right(Remove, Len(Remove) - ls)
    Wend
    
FuncExit:
    If Remove <> TargetASource Then
        StringMod = True
    End If
    Recursion = Recursion - 1
    If Recursion = 0 Then
        oTaS = vbNullString
    End If
End Function                                     ' O_Goodies.Remove

'---------------------------------------------------------------------------------------
' Method : Function RemoveChars
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RemoveChars(ByVal orig As String, ByVal NoNoChars As String) As String
Dim i As Long

    RemoveChars = orig
    For i = 1 To Len(NoNoChars)
        RemoveChars = Replace(RemoveChars, Mid(NoNoChars, i, 1), vbNullString)
    Next i
End Function                                     ' O_Goodies.RemoveChars

'---------------------------------------------------------------------------------------
' Method : Function RemoveDoubleBlanks
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RemoveDoubleBlanks(S As String) As String
    RemoveDoubleBlanks = S
    While InStr(RemoveDoubleBlanks, B2) > 0
        RemoveDoubleBlanks = Replace(RemoveDoubleBlanks, B2, b)
    Wend
End Function                                     ' O_Goodies.RemoveDoubleBlanks

'---------------------------------------------------------------------------------------
' Method : Function RemoveWord
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RemoveWord(Source As String, pattern As String, Optional split As String) As String
Dim i As Long, lsplit As String, actpattern As String
Dim N As Long, j As Long, Word As String
Dim skiPropTailWild As Boolean

    RemoveWord = Source
    
    If LenB(split) = 0 Then
        lsplit = b
    Else
        lsplit = split
    End If
    
    actpattern = pattern
    If Left(actpattern, 1) = "*" Then
        actpattern = Mid(actpattern, 2)
        skiPropTailWild = True
    End If
    If Right(actpattern, 1) = "*" Then
        actpattern = Left(actpattern, Len(actpattern) - 1)
        skiPropTailWild = True
    End If
    
    For N = 1 To Len(RemoveWord)
        i = InStr(N, RemoveWord, actpattern)
        If i = 0 Then
            GoTo ProcRet                         ' no pattern Match
        Else
            Word = GetThisWord(RemoveWord, i, Len(actpattern), lsplit, j)
            If skiPropTailWild Or j = i Then
                If skiPropTailWild Or i + Len(actpattern) = j + Len(Word) Then
                    RemoveWord = Left(RemoveWord, j - 1) & Mid(RemoveWord, j + Len(Word) + 1)
                    StringsRemoved = StringsRemoved & "entfernt: " & Word & " at pos. " & i & vbCrLf
                End If
            End If
            N = i + Len(actpattern)
        End If
    Next N
ProcRet:
End Function                                     ' O_Goodies.RemoveWord

'---------------------------------------------------------------------------------------
' Method : Repeat
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Repeat a string n times with sep between (but not for last)
'---------------------------------------------------------------------------------------
Function Repeat(str As String, times As Long, Optional sep As String = " ") As String
Dim i As Long
    
    For i = 1 To times - 1
        Repeat = Repeat & str & sep
    Next i
    Repeat = Repeat & str
End Function                                     ' O_Goodies.Repeat

'---------------------------------------------------------------------------------------
' Method : Function ReplaceAll
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ReplaceAll(ByVal S As String, f As String, DS As String, Optional ByVal compType As VbCompareMethod = -1) As String
Dim ach As String

    ReplaceAll = S
    While ach <> ReplaceAll
        ach = ReplaceAll
        If compType = -1 Then
            ReplaceAll = Replace(ReplaceAll, f, DS, 1, -1)
        Else
            ReplaceAll = Replace(ReplaceAll, f, DS, 1, -1, compType)
        End If
    Wend
End Function                                     ' O_Goodies.ReplaceAll

'---------------------------------------------------------------------------------------
' Method : RString
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Righ path to fixed length
'          needed only because there is no Function, only Statement RSet
'---------------------------------------------------------------------------------------
Function RString(x As Variant, L As Long) As String
    RString = String(L, b)
    RSet RString = x
End Function                                     ' O_Goodies.RString

'---------------------------------------------------------------------------------------
' Method : RTail
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Schneide String hinter Sat ab, Suche von hinten. Optional Front Teil vor Sat
'---------------------------------------------------------------------------------------
Function RTail(S As String, Sat As String, Optional Front As String, Optional compType As VbCompareMethod = vbTextCompare) As String
Dim Bpos As Long

    Bpos = InStrRev(S, Sat, -1, compType)
    If Bpos > 0 Then
        RTail = Mid(S, Bpos + Len(Sat))
        Front = Left(S, Bpos - 1)
    Else
        RTail = S
        Front = vbNullString
    End If
End Function                                     ' O_Goodies.RTail

'---------------------------------------------------------------------------------------
' Method : Function RTrimC
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RTrimC(S As String, Optional DQ As String = b) As String
Dim k As Long
Dim lt As Long
    lt = Len(DQ)
    k = Len(S) - lt + 1
    While k > 0 And Mid(S, k, lt) = DQ
        k = k - lt
    Wend
    RTrimC = Mid(S, 1, k)
End Function                                     ' O_Goodies.RTrimC

'---------------------------------------------------------------------------------------
' Method : Sub SetGlobal
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SetGlobal(VarName As String, VarValue As String)

Const zKey As String = "O_Goodies.SetGlobal"
Const MyId As String = "SetGlobal"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="O_Goodies")

Dim SetX As String
Dim ShellRet As Long
    
    SetX = "SETX " & VarName & b & Quote(VarValue) ' do NOT use Quote here ???
    If LogAllErrors Then
        ShellRet = Shell("CMD /K echo Ausführen von " & SetX & " & " & SetX, vbNormalFocus)
    Else
        ShellRet = Shell("CMD /C " & SetX, vbMinimizedFocus)
    End If

FuncExit:
    If VarName = "Test" Then                     ' check for selftest
        TestTail = b & Trim(Mid(Testvar, (InStr(Testvar, "|") + 1)))
        aDebugState = InStr(TestTail, b & MyId) > 0
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' O_Goodies.SetGlobal

'---------------------------------------------------------------------------------------
' Method : Sub ShowAsc
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowAsc(ByVal S As String)
Dim i As Long

    For i = 1 To Len(S)
        Debug.Print i, Asc(Mid(S, i, 1))
    Next i
End Sub                                          ' O_Goodies.ShowAsc

'---------------------------------------------------------------------------------------
' Method : Function SimpleAsciiText
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SimpleAsciiText(aText As String, Optional firstcontrolcharPos As Long, Optional stopimmediately As Boolean) As String
Dim work As String
Dim C As String
Dim i As Long

    firstcontrolcharPos = 0                      ' none found so far
    work = aText
    For i = 1 To Len(work)
        C = Mid(work, i, 1)
        If Asc(C) < 32 Then                      ' replace all control chars by blanks
            If firstcontrolcharPos = 0 Then
                firstcontrolcharPos = i
            End If
            Mid(work, i, 1) = b
            If stopimmediately Then              ' stop after replacing first control char
                Exit For
            End If
        End If
    Next i
    SimpleAsciiText = work
End Function                                     ' O_Goodies.SimpleAsciiText

'---------------------------------------------------------------------------------------
' Method : Function SplitAtWordsWhenCaseChanges
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SplitAtWordsWhenCaseChanges(S As String, Optional modified As Boolean) As String
Dim i As Long
Dim L As Long
Dim caseIsUp As Boolean
Dim ncaseIsUp As Boolean
Dim WordStarted As Boolean                   ' Cap followed by LowerCase is not a new word
Dim work As String
Dim A As String                              ' char at pos i
Dim N As String                              ' char after  i

    modified = False
    L = Len(S)
    If L < 2 Then
        SplitAtWordsWhenCaseChanges = S
        GoTo ProcRet
    End If
    
    A = Left(S, 1)
    caseIsUp = IsUcase(A)
    WordStarted = caseIsUp
    N = Mid(S, i + 1, 1)
    For i = 1 To L - 1
        N = Mid(S, i + 1, 1)
        ncaseIsUp = IsUcase(N)
        If N = b Or A = b Then                   ' already seperate word
            WordStarted = True
        End If
        If ncaseIsUp = caseIsUp Or WordStarted Or WordStarted <> caseIsUp Then
            work = work & A
            WordStarted = False
        Else
            work = work & A & b
            WordStarted = ncaseIsUp
            modified = True
        End If
        caseIsUp = ncaseIsUp
        A = N
    Next i
    SplitAtWordsWhenCaseChanges = work & A
    
ProcRet:
End Function                                     ' O_Goodies.SplitAtWordsWhenCaseChanges

' split string into non-blank words as variant array of strings
Function SplitNBWords(aString As String) As Variant
Dim j As Long, A As Variant, S As Variant
    j = InStr(aString, B2)                       ' find double blanks
    While j > 0
        aString = Replace(aString, B2, b)
        j = InStr(aString, B2)                   ' find double blanks
    Wend
    If LenB(Trim(aString) = 0) Then
        SplitNBWords = vbNullString
        GoTo FuncExit
    End If
    S = split(aString, b)
    j = 0
    For Each A In S
        A = Trim(A)
        If LenB(A) > 0 Then
            S(j) = A
            j = j + 1
        End If
    Next A
    ReDim Preserve S(j - 1)
    SplitNBWords = S

FuncExit:

End Function                                     ' O_Goodies.SplitNBWords

'---------------------------------------------------------------------------------------
' Method : Function StrAssignIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function StrAssignIfChanged(Target As String, Source As String, modified As Boolean) As String
    If Source <> Target Then                     ' only valid for String not WITHIN Objects
        If ChangeAssignReverse Then
            StrAssignIfChanged = Target
        Else
            StrAssignIfChanged = Source
        End If
        modified = True
    End If

End Function                                     ' O_Goodies.StrAssignIfChanged

'---------------------------------------------------------------------------------------
' Method : Sub StrBetween
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Find the string between two delimiting strings
'---------------------------------------------------------------------------------------
Sub StrBetween(Source As String, startS As String, endS As String, inS As String, Optional KeepDelims As Boolean, Optional StartPos As Long)
Dim A As Long
Dim P As Long

    inS = vbNullString
    A = InStr(Source, startS)
    If A > 0 Then
        P = InStr(A + Len(startS), Source, endS)
        If P > 0 Then
            If KeepDelims Then
                StartPos = A
                inS = Mid(Source, StartPos, P + Len(endS) - 1)
            Else
                StartPos = A + Len(startS)
                inS = Mid(Source, StartPos, P - Len(endS) - 1)
            End If
        End If
    End If

End Sub                                          ' O_Goodies.StrBetween

'---------------------------------------------------------------------------------------
' Method : Sub StringDiff
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub StringDiff(ByVal S1 As String, ByVal S2 As String)

Dim i As Long, lmax As Long, lmin As Long
Dim uline As String
Dim allsame As Boolean

    allsame = True
    S1 = Trim(S1)
    S2 = Trim(S2)
    lmin = Min(Len(S1), Len(S2))
    lmax = Max(Len(S1), Len(S2))
    If lmax = 0 Then
        Debug.Print "both S1 and S2 are null strings"
    Else
        If lmin = 0 Then
            If Len(S1) = 0 Then
                Debug.Print "S1 is null string"
                Debug.Print "S2=" & Quote(S2)
            ElseIf Len(S2) = 0 Then
                Debug.Print "S2 is null string"
                Debug.Print "S1=" & Quote(S1)
                Debug.Print "S2 is null string"
            End If
        Else
            For i = 1 To lmin
                If i > lmax Then
                    allsame = False
                    uline = uline & "-"
                Else
                    If Mid(S1, i, 1) = Mid(S2, i, 1) Then
                        uline = uline & b
                    Else
                        allsame = False
                        uline = uline & "|"
                    End If
                End If
            Next i
            If allsame Then
                Debug.Print "===" & Quote(S1)
            Else
                Debug.Print "^^^ " & uline & String(lmax - lmin, "-")
                Debug.Print "S1=" & Quote(S1)
                Debug.Print "S2=" & Quote(S2)
            End If
        End If
    End If

ProcRet:
End Sub                                          ' O_Goodies.StringDiff

' Insert a text at specified position and return the position at the end of this insertion
Function StringInsert(inToString As String, atpos As Long, ByVal InsertText As String, ByVal addSep As String) As Long
Dim L As Long
Dim paddedInsert As String
    L = Len(inToString)
    paddedInsert = addSep & InsertText & addSep
    If atpos >= L Then
        Call AppendTo(inToString, InsertText, addSep)
    Else
        inToString = Left(inToString, atpos) & paddedInsert & Mid(inToString, atpos + 1)
    End If
    atpos = InStr(inToString, paddedInsert)
    
    StringInsert = atpos + Len(paddedInsert)     ' end of inserted text at this position

End Function                                     ' O_Goodies.StringInsert

'---------------------------------------------------------------------------------------
' Method : Function StringRemove
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function StringRemove(inputString As String, droplist As String, sep As String) As String
Dim D As Variant
Dim S As Variant
Const MyId As String = "StringRemove"

    StringRemove = Trim(inputString)
    If LenB(StringRemove) = 0 Then
        GoTo FuncExit
    End If
    If Right(StringRemove, 1) <> Left(sep, 1) Then
        If Len(sep) = 2 And Right(sep, 1) = b Then
            StringRemove = StringRemove & Left(sep, 1)
        End If
    End If
    D = split(droplist, Left(sep, 1))
    For Each S In D
        S = Trim(S)
        If LenB(S) > 0 Then
            StringRemove = Replace(StringRemove, S & sep, vbNullString)
            If Len(sep) = 2 And Right(sep, 1) = b Then
                StringRemove = Replace(StringRemove, S & Left(sep, 1), b)
            End If
            StringRemove = Trim(StringRemove)
        End If
    Next S
    
FuncExit:
    If (DebugMode Or aDebugState) And ShowFunctionValues Then
        Call ShowFunctionValue(MyId, CStr(StringRemove), False)
    End If
End Function                                     ' O_Goodies.StringRemove

'---------------------------------------------------------------------------------------
' Method : Sub StrReplBetween
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Replace a string between two delimiting strings
'---------------------------------------------------------------------------------------
Sub StrReplBetween(Source As String, startS As String, endS As String, res As String, Optional RemovePart As String, Optional withStartS As Boolean, Optional withEndS As Boolean, Optional withWhatS As String)

    Call StrBetween(Source, startS, endS, RemovePart, KeepDelims:=False)
    If LenB(RemovePart) > 0 Then
        If withStartS Then
            RemovePart = startS & RemovePart
        End If
        If withEndS Then
            RemovePart = RemovePart & endS
        End If
        res = Replace(Source, RemovePart, withWhatS, Count:=1)
    Else
        res = Source
    End If

End Sub                                          ' O_Goodies.StrReplBetween

'---------------------------------------------------------------------------------------
' Method : Sub Swap
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Swap(ByRef A As Variant, ByRef b As Variant, Optional asObject As Boolean)
Dim C As Variant
    
    If asObject Then
        GoTo TreatAsObject
    End If
    
    If VarType(A) = vbObject Then
TreatAsObject:
        Set C = A
        Set A = b
        Set b = C
    Else
        C = A
        A = b
        b = C
    End If
    Set C = Nothing
End Sub                                          ' O_Goodies.Swap

'---------------------------------------------------------------------------------------
' Method : Tail
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Schneide String hinter Sat ab, Suche von vorn. Optional Front Teil vor Sat
'---------------------------------------------------------------------------------------
Function Tail(S As String, Sat As String, Optional Front As String) As String
Dim i As Long, lFront As String
    
    lFront = S
    i = InStr(S, Sat)
    If i > 0 Then
        Tail = Mid(S, i + Len(Sat))
        Front = Left(S, i - 1)
    Else                                         ' all is tail, and Front empty if chars not found
        Tail = S
        Front = vbNullString                               ' do not duplicate
    End If
    
    If DebugMode Then
        If ShowFunctionValues Then
            If LenB(S) > 0 Then
                Debug.Print "Tail='" & Tail & "'", "Front='" & Front & "'", "Sat='" & Sat & "'"
            End If
        End If
    End If
    
FuncExit:
    If (DebugMode Or aDebugState) And ShowFunctionValues Then
        Debug.Print "Tail='" & CStr(Tail) & "'", "Front='" & Front & "'", "Sat='" & Sat & "'"
    End If
End Function                                     ' O_Goodies.Tail

'---------------------------------------------------------------------------------------
' Method : Function TextEdit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TextEdit(Text As String, Optional ReplCrLf As String = "|") As String
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "O_Goodies.TextEdit"
Dim zErr As cErr
    
    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction) ' or drop for ) ' Z_Type

    Call frmLongText.TextEdit(Text, ReplCrLf)

ProcReturn:
    Call ProcExit(zErr)
End Function                                     ' O_Goodies.TextEdit

'---------------------------------------------------------------------------------------
' Method : TimerNow
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get the present time using time and exact timer
'---------------------------------------------------------------------------------------
Function TimerNow() As String
Dim nw As Date
Dim TT As Single
Dim nwl As Single

Const spd As Long = 86400
    
    TT = Timer
    nw = Time
    nwl = CSng(nw) * spd

    TimerNow = nw & "," & Mid(TT - nwl, 3)

End Function                                     ' O_Goodies.TimerNow

'---------------------------------------------------------------------------------------
' Method : Function Trail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Trail(S As String, at As String) As String
Dim Bpos As Long

    Bpos = InStrRev(S, at)
    If Bpos > 0 Then
        Trail = Mid(S, Bpos + Len(at))
    Else
        Trail = S
    End If
End Function                                     ' O_Goodies.Trail

' what / tothis can be single string or array of strings
' if single string, Array is generated using splitter
' splitter="" means what is no array, but atothis is
' splitter="$$1$$x" selects x only for what
' splitter="$$2$$x" selects x only for atothis
Function Translate(S As String, what As Variant, tothis As Variant, Optional splitter As String = b, Optional compType As VbCompareMethod = vbBinaryCompare) As String
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "O_Goodies.Translate"
Dim zErr As cErr

Dim aWhat As Variant
Dim aToThis As Variant
Dim i As Long
Dim uT As Long
Dim ms1 As String
Dim ms2 As String
Dim A As String
Dim DS As String

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tFunction)
    
    ms1 = Replace(splitter, "$$1$$", vbNullString)
    ms2 = RTail(splitter, "$$2$$")
    ms1 = Trunc(1, ms1, "$$")
    aWhat = GenArr(what, ms1)
    aToThis = GenArr(tothis, ms2)
    uT = UBound(aToThis)
    Translate = S
    For i = 0 To UBound(aWhat)
        A = aWhat(i)
        If i > uT Then                           ' fewer replacements than replacers
            DS = aToThis(uT)
        Else
            DS = aToThis(i)
        End If
        Translate = ReplaceAll(Translate, A, DS, compType)
    Next i

FuncExit:
    zErr.atFuncResult = CStr(Translate)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' O_Goodies.Translate

'---------------------------------------------------------------------------------------
' Method : Function TrimFront
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TrimFront(N As String, C As String) As String
Dim j As Long

    TrimFront = N
    j = InStrRev(N, C)
    If j > Len(C) Then                           ' if c used like a string quote, remove that
        If Right(N, Len(C)) = Left(N, Len(C)) Then ' front == tail
            TrimFront = Mid(N, j + Len(C), Len(N) - 2 * Len(C)) ' remove c from front and tail completely
        End If
    End If
    
    j = InStr(TrimFront, C)                      ' cut off the part in front of c
    While j > 0                                  '
        TrimFront = Mid(TrimFront, j + Len(C))
        j = InStr(TrimFront, C)                  ' repeat if it contains several c
    Wend
    
ProcRet:
End Function                                     ' O_Goodies.TrimFront

'---------------------------------------------------------------------------------------
' Method : Function TrimNul
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TrimNul(Item As String) As String
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
   
Dim iPos As Long                             ' Nul markiert String-Ende

    iPos = InStr(Item, vbNullChar)
    TrimNul = IIf(iPos > 0, Left$(Item, iPos - 1), Item)

End Function                                     ' O_Goodies.TrimNul

'---------------------------------------------------------------------------------------
' Method : Function TrimRemove
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TrimRemove(N As String, Trimlist As Variant) As String
Dim thisTrimArray As Variant
Dim trimalso As Boolean
Dim L As Variant
    
    If isArray(Trimlist) Then
        thisTrimArray = Trimlist
    Else
        thisTrimArray = split(Trimlist, b)
    End If
    TrimRemove = N
    For Each L In thisTrimArray
        If LenB(L) = 0 Then
            trimalso = True
        Else
            While Left(TrimRemove, Len(L)) = L
                TrimRemove = Mid(TrimRemove, Len(L) + 1)
            Wend
            While Right(TrimRemove, Len(L)) = L
                TrimRemove = Mid(TrimRemove, 1, Len(TrimRemove) - Len(L))
            Wend
        End If
        If trimalso Then
            TrimRemove = Trim(TrimRemove)
        End If
    Next L
End Function                                     ' O_Goodies.TrimRemove

' remove all chars following last c
Function TrimTail(N As String, C As String) As String

Dim j As Long
    j = InStrRev(N, C)
    If j > 0 Then
        If j > 1 Then
            TrimTail = Left(N, j - 1)
        Else
            TrimTail = vbNullString
        End If
    Else
        TrimTail = N
    End If
    
FuncExit:
    If (DebugMode Or aDebugState) And ShowFunctionValues Then
        Call ShowFunctionValue("TrimTail", TrimTail, False)
    End If
End Function                                     ' O_Goodies.TrimTail

'---------------------------------------------------------------------------------------
' Method : Function Trunc
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Trunc(StartPos As Long, inS As String, bS As String, Optional Tail As Variant, Optional compType As VbCompareMethod = vbTextCompare, Optional Bpos As Long) As String
    If StartPos <= 0 Then
        StartPos = 1
    End If
    Bpos = InStr(StartPos, inS, bS, compType)
    If Bpos > 0 Then
        Trunc = Mid(inS, StartPos, Bpos - StartPos)
        Tail = Mid(inS, Bpos + Len(bS))
    Else
        Trunc = Mid(inS, StartPos)
        Tail = vbNullString
    End If
End Function                                     ' O_Goodies.Trunc

'---------------------------------------------------------------------------------------
' Method : Function UnQuote
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function UnQuote(S As String) As String
    UnQuote = Mid(S, 2, Len(S) - 1)
End Function                                     ' O_Goodies.UnQuote

'---------------------------------------------------------------------------------------
' Method : Function VarAssignIfChanged
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function VarAssignIfChanged(Target As Variant, Source As Variant, modified As Boolean) As Variant


    If Source <> Target Then                     ' only valid for Variant not WITHIN Objects
        If ChangeAssignReverse Then
            VarAssignIfChanged = Target
        Else
            VarAssignIfChanged = Source
        End If
        modified = True
    End If

End Function                                     ' O_Goodies.VarAssignIfChanged

'---------------------------------------------------------------------------------------
' Method : Verify
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: delivers first position of valid char (Valid defined by letters in OK)
'          if negate:=True,delivers first position of invalid char
'---------------------------------------------------------------------------------------
Function Verify(S As String, OK As String, Optional negate As Boolean) As Long
Dim i As Long
Dim j As Long

    For i = 1 To Len(S)
        j = InStr(OK, Mid(S, i, 1))
        If j = 0 Then                            ' not in ok
            If Not negate Then
                Exit For
            End If
        Else
            If negate Then
                Exit For
            End If
        End If
    Next i
    If i <= Len(S) Then
        Verify = i                               ' else = 0
    End If
End Function                                     ' O_Goodies.Verify

'---------------------------------------------------------------------------------------
' Method : Function Wait
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Wait(sec As Single, Optional trueStart As Variant, Optional TotalTime As Variant, Optional Retries As Long = 0, Optional Title As String, Optional DebugOutput As Boolean) As Boolean

Const zKey As String = "O_Goodies.Wait"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                    & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True
    Call DoCall(zKey, tFunction, eQzMode)

Dim starttime As Variant
Dim LastInDate As Date
Dim PrevEntrySec As Double
Dim LastInSec As Double
Static thisLong As Variant

    LastInDate = E_Active.atLastInDate        ' remember state before wait
    PrevEntrySec = E_Active.atPrevEntrySec
    LastInSec = E_Active.atLastInSec
    
    If IsMissing(trueStart) Or isEmpty(trueStart) Then
        trueStart = 0
    End If
    If LenB(Title) = 0 Then
        Title = "Waited for dialog to complete for "
    End If
    
    starttime = Timer
    If trueStart = 0 Or Retries = 0 Then
        trueStart = starttime
        thisLong = 0
    End If
    Retries = Retries + 1
    Call Sleep(sec)                              ' wait here with modal box open
    Call ShowStatusUpdate
    thisLong = thisLong + Timer - trueStart
    TotalTime = thisLong
    If DebugOutput And DebugMode And (Retries) Mod 5 = 0 Then
        Debug.Print Title & thisLong & " sec"
    End If
    If DebugMode Or TotalTime > 60# Then
        ' Don'W.xlTSheet wait,  because sleep causes problems in debugmode
        Debug.Print "Please continue debug mode, can't wait for timer in debug mode" & vbCrLf _
                  & "==> assuming wait has completed, so either left button was pressed or Continue/Step (F5/F8)"
        If Not aNonModalForm Is Nothing Then
            ErrStatusFormUsable = True
            Call BugEval
        End If
        Wait = True                              ' End of wait
        ' Debug.Assert False on caller's side recommended
    End If
    
FuncExit:

    E_Active.atLastInDate = LastInDate        ' restore state before Wait
    E_Active.atPrevEntrySec = PrevEntrySec
    E_Active.atLastInSec = LastInSec
    Recursive = False

zExit:
    Call DoExit(zKey, "Waited " & TotalTime)
    
ProcRet:
End Function                                     ' O_Goodies.Wait

'---------------------------------------------------------------------------------------
' Method : Z_GetApplication
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Use Named Application or open if not running yet. Returns Used or Opened
'---------------------------------------------------------------------------------------
Function Z_GetApplication(AppName As String, Optional OpenedHere As Boolean) As Object
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "O_Goodies.Z_GetApplication"
    Call DoCall(zKey, "Sub", eQzMode)

    OpenedHere = False
startOver:
    If Not ShutUpMode Then
        Debug.Print LString(String(OffCal, b) & zKey _
                          & " Started get Application " & AppName & ", Time: ", OffTim) & Format(Timer, " #####.00") _
            
    End If
    
    Call ErrEx.Disable                           ' Must disable, it would set Exception Options
    If ErrStatusFormUsable Then
        frmErrStatus.fErrAppl = vbNullString
        Call BugEval
    End If
    
    aBugTxt = "Get Excel Object"                 ' Extras\Options\General\Stop on not handled Errors
    Call Try(testAll)
    Set Z_GetApplication = GetObject(, AppName & ".Application")
    If Z_GetApplication Is Nothing Then          ' returns error 429 if application not running
        If CatchNC Then
            Debug.Print LString(String(OffCal, b) & AppName _
                              & " is not found running at ", OffTim) _
                              & Format(Timer, " #####.00")
            If Err.Number <> 429 Then
                Debug.Print "Error " _
                          & Err.Number & vbCrLf & Err.Description
                Call ErrReset(4)
            End If
        End If
    Else
        If Not ShutUpMode Then
            Debug.Print LString(String(OffCal, b) & zKey _
                              & " found Application " & AppName _
                              & " already running, Time: ", OffTim) _
                              & Format(Timer, " #####.00")
        End If
        Z_GetApplication = False
        GoTo FuncExit
    End If
    
    aBugTxt = "CreateObject(AppName" & Quote(".Application", Bracket)
    Call Try
    Set Z_GetApplication = CreateObject(AppName & ".Application")
    If Catch Then
        Debug.Print LString(String(OffCal, b) & AppName _
                          & " Application not available, Time: ", OffTim) _
                          & Format(Timer, " #####.00") _
                          & vbCrLf & " Error " & Err.Number _
                          & vbCrLf & Err.Description
    Else
        Debug.Print LString(String(OffCal, b) & AppName _
                          & " Application successfully launched, Time: ", OffTim) _
                          & Format(Timer, " #####.00")
        OpenedHere = True
    End If

FuncExit:
    If LenB(UseErrExOn) > 0 Then
        Call ErrEx.Enable(UseErrExOn)
        If ErrStatusFormUsable Then
            frmErrStatus.fErrAppl = AppName
            Call BugEval
        End If
    End If

zExit:
    Call DoExit(zKey)

End Function                                     ' O_Goodies.Z_GetApplication

'---------------------------------------------------------------------------------------
' Method : Function SetOffline
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SetOffline(Optional OnlineMode As OlExchangeConnectionMode = olCachedDisconnected, Optional withProlog As Boolean = False, Optional Exact As Boolean, Optional RepeatLimit As Long = 5) As Boolean
Const zKey As String = "O_Goodies.SetOffline"
Static zErr As New cErr

    If aRDOSession Is Nothing Then
        withProlog = True
        Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")
    End If

    SetOffline = SetOnline(OnlineMode, withProlog, RepeatLimit)

ProcReturn:
    If withProlog Then
        Call ProcExit(zErr)
    End If
    
ProcRet:
End Function                                 ' O_Goodies.SetOffline

'---------------------------------------------------------------------------------------
' Method : Sub SetOnline
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Change Online Status,  ToggleMode is only method working
'---------------------------------------------------------------------------------------
Function SetOnline(Optional RequestStatus As OlExchangeConnectionMode = olCachedConnectedFull, Optional withProlog As Boolean = False, Optional RepeatLimit As Long = 10) As Boolean
Const zKey As String = "O_Goodies.SetOnline"
Static zErr As New cErr

Dim oriStatus As OlExchangeConnectionMode
Dim oriMode As String
Dim actStatus As OlExchangeConnectionMode
Dim actMode As String
Dim reqMode As String

Dim Retries As Long
Dim ModeWanted As String
Dim ModeNow As String

Dim debugthis As Boolean
    debugthis = DebugMode And LogZProcs
                                                
    If aRDOSession Is Nothing Then
        withProlog = True
        Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="O_Goodies")
    End If

    If Not aRDOSession.LoggedOn Then
        Call aRDOSession.Logon(aRDOSession.Profiles(1))
    End If
    
    Call getOnlineStatus(oriStatus, oriMode, ModeNow)
    actMode = oriMode
    actStatus = oriStatus
    
    reqMode = ExStatusToText(RequestStatus, ModeWanted)
    
    If ErrStatusFormUsable Then
        If debugthis Then
            If (actOnlineStatus <> ModeNow And ModeWanted <> ModeNow) Then
                DoVerify InStr(frmErrStatus.fOnline.Caption, ModeNow) > 0, _
                    "*** ErrStatus displays online status=" _
                    & frmErrStatus.fOnline.Caption & ", actually=" & ModeNow _
                    & ", wanting=" & ModeWanted
            End If
        End If
        frmErrStatus.fOnline.Caption = ModeNow
    End If
    actOnlineStatus = ModeNow               ' globally correct(ed) state
    
    Do While InStr(ModeWanted, ModeNow) = 0 And Retries < RepeatLimit
        Retries = Retries + 1
        ' toggle mode because desired state not reached
        Call olApp.ActiveExplorer.CommandBars.ExecuteMso("ToggleOnline")
        DoEvents
        Call Sleep(200)
        
        Call getOnlineStatus(actStatus, actMode, ModeNow)
        If debugthis And InStr(ModeWanted, ModeNow) = 0 Then
            Debug.Print "shit, Wanting " & ModeWanted & "<>" & ModeNow, Retries
            Debug.Assert True
        Else
            Exit Do
        End If
    Loop
    DoVerify Retries < RepeatLimit, "** Unable to reach connection mode " & ModeWanted
    
    SetOnline = oriMode <> actMode              ' indicate changed mode if not bad, else false
    If debugthis Then
        If actStatus = RequestStatus Then       ' no change necessary
            ModeWanted = "exact " & ModeNow
        ElseIf ModeWanted = ModeNow Then
            ModeWanted = "Equivalent " & ModeWanted
        End If
        Call LogEvent("original Connection State=" & oriStatus _
                    & "(" & oriMode _
                    & ")" & vbCrLf & " current State=" & actStatus _
                    & "(" & actMode & ")" & vbCrLf & " Wanted: " _
                    & ModeWanted & "(" & RequestStatus & ") " _
                    & ", tried " & Retries & " times")
    End If
    
    actOnlineStatus = ModeNow
    If ErrStatusFormUsable Then
        If frmErrStatus.fOnline.Caption <> ModeNow Then
            frmErrStatus.fOnline.Caption = ModeNow
        End If
    End If
    
ProcReturn:
    If withProlog Then
        Call ProcExit(zErr, CStr(SetOnline))
    End If

ProcRet:
End Function                                 ' O_Goodies.SetOnline

Sub getOnlineStatus(isStatus As OlExchangeConnectionMode, isMode As String, isNow As String)
    isStatus = aRDOSession.ExchangeConnectionMode
    isMode = ExStatusToText(isStatus, isNow)
End Sub                                     ' O_Goodies.getOnlineStatus

'---------------------------------------------------------------------------------------
' Method : ExStatusToText
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Convert
'---------------------------------------------------------------------------------------
Function ExStatusToText(aStatus As OlExchangeConnectionMode, aState As String) As String
' ----  Trivial Proc ----

    If aStatus >= olCachedConnectedHeaders Then
        aState = "Online"
    Else
        aState = "Offline"
    End If
    ExStatusToText = ExModeNames(aStatus / 100)

End Function                                     ' O_Goodies.ExStatusToText



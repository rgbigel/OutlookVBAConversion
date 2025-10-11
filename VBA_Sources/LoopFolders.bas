Attribute VB_Name = "LoopFolders"
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Sub DeferredActionAdd
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DeferredActionAdd(curObj As Object, curAction As Long, Optional NoChecking As Boolean)
Const zKey As String = "LoopFolders.DeferredActionAdd"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="add Email: " & curObj.Subject)

Dim AO As cActionObject
Dim aSubject As String
Dim aPath As String
Dim aCat As String
Dim LogMsg As String
Dim objTimeType As String
Dim aTime As String
    
    If NoChecking Then
        LogMsg = " preselected from "
    Else
        If Not IsMailLike(curObj) Then
            Call LogEvent("---- Object is not Mail-Like, EntryID=" & curObj.EntryID _
                    & vbCrLf & "   Class=" & curObj.Class & "/" & TypeName(curObj) _
                    & vbCrLf & "   so it will not be processed", eLall)
            GoTo ProcReturn
        End If
        LogMsg = b & LOGGED & " is not specified as Categories in "
    End If

    Set AO = New cActionObject
    aBugTxt = "Set up deferred action"
    Call Try
    AO.aoObjID = curObj.EntryID
    If Catch Then
        GoTo SkipUntilLater
    End If
    AO.ActionID = curAction
    aPindex = 1

    Call GetITMClsModel(curObj, aPindex)                    ' changes aID(aPindex)=aItmDsc, uses or makes aOD
    
    Call aItmDsc.UpdItmClsDetails(curObj)
    
    If NoChecking Then                                      ' already checked and filtered
        aPath = curObj.Parent.FolderPath
        aSubject = Quote(curObj.Subject)
        aCat = curObj.Categories
        GoTo TestLimit
    End If
    
    aSubject = "Item without Subject Property: " & curObj.EntryID ' set default for none there
    aBugTxt = "get Subject for EntryID=" & curObj.EntryID
    Call Try
    aSubject = Quote(curObj.Subject)
    Catch
    aBugTxt = "get folder path=Parent of " & aSubject ' it may not have Subject, we accept that
    Call Try
    aPath = curObj.Parent.FolderPath
    Catch
    aCat = vbNullString
    aBugTxt = "get Categories of " & aSubject
    Call Try
    aCat = curObj.Categories
    Catch
    
    If CurIterationSwitches.ReProcessDontAsk Or CurIterationSwitches.ReprocessLOGGEDItems Then
        GoTo TestLimit
    End If
    If InStr(aCat, LOGGED) = 0 Then                         ' filter not always active ?? so filter again
TestLimit:
        If RestrictedItems Is Nothing Then
            DeferredLimitExceeded = False
        ElseIf RestrictedItems.Count = 0 Then
            DeferredLimitExceeded = False
        Else
            If Deferred.Count >= DeferredLimit - 1 Then
                Call LogEvent(LString("* exceeding the maximum number (" _
                        & RString(DeferredLimit, 3) _
                        & ") of Deferred items on stack", OffObj) _
                        & "remaining items in '" & aPath & "'will be done next time")
                DeferredLimitExceeded = True
            End If
            Deferred.Add AO                                 ' add as deferred action object
            EventHappened = True
        End If
            
        DoVerify LenB(aPath) > 0
        
        aTime = LString(" no time info", 30) & b
        objTimeType = aOD(aPindex).objTimeType
        If LenB(objTimeType) > 0 Then
            If InStr("SentOn Sent LastModificationTime CreationTime", objTimeType) > 0 Then
                aTime = LString(b & LString(objTimeType, 8) & b & aID(aPindex).idTimeValue, 30) & b
            End If
        End If
    
        Call LogEvent(LString("added #" & CStr(Deferred.Count) _
                    & LogMsg & Quote(aPath), 60) & aTime _
                    & LString(aOD(aPindex).objTypeName, 15) _
                    & b & aSubject, eLdebug)
    Else
        Call LogEvent(LString("did not add #" & CStr(Deferred.Count) _
                & " action for " & LOGGED & " already done in " _
                & Quote(aPath), OffAdI) & aSubject, eLmin)
    End If
SkipUntilLater:
    Set AO = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.DeferredActionAdd

'---------------------------------------------------------------------------------------
' Method : Sub DefineLocalEnvironment
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DefineLocalEnvironment()
Dim zErr As cErr
Const zKey As String = "LoopFolders.DefineLocalEnvironment"

'------------------- gated Entry -------------------------------------------------------
    If Not LF_CurActionObj Is Nothing Then          '
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Set LF_CurActionObj = New cActionObject
    Set CurIterationSwitches = New cIterationSwitches
    Set LF_CurActionObj.IterationSettings = CurIterationSwitches

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.DefineLocalEnvironment

'---------------------------------------------------------------------------------------
' Method : Function FindbeginInFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindBeginInFolder(inFolder As String, beginInFolder As Object, noSearchFolders As Boolean) As Folder
Dim zErr As cErr
Const zKey As String = "LoopFolders.FindbeginInFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim curFolder As Folder
Dim RootFolderName As String
    
    inFolder = Replace(inFolder, "/", "\")
    If olApp.ActiveExplorer Is Nothing Then
        Set curFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
    Else
        Set curFolder = olApp.ActiveExplorer.CurrentFolder
    End If
    If beginInFolder Is Nothing Then
        Set beginInFolder = topFolder
    End If
    If beginInFolder.Class = olNamespace Then
        Set beginInFolder = topFolder
    End If
    If curFolder Is Nothing Then
        DoVerify False
    End If
    
    If topFolder Is Nothing Then
        Set topFolder = curFolder
    End If
    If Not topFolder Is Nothing Then
            ' -1 means first is bigger, +1 first is smaller, =0 if same
        If StrComp(inFolder, topFolder.FolderPath, vbTextCompare) = 0 Then ' exact match:
            Set FindBeginInFolder = topFolder     ' ignoring NoSearchFolders, allowed always
            GoTo ProcReturn                       ' no need to search
        End If
        RootFolderName = TrimTail(curFolder.FolderPath, "\")
        If InStr(1, topFolder.FolderPath, "Suchordner", vbTextCompare) _
                > 0 And noSearchFolders Then
            ' already is: Set FindbeginInFolder = Nothing
            Call LogEvent(Quote(RootFolderName & "\" & inFolder) _
                    & " nicht gesucht: topFolder ist Suchordner")
            GoTo ProcReturn                       ' no need to search, returning "Nothing"
        End If
    End If
        
    If Not curFolder Is Nothing Then            ' exact match:
        If StrComp(inFolder, curFolder.FolderPath, vbTextCompare) = 0 Then
            Set FindBeginInFolder = curFolder     ' ignoring NoSearchFolders
            GoTo ProcReturn                       ' no need to search, returning "Nothing"
        End If
    End If
   
    If beginInFolder Is Nothing Then  ' we are looking everywhere
        ' start with topFolder parent object
        Set FindBeginInFolder = FindBeginInFolder(inFolder, topFolder.Parent, noSearchFolders)
    Else
        If beginInFolder.Class = olFolder Then
            Set FindBeginInFolder = beginInFolder
        Else
            Set FindBeginInFolder = Nothing
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' LoopFolders.FindbeginInFolder

'---------------------------------------------------------------------------------------
' Method : Function FolderFromName
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get folder from namespace by name (Case Sensitive)
'---------------------------------------------------------------------------------------
Function FolderFromName(FolderPath As String) As Folder
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "LoopFolders.FolderFromName"
    Call DoCall(zKey, "Function", eQzMode)

Dim FolderName As String
Dim LookInPath As String
Dim Rest As String
Dim i As Long
Dim j As Long
Dim beginInFolder As Object

    If Left(FolderPath, 2) = "\\" Then
        i = 3                               ' initially, skip "\\"
        Set beginInFolder = olApp.GetNamespace("MAPI")  ' = aNameSpace
        LookInPath = "MAPI Namespace"
    Else
        DoVerify False, "FolderFromName requires fully qualified inFolder path"
    End If
    
NextLevel:
    FolderName = Trunc(i, FolderPath, "\", Rest, vbBinaryCompare, j)
    If InStr(FolderName, "Suchordner") > 0 Then
        i = j
        FolderName = Rest
    End If
    aBugTxt = "FolderName " & Quote(FolderName) & " in " & LookInPath
    Call Try(allowNew)
    Set FolderFromName = beginInFolder.Folders(FolderName)
    Catch
    If FolderFromName Is Nothing Then
        GoTo FuncExit
    End If
    
    If j > 0 Then                           ' more parts delimited by \
        LookInPath = Left(Rest, j - 1)
        i = j + 1                           ' start after \ next time
        Set beginInFolder = FolderFromName
        LookInPath = beginInFolder.FolderPath
        GoTo NextLevel
    End If
    
FuncExit:
    Set beginInFolder = Nothing

zExit:
    Call DoExit(zKey)

End Function ' LoopFolders.FolderFromName

'---------------------------------------------------------------------------------------
' Method : Function GetFolderByName
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetFolderByName(inFolder As String, Optional beginInFolder As Object = Nothing, Optional noSearchFolders As Boolean = True, Optional MaxDepth As Long = 0, Optional CaseKnown As Boolean = True) As Folder
Dim zErr As cErr
Const zKey As String = "LoopFolders.GetFolderByName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim Folder As Folder
Dim RootFolderName As String
Dim FolderName As String
Dim startFolder As Folder
Dim workFolderName As String
Dim nFolderName As String
Dim nRootFolderName As String
Dim i As Long

Dim myRecursionDepth As Long                             ' *** try to eliminate this static

    If Left(inFolder, 2) = "\\" Then
        If CaseKnown Then
            Set GetFolderByName = FolderFromName(inFolder)
            If Not GetFolderByName Is Nothing Then
                GoTo ProcReturn
            End If
        End If
    End If
    
    If myRecursionDepth = 0 Then
        StringMod = False
        myRecursionDepth = E_AppErr.atRecursionLvl
    End If
    
    FolderName = inFolder
    If beginInFolder Is Nothing Then
        If Left(inFolder, 2) = "\\" Then
            Set beginInFolder = aNameSpace
            workFolderName = inFolder
            i = 3                               ' initially, skip "\\"
innerFolder:
            i = InStr(i, workFolderName, "\")
            If i = 0 Then
                nFolderName = workFolderName
            Else
                nFolderName = Left(workFolderName, i - 1)
            End If
            For Each Folder In beginInFolder.Folders
                If StrComp(Folder.Name, nFolderName, vbTextCompare) = 0 Then
                    Set GetFolderByName = Folder
                    If StrComp(GetFolderByName.Name, inFolder, vbTextCompare) = 0 Then
                        If i > 0 Then           ' more inner folders
                            i = i + 1           ' skip to next "\"
                            Set beginInFolder = GetFolderByName
                            GoTo innerFolder
                        Else                    ' found final full match
                            GoTo SearchExit ' with message
                        End If
                    End If
                End If
                Set GetFolderByName = Nothing
            Next Folder
            GoTo SearchExit
        Else
            Set beginInFolder = FindBeginInFolder(inFolder, beginInFolder, noSearchFolders)
        End If
    ElseIf MaxDepth = 1 _
        And beginInFolder.Class = olNamespace Then ' all Folders in top level only
        ' do all Folders in NameSpace
        For Each Folder In beginInFolder.Folders
            If StrComp(Folder.Name, FolderName, vbTextCompare) = 0 Then
                Set GetFolderByName = Folder
                GoTo SearchExit ' with message
            End If
        Next Folder
    End If
    If beginInFolder Is Nothing Then  ' unable to find start object
        GoTo ProcReturn
    End If
    myRecursionDepth = myRecursionDepth + 1
    If myRecursionDepth > MaxDepth And MaxDepth <> 0 Then
        GoTo SearchExit
    End If
    If beginInFolder.Class = olFolder Then
        Set startFolder = beginInFolder
    ElseIf beginInFolder.Class = olNamespace Then ' all Folders from top level down
        ' do all Folders in NameSpace
        For Each Folder In beginInFolder.Folders
            Set GetFolderByName = GetFolderByName(inFolder, Folder, noSearchFolders)
            If Not GetFolderByName Is Nothing Then
                GoTo SearchExit ' with message
            End If
        Next Folder
        GoTo SearchExit
    Else
        DoVerify False, " impossible?"
    End If
    ' note: startFolder is a Folder now, can never be Nothing
    ' temporarily remove initial \\ before we search subFolders (added again later)
    RootFolderName = Replace(startFolder.FullFolderPath, "\\", vbNullString)
    nFolderName = RTail(RootFolderName, "\", nRootFolderName)
    nRootFolderName = "\\" & nRootFolderName
    If nRootFolderName = "\\" Then  ' not fully qualified FolderPath
        RootFolderName = startFolder.FullFolderPath ' fix this
        If RootFolderName Like FolderName Then
            workFolderName = RootFolderName
        Else
            If InStr(FolderName, "\\") = 1 Then
                workFolderName = FolderName
            Else
                ' compose a fully qualified name
                workFolderName = RootFolderName & "\" & FolderName
            End If
        End If
    Else
        workFolderName = RootFolderName & "\" & nFolderName
    End If
    Set Folder = startFolder
    If StrComp(Folder.FullFolderPath, workFolderName, vbTextCompare) = 0 Then
        Set GetFolderByName = Folder    ' explict and complete match, use it
        GoTo SearchExit
    Else
        If noSearchFolders Then
            If InStr(1, Folder.FolderPath, "Suchordner", vbTextCompare) _
                    > 0 Then ' search Folders are undesired:
                FolderName = RTail(Folder.FolderPath, "\", RootFolderName)
                Call RTail(RootFolderName, "\", RootFolderName) ' remove SearchFolder part
                ' compose a fully qualified name
                workFolderName = RootFolderName & "\" & FolderName
            End If
        End If
    End If
    If startFolder Is Nothing Then
        For Each Folder In LookupFolders
            Set GetFolderByName = getSubFolderByName(workFolderName, Folder)
            If Not GetFolderByName Is Nothing Then
                GoTo SearchExit
            End If
        Next Folder
    Else
        Set GetFolderByName = getSubFolderByName(workFolderName, startFolder)
    End If
SearchExit:
    If DebugMode Then
        If GetFolderByName Is Nothing Then
            Call LogEvent("In " & Quote(RootFolderName) _
                & " den gesuchten Ordner " & Quote(inFolder) _
                & " nicht gefunden", eLall)
        ElseIf myRecursionDepth <= 2 Then
            myRecursionDepth = 2
            Call LogEvent("Der gesuchte Ordner wurde gefunden: " _
                & GetFolderByName.FolderPath, eLall)
        End If
    End If
RecursionExit:
    myRecursionDepth = myRecursionDepth - 1
    
FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' LoopFolders.GetFolderByName

'---------------------------------------------------------------------------------------
' Method : Function getDefaultFolderType
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function getDefaultFolderType(curObj As Variant) As Folder
Dim zErr As cErr
Const zKey As String = "LoopFolders.getDefaultFolderType"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    'On Error GoTo 0
    If TypeName(curObj) = "Variant()" Then
        Set ActItemObject = curObj(1)
    ElseIf TypeName(curObj) = "Collection" Then
        If curObj.Count = 0 Then
            GoTo gotNone
        End If
        Set ActItemObject = curObj.Item(1)
    ElseIf curObj.Class = olSelection Then
        If curObj.Count = 0 Then
            GoTo gotNone
        End If
        Set ActItemObject = curObj.Item(1)
    ElseIf curObj.Class = olFolder Then
        If curObj.Items.Count = 0 Then
            GoTo gotNone
        End If
        Set ActItemObject = curObj.Items(1)
    Else
gotNone:
        If DebugMode Then DoVerify False, " no item in curobj"
        Set getDefaultFolderType = Nothing
        GoTo ProcReturn
    End If
    If ActItemObject.Parent.Class = olAppointment Then
        Set getDefaultFolderType = ActItemObject.Parent.Parent
    Else
        Set getDefaultFolderType = ActItemObject.Parent
    End If
    Call BestObjProps(getDefaultFolderType, ActItemObject, withValues:=False) ' seek only if not given

FuncExit:
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' LoopFolders.getDefaultFolderType

' InFolder must be a fully qualified FolderPath
Function getSubFolderByName(inFolder As String, ByVal startFolder As Folder, Optional Sfstart As Long = 0) As Folder
Dim zErr As cErr
Const zKey As String = "LoopFolders.getSubFolderByName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim oFolder As Folder
Dim i As Long
Dim SubFolderName As String
Dim RootFolderName As String
    
    'On Error GoTo 0
    SearchFolderNameResult = vbNullString
    SubFolderName = RTail(inFolder, "\", RootFolderName)
    If Not startFolder Is Nothing Then
        If InStr(1, QuoteWithDoubleQ(startFolder, "\"), _
                QuoteWithDoubleQ(SubFolderName, "\"), vbTextCompare) > 0 Then
            Set getSubFolderByName = startFolder
            SearchFolderNameResult = startFolder & " Matches InFolder " & inFolder
            GoTo FuncExit
        End If
        If InStr(1, startFolder.FolderPath, inFolder, vbTextCompare) Then
            Set getSubFolderByName = startFolder
            SearchFolderNameResult = startFolder.FolderPath _
                & " Matches InFolder " & inFolder
            GoTo FuncExit
        End If
    End If
    If InStr(1, inFolder, "Suchordner", vbTextCompare) > 0 Then
isSearchFolder:
        i = 0
        For Each aStore In oStores
            Set oSearchFolders = aStore.GetSearchFolders
            For Each oFolder In oSearchFolders
                i = i + 1
            ' find first SearchFolder if 0, else the one after Sfstart
                If (Sfstart = 0 Or i > Sfstart) _
                And InStr(1, oFolder.FolderPath, inFolder, vbTextCompare) > 0 Then
                    Set getSubFolderByName = oFolder
                    Sfstart = i ' occurrence in matching search Folders
                    SearchFolderNameResult = "Store Search Folder " _
                        & oFolder.FolderPath & " Matches InFolder " & inFolder
                    GoTo FuncExit
                End If
            Next oFolder
        Next aStore
        Set getSubFolderByName = Nothing
        GoTo FuncExit
    End If
        
    If startFolder Is Nothing Then
        DoVerify False, " ???"
    ' could we be looking at a search Folder
        If startFolder.Parent Is Nothing Then   ' multiFolder-table
            Set getSubFolderByName = startFolder
        Else
            For i = LBound(DeferredFolder) To UBound(DeferredFolder)
                Set oFolder = DeferredFolder(i)
                If oFolder Is Nothing Then
                    Exit For
                Else
                    If oFolder.FolderPath = startFolder.FolderPath Then
                        If oFolder.Items.Count > 0 Then
                            Set topFolder = oFolder.Items.Item(1).Parent
                        Else
                            DoVerify False, " what else can we do..."
                        End If
                        Exit For
                    End If
                End If
            Next i
            Set getSubFolderByName = topFolder
        End If
    Else    ' startFolder is specified
        If InStr(1, startFolder.FolderPath, "Suchordner", vbTextCompare) > 0 Then
            GoTo isSearchFolder ' this can never have subFolders, Folders property invalid
        Else
            Set oSearchFolders = startFolder.Folders
        End If
        For Each oFolder In oSearchFolders
            If InStr(1, oFolder.FolderPath, inFolder, vbTextCompare) > 0 Then
                Set getSubFolderByName = oFolder
                SearchFolderNameResult = "Search Folder " _
                        & oFolder.FolderPath & " Matches InFolder " & inFolder
            GoTo FuncExit
            End If
        Next oFolder
        ' not found yet: try recursion
        For Each oFolder In oSearchFolders
            ' not on original level: recurse subFolders
            If oFolder.Folders.Count > 0 Then
                Set getSubFolderByName = getSubFolderByName(inFolder, oFolder)
                If Not getSubFolderByName Is Nothing Then
                    SearchFolderNameResult = "Located " & oFolder.FolderPath _
                        & " walking SubFolders, InFolder " & inFolder
                    GoTo FuncExit
                End If
            End If
        Next oFolder
    End If

FuncExit:
    If LenB(SearchFolderNameResult) = 0 Then
        SearchFolderNameResult = "Found no suitable Folder for " & inFolder
    End If
    If DebugMode Then
        Debug.Print SearchFolderNameResult
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' LoopFolders.getSubFolderByName

' Process (all or selected) items in a Folder
Sub ItemActions(curObj As Object)
Dim zErr As cErr
Const zKey As String = "LoopFolders.ItemActions"
    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

Dim itemNo As Long
Dim NoOfItems As Long
Dim aResult As VbMsgBoxResult
    
    ' Process Folder ITEMS
    ' ====================
    NoOfItems = CountItemsIn(curObj, LF_ItemCount)
    If NoOfItems = 0 And Not curObj Is Nothing Then
        Call LogEvent("         No items left in " & curObj, eLmin)
        Call ErrReset(0)
    End If
    DateSkipCount = 0
    For itemNo = NoOfItems To 1 Step -1
        NoOfItems = CountItemsIn(curObj, LF_ItemCount)
        If NoOfItems < itemNo Then  ' itemNo has to be corrected
            itemNo = NoOfItems      ' some items of Folder have gone (e.g. Processed, or done by rule)
        End If
        If NoOfItems < 1 Then
            Exit For
        End If
        
        Call ShowStatusUpdate
        If DebugLogging Or LF_ItemCount Mod 100 = 0 Then
            If eOnlySelectedItems Then
                Debug.Print "Now processing item no. " & itemNo _
                    & " of " & NoOfItems & " selected items"
            Else
                Debug.Print "Now processing item no. " & itemNo _
                    & " of " & NoOfItems & " in Folder No. " _
                    & LF_DoneFldrCount + 1 _
                    & b & Quote(curObj.FolderPath)
            End If
            Call ShowStatusUpdate
        End If
        LF_ItemCount = LF_ItemCount + 1
Restart:
        If eOnlySelectedItems Then
            aResult = FolderActions(curObj, -itemNo)
            If itemNo <= curObj.Count Then
                curObj.Remove itemNo
                itemNo = itemNo - 1
            End If
        Else
            aResult = FolderActions(curObj, itemNo)
        End If
        If ActionID = atDoppelteItemslöschen Then
            itemNo = NoOfItems ' this stops outer loop, we process all selected directly in action
        End If
        Select Case aResult
            Case vbCancel
                If TerminateRun Then
                    GoTo ProcReturn
                End If
                GoTo ProcReturn
            Case 0      ' user just wanted NOTHING
                GoTo FuncExit
            Case vbOK
            Case vbNo
            Case vbIgnore
            Case vbCancel
                GoTo ProcReturn
            Case Else
                DoVerify False
        End Select
nextOne:
        If DeferredLimitExceeded Then
            Exit For
        End If
    Next itemNo

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.ItemActions

'---------------------------------------------------------------------------------------
' Method : Sub LoopToDoItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub LoopToDoItems() ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "LoopFolders.LoopToDoItems"
Static zErr As New cErr

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & zKey & "Ignored, recursion from " _
                                          & P_Active.DbgId
        End If
        GoTo ProcRet
    End If
    Recursive = True                        ' restored by    Recursive = False ProcRet:

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="LoopToDoItems")

    Call FldActions2Do                      ' (must) have (at least 1) open items

ProcReturn:
    Call ProcExit(zErr)
    Recursive = False

ProcRet:
End Sub ' LoopFolders.LoopToDoItems

'---------------------------------------------------------------------------------------
' Method : Sub LoopFoldersDialog
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub LoopFoldersDialog() ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "LoopFolders.LoopFoldersDialog"
Static zErr As New cErr
Dim ReShowFrmErrStatus As Boolean

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="LoopFolders")

Dim OldStatusOfNowNotLater As Boolean

    IsEntryPoint = True
    E_Active.EventBlock = False                            ' overriding application NoEvent Stop
    If ErrStatusFormUsable Then
        frmErrStatus.fNoEvents = E_Active.EventBlock
        Call BugEval
    End If
    
    ' Statistics inits
    LF_DontAskAgain = False
    LF_ItemCount = 0
    LF_DoneFldrCount = 0
    LF_ItmChgCount = 0
    ' State and Action Presets
    Call DefineLocalEnvironment
    CurIterationSwitches.ReprocessLOGGEDItems = False
    CurIterationSwitches.CategoryConfirmation = False
    
    AllPublic.eActFolderChoice = True
    
    ' Processing Mode Now/Later
    If Not NoEventOnAddItem Then 'save for later
    ' maybe delayed processing present, now to be done?
        OldStatusOfNowNotLater = NoEventOnAddItem
        NoEventOnAddItem = True     ' let's not be interrupted
        If Not StopRecursionNonLogged Then
            Call DoDeferred
        End If
    End If
    
    If frmErrStatus.Visible Then
        Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
        ReShowFrmErrStatus = True
    End If
    Set FRM = New frmFolderLoop
    Call ShowOrHideForm(FRM, ShowIt:=True)
    ActionID = LF_UsrRqAtionId
    aPindex = 1
    Call LoopRecursiveFolders(AllPublic.eActFolderChoice)
    
    ' Processing Mode "Later"
    Call DoDeferred
    rsp = MsgBox("LoopFolder beendet, Anzahl veränderte Objekte:  " _
                & LF_ItmChgCount)
    Call LogEvent("==== Total number of items modified: " _
        & LF_ItmChgCount, eLmin)
    NoEventOnAddItem = OldStatusOfNowNotLater
    If TerminateRun Then
        GoTo FuncExit
    End If
    
FuncExit:
    Set FRM = Nothing
    If ReShowFrmErrStatus Then
        Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.LoopFoldersDialog

'---------------------------------------------------------------------------------------
' Method : Sub LoopRecursiveFolders
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub LoopRecursiveFolders(ChooseTopFolder As Boolean)
Dim zErr As cErr
Const zKey As String = "LoopFolders.LoopRecursiveFolders"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim i As Long
    AskEveryFolder = True
    WantConfirmation = True
    
    If LF_UsrRqAtionId = atFindealleDeferredSuchordner Then    ' all-Folder Operations
        eOnlySelectedItems = False
        eActFolderChoice = False
        ChooseTopFolder = False
        Call LogEvent("Checking all Search Folders for " _
                & SpecialSearchFolderName)
        Call Recurse(topFolder)
        GoTo RecursionExit
    End If
        
    If LF_UsrRqAtionId = atOrdnerinhalteZusammenführen Then    ' 2-Folder Operation
        AskEveryFolder = True
        WantConfirmation = True
        Call Recurse(Nothing)
        Call AddItemDataToOlderFolder
    ElseIf eAllFoldersOfType Then
        If eOnlySelectedItems Then ' selection done, process all items
            If SelectedItems.Count = 0 Then
                Call MakeSelection(1, _
                    "bitte selektierien Sie die gewünschten Objekte im maßgeblichen Ordner ", _
                    "Auswahl selektierter Objekte", "OK", "Cancel")
            End If
        Else
            Call PickAFolder(1, "bitte bestätigen oder wählen Sie den maßgeblichen Ordner ", _
                          "Auswahl des Ordners zu dem gleichartige bestehen", "OK", "Cancel")
            Call getFoldersOfType(Folder(1))
            ' all relevant items are collected from all Folders we found
            Set SelectedItems = New Collection
            For i = 1 To UBound(lFolders)
                Set topFolder = lFolders(i)
                Call Recurse(topFolder)
            Next i
            eOnlySelectedItems = True    ' now work on them
            Call Recurse(SelectedItems)
            ' free space
            Erase lFolders
        End If
    ElseIf ChooseTopFolder And Not eOnlySelectedItems Then
        eOnlySelectedItems = False
        Set SelectedItems = New Collection
        Set topFolder = Nothing
        Call PickAFolder(1, "bitte bestätigen oder wählen Sie den obersten Ordner ", _
                      "Auswahl des Ordners auf der obersten Ebene", "OK", "Cancel")
        Set topFolder = Folder(1)
        LF_DoneFldrCount = 1
        curFolderPath = topFolder.FolderPath
        ' compute topmost Folder above the current one:
        If Left(curFolderPath, 2) = "\\" Then
            FullFolderPath(FolderPathLevel) = "\\" & Trunc(3, curFolderPath, "\")
        Else
            FullFolderPath(FolderPathLevel) = curFolderPath
        End If
        Call LogEvent("==== User has selected Folder " & curFolderPath, eLall)
        Call Recurse(topFolder)
    Else
        If eOnlySelectedItems Then
            eOnlySelectedItems = True
            If ActiveExplorer.Selection.Count < 2 Then
                Call MakeSelection(1, _
                    "bitte selektierien Sie die gewünschten Objekte im maßgeblichen Ordner ", _
                    "Auswahl selektierter Objekte", "OK", "Cancel")
            Else
                Call LogEvent("==== Es wurden " & ActiveExplorer.Selection.Count _
                        & " Objekte vorselektiert.", eLall)
            End If
            Set SelectedItems = New Collection
            Call GetSelectedItems(olApp.ActiveExplorer.Selection)
            Set LF_CurLoopFld = Folder(1)
            If BeforeFolderActions() = vbNo Then
                GoTo ProcReturn    ' recursion makes no sense
            End If
            Call Recurse(SelectedItems)
            GoTo RecursionExit
        Else
            Call LogEvent("==== User requested all Folders to be processed.", eLall)
            For i = 1 To LookupFolders.Count
                Set topFolder = LookupFolders.Item(i)
                curFolderPath = topFolder.FolderPath
                FolderPathLevel = 0
                FullFolderPath(FolderPathLevel) = vbNullString
            
                Call LogEvent("==== now recursing " & Quote(topFolder.FullFolderPath))
                Call Recurse(topFolder)
                If eOnlySelectedItems Then
                    GoTo RecursionExit  ' do not ProcCall incorrect selection mode
                End If
            Next i
        End If
    End If
RecursionExit:
    If Not xlApp Is Nothing Then
        Call EndAllWorkbooks
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.LoopRecursiveFolders

'---------------------------------------------------------------------------------------
' Method : Function NonLoopFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function NonLoopFolder(FolderName As String) As Boolean
Dim zErr As cErr
Const zKey As String = "LoopFolders.NonLoopFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ShortName As String
    ShortName = LCase(Left(FolderName, 4))
    If ShortName = "dele" Then GoTo asNonLoopFolder
    If ShortName = "gelö" Then GoTo asNonLoopFolder
    If ShortName = "tras" Then GoTo asNonLoopFolder
    If ShortName = "junk" Then GoTo asNonLoopFolder
    If ShortName = "spam" Then GoTo asNonLoopFolder
    If ShortName = "uner" Then GoTo asNonLoopFolder
    If ShortName = "dupl" Then GoTo asNonLoopFolder
    NonLoopFolder = False
    isNonLoopFolder = False
    GoTo ProcReturn
asNonLoopFolder:
    NonLoopFolder = True    ' user can revise this
    isNonLoopFolder = True
    If ActionID > 0 Then
        rsp = MsgBox("Aktion " & Quote(ActionTitle(ActionID)) _
                    & vbCrLf & "    ist eventuell sinnlos in " & Quote(FolderName) _
                    & vbCrLf & "Empfehlung: Lass den Quatsch   Ja!" _
                    & vbCrLf & "Trotzdem ausführen:            Nein" _
                    & vbCrLf & "immer ausführen:    Cancel", vbYesNoCancel)
        If rsp = vbNo Then
            NonLoopFolder = False
            LF_DontAskAgain = False
            GoTo ProcReturn
        ElseIf rsp = vbYes Then
            Call LogEvent("<======> skipping item action " _
                & Quote(ActionTitle(ActionID)) _
                & " because it is inFolder " & Quote(FolderName) _
                & " loop item " & WorkIndex(1) _
                & " Time: " & Now(), eLmin)
            LF_DontAskAgain = False
            GoTo ProcReturn  ' pretend normal Folder, returning false
        ElseIf rsp = vbCancel Then
            NonLoopFolder = False
            LF_DontAskAgain = True
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' LoopFolders.NonLoopFolder

'---------------------------------------------------------------------------------------
' Method : Sub Recurse
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Recurse(curObj As Object)
Dim zErr As cErr
Const zKey As String = "LoopFolders.Recurse"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim iFolder As Folder
Dim subFolderCount As Long
Dim recursedFolder As Folder
      
    If curObj Is Nothing Then
        subFolderCount = 0
        Set LF_CurLoopFld = Nothing
    ElseIf eOnlySelectedItems Then
        If SelectedItems.Count = 0 Then
            GoTo ItemActionsDone    ' finished on selection/Collection
        End If
        Set LF_CurLoopFld = SelectedItems.Item(1).Parent
        If LF_CurLoopFld Is Nothing Then
            DoVerify False, " item without Folder as parent"
        End If
        If eAllFoldersOfType Then
            Set LF_CurLoopFld = getDefaultFolderType(curObj)
            Set topFolder = LF_CurLoopFld
            GoTo selOnly
        End If
        Call ItemActions(SelectedItems)
    Else
        If Not curObj.Parent Is Nothing Then    ' Search Folder: no recursion
            If curObj.Parent.Class = olNamespace Then   ' it is a top Folder
                ' top Folders have no DefaultItemType
                MsgBox "Rekursion sinnlos für oberste Ordnerebenen"
                GoTo ProcReturn
            End If
        End If
        Set LF_CurLoopFld = curObj
        If LF_CurLoopFld.Parent Is Nothing Then
            subFolderCount = 0  ' no parent => no kid Folders
        Else                                    ' instance of subfoldrCount remains 0
            subFolderCount = LF_CurLoopFld.Folders.Count ' do recurse
        End If
    End If
    
    If BeforeFolderActions() = vbNo Then
        GoTo ProcReturn    ' recursion makes no sense
    End If
    Set recursedFolder = LF_CurLoopFld  ' remember for recursion exit
    If eAllFoldersOfType Then
        GoTo selOnly    ' no recursion in this case
    End If
    For LF_recursedFldInx = 1 To subFolderCount
    ' Prolog for recursion
    ' ====================
        LF_DoneFldrCount = LF_DoneFldrCount + 1
        Set iFolder = recursedFolder.Folders(LF_recursedFldInx)
        If iFolder = topFolder And FolderPathLevel = 0 Then
            ' no op because this is the selected Folder, level will always be =1
        Else
            curFolderPath = FullFolderPath(FolderPathLevel) _
                            & "\" & iFolder.Name
        End If
        Set Folder(1) = iFolder
        If iFolder.DefaultItemType = olContactItem Then
            If Not iFolder.ShowAsOutlookAB Then
                Call LogEvent("<======> skipping Folder " & iFolder.FolderPath & ": not in Addressbook")
                GoTo noRecurse
            End If
        End If
        If NonLoopFolder(iFolder.Name) Then ' could contain any item type! ??? *** what about Entw ?
skipThis:
            Call LogEvent("<======> skipping Folder " & curFolderPath _
                    & " Time: " & Now(), eLmin)
            GoTo noRecurse
        End If
        
        ' Entry to recursion
        ' ====================
        FolderPathLevel = FolderPathLevel + 1
        FullFolderPath(FolderPathLevel) = curFolderPath
        Call Recurse(iFolder)
        
        ' Epilog of recursion
        ' ====================
        FolderPathLevel = FolderPathLevel - 1
        curFolderPath = FullFolderPath(FolderPathLevel)
noRecurse:
        Set LF_CurLoopFld = recursedFolder  ' restore from recursion
    Next LF_recursedFldInx
selOnly:
    If eOnlySelectedItems Then ' only selection in one Folder, no recursion
        If SelectedItems.Count = 0 Then
            Call LogEvent("Es wurden nichts (mehr) selektiert, Verarbeitung wird beendet.", _
                    eLall)
            If TerminateRun Then
                GoTo ProcReturn
            End If
        Else
            Call LogEvent("     Bearbeitung von " & SelectedItems.Count _
                & " Selektionen in " & Quote(LF_CurLoopFld.FullFolderPath) _
                & " . Beginn: " & Now())
        End If
    ElseIf subFolderCount > 0 Then
        Call LogEvent("=======> All Ordner unterhalb " & curFolderPath)
        Call LogEvent("         aktuell: " _
            & Quote(LF_CurLoopFld.FullFolderPath) _
            & " mit den enthaltenen " & LF_CurLoopFld.Items.Count _
            & " Items. Zeit: " _
            & Now())
    Else
        If Not LF_CurLoopFld Is Nothing Then
            Call LogEvent("=======> keine Unterordner in " _
                & Quote(LF_CurLoopFld.FullFolderPath) & ". Zeit: " & Now())
        End If
    End If
skipRecursion:
    If BeforeItemActions() = vbOK Then ' working on LF_CurLoopFld
        If eOnlySelectedItems Then
            Call ItemActions(SelectedItems)
        Else
            Call ItemActions(curObj)
        End If
    End If
ItemActionsDone:
    Call PostFolderActions

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' LoopFolders.Recurse




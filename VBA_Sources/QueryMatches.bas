Attribute VB_Name = "QueryMatches"
Option Explicit

Private DidFindInits As Boolean ' used to init only once for Entry Points

'---------------------------------------------------------------------------------------
' Method : Sub WasEmailProcessed
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub WasEmailProcessed() ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "QueryMatches.WasEmailProcessed"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="QueryMatches")
    
    Call FindEntryInit("User initiated Event")
    eOnlySelectedItems = True
    Call FirstPrepare                           ' Folder by default (no user interaction)
    Call SelectAndFind

    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.WasEmailProcessed

'---------------------------------------------------------------------------------------
' Method : Sub SelectAndFind
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SelectAndFind()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "QueryMatches.SelectAndFind"
    Call DoCall(zKey, "Sub", eQzMode)

    xReportExcel = True  ' maybe other values?
    Call MatchingItems(MatchMode:=1)

zExit:
    Call DoExit(zKey)

End Sub ' QueryMatches.SelectAndFind

'---------------------------------------------------------------------------------------
' Method : Sub FindEntryInit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FindEntryInit(MyExplanation As String)  ' call this in QueryMatches-related Entry Points
                                            ' and call ReturnEP at Exit there???
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "QueryMatches.FindEntryInit"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:=MyExplanation)

    IsEntryPoint = True                     ' common for several Entry Points
    Call SetEventMode(force:=True)
    
    xUseExcel = False                       ' Defaults for QueryMatches only
    xDeferExcel = False
    xReportExcel = False
    quickChecksOnly = True
    SelectOnlyOne = True
    SelectMulti = False
    
    ActionID = 0
    ActionTitle(0) = "dynamic action SelectAndFind"

ProcReturn:
    Call ProcExit(zErr)

End Sub ' QueryMatches.FindEntryInit

'---------------------------------------------------------------------------------------
' Method : Sub FirstPrepare
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FirstPrepare()
Dim zErr As cErr
Const zKey As String = "QueryMatches.FirstPrepare"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If ChosenTargetFolder Is Nothing Then
        Set ChosenTargetFolder = GetFolderByName("Erhalten", _
                                    beginInFolder:=FolderBackup, _
                                    noSearchFolders:=True)
    End If
    Set Folder(2) = ChosenTargetFolder
    If Folder(1) Is Nothing Then
        Set Folder(1) = ChosenTargetFolder
    End If
    Set LF_CurLoopFld = Folder(1)
    Call Initialize_UI     ' displays options dialogue
    Select Case rsp
    Case vbCancel
        Call LogEvent("=======> Stopped before processing any items . Time: " _
                        & Now(), eLnothing)
        If TerminateRun Then
            GoTo ProcReturn
        End If
        End ' abort
        GoTo ProcReturn
    Case Else   ' loop Candidates
        If topFolder Is Nothing Then
            Set topFolder = LookupFolders.Item(LF_DoneFldrCount)
        End If
        Call FindTrashFolder
    End Select ' rsp (response from InitializeUserID)
    Call InitFindSelect
    DidFindInits = True

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.FirstPrepare

'---------------------------------------------------------------------------------------
' Method : Sub CheckItemProcessed
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CheckItemProcessed(oneItem As Object)           ' used in LoopFolders for each item
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "QueryMatches.CheckItemProcessed"
    Call DoCall(zKey, "Sub", eQzMode)

    DidFindInits = True
    Set ActItemObject = oneItem
    Call SelectAndFind

FuncExit:

zExit:
    Call DoExit(zKey)

End Sub ' QueryMatches.CheckItemProcessed

'---------------------------------------------------------------------------------------
' Method : Sub InitFindModel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitFindModel(Item As Object)
Dim zErr As cErr
Const zKey As String = "QueryMatches.InitFindModel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    TrueCritList = vbNullString                       ' just make consistent, do not change
    eOnlySelectedFolder = False
    FindMatchingItems = True
    Call getCriteriaList
    
    IsComparemode = True                    ' as opposed to delete/doubles
    
    aPindex = 1
    Call GetITMClsModel(Item, aPindex)
    Call aItmDsc.SetDscValues(Item, withValues:=False)
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.InitFindModel

'---------------------------------------------------------------------------------------
' Method : Sub InitFindSelect
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitFindSelect()
Dim zErr As cErr
Const zKey As String = "QueryMatches.InitFindSelect"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call SelectAndCompare(DontDecode:=True)  ' only one, no decode, aID(1).idObjItem
    Call InitFindModel(ActiveExplorerItem(1))   ' use this for New cObjDsc

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.InitFindSelect

'---------------------------------------------------------------------------------------
' Method : Function MatchingItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function MatchingItems(MatchMode As Long) As Long
Dim zErr As cErr
Const zKey As String = "QueryMatches.MatchingItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

'                      = -1: just get the count of items that match
'                      = 0:  delete duplicates except the newest
'                      = 1:  let user decide what to delete
'                      = 2:  Interactive Answer
Dim MandatoryWorkRule As cNameRule
Dim MyMsg As String
    If DidFindInits Then
        DidFindInits = False ' Preselection valid only once
    Else
        Call InitFindSelect
    End If
    Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)
    Call find_Corresponding(ActItemObject, _
                    CritList:=MandatoryWorkRule, _
                    howmany:=MatchingItems, _
                    eliminateID:=False)
                    
    Matches = MatchingItems
    If Matches = 0 Then
        MyMsg = "Es wurden keine übereinstimmende Items"
    ElseIf Matches = 1 Then
        MyMsg = "Es wurde nur ein übereinstimmendes Item"
    Else
        MyMsg = "Es wurden " & Matches & " übereinstimmende Items"
    End If
    MyMsg = MyMsg & " in " & Quote(Folder(2).FolderPath) & " gefunden"
    MyMsg = MyMsg & vbCrLf & "   Kriterien: " _
                & Replace(MandatoryWorkRule.CritRestrictString, " And ", _
                            vbCrLf & vbTab & "And ")
                ' nb: kleines "and" wird nicht ersetzt
    
    If MatchMode = -1 Then
        Call CleanUpRun(False)
        GoTo ProcReturn
    End If
    
    If DebugMode Then
        Debug.Print MyMsg
    End If
    If MatchMode = 0 And MatchingItems < 2 Then ' too few to remove duplicates
        GoTo Finish ' with termination E
    End If
    
    If MatchMode > 0 Then
        If Matches = 0 Then
            Folder(1).Display
            MsgBox MyMsg
            GoTo Finish
        Else
            Folder(2).Display   ' prepare for FilterDisplay
            ActiveExplorer.ClearSelection
            ActiveExplorer.AddToSelection ActItemObject
            If Matches = 1 Then
                Set ActItemObject = RestrictedItemCollection.Item(1)
                rsp = MsgBox(MyMsg _
                    & vbCrLf & "   Yes: das Item wird dargestellt" _
                    & vbCrLf & "   No:  das Item wird selektiert" _
                    & vbCrLf & "   Cancel: nichts tun, weiter ausführen" _
                    , vbYesNoCancel + vbDefaultButton2)
                If rsp = vbYes Then
                    ActItemObject.Display
                    GoTo Finish
                ElseIf rsp = vbNo Then
                    
                    Call FilterDisplay(MandatoryWorkRule.CritFilterString)
                    GoTo Finish
                Else    ' cancel, no abort: it just skips all actions
                    GoTo Finish
                End If
            Else
                Call FilterDisplay(MandatoryWorkRule.CritFilterString)
                GoTo Finish
            End If
        End If
    End If
        
    If xReportExcel Then   ' decode for display in excel not wanted
        aPindex = 1
        Call startReportToExcel
        Call ReportMatchItems(MatchingItems)
        Call endReportToExcel
    Else
        If MatchMode = 0 Then
            For dcCount = 1 To Matches
                DoVerify False
            Next dcCount
            dcCount = Matches
            Call QueryAboutDelete(CStr(dcCount) & " Items matching selection item")
           
        ElseIf MatchMode = 1 Then
            Call PutSelectedItemDataIntoList
        Else
            MsgBox MyMsg
        End If
        ActiveExplorer.ClearSearch
        GoTo Finish
    End If
    
    Call DoTheDeletes
Finish:
    Call CleanUpRun

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.MatchingItems

'---------------------------------------------------------------------------------------
' Method : Sub FilterDisplay
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub FilterDisplay(Restrictions As String)
Dim zErr As cErr
Const zKey As String = "QueryMatches.FilterDisplay"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Filter As String
    Filter = Restrictions
    If InStr(Filter, "[") > 0 Then
        Filter = Replace(Filter, "] = ", ":")
        Filter = Replace(Filter, "] > ", ":>")
        Filter = Replace(Filter, "] < ", ":<")
        Filter = Replace(Filter, "] >= ", ":>=")
        Filter = Replace(Filter, "] <= ", ":<=")
        Filter = Replace(Filter, "'", Q, 1, -1)
        Filter = Replace(Filter, "[", vbNullString)
        
        ' Filter=replace(filter,a,b)  ' for each word not following syntax or translated
        Filter = Replace(Filter, " and ", " A ", 1, -1, vbTextCompare)
        Filter = Replace(Filter, " or ", " OR ", 1, -1, vbTextCompare)
        Filter = Replace(Filter, "subject:", "betreff:", 1, -1, vbTextCompare)
        Filter = Replace(Filter, "sendername:", "von:", 1, -1, vbTextCompare)
        Filter = Replace(Filter, "senton:", "gesendet:", 1, -1, vbTextCompare)
        Filter = Replace(Filter, "received:", "erhalten:", 1, -1, vbTextCompare)
    End If
    If DebugMode Then
        Debug.Print Restrictions
        Debug.Print Filter
    End If
    ActiveExplorer.Search Filter, olSearchScopeCurrentFolder

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.FilterDisplay

'---------------------------------------------------------------------------------------
' Method : Sub PutSelectedItemDataIntoList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PutSelectedItemDataIntoList()
Dim zErr As cErr
Const zKey As String = "QueryMatches.PutSelectedItemDataIntoList"
Dim ReShowFrmErrStatus As Boolean

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
    Call AddItemToList(CStr(Matches), "Gefundene übereinstimmende Items zu " _
                                    & RestrictCriteriaString, vbNullString, vbNullString)
    For i = 1 To Matches
        Call AddItemToList(CStr(i), RestrictedItemCollection(i).Subject, vbNullString, vbNullString)
    Next i
    If DateSkipCount > 0 Then
        ListContent(ListCount).MatchData = "vor dem"
        ListContent(ListCount).DiffsRecognized = CStr(CutOffDate)
    End If
    
    If ListContent.Count > 0 Then
        If frmErrStatus.Visible Then
            Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
            ReShowFrmErrStatus = True
        End If
        Set FRM = New frmDeltaList
        Call ShowOrHideForm(FRM, ShowIt:=True)
        Set FRM = Nothing
    End If
endsub:
    Set ListContent = Nothing

FuncExit:
    If ReShowFrmErrStatus Then
        Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.PutSelectedItemDataIntoList

'---------------------------------------------------------------------------------------
' Method : Function FindSchema
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindSchema(MapiItemType As String, adName As String) As String
Dim zErr As cErr
Const zKey As String = "QueryMatches.FindSchema"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ResCell As Range
Dim ResRow As Long
Dim LastPos As Long
Dim ResType As String
Dim CaseMatch As Boolean
    
    If LenB(adName) = 0 Then
        GoTo FunExit
    End If
    If S Is Nothing Then
        Call OpenAllSchemata
    Else
        S.xlTSheet.Select
    End If
    With S.xlTSheet
        CaseMatch = True
        .Range("A1").Select
trymore:
        Set ResCell = .Cells.Find(what:=adName, _
                After:=xlApp.ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=CaseMatch, SearchFormat:=False)
        If ResCell Is Nothing Then
            If CaseMatch Then
                CaseMatch = False
                GoTo trymore
            End If
NotThere:
            FindSchema = "*" & adName & " noch nicht in " _
                        & MapiItemType & " gefunden*"
            Debug.Print FindSchema
            GoTo FunExit
        End If
        
        If ResCell.Row < LastPos Then
            If ResRow > ResCell.Row Then
                GoTo NotThere
            End If
            
            GoTo incLast
        End If
        
        ResRow = ResCell.Row
        .Range("A" & ResRow).Select
        ResType = LCase(xlApp.Selection.Value)
        
        If ResType <> LCase(MapiItemType) Then
            CaseMatch = True
incLast:
            If LastPos > ResRow Then
                LastPos = LastPos + 1
            Else
                LastPos = ResRow + 1
            End If
            .Range("A" & LastPos).Select
            ResRow = LastPos
            GoTo trymore
        End If
    ' got a match!
    End With ' S.xlTSheet
    FindSchema = xlApp.Selection.Cells(1, 4).Text
FunExit:

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.FindSchema

'---------------------------------------------------------------------------------------
' Method : Function ModObligMatches
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ModObligMatches(ByVal NewObligAttribs As String) As Boolean
Dim zErr As cErr
Const zKey As String = "QueryMatches.ModObligMatches"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If sRules Is Nothing Then
        Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, withMatchBits:=False)
        Set iRules = Nothing
    End If
    sRules.clsObligMatches.ChangeTo = NewObligAttribs
    ' in ChangeTo, .clsObligMatches.RuleMatches signals change
    ModObligMatches = sRules.clsObligMatches.RuleMatches
    If ModObligMatches Then  ' re-check consistency before use
        sRules.RuleInstanceValid = False ' never is before we check
        If Not aTD Is Nothing Then
            ' Call aTD.adRules.AllRulesCopy(InstanceRule, sRules, withMatchBits:=False)
            Call Get_iRules(aTD)
        End If
        SelectedAttributes = vbNullString ' append Similarities later
    Else
        If Not aTD Is Nothing Then
            Call Get_iRules(aTD)
        End If
    End If
    ' the following True... are shortcuts only for speed
    TrueCritList = Trim(iRules.clsObligMatches.aRuleString)
    TrueImportantProperties = iRules.clsObligMatches.MatchesList

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.ModObligMatches

'---------------------------------------------------------------------------------------
' Method : Sub GenCrit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GenCrit(Criteria As cNameRule)
Dim zErr As cErr
Const zKey As String = "QueryMatches.GenCrit"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim j As Long
Dim tDictIndex As Long
Dim gItemProp As ItemProperty
Dim sT As String
Dim PorOP As String ' CritPropName or operator
Dim nextOperator As String
Dim CritRestrictElement As cFilterCriterium
Dim CritFilterElement As cFilterCriterium
Dim UseTimeCompare  As Boolean  ' Time without seconds only, >= instead of =
Dim TimeString As String
Dim eTimeString As String
Dim RawValue As Variant
Dim PropertyIdent As String
Dim TimeEndAdder As Long
Dim ValSep As String
Dim tempstr As String
    
    ' #########################
    j = LBound(Criteria.CleanMatches)
    i = LBound(Criteria.MatchesList)
    TimeEndAdder = 1 ' normally, one minute
    Set SQLpropC = Nothing  ' new criteria
    nextOperator = " And "  ' default op
    
startloop:
    While i <= UBound(Criteria.MatchesList) And j <= UBound(Criteria.CleanMatches)
        Set CritRestrictElement = New cFilterCriterium
        CritRestrictElement.CritType = 1
        
        With CritRestrictElement
            UseTimeCompare = False
            .ValueIsTimeType = UseTimeCompare
            If SQLpropC Is Nothing Then
                Set SQLpropC = New Collection
            End If
            .CritIndex = j + 1
            .CritPropName = Criteria.CleanMatches(j)
            PropertyNameX = .CritPropName
            ' and formatted value
            .adFormattedValue = aID(1).idAttrDict.Item(.CritIndex).adFormattedValue
            
            ' get itemProperty for this attribute
            tDictIndex = CInt(aID(1).idAttrDict.Item(.CritIndex).adtrueIndex)
            Set gItemProp = aID(1).GetAttrDsc4Prop(tDictIndex).adItemProp
            DoVerify aID(1).idAttrDict.Item(.CritIndex).adName = gItemProp.Name
            DoVerify Criteria.CleanMatches(j) = gItemProp.Name
            ' get original raw value as variant if possible
            RawValue = vbNullString
            aBugTxt = "Get Item Property Value"
            Call Try
            RawValue = gItemProp.Value
            Catch
            
            PorOP = Criteria.MatchesList(i)
            sT = Left(PorOP, 1)
            If sT = "!" Then
                ' ! means value to stop building CriteriaString used in FI,
                ' keeping most relevant
                GoTo EndWhile
            End If
            
            If UCase(PorOP) = "OR" Then
                nextOperator = " Or "
                GoTo swallowit
            ElseIf UCase(PorOP) = "NOT" Then
                nextOperator = " Not "
                GoTo swallowit
            ElseIf UCase(PorOP) = "TOETIME" Then
                i = i + 1   ' skip this, use next as parameter
                TimeEndAdder = Criteria.MatchesList(i)
                GoTo swallowit
            ElseIf UCase(PorOP) = "A" Then
                nextOperator = " And "
swallowit:
                GoTo incI
            Else
                nextOperator = " And "
            End If
buildIt:
           If LenB(PorOP) = 0 Then   ' was an operator only, not just prefix
                PorOP = Criteria.MatchesList(i)
                GoTo SkipOpOnly
            End If
            While Right(PorOP, 1) = ")"
                .cBracket = .cBracket & ")"
                PorOP = Mid(PorOP, 1, Len(PorOP) - 1)
                .BracketOpenCount = .BracketOpenCount - 1
            Wend
            While Left(PorOP, 1) = "("
                .oBracket = .oBracket & "("
                PorOP = Mid(PorOP, 2)
                .BracketOpenCount = .BracketOpenCount + 1
            Wend
            
            sT = Left(PorOP, 1)
            Select Case sT  ' first character of porop
                Case "+"
                    .Comparator = " = "
                    .Operator = " And "
                    PorOP = Mid(PorOP, 2)
                    GoTo buildIt
                Case "|"
                    .Comparator = " = "
                    .Operator = " Or "
                    PorOP = Mid(PorOP, 2)
                    GoTo buildIt
                Case "-"
                    .Comparator = " <> "  ' And Not ?? operator = " And Not "
                    PorOP = Mid(PorOP, 2)
                    GoTo buildIt
                Case "%"
                    DoVerify isSQL, "GenCrit Criteria might not work"
                    .Comparator = " Like "
                    isSQL = True
                    PorOP = Mid(PorOP, 2)
                    GoTo buildIt
                Case "~"
                    .Comparator = vbNullString
                    PorOP = Mid(PorOP, 2)
                    StringMod = False
                    tempstr = Append(Trim(sRules.clsSimilarities.aRuleString), PorOP)
                    If StringMod Then
                        sRules.clsSimilarities.ChangeTo = _
                                tempstr _
                                ' ??? clean this assignment
                    End If
                Case Else
                    If LenB(.Comparator) = 0 Then
                        .Comparator = " = "
                    End If
            End Select
            If .CritIndex = 1 Then
                If isSQL Then
                    .Operator = "@SQL="
                Else
                    .Operator = vbNullString
                End If
            Else
                .Operator = nextOperator
                .BracketOpenCount = SQLpropC.Item(.CritIndex - 1).BracketOpenCount
            End If
            DoVerify .BracketOpenCount >= 0, " error: more closing brackets than open ones!"
           ' special for date/time
            If IsDate(.adFormattedValue) Then
                If InStr(.adFormattedValue, ":") > 0 Then ' contains time
                    UseTimeCompare = True
                    .ValueIsTimeType = UseTimeCompare
                End If
            End If
            
            If LenB(.Comparator) > 0 Then   ' any valid condition?
                If isSQL Then
                    .ValueSeperator = "'%"
                Else
                    .ValueSeperator = "'"
                End If
                PropertyIdent = fixPropertyname(PorOP, isSQL)
            End If
            If .ValueIsTimeType Then
                ' NO seconds, check endtime, conditionally UTC base
                TimeString = StandardTime(gItemProp, (RawValue), UTCisUsed)
                TimeString = Format(TimeString, "dd.mm.yyyy hh:nn")
                ' add endtime TimeEndAdder (+= 1 minute as default)
                eTimeString = DateAdd("n", TimeEndAdder, TimeString)
                eTimeString = Format(eTimeString, "dd.mm.yyyy hh:nn")
                If TimeEndAdder < 0 Then
                    Call Swap(TimeString, eTimeString)
                End If
                .adFormattedValue = TimeString
                Call AppendTo(.oBracket, "(", always:=True, ToFront:=True)
                .Comparator = " >= "
                .AttrRawValue = RawValue
                .PropertyIdent = PropertyIdent
                ValSep = .ValueSeperator    ' save for next CritRestrictElement
                .addTo SQLpropC   ' add start time to collection
                Set CritFilterElement = CritRestrictElement.Clone(2)
                CritFilterElement.addTo SQLpropC
                Set CritRestrictElement = New cFilterCriterium
            End If
        End With ' CritRestrictElement
        With CritRestrictElement  ' start over because of change if UseTimeCompare
            If UseTimeCompare Then
                .ValueIsTimeType = True
                .ValueSeperator = ValSep    ' restore from prev CritRestrictElement
                ' create End-Time aspect
                .CritPropName = Criteria.CleanMatches(j)
                .CritIndex = SQLpropC.Count + 1
                .adFormattedValue = eTimeString
                .Comparator = " <= "
                Call AppendTo(.cBracket, ")", always:=True, ToFront:=False)
                ' operator now and
                .Operator = " and "
            End If
            .AttrRawValue = RawValue
            .PropertyIdent = PropertyIdent
            .CritType = 1
            .addTo SQLpropC   ' add it to collection
SkipOpOnly:
            j = j + 1
incI:
            i = i + 1
        End With ' CritRestrictElement
        Set CritFilterElement = CritRestrictElement.Clone(2)
        CritFilterElement.addTo SQLpropC
    Wend    ' criteria loop
EndWhile:
    DoVerify CritRestrictElement.BracketOpenCount = 0, " error: more opening brackets than terms!"

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.GenCrit

' build Match Criteria as string (e.g. for .Restrict) from criteria
Function BuildFindCriteria(Criteria As cNameRule) As String
Dim zErr As cErr
Const zKey As String = "QueryMatches.BuildFindCriteria"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim CritRestrictElement As cFilterCriterium
Dim RestrictCompareTerm As String
    If Criteria Is Nothing Then ' specific call to use RestrictCriteriaString
        BuildFindCriteria = RestrictCriteriaString ' without rule applied
        ' make sure we use intended sRules.clsObligMatches
        If Not sRules Is Nothing Then    ' ! really???
            DoVerify False
            sRules.clsObligMatches.CritRestrictString = RestrictCriteriaString
        End If
    Else
        Call GenCrit(Criteria)
        RestrictCriteriaString = vbNullString
        FilterCriteriaString = vbNullString
            
        For i = 1 To SQLpropC.Count
            Set CritRestrictElement = SQLpropC.Item(i)
            If LenB(CritRestrictElement.Comparator) > 0 Then
                If LenB(CritRestrictElement.PropertyIdent) > 0 _
                And Mid(CritRestrictElement.PropertyIdent, 2, 1) <> "*" Then
                    RestrictCompareTerm = CritRestrictElement.Operator _
                        & CritRestrictElement.oBracket _
                        & CritRestrictElement.PropertyIdent _
                        & CritRestrictElement.Comparator _
                        & QuoteWithDoubleQ(CritRestrictElement.adFormattedValue, _
                                CritRestrictElement.ValueSeperator) _
                        & CritRestrictElement.cBracket ' no term yet
                    If isEmpty(CritRestrictElement.AttrRawValue) Then
                        If DebugMode Then DoVerify False, " value missing in object"
                    Else
                        If CritRestrictElement.CritType = 1 Then
                            RestrictCriteriaString = RestrictCriteriaString _
                                                & RestrictCompareTerm
                        Else
                            FilterCriteriaString = FilterCriteriaString _
                                                & RestrictCompareTerm
                        End If
                    End If
                End If
            End If
        Next i
        Criteria.CritRestrictString = RestrictCriteriaString
        Criteria.CritFilterString = FilterCriteriaString
    End If
    BuildFindCriteria = RestrictCriteriaString

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.BuildFindCriteria

' gets RestrictedItems and adds entries into RestrictedItemCollection
Sub GetRestrictedItems(matchToItemO As Object, FolderToSeekIn As Folder, Criteria As cNameRule, findCount As Long, eliminateID As Variant)
Dim zErr As cErr
Const zKey As String = "QueryMatches.GetRestrictedItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim limCount As Long
Dim A As String
Dim pushFolder As Folder
Dim ReviseTextForm As Object
Dim MandatoryWorkRule As cNameRule
Dim WeRevisedRestrictCriteriaString As Boolean
Dim LocalRestrictString As String
Dim RuleString As String

    If Criteria Is Nothing Then                 ' use RestrictCriteriaString as is
        Call BestObjProps(FolderToSeekIn, matchToItemO, withValues:=False)
        Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)
        ' will not set up new RestrictCriteriaString unless in error
        DoVerify IsMailLike(matchToItemO), _
            "design expects only Mail-like Items, but is " _
            & MandatoryWorkRule.PropAllRules.ARName
        RuleString = matchToItemO.Subject
        If InStr(RuleString, "'") > 0 Then      ' this could contain Quotes, so double these
            RuleString = DoubleInternalQuotes(RuleString, "'")
        End If
        If InStr(RuleString, Q) > 0 Then        ' or even Double quotes "
            RuleString = DoubleInternalQuotes(RuleString, Q)
        End If
        
        LocalRestrictString = "[" & MandatoryWorkRule.aRuleString & "] = '" _
                    & RuleString & "' And " & RestrictCriteriaString
        MandatoryWorkRule.MatchOn = "Working=" & Quote(LocalRestrictString)
    Else
        Set MandatoryWorkRule = Criteria    ' use RestrictCriteriaString from Rule
        RestrictCriteriaString = MandatoryWorkRule.aRuleString
        LocalRestrictString = RestrictCriteriaString
    End If
    Message = vbNullString
    If RestrictedItemCollection.Count > 0 Then
        Set RestrictedItemCollection = New Collection
    End If
    aBugTxt = "Restrict folder " & FolderToSeekIn.FolderPath _
                    & " Using " & LocalRestrictString
    Call Try(allowNew)
    Set RestrictedItems = FolderToSeekIn.Items.Restrict(RestrictCriteriaString)
    If Catch Then
        DoVerify False, "Unable to continue"
    End If
    
    If TestCriteriaEditing <= vbCancel Then ' undefined or OK: get default again
        If TestCriteriaEditing <> vbOK And TestCriteriaEditing <> vbCancel Then
            If DebugLogging And TestCriteriaEditing <> vbOK _
            Or TestCriteriaEditing = vbCancel Then
                TestCriteriaEditing = vbYes
            Else
                TestCriteriaEditing = vbNo
            End If
        End If
    End If
    
tryAgain:
    While Err.Number <> 0 Or TestCriteriaEditing = vbYes
        Set ReviseTextForm = New frmLongText    ' init asks if edit wanted=VbYes / cancel stops
        If ErrorCaught <> 0 Then
            Message = "Suche passende items in " & Quote(FolderToSeekIn.FolderPath) _
                & " ergab Fehler " & Err.Description
        Else
            Message = "Suche passende items in " & Quote(FolderToSeekIn.FolderPath) _
                & " ergab " & RestrictedItems.Count & " Ergebnisse"
        End If
        Call ErrReset(0)
        Call ReviseTextForm.UserMsg(Message)
        Call ReviseTextForm.SetText(RestrictCriteriaString, " And ")
        If Not ReviseTextForm.Visible Then
            ReviseTextForm.Show
        End If
        RestrictCriteriaString = ReviseTextForm.EditedText()
        Call ErrReset(4)
        If TestCriteriaEditing = vbCancel Then  ' that indicates we did not change now and exit the loop
            limCount = 0
        ElseIf TestCriteriaEditing = vbOK Then  ' that indicates we are happy with the selection criteria
        Else
            ' start over with new RestrictCriteriaString: so delete previous
            Set RestrictedItemCollection = New Collection
            ErrReset
            Set RestrictedItems = FolderToSeekIn.Items.Restrict(LocalRestrictString)   ' replace name appendage
            WeRevisedRestrictCriteriaString = True
        End If
    Wend
    Set ReviseTextForm = Nothing
    findCount = Add2RestrictedCol(matchToItemO, eliminateID)
    
    limCount = RestrictedItemCollection.Count
    If eliminateID Then
        A = " andere passenden Items"
        findCount = limCount - 1
    Else
        A = " passende Items"
        findCount = limCount
    End If
    If findCount < 0 Then   ' ==> we did elimiateID the self search already
        A = "kein weiteres passendes Item in " & Quote(matchToItemO.Parent.FolderPath)
        GoTo CheckConditions
    ElseIf limCount < 1 Or (eliminateID And limCount < 2) Then
        Message = A & " nicht gefunden"
    Else
        Message = findCount & A & " gefunden"
    End If
    If findCount > 20 Then
        Call LogEvent(Message)
        rsp = MsgBox(Message & " Alle verwenden?" & vbCrLf _
            & RestrictCriteriaString & vbCrLf & vbCrLf _
            & " (Nein: nur 20 items)", vbYesNoCancel, matchToItemO.Subject)
        If rsp = vbCancel Then
            GoTo cleanup
        ElseIf rsp = vbNo Then
            Debug.Print "verarbeite nur 20 der " & findCount & " items"
            limCount = 20
        Else
            Debug.Print "alle " & limCount & " passenden items werden verwendet!"
        End If
    End If
    
CheckConditions:
    Message = Message & " in " & Quote(FolderToSeekIn.FolderPath) _
            & " Kriterien:" & vbCrLf & vbTab _
            & Replace(LocalRestrictString, " And ", vbCrLf & vbTab & "And ")
    If ErgebnisseAlsListe Then
        Debug.Print Message
        Call RestrictedItemsShow(limCount)
    End If
    
    If findCount > 0 And (limCount = 0 _
    Or (Not eliminateID And findCount < 1)) _
    And (DebugLogging Or ActionID = 0) Then
        Call LogEvent(Message)
        If ActionID = 0 Then
            TestCriteriaEditing = MsgBox(Message & vbCrLf _
                & "Kriterien korrekt?              (Nein: erlaube Änderung)", _
                vbYesNo, "Gesuchtes Objekt nicht vorhanden")
            If TestCriteriaEditing = vbYes Then
                GoTo cleanup
            End If
        End If
        Message = Message & vbCrLf _
            & "Ja: Direktes Verändern der verwendeten Regeln " & vbCrLf _
            & "Nein: Direktes Verändern der generierten Abfrage " & vbCrLf _
            & "Abbruch: Problem ignorieren " & vbCrLf _
            & "     ggf. Zeilenumbruch bei '|' beachten!"
        rsp = MsgBox(Message, vbYesNoCancel, "Alternative Suche")
        If rsp <> vbNo Then
            Call ReviseTextForm.UserMsg(Message)
            Call frmLongText.UserMsg(Message)
            TrueCritList = frmLongText.TextEdit(TrueCritList)
            frmLongText.Show
            If rsp = vbCancel Then  ' use default again next time
                GoTo cleanup
            End If
            Call SplitMandatories(TrueCritList, MandatoryWorkRule)
            RestrictCriteriaString = BuildFindCriteria(MandatoryWorkRule)
            GoTo testCriteria
        ElseIf TestCriteriaEditing = vbYes Then  ' allow user to edit
            RestrictCriteriaString = frmLongText.TextEdit(RestrictCriteriaString)
            frmLongText.Show
            If TestCriteriaEditing = vbCancel Then
                GoTo cleanup
            End If
testCriteria:
            If LenB(RestrictCriteriaString) = 0 Then   ' restore the original RuleString
                TrueCritList = MandatoryWorkRule.aRuleString
                Set pushFolder = LF_CurLoopFld      ' save
                Set LF_CurLoopFld = FolderToSeekIn
                Call Initialize_UI
                If DebugMode Then   ' not expecting Debug.Assert False:
                    DoVerify LF_CurLoopFld.FolderPath = pushFolder.FolderPath
                End If
                Set LF_CurLoopFld = pushFolder      ' restore
                Call BuildFindCriteria(Nothing)  ' special call preserving RestrictCriteriaString
            End If
            Set RestrictedItemCollection = New Collection
            WeRevisedRestrictCriteriaString = True
            GoTo tryAgain
        End If
    Else
        Call LogEvent(Message)
    End If
cleanup:
    TestCriteriaEditing = vbOK  ' default used again next time

FuncExit:
    Set ReviseTextForm = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.GetRestrictedItems

'---------------------------------------------------------------------------------------
' Method : Function Add2RestrictedCol
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Add2RestrictedCol(matchToItemO As Object, eliminateID As Variant) As Long
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "QueryMatches.Add2RestrictedCol"
    Call DoCall(zKey, "Function", eQzMode)

Dim i As Long
Dim aFileSpec As Object

    Add2RestrictedCol = 0
    For i = 1 To RestrictedItems.Count
        Set aFileSpec = RestrictedItems.Item(i)
        
        If eliminateID = True And matchToItemO.EntryID = aFileSpec.EntryID Then ' do not look for identical ones
            Debug.Print vbNullString, "= " & i, aFileSpec.EntryID
            If eliminateID Then
                Call LogEvent(vbTab & "= " & i & vbTab _
                    & "nicht berücksichtigt weil identisch mit Suchobjekt")
            End If
        Else
            If matchToItemO.Subject = aFileSpec.Subject Then ' double check!
                Add2RestrictedCol = Add2RestrictedCol + 1
                RestrictedItemCollection.Add aFileSpec
            Else    ' should never happen: but RestrictedItems could have changed...
                    ' or the subjects are very similar.
                If DebugLogging Then DoVerify False
            End If
        End If
    Next i

FuncExit:
    Set RestrictedItems = Nothing

zExit:
    Call DoExit(zKey)

End Function ' QueryMatches.Add2RestrictedCol

'---------------------------------------------------------------------------------------
' Method : Sub find_Corresponding
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub find_Corresponding(matchToItemO As Object, CritList As cNameRule, howmany As Long, eliminateID As Variant)
Dim zErr As cErr
Const zKey As String = "QueryMatches.find_Corresponding"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim NondefaultMatches As Variant
    If Folder(2) Is Nothing Then
        Call PickAFolder(2, "in diesem Ordner ähnliche suchen", "Wähle Ordner", "OK", "Cancel")
    End If
    Stop ' ???
    Set aID(1).idAttrDict = New Dictionary
    AttributeUndef(1) = 0
    Set aID(2).idAttrDict = New Dictionary
    AttributeUndef(2) = 0
    Set matchToItemO = GetAobj(1, -1)
    objTypName = DecodeObjectClass(getValues:=False)
    NondefaultMatches = TrueImportantProperties
    sRules.ARName = aObjDsc.objTypeName         ' do NOT reinitialize ..ImportantProperties
    Call SetCriteria                            ' using TrueImportantProperties
           
    If quickChecksOnly Then
        AttributeIndex = -2                     ' check and decode all mandatory properties
    Else
        AttributeIndex = -1                     ' check all ( most important ONLY would be -2 )
    End If
    MaxPropertyCount = Max(MaxPropertyCount, aID(1).idAttrDict.Count)
    
    If quickChecksOnly Then
        AttributeIndex = -2  ' next time, do all ???
    Else
        AttributeIndex = -1
    End If
    Call BuildFindCriteria(CritList)
    Call GetRestrictedItems(matchToItemO, Folder(2), CritList, _
            howmany, eliminateID:=eliminateID)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.find_Corresponding

' delivers duplicates in RestrictedItemCollection or RestrictedItems
Sub findUniqueEmailItems(matchToItemO As Object, inFolder As Folder, GetFirstOnly As Variant, Optional howmany As Long, Optional maxTimeDiff As Double = 1, Optional WhichTime As String = "SentOn")
Dim zErr As cErr
Const zKey As String = "QueryMatches.findUniqueEmailItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim findUniqueEmailItem As MailItem
Dim tS As String
Dim tEnd As String
Dim absMtDiff As Double
Dim oDate As Date
Dim timeFilter As String

    Select Case matchToItemO.Class
    Case olMail, olMeetingRequest, olMeetingResponseTentative, olMeetingResponseNegative, olMeetingResponsePositive
    Case olReport, olMeetingCancellation, olTaskRequest
            howmany = 0     ' these "mails" have no Sent on or Received Time, so that's ok
            Set findUniqueEmailItem = Nothing   ' it's not of class MailItem
            GoTo ProcReturn
    Case Else      ' Other cases not expected yet
            DoVerify False
    End Select
    
    If WhichTime = "SentOn" Then
        oDate = matchToItemO.SentOn
    End If
    If oDate = BadDate Then
        oDate = matchToItemO.ReceivedTime
        WhichTime = "ReceivedTime"
    End If
    If LenB(aTimeFilter) = 0 Then
        aTimeFilter = WhichTime
    ElseIf DebugMode Then
        If WhichTime <> aTimeFilter Then
            If DebugLogging Then
                Debug.Print "Unusual comparison: WhichTime <> aTimeFilter " _
                        & WhichTime & " <> " & aTimeFilter
            End If
        End If
    End If
    If maxTimeDiff < 0 Then
        tS = Format(DateAdd("n", maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
        tEnd = Format(DateAdd("n", -maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
        absMtDiff = -maxTimeDiff
    Else
        tS = Format(oDate, "dd.mm.yyyy hh:mm")  ' start +0
        tEnd = Format(DateAdd("n", maxTimeDiff, oDate), "dd.mm.yyyy hh:mm")
        absMtDiff = maxTimeDiff
    End If
                                    ' using >= to avoid a bug (will not find exact match)
    RestrictCriteriaString = "[SenderName] = " & Quote1(matchToItemO.SenderName)
    timeFilter = " And [" & aTimeFilter & "] >= " & Quote1(tS)
    If absMtDiff >= 1# Then  ' range must be at least 1 Minutes (Seconds not available)
        timeFilter = timeFilter & " And [" & aTimeFilter & "] < " & Quote1(tEnd)
    End If
    Call LogEvent("  -- locating similar items in " & Quote(inFolder.FolderPath) _
                        & vbCrLf & "      for " & Quote(matchToItemO.Subject) _
                        & ", original item size=" & matchToItemO.Size _
                        & vbCrLf & LString("       0", 4) _
                        & matchToItemO.EntryID & vbCrLf _
                        & "      at around " & tEnd)
    RestrictCriteriaString = RestrictCriteriaString & timeFilter
    ' Criteria = Nothing means: RestrictCriteriaString is built already
    ' Call BuildFindCriteria(inFolder, Nothing)  ' so we assume this was already done on caller side
    Call GetRestrictedItems(matchToItemO, _
                            FolderToSeekIn:=inFolder, _
                            Criteria:=Nothing, _
                            findCount:=howmany, _
                            eliminateID:=GetFirstOnly)  ' nothing forces use of RestrictCriteriaString
    howmany = FindNonUnique(matchToItemO, oDate, maxTimeDiff)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.findUniqueEmailItems

'---------------------------------------------------------------------------------------
' Method : Function FindNonUnique
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function FindNonUnique(matchToItemO As Object, oDate As Date, mtDiff As Double) As Long
Dim zErr As cErr
Const zKey As String = "QueryMatches.FindNonUnique"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim tS As String
Dim sameID As Boolean
Dim sufficientlySameButDifferentId As Boolean
Dim aDiffTimeInMinutes As Double
Dim afilespecObject As Object
Dim aMtDiff As Double
Dim sign As String
Dim eDate As Date
Dim dDate As Date

    FindNonUnique = RestrictedItemCollection.Count
    If FindNonUnique <= 0 Then
        GoTo ProcReturn
    End If
    If FindNonUnique = 1 Then   ' Only the original was found
        If DebugMode Then
            Call checkDates(matchToItemO, RestrictedItemCollection(1))
        End If
        GoTo ProcReturn
    End If
    If mtDiff < 0 Then
        sign = Chr(177)     ' +-
        aMtDiff = -mtDiff
    Else
        sign = "+"
        aMtDiff = mtDiff
    End If
    If DebugMode Then
        Debug.Print "        we found " & FindNonUnique & " items which match"
    End If
    For i = 1 To FindNonUnique  ' we start at 1 because we know nothing about the sequence in the RestrictedItemCollection
        tS = vbNullString
        If matchToItemO.Class <> RestrictedItemCollection.Item(i).Class Then
            RestrictedItemCollection.Remove i
            FindNonUnique = RestrictedItemCollection.Count
            If i > FindNonUnique Then
                Exit For
            End If
        End If
        Set afilespecObject = RestrictedItemCollection.Item(i)
        If matchToItemO.Subject <> afilespecObject.Subject Then
            tS = " ??? item " & i _
                & " has a different Subject " & Quote(afilespecObject.Subject)
            GoTo someSecondsOff
        End If
        If aTimeFilter = "SentOn" Then
            eDate = afilespecObject.SentOn
        Else
            eDate = afilespecObject.ReceivedTime
        End If
        dDate = DateDiff("n", eDate, oDate)
        sameID = matchToItemO.EntryID = afilespecObject.EntryID
        aDiffTimeInMinutes = Abs(dDate)
        If sameID Then
            If aDiffTimeInMinutes = 0 Then
                tS = ", Exact time "
                GoTo someSecondsOff
            Else    ' can'tS have differrent time?
                DoVerify False
            End If
            GoTo littleDiff
        End If
        If aDiffTimeInMinutes = 0 Then
            sufficientlySameButDifferentId = True
        ElseIf aDiffTimeInMinutes > 0 Then
littleDiff:
            tS = ", time "
            If aDiffTimeInMinutes < 3 Then
                sufficientlySameButDifferentId = True   ' ignore diff in IDs when times are very similar
                tS = tS & "delta< "
            ElseIf aDiffTimeInMinutes > aMtDiff Then
                tS = tS & "off (by "
            Else
                tS = tS & "delta> "
                 sameID = sufficientlySameButDifferentId   ' ignore diff in IDs when times are  similar
            End If
            tS = tS & Format(aDiffTimeInMinutes, "@0") & ":00" _
                            & " min, Tolerance = " _
                            & sign & Format(aMtDiff, "0#") & ":00"
        End If
        ' check the item we know is there, because it is identical by ID
        If matchToItemO.Size <> afilespecObject.Size Then
            On Error GoTo noBodyFormat    ' BodyFormat may not exist for (e.g. SMS)
            If matchToItemO.BodyFormat <> afilespecObject.BodyFormat _
            And afilespecObject.BodyFormat <> olFormatHTML Then
                If matchToItemO.Body <> afilespecObject.Body Then
                    tS = tS & " Size differs by " _
                    & afilespecObject.Size - matchToItemO.Size _
                        & ", size=" & afilespecObject.Size
                End If
noBodyFormat:
                sameID = False  ' can'tS be same ID ?!
                sufficientlySameButDifferentId = False
            Else
                tS = tS & ", BodyFormat in target had been changed to HTML"
            End If
        End If              ' test on equal size
        
        If sameID Or sufficientlySameButDifferentId Then
someSecondsOff:
            FindNonUnique = FindNonUnique - 1   ' correct for true elements
        End If
        If DebugMode Or DebugLogging Then
            Call checkDates(RestrictedItemCollection(1), afilespecObject)
        End If
        If sameID And i > 1 _
        And RestrictedItemCollection(1).EntryID <> afilespecObject.EntryID Then
            If RestrictedItemCollection(1).CreationTime <= afilespecObject.CreationTime Then
                                    ' move the best=oldest match to position 1
                                    ' so we never delete the best match
                                    ' note: now RestrictedItems are not conform regarding index
            ' unfortunately, swapping items in a collection does not work. So we try to delete the younger one instead.
            ' Call Swap(RestrictedItemCollection.Item(i), RestrictedItemCollection.Item(1), asObject:=True)
                RestrictedItemCollection.Item(1).Delete         ' remove from folder
                RestrictedItemCollection.Remove 1               ' remove from collection (shifting down)
                Call LogEvent("Removed RestrictedItemCollection positions 1 because it is younger than that in pos=" & i)
            Else
                RestrictedItemCollection.Item(i).Delete         ' remove from folder
                RestrictedItemCollection.Remove i               ' remove from collection (shifting down)
                Call LogEvent("Removed RestrictedItemCollection pos=" & i & " 1 because it is younger than that in pos=1")
            End If
            i = i - 1                                           ' do not advance loop
        End If
        If FindNonUnique <= 0 Then
            Exit For
        End If
    Next i

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.FindNonUnique

'---------------------------------------------------------------------------------------
' Method : Function fixPropertyname
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function fixPropertyname(PropName As String, sql As Boolean) As String
Dim zErr As cErr
Const zKey As String = "QueryMatches.fixPropertyname"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim msg As String
Dim rsp As VbMsgBoxResult
    If sql Then
        Select Case PropName
        Case "Categories"
            fixPropertyname = "urn:schemas-microsoft-com:office:office#Keywords"
        Case "LastName"
            fixPropertyname = "urn:schemas:contacts:sn"
        Case "FullName"
            fixPropertyname = "urn:schemas:contacts:cn"
        Case "FirstName"
            fixPropertyname = "urn:schemas:contacts:givenName"
        Case "MiddleName"
            fixPropertyname = "urn:schemas:contacts:middlename"
        Case "FileAs"
            fixPropertyname = "urn:schemas:contacts:fileas"
        Case Else
            msg = "Kenne das urn schema (noch) nicht für " & PropName
            
            fixPropertyname = FindSchema(aOD(1).objItemClassName, PropName)
            If Left(fixPropertyname, 1) = "*" Then
                rsp = MsgBox(msg & vbCrLf _
                            & "  Ja: Ignorieren" & vbCrLf _
                            & "  Nein: Editieren in Excel", _
                            vbYesNoCancel, "Abfrage/Vergleich ignorieren?")
                If rsp = vbCancel Then
                    Debug.Print msg
                    DoVerify False
                    End
                ElseIf rsp = vbNo Then
                    fixPropertyname = FindSchema(aOD(1).objItemClassName, _
                                        PropName)
                Else
                    Debug.Print msg
                    Debug.Print "Abfrage/Vergleich von " _
                            & PropName & " wird ignoriert"
                    fixPropertyname = vbNullString
                    GoTo FunExit
                End If
            End If
        End Select
        fixPropertyname = Chr(34) & fixPropertyname & Chr(34)
    Else
        fixPropertyname = "[" & PropName & "]"
    End If
FunExit:

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.fixPropertyname

'---------------------------------------------------------------------------------------
' Method : Sub getCriteriaList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub getCriteriaList()
Dim zErr As cErr
Const zKey As String = "QueryMatches.getCriteriaList"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aItmClass As OlObjectClass
Dim DefaultCritListForClass As String

    ' operators are
    ' + A
    ' - NOT
    ' | OR
    ' ~ relaxed, no Match required
    ' % like
    ' ! ignore this and all following Criteria for FIND
    ' debugMode = True
    DefaultCritListForClass = TrueCritList
    If SelectedItems Is Nothing Then
useSorted:
        aItmClass = sortedItems(1).Item(1).Class
        'SelectedItems.Add sortedItems(1) ???
    Else
        If SelectedItems.Count = 0 Then
            GoTo useSorted
        End If
        aItmClass = SelectedItems.Item(1).Class
    End If
    ' Now, we always have item in "sortedItems"
    If aItmClass = olContactItem Or aItmClass = olContact Then
        ' "%LastName |%FirstName"
        If LenB(DefaultCritListForClass) = 0 _
        And LenB(Trim(sRules.clsObligMatches.aRuleString)) = 0 Then
            DefaultCritListForClass = "LastName %FirstName" ' this is like a parameter, used for test only
        Else
            DefaultCritListForClass = Trim(sRules.clsObligMatches.aRuleString)
        End If
    ElseIf aItmClass = olMail Then
        DefaultCritListForClass = "Subject SenderName SentOn "
    ElseIf aItmClass = olAppointment Then
        DefaultCritListForClass = "Subject Start End IsRecurring Exceptions "
    Else
        If DebugMode Then DoVerify False, " must define value"
    End If
    
    If LenB(TrueCritList) = 0 Then     ' no change, just make consistent
        DefaultCritListForClass = Trim(sRules.clsObligMatches.aRuleString)
        TrueCritList = DefaultCritListForClass
    Else
        Call SplitMandatories(DefaultCritListForClass)
    End If
    
    If InStr(DefaultCritListForClass, "%") > 0 Then
        isSQL = True
    Else
        isSQL = False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.getCriteriaList

'---------------------------------------------------------------------------------------
' Method : Sub OpenAllSchemata
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub OpenAllSchemata()
Dim zErr As cErr
Const zKey As String = "QueryMatches.OpenAllSchemata"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aHdl As String

    aHdl = "ItemType- ADName---------------  AttrNameOhneLeer------------- AttributeAccessString------------------------------------------"
    Set S = xlWBInit(xlA, TemplateFile, "AllSchemata", _
            aHdl, showWorkbook:=DebugMode, mustClear:=False)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.OpenAllSchemata

'---------------------------------------------------------------------------------------
' Method : Sub CleanUpRun
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CleanUpRun(Optional Terminate As Boolean = True)
Dim zErr As cErr
Const zKey As String = "QueryMatches.CleanUpRun"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If Not xlApp Is Nothing Then
        Call ClearWorkSheet(xlA, O)
    End If
    StopRecursionNonLogged = False
    If Terminate Then
        If DebugMode Then DoVerify False
        If TerminateRun Then
            GoTo ProcReturn
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.CleanUpRun

'---------------------------------------------------------------------------------------
' Method : Sub ReportMatchItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ReportMatchItems(reportedMatches As Long)
Dim zErr As cErr
Const zKey As String = "QueryMatches.ReportMatchItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim same As Boolean
Dim matching As Long
Dim Item As Object
    
    aPindex = 1
    Call ItemReportToExcel(1)
    aPindex = 2
    targetIndex = aPindex
    sourceIndex = 0                                 ' do not copy object item properties, only description
    
    For i = 1 To reportedMatches
        workingOnNonspecifiedItem = False
        BaseAndSpecifiedDiffer = False
        Set Item = RestrictedItemCollection.Item(i)
        If Item.Class <> aID(1).idObjItem.Class Then
            Debug.Print "!!!! target Folder contains items of other type than source"
            DoVerify False, " target Folder contains items of other type"
            GoTo skipitem
        End If
        Set aOD(2) = GetITMClsModel(Item, aPindex).idObjDsc
        Call aItmDsc.SetDscValues(Item, withValues:=True, aRules:=sRules)
        
        aOD(0).objDumpMade = 1
        Set aID(2).idAttrDict = New Dictionary      ' object/item has not been decoded
        WorkIndex(1) = 1                            ' our base item ???
        WorkIndex(2) = i
        SelectOnlyOne = False                       ' takes 2 to compare
        fiMain(2) = fiMain(1)                       ' by find/findnext
        same = ItemIdentity()
        If same Then
            matching = matching + 1
            Call ItemReportToExcel(i)
            If DebugMode Then DoVerify False, " testing phase only"
        End If
skipitem:
        Set aID(2).idAttrDict = New Dictionary
        aOD(0).objDumpMade = 1
    Next i
    
    If matching = 0 Then
        If DebugMode Then
            Debug.Print "Found no complete match in " _
                        & reportedMatches - 1 & " relevant matches"
        End If
        Call CleanUpRun                             ' with termination E
    Else
        If DebugMode Then
            Debug.Print "Found " & matching & " relevant matches"
        End If
    End If
    
    FindMatchingItems = False

FuncExit:
    Set Item = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.ReportMatchItems

'---------------------------------------------------------------------------------------
' Method : Function SetCriteria
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SetCriteria() As Boolean
Dim zErr As cErr
Const zKey As String = "QueryMatches.SetCriteria"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)
    
' Out or confirmed values are:
' Globals in AllPublic: TrueImportantProperties (array of strings)
'                       Similarities       (array of cleanMatches, these are appended to:)
'                       SelectedAttributes (in/out: must be vbNullString if we want to set new (incl. Similarities)
'                       TrueCritList       (out as cleaned string)
'                       MandatoryMatches   (== _Rulestring, out, criteria with OPs, not clean,
'                                           Similarities are not included here)
Dim i As Long
Dim sT As String
Dim sr As String
Dim sc As String

    If sRules Is Nothing Then
        If Not aObjDsc Is Nothing Then
            Set sRules = aObjDsc.objClsRules
        End If
        If sRules Is Nothing Then
            DoVerify False, "remove block if no hit ???"
            Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, withMatchBits:=False)
        End If
        Set iRules = Nothing
    End If
    sr = Trim(sRules.clsObligMatches.aRuleString)
    sT = Trim(sRules.clsObligMatches.CleanMatchesString)
    sc = Trim(sRules.clsSimilarities.CleanMatchesString)
    If LenB(sr) > 0 _
    And TrueCritList = sr _
    And SelectedAttributes = sT & b & sc Then   ' alles da, NOP
        GoTo ProcReturn                         ' there is no change for sRules
    End If

    SetCriteria = True                          ' we are changing sRule for this class
    If isEmpty(TrueImportantProperties) Then
        SelectedAttributes = sRules.clsObligMatches.CleanMatchesString
    Else
        ' reconstruct from TrueImportantProperties:
                                                ' rebuild _RuleString, but controlled:
                                                ' from TrueImportantProperties A Similarities (avoiding doubles)
        TrueCritList = vbNullString                       '  rebuilt WITH special chars from both
        SelectedAttributes = vbNullString                 '         ONLY from TrueImportantProperties
                                                ' built SelectedAttributes, without special chars
        For i = LBound(TrueImportantProperties) _
            To UBound(TrueImportantProperties)
            sr = TrueImportantProperties(i)
            sc = sr
shorten:
            sT = Left(sc, 1)
            If LenB(sT) > 0 Then
                If sT < "A" Or sT > "z" Then    ' allow ASCII letters only in front
                    sc = Mid(sc, 2)
                    GoTo shorten
                End If
            End If
            
            Call AppendTo(TrueCritList, _
                                 sr, b, _
                                 always:=False, ToFront:=False) ' special chars Not removed
            ' add only if new in critlist, always if unique in TrueItemProperties
            ' special chars were removed in sr
            Call AppendTo(SelectedAttributes, sc, b, _
                                 always:=False, ToFront:=False) ' special chars are removed
        Next i
        sRules.clsObligMatches.ChangeTo = TrueCritList
    End If
    ' use what's clean already and append Similarities
    If LenB(TrueCritList) > 0 Then                              ' some criteria wanted:
        If LenB(sRules.clsSimilarities.CleanMatchesString) > 0 Then
            ' include the similarities and default matches (without operators) into SelectedAttributes
            Call AppendTo(SelectedAttributes, _
                            sRules.clsSimilarities.CleanMatchesString _
                            & b & RemoveChars(aObjDsc.objDftMatches, "*!%-+|:^()"), _
                            b, False, False)
        End If
    End If
    While InStr(SelectedAttributes, B2) > 0                     ' reduce double blanks
        SelectedAttributes = Replace(SelectedAttributes, B2, b)
    Wend
    sRules.clsSimilarities.CleanMatches = split(SelectedAttributes)
    TotalPropertyCount = UBound(sRules.clsSimilarities.CleanMatches) + 1
    TrueCritList = Trim(TrueCritList)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.SetCriteria

'---------------------------------------------------------------------------------------
' Method : Sub startReportToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub startReportToExcel()
Dim zErr As cErr
Const zKey As String = "QueryMatches.startReportToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim j As Long
Dim col As Long
Dim aText As String
Dim sColList As String
Dim skipAttr As Boolean

    sColList = vbNullString
    If Not xReportExcel Then
        GoTo ProcReturn
    End If
    If aPindex = 1 Then
            cHdl = vbNullString     ' put1intoExcel must match this:
        cHdl = "Item"                    ' 1
        With aID(aPindex)
            j = Min(.idAttrDict.Count, UBound(pArr))
            col = 2
            For i = 1 To j
                skipAttr = False
                aText = .idAttrDict.Keys(i)
                If Left(aText, 2) = "==" Then  ' skip seperators ... and DontCompare
                    skipAttr = True
                Else
                    If InStr(sRules.clsNeverCompare.CleanMatchesString, aText) > 0 Then
                        skipAttr = True
                    End If
                    If InStr(sRules.clsSimilarities.CleanMatchesString, aText) > 0 Then
                        skipAttr = False
                    End If
                    If InStr(sRules.clsObligMatches.CleanMatchesString, aText) > 0 Then
                        skipAttr = False
                    End If
                End If
                If skipAttr Then
                    If DebugMode Then
                        DoVerify False
                    End If
                Else
                    cHdl = cHdl & b & aText
                    sColList = sColList & b & CStr(col)
                    col = col + 1
                End If
            Next i
    
            Set O = xlWBInit(O.xlTWBook, TemplateFile, "Report", _
                            cHdl, showWorkbook:=DebugMode, mustClear:=True)
        End With ' aID(aPindex)
        
        xlApp.ScreenUpdating = False
        O.xlTSheet.EnableCalculation = False  ' for attribute rules do not calculate
        xlApp.Cursor = xlWait
        O.xlTHeadline = cHdl
        O.xlTHead = split(cHdl)
        xColList = split(Trim(sColList))
        If Not aID(0) Is Nothing Then
            aOD(0).objDumpMade = 0                          ' no dumps made so far
        End If
    End If  ' first line

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.startReportToExcel

'---------------------------------------------------------------------------------------
' Method : Sub endReportToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub endReportToExcel()
Dim zErr As cErr
Const zKey As String = "QueryMatches.endReportToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

   ' Note: leaving W again (if we can)
    xlApp.ScreenUpdating = True
    xlApp.Cursor = xlDefault
    If Not O Is Nothing Then
        Set x = O
    End If
    If x Is Nothing Then
        Set W.xlTSheet = Nothing
    Else
        Set W.xlTSheet = x.xlTSheet
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.endReportToExcel

'---------------------------------------------------------------------------------------
' Method : Sub ItemReportToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ItemReportToExcel(Line As Long)
Dim zErr As cErr
Const zKey As String = "QueryMatches.ItemReportToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim aText As String
Dim Finished As Boolean
Dim maxCols As Long
Dim col As Long

    'On Error GoTo 0
    With aID(aPindex).idAttrDict
        pArr(1) = Line
        i = LBound(xColList)
        maxCols = UBound(xColList)
        maxCols = Min(UBound(pArr), maxCols)
        Finished = maxCols = 0
        Do Until Finished  ' get the cAttrDsc items
            If i > maxCols Then
                Finished = True
                GoTo skipitem
            End If
            col = CLng(xColList(i))
            Set aTD = .Item(col - 1).Item
            i = i + 1
            aText = aTD.adFormattedValue
            pArr(col) = aText
            If DebugLogging Or i Mod 10 = 0 Or col >= maxCols Then
                Debug.Print Format(Timer, "0#####.00") & vbTab _
                    & i & b & x.xlTHead(col - 1) & " = " & Quote(aText) _
                    & " into Excel Sheet " & W.xlTName & " column " & col
            End If
skipitem:
        Loop
        Call addLine(x, Line - 1, pArr)
    End With ' aID(aPindex)
    Call DisplayExcel(W, EnableEvents:=False, unconditionallyShow:=True)
    Set aTD = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.ItemReportToExcel

' derived from: http://www.slipstick.com/developer/instant-search-messages-selected-contact/
' Search For Messages From Email or Contact
' Creates an instant search for all messages to or from the sender, including messages you sent.
' Two versions: 1 - Uses selected Contact. 2 - Uses selected Message
' Searches for messages from all three email addresses on a contact, if additional addresses exist.

'---------------------------------------------------------------------------------------
' Method : Sub SearchAssociatedItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SearchAssociatedItems()
Dim zErr As cErr
Const zKey As String = "QueryMatches.SearchAssociatedItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    IsEntryPoint = True
    
    Call SearchByItem(Nothing)      ' Select item from ActiveExplorer

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.SearchAssociatedItems

'---------------------------------------------------------------------------------------
' Method : Sub SearchByItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SearchByItem(oBaseItem As Object, Optional SenderOrRecipient As Long = 3)
Dim zErr As cErr
Const zKey As String = "QueryMatches.SearchByItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim strFilter As String
Dim oFolder As Folder
Dim Operator As String
Dim noOfEmails As Long
    If oBaseItem Is Nothing Then
        Set oBaseItem = ActiveExplorer.Selection.Item(1)
    End If
    
    ' get result window as (new) ActiveExplorer
    If oBaseItem.Parent Is Nothing Then
        Set oFolder = aNameSpace.GetDefaultFolder(olFolderInbox)
        Set olApp.ActiveExplorer.CurrentFolder = oFolder
    Else
        Set oFolder = getParentFolder(oBaseItem)
        If oFolder Is Nothing Then
            DoVerify False, " item has no parent"
        Else
            oFolder.Display ' providing ActiveExplorer of correct type
        End If
    End If
      
    If IsMailLike(oBaseItem) Then
        noOfEmails = noOfEmails + 1
        If InStr(LCase(oFolder.FolderPath), "sen") > 0 Then
            strFilter = "an:" & oBaseItem.SenderEmailAddress
        Else
            strFilter = "von:" & oBaseItem.SenderEmailAddress
        End If
    ElseIf oBaseItem.Class = olContact Then
        If LenB(oBaseItem.Email1Address) > 0 Then
            strFilter = Chr(34) & oBaseItem.Email1Address & Chr(34)
            Operator = " OR "
            noOfEmails = noOfEmails + 1
        End If
        If LenB(oBaseItem.Email2Address) > 0 Then
            strFilter = strFilter & Operator & Chr(34) & oBaseItem.Email2Address & Chr(34)
            Operator = " OR "
            noOfEmails = noOfEmails + 1
        End If
        
        If LenB(oBaseItem.Email3Address) > 0 Then
            strFilter = strFilter & Operator & Chr(34) & oBaseItem.Email3Address & Chr(34)
            ' Operator = " OR "
            noOfEmails = noOfEmails + 1
        End If
        If SenderOrRecipient = 3 Then       ' both
            strFilter = "von:(" & strFilter & ") OR an:(" & strFilter & ")"
        ElseIf SenderOrRecipient = 2 Then   ' only sender
            strFilter = "an:" & strFilter
        ElseIf SenderOrRecipient = 1 Then   ' only recipient
            strFilter = "von:" & strFilter
        End If
Else
        DoVerify False, " not implemented"
    End If
    If noOfEmails = 0 Then
        MsgBox "Keine Email Adresse vorhanden für " & oBaseItem.Subject, vbExclamation
        GoTo ProcReturn
    End If
    ' filter result window
    olApp.ActiveExplorer.Search strFilter, olSearchScopeAllFolders

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.SearchByItem

'---------------------------------------------------------------------------------------
' Method : Sub reportOnSelectedEmail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub reportOnSelectedEmail()
Dim zErr As cErr
Const zKey As String = "QueryMatches.reportOnSelectedEmail"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim msg As String

    IsEntryPoint = True
    
    msg = IsSelectedEmailKnown(3)   ' all 3 are reported
    MsgBox msg, vbExclamation

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' QueryMatches.reportOnSelectedEmail

' get Contacts for email address
Function IsSelectedEmailKnown(TestOnly As Boolean, Optional Contacts As Collection) As String
Dim zErr As cErr
Const zKey As String = "QueryMatches.IsSelectedEmailKnown"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim strFilter As String
Dim oFolder As Folder
Dim noOfEmails As Long
Dim i As Long
Dim oBaseItem As Object
    Set Contacts = New Collection
    Set oBaseItem = ActiveExplorer.Selection.Item(1)    ' Expecting mail object
    If Not IsMailLike(oBaseItem) Then
        IsSelectedEmailKnown = "*** Das selektierte Item ist keine Mail"
        GoTo ProcReturn       ' hat also keine
    End If
    Set oFolder = getParentFolder(oBaseItem)
    If oFolder Is Nothing Then
        DoVerify False, " item has no parent"
    End If
    strFilter = oBaseItem.SenderEmailAddress
    
    If IsEmailKnown(strFilter, Contacts) Then
        noOfEmails = Contacts.Count
        IsSelectedEmailKnown = "Es wurden " & noOfEmails _
            & " Kontakteinträge für " & strFilter & " gefunden " & vbCrLf
        For i = 1 To noOfEmails
            IsSelectedEmailKnown = IsSelectedEmailKnown _
                & "Email " & noOfEmails & " ist Email" & Contacts(i).Alias _
                & " von " & Quote(Contacts(i).ValueOfItem.Subject)
            If TestOnly < 3 Then    ' Only check for one email
                GoTo ProcReturn
            End If
        Next i
    Else
        IsSelectedEmailKnown = "*** Kein Kontakt vorhanden für " & strFilter
        GoTo ProcReturn
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.IsSelectedEmailKnown

'---------------------------------------------------------------------------------------
' Method : Function IsEmailKnown
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsEmailKnown(MailAddress As String, Contacts As Collection) As Boolean
Dim zErr As cErr
Const zKey As String = "QueryMatches.IsEmailKnown"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    IsEmailKnown = IsEmailInAddressLists(MailAddress, Contacts)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.IsEmailKnown

'---------------------------------------------------------------------------------------
' Method : Function IsEmailInAddressLists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsEmailInAddressLists(MailAddress As String, Contacts As Collection) As Boolean
Dim zErr As cErr
Const zKey As String = "QueryMatches.IsEmailInAddressLists"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim myaddresslists As AddressLists
Dim oneAddressList As AddressList
    Set myaddresslists = aNameSpace.AddressLists
    For Each oneAddressList In myaddresslists
        If IsEmailInThisAddressList(oneAddressList, MailAddress, Contacts) Then
            IsEmailInAddressLists = True
            Exit For
        End If
    Next oneAddressList

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.IsEmailInAddressLists

'---------------------------------------------------------------------------------------
' Method : Function IsEmailInThisAddressList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsEmailInThisAddressList(oneAddressList As AddressList, MailAddress As String, Contacts As Collection) As Boolean
Dim zErr As cErr
Const zKey As String = "QueryMatches.IsEmailInThisAddressList"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim myAddressEntries As AddressEntries
Dim ContactEntry As cNumbItem
Dim thisContact As ContactItem
Dim thisAddressEntry As AddressEntry
Dim i As Long
    If oneAddressList.AddressListType <= olOutlookAddressList Then
        Set myAddressEntries = oneAddressList.AddressEntries
        For i = 1 To myAddressEntries.Count
            Set thisAddressEntry = myAddressEntries.Item(i)
            Set thisContact = thisAddressEntry.GetContact
            If thisContact.Email1Address = MailAddress Then
                Set ContactEntry = New cNumbItem
                ContactEntry.Key = "1"
                Set ContactEntry.ValueOfItem = thisContact
                ContactEntry.Alias = MailAddress
                Contacts.Add ContactEntry
            End If
            If thisContact.Email2Address = MailAddress Then
                Set ContactEntry = New cNumbItem
                ContactEntry.Key = "2"
                ContactEntry.ValueOfItem = thisContact
                ContactEntry.Alias = MailAddress
                Contacts.Add ContactEntry
            End If
            If thisContact.Email3Address = MailAddress Then
                Set ContactEntry = New cNumbItem
                ContactEntry.Key = "3"
                ContactEntry.ValueOfItem = thisContact
                ContactEntry.Alias = MailAddress
                Contacts.Add ContactEntry
            End If
        Next i
    Else
        DoVerify False, " not implemented"
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function ' QueryMatches.IsEmailInThisAddressList


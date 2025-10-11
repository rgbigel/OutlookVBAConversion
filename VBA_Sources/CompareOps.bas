Attribute VB_Name = "CompareOps"
Option Explicit

Dim CompareState As Boolean
Dim RrX As String, RrY As String
Dim LocalPropertyCount As Long
Dim undecodedPropertyCount As Long
Dim MatchIndicator As String
Dim currPropMisMatch As Boolean
Dim MisMatchIgnore As Boolean
Dim UserAnswer As String
Dim xQuickchecksonly As Boolean
Dim xTDeferExcel As Boolean
Dim xTuseExcel As Boolean
Dim doingTheRest As Boolean
Dim AbortSignal As Long
Dim sTime As Variant

'---------------------------------------------------------------------------------------
' Method : Sub C1SI
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub C1SI()                                         ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "CompareOps.C1SI"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")
    
    xUseExcel = False
    xDeferExcel = True
    displayInExcel = True
    SelectOnlyOne = True
    Call SelectAndCompare
    StopRecursionNonLogged = False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.C1SI

'---------------------------------------------------------------------------------------
' Method : Sub C2SI
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub C2SI()                                         ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "CompareOps.C2SI"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")
    
    IsEntryPoint = True

    xUseExcel = False
    xDeferExcel = False
    SelectOnlyOne = False
    FindMatchingItems = True                       ' do not dump item model attributes
    UI_DontUse_Sel = True
    Call SelectAndCompare
    StopRecursionNonLogged = False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.C2SI

'---------------------------------------------------------------------------------------
' Method : Sub Compare2SelectedItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Compare2SelectedItems(Optional forItem1 As Boolean = True, Optional forItem2 As Boolean = True)
Dim zErr As cErr
Const zKey As String = "CompareOps.Compare2SelectedItems"
    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

Dim ExcelDefer As Boolean
    ' Process Folder items
Restart:
    AllProps = False
    AllPropsDecoded = InitsForDecoding(forItem1, forItem2)
    ' Use main Object Identification as used by sort!
    Matchcode = GetObjectAttributes(True, True)
    
    Message = vbCrLf & "Prüfung von " & Folder(1).FolderPath & ": " & WorkIndex(1)
    If Not forItem1 And forItem2 Then
        fiMain(2) = fiMain(1)
    Else
        Message = Message & " und " & Folder(2).FolderPath & ": " & WorkIndex(2)
    End If
    Message = Message & ", " & MainObjectIdentification & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1)
    If fiMain(1) <> fiMain(2) Then
        Message = Message & vbCrLf & WorkIndex(2) & ": " & fiMain(2)
    End If
    Call LogEvent(String(60, "_") & Message)
    
    ExcelDefer = xDeferExcel                       ' save here cause localy changed
    If Matchcode = Passt_Synch Then
        If ItemIdentity(AllPropsDecoded) Then
            ' ==================================================================
            ListContent(ListCount).Compares = "="
            If UserDecisionRequest And Not UserDecisionEffective And (xDeferExcel Or Not xUseExcel) Then
                rsp = MsgBox("Selected items Match. OK to see all details." & vbCrLf & Message, vbOKCancel)
            Else
                rsp = vbNo
            End If
        Else                                       ' ###########################################################
            ListContent(ListCount).Compares = "<>"
            ListContent(ListCount).MatchData = MatchData
            If IsComparemode And rsp <> vbCancel Then
                rsp = MsgBox("Selected items do not Match. OK to see all details." & vbCrLf & Message, vbOKCancel)
            End If
        End If
        If rsp = vbOK Then
            If xlA Is Nothing Then
                Call ShowDetails
                If UserDecisionRequest Then        ' Show decode and Show in excel
                    xDeferExcel = False
                    displayInExcel = True
                    quickChecksOnly = False
                    OnlyMostImportantProperties = False
                    IsComparemode = True
                    ' ?                    doingTheRest = True
                    GoTo Restart
                End If
            Else
                Call DisplayExcel(O)
            End If
        End If
    End If
    Set ListContent = Nothing
    Set sortedItems(1) = Nothing
    Set sortedItems(2) = Nothing
    xDeferExcel = ExcelDefer                       ' restore detailed analysis

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.Compare2SelectedItems

'---------------------------------------------------------------------------------------
' Method : Sub CompareAllStackedAttributes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CompareAllStackedAttributes()
Dim zErr As cErr
Const zKey As String = "CompareOps.CompareAllStackedAttributes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    '       Textual results of Comparison also generated here.
    '       note that aDecProp C may be filled up to MostRelevant only in some cases
    If Not doingTheRest Then
        relevantIndex = 1                          ' start here
    Else
        If DebugMode And relevantIndex <> 1 Then   ' interesting ???
            DoVerify False
        End If
    End If
    LocalPropertyCount = aID(1).idAttrDict.Count - 1
    stpcnt = aID(2).idAttrDict.Count - 1
    If relevantIndex = 1 Then
        Call initializeComparison
        undecodedPropertyCount = 0
        CompareState = True
        LocalPropertyCount = MaxPropertyCount
    Else
        LocalPropertyCount = Max(LocalPropertyCount, stpcnt)
    End If
    If Not ShutUpMode Then
        Debug.Print Format(Timer, "0#####.00") & vbTab & "comparing first property"
    End If
    
    For AttributeIndex = relevantIndex To stpcnt   ' Y y=x count, depending on quickChecksOnly
        OneDiff = vbNullString
        OneDiff_qualifier = vbNullString
        currPropMisMatch = False
        MisMatchIgnore = False
        ' get display values into Rrx, RrY
        If synchedNames(RrX, RrY) Then             ' synchs attribute names and more...
            AbortSignal = 1                        ' did not work: no way to success here
            If DebugMode Then DoVerify False
            GoTo ProcReturn
        End If
        
        ' done in synchedNames: Set aTd = GetAttrDsc(PropertyNameX)
        '                                 also sets correct iRules
        pArr(1) = PropertyNameX
        pArr(2) = RrX
        pArr(3) = RrY
        If PropertyNameX = "HTMLBody" Then DoVerify False
       
        If InStr(PropertyNameX, "EntryID") > 0 Then ' should NOT match
            If RrX = RrY _
            And LenB(RrX) > 0 Then              ' Vergleich nicht sinnvoll für EntryID
                Message = PropertyNameX & " identisch, vermutlich gleiches Objekt"
                If PropertyNameX = "EntryID" Then
                    Message = Message & vbCrLf _
                              & "<> Vergleichen der Objekte gibt keinen Sinn: " _
                              & "beendet wegen Fehler: " _
                              & vbCrLf & Err.Description
                    CompareState = False
                    Call logMatchInfo
                End If
            End If
        End If
DoItAnyway:
        If aTD Is Nothing Then
            If aDecProp(1) Is Nothing Or aDecProp(2) Is Nothing Then
                If DebugMode Then DoVerify False
                GoTo skipCompare
            End If
            Set iRules = sRules
            If DebugMode Then DoVerify False
        End If
        
        If iRules.clsObligMatches.RuleMatches _
        Or (aDecProp(1).adOrigValDecodingOK _
            And aDecProp(2).adOrigValDecodingOK) Then
            
            If Left(RrX, 1) = "#" Then             ' like unequal, but always same
                If iRules.clsNeverCompare.RuleMatches Then
                    IgString = " #. "
                    GoTo iggit
                Else
                    GoTo genMessage
                End If
            End If
            If RrX <> RrY Then
                currPropMisMatch = True
                IgString = vbNullString
                MatchIndicator = "<>  "
                AttributeIndex = aTD.adtrueIndex
                logDecodedProperty aDecProp(1).adShowValue, MatchIndicator
                logDecodedProperty aDecProp(2).adShowValue, MatchIndicator
                
                If ShowEmptyAttributes _
                Or LenB(RrX) > 0 _
                Or LenB(RrY) > 0 Then
                    If LenB(killStringMsg) > 0 Then
                        MatchData = MatchData & vbCrLf & killStringMsg
                    End If
                End If
genMessage:
                If Not iRules.RuleIsSpecific Then  ' just a mismatch
                    MatchIndicator = "...  "
                    GoTo iggit
                End If
                
                If iRules.clsSimilarities.RuleMatches Then
                    MatchIndicator = "~~  "
                    
                    UserDecisionRequest = UserDecisionRequest _
                                          Or AcceptCloseMatches
                    OneDiff = "~~~ Bedingte Abweichung bei Eigenschaft " _
                              & iRuleBits & vbCrLf
                    MisMatchIgnore = True
                    SimilarityCount = SimilarityCount + 1
                    GoTo Logshort
                ElseIf iRules.clsNeverCompare.RuleMatches Then
                    MatchIndicator = "__  "
iggit:
                    IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
                    If iRules.clsNotDecodable.RuleMatches Then
                        undecodedPropertyCount = undecodedPropertyCount + 1
                        OneDiff = "§§§ kein Vergleichsergebnis, nicht dekodiert " _
                                  & iRuleBits & vbCrLf
                    Else
                        OneDiff = "... Abweichung ignoriert für " _
                                  & iRuleBits & vbCrLf
                    End If
                    MisMatchIgnore = True
                    GoTo Logshort
                ElseIf iRules.clsObligMatches.RuleMatches Then
                    CompareState = False
                    If Left(RrX, 1) = "#" Then
                        OneDiff = "undefinierte Abweichung bei Eigenschaft " _
                                  & iRuleBits & vbCrLf
                    Else
                        cMisMatchesFound = cMisMatchesFound + 1
                        OneDiff = cMisMatchesFound _
                                  & ". relevante Abweichung bei Eigenschaft " _
                                  & iRuleBits & vbCrLf
                    End If
                    SuperRelevantMisMatch = True
                    MisMatchIgnore = False
doLog:
                    DiffsRecognized = DiffsRecognized & vbCrLf & "#### " _
                                      & OneDiff
                    GoTo setDiff
Logshort:
                    DiffsRecognized = DiffsRecognized & vbCrLf & OneDiff & _
                                      vbCrLf & aDecProp(2).adKillMsg
setDiff:
                    OneDiff_qualifier = "    " _
                                        & WorkIndex(1) & ": " & aDecProp(1).adShowValue _
                                        & vbCrLf & "    " _
                                        & WorkIndex(2) & ": " & aDecProp(2).adShowValue
                Else
                    If iRules.clsSimilarities.RuleMatches Then
                        OneDiff = OneDiff & "~~~ Abweichung, Ähnlichkeit prüfen " _
                                  & iRuleBits & vbCrLf
                    ElseIf iRules.clsNeverCompare.RuleMatches Then
                        OneDiff = OneDiff & "??? von Vergleich ausgeschlossen " _
                                  & iRuleBits & vbCrLf
                    ElseIf iRules.clsNotDecodable.RuleMatches Then
                        OneDiff = OneDiff & "nicht dekodiert " _
                                  & iRuleBits & vbCrLf
                    Else
                        OneDiff = OneDiff & "§§§ Wertabweichung bei " _
                                  & iRuleBits & vbCrLf
                        cMisMatchesFound = cMisMatchesFound + 1
                        MisMatchIgnore = False
                        CompareState = False
                        MatchIndicator = "##  "
                        DiffsRecognized = DiffsRecognized & vbCrLf & "#### " _
                                          & OneDiff
                        GoTo Logshort
                    End If
                    ' ignored misMatch for whatever reason
                    MisMatchIgnore = True
                    MatchIndicator = "--  "
                    IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
                    GoTo Logshort                  ' what shall we do with onediff???
                End If
                Call logDiffInfo(MatchIndicator & OneDiff & OneDiff_qualifier)
                If MisMatchIgnore Then
                    DiffsIgnored = DiffsIgnored & vbCrLf _
                                   & OneDiff & OneDiff_qualifier
                Else
                    DiffsRecognized = DiffsRecognized & vbCrLf _
                                      & OneDiff_qualifier
                End If
            Else                                   ' the property Matches !
                MatchIndicator = "    "
                Matches = Matches + 1
                If (ShowEmptyAttributes Or LenB(RrX) > 0) _
                   And MinimalLogging < eLmin Then
                    MatchData = MatchData & vbCrLf _
                                & MatchIndicator & WorkIndex(1) & b _
                                & PropertyNameX & "=" & Quote(Left(RrX, 80))
                End If
            End If
        Else
            If RrX <> RrY Then
                currPropMisMatch = True
                MatchIndicator = "..  "
                OneDiff = "... Abweichung nicht gewertet zu *" & iRuleBits & vbCrLf
                MisMatchIgnore = True
                IgnoredPropertyComparisons = IgnoredPropertyComparisons + 1
                GoTo Logshort
            End If
        End If
        
        If LenB(OneDiff) > 0 Then
            AllItemDiffs = AllItemDiffs & OneDiff & OneDiff_qualifier & vbCrLf
            If displayInExcel Then                 ' And (xDeferExcel Or xUseExcel) ???
                If O Is Nothing Then               ' not defined in excel ==> add at end
                    Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
                                     sHdl, False, DebugMode) ' no: put it there
                End If
                pArr(4) = MatchIndicator
                pArr(5) = Trunc(1, OneDiff, vbCrLf)
                Call addLine(O, AttributeIndex, pArr)
                If Not aCell Is Nothing Then
                    Set aCell = O.xlTSheet.Cells(AttributeIndex + 1, ValidCol)
                    If DebugMode Or DebugLogging Then
                        aCell.Select
                    End If
                    If MisMatchIgnore Then
                        aCell.Interior.pattern = xlSolid
                        aCell.Interior.PatternColorIndex = xlAutomatic
                        aCell.Interior.ThemeColor = xlThemeColorAccent6
                    Else
                        aCell.Font.Color = -16776961 ' ROT
                        O.xlTSheet.Cells(AttributeIndex + 1, 1).Font.Color = -16776961
                    End If
                End If
            End If
        End If
        
        If currPropMisMatch And Not MisMatchIgnore Then
            If SuperRelevantMisMatch Then          ' some prop misMatch that counts
                If AcceptCloseMatches Then         ' check approximate Matches
                    If cMisMatchesFound > MaxMisMatchesForCandidates _
                    And quickChecksOnly Then
                        Exit For                   ' no point to compare more
                    End If
                Else                               ' refuse approximate Matches
                    Exit For
                End If
            End If
        End If

        If AttributeIndex Mod 10 = 0 Then
            If Not ShutUpMode Then
                Debug.Print Format(Timer, "0#####.00") _
        & vbTab & "comparing property # " & AttributeIndex
            End If
        End If
skipCompare:
    Next AttributeIndex
    
    ' hurra, alle vergleiche sind gemacht. Zeig das Ergebnis.
    Debug.Print Format(Timer, "0#####.00") & vbTab & " Match=" & CompareState _
                                           & vbTab & "compared last property, # " & AttributeIndex - 1
    AllPropsDecoded = (MaxPropertyCount >= TotalPropertyCount)
    
    Call GenerateSummary

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.CompareAllStackedAttributes

'---------------------------------------------------------------------------------------
' Method : Function CreateStatisticOutput
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CreateStatisticOutput() As String
Dim zErr As cErr
Const zKey As String = "CompareOps.CreateStatisticOutput"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    Message = "    " & Matches & " Übereinstimmungen / " _
              & cMisMatchesFound & " relevante Abweichungen" _
              & " / " & SimilarityCount & " Ähnlichkeiten"
        
    Message = Message & vbCrLf & "    es wurden " _
              & IgnoredPropertyComparisons _
              & " Eigenschaften ignoriert"
        
    If AttributeUndef(1) > 0 Then
        displayInExcel = True
        Message = Message & vbCrLf & "  ## Attribut " _
                  & AttributeUndef(1) + 1 _
                  & " inkonsistent in item " & WorkIndex(1)
    End If
    If AttributeUndef(2) > 0 Then
        displayInExcel = True
        Message = Message & vbCrLf & "  ## Attribut " _
                  & AttributeUndef(2) + 1 _
                  & " inkonsistent in item " & WorkIndex(2)
    End If
    
    If saveItemNotAllowed Then
        Message = Message & vbCrLf & "    " & YleadsXby _
                  & " / " & NotDecodedProperties _
                  & " Eigenschaften sind nicht in beiden Items enthalten" _
                  & " / nicht decodiert"
    End If
    Call logMatchInfo
    CreateStatisticOutput = Message

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' CompareOps.CreateStatisticOutput

'---------------------------------------------------------------------------------------
' Method : Sub DisplayWithExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DisplayWithExcel(Statistics As String)
Dim zErr As cErr
Const zKey As String = "CompareOps.DisplayWithExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If displayInExcel Then
        ' xUseExcel = True obsolete ***???***
        If O Is Nothing Then                       ' defined in excel ?
            Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
                             sHdl, DebugMode)      ' no: put it there
        End If
        Call StckedAttrs2Xcel(O)
        If WorkIndex(1) > 0 And LenB(Statistics) = 0 Then
            Call CompareAllStackedAttributes       ' not necessarily all that exist in Items
            Statistics = CreateStatisticOutput
        End If
        
        If LenB(Statistics) > 0 Then
            PutStatisticOutputToExcel (Statistics)
        End If
        Call ExcelEditSession(0)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.DisplayWithExcel

'---------------------------------------------------------------------------------------
' Method : Sub DisplayWithoutExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DisplayWithoutExcel(oMessage As String)
Dim zErr As cErr
Const zKey As String = "CompareOps.DisplayWithoutExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim textmod As String
    If Not xlApp Is Nothing Then
        If Not (displayInExcel Or O Is Nothing) Then
            If Not UserDecisionRequest Then
                If Not AllPropsDecoded Then
                    textmod = "(Partiellen) "
                End If
                If vbYes = MsgBox(textmod _
                        & "Vergleich in Excel anzeigen?", vbYesNo) Then
                    displayInExcel = True
                Else
                    displayInExcel = False
                    If O Is Nothing Then
                        UserDecisionRequest = True
                    Else
                        If UserDecisionEffective Then
                            GoTo ProcReturn
                        End If
                        Set O = Nothing            ' assume No as an answer next time
                    End If
                End If
            End If
        End If
    End If

    If Not displayInExcel Then
showResults:
        If UserDecisionRequest Or Not (xDeferExcel Or xUseExcel) Then
            CompareState = AskUserAndInterpretAnswer(oMessage)
        Else
            displayInExcel = True                  ' displayWithExcel always called AFTER this !!!
            mustDecodeRest = Not AllPropsDecoded   ' if still incomplete, do rest now
        End If                                     ' use the decision we have at this point
CheckWithUser:
        If rsp = vbRetry Then                      ' we ALWAYS need full comparison now!
            UserDecisionEffective = False
            displayInExcel = True
            mustDecodeRest = Not AllPropsDecoded   ' if still incomplete, do rest now
            quickChecksOnly = False                ' do a complete decode; value is restored on exit
        End If
        If displayInExcel And xlApp Is Nothing Then
            Call XlgetApp
        End If
    End If                                         ' if not DisplayInExcel

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.DisplayWithoutExcel

' Evaluation ExplainS
' 0: PerformChangeOpsForMapiItems
' 1: ModRuleTab
Sub ExcelEditSession(EvaluationMode As Long)
Dim zErr As cErr
Const zKey As String = "CompareOps.ExcelEditSession"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    rsp = vbNo
    Call Try(allowAll)                                ' Try anything, autocatch, Err.Clear
    xlApp.EnableEvents = False
    Catch
    With x.xlTSheet
        .Cells(1, clickColumn) = "Hier Klicken:"
        .Cells(1, moreColumn) = vbNullString
        .Cells(1, clickColumn).Interior.ColorIndex = 35 ' hell GRÜN
        .Cells(1, moreColumn).Interior.ColorIndex = 35
        .Cells(1, promptColumn).Interior.ColorIndex = 0
        .Cells(1, promptColumn).Font.Color = -16776961 ' ROT
        .Cells(1, promptColumn) = "Return -> Outlook"
        .Cells(1, changeCounter) = 0
        Call DisplayExcel(x, EnableEvents:=False, unconditionallyShow:=True)
    
        xlApp.EnableEvents = False
        Catch
        .Cells(1, WatchingChanges) = True          ' set for editing
        .Cells(1, WatchingChanges).Interior.ColorIndex = 35
        If .Cells(1, flagColumn) = "Flag" Then
            .Cells(1, flagColumn).Interior.ColorIndex = 35
        End If
        .Activate
        Catch
        
        xlApp.EnableEvents = True
        sTime = 0
        If Not ShutUpMode Then
            Debug.Print Format(Timer, "0#####.00") & vbTab _
                        & "Beginning Excel Edit Session"
        End If
        
        Call DisplayWindowInFront(xLMainWindowHdl, 1)
        xlApp.Cursor = xlDefault
stillEditing:
        If xlApp.ActiveWindow Is Nothing Then
            Call ClearWorkSheet(xlA, x)
            
            rsp = vbYes                            ' user closed the window, so it is like edit aborted
            If AllPropsDecoded Then
                Call LogEvent("Excel Window wurde geschlossen, Änderungen " _
                              & "werden nicht aus Excel übernommen.", eLall)
                rsp = vbCancel
            Else                                   ' closed excel, but we did not Show all properties:
                Call LogEvent("Excel geschlossen (kein Edit), " _
                              & "es erfolgt eine neue, vollständige Darstellung aller Attribute", eLall)
                rsp = vbRetry                      ' do a retry with all
            End If
            GoTo canMod
        ElseIf Not xlApp.Visible Then
            xlApp.EnableEvents = False
            GoTo showResults
        Else
showResults:
            xlApp.Cursor = xlDefault
            Catch
            
            If DebugMode Then
                If MsgBox("waiting here for end of edit in Excel, Click into cell(1, " _
                          & clickColumn & ")", vbOKCancel) = vbCancel Then
                    DoVerify False
                End If
            End If
excelFinished:
            If xlApp.ActiveSheet Is Nothing Then
                Debug.Print "user has closed excel sheet or application"
                Set x = Nothing
                Set x = xlWBInit(xlA, TemplateFile, cOE_SheetName, _
                                 sHdl, showWorkbook:=DebugMode) ' no: put it there
                rsp = vbRetry
                quickChecksOnly = False
                mustDecodeRest = True
                GoTo canMod
            End If
            If InStr(.Cells(1, promptColumn).Text, "Go Outlook") > 0 Then
                displayInExcel = False
                xlApp.Visible = False
                xlApp.EnableEvents = False
                x.xlTSheet.EnableCalculation = False
                x.xlTSheet.EnableFormatConditionsCalculation = False
                Catch
                Call LogEvent(Format(Timer, "0#####.00") & " Edit in Excel fertig", eLall)
                Call DisplayWindowInFront(FWP_xLW_Hdl, 2)
                rsp = vbYes                        ' Excel Return is like Answer = yes
                Select Case EvaluationMode
                    Case 0:
                        Call PerformChangeOpsForMapiItems
                    Case 1:
                        Call ModRuleTab
                        GoTo ProcReturn
                    Case Else
                        DoVerify False, " not imp"
                End Select
                If Folder(2) Is Nothing Then
                    If Folder(1).Parent Is Nothing Then
                        Message = vbNullString
                    Else
                        Message = "Item ist in " & Folder(1).FolderPath
                    End If
                ElseIf Folder(1).Parent Is Nothing Or Folder(2).Parent Is Nothing Then
                    GoTo setFmsg
                ElseIf Folder(1).Parent.Name = Folder(2).Parent.Name Then
                    Message = vbNullString
                Else
setFmsg:
                    Message = "Item 2 ist in " & Folder(2).FolderPath
                End If
            Else
                If Wait(0.5, sTime, "Warte auf Ende in Excel. Klicken Sie auf 'Hier Klicken' ") Then
                    DoVerify False
                    GoTo excelFinished
                End If
                Call ShowStatusUpdate
                GoTo stillEditing
            End If
        End If
canMod:
        If rsp = vbYes Then
            Call SaveItemsIfChanged(True)
            AllPropsDecoded = True                 ' do not get more attributes
        ElseIf rsp = vbCancel Then
            If TerminateRun Then
                GoTo ProcReturn
            End If
        ElseIf rsp = vbNo Then
            GoTo showResults
            '               vbRetry is also possible, user closed excel ==> decode all
        End If
    End With                                       ' X.xlTSheet

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.ExcelEditSession

'---------------------------------------------------------------------------------------
' Method : Sub ExcelShowItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ExcelShowItem()                                ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "CompareOps.ExcelShowItem"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="CompareOps")
    
    IsEntryPoint = True

    ActionTitle(0) = "Attribute in Excel zeigen für selektiertes Element in"
    AttributeUndef(1) = 0
    AttributeUndef(2) = 0
    aPindex = 1
        
    xUseExcel = True
    xDeferExcel = False
    displayInExcel = True
    SelectOnlyOne = True
    IsComparemode = True
    eOnlySelectedFolder = False
    Set LF_CurLoopFld = Nothing
    UI_DontUse_Sel = True                          ' no special selection/filter parameters to be used
    UI_DontUseDel = True                           ' no deletion rules to be used
    E_Active.EventBlock = True
    Call SelectAndCompare                          ' converts and displays in Excel sheet O (Objekteigenschaften)
    E_Active.EventBlock = False
    If T_DC.TermRQ Then
        Call TerminateRun
        GoTo ProcReturn
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.ExcelShowItem

'---------------------------------------------------------------------------------------
' Method : Sub GenerateSummary
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GenerateSummary()
Dim zErr As cErr
Const zKey As String = "CompareOps.GenerateSummary"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If quickChecksOnly Then                        ' !! idObjItem is used to access oTi !!
        NotDecodedProperties = _
                             Max(aID(1).idObjItem.ItemProperties.Count, _
                                 aID(2).idObjItem.ItemProperties.Count) _
        - LocalPropertyCount
    Else
        NotDecodedProperties = LocalPropertyCount - (AttributeIndex - 1)
    End If
    NotDecodedProperties = NotDecodedProperties + undecodedPropertyCount
    
    If quickChecksOnly Or Not (MisMatchIgnore Or IsComparemode) Then
        If CompareState Then
            If Not quickChecksOnly Then
                ' we must continue at least until the cmisMatchesfound-limit is exceeded
                GoTo keepChecking
            End If
        End If
        
        If NotDecodedProperties > 0 Then
            saveItemNotAllowed = True
            If Not quickChecksOnly _
            Or mustDecodeRest Then
                ' we must continue at least until the cmisMatchesfound-limit is exceeded
                GoTo keepChecking
            End If
            If displayInExcel Then
                pArr(1) = "*** es wurden nicht alle Merkmale verglichen"
                Call addLine(O, aID(2).idAttrDict.Count + 2, pArr)
            End If
            AllItemDiffs = AllItemDiffs & vbCrLf & pArr(1)
            IgnoredPropertyComparisons = IgnoredPropertyComparisons _
                                         + NotDecodedProperties
        End If
    Else
keepChecking:
        doingTheRest = True
                
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.GenerateSummary

'---------------------------------------------------------------------------------------
' Method : Sub GetCompareItems
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetCompareItems()
Dim zErr As cErr
Const zKey As String = "CompareOps.GetCompareItems"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    rsp = 0                                        ' not a valid value at all!!
    LocalPropertyCount = 0
    currPropMisMatch = False
    YleadsXby = 0
    
    sHdl = "CritPropName---------------" _
           & " Objekt-" & WorkIndex(1) & "----------------------" _
           & " Objekt-" & WorkIndex(2) & "----------------------" _
           & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
    Call InitsForPropertyDecoding(doingTheRest)
    
    If mustDecodeRest And Not quickChecksOnly Then
        Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
             ContinueAfterMostRelevant:=mustDecodeRest, _
             onlyItemNo:=3)
    Else
        If isEmpty(MostImportantProperties) Then
            GoTo allOfThem
        End If
        Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
             ContinueAfterMostRelevant:=mustDecodeRest, _
             onlyItemNo:=1)
allOfThem:
        If displayInExcel _
        And Not xlApp _
        And Not O Is Nothing Then               ' unless we defer
            Call StckedAttrs2Xcel(O)
        End If
        Call DecodeAllPropertiesFor2Items(StopAfterMostRelevant:=quickChecksOnly, _
             ContinueAfterMostRelevant:=mustDecodeRest, _
             onlyItemNo:=2)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.GetCompareItems

'---------------------------------------------------------------------------------------
' Method : Function GetObjectAttributes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetObjectAttributes(Optional forItem1 As Boolean = True, Optional forItem2 As Boolean = True) As Long
Dim zErr As cErr
Const zKey As String = "CompareOps.GetObjectAttributes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim MainCompare As Long

    aCloneMode = withNewValues
    
    If forItem1 Then
        If SelectOnlyOne And forItem2 Then
            DoVerify False, " inconsistent ???"
        End If
        Call GetAobj(1, WorkIndex(1))
        objTypName = DecodeObjectClass(getValues:=True)
        DoVerify WorkIndex(1) = aItmIndex
        Call AddItemToList((0), fiMain(1), (Passt_Deleted), vbNullString)
    End If
    If forItem2 Then
        If SelectOnlyOne And forItem1 Then
            DoVerify False, "SelectOnlyOne inconsistent with: decode for item 1 when decoding item 2"
        End If
        Call GetAobj(2, WorkIndex(2))
        objTypName = DecodeObjectClass(getValues:=True)
        If objTypName = "not defined" Then
            WorkIndex(2) = 0
            GoTo ProcReturn                        ' value 0
        End If
        Call AddItemToList(vbNullString, fiMain(2), (Passt_Inserted), (0))
    End If
    
    If SelectOnlyOne Then
        MainCompare = True
    Else
        MainCompare = StrComp(fiMain(1), fiMain(2), vbTextCompare)
    End If
    GetObjectAttributes = Passt_Synch

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' CompareOps.GetObjectAttributes

'---------------------------------------------------------------------------------------
' Method : Sub IdentityCheck
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub IdentityCheck()
Dim zErr As cErr
Const zKey As String = "CompareOps.IdentityCheck"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Statistics As String
Dim DontSkipActions As Boolean
    Call GetCompareItems
    ' FindMatchingItems = true ==> selected main properties by search already
    CompareState = (fiMain(1) = fiMain(2)) Or FindMatchingItems
    UserDecisionEffective = False
    If quickChecksOnly Then
        If Not CompareState Then
            GoTo ProcReturn                        ' primary ids are different, we don't care about other details
        End If
    End If
    
    If Not AllPropsDecoded Then
        If Not quickChecksOnly Then
            quickChecksOnly = False
Retry:
            If AllPropsDecoded Then
                Call initializeComparison          ' restart here *** for manual debug
                Call initializeExcel
            End If                                 ' else, start after mandatories
            AllPropsDecoded = False                ' rules could have changed
            mustDecodeRest = True
            aPindex = 1
            If aOD(0).objMaxAttrCount <> aID(1).idAttrDict.Count - 1 Then
                aOD(0).objDumpMade = -1            ' dump AttributeDescriptors again (seldom)
            End If
            aOD(1).objDumpMade = -1
            aOD(2).objDumpMade = -1
            Call GetCompareItems
        End If
    End If
        
    Call CompareAllStackedAttributes               ' not necessarily all that exist in Items
    
    DontSkipActions = True
    If CompareState Then                           ' we found identical one
        Select Case ActionID
            Case atOrdnerinhalteZusammenführen
                DontSkipActions = False
                If DebugMode Or DebugLogging Then
                    Debug.Print "Fast Check completed, items Match with " _
                                & Matches & " Attributes, no relevant misMatches"
                End If
            Case Else
        End Select
    End If
    If DontSkipActions And Not UserDecisionEffective And Not xReportExcel Then
        Statistics = CreateStatisticOutput
        Call DisplayWithoutExcel(Statistics)
        If rsp = vbRetry Then
            GoTo Retry
        End If
        Call DisplayWithExcel(Statistics)
        If rsp = vbRetry Then
            GoTo Retry
        End If
        Call SaveItemsIfChanged                    ' sets SaveItemRequested if we saved 1 or 2
        Call QueryAboutDelete(Statistics)
        If rsp = vbRetry Then
            GoTo Retry
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.IdentityCheck

'---------------------------------------------------------------------------------------
' Method : Sub InitObjectSelection
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitObjectSelection()
Dim zErr As cErr
Const zKey As String = "CompareOps.InitObjectSelection"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    WorkIndex(1) = 1
    If aPindex > 0 Then                            ' there is no (need no) aID(0).idAttrDict
        ' set dynamic headline
        If SelectOnlyOne Then
            sHdl = "CritPropName---------------" _
                   & " Objekt-----------------------" _
                   & " -----------------------------" _
                   & " Comp-- Info-- Flag-- ign.parts1"
            WorkIndex(2) = inv                     ' undef
        Else
            sHdl = "CritPropName---------------" _
                   & " Objekt-" & WorkIndex(1) & "----------------------" _
                   & " Objekt-" & WorkIndex(2) & "----------------------" _
                   & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
            WorkIndex(2) = 2
        End If
    End If
    
    ' if we compare, close matches are ok if no quickchecks
    AcceptCloseMatches = IsComparemode And Not quickChecksOnly
    OnlyMostImportantProperties = quickChecksOnly  ' decode OnlyMostImportantProperties
    MinimalLogging = 3
    WantConfirmation = True
    MatchMin = 1000
    MaxMisMatchesForCandidates = MaxMisMatchesForCandidatesDefault
    ListCount = 0
    MaxPropertyCount = 0
    Set sortedItems(1) = Nothing
    Set sortedItems(2) = Nothing
    Set Folder(2) = ChosenTargetFolder
    If LF_UsrRqAtionId <> atBearbeiteAlleÜbereinstimmungenzueinerSuche Then
        Set SelectedItems = New Collection
        Set Folder(1) = Nothing
        Set ChosenTargetFolder = Nothing           ' default: let user find Folder next time
    End If
    Set ListContent = Nothing
    
    AllDetails = vbNullString
    eOnlySelectedItems = True
    
    bDefaultButton = "No"
    rsp = vbOK

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.InitObjectSelection

'---------------------------------------------------------------------------------------
' Method : Function InitsForDecoding
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function InitsForDecoding(Optional forItem1 As Boolean = True, Optional forItem2 As Boolean = True) As Boolean
Dim zErr As cErr
Const zKey As String = "CompareOps.InitsForDecoding"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    DeletedItem = inv
    DeleteIndex = -1
    If Not forItem1 And forItem2 Then
        InitsForDecoding = True
    Else
        Set LF_CurLoopFld = Folder(1)
        Call BestObjProps(LF_CurLoopFld, withValues:=True)
        Call Initialize_UI
        If StopLoop Then
            GoTo ProcReturn
        End If
        InitsForDecoding = False
    End If
    
    Call FindTopFolder(LF_CurLoopFld)
    
    If xUseExcel Then
        If xlApp Is Nothing Then
            Call XlgetApp
            Call XlopenObjAttrSheet(xlA)
        ElseIf O Is Nothing Then
            Call XlopenObjAttrSheet(xlA)
        Else
            If O.xlTSheet.Name <> cOE_SheetName Then
                Call XlopenObjAttrSheet(xlA)
            End If
        End If
    End If
    
    FindTrashFolder

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' CompareOps.InitsForDecoding

'---------------------------------------------------------------------------------------
' Method : Function ItemIdentity
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ItemIdentity(Optional ByVal DecodingFinished As Boolean) As Boolean
Dim zErr As cErr
Const zKey As String = "CompareOps.ItemIdentity"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    AbortSignal = 0                                ' all is ok so far
    AllPropsDecoded = DecodingFinished
    xQuickchecksonly = quickChecksOnly             ' keep for restore at end of Function
    xTDeferExcel = xDeferExcel
    xTuseExcel = xUseExcel
    ' ActionID = 0, 5, flagcolumn, normally lead to editable excel display
    displayInExcel = (xUseExcel Or xDeferExcel) _
                     And (ActionID = 0 _
                     Or ActionID = atNormalrepräsentationerzwingen _
                     Or ActionID = atOrdnerinhalteZusammenführen)
   
    If Not quickChecksOnly Then                    ' mostimportant ones Match already
        mustDecodeRest = True                      ' we already did those, now the rest
    End If
    
    Call IdentityCheck
    If Not (DebugMode Or DebugLogging) Then
        If Not xlApp Is Nothing Then
            If Not CheckExcelOK Then
                Call xlEndApp
            End If
            If Not xlC Is Nothing Then
                Call ClearWorkSheet(xlC, O)        ' it is NOT closed here!
            End If
        End If
        rsp = vbCancel                             ' do not redisplay user questions
    End If
    
    quickChecksOnly = xQuickchecksonly
    xDeferExcel = xTDeferExcel
    xUseExcel = xTuseExcel
    ItemIdentity = CompareState

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' CompareOps.ItemIdentity

'---------------------------------------------------------------------------------------
' Method : Sub multiLinesToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub multiLinesToExcel(Message As String, xw As cXLTab, LineNo As Long, col As Long)
Dim zErr As cErr
Const zKey As String = "CompareOps.multiLinesToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim MultiLine As Variant
Dim MsgLine As Variant
    MultiLine = split(Message, vbCrLf)
    For Each MsgLine In MultiLine
        pArr(col) = MsgLine
        LineNo = LineNo + 1
        Call addLine(xw, LineNo, pArr)
    Next MsgLine

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.multiLinesToExcel

'---------------------------------------------------------------------------------------
' Method : Sub PutStatisticOutputToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PutStatisticOutputToExcel(Message As String)
Dim zErr As cErr
Const zKey As String = "CompareOps.PutStatisticOutputToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If Not displayInExcel Then
        GoTo ProcReturn
    End If
    stpcnt = LocalPropertyCount + 1                ' excel Row number for messages
    If Not O Is Nothing Then
        Call multiLinesToExcel(Message, O, LineNo:=stpcnt, col:=1)
        If LenB(DiffsRecognized) > 0 Then
            Call multiLinesToExcel(DiffsRecognized, O, LineNo:=stpcnt, col:=1)
        End If
        If LenB(DiffsIgnored) > 0 Then
            Call multiLinesToExcel(DiffsIgnored, O, LineNo:=stpcnt, col:=1)
        End If
        W.xlTSheet.Activate
        W.xlTSheet.Cells(LocalPropertyCount + 3, 1).Select
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.PutStatisticOutputToExcel

'---------------------------------------------------------------------------------------
' Method : Sub QueryAboutDelete
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub QueryAboutDelete(oMessage As String)
Dim zErr As cErr
Const zKey As String = "CompareOps.QueryAboutDelete"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    ' find out if we must query about deletes and if its reasonable
    If UserDecisionEffective Then
        GoTo ProcReturn                            ' we already asked user
    End If
    If ActionID = 0 _
    Or cMisMatchesFound + SimilarityCount < MaxMisMatchesForCandidates Then
        If IsComparemode _
        And NotDecodedProperties > 0 _
        And Not UserDecisionRequest _
        And Not SuperRelevantMisMatch Then      ' Automatic decision or ...
            If NotDecodedProperties > 0 Then
                UserAnswer = Quote(fiMain(1)) _
                    & "  sollte bei unvollständigem Vergleich nicht gelöscht werden " _
                    & vbCrLf & vbCrLf & AllItemDiffs
                If AllPropsDecoded Then
                    CompareState = AskUserAndInterpretAnswer(oMessage)
                ElseIf displayInExcel Then
                    Call DisplayExcel(O, relevant_only:=True, _
                                      unconditionallyShow:=True)
                    Call PerformChangeOpsForMapiItems
                Else
                    CompareState = AskUserAndInterpretAnswer(oMessage)
                End If
            End If
            CompareState = (cMisMatchesFound = 0) _
                            Or ((cMisMatchesFound <= MaxMisMatchesForCandidates) _
                                And AcceptCloseMatches)
        Else                                       ' more complex cases
            If ActionID <> atOrdnerinhalteZusammenführen Then
                ' merge Folders without delete (normally)
                CompareState = AskUserAndInterpretAnswer(oMessage)
            End If
        End If
    End If                                         ' must query

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.QueryAboutDelete

'---------------------------------------------------------------------------------------
' Method : Sub SelectAndCompare
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SelectAndCompare(Optional DontDecode As Boolean)
Dim zErr As cErr
Const zKey As String = "CompareOps.SelectAndCompare"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim action As String
Dim SelCount As Long

    If SelectedItems Is Nothing Then
        GoTo prep
    End If
    If Not (eOnlySelectedItems And SelectedItems.Count > 0) Then
prep:
        Call InitObjectSelection
        SelCount = ActiveExplorer.Selection.Count
    Else
        SelCount = SelectedItems.Count
        GoTo matchGet
    End If
    If SelCount >= 1 Then
        Set ActiveExplorerItem(1) = GetDItem_P()
        If SelectMulti Then
            GoTo gothemAll
        End If
    End If
    
    If SelCount > 1 Then                           ' do we have enough to work on?
        If SelectOnlyOne Then
            ActiveExplorer.ClearSelection          ' too much
            GoTo ask1again
        End If
    Else                                           ' we can work with this
        ' ====== AllPublic Global value used, set VVVVVVVVVVVVV before call
        If SelCount = 0 _
        Or (SelCount < 2 And Not SelectOnlyOne) _
        Or (SelectOnlyOne And SelCount <> 1) Then
            action = " und bestätigen Sie die Auswahl mit 'OK'"
ask1again:
            If SelectOnlyOne Then
                rsp = NonModalMsgBox("bitte wählen Sie ein Objekt im richtigen Ordner!" _
                                     & action, _
                                     "OK", "Cancel", "Auswahl Item für Ähnlichkeitssuche")
            Else
                rsp = NonModalMsgBox("bitte wählen Sie einen" _
                                     & " oder zwei zu vergleichende Objekte" _
                                     & action, _
                                     "OK", "Cancel", "Auswahl von Objekten")
            End If
        End If
    End If
    
    Select Case rsp
        Case vbOK
            If ActiveExplorer.Selection.Count = 1 Then
                Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
                If SelectOnlyOne Then
                    GoTo gothemAll
                End If
            ElseIf ActiveExplorer.Selection.Count = 2 Then
                Set ActiveExplorerItem(1) = ActiveExplorer.Selection.Item(1)
                Set ActiveExplorerItem(2) = ActiveExplorer.Selection.Item(2)
                GoTo gothemAll
            Else
                action = " (nur je ein oder genau 2 Item(s) selektieren, bitte)"
                GoTo ask1again
            End If
        Case vbCancel
            Call LogEvent("=======> Stopped before processing any Folders . Time: " _
                          & Now())
            If TerminateRun Then
                GoTo ProcReturn
            End If
    End Select
    action = " und bestätigen Sie die Auswahl mit 'OK'"
ask2again:
    rsp = NonModalMsgBox("bitte wählen Sie das zweite zu vergleichende Objekt" _
                         & action, _
                         "OK", "Cancel", "Auswahl der zu vergleichenden Items")
    
    Select Case rsp
        Case vbOK
            If ActiveExplorer.Selection.Count = 1 Then
                Set ActiveExplorerItem(2) = ActiveExplorer.Selection.Item(1)
            Else
                action = " (nur ein Item, bitte)"
                GoTo ask2again
            End If
        Case vbCancel
            Call LogEvent("=======> Stopped before processing any Folders . Time: " _
                          & Now())
            If TerminateRun Then
                GoTo ProcReturn
            End If
    End Select
    
gothemAll:
    Set SelectedObjects = ActiveExplorer.Selection
    Call GetSelectedItems(ActiveExplorerItem)      ' Selection -> SelectedItems
matchGet:
    If Not sRules Is Nothing Then
        If sRules.RuleObjDsc Is Nothing Then
            GoTo NoObjDsc
        End If
        If SelectedItems.Item(1).Class <> sRules.RuleObjDsc.objItemClass Then
NoObjDsc:
            Set sRules = Nothing                   ' Rules can not be used for differrent class
            Set aTD = Nothing
            Set sDictionary = Nothing
        End If
    End If
    ' get all attributes; relevant properties are located
    dcCount = 0
    Set DeletionCandidates = New Dictionary
    LöListe = vbNullString
    
    If SelectOnlyOne Then
        aItmIndex = 1
        AllPropsDecoded = InitsForDecoding(forItem2:=False)
        If DontDecode Then
            GoTo preponly                          ' no decode or get any attributes yet
        End If
        ' Use main Object Identification as used by sort!
        Stop                                       ' hier geht's daneben:
        Matchcode = GetObjectAttributes(True, False)
        
        Call LogEvent(String(60, "_") & vbCrLf & "Attribute von " _
                      & Folder(1).FolderPath & ": " & WorkIndex(1) _
                      & b & MainObjectIdentification _
                      & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))
            
        If displayInExcel Then
            If xlApp Is Nothing Then
                Call XlgetApp
OisN:
                Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl)
                Set O = x                          ' open default but don't Show it
            End If
            If O Is Nothing Then
                GoTo OisN
            End If
            Call StckedAttrs2Xcel(O)
        End If
        
        aOD(0).objMaxAttrCount = aOD(1).objMaxAttrCount
        If Not xlApp Is Nothing Then
            Call DisplayExcel(O, relevant_only:=True, _
                              EnableEvents:=False, _
                              unconditionallyShow:=True)
            Call DisplayWindowInFront(xLMainWindowHdl, 1)
        End If
        If Not ShutUpMode Then
            Debug.Print "done, one Selected Compare Item, Excel visible and not waiting"
        End If
        GoTo preponly
    Else
        aItmIndex = 0
        Call Compare2SelectedItems
        Call DoTheDeletes
    End If
    
FuncExit:
    If TerminateRun Then
        GoTo ProcReturn
    End If
preponly:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.SelectAndCompare

'---------------------------------------------------------------------------------------
' Method : Function setDefaultWindowName
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function setDefaultWindowName(wTx As Long) As String
Dim zErr As cErr
Const zKey As String = "CompareOps.setDefaultWindowName"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    Static WindowNamePattern(1 To 2) As String
    If LenB(WindowNamePattern(wTx)) = 0 Then
        Select Case wTx
            Case 1                                 ' Excel
                WindowNamePattern(wTx) = RTail(TemplateFile, "\") _
        & " - " & Replace(xlApp.Name, "Microsoft ", "*") ' geht bei [Schreibgeschützt] nicht
            Case 2                                 ' Outlook
                WindowNamePattern(wTx) = "* - Microsoft " & olApp.Name
            Case Else
                DoVerify False
        End Select
    End If
    setDefaultWindowName = WindowNamePattern(wTx)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' CompareOps.setDefaultWindowName

'---------------------------------------------------------------------------------------
' Method : Sub ShowDetails
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowDetails()
Dim zErr As cErr
Const zKey As String = "CompareOps.ShowDetails"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    ListContent(ListCount).MatchData = Message _
                                       & vbCrLf & MainObjectIdentification _
                                       & ": " & fiMain(1) & vbCrLf & AllDetails
    ListContent(ListCount).DiffsRecognized = AllItemDiffs
    rID = ListCount
    frmCompareInfo.Show

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' CompareOps.ShowDetails

'---------------------------------------------------------------------------------------
' Method : Sub AddItemDataToOlderFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AddItemDataToOlderFolder()
Dim zErr As cErr
Const zKey As String = "FolderOps.AddItemDataToOlderFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call Pick2Folders
    AcceptCloseMatches = True                      ' User may refine comparison result if closely related items
    quickChecksOnly = True                         ' initially, only check most relevant
    Call InitFolderOps(False)                      'sort both Folders, not only second
    Call getCriteriaList
    
    Call MergeItemsNewIntoOld
    If TerminateRun Then
        GoTo ProcReturn
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.AddItemDataToOlderFolder

'---------------------------------------------------------------------------------------
' Method : Sub AskForUserAction
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AskForUserAction(Missing As Boolean, Optional SomeMatch As Boolean)
Dim zErr As cErr
Const zKey As String = "FolderOps.AskForUserAction"
    Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

Dim movedItem As Object
Dim targetItem As Object
Dim DelObjectItem As cDelObjectsEntry

    Set targetItem = aID(1).idObjItem
    Set DelObjectItem = New cDelObjectsEntry
    DelObjectItem.DelObjPindex = 1
    DelObjectItem.DelObjPos = WorkIndex(1)
    DelObjectItem.DelObjInd = sortedItems(1).Count > 0 ' from sorted items?
    
    If Missing Then
        rsp = MsgBox("Wollen Sie das Item in den Ordner " _
                     & Quote(Folder(2).FolderPath) & "  übernehmen?" _
                     & vbCrLf & "    " & Quote(fiMain(1)), vbYesNoCancel, _
                     "Behandlung eines Items in " & Quote(Folder(1).FolderPath))
        If rsp = vbYes Then
            ' if copy fails, it is general error so we do no ErrTry()
            Set movedItem = aID(1).idObjItem.Copy
            If movedItem.Class = olAppointment Then
                movedItem.Subject = aID(1).idObjItem.Subject ' get rid of rename (seen on Appointment)
            Else
                'On Error GoTo 0
                If movedItem.Subject <> aID(1).idObjItem.Subject Then
                    DoVerify False
                End If
            End If
            aBugTxt = "move copied item to folder " & Folder(2).FolderPath
            Call Try
            Set movedItem = movedItem.Move(Folder(2))
            If Catch Then
                Message = "Item Copy " & aID(1).idObjItem.Subject _
        & " NOT moved to " & Folder(2).FolderPath _
        & " reason: " & E_AppErr.Description
            Else
                Message = "Item Copy successfully moved into " _
                          & Folder(2).FolderPath _
                          & " as " & Quote(movedItem.Subject) & b
            End If
            Call LogEvent(Message, eLall)
            
        ElseIf rsp = vbNo Then
            Set targetItem = aID(1).idObjItem
delOption:
            rsp = MsgBox("Wollen Sie das original-Item aus Ordner " _
                         & Quote(targetItem.Parent.FolderPath) & "  löschen?" _
                         & vbCrLf & "    " & Quote(fiMain(1)) & "  -> " _
                         & Quote(TrashFolder.FolderPath), _
                         vbYesNoCancel, "Behandlung eines Items in " _
                                       & Quote(targetItem.Parent.FolderPath))
            If rsp = vbYes Then
                Call TrashOrDeleteItem(DelObjectItem)
            ElseIf rsp = vbCancel Then
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            Else                                   ' No: do nothing
            End If
        ElseIf rsp = vbCancel Then
            If TerminateRun Then
                GoTo ProcReturn
            End If
        End If
    Else                                           ' NOT missing, but evtl. differs or duplikate
        If SomeMatch Or ActionID = atOrdnerinhalteZusammenführen Then
            If ActionID = atOrdnerinhalteZusammenführen Then
                ' do not copy to target, identical item exists
            Else
                DoVerify False, " strange or not implemented ???"
            End If
        Else
            DelObjectItem.DelObjPindex = 2
            DelObjectItem.DelObjPos = WorkIndex(2)
            DelObjectItem.DelObjInd = sortedItems(2).Count > 0 ' from sorted items?
            Set targetItem = aID(2).idObjItem
            GoTo delOption
        End If
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.AskForUserAction

'---------------------------------------------------------------------------------------
' Method : Sub CompareItemsIn2Folders
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CompareItemsIn2Folders()
Dim zErr As cErr
Const zKey As String = "FolderOps.CompareItemsIn2Folders"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Call Pick2Folders
    Call InitFolderOps(False)                      'sort both Folders, not only second
    Call CompareItemsNew2Old
    If TerminateRun Then
        GoTo ProcReturn
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.CompareItemsIn2Folders

'---------------------------------------------------------------------------------------
' Method : Sub CompareItemsNew2Old
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CompareItemsNew2Old()
Dim zErr As cErr
Const zKey As String = "FolderOps.CompareItemsNew2Old"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    ' Process Folder items
    DeletedItem = inv
    DeleteIndex = -1
    DateSkipCount = 0
        
    WorkIndex(1) = 0
    WorkIndex(2) = 0
    If 0 >= Ictr(1) Or 0 >= Ictr(2) Then           ' nothing to compare
        GoTo endsub
    End If
    DateSkipCount = 0
    Do                                             ' loop follows a synch as long as there's something left to do
        WorkIndex(1) = WorkIndex(1) + 1
        If ItemDateFilter(sortedItems(1).Item(WorkIndex(1))) = vbNo Then
            GoTo nextOne
        End If
        WorkIndex(2) = WorkIndex(2) + 1
        If ItemDateFilter(sortedItems(2).Item(WorkIndex(2))) = vbNo Then
            WorkIndex(1) = WorkIndex(1) - 1
            GoTo nextOne
        End If
        AllDetails = vbNullString
                         
        ' Use main Object Identification as used by sort!
        Matchcode = quickCheck()
        
        If Matchcode = inv Then
            GoTo nextOne
        End If
        
        Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                      & Folder(1).FolderPath & ": " & WorkIndex(1) & " und " _
                      & WorkIndex(2) & ", " & MainObjectIdentification _
                      & ": " & vbCrLf & WorkIndex(1) & ": " & fiMain(1))
        If fiMain(1) <> fiMain(2) Then
            Call LogEvent(WorkIndex(2) & ": " & fiMain(2))
        End If
        
        If Matchcode = Passt_Synch Then
            If ItemIdentity() Then
                ' ==================================================================
                ListContent(ListCount).Compares = "="
            Else
                ListContent(ListCount).Compares = "<>"
            End If
            ListContent(ListCount).MatchData = MatchData
            ListContent(ListCount).DiffsRecognized = DiffsRecognized
        End If
nextOne:
    Loop Until WorkIndex(1) >= Ictr(1) Or WorkIndex(2) >= Ictr(2)
    
    If DateSkipCount > 0 Then
        Call AddItemToList((DateSkipCount), "Einträge übersprungen weil Datum", vbNullString, vbNullString)
        ListContent(ListCount).MatchData = "vor dem"
        ListContent(ListCount).DiffsRecognized = CStr(CutOffDate)
    End If
    
    If ListContent.Count > 0 Then
        frmDeltaList.Show
    End If
endsub:
    Set ListContent = Nothing
    Set sortedItems(1) = Nothing

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.CompareItemsNew2Old

'---------------------------------------------------------------------------------------
' Method : Sub MergeItemsNewIntoOld
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MergeItemsNewIntoOld()
Dim zErr As cErr
Const zKey As String = "FolderOps.MergeItemsNewIntoOld"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim reportedMatches As Long
Dim ix1 As Long
Dim ix2 As Long
Dim foundItem As Object
Dim MandatoryWorkRule As cNameRule
    DeletedItem = inv
    DeleteIndex = -1
    ListCount = 0
    ' work rule is potentially changed interactively
    Set MandatoryWorkRule = sRules.clsObligMatches.Clone(False)
   
    For ix1 = 1 To Ictr(1)
        Set foundItem = GetAobj(1, ix1)
        
        Call find_Corresponding(foundItem, _
                                MandatoryWorkRule, _
                                reportedMatches, _
                                eliminateID:=True)
        
        If reportedMatches > 0 Then                ' found one or more
            Call LogEvent("   " & SelectedItems.Count _
                          & " corresponding item(s) found for " _
                          & RestrictCriteriaString)
            
            For ix2 = 1 To reportedMatches
     
                sHdl = "CritPropName---------------" _
                       & " Quelle-" & ix1 & "----------------------" _
                       & " Ziel-" & ix2 & "------------------------" _
                       & " Comp-- Info-- Flag-- ign.parts1 ign.parts2"
                ' loop follows a synch as long as there's something left to do
                Call GetAobj(aPindex, ix2)
                If objTypName <> DecodeObjectClass(True) Then
                    GoTo nextOne
                End If
                AllDetails = vbNullString
                Call AddItemToList((ix1), fiMain(2), vbNullString, (ix2))
                Call LogEvent(String(60, "_") & vbCrLf & "Prüfung von " _
                              & Folder(1).FolderPath & ": " & ix1 & " und " _
                              & Folder(2).FolderPath & ": " & ix2 & ", " _
                              & MainObjectIdentification _
                              & ": " & vbCrLf & ix1 & ": " & fiMain(1))
                If fiMain(1) <> fiMain(2) Then
                    Call LogEvent(ix2 & ": " & fiMain(2))
                End If
            
                If ItemIdentity() Then
                    ' ==================================================================
                    ListContent(ListCount).Compares = "="
                    Call AskForUserAction(False, True)
                Else
                    ListContent(ListCount).Compares = "<>"
                    Call AskForUserAction(False, False)
                End If
                ListContent(ListCount).MatchData = MatchData
                ListContent(ListCount).DiffsRecognized = DiffsRecognized
nextOne:
            Next ix2
        Else
            Call LogEvent("   No corresponding item found for " _
                          & RestrictCriteriaString)
            Call AskForUserAction(True, False)
        End If
        If ActionID <> 6 Then
            frmDeltaList.Show
        End If
        Set ListContent = Nothing
    Next ix1
    
    Set sortedItems(1) = Nothing
    Set SelectedItems = New Collection

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.MergeItemsNewIntoOld

'---------------------------------------------------------------------------------------
' Method : Sub Pick2Folders
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Pick2Folders()
Dim zErr As cErr
Const zKey As String = "FolderOps.Pick2Folders"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim action As String
Dim afolder As Folder
    Call InitObjectSelection
    eOnlySelectedItems = False
    action = " und bestätigen Sie die Auswahl mit 'OK'"
    bDefaultButton = "Go"
    Call PickAFolder(1, _
                     "bitte wählen Sie den Ordner mit den neueren (=Quell) Objekten" _
                     & action, "Auswahl der zu vergleichenden Ordner", "OK", "Cancel")
    
    If topFolder Is Nothing Then
        Set topFolder = Folder(1)
    Else
    End If
    
    curFolderPath = Folder(1).FolderPath
    
    Set afolder = Folder(1)
    
    Call LogEvent("==== User has selected Source " & curFolderPath, eLall)
    Call PickAFolder(2, _
                     "bitte wählen Sie den Ordner mit den älteren (=Ziel) Objekten" _
                     & action, "Auswahl der zu vergleichenden Ordner", _
                     "OK", "Cancel")
    Call LogEvent("==== Target Folder is called  " _
                  & Folder(2).FolderPath, eLall)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.Pick2Folders

'---------------------------------------------------------------------------------------
' Method : Sub MakeSelection
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MakeSelection(i As Long, msg1 As String, Title As String, But1 As String, But2 As String)
Dim zErr As cErr
Const zKey As String = "FolderOps.MakeSelection"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aFolderpath As String

doAskForSelection:
    rsp = NonModalMsgBox(msg1, But1, But2, Title)
    Select Case rsp
        Case vbOK
            If ActiveExplorer.Selection.Count = 0 Then
                GoTo doAskForSelection
            End If
            If ActiveExplorer.CurrentFolder Is Nothing Then
                Set Folder(i) = aNameSpace.PickFolder
                GoTo doAskForSelection
            Else
                Set Folder(i) = ActiveExplorer.CurrentFolder
                ' DoVerify False, " watchout: volatile result, but sometimes this works, sometimes not"
                aFolderpath = Folder(i).FolderPath
                ' *** bug: setting Folder sets/ruins Folder(1) although i=2, both are set to same Folder
                If i > 1 Then
                    Set Folder(i) = GetFolderByName(aFolderpath, Folder(i)) ' get a non-volalite pointer
                End If
            End If
            SkipNextInteraction = True
        Case vbCancel
            SkipNextInteraction = False
            If But2 = "Cancel" Then
                Call LogEvent("=======> Keine Auswahl, Abbruch: " & Now(), eLnothing)
                End
            End If
    End Select
    Call LogEvent("=======> Es wurden " & ActiveExplorer.Selection.Count _
                  & " Items aus dem Ordner " & Quote(aFolderpath) _
                  & " gewählt. Time: " & Now(), eLnothing)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.MakeSelection

'---------------------------------------------------------------------------------------
' Method : Sub PickAFolder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PickAFolder(i As Long, msg1 As String, Title As String, But1 As String, But2 As String)
Dim zErr As cErr
Const zKey As String = "FolderOps.PickAFolder"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aFolderpath As String
    rsp = vbAbort
    If LF_UsrRqAtionId = atPostEingangsbearbeitungdurchführen Then
        If ActiveExplorer.CurrentFolder = SpecialSearchFolderName Then
            rsp = vbOK
        End If
    End If
    If rsp = vbAbort Then
doAskForFolder:
        rsp = NonModalMsgBox(msg1, But1, But2, Title)
    End If
    
    Select Case rsp
        Case vbOK
            ' *** bug: this sets/ruins Folder(1) although i=2, both are set to same Folder
            If ActiveExplorer.CurrentFolder Is Nothing Then
                Set Folder(i) = aNameSpace.PickFolder
            Else
                Set Folder(i) = ActiveExplorer.CurrentFolder
                If DebugMode Then
                    Debug.Print "watchout: volatile result, but sometimes this works, sometimes not"
                End If
                aFolderpath = Folder(i).FolderPath
                Set Folder(i) = GetFolderByName(aFolderpath, Folder(i), noSearchFolders:=True) ' get a non-volalite pointer
            End If
            If Folder(i) Is Nothing Then
                PickTopFolder = True
                GoTo doAskForFolder
            Else
                PickTopFolder = False
            End If
            SkipNextInteraction = True
        Case vbCancel
            SkipNextInteraction = False
            If But2 = "Cancel" Then
                Call LogEvent("=======> Kein Ordner gewählt, Abbruch: " & Now(), eLnothing)
                End
            End If
            Call LogEvent("=======> Walking all Folders. Time: " & Now(), eLnothing)
    End Select

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.PickAFolder

'---------------------------------------------------------------------------------------
' Method : Sub InitFolderOps
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitFolderOps(Optional SortingOnlySecond As Boolean)
Dim zErr As cErr
Const zKey As String = "FolderOps.InitFolderOps"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim Filter As String
    If Folder(1).DefaultItemType <> Folder(2).DefaultItemType Then
        rsp = NonModalMsgBox("Discrepancy in Folder Type", Folder(1).Name, _
                             Folder(2).Name, _
                             "Folder Types do not Match: Choose primary type")
        If rsp = vbOK Then
            Matchcode = 1
        ElseIf rsp = vbCancel Then
            Matchcode = 2
        Else
            If TerminateRun Then
                GoTo ProcReturn
            End If
        End If
    Else
        Matchcode = 1                              ' the Folder types are identical, use only first to determine GetObjectAttributess
    End If
    
    Call BestObjProps(Folder(Matchcode), withValues:=False)
Restart:
    If SortingOnlySecond Then
        If getFolderFilter(Folder(2).Items.Item(1), CutOffDate, _
                           Filter, ">=") Then
            Set sortedItems(1) = Folder(2).Items.Restrict(Filter)
        Else
            Set sortedItems(1) = Folder(2).Items
        End If
    Else
        If Folder(1).Items.Count = 0 Then
            Set sortedItems(1) = Folder(1).Items
        ElseIf getFolderFilter(Folder(1).Items.Item(1), CutOffDate, _
                               Filter, ">=") Then
            Set sortedItems(1) = Folder(1).Items.Restrict(Filter)
        Else
            Set sortedItems(1) = Folder(1).Items
        End If
    End If
    aBugTxt = "Sort Matches"
    Call Try                                       ' Try anything, autocatch
    sortedItems(1).sort SortMatches, False
    If Catch Then
        Message = vbCrLf & E_AppErr.Description _
                  & vbCrLf & vbCrLf & "Bitte ändern Sie die [Sortierparameter]"
        b1text = "Weiter"
        b2text = vbNullString
        Set FRM = New frmDelParms
        FRM.Show
        Set FRM = Nothing
        If rsp = vbCancel Then
            If TerminateRun Then
                GoTo ProcReturn
            End If
        End If
        GoTo Restart
    End If
    
    Ictr(1) = sortedItems(1).Count
    If Not SortingOnlySecond Then
        Set sortedItems(2) = Folder(2).Items
        Call Try                                   ' Try anything, autocatch
        sortedItems(2).sort SortMatches, False
        If Catch Then
            Message = vbCrLf & E_AppErr.Description & vbCrLf & vbCrLf _
                      & "Bitte ändern Sie die [Sortierparameter]"
            b1text = "Weiter"
            b2text = vbNullString
            Set FRM = Nothing
            Set FRM = New frmDelParms
            FRM.Show
            Set FRM = Nothing
            If rsp = vbCancel Then
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            End If
            GoTo Restart
        Else
            Call ErrReset(0)
            Message = "Sortierung erfolgreich." & vbCrLf _
                      & "Anzahl der enthaltenen Objekte:" & vbCrLf _
                      & "Ordner 1=" & Folder(1).FolderPath _
                      & ", Items: " & Ictr(1) _
                      & " Ordner 2=" & Folder(2).FolderPath _
                      & ", Items: " & sortedItems(2).Count & vbCrLf _
                      & "Bitte bestätigen Sie die Parameter für die Vergleichsoperationen"
            b1text = "Weiter"
            b2text = vbNullString
            b3text = "Cancel"
            Set FRM = Nothing
            Set FRM = New frmDelParms
            FRM.Show
            If DebugMode Then
                DoVerify False
            End If
            Set FRM = Nothing
            If rsp = vbCancel Then
                If TerminateRun Then
                    GoTo ProcReturn
                End If
            End If
        End If
        Ictr(2) = sortedItems(2).Count
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.InitFolderOps

'---------------------------------------------------------------------------------------
' Method : Sub DetailsToPrintFile
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DetailsToPrintFile(FileName As String)
Dim zErr As cErr
Const zKey As String = "FolderOps.DetailsToPrintFile"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    On Error GoTo ficl
    Open FileName For Output As #2
    Print #2, "Details zu: " _
             & MainObjectIdentification & ": " & fiMain(1)
    Print #2, vbCrLf & AllDetails
ficl:
    Close #2

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' FolderOps.DetailsToPrintFile

'---------------------------------------------------------------------------------------
' Method : Function quickCheck
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function quickCheck() As Long
Dim zErr As cErr
Const zKey As String = "FolderOps.quickCheck"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i1 As Long
Dim i2 As Long
Dim MainCompare As Long
    quickCheck = Passt_Synch
    
Synch:
    For i1 = WorkIndex(1) To Ictr(1)
        
        Call GetAobj(1, i1)
        objTypName = DecodeObjectClass(getValues:=False)
        
        For i2 = WorkIndex(2) To Ictr(2)
            Call GetAobj(1, i2)
            objTypName = DecodeObjectClass(getValues:=False)
            
            MainCompare = StrComp(fiMain(1), fiMain(2), vbTextCompare)
            If MainCompare = 0 Then                ' Identische strings
                quickCheck = Passt_Synch
                If ListCount > 0 Then              ' schon ein Eintrag da?
                    If LenB(ListContent(ListCount).Index2) = 0 _
                    And ListContent(ListCount).Index1 = i1 Then ' der linke Eintrag ist schon da: zusammenfassen
                        ListContent(ListCount).Index2 = i2 ' dann rechts mit diesem vergleichen
                        ListContent(ListCount).Compares = (Passt_Deleted) ' deleted, weil rechts i1 nicht gefunden wurde
                    ElseIf ListContent(ListCount).Index1 = vbNullString _
                        And ListContent(ListCount).Index2 = i2 Then ' der rechte Eintrag ist schon da: zusammenfassen
                            ListContent(ListCount).Index1 = i1 ' dann rechts mit diesem vergleichen
                            ListContent(ListCount).Compares = (Passt_Deleted) ' deleted, weil rechts i2 nicht gefunden wurde
                    Else
                        If ListContent(ListCount).Index1 <> i1 _
                        Or ListContent(ListCount).Index2 <> i2 Then ' Doublette möglich
                            Call AddItemToList((i1), fiMain(1), (Passt_Synch), (i2))
                        Else
                            DoVerify False, " wie geht das denn ?"
                        End If
                    End If
                Else
                    Call AddItemToList((i1), fiMain(1), (Passt_Synch), (i2))
                End If
                WorkIndex(1) = i1                  ' we continue here on next entry
                WorkIndex(2) = i2
                GoTo ProcReturn
            ElseIf MainCompare < 0 Then            ' links ist das kleinere, als erstes zufügen
                Call AddItemToList((i1), fiMain(1), (Passt_Deleted), vbNullString)
                WorkIndex(1) = i1
                WorkIndex(2) = i2
                ListContent(ListCount).DiffsRecognized = _
                                                       MainObjectIdentification _
                                                       & "(2): " & fiMain(2) & " kleiner als " & fiMain(1)
                GoTo ni1                           ' do not inc i2
            Else                                   ' fiMain(1) > fiMain(2)    : rechts ist das kleinere, als erstes zufügen
                Call AddItemToList(vbNullString, fiMain(2), (Passt_Inserted), (i2))
                WorkIndex(1) = i1
                WorkIndex(2) = i2
                ListContent(ListCount).DiffsRecognized = _
                                                       MainObjectIdentification _
                                                       & "(1): " & fiMain(1) & " größer als Rechts"
            End If
        Next i2
ni1:
    Next i1
    quickCheck = inv
    WorkIndex(1) = i1
    WorkIndex(2) = i2

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' FolderOps.quickCheck

'---------------------------------------------------------------------------------------
' Method : Function SetFilter
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function SetFilter(adName As String, AttrValue As String) As String
Dim zErr As cErr
Const zKey As String = "FolderOps.SetFilter"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If LenB(Trim(AttrValue)) > 0 Then
        SetFilter = adName & " = " & Quote(AttrValue)
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' FolderOps.SetFilter



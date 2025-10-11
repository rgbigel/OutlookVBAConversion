Attribute VB_Name = "RuleSetup"
Option Explicit

Public CategoryString As String
Public RuleTable As Collection
Public staticRuleTable As Boolean                ' if true, uses stored tabs,
' else get them from Excel
Public UseExcelRuleTable As Boolean              ' Dynamischer Excel-modus wenn True
                                
Public Const CategoryKeepList As String = "Unbekannt; "
Public CategoryDroplist As String
Public RulesExplained As String

Dim modCell As Long
Dim MatchMode As String

'---------------------------------------------------------------------------------------
' Method : Function BestRule
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function BestRule(Optional withIrule As Boolean = False) As cAllNameRules
Dim zErr As cErr
Const zKey As String = "RuleSetup.BestRule"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

    If withIrule And Not iRules Is Nothing Then
        Set BestRule = iRules
    ElseIf Not sRules Is Nothing Then
        Set BestRule = sRules
        Set iRules = Nothing
    Else
        Set BestRule = dftRule
        Set aID(1).idAttrDict = Nothing
        Set aID(2).idAttrDict = Nothing
        Set sRules = Nothing
        Set aTD = Nothing
        Set sDictionary = Nothing
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' RuleSetup.BestRule

'---------------------------------------------------------------------------------------
' Method : Sub CreateIRule
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub CreateIRule(newPropKey As String)
Dim zErr As cErr
Const zKey As String = "RuleSetup.CreateIRule"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim aDictItem As cAttrDsc

    If aTD Is Nothing Then
        GoTo make_new
    Else
        newPropKey = aTD.adKey
        If Not aTD.adRules Is Nothing Then
            If aTD.adKey = newPropKey Then
                If aTD.adRules.RuleInstanceValid Then
                    Call Get_iRules(aTD)
                    GoTo FuncExit                ' it is there already
                Else
                    GoTo use_old
                End If
            Else
                GoTo make_new
            End If
        ElseIf aTD.adKey <> newPropKey Then      ' wrong trail here, define again
            DoVerify False, "does this make sense in this version ***???"
            Set aTD = Nothing
            Call GetAttrKey(newPropKey, noget:=False)
            If aTD Is Nothing Then
                GoTo make_new
            End If
        ElseIf aTD.adNr = 0 Then
            AttributeIndex = -1                  ' no aID(1).idAttrDict yet
        Else
            aTD.adNr = AttributeIndex
            GoTo use_old
        End If
    End If
    If aTD.adisUserAttr <> isUserProperty Then   ' Misspelled Ucase/Lcase (!!!)
        DoVerify False
        GoTo make_new
    End If
    GoTo use_old
    
make_new:
    ' reached for some specials, like Seperator lines
    aCloneMode = FullCopy
    Set aTD = New cAttrDsc
        
    If aID(aPindex).idAttrDict Is Nothing Then   ' somone kills it???*** (Termination of temporary instance after .addItem).idAttrDict
        DoVerify False, " CRAP:  but there is an easy fix"
        Set aID(aPindex).idAttrDict = aID(aPindex).idAttrDict
        If aID(aPindex).idAttrDict Is Nothing Then DoVerify False, "double check, Debug.Assert False if not fixed"
    End If
    If aTD Is Nothing Then
        With aID(aPindex).idAttrDict
            If .Exists(newPropKey) Then          ' fix aTd
                Set aDictItem = .Item(newPropKey)
                Set aTD = aDictItem
                DoVerify PropertyNameX = aTD.adName
                GoTo use_old
            Else
                DoVerify False
            End If
        End With                                 ' aID(aPindex).idAttrDict
    End If
    
use_old:
    Set iRules = sRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=False)
    iRules.ARName = newPropKey                   ' !!! not PropertyNameX!!!
    Set aTD.adRules = iRules
    Call SplitDescriptor(aTD)                    ' determine Rules for this Attribute
    aTD.adRuleIsModified = False                 ' straight from sRules, but now RuleIsSpecific and RuleInstanceValid

FuncExit:
    Set aDictItem = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.CreateIRule

'---------------------------------------------------------------------------------------
' Method : Sub RulesToExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub RulesToExcel(px As Long, withAttributes As Boolean)
Dim zErr As cErr
Const zKey As String = "RuleSetup.adRulesToExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If withAttributes Then
        If xlApp Is Nothing Then                 ' no: put it there
            Call XlgetApp
        End If
        Call AttrDscs2Excel
        If xlApp Is Nothing Then
            GoTo FuncExit
        End If
    End If
    If SelectOnlyOne Then
        If Not aID(2) Is Nothing Then
            If Not aID(2).idAttrDict Is Nothing Then
                DoVerify False, "??? why two operands for this?"
            End If
        End If
        If xDeferExcel Then
            displayInExcel = True
            If xlApp Is Nothing Then             ' no: put it there
                Call XlgetApp
                Set O = xlWBInit(xlA, TemplateFile, _
                                 cOE_SheetName, sHdl, showWorkbook:=DebugMode)
            ElseIf O Is Nothing Then
                Set O = xlWBInit(xlA, TemplateFile, cOE_SheetName, sHdl, showWorkbook:=DebugMode)
            End If
            aOD(px).objDumpMade = 0
            xlApp.ScreenUpdating = False
            Call StckedAttrs2Xcel(O)
            xlApp.ScreenUpdating = True
        End If
    End If
    ' are we done with all existing properties?
    If aID(px).idAttrDict.Count >= TotalPropertyCount Then
        AllPropsDecoded = True
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.adRulesToExcel

'---------------------------------------------------------------------------------------
' Method : Sub Get_iRules
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Get_iRules(xTD As cAttrDsc)
Dim zErr As cErr
Const zKey As String = "RuleSetup.Get_iRules"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If iRules Is Nothing Then
        If xTD.adRules.RuleInstanceValid Then
            Set iRules = xTD.adRules
            GoTo couldMatch
        End If
        Set iRules = xTD.adRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=True)
    Else
couldMatch:
        If iRules.ARName <> xTD.adRules.ARName Then ' wrong one:
            If xTD.adRules.RuleType = "InstanceRule" Then
                Set iRules = xTD.adRules
            Else                                 ' create fresh instance rule
                Set iRules = xTD.adRules.AllRulesClone(InstanceRule, aObjDsc, withMatchBits:=True)
                iRules.RuleIsSpecific = True
                iRules.RuleInstanceValid = True
            End If
        End If
    End If
    Call GetRuleBits(xTD)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.Get_iRules

'---------------------------------------------------------------------------------------
' Method : GetRuleBits
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: set the iRule Bits including AllPublic.iRuleBits
'---------------------------------------------------------------------------------------
Sub GetRuleBits(xTD As cAttrDsc)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "RuleSetup.Get_iRules"
    #If MoreDiagnostics Then
        Call DoCall(zKey, "Sub", eQzMode)
    #End If

    If iRules.RuleInstanceValid Then
        If LenB(xTD.adRuleBits) = 0 Or Left(xTD.adRuleBits, 1) = "(" Then
            iRuleBits = xTD.adName & _
                        " E:" & Left(iRules.RuleIsSpecific, 1) & _
                        " M:" & Left(iRules.clsObligMatches.RuleMatches, 1) & _
                        " D:" & Left(iRules.clsNeverCompare.RuleMatches, 1) & _
                        " N:" & Left(iRules.clsNotDecodable.RuleMatches, 1) & _
                        " S:" & Left(iRules.clsSimilarities.RuleMatches, 1)
            xTD.adRuleBits = iRuleBits
        Else
            iRuleBits = xTD.adRuleBits
        End If
    Else
        iRuleBits = "(not val.)"
        xTD.adRuleBits = iRuleBits
    End If

zExit:
    Call DoExit(zKey)
End Sub                                          ' RuleSetup.GetRuleBits

'---------------------------------------------------------------------------------------
' Method : Function IsAMandatory
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsAMandatory() As Boolean
Dim zErr As cErr
Const zKey As String = "RuleSetup.IsAMandatory"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
    ' compare full word Match, wildcards or ":" will not work
    i = InStr(b & Trim(sRules.clsObligMatches.aRuleString) & b _
            & ExtendedAttributeList & b, _
              b & PropertyNameX & b)
    IsAMandatory = (i > 0)
    If DebugLogging Then
        Debug.Print "MandatoryAttribute", PropertyNameX, IsAMandatory
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' RuleSetup.IsAMandatory

'---------------------------------------------------------------------------------------
' Method : Sub SplitDescriptor
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SplitDescriptor(xTD As cAttrDsc)
Dim zErr As cErr
Const zKey As String = "RuleSetup.SplitDescriptor"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim LookupName As String

    If xTD Is Nothing Then
        GoTo ProcReturn                          ' no can do at this point
    End If
    If xTD.adRules Is Nothing Then
        GoTo ProcReturn                          ' Seperator Lines ========= etc
    End If
    If xTD.adRules.RuleInstanceValid Then        ' if yes, did split already
        If xTD.adName <> PropertyNameX Then      ' but for wrong one...
            LookupName = xTD.adName
            GoTo getit                           ' not existing, may need to create it
        End If
        Call Get_iRules(xTD)                     ' use results of previous splitDescriptor
        GoTo ProcReturn                          ' already consistent
    Else
        If xTD.adName <> PropertyNameX Then      ' valid, but the wrong one
            If LenB(PropertyNameX) = 0 Then
                If DebugMode Then
                    DoVerify False, "check if aID(aPindex).odItemDict ok? --> fixed next"
                End If
                LookupName = xTD.adName
            End If
getit:
            Set xTD = GetAttrDsc(LookupName, Get_aTD:=False)
            If xTD Is Nothing Then
                GoTo ProcReturn                  ' no can do at this point
            End If
            PropertyNameX = LookupName
        Else
            Call Get_iRules(xTD)
        End If
    End If
    
    With iRules
        If (.clsObligMatches.bConsistent _
        And .clsSimilarities.bConsistent _
        And .clsNotDecodable.bConsistent _
        And .clsNeverCompare.bConsistent) Then
            Call iRules.CheckAllRules(PropertyNameX, "->") ' re-use value settings
        Else
            Call iRules.CheckAllRules(PropertyNameX, vbNullString) ' original value setting
        End If
        
        ' check logic and set result of logic as message
        If .ARName <> vbNullString And .clsNotDecodable.RuleMatches Then ' unMatchable because undecodable
            IgString = "  --non-decodable prop."
        End If
        
        If .RuleIsSpecific Then                  ' some rule explicitly defined
            If .clsNotDecodable.RuleMatches Then
                .clsNeverCompare.RuleMatches = True
                .clsSimilarities.RuleMatches = False
                .clsObligMatches.RuleMatches = False
                If LenB(Trim(IgString)) = 0 Then
                    IgString = "  (will not decode or compare) "
                End If
            ElseIf .clsObligMatches.RuleMatches Then
                .clsNeverCompare.RuleMatches = False
                .clsSimilarities.RuleMatches = False
            ElseIf .clsNeverCompare.RuleMatches Then
                ' although not compared, first decode it
                IgString = "   --dontcompare: "
            End If
                
            If .clsSimilarities.RuleMatches Then
                If .clsNeverCompare.RuleMatches Then
                    .clsNeverCompare.RuleMatches = False
                    IgString = "  ignoring don't compare, check if similar: "
                End If
            End If
            
            If LenB(TrueCritList) > 0 Then       ' we do not want any re-ordered attributes
                If .clsObligMatches.RuleMatches _
                Or .clsSimilarities.RuleMatches Then
                    Call AppendTo(SelectedAttributes, PropertyNameX, b)
                End If
            End If
        End If
    End With                                     ' iRules

FuncExit:
    xTD.adRules.RuleInstanceValid = True
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.SplitDescriptor

'---------------------------------------------------------------------------------------
' Method : Sub SplitMandatories
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub SplitMandatories(ByVal MatchRequest As String, Optional ByRef MandatoryWorkRule As cNameRule)
Dim zErr As cErr
Const zKey As String = "RuleSetup.SplitMandatories"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                        & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcReturn
    End If
    Recursive = True                             ' restored by    Recursive = False ProcReturn:

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    ' also creates sort criteria and sets TrueCritList - string
Dim i As Long
Dim j As Long
Dim Ci As Long
Dim PropName As Variant
Dim NoSkipSortAttrs As Boolean

    'On Error GoTo 0
    NoSkipSortAttrs = True
    
    If sRules Is Nothing Then
        If MandatoryWorkRule Is Nothing Then
            Set sRules = dftRule.AllRulesClone(ClassRules, aObjDsc, False)
            Set MandatoryWorkRule = sRules.clsObligMatches
        Else
            Set sRules = MandatoryWorkRule.PropAllRules
            If Not sRules Is Nothing Then
                GoTo GotIt
            End If
        End If
    Else
GotIt:
        If LenB(Trim(sRules.clsObligMatches.aRuleString)) = 0 Then
            j = 0
        Else
            Set MandatoryWorkRule = sRules.clsObligMatches
            If Not sRules.RuleInstanceValid Then
                GoTo noRebuildFromRule
            End If
            For i = LBound(sRules.clsObligMatches.CleanMatches) To LBound(sRules.clsObligMatches.CleanMatches)
                j = InStr(MatchRequest, sRules.clsObligMatches.CleanMatches(i))
                If j = 0 Then                    ' invalid list, (some) Matchrequests in sRules irregular
                    Exit For
                End If
            Next i
        End If                                   ' .aRuleString ? vbNullString
        If j > 0 Then                            ' no override wanted
            MatchRequest = Trim(sRules.clsObligMatches.aRuleString)
        End If
    End If
noRebuildFromRule:
    If ModObligMatches(MatchRequest) Then        ' Will always set sRules
        If MandatoryWorkRule Is Nothing Then
            Set MandatoryWorkRule = sRules.clsObligMatches
        End If
        SortMatches = vbNullString
        ExtendedAttributeList = vbNullString
        j = 0
        i = 0
        With MandatoryWorkRule
            .CleanMatchesString = vbNullString
            Ci = UBound(.CleanMatches)
            For Each PropName In .CleanMatches
                If LenB(PropName) > 0 _
                And LCase(PropName) <> "or" _
                And LCase(PropName) <> "and" _
                And LCase(PropName) <> "not" Then ' remove empty or operators
                    If i <> j Then
                        MandatoryWorkRule.CleanMatches(j) = PropName ' pull down to correct position
                        Ci = Ci - 1              ' cut end of array
                    End If
                    .CleanMatchesString = .CleanMatchesString & b & PropName
                    j = j + 1                    ' copy into this position when true CritPropName NEXT time
                    ' find out if we want to Debug.Assert False building objSortMatches
                    If InStr(.MatchesList(i), "|") > 0 Then
                        NoSkipSortAttrs = False  ' stays this way for this call
                    End If
                    If NoSkipSortAttrs Then
                        If InStr(ExtendedAttributeList & b, .CleanMatches(i) & b) = 0 Then
                            SortMatches = SortMatches & "[" & PropName & "] "
                        End If
                    Else
                        ExtendedAttributeList = ExtendedAttributeList & b & PropName
                    End If
                Else
                    j = j                        ' no step of target array
                End If
                i = i + 1
            Next PropName
            ' correction of element count in TrueImportantProperties
            If j < 1 Then
                Erase TrueImportantProperties
            ElseIf Ci < UBound(.CleanMatches) Then
                ' a vbNullString Property was removed
                ReDim Preserve TrueImportantProperties(j - 1)
            End If
            .CleanMatchesString = Trim(.CleanMatchesString)
            ExtendedAttributeList = Trim(ExtendedAttributeList)
            TrueCritList = Trim(.aRuleString)
            MostImportantProperties = .CleanMatches
            MostImportantAttributes = .CleanMatchesString
            MainObjectIdentification = .CleanMatches(0)
        End With                                 ' MandatoryWorkRule
    End If
    If MandatoryWorkRule Is Nothing Then
        MostImportantProperties = Array(vbNullString)
        MostImportantAttributes = vbNullString
    Else
        MostImportantProperties = MandatoryWorkRule.CleanMatches
        MostImportantAttributes = MandatoryWorkRule.CleanMatchesString
        If Not isEmpty(MandatoryWorkRule.CleanMatches) Then
            MainObjectIdentification = MandatoryWorkRule.CleanMatches(0)
        End If
    End If
    SortMatches = Trim(SortMatches)
    Call AppendTo(SortMatches, "[LastModificationTime]", b)
    sRules.RuleInstanceValid = True

FuncExit:
    Recursive = False
ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                          ' RuleSetup.SplitMandatories

'---------------------------------------------------------------------------------------
' Method : Sub StckedAttrs2Xcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub StckedAttrs2Xcel(aTab As cXLTab)
Dim zErr As cErr
Const zKey As String = "RuleSetup.StckedAttrs2Xcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim px As Long
Dim DidIndex As Long
Dim i As Long
Dim FilterOff As Boolean

    If aOD(0).objDumpMade >= 2 Then
        GoTo ProcReturn
    End If
    
    aOD(0).objDumpMade = 0
    If xlApp.ActiveSheet.Name <> aTab.xlTName Then
        Call XlopenObjAttrSheet(xlA)
    End If
    aTab.xlTSheet.Activate
    
noquickclear:
    'On Error GoTo 0
    aTab.xHdl = sHdl                             ' sets first line as headline ( property let )
    If Not (DebugMode Or DebugLogging) Then
        xlApp.Visible = False
    End If
    xlApp.Cursor = xlWait
    If SelectOnlyOne Then
        FilterOff = True
        If WorkItemMod(1) Then
            i = 2
        Else
            i = 1
        End If
    Else
        i = 2
    End If
    
    For px = aOD(0).objDumpMade + 1 To i
        If Not aID(px).idAttrDict Is Nothing Then
            If aID(px).idAttrDict.Count > 0 Then
                Call StckedAttrLoop(px)
                DidIndex = px
                If aID(px) Is Nothing Then
                Else
                    aOD(px).objDumpMade = px
                End If
            End If
        End If
    Next px
    px = DidIndex
    aOD(0).objDumpMade = px
    If FilterOff Then
        aTab.xlTSheet.Range("$A$1").AutoFilter
    Else
        aTab.xlTSheet.Range("$A$1:$H$1").AutoFilter _
        Field:=6, Criteria1:="="
    End If
    aTab.xlTSheet.Cells(2, 1).Select

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.StckedAttrs2Xcel

'---------------------------------------------------------------------------------------
' Method : Sub StckedAttrLoop
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub StckedAttrLoop(px As Long)
Dim zErr As cErr
Const zKey As String = "RuleSetup.StckedAttrLoop"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim N As Long
    aPindex = px
    If xlApp Is Nothing Then
        GoTo ProcReturn
    End If
    N = 1
    For i = 1 To aID(px).idAttrDict.Count
        AttributeIndex = i
        Set aTD = aID(px).idAttrDict.Items(i)
        If OnlyMostImportantProperties Then
            If InStr(SelectedAttributes & b, aTD.adKey & b) = 0 Then
                GoTo nextInLoop
            End If
            N = N + 1
            Call put2IntoExcel(px, N)
        End If
        If N = 1 Or N Mod 10 = 0 Then
            If Not ShutUpMode Then
                Debug.Print Format(Timer, "0#####.00") & vbTab & px, _
                                                       "inserting attribute # " & i _
                                                     & " into Sheet " & W.xlTName
            End If
        End If
nextInLoop:
    Next i
    
    Debug.Print Format(Timer, "0#####.00") & vbTab & px, _
                                           "last cAttrDsc = " & N & " into Sheet " & W.xlTName

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.StckedAttrLoop

'---------------------------------------------------------------------------------------
' Method : Sub ChkCatLogic
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub ChkCatLogic()
Dim zErr As cErr
Const zKey As String = "RuleSetup.ChkCatLogic"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim temp As Boolean
    With CurIterationSwitches
        temp = .ReProcessDontAsk _
             Or (.ReprocessLOGGEDItems And Not eOnlySelectedItems)
        If temp <> .ReProcessDontAsk Then
            .ReProcessDontAsk = temp
        End If
    End With                                     ' CuriterationSwitches

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.ChkCatLogic

'---------------------------------------------------------------------------------------
' Method : Function CategorizeItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CategorizeItem(ByVal curItem As Object) As String
Dim zErr As cErr
Const zKey As String = "RuleSetup.CategorizeItem"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim i As Long
Dim pos As Long
Dim wLen As Long
Dim ca As Variant
Dim aAttributeValue As String
Dim Matched As Boolean
Dim Prefix As String
Dim categoryKeepV As Variant
Dim CategoryCandidates As String
Dim WordStart As Boolean
Dim WordEnd As Boolean

    CategoryCandidates = vbNullString
    aAttributeValue = curItem.Categories
    If Not CurIterationSwitches.ResetCategories Then
        If isEmpty(categoryKeepV) Then           ' never dropped
            categoryKeepV = split(CategoryKeepList, "; ")
        End If
        If Not isEmpty(categoryKeepV) Then
            For i = 0 To UBound(categoryKeepV)
                Prefix = Trim(categoryKeepV(i))
                If Prefix <> vbNullString And InStr(aAttributeValue, Prefix) > 0 Then
                    Call AppendTo(CategoryCandidates, Prefix, ";")
                End If
            Next i
        End If
    End If
    Prefix = "     "
    CategorizeItem = vbNullString
    
    For i = 1 To RuleTable.Count
        aAttributeValue = vbNullString
        Matched = False
        Set ca = RuleTable.Item(i)
        WordStart = False
        WordEnd = False
        
        With ca
            If LenB(.typ) = 0 Then
                Exit For                         ' invalid, terminate loop
            End If
            '           .typ                       ' col  1
            '           .checkitem                 '      2
            '           .category                  '      3
            '           .final                     '      4
            '           .exact                     '      5
            '           .never                     '      6
            '           .pattern                   '      7
 
            wLen = Len(.checkitem)
            MatchMode = vbNullString                       ' 4 states possible:
            If .final Then
                MatchMode = "="
            End If
            If .never Then
                MatchMode = "-"
            End If
            If .Exact Then
                MatchMode = "*"
            End If
            If Not .pattern Then
                MatchMode = "="
            End If
        
            Select Case LCase(.typ)
            Case "subject"
                aAttributeValue = UCase(curItem.Subject)
            Case "sender"
                If aObjDsc.objHasSenderName Then
                    aAttributeValue = UCase(curItem.SenderEmailAddress)
                ElseIf aObjDsc.objHasSentOnBehalfOf Then
                    aAttributeValue = UCase(curItem.SenderName & b & curItem.SentOnBehalfOfName)
                Else
                    aAttributeValue = "# anonymous " & ca.typ
                End If
            Case "body"
                If Not aObjDsc.objHasHtmlBodyFlag Then
                    GoTo noHTML                  ' some Classes do not have that, so body is always text
                End If
                If curItem.BodyFormat <> olFormatHTML Then
noHTML:
                    Call ErrReset(0)
                    If Len(curItem.Body) > 300000 Then
                        Debug.Assert False
                    End If
                    aAttributeValue = UCase(curItem.Body)
                Else
                    If Len(curItem.HTMLBody) > 500000 Then
                        'Debug.Assert False      #### Switch this on?
                    End If
                    Call Try                     ' Try anything, autocatch, Err.Clear
                    aAttributeValue = UCase(curItem.HTMLBody)
                    If Catch Then
                        GoTo noHTML
                    End If
                    Call ErrReset(0)
                End If
            Case Else
                DoVerify False, " not implemented as a rule"
            End Select
            
            pos = InStr(aAttributeValue, UCase(.checkitem))
            
            If pos > 0 Then                      ' if found at all
                If pos = 1 Or .pattern Then
                    WordStart = True
                Else
                    If InStr(WordSep, Mid(aAttributeValue, pos - 1, 1)) > 0 Then
                        WordStart = True         ' WordSep preceedes it
                    ElseIf Asc(Mid(aAttributeValue, pos - 1, 1)) < Asc(b) Then
                        WordStart = True         ' match on nonprintable start seperators
                    End If
                End If
                If wLen <= Len(aAttributeValue) Then
                    If InStr(WordSep, Mid(aAttributeValue, pos + wLen, 1)) > 0 Then
                        WordEnd = True           ' WordSep follows it
                    ElseIf Asc(Mid(aAttributeValue, pos + wLen, 1)) < Asc(b) Then
                        WordEnd = True           ' match on nonprintable end seperators
                    End If
                End If
            
                If WordStart And WordEnd Then
                    If CurIterationSwitches.CategoryConfirmation Then
                        RulesExplained = Append(RulesExplained, TypeName(curItem) & _
                                                " chosen Category: " & .category & _
                                                " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
                        Call AppendTo(CategoryCandidates, .category, "; ")
                    Else                         ' complex rules
                        Select Case (MatchMode)
                        Case "-"                 ' if .never then
                            RulesExplained = Append(RulesExplained, TypeName(curItem) & _
                                                    " can not be Category " & .category & _
                                                    " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
                            Call AppendTo(CategoryDroplist, .category, "; ")
                            CategoryCandidates = StringRemove(CategoryCandidates, _
                                                              CategoryDroplist, "; ")
                        Case vbNullString                  ' Category is not uniquely set (may have several)
                            RulesExplained = Append(RulesExplained, TypeName(curItem) & _
                                                    " fits Category " & .category & _
                                                    " due to tag '" & .typ & "' " & .checkitem) & vbCrLf
                            Call AppendTo(CategoryCandidates, .category, "; ")
                        Case "="                 ' Category is final (multiple) + LOGGED always)
                            RulesExplained = Append(RulesExplained, TypeName(curItem) _
                                                  & " given final categories " & .category _
                                                  & " due to tag '" & .typ & "' " & .checkitem)
                            If LenB(CategoryCandidates) > 0 Then
                                RulesExplained = RulesExplained & ", previous categories " _
                                               & Quote(CategoryCandidates) & " kept" & vbCrLf
                            Else
                                RulesExplained = RulesExplained & vbCrLf
                            End If
                            Call AppendTo(CategoryCandidates, .category, "; ")
                            GoTo FunExit
                        Case Else                ' if .exact; should be same as "*", final unique cat.
                            RulesExplained = Append(RulesExplained, TypeName(curItem) _
                                                  & " given unique categories " & .category _
                                                  & " due to tag '" & .typ & "' " & .checkitem)
                            If LenB(CategoryCandidates) > 0 Then
                                RulesExplained = RulesExplained & ", previous categories " _
                                               & StringRemove(CategoryCandidates, .category, "; ") _
                                               & " dropped" & vbCrLf
                            Else
                                RulesExplained = RulesExplained & vbCrLf
                            End If
                            CategoryCandidates = .checkitem
                            GoTo FunExit
                        End Select
                    End If
                Else
                    RulesExplained = RulesExplained & Prefix & TypeName(curItem) _
                                & " has no word-Match in " & .typ & " for " _
                                & .checkitem & ": " & Quote(Mid(aAttributeValue, pos, wLen + 2)) _
                                & vbCrLf
                End If
            Else
                RulesExplained = RulesExplained & Prefix & TypeName(curItem) _
                            & " has no Match in " & .typ & " for " _
                            & .checkitem & ": " & Quote(Mid(aAttributeValue, 1, wLen + 2) & "... ") _
                            & vbCrLf
            End If
        End With                                 ' ca
    Next i

FunExit:
    If CurIterationSwitches.CategoryConfirmation Then
        Call LogEvent("Category Candidates are: " & CategoryCandidates)
    Else
        If LenB(CategoryCandidates) = 0 Then
            Call LogEvent(Prefix & TypeName(curItem) & _
                          " not assigned specific Category, no tags Matched", eLmin)
        Else
            Call LogEvent(Prefix & TypeName(curItem) & _
                          " was assigned to the following categories: " _
                        & CategoryCandidates)
        End If
    End If
    CategorizeItem = CategoryCandidates
    If ShowFunctionValues Then
        Call LogEvent(Prefix & RulesExplained)
    End If

FuncExit:
    RulesExplained = vbNullString
    Call ErrReset(4)

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' RuleSetup.CategorizeItem

'---------------------------------------------------------------------------------------
' Method : Function DetectCategory
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DetectCategory(TargetFolder As Folder, curItem As Object, ByRef OldCategory As String) As String
Dim zErr As cErr
Const zKey As String = "RuleSetup.DetectCategory"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ShortName As String
Dim XlwasOpenedHere As Boolean
Dim editmode As Boolean
Dim NewCategory As String
Dim tempCategory As String
Dim SourceFolder As Folder
Dim SourceFolderPath As String

    RulesExplained = vbNullString
    If RuleTable Is Nothing Then
        Set RuleTable = New Collection
    End If
    Set SourceFolder = curItem.Parent
    SourceFolderPath = SourceFolder.FullFolderPath

ReEstablishCategoryRules:
    If CurIterationSwitches.ResetCategories Then
        NewCategory = vbNullString
    Else
        NewCategory = OldCategory
    End If
    If RuleTable.Count = 0 Then
        If staticRuleTable And Not UseExcelRuleTable Then
            Call InitRuleTable
        Else
            GoTo getOrEdittags
        End If
    Else
        If UseExcelRuleTable Then
getOrEdittags:
            If OpenRuleTable(XlwasOpenedHere) Then
                Call Xl2RuleTable(editmode)      ' edit or not depends on Z_AppEntry
                Call GotRuleTabXl(XlwasOpenedHere)
            End If
        Else
            ' can use the existing static rule table
        End If
    End If
    
    OldCategory = curItem.Categories
    MatchMode = vbNullString
    
    If aObjDsc.objIsMailLike Then
        ShortName = Left(TargetFolder.Name, 4)
        If ShortName = "Junk" _
        Or ShortName = "Spam" _
        Or ShortName = "Uner" _
           Then
            NewCategory = "Junk"
            Call LogEvent("Mail assigned OldCategory " & NewCategory _
                        & " due to Folder name " & TargetFolder.FullFolderPath, eLmin)
        Else
            If CurIterationSwitches.ResetCategories _
            Or InStr(OldCategory, LOGGED) = 0 Then
                tempCategory = CategorizeItem(curItem)
                If CurIterationSwitches.ResetCategories Then
                    NewCategory = tempCategory
                ElseIf LenB(tempCategory) > 0 Then ' keep the existing value if no Rule
                    NewCategory = tempCategory
                    Call AppendTo(NewCategory, curItem.Categories, "; ")
                End If
            Else                                 ' default: unchanged by rule
                NewCategory = curItem.Categories
            End If
        End If
    End If
OKOK:
    If InStr(1, NewCategory, Unbekannt, vbTextCompare) > 0 Then
        GoTo unk
    ElseIf Not CurIterationSwitches.ResetCategories Then ' reset Cat ==> may correct source folder
        If InStr(1, SourceFolderPath, Unbekannt, vbTextCompare) > 0 Then
            GoTo unk1
        End If
    Else                                         ' test on unknown email
        If InStr(curItem.Parent.FullFolderPath, "SMS") > 0 Then
            Call AppendTo(NewCategory, "SMS", "; ")
        End If
        If Not aObjDsc.objHasSenderName Then
            GoTo unk1
        End If
        If IsUnkContact(curItem.SenderEmailAddress) Then
unk1:
            Call AppendTo(NewCategory, Unbekannt, "; ")
unk:
            If FolderUnknown Is Nothing Then
                Set TargetFolder = SourceFolder  ' not moving anything
            Else
                Set TargetFolder = FolderUnknown
            End If
        Else
            If InStr(1, SourceFolderPath, "BACKUP", vbTextCompare) > 0 Then
                If CurIterationSwitches.ResetCategories Then
                    Set TargetFolder = FolderInbox ' no longer unknown sender
                Else
                    Set TargetFolder = SourceFolder ' remain in Folder (do not move)
                End If
            Else
                If TargetFolder Is Nothing Then
                    Set TargetFolder = FolderInbox
                End If
            End If
        End If
    End If
    NewCategory = Append(NewCategory, LOGGED, "; ", ToFront:=True)
    If CurIterationSwitches.ReProcessDontAsk Then ' user interaction has been turned off
        NewCategory = NewCategory
        GoTo FunExit
    End If
    If CurIterationSwitches.CategoryConfirmation Then
        If LenB(RulesExplained) = 0 Then
            RulesExplained = "Keine Regel gefunden => keine besondere Kategorie"
        End If
        On Error GoTo FunExit
        frmStrEdit.Caption = curItem.Subject
        frmStrEdit.chSaveItemRequested = CurIterationSwitches.SaveItemRequested
        frmStrEdit.StringModifierCancelLabel.Caption = "alt:"
        frmStrEdit.StringModifierCancelValue.Text = OldCategory
        frmStrEdit.StringModifierExpectation = _
                                             "Die folgenden Kategorien sind u.W2. für dieses " _
                                           & TypeName(curItem) _
                                           & " geeignet. Bitte prüfen und ggf. korrigieren."
        frmStrEdit.StringToConfirm = NewCategory
        frmStrEdit.Explanations = RulesExplained
        CurIterationSwitches.CategoryConfirmation = False ' once assumed remember globally
        frmStrEdit.Show
        rsp = frmStrEdit.StringModifierRsp
        Select Case rsp
        Case vbOK                                ' user closed this and/or did not answer...
            If frmStrEdit.CategoryConfirmation And Not frmStrEdit.ReProcessDontAsk Then
                RulesExplained = "<=== Categorie-Regeln werden erneut angewendet ===>"
                GoTo ReEstablishCategoryRules
            End If
            NewCategory = frmStrEdit.StringToConfirm
        Case vbYes                               ' want to edit the rules
            editmode = True
            frmMaintenance.someAction = 1
            GoTo getOrEdittags
        Case vbNo
            NewCategory = OldCategory            ' not changing anything and no save
        Case Else                                ' no, cancel, retry: user closed this and/or did not answer...
            Call TerminateRun
            End
        End Select                               ' rsp
    End If
FunExit:
    Set frmStrEdit = Nothing
    If NewCategory <> OldCategory Then
        DetectCategory = NewCategory
        MailModified = True
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' RuleSetup.DetectCategory

' Achtung Generierter Code, nach frmMaintenance ersetzen
Sub InitRuleTable()
Dim zErr As cErr
Const zKey As String = "RuleSetup.InitRuleTable"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Set RuleTable = New Collection
    Call SetRuleTabFlds("Subject", "empfie", "Newsflash", True, False, False, True)
    Call SetRuleTabFlds("Subject", "statusreport", "Log", False, True, False, True)
    Call SetRuleTabFlds("Body", "status report", "Log", False, True, False, True)
    Call SetRuleTabFlds("Sender", "Stern", "Stern", False, True, False, False)
    Call SetRuleTabFlds("Sender", "Save.TV", "SaveTV", True, False, False, False)
    Call SetRuleTabFlds("Sender", "redaktion@JungeWelt", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Sender", "support", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Sender", "infomail", "Newsflash", False, True, False, True)
    Call SetRuleTabFlds("Sender", "jaxx", "Lotto", False, True, False, False)
    Call SetRuleTabFlds("Sender", "MS", "MS", False, False, False, False)
    Call SetRuleTabFlds("Sender", "bounce", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Sender", "news", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Sender", "@facebookmail.com", "Newsflash", True, False, False, True)
    Call SetRuleTabFlds("Sender", "one4all", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Sender", "wissenschaft-online", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Sender", "Sparaktion", "Newsflash", True, False, False, True)
    Call SetRuleTabFlds("Sender", "eilmeldung", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Sender", "oxygen", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Body", "newsletter regist", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Body", "newsletter", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Sender", "orders", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Sender", "payment@paypal", "Rechnung", False, True, False, False)
    Call SetRuleTabFlds("Body", "MS", "MS", True, False, False, False)
    Call SetRuleTabFlds("Sender", "Lockergnome", "Lockergome", False, True, False, False)
    Call SetRuleTabFlds("Body", "Lockergnome", "Lockergome", False, False, False, True)
    Call SetRuleTabFlds("Body", "Fritz!Box", "FritzBox", True, False, False, False)
    Call SetRuleTabFlds("Body", "Redmond", "MS", True, False, False, False)
    Call SetRuleTabFlds("Body", "codeproject", "MS", True, False, False, False)
    Call SetRuleTabFlds("Body", "SmartTools", "SmartTools", False, False, False, False)
    Call SetRuleTabFlds("Body", "Order Confirmation", "MailOrder", False, True, False, False)
    Call SetRuleTabFlds("Body", "Ausgangsbestätigung", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Body", "Rücksendung", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Body", "Rechnung", "Rechnung", False, False, False, False)
    Call SetRuleTabFlds("Body", "payment", "Rechnung", True, False, False, True)
    Call SetRuleTabFlds("Sender", "anja.weber", "newsflash", False, True, False, False)
    Call SetRuleTabFlds("Body", "eur statt", "newsflash", False, False, False, False)
    Call SetRuleTabFlds("Body", "Bestellung", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Body", "key", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "licen", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "Anmeldung", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "Zugang", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "lizen", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "Auftrag", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Body", "account", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "passwor", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "Anmeld", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Body", "Abmeld", "Saved Mail", False, False, False, False)
    Call SetRuleTabFlds("Sender", "pearl", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Sender", "digitalo", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Sender", "valentins", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Sender", "Oxygen3", "Junk", False, False, False, False)
    Call SetRuleTabFlds("Sender", "kundenservice", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Sender", "vmwareteam", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Sender", "SmartTools", "SmartTools", False, False, False, False)
    Call SetRuleTabFlds("Body", "Angebot", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Body", "Kauf", "MailOrder", False, False, False, False)
    Call SetRuleTabFlds("Sender", "supp0rt", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Sender", "Bercht", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Sender", "Igel-Soft", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Body", "Newsgroup", "Groups", False, False, False, False)
    Call SetRuleTabFlds("Body", "web.de informiert", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Body", "News", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Body", "event", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Body", "Angebote", "Newsflash", False, False, False, True)
    Call SetRuleTabFlds("Body", "Sparen", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Body", "Sparbrief", "Junk", True, False, False, False)
    Call SetRuleTabFlds("Body", "web.de gmbh", "Newsflash", False, True, False, False)
    Call SetRuleTabFlds("Body", "n e w s", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Body", "Gutschein", "Newsflash", False, False, False, False)
    Call SetRuleTabFlds("Body", "linke zeitung", "Newsflash", False, True, False, True)
    Call SetRuleTabFlds("Body", "cxtreme", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Body", "conrad", "Newsflash", True, False, False, False)
    Call SetRuleTabFlds("Body", "Lockergnome", "Lockergnome", True, False, False, False)
    Call SetRuleTabFlds("Body", "Windows Fanatics", "Lockergnome", True, False, False, False)
    Call SetRuleTabFlds("Body", "Stern", "Stern", True, False, False, False)
    Call SetRuleTabFlds("Body", "Jubiläum", "Junk", False, True, False, False)
    Call SetRuleTabFlds("Body", "MSN Groups", "Groups", False, False, False, False)
    Call SetRuleTabFlds("Body", "bercht", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Body", "Marc", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Body", "Saskia", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Body", "Franzi", "Rolf", False, False, False, False)
    Call SetRuleTabFlds("Body", "regist", "Saved Mail", False, False, False, False)
staticRuleTable = True
    UseExcelRuleTable = False                    ' execute sub InitRuleTable

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.InitRuleTable

'---------------------------------------------------------------------------------------
' Method : Function OpenRuleTable
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function OpenRuleTable(XlwasOpenedHere As Boolean) As Boolean
Dim zErr As cErr
Const zKey As String = "RuleSetup.OpenRuleTable"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim ntag(1 To 7) As Variant
Dim aHdl As String

    aHdl = "Typ------------ CheckItem----------- Category---------- Final--- Exact--- Never--- Pattern-"
    modCell = UBound(ntag) + 1                   ' this cell will Show if we changed anything
    Set E = xlWBInit(xlA, TemplateFile, "RuleCategoriesTable", aHdl, showWorkbook:=DebugMode)
    XlwasOpenedHere = XlOpenedHere And Not (xUseExcel Or xDeferExcel)
    If DebugMode Then
        Call DisplayExcel(E, relevant_only:=False, _
                          EnableEvents:=True, xlY:=xlA)
    End If
    E.xlTSheet.Activate
    Call GetLine(1, ntag)
    If InStr(aHdl, ntag(1)) <> 1 Then
        OpenRuleTable = False                    ' incorrect Headline, do not read this RuleCategoriesTable
    Else
        OpenRuleTable = True
    End If
    If LCase(E.xlTSheet.Cells(1, modCell)) = "modified" Then
        UseExcelRuleTable = True                 ' and use them from now on because inline code can not be changed here
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' RuleSetup.OpenRuleTable

'---------------------------------------------------------------------------------------
' Method : Sub Xl2RuleTable
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Xl2RuleTable(editmode As Boolean)            ' only makes sense after OpenRuleTable = True
Dim zErr As cErr
Const zKey As String = "RuleSetup.Xl2RuleTable"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim ntag(1 To 7) As Variant
Dim i As Long
Dim nFile As Long
    Set RuleTable = New Collection
    ' = am anfang des CheckItem heißt: finale Zuordnung, nicht weitersuchen (ab hier)
    ' *                                finale Zuordnung wie "=", exakt diese eine Kategorie
    ' -                                diese Kategorie sicher nicht setzen
    ' sonstige                         diese Kategorie zuordnen (bis evtl. Widerruf)
    ' = am Ende: Wildcard-modus AUS (Wichtig!)
    '                Typ      CheckItem     Category   (ordered, high priority first)
    If LCase(W.xlTSheet.Cells(1, clickColumn)) = "modified" Then
        UseExcelRuleTable = True                 ' and use them from now on because inline code can not be changed here
    Else
        W.xlTSheet.Cells(1, clickColumn).Value = vbNullString
    End If
    If editmode Then
        nFile = -2
        Call openFile(nFile, genCodePath, "InitRuleTable", ".bas", "Output")
        Print #nFile, " & quote( Achtung Generierter Code, nach frmMaintenance ersetzen"
        Print #nFile, "Sub InitRuleTable()"
        Print #nFile, "    Set RuleTable = New Collection"
    End If
    i = 2                                        ' skip headline
    While i > 1
        Call GetLine(i, ntag)
        If LenB(ntag(1)) = 0 Then
            GoTo LoopEnd                         ' end of File
        End If
        Call SetRuleTabFlds(ntag(1), ntag(2), ntag(3), _
                            CBool(ntag(4)), CBool(ntag(5)), _
                            CBool(ntag(6)), CBool(ntag(7)))
        If editmode Then
            Print #nFile, "    Call SetRuleTabFlds(" _
                       & Quote(ntag(1)) & ", " _
                       & Quote(ntag(2)) & ", " _
                       & Quote(ntag(3)) & ", " _
                       & DeBoolToEn(ntag(4)) & ", " _
                       & DeBoolToEn(ntag(5)) & ", " _
                       & DeBoolToEn(ntag(6)) & ", " _
                       & DeBoolToEn(ntag(7)) & ")"
        End If
        i = i + 1
    Wend
LoopEnd:
    If editmode Then
        Print #nFile, "StaticRuleTable =  True"
        Print #nFile, "    UseExcelRuleTable = False ' execute sub InitRuleTable"
        Debug.Print "En" & "d Su" & "b      ' Xl2RuleTable" & vbCrLf
        Print #nFile, " & quote( Ende des Generierten Codes"
        Close #nFile
        Call LogEvent("Created new file " & Quote(genCodePath & "\InitRuleTable.bas") _
                    & " for " & RuleTable.Count & " Rules", eLall)
    End If
    UseExcelRuleTable = True                     ' until we execute sub InitRuleTable

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.Xl2RuleTable

'---------------------------------------------------------------------------------------
' Method : Sub SetRuleTabFlds
' Author : rgbig
' Date   : 20211108@11_47
' Purpose: Trivial, Inline for setting and adding RuleTab via parameter list.
'---------------------------------------------------------------------------------------
Sub SetRuleTabFlds(typ, checkitem, category, final, Exact, never, pattern)

Dim xTag As cRuleCat

    Set xTag = New cRuleCat
    With xTag
        .typ = typ
        .checkitem = checkitem
        .category = category
        .final = CBool(final)
        .Exact = CBool(Exact)
        .never = CBool(never)
        .pattern = CBool(pattern)
    End With                                     ' xTag
    RuleTable.Add xTag

End Sub                                          ' RuleSetup.SetRuleTabFlds

'---------------------------------------------------------------------------------------
' Method : Sub GotRuleTabXl
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GotRuleTabXl(XlwasOpenedHere As Boolean)
Dim zErr As cErr
Const zKey As String = "RuleSetup.GotRuleTabXl"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If RuleTable.Count < 1 Then
        DoVerify False, " Excel Rules table is empty???"
    staticRuleTable = True
        Call InitRuleTable                       ' use the static ones
    Else
        Call LogEvent("RuleSetup is using Excel tags", _
                      eLmin)
    End If
    If XlwasOpenedHere Then
        Call xlEndApp
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.GotRuleTabXl

'---------------------------------------------------------------------------------------
' Method : Sub setItmCats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub setItmCats(curItem As Object, addCatList As String, dropCatList As String, Optional oldCatList As String)
Dim zErr As cErr
Const zKey As String = "RuleSetup.setItmCats"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim sa As Variant
Dim sO As Variant
Dim sX As Variant
Dim category As String
Dim msg As String
Dim sT As String
    
    oldCatList = curItem.Categories
    sa = split(addCatList, ";")
    category = vbNullString
    
    If MatchMode = "*" _
    Or LenB(oldCatList) = 0 _
    Or CurIterationSwitches.ResetCategories _
       Then                                      ' override Category by new value(sX) (forget old)
        msg = "     " & TypeName(curItem) & " categories "
    Else
        sO = split(oldCatList, ";")
        For Each sX In sO
            sX = Trim(sX)
            sT = sX & "; "
            If InStr(dropCatList, sT) = 0 Then
                If InStr(category & "; ", sT) = 0 And LenB(sX) > 0 Then
                    Call AppendTo(category, Trim(sX), "; ")
                End If
            End If
        Next sX
        msg = "     " & TypeName(curItem) & vbTab _
                                        & "assigned categories: "
    End If
    For Each sX In sa
        If InStr(category, sX) = 0 And LenB(sX) > 0 Then
            Call AppendTo(category, Trim(sX), "; ")
        End If
    Next sX
    
    category = Trim(LOGGED & "; " _
                  & StringRemove(category, dropCatList, "; "))
    category = Replace(category, "; ;", ";")
    category = Replace(category, ";;", ";")
    category = RCut(category, 1)
        
    curItem.UnRead = False                       ' should normally cause sA change
    If oldCatList <> category Then
        curItem.Categories = category            ' source-side modification
        category = "changed from " & Quote(oldCatList) & " to "
    Else
        category = "not changed from "
    End If
    
    ' always save original (which was very likely changed)
    ' because item.saved is not reliable when Categories change (IMAP has no categories)
    curItem.Save
    If T_DC.DCerrNum = 0 Then
        MailModified = False                     ' so far, so good
        Call LogEvent(msg & category & Quote(curItem.Categories) & " and saved", eLall)
    Else
        Call LogEvent(msg & category & Quote(curItem.Categories) & " NOT saved", eLall)
    End If
    
    aNewCat = curItem.Categories

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.setItmCats

'---------------------------------------------------------------------------------------
' Method : Sub ModRuleTab
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ModRuleTab()                                 ' called in ExcelEditSession
Dim zErr As cErr
Const zKey As String = "RuleSetup.ModRuleTab"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim modified As Boolean
Dim AR As Range

    Call Xl2RuleTable(editmode:=True)
    modified = x.xlTSheet.Cells(1, changeCounter).Value > 0
    Call Xl2RuleTable(editmode:=True)
    If modified Then
        Debug.Print " & quote( Regeln wurden geändert, bitte die folgenden " _
                  & "Zeilen in Rule-Wizard ersetzen"
        Debug.Print " & quote( *** Anfang des aus Excel generierten Codes"
        Debug.Print " & quote( *** Ende des generierten Codes"
        x.xlTSheet.Cells(1, changeCounter).Clear
    Else
        Call LogEvent("Es wurden keine Änderungen in Excel durchgeführt", eLall)
        UserDecisionEffective = True
    End If
    ' clear changes done in this session
    x.xlTLastLine = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Row - 1
    x.xlTLastCol = ActiveSheet.UsedRange.columns.Count + ActiveSheet.UsedRange.Column - 1
    Set AR = ActiveSheet.Range(Cells(2, clickColumn - 1), Cells(x.xlTLastLine, x.xlTLastCol))
    Set E = x
    AR.Clear
    'Range(Cells(2, clickColumn - 1), _
    A.xlTSheet.Cells.SpecialCells(xlCellTypeLastCell)).Clear ??? *** old code to remove

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.ModRuleTab

'---------------------------------------------------------------------------------------
' Method : Sub EditRulesTable
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub EditRulesTable()
Dim zErr As cErr
Const zKey As String = "RuleSetup.EditRulesTable"
    Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tSub, ExplainS:="")

Dim XlwasOpenedHere As Boolean
    frmMaintenance.Hide
    If OpenRuleTable(XlwasOpenedHere) Then
        DoVerify False
        Call ExcelEditSession(1)
        Call GotRuleTabXl(XlwasOpenedHere)
    Else                                         ' we could not get new rules
        UseExcelRuleTable = False
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' RuleSetup.EditRulesTable



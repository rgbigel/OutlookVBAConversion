VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItmDsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this describes simple access to itemf of class "MapiItem" (Object)
' and references collection of attributes of this item, with decode information


Public idObjDsc As cObjDsc
Public idEntryId As String
Public idTimeValue As Date
Public idPindex As Long
Public idObjItem As Object                         ' the object we currently work on
Public idAttrCount As Long                         ' number-1 of non-null items in idAttrDict
Public idAttrDict As Dictionary                    ' Dictionary of Attributes (cAttrDsc, contains only cDictItem
Public idRules As cAllNameRules                    ' individual rules for this item
Public idFullyDecoded As Boolean                   ' all Dictionary values set for
Public idSelectedAttrs As String                   '  these attributes (vbNullString if not partially decoded)

'---------------------------------------------------------------------------------------
' Method : SetDscValues
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get New Object-dependent Values into Object Class Description
'---------------------------------------------------------------------------------------
Sub SetDscValues(Item As Object, withValues As Boolean, Optional aRules As cAllNameRules, Optional SD As String)
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cItmDsc.SetDscValues"
Static zErr As New cErr

Dim aClassKey As String
Dim bugInfo As String

    If Item Is Nothing Then                        ' not gated: just get for proper Explanation
        aClassKey = "Px=0"
    Else
        aClassKey = CStr(Item.Class) & SD
    End If
        
    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cItmDsc " & aClassKey)
    
    If D_TC.Exists(aClassKey) Then
        bugInfo = "D_TC(" & aClassKey & ", " & D_TC.Item(aClassKey).objItemClassName & ") exists"
        aBugVer = aObjDsc Is D_TC.Item(aClassKey)
        aBugTxt = "design check, should match or switch class ??? aClassKey=" _
                  & aClassKey & " ?? aObjDsc=" & aObjDsc.objClassKey & b & bugInfo
        If DoVerify(aBugVer, aBugTxt) Then
            Set aObjDsc = D_TC.Item(aClassKey)
        End If
    Else
        DoVerify aClassKey = "Px=0", "design check, should exist: ??? aClassKey=" & aClassKey
        Set aObjDsc = New cObjDsc
        Set aObjDsc.objSeqInImportant = New Collection
    End If
         
    If aPindex = 0 Then
        aPindex = 1
    End If
    
    Set aItmDsc = Nothing                          ' forcing SetupAttribs to do something
    Call SetupAttribs(Item, aPindex, withValues:=withValues)
    
    If aRules Is Nothing Then
        Set aRules = aObjDsc.objClsRules
        If aRules Is Nothing Then
            Set aRules = aOD(aPindex).objClsRules
        End If
    End If
    If aRules Is Nothing Then
        DoVerify False, "Design to be verified *** ???"
        GoTo designChk
    Else
        Set aObjDsc.objClsRules = aRules
        Set aRules.RuleObjDsc = aObjDsc            ' modified if needed later
        If LenB(aRules.ARName) = 0 Then            ' new, copied from dftRule, but with Critlist
            aRules.ARName = aObjDsc.objTypeName
        ElseIf aRules.ARName <> aObjDsc.objTypeName Then ' wrong critlist: not OK
            aBugTxt = "Design check aRules.ARName = aObjDsc.objTypeName ??? " & aRules.ARName
            Debug.Print aBugTxt ' ??? should be DoVerify False
designChk:
            ' ??? Set aRules = Nothing
            Set aTD = Nothing
            Set sDictionary = Nothing
        End If
    End If
    
    aObjDsc.objSortMatches = AllPublic.SortMatches
    
    Set sRules = aRules

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' cItmDsc.SetDscValues

'---------------------------------------------------------------------------------------
' Method : Function GetAttrDsc4Prop
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Function GetAttrDsc4Prop(ByVal trueindex As Long, Optional sPindex As Long) As cAttrDsc
Dim zErr As cErr
Const zKey As String = "cObjDsc.GetAttrDsc4Prop"

Dim trapme As Long

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cObjDsc")
    
    If sPindex = 0 Then                            ' VB can't default to another module's defined value
        If sourceIndex > 0 Then
            sPindex = sourceIndex
        Else
            If idPindex = 0 Then
                sPindex = 1
            Else
                sPindex = Me.idPindex              ' same as last time
            End If
        End If
    End If
    
    Set aProp = aProps(trueindex)
    aBugVer = PropertyNameX = aProp.Name
    If DoVerify(aBugVer, "design check PropertyNameX = aProp.Name ???") Then
        PropertyNameX = aProp.Name
    End If
    
    If Not aItmDsc Is aID(sPindex) Then
        Set aItmDsc = aID(sPindex)
    End If
    
    If aItmDsc.idAttrDict Is Nothing Then
        Set GetAttrDsc4Prop = Nothing
        GoTo FuncExit
    End If
    
    trapme = aItmDsc.idAttrDict.Count
    If aItmDsc.idAttrDict.Exists(aProp.Name) Then
        aBugVer = trapme = aItmDsc.idAttrDict.Count
        DoVerify aBugVer, "trapme = aItmDsc.idAttrDict.Count ???"
        If isEmpty(aItmDsc.idAttrDict.Item(aProp.Name)) Then
            Set aTD = New cAttrDsc                 ' *** correct the bug: Exists generated empty Item ???
            Set aItmDsc.idAttrDict.Item(aProp.Name) = aTD
            Set GetAttrDsc4Prop = aItmDsc.idAttrDict.Item(aProp.Name)
        Else
            Set GetAttrDsc4Prop = aItmDsc.idAttrDict.Item(aProp.Name)
        End If
    Else
        Set GetAttrDsc4Prop = Nothing
        GoTo FuncExit
    End If
    apropTrueIndex = trueindex
    
    Call GetMiAttrNr                               ' find out if we specifically need this attr value
    If GetAttrDsc4Prop.adisSel Then
        GetAttrDsc4Prop.adInfo.iType = inv         ' Force new values
        Call GetAttrDsc4Prop.GetScalarValue        ' at this time, without formatting
    End If
    
FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cObjDsc.GetAttrDsc4Prop

'---------------------------------------------------------------------------------------
' Method : Function IDictClone
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Function IDictClone(withValues As Boolean) As Dictionary
Dim zErr As cErr
Const zKey As String = "cItmDsc.IDictClone"

Dim i As Long
Dim aDictItemi As cAttrDsc
Dim thisAD As cAttrDsc
Dim cloneDI As cAttrDsc
Dim xObjDsc As cObjDsc
Dim aClassKey As String
    
'------------------- gated Entry -------------------------------------------------------

    If idAttrDict Is Nothing Then
        GoTo pExit                                 ' nothing there to clone
    Else
        aBugVer = aID(aPindex) Is aItmDsc          ' if this must always be, use aItmDsc hereafter
        aBugVer = aBugVer And aItmDsc Is Me
        If DoVerify(aBugVer, "aID(aPindex) Is aItmDsc ???") Then ' runtime check
            GoTo pExit
        End If
        
        aClassKey = aItmDsc.idObjDsc.objClassKey
        If D_TC.Exists(aClassKey) Then
            If aObjDsc.objClassKey = aClassKey Then
                Set xObjDsc = aObjDsc
            Else
                Set xObjDsc = D_TC.Item(aClassKey)
            End If
        Else
            Set xObjDsc = aItmDsc.idObjDsc
        End If
        
        aBugVer = xObjDsc.objClassKey = aClassKey
        aBugTxt = "messed up class key ???"
        If DoVerify Then
            GoTo pExit                             ' wrong type of clone
        End If
            
        Set ActItemObject = idObjItem
        Set sDictionary = aItmDsc.idAttrDict
        aBugVer = Not sDictionary Is Nothing
        aBugTxt = "no dictionary"              ' runtime check
        If Not aBugVer Then
            DoVerify False, aBugTxt & " design check, do remove quickly ???"
            GoTo pExit
        End If
        If sDictionary.Count = 0 And Not withValues Then
            GoTo pExit                             ' Dictionary still empty, no need to clone
        End If
        aBugVer = sDictionary.Count >= ActItemObject.ItemProperties.Count
        aBugTxt = "not enough ItemProperties in it ???" ' runtime check
        If DoVerify Then
            GoTo pExit
        End If
    End If
    
    ' ----------------- end Gate -----------------------------------------------------------

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cItmDsc")

    Set IDictClone = New Dictionary
    
    IDictClone.Add sDictionary.Keys(0), sDictionary.Items(0) ' NOT a cDictItem: String -> ClassKey
    
    aCloneMode = withNewValues                     ' use ADItmDsc, rules etc, but not values
    For i = 1 To sDictionary.Count - 1
        Set thisAD = sDictionary.Items(i)
        PropertyNameX = thisAD.adKey
        Set cloneDI = New cAttrDsc
        cloneDI.adKey = thisAD.adKey
        aTD.adtrueIndex = thisAD.adtrueIndex       ' aTD returned from new ==> ItemProperties
        IDictClone.Add cloneDI.adKey, cloneDI
        If withValues Then
            Set aProp = ActItemObject.ItemProperties.Item(PropertyNameX)
            Call aTD.GetScalarValue
            Call PrepDecodeProp
            Call StackAttribute
        End If
    Next i
    
FunExit:
    Set aDictItemi = Nothing
    Set cloneDI = Nothing
    Set xObjDsc = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cItmDsc.IDictClone

'---------------------------------------------------------------------------------------
' Method : Sub UpdItmClsDetails
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Update existing  class Time selection details, object values not decoded
'---------------------------------------------------------------------------------------
Function UpdItmClsDetails(Item As Object) As Boolean
Dim zErr As cErr
Const zKey As String = "cObjDsc.UpdItmClsDetails"
    
'------------------- gated Entry -------------------------------------------------------

    If idEntryId <> Item.EntryID Then
        idEntryId = Item.EntryID
        idTimeValue = 0
    End If
    If Not idObjItem Is Item Then
        Set idObjItem = Item
        idTimeValue = 0
    End If
    If idTimeValue <> 0 Then
        GoTo ProcRet
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)
    
    UpdItmClsDetails = True                        ' update is needed
    Call UpdItmTime

ProcReturn:
    Call ProcExit(zErr, "Time updated for item")

ProcRet:
End Function                                       ' cObjDsc.UpdItmClsDetails

'---------------------------------------------------------------------------------------
' Method : Sub UpdItmTime
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Update the distinction time of an Object safely
'---------------------------------------------------------------------------------------
Sub UpdItmTime()

Const zKey As String = "cObjDsc.UpdItmTime"
    Call DoCall(zKey, tSub, eQzMode)

Dim A As AppointmentItem
Dim M As MailItem
Dim i As MeetingItem
Dim T As TaskItem
Dim C As TaskRequestAcceptItem
Dim D As TaskRequestDeclineItem
Dim Q As TaskRequestItem
Dim W As TaskRequestUpdateItem

Dim mClass As String

    aBugTxt = "Get MessageClass"
    Call Try
    mClass = idObjItem.MessageClass
    If Catch Then
        GoTo badFin
    End If
        
    Select Case TypeName(idObjItem)
        Case "MailItem"
            Set M = idObjItem
            idTimeValue = M.SentOn
            DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
            Set M = Nothing
        Case "AppointmentItem"
            Set A = idObjItem                      ' Explicit object to Outlook Type conversions
            idTimeValue = A.start
            DoVerify aObjDsc.objTimeType = "Start", "Check init of aObjDsc"
            Set A = Nothing
        Case "MeetingItem"
            Set i = idObjItem
            If (mClass = "IPM.Schedule.Meeting.Request") Then
                ' Meeting Request: Time is in associated Appointment
                Set A = i.GetAssociatedAppointment(False)
                If Not A Is Nothing Then            ' appointment could be gone/cancelled
                    idTimeValue = A.start           ' schedule the associated Appointment
                    Set A = Nothing
                End If
            Else
                ' Other MeetingItem
                aBugTxt = "SentOn Time of Item"
                Call Try
                idTimeValue = i.SentOn
                Set i = Nothing
noSent:
                If Catch Then
                    Call LogEvent("SentOn Time of Item " & aObjDsc.objTimeType _
                                  & " not available for Item Class " _
                                  & aObjDsc.objItemClassName, 0)
                End If
            End If
        Case "TaskItem"
            Set T = idObjItem
            DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
            idTimeValue = T.SentOn
            Set T = Nothing
            GoTo noSent
        Case "TaskRequestAcceptItem"
            Set C = idObjItem
            DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
            idTimeValue = C.SentOn
            Set C = Nothing
            GoTo noSent
        Case "TaskRequestDeclineItem"
            Set D = idObjItem
            DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
            idTimeValue = D.SentOn
            Set D = Nothing
            GoTo noSent
        Case "TaskRequestItem"
            Set Q = idObjItem
            DoVerify aObjDsc.objTimeType = "CreationTime", "Check init of aObjDsc"
            Set Q = Nothing
            GoTo noCreation
        Case "TaskRequestUpdateItem"
            Set W = idObjItem
            DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
            idTimeValue = W.SentOn
            Set W = Nothing
            GoTo noSent
        Case Else
            aObjDsc.objTimeType = "CreationTime"
            idTimeValue = idObjItem.CreationTime
noCreation:
            If Catch Then
                Call LogEvent("CreationTime of Item " & aObjDsc.objTimeType _
                              & " not available for Item Class " _
                              & aObjDsc.objItemClassName, 0)
            End If
    End Select
    If DebugMode Or DebugLogging Then              ' *** design check only
        Debug.Print "Type Name=" & TypeName(idObjItem), _
        LString(LString(aObjDsc.objTimeType, 15) & CStr(idTimeValue), 40), _
        Quote(idObjItem.Subject)
badFin:
        If LenB(mClass) > 0 Then
            Debug.Print "MessageClass=" & mClass & b & T_DC.DCerrMsg
        Else
            Debug.Print "UpdItmTime Error: "; T_DC.DCerrMsg
        End If
    End If
    Call ErrReset

zExit:
    Call DoExit(zKey)

End Sub                                            ' cObjDsc.UpdItmTime

'---------------------------------------------------------------------------------------
' Method : Function getObjDsc4Itm
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Get ObjectDesc, including optional SD Extensions
' Note   : will not get new values, but will init Item Class Model if new
'---------------------------------------------------------------------------------------
Function getObjDsc4Itm(Item As Object, Optional SD As String) As cItmDsc
    
Dim aClassKey As String

    DoVerify LenB(SD) = 0, "SD version use is new"
    aClassKey = TypeName(Item) & SD
    
    If D_TC.Exists(aClassKey) Then
        Set getObjDsc4Itm = D_TC.Item(aClassKey)
    Else
        Call GetITMClsModel(Item, aPindex)
        DoVerify aID(aPindex) Is getObjDsc4Itm, "** omit next if no hit ???"
        Set getObjDsc4Itm = aObjDsc
    End If
    
    If aItmDsc Is Nothing Then
        Set aItmDsc = New cItmDsc
    ElseIf aItmDsc.idObjDsc.objClassKey <> CStr(aClassKey) Then
        Set aItmDsc = New cItmDsc
    End If

End Function                                       ' cObjDsc.getObjDsc4Itm



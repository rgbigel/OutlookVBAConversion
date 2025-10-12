# Converted from cItmDsc.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cItmDsc"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# ' this describes simple access to itemf of class "MapiItem" (Object)
# ' and references collection of attributes of this item, with decode information


# Public idObjDsc As cObjDsc
# Public idEntryId As String
# Public idTimeValue As Date
# Public idPindex As Long
# Public idObjItem As Object                         ' the object we currently work on
# Public idAttrCount As Long                         ' number-1 of non-null items in idAttrDict
# Public idAttrDict As Dictionary                    ' Dictionary of Attributes (cAttrDsc, contains only cDictItem
# Public idRules As cAllNameRules                    ' individual rules for this item
# Public idFullyDecoded As Boolean                   ' all Dictionary values set for
# Public idSelectedAttrs As String                   '  these attributes (vbNullString if not partially decoded)

# '---------------------------------------------------------------------------------------
# ' Method : SetDscValues
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get New Object-dependent Values into Object Class Description
# '---------------------------------------------------------------------------------------
def setdscvalues():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cItmDsc.SetDscValues"
    # Static zErr As New cErr

    # Dim aClassKey As String
    # Dim bugInfo As String

    if Item Is Nothing Then                        ' not gated: just get for proper Explanation:
    # aClassKey = "Px=0"
    else:
    # aClassKey = CStr(Item.Class) & SD

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cItmDsc " & aClassKey)

    if D_TC.Exists(aClassKey) Then:
    # bugInfo = "D_TC(" & aClassKey & ", " & D_TC.Item(aClassKey).objItemClassName & ") exists"
    # aBugVer = aObjDsc Is D_TC.Item(aClassKey)
    # aBugTxt = "design check, should match or switch class ??? aClassKey=" _
    # & aClassKey & " ?? aObjDsc=" & aObjDsc.objClassKey & b & bugInfo
    if DoVerify(aBugVer, aBugTxt) Then:
    # Set aObjDsc = D_TC.Item(aClassKey)
    else:
    # DoVerify aClassKey = "Px=0", "design check, should exist: ??? aClassKey=" & aClassKey
    # Set aObjDsc = New cObjDsc
    # Set aObjDsc.objSeqInImportant = New Collection

    if aPindex = 0 Then:
    # aPindex = 1

    # Set aItmDsc = Nothing                          ' forcing SetupAttribs to do something
    # Call SetupAttribs(Item, aPindex, withValues:=withValues)

    if aRules Is Nothing Then:
    # Set aRules = aObjDsc.objClsRules
    if aRules Is Nothing Then:
    # Set aRules = aOD(aPindex).objClsRules
    if aRules Is Nothing Then:
    # DoVerify False, "Design to be verified *** ???"
    # GoTo designChk
    else:
    # Set aObjDsc.objClsRules = aRules
    # Set aRules.RuleObjDsc = aObjDsc            ' modified if needed later
    if LenB(aRules.ARName) = 0 Then            ' new, copied from dftRule, but with Critlist:
    # aRules.ARName = aObjDsc.objTypeName
    elif aRules.ARName <> aObjDsc.objTypeName Then ' wrong critlist: not OK:
    # aBugTxt = "Design check aRules.ARName = aObjDsc.objTypeName ??? " & aRules.ARName
    print(Debug.Print aBugTxt ' ??? should be DoVerify False)
    # designChk:
    # ' ??? Set aRules = Nothing
    # Set aTD = Nothing
    # Set sDictionary = Nothing

    # aObjDsc.objSortMatches = AllPublic.SortMatches

    # Set sRules = aRules

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetAttrDsc4Prop
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function GetAttrDsc4Prop(ByVal trueindex As Long, Optional sPindex As Long) As cAttrDsc
# Dim zErr As cErr
# Const zKey As String = "cObjDsc.GetAttrDsc4Prop"

# Dim trapme As Long

# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cObjDsc")

if sPindex = 0 Then                            ' VB can't default to another module's defined value:
if sourceIndex > 0 Then:
# sPindex = sourceIndex
else:
if idPindex = 0 Then:
# sPindex = 1
else:
# sPindex = Me.idPindex              ' same as last time

# Set aProp = aProps(trueindex)
# aBugVer = PropertyNameX = aProp.Name
if DoVerify(aBugVer, "design check PropertyNameX = aProp.Name ???") Then:
# PropertyNameX = aProp.Name

if Not aItmDsc Is aID(sPindex) Then:
# Set aItmDsc = aID(sPindex)

if aItmDsc.idAttrDict Is Nothing Then:
# Set GetAttrDsc4Prop = Nothing
# GoTo FuncExit

# trapme = aItmDsc.idAttrDict.Count
if aItmDsc.idAttrDict.Exists(aProp.Name) Then:
# aBugVer = trapme = aItmDsc.idAttrDict.Count
# DoVerify aBugVer, "trapme = aItmDsc.idAttrDict.Count ???"
if isEmpty(aItmDsc.idAttrDict.Item(aProp.Name)) Then:
# Set aTD = New cAttrDsc                 ' *** correct the bug: Exists generated empty Item ???
# Set aItmDsc.idAttrDict.Item(aProp.Name) = aTD
# Set GetAttrDsc4Prop = aItmDsc.idAttrDict.Item(aProp.Name)
else:
# Set GetAttrDsc4Prop = aItmDsc.idAttrDict.Item(aProp.Name)
else:
# Set GetAttrDsc4Prop = Nothing
# GoTo FuncExit
# apropTrueIndex = trueindex

# Call GetMiAttrNr                               ' find out if we specifically need this attr value
if GetAttrDsc4Prop.adisSel Then:
# GetAttrDsc4Prop.adInfo.iType = inv         ' Force new values
# Call GetAttrDsc4Prop.GetScalarValue        ' at this time, without formatting

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function IDictClone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function IDictClone(withValues As Boolean) As Dictionary
# Dim zErr As cErr
# Const zKey As String = "cItmDsc.IDictClone"

# Dim i As Long
# Dim aDictItemi As cAttrDsc
# Dim thisAD As cAttrDsc
# Dim cloneDI As cAttrDsc
# Dim xObjDsc As cObjDsc
# Dim aClassKey As String

# '------------------- gated Entry -------------------------------------------------------

if idAttrDict Is Nothing Then:
# GoTo pExit                                 ' nothing there to clone
else:
# aBugVer = aID(aPindex) Is aItmDsc          ' if this must always be, use aItmDsc hereafter
# aBugVer = aBugVer And aItmDsc Is Me
if DoVerify(aBugVer, "aID(aPindex) Is aItmDsc ???") Then ' runtime check:
# GoTo pExit

# aClassKey = aItmDsc.idObjDsc.objClassKey
if D_TC.Exists(aClassKey) Then:
if aObjDsc.objClassKey = aClassKey Then:
# Set xObjDsc = aObjDsc
else:
# Set xObjDsc = D_TC.Item(aClassKey)
else:
# Set xObjDsc = aItmDsc.idObjDsc

# aBugVer = xObjDsc.objClassKey = aClassKey
# aBugTxt = "messed up class key ???"
if DoVerify Then:
# GoTo pExit                             ' wrong type of clone

# Set ActItemObject = idObjItem
# Set sDictionary = aItmDsc.idAttrDict
# aBugVer = Not sDictionary Is Nothing
# aBugTxt = "no dictionary"              ' runtime check
if Not aBugVer Then:
# DoVerify False, aBugTxt & " design check, do remove quickly ???"
# GoTo pExit
if sDictionary.Count = 0 And Not withValues Then:
# GoTo pExit                             ' Dictionary still empty, no need to clone
# aBugVer = sDictionary.Count >= ActItemObject.ItemProperties.Count
# aBugTxt = "not enough ItemProperties in it ???" ' runtime check
if DoVerify Then:
# GoTo pExit

# ' ----------------- end Gate -----------------------------------------------------------

# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cItmDsc")

# Set IDictClone = New Dictionary

# IDictClone.Add sDictionary.Keys(0), sDictionary.Items(0) ' NOT a cDictItem: String -> ClassKey

# aCloneMode = withNewValues                     ' use ADItmDsc, rules etc, but not values
# Set thisAD = sDictionary.Items(i)
# PropertyNameX = thisAD.adKey
# Set cloneDI = New cAttrDsc
# cloneDI.adKey = thisAD.adKey
# aTD.adtrueIndex = thisAD.adtrueIndex       ' aTD returned from new ==> ItemProperties
# IDictClone.Add cloneDI.adKey, cloneDI
if withValues Then:
# Set aProp = ActItemObject.ItemProperties.Item(PropertyNameX)
# Call aTD.GetScalarValue
# Call PrepDecodeProp
# Call StackAttribute

# FunExit:
# Set aDictItemi = Nothing
# Set cloneDI = Nothing
# Set xObjDsc = Nothing

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub UpdItmClsDetails
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Update existing  class Time selection details, object values not decoded
# '---------------------------------------------------------------------------------------
def upditmclsdetails():
    # Dim zErr As cErr
    # Const zKey As String = "cObjDsc.UpdItmClsDetails"

    # '------------------- gated Entry -------------------------------------------------------

    if idEntryId <> Item.EntryID Then:
    # idEntryId = Item.EntryID
    # idTimeValue = 0
    if Not idObjItem Is Item Then:
    # Set idObjItem = Item
    # idTimeValue = 0
    if idTimeValue <> 0 Then:
    # GoTo ProcRet

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    # UpdItmClsDetails = True                        ' update is needed
    # Call UpdItmTime

    # ProcReturn:
    # Call ProcExit(zErr, "Time updated for item")

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub UpdItmTime
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Update the distinction time of an Object safely
# '---------------------------------------------------------------------------------------
def upditmtime():

    # Const zKey As String = "cObjDsc.UpdItmTime"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim A As AppointmentItem
    # Dim M As MailItem
    # Dim i As MeetingItem
    # Dim T As TaskItem
    # Dim C As TaskRequestAcceptItem
    # Dim D As TaskRequestDeclineItem
    # Dim Q As TaskRequestItem
    # Dim W As TaskRequestUpdateItem

    # Dim mClass As String

    # aBugTxt = "Get MessageClass"
    # Call Try
    # mClass = idObjItem.MessageClass
    if Catch Then:
    # GoTo badFin

    match TypeName(idObjItem):
        case "MailItem":
    # Set M = idObjItem
    # idTimeValue = M.SentOn
    # DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
    # Set M = Nothing
        case "AppointmentItem":
    # Set A = idObjItem                      ' Explicit object to Outlook Type conversions
    # idTimeValue = A.start
    # DoVerify aObjDsc.objTimeType = "Start", "Check init of aObjDsc"
    # Set A = Nothing
        case "MeetingItem":
    # Set i = idObjItem
    if (mClass = "IPM.Schedule.Meeting.Request") Then:
    # ' Meeting Request: Time is in associated Appointment
    # Set A = i.GetAssociatedAppointment(False)
    if Not A Is Nothing Then            ' appointment could be gone/cancelled:
    # idTimeValue = A.start           ' schedule the associated Appointment
    # Set A = Nothing
    else:
    # ' Other MeetingItem
    # aBugTxt = "SentOn Time of Item"
    # Call Try
    # idTimeValue = i.SentOn
    # Set i = Nothing
    # noSent:
    if Catch Then:
    # Call LogEvent("SentOn Time of Item " & aObjDsc.objTimeType _
    # & " not available for Item Class " _
    # & aObjDsc.objItemClassName, 0)
        case "TaskItem":
    # Set T = idObjItem
    # DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
    # idTimeValue = T.SentOn
    # Set T = Nothing
    # GoTo noSent
        case "TaskRequestAcceptItem":
    # Set C = idObjItem
    # DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
    # idTimeValue = C.SentOn
    # Set C = Nothing
    # GoTo noSent
        case "TaskRequestDeclineItem":
    # Set D = idObjItem
    # DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
    # idTimeValue = D.SentOn
    # Set D = Nothing
    # GoTo noSent
        case "TaskRequestItem":
    # Set Q = idObjItem
    # DoVerify aObjDsc.objTimeType = "CreationTime", "Check init of aObjDsc"
    # Set Q = Nothing
    # GoTo noCreation
        case "TaskRequestUpdateItem":
    # Set W = idObjItem
    # DoVerify aObjDsc.objTimeType = "SentOn", "Check init of aObjDsc"
    # idTimeValue = W.SentOn
    # Set W = Nothing
    # GoTo noSent
        case _:
    # aObjDsc.objTimeType = "CreationTime"
    # idTimeValue = idObjItem.CreationTime
    # noCreation:
    if Catch Then:
    # Call LogEvent("CreationTime of Item " & aObjDsc.objTimeType _
    # & " not available for Item Class " _
    # & aObjDsc.objItemClassName, 0)
    if DebugMode Or DebugLogging Then              ' *** design check only:
    print(Debug.Print "Type Name=" & TypeName(idObjItem), _)
    # LString(LString(aObjDsc.objTimeType, 15) & CStr(idTimeValue), 40), _
    # Quote(idObjItem.Subject)
    # badFin:
    if LenB(mClass) > 0 Then:
    print(Debug.Print "MessageClass=" & mClass & b & T_DC.DCerrMsg)
    else:
    print(Debug.Print "UpdItmTime Error: "; T_DC.DCerrMsg)
    # Call ErrReset

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function getObjDsc4Itm
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get ObjectDesc, including optional SD Extensions
# ' Note   : will not get new values, but will init Item Class Model if new
# '---------------------------------------------------------------------------------------
def getobjdsc4itm():

    # Dim aClassKey As String

    # DoVerify LenB(SD) = 0, "SD version use is new"
    # aClassKey = TypeName(Item) & SD

    if D_TC.Exists(aClassKey) Then:
    # Set getObjDsc4Itm = D_TC.Item(aClassKey)
    else:
    # Call GetITMClsModel(Item, aPindex)
    # DoVerify aID(aPindex) Is getObjDsc4Itm, "** omit next if no hit ???"
    # Set getObjDsc4Itm = aObjDsc

    if aItmDsc Is Nothing Then:
    # Set aItmDsc = New cItmDsc
    elif aItmDsc.idObjDsc.objClassKey <> CStr(aClassKey) Then:
    # Set aItmDsc = New cItmDsc



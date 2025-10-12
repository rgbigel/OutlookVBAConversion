# Converted from cObjDsc.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cObjDsc"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# ' this describes one Type of class "MapiItem" (Object)

# Public objClassKey As String
# Attribute objClassKey.VB_VarUserMemId = 0
# Attribute objClassKey.VB_VarDescription = "Object Class Key"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute objClassKey.VB_VarUserMemId = 0
# ' Attribute objClassKey.VB_VarDescription = "Object Class Key"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public objItemClass As OlObjectClass
# Public objItemType As OlItemType
# Public objItemClassName As String
# Public objTypeName As String                       ' default value of class <self> ****
# Public objClsRules As cAllNameRules                ' current (s-)Rule for this class
# Public objDefaultIdent As String
# Public objNameExt As String

# Public objIsMailLike As Boolean                    ' some of these have missing Properties
# Public objTimeType As String
# Public objHasReceivedTime As Boolean
# Public objHasHtmlBodyFlag As Boolean
# Public objHasSenderName As Boolean
# Public objHasSentOnBehalfOf As Boolean
# Public objDumpMade As Long                         ' last item dumped or -1 when never

# ' Note: all following are "not quite static" i.e. may depend on user options and can "grow over time"
# Public objDftMatches As String
# Public objSortMatches As String
# Public objMaxAttrCount As Long                     ' max number defined slots; "nearly" invariant for each ObjectType
# Public objMinAttrCount As Long                     ' number of Attrs before Recurrences/Exceptions
# Public objSeqInImportant As Collection

# '---------------------------------------------------------------------------------------
# ' Note   : When a New cObjDsc is created from cObjDsc, it provides no objClsRules, but
# '          will then copy / clone the rule parts (from sRules) or make them if Nothing
# '---------------------------------------------------------------------------------------
# Public Sub ODescClone(aClassKey As String, Optional sITMDsc As cItmDsc)
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "cObjDsc.ODescClone"
# Call DoCall(zKey, tSub, eQzMode)

if objClassKey <> Me.objClassKey Then:
# objClassKey = aClassKey
# D_TC.Add aClassKey, Me
# Set aObjDsc = Me

# With aObjDsc
if sITMDsc Is Nothing Then:
# Set sITMDsc = Me
# .objDumpMade = -1
else:
if Not sITMDsc.idObjDsc Is Nothing Then:
# DoVerify (objClassKey = sITMDsc.idObjDsc.objClassKey _
# Or objClassKey = CStr(sITMDsc.idObjItem.Class) _
# & sITMDsc.idObjDsc.objNameExt), _
# "Keys not canonical ???"
# Set sITMDsc.idObjDsc = aObjDsc

# .objItemClass = Me.objItemClass
# .objTypeName = Me.objTypeName
# .objDefaultIdent = Me.objDefaultIdent

if LenB(Me.objSortMatches) = 0 Then:
# Me.objSortMatches = AllPublic.SortMatches
# .objDftMatches = Me.objDftMatches
# .objSortMatches = Me.objSortMatches

if Me.objClsRules Is Nothing Then:
# Set .objClsRules = dftRule
if sITMDsc.idObjDsc Is Nothing Then:
# Set sITMDsc.idObjDsc = aObjDsc
# Set .objClsRules = Me.objClsRules.AllRulesClone(ClassRules, sITMDsc.idObjDsc, True)
# Set .objClsRules.clsNeverCompare.PropAllRules = .objClsRules
# Set .objClsRules.clsObligMatches.PropAllRules = .objClsRules
# Set .objClsRules.clsNotDecodable.PropAllRules = .objClsRules
# Set .objClsRules.clsSimilarities.PropAllRules = .objClsRules
# Set .objSeqInImportant = Me.objSeqInImportant

# .objMaxAttrCount = Me.objMaxAttrCount
# .objMinAttrCount = Me.objMinAttrCount
# .objItemClassName = Me.objItemClassName
# .objIsMailLike = Me.objIsMailLike
# .objHasReceivedTime = Me.objHasReceivedTime
# .objHasHtmlBodyFlag = Me.objHasHtmlBodyFlag
# .objHasSenderName = Me.objHasSenderName
# .objTimeType = Me.objTimeType
# .objHasSentOnBehalfOf = Me.objHasSentOnBehalfOf
# .objDumpMade = Me.objDumpMade
# End With                                       ' aObjDsc

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function IsSame
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def issame():
    # Const zKey As String = "cItemClsProps.IsSame"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim NrOfDiffs As Long

    if other Is Nothing Then                       ' invalid: return False:
    # GoTo zExit
    if Me Is other Then:
    # IsSame = True
    # GoTo zExit

    # ' incomplete=True stops some of the comparisons
    # NrOfDiffs = 0
    if Not objItemClassName <> other.objItemClassName Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " & "objItemClassName " _)
    # & objItemClassName & "<>" & other.objItemClassName
    else:
    # GoTo OneDiffIsEnough

    if objItemClass <> other.objItemClass Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objItemClass " & objItemClass & "<>" & other.objItemClass
    else:
    # GoTo OneDiffIsEnough

    if objItemClassName <> other.objItemClassName Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objItemClassName " & objItemClassName & "<>" & other.objItemClassName
    else:
    # GoTo OneDiffIsEnough
    if objItemType <> other.objItemType Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objItemType " & objItemType & "<>" & other.objItemType
    else:
    # GoTo OneDiffIsEnough
    if objIsMailLike <> other.objIsMailLike Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objIsMailLike " & objIsMailLike & "<>" & other.objIsMailLike
    else:
    # GoTo OneDiffIsEnough
    if objHasReceivedTime <> other.objHasReceivedTime Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objHasReceivedTime " & objHasReceivedTime & "<>" & other.objHasReceivedTime
    else:
    # GoTo OneDiffIsEnough
    if objHasHtmlBodyFlag <> other.objHasHtmlBodyFlag Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objHasHtmlBodyFlag " & objHasHtmlBodyFlag & "<>" & other.objHasHtmlBodyFlag
    else:
    # GoTo OneDiffIsEnough
    if objHasSenderName <> other.objHasSenderName Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objHasSenderName " & objHasSenderName & "<>" & other.objHasSenderName
    else:
    # GoTo OneDiffIsEnough
    if objHasSentOnBehalfOf <> other.objHasSentOnBehalfOf Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "objHasSentOnBehalfOf " _
    # & objHasSentOnBehalfOf & "<>" & other.objHasSentOnBehalfOf
    else:
    # GoTo OneDiffIsEnough

    # IsSame = True
    # GoTo zExit

    # OneDiffIsEnough:
    # IsSame = (NrOfDiffs = 0)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : objDictClone
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Copy Class Model's Dictionary to a new Dictionary
# '---------------------------------------------------------------------------------------
def objdictclone():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cObjDsc.objDictClone"
    # #If MoreDiagnostics Then
    # Call DoCall(zKey, "Function", eQzMode)
    # #End If

    # Dim i As Long
    # Dim cloneDe As cObjDsc
    # Dim cloneID As cAttrDsc
    # Dim thisClone As Object
    # Dim adKey As String

    if ModelDict.Count < 2 Then:
    # GoTo zExit

    # Set objDictClone = New Dictionary

    # adKey = ModelDict.Items(0).objClassKey
    # Set cloneDe = ModelDict.Items(0)
    # objDictClone.Add adKey, cloneDe

    # Set thisClone = ModelDict.Items(i)
    # PropertyNameX = thisClone.adKey
    # Set cloneID = thisClone.adictClone
    # ' Set cloneID = New cAttrDsc
    # ' cloneID.adKey = thisAD.adKey
    # ' aTD.adtrueIndex = thisAD.adtrueIndex
    # objDictClone.Add cloneID.adKey, cloneID

    # Set thisClone = Nothing
    # Set cloneDe = Nothing
    # Set cloneID = Nothing

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:


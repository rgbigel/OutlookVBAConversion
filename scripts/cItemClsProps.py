# Converted from cItemClsProps.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cItemClsProps"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public aCitemClass As OlObjectClass
# Public aCitemClassName As String
# Public aCITMDsc As cItmDsc                         ' Parent

# Public aCitemType As OlItemType

# Public aCisMailLike As Boolean                     ' some of these have missing Properties
# Public aHasReceivedTime As Boolean
# Public aHasHtmlBodyFlag As Boolean
# Public aHasSenderName As Boolean
# Public aHasSentOnBehalfOf As Boolean

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
    if aCITMDsc Is Nothing Then                    ' invalid: return False:
    # GoTo zExit
    if Me Is other Then:
    # IsSame = True
    # GoTo zExit

    # ' incomplete=True stops some of the comparisons
    # NrOfDiffs = 0
    if Not aCITMDsc.idObjDsc.objItemClassName _:
    # <> other.aCITMDsc.idObjDsc.objItemClassName Then
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aCITMDsc.idObjDsc.objItemClassName " _
    # & aCITMDsc.idObjDsc.objItemClassName _
    # & "<>" & other.aCITMDsc.idObjDsc.objItemClassName
    else:
    # GoTo OneDiffIsEnough

    if aCitemClass <> other.aCitemClass Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aCitemClass " & aCitemClass _
    # & "<>" & other.aCitemClass
    else:
    # GoTo OneDiffIsEnough

    if aCitemClassName <> other.aCitemClassName Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aCitemClassName " & aCitemClassName _
    # & "<>" & other.aCitemClassName
    else:
    # GoTo OneDiffIsEnough
    if aCitemType <> other.aCitemType Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aCitemType " & aCitemType _
    # & "<>" & other.aCitemType
    else:
    # GoTo OneDiffIsEnough
    if aCisMailLike <> other.aCisMailLike Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aCisMailLike " & aCisMailLike _
    # & "<>" & other.aCisMailLike
    else:
    # GoTo OneDiffIsEnough
    if aHasReceivedTime <> other.aHasReceivedTime Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aHasReceivedTime " & aHasReceivedTime _
    # & "<>" & other.aHasReceivedTime
    else:
    # GoTo OneDiffIsEnough
    if aHasHtmlBodyFlag <> other.aHasHtmlBodyFlag Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aHasHtmlBodyFlag " & aHasHtmlBodyFlag _
    # & "<>" & other.aHasHtmlBodyFlag
    else:
    # GoTo OneDiffIsEnough
    if aHasSenderName <> other.aHasSenderName Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aHasSenderName " & aHasSenderName _
    # & "<>" & other.aHasSenderName
    else:
    # GoTo OneDiffIsEnough
    if aHasSentOnBehalfOf <> other.aHasSentOnBehalfOf Then:
    # NrOfDiffs = NrOfDiffs + 1
    if showdiffs Then:
    print(Debug.Print "  not same: " _)
    # & "aHasSentOnBehalfOf " _
    # & aHasSentOnBehalfOf _
    # & "<>" & other.aHasSentOnBehalfOf
    else:
    # GoTo OneDiffIsEnough

    # IsSame = True
    # GoTo zExit

    # OneDiffIsEnough:
    # IsSame = (NrOfDiffs = 0)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Clone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clone():

    # Set Clone = New cItemClsProps

    # Clone.aCitemType = aCitemType
    # Set Clone.aCITMDsc = aCITMDsc

    # Clone.aCisMailLike = aCisMailLike
    # Clone.aHasReceivedTime = aHasReceivedTime
    # Clone.aHasHtmlBodyFlag = aHasHtmlBodyFlag
    # Clone.aHasSenderName = aHasSenderName
    # Clone.aHasSentOnBehalfOf = aHasSentOnBehalfOf



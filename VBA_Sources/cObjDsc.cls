VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObjDsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this describes one Type of class "MapiItem" (Object)

Public objClassKey As String
Attribute objClassKey.VB_VarUserMemId = 0
Attribute objClassKey.VB_VarDescription = "Object Class Key"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute objClassKey.VB_VarUserMemId = 0
' Attribute objClassKey.VB_VarDescription = "Object Class Key"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public objItemClass As OlObjectClass
Public objItemType As OlItemType
Public objItemClassName As String
Public objTypeName As String                       ' default value of class <self> ****
Public objClsRules As cAllNameRules                ' current (s-)Rule for this class
Public objDefaultIdent As String
Public objNameExt As String

Public objIsMailLike As Boolean                    ' some of these have missing Properties
Public objTimeType As String
Public objHasReceivedTime As Boolean
Public objHasHtmlBodyFlag As Boolean
Public objHasSenderName As Boolean
Public objHasSentOnBehalfOf As Boolean
Public objDumpMade As Long                         ' last item dumped or -1 when never

' Note: all following are "not quite static" i.e. may depend on user options and can "grow over time"
Public objDftMatches As String
Public objSortMatches As String
Public objMaxAttrCount As Long                     ' max number defined slots; "nearly" invariant for each ObjectType
Public objMinAttrCount As Long                     ' number of Attrs before Recurrences/Exceptions
Public objSeqInImportant As Collection

'---------------------------------------------------------------------------------------
' Note   : When a New cObjDsc is created from cObjDsc, it provides no objClsRules, but
'          will then copy / clone the rule parts (from sRules) or make them if Nothing
'---------------------------------------------------------------------------------------
Public Sub ODescClone(aClassKey As String, Optional sITMDsc As cItmDsc)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cObjDsc.ODescClone"
    Call DoCall(zKey, tSub, eQzMode)

    If objClassKey <> Me.objClassKey Then
        objClassKey = aClassKey
        D_TC.Add aClassKey, Me
    End If
    Set aObjDsc = Me
    
    With aObjDsc
        If sITMDsc Is Nothing Then
            Set sITMDsc = Me
            .objDumpMade = -1
        Else
            If Not sITMDsc.idObjDsc Is Nothing Then
                DoVerify (objClassKey = sITMDsc.idObjDsc.objClassKey _
                            Or objClassKey = CStr(sITMDsc.idObjItem.Class) _
                        & sITMDsc.idObjDsc.objNameExt), _
                        "Keys not canonical ???"
            End If
            Set sITMDsc.idObjDsc = aObjDsc
        End If
        
        .objItemClass = Me.objItemClass
        .objTypeName = Me.objTypeName
        .objDefaultIdent = Me.objDefaultIdent
    
        If LenB(Me.objSortMatches) = 0 Then
            Me.objSortMatches = AllPublic.SortMatches
        End If
        .objDftMatches = Me.objDftMatches
        .objSortMatches = Me.objSortMatches
        
        If Me.objClsRules Is Nothing Then
            Set .objClsRules = dftRule
        End If
        If sITMDsc.idObjDsc Is Nothing Then
            Set sITMDsc.idObjDsc = aObjDsc
        End If
        Set .objClsRules = Me.objClsRules.AllRulesClone(ClassRules, sITMDsc.idObjDsc, True)
        Set .objClsRules.clsNeverCompare.PropAllRules = .objClsRules
        Set .objClsRules.clsObligMatches.PropAllRules = .objClsRules
        Set .objClsRules.clsNotDecodable.PropAllRules = .objClsRules
        Set .objClsRules.clsSimilarities.PropAllRules = .objClsRules
        Set .objSeqInImportant = Me.objSeqInImportant
        
        .objMaxAttrCount = Me.objMaxAttrCount
        .objMinAttrCount = Me.objMinAttrCount
        .objItemClassName = Me.objItemClassName
        .objIsMailLike = Me.objIsMailLike
        .objHasReceivedTime = Me.objHasReceivedTime
        .objHasHtmlBodyFlag = Me.objHasHtmlBodyFlag
        .objHasSenderName = Me.objHasSenderName
        .objTimeType = Me.objTimeType
        .objHasSentOnBehalfOf = Me.objHasSentOnBehalfOf
        .objDumpMade = Me.objDumpMade
    End With                                       ' aObjDsc

zExit:
    Call DoExit(zKey)

End Sub                                            ' cObjDsc.ODescClone

'---------------------------------------------------------------------------------------
' Method : Function IsSame
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsSame(other As cObjDsc, Optional showdiffs = False)
Const zKey As String = "cItemClsProps.IsSame"
    Call DoCall(zKey, tFunction, eQzMode)

Dim NrOfDiffs As Long
    
    If other Is Nothing Then                       ' invalid: return False
        GoTo zExit
    End If
    If Me Is other Then
        IsSame = True
        GoTo zExit
    End If

    ' incomplete=True stops some of the comparisons
    NrOfDiffs = 0
    If Not objItemClassName <> other.objItemClassName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " & "objItemClassName " _
                        & objItemClassName & "<>" & other.objItemClassName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    
    If objItemClass <> other.objItemClass Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objItemClass " & objItemClass & "<>" & other.objItemClass
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    
    If objItemClassName <> other.objItemClassName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objItemClassName " & objItemClassName & "<>" & other.objItemClassName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objItemType <> other.objItemType Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objItemType " & objItemType & "<>" & other.objItemType
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objIsMailLike <> other.objIsMailLike Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objIsMailLike " & objIsMailLike & "<>" & other.objIsMailLike
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objHasReceivedTime <> other.objHasReceivedTime Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objHasReceivedTime " & objHasReceivedTime & "<>" & other.objHasReceivedTime
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objHasHtmlBodyFlag <> other.objHasHtmlBodyFlag Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objHasHtmlBodyFlag " & objHasHtmlBodyFlag & "<>" & other.objHasHtmlBodyFlag
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objHasSenderName <> other.objHasSenderName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objHasSenderName " & objHasSenderName & "<>" & other.objHasSenderName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If objHasSentOnBehalfOf <> other.objHasSentOnBehalfOf Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "objHasSentOnBehalfOf " _
                        & objHasSentOnBehalfOf & "<>" & other.objHasSentOnBehalfOf
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    
    IsSame = True
    GoTo zExit
    
OneDiffIsEnough:
    IsSame = (NrOfDiffs = 0)

zExit:
    Call DoExit(zKey)

End Function                                       ' cItemClsProps.IsSame

'---------------------------------------------------------------------------------------
' Method : objDictClone
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Copy Class Model's Dictionary to a new Dictionary
'---------------------------------------------------------------------------------------
Function objDictClone(ModelDict As Dictionary) As Dictionary
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cObjDsc.objDictClone"
#If MoreDiagnostics Then
        Call DoCall(zKey, "Function", eQzMode)
#End If

Dim i As Long
Dim cloneDe As cObjDsc
Dim cloneID As cAttrDsc
Dim thisClone As Object
Dim adKey As String

    If ModelDict.Count < 2 Then
        GoTo zExit
    End If
    
    Set objDictClone = New Dictionary
    
    adKey = ModelDict.Items(0).objClassKey
    Set cloneDe = ModelDict.Items(0)
    objDictClone.Add adKey, cloneDe
    
    For i = 1 To ModelDict.Count - 1
        Set thisClone = ModelDict.Items(i)
        PropertyNameX = thisClone.adKey
        Set cloneID = thisClone.adictClone
        ' Set cloneID = New cAttrDsc
        ' cloneID.adKey = thisAD.adKey
        ' aTD.adtrueIndex = thisAD.adtrueIndex
        objDictClone.Add cloneID.adKey, cloneID
    Next i

    Set thisClone = Nothing
    Set cloneDe = Nothing
    Set cloneID = Nothing

zExit:
    Call DoExit(zKey)
ProcRet:
End Function                                       ' cObjDsc.objDictClone



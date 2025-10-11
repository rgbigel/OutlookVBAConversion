VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemClsProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public aCitemClass As OlObjectClass
Public aCitemClassName As String
Public aCITMDsc As cItmDsc                         ' Parent

Public aCitemType As OlItemType

Public aCisMailLike As Boolean                     ' some of these have missing Properties
Public aHasReceivedTime As Boolean
Public aHasHtmlBodyFlag As Boolean
Public aHasSenderName As Boolean
Public aHasSentOnBehalfOf As Boolean

'---------------------------------------------------------------------------------------
' Method : Function IsSame
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsSame(other As cItemClsProps, Optional showdiffs = False)
Const zKey As String = "cItemClsProps.IsSame"
    Call DoCall(zKey, tFunction, eQzMode)

Dim NrOfDiffs As Long
    
    If other Is Nothing Then                       ' invalid: return False
        GoTo zExit
    End If
    If aCITMDsc Is Nothing Then                    ' invalid: return False
        GoTo zExit
    End If
    If Me Is other Then
        IsSame = True
        GoTo zExit
    End If

    ' incomplete=True stops some of the comparisons
    NrOfDiffs = 0
    If Not aCITMDsc.idObjDsc.objItemClassName _
       <> other.aCITMDsc.idObjDsc.objItemClassName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aCITMDsc.idObjDsc.objItemClassName " _
                        & aCITMDsc.idObjDsc.objItemClassName _
                        & "<>" & other.aCITMDsc.idObjDsc.objItemClassName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    
    If aCitemClass <> other.aCitemClass Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aCitemClass " & aCitemClass _
                        & "<>" & other.aCitemClass
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    
    If aCitemClassName <> other.aCitemClassName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aCitemClassName " & aCitemClassName _
                        & "<>" & other.aCitemClassName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aCitemType <> other.aCitemType Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aCitemType " & aCitemType _
                        & "<>" & other.aCitemType
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aCisMailLike <> other.aCisMailLike Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aCisMailLike " & aCisMailLike _
                        & "<>" & other.aCisMailLike
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aHasReceivedTime <> other.aHasReceivedTime Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aHasReceivedTime " & aHasReceivedTime _
                        & "<>" & other.aHasReceivedTime
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aHasHtmlBodyFlag <> other.aHasHtmlBodyFlag Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aHasHtmlBodyFlag " & aHasHtmlBodyFlag _
                        & "<>" & other.aHasHtmlBodyFlag
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aHasSenderName <> other.aHasSenderName Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aHasSenderName " & aHasSenderName _
                        & "<>" & other.aHasSenderName
        Else
            GoTo OneDiffIsEnough
        End If
    End If
    If aHasSentOnBehalfOf <> other.aHasSentOnBehalfOf Then
        NrOfDiffs = NrOfDiffs + 1
        If showdiffs Then
            Debug.Print "  not same: " _
                        & "aHasSentOnBehalfOf " _
                        & aHasSentOnBehalfOf _
                        & "<>" & other.aHasSentOnBehalfOf
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
' Method : Clone
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Clone() As cItemClsProps

    Set Clone = New cItemClsProps
    
    Clone.aCitemType = aCitemType
    Set Clone.aCITMDsc = aCITMDsc
    
    Clone.aCisMailLike = aCisMailLike
    Clone.aHasReceivedTime = aHasReceivedTime
    Clone.aHasHtmlBodyFlag = aHasHtmlBodyFlag
    Clone.aHasSenderName = aHasSenderName
    Clone.aHasSentOnBehalfOf = aHasSentOnBehalfOf

End Function                                       ' cItemClsProps.Clone



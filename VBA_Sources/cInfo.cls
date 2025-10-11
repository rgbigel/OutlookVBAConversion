VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this describes most significant Object vars, inspecting Variants/Objects of any type

Public iAssignmentMode As Long                     ' 0/-99 - no value to assign/not initialized
' 1 - normal Assign
' 2 - using set
Public iClass As Long                              ' Class of original ItemProperty/Variable/Value to check, -99 if none
Public iType As VbVarType                          ' Type of original ItemProperty/Variable/Value
Public iTypeName As String                         ' TypeName of original
Public iScalarType As VbVarType                    ' Type corresponding to TypeName
Public iIsArray As Boolean                         ' value in ivalue is Array with Count
Public iArraySize As Long                          ' used by getInfo and TypeDecode - Procs
Public iValue As Variant                           ' This scalar, variant, Object
Public DecodedStringValue As String                ' if iValue can be converted to String, this is it
Public DecodeMessage As String                     ' if something must be said to Decoding...

Public iDepth As Long                              ' how many times we successfully set iValue = ivalue.Value
Public iUp As cInfo
Public iDown As cInfo

'---------------------------------------------------------------------------------------
' Method : Class_Initialize
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Defaults values for use in getInfo and TypeDecode Procs
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()

    iAssignmentMode = inv                          ' undefined
    
    iClass = inv                                   ' Class of original ItemProperty/Variable/Value
    iType = inv
    iTypeName = vbNullString                       ' TypeName of original
    iScalarType = inv
    iIsArray = False
    iArraySize = inv                               ' used by getInfo and TypeDecode - Procs. -2 means: is fresh.
    iValue = vbNullString                          ' value converted (or # CantDoThat) to a string
    
    iDepth = 0                                     ' how many times we successfully assign iValue = ivalue.Value
    Set iUp = Nothing
    Set iDown = Nothing
    Set iValue = Nothing                           ' This scalar, variant, Object
End Sub                                            ' cInfo.Class_Initialize

'---------------------------------------------------------------------------------------
' Method : Function DrillDown
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function DrillDown(aValue As Variant) As cInfo
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.DrillDown"
    Call DoCall(zKey, "Sub", eQzMode)

    iScalarType = IsScalar(TypeName(aValue))
    If iScalarType > 0 Then                        ' this does not always match iType
        iAssignmentMode = 1                        ' scalar
    ElseIf iScalarType < 0 Then
        iAssignmentMode = 0                        ' do not assign
    Else
        iAssignmentMode = 2                        ' not scalar, object
    End If
    
    Set DrillDown = New cInfo
    Set iDown = DrillDown
    
    With DrillDown
        Set .iUp = Me
        .iDepth = Me.iDepth + 1
        Select Case iAssignmentMode
            Case 0                                 ' do not assign
            Case 1
                .iType = VarType(aValue)
                .iTypeName = TypeName(aValue)
                .iValue = aValue                   ' scalar
            Case 2
                .iType = VarType(aValue)
                .iTypeName = TypeName(aValue)
                Set .iValue = aValue               ' not scalar, object
            Case Else
                DoVerify False, "invalid Assignment mode " & iAssignmentMode
        End Select                                 ' iAssignmentMode
    End With                                       ' DrillDown
    
zExit:
    Call DoExit(zKey)

End Function                                       ' cInfo.DrillDown

'---------------------------------------------------------------------------------------
' Method : Function Top
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Top() As cInfo
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.Top"
    Call DoCall(zKey, "Function", eQzMode)

    Set Top = Me
    While Not Top.iUp Is Nothing
        Set Top = Top.iUp
    Wend                                           ' Top -> iUp
    
zExit:
    Call DoExit(zKey)

End Function                                       ' cInfo.Top

'---------------------------------------------------------------------------------------
' Method : Function Find
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Find(aAttrName As String, Optional Direction As Long = 1) As cInfo
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.Find"
    Call DoCall(zKey, "Function", eQzMode)

    If Direction = 0 Then                          ' start search at top
        Set Find = Me.Top
        Direction = 1
    Else
        Set Find = Me                              ' start here
    End If
    
    While Find.iTypeName <> aAttrName
        If Direction = 1 Then                      ' follow up / else down
            Find = Me.iDown
            If Find Is Nothing Then
                GoTo zExit
            End If
        Else
            Find = Me.iUp
            If Find Is Nothing Then
                GoTo zExit
            End If
        End If
    Wend                                           ' Find.iTypeName <> aAttrName
    
zExit:
    Call DoExit(zKey)
    ' returns "Nothing" if not found
End Function                                       ' cInfo.Find

'---------------------------------------------------------------------------------------
' Method : Function IsArrayProperty
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function IsArrayProperty(PropValue As Variant, Optional vArraySize As Long) As Boolean
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.IsArrayProperty"
    Call DoCall(zKey, tFunction, eQzMode)

    Const toDepth As Long = 2
Dim haveDepth As Long
Dim PropObj As Object

    If DebugMode Then
        If vArraySize >= 0 Then
            DoVerify False, " ??? not expected, who determined this value???"
        End If
    End If
    vArraySize = -10000                            ' is not an array property
    
    Set PropObj = PropValue
GoDeep:
    With PropObj
        Select Case .Class
            Case olActions, olAttachments, olUserProperties, olLinks, _
                 olRecipients, olConflicts         ' not defined: , olReplyRecipients, olMemberCount
                aBugTxt = "value depth=" & haveDepth
                Call Try
                vArraySize = .Count
                If Not Catch Then
                    IsArrayProperty = True
                End If
                Set PropValue = PropObj
                GoTo FuncExit
            Case Else
                haveDepth = haveDepth + 1
                aBugTxt = "Get Object value, depth=" & haveDepth
                Call Try(allowAll)                    ' Try anything, autocatch, Err.Clear
                Set PropObj = PropValue.Value
                If Catch Then
                    GoTo FuncExit
                End If
                If haveDepth < toDepth Then
                    GoTo GoDeep
                End If
        End Select
    End With                                       ' PropValue.value
    
FuncExit:
    Set PropObj = Nothing
    
zExit:
    Call DoExit(zKey)

End Function                                       ' cInfo.IsArrayProperty

'---------------------------------------------------------------------------------------
' Method : Function ShowInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ShowInfo() As String
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.ShowInfo"

    If LenB(iTypeName) = 0 Then
        ShowInfo = "This cInfo has no values"
    Else
        ShowInfo = iDepth & " TypeName " & iTypeName & "(" & iType _
                   & "), ScalarType=" & iScalarType & ", Class="
        If iClass = inv Then
            ShowInfo = ShowInfo & "None, "
        Else
            ShowInfo = ShowInfo & iClass & ", "
        End If
        If iIsArray Then
            ShowInfo = ShowInfo & "is array with " & iArraySize & " elements, "
        Else
            If iArraySize > 0 Then
                ShowInfo = ShowInfo & "array element nr. " & iArraySize & ", "
            Else
                ShowInfo = ShowInfo & "no array, "
            End If
        End If
        If iAssignmentMode <= 0 Then
            ShowInfo = ShowInfo & "# not decoded or not decodable"
        ElseIf iAssignmentMode = 1 Then
            ShowInfo = ShowInfo & "AssignmentMode=1, value: '" & DecodedStringValue & "'"
        ElseIf isEmpty(iValue) Then
            ShowInfo = ShowInfo & "AssignmentMode=2, # Empty"
        ElseIf iValue Is Nothing Then
            ShowInfo = ShowInfo & "AssignmentMode=2, # Nothing"
        Else
            ShowInfo = ShowInfo & "AssignmentMode=2, # Object"
        End If
    End If

End Function                                       ' cInfo.ShowInfo

'---------------------------------------------------------------------------------------
' Method : Function ShowType
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function ShowType(Optional withLeads As String = "   |") As String
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "cInfo.ShowType"

Dim tInfo As cInfo
Dim aLead As String
Dim ArrayElements As Boolean
    
    Set tInfo = Me
    Do
        ShowType = ShowType & vbCrLf & aLead & tInfo.ShowInfo
        aLead = aLead & withLeads
        If tInfo.iIsArray Then
            ArrayElements = True
        End If
        If ArrayElements Then
            If tInfo.iArraySize = 0 Then
                aLead = Mid(aLead, Len(withLeads) + 1)
                ArrayElements = False
            End If
        End If
        Set tInfo = tInfo.iDown
    Loop Until tInfo Is Nothing
    
    ShowType = Mid(ShowType, 3)                    ' no initial vbCrLf

End Function                                       ' cInfo.ShowType

'---------------------------------------------------------------------------------------
' Method : Sub PrintType
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub PrintType()
    Debug.Print ShowType
End Sub                                            ' cInfo.printtype



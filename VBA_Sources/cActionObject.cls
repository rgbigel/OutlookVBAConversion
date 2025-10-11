VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActionObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public aoObjID As String
Attribute aoObjID.VB_VarUserMemId = 0
Attribute aoObjID.VB_VarDescription = "Display Object ID"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute aoObjID.VB_VarUserMemId = 0
' Attribute aoObjID.VB_VarDescription = "Display Object ID"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public IterationSettings As cIterationSwitches
Public ActionID As Long

Property Get Value() As String                     ' Default Property of the Class <self> ****
    Value = ActionTitle(ActionID) & ": " & aoObjID
End Property                                       ' cActionObject.Value Get

'---------------------------------------------------------------------------------------
' Method : assign
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Return Method S copied
'---------------------------------------------------------------------------------------
Public Function Assign(Source As cActionObject) As cActionObject

Const zKey As String = "cActionObject.Assign"
    Call DoCall(zKey, tFunction, eQzMode)

    ActionID = Source.ActionID
    Set IterationSettings = Source.Assign(Source.IterationSettings)
    aoObjID = Source.aoObjID

zExit:
    Call DoExit(zKey)

End Function                                       ' cActionObject.Assign


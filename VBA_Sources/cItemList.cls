VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Items As Collection

'---------------------------------------------------------------------------------------
' Method : Sub Add
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub Add(ItemDesc As cAttrDsc, Optional Clone As Boolean)

Const zKey As String = "cItemList.Add"
    Call DoCall(zKey, tSub, eQzMode)

Dim thisClone As cAttrDsc

    If GetDItem_P(Items(ItemDesc.adItem.EntryID)) Is Nothing Then
        If Clone Then
            Set thisClone = ItemDesc.adictClone
            Me.addItem thisClone
        Else
            Me.addItem ItemDesc
        End If
    Else
        DoVerify False, " item exists, can not add it"
    End If

zExit:
    Call DoExit(zKey)

End Sub                                            ' cItemList.Add

'---------------------------------------------------------------------------------------
' Method : Sub addItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub addItem(ItemDesc As cAttrDsc)

Const zKey As String = "cItemList.addItem"
    Call DoCall(zKey, tSub, eQzMode)

    Items.Add ItemDesc, ItemDesc.adItem.EntryID

zExit:
    Call DoExit(zKey)

End Sub                                            ' cItemList.addItem

'---------------------------------------------------------------------------------------
' Method : Function GetDItem_P
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Function GetDItem_P(EntryID As String) As cAttrDsc
Dim zErr As cErr
Const zKey As String = "cItemList.GetDItem_P"

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cItemList")
    aBugTxt = "could not get item " & EntryID
    Call Try                                       ' Try anything, autocatch
    Set GetDItem_P = Items.Item(EntryID)
    Catch

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cItemList.GetDItem_P


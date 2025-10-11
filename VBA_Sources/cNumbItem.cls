VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNumbItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NuIndex As Long
Public Key As String                               ' default value of class <self> ****
Public Alias As String
Public Subfields As String

Public ValueOfItem As Object

'---------------------------------------------------------------------------------------
' Method : Function Exists
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Exists(aColl As Collection, sKey As String, Optional aAlias As String) As Object

Const zKey As String = "cNumbItem.Exists"
    Call DoCall(zKey, tFunction, eQzMode)

Dim aitem As Object
    
    For Each aitem In aColl
        If IsMissing(aAlias) Then
keyCheck:
            If aitem.Key = sKey Then
                Set Exists = aitem
                GoTo FuncExit
            End If
        Else
            If aitem.Alias = aAlias Then
                If LenB(sKey) > 0 Then
                    GoTo keyCheck
                End If
                Set Exists = aitem
                GoTo FuncExit
            End If
        End If
    Next aitem

FuncExit:
    Set aitem = Nothing

zExit:
    Call DoExit(zKey)

End Function                                       ' cNumbItem.Exists


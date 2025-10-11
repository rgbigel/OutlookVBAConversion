VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTriState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim StateIs As Long

Public Property Get Value() As Long
    Value = StateIs
End Property                                       ' cTriState.Value Get

Public Property Let Value(nv As Long)
    Select Case nv
        Case TristateTrue
            StateIs = TristateTrue
        Case TriStateFalse
            StateIs = TriStateFalse
        Case TriStateMixed
            StateIs = TriStateMixed
        Case TriStatsUndefined
            StateIs = TriStatsUndefined
        Case Else
            If nv <> TriStatsUndefined Then
                rsp = MsgBox("Tristate variable has invalid value " & CStr(nv), vbOKCancel)
                If rsp = vbCancel Then
                    Err.Raise Hell
                End If
            End If
            StateIs = TriStatsUndefined
    End Select
End Property                                       ' cTriState.Value Let


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmProgressBox.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
' Implements a progress box with a progress bar, optional Cancel button and space for user-defined text above the bar/button
' Uses Microsoft's Forms library (by default available with all office/VBA installations)
' To use in your VBA project:
' 1) Make sure that the "Microsoft Forms" object library is checked in Tools/References
' 2) Insert a blank User Form
' 3) Rename the user form to "frmProgressBox"
' 4) Set the user form property "showModal" to false (so you can do other things while the dialog is displayed)
' 5) Add a command button to the form
' 6) Rename the command button to "CancelButton"
' 7) Show the code for the User Form, and highlight/delete everything
' 8) Insert this file (using insert/file) into the code for the User Form
'    From: http://www.outlookcode.com/codedetail.aspx?id=1880
' 9) Add appropriate code to your VBA routine where you want to Show progress:
'       * frmProgressBox.Show --- shows the progress box. Include this before starting processing.
'       * frmProgressBox.ShowPercent --- shows or hides the progress bar
'       * frmProgressBox.ShowCancel --- shows or hides the progress bar
'       * frmProgressBox.Increment newPercent (single), NewText (optional string) --- updates the progress bar and optionally changes the text and repaints
'       * frmProgressBox.Update NewText (string) --- changes the text and repaints
'       * frmProgressBox.Cancelled (property) --- should be tested at appropriate points in your processing - if true, the user pressed cancel
'       * frmProgressBox.Hide --- removes the progress bar. Include this at the end of processing.
' 10) Optionally, you can get/set the percentage and the text individually using the "Percent" and "Text" properties, followed by calling frmProgressBox.repaint
 
Private Const DefaultTitle = "Progress"
Private myText As String
Private myPercent As Single
Private myCancelled As Boolean
Private myShowPercent As Boolean
Private myShowCancel As Boolean
Private myHeight As Long                         ' stops dialog shrinking once it has expanded for multi lines of text

' Record that the cancel button was clicked
Private Sub CancelButton_Click()
    frmProgressBox.Cancelled = True
End Sub                                          ' frmProgressBox.CancelButton_Click

' Cancelled property resets/indicates that cancel button was pressed
Public Property Let Cancelled(newCancelled As Boolean)
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Cancelled"
    Call ProcCall(zErr, zKey, eQxMode, tPropLet, vbNullString)

    If newCancelled <> myCancelled Then
        myCancelled = newCancelled
        Call updateCancelled
    End If
    doMyEvents

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Cancelled Let

Public Property Get Cancelled() As Boolean
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Cancelled"
    Call ProcCall(zErr, zKey, eQxMode, tPropGet, vbNullString)

    Cancelled = myCancelled

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Cancelled Get

' Increment method enables the percent and optionally the text to be updated at same time
Public Sub Increment(ByVal newPercent As Single, Optional ByVal NewText As String)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.Increment"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

    Me.Percent = newPercent
    If LenB(NewText) > 0 Then
        Me.Text = NewText
    End If
    Call updateTitle
    doMyEvents
    Me.Repaint

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.Increment

' Percent property alters progress shown on the progress bar
Public Property Let Percent(newPercent As Single)
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Percent"
    Call ProcCall(zErr, zKey, eQxMode, tPropLet, vbNullString)

    If newPercent <> myPercent Then              ' limit percent to between 0 and 100
        myPercent = Min(Max(newPercent, 0#), 100#)
        Call updateProgress
    End If
    doMyEvents

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Percent Let

Public Property Get Percent() As Single
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Percent"
    Call ProcCall(zErr, zKey, eQxMode, tPropGet, vbNullString)

    Percent = myPercent

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Percent Get

' Removes any current controls, add the needed controls ...
Private Sub setupControls()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.setupControls"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

Dim i As Long
Dim aControl As Object

    ' Setup cancel button (should already exist on the form)
    Set aControl = Me.Controls("CancelButton")
  
    With aControl
        .Tag = "CancelButton"
        .Caption = "Cancel"
        .AutoSize = False
        .height = 20
        .width = 60
        .Font.Size = 8
        .TakeFocusOnClick = False
    End With
    ' remove existing controls - all except the cancel button
    For i = Me.Controls.Count - 1 To 0 Step -1
        If Me.Controls(i).Tag <> "CancelButton" Then Me.Controls(i).Remove
    Next i
    ' add user text - don't worry about positioning as "sizeToFit" takes care of this
    Set aControl = Me.Controls.Add("Forms.Label.1", "UserText", True)
    With aControl
        .Caption = vbNullString
        .AutoSize = True
        .WordWrap = True
        .Font.Size = 8
    End With
    ' add progressFrame - don't worry about positioning as "sizeToFit" takes care of this
    Set aControl = Me.Controls.Add("Forms.Label.1", "ProgressFrame", True)
    With aControl
        .Caption = vbNullString
        .height = 20
        .SpecialEffect = fmSpecialEffectSunken
    End With
    ' add user text - don't worry about positioning as "sizeToFit" takes care of this
    Set aControl = Me.Controls.Add("Forms.Label.1", "ProgressBar", True)
    With aControl
        .Caption = vbNullString
        .height = 18
        .BackStyle = fmBackStyleOpaque
        .BackColor = &HFF0000                    ' Blue
    End With
    ' position the controls and size the progressBox
    Call sizeToFit
ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.setupControls

' ShowCancelled property shows or hides the cancel button
Public Property Let ShowCancel(newShowCancel As Boolean)
Dim zErr As cErr
Const zKey As String = "frmProgressBox.ShowCancel"
    Call ProcCall(zErr, zKey, eQxMode, tPropLet, vbNullString)

    If newShowCancel <> myShowCancel Then
        ' Show the cancel button
        myShowCancel = newShowCancel
        If newShowCancel Then myCancelled = False ' Button is being added, start with cancel not pressed
        If Not newShowCancel Then myHeight = 0   ' Button is being removed, force re-evaluation of height
        Call sizeToFit
    End If
    doMyEvents

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.ShowCancel Let

Public Property Get ShowCancel() As Boolean
Dim zErr As cErr
Const zKey As String = "frmProgressBox.ShowCancel"
    Call ProcCall(zErr, zKey, eQxMode, tPropGet, vbNullString)

    ShowCancel = myShowCancel

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.ShowCancel Get

' ShowPercent property shows or hides the progress bar
Public Property Let ShowPercent(newShowPercent As Boolean)
    If newShowPercent <> myShowPercent Then
        ' Show the progress bar
        myShowPercent = newShowPercent
        If Not newShowPercent Then myHeight = 0  ' Progress bar is being removed, force re-evaluation of height
        Call sizeToFit
    End If
    doMyEvents
End Property                                     ' frmProgressBox.ShowPercent Let

Public Property Get ShowPercent() As Boolean
    ShowPercent = myShowPercent
End Property                                     ' frmProgressBox.ShowPercent Get

' Adjusts positioning of controls/size of form depending on size of user text
Private Sub sizeToFit()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.sizeToFit"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

Dim newHeight As Long
 
    Me.width = 300                               ' setup width of progress box
    
    ' user-supplied text should be topmost, taking up the appropriate size ...
    Me.Controls("UserText").Top = 6
    Me.Controls("UserText").Left = 6
    Me.Controls("UserText").AutoSize = False
    Me.Controls("UserText").Font.Size = 8
    Me.Controls("UserText").width = Me.InsideWidth - 12
    Me.Controls("UserText").AutoSize = True
    ' Cancel button should be below user text, if visible
    If myShowCancel Then
        Me.Controls("CancelButton").Visible = True
        Me.Controls("CancelButton").Top = Me.Controls("UserText").Top + Me.Controls("UserText").height + 6
        Me.Controls("CancelButton").Left = Me.InsideWidth - Me.Controls("CancelButton").width - 6
        Call updateCancelled                     ' update Cancel button text/enabled
        ' finally, height of progress box should fit around text and Cancel button & allow for title/box frame
        newHeight = Max(newHeight, Me.Controls("CancelButton").Top + Me.Controls("CancelButton").height + 6 + (Me.height - Me.InsideHeight))
    Else
        Me.Controls("CancelButton").Visible = False
        ' finally, height of progress box should fit around text & allow for title/box frame
        newHeight = Max(newHeight, Me.Controls("UserText").Top + Me.Controls("UserText").height + 6 + (Me.height - Me.InsideHeight))
    End If
    ' progress frame/bar should be below user text (if visible) and to left of Cancel button
    If myShowPercent Then
        Me.Controls("ProgressFrame").Visible = True
        Me.Controls("ProgressFrame").Top = Me.Controls("UserText").Top + Me.Controls("UserText").height + 6
        Me.Controls("ProgressFrame").Left = 6
        If myShowCancel Then
            Me.Controls("ProgressFrame").width = Me.Controls("CancelButton").Left - 12
        Else
            Me.Controls("ProgressFrame").width = Me.InsideWidth - 12
        End If
        Me.Controls("ProgressBar").Visible = True
        Me.Controls("ProgressBar").Top = Me.Controls("ProgressFrame").Top + 1
        Me.Controls("ProgressBar").Left = Me.Controls("ProgressFrame").Left + 1
        Call updateProgress                      ' update ProgressBar width
        ' finally, height of progress box should fit around text and progress bar & allow for title/box frame
        newHeight = Max(newHeight, Me.Controls("ProgressFrame").Top + Me.Controls("ProgressFrame").height + 6 + (Me.height - Me.InsideHeight))
    Else
        Me.Controls("ProgressFrame").Visible = False
        Me.Controls("ProgressBar").Visible = False
        ' finally, height of progress box should fit around text & allow for title/box frame
        newHeight = Max(newHeight, Me.Controls("UserText").Top + Me.Controls("UserText").height + 6 + (Me.height - Me.InsideHeight))
    End If
    ' Don't decrease the height once it's been set/increased ...
    myHeight = Max(myHeight, newHeight)
    Me.height = myHeight
    ' And ensure that the progress bar and cancel button (if visible) stay stuck to the bottom
    If myShowPercent Then
        Me.Controls("ProgressFrame").Top = Me.InsideHeight - 6 - Me.Controls("ProgressFrame").height
        Me.Controls("ProgressBar").Top = Me.Controls("ProgressFrame").Top + 1
    End If
    If myShowCancel Then
        Me.Controls("CancelButton").Top = Me.InsideHeight - 6 - Me.Controls("CancelButton").height
    End If

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.sizeToFit

' Text property shows user-defined text above the progress bar
Public Property Let Text(NewText As String)
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Text"
    Call ProcCall(zErr, zKey, eQxMode, tPropLet, vbNullString)

    If NewText <> myText Then
        myText = NewText
        Me.Controls("UserText").Caption = myText
        Call sizeToFit
    End If
    doMyEvents

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Text Let

Public Property Get Text() As String
Dim zErr As cErr
Const zKey As String = "frmProgressBox.Text"
    Call ProcCall(zErr, zKey, eQxMode, tPropGet, vbNullString)

    Text = myText

ProcReturn:
    Call ProcExit(zErr)

End Property                                     ' frmProgressBox.Text Get

' Update method enables the text to be updated (with a repaint)
Public Sub Update(ByVal NewText As String)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.Update"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

    Me.Text = NewText
    Call updateTitle
    doMyEvents
    Me.Repaint

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.Update

' updates the text and whether button is enabled based on private variables
Private Sub updateCancelled()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.updateCancelled"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

    If Not myShowCancel Then
        Me.Controls("CancelButton").Visible = False
    Else
        Me.Controls("CancelButton").Visible = True
        If Not myCancelled Then
            Me.Controls("CancelButton").Caption = "Cancel"
            Me.Controls("CancelButton").Enabled = True
        Else
            Me.Controls("CancelButton").Caption = "Cancelling ..."
            Me.Controls("CancelButton").Enabled = False
        End If
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.updateCancelled

' updates the width of the progress bar to Match the current percentage
Private Sub updateProgress()
    If (Not myShowPercent) Or (myPercent = 0) Then
        Me.Controls("ProgressBar").Visible = False
    Else
        Me.Controls("ProgressBar").Visible = True
        Me.Controls("ProgressBar").width = Int((Me.Controls("ProgressFrame").width - 2) * myPercent / 100)
    End If
End Sub                                          ' frmProgressBox.updateProgress

' updates the caption of the progress box to keep track of progress
Private Sub updateTitle()

Const zKey As String = "frmProgressBox.updateTitle"
    If myShowPercent Then
        If (Int(myPercent) Mod 5) = 0 Then
            Me.Caption = DefaultTitle & " - " & Format(Int(myPercent), "@@") & "% Complete"
        End If
    Else
        Me.Caption = DefaultTitle
    End If
End Sub                                          ' frmProgressBox.updateTitle

' Setup the progress dialog - title, control layout/size etc.
Private Sub UserForm_Initialize()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmProgressBox.UserForm_Initialize"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)
 
    myShowPercent = False
    myShowCancel = False
    myCancelled = False
    Call setupControls
    Call updateTitle

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmProgressBox.UserForm_Initialize

' Prevents use of the Close button. BOTH PARMS MUST BE Integer
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub                                          ' frmProgressBox.UserForm_QueryClose


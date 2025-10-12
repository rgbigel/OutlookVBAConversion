# Converted from frmLongText.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLongText
# Caption         =   "FormLongText"
# ClientHeight    =   7860
# ClientLeft      =   45
# ClientTop       =   375
# ClientWidth     =   9960.001
# OleObjectBlob   =   "frmLongText.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmLongText"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit
# Const dftReplCrLf As String = " ||  "
# Public HasChanged As Boolean
# Dim HasCrLfs As Boolean
# Dim HasReplCrLfs As Boolean
# Dim HasFold As Variant
# Dim IsFolded As Boolean
# Dim FoldBy As Variant
# Dim OriginalText As String
# Dim NewText As String
# Dim ReplacementCrLf As String
# Dim Locked As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Sub SetText
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub SetText(Text As String, Optional FoldAt As Variant = vbNullString, Optional ReplCrLf As String = dftReplCrLf)
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmLongText.SetText"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# OriginalText = Text
# ReplacementCrLf = ReplCrLf
# NewText = Text                               ' force nochange in GetState below
# Call GetState
# NewText = Fold(Text, FoldAt, ReplCrLf)
# EditText.Value = NewText
# TestCriteriaEditing = vbNo
# Locked = False

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function Fold
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function Fold(Text As String, Optional FoldAt As Variant = vbNullString, Optional ReplCrLf As String = dftReplCrLf) As String
# Dim i As Long

# ' for beauty only: fold for display by splitting
# HasFold = Replace(FoldAt, vbCrLf, ReplCrLf)  ' FoldAt must not contain vbCrLf ever
# HasFold = split(HasFold, ",")                ' can pass a csv
try:
    if LenB(FoldAt) = 0 Then:
    # FoldBy = vbNullString
    # Fold = Text
    # IsFolded = False
    else:
    # IsFolded:
    # Err.Clear
    # FoldBy = HasFold                         ' full array now
    # Fold = Text
    # FoldBy(i) = vbCrLf & HasFold(i)
    if LenB(HasFold(i)) > 0 Then:
    # IsFolded = True
    # Fold = Replace(Fold, HasFold(i), FoldBy(i))

# '---------------------------------------------------------------------------------------
# ' Method : Function UnFold
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function UnFold(Text As String) As String
# Dim i As Long
if IsFolded Then:
# UnFold = Text
if LenB(HasFold(i)) > 0 Then:
# UnFold = Replace(UnFold, FoldBy(i), HasFold(i))
else:
# UnFold = Text
# IsFolded = False

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserMsg
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub UserMsg(Text As String)
# Me.MsgToUser = Text

# '---------------------------------------------------------------------------------------
# ' Method : Sub GetState
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub GetState()
if InStr(NewText, vbCrLf) > 0 Then:
# HasCrLfs = True
else:
# HasCrLfs = False
if InStr(NewText, ReplacementCrLf) > 0 Then:
# HasReplCrLfs = True
else:
# HasReplCrLfs = False

# '---------------------------------------------------------------------------------------
# ' Method : Function EditedText
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function EditedText() As String
# Dim zErr As cErr
# Const zKey As String = "frmLongText.EditedText"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

# EditedText = UnFold(EditText.Value)

# ProcReturn:
# Call ProcExit(zErr)

# ' this is the important one!
# Public Function TextEdit(Text As String, Optional FoldAt As Variant = vbNullString, Optional ReplCrLf As String = dftReplCrLf) As String
# Dim zErr As cErr
# Const zKey As String = "frmLongText.TextEdit"
# Call ProcCall(zErr, zKey, Qmode:=eQAsMode, CallType:=tFunction)

# Call SetText(Text, FoldAt, ReplCrLf)
# TextEdit = EditText.Value                    ' returnvalue is current (folded) text
# ' allow more functions before we actually Show the form. Use .Show to Show it
# TestCriteriaEditing = vbOK

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub AcceptEdits_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub AcceptEdits_Click()
# Call EditText_AfterUpdate
# Call formCloser
# TestCriteriaEditing = vbOK                   ' do not ask again

# '---------------------------------------------------------------------------------------
# ' Method : Sub CancelEdits_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub CancelEdits_Click()
# Locked = True
# NewText = OriginalText
# EditText.Value = OriginalText
# Call formCloser
# TestCriteriaEditing = vbCancel

# '---------------------------------------------------------------------------------------
# ' Method : Sub EditText_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EditText_AfterUpdate()
if Not Locked Then:
if EditText.Value <> NewText Then:
# NewText = EditText.Value

# '---------------------------------------------------------------------------------------
# ' Method : Sub EditYes_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EditYes_Click()
# Call EditText_AfterUpdate
# Call formCloser
# TestCriteriaEditing = vbYes                  ' do ask again

# '---------------------------------------------------------------------------------------
# ' Method : Sub FoldUnfold_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub FoldUnfold_Click()
if HasCrLfs Then:
# NewText = Replace(EditText.Value, vbCrLf, ReplacementCrLf)
if NewText <> EditText.Value Then:
# EditText.Value = NewText
elif HasReplCrLfs Then:
# NewText = Replace(EditText.Value, ReplacementCrLf, vbCrLf)
if NewText <> EditText.Value Then:
# EditText.Value = NewText
# Call GetState


# '---------------------------------------------------------------------------------------
# ' Method : Sub formCloser
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub formCloser()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmLongText.formCloser"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# Call GetState
# Me.Hide
# Locked = False
# Me.MsgToUser = vbNullString

# ProcReturn:
# Call ProcExit(zErr)


# Private Sub UserForm_Initialize()
if TestCriteriaEditing = 0 Then:
if TestCriteriaEditing = vbCancel Then DoVerify False:

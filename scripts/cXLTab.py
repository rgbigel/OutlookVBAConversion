# Converted from cXLTab.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cXLTab"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public xlTName As String                           ' this is the string value
# Attribute xlTName.VB_VarUserMemId = 0
# Attribute xlTName.VB_VarDescription = "Excel Table Name"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute xlTName.VB_VarUserMemId = 0
# ' Attribute xlTName.VB_VarDescription = "Excel Table Name"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public xlTWBook As WorkBook
# Public xlTSheet As Worksheet                       ' default value of class <self> ****
# Public xlTHeadline As String
# Public xlTHead As Variant                          ' xlTHeadLine split into columns
# Public xlTLastLine As Long
# Public xlTLastCol As Long
# Public xlTabIsEmpty As Long                        ' 0   = undefined,
# ' 1   = non-empty:defined, with headline
# ' -1  = empty:defined
# ' >1  last row added with addline

# '---------------------------------------------------------------------------------------
# ' Method : Function GetLastLine
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get the last line via UsedRange
# '---------------------------------------------------------------------------------------
def getlastline():
    # Const zKey As String = "cXLTab.GetLastLine"

    # xlTSheet.Select
    # GetLastLine = xlTSheet.UsedRange.Rows.Count + xlTSheet.UsedRange.Row - 1


# '---------------------------------------------------------------------------------------
# ' Method : Sub RemoveEmpty
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Delete Entire Rows with empty content in column C(dft=1) beginning Row(dft=2)
# '---------------------------------------------------------------------------------------
def removeempty():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "cXLTab.RemoveEmpty"

    # Call DoCall(zKey, "Sub", eQzMode)

    # Dim CellValueInspected As String
    # Dim DelCount As Long
    # Dim RowNo As Long
    # Dim SUstate As Boolean

    # SUstate = xlApp.ScreenUpdating
    if DebugMode Then:
    # Application.ScreenUpdating = True
    else:
    # Application.ScreenUpdating = False
    # With xlTSheet
    # NoStep:
    if Me.xlTLastLine < RowNo Then:
    # Exit For
    # CellValueInspected = Trim(.Cells(RowNo, ColNo))
    if LenB(CellValueInspected) = 0 Then:
    # .Cells(RowNo, ColNo).Select
    # .Rows(RowNo).EntireRow.Delete
    # DelCount = DelCount + 1
    # Me.xlTLastLine = Me.xlTLastLine - 1
    if Me.xlTLastLine < 1 Then:
    # Exit For
    # GoTo NoStep
    if DelCount Mod 50 = 1 Then:
    # Application.StatusBar = "Removed " & DelCount _
    # & " Lines in " & .Name & ", now " _
    # & Me.xlTLastLine - RowNo

    # Application.ScreenUpdating = SUstate
    if DebugMode Or DebugLogging Then:
    # Call LogEvent(DelCount _
    # & " Lines with Empty Col=" & ColNo _
    # & " removed in " & .Name & ", now LastLine=" _
    # & Me.xlTLastLine)
    # End With                                       ' xlTSheet

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function ExcelSheetIsEmpty
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def excelsheetisempty():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "cXLTab.ExcelSheetIsEmpty"
    # Const zKey As String = "cXLTab.ExcelSheetIsEmpty"

    # Call DoCall(zKey, "Function", eQzMode)

    # Dim AR As Range

    try:
        if xlTSheet Is Nothing Then:
        # GoTo isEmpty
        else:
        # xlTSheet.Activate
        if xlTSheet.Name = "Tabelle1" Then             ' *** correct this logic to set proper sheet:
        # GoTo isEmpty

        # Call Try                                       ' Try anything, autocatch, Err.Clear
        # xlTLastLine = xlTSheet.UsedRange.Rows.Count + xlTSheet.UsedRange.Row - 1
        # xlTLastCol = xlTSheet.UsedRange.columns.Count + xlTSheet.UsedRange.Column - 1
        # aBugTxt = "get Last cell in sheet " & xlTSheet.Name
        # Set AR = xlTSheet.Range(Cells(1, 1), Cells(xlTLastLine, xlTLastCol))
        if Catch Then:
        # isEmpty:
        # Call ErrReset(0)
        # xlTabIsEmpty = 0                           ' undefined
        # ExcelSheetIsEmpty = True                   ' and empty= -1
        else:
        if AR.Row <= 2 _:
        # Or (AR.Row + AR.Column = 2 _
        # And Trim(AR.Cells(2, 1).Value) = vbNullString) Then
        # ExcelSheetIsEmpty = True
        # ' myTabIsEmpty probably is = 1: defined and not empty

        # zExit:
        # Call DoExit(zKey)
        # ProcRet:

# ' Dim Hdl As Variant -> O.xlTHead
# ' Dim sHdl As String -> O.xlTHeadline

# Property Let xHdl(headline As String)
if LenB(xlTSheet.Cells(1, 1).Text) = 0 Then:
# xlTHeadline = vbNullString
if xlTHeadline <> headline Then:
# xlTHead = split(Chr(0) & b & headline, b)
# Call addLine(Me, 0, xlTHead)               ' 0 is headline
# xlTHeadline = headline
# End Property                                       ' cXlTab.xHdl Let


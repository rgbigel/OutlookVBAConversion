# Converted from XLAccessingOL.py

# Attribute VB_Name = "XLAccessingOL"
# Option Explicit

# Public xlApp As Object                           ' truly it should be Object/Excel.Application
# Public xlA As Excel.WorkBook                     ' Selected workbook (=file)
# Public xlAS As Dictionary                        ' Korrespondiert zu xlA.Sheets: Worksheets
# Public xlC As Excel.WorkBook                     ' Possible second Workbook instance
# Public xlCS As Dictionary                        ' Korrespondiert zu xlC.Sheets: Worksheets
# ' quick access sheets
# Public E As cXLTab                               ' Editing RuleTableRule sheet
# Public O As cXLTab                               ' selektiertes sheet "Objekteigenschaften"
# Public S As cXLTab                               ' Table showing allSchemata rules
# Public W As cXLTab                               ' Which sheet=<ItemTypeName>
# Public x As cXLTab                               ' Aktuelle cXlTab gem SheetName

# Public xLMainWindowHdl As cFindWindowParms       ' Handle des Excel Hauptfensters

# '---------------------------------------------------------------------------------------
# ' Method : xlOlClear
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Clear associated variables for xlApp
# '---------------------------------------------------------------------------------------
def xlolclear():
    # '''' Proc Must ONLY CALL Z_-, Y_- or Y_Type PROCS                         ' trivial proc
    # Const zKey As String = "XLAccessingOL.xlOlClear"

    # Call DoCall(zKey, "Sub", eQzMode)

    # Set xlA = Nothing                            ' Selected workbook (=file)
    # Set xlAS = Nothing                           ' Korrespondiert zu xlA.Sheets: Worksheets
    # Set xlC = Nothing                            ' Possible second Workbook instance
    # Set xlCS = Nothing                           ' Korrespondiert zu xlC.Sheets: Worksheets
    # ' quick access sheets
    # Set E = Nothing                              ' Editing RuleTableRule sheet
    # Set O = Nothing                              ' selektiertes sheet "Objekteigenschaften"
    # Set S = Nothing                              ' Table showing allSchemata rules
    # Set W = Nothing                              ' Which sheet=<ItemTypeName>
    # Set x = Nothing                              ' Aktuelle cXlTab gem SheetName

    # Set xLMainWindowHdl = Nothing                ' Handle des Excel Hauptfensters

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function CheckExcelOK
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def checkexcelok():
    # Dim zErr As cErr
    # Const zKey As String = "cXlObject.CheckExcelOK"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cXlObject")

    if xlApp Is Nothing Then:
    # Call KillExcel
    elif TypeName(xlApp) <> "Application" Then:
    # Call KillExcel
    elif Not xlA Is Nothing Then:
    if xlA.ActiveSheet Is Nothing Then:
    # Call KillExcel
    else:
    # CheckExcelOK = True

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ClearSheetLines
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clearsheetlines():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cXlObject.ClearSheetLines"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    # Dim oldcontent As Range
    # Dim eVinit As Boolean
    # Dim ws As Worksheet
    # Dim ns As Worksheet
    # Dim RS As Worksheet
    # Dim isVis As Boolean
    # Dim ShNam As String

    # Set ws = xw.xlTSheet
    # ws.Activate
    # xw.xlTLastLine = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    if FromLine = 0 Then:
    # xw.xlTLastLine = 1
    # xw.xlTLastCol = 0
    # GoTo eraseIt
    # xw.xlTLastLine = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    # xw.xlTabIsEmpty = xw.xlTLastLine             ' defined 3-state variable: is empty if <= 1
    if xw.xlTLastLine <= 1 Then:
    # xw.xlTabIsEmpty = 1
    # GoTo ProcReturn

    # isVis = xlApp.Visible
    # eVinit = xlApp.EnableEvents
    # xlApp.EnableEvents = True
    # ws.Cells(2, 1).Value = "ClearMe"
    # ws.Cells(2, 1).Select                        ' this triggers eventroutine in Excel which clears WS
    # xlApp.EnableEvents = False
    # xlApp.AutoRecover.Enabled = False

    # ws.Activate
    # Set oldcontent = ws.UsedRange
    # xw.xlTLastLine = oldcontent.Rows.Count + oldcontent.Row - 1
    # xw.xlTLastCol = oldcontent.columns.Count + oldcontent.Column - 1

    if xw.xlTLastLine > 1 Then                   ' not empty except headline:
    # ShNam = ws.Name
    if ShNam = "Objekteigenschaften" Then:
    if Not ShutUpMode Then:
    print(Debug.Print "Col, Row of last cell = ", xw.xlTLastCol, xw.xlTLastLine, "Restoring Sheet")
    # Set RS = ws.Parent.Sheets("OEdefault")
    # RS.Copy Before:=ws
    # Set ns = ws.Parent.Sheets("OEdefault (2)")
    else:
    if Not ShutUpMode Then:
    print(Debug.Print "Col, Row of last cell = ", xw.xlTLastCol, xw.xlTLastLine, "Clearing Sheet")
    # Set ns = ws.Parent.Sheets.Add        ' no code in this sheet
    # eraseIt:
    # ws.Activate
    # xlApp.DisplayAlerts = False
    # ws.Delete
    # xlApp.DisplayAlerts = True
    # ns.Name = ShNam
    # ns.Visible = xlSheetVisible
    # Set xw.xlTSheet = ns
    # ns.Activate
    # ' housekeeping
    if Not (DebugMode Or DebugLogging) Then:
    # xlApp.Visible = isVis
    # xw.xlTabIsEmpty = 1
    # xlApp.EnableEvents = eVinit

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ClearWorkSheet
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clearworksheet():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cXlObject.ClearWorkSheet"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    # DoVerify WBook Is ws.xlTWBook, "Design Check on correct Worksheet parent ???"
    if xlApp Is Nothing Then:
    # Set WBook = Nothing
    # Set ws = Nothing
    # GoTo ProcReturn
    elif WBook Is Nothing Then:
    # Set ws = Nothing
    # GoTo ProcReturn
    elif ws Is Nothing Then:
    # Call CreateOrUse(WBook, WBook.Sheets(1).Name, ws)
    if DebugMode Then:
    # DoVerify False, "about to clear sheet " & ws.xlTSheet.Name & " are you shure???"
    # GoTo usit
    else:
    # usit:
    # Call ClearSheetLines(ws, 0)

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub CreateOrUse
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Uses XLB workbook, Set cu.xlTSheet, makes or uses Sheet(Sheetname), Set X==cu
# '          Only valid for xlAS matching WBook
# '---------------------------------------------------------------------------------------
def createoruse():
    # Dim zErr As cErr
    # Const zKey As String = "cXlObject.CreateOrUse"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cXlObject")

    if cu Is Nothing Then                        ' we must find it:
    # Set cu = WorksheetDescriptorFind(xlAS, Sheetname)
    else:
    if cu.xlTSheet.Name <> Sheetname Then    ' find only if not the right one:
    # Set cu = WorksheetDescriptorFind(xlAS, Sheetname)
    if cu Is Nothing Then                        ' not found: make a new Sheet by that name:
    # Set cu = New cXLTab
    # Set cu.xlTWBook = WBook
    # Set cu.xlTSheet = xlA.Sheets.Add(After:=xlA.Sheets.Item(xlA.Sheets.Count))
    # cu.xlTSheet.Name = Sheetname
    # xlAS.Add Sheetname, cu.xlTSheet          ' add to Sheet list xlAS
    # Set x = cu
    # cu.xlTSheet.Select

    # cu.xlTSheet.EnableCalculation = False        ' default for our sheets
    # cu.xlTSheet.EnableFormatConditionsCalculation = False

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def displayexcel():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cXlObject.DisplayExcel"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    if unconditionallyShow Or displayInExcel Then:
    if Not xlW Is Nothing Then:
    if IsMissing(xlY) Or xlY Is Nothing Then:
    # Set xlY = xlW.xlTSheet           ' default worksheet of xlW=clsXlTab
    if Not xlApp Is Nothing Then:
    if xlY Is Nothing Then               ' ActiveWorkbook:
    # DoVerify False
    else:
    # xlApp.Visible = True
    # xlY.Cells(1, 20).Value = relevant_only
    # xlY.EnableCalculation = True
    # xlY.EnableFormatConditionsCalculation = True
    # xlApp.EnableEvents = False
    # xlY.Activate
    if xlApp.Visible Then:
    # xlApp.EnableEvents = EnableEvents
    else:
    # xlY.Activate
    # xlApp.EnableEvents = EnableEvents
    if xlApp.EnableEvents Then:
    if Not (DebugMode Or DebugLogging) Then:
    # xlApp.ScreenUpdating = False
    # xlApp.Cursor = xlWait
    else:
    # xlApp.ScreenUpdating = True
    # xlApp.Cursor = xlDefault

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub EndAllWorkbooks
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def endallworkbooks():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cXlObject.EndAllWorkbooks"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    # Dim onlyMySheets As Boolean

    # onlyMySheets = True
    if xlApp Is Nothing Then:
    # GoTo ProcReturn
    if xlApp Is Nothing Then:
    # GoTo ProcReturn
    # xlApp.EnableEvents = False
    # xlApp.AutoRecover.Enabled = True
    # xlApp.Cursor = xlDefault
    if Workbooks.Count > 0 Then:
    # xlApp.CalculateBeforeSave = True
    # xlApp.Calculation = xlCalculationAutomatic
    for xla in xlapp:
    # aBugTxt = "close Workbook " & Quote(xlA.Name)
    if InStr(UCase(xlA.Name), "PERSONAL") > 0 Then:
    elif InStr(UCase(xlA.Name), "OUTLOOKDEFAULT") > 0 Then:
    # aBugTxt = "Close WorkBook " & xlA.Name
    # Call Try                             ' Try anything, autocatch
    # xlA.Close False                      ' error here never matter
    elif Left(xlA.Name, 5) = "Mappe" Then:
    # Call Try                             ' Try anything, autocatch
    # xlA.Close False                      ' error here never matter
    else:
    # onlyMySheets = False
    # XlOpenedHere = False
    # Catch

    # Call xlOlClear

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub xlEndApp
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlendapp():
    # Const zKey As String = "XLAccessingOL.xlEndApp"
    # Static zErr As New cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="XLAccessingOL")

    # Call EndAllWorkbooks
    if XlOpenedHere Or xlApp.Workbooks.Count = 0 Then:
    # aBugTxt = "Close/Quit Excel"
    # Call Try                                 ' Try anything, autocatch
    # xlApp.Quit
    if Not Catch Then:
    if Not ShutUpMode Then:
    print(Debug.Print "Excel has been closed successfully")
    # Call xlOlClear
    # Set xlApp = Nothing
    else:
    # Call LogEvent("Excel not closed because (we know) other workbooks are open")
    # Call ErrReset(0)

    # XlOpenedHere = False

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Function xlWBInit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlwbinit():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cXlObject.xlWBInit"
    # Dim zErr As cErr

    # Dim RetryCounter As Long
    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    # Call ErrReset(4)
    if xlApp Is Nothing Then:
    # Call xlOlClear
    # ExcelDied:
    # Call ErrReset(4)
    # Call XlgetApp
    if xlApp Is Nothing Then:
    # Call LogEvent("unable to obtain Excel Application Object, Retries=" & RetryCounter)
    # GoTo Retry
    elif TypeName(xlApp) <> "Application" Then:
    # Retry:
    # Call ErrReset(4)
    # RetryCounter = RetryCounter + 1
    # Call xlEndApp
    if RetryCounter > 4 Then:
    # Call LogEvent("Unable to use Excel with " & TemplateFile)
    # GoTo ExcelDied

    if xlA Is Nothing Or TypeName(xlA) = "Object" Then:
    if xlApp.Workbooks.Count = 0 Then:
    # Call LogEvent("Excel open, but no Workbooks open, need: " & TemplateFile)
    # GoTo useApplication                  ' open one
    for xla in xlapp:
    if xlA.FullName = TemplateFile Then:
    if DebugLogging Then:
    # Call LogEvent("Found the correct Workbook for xlA: " & TemplateFile)
    # GoTo useApplication              ' found the TemplateFile Workbook
    # GoTo Retry

    if Err.Number <> 0 Then:
    # broken:
    # XlOpenedHere = True
    # GoTo Retry

    # UseWorkBook:
    if xlWBInit Is Nothing Or TypeName(xlWBInit) = "Object" Then:
    # Set xlWBInit = Nothing
    # GoTo useApplication

    # useApplication:

    try:
        if xlA Is Nothing Or TypeName(xlA) = "Object" Then:
        # Call xlOlClear
        # xlApp.Cursor = xlWait
        if ActiveWorkbook Is Nothing Then:
        # GoTo openWB
        elif ActiveWorkbook.FullName = TemplateFile Then:
        if Not ShutUpMode Then:
        print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
        # & "workbook file already open:", TemplateFile
        else:
        # openWB:
        # xlApp.EnableEvents = True            ' do the xl-side inits
        # xlApp.DisplayAlerts = False
        # aBugTxt = "open file: " & TemplateFile
        # Call Try
        # xlApp.Workbooks.Open FileName:=TemplateFile
        if Catch Then:
        # Call xlOlClear
        # Set xlApp = Nothing
        # GoTo ProcReturn
        # xlApp.DisplayAlerts = True
        # xlApp.EnableEvents = False

        # xlApp.Calculation = xlCalculationAutomatic
        # xlApp.CalculateBeforeSave = False
        # Set xlA = xlApp.ActiveWorkbook
        if xlA Is Nothing Then                   ' create an empty Workbook:
        # Set xlA = Workbooks.Add
        # xlA.Activate

        if xlAS Is Nothing Then:
        # Call MaintainSheetList(WBook, xlAS)
        elif xlAS.Count = 0 Then:
        # Call MaintainSheetList(WBook, xlAS)
        # ' publish .xlWBInit xlTab, matching worksheet
        # Call CreateOrUse(WBook, Sheetname, xlWBInit)

        # xlWBInit.xlTSheet.Activate
        if mustClear Then:
        # Call ClearSheetLines(xlWBInit)
        # xlWBInit.xlTSheet.Cells.VerticalAlignment = xlTop
        # xlWBInit.xlTSheet.EnableCalculation = enableSheetCalculation
        # xlWBInit.xlTSheet.EnableFormatConditionsCalculation = enableSheetCalculation

        if xlWBInit.xlTHeadline <> headline Or xlWBInit.xlTabIsEmpty < 1 Then:
        # xlWBInit.xHdl = headline                 ' this is a Property Let

        if showWorkbook Or DebugLogging Then:
        # Call DisplayExcel(xlWBInit, EnableEvents:=True, unconditionallyShow:=showWorkbook)
        # xlApp.Cursor = xlDefault

        if mustClear Then:
        print(Debug.Print Format(Timer, "0#####.00") & vbTab & "workbook ErrHdlInited, sheet", Sheetname)
        else:
        print(Debug.Print Format(Timer, "0#####.00") & vbTab & "workbook opened (with content), sheet", Sheetname)

        # ProcReturn:
        # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub MaintainSheetList
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def maintainsheetlist():

    # Const zKey As String = "cXlObject.MaintainSheetList"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim i As Long
    # Dim tW As cXLTab

    if SheetList Is Nothing Then:
    # Set SheetList = New Dictionary
    elif SheetList.Count <> WBook.Sheets.Count Then:
    # Set SheetList = New Dictionary
    elif SheetList.Count > 0 _:
    # And SheetList.Count - 1 = WBook.Sheets.Count Then
    if SheetList.Keys(1) = WBook.Name Then:
    # GoTo FuncExit
    else:
    # Set SheetList = New Dictionary
    else:
    # Set SheetList = New Dictionary

    # SheetList.Add WBook.Name, WBook              ' items(0) is parent
    # Set tW = New cXLTab
    # Set tW.xlTWBook = WBook                  ' Parent WB
    # Set tW.xlTSheet = WBook.Sheets(i)
    # tW.xlTSheet.Name = tW.xlTSheet.Name
    # tW.xlTName = tW.xlTSheet.Name
    # SheetList.Add tW.xlTSheet.Name, tW       ' sheets can never have duplicate names

    # FuncExit:
    # Set tW = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Function WorksheetDescriptorFind
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def worksheetdescriptorfind():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.WorksheetDescriptorFind"
    # Const zKey As String = "XLAccessingOL.WorksheetDescriptorFind"

    # Call DoCall(zKey, "Function", eQzMode)

    if D_S.Exists(Sheetname) Then:
    # Set WorksheetDescriptorFind = D_S.Item(Sheetname)


    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub XlgetApp
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlgetapp():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.XlgetApp"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if xlApp Is Nothing Then:
    # Set xlApp = Z_GetApplication("Excel", XlOpenedHere) ' start or use Excel
    # Call ErrReset(4)
    if CheckExcelOK() Then:
    if DebugMode Then:
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab & "Excel running, ";)
    if xlA Is Nothing Then:
    print(Debug.Print "No administered Workbook set yet")
    else:
    if TypeName(xlA) = "Object" Then:
    if xlApp.Workbooks.Count > 0 Then:
    if Not ShutUpMode Then:
    print(Debug.Print "Some Workbooks are open")
    # GoTo ProcReturn
    # DoVerify False, "Workbook no longer available"
    else:
    # DoVerify ShutUpMode, "Name of open Workbook is " & xlA.Name

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub KillExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def killexcel():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.KillExcel"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if xlApp Is Nothing Then:
    # GoTo FuncExit
    elif TypeName(xlApp) = "Nothing" Then:
    # GoTo FuncExit

    if XlOpenedHere Then:
    if DebugMode Then:
    else:
    # rsp = vbOK
    else:
    # & vbCrLf & vbCrLf & "Only Kill in case of severe Problem." _
    # & vbCrLf & vbCrLf & "About to Kill Excel Application", vbOK + vbCancel)
    if rsp <> vbCancel Then:
    if Not xlApp Is Nothing Then:
    # xlApp.Quit
    # Set xlApp = Nothing
    # Call xlEndApp                            ' set xlApp and associated values = Nothing

    # FuncExit:
    # XlOpenedHere = False

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Function XlForceTabOpen
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlforcetabopen():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.XlForceTabOpen"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if Not CheckExcelOK() Then:
    # Call XlgetApp                            ' if not, open it again.
    if xlA Is Nothing Or XlForceTabOpen Is Nothing Then ' this does not depend on xUseExcel !:
    # GoTo GetSheet
    elif XlForceTabOpen.xlTName <> Sheetname _:
    # Or (XlForceTabOpen.xlTHeadline <> headline And mustClear) Then
    # GetSheet:
    # Set XlForceTabOpen = xlWBInit(WBook, TemplateFile, Sheetname, _
    # headline, showWorkbook, mustClear)
    if Not xlApp Is Nothing Then:
    # xlApp.EnableEvents = False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub XlopenObjAttrSheet
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlopenobjattrsheet():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.XlopenObjAttrSheet"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Set O = XlForceTabOpen(WBook, TemplateFile, cOE_SheetName, _
    # sHdl, showWorkbook:=DebugMode, mustClear:=False)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Xlvisible
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def xlvisible():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.Xlvisible"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    if xlApp Is Nothing Then:
    print(Debug.Print "no Excel Object")
    # GoTo ProcReturn
    # Call XlgetApp
    if xlApp Is Nothing Then:
    print(Debug.Print "no Excel active")
    # GoTo ProcReturn
    # xlApp.Visible = Not off

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayWindowInFront
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def displaywindowinfront():
    # Dim zErr As cErr
    # Const zKey As String = "XlAccessingOl.DisplayWindowInFront"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim windowTitle As String
    if xHdl Is Nothing Then:
    # GoTo isNew
    if LenB(xHdl.strTitle) = 0 Then:
    # isNew:
    # windowTitle = setDefaultWindowName(wTitleIndex)
    else:
    # windowTitle = xHdl.strTitle
    # rsp = WindowSetForeground(windowTitle, xHdl)

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub addLine
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def addline():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.addLine"
    # Const zKey As String = "XLAccessingOL.addLine"

    # Call DoCall(zKey, "Sub", eQzMode)

    # Dim j As Long
    # Dim k As Long
    # Dim mustReplace As Boolean
    # Dim mCells As Range
    # Dim MyIndex As Long
    # Dim xwIsObjekteigenschaften As Boolean
    # Dim CellColor As Long
    # Dim nVar As Variant

    # xwIsObjekteigenschaften = (xw.xlTName = cOE_SheetName)
    # k = i + 1
    # MyIndex = 1
    # Call N_ClearAppErr
    # Set aCell = Nothing

    # xw.xlTSheet.Activate
    # ' Headline already present? Judge by LineNumber=0 and nothing in first cell
    if i > 0 And LenB(xw.xlTSheet.Cells(1, 1).Text) = 0 Then:
    # Call DisplayExcel(xw, EnableEvents:=False, unconditionallyShow:=True)
    # xw.xHdl = x.xlTHeadline

    # nVar = cvExcelVal(varr(j))
    # Set aCell = xw.xlTSheet.Cells(k, j)
    # aCell.Font.Color = 0

    if i <> 0 Then                           ' setting up headline only:
    if Not isEmpty(MostImportantProperties) Then:
    if xwIsObjekteigenschaften And j = 1 And k > 1 And k <= UBound(MostImportantProperties) + 1 Then ' that many mandatory props:
    # xw.xlTSheet.Cells(k, j).Borders.ColorIndex = 45 ' Borders brown for Mandatories in col 1
    if nVar <> Chr(0) Then:
    if Addit_Text And j = 1 Then:
    # aCell.EntireRow.Insert           ' this moves aCell, so we correct it:
    # Set aCell = xw.xlTSheet.Cells(k, j)
    # Addit_Text = False
    # aCell = nVar
    # aCell.HorizontalAlignment = xlLeft
    if i = 0 And InStr(varr(j), "-") > 0 Then:
    # mustReplace = True
    if xwIsObjekteigenschaften And j < 7 Then ' non-data columns:
    if InStr(varr(j), vbCrLf) = 0 Then:
    # aCell.WrapText = False
    else:
    # With aCell
    # .Activate
    # .HorizontalAlignment = xlGeneral
    # .VerticalAlignment = xlTop
    # .Orientation = 0
    # .MergeCells = False
    # .AddIndent = False
    # .IndentLevel = 0
    # .ShrinkToFit = False
    # .WrapText = True
    # .ShrinkToFit = True
    # End With                     ' aCell
    if xNoColAdjust Then:
    if j = 1 And i > MaxPropertyCount Then:
    # Set mCells = xw.xlTSheet.Range("A" & k & ":C" & k)
    # With mCells
    if DebugMode Then:
    # .Select
    # .HorizontalAlignment = xlLeft
    # .WrapText = True
    # .Orientation = 0
    # .AddIndent = False
    # .IndentLevel = 0
    # .ShrinkToFit = True
    # .ReadingOrder = xlContext
    # .MergeCells = False
    # .Rows.AutoFit
    # .Merge
    # End With             ' mCells
    else:
    # aCell.EntireColumn.AutoFit
    # aCell.WrapText = True
    # aCell.MergeCells = True
    # aCell.EntireColumn.AutoFit
    if i > 1 And (j = 3 Or j = 8) Then:
    if MyIndex <> 2 Then:
    # ' Debug.Assert False    ' eigentlich sollte MyIndex parameter sein ???
    # MyIndex = 2
    else:
    # aCell.WrapText = False
    elif DebugMode Then:
    # aCell.Select                         ' not changed, just Show this cell
    # varr(j) = Chr(0)

    # insertline:
    if i = 0 Then                                ' I=0: special operations on Headline:
    # xNoColAdjust = False
    # xw.xlTSheet.Rows(1).WrapText = False
    # xw.xlTSheet.Rows(1).columns.AutoFit
    if mustReplace Then:
    # xlApp.DisplayAlerts = False          ' suppress Excel message
    # xw.xlTSheet.Rows(1).Replace what:="-", Replacement:=b, LookAt:=xlPart, _
    # SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    # xlApp.DisplayAlerts = True
    # xw.xlTSheet.Activate
    # With xlApp.ActiveWindow
    # .SplitColumn = 0
    # .SplitRow = 1
    if Not .FreezePanes Then:
    # aBugTxt = "Freeze top line (accepted)"
    # Call Try
    # .FreezePanes = True
    # Call ErrReset(4)
    # End With                                 ' xlApp.ActiveWindow
    # xw.xlTSheet.Rows(1).Borders(xlDiagonalDown).LineStyle = xlNone
    # xw.xlTSheet.Rows(1).Borders(xlDiagonalUp).LineStyle = xlNone
    # xw.xlTSheet.Rows(1).Borders(xlEdgeLeft).LineStyle = xlNone
    # xw.xlTSheet.Rows(1).Borders(xlEdgeTop).LineStyle = xlNone
    # With xw.xlTSheet.Rows(1).Borders(xlEdgeBottom)
    # .LineStyle = xlContinuous
    # .ColorIndex = 0
    # .TintAndShade = 0
    # .Weight = xlThin
    # End With                                 ' T.Rows(1).Borders(xlEdgeBottom)
    # xNoColAdjust = True
    elif MyIndex = 2 And xwIsObjekteigenschaften Then:
    # CellColor = 35                           ' Grn
    if Left(xw.xlTSheet.Cells(k, 7), 2) = "# " Then:
    # CellColor = 0                        ' keine Farbe, Wei
    elif xw.xlTSheet.Cells(k, 7) <> xw.xlTSheet.Cells(k, 8) Then:
    # CellColor = 45                       ' Orange
    if BestRule(True).clsNeverCompare.RuleMatches Then:
    # CellColor = CellColor + 9
    elif xw.xlTSheet.Cells(k, 2) <> xw.xlTSheet.Cells(k, 3) Then:
    # CellColor = 6                        ' Gelb
    if BestRule(True).clsNeverCompare.RuleMatches Then:
    # CellColor = CellColor + 9
    if CellColor = 0 Then:
    # xw.xlTSheet.Cells(k, ChangeCol3).Interior.ColorIndex = -4142 ' wei
    else:
    # xw.xlTSheet.Cells(k, ChangeCol3).Interior.pattern = xlSolid
    # xw.xlTSheet.Cells(k, ChangeCol3).Interior.PatternColorIndex = xlAutomatic
    # xw.xlTSheet.Cells(k, ChangeCol3).Interior.ColorIndex = CellColor
    # xw.xlTabIsEmpty = k                          ' definitely not empty, >=1,
    # ' this is NOT=MaxPropertycount, but line number in Excel (offby 1)
    # Set aCell = Nothing
    # xw.xlTSheet.Cells(k, 1).Select


    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub AttrDscs2Excel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def attrdscs2excel():
    # Const zKey As String = "XLAccessingOL.AttrDscs2Excel"
    # Dim zErr As cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="XLAccessingOL")

    # Dim i As Long
    # Dim xOD As cObjDsc
    # Dim xC As Long

    if Not (xUseExcel Or xDeferExcel Or displayInExcel) Then:
    # GoTo ProcReturn
    # cHdl = vbNullString                                    ' put1intoExcel must match this:
    # cHdl = cHdl & b & "AIndex"                   ' 1
    # cHdl = cHdl & b & "ADName---------"          ' 2
    # cHdl = cHdl & b & "TIndex"                   ' 3
    # cHdl = cHdl & b & "px"                       ' 4
    # cHdl = cHdl & b & "MandAttr"                 ' 5
    # cHdl = cHdl & b & "DontComp"                 ' 6
    # cHdl = cHdl & b & "CantDec-"                 ' 7
    # cHdl = cHdl & b & "Similar-"                 ' 8
    # cHdl = cHdl & b & "Specific"                 ' 9
    # cHdl = cHdl & b & "PMandatory-----"          ' 10
    # cHdl = cHdl & b & "PDontCompare---"          ' 11
    # cHdl = cHdl & b & "PDecodable----"           ' 12
    # cHdl = cHdl & b & "PSimilar-------"          ' 13
    # cHdl = cHdl & b & "PIgnoreProp----"          ' 14
    # cHdl = cHdl & b & "PReserve-------"          ' 15

    # cHdl = Trim(cHdl)

    # Set xOD = D_TC.Items(xC)

    # xlApp.ScreenUpdating = DebugMode
    # xlApp.Cursor = xlWait
    if Not aID(0) Is Nothing Then:
    # aOD(0).objDumpMade = 0               ' no dumps made so far
    # Set W = xlWBInit(xlA, TemplateFile, xOD.objTypeName, _
    # cHdl, showWorkbook:=DebugMode, mustClear:=False)
    # ' NOTE: cleared completely and Headline is set to cHdl

    if xOD.objDumpMade < 1 Then:
    # GoTo noMatch                         ' never did Excel Output for ItemType

    # ' test if we have same content as before
    # i = W.xlTSheet.Cells(1, 16)
    if i - 3 <> aID(aPindex).idAttrDict.Count Then GoTo noMatch:
    if W.xlTSheet.Cells(i, 4) <> TrueCritList Then GoTo noMatch:
    if W.xlTSheet.Cells(i, 10) <> Trim(sRules.clsObligMatches.aRuleString) Then GoTo noMatch:
    if W.xlTSheet.Cells(i, 11) <> Trim(sRules.clsNeverCompare.aRuleString) Then GoTo noMatch:
    if W.xlTSheet.Cells(i, 12) <> Trim(sRules.clsNotDecodable.aRuleString) Then GoTo noMatch:
    if W.xlTSheet.Cells(i, 13) <> Trim(sRules.clsSimilarities.aRuleString) Then GoTo noMatch:
    # xOD.objDumpMade = 1                      ' dump was done, we assume headline is ok
    # W.xlTHeadline = cHdl
    # W.xlTHead = split(cHdl)

    # GoTo ProcReturn                          ' it's all in there already

    # noMatch:
    # W.xlTSheet.EnableCalculation = False     ' for attribute rules do not calculate
    # Call ClearSheetLines(W, 2)
    # With aID(aPindex)
    if TypeName(.idAttrDict.Items(i)) <> "cAttrDsc" Then:
    # pArr(1) = vbNullString
    # pArr(2) = "'" & "==========="
    # pArr(3) = i
    # pArr(15) = "Seperator"
    # Call addLine(W, i, pArr)
    else:
    # Set aTD = .idAttrDict.Items(i) ' for numeric index pos, must use Items instead of Item
    # Call put1IntoExcel(W, aTD, i)
    if i Mod 10 = 0 Then:
    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
    # & "inserted cAttrDsc # " & i _
    # & " into Excel Sheet " & W.xlTName
    # End With                                 ' aID(aPindex)

    if Not ShutUpMode Then:
    print(Debug.Print Format(Timer, "0#####.00") & vbTab _)
    # & "last cAttrDsc = " & i - 1 _
    # & " in Excel Sheet " & W.xlTName
    # i = i + 1
    # pArr(1) = i + 1
    # pArr(2) = "Matchlisten (aktuelle Grundlage der Attribut Deskriptoren)"
    # pArr(3) = b                              'Empty cell but force line wrap
    if sRules Is Nothing Then:
    # pArr(15) = "Invalid Class Rule"
    else:
    # pArr(4) = TrueCritList
    # pArr(10) = Trim(sRules.clsObligMatches.aRuleString)
    # pArr(11) = Trim(sRules.clsNeverCompare.aRuleString)
    # pArr(12) = Trim(sRules.clsNotDecodable.aRuleString)
    # pArr(13) = Trim(sRules.clsSimilarities.aRuleString)
    # ' pArr(14) = Trim(sRules.clsIgnoreProperties.aRuleString)
    # Call addLine(W, i, pArr)
    # i = i + 1
    # W.xlTSheet.Cells(i, 1).ShrinkToFit = False
    # W.xlTSheet.Cells(i, 1).EntireRow.VerticalAlignment = xlTop
    # W.xlTSheet.Cells(i, 1).EntireRow.WrapText = True
    # '    R4   Zi    R9   Zi
    # W.xlTSheet.Range("D" & i & ":I" & i).Merge
    # W.xlTSheet.Cells(1, 16) = i
    # W.xlTSheet.Cells(2, 1).Select
    # xOD.objDumpMade = 1
    # W.xlTabIsEmpty = i

    # ' Note: leaving W again (if we can)
    # xlApp.ScreenUpdating = True
    # xlApp.Cursor = xlDefault
    if Not O Is Nothing Then:
    # Set x = O
    if Not x Is x Then:
    if x Is Nothing Then:
    # Set W.xlTSheet = Nothing
    else:
    # Set W.xlTSheet = x.xlTSheet

    # FuncExit:
    # Set xOD = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowItemProps
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showitemprops():

    # Dim Line As String
    # Dim line2 As String
    # Dim what1 As String
    # Dim what2 As String
    # Dim a1 As cAttrDsc
    # Dim a2 As cAttrDsc

    # Dim i As Long
    # Dim iTop As Long
    # Dim Px2 As Long

    # 'On Error GoTo 0
    # Px2 = thisPx
    # what1 = what
    # what2 = what
    if thisPx = 2 Then:
    # DoVerify aID(2).idAttrCount < aID(2).idAttrDict.Count - 1, _
    # "The ItemPropertiesArray incomplete for index " & Px2
    # iTop = aID(2).idAttrCount
    elif thisPx = 3 Then:
    if aID(2).idAttrCount < aID(2).idAttrDict.Count - 1 Then:
    # Px2 = 1
    # iTop = aID(2).idAttrCount
    # Line = "The ItemPropertiesArray incomplete for index 2"
    if aID(1).idAttrCount < aID(1).idAttrDict.Count - 1 Then:
    if Px2 = 1 Then:
    # Px2 = 0
    # Line = "The ItemPropertiesArray is incomplete for both index 1 and 2"
    else:
    # Px2 = 2
    # Line = "The ItemPropertiesArray is incomplete for index 1"
    else:
    if aID(1).idAttrCount < aID(1).idAttrDict.Count - 1 Then:
    # Line = "The ItemPropertiesArray incomplete for index 1"
    # iTop = aID(1).idAttrCount
    else:
    # Px2 = 1
    # iTop = aID(1).idAttrDict.Count - 1
    if LenB(Line) > 0 Then:
    print(Debug.Print Line)

    if Px2 = 0 Then:
    # DoVerify False, Line
    # GoTo FuncExit

    if Px2 = 1 Then:
    # Call getWhats(aID(1).GetAttrDsc4Prop(i), Px2, Line, what1, what2)
    print(Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _)
    # & LString(what1, 6) & what2
    elif Px2 = 2 Then:
    # Call getWhats(aID(2).GetAttrDsc4Prop(i), Px2, Line, what1, what2)
    print(Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _)
    # & LString(what1, 6) & what2
    else:
    # Set a1 = aID(1).GetAttrDsc4Prop(i)
    # Set a2 = aID(2).GetAttrDsc4Prop(i)
    # Call getWhats(a1.adItemProp, Px2, Line, what1, vbNullString)
    # Call getWhats(a2.adItemProp, Px2, line2, what2, vbNullString)
    print(Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _)
    # & LString(what1, 6) & LString(line2, 20) & what2
    # FuncExit:
    # Set a1 = Nothing
    # Set a2 = Nothing


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowObjDesc
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showobjdesc():

    # Dim i As Long
    # Dim aDi As cAttrDsc

    # Set aDi = New cAttrDsc
    # ' aDi.iPindex = aPindex ???
    # i = 0
    for adi in aditmdsc:
    print(Debug.Print i, "dKey" & aPindex, "key=" _)
    # & aDi.adName _
    # & " has key=" & aDi.adKey;
    # i = i + 1
    if aDi.adItemProp Is Nothing Then:
    print(Debug.Print ", but has no ItemProperty yet")
    else:
    print(Debug.Print ", Property resolved is named " & aDi.adItemProp.Name)


# '---------------------------------------------------------------------------------------
# ' Method : Sub getWhats
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getwhats():

    # Const zKey As String = "XLAccessingOL.getWhats"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim Value As String

    # Line = adItemProp.Name
    # Value = "*** Error ***"
    # Call Try
    match what:
        case 0:

        case 1:
    # what1 = "W.xlTSheet: " & adItemProp.Type
    # Value = vbNullString
        case 2:
    # Value = "V: " & CStr(adItemProp.Value)
        case _:
    # what1 = "W.xlTSheet: " & adItemProp.Type
    # Value = "V: " & CStr(adItemProp.Value)
    # what2 = LString(Value, 19) & b

    # FuncExit:
    # Call ErrReset(0)

    # zExit:
    # Call DoExit(zKey, Value)


# '---------------------------------------------------------------------------------------
# ' Method : Sub put1IntoExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def put1intoexcel():
    # '--- Proc MAY ONLY CALL Z_Type PROCS
    # Const zKey As String = "XLAccessingOL.put1IntoExcel"

    # Call DoCall(zKey, "Sub", eQzMode)

    # With AP
    # pArr(1) = .adDictIndex
    # pArr(2) = .adKey
    # pArr(3) = .adtrueIndex
    if AP.adRules Is Nothing Then:
    # pArr(15) = "Seperator"
    else:
    # pArr(4) = " --"                      ' ??? war AP.adRules.ruleattrdsc.addictindex
    # pArr(5) = AP.adRules.clsObligMatches.RuleMatches
    # pArr(6) = AP.adRules.clsNeverCompare.RuleMatches
    # pArr(7) = AP.adRules.clsNotDecodable.RuleMatches
    # pArr(8) = AP.adRules.clsSimilarities.RuleMatches
    # pArr(9) = AP.adRules.RuleIsSpecific
    # pArr(10) = AP.adRules.clsObligMatches.MatchOn
    # pArr(11) = AP.adRules.clsNeverCompare.MatchOn
    # pArr(12) = AP.adRules.clsNotDecodable.MatchOn
    # pArr(13) = AP.adRules.clsSimilarities.MatchOn
    # ' pArr(14) = AP.adRules.clsIgnoreProperties.MatchOn
    # pArr(15) = AP.adRules.RuleType
    # End With                                     ' AP

    # Call addLine(xT, AttributeIndex, pArr)

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub put2IntoExcel
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def put2intoexcel():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.put2IntoExcel"
    # Const zKey As String = "XLAccessingOL.put2IntoExcel"

    # Call DoCall(zKey, "Sub", eQzMode)

    if xlApp Is Nothing Then:
    # GoTo FuncExit                            ' no Excel (yet)
    elif O Is Nothing Then:
    # GoTo FuncExit                            ' no Excel sheet (yet)
    if px = 1 And AttributeIndex = 1 Then:
    # Call ClearSheetLines(O, 2)
    # pArr(1) = aTD.adName                         ' = PropertyNameX, col1
    # ' cols 2 and 3
    # pArr(1 + px) = aTD.adShowValue               ' true value
    # ' parr(4) = compare indication
    if px = 2 Then:
    if Left(aTD.adDecodedValue, 1) = "#" Then:
    # pArr(6) = True                       ' visibility Flag, if empty: Show it
    # pArr(5) = "***"                      ' probably no value accessible or exists
    # pArr(4) = "___"
    elif aTD.adShowValue _:
    # <> aTD.adDecodedValue Then
    if Left(aTD.adShowValue, 2) = "# " Then:
    # pArr(6) = False                  ' visible because value ignored
    # pArr(4) = "..."
    # ' cols 7 and 8
    if LenB(CStr(aTD.adDecodedValue)) > 0 Then   ' non-nomalized phone number:
    # pArr(6 + px) = CStr(aTD.adDecodedValue)
    # ' cols 9 and 10
    # pArr(8 + px) = aTD.adKillMsg
    # Call addLine(O, Line, pArr)

    # FuncExit:

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:


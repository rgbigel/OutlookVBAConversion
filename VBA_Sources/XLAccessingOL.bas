Attribute VB_Name = "XLAccessingOL"
Option Explicit

Public xlApp As Object                           ' truly it should be Object/Excel.Application
Public xlA As Excel.WorkBook                     ' Selected workbook (=file)
Public xlAS As Dictionary                        ' Korrespondiert zu xlA.Sheets: Worksheets
Public xlC As Excel.WorkBook                     ' Possible second Workbook instance
Public xlCS As Dictionary                        ' Korrespondiert zu xlC.Sheets: Worksheets
' quick access sheets
Public E As cXLTab                               ' Editing RuleTableRule sheet
Public O As cXLTab                               ' selektiertes sheet "Objekteigenschaften"
Public S As cXLTab                               ' Table showing allSchemata rules
Public W As cXLTab                               ' Which sheet=<ItemTypeName>
Public x As cXLTab                               ' Aktuelle cXlTab gem‰ﬂ SheetName

Public xLMainWindowHdl As cFindWindowParms       ' Handle des Excel Hauptfensters

'---------------------------------------------------------------------------------------
' Method : xlOlClear
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Clear associated variables for xlApp
'---------------------------------------------------------------------------------------
Sub xlOlClear()
'''' Proc Must ONLY CALL Z_-, Y_- or Y_Type PROCS                         ' trivial proc
Const zKey As String = "XLAccessingOL.xlOlClear"

    Call DoCall(zKey, "Sub", eQzMode)

    Set xlA = Nothing                            ' Selected workbook (=file)
    Set xlAS = Nothing                           ' Korrespondiert zu xlA.Sheets: Worksheets
    Set xlC = Nothing                            ' Possible second Workbook instance
    Set xlCS = Nothing                           ' Korrespondiert zu xlC.Sheets: Worksheets
    ' quick access sheets
    Set E = Nothing                              ' Editing RuleTableRule sheet
    Set O = Nothing                              ' selektiertes sheet "Objekteigenschaften"
    Set S = Nothing                              ' Table showing allSchemata rules
    Set W = Nothing                              ' Which sheet=<ItemTypeName>
    Set x = Nothing                              ' Aktuelle cXlTab gem‰ﬂ SheetName

    Set xLMainWindowHdl = Nothing                ' Handle des Excel Hauptfensters

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                          ' XLAccessingOL.xlOlClear

'---------------------------------------------------------------------------------------
' Method : Function CheckExcelOK
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function CheckExcelOK() As Boolean
Dim zErr As cErr
Const zKey As String = "cXlObject.CheckExcelOK"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cXlObject")

    If xlApp Is Nothing Then
        Call KillExcel
    ElseIf TypeName(xlApp) <> "Application" Then
        Call KillExcel
    ElseIf Not xlA Is Nothing Then
        If xlA.ActiveSheet Is Nothing Then
            Call KillExcel
        End If
    Else
        CheckExcelOK = True
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' cXlObject.CheckExcelOK

'---------------------------------------------------------------------------------------
' Method : Sub ClearSheetLines
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClearSheetLines(xw As cXLTab, Optional FromLine As Long = 2)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cXlObject.ClearSheetLines"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

Dim oldcontent As Range
Dim eVinit As Boolean
Dim ws As Worksheet
Dim ns As Worksheet
Dim RS As Worksheet
Dim isVis As Boolean
Dim ShNam As String

    Set ws = xw.xlTSheet
    ws.Activate
    xw.xlTLastLine = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    If FromLine = 0 Then
        xw.xlTLastLine = 1
        xw.xlTLastCol = 0
        GoTo eraseIt
    End If
    xw.xlTLastLine = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    xw.xlTabIsEmpty = xw.xlTLastLine             ' defined 3-state variable: is empty if <= 1
    If xw.xlTLastLine <= 1 Then
        xw.xlTabIsEmpty = 1
        GoTo ProcReturn
    End If
    
    isVis = xlApp.Visible
    eVinit = xlApp.EnableEvents
    xlApp.EnableEvents = True
    ws.Cells(2, 1).Value = "ClearMe"
    ws.Cells(2, 1).Select                        ' this triggers eventroutine in Excel which clears WS
    xlApp.EnableEvents = False
    xlApp.AutoRecover.Enabled = False
    
    ws.Activate
    Set oldcontent = ws.UsedRange
    xw.xlTLastLine = oldcontent.Rows.Count + oldcontent.Row - 1
    xw.xlTLastCol = oldcontent.columns.Count + oldcontent.Column - 1
    
    If xw.xlTLastLine > 1 Then                   ' not empty except headline
        ShNam = ws.Name
        If ShNam = "Objekteigenschaften" Then
            If Not ShutUpMode Then
                Debug.Print "Col, Row of last cell = ", xw.xlTLastCol, xw.xlTLastLine, "Restoring Sheet"
            End If
            Set RS = ws.Parent.Sheets("OEdefault")
            RS.Copy Before:=ws
            Set ns = ws.Parent.Sheets("OEdefault (2)")
        Else
            If Not ShutUpMode Then
                Debug.Print "Col, Row of last cell = ", xw.xlTLastCol, xw.xlTLastLine, "Clearing Sheet"
            End If
            Set ns = ws.Parent.Sheets.Add        ' no code in this sheet
        End If
eraseIt:
        ws.Activate
        xlApp.DisplayAlerts = False
        ws.Delete
        xlApp.DisplayAlerts = True
        ns.Name = ShNam
        ns.Visible = xlSheetVisible
        Set xw.xlTSheet = ns
        ns.Activate
    End If
    ' housekeeping
    If Not (DebugMode Or DebugLogging) Then
        xlApp.Visible = isVis
    End If
    xw.xlTabIsEmpty = 1
    xlApp.EnableEvents = eVinit

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' cXlObject.ClearSheetLines

'---------------------------------------------------------------------------------------
' Method : Sub ClearWorkSheet
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ClearWorkSheet(ByRef WBook As WorkBook, ByRef ws As cXLTab)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cXlObject.ClearWorkSheet"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    DoVerify WBook Is ws.xlTWBook, "Design Check on correct Worksheet parent ???"
    If xlApp Is Nothing Then
        Set WBook = Nothing
        Set ws = Nothing
        GoTo ProcReturn
    ElseIf WBook Is Nothing Then
        Set ws = Nothing
        GoTo ProcReturn
    ElseIf ws Is Nothing Then
        Call CreateOrUse(WBook, WBook.Sheets(1).Name, ws)
        If DebugMode Then
            DoVerify False, "about to clear sheet " & ws.xlTSheet.Name & " are you shure???"
        End If
        GoTo usit
    Else
usit:
        Call ClearSheetLines(ws, 0)
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' cXlObject.ClearWorkSheet

'---------------------------------------------------------------------------------------
' Method : Sub CreateOrUse
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Uses XLB workbook, Set cu.xlTSheet, makes or uses Sheet(Sheetname), Set X==cu
'          Only valid for xlAS matching WBook
'---------------------------------------------------------------------------------------
Sub CreateOrUse(ByRef WBook As WorkBook, Sheetname As String, ByRef cu As cXLTab)
Dim zErr As cErr
Const zKey As String = "cXlObject.CreateOrUse"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cXlObject")

    If cu Is Nothing Then                        ' we must find it
        Set cu = WorksheetDescriptorFind(xlAS, Sheetname)
    Else
        If cu.xlTSheet.Name <> Sheetname Then    ' find only if not the right one
            Set cu = WorksheetDescriptorFind(xlAS, Sheetname)
        End If
    End If
    If cu Is Nothing Then                        ' not found: make a new Sheet by that name
        Set cu = New cXLTab
        Set cu.xlTWBook = WBook
        Set cu.xlTSheet = xlA.Sheets.Add(After:=xlA.Sheets.Item(xlA.Sheets.Count))
        cu.xlTSheet.Name = Sheetname
        xlAS.Add Sheetname, cu.xlTSheet          ' add to Sheet list xlAS
    End If
    Set x = cu
    cu.xlTSheet.Select
    
    cu.xlTSheet.EnableCalculation = False        ' default for our sheets
    cu.xlTSheet.EnableFormatConditionsCalculation = False

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' cXlObject.CreateOrUse

'---------------------------------------------------------------------------------------
' Method : Sub DisplayExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DisplayExcel(xlW As cXLTab, Optional relevant_only As Boolean, Optional EnableEvents As Boolean = True, Optional unconditionallyShow As Boolean, Optional xlY As Excel.Worksheet)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cXlObject.DisplayExcel"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    If unconditionallyShow Or displayInExcel Then
        If Not xlW Is Nothing Then
            If IsMissing(xlY) Or xlY Is Nothing Then
                Set xlY = xlW.xlTSheet           ' default worksheet of xlW=clsXlTab
            End If
        End If
        If Not xlApp Is Nothing Then
            If xlY Is Nothing Then               ' ActiveWorkbook
                DoVerify False
            Else
                xlApp.Visible = True
                xlY.Cells(1, 20).Value = relevant_only
                xlY.EnableCalculation = True
                xlY.EnableFormatConditionsCalculation = True
                xlApp.EnableEvents = False
                xlY.Activate
                If xlApp.Visible Then
                    xlApp.EnableEvents = EnableEvents
                Else
                    xlY.Activate
                    xlApp.EnableEvents = EnableEvents
                End If
            End If
        End If
        If xlApp.EnableEvents Then
            If Not (DebugMode Or DebugLogging) Then
                xlApp.ScreenUpdating = False
            End If
            xlApp.Cursor = xlWait
        Else
            xlApp.ScreenUpdating = True
            xlApp.Cursor = xlDefault
        End If
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' cXlObject.DisplayExcel

'---------------------------------------------------------------------------------------
' Method : Sub EndAllWorkbooks
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub EndAllWorkbooks()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cXlObject.EndAllWorkbooks"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

Dim onlyMySheets As Boolean

    onlyMySheets = True
    If xlApp Is Nothing Then
        GoTo ProcReturn
    End If
    If xlApp Is Nothing Then
        GoTo ProcReturn
    End If
    xlApp.EnableEvents = False
    xlApp.AutoRecover.Enabled = True
    xlApp.Cursor = xlDefault
    If Workbooks.Count > 0 Then
        xlApp.CalculateBeforeSave = True
        xlApp.Calculation = xlCalculationAutomatic
    End If
    For Each xlA In xlApp.Workbooks
        aBugTxt = "close Workbook " & Quote(xlA.Name)
        If InStr(UCase(xlA.Name), "PERSONAL") > 0 Then
        ElseIf InStr(UCase(xlA.Name), "OUTLOOKDEFAULT") > 0 Then
            aBugTxt = "Close WorkBook " & xlA.Name
            Call Try                             ' Try anything, autocatch
            xlA.Close False                      ' error here never matter
        ElseIf Left(xlA.Name, 5) = "Mappe" Then
            Call Try                             ' Try anything, autocatch
            xlA.Close False                      ' error here never matter
        Else
            onlyMySheets = False
            XlOpenedHere = False
        End If
        Catch
    Next xlA
    
    Call xlOlClear

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' cXlObject.EndAllWorkbooks

'---------------------------------------------------------------------------------------
' Method : Sub xlEndApp
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub xlEndApp()
Const zKey As String = "XLAccessingOL.xlEndApp"
Static zErr As New cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="XLAccessingOL")
    
    Call EndAllWorkbooks
    If XlOpenedHere Or xlApp.Workbooks.Count = 0 Then
        aBugTxt = "Close/Quit Excel"
        Call Try                                 ' Try anything, autocatch
        xlApp.Quit
        If Not Catch Then
            If Not ShutUpMode Then
                Debug.Print "Excel has been closed successfully"
            End If
        End If
        Call xlOlClear
        Set xlApp = Nothing
    Else
        Call LogEvent("Excel not closed because (we know) other workbooks are open")
    End If
    Call ErrReset(0)
    
    XlOpenedHere = False

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' cXlObject.xlEndApp

'---------------------------------------------------------------------------------------
' Method : Function xlWBInit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function xlWBInit(ByRef WBook As WorkBook, TemplateFile As String, Sheetname As String, headline As String, Optional showWorkbook As Boolean, Optional mustClear As Boolean, Optional enableSheetCalculation As Boolean = False) As cXLTab
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cXlObject.xlWBInit"
Dim zErr As cErr

Dim RetryCounter As Long
    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cXlObject")

    Call ErrReset(4)
    If xlApp Is Nothing Then
        Call xlOlClear
ExcelDied:
        Call ErrReset(4)
        Call XlgetApp
        If xlApp Is Nothing Then
            Call LogEvent("unable to obtain Excel Application Object, Retries=" & RetryCounter)
            GoTo Retry
        End If
    ElseIf TypeName(xlApp) <> "Application" Then
Retry:
        Call ErrReset(4)
        RetryCounter = RetryCounter + 1
        Call xlEndApp
        If RetryCounter > 4 Then
            Call LogEvent("Unable to use Excel with " & TemplateFile)
        End If
        GoTo ExcelDied
    End If
    
    If xlA Is Nothing Or TypeName(xlA) = "Object" Then
        If xlApp.Workbooks.Count = 0 Then
            Call LogEvent("Excel open, but no Workbooks open, need: " & TemplateFile)
            GoTo useApplication                  ' open one
        End If
        For Each xlA In xlApp.Workbooks
            If xlA.FullName = TemplateFile Then
                If DebugLogging Then
                    Call LogEvent("Found the correct Workbook for xlA: " & TemplateFile)
                End If
                GoTo useApplication              ' found the TemplateFile Workbook
            End If
        Next xlA
        GoTo Retry
    End If
    
    If Err.Number <> 0 Then
broken:
        XlOpenedHere = True
        GoTo Retry
    End If
        
UseWorkBook:
    If xlWBInit Is Nothing Or TypeName(xlWBInit) = "Object" Then
        Set xlWBInit = Nothing
        GoTo useApplication
    End If
    
useApplication:
    
    On Error GoTo broken
    If xlA Is Nothing Or TypeName(xlA) = "Object" Then
        Call xlOlClear
        xlApp.Cursor = xlWait
        If ActiveWorkbook Is Nothing Then
            GoTo openWB
        ElseIf ActiveWorkbook.FullName = TemplateFile Then
            If Not ShutUpMode Then
                Debug.Print Format(Timer, "0#####.00") & vbTab _
                                                     & "workbook file already open:", TemplateFile
            End If
        Else
openWB:
            xlApp.EnableEvents = True            ' do the xl-side inits
            xlApp.DisplayAlerts = False
            aBugTxt = "open file: " & TemplateFile
            Call Try
            xlApp.Workbooks.Open FileName:=TemplateFile
            If Catch Then
                Call xlOlClear
                Set xlApp = Nothing
                GoTo ProcReturn
            End If
            xlApp.DisplayAlerts = True
            xlApp.EnableEvents = False
        End If
        
        xlApp.Calculation = xlCalculationAutomatic
        xlApp.CalculateBeforeSave = False
        Set xlA = xlApp.ActiveWorkbook
        If xlA Is Nothing Then                   ' create an empty Workbook
            Set xlA = Workbooks.Add
            xlA.Activate
        End If
    End If
    
    If xlAS Is Nothing Then
        Call MaintainSheetList(WBook, xlAS)
    ElseIf xlAS.Count = 0 Then
        Call MaintainSheetList(WBook, xlAS)
    End If
    ' publish .xlWBInit xlTab, matching worksheet
    Call CreateOrUse(WBook, Sheetname, xlWBInit)
    
    xlWBInit.xlTSheet.Activate
    If mustClear Then
        Call ClearSheetLines(xlWBInit)
        xlWBInit.xlTSheet.Cells.VerticalAlignment = xlTop
    End If
    xlWBInit.xlTSheet.EnableCalculation = enableSheetCalculation
    xlWBInit.xlTSheet.EnableFormatConditionsCalculation = enableSheetCalculation
        
    If xlWBInit.xlTHeadline <> headline Or xlWBInit.xlTabIsEmpty < 1 Then
        xlWBInit.xHdl = headline                 ' this is a Property Let
    End If
    
    If showWorkbook Or DebugLogging Then
        Call DisplayExcel(xlWBInit, EnableEvents:=True, unconditionallyShow:=showWorkbook)
        xlApp.Cursor = xlDefault
    End If

    If mustClear Then
        Debug.Print Format(Timer, "0#####.00") & vbTab & "workbook ErrHdlInited, sheet", Sheetname
    Else
        Debug.Print Format(Timer, "0#####.00") & vbTab & "workbook opened (with content), sheet", Sheetname
    End If

ProcReturn:
    Call ProcExit(zErr)

End Function                                     ' cXlObject.xlWBInit

'---------------------------------------------------------------------------------------
' Method : Sub MaintainSheetList
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub MaintainSheetList(WBook As WorkBook, ByRef SheetList As Dictionary)

Const zKey As String = "cXlObject.MaintainSheetList"
    Call DoCall(zKey, tSub, eQzMode)

Dim i As Long
Dim tW As cXLTab

    If SheetList Is Nothing Then
        Set SheetList = New Dictionary
    ElseIf SheetList.Count <> WBook.Sheets.Count Then
        Set SheetList = New Dictionary
    ElseIf SheetList.Count > 0 _
        And SheetList.Count - 1 = WBook.Sheets.Count Then
        If SheetList.Keys(1) = WBook.Name Then
            GoTo FuncExit
        Else
            Set SheetList = New Dictionary
        End If
    Else
        Set SheetList = New Dictionary
    End If
    
    SheetList.Add WBook.Name, WBook              ' items(0) is parent
    For i = 1 To WBook.Sheets.Count              ' contains all current sheets
        Set tW = New cXLTab
        Set tW.xlTWBook = WBook                  ' Parent WB
        Set tW.xlTSheet = WBook.Sheets(i)
        tW.xlTSheet.Name = tW.xlTSheet.Name
        tW.xlTName = tW.xlTSheet.Name
        SheetList.Add tW.xlTSheet.Name, tW       ' sheets can never have duplicate names
    Next i

FuncExit:
    Set tW = Nothing
    
zExit:
    Call DoExit(zKey)

End Sub                                          ' cXlObject.MaintainSheetList

'---------------------------------------------------------------------------------------
' Method : Function WorksheetDescriptorFind
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function WorksheetDescriptorFind(D_S As Dictionary, Sheetname As String) As cXLTab
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.WorksheetDescriptorFind"
Const zKey As String = "XLAccessingOL.WorksheetDescriptorFind"

    Call DoCall(zKey, "Function", eQzMode)
      
    If D_S.Exists(Sheetname) Then
        Set WorksheetDescriptorFind = D_S.Item(Sheetname)
    End If


zExit:
    Call DoExit(zKey)
ProcRet:
End Function                                     ' cXlObject.WorksheetDescriptorFind

'---------------------------------------------------------------------------------------
' Method : Sub XlgetApp
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub XlgetApp()
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.XlgetApp"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If xlApp Is Nothing Then
        Set xlApp = Z_GetApplication("Excel", XlOpenedHere) ' start or use Excel
        Call ErrReset(4)
    End If
    If CheckExcelOK() Then
        If DebugMode Then
            If Not ShutUpMode Then
                Debug.Print Format(Timer, "0#####.00") & vbTab & "Excel running, ";
            End If
            If xlA Is Nothing Then
                Debug.Print "No administered Workbook set yet"
            Else
                If TypeName(xlA) = "Object" Then
                    If xlApp.Workbooks.Count > 0 Then
                        If Not ShutUpMode Then
                            Debug.Print "Some Workbooks are open"
                        End If
                        GoTo ProcReturn
                    End If
                    DoVerify False, "Workbook no longer available"
                Else
                    DoVerify ShutUpMode, "Name of open Workbook is " & xlA.Name
                End If
            End If
        End If
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XlAccessingOl.XlgetApp

'---------------------------------------------------------------------------------------
' Method : Sub KillExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub KillExcel()
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.KillExcel"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    
    If xlApp Is Nothing Then
        GoTo FuncExit
    ElseIf TypeName(xlApp) = "Nothing" Then
        GoTo FuncExit
    End If

    If XlOpenedHere Then
        If DebugMode Then
            rsp = MsgBox("About to Kill Excel Application", vbExclamation + vbCancel)
        Else
            rsp = vbOK
        End If
    Else
        rsp = MsgBox("Beware: at least one Excel Session was not opened here!" _
                   & vbCrLf & vbCrLf & "Only Kill in case of severe Problem." _
                   & vbCrLf & vbCrLf & "About to Kill Excel Application", vbOK + vbCancel)
    End If
    If rsp <> vbCancel Then
        If Not xlApp Is Nothing Then
            xlApp.Quit
        End If
        Set xlApp = Nothing
        Call xlEndApp                            ' set xlApp and associated values = Nothing
    End If

FuncExit:
    XlOpenedHere = False

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XlAccessingOl.KillExcel

'---------------------------------------------------------------------------------------
' Method : Function XlForceTabOpen
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function XlForceTabOpen(WBook As WorkBook, TemplateFile As String, Sheetname As String, headline As String, Optional showWorkbook As Boolean, Optional mustClear As Boolean = True) As cXLTab
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.XlForceTabOpen"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If Not CheckExcelOK() Then
        Call XlgetApp                            ' if not, open it again.
    End If
    If xlA Is Nothing Or XlForceTabOpen Is Nothing Then ' this does not depend on xUseExcel !
        GoTo GetSheet
    ElseIf XlForceTabOpen.xlTName <> Sheetname _
        Or (XlForceTabOpen.xlTHeadline <> headline And mustClear) Then
GetSheet:
        Set XlForceTabOpen = xlWBInit(WBook, TemplateFile, Sheetname, _
                                      headline, showWorkbook, mustClear)
    End If
    If Not xlApp Is Nothing Then
        xlApp.EnableEvents = False
    End If
    
FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' XlAccessingOl.XlForceTabOpen

'---------------------------------------------------------------------------------------
' Method : Sub XlopenObjAttrSheet
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub XlopenObjAttrSheet(WBook As WorkBook)
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.XlopenObjAttrSheet"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Set O = XlForceTabOpen(WBook, TemplateFile, cOE_SheetName, _
                           sHdl, showWorkbook:=DebugMode, mustClear:=False)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XlAccessingOl.XlopenObjAttrSheet

'---------------------------------------------------------------------------------------
' Method : Sub Xlvisible
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Xlvisible(Optional off As Boolean)
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.Xlvisible"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    If xlApp Is Nothing Then
        Debug.Print "no Excel Object"
        GoTo ProcReturn
    End If
    Call XlgetApp
    If xlApp Is Nothing Then
        Debug.Print "no Excel active"
        GoTo ProcReturn
    End If
    xlApp.Visible = Not off

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XlAccessingOl.Xlvisible

'---------------------------------------------------------------------------------------
' Method : Sub DisplayWindowInFront
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub DisplayWindowInFront(xHdl As cFindWindowParms, wTitleIndex As Long)
Dim zErr As cErr
Const zKey As String = "XlAccessingOl.DisplayWindowInFront"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim windowTitle As String
    If xHdl Is Nothing Then
        GoTo isNew
    End If
    If LenB(xHdl.strTitle) = 0 Then
isNew:
        windowTitle = setDefaultWindowName(wTitleIndex)
    Else
        windowTitle = xHdl.strTitle
    End If
    rsp = WindowSetForeground(windowTitle, xHdl)

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XlAccessingOl.DisplayWindowInFront

'---------------------------------------------------------------------------------------
' Method : Sub addLine
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub addLine(ByRef xw As cXLTab, i As Long, varr As Variant)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.addLine"
Const zKey As String = "XLAccessingOL.addLine"

    Call DoCall(zKey, "Sub", eQzMode)

Dim j As Long
Dim k As Long
Dim mustReplace As Boolean
Dim mCells As Range
Dim MyIndex As Long
Dim xwIsObjekteigenschaften As Boolean
Dim CellColor As Long
Dim nVar As Variant

    xwIsObjekteigenschaften = (xw.xlTName = cOE_SheetName)
    k = i + 1
    MyIndex = 1
    Call N_ClearAppErr
    Set aCell = Nothing
    
    xw.xlTSheet.Activate
    ' Headline already present? Judge by LineNumber=0 and nothing in first cell
    If i > 0 And LenB(xw.xlTSheet.Cells(1, 1).Text) = 0 Then
        Call DisplayExcel(xw, EnableEvents:=False, unconditionallyShow:=True)
        xw.xHdl = x.xlTHeadline
    End If
    
    For j = 1 To UBound(varr)                    ' not using lbound = 0 as after split
        nVar = cvExcelVal(varr(j))
        Set aCell = xw.xlTSheet.Cells(k, j)
        aCell.Font.Color = 0
        
        If i <> 0 Then                           ' setting up headline only
            If Not isEmpty(MostImportantProperties) Then
                If xwIsObjekteigenschaften And j = 1 And k > 1 And k <= UBound(MostImportantProperties) + 1 Then ' that many mandatory props
                    xw.xlTSheet.Cells(k, j).Borders.ColorIndex = 45 ' Borders brown for Mandatories in col 1
                End If
            End If
        End If
        If nVar <> Chr(0) Then
            If Addit_Text And j = 1 Then
                aCell.EntireRow.Insert           ' this moves aCell, so we correct it:
                Set aCell = xw.xlTSheet.Cells(k, j)
            End If
            Addit_Text = False
            aCell = nVar
            aCell.HorizontalAlignment = xlLeft
            If i = 0 And InStr(varr(j), "-") > 0 Then
                mustReplace = True
            End If
            If xwIsObjekteigenschaften And j < 7 Then ' non-data columns
                If InStr(varr(j), vbCrLf) = 0 Then
                    aCell.WrapText = False
                Else
                    With aCell
                        .Activate
                        .HorizontalAlignment = xlGeneral
                        .VerticalAlignment = xlTop
                        .Orientation = 0
                        .MergeCells = False
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .WrapText = True
                        .ShrinkToFit = True
                    End With                     ' aCell
                    If xNoColAdjust Then
                        If j = 1 And i > MaxPropertyCount Then
                            Set mCells = xw.xlTSheet.Range("A" & k & ":C" & k)
                            With mCells
                                If DebugMode Then
                                    .Select
                                End If
                                .HorizontalAlignment = xlLeft
                                .WrapText = True
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = True
                                .ReadingOrder = xlContext
                                .MergeCells = False
                                .Rows.AutoFit
                                .Merge
                            End With             ' mCells
                        End If
                    Else
                        aCell.EntireColumn.AutoFit
                        aCell.WrapText = True
                        aCell.MergeCells = True
                        aCell.EntireColumn.AutoFit
                    End If
                End If
                If i > 1 And (j = 3 Or j = 8) Then
                    If MyIndex <> 2 Then
                        ' Debug.Assert False    ' eigentlich sollte MyIndex parameter sein ???
                        MyIndex = 2
                    End If
                End If
            Else
                aCell.WrapText = False
            End If
        ElseIf DebugMode Then
            aCell.Select                         ' not changed, just Show this cell
        End If
        varr(j) = Chr(0)
    Next j
    
insertline:
    If i = 0 Then                                ' I=0: special operations on Headline
        xNoColAdjust = False
        xw.xlTSheet.Rows(1).WrapText = False
        xw.xlTSheet.Rows(1).columns.AutoFit
        If mustReplace Then
            xlApp.DisplayAlerts = False          ' suppress Excel message
            xw.xlTSheet.Rows(1).Replace what:="-", Replacement:=b, LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            xlApp.DisplayAlerts = True
        End If
        xw.xlTSheet.Activate
        With xlApp.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            If Not .FreezePanes Then
                aBugTxt = "Freeze top line (accepted)"
                Call Try
                .FreezePanes = True
                Call ErrReset(4)
            End If
        End With                                 ' xlApp.ActiveWindow
        xw.xlTSheet.Rows(1).Borders(xlDiagonalDown).LineStyle = xlNone
        xw.xlTSheet.Rows(1).Borders(xlDiagonalUp).LineStyle = xlNone
        xw.xlTSheet.Rows(1).Borders(xlEdgeLeft).LineStyle = xlNone
        xw.xlTSheet.Rows(1).Borders(xlEdgeTop).LineStyle = xlNone
        With xw.xlTSheet.Rows(1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With                                 ' T.Rows(1).Borders(xlEdgeBottom)
        xNoColAdjust = True
    ElseIf MyIndex = 2 And xwIsObjekteigenschaften Then
        CellColor = 35                           ' Gr¸n
        If Left(xw.xlTSheet.Cells(k, 7), 2) = "# " Then
            CellColor = 0                        ' keine Farbe, Weiﬂ
        ElseIf xw.xlTSheet.Cells(k, 7) <> xw.xlTSheet.Cells(k, 8) Then
            CellColor = 45                       ' Orange
            If BestRule(True).clsNeverCompare.RuleMatches Then
                CellColor = CellColor + 9
            End If
        ElseIf xw.xlTSheet.Cells(k, 2) <> xw.xlTSheet.Cells(k, 3) Then
            CellColor = 6                        ' Gelb
            If BestRule(True).clsNeverCompare.RuleMatches Then
                CellColor = CellColor + 9
            End If
        End If
        If CellColor = 0 Then
            xw.xlTSheet.Cells(k, ChangeCol3).Interior.ColorIndex = -4142 ' weiﬂ
        Else
            xw.xlTSheet.Cells(k, ChangeCol3).Interior.pattern = xlSolid
            xw.xlTSheet.Cells(k, ChangeCol3).Interior.PatternColorIndex = xlAutomatic
            xw.xlTSheet.Cells(k, ChangeCol3).Interior.ColorIndex = CellColor
        End If
    End If
    xw.xlTabIsEmpty = k                          ' definitely not empty, >=1,
    ' this is NOT=MaxPropertycount, but line number in Excel (offby 1)
    Set aCell = Nothing
    xw.xlTSheet.Cells(k, 1).Select
    
    
zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                          ' XlAccessingOl.addLine

'---------------------------------------------------------------------------------------
' Method : Sub AttrDscs2Excel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub AttrDscs2Excel()
Const zKey As String = "XLAccessingOL.AttrDscs2Excel"
Dim zErr As cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="XLAccessingOL")

Dim i As Long
Dim xOD As cObjDsc
Dim xC As Long

    If Not (xUseExcel Or xDeferExcel Or displayInExcel) Then
        GoTo ProcReturn
    End If
    cHdl = vbNullString                                    ' put1intoExcel must match this:
    cHdl = cHdl & b & "AIndex"                   ' 1
    cHdl = cHdl & b & "ADName---------"          ' 2
    cHdl = cHdl & b & "TIndex"                   ' 3
    cHdl = cHdl & b & "px"                       ' 4
    cHdl = cHdl & b & "MandAttr"                 ' 5
    cHdl = cHdl & b & "DontComp"                 ' 6
    cHdl = cHdl & b & "CantDec-"                 ' 7
    cHdl = cHdl & b & "Similar-"                 ' 8
    cHdl = cHdl & b & "Specific"                 ' 9
    cHdl = cHdl & b & "PMandatory-----"          ' 10
    cHdl = cHdl & b & "PDontCompare---"          ' 11
    cHdl = cHdl & b & "PDecodable----"           ' 12
    cHdl = cHdl & b & "PSimilar-------"          ' 13
    cHdl = cHdl & b & "PIgnoreProp----"          ' 14
    cHdl = cHdl & b & "PReserve-------"          ' 15
    
    cHdl = Trim(cHdl)
    
    For xC = 1 To D_TC.Count - 1
        Set xOD = D_TC.Items(xC)
    
        xlApp.ScreenUpdating = DebugMode
        xlApp.Cursor = xlWait
        If Not aID(0) Is Nothing Then
            aOD(0).objDumpMade = 0               ' no dumps made so far
        End If
        Set W = xlWBInit(xlA, TemplateFile, xOD.objTypeName, _
                         cHdl, showWorkbook:=DebugMode, mustClear:=False)
        ' NOTE: cleared completely and Headline is set to cHdl

        If xOD.objDumpMade < 1 Then
            GoTo noMatch                         ' never did Excel Output for ItemType
        End If
        
        ' test if we have same content as before
        i = W.xlTSheet.Cells(1, 16)
        If i - 3 <> aID(aPindex).idAttrDict.Count Then GoTo noMatch
        If W.xlTSheet.Cells(i, 4) <> TrueCritList Then GoTo noMatch
        If W.xlTSheet.Cells(i, 10) <> Trim(sRules.clsObligMatches.aRuleString) Then GoTo noMatch
        If W.xlTSheet.Cells(i, 11) <> Trim(sRules.clsNeverCompare.aRuleString) Then GoTo noMatch
        If W.xlTSheet.Cells(i, 12) <> Trim(sRules.clsNotDecodable.aRuleString) Then GoTo noMatch
        If W.xlTSheet.Cells(i, 13) <> Trim(sRules.clsSimilarities.aRuleString) Then GoTo noMatch
        xOD.objDumpMade = 1                      ' dump was done, we assume headline is ok
        W.xlTHeadline = cHdl
        W.xlTHead = split(cHdl)
        
        GoTo ProcReturn                          ' it's all in there already

noMatch:
        W.xlTSheet.EnableCalculation = False     ' for attribute rules do not calculate
        Call ClearSheetLines(W, 2)
        With aID(aPindex)
            For i = 1 To .idAttrDict.Count - 1   ' get the cAttrDsc items
                If TypeName(.idAttrDict.Items(i)) <> "cAttrDsc" Then
                    pArr(1) = vbNullString
                    pArr(2) = "'" & "==========="
                    pArr(3) = i
                    pArr(15) = "Seperator"
                    Call addLine(W, i, pArr)
                Else
                    Set aTD = .idAttrDict.Items(i) ' for numeric index pos, must use Items instead of Item
                    Call put1IntoExcel(W, aTD, i)
                End If
                If i Mod 10 = 0 Then
                    If Not ShutUpMode Then
                        Debug.Print Format(Timer, "0#####.00") & vbTab _
                                                             & "inserted cAttrDsc # " & i _
                                                             & " into Excel Sheet " & W.xlTName
                    End If
                End If
            Next i
        End With                                 ' aID(aPindex)

        If Not ShutUpMode Then
            Debug.Print Format(Timer, "0#####.00") & vbTab _
                                                 & "last cAttrDsc = " & i - 1 _
                                                 & " in Excel Sheet " & W.xlTName
        End If
        i = i + 1
        pArr(1) = i + 1
        pArr(2) = "Matchlisten (aktuelle Grundlage der Attribut Deskriptoren)"
        pArr(3) = b                              'Empty cell but force line wrap
        If sRules Is Nothing Then
            pArr(15) = "Invalid Class Rule"
        Else
            pArr(4) = TrueCritList
            pArr(10) = Trim(sRules.clsObligMatches.aRuleString)
            pArr(11) = Trim(sRules.clsNeverCompare.aRuleString)
            pArr(12) = Trim(sRules.clsNotDecodable.aRuleString)
            pArr(13) = Trim(sRules.clsSimilarities.aRuleString)
            ' pArr(14) = Trim(sRules.clsIgnoreProperties.aRuleString)
        End If
        Call addLine(W, i, pArr)
        i = i + 1
        W.xlTSheet.Cells(i, 1).ShrinkToFit = False
        W.xlTSheet.Cells(i, 1).EntireRow.VerticalAlignment = xlTop
        W.xlTSheet.Cells(i, 1).EntireRow.WrapText = True
        '    R4   Zi    R9   Zi
        W.xlTSheet.Range("D" & i & ":I" & i).Merge
        W.xlTSheet.Cells(1, 16) = i
        W.xlTSheet.Cells(2, 1).Select
        xOD.objDumpMade = 1
        W.xlTabIsEmpty = i
    Next xC
       
    ' Note: leaving W again (if we can)
    xlApp.ScreenUpdating = True
    xlApp.Cursor = xlDefault
    If Not O Is Nothing Then
        Set x = O
    End If
    If Not x Is x Then
        If x Is Nothing Then
            Set W.xlTSheet = Nothing
        Else
            Set W.xlTSheet = x.xlTSheet
        End If
    End If

FuncExit:
    Set xOD = Nothing

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' XLAccessingOL.AttrDscs2Excel

'---------------------------------------------------------------------------------------
' Method : Sub ShowItemProps
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowItemProps(thisPx As Long, Optional what As Long = 0)

Dim Line As String
Dim line2 As String
Dim what1 As String
Dim what2 As String
Dim a1 As cAttrDsc
Dim a2 As cAttrDsc

Dim i As Long
Dim iTop As Long
Dim Px2 As Long
    
'On Error GoTo 0
    Px2 = thisPx
    what1 = what
    what2 = what
    If thisPx = 2 Then
        DoVerify aID(2).idAttrCount < aID(2).idAttrDict.Count - 1, _
                                                              "The ItemPropertiesArray incomplete for index " & Px2
        iTop = aID(2).idAttrCount
    ElseIf thisPx = 3 Then
        If aID(2).idAttrCount < aID(2).idAttrDict.Count - 1 Then
            Px2 = 1
            iTop = aID(2).idAttrCount
            Line = "The ItemPropertiesArray incomplete for index 2"
        End If
        If aID(1).idAttrCount < aID(1).idAttrDict.Count - 1 Then
            If Px2 = 1 Then
                Px2 = 0
                Line = "The ItemPropertiesArray is incomplete for both index 1 and 2"
            Else
                Px2 = 2
                Line = "The ItemPropertiesArray is incomplete for index 1"
            End If
        End If
    Else
        If aID(1).idAttrCount < aID(1).idAttrDict.Count - 1 Then
            Line = "The ItemPropertiesArray incomplete for index 1"
            iTop = aID(1).idAttrCount
        Else
            Px2 = 1
            iTop = aID(1).idAttrDict.Count - 1
        End If
    End If
    If LenB(Line) > 0 Then
        Debug.Print Line
    End If
    
    If Px2 = 0 Then
        DoVerify False, Line
        GoTo FuncExit
    End If
    
    For i = 0 To iTop + 1
        If Px2 = 1 Then
            Call getWhats(aID(1).GetAttrDsc4Prop(i), Px2, Line, what1, what2)
            Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _
      & LString(what1, 6) & what2
        ElseIf Px2 = 2 Then
            Call getWhats(aID(2).GetAttrDsc4Prop(i), Px2, Line, what1, what2)
            Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _
      & LString(what1, 6) & what2
        Else
            Set a1 = aID(1).GetAttrDsc4Prop(i)
            Set a2 = aID(2).GetAttrDsc4Prop(i)
            Call getWhats(a1.adItemProp, Px2, Line, what1, vbNullString)
            Call getWhats(a2.adItemProp, Px2, line2, what2, vbNullString)
            Debug.Print LString(thisPx & "." & i & B2 & Line, 40) _
      & LString(what1, 6) & LString(line2, 20) & what2
        End If
    Next i
FuncExit:
    Set a1 = Nothing
    Set a2 = Nothing

End Sub                                          ' XLAccessingOL.ShowItemProps

'---------------------------------------------------------------------------------------
' Method : Sub ShowObjDesc
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowObjDesc(adItmDsc As cItmDsc)

Dim i As Long
Dim aDi As cAttrDsc

    Set aDi = New cAttrDsc
    ' aDi.iPindex = aPindex ???
    i = 0
    For Each aDi In adItmDsc.idAttrDict
        Debug.Print i, "dKey" & aPindex, "key=" _
                                      & aDi.adName _
                                      & " has key=" & aDi.adKey;
        i = i + 1
        If aDi.adItemProp Is Nothing Then
            Debug.Print ", but has no ItemProperty yet"
        Else
            Debug.Print ", Property resolved is named " & aDi.adItemProp.Name
        End If
    Next aDi

End Sub                                          ' XLAccessingOL.ShowObjDesc

'---------------------------------------------------------------------------------------
' Method : Sub getWhats
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub getWhats(adItemProp As ItemProperty, what As Long, Line As String, what1 As String, what2 As String)

Const zKey As String = "XLAccessingOL.getWhats"
    Call DoCall(zKey, tSub, eQzMode)
    
Dim Value As String

    Line = adItemProp.Name
    Value = "*** Error ***"
    Call Try
    Select Case what
    Case 0
        
    Case 1
        what1 = "W.xlTSheet: " & adItemProp.Type
        Value = vbNullString
    Case 2
        Value = "V: " & CStr(adItemProp.Value)
    Case Else
        what1 = "W.xlTSheet: " & adItemProp.Type
        Value = "V: " & CStr(adItemProp.Value)
    End Select
    what2 = LString(Value, 19) & b

FuncExit:
    Call ErrReset(0)

zExit:
    Call DoExit(zKey, Value)

End Sub                                          ' XLAccessingOL.getWhats

'---------------------------------------------------------------------------------------
' Method : Sub put1IntoExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub put1IntoExcel(xT As cXLTab, AP As cAttrDsc, AttributeIndex As Long)
'--- Proc MAY ONLY CALL Z_Type PROCS
Const zKey As String = "XLAccessingOL.put1IntoExcel"

    Call DoCall(zKey, "Sub", eQzMode)
    
    With AP
        pArr(1) = .adDictIndex
        pArr(2) = .adKey
        pArr(3) = .adtrueIndex
        If AP.adRules Is Nothing Then
            pArr(15) = "Seperator"
        Else
            pArr(4) = " --"                      ' ??? war AP.adRules.ruleattrdsc.addictindex
            pArr(5) = AP.adRules.clsObligMatches.RuleMatches
            pArr(6) = AP.adRules.clsNeverCompare.RuleMatches
            pArr(7) = AP.adRules.clsNotDecodable.RuleMatches
            pArr(8) = AP.adRules.clsSimilarities.RuleMatches
            pArr(9) = AP.adRules.RuleIsSpecific
            pArr(10) = AP.adRules.clsObligMatches.MatchOn
            pArr(11) = AP.adRules.clsNeverCompare.MatchOn
            pArr(12) = AP.adRules.clsNotDecodable.MatchOn
            pArr(13) = AP.adRules.clsSimilarities.MatchOn
            ' pArr(14) = AP.adRules.clsIgnoreProperties.MatchOn
            pArr(15) = AP.adRules.RuleType
        End If
    End With                                     ' AP
    
    Call addLine(xT, AttributeIndex, pArr)

FuncExit:

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                          ' XLAccessingOL.put1IntoExcel

'---------------------------------------------------------------------------------------
' Method : Sub put2IntoExcel
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub put2IntoExcel(px As Long, Line As Long)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard procConst zKey As String = "XLAccessingOL.put2IntoExcel"
Const zKey As String = "XLAccessingOL.put2IntoExcel"

    Call DoCall(zKey, "Sub", eQzMode)
    
    If xlApp Is Nothing Then
        GoTo FuncExit                            ' no Excel (yet)
    ElseIf O Is Nothing Then
        GoTo FuncExit                            ' no Excel sheet (yet)
    End If
    If px = 1 And AttributeIndex = 1 Then
        Call ClearSheetLines(O, 2)
    End If
    pArr(1) = aTD.adName                         ' = PropertyNameX, col1
    ' cols 2 and 3
    pArr(1 + px) = aTD.adShowValue               ' true value
    ' parr(4) = compare indication
    If px = 2 Then
        If Left(aTD.adDecodedValue, 1) = "#" Then
            pArr(6) = True                       ' visibility Flag, if empty: Show it
            pArr(5) = "***"                      ' probably no value accessible or exists
            pArr(4) = "___"
        ElseIf aTD.adShowValue _
               <> aTD.adDecodedValue Then
            If Left(aTD.adShowValue, 2) = "# " Then
                pArr(6) = False                  ' visible because value ignored
            End If
            pArr(4) = "..."
        End If
    End If
    ' cols 7 and 8
    If LenB(CStr(aTD.adDecodedValue)) > 0 Then   ' non-nomalized phone number
        pArr(6 + px) = CStr(aTD.adDecodedValue)
    End If
    ' cols 9 and 10
    pArr(8 + px) = aTD.adKillMsg
    Call addLine(O, Line, pArr)

FuncExit:

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                          ' XLAccessingOL.put2IntoExcel



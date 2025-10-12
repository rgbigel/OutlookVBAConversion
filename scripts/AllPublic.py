# Converted from AllPublic.py

# Attribute VB_Name = "AllPublic"
# Option Explicit

# ' Project Name: BerchtWare, stored in VBAProject.OTM
# '               All Files should be approximately the same
# '               Exception: the order of the dependencies always has the Master ahead of Slave App

# Option Base 1                                      ' arrays begin with index 1

# Public CallDepth As Long
# Public FastMode As Boolean                         ' requestd current Mode. Using current Error .atFastMode
# Public va As Variant                               ' dummies usable in Direct Window
# Public vb As Variant
# Public bb As Boolean
# Public ba As Boolean
# Public sa As String
# Public sb As String
# Public la As String
# Public lb As String


# Public TheMasterIsExcel As Boolean                 ' distinges running Master App from Slave
# Public olApp As Outlook.Application
# Public OlSession As Object

# Public Const UseTestStartDft As Boolean = False    ' sets various other debug options
# Public UseTestStart As Boolean                     ' init to UseTestStartDft
# Public CloneCounter As Long
# Public ErrExConstructed As Boolean                 ' "
# Public DidStop As Boolean                          ' "
# Public CallLogging As Boolean                      ' log Procs called by ProcStart
# Public LogZProcs As Boolean                        ' Additionally Log Z_Type Procs, no reset in N_DeInit
# Public UseStartUp As Long                          ' not reset in N_OnError and N_DeInit
# Public ExLiveCheckLog As Boolean                   ' not cleared by N_DeInit
# Public ExLiveDscGen As Boolean                        ' generate cProcDsc and cErr for all procs from LiveStack
# Public ZStartApp As cProcItem                     '  | K:Thisoutlooksession.ApplicationStartup

# Public Const b As String = " "                     ' single Blank
# Public Const B2 As String = "  "                   ' 2 Blanks
# Public Const Q As String = """"                    ' double Quote
# Public Const Bracket As String = "()"              ' open/close brackets (for Quote-Function)
# Public Const Pop As Boolean = False
# Public Const Push As Boolean = True                ' for simple stack calls readability
# Public Const WordSep As String = " !""$%&/()=?`@*+~'#_:;-.,><|"

# Public Const IsActive As String = "aktiv"
# Public Const InActive As String = "NICHT aktiv"
# Public Const LOGGED As String = "LOGGED"            ' Categories for items processed must contain this
# Public Const ErrorStatusCaption As String = "Error and Debug Status "
# Public Const IgnoredHeader As String = "' ###"
# Public Const Unbekannt As String = "Unbekannt"
# Public Const inv As Long = -999999                  ' indicates not inited value, associated object value is invalid for use
# Public Const testOne As String = "*"                ' all Errors allowed, only once, err is returned
# Public Const testAll As String = "**"               ' all allowed, Permitted stays, err is returned
# Public Const allowAll As String = "**0"             ' all allowed, Permitted stays, err is cleared
# Public Const allowNew As String = "*0"              ' all Errors allowed, only once, err is cleared

# #Const VBA71 = True                                 ' true for Office 16

# ' The following Object Libraries must be included IN THIS ORDER (dt. "Verweise")
# ' Visual Basic For Applications                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL
# ' Microsoft Outlook 16.0 ObjectLibrary              C:\Program Files (x86)\Microsoft Office\Root\Office16\MSOUTL.OLB
# ' OLE Automation                                    C:\Windows\SysWOW64\stdole2.tlb
# ' Microsoft Forms 2.0 Object Library                C:\Windows\SysWoW64\FM20.DLL
# ' Microsoft Excel 16.0 Object Library               C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE
# ' Redemption Outlook and MAPI COM Library           C:\Program Files (x86)\Redemption\Redemption.dll
# '   (Microsoft CDO 1.21 Library                     C:\Windows\System32\cdosys.dll atm not used)
# ' Microsoft Scripting Runtime                       C:\Windows\SysWOW64\scrrun.dll
# ' Microsoft Office 16.0 ObjectLibrary               C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL
# ' OutSpy Library                                    C:\Program Files (x86)\OutlookSpy\OutSpy.dll  optional, not used)
# ' Microsoft Word 16.0 Library                       C:\Program Files (x86)\Microsoft Office\Root\Office16\MSWORD.OlB
# ' Dialoge von c't Ralf Nebelo (manual register!)    D:\OneDrive\Berchtware\Outlook\Ct_DLLS\Dialoge.ocx
# ' (nur per Installation MZTools3 &| MZTools8)       C:\Users\rgbigel\AppData\Local\MZTools Software\MZTools8VBA\MZTools8VBA.dll
# ' (EverythingAccess: Errex, Errex_Helper, Errex_Callstack, Errex_DialogOptions, ErrexVariables)

# ' for Test Support, using ErrEx -> vbWatchDog, which must be installed
# ' http://www.everythingaccess.com/vbwatchdog.htm

# Public BugState As Long                            ' State when error occurred
# Public BugStateAsStr As String                     ' dito as String
# Public BugDlgRsp As Long                           ' State to use after user chose in ShowErrorDialog
# Public BugDlgRspAsStr As String                    ' dito as String
# Public BugWillPropagateTo As String                ' error handler propagation (to unwind ErrStack)

# ' **************************************************************************************
# ' If you have Classes which have displayable values, you can define one of them as  ****
# ' the default Class Value <ValueName>. To do this, first export the <ClassName>.cls ****
# ' if that worked, you should delete the Class in Outlook. Then, uing a text editor, ****
# ' insert the lines into <ClassName>.cls(with <values> changed, of course)           ****
# ' then Import <ClassName>.cls into Outlook. These are the pattern lines:            ****
# ' Attribute <ValueName>.VB_VarUserMemId = 0
# ' Attribute <ValueName>.VB_VarDescription = "Display Class Instance ID"
# ' lines above (without Comment ' ) must be placed immediately after the Declaration ****
# ' **************************************************************************************

# ' Alternatively do this with "Property Get/ Let <PropValue> so you could read/assign to the
# ' class directly (NOTE: The Prefix "VB_Var" then must be only "VB_")
# ' You can NOT have both a default member variable and a default procedure
# ' Attribute <PropValue>.VB_Description = vbNullString
# ' Attribute <PropValue>.VB_UserMemId = 0
# ' place the lines into <ClassName>.cls immediately after "Property Get <Value>"
# ' Constants that must be adjusted depending on Installation
# ' **************************************************************************************
# ' affected Classes at time of writing:           (In Property Get named:)
# ' cErr                                           Default
# ' cProcItem                                      -- .Key --

# Public Const EditProg As String = "D:\Tools\Notepad++\notepad++.exe"
# Public Const aPfad As String = "C:\Outlook-Dateien\Attachments\"
# Public Const lPfad As String = "C:\Outlook-Dateien\Logs\"
# Public Const LimitLog As Long = 100    ' allow for flush of immediate window
# Public Const MaxCharsPerLogFile As Long = 40000    ' about 5000 lines
# Public Const cPfad As String = "D:\OneDrive\@Faces\"
# Public Const tPfad As String = "D:\OneDrive\Berchtware\Outlook\"
# Public Const genCodePath As String = tPfad & "VbaProjectOTM\Working"
# Public Const iPfad As String = "C:\Users\rgbig\AppData\Local\Temp\OutlookItemCompare.txt"
# Public Const TemplateFolder As String = tPfad
# Public Const TemplateFile As String = TemplateFolder & "OutlookDefault.xlsm"
# Public Const cOE_SheetName As String = "Objekteigenschaften"

# Public Const MailAccount1 As String = "rgbigel@web.de"
# Public Const MailAccount2 As String = "Rolf.Bercht@Gmail.com"

# Public Const BackupPSTname As String = "Backup"    ' Dieser MAPI-Ordner muss vorher schon existieren!!
# ' als oberster Ordner eines PST-Files
# ' Der Dateiordner UNTER aPfad heisst genau so!
# ' auch dieser Dateiordner muss angelegt sein.

# Public Const StdInboxFolder As String = "Posteingang" ' Language dependent!
# Public Const StdSentFolder As String = "Gesendet"
# Public Const BackupInboxFolder As String = "Erhalten" ' Ordner im BackupPSTname
# Public Const BackupAggregatedInbox As String = "AggregatedInbox" ' Ordner im BackupPSTname
# Public Const BackupUnknownFolder As String = "Unbekannte"
# Public Const BackupSMSFolder As String = "SMS"
# Public Const BackupSentFolder As String = "Gesendete Elemente"
# Public Const SpecialSearchFolderName As String = "NLOGGED"
# Public Const WebInboxFolder As String = "Inbox"

# Public Const SchemaPropTag As String = "http://schemas.microsoft.com/mapi/proptag/"
# Public Const SchemaMapiID As String = "http://schemas.microsoft.com/mapi/id/"
# Public Const SchemaMapiString = "http://schemas.microsoft.com/mapi/string/"
# Public Const IgnoredAttachmentNames As String = "image.* *.png *.gif" ' split below once
# Public Const MailEventsViaRules As Boolean = False ' CopyTo done by Outlook Rules
# ' copying mail events done by Outlook Rules
# ' if False: will do the CopyTo here
# ' Rules are reliably lost (bug) Value=>>False
# Public Const CopyOriginal As Boolean = False       ' bei Papa False, Bei Marc True
# Public Const SetAllToHTML As Boolean = True
# Public Const SaveAttachmentMode As Boolean = False ' Attachments im HTML-Teil auch speichern?
# Public Const DelSavedAttachments As Boolean = False ' leave in mail or not
# Public Const TriStatsUndefined As Long = 99        ' used in Class cTriState
# Public Const maxDeferredLimit As Long = 50         ' number of items currently not processed because we could have storage problems
# Public Const DropIndent As Long = 1                ' .atCallDepth - DropIndent is used to indent visual log

# Public Const BadDate As String = "01.01.4501"
# Public Const DefaultDebugSleep As Long = 5         ' wait time to see debug output in direct window
# Public Const dModule As String = "BugHelp"
# Public dModuleWithP As String                      ' = dModule & "." done in Z_ProcInitErrHdl
# Public Const pad8 As String = "!@@@@@@@@"
# Public Const padDbl8 As String = "#####0.00"
# Public Const DebugControlsWanted As Boolean = True ' !!!!!!! Show aNonModalForm (implementation choice)
# Public Const ZEntryKey As String = "ErrorHandler.ProcCall"
# Public Const ZEntryType As String = "Function"
# Public Const tPropGet As String = "Property Get"
# Public Const tPropLet As String = "Property Let"
# Public Const tPropSet As String = "Property Set"
# Public Const tAppl As String = "Application"
# Public Const tFunction As String = "Function"
# Public Const tSub As String = "Sub"
# Public Const tSubEP As String = "Sub EP"

# Public Const ScalarTypeNames As String = "Integer Long Single Double Date String Boolean LongLong "
# Public ScalarTypes As Variant                      ' array of long
# Public ScalarTypeV As Variant                      ' array of long
# Public dSType As Dictionary                        ' for reverse lookup

# Public Const MaxIndent As Long = 20
# Public Const Hell As Long = 99                     ' error associated with DebugControl.TermRQ
# Public Const lKeyM As Long = 58                    ' length we can afford for printing ProcName / Key (incl. indent)
# Public Const lDbgM As Long = 20                    ' length we can afford for printing DbgID
# Public Const lCallInfo As Long = 16                ' length we can afford for printing Live Caller Information

# Public Const EntryLead As String = " -> "
# Public Const ExitLead As String = " <- "
# Public Const RetLead As String = " -- "
# Public Const ErrLifeTime As Long = 200             ' max size of C_CallTrace
# Public Const ErrLifeKept As Long = 100             ' kept size of C_CallTrace after reduce

# Public ZApplication_Startup As Long
# Public ZCopyAllBackupCategories As Long
# Public ZCopyAllHotmailCategories As Long
# Public ZCreateRules As Long
# Public ZLoopFoldersDialog As Long
# Public ZMPEmap As Long
# Public ZNoDupes As Long
# Public ZRunMissedRules As Long
# Public ZSelectAndFind As Long
# Public ZExcelShowItem As Long
# Public ZWasEmailProcessed As Long

# ' Error Handler interface -------------------------->>> See Module BugHelp / Z_ErrIf <- Generated by ZZZIfGen...

# ' Constants for Outlook Macros
# Public Const Passt_Synch As Long = 0
# Public Const Passt_Inserted As Long = 1
# Public Const Passt_Deleted As Long = 2
# Public Const MaxMisMatchesForCandidatesDefault As Long = 5

# Public Const InstanceRule As String = "InstanceRule"
# Public Const ClassRules As String = "ClassRules"
# Public Const DefaultRule As String = "DefaultRule"

# Public Const minPropCountForFullItem As Long = 10
# Public Const DaySeconds As Double = 1# / 86400#
# Public Const DayMinutes = DaySeconds * 60#
# Public Const DftBugStateAge As Double = 3#         ' frequency for querying ErrorState Changes in Seconds

# ' Interface Outlook<=>Excel for Editing
# ' the following constants must be the same as in "OutlookDefault.XlSM" EditWatch
# Public Const ChangeCol1 As Long = 2
# Public Const ChangeCol2 As Long = 3
# Public Const ChangeCol3 As Long = 4
# Public Const ValidCol As Long = 5
# Public Const flagColumn As Long = 6
# Public Const clickColumn As Long = 9
# Public Const moreColumn As Long = 10
# Public Const promptColumn As Long = 11
# Public Const OriginalCols As Long = 13
# Public Const WatchingChanges As Long = 14
# Public Const changeCounter As Long = 15
# Public Const Old1 As Long = 16
# Public Const Old2 As Long = 17

# ' Global Interface variables of ErrorHandler

# Public HasRunBefore As Boolean                     ' NONE of these up to blank line are restored in N_DeInit
# Public NoPrintLog As Boolean
# Public UseErrExOn As String                        ' Name of global error handler to choose
# Public LastErrExOn As String
# Public ErrExActive As Boolean

# Public ZErrSnoCatch As Boolean                    ' No ErrHandler Recursion;  NO N_PublishBugState
# Public ZErrNoRec As Boolean                       ' No ErrHandler Recursion; use N_PublishBugState

# Public ProcDbg As String
# Public WithLiveCheck As Boolean

# Public ExternalEntryCount As Long                  ' this is NOT cleared in N_OnError!!

# Public S_ActModule As String
# Public S_AppIndex As Long                          ' D_AppStack.Count - 1
# Public S_AppKey As String                          ' Key of Active App-Lvlel Proc active
# Public S_DbgId As String                           ' active proc (after EnterProc/ProcReturn)
# Public S_Key As String                             ' called key after last ProcCall (gen. ProcMakerDsc)
# Public S_ActKey As String                          ' active        on last ProcCall  vbNullString

# ' Variants used for the not existing Method "ToString": not constants because can't set constant string array

# Public Const ExStackProcString As String = "S0 Sub " & "Function " & "GET " & "LET " & "SET "
# Public ExStackProcNames As Variant                 ' inited in N_Prepare

# Public MacroArray As Variant                       ' Split from MacroList

# Public QModeNames As Variant                       ' names(Qmode)
# Public Const QModeString As String = "noDef zMode yMode xMode RecMode quietMode EPMode stdMode AppRec"

# ' QmodeNames = split(QmodeString)                ' see Inits for ErrHandlerModule in Proc N_Prepare
# Public Enum eQMode                                 ' Default .ErrLevels below when using Do/Proc Call()

# eQnoDef = 0                                    ' used as a mark for a proc NOT using DoCall or ProcCall
# ' not part of D_ErrInterface unless generating in DoCall
# eQzMode = 1                                    ' .ErrLevel Lcrit.= 4   Use DoCall/DoExit, not traced
# '    may not use ProcCall, must not use ProcExit, trivial proc
# eQyMode = 2                                    ' like eQzMode, (traced ???,) logged If StackDebug >= 8
# eQxMode = 3                                    ' .ErrLevel Lmin  = 3   Proc is in D_ErrInterface, not checking recursion
# '  may not use Z_EntryPoint/Z_AppExit/Z_AppStackRem,
# '   will define via DefDsc
# eQrMode = 4                                    ' .ErrLevel Lmin  = 3   Recursive module
# eQuMode = 5                                    ' .ErrLevel Lmin  = 3   do not log routine
# ' following are on      Application Level       (with D_AppStack and C_CallTrace)
# eQEPMode = 6                                   ' .ErrLevel Lall  = 1   Entry Point mode
# eQAsMode = 7                                   ' .ErrLevel Lmin  = 3   use C_CallTrace, and log as normal
# eQArMode = 8                                   ' .ErrLevel Lsome = 2   Recursive Application module
# End Enum

# Public CStateNames As Variant                      ' names(CState)
# Public Const CStateString As String = "undef Exited Paused Active"
# ' CStateNames = split(CStateString)              ' see Inits for BugHelp in Proc: N_Prepare
# Public Enum eCState                                ' values of CallStatus
# eCUndef = 0                                    ' proc defined, but never called
# eCExited = 1                                   ' proc exited via ProcReturn (Z_AppExit)
# eCpaused = 2                                   ' proc running, but interrupted
# eCActive = 3                                   ' proc currently active
# End Enum

# Public LogLevelNames As Variant                    ' names(LogLevel)
# Public Const LogLevelString As String = "*debugmode* LogAll LogSome LogMin LogCritical LogNothing"
# ' LogLevelNames = split(LogLevelString)          ' see Inits for BugHelp in Proc N_Prepare
# Public Enum eLogLevel
# eLdebug = 0
# eLall = 1
# eLSome = 2
# eLmin = 3
# eLcritical = 4
# eLnothing = 5
# End Enum

# Public OkValueNames As Variant                     ' names(OkValue)
# Public Const OkValueString As String = "unspec bad triv ok obj errObj"
# ' OkValueNames = split(OkValueString)            ' see Inits for BugHelp in Proc N_Prepare
# Public Enum eOkValue
# eVunspec = 0
# eVbad = 1
# eVtriv = 2
# eVok = 3
# eVobj = 4
# eVerrObj = 5
# End Enum

# Public PushTypes As Variant                        ' names(OkValue)
# Public Const PushTypeString As String = "unk Lifo FiFo"
# ' PushTypes = split(PushTypeString)         ' see Inits for BugHelp in Proc N_Prepare
# Public Enum ePPushType
# ePunk = 0
# ePLiFo = 1
# ePFiFo = 2
# End Enum

# Public AccountTypeNames As Variant                 ' names(AccountTypeNames)
# Public Const AccountTypeString As String = "Exchange IMAP POP3 HTTP EAS unknown"
# ' AccountTypeNames = split(AccountTypeString)    ' see Inits for BugHelp in Proc N_Prepare
# Public Enum eAccountType
# eAunk = 0
# eAExchange = 1
# eAIMAP = 2
# eAPOP3 = 3
# eAHTTP = 4
# eAEAS = 5
# End Enum

# Public ExModeNames As Variant                      ' names(ExchangeConnectionMode)
# Public Const ExModeNamesString As String = "NoExchange Offline CachedOffline DisConnected " & _
# "CachedDisconnected CachedConnectedHeaders " & _
# "CachedConnectedDrizzle CachedConnectedFull Online"

# ' *********** Make sure all Vars, Objects and Arrays below are included in N_DeInit **************
# ' Simple Variables

# Public aBugVer As Boolean
# Public aBugTxt As String
# Public aCallState As String                        ' short - lived result of last cTraceEntry Get/Let/ECallTrace
# Public aAccountNumber As Long
# Public aAccountType As OlAccountType
# Public aAccountTypeName As String
# Public AcceptCloseMatches As Boolean
# Public actDate As String
# Public ActionID As Long
# Public AdditionalLine As String                    ' used after obtaining performance data
# Public Addit_Text As Boolean
# Public aDebugState As Boolean
# Public aItmIndex As Long
# Public loopItmIndex As Long
# Public TotalDeferred As Long
# Public AllDetails As String
# Public AllItemDiffs As String
# Public AllProps As Boolean
# Public AllPropsDecoded As Boolean
# Public aNewCat As String
# Public aPindex As Long                             ' current index in aDecProp/c
# Public apropTrueIndex As Long                      ' position of aProp in ItemProperties, base 0
# Public AskEveryFolder As Boolean
# Public askforParams As Boolean
# Public aStringValue As String                      ' attribute string after decoding
# Public aTimeFilter As String
# Public AttributeIndex As Long
# ' Public ADKey As String === aTD.ADKey
# Public b1text As String                            ' Labels for buttons in dialogs
# Public b2text As String
# Public b3text As String
# Public BaseAndSpecifiedDiffer As Boolean           ' if specified item type is exception or occurrence then aOjb(px) is standard object type and px+2 is exc/occ
# Public bDefaultButton As String
# Public CallNr As Long                              ' number of calls, for log
# Public cHdl As String
# Public ClientIsDefProc As Boolean                  ' valid only if ProcCall is active
# Public CloningMode As Boolean                      ' modifies Class_Initialize for some classes
# Public cMisMatchesFound As Long
# Public cMissingPropertiesAdded As Long
# Public CondensedPhoneNumber As String
# Public ContactFolder As Folder
# Public FolderAggregatedInbox As Folder
# Public ContactFolderName As String
# Public ContactFolderPath As String
# Public curFolderPath As String
# Public CutOffDate As Date
# Public DateId As String
# Public DateIdNB As String
# Public DateSkipCount As Long
# Public dcCount As Long
# Public BugStateElapsed As Double
# Public DebugControlsUsable As Boolean
# Public DebugLogging As Boolean
# Public DebugMode As Boolean                        ' use SetDebugModes for permanent values
# Public DebugSleep As Long
# Public dontIncrementCallDepth As Boolean           ' needed for predefines
# Public ElapsedTime As Double                       ' time consumed until last Exit or Pause
# Public ErrStatusFormUsable As Boolean
# Public DeletedItem As Long
# Public DeleteIndex As Long                         ' 1 or 2: which was deleted? When 0: nothing, when inv no reuse possible
# Public DeleteNow As Boolean
# Public DftItemClass As OlObjectClass
# Public DftItemType As OlItemType
# Public DftItemTypeName As String
# Public DidItAlready As Boolean                     ' indicator for De_Init done, force to reset
# Public Diffs As String
# Public DiffsIgnored As String
# Public DiffsRecognized As String
# Public displayInExcel As Boolean
# Public DontCompareListDefault As String
# Public dontLog As Boolean
# Public eActFolderChoice As Boolean                 ' SelectorMode 1
# Public eAllFoldersOfType As Boolean                ' SelectorMode 2
# Public eOnlySelectedFolder As Boolean              ' Selectormode 4
# Public eOnlySelectedItems As Boolean               ' Selectormode 3
# Public EntryPointCtr As Long
# Public EPCalled As Boolean                         ' called via ***Entry Point***
# Public ErgebnisseAlsListe As Boolean
# Public ErrStatusHide As Boolean                    ' determines visibility of status display (frmErrStatus)
# Public ErrDisplayModify As Boolean                 ' refresh of error status display needed (frmErrStatus)
# Public ErrorCaught As Long                         ' raw error last set by N_OnError
# Public ErrStackProcs As Long
# Public BlockEvents As Boolean                      ' Stop responding to Events (seperate from BugTimer)
# Public OnlyMostImportantProperties As Boolean      ' current rule, derived from user-specified
# Public EventHappened As Boolean
# Public ExceptionProcessing As Boolean              ' processing exception of recursion pattern, do not recurse this
# Public ExtendedAttributeList As String             ' Properties not used for Sorting
# Public FavNoLogCtr As Long
# Public FilterCriteriaString As String              ' used for display filters
# Public FindMatchingItems As Boolean
# Public FirstAddPos As Long
# Public FldCnt As Long
# Public FolderLoopIndex As Long
# Public FolderPathLevel As Long
# Public HeadlineName As String
# Public IgnoredPropertyComparisons As Long
# Public IgnoreUnhandledError As Boolean             ' Inside some routines, do not log bad behaviour
# Public IgString As String
# Public InlineZ_ProcLog As String                   ' procs traced seperated by CRLF
# Public InOrOut As String
# Public iRuleBits As String                         ' ??? for debugging only
# Public IsComparemode As Boolean
# Public IsEntryPoint As Boolean
# Public isNonLoopFolder As Boolean                  ' the real result of NonLoopFolder without user interaction
# Public isSpecialName As Boolean                    ' key with #B
# Public isSQL As Boolean
# Public isUserProperty As Boolean                   ' key with #W2
# Public ItemInIMAPFolder As Boolean
# Public ItsAPhoneNumber As Boolean
# Public killMsg As String
# Public killStringMsg As String
# Public killType As String
# Public LastAddPos As Long

# Public lHeadLine As Long
# Public MainProfileAccount As String
# Public lMinus As Long
# Public MinusLine As String
# Public ModelLine(1 To 3) As String                 ' for dynamic progress
# Public NoTimerEvent As Boolean                     ' suppress setting timer events if true

# Public LineToAdd As String
# Public ListCount As Long
# Public LoeschbesttigungCaption As String
# Public LogImmediate As Boolean
# Public LogAllErrors As Boolean
# Public LogAppStack As Boolean                      ' not in N_OnError!!
# Public LogicTrace As String
# Public LogPerformance As Boolean                   ' not in N_OnError!!
# Public LogSelection As String
# Public LListe As String
# Public LSD As Long                                 ' Minimal Stack Depth of interest, LSD = ZAppStart.Erractive.atLiveLevel -1
# Public MailModified As Boolean
# Public MainObjectIdentification As String          ' the one we use
# Public Matchcode As Long
# Public MatchData As String
# Public Matches As Long
# Public MatchMin As Long
# Public MaxMisMatchesForCandidates As Long
# Public MaxPropertyCount As Long
# Public MayChangeErr As Boolean
# Public Message As String
# Public MinimalLogging As Long
# Public MPEItemDiffs As String
# Public mustDecodeRest As Boolean
# Public NoEventOnAddItem As Boolean
# Public NoInboxRecursion As Boolean
# Public NormalizedPhoneNumber As String
# Public NoSentRecursion As Boolean
# Public NotDecodedProperties As Long
# Public objTypName As String
# Public OffAdI As Long                              ' Various Head Line Columns
# Public OffPrN As Long
# Public OffMTS As Long
# Public OffLvl As Long
# Public OffCal As Long                              ' offset of Head Line Call-Column
# Public OffObj As Long
# Public OffTim As Long
# Public OneDiff As String
# Public OneDiff_qualifier As String
# Public onlyNew As Boolean                          ' true: do not look for existing property names
# Public OlOpenedHere As Boolean
# Public XlOpenedHere As Boolean
# Public QuitStarted As Boolean
# Public Output As String
# Public PhoneNumberNormalized As Boolean
# Public PickTopFolder As Boolean
# Public AppStartComplete As Boolean
# Public PropertyNameX As String
# Public PropertyNameY As String
# Public PropertyX As String
# Public PropertyY As String
# Public PropStatesString As String
# Public ProtectStackState As Long                   ' note: can only change when stack depleted
# Public quickChecksOnly As Boolean                  ' rule specified by User dialog
# Public relevantIndex As Long
# Public RestrictCriteriaString As String            ' Selection Parameters
# Public rID As Long
# Public RootCreated As Boolean
# Public rPTrueIndex As Long
# Public SaveAttachments As Boolean
# Public saveItemNotAllowed As Boolean
# Public SearchCriteriaString As String
# Public SearchFolderNameResult As String
# Public SelectedAttributes As String                ' Attribute names which are MustMatch or Similar
# Public SelectMulti As Boolean
# Public SelectOnlyOne As Boolean
# Public SelectorMode As Long                        ' selected SelectorButton x
# Public sHdl As String
# Public ShowEmptyAttributes As Boolean
# Public ShowFunctionValues As Boolean
# Public ShutUpMode As Boolean                       ' do not create log output
# Public SIC As Long
# Public SimilarityCount As Long
# Public SkipDontCompare As Boolean
# Public SkipNextInteraction As Boolean
# Public SkipedEventsCounter As Long                 ' count number of times we skipped "add" action
# Public SortMatches As String
# Public SpecialObjectNameAddition As String         ' Name extended by this string for px>2 (#E, #O)
# Public StackDebug As Long                          ' Level of debugging, 0=none, >0 More and more
# Public StackDebugOverride As Long                  ' override first getDebugMode Value

# Public StopLoop As Boolean
# Public StopRecursionNonLogged As Boolean
# Public stpcnt As Long
# Public StringMod As Boolean                        ' side-result of AppendTo: and Remove: true if change occurred
# Public StringsRemoved As String
# Public SuperRelevantMisMatch As Boolean
# Public SuppressStatusFormUpdate As Boolean
# Public testNonValueProperties As String
# Public TestTail As String                          ' the part of Testvar following "|", + leading blank
# Public Testvar As String
# Public TotalPropertyCount As Long                  ' count of properties in decoded item(s). must be same for two compared items
# Public ItemsToDoCount As Long
# Public TraceMode As Boolean
# Public TraceTop As Long                            ' top position in (circular) Trace buffer
# Public TrashFolderPath As String
# Public TrueCritList As String                      ' as requested by user
# Public UInxDeferred As Long
# Public UInxDeferredIsValid As Boolean
# Public UI_DontUseDel As Boolean                    ' request override of frmDelParms
# Public UI_DontUse_Sel As Boolean                   ' request override of SelectionParameter
# Public UI_Show_Del As Boolean                      ' allow user to choose override
# Public UI_Show_Sel As Boolean                      ' allow user to chose from SelectionParameter
# Public DeferredLimit As Long                       ' number of items currently not processed because we could have storage problems
# Public DeferredLimitExceeded As Boolean
# Public UserDecisionEffective As Boolean
# Public UserDecisionRequest As Boolean
# Public UTCisUsed As Boolean
# Public Z_ExceptionList As String                   ' list of DbgIds of Z_Type procs. Dynamic build, but could be constant when stable

# Public WantConfirmation As Boolean
# Public WantConfirmationThisFolder As Boolean
# Public MaintenanceAction As Long
# Public WithProcVal As Variant                      ' Function wants a log of the value it produces
# Public workingOnNonspecifiedItem As Boolean        ' specified aID(px + 2) is exception or occurrence
# Public xDeferExcel As Boolean                      ' do not create excel table until required
# Public xNoColAdjust As Boolean
# Public xReportExcel As Boolean
# Public xUseExcel As Boolean
# Public YleadsXby As Long

# Public actOnlineStatus As String

# ' Objects that can be set to Nothing (see in DeInitialize)

# Public aAccountDsc As cNumbItem
# Public aCell As Excel.Range
# Public aNonModalForm As Object                     ' if there an active Form with ShowModal=False: <> nothing
# Public ActItemObject As Object                     ' (Mapi-)item currently selected; e.g. via GetAobj, Loops
# Public aInfo As New cInfo                          ' for value decoding in hasValue (volatile!! No N_DeInit)
# Public aItemList As cItemList                      ' if we need overview of many items
# Public aNameSpace As NameSpace
# Public aObjDsc As cObjDsc                          ' the currently selected Object Description
# Public aItmDsc As cItmDsc                          ' the currently selected Item Description
# Public aProp As ItemProperty                       ' after selecting from ItemProperties/aProps (or making aTD)
# Public aProps As ItemProperties
# Public aRDOSession As Redemption.RDOSession
# Public aStore As Outlook.Store                     ' selected item in oStores
# Public aTD As cAttrDsc
# Public CalendarFolder As Folder
# Public ChosenTargetFolder As Folder                ' target Folder for Find / Match operations
# Public ClassDescriptorItems As Dictionary
# Public CopiedObjDsc As cObjDsc                     ' new objDesc after Copy or Move item
# Public oStores As Outlook.Stores
# Public CurIterationSwitches As cIterationSwitches
# Public DeletionCandidates As Dictionary
# Public dftRule As cAllNameRules                    ' DefaultRule
# Public FInFolderColl As Collection
# Public FolderBackup As Folder
# Public FolderInbox As Folder
# Public FolderSent As Folder
# Public FolderSMS As Folder
# Public FolderTasks As Folder
# Public FolderUnknown As Folder
# Public FRM As Object                               ' used for Forms with
# Public FWP_LBF_Hdl As cFindWindowParms
# Public FWP_xLW_Hdl As cFindWindowParms
# Public FWP_xMainWinHdl As cFindWindowParms
# Public iRules As cAllNameRules                     ' current InstanceRule
# Public killWords As Collection
# Public LBF As Object
# Public ListContent As Collection                   ' List of Items to process (1, 2 or all in loop)
# Public logItem As MailItem
# Public LookupFolders As Folders
# Public LCA As cCallEnv                             ' extracted from ErrEx.LiveCallStack for calling Procedure
# Public LCS As ErrExCallstack                       ' used to store ErrEx.LiveCallStack
# Public MainFolderContacts As Folder
# Public MainFolderInbox As Folder
# Public MainFolderSent As Folder
# Public MSR As Object
# Public OlBackupHome As Folder
# Public OlGooMailHome As Folder
# Public OlHotMailHome As Folder                     ' the top priority Folder(s) for which we define events:
# Public OlWEBmailHome As Folder
# Public oSearchFolders As Outlook.Folders
# Public D_TC As Dictionary                          ' all cObjDsc
# Public ParentFolder As Folder
# Public RestrictedItemCollection As Collection      ' set up by last getRestrictedItems
# Public RestrictedItems As Items                    ' Items after getRestrictedItems (corresponds to Collection)
# Public sDictionary As Dictionary                   ' source used for clones / reusing Dictionary structure
# Public SelectedItems As Collection
# Public SelectedObjects As Object
# Public SessionAccounts As Accounts
# Public SQLpropC As Collection
# Public sRules As cAllNameRules                     ' current ClassRules
# Public topFolder As Folder
# Public TopFolders As Folders
# Public TrashFolder As Folder
# Public Deferred As Collection                      ' collection of action objects
# Public UserRule As cAllNameRules
# Public OlExplorer As cOlExplorer                   ' ExplorerEvents
# Public NCall As cProcItem                         ' DoCall atDsc

# ' Variant Objects

# Public MostImportantProperties As Variant          ' Same as below, but without Operators, Array
# Public TrueImportantProperties As Variant          ' Array of Operators and Property Names must be decoded and Matched to search/sort criteria (if poss.)
# Public xColList As Variant
# Public MostImportantAttributes As String           ' MostImportantAttributes as String

# ' BugHelp Interfaces that need New

# Public BugTimer As New cBugTimer                   ' do not De_Init
# Public ExternCaller As New cProcItem               ' simulated ProcDsc for external caller | K:Extern.Caller R:Extern.Caller
# Public C_AllPurposeStack As New Collection         ' any class of objects in this
# Public C_CallTrace As New Collection               ' Collectionof cErr for tracing calls
# Public C_ProtectedStack As New Collection          ' see N_Suppress for example
# Public C_PushLogStack As New Collection            ' Collectionof Strings meant for logging
# Public D_AccountDscs As Object
# Public D_ErrInterface As New Dictionary            ' the entire set of Procs, names, Keys, types, etc.
# Public D_LiveStack As New Dictionary               ' when accessing LiveStack, this is the current setting
# Public D_PushDict As New Dictionary                ' Dictionary of cPush objects
# Public E_Active As New cErr                        ' Active Error Environment
# Public E_AppErr As New cErr                        ' Current Application cErr. Need not match .ErrActive
# Public E_PrevCallTrace As New cErr                 ' Previous cErr on C_CallTrace if tracing.
# ' (Won't match P_Previous.ErrAct usually)
# Public P_Active As New cProcItem                   ' Active Proc for E_Active, == E_Active.atDsc
# Public P_CurrEntry As New cProcItem                ' the technical base of all calls (below root, above InitErrHandler) | K:CurrEntry
# Public P_EntryPoint As New cProcItem               ' Z_AppEntry Description. .ErrActive Identifies last active (requires .atRecursionOK=True) | K:N_CallPointDsc
# Public P_LastEP As New cProcItem                   ' Previous Entry Point Proc | K:LastEP
# Public T_DC As New cTermination

# ' Loop Folders Module Globals

# Public LF_UsrRqAtionId As Long                     ' Corresponds to LoopFoldersDialog:
# ' Action selections 1 ... 8 (used so far)
# Public LF_ItemCount As Long
# Public LF_ItmChgCount As Long
# Public LF_DoneFldrCount As Long
# Public LF_CurLoopFld As Folder
# Public LF_recursedFldInx As Long
# Public LF_CurActionObj As cActionObject
# Public LF_DontAskAgain As Boolean

# ' Arrays

# Public ActiveExplorerItem(1 To 2) As Object
# Public aDecProp(1 To 4) As cAttrDsc                ' note: decoded properties are called attributes
# Public aOD(0 To 4) As cObjDsc                      ' describe current objects,
# Public aID(0 To 4) As cItmDsc                      ' describe current objects,
# ' index=0 is for O = Object Properties (Union of aID(*)
# ' set of decoded properties, i.e. attributes,
# ' contained in idAttrDict (aID(apindex).idAttrDict) as item(s)
# Public AttributeUndef(1 To 4) As Long              ' Internal structure wrong if >0
# Public delSource(0 To 2) As Long
# Public Fctr(2) As Long
# Public fiBody(1 To 2) As String
# Public fiMain(1 To 2) As String
# Public Folder(2) As Folder
# Public FullFolderPath(0 To 5) As String
# Public Ictr(2) As Long

# Public WorkIndex(1 To 2) As Long
# Public WorkItem(1 To 2) As Object
# Public WorkItemMod(2) As Boolean
# Public lFolders() As Variant                        ' As Folder no longer valid in O15
# Public MatchPoints(1 To 2) As Long
# Public pArr(1 To 30) As String
# Public rP(1 To 2) As Outlook.RecurrencePattern
# Public sortedItems(2) As Items                      ' is the sorted Folder(x) items list!
# Public DeferredFolder(6) As Folder
# Public LoggableFolders As New Dictionary            ' key: FolderPath, Items: Folder
# Public NLoggedName As String                        ' "Logged & NLoggedName"

# ' *********** Make sure all Vars, Objects and Arrays above are included in N_OnError **************

# ' ### folgende Zeilen werden automatisch im debuglog ausgegeben,
# ' ### bei nderungen der ActionTitle mssen sie updated werden

# ' ### Start of generated code
# Public ActionTitle(0 To 8) As String
# Public Const atDefaultAktion As Long = 1
# Public Const atKategoriederMailbestimmen As Long = 2
# Public Const atPostEingangsbearbeitungdurchfhren As Long = 3
# Public Const atDoppelteItemslschen As Long = 4
# Public Const atNormalreprsentationerzwingen As Long = 5
# Public Const atOrdnerinhalteZusammenfhren As Long = 6
# Public Const atFindealleDeferredSuchordner As Long = 7
# Public Const atBearbeiteAllebereinstimmungenzueinerSuche As Long = 8
# Public Const atContactFixer As Long = 9

# ' Constants for Cloning of Object Class instances; Enum does not work (VBA bug)
# Public Const DummyTarget As Long = -2
# Public Const FullCopy As Long = 1
# Public Const withNewValues As Long = 2

# Public sourceIndex As Long
# Public targetIndex As Long
# Public aCloneMode As Long                          ' Cloning mode selected for next Clone of Class Instance
# Public f As Long                                   ' implicit 'for' loops

# ' ### End of generated code

# '---------------------------------------------------------------------------------------
# ' Method : Sub N_DeInit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_deinit():
    # '--- Proc atDsc is Z_Type: Calls nothing
    # '------------------- gated Entry -------------------------------------------------------
    if OlSession Is Nothing Then:
    if Not HasRunBefore Then                   ' really first run:
    # HasRunBefore = True
    # DidItAlready = True
    # GoTo ProcRet

    # Dim i As Long
    # Dim hDsc As cProcItem

    # Err.Clear

    if IsMissing(AlreadyDone) Then:
    # DidItAlready = False
    else:
    # DidItAlready = Not CBool(AlreadyDone)

    if DidItAlready Then:
    print(Debug.Print "N_DeInit Skipped because we DidItAlready")
    # Err.Clear
    # GoTo FuncExit
    else:
    print(Debug.Print "N_DeInit is starting")
    if Not D_ErrInterface Is Nothing Then      ' procs must be undefined:
    if Not isEmpty(D_ErrInterface.Items(i)) Then:
    # Set hDsc = D_ErrInterface.Items(i)
    # Call N_Undefine(hDsc, i)
    # Set D_ErrInterface = New Dictionary

    # DidItAlready = True

    # ' S_ActKey =""                              omitted here on purpose!
    # ' LogPerformance = False                    omitted here on purpose!
    # ' Testvar = vbNullString
    # 'CallDepth = 0
    # aAccountNumber = 0
    # aAccountType = 0
    # aAccountTypeName = vbNullString
    # AcceptCloseMatches = False
    # actDate = vbNullString
    # ActionID = 0
    # AdditionalLine = vbNullString
    # Addit_Text = False
    # aBugVer = False
    # aDebugState = False
    # ' ADKey = vbNullString === aTD.ADKey
    # aItmIndex = 0
    # TotalDeferred = 0
    # AllDetails = vbNullString
    # AllItemDiffs = vbNullString
    # AllProps = False
    # AllPropsDecoded = False
    # aNewCat = vbNullString
    # aPindex = 0
    # AppStartComplete = False
    # apropTrueIndex = 0                         ' position of aProp in ItemProperties, base 0
    # AskEveryFolder = False
    # askforParams = False
    # aStringValue = vbNullString                ' attribute string after decoding
    # aTimeFilter = vbNullString
    # AttributeIndex = 0
    # b1text = vbNullString                      ' Labels for buttons in dialogs
    # b2text = vbNullString
    # b3text = vbNullString
    # BaseAndSpecifiedDiffer = False             ' if specified item type is exception or occurrence then aOjb(px) is standard object type and px+2 is exc/occ
    # bDefaultButton = vbNullString
    # ' not wanted: CallLogging = False
    # CallNr = 0
    # cHdl = vbNullString
    # CloneCounter = 0
    # CloningMode = False                        ' modifies Class_Initialize for some classes
    # cMisMatchesFound = 0
    # cMissingPropertiesAdded = 0
    # CondensedPhoneNumber = vbNullString
    # ContactFolderName = vbNullString
    # curFolderPath = vbNullString
    # CutOffDate = 0
    # DateId = vbNullString
    # DateIdNB = vbNullString
    # DateSkipCount = 0
    # dcCount = 0
    # DebugControlsUsable = False
    # ' DebugLogging = False
    # ' DebugMode = False                     ' use SetDebugModes for permanent values
    # DebugSleep = 0
    # DeferredLimit = 0                          ' number of items currently not processed because we could have storage problems
    # DeferredLimitExceeded = False
    # DeletedItem = 0
    # DeleteIndex = 0
    # DeleteNow = False
    # DftItemClass = 0
    # DftItemType = 0
    # DftItemTypeName = vbNullString
    # Diffs = vbNullString
    # DiffsIgnored = vbNullString
    # DiffsRecognized = vbNullString
    # displayInExcel = False
    # DontCompareListDefault = vbNullString
    # dontIncrementCallDepth = False
    # dontLog = False
    # BugStateElapsed = 0
    # eActFolderChoice = False                   ' SelectorMode 1
    # eAllFoldersOfType = False                  ' SelectorMode 2
    # ElapsedTime = 0
    # EntryPointCtr = 0
    # eOnlySelectedFolder = False                ' Selectormode 4
    # eOnlySelectedItems = False                 ' Selectormode 3
    # EPCalled = False                           ' called via ***Entry Point***
    # ErgebnisseAlsListe = False
    # ErrStatusHide = False
    # ErrDisplayModify = False
    # ErrorCaught = 0                            ' raw error last set by N_OnError
    # ErrStackProcs = 0
    # ErrStatusFormUsable = False
    # EventHappened = False
    # ExceptionProcessing = False                ' processing exception of recursion pattern, do not recurse this
    # ExtendedAttributeList = vbNullString       ' Properties not used for Sorting
    # FavNoLogCtr = 0
    # FilterCriteriaString = vbNullString        ' used for display filters
    # FindMatchingItems = False
    # FirstAddPos = 0
    # FldCnt = 0
    # FolderLoopIndex = 0
    # FolderPathLevel = 0
    # HeadlineName = vbNullString
    # IgnoredPropertyComparisons = 0
    # IgnoreUnhandledError = False               ' Inside some routines, do not log bad behaviour
    # IgString = vbNullString
    # InlineZ_ProcLog = vbNullString
    # InOrOut = vbNullString
    # iRuleBits = vbNullString                   ' ??? for debugging only
    # IsComparemode = False
    # isNonLoopFolder = False                    ' the real result of NonLoopFolder without user interaction
    # isSpecialName = False                      ' key with #B
    # isSQL = False
    # isUserProperty = False                     ' key with #W2
    # ItemInIMAPFolder = False
    # ItemsToDoCount = 0
    # ItsAPhoneNumber = False
    # killMsg = vbNullString
    # killStringMsg = vbNullString
    # killType = vbNullString
    # LastAddPos = 0
    # lHeadLine = 0
    # LineToAdd = vbNullString
    # ListCount = 0
    # LoeschbesttigungCaption = vbNullString
    # LogAllErrors = False
    # LogicTrace = vbNullString
    # LogImmediate = True
    # T_DC.LogName = vbNullString
    # T_DC.LogNameNext = vbNullString
    # T_DC.LogNamePrev = vbNullString
    # LogSelection = vbNullString
    # ' not wanted: LogZProcs = False
    # loopItmIndex = 0
    # LListe = vbNullString
    # MailModified = False
    # MainObjectIdentification = vbNullString    ' the one we use
    # MainProfileAccount = vbNullString
    # MaintenanceAction = 0
    # Matchcode = 0
    # MatchData = vbNullString
    # Matches = 0
    # MatchMin = 0
    # MaxMisMatchesForCandidates = 0
    # MaxPropertyCount = 0
    # MayChangeErr = True
    # Message = vbNullString
    # MinimalLogging = 0
    # MostImportantAttributes = vbNullString
    # MPEItemDiffs = vbNullString
    # mustDecodeRest = False
    # NoEventOnAddItem = False
    # NoInboxRecursion = False
    # NormalizedPhoneNumber = vbNullString
    # NoSentRecursion = False
    # NotDecodedProperties = 0
    # NoTimerEvent = False
    # objTypName = vbNullString
    # OffAdI = 0
    # OffCal = 0
    # OffLvl = 0
    # OffMTS = 0
    # OffObj = 0
    # OffPrN = 0
    # OffTim = 0
    # OneDiff = vbNullString
    # OneDiff_qualifier = vbNullString
    # OnlyMostImportantProperties = False        ' current rule, derived from user-specified
    # onlyNew = False                            ' true: do not look for existing property names
    # Output = vbNullString
    # PhoneNumberNormalized = False
    # PickTopFolder = False
    # PropertyNameX = vbNullString
    # PropertyNameY = vbNullString
    # PropertyX = vbNullString
    # PropertyY = vbNullString
    # PropStatesString = vbNullString
    # ProtectStackState = inv
    # quickChecksOnly = False                    ' rule specified by User dialog
    # relevantIndex = 0
    # RestrictCriteriaString = vbNullString      ' Selection Parameters
    # rID = 0
    # rPTrueIndex = 0
    # rsp = 0
    # SaveAttachments = False
    # saveItemNotAllowed = False
    # SearchCriteriaString = vbNullString
    # SearchFolderNameResult = vbNullString
    # SelectedAttributes = vbNullString          ' Attribute names which are MustMatch or Similar
    # SelectMulti = False
    # SelectOnlyOne = False
    # SelectorMode = 0                           ' selected SelectorButton x
    # sHdl = vbNullString
    # ShowEmptyAttributes = False
    # ShowFunctionValues = False
    # ShutUpMode = False
    # SIC = 0
    # SimilarityCount = 0
    # SkipDontCompare = False
    # SkipedEventsCounter = 0                    ' count number of times we skipped "add" action
    # SkipNextInteraction = False
    # SortMatches = vbNullString
    # SpecialObjectNameAddition = vbNullString   ' Name extended by this string for px>2 (#E, #O)
    # ExLiveDscGen = False
    # StackDebug = 0
    # StackDebugOverride = 0                     ' override first getDebugMode Value
    # StopLoop = False
    # StopRecursionNonLogged = False
    # stpcnt = 0
    # StringMod = False                          ' side-result of AppendTo: and Remove: true if change occurred
    # StringsRemoved = vbNullString
    # SuperRelevantMisMatch = False
    # SuppressStatusFormUpdate = False
    # S_ActModule = vbNullString
    # S_AppIndex = inv                           ' D_AppStack.Count - 1
    # S_AppKey = vbNullString
    # S_DbgId = vbNullString
    # TestCriteriaEditing = 0
    # testNonValueProperties = vbNullString
    # TotalPropertyCount = 0                     ' count of properties in decoded item(s). must be same for two compared items
    # TraceMode = False
    # TraceTop = 0
    # TrashFolderPath = vbNullString
    # TrueCritList = vbNullString
    # UInxDeferred = 0
    # UInxDeferredIsValid = False
    # UI_DontUseDel = False                      ' request override of frmDelParms
    # UI_DontUse_Sel = False                     ' request override of SelectionParameter
    # UI_Show_Del = False                        ' allow user to choose override
    # UI_Show_Sel = False                        ' allow user to chose from SelectionParameter
    # UserDecisionEffective = False
    # UserDecisionRequest = False
    # UTCisUsed = False
    # WantConfirmation = False
    # WantConfirmationThisFolder = False
    # workingOnNonspecifiedItem = False          ' specified aID(px + 2) is exception or occurrence
    # xDeferExcel = False                        ' do not create excel table until required
    # XlOpenedHere = False
    # xNoColAdjust = False
    # xReportExcel = False
    # xUseExcel = False
    # YleadsXby = 0

    # ' Objects dynamically used
    # Set aObjDsc = Nothing
    # Set aAccountDsc = Nothing
    # Set aCell = Nothing
    # Set ActItemObject = Nothing
    # Set aNonModalForm = Nothing
    # Set aItemList = Nothing
    # Set aNameSpace = Nothing
    # Set aProp = Nothing
    # Set aProps = Nothing
    # Set aRDOSession = Nothing
    # Set aStore = Nothing
    # Set aTD = Nothing
    # Set olApp = Nothing
    # Set CalendarFolder = Nothing
    # Set ChosenTargetFolder = Nothing
    # Set ClassDescriptorItems = Nothing
    # Set oStores = Nothing
    # Set CurIterationSwitches = Nothing
    # Set D_LiveStack = Nothing
    # Set aNonModalForm = Nothing
    # Set DeletionCandidates = Nothing
    # Set dftRule = Nothing
    # Set FInFolderColl = Nothing
    # Set FolderBackup = Nothing
    # Set FolderInbox = Nothing
    # Set FolderSent = Nothing
    # Set FolderSMS = Nothing
    # Set FolderTasks = Nothing
    # Set FolderUnknown = Nothing
    # Set FRM = Nothing
    # Set FWP_LBF_Hdl = Nothing
    # Set FWP_xLW_Hdl = Nothing
    # Set FWP_xMainWinHdl = Nothing
    # Set iRules = Nothing
    # Set killWords = Nothing
    # Set LBF = Nothing
    # Set ListContent = Nothing
    # Set logItem = Nothing
    # Set LoggableFolders = New Dictionary
    # Set LookupFolders = Nothing
    # Set MainFolderContacts = Nothing
    # Set MainFolderInbox = Nothing
    # Set MainFolderSent = Nothing
    # Set MSR = Nothing
    # Set OlBackupHome = Nothing
    # Set OlGooMailHome = Nothing
    # Set OlHotMailHome = Nothing
    # Set OlWEBmailHome = Nothing
    # Set oSearchFolders = Nothing
    # Set D_TC = Nothing
    # Set ParentFolder = Nothing
    # Set QModeNames = Nothing
    # Set RestrictedItemCollection = Nothing
    # Set RestrictedItems = Nothing
    # Set sDictionary = Nothing
    # Set SelectedItems = Nothing
    # Set SelectedObjects = Nothing
    # Set SessionAccounts = Nothing
    # Set SQLpropC = Nothing
    # Set sRules = Nothing
    # Set topFolder = Nothing
    # Set TopFolders = Nothing
    # Set TrashFolder = Nothing
    # Set Deferred = Nothing
    # Set UserRule = Nothing
    # Set WithProcVal = Nothing
    # Set xlApp = Nothing

    # ' Variant Objects

    # Set AccountTypeNames = Nothing
    # Set MostImportantProperties = Nothing
    # Set TrueImportantProperties = Nothing
    # Set xColList = Nothing

    # ' Other variants not meant as constants
    # Erase ActiveExplorerItem
    # Erase aDecProp
    # Erase aID
    # Erase AttributeUndef
    # Erase delSource
    # Erase Fctr
    # Erase fiBody
    # Erase fiMain
    # Erase Folder
    # Erase FullFolderPath
    # Erase Ictr
    # Erase WorkIndex
    # Erase WorkIndex
    # Erase WorkItemMod
    # Erase lFolders
    # Erase MatchPoints
    # Erase pArr
    # Erase rP
    # Erase sortedItems
    # Erase DeferredFolder

    # ' BugHelp Interfaces that need to be New

    # Set ExternCaller = New cProcItem           ' simulated ProcDsc for external caller | K:Extern.Caller R:Extern.Caller
    # Set C_AllPurposeStack = New Collection     ' any class of objects in this
    # Set C_CallTrace = New Collection           ' Collectionof cErr for tracing calls
    # Set C_ProtectedStack = New Collection      ' see N_Suppress for example
    # Set C_PushLogStack = New Collection        ' Collectionof Strings meant for logging
    # Set D_AccountDscs = Nothing
    # Set D_ErrInterface = New Dictionary        ' Dictionary of all defined cProcItem s
    # Set D_ErrInterface = New Dictionary        ' the entire set of Procs, names, Keys, types, etc.
    # Set D_LiveStack = New Dictionary           ' when accessing LiveStack, this is the current setting
    # Set D_PushDict = New Dictionary            ' Dictionary of cPush objects
    # Set E_Active = New cErr                    ' Active Error Environment
    # Set E_AppErr = New cErr                    ' Current Application cErr. Need not match .ErrActive
    # Set E_AppErr = New cErr                    ' used for changing Application cErr during transitions
    # Set E_PrevCallTrace = New cErr             ' Previous cErr on C_CallTrace
    # Set P_Active = New cProcItem               ' ProcDsc for E_Active
    # Set P_CurrEntry = New cProcItem            ' the technical base of all calls (below root, above InitErrHandler) | K:CurrEntry
    # Set P_EntryPoint = New cProcItem           ' Z_AppEntry Description. .ErrActive Identifies last active (requires .atRecursionOK=True) | K:N_CallPointDsc
    # Set P_LastEP = New cProcItem               ' Previous Entry Point Proc | K:LastEP
    # Set T_DC = New cTermination

    print(Debug.Print "N_DeInit has finished")

    # FuncExit:
    # Err.Clear
    # Set hDsc = Nothing
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EditWatch
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def editwatch():
    # Dim zErr As cErr
    # Const zKey As String = "AllPublic.EditWatch"
    # Call ProcCall(zErr, zKey, Qmode:=eQrMode, CallType:=tSub, ExplainS:="")

    # Call TextEdit(Testvar)
    # Call getDebugMode(False)

    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : ShowOrHideForm
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: ShowIt or Hide as requested
# '          observe restriction that only one modal form can be in foreground
# '---------------------------------------------------------------------------------------
def showorhideform():
    # '''' Proc Must ONLY CALL Z_Type PROCS                           ' trivial proc
    # Const zKey As String = "AllPublic.ShowOrHideForm"
    # Dim explainVis As String
    # explainVis = IIf(ShowIt, "Show", "Hide")
    # Call DoCall(zKey, "Sub", eQzMode, ExplainS:="ErrStatus " & explainVis)

    # Static frmErrStatusSuppressed As Boolean

    if oForm Is Nothing Then                       ' just defaulting some:
    if ErrStatusFormUsable Then:
    # Set oForm = aNonModalForm
    elif DebugControlsUsable Then:
    # Set oForm = New frmErrStatus
    # Set aNonModalForm = oForm
    # ErrStatusFormUsable = True
    # frmErrStatus.fHideMe = False
    else:
    if oForm.Visible Then:
    if frmErrStatus Is oForm Then          ' must use change event:
    # frmErrStatus.fHideMe = Not ShowIt  ' if change event, causes hiding
    else:
    if Not ShowIt Then:
    print(Debug.Print "Form Hidden: " & oForm.Name)
    # oForm.Hide                     ' blocking form hidden now, can reshow frmErrStatus
    if Not frmErrStatus.fHideMe Then ' ... if it was supposed to be showing:
    # frmErrStatus.fHideMe = False
    else:
    if oForm Is aNonModalForm Then:
    if ShowIt Then                     ' false + not visible = initialize just before:
    if aNonModalForm Is frmErrStatus Then:
    # Call QueryErrStatusChange(False) ' .f values are changable
    # ErrDisplayModify = True    ' force ShowIt
    # frmErrStatus.fHideMe = False
    # Call frmErrStatus.UpdInfo
    else:
    # oForm.Show vbModeless      ' other form in modeless request
    else:
    if frmErrStatus Is oForm Then:
    if DebugMode Then:
    print(Debug.Print "Form will be Hidden: " _)
    # & frmErrStatus.Name _
    # & " Show other Form " & oForm.Name
    # frmErrStatus.fHideMe = True ' change causes hide
    else:
    # oForm.Hide
    else:
    if ShowIt Then:
    if Not aNonModalForm Is Nothing Then ' but a nonmodal is there:
    if aNonModalForm.Visible Then '  ... and visible:
    print(Debug.Print "Form Hidden: " _)
    # & aNonModalForm.Name _
    # & " to ShowIt " & oForm.Name
    # aNonModalForm.Hide     ' just hide, no fHideMe used at all
    if frmErrStatus Is oForm Then  ' do not reopen frmErrStatus:
    if DebugMode Then:
    print(Debug.Print "** showing modal form " & oForm.Name & " <> frmErrStatus")
    # Call oForm.Show
    else:
    # oForm.Hide
    if Not aNonModalForm Is Nothing Then:
    if aNonModalForm.Visible Then:
    print(Debug.Print "Form to reshow Show: " & aNonModalForm.Name)
    if aNonModalForm Is frmErrStatus Then:
    # frmErrStatus.fHideMe = False
    else:
    # aNonModalForm.Show vbModeless

    if DebugControlsUsable Then:
    # ErrStatusFormUsable = Not aNonModalForm Is Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub InitApp
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initapp():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.InitApp"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")
    # ' do not forget to Call N_ReturnProc(zErr)

    # IsEntryPoint = False
    # DoVerify False, "*** this is not tested yet ???"


    # Call LogEvent("Explicit application Re-init via InitApp - Macro", eLall)
    # E_AppErr.Explanations = vbNullString

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub LogLevelChecks
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def loglevelchecks():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.LogLevelChecks"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")
    # ' do not forget to Call N_ReturnProc(zErr)

    # With meForm
    # MinimalLogging = .LpLogLevel.ListIndex
    # LogSelection = .LpLogLevel.Value
    if .LpLogLevel.ListIndex = eLdebug Then:
    # Call SetDebugModes(volltest:=True)
    elif .LpLogLevel.ListIndex = eLnothing Then:
    # Call SetDebugMode(False, True)
    # End With                                       ' meForm

    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub LPlogLevel_define
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def lploglevel_define():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.LPlogLevel_define"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")
    # ' do not forget to Call N_ReturnProc(zErr)

    # Dim i As Long
    if LenB(LogSelection) = 0 Then:
    # LogSelection = LogLevelNames(eLmin)
    # With meForm
    # .LpLogLevel.Clear
    # .LpLogLevel.addItem LogLevelNames(i), i
    if LogLevelNames(i) = LogSelection Then:
    # MinimalLogging = i
    # End With                                       ' meForm

    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub testML
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def testml():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.testML"
    # Static zErr As New cErr
    # Dim ReShowFrmErrStatus As Boolean


    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="AllPublic")

    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
    # ReShowFrmErrStatus = True
    # Set FRM = New frmMacroSelRun

    # Call ShowOrHideForm(FRM, ShowIt:=True)
    # Set FRM = Nothing

    if ReShowFrmErrStatus Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub MacroCall
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def macrocall():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.MacroCall"
    # Dim zErr As cErr
    # Dim ReShowFrmErrStatus As Boolean

    # Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSubEP, ExplainS:="")

    if frmErrStatus.Visible Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=False)
    # ReShowFrmErrStatus = True

    # Set MSR = New frmMacroSelRun
    # MSR.Show

    # FuncExit:
    # Set MSR = Nothing

    if ReShowFrmErrStatus Then:
    # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)

    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SetStaticActionTitles
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setstaticactiontitles():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "AllPublic.SetStaticActionTitles"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")
    # ' do not forget to Call N_ReturnProc(zErr)

    # ActionTitle(0) = vbNullString
    # ActionTitle(1) = "Default-Aktion"
    # ActionTitle(2) = "Kategorie der Mail bestimmen"
    # ActionTitle(3) = "Post-Eingangsbearbeitung  durchfhren"
    # ActionTitle(4) = "Doppelte Items lschen"
    # ActionTitle(5) = "Normalreprsentation erzwingen"
    # ActionTitle(6) = "Ordnerinhalte Zusammenfhren"
    # ActionTitle(7) = "Finde alle 'NLOGGED' - Suchordner"
    # ActionTitle(8) = "Bearbeite alle bereinstimmungen zu einer Suche"

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub Restart
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def restart():
    # Call Quit(False)
    if StartUpAgain Then:
    # Call StartUp

# '---------------------------------------------------------------------------------------
# ' Method : Sub Quit
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def quit():
    # LogImmediate = True
    # QuitStarted = True                             ' disable any Entry/Exit
    # Call LogEvent("at Quit Entry", eLall)
    # Call Stop_BugTimer                             ' calling Friend
    # Call LogEvent("dStop_BugTimer", eLall)
    # Call ShowLiveStack(True)
    # Call LogEvent("done show livestack", eLall)
    # Call ErrEx.Disable
    # Call LogEvent("ErrEx disabled: done Quit Entry", eLall)
    if DoEnd Then:
    # Call CloseLog(KeepName:=True, msg:="Quit Completed")
    # End
    else:
    # Call N_DeInit
    # Call CloseLog(KeepName:=True, msg:="Completed and all Erased, Not Ending")
    # QuitStarted = False

# '---------------------------------------------------------------------------------------
# ' Method : Function TerminateRun
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def terminaterun():

    # Const zKey As String = "AllPublic.TerminateRun"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True
    # Call DoCall(zKey, tFunction, eQzMode)


    # SkipedEventsCounter = 0
    # NoEventOnAddItem = False
    # StopRecursionNonLogged = False

    if DebugMode Or withStop Then:
    # rsp = vbNo                                 ' if user just closes, like vbNo
    # & "ignore (=Yes), or do NOT Terminate (=No)?" _
    # & vbCrLf & "or DO TERMINATE RUN (Cancel)", _
    # vbYesNoCancel, "Debug Debug.Assert False")
    if rsp = vbYes Then:
    print(Debug.Print "Termination ON APP LEVEL attempted by User.")
    print(Debug.Print "Error Condition Cleared and invoking App should continue.")
    # DoVerify False
    # Call T_DC.N_ClearTermination
    # Call TerminateApp(False)
    elif rsp = vbNo Then:
    # Call T_DC.N_ClearTermination
    print(Debug.Print "* TERMINATION CANCELLED BY USER. Error Condition Cleared")
    # GoTo FuncExit
    elif rsp = vbCancel Then:
    print(Debug.Print "TERMINATION performed by User")
    # T_DC.TermRQ = True
    # End                                    ' Debug.Assert False dead

    # Call CloseLog
    # Set Deferred = New Collection
    # Set sortedItems(1) = Nothing
    # Set sortedItems(2) = Nothing
    # Set SelectedItems = Nothing
    # Set aCell = Nothing

    # FuncExit:
    # Recursive = False
    # Call ErrReset(0)

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : ShowStatusUpdate
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Do Events and Show Status Update in ErrStatusForm if open
# '---------------------------------------------------------------------------------------
def showstatusupdate():
    # Const zKey As String = "AllPublic.ShowStatusUpdate"
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug > 8 And Not SuppressStatusFormUpdate Then:
    print(Debug.Print String(OffCal, b) & zKey & " ignored, recursing from " _)
    # & P_Active.DbgId
    # GoTo ProcRet
    # Recursive = True

    if Not SuppressStatusFormUpdate Then:
    if ErrStatusFormUsable Then:
    # Call N_Suppress(Push, zKey)
    if Not NoEventOnAddItem Then:
    # Call doMyEvents
    if frmErrStatus.Visible Then           ' object only similar to frmErrStatus:
    # Call QueryErrStatusChange(True)    ' update the StatusForm Information
    # Call N_Suppress(Pop, zKey)
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub doMyEvents
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def domyevents():
    # Const zKey As String = "AllPublic.doMyEvents"

    # Static maxDoTime As Double
    # Dim sBlockEvents As Boolean
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # sBlockEvents = E_Active.EventBlock
    if E_Active.EventBlock Or NoEventOnAddItem Then:
    if DebugLogging Then:
    print(Debug.Print "DoMyEvents skipped due to Event Block")
    # GoTo ProcRet
    if DebugLogging And LogZProcs And DebugMode And Not LogPerformance Then:
    # E_Active.EventBlock = True                 ' do not recurse events
    # BugTimer.BugStateLast = Timer
    # DoEvents
    # BugTimer.BugStateElapsed = Timer - BugTimer.BugStateLast
    if BugTimer.BugStateElapsed > maxDoTime Then:
    # maxDoTime = BugTimer.BugStateElapsed
    print(Debug.Print "DoEvents took " & RString(BugTimer.BugStateElapsed, 15) _)
    # & " ticks, maximum is " & RString(maxDoTime, 15)
    # E_Active.EventBlock = sBlockEvents         ' restore state at call of doMyEvents
    else:
    # E_Active.EventBlock = True                 ' do not recurse events
    # DoEvents
    # E_Active.EventBlock = sBlockEvents         ' restore state at call of doMyEvents
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call BugEval

    # Recursive = False
    # ProcRet:


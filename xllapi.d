module xllapi;
import xlltypes;
/** 
 This module containts the C api exported from execel
 */
private { 
 	enum xlCommand = 0x8000;
	enum xlSpecial = 0x4000;
	enum xlIntl = 0x2000;
	enum xlPrompt = 0x1000;
	
//	extern(C) int Excel4f(XlFn xlfn, Xlxoper4 *operRes, int count,... );
	extern(C) int Excel12f(XlFn xlfn, Xloper12 *operRes, int count,... );
}

enum XlFn {
	xlFree = (0 | xlSpecial),
	xlStack = (1 | xlSpecial),
	xlCoerce = (2 | xlSpecial),
	xlSet = (3 | xlSpecial),
	xlSheetId = (4 | xlSpecial),
	xlSheetNm = (5 | xlSpecial),
	xlAbort = (6 | xlSpecial),
	xlGetInst = (7 | xlSpecial), /** Returns application's hinstance as an integer value, supported on 32-bit platform only */
	xlGetHwnd = (8 | xlSpecial),
	xlGetName = (9 | xlSpecial),
	xlEnableXLMsgs = (10 | xlSpecial),
	xlDisableXLMsgs = (11 | xlSpecial),
	xlDefineBinaryName = (12 | xlSpecial),
	xlGetBinaryName = (13 | xlSpecial), /** GetFooInfo are valid only for calls to LPenHelper */
	xlGetFmlaInfo = (14 | xlSpecial),
	xlGetMouseInfo = (15 | xlSpecial),
	xlAsyncReturn = (16 | xlSpecial),	/** Set return value from an asynchronous function call*/
	xlEventRegister = (17 | xlSpecial),	/** Register an XLL event*/
	xlRunningOnCluster = (18 | xlSpecial),	/** Returns true if running on Compute Cluster*/
	xlGetInstPtr = (19 | xlSpecial),	/** Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */

	xlcBeep = (0 | xlCommand),
	xlcOpen = (1 | xlCommand),
	xlcOpenLinks = (2 | xlCommand),
	xlcCloseAll = (3 | xlCommand),
	xlcSave = (4 | xlCommand),
	xlcSaveAs = (5 | xlCommand),
	xlcFileDelete = (6 | xlCommand),
	xlcPageSetup = (7 | xlCommand),
	xlcPrint = (8 | xlCommand),
	xlcPrinterSetup = (9 | xlCommand),
	xlcQuit = (10 | xlCommand),
	xlcNewWindow = (11 | xlCommand),
	xlcArrangeAll = (12 | xlCommand),
	xlcWindowSize = (13 | xlCommand),
	xlcWindowMove = (14 | xlCommand),
	xlcFull = (15 | xlCommand),
	xlcClose = (16 | xlCommand),
	xlcRun = (17 | xlCommand),
	xlcSetPrintArea = (22 | xlCommand),
	xlcSetPrintTitles = (23 | xlCommand),
	xlcSetPageBreak = (24 | xlCommand),
	xlcRemovePageBreak = (25 | xlCommand),
	xlcFont = (26 | xlCommand),
	xlcDisplay = (27 | xlCommand),
	xlcProtectDocument = (28 | xlCommand),
	xlcPrecision = (29 | xlCommand),
	xlcA1R1c1 = (30 | xlCommand),
	xlcCalculateNow = (31 | xlCommand),
	xlcCalculation = (32 | xlCommand),
	xlcDataFind = (34 | xlCommand),
	xlcExtract = (35 | xlCommand),
	xlcDataDelete = (36 | xlCommand),
	xlcSetDatabase = (37 | xlCommand),
	xlcSetCriteria = (38 | xlCommand),
	xlcSort = (39 | xlCommand),
	xlcDataSeries = (40 | xlCommand),
	xlcTable = (41 | xlCommand),
	xlcFormatNumber = (42 | xlCommand),
	xlcAlignment = (43 | xlCommand),
	xlcStyle = (44 | xlCommand),
	xlcBorder = (45 | xlCommand),
	xlcCellProtection = (46 | xlCommand),
	xlcColumnWidth = (47 | xlCommand),
	xlcUndo = (48 | xlCommand),
	xlcCut = (49 | xlCommand),
	xlcCopy = (50 | xlCommand),
	xlcPaste = (51 | xlCommand),
	xlcClear = (52 | xlCommand),
	xlcPasteSpecial = (53 | xlCommand),
	xlcEditDelete = (54 | xlCommand),
	xlcInsert = (55 | xlCommand),
	xlcFillRight = (56 | xlCommand),
	xlcFillDown = (57 | xlCommand),
	xlcDefineName = (61 | xlCommand),
	xlcCreateNames = (62 | xlCommand),
	xlcFormulaGoto = (63 | xlCommand),
	xlcFormulaFind = (64 | xlCommand),
	xlcSelectLastCell = (65 | xlCommand),
	xlcShowActiveCell = (66 | xlCommand),
	xlcGalleryArea = (67 | xlCommand),
	xlcGalleryBar = (68 | xlCommand),
	xlcGalleryColumn = (69 | xlCommand),
	xlcGalleryLine = (70 | xlCommand),
	xlcGalleryPie = (71 | xlCommand),
	xlcGalleryScatter = (72 | xlCommand),
	xlcCombination = (73 | xlCommand),
	xlcPreferred = (74 | xlCommand),
	xlcAddOverlay = (75 | xlCommand),
	xlcGridlines = (76 | xlCommand),
	xlcSetPreferred = (77 | xlCommand),
	xlcAxes = (78 | xlCommand),
	xlcLegend = (79 | xlCommand),
	xlcAttachText = (80 | xlCommand),
	xlcAddArrow = (81 | xlCommand),
	xlcSelectChart = (82 | xlCommand),
	xlcSelectPlotArea = (83 | xlCommand),
	xlcPatterns = (84 | xlCommand),
	xlcMainChart = (85 | xlCommand),
	xlcOverlay = (86 | xlCommand),
	xlcScale = (87 | xlCommand),
	xlcFormatLegend = (88 | xlCommand),
	xlcFormatText = (89 | xlCommand),
	xlcEditRepeat = (90 | xlCommand),
	xlcParse = (91 | xlCommand),
	xlcJustify = (92 | xlCommand),
	xlcHide = (93 | xlCommand),
	xlcUnhide = (94 | xlCommand),
	xlcWorkspace = (95 | xlCommand),
	xlcFormula = (96 | xlCommand),
	xlcFormulaFill = (97 | xlCommand),
	xlcFormulaArray = (98 | xlCommand),
	xlcDataFindNext = (99 | xlCommand),
	xlcDataFindPrev = (100 | xlCommand),
	xlcFormulaFindNext = (101 | xlCommand),
	xlcFormulaFindPrev = (102 | xlCommand),
	xlcActivate = (103 | xlCommand),
	xlcActivateNext = (104 | xlCommand),
	xlcActivatePrev = (105 | xlCommand),
	xlcUnlockedNext = (106 | xlCommand),
	xlcUnlockedPrev = (107 | xlCommand),
	xlcCopyPicture = (108 | xlCommand),
	xlcSelect = (109 | xlCommand),
	xlcDeleteName = (110 | xlCommand),
	xlcDeleteFormat = (111 | xlCommand),
	xlcVline = (112 | xlCommand),
	xlcHline = (113 | xlCommand),
	xlcVpage = (114 | xlCommand),
	xlcHpage = (115 | xlCommand),
	xlcVscroll = (116 | xlCommand),
	xlcHscroll = (117 | xlCommand),
	xlcAlert = (118 | xlCommand),
	xlcNew = (119 | xlCommand),
	xlcCancelCopy = (120 | xlCommand),
	xlcShowClipboard = (121 | xlCommand),
	xlcMessage = (122 | xlCommand),
	xlcPasteLink = (124 | xlCommand),
	xlcAppActivate = (125 | xlCommand),
	xlcDeleteArrow = (126 | xlCommand),
	xlcRowHeight = (127 | xlCommand),
	xlcFormatMove = (128 | xlCommand),
	xlcFormatSize = (129 | xlCommand),
	xlcFormulaReplace = (130 | xlCommand),
	xlcSendKeys = (131 | xlCommand),
	xlcSelectSpecial = (132 | xlCommand),
	xlcApplyNames = (133 | xlCommand),
	xlcReplaceFont = (134 | xlCommand),
	xlcFreezePanes = (135 | xlCommand),
	xlcShowInfo = (136 | xlCommand),
	xlcSplit = (137 | xlCommand),
	xlcOnWindow = (138 | xlCommand),
	xlcOnData = (139 | xlCommand),
	xlcDisableInput = (140 | xlCommand),
	xlcEcho = (141 | xlCommand),
	xlcOutline = (142 | xlCommand),
	xlcListNames = (143 | xlCommand),
	xlcFileClose = (144 | xlCommand),
	xlcSaveWorkbook = (145 | xlCommand),
	xlcDataForm = (146 | xlCommand),
	xlcCopyChart = (147 | xlCommand),
	xlcOnTime = (148 | xlCommand),
	xlcWait = (149 | xlCommand),
	xlcFormatFont = (150 | xlCommand),
	xlcFillUp = (151 | xlCommand),
	xlcFillLeft = (152 | xlCommand),
	xlcDeleteOverlay = (153 | xlCommand),
	xlcNote = (154 | xlCommand),
	xlcShortMenus = (155 | xlCommand),
	xlcSetUpdateStatus = (159 | xlCommand),
	xlcColorPalette = (161 | xlCommand),
	xlcDeleteStyle = (162 | xlCommand),
	xlcWindowRestore = (163 | xlCommand),
	xlcWindowMaximize = (164 | xlCommand),
	xlcError = (165 | xlCommand),
	xlcChangeLink = (166 | xlCommand),
	xlcCalculateDocument = (167 | xlCommand),
	xlcOnKey = (168 | xlCommand),
	xlcAppRestore = (169 | xlCommand),
	xlcAppMove = (170 | xlCommand),
	xlcAppSize = (171 | xlCommand),
	xlcAppMinimize = (172 | xlCommand),
	xlcAppMaximize = (173 | xlCommand),
	xlcBringToFront = (174 | xlCommand),
	xlcSendToBack = (175 | xlCommand),
	xlcMainChartType = (185 | xlCommand),
	xlcOverlayChartType = (186 | xlCommand),
	xlcSelectEnd = (187 | xlCommand),
	xlcOpenMail = (188 | xlCommand),
	xlcSendMail = (189 | xlCommand),
	xlcStandardFont = (190 | xlCommand),
	xlcConsolidate = (191 | xlCommand),
	xlcSortSpecial = (192 | xlCommand),
	xlcGallery3dArea = (193 | xlCommand),
	xlcGallery3dColumn = (194 | xlCommand),
	xlcGallery3dLine = (195 | xlCommand),
	xlcGallery3dPie = (196 | xlCommand),
	xlcView3d = (197 | xlCommand),
	xlcGoalSeek = (198 | xlCommand),
	xlcWorkgroup = (199 | xlCommand),
	xlcFillGroup = (200 | xlCommand),
	xlcUpdateLink = (201 | xlCommand),
	xlcPromote = (202 | xlCommand),
	xlcDemote = (203 | xlCommand),
	xlcShowDetail = (204 | xlCommand),
	xlcUngroup = (206 | xlCommand),
	xlcObjectProperties = (207 | xlCommand),
	xlcSaveNewObject = (208 | xlCommand),
	xlcShare = (209 | xlCommand),
	xlcShareName = (210 | xlCommand),
	xlcDuplicate = (211 | xlCommand),
	xlcApplyStyle = (212 | xlCommand),
	xlcAssignToObject = (213 | xlCommand),
	xlcObjectProtection = (214 | xlCommand),
	xlcHideObject = (215 | xlCommand),
	xlcSetExtract = (216 | xlCommand),
	xlcCreatePublisher = (217 | xlCommand),
	xlcSubscribeTo = (218 | xlCommand),
	xlcAttributes = (219 | xlCommand),
	xlcShowToolbar = (220 | xlCommand),
	xlcPrintPreview = (222 | xlCommand),
	xlcEditColor = (223 | xlCommand),
	xlcShowLevels = (224 | xlCommand),
	xlcFormatMain = (225 | xlCommand),
	xlcFormatOverlay = (226 | xlCommand),
	xlcOnRecalc = (227 | xlCommand),
	xlcEditSeries = (228 | xlCommand),
	xlcDefineStyle = (229 | xlCommand),
	xlcLinePrint = (240 | xlCommand),
	xlcEnterData = (243 | xlCommand),
	xlcGalleryRadar = (249 | xlCommand),
	xlcMergeStyles = (250 | xlCommand),
	xlcEditionOptions = (251 | xlCommand),
	xlcPastePicture = (252 | xlCommand),
	xlcPastePictureLink = (253 | xlCommand),
	xlcSpelling = (254 | xlCommand),
	xlcZoom = (256 | xlCommand),
	xlcResume = (258 | xlCommand),
	xlcInsertObject = (259 | xlCommand),
	xlcWindowMinimize = (260 | xlCommand),
	xlcSize = (261 | xlCommand),
	xlcMove = (262 | xlCommand),
	xlcSoundNote = (265 | xlCommand),
	xlcSoundPlay = (266 | xlCommand),
	xlcFormatShape = (267 | xlCommand),
	xlcExtendPolygon = (268 | xlCommand),
	xlcFormatAuto = (269 | xlCommand),
	xlcGallery3dBar = (272 | xlCommand),
	xlcGallery3dSurface = (273 | xlCommand),
	xlcFillAuto = (274 | xlCommand),
	xlcCustomizeToolbar = (276 | xlCommand),
	xlcAddTool = (277 | xlCommand),
	xlcEditObject = (278 | xlCommand),
	xlcOnDoubleclick = (279 | xlCommand),
	xlcOnEntry = (280 | xlCommand),
	xlcWorkbookAdd = (281 | xlCommand),
	xlcWorkbookMove = (282 | xlCommand),
	xlcWorkbookCopy = (283 | xlCommand),
	xlcWorkbookOptions = (284 | xlCommand),
	xlcSaveWorkspace = (285 | xlCommand),
	xlcChartWizard = (288 | xlCommand),
	xlcDeleteTool = (289 | xlCommand),
	xlcMoveTool = (290 | xlCommand),
	xlcWorkbookSelect = (291 | xlCommand),
	xlcWorkbookActivate = (292 | xlCommand),
	xlcAssignToTool = (293 | xlCommand),
	xlcCopyTool = (295 | xlCommand),
	xlcResetTool = (296 | xlCommand),
	xlcConstrainNumeric = (297 | xlCommand),
	xlcPasteTool = (298 | xlCommand),
	xlcPlacement = (300 | xlCommand),
	xlcFillWorkgroup = (301 | xlCommand),
	xlcWorkbookNew = (302 | xlCommand),
	xlcScenarioCells = (305 | xlCommand),
	xlcScenarioDelete = (306 | xlCommand),
	xlcScenarioAdd = (307 | xlCommand),
	xlcScenarioEdit = (308 | xlCommand),
	xlcScenarioShow = (309 | xlCommand),
	xlcScenarioShowNext = (310 | xlCommand),
	xlcScenarioSummary = (311 | xlCommand),
	xlcPivotTableWizard = (312 | xlCommand),
	xlcPivotFieldProperties = (313 | xlCommand),
	xlcPivotField = (314 | xlCommand),
	xlcPivotItem = (315 | xlCommand),
	xlcPivotAddFields = (316 | xlCommand),
	xlcOptionsCalculation = (318 | xlCommand),
	xlcOptionsEdit = (319 | xlCommand),
	xlcOptionsView = (320 | xlCommand),
	xlcAddinManager = (321 | xlCommand),
	xlcMenuEditor = (322 | xlCommand),
	xlcAttachToolbars = (323 | xlCommand),
	xlcVbaactivate = (324 | xlCommand),
	xlcOptionsChart = (325 | xlCommand),
	xlcVbaInsertFile = (328 | xlCommand),
	xlcVbaProcedureDefinition = (330 | xlCommand),
	xlcRoutingSlip = (336 | xlCommand),
	xlcRouteDocument = (338 | xlCommand),
	xlcMailLogon = (339 | xlCommand),
	xlcInsertPicture = (342 | xlCommand),
	xlcEditTool = (343 | xlCommand),
	xlcGalleryDoughnut = (344 | xlCommand),
	xlcChartTrend = (350 | xlCommand),
	xlcPivotItemProperties = (352 | xlCommand),
	xlcWorkbookInsert = (354 | xlCommand),
	xlcOptionsTransition = (355 | xlCommand),
	xlcOptionsGeneral = (356 | xlCommand),
	xlcFilterAdvanced = (370 | xlCommand),
	xlcMailAddMailer = (373 | xlCommand),
	xlcMailDeleteMailer = (374 | xlCommand),
	xlcMailReply = (375 | xlCommand),
	xlcMailReplyAll = (376 | xlCommand),
	xlcMailForward = (377 | xlCommand),
	xlcMailNextLetter = (378 | xlCommand),
	xlcDataLabel = (379 | xlCommand),
	xlcInsertTitle = (380 | xlCommand),
	xlcFontProperties = (381 | xlCommand),
	xlcMacroOptions = (382 | xlCommand),
	xlcWorkbookHide = (383 | xlCommand),
	xlcWorkbookUnhide = (384 | xlCommand),
	xlcWorkbookDelete = (385 | xlCommand),
	xlcWorkbookName = (386 | xlCommand),
	xlcGalleryCustom = (388 | xlCommand),
	xlcAddChartAutoformat = (390 | xlCommand),
	xlcDeleteChartAutoformat = (391 | xlCommand),
	xlcChartAddData = (392 | xlCommand),
	xlcAutoOutline = (393 | xlCommand),
	xlcTabOrder = (394 | xlCommand),
	xlcShowDialog = (395 | xlCommand),
	xlcSelectAll = (396 | xlCommand),
	xlcUngroupSheets = (397 | xlCommand),
	xlcSubtotalCreate = (398 | xlCommand),
	xlcSubtotalRemove = (399 | xlCommand),
	xlcRenameObject = (400 | xlCommand),
	xlcWorkbookScroll = (412 | xlCommand),
	xlcWorkbookNext = (413 | xlCommand),
	xlcWorkbookPrev = (414 | xlCommand),
	xlcWorkbookTabSplit = (415 | xlCommand),
	xlcFullScreen = (416 | xlCommand),
	xlcWorkbookProtect = (417 | xlCommand),
	xlcScrollbarProperties = (420 | xlCommand),
	xlcPivotShowPages = (421 | xlCommand),
	xlcTextToColumns = (422 | xlCommand),
	xlcFormatCharttype = (423 | xlCommand),
	xlcLinkFormat = (424 | xlCommand),
	xlcTracerDisplay = (425 | xlCommand),
	xlcTracerNavigate = (430 | xlCommand),
	xlcTracerClear = (431 | xlCommand),
	xlcTracerError = (432 | xlCommand),
	xlcPivotFieldGroup = (433 | xlCommand),
	xlcPivotFieldUngroup = (434 | xlCommand),
	xlcCheckboxProperties = (435 | xlCommand),
	xlcLabelProperties = (436 | xlCommand),
	xlcListboxProperties = (437 | xlCommand),
	xlcEditboxProperties = (438 | xlCommand),
	xlcPivotRefresh = (439 | xlCommand),
	xlcLinkCombo = (440 | xlCommand),
	xlcOpenText = (441 | xlCommand),
	xlcHideDialog = (442 | xlCommand),
	xlcSetDialogFocus = (443 | xlCommand),
	xlcEnableObject = (444 | xlCommand),
	xlcPushbuttonProperties = (445 | xlCommand),
	xlcSetDialogDefault = (446 | xlCommand),
	xlcFilter = (447 | xlCommand),
	xlcFilterShowAll = (448 | xlCommand),
	xlcClearOutline = (449 | xlCommand),
	xlcFunctionWizard = (450 | xlCommand),
	xlcAddListItem = (451 | xlCommand),
	xlcSetListItem = (452 | xlCommand),
	xlcRemoveListItem = (453 | xlCommand),
	xlcSelectListItem = (454 | xlCommand),
	xlcSetControlValue = (455 | xlCommand),
	xlcSaveCopyAs = (456 | xlCommand),
	xlcOptionsListsAdd = (458 | xlCommand),
	xlcOptionsListsDelete = (459 | xlCommand),
	xlcSeriesAxes = (460 | xlCommand),
	xlcSeriesX = (461 | xlCommand),
	xlcSeriesY = (462 | xlCommand),
	xlcErrorbarX = (463 | xlCommand),
	xlcErrorbarY = (464 | xlCommand),
	xlcFormatChart = (465 | xlCommand),
	xlcSeriesOrder = (466 | xlCommand),
	xlcMailLogoff = (467 | xlCommand),
	xlcClearRoutingSlip = (468 | xlCommand),
	xlcAppActivateMicrosoft = (469 | xlCommand),
	xlcMailEditMailer = (470 | xlCommand),
	xlcOnSheet = (471 | xlCommand),
	xlcStandardWidth = (472 | xlCommand),
	xlcScenarioMerge = (473 | xlCommand),
	xlcSummaryInfo = (474 | xlCommand),
	xlcFindFile = (475 | xlCommand),
	xlcActiveCellFont = (476 | xlCommand),
	xlcEnableTipwizard = (477 | xlCommand),
	xlcVbaMakeAddin = (478 | xlCommand),
	xlcInsertdatatable = (480 | xlCommand),
	xlcWorkgroupOptions = (481 | xlCommand),
	xlcMailSendMailer = (482 | xlCommand),
	xlcAutocorrect = (485 | xlCommand),
	xlcPostDocument = (489 | xlCommand),
	xlcPicklist = (491 | xlCommand),
	xlcViewShow = (493 | xlCommand),
	xlcViewDefine = (494 | xlCommand),
	xlcViewDelete = (495 | xlCommand),
	xlcSheetBackground = (509 | xlCommand),
	xlcInsertMapObject = (510 | xlCommand),
	xlcOptionsMenono = (511 | xlCommand),
	xlcNormal = (518 | xlCommand),
	xlcLayout = (519 | xlCommand),
	xlcRmPrintArea = (520 | xlCommand),
	xlcClearPrintArea = (521 | xlCommand),
	xlcAddPrintArea = (522 | xlCommand),
	xlcMoveBrk = (523 | xlCommand),
	xlcHidecurrNote = (545 | xlCommand),
	xlcHideallNotes = (546 | xlCommand),
	xlcDeleteNote = (547 | xlCommand),
	xlcTraverseNotes = (548 | xlCommand),
	xlcActivateNotes = (549 | xlCommand),
	xlcProtectRevisions = (620 | xlCommand),
	xlcUnprotectRevisions = (621 | xlCommand),
	xlcOptionsMe = (647 | xlCommand),
	xlcWebPublish = (653 | xlCommand),
	xlcNewwebquery = (667 | xlCommand),
	xlcPivotTableChart = (673 | xlCommand),
	xlcOptionsSave = (753 | xlCommand),
	xlcOptionsSpell = (755 | xlCommand),
	xlcHideallInkannots = (808 | xlCommand),
}

private {
	enum ExcelVersion {
		Excel4 = 4,
		Excel12 = 12,
	}
	
	string genArgsMixin(ExcelVersion eVer, int n) /*@internal*/ {
		if (!__ctfe) {assert(0, "This function should not be called outside of CTFE.");}
		import std.format : format;
		import std.range : iota, join;
		import std.algorithm : map;

		string argString = iota(0, n)
			.map!(i => format(q{&args[%d]}, i))
			.join(", ");

		string mxin = format(
		q{retval = Excel%df(xlfn, &result, %d, %s);},
		eVer, n, argString);

		return mxin;
	}
}
 

Xloper4 Excel4(int N)(XlFn xlfn, Xloper4[N] args) {
	Xloper4 result;
	int retval;
	mixin(genArgsMixin(ExcelVersion.Excel4, N));
	//TODO figure out meaning of retval
	// assert(retval);
	return  result; 	
}



Xloper12 Excel12(int N)(XlFn xlfn, Xloper12[N] args) {
	Xloper12 result;
	int retval;
	mixin(genArgsMixin(ExcelVersion.Excel12, N));
	pragma(msg, genArgsMixin(ExcelVersion.Excel12, N));
	// assert(retval);
	return  result; 	
}


export extern(Windows) short showVal(short val) {
	Xloper12 xStr;

	xStr = Excel12(XlFn.xlCoerce, [Xloper12(12), Xloper12(XloperType.xltypeInt)]);

	Excel12(XlFn.xlcAlert,  [xStr]);
	Excel12(XlFn.xlFree, [xStr]);

	return 1;
}
void _unittest() {
	
}
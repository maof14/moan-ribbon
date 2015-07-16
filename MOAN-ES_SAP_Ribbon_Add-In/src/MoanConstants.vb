﻿''' <summary>
''' A module to contain all the Excel constants. For use mostly in the CTemplateGenerator class. 
''' </summary>
''' <remarks></remarks>
Module MoanConstants

    Public Enum UnclassifiedConstants
        xl3DBar = -4099
        xl3DEffects1 = 13
        xl3DEffects2 = 14
        xl3DSurface = -4103
        xlAbove = 0
        xlAccounting1 = 4
        xlAccounting2 = 5
        xlAccounting4 = 17
        xlAdd = 2
        xlAll = -4104
        xlAccounting3 = 6
        xlAllExceptBorders = 7
        xlAutomatic = -4105
        xlBar = 2
        xlBelow = 1
        xlBidi = -5000
        xlBidiCalendar = 3
        xlBoth = 1
        xlBottom = -4107
        xlCascade = 7
        xlCenter = -4108
        xlCenterAcrossSelection = 7
        xlChart4 = 2
        xlChartSeries = 17
        xlChartShort = 6
        xlChartTitles = 18
        xlChecker = 9
        xlCircle = 8
        xlClassic1 = 1
        xlClassic2 = 2
        xlClassic3 = 3
        xlClosed = 3
        xlColor1 = 7
        xlColor2 = 8
        xlColor3 = 9
        xlColumn = 3
        xlCombination = -4111
        xlComplete = 4
        xlConstants = 2
        xlContents = 2
        xlContext = -5002
        xlCorner = 2
        xlCrissCross = 16
        xlCross = 4
        xlCustom = -4114
        xlDebugCodePane = 13
        xlDefaultAutoFormat = -1
        xlDesktop = 9
        xlDiamond = 2
        xlDirect = 1
        xlDistributed = -4117
        xlDivide = 5
        xlDoubleAccounting = 5
        xlDoubleClosed = 5
        xlDoubleOpen = 4
        xlDoubleQuote = 1
        xlDrawingObject = 14
        xlEntireChart = 20
        xlExcelMenus = 1
        xlExtended = 3
        xlFill = 5
        xlFirst = 0
        xlFixedValue = 1
        xlFloating = 5
        xlFormats = -4122
        xlFormula = 5
        xlFullScript = 1
        xlGeneral = 1
        xlGray16 = 17
        xlGray25 = -4124
        xlGray50 = -4125
        xlGray75 = -4126
        xlGray8 = 18
        xlGregorian = 2
        xlGrid = 15
        xlGridline = 22
        xlHigh = -4127
        xlHindiNumerals = 3
        xlIcons = 1
        xlImmediatePane = 12
        xlInside = 2
        xlInteger = 2
        xlJustify = -4130
        xlLast = 1
        xlLastCell = 11
        xlLatin = -5001
        xlLeft = -4131
        xlLeftToRight = 2
        xlLightDown = 13
        xlLightHorizontal = 11
        xlLightUp = 14
        xlLightVertical = 12
        xlList1 = 10
        xlList2 = 11
        xlList3 = 12
        xlLocalFormat1 = 15
        xlLocalFormat2 = 16
        xlLogicalCursor = 1
        xlLong = 3
        xlLotusHelp = 2
        xlLow = -4134
        xlLTR = -5003
        xlMacrosheetCell = 7
        xlManual = -4135
        xlMaximum = 2
        xlMinimum = 4
        xlMinusValues = 3
        xlMixed = 2
        xlMixedAuthorizedScript = 4
        xlMixedScript = 3
        xlModule = -4141
        xlMultiply = 4
        xlNarrow = 1
        xlNextToAxis = 4
        xlNoDocuments = 3
        xlNone = -4142
        xlNotes = -4144
        xlOff = -4146
        xlOn = 1
        xlOpaque = 3
        xlOpen = 2
        xlOutside = 3
        xlPartial = 3
        xlPartialScript = 2
        xlPercent = 2
        xlPlus = 9
        xlPlusValues = 2
        xlReference = 4
        xlRight = -4152
        xlRTL = -5004
        xlScale = 3
        xlSemiautomatic = 2
        xlSemiGray75 = 10
        xlShort = 1
        xlShowLabel = 4
        xlShowLabelAndPercent = 5
        xlShowPercent = 3
        xlShowValue = 2
        xlSimple = -4154
        xlSingle = 2
        xlSingleAccounting = 4
        xlSingleQuote = 2
        xlSquare = 1
        xlStar = 5
        xlStError = 4
        xlStrict = 2
        xlSubtract = 3
        xlSystem = 1
        xlTextBox = 16
        xlTiled = 1
        xlTitleBar = 8
        xlToolbar = 1
        xlToolbarButton = 2
        xlTop = -4160
        xlTopToBottom = 1
        xlTransparent = 2
        xlTriangle = 3
        xlVeryHidden = 2
        xlVisible = 12
        xlVisualCursor = 2
        xlWatchPane = 11
        xlWide = 3
        xlWorkbookTab = 6
        xlWorksheet4 = 1
        xlWorksheetCell = 3
        xlWorksheetShort = 5
    End Enum


    Public Enum xlAboveBelow
        XlAboveAverage = 0
        XlAboveStdDev = 4
        XlBelowAverage = 1
        XlBelowStdDev = 5
        XlEqualAboveAverage = 2
        XlEqualBelowAverage = 3
    End Enum


    Public Enum xlActionType
        xlActionTypeDrillthrough = 256
        xlActionTypeReport = 128
        xlActionTypeRowset = 16
        xlActionTypeUrl = 1
    End Enum


    Public Enum xlAllocation
        xlAutomaticAllocation = 2
        xlManualAllocation = 1
    End Enum


    Public Enum xlAllocationMethod
        xlEqualAllocation = 1
        xlWeightedAllocation = 2
    End Enum


    Public Enum xlAllocationValue
        xlAllocateIncrement = 2
        xlAllocateValue = 1
    End Enum


    Public Enum XlApplicationInternational
        xl24HourClock = 33
        xl4DigitYears = 43
        xlAlternateArraySeparator = 16
        xlColumnSeparator = 14
        xlCountryCode = 1
        xlCountrySetting = 2
        xlCurrencyBefore = 37
        xlCurrencyCode = 25
        xlCurrencyDigits = 27
        xlCurrencyLeadingZeros = 40
        xlCurrencyMinusSign = 38
        xlCurrencyNegative = 28
        xlCurrencySpaceBefore = 36
        xlCurrencyTrailingZeros = 39
        xlDateOrder = 32
        xlDateSeparator = 17
        xlDayCode = 21
        xlDayLeadingZero = 42
        xlDecimalSeparator = 3
        xlGeneralFormatName = 26
        xlHourCode = 22
        xlLeftBrace = 12
        xlLeftBracket = 10
        xlListSeparator = 5
        xlLowerCaseColumnLetter = 9
        xlLowerCaseRowLetter = 8
        xlMDY = 44
        xlMetric = 35
        xlMinuteCode = 23
        xlMonthCode = 20
        xlMonthLeadingZero = 41
        xlMonthNameChars = 30
        xlNoncurrencyDigits = 29
        xlNonEnglishFunctions = 34
        xlRightBrace = 13
        xlRightBracket = 11
        xlRowSeparator = 15
        xlSecondCode = 24
        xlThousandsSeparator = 4
        xlTimeLeadingZero = 45
        xlTimeSeparator = 18
        xlUpperCaseColumnLetter = 7
        xlUpperCaseRowLetter = 6
        xlWeekdayNameChars = 31
        xlYearCode = 19
    End Enum


    Public Enum XlApplyNamesOrder
        xlColumnThenRow = 2
        xlRowThenColumn = 1
    End Enum


    Public Enum XlArabicModes
        xlArabicBothStrict = 3
        xlArabicNone = 0
        xlArabicStrictAlefHamza = 1
        xlArabicStrictFinalYaa = 2
    End Enum


    Public Enum XlArrangeStyle
        xlArrangeStyleCascade = 7
        xlArrangeStyleHorizontal = -4128
        xlArrangeStyleTiled = 1
        xlArrangeStyleVertical = -4166
    End Enum


    Public Enum XlArrowHeadLength
        xlArrowHeadLengthLong = 3
        xlArrowHeadLengthMedium = -4138
        xlArrowHeadLengthShort = 1
    End Enum


    Public Enum XlArrowHeadStyle
        xlArrowHeadStyleClosed = 3
        xlArrowHeadStyleDoubleClosed = 5
        xlArrowHeadStyleDoubleOpen = 4
        xlArrowHeadStyleNone = -4142
        xlArrowHeadStyleOpen = 2
    End Enum


    Public Enum XlArrowHeadWidth
        xlArrowHeadWidthMedium = -4138
        xlArrowHeadWidthNarrow = 1
        xlArrowHeadWidthWide = 3
    End Enum


    Public Enum XlAutoFillType
        xlFillCopy = 1
        xlFillDays = 5
        xlFillDefault = 0
        xlFillFormats = 3
        xlFillMonths = 7
        xlFillSeries = 2
        xlFillValues = 4
        xlFillWeekdays = 6
        xlFillYears = 8
        xlGrowthTrend = 10
        xlLinearTrend = 9
    End Enum


    Public Enum XlAutoFilterOperator
        xlAnd = 1
        xlBottom10Items = 4
        xlBottom10Percent = 6
        xlOr = 2
        xlTop10Items = 3
        xlTop10Percent = 5
    End Enum


    Public Enum XlAxisCrosses
        xlAxisCrossesAutomatic = -4105
        xlAxisCrossesCustom = -4114
        xlAxisCrossesMaximum = 2
        xlAxisCrossesMinimum = 4
    End Enum


    Public Enum XlAxisGroup
        xlPrimary = 1
        xlSecondary = 2
    End Enum


    Public Enum XlAxisType
        xlCategory = 1
        xlSeriesAxis = 3
        xlValue = 2
    End Enum


    Public Enum XlBackground
        xlBackgroundAutomatic = -4105
        xlBackgroundOpaque = 3
        xlBackgroundTransparent = 2
    End Enum


    Public Enum XlBarShape
        xlBox = 0
        xlConeToMax = 5
        xlConeToPoint = 4
        xlCylinder = 3
        xlPyramidToMax = 2
        xlPyramidToPoint = 1
    End Enum


    Public Enum XlBordersIndex
        xlDiagonalDown = 5
        xlDiagonalUp = 6
        xlEdgeBottom = 9
        xlEdgeLeft = 7
        xlEdgeRight = 10
        xlEdgeTop = 8
        xlInsideHorizontal = 12
        xlInsideVertical = 11
    End Enum


    Public Enum XlBorderWeight
        xlHairline = 1
        xlMedium = -4138
        xlThick = 4
        xlThin = 2
    End Enum


    Public Enum XlBuiltInDialog
        xlDialogActivate = 103
        xlDialogActiveCellFont = 476
        xlDialogAddChartAutoformat = 390
        xlDialogAddinManager = 321
        xlDialogAlignment = 43
        xlDialogApplyNames = 133
        xlDialogApplyStyle = 212
        xlDialogAppMove = 170
        xlDialogAppSize = 171
        xlDialogArrangeAll = 12
        xlDialogAssignToObject = 213
        xlDialogAssignToTool = 293
        xlDialogAttachText = 80
        xlDialogAttachToolbars = 323
        xlDialogAutoCorrect = 485
        xlDialogAxes = 78
        xlDialogBorder = 45
        xlDialogCalculation = 32
        xlDialogCellProtection = 46
        xlDialogChangeLink = 166
        xlDialogChartAddData = 392
        xlDialogChartLocation = 527
        xlDialogChartOptionsDataLabelMultiple = 724
        xlDialogChartOptionsDataLabels = 505
        xlDialogChartOptionsDataTable = 506
        xlDialogChartSourceData = 540
        xlDialogChartTrend = 350
        xlDialogChartType = 526
        xlDialogChartWizard = 288
        xlDialogCheckboxProperties = 435
        xlDialogClear = 52
        xlDialogColorPalette = 161
        xlDialogColumnWidth = 47
        xlDialogCombination = 73
        xlDialogConditionalFormatting = 583
        xlDialogConsolidate = 191
        xlDialogCopyChart = 147
        xlDialogCopyPicture = 108
        xlDialogCreateList = 796
        xlDialogCreateNames = 62
        xlDialogCreatePublisher = 217
        xlDialogCreateRelationship = 1272
        xlDialogCustomizeToolbar = 276
        xlDialogCustomViews = 493
        xlDialogDataDelete = 36
        xlDialogDataLabel = 379
        xlDialogDataLabelMultiple = 723
        xlDialogDataSeries = 40
        xlDialogDataValidation = 525
        xlDialogDefineName = 61
        xlDialogDefineStyle = 229
        xlDialogDeleteFormat = 111
        xlDialogDeleteName = 110
        xlDialogDemote = 203
        xlDialogDisplay = 27
        xlDialogDocumentInspector = 862
        xlDialogEditboxProperties = 438
        xlDialogEditColor = 223
        xlDialogEditDelete = 54
        xlDialogEditionOptions = 251
        xlDialogEditSeries = 228
        xlDialogErrorbarX = 463
        xlDialogErrorbarY = 464
        xlDialogErrorChecking = 732
        xlDialogEvaluateFormula = 709
        xlDialogExternalDataProperties = 530
        xlDialogExtract = 35
        xlDialogFileDelete = 6
        xlDialogFileSharing = 481
        xlDialogFillGroup = 200
        xlDialogFillWorkgroup = 301
        xlDialogFilter = 447
        xlDialogFilterAdvanced = 370
        xlDialogFindFile = 475
        xlDialogFont = 26
        xlDialogFontProperties = 381
        xlDialogFormatAuto = 269
        xlDialogFormatChart = 465
        xlDialogFormatCharttype = 423
        xlDialogFormatFont = 150
        xlDialogFormatLegend = 88
        xlDialogFormatMain = 225
        xlDialogFormatMove = 128
        xlDialogFormatNumber = 42
        xlDialogFormatOverlay = 226
        xlDialogFormatSize = 129
        xlDialogFormatText = 89
        xlDialogFormulaFind = 64
        xlDialogFormulaGoto = 63
        xlDialogFormulaReplace = 130
        xlDialogFunctionWizard = 450
        xlDialogGallery3dArea = 193
        xlDialogGallery3dBar = 272
        xlDialogGallery3dColumn = 194
        xlDialogGallery3dLine = 195
        xlDialogGallery3dPie = 196
        xlDialogGallery3dSurface = 273
        xlDialogGalleryArea = 67
        xlDialogGalleryBar = 68
        xlDialogGalleryColumn = 69
        xlDialogGalleryCustom = 388
        xlDialogGalleryDoughnut = 344
        xlDialogGalleryLine = 70
        xlDialogGalleryPie = 71
        xlDialogGalleryRadar = 249
        xlDialogGalleryScatter = 72
        xlDialogGoalSeek = 198
        xlDialogGridlines = 76
        xlDialogImportTextFile = 666
        xlDialogInsert = 55
        xlDialogInsertHyperlink = 596
        xlDialogInsertObject = 259
        xlDialogInsertPicture = 342
        xlDialogInsertTitle = 380
        xlDialogLabelProperties = 436
        xlDialogListboxProperties = 437
        xlDialogMacroOptions = 382
        xlDialogMailEditMailer = 470
        xlDialogMailLogon = 339
        xlDialogMailNextLetter = 378
        xlDialogMainChart = 85
        xlDialogMainChartType = 185
        xlDialogManageRelationships = 1271
        xlDialogMenuEditor = 322
        xlDialogMove = 262
        xlDialogMyPermission = 834
        xlDialogNameManager = 977
        xlDialogNew = 119
        xlDialogNewName = 978
        xlDialogNewWebQuery = 667
        xlDialogNote = 154
        xlDialogObjectProperties = 207
        xlDialogObjectProtection = 214
        xlDialogOpen = 1
        xlDialogOpenLinks = 2
        xlDialogOpenMail = 188
        xlDialogOpenText = 441
        xlDialogOptionsCalculation = 318
        xlDialogOptionsChart = 325
        xlDialogOptionsEdit = 319
        xlDialogOptionsGeneral = 356
        xlDialogOptionsListsAdd = 458
        xlDialogOptionsME = 647
        xlDialogOptionsTransition = 355
        xlDialogOptionsView = 320
        xlDialogOutline = 142
        xlDialogOverlay = 86
        xlDialogOverlayChartType = 186
        xlDialogPageSetup = 7
        xlDialogParse = 91
        xlDialogPasteNames = 58
        xlDialogPasteSpecial = 53
        xlDialogPatterns = 84
        xlDialogPermission = 832
        xlDialogPhonetic = 656
        xlDialogPivotCalculatedField = 570
        xlDialogPivotCalculatedItem = 572
        xlDialogPivotClientServerSet = 689
        xlDialogPivotFieldGroup = 433
        xlDialogPivotFieldProperties = 313
        xlDialogPivotFieldUngroup = 434
        xlDialogPivotShowPages = 421
        xlDialogPivotSolveOrder = 568
        xlDialogPivotTableOptions = 567
        xlDialogPivotTableSlicerConnections = 1183
        xlDialogPivotTableWhatIfAnalysisSettings = 1153
        xlDialogPivotTableWizard = 312
        xlDialogPlacement = 300
        xlDialogPrint = 8
        xlDialogPrinterSetup = 9
        xlDialogPrintPreview = 222
        xlDialogPromote = 202
        xlDialogProperties = 474
        xlDialogPropertyFields = 754
        xlDialogProtectDocument = 28
        xlDialogProtectSharing = 620
        xlDialogPublishAsWebPage = 653
        xlDialogPushbuttonProperties = 445
        xlDialogRecommendedPivotTables = 1258
        xlDialogReplaceFont = 134
        xlDialogRoutingSlip = 336
        xlDialogRowHeight = 127
        xlDialogRun = 17
        xlDialogSaveAs = 5
        xlDialogSaveCopyAs = 456
        xlDialogSaveNewObject = 208
        xlDialogSaveWorkbook = 145
        xlDialogSaveWorkspace = 285
        xlDialogScale = 87
        xlDialogScenarioAdd = 307
        xlDialogScenarioCells = 305
        xlDialogScenarioEdit = 308
        xlDialogScenarioMerge = 473
        xlDialogScenarioSummary = 311
        xlDialogScrollbarProperties = 420
        xlDialogSearch = 731
        xlDialogSelectSpecial = 132
        xlDialogSendMail = 189
        xlDialogSeriesAxes = 460
        xlDialogSeriesOptions = 557
        xlDialogSeriesOrder = 466
        xlDialogSeriesShape = 504
        xlDialogSeriesX = 461
        xlDialogSeriesY = 462
        xlDialogSetBackgroundPicture = 509
        xlDialogSetManager = 1109
        xlDialogSetMDXEditor = 1208
        xlDialogSetPrintTitles = 23
        xlDialogSetTupleEditorOnColumns = 1108
        xlDialogSetTupleEditorOnRows = 1107
        xlDialogSetUpdateStatus = 159
        xlDialogShowDetail = 204
        xlDialogShowToolbar = 220
        xlDialogSize = 261
        xlDialogSlicerCreation = 1182
        xlDialogSlicerPivotTableConnections = 1184
        xlDialogSlicerSettings = 1179
        xlDialogSort = 39
        xlDialogSortSpecial = 192
        xlDialogSparklineInsertColumn = 1134
        xlDialogSparklineInsertLine = 1133
        xlDialogSparklineInsertWinLoss = 1135
        xlDialogSplit = 137
        xlDialogStandardFont = 190
        xlDialogStandardWidth = 472
        xlDialogStyle = 44
        xlDialogSubscribeTo = 218
        xlDialogSubtotalCreate = 398
        xlDialogSummaryInfo = 474
        xlDialogTable = 41
        xlDialogTabOrder = 394
        xlDialogTextToColumns = 422
        xlDialogUnhide = 94
        xlDialogUpdateLink = 201
        xlDialogVbaInsertFile = 328
        xlDialogVbaMakeAddin = 478
        xlDialogVbaProcedureDefinition = 330
        xlDialogView3d = 197
        xlDialogWebOptionsBrowsers = 773
        xlDialogWebOptionsEncoding = 686
        xlDialogWebOptionsFiles = 684
        xlDialogWebOptionsFonts = 687
        xlDialogWebOptionsGeneral = 683
        xlDialogWebOptionsPictures = 685
        xlDialogWindowMove = 14
        xlDialogWindowSize = 13
        xlDialogWorkbookAdd = 281
        xlDialogWorkbookCopy = 283
        xlDialogWorkbookInsert = 354
        xlDialogWorkbookMove = 282
        xlDialogWorkbookName = 386
        xlDialogWorkbookNew = 302
        xlDialogWorkbookOptions = 284
        xlDialogWorkbookProtect = 417
        xlDialogWorkbookTabSplit = 415
        xlDialogWorkbookUnhide = 384
        xlDialogWorkgroup = 199
        xlDialogWorkspace = 95
        xlDialogZoom = 256
    End Enum


    Public Enum XlCalcFor
        xlAllValues = 0
        xlColGroups = 2
        xlRowGroups = 1
    End Enum


    Public Enum XlCalcMemNumberFormatType
        xlNumberFormatTypeDefault = 0
        xlNumberFormatTypeNumber = 1
        xlNumberFormatTypePercent = 2
    End Enum


    Public Enum XlCalculatedMemberType
        xlCalculatedMeasure = 2
        xlCalculatedMember = 0
        xlCalculatedSet = 1
    End Enum


    Public Enum XlCalculation
        xlCalculationAutomatic = -4105
        xlCalculationManual = -4135
        xlCalculationSemiautomatic = 2
    End Enum


    Public Enum XlCalculationInterruptKey
        xlAnyKey = 2
        xlEscKey = 1
        xlNoKey = 0
    End Enum


    Public Enum XlCalculationState
        xlCalculating = 1
        xlDone = 0
        xlPending = 2
    End Enum


    Public Enum XlCategoryLabelLevel
        xlCategoryLabelLevelAll = -1
        xlCategoryLabelLevelCustom = -2
        xlCategoryLabelLevelNone = -3
    End Enum


    Public Enum XlCategoryType
        xlAutomaticScale = -4105
        xlCategoryScale = 2
        xlTimeScale = 3
    End Enum


    Public Enum XlCellChangedState
        xlCellChangeApplied = 3
        xlCellChanged = 2
        xlCellNotChanged = 1
    End Enum


    Public Enum XlCellInsertionMode
        xlInsertDeleteCells = 1
        xlInsertEntireRows = 2
        xlOverwriteCells = 0
    End Enum


    Public Enum XlCellType
        xlCellTypeAllFormatConditions = -4172
        xlCellTypeAllValidation = -4174
        xlCellTypeBlanks = 4
        xlCellTypeComments = -4144
        xlCellTypeConstants = 2
        xlCellTypeFormulas = -4123
        xlCellTypeLastCell = 11
        xlCellTypeSameFormatConditions = -4173
        xlCellTypeSameValidation = -4175
        xlCellTypeVisible = 12
    End Enum


    Public Enum XlChartElement
        xlChartElementPositionAutomatic = -4105
        xlChartElementPositionCustom = -4114
    End Enum


    Public Enum XlChartGallery
        xlAnyGallery = 23
        xlBuiltIn = 21
        xlUserDefined = 22
    End Enum


    Public Enum XlChartItem
        xlAxis = 21
        xlAxisTitle = 17
        xlChartArea = 2
        xlChartTitle = 4
        xlCorners = 6
        xlDataLabel = 0
        xlDataTable = 7
        xlDisplayUnitLabel = 30
        xlDownBars = 20
        xlDropLines = 26
        xlErrorBars = 9
        xlFloor = 23
        xlHiLoLines = 25
        xlLeaderLines = 29
        xlLegend = 24
        xlLegendEntry = 12
        xlLegendKey = 13
        xlMajorGridlines = 15
        xlMinorGridlines = 16
        xlNothing = 28
        xlPivotChartDropZone = 32
        xlPivotChartFieldButton = 31
        xlPlotArea = 19
        xlRadarAxisLabels = 27
        xlSeries = 3
        xlSeriesLines = 22
        xlShape = 14
        xlTrendline = 8
        xlUpBars = 18
        xlWalls = 5
        xlXErrorBars = 10
        xlYErrorBars = 11
    End Enum


    Public Enum XlChartLocation
        xlLocationAsNewSheet = 1
        xlLocationAsObject = 2
        xlLocationAutomatic = 3
    End Enum


    Public Enum XlChartPicturePlacement
        xlAllFaces = 7
        xlEnd = 2
        xlEndSides = 3
        xlFront = 4
        xlFrontEnd = 6
        xlFrontSides = 5
        xlSides = 1
    End Enum


    Public Enum XlChartPictureType
        xlStack = 2
        xlStackScale = 3
        xlStretch = 1
    End Enum


    Public Enum XlChartSplitType
        xlSplitByCustomSplit = 4
        xlSplitByPercentValue = 3
        xlSplitByPosition = 1
        xlSplitByValue = 2
    End Enum


    Public Enum XlChartType
        xl3DArea = -4098
        xl3DAreaStacked = 78
        xl3DAreaStacked100 = 79
        xl3DBarClustered = 60
        xl3DBarStacked = 61
        xl3DBarStacked100 = 62
        xl3DColumn = -4100
        xl3DColumnClustered = 54
        xl3DColumnStacked = 55
        xl3DColumnStacked100 = 56
        xl3DLine = -4101
        xl3DPie = -4102
        xl3DPieExploded = 70
        xlArea = 1
        xlAreaStacked = 76
        xlAreaStacked100 = 77
        xlBarClustered = 57
        xlBarOfPie = 71
        xlBarStacked = 58
        xlBarStacked100 = 59
        xlBubble = 15
        xlBubble3DEffect = 87
        xlColumnClustered = 51
        xlColumnStacked = 52
        xlColumnStacked100 = 53
        xlConeBarClustered = 102
        xlConeBarStacked = 103
        xlConeBarStacked100 = 104
        xlConeCol = 105
        xlConeColClustered = 99
        xlConeColStacked = 100
        xlConeColStacked100 = 101
        xlCylinderBarClustered = 95
        xlCylinderBarStacked = 96
        xlCylinderBarStacked100 = 97
        xlCylinderCol = 98
        xlCylinderColClustered = 92
        xlCylinderColStacked = 93
        xlCylinderColStacked100 = 94
        xlDoughnut = -4120
        xlDoughnutExploded = 80
        xlLine = 4
        xlLineMarkers = 65
        xlLineMarkersStacked = 66
        xlLineMarkersStacked100 = 67
        xlLineStacked = 63
        xlLineStacked100 = 64
        xlPie = 5
        xlPieExploded = 69
        xlPieOfPie = 68
        xlPyramidBarClustered = 109
        xlPyramidBarStacked = 110
        xlPyramidBarStacked100 = 111
        xlPyramidCol = 112
        xlPyramidColClustered = 106
        xlPyramidColStacked = 107
        xlPyramidColStacked100 = 108
        xlRadar = -4151
        xlRadarFilled = 82
        xlRadarMarkers = 81
        xlStockHLC = 88
        xlStockOHLC = 89
        xlStockVHLC = 90
        xlStockVOHLC = 91
        xlSurface = 83
        xlSurfaceTopView = 85
        xlSurfaceTopViewWireframe = 86
        xlSurfaceWireframe = 84
        xlXYScatter = -4169
        xlXYScatterLines = 74
        xlXYScatterLinesNoMarkers = 75
        xlXYScatterSmooth = 72
        xlXYScatterSmoothNoMarkers = 73
    End Enum


    Public Enum XlCheckInVersionType
        xlCheckInMajorVersion = 1
        xlCheckInMinorVersion = 0
        xlCheckInOverwriteVersion = 2
    End Enum


    Public Enum XlClipboardFormat
        xlClipboardFormatBIFF = 8
        xlClipboardFormatBIFF2 = 18
        xlClipboardFormatBIFF3 = 20
        xlClipboardFormatBIFF4 = 30
        xlClipboardFormatBinary = 15
        xlClipboardFormatBitmap = 9
        xlClipboardFormatCGM = 13
        xlClipboardFormatCSV = 5
        xlClipboardFormatDIF = 4
        xlClipboardFormatDspText = 12
        xlClipboardFormatEmbeddedObject = 21
        xlClipboardFormatEmbedSource = 22
        xlClipboardFormatLink = 11
        xlClipboardFormatLinkSource = 23
        xlClipboardFormatLinkSourceDesc = 32
        xlClipboardFormatMovie = 24
        xlClipboardFormatNative = 14
        xlClipboardFormatObjectDesc = 31
        xlClipboardFormatObjectLink = 19
        xlClipboardFormatOwnerLink = 17
        xlClipboardFormatPICT = 2
        xlClipboardFormatPrintPICT = 3
        xlClipboardFormatRTF = 7
        xlClipboardFormatScreenPICT = 29
        xlClipboardFormatStandardFont = 28
        xlClipboardFormatStandardScale = 27
        xlClipboardFormatSYLK = 6
        xlClipboardFormatTable = 16
        xlClipboardFormatText = 0
        xlClipboardFormatToolFace = 25
        xlClipboardFormatToolFacePICT = 26
        xlClipboardFormatVALU = 1
        xlClipboardFormatWK1 = 10
    End Enum


    Public Enum XlCmdType
        xlCmdCube = 1
        xlCmdDAX = 8
        xlCmdDefault = 4
        xlCmdExcel = 7
        xlCmdList = 5
        xlCmdSql = 2
        xlCmdTable = 3
        xlCmdTableCollection = 6
    End Enum


    Public Enum XlColorIndex
        xlColorIndexAutomatic = -4105
        xlColorIndexNone = -4142
    End Enum


    Public Enum XlColumnDataType
        xlDMYFormat = 4
        xlDYMFormat = 7
        xlEMDFormat = 10
        xlGeneralFormat = 1
        xlMDYFormat = 3
        xlMYDFormat = 6
        xlSkipColumn = 9
        xlTextFormat = 2
        xlYDMFormat = 8
        xlYMDFormat = 5
    End Enum


    Public Enum XlCommandUnderlines
        xlCommandUnderlinesAutomatic = -4105
        xlCommandUnderlinesOff = -4146
        xlCommandUnderlinesOn = 1
    End Enum


    Public Enum XlCommentDisplayMode
        xlCommentAndIndicator = 1
        xlCommentIndicatorOnly = -1
        xlNoIndicator = 0
    End Enum


    Public Enum XlConditionValueTypes
        xlConditionValueAutomaticMax = 7
        xlConditionValueAutomaticMin = 6
        xlConditionValueFormula = 4
        xlConditionValueHighestValue = 2
        xlConditionValueLowestValue = 1
        xlConditionValueNone = -1
        xlConditionValueNumber = 0
        xlConditionValuePercent = 3
        xlConditionValuePercentile = 5
    End Enum


    Public Enum XlConnectionType
        xlConnectionTypeDATAFEED = 6
        xlConnectionTypeMODEL = 7
        xlConnectionTypeODBC = 2
        xlConnectionTypeOLEDB = 1
        xlConnectionTypeTEXT = 4
        xlConnectionTypeWEB = 5
        xlConnectionTypeWORKSHEET = 8
        xlConnectionTypeXMLMAP = 3
    End Enum


    Public Enum XlConsolidationFunction
        xlAverage = -4106
        xlCount = -4112
        xlCountNums = -4113
        xlmax = -4136
        xlMin = -4139
        xlProduct = -4149
        xlStDev = -4155
        xlStDevP = -4156
        xlSum = -4157
        xlUnknown = 1000
        xlVar = -4164
        xlVarP = -4165
    End Enum


    Public Enum XlContainsOperator
        xlBeginsWith = 2
        xlContains = 0
        xlDoesNotContain = 1
        xlEndsWith = 3
    End Enum


    Public Enum XlCopyPictureFormat
        xlBitmap = 2
        xlPicture = -4147
    End Enum


    Public Enum XlCorruptLoad
        xlExtractData = 2
        xlNormalLoad = 0
        xlRepairFile = 1
    End Enum


    Public Enum XlCreator
        xlCreatorCode = 1480803660
    End Enum


    Public Enum XlCredentialsMethod
        CredentialsMethodIntegrated = 0
        CredentialsMethodNone = 1
        CredentialsMethodStored = 2
    End Enum


    Public Enum XlCubeFieldSubType
        xlCubeAttribute = 4
        xlCubeCalculatedMeasure = 5
        xlCubeHierarchy = 1
        xlCubeImplicitMeasure = 11
        xlCubeKPIGoal = 7
        xlCubeKPIStatus = 8
        xlCubeKPITrend = 9
        xlCubeKPIValue = 6
        xlCubeKPIWeight = 10
        xlCubeMeasure = 2
        xlCubeSet = 3
    End Enum


    Public Enum XlCubeFieldType
        xlHierarchy = 1
        xlMeasure = 2
        xlSet = 3
    End Enum


    Public Enum XlCutCopyMode
        xlCopy = 1
        xlCut = 2
    End Enum


    Public Enum XlCVError
        xlErrDiv0 = 2007
        xlErrNA = 2042
        xlErrName = 2029
        xlErrNull = 2000
        xlErrNum = 2036
        xlErrRef = 2023
        xlErrValue = 2015
    End Enum


    Public Enum XlDataBarAxisPosition
        xlDataBarAxisAutomatic = 0
        xlDataBarAxisMidpoint = 1
        xlDataBarAxisNone = 2
    End Enum


    Public Enum XlDataBarBorderType
        xlDataBarBorderNone = 0
        xlDataBarBorderSolid = 1
    End Enum


    Public Enum XlDataBarFillType
        xlDataBarFillGradient = 1
        xlDataBarFillSolid = 0
    End Enum


    Public Enum XlDataBarNegativeColorType
        xlDataBarColor = 0
        xlDataBarSameAsPositive = 1
    End Enum


    Public Enum XlDataLabelPosition
        xlLabelPositionAbove = 0
        xlLabelPositionBelow = 1
        xlLabelPositionBestFit = 5
        xlLabelPositionCenter = -4108
        xlLabelPositionCustom = 7
        xlLabelPositionInsideBase = 4
        xlLabelPositionInsideEnd = 3
        xlLabelPositionLeft = -4131
        xlLabelPositionMixed = 6
        xlLabelPositionOutsideEnd = 2
        xlLabelPositionRight = -4152
    End Enum


    Public Enum XlDataLabelSeparator
        xlDataLabelSeparatorDefault = 1
    End Enum


    Public Enum XlDataLabelsType
        xlDataLabelsShowBubbleSizes = 6
        xlDataLabelsShowLabel = 4
        xlDataLabelsShowLabelAndPercent = 5
        xlDataLabelsShowNone = -4142
        xlDataLabelsShowPercent = 3
        xlDataLabelsShowValue = 2
    End Enum


    Public Enum XlDataSeriesDate
        xlDay = 1
        xlMonth = 3
        xlWeekday = 2
        xlYear = 4
    End Enum


    Public Enum XlDataSeriesType
        xlAutoFill = 4
        xlChronological = 3
        xlDataSeriesLinear = -4132
        xlGrowth = 2
    End Enum


    Public Enum XlDeleteShiftDirection
        xlShiftToLeft = -4159
        xlShiftUp = -4162
    End Enum


    Public Enum XlDirection
        xlDown = -4121
        xlToLeft = -4159
        xlToRight = -4161
        xlUp = -4162
    End Enum


    Public Enum XlDisplayBlanksAs
        xlInterpolated = 3
        xlNotPlotted = 1
        xlZero = 2
    End Enum


    Public Enum XlDisplayDrawingObjects
        xlDisplayShapes = -4104
        xlHide = 3
        xlPlaceholders = 2
    End Enum


    Public Enum XlDisplayUnit
        xlHundredMillions = -8
        xlHundreds = -2
        xlHundredThousands = -5
        xlMillionMillions = -10
        xlMillions = -6
        xlTenMillions = -7
        xlTenThousands = -4
        xlThousandMillions = -9
        xlThousands = -3
    End Enum


    Public Enum XlDupeUnique
        xlDuplicate = 1
        xlUnique = 0
    End Enum


    Public Enum XlDVAlertStyle
        xlValidAlertInformation = 3
        xlValidAlertStop = 1
        xlValidAlertWarning = 2
    End Enum


    Public Enum XlDVType
        xlValidateCustom = 7
        xlValidateDate = 4
        xlValidateDecimal = 2
        xlValidateInputOnly = 0
        xlValidateList = 3
        xlValidateTextLength = 6
        xlValidateTime = 5
        xlValidateWholeNumber = 1
    End Enum


    Public Enum XlDynamicFilterCriteria
        xlFilterAboveAverage = 33
        xlFilterAllDatesInPeriodApril = 24
        xlFilterAllDatesInPeriodAugust = 28
        xlFilterAllDatesInPeriodDecember = 32
        xlFilterAllDatesInPeriodFebruray = 22
        xlFilterAllDatesInPeriodJanuary = 21
        xlFilterAllDatesInPeriodJuly = 27
        xlFilterAllDatesInPeriodJune = 26
        xlFilterAllDatesInPeriodMarch = 23
        xlFilterAllDatesInPeriodMay = 25
        xlFilterAllDatesInPeriodNovember = 31
        xlFilterAllDatesInPeriodOctober = 30
        xlFilterAllDatesInPeriodQuarter1 = 17
        xlFilterAllDatesInPeriodQuarter2 = 18
        xlFilterAllDatesInPeriodQuarter3 = 19
        xlFilterAllDatesInPeriodQuarter4 = 20
        xlFilterAllDatesInPeriodSeptember = 29
        xlFilterBelowAverage = 34
        xlFilterLastMonth = 8
        xlFilterLastQuarter = 11
        xlFilterLastWeek = 5
        xlFilterLastYear = 14
        xlFilterNextMonth = 9
        xlFilterNextQuarter = 12
        xlFilterNextWeek = 6
        xlFilterNextYear = 15
        xlFilterThisMonth = 7
        xlFilterThisQuarter = 10
        xlFilterThisWeek = 4
        xlFilterThisYear = 13
        xlFilterToday = 1
        xlFilterTomorrow = 3
        xlFilterYearToDate = 16
        xlFilterYesterday = 2
    End Enum


    Public Enum XlEditionFormat
        xlBIFF = 2
        xlPICT = 1
        xlRTF = 4
        xlVALU = 8
    End Enum


    Public Enum XlEditionOptionsOption
        xlAutomaticUpdate = 4
        xlCancel = 1
        xlChangeAttributes = 6
        xlManualUpdate = 5
        xlOpenSource = 3
        xlSelect = 3
        xlSendPublisher = 2
        xlUpdateSubscriber = 2
    End Enum


    Public Enum XlEditionType
        xlPublisher = 1
        xlSubscriber = 2
    End Enum


    Public Enum XlEnableCancelKey
        xlDisabled = 0
        xlErrorHandler = 2
        xlInterrupt = 1
    End Enum


    Public Enum XlEnableSelection
        xlNoRestrictions = 0
        xlNoSelection = -4142
        xlUnlockedCells = 1
    End Enum


    Public Enum XlEndStyleCap
        xlCap = 1
        xlNoCap = 2
    End Enum


    Public Enum XlErrorBarDirection
        xlX = -4168
        xlY = 1
    End Enum


    Public Enum XlErrorBarInclude
        xlErrorBarIncludeBoth = 1
        xlErrorBarIncludeMinusValues = 3
        xlErrorBarIncludeNone = -4142
        xlErrorBarIncludePlusValues = 2
    End Enum


    Public Enum XlErrorBarType
        xlErrorBarTypeCustom = -4114
        xlErrorBarTypeFixedValue = 1
        xlErrorBarTypePercent = 2
        xlErrorBarTypeStDev = -4155
        xlErrorBarTypeStError = 4
    End Enum


    Public Enum XlErrorChecks
        xlEmptyCellReferences = 7
        xlEvaluateToError = 1
        xlInconsistentFormula = 4
        xlListDataValidation = 8
        xlNumberAsText = 3
        xlOmittedCells = 5
        xlTextDate = 2
        xlUnlockedFormulaCells = 6
    End Enum


    Public Enum XlFileAccess
        xlReadOnly = 3
        xlReadWrite = 2
    End Enum


    Public Enum XlFileFormat
        xlAddIn = 18
        xlAddIn8 = 18
        xlCSV = 6
        xlCSVMac = 22
        xlCSVMSDOS = 24
        xlCSVWindows = 23
        xlCurrentPlatformText = -4158
        xlDBF2 = 7
        xlDBF3 = 8
        xlDBF4 = 11
        xlDIF = 9
        xlExcel12 = 50
        xlExcel2 = 16
        xlExcel2FarEast = 27
        xlExcel3 = 29
        xlExcel4 = 33
        xlExcel4Workbook = 35
        xlExcel5 = 39
        xlExcel7 = 39
        xlExcel8 = 56
        xlExcel9795 = 43
        xlHtml = 44
        xlIntlAddIn = 26
        xlIntlMacro = 25
        xlOpenDocumentSpreadsheet = 60
        xlOpenXMLAddIn = 55
        xlOpenXMLStrictWorkbook = 61
        xlOpenXMLTemplate = 54
        xlOpenXMLTemplateMacroEnabled = 53
        xlOpenXMLWorkbook = 51
        xlOpenXMLWorkbookMacroEnabled = 52
        xlSYLK = 2
        xlTemplate = 17
        xlTemplate8 = 17
        xlTextMac = 19
        xlTextMSDOS = 21
        xlTextPrinter = 36
        xlTextWindows = 20
        xlUnicodeText = 42
        xlWebArchive = 45
        xlWJ2WD1 = 14
        xlWJ3 = 40
        xlWJ3FJ3 = 41
        xlWK1 = 5
        xlWK1ALL = 31
        xlWK1FMT = 30
        xlWK3 = 15
        xlWK3FM3 = 32
        xlWK4 = 38
        xlWKS = 4
        xlWorkbookDefault = 51
        xlWorkbookNormal = -4143
        xlWorks2FarEast = 28
        xlWQ1 = 34
        xlXMLSpreadsheet = 46
    End Enum


    Public Enum XlFileValidationPivotMode
        xlFileValidationPivotDefault = 0
        xlFileValidationPivotRun = 1
        xlFileValidationPivotSkip = 2
    End Enum


    Public Enum XlFillWith
        xlFillWithAll = -4104
        xlFillWithContents = 2
        xlFillWithFormats = -4122
    End Enum


    Public Enum XlFilterAction
        xlFilterCopy = 2
        xlFilterInPlace = 1
    End Enum


    Public Enum XlFilterAllDatesInPeriod
        xlFilterAllDatesInPeriodDay = 2
        xlFilterAllDatesInPeriodHour = 3
        xlFilterAllDatesInPeriodMinute = 4
        xlFilterAllDatesInPeriodMonth = 1
        xlFilterAllDatesInPeriodSecond = 5
        xlFilterAllDatesInPeriodYear = 0
    End Enum


    Public Enum XlFilterStatus
        xlFilterStatusOK = 0
        xlFilterStatusDateWrongOrder = 1
        xlFilterStatusDateHasTime = 2
        xlFilterStatusInvalidDate = 3
    End Enum


    Public Enum XlFindLookIn
        xlComments = -4144
        xlFormulas = -4123
        xlValues = -4163
    End Enum


    Public Enum XlFixedFormatQuality
        xlQualityMinimum = 1
        xlQualityStandard = 0
    End Enum


    Public Enum XlFixedFormatType
        xlTypePDF = 0
        xlTypeXPS = 0
    End Enum


    Public Enum XlFormatConditionOperator
        xlBetween = 1
        xlEqual = 3
        xlGreater = 5
        xlGreaterEqual = 7
        xlLess = 6
        xlLessEqual = 8
        xlNotBetween = 2
        xlNotEqual = 4
    End Enum


    Public Enum XlFormatConditionType
        xlAboveAverageCondition = 12
        xlBlanksCondition = 10
        xlCellValue = 1
        xlColorScale = 3
        xlDatabar = 4
        xlErrorsCondition = 16
        xlExpression = 2
        XlIconSet = 6
        xlNoBlanksCondition = 13
        xlNoErrorsCondition = 17
        xlTextString = 9
        xlTimePeriod = 11
        xlTop10 = 5
        xlUniqueValues = 8
    End Enum


    Public Enum XlFormatFilterTypes
        FilterBottom = 0
        FilterBottomPercent = 2
        FilterTop = 1
        FilterTopPercent = 3
    End Enum


    Public Enum XlFormControl
        xlButtonControl = 0
        xlCheckBox = 1
        xlDropDown = 2
        xlEditBox = 3
        xlGroupBox = 4
        xlLabel = 5
        xlListBox = 6
        xlOptionButton = 7
        xlScrollBar = 8
        xlSpinner = 9
    End Enum


    Public Enum XlFormulaLabel
        xlColumnLabels = 2
        xlMixedLabels = 3
        xlNoLabels = -4142
        xlRowLabels = 1
    End Enum


    Public Enum XlGenerateTableRefs
        xlA1TableRefs = 0
        xlTableNames = 1
    End Enum


    Public Enum XlGradientFillType
        GradientFillLinear = 0
        GradientFillPath = 1
    End Enum


    Public Enum XlHAlign
        xlHAlignCenter = -4108
        xlHAlignCenterAcrossSelection = 7
        xlHAlignDistributed = -4117
        xlHAlignFill = 5
        xlHAlignGeneral = 1
        xlHAlignJustify = -4130
        xlHAlignLeft = -4131
        xlHAlignRight = -4152
    End Enum


    Public Enum XlHebrewModes
        xlHebrewFullScript = 0
        xlHebrewMixedAuthorizedScript = 3
        xlHebrewMixedScript = 2
        xlHebrewPartialScript = 1
    End Enum


    Public Enum XlHighlightChangesTime
        xlAllChanges = 2
        xlNotYetReviewed = 3
        xlSinceMyLastSave = 1
    End Enum


    Public Enum XlHtmlType
        xlHtmlCalc = 1
        xlHtmlChart = 3
        xlHtmlList = 2
        xlHtmlStatic = 0
    End Enum


    Public Enum XlIcon
        xlIcon0Bars = 37
        xlIcon0FilledBoxes = 52
        xlIcon1Bar = 38
        xlIcon1FilledBox = 51
        xlIcon2Bars = 39
        xlIcon2FilledBoxes = 50
        xlIcon3Bars = 40
        xlIcon3FilledBoxes = 49
        xlIcon4Bars = 41
        xlIcon4FilledBoxes = 48
        xlIconBlackCircle = 32
        xlIconBlackCircleWithBorder = 13
        xlIconCircleWithOneWhiteQuarter = 33
        xlIconCircleWithThreeWhiteQuarters = 35
        xlIconCircleWithTwoWhiteQuarters = 34
        xlIconGoldStar = 42
        xlIconGrayCircle = 31
        xlIconGrayDownArrow = 6
        xlIconGrayDownInclineArrow = 28
        xlIconGraySideArrow = 5
        xlIconGrayUpArrow = 4
        xlIconGrayUpInclineArrow = 27
        xlIconGreenCheck = 22
        xlIconGreenCheckSymbol = 19
        xlIconGreenCircle = 10
        xlIconGreenFlag = 7
        xlIconGreenTrafficLight = 14
        xlIconGreenUpArrow = 1
        xlIconGreenUpTriangle = 45
        xlIconHalfGoldStar = 43
        xlIconNoCellIcon = -1
        xlIconPinkCircle = 30
        xlIconRedCircle = 29
        xlIconRedCircleWithBorder = 12
        xlIconRedCross = 24
        xlIconRedCrossSymbol = 21
        xlIconRedDiamond = 18
        xlIconRedDownArrow = 3
        xlIconRedDownTriangle = 47
        xlIconRedFlag = 9
        xlIconRedTrafficLight = 16
        xlIconSilverStar = 44
        xlIconWhiteCircleAllWhiteQuarters = 36
        xlIconYellowCircle = 11
        xlIconYellowDash = 46
        xlIconYellowDownInclineArrow = 26
        xlIconYellowExclamation = 23
        xlIconYellowExclamationSymbol = 20
        xlIconYellowFlag = 8
        xlIconYellowSideArrow = 2
        xlIconYellowTrafficLight = 15
        xlIconYellowTriangle = 17
        xlIconYellowUpInclineArrow = 25
    End Enum


    Public Enum XlIconSet
        xl3Arrows = 1
        xl3ArrowsGray = 2
        xl3Flags = 3
        xl3Signs = 6
        xl3Symbols = 7
        xl3TrafficLights1 = 4
        xl3TrafficLights2 = 5
        xl4Arrows = 8
        xl4ArrowsGray = 9
        xl4CRV = 11
        xl4RedToBlack = 10
        xl4TrafficLights = 12
        xl5Arrows = 13
        xl5ArrowsGray = 14
        xl5CRV = 15
        xl5Quarters = 16
    End Enum


    Public Enum XlIMEMode
        xlIMEModeAlpha = 8
        xlIMEModeAlphaFull = 7
        xlIMEModeDisable = 3
        xlIMEModeHangul = 10
        xlIMEModeHangulFull = 9
        xlIMEModeHiragana = 4
        xlIMEModeKatakana = 5
        xlIMEModeKatakanaHalf = 6
        xlIMEModeNoControl = 0
        xlIMEModeOff = 2
        xlIMEModeOn = 1
    End Enum


    Public Enum XlImportDataAs
        xlPivotTableReport = 1
        xlQueryTable = 0
    End Enum


    Public Enum XlInsertFormatOrigin
        xlFormatFromLeftOrAbove = 0
        xlFormatFromRightOrBelow = 1
    End Enum


    Public Enum XlInsertShiftDirection
        xlShiftDown = -4121
        xlShiftToRight = -4161
    End Enum


    Public Enum XlLayoutFormType
        xlOutline = 1
        xlTabular = 0
    End Enum


    Public Enum XlLayoutRowType
        xlCompactRow = 0
        xlOutlineRow = 2
        xlTabularRow = 1
    End Enum


    Public Enum XlLegendPosition
        xlLegendPositionBottom = -4107
        xlLegendPositionCorner = 2
        xlLegendPositionLeft = -4131
        xlLegendPositionRight = -4152
        xlLegendPositionTop = -4160
    End Enum


    Public Enum XlLineStyle
        xlContinuous = 1
        xlDash = -4115
        xlDashDot = 4
        xlDashDotDot = 5
        xlDot = -4118
        xlDouble = -4119
        xlLineStyleNone = -4142
        xlSlantDashDot = 13
    End Enum


    Public Enum XlLink
        xlExcelLinks = 1
        xlOLELinks = 2
        xlPublishers = 5
        xlSubscribers = 6
    End Enum


    Public Enum XlLinkInfo
        xlEditionDate = 2
        xlLinkInfoStatus = 3
        xlUpdateState = 1
    End Enum


    Public Enum XlLinkInfoType
        xlLinkInfoOLELinks = 2
        xlLinkInfoPublishers = 5
        xlLinkInfoSubscribers = 6
    End Enum


    Public Enum XlLinkStatus
        xlLinkStatusCopiedValues = 10
        xlLinkStatusIndeterminate = 5
        xlLinkStatusInvalidName = 7
        xlLinkStatusMissingFile = 1
        xlLinkStatusMissingSheet = 2
        xlLinkStatusNotStarted = 6
        xlLinkStatusOK = 0
        xlLinkStatusOld = 3
        xlLinkStatusSourceNotCalculated = 4
        xlLinkStatusSourceNotOpen = 8
        xlLinkStatusSourceOpen = 9
    End Enum


    Public Enum XlLinkType
        xlLinkTypeExcelLinks = 1
        xlLinkTypeOLELinks = 2
    End Enum


    Public Enum XlListConflict
        xlListConflictDialog = 0
        xlListConflictDiscardAllConflicts = 2
        xlListConflictError = 3
        xlListConflictRetryAllConflicts = 1
    End Enum


    Public Enum XlListDataType
        xlListDataTypeCheckbox = 9
        xlListDataTypeChoice = 6
        xlListDataTypeChoiceMulti = 7
        xlListDataTypeCounter = 11
        xlListDataTypeCurrency = 4
        xlListDataTypeDateTime = 5
        xlListDataTypeHyperLink = 10
        xlListDataTypeListLookup = 8
        xlListDataTypeMultiLineRichText = 12
        xlListDataTypeMultiLineText = 2
        xlListDataTypeNone = 0
        xlListDataTypeNumber = 3
        xlListDataTypeText = 1
    End Enum


    Public Enum XlListObjectSourceType
        xlSrcExternal = 0
        xlSrcModel = 4
        xlSrcQuery = 3
        xlSrcRange = 1
        xlSrcXml = 2
    End Enum


    Public Enum XlLocationInTable
        xlColumnHeader = -4110
        xlColumnItem = 5
        xlDataHeader = 3
        xlDataItem = 7
        xlPageHeader = 2
        xlPageItem = 6
        xlRowHeader = -4153
        xlRowItem = 4
        xlTableBody = 8
    End Enum


    Public Enum XlLookAt
        xlPart = 2
        xlWhole = 1
    End Enum


    Public Enum XlLookFor
        LookForBlanks = 0
        LookForErrors = 1
        LookForFormulas = 2
    End Enum


    Public Enum XlMailSystem
        xlMAPI = 1
        xlNoMailSystem = 0
        xlPowerTalk = 2
    End Enum


    Public Enum XlMarkerStyle
        xlMarkerStyleAutomatic = -4105
        xlMarkerStyleCircle = 8
        xlMarkerStyleDash = -4115
        xlMarkerStyleDiamond = 2
        xlMarkerStyleDot = -4118
        xlMarkerStyleNone = -4142
        xlMarkerStylePicture = -4147
        xlMarkerStylePlus = 9
        xlMarkerStyleSquare = 1
        xlMarkerStyleStar = 5
        xlMarkerStyleTriangle = 3
        xlMarkerStyleX = -4168
    End Enum


    Public Enum XlMeasurementUnits
        xlCentimeters = 1
        xlInches = 0
        xlMillimeters = 2
    End Enum


    Public Enum XlMouseButton
        xlNoButton = 0
        xlPrimaryButton = 1
        xlSecondaryButton = 2
    End Enum


    Public Enum XlMousePointer
        xlDefault = -4143
        xlIBeam = 3
        xlNorthwestArrow = 1
        xlWait = 2
    End Enum


    Public Enum XlMSApplication
        xlMicrosoftAccess = 4
        xlMicrosoftFoxPro = 5
        xlMicrosoftMail = 3
        xlMicrosoftPowerPoint = 2
        xlMicrosoftProject = 6
        xlMicrosoftSchedulePlus = 7
        xlMicrosoftWord = 1
    End Enum


    Public Enum XlOartHorizontalOverflow
        xlOartHorizontalOverflowClip = 1
        xlOartHorizontalOverflowOverflow = 0
    End Enum


    Public Enum XlOartVerticalOverflow
        xlOartVerticalOverflowClip = 1
        xlOartVerticalOverflowEllipsis = 2
        xlOartVerticalOverflowOverflow = 0
    End Enum


    Public Enum XlObjectSize
        xlFitToPage = 2
        xlFullPage = 3
        xlScreenSize = 1
    End Enum


    Public Enum XlOLEType
        xlOLEControl = 2
        xlOLEEmbed = 1
        xlOLELink = 0
    End Enum


    Public Enum XlOLEVerb
        xlVerbOpen = 2
        xlVerbPrimary = 1
    End Enum


    Public Enum XlOrder
        xlDownThenOver = 1
        xlOverThenDown = 2
    End Enum


    Public Enum XlOrientation
        xlDownward = -4170
        xlHorizontal = -4128
        xlUpward = -4171
        xlVertical = -4166
    End Enum


    Public Enum XlPageBreak
        xlPageBreakAutomatic = -4105
        xlPageBreakManual = -4135
        xlPageBreakNone = -4142
    End Enum


    Public Enum XlPageBreakExtent
        xlPageBreakFull = 1
        xlPageBreakPartial = 2
    End Enum


    Public Enum XlPageOrientation
        xlLandscape = 2
        xlPortrait = 1
    End Enum


    Public Enum XlPaperSize
        xlPaper10x14 = 16
        xlPaper11x17 = 17
        xlPaperA3 = 8
        xlPaperA4 = 9
        xlPaperA4Small = 10
        xlPaperA5 = 11
        xlPaperB4 = 12
        xlPaperB5 = 13
        xlPaperCsheet = 24
        xlPaperDsheet = 25
        xlPaperEnvelope10 = 20
        xlPaperEnvelope11 = 21
        xlPaperEnvelope12 = 22
        xlPaperEnvelope14 = 23
        xlPaperEnvelope9 = 19
        xlPaperEnvelopeB4 = 33
        xlPaperEnvelopeB5 = 34
        xlPaperEnvelopeB6 = 35
        xlPaperEnvelopeC3 = 29
        xlPaperEnvelopeC4 = 30
        xlPaperEnvelopeC5 = 28
        xlPaperEnvelopeC6 = 31
        xlPaperEnvelopeC65 = 32
        xlPaperEnvelopeDL = 27
        xlPaperEnvelopeItaly = 36
        xlPaperEnvelopeMonarch = 37
        xlPaperEnvelopePersonal = 38
        xlPaperEsheet = 26
        xlPaperExecutive = 7
        xlPaperFanfoldLegalGerman = 41
        xlPaperFanfoldStdGerman = 40
        xlPaperFanfoldUS = 39
        xlPaperFolio = 14
        xlPaperLedger = 4
        xlPaperLegal = 5
        xlPaperLetter = 1
        xlPaperLetterSmall = 2
        xlPaperNote = 18
        xlPaperQuarto = 15
        xlPaperStatement = 6
        xlPaperTabloid = 3
        xlPaperUser = 256
    End Enum


    Public Enum XlParameterDataType
        xlParamTypeBigInt = -5
        xlParamTypeBinary = -2
        xlParamTypeBit = -7
        xlParamTypeChar = 1
        xlParamTypeDate = 9
        xlParamTypeDecimal = 3
        xlParamTypeDouble = 8
        xlParamTypeFloat = 6
        xlParamTypeInteger = 4
        xlParamTypeLongVarBinary = -4
        xlParamTypeLongVarChar = -1
        xlParamTypeNumeric = 2
        xlParamTypeReal = 7
        xlParamTypeSmallInt = 5
        xlParamTypeTime = 10
        xlParamTypeTimestamp = 11
        xlParamTypeTinyInt = -6
        xlParamTypeUnknown = 0
        xlParamTypeVarBinary = -3
        xlParamTypeVarChar = 12
        xlParamTypeWChar = -8
    End Enum


    Public Enum XlParameterType
        xlConstant = 1
        xlPrompt = 0
        xlRange = 2
    End Enum


    Public Enum XlPasteSpecialOperation
        xlPasteSpecialOperationAdd = 2
        xlPasteSpecialOperationDivide = 5
        xlPasteSpecialOperationMultiply = 4
        xlPasteSpecialOperationNone = -4142
        xlPasteSpecialOperationSubtract = 3
    End Enum


    Public Enum XlPasteType
        xlPasteAll = -4104
        xlPasteAllExceptBorders = 7
        xlPasteColumnWidths = 8
        xlPasteComments = -4144
        xlPasteFormats = -4122
        xlPasteFormulas = -4123
        xlPasteFormulasAndNumberFormats = 11
        xlPasteValidation = 6
        xlPasteValues = -4163
        xlPasteValuesAndNumberFormats = 12
    End Enum


    Public Enum XlPattern
        xlPatternAutomatic = -4105
        xlPatternChecker = 9
        xlPatternCrissCross = 16
        xlPatternDown = -4121
        xlPatternGray16 = 17
        xlPatternGray25 = -4124
        xlPatternGray50 = -4125
        xlPatternGray75 = -4126
        xlPatternGray8 = 18
        xlPatternGrid = 15
        xlPatternHorizontal = -4128
        xlPatternLightDown = 13
        xlPatternLightHorizontal = 11
        xlPatternLightUp = 14
        xlPatternLightVertical = 12
        xlPatternNone = -4142
        xlPatternSemiGray75 = 10
        xlPatternSolid = 1
        xlPatternUp = -4162
        xlPatternVertical = -4166
        xlSolid = 1
    End Enum


    Public Enum XlPhoneticAlignment
        xlPhoneticAlignCenter = 2
        xlPhoneticAlignDistributed = 3
        xlPhoneticAlignLeft = 1
        xlPhoneticAlignNoControl = 0
    End Enum


    Public Enum XlPhoneticCharacterType
        xlHiragana = 2
        xlKatakana = 1
        xlKatakanaHalf = 0
        xlNoConversion = 3
    End Enum


    Public Enum XlPictureAppearance
        xlPrinter = 2
        xlScreen = 1
    End Enum


    Public Enum XlPictureConvertorType
        xlBMP = 1
        xlCGM = 7
        xlDRW = 4
        xlDXF = 5
        xlEPS = 8
        xlHGL = 6
        xlPCT = 13
        xlPCX = 10
        xlPIC = 11
        xlPLT = 12
        xlTIF = 9
        xlWMF = 2
        xlWPG = 3
    End Enum


    Public Enum XlPieSliceIndex
        xlCenterPoint = 5
        xlInnerCenterPoint = 8
        xlInnerClockwisePoint = 7
        xlInnerCounterClockwisePoint = 9
        xlMidClockwiseRadiusPoint = 4
        xlMidCounterClockwiseRadiusPoint = 6
        xlOuterCenterPoint = 2
        xlOuterClockwisePoint = 3
        xlOuterCounterClockwisePoint = 1
    End Enum


    Public Enum XlPieSliceLocation
        xlHorizontalCoordinate = 1
        xlVerticalCoordinate = 2
    End Enum


    Public Enum XlPivotCellType
        xlPivotCellBlankCell = 9
        xlPivotCellCustomSubtotal = 7
        xlPivotCellDataField = 4
        xlPivotCellDataPivotField = 8
        xlPivotCellGrandTotal = 3
        xlPivotCellPageFieldItem = 6
        xlPivotCellPivotField = 5
        xlPivotCellPivotItem = 1
        xlPivotCellSubtotal = 2
        xlPivotCellValue = 0
    End Enum


    Public Enum XlPivotConditionScope
        xlDataFieldScope = 2
        xlFieldsScope = 1
        xlSelectionScope = 0
    End Enum


    Public Enum XlPivotFieldCalculation
        xlDifferenceFrom = 2
        xlIndex = 9
        xlNoAdditionalCalculation = -4143
        xlPercentDifferenceFrom = 4
        xlPercentOf = 3
        xlPercentOfColumn = 7
        xlPercentOfParent = 12
        xlPercentOfParentColumn = 11
        xlPercentOfParentRow = 10
        xlPercentOfRow = 6
        xlPercentOfTotal = 8
        xlPercentRunningTotal = 13
        xlRankAscending = 14
        xlRankDecending = 15
        xlRunningTotal = 5
    End Enum


    Public Enum XlPivotFieldDataType
        xlDate = 2
        xlNumber = -4145
        xlText = -4158
    End Enum


    Public Enum XlPivotFieldOrientation
        xlColumnField = 2
        xlDataField = 4
        xlHidden = 0
        xlPageField = 3
        xlRowField = 1
    End Enum


    Public Enum XlPivotFieldRepeatLabels
        xlDoNotRepeatLabels = 1
        xlRepeatLabels = 1
    End Enum


    Public Enum XlPivotFilterType
        xlBefore = 31
        xlBeforeOrEqualTo = 32
        xlAfter = 33
        xlAfterOrEqualTo = 34
        xlAllDatesInPeriodJanuary = 53
        xlAllDatesInPeriodFebruary = 54
        xlAllDatesInPeriodMarch = 55
        xlAllDatesInPeriodApril = 56
        xlAllDatesInPeriodMay = 57
        xlAllDatesInPeriodJune = 58
        xlAllDatesInPeriodJuly = 59
        xlAllDatesInPeriodAugust = 60
        xlAllDatesInPeriodSeptember = 61
        xlAllDatesInPeriodOctober = 62
        xlAllDatesInPeriodNovember = 63
        xlAllDatesInPeriodDecember = 64
        xlAllDatesInPeriodQuarter1 = 49
        xlAllDatesInPeriodQuarter2 = 50
        xlAllDatesInPeriodQuarter3 = 51
        xlAllDatesInPeriodQuarter4 = 52
        xlBottomCount = 2
        xlBottomPercent = 4
        xlBottomSum = 6
        xlCaptionBeginsWith = 17
        xlCaptionContains = 21
        xlCaptionDoesNotBeginWith = 18
        xlCaptionDoesNotContain = 22
        xlCaptionDoesNotEndWith = 20
        xlCaptionDoesNotEqual = 16
        xlCaptionEndsWith = 19
        xlCaptionEquals = 15
        xlCaptionIsBetween = 27
        xlCaptionIsGreaterThan = 23
        xlCaptionIsGreaterThanOrEqualTo = 24
        xlCaptionIsLessThan = 25
        xlCaptionIsLessThanOrEqualTo = 26
        xlCaptionIsNotBetween = 28
        xlDateBetween = 32
        xlDateLastMonth = 41
        xlDateLastQuarter = 44
        xlDateLastWeek = 38
        xlDateLastYear = 47
        xlDateNextMonth = 39
        xlDateNextQuarter = 42
        xlDateNextWeek = 36
        xlDateNextYear = 45
        xlDateThisMonth = 40
        xlDateThisQuarter = 43
        xlDateThisWeek = 37
        xlDateThisYear = 46
        xlDateToday = 34
        xlDateTomorrow = 33
        xlDateYesterday = 35
        xlNotSpecificDate = 30
        xlSpecificDate = 29
        xlTopCount = 1
        xlTopPercent = 3
        xlTopSum = 5
        xlValueDoesNotEqual = 8
        xlValueEquals = 7
        xlValueIsBetween = 13
        xlValueIsGreaterThan = 9
        xlValueIsGreaterThanOrEqualTo = 10
        xlValueIsLessThan = 11
        xlValueIsLessThanOrEqualTo = 12
        xlValueIsNotBetween = 14
        xlYearToDate = 48
    End Enum


    Public Enum XlPivotFormatType
        xlPTClassic = 20
        xlPTNone = 21
        xlReport1 = 0
        xlReport10 = 9
        xlReport2 = 1
        xlReport3 = 2
        xlReport4 = 3
        xlReport5 = 4
        xlReport6 = 5
        xlReport7 = 6
        xlReport8 = 7
        xlReport9 = 8
        xlTable1 = 10
        xlTable10 = 19
        xlTable2 = 11
        xlTable3 = 12
        xlTable4 = 13
        xlTable5 = 14
        xlTable6 = 15
        xlTable7 = 16
        xlTable8 = 17
        xlTable9 = 18
    End Enum


    Public Enum XlPivotLineType
        xlPivotLineBlank = 3
        xlPivotLineGrandTotal = 2
        xlPivotLineRegular = 0
        xlPivotLineSubtotal = 1
    End Enum


    Public Enum XlPivotTableMissingItems
        xlMissingItemsDefault = -1
        xlMissingItemsMax = 32500
        xlMissingItemsMax2 = 1048576
        xlMissingItemsNone = 0
    End Enum


    Public Enum XlPivotTableSourceType
        xlConsolidation = 3
        xlDatabase = 1
        xlExternal = 2
        xlPivotTable = -4148
        xlScenario = 4
    End Enum


    Public Enum XlPivotTableVersionList
        xlPivotTableVersion2000 = 0
        xlPivotTableVersion10 = 1
        xlPivotTableVersion11 = 2
        xlPivotTableVersion12 = 3
        xlPivotTableVersion14 = 4
        xlPivotTableVersion15 = 5
        xlPivotTableVersionCurrent = -1
    End Enum


    Public Enum XlPlacement
        xlFreeFloating = 3
        xlMove = 2
        xlMoveAndSize = 1
    End Enum


    Public Enum XlPlatform
        xlMacintosh = 1
        xlMSDOS = 3
        xlWindows = 2
    End Enum


    Public Enum XlPortugueseReform
        xlPortugueseBoth = 3
        xlPortuguesePostReform = 2
        xlPortuguesePreReform = 1
    End Enum


    Public Enum XlPrintErrors
        xlPrintErrorsBlank = 1
        xlPrintErrorsDash = 2
        xlPrintErrorsDisplayed = 0
        xlPrintErrorsNA = 3
    End Enum


    Public Enum XlPrintLocation
        xlPrintInPlace = 16
        xlPrintNoComments = -4142
        xlPrintSheetEnd = 1
    End Enum


    Public Enum XlPriority
        xlPriorityHigh = -4127
        xlPriorityLow = -4134
        xlPriorityNormal = -4143
    End Enum


    Public Enum XlPropertyDisplayedIn
        xlDisplayPropertyInPivotTable = 1
        xlDisplayPropertyInPivotTableAndTooltip = 3
        xlDisplayPropertyInTooltip = 2
    End Enum


    Public Enum XlProtectedViewCloseReason
        xlProtectedViewCloseEdit = 1
        xlProtectedViewCloseForced = 2
        xlProtectedViewCloseNormal = 0
    End Enum


    Public Enum XlProtectedViewWindowState
        xlProtectedViewWindowMaximized = 2
        xlProtectedViewWindowMinimized = 1
        xlProtectedViewWindowNormal = 0
    End Enum


    Public Enum XlPTSelectionMode
        xlBlanks = 4
        xlButton = 15
        xlDataAndLabel = 0
        xlDataOnly = 2
        xlFirstRow = 256
        xlLabelOnly = 1
        xlOrigin = 3
    End Enum


    Public Enum XlQueryType
        xlADORecordset = 7
        xlDAORecordset = 2
        xlODBCQuery = 1
        xlOLEDBQuery = 5
        xlTextImport = 6
        xlWebQuery = 4
    End Enum


    Public Enum XlQuickAnalysisMode
        xlLensOnly = 0
        xlFormatConditions = 1
        xlRecommendedCharts = 2
        xlTotals = 3
        xlTables = 4
        xlSparklines = 5
    End Enum


    Public Enum XlRangeAutoFormat
        xlRangeAutoFormat3DEffects1 = 13
        xlRangeAutoFormat3DEffects2 = 14
        xlRangeAutoFormatAccounting1 = 4
        xlRangeAutoFormatAccounting2 = 5
        xlRangeAutoFormatAccounting3 = 6
        xlRangeAutoFormatAccounting4 = 17
        xlRangeAutoFormatClassic1 = 1
        xlRangeAutoFormatClassic2 = 2
        xlRangeAutoFormatClassic3 = 3
        xlRangeAutoFormatClassicPivotTable = 31
        xlRangeAutoFormatColor1 = 7
        xlRangeAutoFormatColor2 = 8
        xlRangeAutoFormatColor3 = 9
        xlRangeAutoFormatList1 = 10
        xlRangeAutoFormatList2 = 11
        xlRangeAutoFormatList3 = 12
        xlRangeAutoFormatLocalFormat1 = 15
        xlRangeAutoFormatLocalFormat2 = 16
        xlRangeAutoFormatLocalFormat3 = 19
        xlRangeAutoFormatLocalFormat4 = 20
        xlRangeAutoFormatNone = -4142
        xlRangeAutoFormatPTNone = 42
        xlRangeAutoFormatReport1 = 21
        xlRangeAutoFormatReport10 = 30
        xlRangeAutoFormatReport2 = 22
        xlRangeAutoFormatReport3 = 23
        xlRangeAutoFormatReport4 = 24
        xlRangeAutoFormatReport5 = 25
        xlRangeAutoFormatReport6 = 26
        xlRangeAutoFormatReport7 = 27
        xlRangeAutoFormatReport8 = 28
        xlRangeAutoFormatReport9 = 29
        xlRangeAutoFormatSimple = -4154
        xlRangeAutoFormatTable1 = 32
        xlRangeAutoFormatTable10 = 41
        xlRangeAutoFormatTable2 = 33
        xlRangeAutoFormatTable3 = 34
        xlRangeAutoFormatTable4 = 35
        xlRangeAutoFormatTable5 = 36
        xlRangeAutoFormatTable6 = 37
        xlRangeAutoFormatTable7 = 38
        xlRangeAutoFormatTable8 = 39
        xlRangeAutoFormatTable9 = 40
    End Enum


    Public Enum XlRangeValueDataType
        xlRangeValueDefault = 10
        xlRangeValueMSPersistXML = 12
        xlRangeValueXMLSpreadsheet = 11
    End Enum


    Public Enum XlReferenceStyle
        xlA1 = 1
        xlR1C1 = -4150
    End Enum


    Public Enum XlReferenceType
        xlAbsolute = 1
        xlAbsRowRelColumn = 2
        xlRelative = 4
        xlRelRowAbsColumn = 3
    End Enum


    Public Enum XlRemoveDocInfoType
        xlRDIAll = 99
        xlRDIComments = 1
        xlRDIContentType = 16
        xlRDIDefinedNameComments = 18
        xlRDIDocumentManagementPolicy = 15
        xlRDIDocumentProperties = 8
        xlRDIDocumentServerProperties = 14
        xlRDIDocumentWorkspace = 10
        xlRDIEmailHeader = 5
        xlRDIExcelDataModel = 23
        xlRDIInactiveDataConnections = 19
        xlRDIInkAnnotations = 11
        xlRDIInlineWebExtensions = 21
        xlRDIPrinterPath = 20
        xlRDIPublishInfo = 13
        xlRDIRemovePersonalInformation = 4
        xlRDIRoutingSlip = 6
        xlRDIScenarioComments = 12
        xlRDISendForReview = 7
        xlRDITaskpaneWebExtensions = 22
    End Enum


    Public Enum XlRgbColor
        rgbAliceBlue = 16775408
        rgbAntiqueWhite = 14150650
        rgbAqua = 16776960
        rgbAquamarine = 13959039
        rgbAzure = 16777200
        rgbBeige = 14480885
        rgbBisque = 12903679
        rgbBlack = 0
        rgbBlanchedAlmond = 13495295
        rgbBlue = 16711680
        rgbBlueViolet = 14822282
        rgbBrown = 2763429
        rgbBurlyWood = 8894686
        rgbCadetBlue = 10526303
        rgbChartreuse = 65407
        rgbCoral = 5275647
        rgbCornflowerBlue = 15570276
        rgbCornsilk = 14481663
        rgbCrimson = 3937500
        rgbDarkBlue = 9109504
        rgbDarkCyan = 9145088
        rgbDarkGoldenrod = 755384
        rgbDarkGray = 11119017
        rgbDarkGreen = 25600
        rgbDarkGrey = 11119017
        rgbDarkKhaki = 7059389
        rgbDarkMagenta = 9109643
        rgbDarkOliveGreen = 3107669
        rgbDarkOrange = 36095
        rgbDarkOrchid = 13382297
        rgbDarkRed = 139
        rgbDarkSalmon = 8034025
        rgbDarkSeaGreen = 9419919
        rgbDarkSlateBlue = 9125192
        rgbDarkSlateGray = 5197615
        rgbDarkSlateGrey = 5197615
        rgbDarkTurquoise = 13749760
        rgbDarkViolet = 13828244
        rgbDeepPink = 9639167
        rgbDeepSkyBlue = 16760576
        rgbDimGray = 6908265
        rgbDimGrey = 6908265
        rgbDodgerBlue = 16748574
        rgbFireBrick = 2237106
        rgbFloralWhite = 15792895
        rgbForestGreen = 2263842
        rgbFuchsia = 16711935
        rgbGainsboro = 14474460
        rgbGhostWhite = 16775416
        rgbGold = 55295
        rgbGoldenrod = 2139610
        rgbGray = 8421504
        rgbGreen = 32768
        rgbGreenYellow = 3145645
        rgbGrey = 8421504
        rgbHoneydew = 15794160
        rgbHotPink = 11823615
        rgbIndianRed = 6053069
        rgbIndigo = 8519755
        rgbIvory = 15794175
        rgbKhaki = 9234160
        rgbLavender = 16443110
        rgbLavenderBlush = 16118015
        rgbLawnGreen = 64636
        rgbLemonChiffon = 13499135
        rgbLightBlue = 15128749
        rgbLightCoral = 8421616
        rgbLightCyan = 9145088
        rgbLightGoldenrodYellow = 13826810
        rgbLightGray = 13882323
        rgbLightGreen = 9498256
        rgbLightGrey = 13882323
        rgbLightPink = 12695295
        rgbLightSalmon = 8036607
        rgbLightSeaGreen = 11186720
        rgbLightSkyBlue = 16436871
        rgbLightSlateGray = 10061943
        rgbLightSteelBlue = 14599344
        rgbLightYellow = 14745599
        rgbLime = 65280
        rgbLimeGreen = 3329330
        rgbLinen = 15134970
        rgbMaroon = 128
        rgbMediumAquamarine = 11206502
        rgbMediumBlue = 13434880
        rgbMediumOrchid = 13850042
        rgbMediumPurple = 14381203
        rgbMediumSeaGreen = 7451452
        rgbMediumSlateBlue = 15624315
        rgbMediumSpringGreen = 10156544
        rgbMediumTurquoise = 13422920
        rgbMediumVioletRed = 8721863
        rgbMidnightBlue = 7346457
        rgbMintCream = 16449525
        rgbMistyRose = 14804223
        rgbMoccasin = 11920639
        rgbNavajoWhite = 11394815
        rgbNavy = 8388608
        rgbNavyBlue = 8388608
        rgbOldLace = 15136253
        rgbOlive = 32896
        rgbOliveDrab = 2330219
        rgbOrange = 42495
        rgbOrangeRed = 17919
        rgbOrchid = 14053594
        rgbPaleGoldenrod = 7071982
        rgbPaleGreen = 10025880
        rgbPaleTurquoise = 15658671
        rgbPaleVioletRed = 9662683
        rgbPapayaWhip = 14020607
        rgbPeachPuff = 12180223
        rgbPeru = 4163021
        rgbPink = 13353215
        rgbPlum = 14524637
        rgbPowderBlue = 15130800
        rgbPurple = 8388736
        rgbRed = 255
        rgbRosyBrown = 9408444
        rgbRoyalBlue = 14772545
        rgbSalmon = 7504122
        rgbSandyBrown = 6333684
        rgbSeaGreen = 5737262
        rgbSeashell = 15660543
        rgbSienna = 2970272
        rgbSilver = 12632256
        rgbSkyBlue = 15453831
        rgbSlateBlue = 13458026
        rgbSlateGray = 9470064
        rgbSnow = 16448255
        rgbSpringGreen = 8388352
        rgbSteelBlue = 11829830
        rgbTan = 9221330
        rgbTeal = 8421376
        rgbThistle = 14204888
        rgbTomato = 4678655
        rgbTurquoise = 13688896
        rgbViolet = 15631086
        rgbWheat = 11788021
        rgbWhite = 16777215
        rgbWhiteSmoke = 16119285
        rgbYellow = 65535
        rgbYellowGreen = 3329434
    End Enum


    Public Enum XlRobustConnect
        xlAlways = 1
        xlAsRequired = 0
        xlNever = 2
    End Enum


    Public Enum XlRoutingSlipDelivery
        xlAllAtOnce = 2
        xlOneAfterAnother = 1
    End Enum


    Public Enum XlRoutingSlipStatus
        xlNotYetRouted = 0
        xlRoutingComplete = 2
        xlRoutingInProgress = 1
    End Enum


    Public Enum XlRowCol
        xlColumns = 2
        xlRows = 1
    End Enum


    Public Enum XlRunAutoMacro
        xlAutoActivate = 3
        xlAutoClose = 2
        xlAutoDeactivate = 4
        xlAutoOpen = 1
    End Enum


    Public Enum XlSaveAction
        xlDoNotSaveChanges = 2
        xlSaveChanges = 1
    End Enum


    Public Enum XlSaveAsAccessMode
        xlExclusive = 3
        xlNoChange = 1
        xlShared = 2
    End Enum


    Public Enum XlSaveConflictResolution
        xlLocalSessionChanges = 2
        xlOtherSessionChanges = 3
        xlUserResolution = 1
    End Enum


    Public Enum XlScaleType
        xlScaleLinear = -4132
        xlScaleLogarithmic = -4133
    End Enum


    Public Enum XlSearchDirection
        xlNext = 1
        xlPrevious = 2
    End Enum


    Public Enum XlSearchOrder
        xlByColumns = 2
        xlByRows = 1
    End Enum


    Public Enum XlSearchWithin
        xlWithinSheet = 1
        xlWithinWorkbook = 2
    End Enum


    Public Enum XlSeriesNameLevel
        xlSeriesNameLevelAll = -1
        xlSeriesNameLevelCustom = -2
        xlSeriesNameLevelNone = -3
    End Enum


    Public Enum XlSheetType
        xlChart = -4109
        xlDialogSheet = -4116
        xlExcel4IntlMacroSheet = 4
        xlExcel4MacroSheet = 3
        xlWorksheet = -4167
    End Enum


    Public Enum XlSheetVisibility
        xlSheetHidden = 0
        xlSheetVeryHidden = 2
        xlSheetVisible = -1
    End Enum


    Public Enum XlSizeRepresents
        xlSizeIsArea = 1
        xlSizeIsWidth = 2
    End Enum


    Public Enum XlSlicerCacheType
        xlSlicer = 1
        xlTimeline = 2
    End Enum


    Public Enum XlSlicerCrossFilterType
        xlSlicerCrossFilterHideButtonsWithNoData = 4
        xlSlicerCrossFilterShowItemsWithDataAtTop = 2
        xlSlicerCrossFilterShowItemsWithNoData = 3
        xlSlicerNoCrossFilter = 1
    End Enum


    Public Enum XlSlicerSort
        xlSlicerSortAscending = 2
        xlSlicerSortDataSourceOrder = 1
        xlSlicerSortDescending = 3
    End Enum


    Public Enum XlSmartTagControlType
        xlSmartTagControlActiveX = 13
        xlSmartTagControlButton = 6
        xlSmartTagControlCheckbox = 9
        xlSmartTagControlCombo = 12
        xlSmartTagControlHelp = 3
        xlSmartTagControlHelpURL = 4
        xlSmartTagControlImage = 8
        xlSmartTagControlLabel = 7
        xlSmartTagControlLink = 2
        xlSmartTagControlListbox = 11
        xlSmartTagControlRadioGroup = 14
        xlSmartTagControlSeparator = 5
        xlSmartTagControlSmartTag = 1
        xlSmartTagControlTextbox = 10
    End Enum


    Public Enum XlSmartTagDisplayMode
        xlButtonOnly = 2
        xlDisplayNone = 1
        xlIndicatorAndButton = 0
    End Enum


    Public Enum XlSortDataOption
        xlSortNormal = 0
        xlSortTextAsNumbers = 1
    End Enum


    Public Enum XlSortMethod
        xlPinYin = 1
        xlStroke = 2
    End Enum


    Public Enum XlSortMethodOld
        xlCodePage = 2
        xlSyllabary = 1
    End Enum


    Public Enum XlSortOn
        SortOnCellColor = 1
        SortOnFontColor = 2
        SortOnIcon = 3
        SortOnValues = 0
    End Enum


    Public Enum XlSortOrder
        xlAscending = 1
        xlDescending = 2
    End Enum


    Public Enum XlSortOrientation
        xlSortColumns = 1
        xlSortRows = 2
    End Enum


    Public Enum XlSortType
        xlSortLabels = 2
        xlSortValues = 1
    End Enum


    Public Enum XlSourceType
        xlSourceAutoFilter = 3
        xlSourceChart = 5
        xlSourcePivotTable = 6
        xlSourcePrintArea = 2
        xlSourceQuery = 7
        xlSourceRange = 4
        xlSourceSheet = 1
        xlSourceWorkbook = 0
    End Enum


    Public Enum XlSpanishModes
        xlSpanishTuteoAndVoseo = 1
        xlSpanishTuteoOnly = 0
        xlSpanishVoseoOnly = 2
    End Enum


    Public Enum XlSparklineRowCol
        SparklineColumnsSquare = 2
        SparklineNonSquare = 0
        SparklineRowsSquare = 1
    End Enum


    Public Enum XlSparkScale
        xlSparkScaleCustom = 3
        xlSparkScaleGroup = 1
        xlSparkScaleSingle = 2
    End Enum


    Public Enum XlSparkType
        xlSparkColumn = 2
        xlSparkColumnStacked100 = 3
        xlSparkLine = 1
    End Enum


    Public Enum XlSpeakDirection
        xlSpeakByColumns = 1
        xlSpeakByRows = 0
    End Enum


    Public Enum XlSpecialCellsValue
        xlErrors = 16
        xlLogical = 4
        xlNumbers = 1
        xlTextValues = 2
    End Enum


    Public Enum XlStdColorScale
        ColorScaleBlackWhite = 3
        ColorScaleGYR = 2
        ColorScaleRYG = 1
        ColorScaleWhiteBlack = 4
    End Enum


    Public Enum XlSubscribeToFormat
        xlSubscribeToPicture = -4147
        xlSubscribeToText = -4158
    End Enum


    Public Enum XlSubtototalLocationType
        xlAtBottom = 2
        xlAtTop = 1
    End Enum


    Public Enum XlSummaryColumn
        xlSummaryOnLeft = -4131
        xlSummaryOnRight = -4152
    End Enum


    Public Enum XlSummaryReportType
        xlStandardSummary = 1
        xlSummaryPivotTable = -4148
    End Enum


    Public Enum XlSummaryRow
        xlSummaryAbove = 0
        xlSummaryBelow = 1
    End Enum


    Public Enum XlTableStyleElementType
        xlBlankRow = 19
        xlColumnStripe1 = 7
        xlColumnStripe2 = 8
        xlColumnSubheading1 = 20
        xlColumnSubheading2 = 21
        xlColumnSubheading3 = 22
        xlFirstColumn = 3
        xlFirstHeaderCell = 9
        xlFirstTotalCell = 11
        xlGrandTotalColumn = 4
        xlGrandTotalRow = 2
        xlHeaderRow = 1
        xlLastColumn = 4
        xlLastHeaderCell = 10
        xlLastTotalCell = 12
        xlPageFieldLabels = 26
        xlPageFieldValues = 27
        xlRowStripe1 = 5
        xlRowStripe2 = 6
        xlRowSubheading1 = 23
        xlRowSubheading2 = 24
        xlRowSubheading3 = 25
        xlSlicerHoveredSelectedItemWithData = 33
        xlSlicerHoveredSelectedItemWithNoData = 35
        xlSlicerHoveredUnselectedItemWithData = 32
        xlSlicerHoveredUnselectedItemWithNoData = 34
        xlSlicerSelectedItemWithData = 30
        xlSlicerSelectedItemWithNoData = 31
        xlSlicerUnselectedItemWithData = 28
        xlSlicerUnselectedItemWithNoData = 29
        xlSubtotalColumn1 = 13
        xlSubtotalColumn2 = 14
        xlSubtotalColumn3 = 15
        xlSubtotalRow1 = 16
        xlSubtotalRow2 = 17
        xlSubtotalRow3 = 18
        xlTimelinePeriodLabels1 = 38
        xlTimelinePeriodLabels2 = 39
        xlTimelineSelectedTimeBlock = 40
        xlTimelineSelectedTimeBlockSpace = 42
        xlTimelineSelectionLabel = 36
        xlTimelineTimeLevel = 37
        xlTimelineUnselectedTimeBlock = 41
        xlTotalRow = 2
        xlWholeTable = 0
    End Enum


    Public Enum XlTabPosition
        xlTabPositionFirst = 0
        xlTabPositionLast = 1
    End Enum


    Public Enum XlTextParsingType
        xlDelimited = 1
        xlFixedWidth = 2
    End Enum


    Public Enum XlTextQualifier
        xlTextQualifierDoubleQuote = 1
        xlTextQualifierNone = -4142
        xlTextQualifierSingleQuote = 2
    End Enum


    Public Enum XlTextVisualLayoutType
        xlTextVisualLTR = 1
        xlTextVisualRTL = 2
    End Enum


    Public Enum XlThemeColor
        xlThemeColorAccent1 = 5
        xlThemeColorAccent2 = 6
        xlThemeColorAccent3 = 7
        xlThemeColorAccent4 = 8
        xlThemeColorAccent5 = 9
        xlThemeColorAccent6 = 10
        xlThemeColorDark1 = 1
        xlThemeColorDark2 = 3
        xlThemeColorFollowedHyperlink = 12
        xlThemeColorHyperlink = 11
        xlThemeColorLight1 = 2
        xlThemeColorLight2 = 4
    End Enum


    Public Enum XlThemeFont
        xlThemeFontMajor = 2
        xlThemeFontMinor = 1
        xlThemeFontNone = 0
    End Enum


    Public Enum XlThreadMode
        xlThreadModeAutomatic = 0
        xlThreadModeManual = 1
    End Enum


    Public Enum XlTickLabelOrientation
        xlTickLabelOrientationAutomatic = -4105
        xlTickLabelOrientationDownward = -4170
        xlTickLabelOrientationHorizontal = -4128
        xlTickLabelOrientationUpward = -4171
        xlTickLabelOrientationVertical = -4166
    End Enum


    Public Enum XlTickLabelPosition
        xlTickLabelPositionHigh = -4127
        xlTickLabelPositionLow = -4134
        xlTickLabelPositionNextToAxis = 4
        xlTickLabelPositionNone = -4142
    End Enum


    Public Enum XlTickMark
        xlTickMarkCross = 4
        xlTickMarkInside = 2
        xlTickMarkNone = -4142
        xlTickMarkOutside = 3
    End Enum


    Public Enum XlTimelineLevel
        xlTimelineLevelYears = 0
        xlTimelineLevelQuarters = 1
        xlTimelineLevelMonths = 2
        xlTimelineLevelDays = 3
    End Enum


    Public Enum XlTimePeriods
        xlLast7Days = 2
        xlLastMonth = 5
        xlLastWeek = 4
        xlNextMonth = 8
        xlNextWeek = 7
        xlThisMonth = 9
        xlThisWeek = 3
        xlToday = 0
        xlTomorrow = 6
        xlYesterday = 1
    End Enum


    Public Enum XlTimeUnit
        xlDays = 0
        xlMonths = 1
        xlYears = 2
    End Enum


    Public Enum XlToolbarProtection
        xlNoButtonChanges = 1
        xlNoChanges = 4
        xlNoDockingChanges = 3
        xlNoShapeChanges = 2
        xlToolbarProtectionNone = -4143
    End Enum


    Public Enum XlTopBottom
        xlTop10Bottom = 0
        xlTop10Top = 1
    End Enum


    Public Enum XlTotalsCalculation
        xlTotalsCalculationAverage = 2
        xlTotalsCalculationCount = 3
        xlTotalsCalculationCountNums = 4
        xlTotalsCalculationCustom = 9
        xlTotalsCalculationMax = 6
        xlTotalsCalculationMin = 5
        xlTotalsCalculationNone = 0
        xlTotalsCalculationStdDev = 7
        xlTotalsCalculationSum = 1
        xlTotalsCalculationVar = 8
    End Enum


    Public Enum XlTrendlineType
        xlExponential = 5
        xlLinear = -4132
        xlLogarithmic = -4133
        xlMovingAvg = 6
        xlPolynomial = 3
        xlPower = 4
    End Enum


    Public Enum XlUnderlineStyle
        xlUnderlineStyleDouble = -4119
        xlUnderlineStyleDoubleAccounting = 5
        xlUnderlineStyleNone = -4142
        xlUnderlineStyleSingle = 2
        xlUnderlineStyleSingleAccounting = 4
    End Enum


    Public Enum XlUpdateLinks
        xlUpdateLinksAlways = 3
        xlUpdateLinksNever = 2
        xlUpdateLinksUserSetting = 1
    End Enum


    Public Enum XlVAlign
        xlVAlignBottom = -4107
        xlVAlignCenter = -4108
        xlVAlignDistributed = -4117
        xlVAlignJustify = -4130
        xlVAlignTop = -4160
    End Enum


    Public Enum XlWBATemplate
        xlWBATChart = -4109
        xlWBATExcel4IntlMacroSheet = 4
        xlWBATExcel4MacroSheet = 3
        xlWBATWorksheet = -4167
    End Enum


    Public Enum XlWebFormatting
        xlWebFormattingAll = 1
        xlWebFormattingNone = 3
        xlWebFormattingRTF = 2
    End Enum


    Public Enum XlWebSelectionType
        xlAllTables = 2
        xlEntirePage = 1
        xlSpecifiedTables = 3
    End Enum


    Public Enum XlWindowState
        xlMaximized = -4137
        xlMinimized = -4140
        xlNormal = -4143
    End Enum


    Public Enum XlWindowType
        xlChartAsWindow = 5
        xlChartInPlace = 4
        xlClipboard = 3
        xlInfo = -4129
        xlWorkbook = 1
    End Enum


    Public Enum XlWindowView
        xlNormalView = 1
        xlPageBreakPreview = 2
        xlPageLayoutView = 3
    End Enum


    Public Enum XlXLMMacroType
        xlCommand = 2
        xlFunction = 1
        xlNotXLM = 3
    End Enum


    Public Enum XlXmlExportResult
        xlXmlExportSuccess = 0
        xlXmlExportValidationFailed = 1
    End Enum


    Public Enum XlXmlImportResult
        xlXmlImportElementsTruncated = 1
        xlXmlImportSuccess = 0
        xlXmlImportValidationFailed = 2
    End Enum


    Public Enum XlXmlLoadOption
        xlXmlLoadImportToList = 2
        xlXmlLoadMapXml = 3
        xlXmlLoadOpenXml = 1
        xlXmlLoadPromptUser = 0
    End Enum


    Public Enum XlYesNoGuess
        xlGuess = 0
        xlNo = 2
        xlYes = 1
    End Enum


    Public Enum XlModelChangeSource
        xlChangeByExcel = 0
        xlChangeByPowerPivotAddIn = 1
    End Enum
End Module

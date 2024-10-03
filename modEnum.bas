Attribute VB_Name = "modEnum"
Option Explicit

''''''''''''''''
'After adding or amending any entry
'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
''''''''''''''''

Public Enum csMenu

    csLogOn = 1001
    csLogOff = 1002
    csResetLastUsed = 1003
    csShowErrorLog = 1004
    csBatches = 1005
    csHaematologyFilm = 1006
    csHaematologyImmuno = 1007
    csHIV = 1008
    csAutoimmuneProfile = 1009
    csSearch = 1010
    csLists = 1011
    csLocations = 1012
    csDefaults = 1013
    cs24HrUrine = 1014
    csStatistics = 1015
    csGeneralStats = 1016
    csExternalStats = 1017
    csStockControl = 1018
    csAddReagent = 1019
    csAdministerReagents = 1020
    csCheckStock = 1021
    csStockConfirmation = 1022
    csStockHistory = 1023
    csUpdateStock = 1024
    csHelp = 1025
    csWindowsHelp = 1026

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csNetAcquireUserManual = 1027
    csTechnicalAssistance = 1028
    csAbout = 1029
    csSystemOptions = 1030
    csUserOptions = 1031
    csLanguage = 1032
    csEnglish = 1033
    csPortuguese = 1034
    csGotoCustomSoftwareWebsite = 1035
    csViewResults = 1036
    csAddParameterAgeSpecificRange = 1037
    csGeneral = 1038
    csEditGeneralChemistry = 1039
    csAuditTrail = 1040
    csFaeces = 1041
    csFaecesCulture = 1042
    csQC = 1043
    csViewWardEnquires = 1044
    csView = 1045
    csComment = 1046

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csBatchReporting = 1047
    csOrganismGroups = 1048
    csOrganismNames = 1049
    csStool = 1050
    csXLD = 1051
    csDCA = 1052
    csSMAC = 1053
    csXLDSub = 1054
    csDCASub = 1055
    csPreston = 1056
    csCCDA = 1057
    csHeadOfFirm = 1058
    csArchive = 1059
    csInvestigationRequested = 1060
    csAddSequenceParameter = 1061
    csOtherTests = 1062
    csWidalBrucellaProt = 1063
    csCD4 = 1064
    csTM = 1065
    csMacroscopicAmount = 1066
    csMedia = 1067
    csCSFStains = 1068
    csFile = 1069
    csHormonesTumourMarkers = 1070
    csTrainingLog = 1071
    csQuarterly = 1072
    csTotals = 1073
    csInventory = 1074

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csSuppliers = 1075
    csSuppliersList = 1076
    csSpecimenStockTracking = 1077
    csStockInStockOut = 1078
    csStockRelocation = 1079
    '1080
    csItemReportStockLocator = 1081
    csEquipmentReportStockLocator = 1082
    csItems = 1083
    csItemsList = 1084
    csEquipment = 1085
    csEquipmentList = 1086
    csStockLocations = 1087
    csUtilities = 1088
    csGeneral2 = 1089
    csFastingMenu = 1090

    csPrintHandlerLocations = 1091
    csPrintHandlerRoles = 1092
    csViewOutstanding = 1093
    csHaemNormalRanges = 1094
    csHideOutstanding = 1095

    csReflexTests = 1096
    csTestCodeMapping = 1097
    csNetAcquire = 2000
End Enum

''''''''''''''''
'After adding or amending any entry
'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
''''''''''''''''

Public Enum csLabels
    csTo1 = 2001
    csFrom = 2002
    csSampleID1 = 2003
    csPatient = 2004
    csHrs = 2005
    csNitrogen = 2007
    csPotassium = 2008
    csSodium = 2009
    csUrea = 2010
    csCalcium = 2011
    csPhosphorus = 2012
    csTProt = 2013
    csMagnesium = 2014
    csChloride = 2015
    csCreatinine = 2016
    csVolume = 2017
    csRunDateTime = 2018
    csName = 2019
    csDoB = 2020
    csChart = 2021
    csSampleID = 2022
    csmmol = 2023
    csMl = 2024

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csBetweenDates = 2025
    csOnlyAbnormals = 2026
    csNormalRanges = 2027
    csLow = 2028
    csHigh = 2029
    csFlagRanges = 2030
    csWardDateSampleIDResult = 2031
    csNameCode = 2032
    csCodeSendToAddressPhoneFax = 2033
    csAgeFromYMDAgeToYMD = 2034
    csNetAcquireLaboratoryInformationSystemVersion = 2035
    csSpecimenChartNameESRReticIMAsot = 2036
    csSampleIdNameResultCommentWardClinicianGp = 2037
    csSampleIDGlucoseProteinBenceJonesFatGlobulesPregnancyHCGSG = 2038
    csSpecimenChartNameHiv = 2039
    csSampleIDpHProteinGlucoseKetonesUroBiliBloodHbWCCRCCCrystalsCastsMiscellaneousMiscellaneousMiscellaneous = 2040
    csSampleIDNOPASDoBNameClinicianPhPCO2PO2HCO3BEO2SATTotCO2 = 2041
    csRangeAgeFromAgeTo = 2042
    csHours = 2043
    csParameterLowHigh = 2044
    csLongNameShortNameCodeBarCodeImmunoCodeUnitsDecPlPrintableKnowntoAnalyserAnalyserCodeInUseEndofDayViewonWardPrintRefRangeResultOptionsIndex = 2045
    csLongNameShortNameCodeAgeFromAgeToNormalMaleLowNormalMaleHighNormalFemaleLowNormalFemaleHighFlagMaleLowFlagMaleHighFlagFemaleLowFlagFemaleHighCodeDecPlAgeFromDays = 2046
    csLongNameShortNameCodePlausibleLowPlausibleHighAutoValLowAutoValHighDoDeltaDeltaAbsoluteIndex = 2047
    csLongNameShortNameCodeOldLipaemicIctericHaemolysedSlightlyHaemalysedGrosslyHaemalysed = 2048
    csTestNameLongName = 2049

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csParameterTests = 2050
    csAgeFromYMDAgeToYMD2 = 2051
    csInUseCodeTitleForeNameSurnameClinicianWard = 2052
    csParameterResultUnits1 = 2053
    csSourceSamplesTestsTS1 = 2054
    csWardsCliniciansGPsDeviceDestination = 2055
    csOutstanding1 = 2056
    csNoAmendmentsreceivedfor = 2057
    csNoNewRecordsreceivedfor = 2058
    csSampleIDNOPASDoBNameClinicianWardGPPtINRApptDDimerFib = 2059
    csSampleIDNameChartGPWardClinicianTests = 2060
    csSampleIDNameDoBWardClinicianGpUreaNaKClGluBiliALPGGTALTAMYCPKASTLDHCaPhosMgUra = 2061
    csAbsRefRangeDiff = 2062
    csFBCResultRefRange = 2063
    csTestResultUnitsRefRangeHLVPCPALComment = 2064
    csParameterResultUnits = 2065
    csTestResultUnitsRefRangeHLVPComment = 2066
    csParameterResultUnitsRefRangeFlagVP = 2067
    csTestResultUnitsRefRangeNEIVPPCComment = 2068
    csTestNumberTestNameResultNormalRangeUnitsSendtoSentDateRetDateSapCode = 2069
    csSampleIDNameDoBWardClinician = 2070
    csSampleIdHBCEGI = 2071
    csFormControlWhatIsTabindexVisible = 2072
    csDateName = 2073
    csSampleIDTimeSerummmolL = 2074
    csNameDateofBirth = 2075
    csSampleIDDateTimeSerumHLUrineHL = 2076
    csCodeInUseGPNameAddrAddrTitleForeNameSurNamePhoneFAXPracticeCFHealthlink = 2077
    csNameTestsSamplesTSTestWorkloadSampleWorkloadLabTestsLabSamples = 2078
    csUserFormOptions = 2079
    csSampleIDNOPASDoBNameClinicianGpWardComment = 2080
    csSpecimenChartNameWardGpClinician = 2081
    csNameSampleIDANAASMAAMAGPC = 2082
    csInUseOperatorNameCodeMemberOfLogOffDelayPasswordDiscipline = 2083
    csDescriptionContent = 2084
    csPanelNameBarCode = 2085
    csInUseCodeTextFAXPrinterAddress = 2086
    csChartNameDoBSexAddressWardClinicianHospital = 2087
    csBHCEIGXRMSHICYRunDateSampleIDNopasChartNameDoBAgeSexAddressWardClinicianGPHospitalSampSurnameForname = 2088

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSourceTotalSamplesCoagSamplesBioSamplesHaemSamples = 2089
    csSampleIdNameDepartmentPrintTimeReportNoInitiatorPagesPrinter = 2090
    csSourceSamplesSampleTestsTestTS = 2091
    csNumberNameNormalRangeUnitsAddressSendTo = 2092
    csDestinationTestTotal = 2093
    csDateTimeHBCIGXMEPhonedToCommentPhonedBy = 2094
    csSampleIdPatientTestSentToSent = 2095
    csChartNoDateTimeSampleIDNameDobOperatorDetails = 2096
    csSampleIDPatientNameChartAnalyteResultExtLab = 2097
    csParameterResultV = 2098
    csParameterResultVComm = 2099
    csReagentNameAmountUsernameDateAddedComment = 2100
    csReagentNameInStockMinStockDept = 2101
    csReagentNameTestNameUnits = 2102
    csCodeReagentNameUnits = 2103
    csDateTimeMessage = 2104
    csAnalyteTypeRawResultRunDateTimeDateTimeArchivedArchivedBy = 2105
    csLaboratoryInformationSystem = 2106
    csOccultBlood = 2107
    csHBA1CFerritinPSA = 2108
    csHistologyCytology = 2109
    csAutoRefresh = 2110
    csSampleIDAnalyser = 2111
    csClickonHeadingtoSort = 2112
    csNotPrintedOutstandingRequests = 2113

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csResultsFor = 2114
    csNoPreviousCoagDetails = 2115
    csTestCodeTestName = 2116
    csOutstanding = 2117
    csOvae = 2118
    csAntibiotics = 2119
    csErrors = 2120
    csControlName = 2121
    csImmunohaem = 2122
    csListOfHospitals = 2123
    csMorphology = 2124
    csTotalsforBiochemistry = 2125
    csSystem = 2126
    csViewonWard = 2127
    csHaematologySummary = 2128
    csSamples = 2129
    csTests = 2130
    csLaboratoryTotals = 2131
    csEnterFaxNumber = 2132
    csEnterSurname = 2133
    csFull = 2134
    csPractice = 2135
    csCompiledReport = 2136
    csAlternativeCode = 2137
    csPhoneLogHistoryforSampleID = 2138

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csDataPoints = 2139
    csFalse = 2140
    csTrue = 2141
    csDescription = 2142
    csContent = 2143
    csDoyouwanttoresettheDefaultDisplayto = 2144
    csFBC1 = 2145
    csTestCode = 2146
    csPanels = 2147
    csFluids = 2148
    csShort = 2149
    csLong = 2150
    '2151
    csMemberOf = 2152
    csAutoLogOffin = 2153
    csMinutes = 2154
    csAutoSelectFrom = 2155
    csPathogens = 2156
    csDateFrom = 2157
    csDateTo = 2158
    csMappedToPrinterName = 2159
    csPrinterName = 2160
    csSystemInfo = 2161
    csVersion = 2162
    csLongName = 2163
    csShortName = 2164
    csCode = 2165
    csSampleType = 2166
    csRecalculate = 2167

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csDeSelectRange = 2168
    csOverRange = 2169
    csUnderRange = 2170
    csAll = 2171
    csMale = 2172
    csFemale = 2173
    csBoth1 = 2174
    csApplicationTitle = 2175
    csDateofBirth = 2176
    csAnalyser = 2177
    csImmunocode = 2178
    csCategory1 = 2179
    csFax1 = 2180
    csPhone = 2181
    csAddress = 2182
    csPrint = 2183
    csSerum = 2184
    csurine = 2185
    csCSF = 2186
    csOrganisms = 2187
    csPrintList = 2188
    csMoveUp = 2189
    csMoveDown = 2190
    csAdd = 2191
    csRemove = 2192

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSecondaryList = 2193
    csBadResults = 2194
    csSite = 2195
    csGeneric = 2196
    csOptions = 2197
    csBiochemistry = 2198
    csHaematology = 2199
    csCoagulation = 2200
    csImmunology = 2201
    csEndocrinology = 2202
    csBloodGas = 2203
    csExternals = 2204
    csDemographics = 2205
    csRepeats = 2206
    csIncomplete = 2207
    csOrdered = 2208
    csDiscipline = 2209
    csToday = 2210
    csESR = 2211
    csFBC = 2212
    csNegative = 2213
    csPositive = 2214
    csInconclusive = 2215
    csNotSeen = 2216
    csPresent = 2217
    csLaboratorialPrinting = 2218
    csWardPrinting = 2219
    csFaxing = 2220
    csTheseNormalRangesonlyapplywhentheAgeSexRelatedOptionisDisabled = 2221
    csFaecesStats = 2222
    csCon = 2223
    csUrineStats = 2224
    csFrozenSection = 2225
    csCytologyStats = 2226


    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csHistologyStats = 2227
    csAgeRanges = 2228
    csBaud = 2229
    csParity = 2230
    csStatsbySource = 2232
    csOdd = 2331
    csEven = 2333
    csNone = 2334
    csMinimum888Maximum888Mean888888SD88888 = 2335
    csLessthan5Datapoints = 2336
    csMinimum = 2337
    csMaximum = 2338
    csMean = 2339
    csSD = 2340
    csCV = 2341
    csAbnormalReportfor = 2342
    csTestName1 = 2343
    csAnalyserCode = 2344
    csAvailableAntibiotics = 2345
    csPrimaryList = 2346

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csRotaVirus = 2347
    csNumberFrom = 2348
    csUrineSample = 2349
    csUrineRequests = 2350
    csUrinaryHCG = 2351
    csFatGlobules = 2352
    csPregnancy = 2353
    'csNumbers = 2354 - duplicated see 6017
    csStopNumber = 2355
    csStartNumber = 2356
    csLac = 2357
    csPur = 2358
    csEndOfDay = 2359
    csAmendAgeRange = 2360
    csKnowntoAnalyser = 2361
    csValue = 2362
    csNumberofDecimalPlaces = 2363
    csBarCode = 2364
    csClinician = 2365
    csValidateallSelectedrows = 2366
    csRunDate1 = 2367
    csSortedby = 2368
    csPatientName = 2369
    csLaboratory_At = 2370
    csLaboratory = 2371

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csQCB = 2372
    csListofBiochemistryPlausibleRanges = 2373
    csCodesUnitsandPrecision = 2374
    csNormalandFlagRanges = 2375
    csPlausibleAutoValandDeltaRanges = 2376
    csMasks = 2377
    csPrintSequence = 2378
    csDonotPrintResultif = 2379
    csExporting = 2380
    csHospital = 2381
    csAge = 2382
    csEnterUnitsfor = 2383
    csAgeSpecificRangefor = 2384
    csRangefor = 2385
    csfor = 2386
    csTestsSample = 2387
    csTotalTests1 = 2388
    csTotalNumberofTests = 2389
    csBetween = 2390
    csTotalSamples = 2391
    csNumber = 2392
    csTestsperSample = 2393
    csto = 2394
    csSurname = 2395
    csForeName = 2396

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csTitle = 2397
    csListofClinicians = 2398
    csSpecificsAppliestoallageranges = 2399
    csPrintable = 2400
    csDecimalPoints = 2401
    csGraph = 2402
    csWard = 2403
    csGP = 2404
    csTotalsbetween = 2405
    csand = 2406
    csTotalabove = 2407
    csExcel = 2408
    csListof = 2409
    csText = 2410
    csCytology = 2411
    csHistology = 2412
    csListType = 2413
    csEndocrinologyTotals = 2414
    csResults = 2415
    csMicro = 2416
    csSendCopyTo = 2417
    csSendOriginalTo = 2418
    csUseDefault = 2419
    csPrinter = 2420
    csFax = 2421
    csUrinaryProtein1 = 2422
    csCreatinineClearance1 = 2423
    csUrinaryProtein = 2424
    csUrineCreatinine = 2425

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csSerumCreatinine = 2426
    csTotalUrinaryVolume = 2427
    csUrineSampleID = 2428
    csSerumSampleID = 2429
    csRunDatesBetween = 2430
    csUrineComments = 2431
    csCSComments = 2432
    csCytologyComments = 2433
    csImmunologyResults1 = 2434
    csHistologyComments = 2435
    csImmunologyComments = 2436
    '2437
    csMicroComments = 2438
    csDemographicComments = 2439
    csSemenComments = 2440
    csBloodGasComments = 2441
    csEndocrinologyComments = 2442
    csBiochemistryComments = 2443
    csHaematologyComments = 2444
    csCoagulationComments = 2445
    csCoagulationSummary = 2446
    csCriticalProblem = 2447
    csTotalRecords = 2448
    csMicrobiology = 2449
    csDailyReportfor = 2450

    csEndofdayreportfor = 2451
    csAE = 2452
    csMRN1 = 2453
    csLeucocytes = 2454
    csBasophils = 2455
    csEosinophils = 2456
    csMonocytes = 2457
    csLymphocytes = 2458

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csNeutrophils = 2459
    csDifferentiatedBy = 2460
    csWBC = 2461
    csCount1 = 2462
    csCount2 = 2463
    csKey = 2464
    csMostRecentlyUsedNumbers = 2465
    csFindLastRecord = 2466
    csPhoneLog = 2467
    csSampleIdentificationNumber = 2468
    csFindNextRelevantSample = 2469
    csFindPreviousRelevantSample = 2470
    csPatientDetails = 2471
    csSearchusingName = 2472
    csSearchusingDateofBirth = 2473
    cssex = 2474
    csFirstName = 2475

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csChartMrnNumber = 2476
    csPasNumber = 2477
    csAENumber = 2478
    csUrgent = 2479
    csResultsNeededUrgently = 2480
    csSampleDate = 2481
    csTestNameCodeUnits = 2482
    csSampleIDResult = 2483
    csTestNameResult = 2484
    csTypeAnalyte = 2485
    csCodeText = 2486
    csSourceTests = 2487
    csCategory = 2488
    csClinicalDetails = 2489
    csPatientLocationInformation = 2490
    csChartNo = 2491
    csClicktochangeLocation = 2492
    csAandE = 2493
    csusername = 2494
    csTodaysDate = 2495
    csDemographicCheck = 2496
    csOrder = 2497
    csOrderTestsforSample = 2498
    csTransfusionDetails = 2499
    csHighlightitemtobemovedthenclickappropriatearrow = 2500

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csHistory = 2501
    csUnvalidatedReport = 2502
    csViewUnvalidatedReports = 2503
    csReports = 2505
    csViewPrintedFaxedReports = 2506
    csOrderExternalTest = 2507
    csSpecimenCondition = 2508
    csPatientMedicalCondition = 2509
    csFilm = 2510
    csBadResult = 2511
    csAntistreptococcalAntibodyTitres = 2512
    csRF = 2513
    csErythrocyteSedimentationRate = 2514
    csPlatelets = 2515
    csMeanPlateletVolume = 2516
    csBiochemistryTotals = 2517
    csListofEndocrinologyPlausibleRanges = 2518
    csListofImmunologyPlausibleRanges = 2519
    csDonotprinttheresultifthesampleis = 2520
    csPrintoutText = 2521
    csTriglyceride = 2522
    csCholesterol = 2523
    csChooseyourPrinterfromthelistbelowNOTEFornormaloperationchooseAutomaticSelection = 2524
    csToolTipText = 2525

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csGeneralChemistry = 2526
    csPlotbetween = 2527
    csST = 2528
    csRunDate = 2529
    csRunTime = 2530
    csRefRanges = 2531
    csNoofResults = 2532
    csCumulativeReportfromBiochemistryDept = 2533
    csRecords = 2534
    csCumulativeReportfromCoagulationDept = 2535
    csEndOfReport = 2536
    csCumulativeReportfromEndocrinologyDept = 2537
    csCumulativeReportfromExternalTests = 2538
    csCumulativeReportfromHaematologyDept = 2539
    csCumulativeReportfromImmunologyDept = 2540
    csUseaFuzzySearch = 2541
    csUseExactDateofBirth = 2542
    csYear = 2543
    csYears1 = 2544

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csddmmyyorddmmyyyy = 2545
    csAddress1 = 2546
    csAddress2 = 2547
    csMPV = 2548
    csPlt = 2549
    csMeanPeroxidaseIndex = 2550
    csAnalyserDifferential = 2551
    csWhiteBloodCount = 2552
    csMPXI = 2553
    csLI = 2554
    csWOC = 2555
    csWIC = 2556
    csAddressLine2 = 2557
    csAddressLine1 = 2558

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csGpID = 2559
    csOutstandingTests = 2560
    csTimeofSample = 2561
    csTimeSampleReceived = 2562
    csReceived1 = 2563
    csPreviousDay = 2564
    csNextDay = 2565
    csSettoToday = 2566
    cssample = 2567
    csOutofHours = 2568
    csRoutine = 2569

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csWarfarin = 2570
    csTestsRepeated = 2571
    csRandomSample = 2572
    csCondition = 2573
    csFLAGS = 2574
    csOutstandingExternals = 2575
    csAlreadyPrinted = 2576
    csMRU = 2577
    csFaxResult = 2578
    csASOT = 2579
    csDate = 2580
    csBiochemistryResults = 2581
    csCoagulationResults = 2582
    csEndocrinologyResults = 2583
    csBloodGasResults = 2584
    csImmunologyResults = 2585
    csBoth = 2586
    csMRN = 2587
    csResultsPhoned = 2588
    csNeut = 2589
    csLymph = 2590
    csMono = 2591

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csEos = 2592
    csBas = 2593
    csLuc = 2594
    csRBC = 2595
    csHgb = 2596
    csHCT = 2597
    csMCV = 2598
    csHDW = 2599
    csMCH = 2600
    csMCHC = 2601
    csCHCM = 2602
    csRDW = 2603
    csNRBC = 2604
    csHYPO = 2605
    csTestName = 2606
    csTestNumber = 2607
    csDeleted = 2608

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSaving = 2609
    csConfirmthisPatienthas = 2610
    csAutomaticSelection = 2611
    csPrintForcedto = 2612
    csWBCP = 2613
    csWBCB = 2614
    csViewingAll = 2615
    csViewingPrimarySplit = 2616
    csViewingSecondarySplit = 2617
    csP1 = 2618
    csc = 2619
    csPC = 2620
    csEnternewValuefor = 2621
    csEnterResultfor = 2622
    csEnterSapCodefor = 2623
    csTotalsforImmunology = 2624

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csTotalsforHaematology = 2625
    csTotalsforEndocrinology = 2626
    csSourceTotals = 2627
    csRequests = 2628
    csEditHistologyCytology = 2629
    csEditSemenAnalysis = 2630
    csEditMicrobiology = 2631
    csWKPos = 2632
    csSTPos = 2633

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csEquiv = 2634
    'csTotalTests = 2635
    csPrintListing = 2636
    csCheck = 2637
    csDefault = 2638
    csUnvalidatetochange = 2639
    csBlockValidate = 2640
    csEndocrinologyEndofDayReportfor = 2641
    csImmunologyEndofDayReportfor = 2642
    csClickonNumberorDiscipline = 2643
    csOutstandingRequests = 2644
    csEPCAdenoRota = 2645
    csOvaParasite = 2646
    csUrineLogIn = 2647
    csBenceJonesMisc = 2648
    csNormalRange = 2649
    csPlausibleRanges = 2650
    csControlChart = 2651
    csIdentification = 2652
    csGramStains = 2653
    csWetPrep = 2654
    csCrystals = 2655
    csCasts = 2656
    csMiscellaneous = 2657

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csPrinters = 2658
    csCreatinineClearance = 2659
    csWorkList = 2660
    csGeneralWorkList = 2661
    csEndofDaySummary = 2662
    csMicroReports = 2663
    csIsolateReport = 2664
    csNCRIReport = 2665
    csImmunologyTotals = 2666

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csLimits = 2668
    csRunningMeans = 2669
    csGlucose = 2670
    csProblemwith = 2671
    csResultsOverview = 2672
    csCoagulationResult = 2673
    csBiochemistryHistory = 2674
    csCoagulationHistory = 2675
    csHaematologyHistory = 2676
    csHaematologyGraphs = 2677
    csCumulativeHaematology = 2678
    csImmunologyHistory = 2679
    csImmunologyGraphs = 2680
    csHaematologyResult = 2681
    csBloodGasResult = 2682
    csBloodGasHistory = 2683
    csCloseProgram = 2684
    csLogInOut = 2685
    csReportTime = 2686
    csUnknown = 2687
    csNotGiven = 2688
    csRequest = 2689
    csNotyetAvailable = 2690
    csReceived = 2691

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csImmunologyCumulative = 2692
    csEndocrinologyCumulative = 2693
    csHaematologyCumulative = 2694
    csBloodGasCumulative = 2695
    csBiochemistryCumulative = 2696
    csCoagulationCumulative = 2697
    csTimeTaken = 2698
    csNotSpecified = 2699
    csRecordXXXXofXXXX = 2700
    csMostRecentRecord = 2701
    csEarliestRecord = 2702
    csRecord = 2703
    csof = 2704
    csFaxCancelled = 2705
    csErrorNumber = 2706
    csReagentName1 = 2707

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCurrentLevel1 = 2708
    csAmountofNewReagent = 2709
    csReagentInformation = 2710
    csUsageperTest = 2711
    csReagentName = 2712
    csMinStock = 2713
    csCurrentLevel = 2714
    csClinicType = 2715
    csRace = 2716
    csMothersName = 2717
    csFathersName = 2718
    csFathersProfession = 2719
    csWhereBorn = 2720
    csHealthFacility = 2721
    csWhenCompleted = 2722
    csBreastFed = 2723
    csAnalyteMean1SD = 2724

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csYears = 2725
    csDays = 2726
    csMonths = 2728
    csH = 2729
    csL = 2730
    csX = 2731
    csTrace = 2732
    csCommunications = 2733

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCautionAdjustingthesesettingsmayhaveadetrimentaleffectontheInterfacePleaseconsultyouradministratorbeforemakinganychanges = 2734
    cs24HrResult = 2735
    csViewDeletedResults = 2736
    csHideDeletedResults = 2737
    csRemovefromResults = 2738
    csPrintHandler = 2739
    csPrintLayout = 2740
    csIfMoreThan = 2741
    csResultsthen = 2742
    csPrintsidebyside = 2743
    csPrintonSecondpage = 2744
    csDRAFTREPORT = 2745
    csPrintedby = 2746
    csValidatedby = 2747
    csClin = 2748
    csCopyClin = 2749
    csPrintedOn = 2750
    csLaboratoryPhone = 2751
    csAnalyserIdentifier = 2752
    csAnalyserID = 2753
    csV = 2754
    csP = 2755
    csat = 2756
    csCreatinineClearanceTest = 2757
    csVolumeCollected = 2758

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csStockAdded = 2759
    csPlasmaCreatinine = 2760
    csClearance = 2761
    csProteinConcentration = 2762
    csReportDate = 2763
    csGlucoseToleranceTest = 2764
    csGlucoseSeries = 2765
    csResultOptions = 2766
    csDisciplineAccessControl = 2767
    csbio = 2768
    csHae = 2769
    csImm = 2770
    csHow = 2771
    csSelectDiscipline = 2772
    csViewPreviousDetails = 2773
    csEndocrinologyResultsPhoned = 2774
    csMicrobiologyResultsPhoned = 2775
    csExternalResultsPhoned = 2776
    csBloodGasResultsPhoned = 2777
    csImmunologyResultsPhoned = 2778

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCoagulationResultsPhoned = 2779
    csBiochemistryResultsPhoned = 2780
    csHaematologyResultsPhoned = 2781
    csPhoneTo = 2782
    csDateTime = 2783
    csPhonedBy = 2784
    csHaemFinal = 2785
    csHaemDraft = 2786
    csBioFinal = 2787
    csImmFinal = 2788
    csEndFinal = 2789
    csCoagFinal = 2790
    csFaxFinal = 2791

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csPage1 = 2792
    csPage2 = 2793
    csNoofReports = 2794
    csPaymentCode = 2795
    csSendTo = 2796
    csEntry = 2797
    csSOP = 2798
    csWindow = 2799
    csAvailablePrinters = 2800
    csMappedTo = 2801
    csReligion = 2802
    csPatientBloodGroup = 2803

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSpecimenAcceptable = 2804
    csTBLeprosyDetails = 2805
    csPatientTBRegisterNumber = 2806
    csTreatmentUnit = 2807
    csAreaLeaderNeighbour = 2808
    csLeprosyRegisterNumber = 2809
    csRegionNumber = 2810
    csTBDistrictNumber = 2811
    csNameOfPersonRequesting = 2812
    csReceiptNumber = 2813
    csAmountPaid = 2814
    csNatureOfSpecimen = 2815
    csSpecimenType = 2816
    csBloodPackHistory = 2817
    csBloodPackHistoryReport = 2818
    csGroup = 2819
    csCollectedDateTime = 2820
    csExpiryDate = 2821
    csProductType = 2822
    csSourceType = 2823
    csEnterPatientChartNumber = 2824

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csSelectBloodGroup = 2825
    csCurrentInStock = 2826
    csTotalReceived = 2827
    csTotalUsed = 2828
    csCells = 2829
    csCellsPreparedBy = 2830
    csLotNumber = 2831
    csACells = 2832
    csBCells = 2833
    csAntiA = 2834
    csAntiB = 2835
    csAntiD = 2836
    csReceivedFrom = 2837

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csBroughtInBy = 2838
    csMeansOfTransport = 2839
    csDateTimeIssued = 2840
    csDateCollected = 2841
    csTransitTime = 2842
    csRhesus = 2843
    csVolumeReceived = 2844
    csAnaerobic5CO2 = 2845
    csPacksReceived = 2846
    csDonorsID = 2847
    csReasonForDonation = 2848
    csOrderedBy = 2849
    csDonorsDetails = 2850
    csPreDonationScreening = 2851
    csPreTestCounselled = 2852
    csHIVTest = 2853
    csHb = 2854
    csBP = 2855
    csBodyWeight = 2856
    csOkForDonation = 2857

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csPhlebotomist = 2858
    csSyphilis = 2859
    csHepatitisA = 2860
    csHepatitisB = 2861
    csHepatitisC = 2862
    csTestedBy = 2863
    csVerification = 2864
    csProductOKForUse = 2865
    csVerifiedBy = 2866
    csCurrentVolume = 2867
    csNewVolume = 2868
    csCreatedBy = 2869
    csReceivedVolume = 2870
    csVolumeAvailable = 2871
    csOriginalPack = 2872
    csNewPack = 2873
    csRequestDetails = 2874
    csRequestedAmount = 2875
    csEnterServerNamesWherePrintHandlerswillberunningintheLaboratoryNetworkDomain = 2876
    csPackDetails = 2877
    csXMatch = 2878

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csXMatchBy = 2879
    csXMatchDateTime = 2880
    csCollectedBy = 2881
    csCultureResults = 2882
    csTransfusionReaction = 2883
    csLesion1 = 2884
    csReaction = 2885
    csReturnedtoLab = 2886
    csReturnedBy = 2887
    csReceivedBy = 2888
    csReasonForTransfusion = 2889
    csTransfused = 2890

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csReasonForNonTranfusion = 2891
    csAddNewReasonForTransfusion = 2892
    csTypeOfSpecimen = 2893
    csDay1 = 2894
    csDay2 = 2895
    csTransportMedia = 2896
    csMacroscopic = 2897
    csAppearance = 2898
    csAmount = 2899
    csMedia1 = 2900
    csMedia2 = 2901
    csMedia3 = 2902
    csMedia4 = 2903

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csIncubationTemp = 2904
    csAnaerobicCO2 = 2905
    csSignificantGrowth = 2906
    csSuspectedOrganism = 2907
    csOrganism1 = 2908
    csOrganism2 = 2909
    csOrganism3 = 2910
    csOrganism4 = 2911
    csNotes = 2912
    csGramStain = 2913
    csAntibiotic = 2914
    csMedicalScientistComment = 2915
    csConsultantComment = 2916
    csAddNewTypeOfSpecimen = 2917
    csAddNewMedicalScientistComment = 2918
    csAddNewTransportMedium = 2919
    csAddNewMacroscopicAppearance = 2920
    csAddNewMacroscopicAmount = 2921
    csAddNewMedium = 2922
    csAddNewIncubationTemperature = 2923

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csAddNewZNStain = 2924
    csAddNewNegativeStain = 2925
    csAddNewLeishmanStain = 2926
    csAddNewOrganism = 2927
    csCoagulase = 2928

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCatalase = 2929
    csOxidase = 2930
    csDnase = 2931
    csMethylRed = 2932
    csVP = 2933
    csIndole = 2934
    csLesion1Site = 2935
    csMotility = 2936
    csSugars = 2937
    csOthers = 2938
    csReinc = 2939
    csSerology = 2940
    csUrineSpecimen = 2941
    csCultureIdentifiedAs = 2942
    csHaematoxilinEosinStain = 2943
    csNucleus = 2944
    csCytoplasm = 2945
    csSlide1 = 2946
    csSlide2 = 2947
    csHistologyNumber = 2948
    csBlocks = 2949
    csFs = 2950
    csStatus = 2951
    csPiece = 2952
    'csMacroscopicAppearance = 2953 - spelling mistake "Aspeito Macroscopico" should be "Aspecto Macroscopico"
    csTechnician = 2954
    csDoctor = 2955
    csPathologist = 2956

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csAutopsyNumber = 2957
    csDateTimeOfDeath = 2958
    csDateTimeOfAutopsy = 2959
    csCytoloyNumber = 2960
    csSpecimen = 2961
    csConsistency = 2962
    csDifferentNetworkDomainSide = 2963
    csCount = 2964
    csMillionPermL = 2965
    csNormalValue25to5mL = 2966
    csNormalValue60MillionmL = 2967
    csMotile = 2968
    csMotileProgressive = 2969
    csMotileNonProgressive = 2970
    csNonMotile = 2971
    csTotal_Percentage = 2972
    csVDRLRPR = 2973
    csTPHA = 2974

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csReferenceLaboratorySerialNumber = 2975
    csRFtest = 2976
    csHepAAntigen = 2977
    csHepBAntigen = 2978
    csHepCAntigen = 2979
    csRightEar = 2980
    csWidal = 2981
    csBrucellaAb = 2982
    csProtOX19 = 2983
    csDateOfAnalysis = 2984
    csBioline = 2985
    csDetermine = 2986
    csUnigold = 2987
    csHIVELISA = 2988
    csELISA1 = 2989
    csELISA2 = 2990
    csCutOffValue = 2991
    csOD = 2992
    csHIVOther = 2993
    csViralLoad = 2994
    csCopiesmL = 2995
    csHIV1DNAPCR = 2996
    csPositiveControl = 2997
    csNegativeControl = 2998
    csBloodSmear = 2999

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csStoolForOB = 3000
    cspH = 3001
    csSG = 3002
    csNitrite = 3003
    csRCBHb = 3004
    csWCC = 3005
    csCultureReading = 3006
    csColonies = 3007
    csProtein = 3008
    csMicroscopicExaminationResults = 3009
    csKetones = 3010
    csUrobilinogen = 3011
    csBilirubin = 3012
    csAddNewBloodSmear = 3013
    csAddNewUrine = 3014
    csTwoHundreadWhiteCells = 3015
    csNegativeControlStained = 3016
    csNegativeControlUnStained = 3017
    csPositiveControlStained = 3018
    csPositiveControlUnStained = 3019
    csSpecimenID = 3020
    csSpecimenDate = 3021
    csLabSerialRegisterNumber = 3022

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csFirstSpecimen = 3023
    csSecondSpecimen = 3024
    csThirdSpecimen = 3025
    csResultsOfDrugSensitivityTesting = 3049
    csLeftEar = 3027
    csStain = 3028
    csTumourMarkers = 3029
    csLesion2Site = 3030
    csDiagnosis = 3031
    csOriginOfRequest = 3032
    csDistrict = 3033
    csCliniciansName = 3034
    csRegion = 3035
    csPosition = 3036
    csTypeOfPatient = 3037
    csLesion2 = 3038
    csLabIdentificationLetter = 3039
    csLaboratorySmearResult = 3040
    csWeekOfIncubation = 3041

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csReferenceLaboratoryResults = 3042
    csSourceDetails = 3043
    csReceptionDetails = 3044
    csReceivedDateTime = 3045
    csLastDonationDate = 3046
    csKg = 3047
    csBloodGroup = 3048
    csScreening = 3049
    csHospitalRegistrationNumber = 3050
    csIssuedBy = 3051
    csIssuedDateTime = 3052
    csEnterBloodPackNumber = 3053
    csBloodGroup2 = 3054
    csPatientHistoryReport = 3055
    csPreparedDate = 3056
    csFollowingsplitedpacksfoundforselectedpacknumber = 3057
    csPleaseselectdetailsfromlistandpressmodifybuttonorpressnewbuttontocreatednewdetails = 3058
    csSpecimenA = 3059
    csSpecimenB = 3060
    csSpecimenC = 3061
    csSpecimenD = 3062
    csSpecimenE = 3063
    csSpecimenF = 3064
    csMACRO = 3065
    csAnalyte = 3066
    csResult = 3067

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csHL = 3068
    csPreviouslyDeletedAnalytes = 3069
    csAbnormalResultsshown = 3070
    csTime = 3071
    csSampleNumbers = 3072
    csAbc = 3073
    csSugar = 3074
    csNegativeStain = 3075
    csZNStain = 3076
    csHormoneTMComments = 3077
    csHIVRapidTest = 3078
    csCD4Comments = 3079
    csHistoricalViewofOrganism = 3080
    csLeishmansStain = 3081
    csDay3 = 3082
    csR = 3083
    css = 3084
    cs1st = 3085
    cs2nd = 3086
    cs3rd = 3087
    csDetails = 3088

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csItem = 3089
    csSampleIDtoPrint = 3090
    csorCulture = 3091
    csMicroscopy = 3092
    csPieces = 3093
    csAddNewGeneric = 3094
    csCurrent = 3095
    csReport = 3096
    csTBQCisRequired = 3097
    csShowMostRecent = 3098
    csStained = 3099
    csUnStained = 3100
    csCalculate = 3101
    csReagent = 3102
    csTestsInRedWillBeTransfered = 3103
    csAddNewParameter = 3104
    csNumberOfFrozenSectionsBetweenTheAboveDates = 3105
    csNumberOfBlocksBetweenTheAboveDates = 3106
    csChildren = 3107
    csOutPatients = 3108
    csPenicillinAllergy = 3109
    csMembersofthisGroup = 3110
    csExistingTests = 3111

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csIncludeInEndOfDayReport = 3112
    csAbsoluteValue = 3113
    csEnterNewControlName = 3114
    csPassword = 3115
    csDonation = 3116
    csReportMethod = 3117
    csTest = 3118
    csAddressSendTo = 3119
    csRCC = 3120
    csOperator = 3121

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csWeek1 = 3122
    csWeek2 = 3123
    csWeek3 = 3124
    csWeek4 = 3125
    csWeek5 = 3126
    csWeek6 = 3127
    csWeek7 = 3128
    csWeek8 = 3129
    csAllowDelete = 3130
    csGPName = 3131
    csIsolate = 3132
    csOrganismName = 3133
    csQualifier = 3134
    csArchivedBy = 3135
    csb = 3136
    css2 = 3137
    csd = 3138
    csT = 3139
    csDateAdded = 3140
    csDept = 3141
    csSentTo = 3142
    csSent = 3143
    csDepartment = 3144
    csPrintTime = 3145
    csReportTo = 3146
    csInitiator = 3147
    csNa = 3148
    csk = 3149
    csCl = 3150
    csGlu = 3151
    csBili = 3152
    csALP = 3153
    csGGT = 3154

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csALT = 3155
    csAMY = 3156
    csCPK = 3157
    csAST = 3158
    csLDH = 3159
    csCa = 3160
    csphos = 3161
    csMg = 3162
    csUrate = 3163
    csChol = 3164
    csTrig = 3165
    csCrea = 3166
    csRun_Number = 3167
    csSource = 3168
    csInvalidSampleID = 3169
    csReportedBy = 3170
    csReportDateTime = 3171
    csReportDetails = 3172
    csRefRange = 3173
    csExtLab = 3174

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csIdentifiedAs = 3175
    csArchiveTime = 3176
    csReportNo = 3177
    csRooH = 3178
    csFaxed = 3179
    csOnWarfarin = 3180
    csDateTimeDemographics = 3181
    csDateTimeHamePrinted = 3182
    csDateTimeBioPrinted = 3183
    csDateTimeCoagPrinted = 3184
    csPregnant = 3185
    csRecDate = 3186
    csHistoValid = 3187

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csCytoValid = 3188
    csHYear = 3189
    csSentToMedRenal = 3190
    csNameOfMother = 3191
    csNameOfFather = 3192
    csBreastFedEnd = 3193
    csProfession = 3194
    csSpecimenState = 3195
    csRequestDate = 3196
    csExported = 3197
    csPrinted = 3198
    csDateTimeOfRecord = 3200

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csNormalLow = 3201
    csNormalHigh = 3202
    csNormalUsed = 3203
    csRetDate = 3204
    csSentDate = 3205
    csSAPCode = 3206
    csOrderList = 3207
    csSaveTime = 3208
    csHa = 3209
    csBi = 3210
    csco = 3211
    csEn = 3212
    csBG = 3213
    csIm = 3214
    csEx = 3215
    csHi = 3216
    cscy = 3217
    csMi = 3218
    csPa = 3219
    csSe = 3220

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csAntibioticName = 3221
    '3222
    '3223
    csPenAll = 3224
    csAgeFrom = 3225
    csAgeTo = 3226
    csAgeFromDays = 3227
    csAgeToDays = 3228
    csPanelName = 3229
    csOrganism = 3230

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSample_Percentage = 3231
    csTest_Percentage = 3232
    csTS = 3233
    csBloodStockReport = 3234
    csHormonesTumourMarkers_Tab = 3235
    csListOfAllAvailableEquipment = 3236
    csFindBy = 3237
    csSupplier = 3238
    csSerialNumber = 3239
    csEnterDescription = 3240
    csItemName = 3241
    csListOfAllAvailableItems = 3242
    csListOfAllAvailableSuppliers = 3243
    csEnterCompanyName = 3244
    csBasicFilter = 3245
    csNoneIncludeAllItems = 3246
    csAdvandedFilter = 3247
    csOnlyItemsWith = 3248
    csServiceDueBefore = 3249
    csSelectedItem = 3250
    'csLocateStoreSpecimen = 3251
    csStoreLocateSpecimen = 3251
    csStockLocation = 3252
    csMultipleSearchResultsFound = 3253

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csStockDate = 3254
    csAddToStock = 3255
    csDeductFromStock = 3256
    csReagentsSuppliers = 3257
    csEquipments = 3258
    csVendor = 3259
    csQuantity = 3260
    csPacks = 3261
    csUnits2 = 3262
    csItemCurrentLocation = 3263
    csUnitsStoredInCurrentLocation = 3264
    csItemNewLocation = 3265
    csUnitsToMoveToNewLocation = 3266
    csDateToDate = 3267
    csSelectedItem2 = 3268
    csSelectedCategory = 3269
    csSelectedSupplier = 3270
    csIncludeInactiveItems = 3271
    csLaboratoryStockReport = 3272
    csSupplierDetails = 3273
    csSupplierName = 3274
    csCompany = 3275
    csContactInformation = 3276
    csMobile = 3277
    csEmail = 3278
    csCreatingNewSupplier = 3279
    csClickSavetocreatenewsupplier = 3280
    csItemDetails = 3281
    csManufacturer = 3282

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csDefaultSupplier = 3283
    csPacking = 3284
    csItemsPerPack = 3285
    csReorderLevel = 3286
    csLeadTime = 3287
    csGrade = 3288
    csStrength = 3289
    csCreatingNewItem = 3290
    csClickSavetocreatenewitem = 3291
    csLocationPath = 3292
    csEnternameoflocationtocreate = 3293
    csEnternameoflocationtofind = 3294
    csModel = 3295
    csSerial = 3296
    csInventory2 = 3297
    csLimitations = 3298
    csWarrantyDate = 3299
    csLastServiceDate = 3300
    csNextServiceDue = 3301
    csRepairContactDetails = 3302
    csCreatingNewEquipment = 3303
    csClickSavetocreatenewEquipment = 3304
    csModifyingexistingEquipment = 3305
    csClickSavetomodifythisEquipment = 3306
    csPleaseselectanitemfirst = 3307
    csSelectanitemtoviewfulllocationpath = 3308

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csEquipmentName = 3310
    csNoneIncludeAllEquipment = 3311
    csSelectedEquipment = 3312
    csOnlyEquipmentWith = 3313
    csEquipmentDetails = 3314
    csStockIn = 3315
    csStockOut = 3316
    csItemStockReport = 3317
    csEquipmentStockReport = 3318
    cssr = 3319
    csRprt = 3320
    csSensitive = 3321
    csResistant = 3322
    csUser = 3323
    csHaematologyAgeSexRelated = 3324
    csIfmorethanxxxxxxBiochemistryResultsthen = 3325
    csServerName = 3326
    csCD4Results = 3327
    csTMResults = 3328
    csFlag = 3329
    csIfThisResult = 3330
    csThenDo = 3331
    csIf = 3332
    csOr = 3333

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    cs24HrUrineExcretion = 4000

End Enum


Public Enum csFormsTabs
    csAbnormals = 4001
    csAddBloodGasTest = 4002
    csAddCoagulationTest = 4004
    csAddAnalyte = 4005
    csAddExternalAddress = 4006
    csAddtoExternalTestsRequested = 4007
    csAges = 4008
    csMicroInformation = 4009
    'csAuditTrail = 4010
    csAddHaematologyResult1 = 4011
    csBadResults = 4012
    csBarCodes = 4013
    csUrineBatchEntry = 4014
    csCategories = 4015
    csBatchEntryEPECRotaVirusAdenoVirus = 4016
    csBatchHIV = 4017
    csUrineBatchSampleLogIn = 4018
    'csOccultBlood = 4019
    csBatchOvaParasites = 4020
    csBatchEntryLactoseUreaPurity = 4021
    csBloodGasParameterDefinitions = 4022

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csBatchCultureEntry = 4023
    csBloodGasDailySummary = 4024
    csBioRelatedImmunologyResults = 4025
    '4026
    csBiochemistryPlausibleRanges = 4027
    csTestDefinitions = 4028
    '4029
    csEndocrinologySplits = 4030
    csImmunologySplits = 4031
    csTotalTests = 4032
    csClinicianList = 4033
    csBloodGasSummary = 4034
    csCreatinineClearance = 4035
    csDisciplineAccessControl = 4036
    csResultOptions = 4037
    csReagentUpdate = 4038
    csSetReagentLevels = 4039
    CS2 = 4040
    csReagentLevels = 4041
    csReagentHistory = 4042
    csReagentList = 4043
    csCoagulationDefinitions = 4044
    csCoagulationControlLimits = 4045
    csCoagulationRepeats = 4046
    csCoagulationTotals = 4047

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCoagulationTotalTests = 4048
    csComments = 4049
    csCoagulationDailySummary = 4050
    csCytologyTotals = 4051
    csDailyReport = 4052
    csBiochemistryEndofDayReportCommonParameters = 4053
    csDemographicConflict = 4054
    csDifferentials = 4055
    csGeneralChemistry = 4056
    csEndocrinologyAbnormals = 4057
    csAddEditEndocrinologyTest = 4058
    csEndocrinologyEndofDayReportCommonParameters = 4059
    csEndocrinologyEndofDayReportCommonParameter = 4060
    csImmunologyEndofDayReportCommonParameters = 4061

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csExternalBatches = 4062
    csEndocrinologyPlausibleRanges = 4063
    csImmunologyPlausibleRanges = 4064
    csEndocrinologyTotalTests = 4065
    csDefineEndocrinologyPanels = 4066
    csDefineImmunologyPanels = 4067
    csEndocrinologyParameterDefinitions = 4068
    csExternalReports = 4069
    csBiochemistryFastingRanges = 4070
    csPrinterOptions = 4071
    csFormTabbingOptionsfor = 4072

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csFullBloodGasHistory = 4073
    csFullPatientHistory = 4074
    csFullCoagulationHistory = 4075
    csFullEndocrinologyHistory = 4076
    csFullExternalHistory = 4078
    csFullHaematologyHistory = 4079
    csFullImmunologyHistory = 4080
    csFullImmunologyParaproteinHistory = 4081
    csMinusGlucoseToleranceTest = 4082
    csGpEntry = 4083
    csGPStats = 4084
    '4085
    csHaematologyBloodFilmDailySummary = 4086
    csHaematologyDefinitions = 4087
    csHaematologyGraphs = 4088
    csAddHaematologyResult = 4089
    csHospitals = 4090
    '    CS = 4091
    '    CS = 4092
    '    CS = 4093
    csImmunologyParameterDefinitions = 4094
    'CS = 4095


    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csControlLimits = 4096
    '    CS = 4097
    '    CS = 4098
    '    CS = 4099
    '    CS = 4100
    '    CS = 4101
    csDefinePanels = 4102
    ' CS = 4103
    'csPrinters = 4104
    'CS = 4105
    '4106
    'CS = 4107
    csEndofDayReportCommonParameters = 4108
    csExternalPanels = 4109
    '    CS = 4110
    '    CS = 4111
    '    CS = 4112

    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in CheckUpdateLanguage()

    '    CS = 4113
    '    CS = 4114
    '    CS = 4115
    '    CS = 4116
    '    CS = 4117
    csTotalStatistics = 4118
    ' CS = 4119
    csExternalStatistics = 4120
    '    CS = 4121
    '    CS = 4122
    '    CS = 4123
    '     CS = 4124
    csQualityControl = 4125
    csPrintResults = 4126
    csListofReasonsForTransfusion = 4127
    csBacteriology = 4128

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csHistologyQC = 4129
    csHistologyCytologyAndrology = 4130
    csRapidHIVQC = 4131
    csParasitology = 4132
    csParasites = 4133
    csDipStick = 4134
    csListofBloodSmear = 4135
    csListofUrine = 4136
    cstbMicroscopy = 4137
    cstbCulture = 4138
    csLeprosySkinSmear = 4139
    csLicencedtoTanzania = 4140
    csMinistryofHealthandSocialWelfare = 4141
    csPatientHistory = 4142
    csStockReport = 4143
    csHaematologyTransfusion = 4144
    csNetAcquireEquipmentStockLocator = 4145
    csNetAcquireSpecimenStorageandTracking = 4146
    csNetAcquireStockInStockOut = 4147
    csNetAcquireStockRelocation = 4148
    csNetAcquireStockReport = 4149
    csNetAcquireAllItems = 4150
    csNetAcquireAllEquipment = 4151
    csNetAcquireAllSuppliers = 4152
    csNetAcquireSuppliers = 4153
    csNetAcquireItemReportStockLocator = 4154
    csNetAcquireReagentsandSupplies = 4155
    csNetAcquireDefineStockLocations = 4156
    csNetAcquireEquipments = 4157
    csNetAcquireListof = 4158
    csNetAcquireAddTest = 4159
    csNetAcquireNormalRanges = 4160
    csNetAcquireFlagRanges = 4161

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csNetAcquireChangeTests = 4162
    csNetAcquirePlausibleandAutoValidationRanges = 4163
    csNetAcquireHaematologyTransfusion = 4164
    csNetAcquireChemistry = 4165
    csNetAcquireFastingRange = 4166
    csNetAcquireControlChart = 4167
    csNetAcquireHistology = 4168
    csNetAcquireParasitology = 4169
    csNetAcquireBacteriology = 4170
    csNetAcquireTB = 4171
    csNetAcquireSerologyImmunology = 4172
    csNetAcquireReception = 4173
    csNetAcquireOrganisms = 4174
    csNetAcquireMicroInformation = 4175
    csNetAcquireSearch = 4176
    csNetAcquireBarCodes = 4177
    csNetTracker = 4178
    csCreateModifyItem = 4179
    csCreateModifySupplier = 4180
    csCreateModifyEquipment = 4181
    csCreateModifyLocation = 4182
    csLaboratoryInventoryManagementSystem = 4183
    csTrackItemItemReport = 4184
    csTrackEquipmentEquipmentReport = 4185
    csReceiveItemsEquipmentIntoStock = 4186
    csRemoveItemEquipmentFromStock = 4187
    csMoveItemEquipmentStockFromOneLocationToAnother = 4188
    csCreateModifyStockLocation = 4189
    csSelectStockLocation = 4190

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPleaseSelectFromTheList = 4191
    csTrackSpecimen = 4192
    csServiceHistory = 4193
    csEquipmentProblems = 4194

End Enum

Public Enum csTooltips
    csUnabletodrawGraph = 6000
End Enum

Public Enum csAllMessages
    csNoSelectedUnitsOk = 6001
    csSystemInformationIsUnavailableAtThisTime = 6002
    csACodeisMandatory = 6003
    csNoSexgivenNormalrangesmaynotberelevant = 6004
    csDoyouwishtoCancelwithoutsaving = 6005
    csThiscodealreadyusedTryAnother = 6006
    csCodemustbeentered = 6007
    csfromtestsrequested = 6008

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csNorowselectfordeletion = 6009
    csResultValidCannotbechanged = 6010
    csResultmustbenumericTryAgain = 6011
    csEnterStartNumber = 6012
    csEnterStopNumber = 6013
    csInvalidStartNumber = 6014
    csInvalidStopNumber = 6015
    csStopNumbermustbegreaterthanStartNumber = 6016
    csNumbers = 6017
    csTooManyNumbersTryLess = 6018
    csResultsoutsidethisrangewillbemarkedasimplausible = 6019
    csResultvaluesoutsidethisrangewillbeflaggedasHighorLow = 6020
    csThesearethenormalrangevaluesprintedonthereportforms = 6021
    csPlausibleLow = 6022
    csPlausibleHigh = 6023
    csCodeLengthmustbe3Chars = 6024
    csEnterCliniciansSurname = 6025
    csSavingClinician = 6026
    csDoyouwishtoEditthisline = 6027
    csNoDobAdultAge25usedforNormalRanges = 6028
    csMaximumSevenDaysTryAgain = 6029
    csNoDisciplineChosenPleasechooseone = 6030
    csNoTestsChosen = 6031
    csCodealreadyusedTryAnother = 6032
    csCodemustbeNumericTryAnother = 6033
    csFirstlineofAddressmustbefilled = 6034
    '6035
    csSelectUnits = 6036
    csTestrequired = 6037

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPanelrequired = 6038
    csNoData = 6039
    csSelectSampleType = 6040
    csAShortNameisMandatory = 6041
    csALongNameisMandatory = 6042
    csPleaseEnteraTestName = 6043
    csNoSexDoBgivenAdultAge25usedforNormalRanges = 6044
    csIncorrectResultResulteitherPNorI = 6045
    csScanEntriesusingBarCodeReader = 6046
    csFrommustbelessthanto = 6047
    csNumbermustbeNumeric = 6048
    csPasswordRequiredtoClose = 6049
    csChartalreadyExistsTryAnother = 6050
    csChartmusthaveanamePleaseEnterone = 6051
    csYoumusthaveaChartPleaseEnterone = 6052
    csAgesarenotEditable = 6053
    csValidatealldemographicsforSelectedrows = 6054
    csDoNotgiveOutResultsOnlyHistoricalResultsforcomparison = 6055
    csTestsstillnotrun = 6056
    csClickHeretoShowFlags = 6057
    csNoPreviousDetails = 6058
    csTestalreadyExistsPleasedeletebeforeadding = 6059
    csisincorrect = 6060
    csDoyouwishalloutstandingrequestsDeleted = 6061
    csEnterBadComment = 6062

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csHaematologynotValidated = 6063
    csSexnotenteredDoyouwanttoentersexnow = 6064
    csYouMustHaveaSampleID = 6065
    csYouMustEnteraWard = 6066
    csYouMustEnteraWardorGp = 6067
    csFaxNumber = 6069
    csNoFaxNumberEnteredFaxCancelled = 6070
    csIncorrectFaxNumberEnteredFaxCancelled = 6071
    csDoyouwishtovalidatedemographics = 6072
    csSaveDemographics = 6073
    csUnvalidatePleaseEnterYourPassword = 6074
    csEnterTimeofDeath = 6075
    csDoyouwishalltoclearallresults = 6076
    csNoChartNoforPreviousDetails = 6077
    csSampleDateAfterRunDatePleaseAmend = 6078
    csRecDateAfterRunDatePleaseAmend = 6079
    csSampleDateAfterRecDatePleaseAmend = 6080
    csRundatenottodayProceed = 6081
    csEnterWetPrepResult = 6082
    csNoCryptosporidiumOocystsSeen = 6083
    csNoOvaorParasitesSeen = 6084
    csCryptosporidiumOocystsSeen = 6085
    csDone = 6086
    csPleasepickaTest = 6087
    csSelectAgeRange = 6088

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPleaseenternewcontrolname = 6089
    csDeleteAllRepeats = 6090
    csCopytoResult = 6091
    csYourPcseemstobeexperiencingtroubleretrievingdatafromtheserverorthenetworkmaybefaultyToremedythisNetAcquireneedstorebootyourPcThankYou = 6092
    csNetAcquireHasEncounteredaCriticalProblem = 6093
    csYoumustenteraDifferentialfirst = 6094
    csDifflessthan100Continue = 6095
    csRetrieveBiochemistryTestsRelevanttoImmunology = 6096
    csBackColourNormalAutomaticPrinterSelectionBackColourRedForced = 6097
    csOnly320Characters = 6098
    csAllowsBadSamplestobeCounted = 6099
    csOnly360Characters = 6100
    csPrinterSelectedAutomatically = 6101
    csExternalSaveEnabled = 6102
    csSampleIDlongerthenrecommended = 6103
    csRefRangeNotAgeSexRelated = 6104
    csRefRangeNotAgeRelated = 6105
    csRefRangeNotSexRelated = 6106
    csNotPrinted = 6107
    csInvalidTime = 6108

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csDemographicsValidated = 6109
    csDemographicDetailshavechangedSave = 6110
    csHaematologyDetailshavechangedSave = 6111
    csBiochemistryDetailshavechangedSave = 6112
    csCoagulationDetailshavechangedSave = 6113
    csEndocrinologyDetailshavechangedSave = 6114
    csBloodGasDetailshavechangedSave = 6115
    csImmunologyDetailshavechangedSave = 6116
    csExternalDetailshavechangedSave = 6117
    csPassword_QuestionMark = 6118
    csDoYouwanttoprintthis = 6119
    csMaintenancenotavailableatthistime = 6120
    csSorryoutofmemorypleaseclosesomeprogramsandtryagain = 6121
    csUnabletofindWebBrowser = 6122
    csOptionsnotavailableatthistime = 6123
    csEnternewLastUsedNumber = 6124
    csLastUsedNumberchangedto = 6125
    csLastUsedNumbernotchanged = 6126
    csPrintersNotLoaded = 6127
    csThefollowingnumberswerepositive = 6128
    csNothingtodoSelectaNametoPrint = 6129
    csNodetailsfound = 6130
    csPrintasGTTReport = 6131
    csPrintasGlucoseSeries = 6132
    csNoneFound = 6133
    csComplied = 6135
    csUnvalidateallSelectedRows = 6136
    csUnvalidSampleID = 6137
    csInvalidPassword = 6138

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csMaximumof50Recordsallowed = 6139
    csCautionYouareabouttoremoveALLreferencestothisSampleNumberThischangeisnotreversibleAlldataforthisSamplenumberwillbelost = 6140
    cs3LoginstriedProgramwillnowclose = 6141
    csContactSystemAdministrator = 6142
    csMusthaveNameCodeandPassword = 6143
    '6144
    csPasswordConfirmdontmatch = 6145
    csRetypePasswordandConfirmation = 6146
    csNamealreadyused = 6147
    csPasswordalreadyused = 6148
    csTypeanotherPassword = 6149
    csLogOffDelayMinutes = 6150
    csDragtochangeListOrder = 6151
    csEnterNewPassWord = 6152
    csVerifyNewPassWord = 6153
    csPasswordsdontmatch = 6154
    csPasswordmustbeatleastsixCharacters = 6155
    csyourPasswordisnowChangedThankYou = 6156
    csSampleIDMustBeNumeric = 6157
    csListofWards = 6158
    csEnterWard = 6159
    csEnterFaxNumber = 6160
    csClicktoToggle1forInuse0fornotInuse = 6161
    csNoCriteriaEntered = 6162
    csInvalidSearchCriteria = 6163
    csNewChartNumber = 6164
    csDatecannotbegreaterthantoday = 6165
    csDoyouwanttochangetheChartNumber = 6166
    csFillinPhoneTo = 6167
    csTheReportisbeingGeneratedPleaseWait = 6169
    csEnterAddress = 6170
    csRemoveFromList = 6171
    csReportsToomanytoprintMaximum50 = 6172
    csReportsYourequestedtoprintAreyousure = 6173
    csClickonSpecificPrinterNametoEdit = 6174
    csCurrentDefaultPrinter = 6175
    csMakeaSelectionfromtheavailableprinters = 6176

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPROCEEDWITHCAUTIONNewPrinterName = 6177
    csTestsinREDtobeTransferred = 6178
    csHighlightTestsToBeTransferred = 6179
    csPleaseChooseaReagent = 6180
    csReasonforLevelchange = 6181
    csConfirmed = 6182
    csYouMustTypeinaComment = 6183
    csPleaseEnteraHospitalName = 6184
    csareyousure = 6185
    csNoCodeshavebeenenteredContactCustomSoftware = 6186
    csNoPrinterInstalledonPc = 6187
    csPleasechoseaReagentaTestandenteranAmount = 6188
    csLowRangeMustBeLessthanHighRange = 6189
    csNoAmountEntered = 6190
    csYouarenotallowedtoEditthisGP = 6191
    csReagentcannotbedeletedasithasStock = 6192
    csYoucannotchangeaHospitalalreadyinuse = 6193
    csEmptystringreceived = 6194
    csTimeoutError = 6195
    csDetailsarenotsavedPleasefillinfieldsmarkedinredandtryagain = 6196
    csDetailsSuccessfullysaved = 6197
    csPleaseenterPackNumberFirst = 6198
    csPleaseenterChartNumberfirst = 6199
    csPleaseenterBloodGroupFirst = 6200
    csPleaseenteranotherpacknumber = 6201
    csNoPreviousdonationdetailsfound = 6202
    csFollowingDonationsfoundforselecteddonor = 6203
    csScreeningTab = 6204
    csEnterPasswordtoUnlock = 6205
    csAsthisbloodpackhasbeenissued = 6206
    csSelectDateofDeathand = 6207
    csYoucannotmodifydetails = 6208
    csLockedDetailsEnterPasswordtoUnlock = 6209

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csYouhaveunsaveddetailsonfollowingtab = 6210
    csPacksReceivedTab = 6211
    csEnterTimeIssued = 6212
    csEnterTimeCollected = 6213
    csSplitTab = 6214
    csDoyouwanttoexitwithoutsavingdetails = 6215
    csDonorsTab = 6216
    csEnterTimeofDonation = 6217
    csSelectDonordateofbirth = 6218
    csSelectpackexpirydate = 6219
    csSelectDateScreenedand = 6220
    csEnterTimeScreened = 6221
    csThispackisalreadybeingusedonsomeothertab = 6222
    csPacknumberhasalreadybeenusedunderpackreceivedtabcannotbeusedagain = 6223
    csThispacknumberdoesntexist = 6224
    csSelectDateofAutopsyand = 6225
    csPackNumberNotNumber = 6226
    csYoucannotfurthersplitthispack = 6227
    csThisisalreadyasplitpack = 6228
    csPackisnotscreened = 6229
    csPleasescreenthispackfirst = 6230
    csSelectDateCollectedand = 6231
    csSelectIssuedDateand = 6232
    csEnterIssuedTime = 6233
    csSelectDateReceivedand = 6234
    csEnterTimeReceived = 6235
    csPleaseselectpatientfirst = 6236
    csPleaseSelectrequestedproductfirst = 6237
    csPackNumberdoesnotexist = 6238
    csBloodpackcannotbeissued = 6239
    csItdoesntmatchwithrequestedproducttype = 6240
    csBloodgroupdoesntmatchwithpatientGroup = 6241
    csPackisexpired = 6242

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csThisbloodpackhasalreadybeenissued = 6243
    csScreeningresultsarenotokfortransfusion = 6244
    csEnterTimeofAutopsy = 6245
    csSorryNodetailsfound = 6246
    csNoPrevioussplitdetailsfound = 6247
    csFollowingsplitsfoundforselectedpack = 6248
    csPleaseselectdetailsfromlistaboveandpressModifySelectedDetailsbutton = 6249
    csORPressCreateNewDetailsbutton = 6250
    csNoCurrentRecordFound = 6251
    csNoChangesMade = 6252
    csNoStockFoundFor = 6253
    csNothingToExport = 6254
    csPleasereentertheTime = 6255
    csPackNotNumber = 6256
    csYouCannotCreateSplitDetails = 6257
    csNothingToPrint = 6258
    csSelectPhonedToetc = 6259
    csPackNumberHasAlreadyBeenUsedUnderDonationTabCannotBeUsedAgain = 6260
    csPleaseselectBloodGroupfirst = 6261
    csSaveError = 6262
    csVarifyNewPassword = 6263
    csCancelwithoutSaving = 6364
    csEnterSampleID = 6365
    csSpecimenIDnotknown = 6366
    '6367
    csSampleNumber = 6368
    csExportCancelled = 6369
    csExportCulturetoExcel = 6370
    csExportMicroscopytoExcel = 6371
    csReceivedDatebeforeSampleDate = 6372
    csRunDatebeforeReceivedDate = 6373
    csIsSpecimenOK = 6374
    csSampleIDnotentered = 6375

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSurnamenotentered = 6376
    csSexnotentered = 6377
    csDateofBirthnotentered = 6378
    csCliniciannotentered = 6379
    csRundateaftertoday = 6380
    csRunDatebeforeSampleDate = 6381
    csNewequipmentdetailsaresaved = 6382
    csEditedequipmentdetailsaresaved = 6383
    csPleaseenterequipmentnamefirst = 6384
    csPleaseselectsupplierfirst = 6385
    csLockedSampleEnterPasswordtoUnlock = 6386
    csUnitsnotentered = 6387
    csWBCResultnotknown = 6388
    csNoItemFound = 6389
    csLastServiceDateCannotBeGreaterThanNextServiceDate = 6390
    csAreYouSureYouWantToRemoveThisStockEntry = 6391
    csIsAlreadyBeingModified = 6392
    csPleaseSaveOrResetFirst = 6393
    csPleaseEnterValidQuantityFirst = 6394
    csInsufficientStockCurrentAvailableStock = 6395
    csFromDateCannotBeGreaterThanToDate = 6396
    csNewStockDetailsAreSaved = 6397
    csModifiedStockDetailsAreSaved = 6398
    csNewItemDetailsAreSaved = 6399
    csModifiedItemDetailsAreSaved = 6400
    csPleaseSelectALocationFirst = 6401
    csSelectedLocationHasSubLocationsPleaseDeleteSubLocationsFirst = 6402
    csAreYouSureYouWantToDeleteThisLocation = 6403
    csSelectedLocationIsInUsePleaseRemoveStockFromThisLocationFirst = 6404
    csPleaseEnterLocationNameFirst = 6405
    csPleaseSelectParentLocationFirst = 6406
    csNoLocationFound = 6407

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csPleaseSelectASupplierFirst = 6408
    csPleaseEnterLotNumberFirst = 6409
    csItemExpiredToday = 6410
    csItemAlreadyExpired = 6411
    csIsExpiryDateCorrect = 6412
    csNoLocationSelectedDoYouStillWantToExit = 6413
    csSelectedLocationHasSubLocationsSoCannotBeSelected = 6414
    csSomeSpecimenHasAlreadyBeenModified = 6415
    csPleaseSelectAtleastOneOfTheFollowings = 6416
    csWrongDateOfBirthDateOfBirthCannotBeGreaterThanTodaysDate = 6417
    csDateOfBirthIsEnteredAsTodaysDateIsThisCorrect = 6418
    csSpecimenIDAlreadyExistsPleaseEnterDifferentSpecimenID = 6419
    csNewSpecimenDetailsAreSaved = 6420
    csModifiedSpecimenDetailsAreSaved = 6421
    csUnitsHasBeenMovedTo = 6422
    csSourceAndTargetLocationsCannotBeSamePleaseSelectDifferentLocation = 6423
    csPleaseEnterUnitsToMoveToNewLocation = 6424
    csUnitsToBeMovedCannotBeGreaterThanAvailableUnits = 6425
    csPleaseSelectItemCategoryFirst = 6426
    csNoFilterIsDefinedSearchWillBeSlowAndWillTakeFewSeconds = 6427
    csDoYouWantToProceed = 6428
    csNewSupplierDetailsAreSaved = 6429
    csModifiedSupplierDetailsAreSaved = 6430
    csPleaseEnterCompanyNameFirst = 6431
    csSelectedLocationIsSystemGeneratedPleaseChooseAnother = 6432
    csPleaseEnterItemNameFirst = 6433
    csPleaseSelectACategoryFirst = 6434

End Enum

Public Enum csControls
    csOK = 7000
    csCancel = 7001
    csAbort = 7002
    csRetry = 7003
    csIgnore = 7004

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csYes1 = 7005
    csNo1 = 7006
    '7007
    '7008
    csExit = 7009
    csUnits = 7010
    csSave = 7011
    csPort = 7012
    csDataBits = 7013
    csStopBits = 7014
    csContinuous = 7015
    csSettings = 7016
    csLayout = 7017
    csCopy = 7018
    csGetNumbers = 7019
    csAllValidNotValidPrintedandNotPrinted1 = 7020
    csValidPrintedorNotPrinted1 = 7021
    csSearchBy = 7022
    csValidnotPrinted1 = 7023
    csBetweenNumbers = 7024
    csResetPassword = 7025
    csAddNewOperator = 7026
    csConfirmPassword = 7027
    '7028
    csPrintReports = 7029
    csAddNewCategory = 7070
    csAddNewUnit = 7031
    csRemoveAgeRange = 7032
    csAddAgeRange = 7033
    csNewSite = 7034
    csNewOrganismGroup = 7035
    csNewOrganism = 7036
    csNewAntibiotic = 7037

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSelectOrganism = 7038
    csStart = 7039
    csShow = 7040
    csExporttoExcel = 7041
    csLastQuarter = 7042
    csLastFullQuarter = 7043
    csLastFullMonth = 7044
    csYeartoDate = 7045
    csLastMonth = 7046
    csLastWeek = 7047
    csTotal = 7048
    csRetics = 7049
    csMonospot = 7050
    csRandom = 7051
    csFasting = 7052
    csSetAnalyserB = 7053
    csSetAnalyserA = 7054
    csClear1 = 7055
    csSetAllNegative = 7056
    csOrganismGroup = 7057
    csSetAllPoly2Negative = 7058
    csSetAllPoly3Negative = 7059
    csSetAllPoly4Negative = 7060
    csSetAll018cKNegative = 7061
    csAdenoVirus = 7062
    csSet = 7063
    csMSU = 7064
    csCSU = 7065
    csBSU = 7066
    csSPA = 7067
    csSpecificGravity = 7068
    csBenceJones = 7069
    csCS = 7070

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSet12and3AllNegative = 7071
    csSetCryptoAllNegative = 7072
    csSetWetPrepAllNegative = 7073
    csSetAuramineAllNegative = 7074
    csUrineResultsBatchEntry = 7075
    csNil = 7076
    csAcid = 7077
    csAlkaline = 7078
    csNeutral = 7079
    csPacked = 7080
    csLactoseUreaPurity = 7081
    csAllNeg = 7082
    csSubSeleniteAllDone = 7083
    csPrimaryCultureAllDone = 7084
    csCampylobacterSetNegative = 7085
    csEColi0157SetAllNegative = 7086
    csEColi0157NotIsolated = 7087
    csNotIsolated = 7088
    csUpdate = 7089
    csinuse = 7090
    csHostCode = 7091
    csIncludeinRunningMean = 7092
    csDoDelta = 7093
    csPrintPriority = 7094
    csCSFOther = 7095
    csPlausible = 7096
    csScanBarCode = 7097
    csSortBy = 7098
    csNOPAS = 7099
    csValidateSelectedRows = 7100
    csValidateSampleID = 7101
    csCopytoImmunology = 7102
    csenter = 7103
    csRemoveItem = 7104
    csMovetoSplit2 = 7105
    csMovetoSplit1 = 7106
    csSecondary = 7107
    csPrimary = 7108
    csClicktoToggle = 7109
    csClicktoEdit = 7110
    csClicktoMove = 7111
    csDeltaCheck = 7112

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csNew1 = 7113
    csCopyData = 7114
    csPasteData = 7115
    csClicktoToggleYesNo = 7116
    csDelete = 7117
    csSemen = 7118
    csAddNew = 7119
    csClear = 7120
    csRefreshNumbers = 7121
    csValidate = 7122
    csStartSearch = 7123
    '7124
    csRestart = 7125
    csOther = 7126
    csPrintListing = 7127
    csSelect = 7128
    csSelectNone = 7129
    csClearAllResults = 7130
    csChangekeysorwording = 7131
    csEnterResults = 7132
    csSaveKeySettings = 7133
    csValid = 7134
    csPrintResult = 7135
    csSaveDetails = 7136
    csSaveChanges = 7137
    csDeleteTest = 7138
    csHaemolysed = 7139
    csSlightlyHaemolysed = 7140
    csLipaemic = 7141
    csOldSample = 7142
    csGrosslyHaemolysed = 7143
    csIcteric = 7144
    csRemoveDuplicates = 7145

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csRemoveResult = 7146
    csAddResult = 7147
    csAddResultManually = 7148
    csResultValidation = 7149
    csViewRepeat = 7150
    csViewRepeatedTests = 7151
    csRePrint = 7152
    csRePrintalreadyPrintedResults = 7153
    csChooseSampleType = 7154
    csChooseUnits = 7155
    csChooseTest = 7156
    csGetBioTests = 7157
    csSaveComment = 7158
    csSaveCommentChanges = 7159
    csSaveReview = 7160
    csViewGraph = 7161
    csMalariaScreen = 7162
    csAutoImmunePanelResults = 7163
    csRun = 7164
    csPrintINR = 7165
    csRunDateTime = 7166
    csAssociatedGlucose1 = 7167
    csAlreadyValidated = 7168
    csViewPatientHistory = 7169
    csExitScreen = 7170
    csPrintAll = 7171
    csClearDiff = 7172
    csPrintReview1 = 7173
    csPrintReview = 7174
    csDeletethisTest = 7175
    csConfirmDeletion = 7176
    csEdit = 7177
    csAddTest = 7178

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPrintSequence = 7179
    csPrintFormat = 7180
    csAddPanels = 7181
    csDefine = 7182
    csAddCategory = 7183
    csAddExternalAddress = 7184
    csStock = 7185
    csSetSourceNames = 7186
    csOn = 7187
    csOff = 7188
    csDown = 7189
    csUpdateTabOrder = 7190
    csFontSelection = 7191
    csSetFont = 7192
    csReferenceRangeColours = 7193
    csUpdateColour = 7194
    csBack1 = 7195
    csFore = 7196
    csSetTabIndex = 7197
    csImplausible = 7198
    csClickonParametertoshowGraph = 7199
    csHealthLink = 7200
    csListofGPs = 7201
    csPrintUnvalidated = 7202
    csUnvalidateSelectedRows = 7203
    csValidnotPrinted = 7204
    csNotValid = 7205
    csValidPrintedorNotPrinted = 7206
    csAllValidNotValidPrintedandNotPrinted = 7207
    csSavePrint = 7208
    csAddNewHospital = 7209
    csGetNames = 7211
    csPrintRefRange = 7212

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csRemoveAllReferences = 7213
    csNot = 7214
    csLookUp = 7215
    csUsers = 7216
    csAddEditPractices = 7217
    csAddOptions = 7218
    csChange1 = 7219
    csAddNewWard = 7220
    csAsYouType = 7221
    csDownload = 7222
    csOtherHospitals = 7223
    csCopytoEdit = 7224
    csHistoric = 7225
    csExactMatch = 7226
    csUseSoundex = 7227
    csLeadingCharacters = 7228
    csTrailingCharacters = 7229
    csWhen = 7230
    csCriteria = 7231
    csAddToTests = 7232
    csAddtoAddress = 7233
    csAddtoList = 7234
    csRemoveAllBloodGasRequests = 7235
    csRemoveAllHaematologyRequests = 7236
    csRemoveAllEndocrinologyRequests = 7237
    csRemoveAllImmunologyRequests = 7238
    csSetAllEndocrinologyStatustoPrinted = 7239
    csSetAllImmunologyStatustoPrinted = 7240
    csSetAllBiochemistryStatustoPrinted = 7241
    csSetAllHaematologyStatustoPrinted = 7242
    csSetAllBloodGasStatustoPrinted = 7243
    csSetAllCoagulationStatustoPrinted = 7244
    csRemoveAllBiochemistryRequests = 7245

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csRemoveAllCoagulationRequests = 7246
    csStopPrint = 7247
    csCheckPrinters = 7248
    csArrangeHorizontal = 7249
    csArrangeVertical = 7250
    csArrangeIcons = 7251
    csChange = 7252
    csToggleGraph = 7253
    '7254
    csManualLogOff = 7255
    csAutoLogOff = 7256
    csRefresh = 7257
    csInsertResult = 7258
    csViewMostRecentRecord = 7259
    csViewNextRecord = 7260
    csViewPreviousRecord = 7261
    csViewEarliestRecord = 7262
    '7263
    csBack = 7264
    csRemoveTests = 7265
    csRemoveAnalytefromResults = 7266
    csRemoveAllResults = 7267
    csNoDeletedResults = 7268
    csSelectAnalytestoberetrieved = 7269
    csUnDeleteselectedAnalytes = 7270
    '7271
    csUp = 7272
    csDefineSampleType = 7273
    csDefineHospital = 7274
    csDefineCategory = 7275
    csyes = 7276
    csno = 7277
    '7278

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csAddNewComment = 7279
    csSelenite = 7280
    csAddNewClinician = 7281
    csDefaultWard = 7282
    csSickleScreen = 7283
    csRheumatoidFactor = 7284
    csMonospotIM = 7285
    csSuspect = 7286
    csAbnormal = 7287
    cscc1 = 7288
    csNew = 7289
    csParameter = 7290
    csNewWindow = 7291
    csCascade = 7292
    csAddNewPrinter = 7293
    csCopyGP = 7294
    csSearchFor = 7295
    csAddGP = 7296
    cscc = 7297
    csSecretarys = 7298
    csPhoned = 7299
    csChecked = 7300
    csPhonedChecked = 7301
    csUnValid = 7302
    csFastingSample = 7303
    csCheckDemographics = 7304
    csBlockPrint = 7305
    csSampleTracking = 7306
    csNameDateofBirth = 7307
    csViewingSecondarySplit = 7308
    csMissing = 7309
    csMaintenance = 7310
    csSplits = 7311

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csFastingValues = 7312
    csDeltaCheckLimits = 7313
    csGPWorkLoad = 7314
    csBatch = 7315
    csByDate = 7316
    csByName = 7317
    csUsage = 7318
    '7319
    csAdministrators = 7320
    csManagers = 7321
    csRemovePanel = 7322
    csAddNewPanel = 7323
    csFollowUp = 7324
    csSputum = 7325
    csSkinSmear = 7326
    csConfirm = 7327
    csExtraDetails = 7328
    csTransfusion = 7329
    csReception = 7330
    csDonors = 7331
    csRecipient = 7332
    csEnquiry = 7333
    csEvent = 7334
    csPackNumber = 7335
    csNumberofReports = 7336
    csPerson = 7337
    '7338
    csIn = 7339
    csOut = 7340
    csBalance = 7341
    csLock = 7342
    csUnlock = 7343
    csImmediate = 7344

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csAftercompletion = 7345
    csPatientisOK = 7346
    csPatientAbsconded = 7347
    csCompatible = 7348
    csIncompatible = 7349
    csONegative = 7350
    csOPositive = 7351
    csANegative = 7352
    csAPositive = 7353
    csBNegative = 7354
    csBPositive = 7355
    csABNegative = 7356
    csABPositive = 7357
    csCulture = 7358
    '7359
    csSensitivity = 7360
    csHistologyReport = 7361
    csPurpleBlue = 7362
    csRed = 7363
    csFail = 7364
    csAutopsyReport = 7365
    csHistologyWorkScreen = 7366
    csSemenAnalysisAndrology = 7367
    csQCRequired = 7368
    csNothingAbnormalDetected = 7369
    csIsoniazid = 7370
    csRifampicin = 7371
    csStreptomycin = 7372
    csEthambutol = 7373
    csNotDone = 7374
    csReset = 7375
    csSelectReportofInterestthenclickViewReport = 7376
    csModifySelectedDetails = 7377

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csCreateNewDetails = 7378
    csModifySelectedPack = 7379
    csCreateNewPack = 7380
    csJaundiced = 7381
    csSetPrinter = 7382
    csRemoveAll = 7383
    csRetrieve = 7384
    csOtherSpecify = 7385
    csFaxedTo = 7386
    csCollectedinLabBy = 7387
    csViewReport = 7388
    csReportResults = 7389
    cs3Points = 7390
    csOneColumn = 7391
    csTwoColumns = 7392
    csChemistry = 7393
    csQCHistory = 7394
    csCopyToFile = 7395
    csSerologyImmunology = 7396
    csTB = 7397
    csDeleteAll = 7398
    csWards = 7399
    csLogin = 7400
    csAllow = 7401
    csExclude = 7402
    csAddNewGroup = 7403
    csSaveListOrder = 7404
    csSetKeysOrText = 7405
    csUndoChanges = 7406
    csFind = 7407
    csRefreshList = 7408
    csPickItem = 7409
    csPickLocation = 7410

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csModify = 7411
    csPickSupplier = 7412
    csOnlyItemsneedtobereordered = 7413
    csOnlyItemsexpiringon = 7414
    csRefrigerated = 7415
    csActive = 7416
    csSpecimenOK = 7417
    csGridResult = 7418
    csGridAnalyser = 7419
    csGridComment = 7420
    csGridRefRange = 7421
    csDateTimeArchived = 7422
    csBlk = 7423

    csStoreTrackSpecimen = 7424
    csReceiveStock = 7425
    csRemoveStock = 7426
    csMoveStock = 7427
    csTrackItemLocation = 7428
    csTrackEquipment = 7429
    'csStockReport = 7430
    csTop50ItemsToBeReordered = 7432
    csItemsExpiringInNext7Days = 7433
    cs10MostRecentlyStoredSpecimens = 7434

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csAutoRefreshEvery10Seconds = 7435
    csRefreshNow = 7436
    csHistoryOfSpecimenFrom = 7437
    csHistoryOfStockReceivedFrom = 7438
    csHistoryOfStockRemovedFrom = 7439
    csTotalQuantityInUnits = 7440
    csSelectAnEquipmentToViewFullPath = 7441
    csListItems = 7442
    csListSuppliers = 7443
    csListEquipment = 7444
    csListLocations = 7445
    csSelectItem = 7446
    csSelectSupplier = 7447
    csSelectEquipment = 7448
    csSelectLocation = 7449

    csEnabled = 7450
    csDisabled = 7451
    csPrintHandleroffline = 7452
    csPrintHandleronline = 7453
    csLaboratoryPrinting = 7454
    csPrintHandlerServer = 7455
    csPrintHandlerVersion = 7456
    csPrintOptions = 7457
    csPrinterSetUp = 7458
    csLaboratorySide = 7459
    csPrnRegion = 7460
    csPrnReportDatePersonresponsible = 7461
    csPrnCountry1 = 7462
    csPrnCountry2 = 7463

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPrnGENERALPURPOSEDIAGNOSISINVESTIGATIONFORM = 7464
    csPrnNameofHospital = 7465
    csPrnHospitalRegistrationNumber = 7466
    csPrnAddressofHospital = 7467
    csPrnSurname = 7468
    csPrnPostalResidentialAddress = 7469
    csPrnDistrict = 7470
    csPrnFirstNames = 7471
    csPrnPhone = 7472
    csPrnFax = 7473
    csPrnDateofBirth = 7474
    csPrnSex = 7475
    csPrnRequesttoClinicalLaboratory = 7476
    csPrnClinicWard = 7477
    csPrnReligion = 7478
    csPrnRequestedDateRequestedByFirm = 7479
    csPrnNameSignature = 7480
    csPrnSpecimenCollectionDate = 7481
    csPrnTime = 7482
    csPrnClinicalNotes = 7483
    csPrnDiagnosis = 7484
    csPrnInvestigation = 7485
    csPrnNatureofSpecimen = 7486
    csPrnBloodGroup = 7487
    csPrnSampleID = 7488
    csPrnSerologyResults = 7489
    csTBCultureResults = 7490
    csPrnParasitologyResults = 7491
    csPrnHistologyReport = 7492
    csPrnDipStickResults = 7493

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPrnNitrite = 7494
    csPrnProtein = 7495
    csPrnBilirubin = 7496
    csPrnHIVReport = 7497
    csPrnHIVRapidTestBioLine = 7498
    csPrnElisaResult1 = 7499
    csPrnElisaResult2 = 7500
    csPrnPCROD = 7501
    csPrnPCRResult = 7502
    csPrnNatureofSpecimenA = 7503
    csPrnNatureofSpecimenB = 7504
    csPrnNatureofSpecimenC = 7505
    csPrnNatureofSpecimenD = 7506
    csPrnMacroReport = 7507
    csPrnMicroReport = 7508
    csPrnCytologyReport = 7509
    csPrnSemenAnalysis = 7510
    csPrnSpermatozoaCount = 7511
    csprnMillionpermL = 7512
    csPrnAndrologyReport = 7513
    csPrnTBMicroscopyResults = 7514
    csPrnCultureReadingWeek = 7515
    csPrnIsoniazidRifampicin = 7516
    csPrnLabSerialnumber = 7517
    csPrnIssuedDateTime = 7518
    csPrnLaboratorySmearResult1st = 7519
    csPrn2nd = 7520
    csPrn3rd = 7521

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csPrnPersonRequesting = 7522
    csPrnLeprosyResults = 7523
    csPrnLeprosy = 7524
    csPrnAutopsy = 7525
    csPrnAndrology = 7526
    csPrnNatureofSpecimenE = 7527
    csPrnNatureofSpecimenF = 7528

    csDifferentialDone = 7529
    csClinicians = 7530
    csAddDifferentialTests = 7531
    csParasitologyReport = 7532
    csHIVResult = 7533
    csElisaResult = 7534
    csScan = 7535
    csViewScan = 7536
    csResolution = 7537
    csNetAcquireScan = 7538
    csClickOnListToViewReport = 7539
    csNetAcquireViewScannedReports = 7540
    csExternalSpecimenID = 7541
    csBW = 7542
    csGrey = 7543
    csColour = 7544
    csShowExtraDetails = 7545
    csShowTBLeprosy = 7546
    csFolderDoesNotExist = 7547
    csDoYouWantToCreateIt = 7548
    csPathToSave = 7549
    csErrorOpeningScanner = 7550
    csErrorSettingScannerResolution = 7551
    csDonationDate = 7552
    csSelectDateofDonationand = 7553
    csQualifiers = 7554
    csSensitivityQualifiers = 7555
    csWholeBlood = 7556
    csCutUp = 7557

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csSpecials = 7558
    csCompleted = 7559
    csWeeks = 7560
    csSlide = 7561
    csNucleusnotentered = 7562
    csCytoplasmnotentered = 7563
    csNucleusResultnotcorrect = 7564
    csCytoplasmResultnotcorrect = 7565
    csOnly500Characters = 7566
    csFindMostRecentRecord = 7567
    csDrawGraph = 7568
    csUserCanPrint = 7569
    csUserCanView = 7570
    csSerologyOrImmunology = 7571
    csAccessRights = 7572

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''
    csPolimeraseChainReaction = 7573
    csFibrinogen = 7574
    csDDimers = 7575
    csBloodParasites = 7576
    csRapidMalariaTest = 7577
    csPlasmodiumInvestigation = 7578
    csTripanosomeInvestigation = 7579
    csFilariaInvestigation = 7580
    csBloodGrouping = 7581
    csInfectiousDiseases = 7582
    csHIV1And2Antibodies = 7583
    csHepatitisPyloriAntibodies = 7584
    csInfectiousMononucleosis = 7585
    csRubella = 7586
    csLiverPanel = 7587
    csAlbumin = 7588
    csTotalProtein = 7589
    csAspartateAminotransferase = 7590
    csAlanineAminotransferase = 7591
    csTotalBilirubin = 7592
    csDirectBilirubin = 7593
    csAlkalinePhosphatase = 7594
    csLacticdehydrogenase = 7595
    csCardiacEnzymes = 7596
    csTroponin = 7597
    csMyoglobin = 7598

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csRenalProfile = 7599
    csUricAcid = 7600
    csPhosphate = 7601
    csAnaemiaProfile = 7602
    csIron = 7603
    csFerritin = 7604
    csFolate = 7605
    csElectrolytes = 7606
    csLipidProfile = 7607
    csAmylase = 7608
    csGlycaemiaProfile = 7609
    csFastingGlucose = 7610
    csRandomGlucose = 7611
    csInsulin = 7612
    csHormones = 7613
    csFreeT3 = 7614
    csFreeT4 = 7615
    csThyroglobulin = 7616
    csProgesterone = 7617
    csProlactin = 7618
    csAlphafetoprotein = 7619
    csElectrophoresis = 7620

    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

    csHaemoglobin = 7621
    csUrinalysis = 7622
    csUrineSedimentation = 7623
    csOrganicLiquidBiochemicalExamination = 7624
    csMicroscopicExamination = 7625
    csCultureAndSensitivity = 7626
    csBloodCulture = 7627
    csBacilli = 7628
    csDrugResistance = 7629
    csMycology = 7630
    csCryptococcusInvestigation = 7631
    csCryptococcusAntigen = 7632
    csBlood = 7633
    csSemenAnalysis = 7635
    csCytolologicalExamination = 7636
    csAscites = 7637

    csMultiDrugResistance = 7638
    csResistantTo = 7639

    csRequestedByReception = 7640
    csSignature = 7641
    csInternal = 7642
    csExternal = 7643
    csInternalExternalNotEntered = 7644
    csSec = 7645

    csErythrocytes = 7646

    csRedefineResults = 7647

    csStudyNumber = 7648

    csApprove = 7649
    csNoResults = 7650
    csReady = 7651
    csNoPrinter = 7652
    csTransmitted = 7653
    csPending = 7654

    csHealthCentre = 7655
    csIntermediate = 7656

    csPleaseWait = 7657

    csVaginalCulture = 7658
    csUrineCulture = 7659
    csFaecalCulture = 7660
    csUrineGlucose = 7661

    csMacroscopicExamination = 7662

    csReactive = 7663
    csNonReactive = 7664

    csIndeterminate = 7665

    csPreview = 7666
    csHeading = 7667
    csNewItem = 7668

    csNumberOfResultsPerPage = 7669

    csPageNumber = 7670

    csEpithelial = 7671
    csTrichomonasVaginalis = 7672
    csYeast = 7673
    csFreshExamination = 7674
    csAmoebae = 7675
    csGranules = 7676

    csMacroscopicAppearance = 7677
    csData = 7678

    csPCR = 7679
    csembedding = 7700
    CsCutting = 7701
    CsImmunohistochemical = 7702
    CsPhase = 7703
    CsInHistology = 7704
    CsWithPathalogist = 7705
    CsAwaitingAuthorisation = 7706
    CsAuthorisedNotPrinted = 7707
    CsExtraRequests = 7708
    CsExternalEventsOut = 7709
    CsCellularPathology = 7710
    CsGoToWorksheet = 7711
    CsGross = 7712
    CsAmendments = 7713
    CsAddendum = 7714
    CsMovementTracker = 7715
    CsAuthorised = 7716
    CsPreliminary = 7717
    CsCase = 7718
    CsAudit = 7719
    CsDiscrepancyLog = 7720
    CsAmendId = 7721
    'CsOrientation = 7722
    CsCutBy = 7723
    CsClickOnCaseIDToGetDetails = 7724
    CsWorkLog = 7725
    CsFilterByDate = 7726
    CsCasesUnreportedAfter = 7727
    CsListofallspecimensbetween = 7728
    CsScheduledForDisposalOn = 7729
    CsListOfAllSpecimensScheduledForDisposalOn = 7730
    CsNotScheduledForDisposalOn = 7731
    CsListOfAllKeptSpecimens = 7732
    CsListOfAllDisposedSpecimensBetween = 7733
    CsDisposal = 7734
    CsPercentageOfCases = 7735
    CsReferrals = 7736
    CsListOfCaseIdsOpenedForEditing = 7737
    CsShowOnlyCasesBetweenAboveDatesThatWereAuthorisedSince = 7738
    CsNoOfCases = 7739
    CsM = 7740
    csBF = 7680
    csLabFigursDateWise = 7681

    csListing = 7682
    csLabFigursTestWise = 7683

    csSpermMolarityPercentageCheck = 7684

    ' Masood 12-0913
    csMonthlyResultSummary = 7685
    csMonthlyReportsDateWise = 7686
    csAuthorizePerson = 7687
    csRegisteredSamples = 7688
    csReason = 7689
    csUnvalidate = 7690
    csReceivedOn = 7691
    csCollectionDate = 7692
    csInvalidDateMoreThan12Month = 7693
    csEnteredBy = 7694
    csSediment = 7695
    csPregnancyTIG = 7697
    cscaseid = 7698
    csContainerLabel = 7699
    csAforAutopsy = 7800
    csHforHistology = 7801
    csCforCytology = 7802
    csBlock = 7803
    csTouchPrep = 7804
    csControl = 7805
    csMforMale = 7806
    csFforFemale = 7807
    csUniqueID = 7808
    csTable = 7809
    csPath = 7810
    csReturned = 7811
    csType = 7812
    csReferredTo = 7813
    csReasonForReferral = 7814
    csTissueTypeId = 7815
    csTissueTypeLetter = 7816
    csTissuePath = 7817
    csTissueTypeListId = 7818
    csAddTissueType = 7819
    csAddCutUpDetails = 7820
    csOpen = 7821
    csDisposeCase = 7822
    csAllEmbedded = 7823
    csAddFrozenSection = 7824
    csAddTouchPrep = 7825
    csAddSingleBlock = 7826
    csAddMultipleBlock = 7827
    csAddSingleSlide = 7828
    csAddMultipleSlide = 7829
    csReferral = 7830
    csAddNumberOfLevels = 7831
    csAddRoutineStain = 7832
    csAddSpecialStain = 7833
    csAddImmunohistochemicalStain = 7834
    csAddExtraLevels = 7835
    csPrintToBlockNumber = 7836
    csAddControl = 7837
    csPCodes = 7838
    csMCodes = 7839
    csQCodes = 7840
    csTCodes = 7841
    csStains = 7842
    csCodes = 7843
    csSpecial = 7844
    csImmunoHistoChemicalStain = 7845
    csCoroners = 7846
    csCounty = 7847
    csOrientation = 7848
    csProcessor = 7849
    csDiscrepancy = 7850
    csAccreditationSettings = 7851
    csNonWorkDays = 7852
    csWorkLogs = 7853
    csDisposals = 7854
    CsAutopsy = 7855
    csByTissueType = 7856
    csDiagnosisSpecificTotals = 7857
    csDiagnosisRangeSearch = 7858
    csGroupedTissueSearch = 7859
    csLocationSpecificSearch = 7860
    csLockedForEditing = 7861
    csNumerical = 7862
    csTurnAroundTime = 7863
    csAuthorisedReports = 7864
    csLoggedin = 7865
    csRecordBeingEditedByYou = 7866
    csRecordBeingEditedBy = 7867
    csPleaseEnterCaseNumber = 7868
    csPleaseSelectANodeFirst = 7869
    csPleaseFillInMandatoryFields = 7870
    csPleaseSelectAPathologist = 7871
    csCaseIDFormatIncorrect = 7872
    csNoDemographicAvailable = 7873
    csRecordSaved = 7874
    csMustEnterPCodeBeforeAuthorisation = 7875
    csMustEnterMCodeBeforeAuthorisation = 7876
    csAreYouSureYouWantToDelete = 7877
    csCaseIDAlreadyExists = 7878
    csPleaseEnterNumberOfLevels = 7879
    csPleaseEnterNumberOfBlocks = 7880
    csPleaseEnterNumberOfSlides = 7881
    csCheckedBy = 7882
    csCutUpBy = 7883
    csPiecesAtEmbedding = 7884
    csAssistedBy = 7885
    csPiecesAtCutUp = 7886
    csEmbeddedBy = 7887
    csTissueID = 7888
    csTissueType = 7889
    csChanges = 7890
    csEvents = 7891
    csLoggedInUser = 7892
    
    
    
    
    
    ' Masood 12-0913
    ''''''''''''''''
    'After adding or amending any entry
    'Add EnsureLanguageEntryExists() line in modLanguage\CheckUpdateLanguage()
    'DONT MAKE CHANGES TO THIS FILE IN ANY PROJECT OTHER THAN NetAcquire
    ''''''''''''''''

End Enum

Attribute VB_Name = "ModGlobal"
Public Sub frmListsGeneric_ChangeLanguage(TypeName As String, TypeNames As String)


On Error GoTo frmListsGeneric_ChangeLanguage_Error

Select Case TypeName
    Case "P Code"
        TypeName = "P" & " " & "Code"
        TypeNames = "PCodes"
    Case "M Code"
        TypeName = "M" & " " & "Code"
        TypeNames = "MCodes"
    Case "T Code"
        TypeName = "T" & " " & "Code"
        TypeNames = "TCodes"
    Case "Q Code"
        TypeName = "Q " & "Code"
        TypeNames = "QCodes"
    Case "Clinician"
        TypeName = "Clinician"
        TypeNames = "Clinicians"
    Case "Coroner"
        TypeName = "Coroners"
        TypeNames = "Coroners"
    Case "County"
        TypeName = "County"
        TypeNames = "County"
    
    
End Select

With frmListsGeneric

    .FrameAdd.Caption = "Add"
    .FrameAdd.Caption = "Add" & " " & TypeName
    .Caption = "Histology --- " & "Lists"
    .Caption = .Caption & " (" & TypeNames & ")"
    .lblSource.Caption = "Source"
    .Label1.Caption = "Code"
    .Label2.Caption = "Text"
    .Label9.Caption = "Login"
    .cmdAdd.Caption = "Add"
    .cmdDelete.Caption = "Delete"
    .cmdSave.Caption = "Save"
    .cmdExit.Caption = "Exit"
End With



Exit Sub


frmListsGeneric_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmListsGeneric_ChangeLanguage", intEL, strES

End Sub
Public Sub frmHistDisposal_ChangeLanguage()

On Error GoTo frmHistDisposal_ChangeLanguage_Error


With frmHistDisposal
    .optList(0).Caption = "List of all specimens between"
    .Label5.Caption = "and"
    .Label4.Caption = "Scheduled For Disposal On"
    .optList(1).Caption = "List Of All Specimens Scheduled For Disposal On)"
    .optList(2).Caption = "List of all specimens between)"
    .Label2.Caption = "Sand"
    .Label3.Caption = "Not Scheduled For Disposal On"
    .optList(3).Caption = "List Of AllKept Specimens"
    .optList(4).Caption = "List Of All DisposedSpecimens Between)"
    .Label1.Caption = "Sand"
    .cmdReloadList.Caption = "Refresh"
    .cmdPrint.Caption = "Print"
    .cmdSave.Caption = "Save"
    .cmdExit.Caption = "Exit"
End With



Exit Sub


frmHistDisposal_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmHistDisposal_ChangeLanguage", intEL, strES

End Sub
Public Sub frmReferralLog_ChangeLanguage()
On Error GoTo frmReferralLog_ChangeLanguage_Error


With frmReferralLog
    .fraDates.Caption = "Between Dates"
    .Label4.Caption = "From"
    .Label2.Caption = "to"
    .optSelect(0).Caption = "All"
    .optSelect(1).Caption = "Completed"
    .optSelect(2).Caption = "Not" & " " & "Completed"

    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "ExporttoExcel"
    .cmdPrint.Caption = "Print"
    .cmdExit.Caption = "Exit"
    .Caption = "Referrals"
End With


Exit Sub


frmReferralLog_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmReferralLog_ChangeLanguage", intEL, strES

End Sub
Public Sub frmCasesCurrentlyOpened_ChangeLanguage()
On Error GoTo frmCasesCurrentlyOpened_ChangeLanguage_Error


With frmCasesCurrentlyOpened
    .Label.Caption = "List Of CaseIds Opened For Editing"
    .cmdReloadList.Caption = "Refresh List"
    .cmdRemove.Caption = "Remove"
    .cmdExit.Caption = "Exit"
End With


Exit Sub


frmCasesCurrentlyOpened_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmCasesCurrentlyOpened_ChangeLanguage", intEL, strES

End Sub

Public Sub frmNCRI_ChangeLanguage()
On Error GoTo frmNCRI_ChangeLanguage_Error


With frmNCRI
    .panSampleDates.Caption = "sample" & " " & "Between Dates"
    .Label1(2).Caption = "From"
    .Label1(3).Caption = "to"
    .chkAuthSince.Caption = "Show Only Cases Between Above Dates That Were Authorised Since"
    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "ExporttoExcel"
    .cmdExit.Caption = "Exit"
    .Label9.Caption = "Loggedin"
End With


Exit Sub


frmNCRI_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmNCRI_ChangeLanguage", intEL, strES

End Sub
Public Sub frmSnomedSearch_ChangeLanguage()
On Error GoTo frmSnomedSearch_ChangeLanguage_Error


With frmSnomedSearch
    .panSampleDates.Caption = "Sample" & " " & "Between Dates"
    .Label1(2).Caption = "From"
    .Label1(3).Caption = "to"
    .fraTCode.Caption = "T" & " " & "Code"
    .fraMCode.Caption = "M" & " " & "Code"
    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "Export to Excel"
    .cmdExit.Caption = "Exit"
    .Label9.Caption = "Logged in"
End With


Exit Sub


frmSnomedSearch_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmSnomedSearch_ChangeLanguage", intEL, strES

End Sub
Public Sub frmNumericalStats_ChangeLanguage()
On Error GoTo frmNumericalStats_ChangeLanguage_Error


With frmNumericalStats
    .Frame1.Caption = "Sample" & " " & "Between Dates"
    .Label4.Caption = "From"
    .Label2.Caption = "to"
    .Label1.Caption = "Source"
    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "ExporttoExcel"
    .cmdExit.Caption = "Exit"
    .Label9.Caption = "Logged in"
End With


Exit Sub


frmNumericalStats_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmNumericalStats_ChangeLanguage", intEL, strES

End Sub
Public Sub frmCutUpEmbed_ChangeLanguage()

On Error GoTo frmCutUpEmbed_ChangeLanguage_Error

PrintControlNamesAndCaptions frmCutUpEmbed

With frmCutUpEmbed

    .Label7.Caption = "Cut Up By"
    .Label3.Caption = "Pieces At Embedding"
    .Label4.Caption = "Embedded By"

    .lblProcessor.Caption = "Processor"
    .Label5.Caption = "Cut Up By"
    .Label6.Caption = "Assisted By"

    .Label1.Caption = "Pieces At Cut Up"
    .Label2.Caption = "Orientation"


    .Label10.Caption = "Logged in"
'    .Label1.Caption = "CheckedBy"
    .cmdAdd.Caption = "Update"
    .cmdExit.Caption = "Exit"
    .cmdPrint.Caption = "Print"
    .cmdReloadList.Caption = "Refresh"
    .cmdSave.Caption = "Save"
    '.lblCaption.Caption = "With Pathalogist"
End With

Exit Sub

frmCutUpEmbed_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmCutUpEmbed_ChangeLanguage", intEL, strES
End Sub

Public Sub frmWithPathologist_ChangeLanguage()

On Error GoTo frmWithPathologist_ChangeLanguage_Error


With frmWithPathologist
    .Caption = "NetAcquire - " & "Histology"
    .Label1.Caption = "CheckedBy"
    .lblCaption.Caption = "With Pathologist"

End With

Exit Sub

frmWithPathologist_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmWithPathologist_ChangeLanguage", intEL, strES
End Sub

Public Sub frmPhase_ChangeLanguage()

On Error GoTo frmPhase_ChangeLanguage_Error

With frmPhase
    .Caption = "Phase"
    .lblWorkPhase(0).Caption = "Cut Up"
    .lblWorkPhase(1).Caption = "Embedding"
    .lblWorkPhase(2).Caption = "Cutting"
    .lblWorkPhase(3).Caption = "Immuno histo chemical"
    .lblWorkPhase(4).Caption = "Specials"
    .lblWorkPhase(5).Caption = "Cytology"
    .Label9.Caption = "Loggedin"
End With

Exit Sub


frmPhase_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmPhase_ChangeLanguage", intEL, strES
End Sub
Public Sub frmWorklist_ChangeLanguage()
On Error GoTo frmWorklist_ChangeLanguage_Error



With frmWorklist
    .Caption = "NetAcquire - " & "Cellular Pathology"
    .cmdInHistology.Caption = "In Histology"
    .lblWorkList(0).Caption = "In Histology"
    
    .lblWorkList(1).Caption = "With Pathologist"
    .lblWorkList(2).Caption = "Awaiting Authorisation"
    .lblWorkList(3).Caption = "Authorised Not Printed"
    .lblWorkList(4).Caption = "Extra Requests"
    .lblWorkList(5).Caption = "External Events Out"

    .cmdRefresh.Caption = "Refresh List"
    .cmdExtensiveSearch.Caption = "Search"
    .cmdAddDemo.Caption = "Demographics"
    .cmdWorkSheet.Caption = "GoToWorksheet"

    .Label9.Caption = "Loggedin"


    .mnuFile.Caption = "File"
    .mnuLogOff.Caption = "LogOff"
    .mnuAudit.Caption = "Audit"
    

    .mnuExit.Caption = "Exit"
    .mnuLists.Caption = "Lists"
    .mnuPCodes.Caption = "PCodes"
    .mnuMCodes.Caption = "MCodes"
    .mnuQCodes.Caption = "QCodes"
    .mnuTCodes.Caption = "TCodes"
    .mnuReferrals.Caption = "Referral"
    .mnuReferredTo.Caption = "ReferredTo"
    .mnuReasonReferral.Caption = "Reason For Referral"
    .mnuDestStains.Caption = "Stains"
    .mnuDestCodes.Caption = "Codes"
    
    .mnuStains.Caption = "Stains"
    
    .mnuRoutineStainList.Caption = "Routine"
    
    .mnuSpecialStainList.Caption = "Special"
    .mnuImmunoStainList.Caption = "Immuno histo chemical"
    .mnuOther.Caption = "Other"
    .mnuWards.Caption = "Wards"
    .mnuCoroners.Caption = "Coroners"
    .mnuClinicians.Caption = "Clinicians"
    .mnuGPs.Caption = "GP"    's
    .mnuCounty.Caption = "County"
    .mnuSource.Caption = "Source"
    .mnuOrientation.Caption = "Orientation"
    .mnuProcessor.Caption = "Processor"
    .mnuDiscrepancy.Caption = "Discrepancy"
    .mnuDiscrepancyType.Caption = "Type"
    .mnuDiscrepancyResolution.Caption = "Resolution"
    .mnuAccreditationSettings.Caption = "Accreditation Settings"
    .mnuNonWorkingDays.Caption = "Non Work Days"
    .mnuWorkLogs.Caption = "Work Logs"
    .mnuCutUpSheet.Caption = "Cut Up"
    .mnuEmbedSheet.Caption = "Embedding"
    .mnuCutting.Caption = "Cutting"
    .mnuDisposal.Caption = "Disposals"
    .mnuHistDisposal.Caption = "Histology"
    .mnuCytDisposal.Caption = "Cytology"
    .mnuAutDisposal.Caption = "Autopsy"
    .mnuReferral.Caption = "Referral"
    .mnuLocked4Editing.Caption = "Locked For Editing"
    .mnuReports.Caption = "Reports"
    .mnuNcri.Caption = "NCRI"
    .mnuSnomed.Caption = "Search"
    .mnuSearch1.Caption = "By Tissue Type"
    .mnuSearch2.Caption = "Diagnosis Specific Totals"
    .mnuSearch3.Caption = "Diagnosis Range Search"
    .mnuSearch4.Caption = "Grouped Tissue Search"
    .mnuSearch5.Caption = "Location Specific Search"
    .mnuStats.Caption = "Statistics"
    .mnuNumerical.Caption = "Numerical"
    .mnuTAT.Caption = "TurnAround Time"
    .mnuTatPCodes.Caption = "PCodes"
    .mnuTatTCodes.Caption = "TCodes"
    .mnuTATPCodeCaseIds.Caption = "PCodes" & " (" & "caseid" & ")"
    .mnuTATTCodeCaseIds.Caption = "TCodes" & " (" & "caseid" & ")"
    .mnuRptDiscrep.Caption = "Discrepancy"
    .mnuAuthorisedReports.Caption = "Authorised Reports"
    .mnuAbout.Caption = "About"
    .mnuLanguage.Caption = "Language"
    .mnuLanguageEnglish.Caption = "English"
    .mnuLanguageRussian.Caption = "Russian"
    .mnuLanguagePortuguese.Caption = "Portuguese"

End With

Exit Sub


frmWorklist_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmWorklist_ChangeLanguage", intEL, strES

End Sub
Public Sub frmWorkSheet_ChangeLanguage()

On Error GoTo frmwor_ChangeLanguage_Error



With frmWorkSheet
    .Caption = "Net Acquire -" & "CellularPathology"
    .Label2.Caption = "caseid"
    .Label1.Caption = "Patient" & " " & "Identification"
    .Label3.Caption = "P" & " " & "Code"
    .Label4.Caption = "Gross"
    .Label5.Caption = "Micro"
    .Label9.Caption = "Loggedin"
    .Label8.Caption = "Q" & " " & "Code"
    .Label6.Caption = "M" & " " & "Code"
    .Label7.Caption = "Addendum" & "" & "Amendments"
    .Label12.Caption = "SampleDate"
    .Label13.Caption = "sample" & "Received"
    .Label14.Caption = "Preliminary" & " " & "Report" & " " & "Date"
    .Label15.Caption = "Authorised" & " " & "Report" & " " & "Date"
    .Label16.Caption = "Movement Tracker"
    .Label17.Caption = "Nature Of Specimen"
    .Label18.Caption = "Container Label"
    .fraCaseState.Caption = "caseid" & " " & "Status"
    .optState(0).Caption = "In Histology"
    .optState(1).Caption = "With Pathologist"
    .optState(2).Caption = "Awaiting Authorisation"
    .optReport(0).Caption = "Preliminary" & " " & "Report"
    .optReport(1).Caption = "Authorised" & " " & "Report"
    .SSTabMovement.TabCaption(0) = "Specimen"
    .SSTabMovement.TabCaption(1) = "Stain"
    .SSTabMovement.TabCaption(2) = "Case"
    .SSTabMovement.TabCaption(3) = "Blocks" & "" & "Slide"

    .cmdClear.Caption = "Clear"
    .cmdSearch.Caption = "Search"
    .cmdEditDemo.Caption = "Edit"
    .cmdCytoHist.Caption = "Histology"
    .cmdAudit.Caption = "Audit"
    .cmdClinicalHist.Caption = "Clinical Details"
    .cmdScanOrder.Caption = "Scan" & " " & "Order"
    .cmdViewScans.Caption = "View Scan"
    .cmdPrnReport.Caption = "Print Reports"
    .cmdDiscrepancyLog.Caption = "Discrepancy Log"
    .cmdSave.Caption = "Save"
    .cmdExit.Caption = "Exit"
    .cmdPrnPreview.Caption = "PrintReview"
    .cmdComments.Caption = "Comments"

    '.mnuMCodesMenu = MCodesMenu
    .mnuMCodesDel.Caption = "Delete"
    '.mnuQCodesMenu = QCodesMenu
    .mnuQCodesDel.Caption = "Delete"
    '.mnuAmendMenu = AmendmentMenu
    .mnuAmendDel.Caption = "Delete"
    '.mnuMoveSpecMenu = MoveSpecMenu
    .mnuMoveSpecDel.Caption = "Delete"
    '.mnuPopUpLevel1 = PopUpLevel1
    .mnuAddTissueType.Caption = "Add Tissue Type"
    .mnuAddCutUp.Caption = "Add Cut Up Details"
    .mnuOpen.Caption = "Open"
    .mnuDisposeCase.Caption = "Dispose Case"
    '.mnuPopupLevel2 = PopUpLevel2
    .mnuEditTissueType.Caption = "Edit"
    .mnuAllEmbedded.Caption = "All Embedded"
    .mnuReferral.Caption = "Referral"
    .mnuFrozenSection.Caption = "Add Frozen Section"
    .mnuTouchPrep.Caption = "Add Touch Prep"
    .mnuSingleBlock.Caption = "Add Single Block"
    .mnuMultipleBlocks.Caption = "Add Multiple Block"
    .mnuSingleSlideLevel2.Caption = "Add Single Slide"
    .mnuMultipleSlidesLevel2.Caption = "Add Multiple Slide"
    '.mnuSeperator3 = -
    .mnuDelTissueType.Caption = "Delete"
    '.mnuPopupLevel3 = PopupLevel3
    .mnuBlockReferral.Caption = "Referral"
    '.mnuSeperator10 = -
    .mnuSingleSlideLevel3.Caption = "Add Single Slide"
    .mnuMultipleSlidesLevel3.Caption = "Add Multiple Slide"
    .mnuNoOfLevelsLevel3.Caption = "Add Number Of Levels"
    .mnuAddControlLevel3.Caption = "Add Control"
    .mnuRoutineStainLevel3.Caption = "Add Routine Stain"
    .mnuSpecialStainLevel3.Caption = "Add Special Stain"
    .mnuImmunoStainLevel3.Caption = "Add Immuno histo chemical Stain"
    .mnuAddExtraLevels.Caption = "AddExtraLevels"
    .mnuPrnBlockNumber.Caption = "Print To Block Number"
    '.mnuSeperator7 = -
    .mnuDelBlock.Caption = "Delete"
    '.mnuPopupLevel4 = PopupLevel4
    .mnuAddControlLevel4.Caption = "Add Control"
    .mnuRoutineStainLevel4.Caption = "Add Routine Stain"
    .mnuSpecialStainLevel4.Caption = "Add Special Stain"
    .mnuImmunoStainLevel4.Caption = "Add Immuno histochemical Stain"
    .mnuNoOfLevelsLevel4.Caption = "Add Number Of Levels"
    '.mnuSeperator8 = -
    .mnuDelSlide.Caption = "Delete"
    '.mnuPopupLevel5 = PopupLevel5
    .mnuDelStain.Caption = "Delete"





End With


Exit Sub


frmwor_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmwor_ChangeLanguage", intEL, strES

End Sub


Public Sub frmChangePass_ChangeLanguage()

On Error GoTo frmChangePass_ChangeLanguage_Error

With frmChangePass

    .Label2 = "ConfirmPassword"
    .Label1 = "EnterNewPassWord"
End With
Exit Sub

frmChangePass_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmChangePass_ChangeLanguage", intEL, strES, sql

End Sub

Public Sub frmSystemManager_ChangeLanguage()


On Error GoTo frmSystemManager_ChangeLanguage_Error

With frmSystemManager

    .Caption = "NetAcquire - " & "Login"
    .fraLogin.Caption = "Login"
    .Label6.Caption = "UserName"
    .Label7.Caption = "PassWord"
    .cmdOK.Caption = "Login"
    .cmdHide.Caption = "Exit"
    .Frame1.Caption = "Add New Operator"
    .Label3.Caption = "First Name"
    .Label2.Caption = "Surname"
    .Label1.Caption = "PassWord"
    .Label11.Caption = "Code"
    .Label4.Caption = "Confirm Password "
    .lblMCRN.Caption = "MCRN"
    .Label8.Caption = "Access Rights"
    .Label5.Caption = "Auto LogOff in"
    .Label9.Caption = "Minutes"
    .cmdSave.Caption = "Save"
    .cmdEdit.Caption = "Edit"
    .cmdExit.Caption = "Exit"
End With

Exit Sub

frmSystemManager_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmSystemManager_ChangeLanguage", intEL, strES
End Sub


Public Sub frmSearch_ChangeLanguage()

On Error GoTo frmSearch_ChangeLanguage_Error

With frmSearch
    .Caption = "Search"
    .fraSurname.Caption = "Surname"
    .fraForename.Caption = "FirstName"
    .Frame4.Caption = "SearchFor"
    .fraSearch.Caption = "How"
    .optFor(0).Caption = "Surname"
    .optFor(1).Caption = "Chart"
    .optFor(2).Caption = "DoB"
    .optFor(3).Caption = "CaseId"
    .optFor(4).Caption = "First Name"
    .optFor(5).Caption = "Surname" & " + " & "First Name"
    .optExact.Caption = "ExactMatch"
    .optLeading.Caption = "LeadingCharacters"
    .chkType.Caption = "As You Type"
    .cmdSearch.Caption = "Search"
    .cmdExit.Caption = "Exit"
    .bcopy.Caption = "Copy to Edit"
    .cmdAudit.Caption = "Audit Trail"
    .lNoPrevious.Caption = "No Previous Details"


End With

Exit Sub

frmSearch_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmSearch_ChangeLanguage", intEL, strES
End Sub

Public Sub frmDemographics_ChangeLanguage()

On Error GoTo frmDemographics_ChangeLanguage_Error

With frmDemographics
    .Caption = "NetAcquire - " & "Demographics"
    .Label9.Caption = "CaseId"
    .Label2.Caption = "Surname"
    .Label1.Caption = "FirstName"
    .Label3.Caption = "Address"
    .lblSampleTaken.Caption = "sample Date"
    .Label10.Caption = "ChartNo"
    .Label14.Caption = "sample Received"
    .Label5.Caption = "Date of Birth"
    .Label6.Caption = "Age"
    .Label7.Caption = "Sex"
    .Label8.Caption = "Phone"
    .lblCoronerClin.Caption = "Clinician"
    .lblMothersName.Caption = "Mothers Name"
    .lblMothersDOB.Caption = "Date of Birth"
    .lblWardGP.Caption = "Ward"
    .lblClinician.Caption = "Clinician"
    .Label4.Caption = "Region"
    .Label25.Caption = "Clinical Details"
    .Label26.Caption = "Login"
    .Label13.Caption = "Patient Comments"
    .Label12.Caption = "Nature Of Specimen"
    .cmdNew.Caption = "New"
    .cmdSave.Caption = "Save"
    .cmdExit.Caption = "Exit"
    .chkUrgent.Caption = "Urgent"
    .Label11.Caption = "Source"
    .Label24.Caption = "Container Label"

End With

Exit Sub

frmDemographics_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "basGlobalTemp", "frmDemographics_ChangeLanguage", intEL, strES
'frmAllEquipments_ChangeLanguage

End Sub

Public Sub frmCaseEventLog_ChangeLanguage()

On Error GoTo frmCaseEventLog_ChangeLanguage_Error

With frmCaseEventLog

    .cmdSearch.Caption = "Search"
    .cmdExit.Caption = "Exit"
    .cmdPrint.Caption = "Print"
    .lblChanges.Caption = "Changes"
    .lblEvents.Caption = "Events"
    .lblAll.Caption = "All"
    .Label1.Caption = "caseid"
    .Label10.Caption = "Loggedin"
    
    
    
    
End With

Exit Sub

frmCaseEventLog_ChangeLanguage_Error:

 Dim strES As String
 Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "ModGlobal", "frmCaseEventLog_ChangeLanguage", intEL, strES, sql
End Sub

Public Sub ChangeFont(frm As Form, FontName As String)

Dim X As Control

On Error GoTo ChangeFont_Error

With frm
    .Font = FontName
    For Each X In .Controls
        If TypeOf X Is Label Or TypeOf X Is CommandButton _
           Or TypeOf X Is MSFlexGrid Or TypeOf X Is Frame _
           Or TypeOf X Is OptionButton Or TypeOf X Is CheckBox _
           Or TypeOf X Is TextBox Or TypeOf X Is RichTextBox _
           Or TypeOf X Is ComboBox Or TypeOf X Is ListBox _
           Or TypeOf X Is TreeView Or TypeOf X Is MaskEdBox Then

            X.Font = "Arial"
        ElseIf TypeOf X Is DTPicker Or TypeOf X Is DataGrid Then
            X.Font.Name = "Arial"
        End If
    Next
End With
Exit Sub

ChangeFont_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "ChangeFont", intEL, strES

End Sub

Public Sub PrintControlNamesAndCaptions(frm As Form)
Dim X As Control

For Each X In frm.Controls
    If TypeOf X Is Label Or TypeOf X Is Frame Or _
       TypeOf X Is OptionButton Or TypeOf X Is CheckBox Or _
       TypeOf X Is Menu Then
        Debug.Print "." & X.Name & ".Caption = " & X.Caption
    End If
Next
End Sub


Public Sub frmtat_ChangeLanguage()
On Error GoTo frmtat_ChangeLanguage_Error


With frmTAT
    .fraDates.Caption = "Between Dates"
    .Label4.Caption = "From"
    .Label2.Caption = "to"
    .lblPathologist.Caption = "Pathologist"
    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "Export to Excel"
    .cmdPrint.Caption = "Print"
    .cmdExit.Caption = "Exit"
    .optCases(0).Caption = "No Of Cases"
    .optCases(1).Caption = "Percentage Of Cases"
End With

 
Exit Sub

 
frmtat_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmtat_ChangeLanguage", intEL, strES

End Sub
Public Sub frmTATcases_ChangeLanguage()
On Error GoTo frmTATcases_ChangeLanguage_Error


With frmTATcases
    .fraDates.Caption = "Between Dates"
    .Label4.Caption = "From"
    .Label2.Caption = "to"
    .lblPathologist.Caption = "Pathologist"
    .lblTissueType.Caption = "Tissue Type"
    .cmdCalc.Caption = "Calculate"
    .cmdExport.Caption = "Export to Excel"
    .cmdPrint.Caption = "Print"
    .cmdExit.Caption = "Exit"
    .lblUnReported1.Caption = "Cases Unreported After"
    .lblDays1.Caption = "Days"
End With

 
Exit Sub

 
frmTATcases_ChangeLanguage_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ModGlobal", "frmTATcases_ChangeLanguage", intEL, strES

End Sub
Public Sub frmDiscrepancyReport_ChangeLanguage()
10    On Error GoTo frmDiscrepancyReport_ChangeLanguage_Error


20    With frmDiscrepancyReport
30        .Frame1.Caption = "Between Dates"
40        .Label4.Caption = "From"
50        .Label2.Caption = "to"
60        .Label1.Caption = "Discrepancy"
70        .Label9.Caption = "Loggedin"
80        .cmdCalc.Caption = "Calculate"
90        .cmdExport.Caption = "Export to Excel"
100       .cmdPrint.Caption = "Print"
110       .cmdExit.Caption = "Exit"
120   End With

       
130   Exit Sub

       
frmDiscrepancyReport_ChangeLanguage_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "ModGlobal", "frmDiscrepancyReport_ChangeLanguage", intEL, strES

End Sub
Public Sub frmAuthorisedReports_ChangeLanguage()
10    On Error GoTo frmAuthorisedReports_ChangeLanguage_Error


20    With frmAuthorisedReports
30        .Label2.Caption = "Filter By Date"
40        .Label1(2).Caption = "From"
50        .Label1(3).Caption = "to"
60        .Label3.Caption = "Advanded Filter"
70        .Label4.Caption = "Click On CaseID To Get Details"
80        .cmdSearch.Caption = "Search"
90        .cmdPrint.Caption = "Print"
100       .cmdExcel.Caption = "Export to Excel"
110       .cmdExit.Caption = "Exit"
120   End With

       
130   Exit Sub

       
frmAuthorisedReports_ChangeLanguage_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "ModGlobal", "frmAuthorisedReports_ChangeLanguage", intEL, strES, sql

End Sub

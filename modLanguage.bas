Attribute VB_Name = "modLanguage"
Option Explicit


Public colLanguage As Collection


Public Sub CheckUpdateLanguage()

'Added 16/10/2009
'    EnsureLanguageEntryExists 1097, "Test Code Mapping", "Mapeamento do Codigo de Teste"
'    EnsureLanguageEntryExists 2783, "Date / Time", ""
'    EnsureLanguageEntryExists 7640, "Requested by Reception", "Solicitado pela Recepção"
'    EnsureLanguageEntryExists 7641, "Signature", "Assinatura"
'    EnsureLanguageEntryExists 7642, "Internal", "Interno"
'    EnsureLanguageEntryExists 7643, "External", "Externo"
'    EnsureLanguageEntryExists 7644, "Internal/External not Entered", "Interno/Externo não informada !"
'    EnsureLanguageEntryExists 7645, "Sec", "Seg"
'    EnsureLanguageEntryExists 7646, "Erythrocytes", "Eritrócitos"
'    EnsureLanguageEntryExists 7647, "Re-define Results", "Re-definir Resultados"
'    EnsureLanguageEntryExists 7648, "Study Number", "Numero de Ensaio"
'    EnsureLanguageEntryExists 7649, "Approve", "Aprovar"
'    EnsureLanguageEntryExists 7650, "No Results", "Sem Resultados"
'    EnsureLanguageEntryExists 7651, "Ready", "Pronto"
'    EnsureLanguageEntryExists 7652, "No Printer", "Sem Impressora"
'    EnsureLanguageEntryExists 7653, "Transmitted", "Transmitidos"
'    EnsureLanguageEntryExists 7654, "Pending", "Pendente"
'    EnsureLanguageEntryExists 7655, "Health Center", "Unidade Sanitária"
'    EnsureLanguageEntryExists 7656, "Intermediate", "Intermédio"
'    EnsureLanguageEntryExists 7657, "Please Wait", "Espere por favor"
'    EnsureLanguageEntryExists 7658, "Vaginal Culture", "Cultura Vaginal"
'    EnsureLanguageEntryExists 7659, "Urine Culture", "Urocultura"
'    EnsureLanguageEntryExists 7660, "Faecal Culture", "Coprocultura"
'    EnsureLanguageEntryExists 7661, "Urine Glucose", "Glicosuria"
'    EnsureLanguageEntryExists 7662, "Macroscopic Examination", "Exame macroscópico"
'    EnsureLanguageEntryExists 7663, "Reactive", "Reactivo"
'    EnsureLanguageEntryExists 7664, "Non Reactive", "Não Reactivo"
'    EnsureLanguageEntryExists 7665, "Indeterminate", "Indeterminado"
'    EnsureLanguageEntryExists 7666, "Preview", "Visualização"
'    EnsureLanguageEntryExists 7667, "Heading", "Titulo"
'    EnsureLanguageEntryExists 7668, "New Item", "Novo Item"
'    EnsureLanguageEntryExists 7669, "Number Of Results Per Page", "Número de Resultados por Página"
'    EnsureLanguageEntryExists 7670, "Page Number", "Página Número"
'    EnsureLanguageEntryExists 7671, "Epithelial Cells", "Células epiteliais"
'    EnsureLanguageEntryExists 7672, "Trichomonas vaginalis", "Trichomonas vaginalis"
'    EnsureLanguageEntryExists 7673, "Yeast", "Leveduras"
'    EnsureLanguageEntryExists 7674, "Fresh Examination", "Fresco Exame"
'    EnsureLanguageEntryExists 7675, "Amoebae", "Amoebae"
'    EnsureLanguageEntryExists 7676, "Granules", "Grânulos"
'    EnsureLanguageEntryExists 7677, "Macroscopic Appearance", "Aspecto Macroscopico"
'    EnsureLanguageEntryExists 7678, "Data", "dados"
'    EnsureLanguageEntryExists 7679, "PCR", "RCP"
'    EnsureLanguageEntryExists 7680, "BF", "FB"
'    EnsureLanguageEntryExists 7681, "Monthly Total Count By Hospital", "mensais contagem total pelo Hospital"
'    EnsureLanguageEntryExists 7682, "Lisiting", "listagem"
'    EnsureLanguageEntryExists 7683, "Monthly Test Counts By Hospital", "Contagens mensais de teste pelo Hospital"
'    EnsureLanguageEntryExists 7684, "Motility total should be 100%. Please correct", "Motilidade total deve ser de 100%. Por favor, corrija"
'
'    EnsureLanguageEntryExists 7685, "Monthly Test Counts", "Contagens de teste mensais"
'    EnsureLanguageEntryExists 7686, "Monthly Test Counts By Day", "Contagens mensais de teste por dia"
'    EnsureLanguageEntryExists 7687, "Consultants", "consultores"
'    EnsureLanguageEntryExists 7688, "Registered Samples", "amostras registradas"
'    EnsureLanguageEntryExists 7689, "Reason", "razão"
'    EnsureLanguageEntryExists 7690, "Unvalidate", "Unvalidate"
'    EnsureLanguageEntryExists 7691, "Received On", "Recebido em"
'    EnsureLanguageEntryExists 7692, "Collection Date", "Data de colheita"
'    EnsureLanguageEntryExists 7693, "Statistics cannot be created for more than 12 months", "As estatísticas não podem ser criados por mais de 12 meses"
'    EnsureLanguageEntryExists 7694, "Entered By", "Introduzido por"
'    EnsureLanguageEntryExists 7695, "Sediment", "Sedimento"
'    EnsureLanguageEntryExists 7696, "Other", "Outros"
'    EnsureLanguageEntryExists 7697, "Pregnancy (TIG)", "TIG"
'    EnsureLanguageEntryExists 7698, "Case ID", "caso id"
'    EnsureLanguageEntryExists 7699, "Container Label", ""
'    EnsureLanguageEntryExists 7700, "Embedding", ""
'    EnsureLanguageEntryExists 7701, "Cutting", ""
'    EnsureLanguageEntryExists 7702, "Immunohistochemical", ""
'    EnsureLanguageEntryExists 7703, "Phase", ""
'    EnsureLanguageEntryExists 7704, "In Histology", ""
'    EnsureLanguageEntryExists 7705, "With Pathologist", ""
'    EnsureLanguageEntryExists 7706, "Awaiting Authorisation", ""
'    EnsureLanguageEntryExists 7707, "Authorised Not Printed", ""
'    EnsureLanguageEntryExists 7708, "Extra Requests", ""
'    EnsureLanguageEntryExists 7709, "External Events Out", ""
'    EnsureLanguageEntryExists 7710, "Cellular Pathology", ""
'    EnsureLanguageEntryExists 7711, "Go To Worksheet", ""
'    EnsureLanguageEntryExists 7712, "Gross", ""
'    EnsureLanguageEntryExists 7713, "Amendments", ""
'    EnsureLanguageEntryExists 7714, "Addendum", ""
'    EnsureLanguageEntryExists 7715, "Movement Tracker", ""
'    EnsureLanguageEntryExists 7716, "Authorised", ""
'    EnsureLanguageEntryExists 7717, "Preliminary", ""
'    EnsureLanguageEntryExists 7718, "Case", ""
'    EnsureLanguageEntryExists 7719, "Audit", ""
'    EnsureLanguageEntryExists 7720, "Discrepancy Log", ""
'    EnsureLanguageEntryExists 7721, "Amend Id", ""
'    'EnsureLanguageEntryExists 7722, "Orientation", ""
'    EnsureLanguageEntryExists 7723, "Cut By", ""
'    EnsureLanguageEntryExists 7724, "Click on Case ID to get details", ""
'    EnsureLanguageEntryExists 7725, "Work Log", ""
'    EnsureLanguageEntryExists 7726, "Filter By Date", ""
'    EnsureLanguageEntryExists 7727, "Cases unreported after", ""
'    EnsureLanguageEntryExists 7728, "List of all specimens between", ""
'    EnsureLanguageEntryExists 7729, "scheduled for disposal on", ""
'    EnsureLanguageEntryExists 7730, "List of all specimens scheduled for disposal on", ""
'    EnsureLanguageEntryExists 7731, "Not scheduled for disposal on", ""
'    EnsureLanguageEntryExists 7732, "List of all Kept specimens", ""
'    EnsureLanguageEntryExists 7733, "List of all disposed specimens between", ""
'    EnsureLanguageEntryExists 7734, "Disposal", ""
'    EnsureLanguageEntryExists 7735, "percentage of Cases", ""
'    EnsureLanguageEntryExists 7736, "Referrals", ""
'    EnsureLanguageEntryExists 7737, "List of Case Ids opened for editing", ""
'    EnsureLanguageEntryExists 7738, "Show Only Cases between above dates that were authorised since", ""
'    EnsureLanguageEntryExists 7739, "No. Of Cases", ""
'    EnsureLanguageEntryExists 7740, "M", ""
'    EnsureLanguageEntryExists 7800, "A", ""
'    EnsureLanguageEntryExists 7801, "H", ""
'    EnsureLanguageEntryExists 7802, "C", ""
'    EnsureLanguageEntryExists 7803, "Block", "Bloco"
'    EnsureLanguageEntryExists 7804, "Touch Prep", ""
'    EnsureLanguageEntryExists 7805, "Control", ""
'    EnsureLanguageEntryExists 7806, "M", "M"
'    EnsureLanguageEntryExists 7807, "F", "F"
'    EnsureLanguageEntryExists 7808, "Unique ID", "exclusivo"
'    EnsureLanguageEntryExists 7809, "Table", "tabela"
'    EnsureLanguageEntryExists 7810, "Path", "caminho"
'    EnsureLanguageEntryExists 7811, "Returned", ""
'    EnsureLanguageEntryExists 7812, "Type", ""
'    EnsureLanguageEntryExists 7813, "Referred To", ""
'    EnsureLanguageEntryExists 7814, "Reason For Referral", ""
'    EnsureLanguageEntryExists 7815, "TissueTypeId", ""
'    EnsureLanguageEntryExists 7816, "TissueTypeLetter", ""
'    EnsureLanguageEntryExists 7817, "TissuePath", ""
'    EnsureLanguageEntryExists 7818, "TissueTypeListId", ""
'    EnsureLanguageEntryExists 7819, "Add Tissue Type", ""
'    EnsureLanguageEntryExists 7820, "Add Cut-Up Details", ""
'    EnsureLanguageEntryExists 7821, "Open", ""
'    EnsureLanguageEntryExists 7822, "Dispose Case", ""
'    EnsureLanguageEntryExists 7823, "All Embedded", ""
'    EnsureLanguageEntryExists 7824, "Add Frozen Section", ""
'    EnsureLanguageEntryExists 7825, "Add Touch Prep", ""
'    EnsureLanguageEntryExists 7826, "Add Single Block", ""
'    EnsureLanguageEntryExists 7827, "Add Multiple Block", ""
'    EnsureLanguageEntryExists 7828, "Add Single Slide", ""
'    EnsureLanguageEntryExists 7829, "Add Multiple Slide", ""
'    EnsureLanguageEntryExists 7830, "Referral", ""
'    EnsureLanguageEntryExists 7831, "Add Number of Levels", ""
'    EnsureLanguageEntryExists 7832, "Add Routine Stain", ""
'    EnsureLanguageEntryExists 7833, "Add Special Stain", ""
'    EnsureLanguageEntryExists 7834, "Add Immunohistochemical Stain", ""
'    EnsureLanguageEntryExists 7835, "Add Extra Levels", ""
'    EnsureLanguageEntryExists 7836, "Print to Block Number", ""
'    EnsureLanguageEntryExists 7837, "Add Control", ""
'    EnsureLanguageEntryExists 7838, "P Codes", ""
'    EnsureLanguageEntryExists 7839, "M Codes", ""
'    EnsureLanguageEntryExists 7840, "Q Codes", ""
'    EnsureLanguageEntryExists 7841, "T Codes", ""
'    EnsureLanguageEntryExists 7842, "Stains", ""
'    EnsureLanguageEntryExists 7843, "Codes", ""
'    EnsureLanguageEntryExists 7844, "Special", ""
'    EnsureLanguageEntryExists 7845, "Immunohistochemical", ""
'    EnsureLanguageEntryExists 7846, "Coroners", ""
'    EnsureLanguageEntryExists 7847, "County", ""
'    EnsureLanguageEntryExists 7848, "Orientation", ""
'    EnsureLanguageEntryExists 7849, "Processor", ""
'    EnsureLanguageEntryExists 7850, "Discrepancy", ""
'    EnsureLanguageEntryExists 7851, "Accreditation Settings", ""
'    EnsureLanguageEntryExists 7852, "Non Work Days", ""
'    EnsureLanguageEntryExists 7853, "Work logs", ""
'    EnsureLanguageEntryExists 7854, "Disposals", ""
'    EnsureLanguageEntryExists 7855, "Autopsy", ""
'    EnsureLanguageEntryExists 7856, "By Tissue Type", ""
'    EnsureLanguageEntryExists 7857, "Diagnosis Specific Totals", ""
'    EnsureLanguageEntryExists 7858, "Diagnosis Range Search", ""
'    EnsureLanguageEntryExists 7859, "Grouped Tissue Search", ""
'    EnsureLanguageEntryExists 7860, "Location Specific Search", ""
'    EnsureLanguageEntryExists 7861, "Locked for Editing", ""
'    EnsureLanguageEntryExists 7862, "Numerical", ""
'    EnsureLanguageEntryExists 7863, "TAT", ""
'    EnsureLanguageEntryExists 7864, "AuthorisedReports", ""
'    EnsureLanguageEntryExists 7865, "Logged In", ""
'    EnsureLanguageEntryExists 7866, "Record Being Edited By You", ""
'    EnsureLanguageEntryExists 7867, "Record Being Edited By", ""
'    EnsureLanguageEntryExists 7868, "Please Enter Case Number", ""
'    EnsureLanguageEntryExists 7869, "Please Select A Node First", ""
'    EnsureLanguageEntryExists 7870, "Please Fill In Mandatory Fields", ""
'    EnsureLanguageEntryExists 7871, "Please Select A Pathologist", ""
'    EnsureLanguageEntryExists 7872, "Case ID Format Incorrect!", ""
'    EnsureLanguageEntryExists 7873, "No Demographics Available", ""
'    EnsureLanguageEntryExists 7874, "Record Saved", ""
'    EnsureLanguageEntryExists 7875, "Must Enter P code before authorisation", ""
'    EnsureLanguageEntryExists 7876, "Must Enter M code before authorisation", ""
'    EnsureLanguageEntryExists 7877, "Are you sure you want to delete?", ""
'    EnsureLanguageEntryExists 7878, "Case ID already exists", ""
'    EnsureLanguageEntryExists 7879, "Please Enter Number Of Levels", ""
'    EnsureLanguageEntryExists 7880, "Please Enter Number Of Blocks", ""
'    EnsureLanguageEntryExists 7881, "Please Enter Number Of Slides", ""
'    EnsureLanguageEntryExists 7882, "Checked By", ""
'    EnsureLanguageEntryExists 7883, "Cut-Up By", ""
'    EnsureLanguageEntryExists 7884, "Pieces At Embedding", ""
'    EnsureLanguageEntryExists 7885, "Assisted By", ""
'    EnsureLanguageEntryExists 7886, "Pieces At Cut-Up", ""
'    EnsureLanguageEntryExists 7887, "Embedded By", ""
'    EnsureLanguageEntryExists 7888, "Tissue ID", ""
'    EnsureLanguageEntryExists 7889, "Tissue Type", ""
'    EnsureLanguageEntryExists 7890, "Changes", ""
'    EnsureLanguageEntryExists 7891, "Events", ""
'    EnsureLanguageEntryExists 7892, "Logged In User", ""
    
     
    
    
    
    
    
    
    
End Sub

'Private Sub EnsureLanguageEntryExists(ByVal Reference As Integer, _
'                                      ByVal English As String, _
'                                      ByVal Portuguese As String)
'
'          Dim tb As Recordset
'          Dim sql As String
'
'10        On Error GoTo EnsureLanguageEntryExists_Error
'
'20        sql = "SELECT Reference, English, Portuguese FROM Language WHERE " & _
'                "Reference = '" & Reference & "'"
'30        Set tb = New Recordset
'40        RecOpenClient 0, tb, sql
'50        If tb.EOF Then
'60            tb.AddNew
'70        End If
'80        tb!Reference = Reference
'90        tb!English = English
'100       tb!Portuguese = Portuguese
'110       tb.Update
'
'120       Exit Sub

'EnsureLanguageEntryExists_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'130       intEL = Erl
'140       strES = Err.Description
'150       LogError "modLanguage", "EnsureLanguageEntryExists", intEL, strES, sql
'
'End Sub
'
Public Sub LoadLanguage(ByVal Language As String)

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo LoadLanguage_Error

20        Set colLanguage = New Collection

30        sql = "SELECT Reference, " & Language & " Lan FROM Language"
40        Set tb = New Recordset
          '50    RecOpenServer 0, tb, sql
50        Set tb = Cnxn(0).Execute(sql)
60        Do While Not tb.EOF
70            s = Trim$(Replace(tb!Lan & "", Chr$(9), ""))
80            colLanguage.Add s, CStr(tb!Reference)
90            tb.MoveNext
100       Loop

110       Exit Sub

LoadLanguage_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "modLanguage", "LoadLanguage", intEL, strES, sql

End Sub
Public Function LS(ResourceID As Integer) As String

10        On Error GoTo LS_Error

20        LS = colLanguage(ResourceID)

30        Exit Function

LS_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "modLanguage", "LS", intEL, strES

End Function



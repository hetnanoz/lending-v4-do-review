Option Explicit


Private Const CLASS_NAME As String = "modReportHelabaMain"

' Production must remain False. Set True only while debugging locally so VBA
' breaks on the original failing statement instead of entering BAT handlers.
Public Const DEV_MODE As Boolean = False

' Shared configuration constants used by the four Report Helaba modules.
' --- Deterministic error numbers (BAT standard, section 5) ---
Public Const ERR_NO_WORK As Long = 1001                    ' guard / nothing to do
Public Const ERR_FILE_NOT_FOUND As Long = 1010             ' file / lookup
Public Const ERR_TABLE_NOT_FOUND As Long = 1011            ' file / lookup
Public Const ERR_WORKBOOK_CLOSE_FAILED As Long = 1012       ' file / lookup
Public Const ERR_CONFIG_MISSING As Long = 1020             ' path / config
Public Const ERR_PATH_MISSING As Long = 1021               ' path / config
Public Const ERR_FOLDER_NOT_FOUND As Long = 1022            ' path / config

' --- Workbook configuration ---
Public Const CONFIG_SHEET_NAME As String = "Main"
Public Const SAVE_NAME_CELL As String = "save_path"

' --- Output workbook configuration ---
Public Const OUTPUT_FILE_PREFIX As String = "Lending_fees_loader_"
Public Const OUTPUT_FILE_EXT As String = ".xlsx"

' --- Microsoft Information Protection sensitivity label ---
Public Const MIP_INTERNAL_LABEL_ID As String = _
    "8ffbc0b8-e97b-47d1-beac-cb0955d66f3b"
Public Const MIP_SITE_ID As String = _
    "614f9c25-bffa-42c7-86d8-964101f55fa2"
Public Const MIP_ASSIGNMENT_METHOD As Long = 1
Public Const MIP_CONTENT_BITS As Long = 2

' --- Report sheet / PDF import configuration ---
Public Const PDF_SHEET_NAME As String = "Report_Helaba"    ' final (static) sheet
Public Const PDF_TMP_SHEET As String = "tmp_ReportImport"  ' transient PQ load sheet
Public Const PDF_TMP_QUERY As String = "tmp_ReportImport"  ' transient PQ query
Public Const PDF_TABLE_ID As String = "Table001"           ' PDF table id: Table001, Table002, Page001...
Public Const PDF_COMBINE_ALL As Boolean = False            ' True = combine all detected tables
Public Const PDF_PROMOTE_HEADERS As Boolean = True         ' True = promote first row to headers

' --- Excel report import configuration (alternative input to the PDF) ---
' The Excel report has a logo/title above the table, so the table is located by
' its first header caption instead of assuming a fixed start cell.
Public Const XL_HEADER_ANCHOR As String = "Fund Name"       ' first header caption
Public Const XL_MAX_SCAN_COLS As Long = 60                 ' safety cap when scanning header width

' --- Fund column configuration (added to Report_Helaba) ---
Public Const FUND_COL_HEADER As String = "Fund"
Public Const FUND_ACCOUNT_COL As Long = 5                  ' Report column E = Account
Public Const TABLE_STYLE As String = "TableStyleLight21"

' --- Fondsliste configuration ---
Public Const FOND_SHEET_NAME As String = "Fondsliste"
Public Const FOND_STATUS_TEXT As String = "Open"           ' value in column B used to filter rows
Public Const FOND_HAS_HEADER As Boolean = True             ' first row of the source file is a meta row
Public Const FOND_STATUS_SRC_COL As Long = 2               ' source column B
Public Const FOND_FUND_SRC_COL As Long = 4                 ' source column D
Public Const FOND_NAME_SRC_COL As Long = 5                 ' source column E
Public Const FOND_ACCOUNT_SRC_COL As Long = 14             ' source column N
Public Const FOND_TEAM_SRC_COL As Long = 15                ' source column O
Public Const FOND_LAST_COLUMN As Long = 15                 ' last column read (O)
Public Const FOND_OUT_COLS As Long = 5                      ' Status, Fund, Name, Account, Team
Public Const FOND_OUT_FUND_COL As Long = 2                 ' Fund in the filtered array
Public Const FOND_OUT_ACCOUNT_COL As Long = 4              ' Account in the filtered array
Public Const FOND_OUT_TEAM_COL As Long = 5                 ' Team in the filtered array
Public Const FOND_NAME_CELL As String = "root_Fondsliste"  ' named cell with the FOLDER path
Public Const FOND_FILE_PATTERN As String = "Fondsliste*.xls*"

' Headers written to the Fondsliste sheet. The first row of the raw file is a
' meta row ("... / Number of selected rows"), not real headers.
Public Const FOND_HDR_STATUS As String = "Status"
Public Const FOND_HDR_FUND As String = "Fund"
Public Const FOND_HDR_NAME As String = "Fund name"
Public Const FOND_HDR_ACCOUNT As String = "Account"
Public Const FOND_HDR_TEAM As String = "Team"

' --- Accrual enrichment configuration (user-picked file) ---
' Match: Report col A (Fund) = accrual file col A, and accrual file col C =
' ACC_BASE_NUMBER & <suffix from prompt> & <currency from Report col C>.
Public Const ACC_BASE_NUMBER As String = "419900"
Public Const ACC_HDR_CODE As String = "Acc_nr_and_suffix"
Public Const ACC_HDR_ACCRUAL As String = "Accrual"
Public Const ACC_HDR_TBB As String = "TBB_on_729000_0"
Public Const ACC_FILE_FUND_COL As Long = 1                 ' accrual file col A = fund
Public Const ACC_FILE_CODE_COL As Long = 3                 ' accrual file col C = account code
Public Const ACC_FILE_VALUE_COL As Long = 7                ' accrual file col G = accrual value
Public Const REPORT_FUND_COL As Long = 1                   ' Report_Helaba col A = Fund
Public Const REPORT_ACCOUNT_COL As Long = FUND_ACCOUNT_COL + 1 ' after Fund is prepended
Public Const REPORT_CURRENCY_COL As Long = 3               ' Report_Helaba col C = CURRENCY
Public Const REPORT_TBB_BASE_COL As Long = 9               ' Report_Helaba col I = Fund received

' --- Loader_input booking rules (TBB = TBB_on_729000_0, ACC = Accrual) ---
'   TBB blank                -> no line
'   TBB = 0 and ACC = 0      -> no line
'   TBB = 0 and ACC > 0      -> 1 line:  DB 510100/68        CR 419900/<suffix>
'   TBB = 0 and ACC < 0      -> 1 line:  DB 419900/<suffix>  CR 510100/68
'   TBB > 0                  -> 2 lines: |ACC| DB 510100/68  CR 419900/<suffix>
'                                        |TBB| DB 510100/68  CR 729000/0
'   TBB < 0                  -> 2 lines: |ACC| DB 510100/68  CR 419900/<suffix>
'                                        |TBB| DB 729000/0   CR 510100/68
'   Amounts are absolute, rounded UP at 2 decimals; a line whose amount rounds to
'   0,00 is never written (mirrors the "no zero booking" rule).
Public Const LOADER_SHEET_NAME As String = "Loader_input"
Public Const LOADER_COL_COUNT As Long = 9                   ' A..I
Public Const LOADER_AMOUNT_COL As Long = 3                 ' column C - BETRAG (text)
Public Const LOADER_DATE_COL As Long = 9                    ' column I - DATUM (text)
Public Const ACC_MAIN As String = "510100"                  ' our main account
Public Const ACC_MAIN_SUFFIX As String = "68"
Public Const ACC_TBB As String = "729000"                   ' TBB account
Public Const ACC_TBB_SUFFIX As String = "0"                 ' always 0 for the TBB account
Public Const LOADER_DESC_PREFIX As String = "Agency Lending Fees"
Public Const HDR_FONDS As String = "FONDS"
Public Const HDR_CURRENCY As String = "WAEHRUNG"
Public Const HDR_AMOUNT As String = "BETRAG"
Public Const HDR_PURPOSE As String = "VERWENDUNGSZWECK"
Public Const HDR_KONTO_DB As String = "KONTO (DB)"
Public Const HDR_SUFFIX_DB As String = "SUFFIX (DB)"
Public Const HDR_KONTO_CR As String = "KONTO (CR)"
Public Const HDR_SUFFIX_CR As String = "SUFFIX (CR)"
Public Const HDR_DATE As String = "DATUM"

' --- File dialog filters ---
Public Const FILTER_REPORT_NAME As String = "Report files (PDF or Excel)"
Public Const FILTER_REPORT As String = "*.pdf;*.xls;*.xlsx;*.xlsm"
Public Const FILTER_XLS_NAME As String = "Excel files"
Public Const FILTER_XLS As String = "*.xls*"

' --- User-facing texts (all English) ---
Public Const MSG_TITLE As String = "Report Helaba"
Public Const MSG_PICK_REPORT As String = "Select the source report file (PDF or Excel)"
Public Const MSG_PICK_ACCRUAL As String = "Select the accrual data file (columns A, C, G)"
Public Const MSG_ASK_SUFFIX As String = "Enter the account suffix (e.g. 5 or 12). The code will be built as " & ACC_BASE_NUMBER & "<suffix><CURRENCY>. Leave empty to skip the accrual matching."
Public Const MSG_ASK_YEAR As String = "Enter the booking year (e.g. 2026). Used in the Loader_input description."
Public Const TITLE_SUFFIX_PROMPT As String = "Accrual suffix"
Public Const TITLE_YEAR_PROMPT As String = "Booking year"
Public Const MSG_SKIP_ACCRUAL As String = "Accrual matching skipped (no file or no suffix). The output workbook will have no accrual columns and no Loader_input sheet."
Public Const MSG_DONE As String = "Done. Output files created:"
Public Const MSG_DONE_WORKBOOK As String = "Workbook:"
Public Const MSG_DONE_DAT As String = "DAT file:"
Public Const DLG_FOLDER_CAPTION As String = "Folder:"
Public Const DLG_OPEN_FOLDER_CAPTION As String = "Move to folder"
Public Const DLG_COPY_PATH_CAPTION As String = "Copy path"
Public Const DLG_COPIED_CAPTION As String = "Copied"
Public Const DLG_CLOSE_CAPTION As String = "Close"
Public Const DLG_ERR_FOLDER_MISSING As String = "The output folder is no longer available: "
Public Const ERR_TXT_NO_REPORT As String = "No report file selected."
Public Const ERR_TXT_NO_CONFIG As String = "The named cell 'root_Fondsliste' was not found or is empty. Create it on worksheet Main and enter the Fondsliste folder path."
Public Const ERR_TXT_NO_SAVE_CONFIG As String = "The named cell 'save_path' was not found or is empty. Create it on worksheet Main and enter the output folder path."
Public Const ERR_TXT_SAVE_FOLDER_NOT_FOUND As String = "The folder configured in 'save_path' does not exist or is not accessible."
Public Const ERR_TXT_NO_FONDFILE As String = "No Fondsliste file (Fondsliste*.xls*) found in the configured folder."
Public Const ERR_TXT_NO_TABLE As String = "The table header '" & XL_HEADER_ANCHOR & "' was not found in the selected Excel file."
Public Const ERR_TXT_TEAM_CANCELLED As String = "Team selection was cancelled. No output file was created."

' Defensive step guards remain in place for invalid or empty results. Handled
' child errors now propagate through the shared BAT ErrorManager.
Public Const ERR_TXT_STEP_FONDSLISTE As String = "The Fondsliste could not be read."
Public Const ERR_TXT_STEP_WRITE As String = "The output sheet could not be written. Sheet: "
Public Const ERR_TXT_STEP_FUND_LOOKUP As String = "The Account to Fund lookup could not be built."
Public Const ERR_TXT_STEP_TEAM_LOOKUP As String = "The Account to Team lookup could not be built."
Public Const ERR_TXT_STEP_IMPORT As String = "The source report could not be imported."
Public Const ERR_TXT_STEP_CLEAN As String = "The empty rows could not be removed from the report."
Public Const ERR_TXT_STEP_FUND_COLUMN As String = "The Fund column could not be added."
Public Const ERR_TXT_STEP_ACCRUAL_FILE As String = "The accrual file could not be read."
Public Const ERR_TXT_STEP_ACCRUAL As String = "The accrual columns could not be added."
Public Const ERR_TXT_STEP_TEAM_COLUMN As String = "The Team column could not be added."
Public Const ERR_TXT_STEP_LOADER As String = "The loader booking lines could not be built."
Public Const ERR_TXT_STEP_DAT As String = "The DAT file could not be created."
Public Const ERR_TXT_STEP_SAVE As String = "The output workbook could not be saved."
Public Const ERR_TXT_NO_FONDROWS As String = "No open funds could be extracted from the Fondsliste."
Public Const ERR_TXT_LOADER_LINE As String = "A loader booking line could not be built."

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    ---
' Returns:       ---
' Description:   Runs the complete Helaba report workflow and handles user-facing
'                messages while restoring the previous Excel application state.
'-------------------------------------------------------------------------------
Public Sub RunReportHelaba()
    Const METHOD_NAME As String = "RunReportHelaba"
    Dim arrFinal As Variant
    Dim arrFond As Variant
    Dim arrLoader As Variant
    Dim arrReport As Variant
    Dim blnAccrual As Boolean
    Dim blnAllFunds As Boolean
    Dim blnAllTeams As Boolean
    Dim blnPreviousEnableEvents As Boolean
    Dim blnPreviousScreenUpdating As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim calculationPrevious As Excel.XlCalculation
    Dim dictAcc As Object
    Dim dictFund As Object
    Dim dictSelectedFunds As Object
    Dim dictSelectedTeams As Object
    Dim dictTeam As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCloseError As Long
    Dim strAccPath As String
    Dim strCloseError As String
    Dim strDatPath As String
    Dim strFolder As String
    Dim strFondFile As String
    Dim strNote As String
    Dim strPlaceholder As String
    Dim strReportPath As String
    Dim strSaveFolder As String
    Dim strSavedAs As String
    Dim strSuffix As String
    Dim strYear As String
    Dim wkbOutput As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousEnableEvents = Application.EnableEvents
    blnPreviousScreenUpdating = Application.ScreenUpdating
    calculationPrevious = Application.Calculation
    blnSettingsCaptured = True

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    strReportPath = GetFilePath(MSG_PICK_REPORT, FILTER_REPORT_NAME, FILTER_REPORT)
    If Len(strReportPath) = 0 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_NO_REPORT
    End If

    strFolder = ResolveFondslisteFolder()
    If Len(strFolder) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, ERR_TXT_NO_CONFIG
    End If

    strFondFile = FindLatestFondsliste(strFolder)
    If Len(strFondFile) = 0 Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_FONDFILE
    End If

    strSaveFolder = ResolveSaveFolder()
    If Len(strSaveFolder) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, ERR_TXT_NO_SAVE_CONFIG
    End If

    Set wkbOutput = Application.Workbooks.Add(xlWBATWorksheet)
    strPlaceholder = wkbOutput.Worksheets(1).Name

    arrFond = ReadAndFilterFondsliste(strFondFile)
    If Not IsArray(arrFond) Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_STEP_FONDSLISTE
    End If

    If Not WriteArrayToSheet(wkbOutput, FOND_SHEET_NAME, arrFond, True, False, Array()) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_WRITE & FOND_SHEET_NAME
    End If

    Set dictFund = BuildFundDict(arrFond)
    If dictFund Is Nothing Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_FUND_LOOKUP
    End If

    Set dictTeam = BuildTeamDict(arrFond)
    If dictTeam Is Nothing Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_TEAM_LOOKUP
    End If

    If IsPdfFile(strReportPath) Then
        arrReport = ImportPdfToArray(wkbOutput, strReportPath)
    Else
        arrReport = ImportExcelReportToArray(strReportPath)
    End If
    If Not IsArray(arrReport) Then
        Err.Raise ERR_TABLE_NOT_FOUND, METHOD_NAME, ERR_TXT_STEP_IMPORT
    End If

    arrReport = RemoveEmptyFirstColumnRows(arrReport)
    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_CLEAN
    End If

    arrFinal = PrependFundColumn(arrReport, dictFund)
    If Not IsArray(arrFinal) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_FUND_COLUMN
    End If

    strAccPath = GetFilePath(MSG_PICK_ACCRUAL, FILTER_XLS_NAME, FILTER_XLS)
    If Len(strAccPath) > 0 Then
        strSuffix = Trim$(InputBox(MSG_ASK_SUFFIX, TITLE_SUFFIX_PROMPT))
        If Len(strSuffix) > 0 Then
            strYear = Trim$(InputBox(MSG_ASK_YEAR, TITLE_YEAR_PROMPT, VBA.Year(Date)))
            blnAccrual = True
        End If
    End If

    If blnAccrual Then
        Set dictAcc = BuildAccrualDict(strAccPath)
        If dictAcc Is Nothing Then
            Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_STEP_ACCRUAL_FILE
        End If

        arrFinal = AppendAccrualColumns(arrFinal, dictAcc, strSuffix)
        If Not IsArray(arrFinal) Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_ACCRUAL
        End If
    End If

    arrFinal = AppendTeamColumn(arrFinal, dictTeam)
    If Not IsArray(arrFinal) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_TEAM_COLUMN
    End If

    If Not WriteArrayToSheet(wkbOutput, PDF_SHEET_NAME, arrFinal, False, True, Array()) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_WRITE & PDF_SHEET_NAME
    End If

    If blnAccrual Then
        If Not SelectTeamsForLoader( _
                arrFinal, blnAllTeams, dictSelectedTeams, _
                blnAllFunds, dictSelectedFunds) Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_TEAM_CANCELLED
        End If

        arrLoader = BuildLoaderRows( _
            arrFinal, strSuffix, strYear, _
            blnAllTeams, dictSelectedTeams, _
            blnAllFunds, dictSelectedFunds)
        If Not IsArray(arrLoader) Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_LOADER
        End If

        If Not WriteArrayToSheet(wkbOutput, LOADER_SHEET_NAME, arrLoader, True, False, _
                                 Array(LOADER_AMOUNT_COL, LOADER_DATE_COL)) Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_STEP_WRITE & LOADER_SHEET_NAME
        End If

        Call modDatFile.CreateDatFileFromLoader( _
            arrLoader, _
            strSaveFolder, _
            strDatPath)

        If Len(strDatPath) = 0 Then
            Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_STEP_DAT
        End If
    End If

    Call RemoveSheet(wkbOutput, strPlaceholder)

    strSavedAs = SaveOutputWorkbook(wkbOutput, strSaveFolder)
    If Len(strSavedAs) = 0 Then
        Err.Raise ERR_PATH_MISSING, METHOD_NAME, ERR_TXT_STEP_SAVE
    End If

    Call CloseWorkbookSafely( _
        wkbOutput, _
        True, _
        lngCloseError, _
        strCloseError)

    If lngCloseError <> 0 Then
        Err.Raise ERR_WORKBOOK_CLOSE_FAILED, METHOD_NAME, _
                  "The output workbook was saved, but Excel could not close it: " & _
                  strCloseError
    End If

    If Not blnAccrual Then
        strNote = MSG_SKIP_ACCRUAL
    End If

    Call ShowCompletionDialog(strDatPath, strSaveFolder, strNote)

ExitPoint:
    If Not wkbOutput Is Nothing Then
        Call CloseWorkbookSafely(wkbOutput, True, lngCloseError, strCloseError)
    End If

    If blnSettingsCaptured Then
        Application.EnableEvents = blnPreviousEnableEvents
        Application.ScreenUpdating = blnPreviousScreenUpdating
        Application.Calculation = calculationPrevious
    End If
    If errNumber <> 0 Then
        Call ErrorManager.display
    End If

    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Sub


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strDatPath - full path of the DAT file, empty when not created
'                strFolder - folder that holds the output files
'                strNote - optional warning shown in the dialog
' Returns:       ---
' Description:   Shows the completion dialog with a copyable path and a button
'                that opens the output folder, then releases the form.
'-------------------------------------------------------------------------------
Private Sub ShowCompletionDialog(ByVal strDatPath As String, _
                                 ByVal strFolder As String, _
                                 ByVal strNote As String)
    Const METHOD_NAME As String = "ShowCompletionDialog"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    frmDoneDialog.ShowCompletion strDatPath, strFolder, strNote

ExitPoint:
    Unload frmDoneDialog

    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strDatPath;strFolder", strDatPath, strFolder)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    ---
' Returns:       String - folder path taken from the named configuration cell
' Description:   Reads the Fondsliste folder from the named cell root_Fondsliste
'                in THIS workbook. A missing or empty name is a configuration
'                error and is raised, not silently ignored.
'-------------------------------------------------------------------------------
Private Function ResolveFondslisteFolder() As String
    Const METHOD_NAME As String = "ResolveFondslisteFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFolder As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolder = ResolveNamedCellText(FOND_NAME_CELL)

    If Len(strFolder) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, ERR_TXT_NO_CONFIG
    End If

    ResolveFondslisteFolder = strFolder

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    ---
' Returns:       String - normalized writable output folder
' Description:   Reads save_path from the workbook or worksheet Main. HTTP URLs
'                are redirected to the current user's Desktop so XLSX and DAT
'                are always written to the same local folder.
'-------------------------------------------------------------------------------
Private Function ResolveSaveFolder() As String
    Const METHOD_NAME As String = "ResolveSaveFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFolder As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolder = ResolveNamedCellText(SAVE_NAME_CELL)

    If Len(strFolder) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, ERR_TXT_NO_SAVE_CONFIG
    End If

    strFolder = NormalizeConfiguredOutputFolder(strFolder)

    If Not FolderExists(strFolder) Then
        Err.Raise ERR_FOLDER_NOT_FOUND, METHOD_NAME, _
                  ERR_TXT_SAVE_FOLDER_NOT_FOUND & vbCrLf & strFolder
    End If

    ResolveSaveFolder = strFolder

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strDefinedName - workbook-level or Main-sheet-level name
' Returns:       String - trimmed text stored in the named cell
' Description:   Resolves configuration names independently of whether they are
'                scoped to the workbook or locally to worksheet Main.
'-------------------------------------------------------------------------------
Private Function ResolveNamedCellText(ByVal strDefinedName As String) As String
    Const METHOD_NAME As String = "ResolveNamedCellText"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objName As Excel.Name
    Dim rngCell As Excel.Range
    Dim strCandidateName As String
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each objName In ThisWorkbook.Names
        strCandidateName = Replace(objName.Name, "'", vbNullString)

        If InStrRev(strCandidateName, "!") > 0 Then
            strCandidateName = Mid$( _
                strCandidateName, _
                InStrRev(strCandidateName, "!") + 1)
        End If

        If StrComp(strCandidateName, strDefinedName, vbTextCompare) = 0 Then
            Set rngCell = objName.RefersToRange
            Exit For
        End If
    Next objName

    If rngCell Is Nothing Then
        ResolveNamedCellText = vbNullString
    Else
        strValue = Trim$(CStr(rngCell.Value2))

        If Len(strValue) >= 2 Then
            If Left$(strValue, 1) = Chr$(34) And _
               Right$(strValue, 1) = Chr$(34) Then
                strValue = Mid$(strValue, 2, Len(strValue) - 2)
            End If
        End If

        ResolveNamedCellText = Trim$(strValue)
    End If

ExitPoint:
    Set objName = Nothing
    Set rngCell = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strDefinedName", strDefinedName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFolderPath - configured output path
' Returns:       String - normalized local path ending with a separator
' Description:   Normalizes local, UNC and synced OneDrive paths. An HTTP URL is
'                redirected to Desktop because VBA text output cannot be written
'                directly to a SharePoint URL with Open ... For Output.
'-------------------------------------------------------------------------------
Private Function NormalizeConfiguredOutputFolder( _
        ByVal strFolderPath As String) As String
    Const METHOD_NAME As String = "NormalizeConfiguredOutputFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFolder As String
    Dim strUserProfile As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolder = Trim$(strFolderPath)

    If StrComp(Left$(strFolder, 4), "http", vbTextCompare) = 0 Then
        strUserProfile = Trim$(Environ$("USERPROFILE"))

        If Len(strUserProfile) = 0 Then
            Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, _
                      "USERPROFILE is unavailable, so Desktop cannot be resolved."
        End If

        strFolder = strUserProfile & Application.PathSeparator & "Desktop"
    End If

    strFolder = Replace(strFolder, "/", Application.PathSeparator)

    If Right$(strFolder, 1) <> Application.PathSeparator Then
        strFolder = strFolder & Application.PathSeparator
    End If

    NormalizeConfiguredOutputFolder = strFolder

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strFolderPath", strFolderPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFolderPath - normalized local or UNC folder
' Returns:       Boolean - True when the folder exists and is accessible
' Description:   Checks the configured destination without creating directories.
'-------------------------------------------------------------------------------
Private Function FolderExists(ByVal strFolderPath As String) As Boolean
    Const METHOD_NAME As String = "FolderExists"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object
    Dim strFolder As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolder = strFolderPath

    Do While Len(strFolder) > 3 And _
             Right$(strFolder, 1) = Application.PathSeparator
        strFolder = Left$(strFolder, Len(strFolder) - 1)
    Loop

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    FolderExists = objFileSystem.FolderExists(strFolder)

ExitPoint:
    Set objFileSystem = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strFolderPath", strFolderPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrReport - final report array containing accrual, TBB and Team
'                blnAllTeams - output flag; True means no team filtering
'                dictSelectedTeams - output dictionary of checked teams
'                blnAllFunds - output flag; True means no fund filtering
'                dictSelectedFunds - output dictionary of checked funds
' Returns:       Boolean - False only when the dialog was cancelled
' Description:   Shows teams with bookable lines and allows an optional fund
'                filter before Loader_input and the DAT file are created.
'-------------------------------------------------------------------------------
Private Function SelectTeamsForLoader(ByVal arrReport As Variant, _
                                      ByRef blnAllTeams As Boolean, _
                                      ByRef dictSelectedTeams As Object, _
                                      ByRef blnAllFunds As Boolean, _
                                      ByRef dictSelectedFunds As Object) As Boolean
    Const METHOD_NAME As String = "SelectTeamsForLoader"
    Dim arrTeamOptions As Variant
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrTeamOptions = BuildBookableTeamOptions(arrReport)

    frmTeamSelector.LoadTeams arrTeamOptions
    frmTeamSelector.Show vbModal

    If Not frmTeamSelector.WasCancelled Then
        blnAllTeams = frmTeamSelector.AllTeamsSelected
        Set dictSelectedTeams = frmTeamSelector.GetSelectedTeams()
        blnAllFunds = frmTeamSelector.AllFundsSelected
        Set dictSelectedFunds = frmTeamSelector.GetSelectedFunds()
        SelectTeamsForLoader = True
    End If

ExitPoint:
    Unload frmTeamSelector

    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "blnAllTeams;blnAllFunds", blnAllTeams, blnAllFunds)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strTitle - dialog caption
'                strFilterName - display name of the file filter
'                strFilter - filter mask, e.g. "*.pdf"
' Returns:       String - selected full path, or an empty string if cancelled
' Description:   Shows a single file picker. UI helper, called only from the entry
'                point so the business logic stays free of dialogs.
'-------------------------------------------------------------------------------
Private Function GetFilePath(ByVal strTitle As String, _
                             ByVal strFilterName As String, _
                             ByVal strFilter As String) As String
    Const METHOD_NAME As String = "GetFilePath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objDialog As Office.FileDialog

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set objDialog = Application.FileDialog(msoFileDialogFilePicker)
    With objDialog
        .Title = strTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add strFilterName, strFilter
        If .Show = -1 Then
            GetFilePath = .SelectedItems(1)
        Else
            GetFilePath = vbNullString
        End If
    End With

ExitPoint:
    Set objDialog = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise( _
            errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If

    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strTitle;strFilterName;strFilter", strTitle, strFilterName, strFilter)
    GoTo ExitPoint
End Function

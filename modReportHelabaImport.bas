Option Explicit


Private Const CLASS_NAME As String = "modReportHelabaImport"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFolder - folder that holds the Fondsliste files
' Returns:       String - full path of the newest matching file
' Description:   Returns the Fondsliste whose NAME carries the newest YYYYMMDD
'                stamp. One directory listing and no per-file timestamp calls
'                (see FileDateKey); an empty folder is raised as a lookup error.
'-------------------------------------------------------------------------------
Public Function FindLatestFondsliste(ByVal strFolder As String) As String
    Const METHOD_NAME As String = "FindLatestFondsliste"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strPath As String
    Dim strFile As String
    Dim strLatest As String
    Dim strKey As String
    Dim strLatestKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strPath = strFolder
    If Right$(strPath, 1) <> Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
    End If

    strFile = Dir$(strPath & FOND_FILE_PATTERN)
    Do While Len(strFile) > 0
        strKey = FileDateKey(strFile)
        If Len(strLatest) = 0 _
           Or strKey > strLatestKey _
           Or (strKey = strLatestKey And StrComp(strFile, strLatest, vbBinaryCompare) > 0) Then
            strLatest = strFile
            strLatestKey = strKey
        End If
        strFile = Dir$()
    Loop

    If Len(strLatest) = 0 Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_FONDFILE
    End If

    FindLatestFondsliste = strPath & strLatest

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
        "strFolder", strFolder)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFolder - folder that holds the Fondsliste file
'                strFileName - exact base name without extension
'                strFileType - exact extension, with or without leading dot
' Returns:       String - full path of the exact configured file
' Description:   Resolves only the exact configured Fondsliste name. No wildcard,
'                date suffix or alternative extension is accepted.
'-------------------------------------------------------------------------------
Public Function FindExactFondsliste( _
        ByVal strFolder As String, _
        ByVal strFileName As String, _
        ByVal strFileType As String) As String
    Const METHOD_NAME As String = "FindExactFondsliste"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strExpectedFile As String
    Dim strExtension As String
    Dim strFound As String
    Dim strPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strPath = Trim$(strFolder)
    If Right$(strPath, 1) <> Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
    End If

    strFileName = Trim$(strFileName)
    strExtension = Trim$(strFileType)

    If Left$(strExtension, 1) <> "." Then
        strExtension = "." & strExtension
    End If

    If InStr(strFileName, "*") > 0 Or InStr(strFileName, "?") > 0 _
       Or InStr(strExtension, "*") > 0 Or InStr(strExtension, "?") > 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, _
                  "Exact Fondsliste file name and extension cannot contain wildcards."
    End If

    If InStr(strFileName, "\") > 0 Or InStr(strFileName, "/") > 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, _
                  "file_name_Fondsliste must contain only the base file name."
    End If

    strExpectedFile = strFileName & strExtension
    strFound = Dir$( _
        strPath & strExpectedFile, _
        vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)

    If Len(strFound) = 0 _
       Or StrComp(strFound, strExpectedFile, vbTextCompare) <> 0 Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, _
                  ERR_TXT_NO_EXACT_FONDFILE & strExpectedFile
    End If

    FindExactFondsliste = strPath & strFound

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
        "strFolder;strFileName;strFileType", strFolder, strFileName, strFileType)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strName - a file name
' Returns:       String - the last 8 digit (YYYYMMDD) group in the name, or ""
' Description:   Extracts the date stamp from a file name so the newest Fondsliste
'                can be chosen without touching the file system. Equal length
'                numeric strings compare correctly as text.
'-------------------------------------------------------------------------------
Private Function FileDateKey(ByVal strName As String) As String
    Const METHOD_NAME As String = "FileDateKey"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngI As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngI = 1 To Len(strName) - 7
        If Mid$(strName, lngI, 8) Like "########" Then
            FileDateKey = Mid$(strName, lngI, 8)
        End If
    Next lngI

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
        "strName", strName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFullPath - full path of the workbook to resolve
' Returns:       Excel.Workbook - matching open workbook, or Nothing
' Description:   Reuses only the workbook opened from the exact configured path,
'                preventing a same-named workbook from another folder from being used.
'-------------------------------------------------------------------------------
Private Function GetOpenWorkbookByFullPath( _
        ByVal strFullPath As String) As Excel.Workbook
    Const METHOD_NAME As String = "GetOpenWorkbookByFullPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbFound As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each wkbFound In Application.Workbooks
        If StrComp(wkbFound.FullName, strFullPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullPath = wkbFound
            Exit For
        End If
    Next wkbFound

ExitPoint:
    Set wkbFound = Nothing

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
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFullPath - full path of the selected Fondsliste file
' Returns:       Variant - filtered array (Status, Fund, Name, Account, Team)
' Description:   Opens the Fondsliste, reads columns A to O in one block and
'                returns only Open rows while restoring changed Excel settings.
'-------------------------------------------------------------------------------
Public Function ReadAndFilterFondsliste(ByVal strFullPath As String) As Variant
    Const METHOD_NAME As String = "ReadAndFilterFondsliste"
    Dim arrSource As Variant
    Dim blnPreviousAskToUpdateLinks As Boolean
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim blnWasOpen As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCloseError As Long
    Dim lngLastRow As Long
    Dim lngPreviousAutomationSecurity As Long
    Dim strCloseError As String
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousDisplayAlerts = Application.DisplayAlerts
    blnPreviousAskToUpdateLinks = Application.AskToUpdateLinks
    lngPreviousAutomationSecurity = Application.AutomationSecurity
    blnSettingsCaptured = True

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set wkbSource = GetOpenWorkbookByFullPath(strFullPath)
    blnWasOpen = Not wkbSource Is Nothing

    If Not blnWasOpen Then
        Set wkbSource = Application.Workbooks.Open( _
            Filename:=strFullPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False)
    End If

    Set wksSource = wkbSource.Worksheets(1)

    lngLastRow = wksSource.Cells(wksSource.Rows.Count, FOND_STATUS_SRC_COL).End(xlUp).Row
    If lngLastRow < 1 Then lngLastRow = 1

    arrSource = wksSource.Range( _
        wksSource.Cells(1, 1), _
        wksSource.Cells(lngLastRow, FOND_LAST_COLUMN)).Value

    ReadAndFilterFondsliste = FilterFondslisteRows(arrSource)
    If Not IsArray(ReadAndFilterFondsliste) Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_FONDROWS
    End If

ExitPoint:
    Set wksSource = Nothing
    Call CloseWorkbookSafely(wkbSource, Not blnWasOpen, lngCloseError, strCloseError)

    If blnSettingsCaptured Then
        Application.DisplayAlerts = blnPreviousDisplayAlerts
        Application.AskToUpdateLinks = blnPreviousAskToUpdateLinks
        Application.AutomationSecurity = lngPreviousAutomationSecurity
    End If

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
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrSource - 1-based 2D array (rows x >=15 cols) from Range.Value
' Returns:       Variant - 1-based 2D array: Status, Fund, Name, Account, Team
' Description:   Pure filter: keeps rows whose column B equals FOND_STATUS_TEXT
'                (trimmed, case-insensitive) and writes our own headers, because
'                the file's first row is a meta row. No worksheet access.
'-------------------------------------------------------------------------------
Public Function FilterFondslisteRows(ByVal arrSource As Variant) As Variant
    Const METHOD_NAME As String = "FilterFondslisteRows"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrCols As Variant
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngMatches As Long
    Dim lngOut As Long
    Dim lngStartRow As Long
    Dim lngR As Long
    Dim lngK As Long
    Dim lngW As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrSource)
    If Not IsArray(arrSource) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Fondsliste source is not an array."
    End If
    Debug.Assert UBound(arrSource, 2) >= FOND_LAST_COLUMN

    arrCols = Array(FOND_STATUS_SRC_COL, FOND_FUND_SRC_COL, FOND_NAME_SRC_COL, _
                    FOND_ACCOUNT_SRC_COL, FOND_TEAM_SRC_COL)
    lngRows = UBound(arrSource, 1)

    If FOND_HAS_HEADER Then
        lngStartRow = 2
    Else
        lngStartRow = 1
    End If

    For lngR = lngStartRow To lngRows
        If StrComp(Trim$(CStr(arrSource(lngR, FOND_STATUS_SRC_COL))), _
                   FOND_STATUS_TEXT, vbTextCompare) = 0 Then
            lngMatches = lngMatches + 1
        End If
    Next lngR

    lngOut = lngMatches
    If FOND_HAS_HEADER Then lngOut = lngOut + 1
    If lngOut = 0 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "No open funds found in the Fondsliste."
    End If

    ReDim arrOut(1 To lngOut, 1 To FOND_OUT_COLS)

    If FOND_HAS_HEADER Then
        lngW = lngW + 1
        arrOut(lngW, 1) = FOND_HDR_STATUS
        arrOut(lngW, 2) = FOND_HDR_FUND
        arrOut(lngW, 3) = FOND_HDR_NAME
        arrOut(lngW, 4) = FOND_HDR_ACCOUNT
        arrOut(lngW, 5) = FOND_HDR_TEAM
    End If

    For lngR = lngStartRow To lngRows
        If StrComp(Trim$(CStr(arrSource(lngR, FOND_STATUS_SRC_COL))), _
                   FOND_STATUS_TEXT, vbTextCompare) = 0 Then
            lngW = lngW + 1
            For lngK = 0 To FOND_OUT_COLS - 1
                arrOut(lngW, lngK + 1) = arrSource(lngR, arrCols(lngK))
            Next lngK
        End If
    Next lngR

    FilterFondslisteRows = arrOut

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
' Parameters:    strPath - full file path
' Returns:       Boolean - True when the file carries a .pdf extension
' Description:   Decides which import route the source report takes.
'-------------------------------------------------------------------------------
Public Function IsPdfFile(ByVal strPath As String) As Boolean
    Const METHOD_NAME As String = "IsPdfFile"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    IsPdfFile = (StrComp(Right$(Trim$(strPath), 4), ".pdf", vbTextCompare) = 0)

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
        "strPath", strPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFullPath - full path of the Excel report to import
' Returns:       Variant - 2D array in the same shape as the PDF import
' Description:   Opens the Excel report read-only, reads its used range and
'                restores all Excel settings changed while opening the source.
'-------------------------------------------------------------------------------
Public Function ImportExcelReportToArray(ByVal strFullPath As String) As Variant
    Const METHOD_NAME As String = "ImportExcelReportToArray"
    Dim arrSource As Variant
    Dim blnPreviousAskToUpdateLinks As Boolean
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim blnWasOpen As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCloseError As Long
    Dim lngPreviousAutomationSecurity As Long
    Dim rngUsed As Excel.Range
    Dim strCloseError As String
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousDisplayAlerts = Application.DisplayAlerts
    blnPreviousAskToUpdateLinks = Application.AskToUpdateLinks
    lngPreviousAutomationSecurity = Application.AutomationSecurity
    blnSettingsCaptured = True

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set wkbSource = GetOpenWorkbookByFullPath(strFullPath)
    blnWasOpen = Not (wkbSource Is Nothing)

    If Not blnWasOpen Then
        Set wkbSource = Application.Workbooks.Open( _
            Filename:=strFullPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False)
    End If

    Set wksSource = wkbSource.Worksheets(1)
    Set rngUsed = wksSource.UsedRange
    arrSource = rngUsed.Value

    ImportExcelReportToArray = ExtractReportTable(arrSource)
    If Not IsArray(ImportExcelReportToArray) Then
        Err.Raise ERR_TABLE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_TABLE
    End If

ExitPoint:
    Set rngUsed = Nothing
    Set wksSource = Nothing
    Call CloseWorkbookSafely(wkbSource, Not blnWasOpen, lngCloseError, strCloseError)

    If blnSettingsCaptured Then
        Application.DisplayAlerts = blnPreviousDisplayAlerts
        Application.AskToUpdateLinks = blnPreviousAskToUpdateLinks
        Application.AutomationSecurity = lngPreviousAutomationSecurity
    End If

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
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrSource - 1-based 2D array of the sheet's used range
' Returns:       Variant - 2D array starting at the table header row
' Description:   Pure transform: locates XL_HEADER_ANCHOR and keeps everything
'                from that row and column to the last non-empty header, dropping
'                the logo and title rows above the table.
'-------------------------------------------------------------------------------
Public Function ExtractReportTable(ByVal arrSource As Variant) As Variant
    Const METHOD_NAME As String = "ExtractReportTable"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngHdrRow As Long
    Dim lngHdrCol As Long
    Dim lngLastCol As Long
    Dim lngScanTo As Long
    Dim lngR As Long
    Dim lngC As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrSource)
    If Not IsArray(arrSource) Then
        Err.Raise ERR_TABLE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_TABLE
    End If

    lngRows = UBound(arrSource, 1)
    lngCols = UBound(arrSource, 2)

    For lngR = 1 To lngRows
        For lngC = 1 To lngCols
            If StrComp(Trim$(CStr(arrSource(lngR, lngC))), XL_HEADER_ANCHOR, vbTextCompare) = 0 Then
                lngHdrRow = lngR
                lngHdrCol = lngC
                Exit For
            End If
        Next lngC
        If lngHdrRow > 0 Then Exit For
    Next lngR

    If lngHdrRow = 0 Then
        Err.Raise ERR_TABLE_NOT_FOUND, METHOD_NAME, ERR_TXT_NO_TABLE
    End If

    lngScanTo = lngHdrCol + XL_MAX_SCAN_COLS - 1
    If lngScanTo > lngCols Then lngScanTo = lngCols
    lngLastCol = lngHdrCol
    For lngC = lngHdrCol To lngScanTo
        If Len(Trim$(CStr(arrSource(lngHdrRow, lngC)))) > 0 Then lngLastCol = lngC
    Next lngC
    Debug.Assert lngLastCol >= lngHdrCol

    ReDim arrOut(1 To lngRows - lngHdrRow + 1, 1 To lngLastCol - lngHdrCol + 1)
    For lngR = lngHdrRow To lngRows
        For lngC = lngHdrCol To lngLastCol
            arrOut(lngR - lngHdrRow + 1, lngC - lngHdrCol + 1) = arrSource(lngR, lngC)
        Next lngC
    Next lngR

    ExtractReportTable = arrOut

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
' Parameters:    wkbTarget - workbook that hosts the transient query and sheet
'                strPdfPath - full path of the PDF to import
' Returns:       Variant - 2D array of the imported PDF table
' Description:   Loads the PDF through Power Query and restores the previous
'                ScreenUpdating value after the temporary import completes.
'-------------------------------------------------------------------------------
Public Function ImportPdfToArray(ByVal wkbTarget As Excel.Workbook, _
                                  ByVal strPdfPath As String) As Variant
    Const METHOD_NAME As String = "ImportPdfToArray"
    Dim blnPreviousScreenUpdating As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim objListObject As Excel.ListObject
    Dim objQueryTable As Excel.QueryTable
    Dim strConnection As String
    Dim strFormula As String
    Dim wksTmp As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousScreenUpdating = Application.ScreenUpdating
    blnSettingsCaptured = True
    Application.ScreenUpdating = True

    Call RemoveSheet(wkbTarget, PDF_TMP_SHEET)
    Call RemoveQuery(wkbTarget, PDF_TMP_QUERY)

    strFormula = BuildPdfQueryFormula(strPdfPath)
    wkbTarget.Queries.Add Name:=PDF_TMP_QUERY, Formula:=strFormula

    Set wksTmp = wkbTarget.Worksheets.Add(After:=wkbTarget.Sheets(wkbTarget.Sheets.Count))
    wksTmp.Name = PDF_TMP_SHEET

    strConnection = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
                    "Location=" & PDF_TMP_QUERY & ";Extended Properties="""""

    Set objListObject = wksTmp.ListObjects.Add( _
        SourceType:=0, _
        Source:=strConnection, _
        Destination:=wksTmp.Range("$A$1"))

    Set objQueryTable = objListObject.QueryTable

    With objQueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & PDF_TMP_QUERY & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With

    ImportPdfToArray = objListObject.Range.Value

    Set objQueryTable = Nothing
    Set objListObject = Nothing
    Set wksTmp = Nothing
    Call RemoveSheet(wkbTarget, PDF_TMP_SHEET)
    Call RemoveQuery(wkbTarget, PDF_TMP_QUERY)

ExitPoint:
    Set objQueryTable = Nothing
    Set objListObject = Nothing
    Set wksTmp = Nothing

    If blnSettingsCaptured Then
        Application.ScreenUpdating = blnPreviousScreenUpdating
    End If

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
        "strPdfPath", strPdfPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strPdfPath - full path of the PDF file
' Returns:       String - a Power Query (M) let-expression for the PDF import
' Description:   Builds the M code for a faithful PDF import, driven by the module
'                constants. No row filtering here: empty first column rows are
'                removed afterwards in VBA by RemoveEmptyFirstColumnRows.
'-------------------------------------------------------------------------------
Private Function BuildPdfQueryFormula(ByVal strPdfPath As String) As String
    Const METHOD_NAME As String = "BuildPdfQueryFormula"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strPath As String
    Dim strFormula As String
    Dim strLastStep As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert Len(strPdfPath) > 0
    If Len(strPdfPath) = 0 Then
        Err.Raise ERR_FILE_NOT_FOUND, METHOD_NAME, "PDF path is empty."
    End If

    strPath = Replace(strPdfPath, """", """""")

    strFormula = "let" & vbCrLf
    strFormula = strFormula & _
        "    Source = Pdf.Tables(File.Contents(""" & strPath & """), [Implementation=""1.3""])," & vbCrLf

    If PDF_COMBINE_ALL Then
        strFormula = strFormula & _
            "    OnlyTables = Table.SelectRows(Source, each [Kind] = ""Table"")," & vbCrLf
        strFormula = strFormula & "    Data = Table.Combine(OnlyTables[Data])"
    Else
        strFormula = strFormula & "    Data = Source{[Id=""" & PDF_TABLE_ID & """]}[Data]"
    End If
    strLastStep = "Data"

    If PDF_PROMOTE_HEADERS Then
        strFormula = strFormula & "," & vbCrLf
        strFormula = strFormula & _
            "    Promoted = Table.PromoteHeaders(Data, [PromoteAllScalars=true])"
        strLastStep = "Promoted"
    End If

    strFormula = strFormula & vbCrLf & "in" & vbCrLf & "    " & strLastStep

    BuildPdfQueryFormula = strFormula

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
        "strPdfPath", strPdfPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFullPath - full path of the user-picked accrual file
' Returns:       Object - dictionary keyed by FUND|ACCOUNTCODE
' Description:   Reads the accrual source in one block and restores all Excel
'                settings changed while opening the source workbook.
'-------------------------------------------------------------------------------
Public Function BuildAccrualDict(ByVal strFullPath As String) As Object
    Const METHOD_NAME As String = "BuildAccrualDict"
    Dim arrSource As Variant
    Dim blnPreviousAskToUpdateLinks As Boolean
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim blnWasOpen As Boolean
    Dim dictAcc As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCloseError As Long
    Dim lngLastRow As Long
    Dim lngPreviousAutomationSecurity As Long
    Dim lngR As Long
    Dim strCloseError As String
    Dim strKey As String
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousDisplayAlerts = Application.DisplayAlerts
    blnPreviousAskToUpdateLinks = Application.AskToUpdateLinks
    lngPreviousAutomationSecurity = Application.AutomationSecurity
    blnSettingsCaptured = True

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set dictAcc = CreateObject("Scripting.Dictionary")

    Set wkbSource = GetOpenWorkbookByFullPath(strFullPath)
    blnWasOpen = Not (wkbSource Is Nothing)

    If Not blnWasOpen Then
        Set wkbSource = Application.Workbooks.Open( _
            Filename:=strFullPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False)
    End If

    Set wksSource = wkbSource.Worksheets(1)
    lngLastRow = wksSource.Cells(wksSource.Rows.Count, ACC_FILE_FUND_COL).End(xlUp).Row
    If lngLastRow < 1 Then lngLastRow = 1

    arrSource = wksSource.Range( _
        wksSource.Cells(1, 1), _
        wksSource.Cells(lngLastRow, ACC_FILE_VALUE_COL)).Value

    If IsArray(arrSource) Then
        For lngR = 1 To UBound(arrSource, 1)
            strKey = UCase$(Trim$(CStr(arrSource(lngR, ACC_FILE_FUND_COL)))) & "|" & _
                     UCase$(Replace(Trim$(CStr(arrSource(lngR, ACC_FILE_CODE_COL))), " ", ""))

            If Len(strKey) > 1 And Not dictAcc.Exists(strKey) Then
                dictAcc.Add strKey, Array( _
                    arrSource(lngR, ACC_FILE_CODE_COL), _
                    arrSource(lngR, ACC_FILE_VALUE_COL))
            End If
        Next lngR
    End If

    Set BuildAccrualDict = dictAcc

ExitPoint:
    Set wksSource = Nothing
    Call CloseWorkbookSafely(wkbSource, Not blnWasOpen, lngCloseError, strCloseError)

    If blnSettingsCaptured Then
        Application.DisplayAlerts = blnPreviousDisplayAlerts
        Application.AskToUpdateLinks = blnPreviousAskToUpdateLinks
        Application.AutomationSecurity = lngPreviousAutomationSecurity
    End If

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
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Function

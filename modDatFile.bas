Option Explicit

Private Const CLASS_NAME As String = "modDatFile"

Private Const ERR_NO_WORK As Long = 1001
Private Const ERR_FILE_EXISTS As Long = 1010
Private Const ERR_FILE_CLOSE As Long = 1011
Private Const ERR_PATH_INVALID As Long = 1020

Private Const DAT_LINE_LENGTH As Long = 1846
Private Const DAT_MAX_COUNTER As Long = 9999
Private Const DAT_FILE_PREFIX As String = "CO_Transact_"
Private Const DAT_TRANSACTION_TYPE As String = "CO02"
Private Const DAT_FILE_EXTENSION As String = ".dat"
Private Const LOADER_COLUMN_COUNT As Long = 9

Private Const COL_FONDS As Long = 1
Private Const COL_CURRENCY As Long = 2
Private Const COL_AMOUNT As Long = 3
Private Const COL_PURPOSE As Long = 4
Private Const COL_ACCOUNT_DB As Long = 5
Private Const COL_SUFFIX_DB As Long = 6
Private Const COL_ACCOUNT_CR As Long = 7
Private Const COL_SUFFIX_CR As Long = 8
Private Const COL_DATE As Long = 9

Private Const HDR_FONDS As String = "FONDS"
Private Const HDR_CURRENCY As String = "WAEHRUNG"
Private Const HDR_AMOUNT As String = "BETRAG"
Private Const HDR_PURPOSE As String = "VERWENDUNGSZWECK"
Private Const HDR_ACCOUNT_DB As String = "KONTO (DB)"
Private Const HDR_SUFFIX_DB As String = "SUFFIX (DB)"
Private Const HDR_ACCOUNT_CR As String = "KONTO (CR)"
Private Const HDR_SUFFIX_CR As String = "SUFFIX (CR)"
Private Const HDR_DATE As String = "DATUM"


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrLoader - one-based two-dimensional loader array
'                strFolderPath - target output folder or SharePoint/OneDrive URL
'                strCreatedFilePath - optional full path returned to the caller
' Returns:       ---
' Description:   Creates a fixed-width DAT file with 1846 characters per record.
'-------------------------------------------------------------------------------
Public Sub CreateDatFileFromLoader(ByVal arrLoader As Variant, _
                                   ByVal strFolderPath As String, _
                                   Optional ByRef strCreatedFilePath As String = "")
    Const METHOD_NAME As String = "CreateDatFileFromLoader"
    Dim arrExpectedHeaders As Variant
    Dim blnFileOpen As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCloseError As Long
    Dim lngColumn As Long
    Dim lngColumnCount As Long
    Dim lngCounter As Long
    Dim lngFileNumber As Long
    Dim lngRow As Long
    Dim lngRowCount As Long
    Dim strActualHeader As String
    Dim strCloseError As String
    Dim strExpectedHeader As String
    Dim strFileName As String
    Dim strFullPath As String
    Dim strLine As String
    Dim strOutputFolder As String
    Dim strSafeUserName As String
    Dim strStamp As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strCreatedFilePath = vbNullString

    If Not IsArray(arrLoader) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "The loader input is not an array."
    End If

    If LBound(arrLoader, 1) <> 1 Or LBound(arrLoader, 2) <> 1 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The loader array must be one-based in both dimensions."
    End If

    lngRowCount = UBound(arrLoader, 1)
    lngColumnCount = UBound(arrLoader, 2)

    If lngRowCount < 2 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "The loader array contains no data rows."
    End If

    If lngColumnCount < LOADER_COLUMN_COUNT Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The loader array must contain at least nine columns."
    End If

    If lngRowCount - 1 > DAT_MAX_COUNTER Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The loader array contains more than 9999 data rows."
    End If

    arrExpectedHeaders = Array( _
        HDR_FONDS, _
        HDR_CURRENCY, _
        HDR_AMOUNT, _
        HDR_PURPOSE, _
        HDR_ACCOUNT_DB, _
        HDR_SUFFIX_DB, _
        HDR_ACCOUNT_CR, _
        HDR_SUFFIX_CR, _
        HDR_DATE)

    For lngColumn = 1 To LOADER_COLUMN_COUNT
        strActualHeader = Trim$(ValueToRecordText(arrLoader(1, lngColumn)))
        strExpectedHeader = CStr(arrExpectedHeaders(lngColumn - 1))

        If StrComp(strActualHeader, strExpectedHeader, vbTextCompare) <> 0 Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, _
                      "Unexpected header in column " & CStr(lngColumn) & _
                      ". Expected '" & strExpectedHeader & "', received '" & _
                      strActualHeader & "'."
        End If
    Next lngColumn

    strOutputFolder = NormalizeOutputFolder(strFolderPath)
    strSafeUserName = SanitizeFileNamePart(Application.UserName)
    strStamp = VBA.Format$(Now, "ddmmyyhhnn")
    strFileName = DAT_FILE_PREFIX & strSafeUserName & "_" & strStamp & DAT_FILE_EXTENSION
    strFullPath = strOutputFolder & strFileName

    If Len(Dir$(strFullPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then
        Err.Raise ERR_FILE_EXISTS, METHOD_NAME, _
                  "The output file already exists: " & strFullPath
    End If

    lngFileNumber = FreeFile
    Open strFullPath For Output Access Write As #lngFileNumber
    blnFileOpen = True

    For lngRow = 2 To lngRowCount
        lngCounter = lngCounter + 1
        strLine = BuildDatLine(arrLoader, lngRow, strStamp, lngCounter)

        If Len(strLine) <> DAT_LINE_LENGTH Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, _
                      "Invalid DAT record length in loader row " & CStr(lngRow) & _
                      ". Expected 1846 characters, received " & CStr(Len(strLine)) & "."
        End If

        Print #lngFileNumber, strLine
    Next lngRow

    If Not CloseDatFileSafely( _
            lngFileNumber, _
            blnFileOpen, _
            lngCloseError, _
            strCloseError) Then

        Err.Raise ERR_FILE_CLOSE, METHOD_NAME, _
                  "The DAT file was written but could not be closed: " & strCloseError
    End If

    strCreatedFilePath = strFullPath

ExitPoint:
    If blnFileOpen Then
        Call CloseDatFileSafely( _
            lngFileNumber, _
            blnFileOpen, _
            lngCloseError, _
            strCloseError)
    End If

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
        "strFolderPath;strCreatedFilePath", strFolderPath, strCreatedFilePath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    lngFileNumber - VBA file handle
'                blnFileOpen - current open-state flag, reset after one attempt
'                lngCloseError - returned cleanup error number
'                strCloseError - returned cleanup error description
' Returns:       Boolean - True when no Close error occurred
' Description:   Closes the DAT text file through an isolated handler so a cleanup
'                failure cannot replace an earlier business error.
'-------------------------------------------------------------------------------
Private Function CloseDatFileSafely(ByVal lngFileNumber As Long, _
                                    ByRef blnFileOpen As Boolean, _
                                    ByRef lngCloseError As Long, _
                                    ByRef strCloseError As String) As Boolean
    Const METHOD_NAME As String = "CloseDatFileSafely"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngCloseError = 0
    strCloseError = vbNullString
    CloseDatFileSafely = True

    If blnFileOpen Then
        Close #lngFileNumber
        blnFileOpen = False
    End If

ExitPoint:
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    lngCloseError = errNumber
    strCloseError = errDescription
    blnFileOpen = False
    CloseDatFileSafely = False

    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "lngFileNumber", lngFileNumber)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrLoader - validated loader array
'                lngRow - current data row
'                strStamp - run timestamp in ddmmyyhhmm format
'                lngCounter - one-based record counter
' Returns:       String - one fixed-width record
' Description:   Maps one loader row to the required 1846-character DAT layout.
'-------------------------------------------------------------------------------
Private Function BuildDatLine(ByVal arrLoader As Variant, _
                              ByVal lngRow As Long, _
                              ByVal strStamp As String, _
                              ByVal lngCounter As Long) As String
    Const METHOD_NAME As String = "BuildDatLine"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dblAccountDb As Double
    Dim strAccountCr As String
    Dim strAccountDb As String
    Dim strAmount As String
    Dim strAmountField As String
    Dim strCsens As String
    Dim strCurrency As String
    Dim strDate As String
    Dim strDateField As String
    Dim strDescriptionField As String
    Dim strFund As String
    Dim strLine As String
    Dim strNecritureField As String
    Dim strNecriturePrefix As String
    Dim strPurpose As String
    Dim strSuffixCr As String
    Dim strSuffixDb As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFund = Trim$(ValueToRecordText(arrLoader(lngRow, COL_FONDS)))
    strCurrency = UCase$(Trim$(ValueToRecordText(arrLoader(lngRow, COL_CURRENCY))))
    strAmount = Trim$(ValueToRecordText(arrLoader(lngRow, COL_AMOUNT)))
    strPurpose = Trim$(ValueToRecordText(arrLoader(lngRow, COL_PURPOSE)))
    strAccountDb = Trim$(ValueToRecordText(arrLoader(lngRow, COL_ACCOUNT_DB)))
    strSuffixDb = Trim$(ValueToRecordText(arrLoader(lngRow, COL_SUFFIX_DB)))
    strAccountCr = Trim$(ValueToRecordText(arrLoader(lngRow, COL_ACCOUNT_CR)))
    strSuffixCr = Trim$(ValueToRecordText(arrLoader(lngRow, COL_SUFFIX_CR)))
    strDate = Trim$(ValueToRecordText(arrLoader(lngRow, COL_DATE)))

    If Len(strAccountDb) > 0 Then
        If Not IsNumeric(strAccountDb) Then
            Err.Raise ERR_NO_WORK, METHOD_NAME, _
                      "The debit account is not numeric in loader row " & _
                      CStr(lngRow) & ": " & strAccountDb
        End If
        dblAccountDb = CDbl(strAccountDb)
    End If

    If StrComp(strCurrency, "EUR", vbTextCompare) <> 0 And _
       dblAccountDb >= 600000# Then
        strCsens = "DB"
    Else
        strCsens = "CR"
    End If

    strSuffixCr = FormatDatSuffix(strSuffixCr)
    strSuffixDb = FormatDatSuffix(strSuffixDb)

    strDateField = Replace(strDate, "/", vbNullString)
    If Len(strDateField) <> 8 Or strDateField Like "*[!0-9]*" Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The date must use DD/MM/YYYY format in loader row " & _
                  CStr(lngRow) & ": " & strDate
    End If

    strAmountField = BuildAmountField(strAmount)

    If lngCounter < 1 Or lngCounter > DAT_MAX_COUNTER Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The DAT record counter must be between 1 and 9999."
    End If

    If StrComp(Left$(strFund, 2), "FA", vbTextCompare) = 0 Then
        strNecriturePrefix = strFund
    Else
        strNecriturePrefix = "F0"
    End If

    strNecritureField = strNecriturePrefix & strStamp & _
                        VBA.Format$(lngCounter, "0000")
    If Len(strNecritureField) > 20 Then
        strNecritureField = Right$(strNecritureField, 20)
    End If
    strNecritureField = PadRight(strNecritureField, 20)

    If StrComp(Left$(strFund, 2), "FA", vbTextCompare) = 0 Then
        strDescriptionField = strFund & "-" & strPurpose
    Else
        strDescriptionField = strPurpose
    End If
    strDescriptionField = Left$(strDescriptionField, 75) & " LOAD"
    strDescriptionField = PadRight(strDescriptionField, 80)

    strLine = PadRight(DAT_TRANSACTION_TYPE, 4)
    strLine = strLine & Space$(2)
    strLine = strLine & PadRight(strCsens, 2)
    strLine = strLine & Space$(8)
    strLine = strLine & PadRight(strCurrency, 3)
    strLine = strLine & Space$(31)
    strLine = strLine & PadRight(strSuffixCr, 10)
    strLine = strLine & PadRight(strSuffixDb, 10)
    strLine = strLine & Space$(32)
    strLine = strLine & PadRight(strAccountCr, 10)
    strLine = strLine & PadRight(strAccountDb, 10)
    strLine = strLine & Space$(44)
    strLine = strLine & PadRight(strDateField, 8)
    strLine = strLine & PadRight(strDateField, 8)
    strLine = strLine & Space$(169)
    strLine = strLine & strAmountField
    strLine = strLine & Space$(119)
    strLine = strLine & strNecritureField
    strLine = strLine & Space$(20)
    strLine = strLine & PadRight(strFund, 20)
    strLine = strLine & strDescriptionField
    strLine = strLine & Space$(1217)

    If Len(strLine) <> DAT_LINE_LENGTH Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The generated DAT record does not contain 1846 characters."
    End If

    BuildDatLine = strLine

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
        "lngRow;strStamp;lngCounter", lngRow, strStamp, lngCounter)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strSuffix - raw suffix value from the loader array
' Returns:       String - suffix containing at least two digits when numeric
' Description:   Formats numeric suffixes below 10 with a leading zero before
'                the fixed-width field is padded to 10 characters.
'-------------------------------------------------------------------------------
Private Function FormatDatSuffix(ByVal strSuffix As String) As String
    Const METHOD_NAME As String = "FormatDatSuffix"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dblSuffix As Double
    Dim strFormattedSuffix As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFormattedSuffix = Trim$(strSuffix)

    If Len(strFormattedSuffix) > 0 And _
       IsNumeric(strFormattedSuffix) Then

        dblSuffix = CDbl(strFormattedSuffix)

        If dblSuffix >= 0# And dblSuffix < 10# Then
            strFormattedSuffix = VBA.Format$(dblSuffix, "00")
        End If
    End If

    FormatDatSuffix = strFormattedSuffix

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
        "strSuffix", strSuffix)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFolderPath - requested target folder or an HTTP URL
' Returns:       String - local folder path ending with Application.PathSeparator
' Description:   Redirects HTTP paths to the desktop and normalizes separators.
'-------------------------------------------------------------------------------
Private Function NormalizeOutputFolder(ByVal strFolderPath As String) As String
    Const METHOD_NAME As String = "NormalizeOutputFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFolder As String
    Dim strUserProfile As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolder = Trim$(strFolderPath)

    If Len(strFolder) = 0 Then
        Err.Raise ERR_PATH_INVALID, METHOD_NAME, _
                  "The output folder path is empty."
    End If

    If Left$(strFolder, 1) = Chr$(34) And _
       Right$(strFolder, 1) = Chr$(34) And _
       Len(strFolder) >= 2 Then
        strFolder = Mid$(strFolder, 2, Len(strFolder) - 2)
    End If

    If StrComp(Left$(strFolder, 4), "http", vbTextCompare) = 0 Then
        strUserProfile = Trim$(Environ$("USERPROFILE"))
        If Len(strUserProfile) = 0 Then
            Err.Raise ERR_PATH_INVALID, METHOD_NAME, _
                      "USERPROFILE is unavailable, so the desktop fallback cannot be resolved."
        End If
        strFolder = strUserProfile & Application.PathSeparator & "Desktop"
    End If

    strFolder = Replace(strFolder, "/", Application.PathSeparator)
    If Right$(strFolder, 1) <> Application.PathSeparator Then
        strFolder = strFolder & Application.PathSeparator
    End If

    NormalizeOutputFolder = strFolder

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
' Parameters:    strValue - raw user name or another file-name component
' Returns:       String - Windows-safe file-name component
' Description:   Replaces forbidden and control characters with underscores.
'-------------------------------------------------------------------------------
Private Function SanitizeFileNamePart(ByVal strValue As String) As String
    Const METHOD_NAME As String = "SanitizeFileNamePart"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCharacter As Long
    Dim lngCharacterCode As Long
    Dim strCharacter As String
    Dim strClean As String
    Dim strInvalidCharacters As String
    Dim strSource As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strInvalidCharacters = "\/:*?""<>|"
    strSource = Trim$(strValue)

    For lngCharacter = 1 To Len(strSource)
        strCharacter = Mid$(strSource, lngCharacter, 1)
        lngCharacterCode = AscW(strCharacter)
        If lngCharacterCode < 0 Then lngCharacterCode = lngCharacterCode + 65536

        If lngCharacterCode < 32 Or _
           strCharacter = " " Or _
           InStr(1, strInvalidCharacters, strCharacter, vbBinaryCompare) > 0 Then
            strClean = strClean & "_"
        Else
            strClean = strClean & strCharacter
        End If
    Next lngCharacter

    Do While Len(strClean) > 0 And _
             (Right$(strClean, 1) = " " Or Right$(strClean, 1) = ".")
        strClean = Left$(strClean, Len(strClean) - 1)
    Loop

    If Len(strClean) = 0 Then strClean = "User"
    If Len(strClean) > 100 Then strClean = Left$(strClean, 100)

    SanitizeFileNamePart = strClean

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
        "strValue", strValue)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strAmount - amount in German text format
' Returns:       Double - parsed numeric amount
' Description:   Removes German thousands separators and parses a decimal point.
'-------------------------------------------------------------------------------
Private Function ParseGermanAmount(ByVal strAmount As String) As Double
    Const METHOD_NAME As String = "ParseGermanAmount"
    Dim errDescription As String
    Dim errNumber As Long
    Dim blnDigitFound As Boolean
    Dim lngCharacter As Long
    Dim lngDecimalPoints As Long
    Dim strCharacter As String
    Dim strNormalized As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strNormalized = Trim$(strAmount)
    strNormalized = Replace(strNormalized, ChrW(160), vbNullString)
    strNormalized = Replace(strNormalized, " ", vbNullString)
    strNormalized = Replace(strNormalized, vbTab, vbNullString)
    strNormalized = Replace(strNormalized, ".", vbNullString)
    strNormalized = Replace(strNormalized, ",", ".")

    If Len(strNormalized) = 0 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "The amount is empty."
    End If

    For lngCharacter = 1 To Len(strNormalized)
        strCharacter = Mid$(strNormalized, lngCharacter, 1)

        If strCharacter >= "0" And strCharacter <= "9" Then
            blnDigitFound = True
        ElseIf strCharacter = "." Then
            lngDecimalPoints = lngDecimalPoints + 1
            If lngDecimalPoints > 1 Then
                Err.Raise ERR_NO_WORK, METHOD_NAME, _
                          "The amount contains more than one decimal separator: " & _
                          strAmount
            End If
        ElseIf (strCharacter = "+" Or strCharacter = "-") And _
               lngCharacter = 1 Then
        Else
            Err.Raise ERR_NO_WORK, METHOD_NAME, _
                      "The amount has an invalid German number format: " & strAmount
        End If
    Next lngCharacter

    If Not blnDigitFound Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The amount does not contain any digits: " & strAmount
    End If

    ParseGermanAmount = Val(strNormalized)

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
        "strAmount", strAmount)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strAmount - amount in German text format
' Returns:       String - 18 absolute cent digits followed by a sign
' Description:   Builds the 19-character MONTANT field.
'-------------------------------------------------------------------------------
Private Function BuildAmountField(ByVal strAmount As String) As String
    Const METHOD_NAME As String = "BuildAmountField"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dblAmount As Double
    Dim dblCents As Double
    Dim dblRoundedCents As Double
    Dim strDigits As String
    Dim strSign As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblAmount = ParseGermanAmount(strAmount)
    dblCents = dblAmount * 100#

    If dblCents >= 0# Then
        dblRoundedCents = Fix(dblCents + 0.5)
    Else
        dblRoundedCents = -Fix(Abs(dblCents) + 0.5)
    End If

    If Abs(dblRoundedCents) >= 1E+18 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The amount exceeds the 18-digit MONTANT capacity: " & strAmount
    End If

    strDigits = VBA.Format$(Abs(dblRoundedCents), "000000000000000000")
    If dblAmount < 0# Then
        strSign = "-"
    Else
        strSign = "+"
    End If

    BuildAmountField = strDigits & strSign

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
        "strAmount", strAmount)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    vValue - value read from the loader array
' Returns:       String - text without line-breaking control characters
' Description:   Converts array values to safe single-line record text.
'-------------------------------------------------------------------------------
Private Function ValueToRecordText(ByVal vValue As Variant) As String
    Const METHOD_NAME As String = "ValueToRecordText"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCharacter As Long
    Dim lngCharacterCode As Long
    Dim strCharacter As String
    Dim strClean As String
    Dim strSource As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The loader array contains an Excel error value."
    End If

    If IsNull(vValue) Or IsEmpty(vValue) Then
        ValueToRecordText = vbNullString
    Else
        strSource = CStr(vValue)

        For lngCharacter = 1 To Len(strSource)
            strCharacter = Mid$(strSource, lngCharacter, 1)
            lngCharacterCode = AscW(strCharacter)
            If lngCharacterCode < 0 Then lngCharacterCode = lngCharacterCode + 65536

            If lngCharacterCode < 32 Or lngCharacterCode = 160 Then
                strClean = strClean & " "
            Else
                strClean = strClean & strCharacter
            End If
        Next lngCharacter

        ValueToRecordText = strClean
    End If

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
' Parameters:    strValue - source text
'                lngWidth - required field width
' Returns:       String - left-truncated and right-padded fixed-width text
' Description:   Fits text into a fixed-width field using trailing spaces.
'-------------------------------------------------------------------------------
Private Function PadRight(ByVal strValue As String, _
                          ByVal lngWidth As Long) As String
    Const METHOD_NAME As String = "PadRight"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If lngWidth < 0 Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The requested fixed-width field length cannot be negative."
    End If

    If Len(strValue) >= lngWidth Then
        strResult = Left$(strValue, lngWidth)
    Else
        strResult = strValue & Space$(lngWidth - Len(strValue))
    End If

    PadRight = strResult

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
        "strValue;lngWidth", strValue, lngWidth)
    GoTo ExitPoint
End Function

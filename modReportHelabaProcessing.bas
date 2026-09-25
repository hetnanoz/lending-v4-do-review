Option Explicit


Private Const CLASS_NAME As String = "modReportHelabaProcessing"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrFond - filtered Fondsliste array
' Returns:       Object - Scripting.Dictionary mapping Account -> Fund number
' Description:   Builds the lookup that fills the Fund column. Account keys are
'                normalised (spaces removed, last 11 characters) as in
'                Tradeversand; the first occurrence of a key wins.
'-------------------------------------------------------------------------------
Public Function BuildFundDict(ByVal arrFond As Variant) As Object
    Const METHOD_NAME As String = "BuildFundDict"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dictFund As Object
    Dim lngR As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictFund = CreateObject("Scripting.Dictionary")

    If IsArray(arrFond) Then
        For lngR = 2 To UBound(arrFond, 1)
            strKey = NormalizeAccount(arrFond(lngR, FOND_OUT_ACCOUNT_COL))
            If Len(strKey) > 0 Then
                If Not dictFund.Exists(strKey) Then
                    dictFund.Add strKey, Trim$(CStr(arrFond(lngR, FOND_OUT_FUND_COL)))
                End If
            End If
        Next lngR
    End If

    Set BuildFundDict = dictFund

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
' Parameters:    arrFond - filtered Fondsliste array
' Returns:       Object - Scripting.Dictionary mapping Account -> Team
' Description:   Builds the lookup used to append Team to Report_Helaba and to
'                filter Loader_input. The first occurrence of an account wins.
'-------------------------------------------------------------------------------
Public Function BuildTeamDict(ByVal arrFond As Variant) As Object
    Const METHOD_NAME As String = "BuildTeamDict"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dictTeam As Object
    Dim lngR As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictTeam = CreateObject("Scripting.Dictionary")
    dictTeam.CompareMode = vbTextCompare

    If IsArray(arrFond) Then
        For lngR = 2 To UBound(arrFond, 1)
            strKey = NormalizeAccount(arrFond(lngR, FOND_OUT_ACCOUNT_COL))
            If Len(strKey) > 0 Then
                If Not dictTeam.Exists(strKey) Then
                    dictTeam.Add strKey, Trim$(CStr(arrFond(lngR, FOND_OUT_TEAM_COL)))
                End If
            End If
        Next lngR
    End If

    Set BuildTeamDict = dictTeam

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
' Parameters:    vValue - raw account value (any type)
' Returns:       String - normalised account key
' Description:   Removes spaces, keeps the last 11 characters and upper-cases the
'                result, so Report and Fondsliste accounts join reliably.
'-------------------------------------------------------------------------------
Private Function NormalizeAccount(ByVal vValue As Variant) As String
    Const METHOD_NAME As String = "NormalizeAccount"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strAccount As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strAccount = Replace(Trim$(CStr(vValue)), " ", "")
    If Len(strAccount) >= 11 Then strAccount = Right$(strAccount, 11)
    NormalizeAccount = UCase$(strAccount)

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
' Parameters:    arrTable - 1-based 2D array; row 1 holds the headers
' Returns:       Variant - the array without data rows empty in column 1
' Description:   Pure filter: keeps the header row and every data row whose first
'                column is not blank, dropping the subtotal ("... Total") rows.
'-------------------------------------------------------------------------------
Public Function RemoveEmptyFirstColumnRows(ByVal arrTable As Variant) As Variant
    Const METHOD_NAME As String = "RemoveEmptyFirstColumnRows"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngKeep As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim lngW As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrTable)
    If Not IsArray(arrTable) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngRows = UBound(arrTable, 1)
    lngCols = UBound(arrTable, 2)

    For lngR = 2 To lngRows
        If Not IsBlankValue(arrTable(lngR, 1)) Then lngKeep = lngKeep + 1
    Next lngR

    ReDim arrOut(1 To lngKeep + 1, 1 To lngCols)
    For lngC = 1 To lngCols
        arrOut(1, lngC) = arrTable(1, lngC)
    Next lngC

    lngW = 1
    For lngR = 2 To lngRows
        If Not IsBlankValue(arrTable(lngR, 1)) Then
            lngW = lngW + 1
            For lngC = 1 To lngCols
                arrOut(lngW, lngC) = arrTable(lngR, lngC)
            Next lngC
        End If
    Next lngR

    RemoveEmptyFirstColumnRows = arrOut

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
' Parameters:    vValue - a cell value (any type)
' Returns:       Boolean - True when the value is blank, empty or an error
' Description:   Safe emptiness test: Empty, Null, error values and whitespace
'                only text all count as blank. Also distinguishes "no accrual
'                match at all" (blank) from a genuine zero.
'-------------------------------------------------------------------------------
Private Function IsBlankValue(ByVal vValue As Variant) As Boolean
    Const METHOD_NAME As String = "IsBlankValue"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Then
        IsBlankValue = True
    ElseIf IsNull(vValue) Then
        IsBlankValue = True
    ElseIf IsEmpty(vValue) Then
        IsBlankValue = True
    Else
        IsBlankValue = (Len(Trim$(CStr(vValue))) = 0)
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
' Parameters:    arrReport - imported report table (Account in column E)
'                dictFund - Account -> Fund lookup
' Returns:       Variant - array with a "Fund" column prepended
' Description:   Pure transform: inserts the Fund column at position 1 (the other
'                columns shift right), filled by matching the normalised Account
'                against the Fondsliste lookup. Unmatched rows stay blank.
'-------------------------------------------------------------------------------
Public Function PrependFundColumn(ByVal arrReport As Variant, ByVal dictFund As Object) As Variant
    Const METHOD_NAME As String = "PrependFundColumn"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrReport)
    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngRows = UBound(arrReport, 1)
    lngCols = UBound(arrReport, 2)
    Debug.Assert lngCols >= FUND_ACCOUNT_COL

    ReDim arrOut(1 To lngRows, 1 To lngCols + 1)

    arrOut(1, 1) = FUND_COL_HEADER
    For lngC = 1 To lngCols
        arrOut(1, lngC + 1) = arrReport(1, lngC)
    Next lngC

    For lngR = 2 To lngRows
        strKey = NormalizeAccount(arrReport(lngR, FUND_ACCOUNT_COL))
        If dictFund.Exists(strKey) Then
            arrOut(lngR, 1) = dictFund(strKey)
        Else
            arrOut(lngR, 1) = vbNullString
        End If
        For lngC = 1 To lngCols
            arrOut(lngR, lngC + 1) = arrReport(lngR, lngC)
        Next lngC
    Next lngR

    PrependFundColumn = arrOut

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
' Parameters:    arrReport - report after Fund and optional accrual enrichment
'                dictTeam - Account -> Team lookup
' Returns:       Variant - report array with Team appended as the last column
' Description:   Keeps all existing columns unchanged and fills Team by matching
'                the report Account against the latest Fondsliste.
'-------------------------------------------------------------------------------
Public Function AppendTeamColumn(ByVal arrReport As Variant, _
                                 ByVal dictTeam As Object) As Variant
    Const METHOD_NAME As String = "AppendTeamColumn"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrReport)
    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngRows = UBound(arrReport, 1)
    lngCols = UBound(arrReport, 2)
    Debug.Assert lngCols >= REPORT_ACCOUNT_COL

    ReDim arrOut(1 To lngRows, 1 To lngCols + 1)

    For lngR = 1 To lngRows
        For lngC = 1 To lngCols
            arrOut(lngR, lngC) = arrReport(lngR, lngC)
        Next lngC
    Next lngR

    arrOut(1, lngCols + 1) = FOND_HDR_TEAM

    For lngR = 2 To lngRows
        strKey = NormalizeAccount(arrReport(lngR, REPORT_ACCOUNT_COL))
        If dictTeam.Exists(strKey) Then
            arrOut(lngR, lngCols + 1) = dictTeam(strKey)
        Else
            arrOut(lngR, lngCols + 1) = vbNullString
        End If
    Next lngR

    AppendTeamColumn = arrOut

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
' Parameters:    arrReport - 1-based report array
'                strHeader - header caption to find in row 1
' Returns:       Long - matching column number
' Description:   Finds report columns by header so Team can remain the final
'                column without breaking the loader's Accrual/TBB lookups.
'-------------------------------------------------------------------------------
Private Function FindReportHeaderColumn(ByVal arrReport As Variant, _
                                        ByVal strHeader As String, _
                                        Optional ByVal blnRequired As Boolean = True) As Long
    Const METHOD_NAME As String = "FindReportHeaderColumn"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngC As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngC = 1 To UBound(arrReport, 2)
        If StrComp(Trim$(CStr(arrReport(1, lngC))), strHeader, vbTextCompare) = 0 Then
            FindReportHeaderColumn = lngC
            Exit For
        End If
    Next lngC

    If FindReportHeaderColumn = 0 And blnRequired Then
        Err.Raise ERR_TABLE_NOT_FOUND, METHOD_NAME, _
                  "Header '" & strHeader & "' was not found in Report_Helaba."
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
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strHeader;blnRequired", strHeader, blnRequired)
    GoTo ExitPoint
End Function


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrReport - final report array containing Fund, Accrual, TBB and Team
' Returns:       Variant - 1-based 2D array: Team key, caption and fund list
' Description:   Returns only teams that would generate at least one loader line,
'                with each team's bookable funds for display and fund filtering.
'-------------------------------------------------------------------------------
Public Function BuildBookableTeamOptions(ByVal arrReport As Variant) As Variant
    Const METHOD_NAME As String = "BuildBookableTeamOptions"
    Dim arrFunds As Variant
    Dim arrOptions As Variant
    Dim arrTeams As Variant
    Dim blnBookable As Boolean
    Dim dblAccrual As Double
    Dim dblTbb As Double
    Dim dictFunds As Object
    Dim dictTeams As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngAccCol As Long
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngR As Long
    Dim lngTbbCol As Long
    Dim lngTeamCol As Long
    Dim strCaption As String
    Dim strFund As String
    Dim strFunds As String
    Dim strTeam As String
    Dim strTmp As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngAccCol = FindReportHeaderColumn(arrReport, ACC_HDR_ACCRUAL)
    lngTbbCol = FindReportHeaderColumn(arrReport, ACC_HDR_TBB)
    lngTeamCol = FindReportHeaderColumn(arrReport, FOND_HDR_TEAM)

    Set dictTeams = CreateObject("Scripting.Dictionary")
    dictTeams.CompareMode = vbTextCompare

    For lngR = 2 To UBound(arrReport, 1)
        If Not IsBlankValue(arrReport(lngR, lngTbbCol)) Then
            dblAccrual = ToDouble(arrReport(lngR, lngAccCol))
            dblTbb = ToDouble(arrReport(lngR, lngTbbCol))
            blnBookable = CeilTwo(Abs(dblAccrual)) > 0 Or CeilTwo(Abs(dblTbb)) > 0

            If blnBookable Then
                strTeam = Trim$(CStr(arrReport(lngR, lngTeamCol)))
                strFund = Trim$(CStr(arrReport(lngR, REPORT_FUND_COL)))

                If Len(strTeam) > 0 Then
                    If Not dictTeams.Exists(strTeam) Then
                        Set dictFunds = CreateObject("Scripting.Dictionary")
                        dictFunds.CompareMode = vbTextCompare
                        dictTeams.Add strTeam, dictFunds
                    Else
                        Set dictFunds = dictTeams(strTeam)
                    End If

                    If Len(strFund) > 0 Then
                        If Not dictFunds.Exists(strFund) Then
                            dictFunds.Add strFund, strFund
                        End If
                    End If
                End If
            End If
        End If
    Next lngR

    If dictTeams.Count = 0 Then
        BuildBookableTeamOptions = Empty
    Else
        arrTeams = dictTeams.Keys

        For lngI = LBound(arrTeams) To UBound(arrTeams) - 1
            For lngJ = lngI + 1 To UBound(arrTeams)
                If StrComp(CStr(arrTeams(lngI)), CStr(arrTeams(lngJ)), _
                           vbTextCompare) > 0 Then
                    strTmp = CStr(arrTeams(lngI))
                    arrTeams(lngI) = arrTeams(lngJ)
                    arrTeams(lngJ) = strTmp
                End If
            Next lngJ
        Next lngI

        ReDim arrOptions(1 To dictTeams.Count, 1 To 3)

        For lngI = LBound(arrTeams) To UBound(arrTeams)
            strTeam = CStr(arrTeams(lngI))
            Set dictFunds = dictTeams(strTeam)
            strFunds = vbNullString

            If dictFunds.Count > 0 Then
                arrFunds = dictFunds.Keys

                For lngJ = LBound(arrFunds) To UBound(arrFunds)
                    If Len(strFunds) > 0 Then strFunds = strFunds & ", "
                    strFunds = strFunds & CStr(arrFunds(lngJ))
                Next lngJ
            End If

            strCaption = strTeam
            If Len(strFunds) > 0 Then
                strCaption = strCaption & " (" & strFunds & ")"
            End If

            arrOptions(lngI - LBound(arrTeams) + 1, 1) = strTeam
            arrOptions(lngI - LBound(arrTeams) + 1, 2) = strCaption
            arrOptions(lngI - LBound(arrTeams) + 1, 3) = _
                Replace(strFunds, ", ", "|")
        Next lngI

        BuildBookableTeamOptions = arrOptions
    End If

ExitPoint:
    Set dictFunds = Nothing
    Set dictTeams = Nothing

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
' Parameters:    strTeam - team value from Report_Helaba
'                blnAllTeams - True to include every row
'                vSelectedTeams - optional dictionary of checked teams
' Returns:       Boolean - True when the report row belongs in Loader_input
' Description:   Applies the team filter chosen in the selection form.
'-------------------------------------------------------------------------------
Private Function TeamIsIncluded(ByVal strTeam As String, _
                                ByVal blnAllTeams As Boolean, _
                                Optional ByVal vSelectedTeams As Variant) As Boolean
    Const METHOD_NAME As String = "TeamIsIncluded"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dictSelectedTeams As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If blnAllTeams Then
        TeamIsIncluded = True
    ElseIf Not IsMissing(vSelectedTeams) Then
        Set dictSelectedTeams = vSelectedTeams

        If Not dictSelectedTeams Is Nothing Then
            TeamIsIncluded = dictSelectedTeams.Exists(Trim$(strTeam))
        End If
    End If

ExitPoint:
    Set dictSelectedTeams = Nothing

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
        "strTeam;blnAllTeams", strTeam, blnAllTeams)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strFund - fund value from Report_Helaba
'                blnAllFunds - True to include every available fund
'                vSelectedFunds - optional dictionary of selected fund numbers
' Returns:       Boolean - True when the report row belongs in Loader_input
' Description:   Applies the fund filter chosen in the selection form.
'-------------------------------------------------------------------------------
Private Function FundIsIncluded(ByVal strFund As String, _
                                ByVal blnAllFunds As Boolean, _
                                Optional ByVal vSelectedFunds As Variant) As Boolean
    Const METHOD_NAME As String = "FundIsIncluded"
    Dim dictSelectedFunds As Object
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If blnAllFunds Then
        FundIsIncluded = True
    ElseIf Not IsMissing(vSelectedFunds) Then
        Set dictSelectedFunds = vSelectedFunds

        If Not dictSelectedFunds Is Nothing Then
            FundIsIncluded = dictSelectedFunds.Exists(Trim$(strFund))
        End If
    End If

ExitPoint:
    Set dictSelectedFunds = Nothing

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
        "strFund;blnAllFunds", strFund, blnAllFunds)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrReport - report array (col 1 = Fund, col 3 = CURRENCY)
'                dictAcc - FUND|CODE dictionary from BuildAccrualDict
'                strSuffix - account suffix entered by the user
' Returns:       Variant - array with Acc_nr_and_suffix, Accrual and
'                TBB_on_729000_0 appended
' Description:   Pure transform: builds ACC_BASE_NUMBER & suffix & currency per
'                row and looks it up together with the fund. A matched row uses
'                the accrual from the LHS file. A missing match defaults to
'                Accrual = 0 and TBB = Fund received, so it remains bookable.
'-------------------------------------------------------------------------------
Public Function AppendAccrualColumns(ByVal arrReport As Variant, _
                                     ByVal dictAcc As Object, _
                                     ByVal strSuffix As String) As Variant
    Const METHOD_NAME As String = "AppendAccrualColumns"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrHit As Variant
    Dim arrOut As Variant
    Dim dblAccrual As Double
    Dim dblBase As Double
    Dim lngC As Long
    Dim lngCols As Long
    Dim lngR As Long
    Dim lngRows As Long
    Dim strCurrency As String
    Dim strExpectedCode As String
    Dim strFund As String
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrReport)
    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngRows = UBound(arrReport, 1)
    lngCols = UBound(arrReport, 2)
    Debug.Assert lngCols >= REPORT_TBB_BASE_COL

    ReDim arrOut(1 To lngRows, 1 To lngCols + 3)

    For lngR = 1 To lngRows
        For lngC = 1 To lngCols
            arrOut(lngR, lngC) = arrReport(lngR, lngC)
        Next lngC
    Next lngR

    arrOut(1, lngCols + 1) = ACC_HDR_CODE
    arrOut(1, lngCols + 2) = ACC_HDR_ACCRUAL
    arrOut(1, lngCols + 3) = ACC_HDR_TBB

    For lngR = 2 To lngRows
        strFund = UCase$(Trim$(CStr(arrReport(lngR, REPORT_FUND_COL))))
        strCurrency = UCase$(Trim$(CStr(arrReport(lngR, REPORT_CURRENCY_COL))))

        If Len(strFund) > 0 And Len(strCurrency) > 0 Then
            strExpectedCode = ACC_BASE_NUMBER & strSuffix & strCurrency
            strKey = strFund & "|" & strExpectedCode
            dblBase = ToDouble(arrReport(lngR, REPORT_TBB_BASE_COL))

            If dictAcc.Exists(strKey) Then
                arrHit = dictAcc(strKey)
                arrOut(lngR, lngCols + 1) = arrHit(0)
                arrOut(lngR, lngCols + 2) = arrHit(1)

                dblAccrual = ToDouble(arrHit(1))
                arrOut(lngR, lngCols + 3) = dblBase - dblAccrual
            Else
                arrOut(lngR, lngCols + 1) = strExpectedCode
                arrOut(lngR, lngCols + 2) = 0
                arrOut(lngR, lngCols + 3) = dblBase
            End If
        End If
    Next lngR

    AppendAccrualColumns = arrOut

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
' Parameters:    vValue - a cell value (any type)
' Returns:       Double - numeric value; 0 for blanks, errors and non-numeric text
' Description:   Safe numeric conversion for the TBB difference and the booking
'                amounts. Accepts numbers and numeric text, also with a decimal
'                comma and space grouping.
'-------------------------------------------------------------------------------
Public Function ToDouble(ByVal vValue As Variant) As Double
    Const METHOD_NAME As String = "ToDouble"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not IsError(vValue) And Not IsNull(vValue) And Not IsEmpty(vValue) Then
        If IsNumeric(vValue) Then
            ToDouble = CDbl(vValue)
        Else
            strValue = Replace(Replace(Trim$(CStr(vValue)), " ", ""), ",", ".")
            If IsNumeric(strValue) Then ToDouble = CDbl(strValue)
        End If
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
' Parameters:    arrReport - final report array, including Team as last column
'                strSuffix - month suffix (e.g. "6")
'                strYear - booking year (e.g. "2026")
'                blnAllTeams - True to build the loader for every team
'                vSelectedTeams - optional dictionary of checked team names
'                blnAllFunds - True to build the loader for every available fund
'                vSelectedFunds - optional dictionary of selected fund numbers
' Returns:       Variant - header row plus one row per booking line
' Description:   Pure transform implementing the booking rules documented in the
'                constants section, optionally filtered by Team and Fund. Accrual,
'                TBB and Team columns are located by header rather than by position.
'-------------------------------------------------------------------------------
Public Function BuildLoaderRows(ByVal arrReport As Variant, _
                                ByVal strSuffix As String, _
                                ByVal strYear As String, _
                                Optional ByVal blnAllTeams As Boolean = True, _
                                Optional ByVal vSelectedTeams As Variant, _
                                Optional ByVal blnAllFunds As Boolean = True, _
                                Optional ByVal vSelectedFunds As Variant) As Variant
    Const METHOD_NAME As String = "BuildLoaderRows"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrBuf As Variant
    Dim arrOut As Variant
    Dim lngRows As Long
    Dim lngTbbCol As Long
    Dim lngAccCol As Long
    Dim lngTeamCol As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim lngW As Long
    Dim strDesc As String
    Dim strDate As String
    Dim strFund As String
    Dim strCurrency As String
    Dim strTeam As String
    Dim dblAccrual As Double
    Dim dblTbb As Double

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrReport)
    If Not IsArray(arrReport) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Report table is not an array."
    End If

    lngRows = UBound(arrReport, 1)
    lngAccCol = FindReportHeaderColumn(arrReport, ACC_HDR_ACCRUAL)
    lngTbbCol = FindReportHeaderColumn(arrReport, ACC_HDR_TBB)
    lngTeamCol = FindReportHeaderColumn(arrReport, FOND_HDR_TEAM, False)
    Debug.Assert lngAccCol > REPORT_CURRENCY_COL

    strDesc = BuildDescription(strSuffix, strYear)
    strDate = FormatDateDdMmYyyy(Date)

    ReDim arrBuf(1 To (lngRows - 1) * 2 + 1, 1 To LOADER_COL_COUNT)
    arrBuf(1, 1) = HDR_FONDS
    arrBuf(1, 2) = HDR_CURRENCY
    arrBuf(1, 3) = HDR_AMOUNT
    arrBuf(1, 4) = HDR_PURPOSE
    arrBuf(1, 5) = HDR_KONTO_DB
    arrBuf(1, 6) = HDR_SUFFIX_DB
    arrBuf(1, 7) = HDR_KONTO_CR
    arrBuf(1, 8) = HDR_SUFFIX_CR
    arrBuf(1, 9) = HDR_DATE
    lngW = 1

    For lngR = 2 To lngRows
        If lngTeamCol > 0 Then
            strTeam = Trim$(CStr(arrReport(lngR, lngTeamCol)))
        Else
            strTeam = vbNullString
        End If

        strFund = Trim$(CStr(arrReport(lngR, REPORT_FUND_COL)))

        If TeamIsIncluded(strTeam, blnAllTeams, vSelectedTeams) _
           And FundIsIncluded(strFund, blnAllFunds, vSelectedFunds) _
           And Not IsBlankValue(arrReport(lngR, lngTbbCol)) Then
            strCurrency = Trim$(CStr(arrReport(lngR, REPORT_CURRENCY_COL)))
            dblAccrual = ToDouble(arrReport(lngR, lngAccCol))
            dblTbb = ToDouble(arrReport(lngR, lngTbbCol))

            If dblTbb = 0 Then
                If dblAccrual >= 0 Then
                    If Not AddLoaderLine(arrBuf, lngW, strFund, strCurrency, dblAccrual, _
                                           strDesc, strDate, ACC_MAIN, ACC_MAIN_SUFFIX, _
                                           ACC_BASE_NUMBER, strSuffix) Then
                        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_LOADER_LINE
                    End If
                Else
                    If Not AddLoaderLine(arrBuf, lngW, strFund, strCurrency, dblAccrual, _
                                           strDesc, strDate, ACC_BASE_NUMBER, strSuffix, _
                                           ACC_MAIN, ACC_MAIN_SUFFIX) Then
                        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_LOADER_LINE
                    End If
                End If
            Else
                If Not AddLoaderLine(arrBuf, lngW, strFund, strCurrency, dblAccrual, _
                                       strDesc, strDate, ACC_MAIN, ACC_MAIN_SUFFIX, _
                                       ACC_BASE_NUMBER, strSuffix) Then
                    Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_LOADER_LINE
                End If
                If dblTbb > 0 Then
                    If Not AddLoaderLine(arrBuf, lngW, strFund, strCurrency, dblTbb, _
                                           strDesc, strDate, ACC_MAIN, ACC_MAIN_SUFFIX, _
                                           ACC_TBB, ACC_TBB_SUFFIX) Then
                        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_LOADER_LINE
                    End If
                Else
                    If Not AddLoaderLine(arrBuf, lngW, strFund, strCurrency, dblTbb, _
                                           strDesc, strDate, ACC_TBB, ACC_TBB_SUFFIX, _
                                           ACC_MAIN, ACC_MAIN_SUFFIX) Then
                        Err.Raise ERR_NO_WORK, METHOD_NAME, ERR_TXT_LOADER_LINE
                    End If
                End If
            End If
        End If
    Next lngR

    ReDim arrOut(1 To lngW, 1 To LOADER_COL_COUNT)
    For lngR = 1 To lngW
        For lngC = 1 To LOADER_COL_COUNT
            arrOut(lngR, lngC) = arrBuf(lngR, lngC)
        Next lngC
    Next lngR

    BuildLoaderRows = arrOut

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
        "strSuffix;strYear;blnAllTeams;blnAllFunds", _
        strSuffix, strYear, blnAllTeams, blnAllFunds)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    arrBuf - target array (ByRef)
'                lngW - current write row (ByRef, incremented on success)
'                strFund, strCurrency - fund number and currency
'                dblAmount - signed amount; the absolute value is booked
'                strDesc - purpose text
'                strDate - booking date, already formatted
'                strDbAcc, strDbSuffix - debit account and suffix
'                strCrAcc, strCrSuffix - credit account and suffix
' Returns:       Boolean - True when the line was appended or intentionally skipped
' Description:   Appends one booking line. Amounts are made absolute, rounded up
'                to 2 decimals and written as text; the direction of the flow is
'                carried by the DB/CR accounts, not by the sign.
'-------------------------------------------------------------------------------
Private Function AddLoaderLine(ByRef arrBuf As Variant, ByRef lngW As Long, _
                               ByVal strFund As String, ByVal strCurrency As String, _
                               ByVal dblAmount As Double, ByVal strDesc As String, _
                               ByVal strDate As String, _
                               ByVal strDbAcc As String, ByVal strDbSuffix As String, _
                               ByVal strCrAcc As String, ByVal strCrSuffix As String) As Boolean
    Const METHOD_NAME As String = "AddLoaderLine"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dblBooked As Double

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblBooked = CeilTwo(Abs(dblAmount))

    AddLoaderLine = True

    If dblBooked <> 0 Then
        lngW = lngW + 1
        arrBuf(lngW, 1) = strFund
        arrBuf(lngW, 2) = strCurrency
        arrBuf(lngW, 3) = FormatAmountDe(dblBooked)
        arrBuf(lngW, 4) = strDesc
        arrBuf(lngW, 5) = strDbAcc
        arrBuf(lngW, 6) = strDbSuffix
        arrBuf(lngW, 7) = strCrAcc
        arrBuf(lngW, 8) = strCrSuffix
        arrBuf(lngW, 9) = strDate
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
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strFund;strCurrency;strDbAcc;strCrAcc", strFund, strCurrency, strDbAcc, strCrAcc)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    dblValue - value to round (expected non-negative)
' Returns:       Double - value rounded UP at 2 decimals
' Description:   Ceiling at 2 decimals (7290.021 -> 7290.03). The intermediate is
'                rounded at 6 decimals first, because binary representation would
'                otherwise turn 5662.61 * 100 into 566260.99... and round it up.
'-------------------------------------------------------------------------------
Public Function CeilTwo(ByVal dblValue As Double) As Double
    Const METHOD_NAME As String = "CeilTwo"
    Dim errDescription As String
    Dim errNumber As Long
    Dim dblScaled As Double

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblScaled = VBA.Round(dblValue * 100, 6)
    If dblScaled <> Int(dblScaled) Then dblScaled = Int(dblScaled) + 1
    CeilTwo = dblScaled / 100

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
        "dblValue", dblValue)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    dblValue - value to format
' Returns:       String - German formatted amount, e.g. 1123098.78 -> 1.123.098,78
' Description:   Builds the string manually instead of using a format mask,
'                because Format$ takes its separators from the system locale while
'                the loader always expects dot thousands and comma decimals.
'-------------------------------------------------------------------------------
Public Function FormatAmountDe(ByVal dblValue As Double) As String
    Const METHOD_NAME As String = "FormatAmountDe"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strRaw As String
    Dim strSign As String
    Dim strInt As String
    Dim strDec As String
    Dim strGrouped As String
    Dim lngDot As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dblValue < 0 Then strSign = "-"

    strRaw = Replace(VBA.Format$(Abs(dblValue), "0.00"), ",", ".")
    lngDot = InStr(strRaw, ".")
    strInt = Left$(strRaw, lngDot - 1)
    strDec = Mid$(strRaw, lngDot + 1)

    Do While Len(strInt) > 3
        strGrouped = "." & Right$(strInt, 3) & strGrouped
        strInt = Left$(strInt, Len(strInt) - 3)
    Loop

    FormatAmountDe = strSign & strInt & strGrouped & "," & strDec

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
        "dblValue", dblValue)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    dtValue - the date to format
' Returns:       String - the date as DD/MM/YYYY
' Description:   Assembled from the date parts rather than a format mask, because
'                "/" in a VBA mask is the locale date separator placeholder and
'                would render as a dot or a dash on some systems.
'-------------------------------------------------------------------------------
Public Function FormatDateDdMmYyyy(ByVal dtValue As Date) As String
    Const METHOD_NAME As String = "FormatDateDdMmYyyy"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strDay As String
    Dim strMonth As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strDay = Right$("0" & CStr(VBA.Day(dtValue)), 2)
    strMonth = Right$("0" & CStr(VBA.Month(dtValue)), 2)
    FormatDateDdMmYyyy = strDay & "/" & strMonth & "/" & CStr(VBA.Year(dtValue))

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
        "dtValue", dtValue)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    strSuffix - month suffix (e.g. "6")
'                strYear - booking year (e.g. "2026")
' Returns:       String - e.g. "Agency Lending Fees June 2026"
' Description:   Builds the purpose text. Month names are hardcoded in English
'                because VBA's MonthName() would follow the system locale.
'-------------------------------------------------------------------------------
Private Function BuildDescription(ByVal strSuffix As String, ByVal strYear As String) As String
    Const METHOD_NAME As String = "BuildDescription"
    Dim errDescription As String
    Dim errNumber As Long
    Dim arrMonths As Variant
    Dim strOut As String
    Dim lngMonth As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrMonths = Array("January", "February", "March", "April", "May", "June", _
                      "July", "August", "September", "October", "November", "December")

    strOut = LOADER_DESC_PREFIX
    If IsNumeric(strSuffix) Then
        lngMonth = CLng(Val(strSuffix))
        If lngMonth >= 1 And lngMonth <= 12 Then
            strOut = strOut & " " & arrMonths(lngMonth - 1)
        End If
    End If
    If Len(strYear) > 0 Then strOut = strOut & " " & strYear

    BuildDescription = strOut

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
        "strSuffix;strYear", strSuffix, strYear)
    GoTo ExitPoint
End Function

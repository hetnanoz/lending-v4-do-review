Option Explicit


Private Const CLASS_NAME As String = "modReportHelabaOutput"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - workbook that receives the sheet
'                strSheetName - target sheet name (recreated on every run)
'                arrData - 1-based 2D array to write
'                blnBoldHeader - bold the first row (ignored when blnAsTable)
'                blnAsTable - format the range as an Excel table
'                vTextCols - array of column indices to pre-format as text
' Returns:       Boolean - True when the sheet was written
' Description:   Recreates the sheet and dumps the array in one block. Columns
'                listed in vTextCols are formatted as text BEFORE writing, so
'                Excel cannot reparse values such as "5.662,61" or "30/07/2026".
'-------------------------------------------------------------------------------
Public Function WriteArrayToSheet(ByVal wkbTarget As Excel.Workbook, _
                                  ByVal strSheetName As String, _
                                  ByVal arrData As Variant, _
                                  ByVal blnBoldHeader As Boolean, _
                                  ByVal blnAsTable As Boolean, _
                                  ByVal vTextCols As Variant) As Boolean
    Const METHOD_NAME As String = "WriteArrayToSheet"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wksOut As Excel.Worksheet
    Dim objTable As Excel.ListObject
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngI As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Debug.Assert IsArray(arrData)
    If Not IsArray(arrData) Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, "Nothing to write to sheet " & strSheetName & "."
    End If

    Call RemoveSheet(wkbTarget, strSheetName)
    Set wksOut = wkbTarget.Worksheets.Add(After:=wkbTarget.Sheets(wkbTarget.Sheets.Count))
    wksOut.Name = strSheetName

    lngRows = UBound(arrData, 1)
    lngCols = UBound(arrData, 2)

    If IsArray(vTextCols) Then
        For lngI = LBound(vTextCols) To UBound(vTextCols)
            wksOut.Columns(CLng(vTextCols(lngI))).NumberFormat = "@"
        Next lngI
    End If

    wksOut.Range("A1").Resize(lngRows, lngCols).Value = arrData

    If blnAsTable Then
        Set objTable = wksOut.ListObjects.Add(xlSrcRange, _
                          wksOut.Range("A1").Resize(lngRows, lngCols), , xlYes)
        objTable.TableStyle = TABLE_STYLE
    ElseIf blnBoldHeader Then
        wksOut.Range("A1").Resize(1, lngCols).Font.Bold = True
    End If

    wksOut.UsedRange.Columns.AutoFit
    WriteArrayToSheet = True

ExitPoint:
    Set objTable = Nothing
    Set wksOut = Nothing

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
        "strSheetName", strSheetName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - the freshly built output workbook
'                strFolder - normalized folder read from save_path
' Returns:       String - full path the workbook was saved to
' Description:   Applies the Internal label and saves the output workbook while
'                restoring the previous DisplayAlerts value.
'-------------------------------------------------------------------------------
Public Function SaveOutputWorkbook( _
        ByVal wkbTarget As Excel.Workbook, _
        ByVal strFolder As String) As String

    Const METHOD_NAME As String = "SaveOutputWorkbook"
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim dtNow As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFullPath As String
    Dim strStamp As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Len(Trim$(strFolder)) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, ERR_TXT_NO_SAVE_CONFIG
    End If

    dtNow = Now
    strStamp = VBA.Format$(dtNow, "yyyymmdd") & "_" & _
               VBA.Format$(dtNow, "hhnnss")
    strFullPath = strFolder & OUTPUT_FILE_PREFIX & strStamp & OUTPUT_FILE_EXT

    Call ApplySensitivityLabel(wkbTarget, MIP_INTERNAL_LABEL_ID)

    blnPreviousDisplayAlerts = Application.DisplayAlerts
    blnSettingsCaptured = True
    Application.DisplayAlerts = False

    wkbTarget.SaveAs Filename:=strFullPath, FileFormat:=xlOpenXMLWorkbook
    SaveOutputWorkbook = wkbTarget.FullName

ExitPoint:
    If blnSettingsCaptured Then
        Application.DisplayAlerts = blnPreviousDisplayAlerts
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
        "strFolder", strFolder)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - output workbook receiving the sensitivity label
'                strLabelId - Microsoft Information Protection label identifier
' Returns:       ---
' Description:   Applies the corporate Internal sensitivity label before SaveAs
'                so Excel does not display the manual label selection prompt.
'-------------------------------------------------------------------------------
Private Sub ApplySensitivityLabel( _
        ByVal wkbTarget As Excel.Workbook, _
        ByVal strLabelId As String)
    Const METHOD_NAME As String = "ApplySensitivityLabel"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objLabelInfo As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbTarget Is Nothing Then
        Err.Raise ERR_NO_WORK, METHOD_NAME, _
                  "The target workbook is not available."
    End If

    If Len(Trim$(strLabelId)) = 0 Then
        Err.Raise ERR_CONFIG_MISSING, METHOD_NAME, _
                  "The sensitivity label identifier is empty."
    End If

    Set objLabelInfo = wkbTarget.SensitivityLabel.CreateLabelInfo()

    With objLabelInfo
        .LabelId = strLabelId
        .SiteId = MIP_SITE_ID
        .AssignmentMethod = MIP_ASSIGNMENT_METHOD
        .ContentBits = MIP_CONTENT_BITS
        .IsEnabled = True
    End With

    Call wkbTarget.SensitivityLabel.SetLabel(objLabelInfo, objLabelInfo)

ExitPoint:
    Set objLabelInfo = Nothing

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
        "strLabelId", strLabelId)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - workbook that may contain the sheet
'                strSheetName - name of the sheet to remove
' Returns:       ---
' Description:   Deletes the requested sheet when present and restores the
'                previous DisplayAlerts setting.
'-------------------------------------------------------------------------------
Public Sub RemoveSheet(ByVal wkbTarget As Excel.Workbook, ByVal strSheetName As String)
    Const METHOD_NAME As String = "RemoveSheet"
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnSettingsCaptured As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim wksTarget As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each wksTarget In wkbTarget.Worksheets
        If StrComp(wksTarget.Name, strSheetName, vbTextCompare) = 0 Then
            Exit For
        End If
    Next wksTarget

    If Not wksTarget Is Nothing Then
        If StrComp(wksTarget.Name, strSheetName, vbTextCompare) <> 0 Then
            Set wksTarget = Nothing
        End If
    End If

    If Not wksTarget Is Nothing Then
        blnPreviousDisplayAlerts = Application.DisplayAlerts
        blnSettingsCaptured = True
        Application.DisplayAlerts = False
        wksTarget.Delete
    End If

ExitPoint:
    If blnSettingsCaptured Then
        Application.DisplayAlerts = blnPreviousDisplayAlerts
    End If

    Set wksTarget = Nothing

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
        "strSheetName", strSheetName)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - workbook that may contain the query
'                strQueryName - name of the Power Query to remove
' Returns:       ---
' Description:   Deletes a Power Query and its auto-created connection, so the
'                tool can be re-run without a duplicate name error.
'-------------------------------------------------------------------------------
Public Sub RemoveQuery(ByVal wkbTarget As Excel.Workbook, ByVal strQueryName As String)
    Const METHOD_NAME As String = "RemoveQuery"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objConnection As Excel.WorkbookConnection
    Dim objQuery As Object
    Dim lngI As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngI = wkbTarget.Connections.Count To 1 Step -1
        Set objConnection = wkbTarget.Connections(lngI)
        If objConnection.Name = "Query - " & strQueryName _
           Or objConnection.Name = strQueryName Then
            objConnection.Delete
        End If
    Next lngI

    For lngI = wkbTarget.Queries.Count To 1 Step -1
        Set objQuery = wkbTarget.Queries(lngI)
        If StrComp(objQuery.Name, strQueryName, vbTextCompare) = 0 Then
            objQuery.Delete
            Exit For
        End If
    Next lngI

ExitPoint:
    Set objConnection = Nothing
    Set objQuery = Nothing

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
        "strQueryName", strQueryName)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-04
' Parameters:    wkbTarget - workbook reference to release
'                blnClose - True when the workbook should be closed
'                lngCloseError - returned cleanup error number
'                strCloseError - returned cleanup error description
' Returns:       Boolean - True when no Close error occurred
' Description:   Closes a workbook through an isolated handler so cleanup errors
'                never replace the original business error.
'-------------------------------------------------------------------------------
Public Function CloseWorkbookSafely(ByRef wkbTarget As Excel.Workbook, _
                                     ByVal blnClose As Boolean, _
                                     ByRef lngCloseError As Long, _
                                     ByRef strCloseError As String) As Boolean
    Const METHOD_NAME As String = "CloseWorkbookSafely"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngCloseError = 0
    strCloseError = vbNullString
    CloseWorkbookSafely = True

    If Not wkbTarget Is Nothing Then
        If blnClose Then
            Application.CutCopyMode = False
            DoEvents
            wkbTarget.Close SaveChanges:=False
        End If
    End If

ExitPoint:
    Set wkbTarget = Nothing
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    lngCloseError = errNumber
    strCloseError = errDescription
    CloseWorkbookSafely = False

    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "blnClose;lngCloseError;strCloseError", blnClose, lngCloseError, strCloseError)
    GoTo ExitPoint
End Function

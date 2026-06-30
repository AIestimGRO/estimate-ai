Attribute VB_Name = "Module6"
' ============================================================
'  MODULE 6: GESN EXCEPTIONS v1.1
'  Reestr utverzhdennyh diapazonov MIN/MAX.
'  Price_Check_Log: odna problema = odna stroka.
'  Prinyatie: postavit 1 v kolonke Approve i zapustit makros.
' ============================================================
Option Explicit

Public Const GESN_EXCEPTIONS_SHEET As String = "GESN_Exceptions"
Public Const PRICE_CHECK_LOG_SHEET As String = "Price_Check_Log"

Private Const LOG_COL_REASON As Long = 2
Private Const LOG_COL_STATUS As Long = 3
Private Const LOG_COL_APPROVE As Long = 4
Private Const LOG_COL_NORM_UNIT As Long = 7
Private Const LOG_COL_NORM_GESN As Long = 8
Private Const LOG_COL_DEM_FLAG As Long = 9
Private Const LOG_COL_SUGGESTED_MIN As Long = 16
Private Const LOG_COL_SUGGESTED_MAX As Long = 17
Private Const LOG_COL_MIN_CATALOG_ROW As Long = 18
Private Const LOG_COL_MAX_CATALOG_ROW As Long = 22
Private Const LOG_COL_NEW_CATALOG_ROW As Long = 26

Public Function GetDemKey(ByVal sourceHasDem As Boolean, ByVal demontazhFilterEnabled As Boolean) As String
    If demontazhFilterEnabled Then
        If sourceHasDem Then
            GetDemKey = "DEM"
        Else
            GetDemKey = "NO_DEM"
        End If
    Else
        GetDemKey = "DEM_FILTER_OFF"
    End If
End Function

Public Function GetGesnExceptionKey(ByVal unitKey As String, ByVal codeKey As String, ByVal demKey As String) As String
    GetGesnExceptionKey = Trim(CStr(unitKey)) & "||" & Trim(CStr(codeKey)) & "||" & UCase$(Trim(CStr(demKey)))
End Function

Public Function BuildGesnExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim ws As Worksheet
    Set ws = EnsureGesnExceptionsSheet()
    If ws Is Nothing Then
        Set BuildGesnExceptionsDict = d
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 2 Then
        Set BuildGesnExceptionsDict = d
        Exit Function
    End If

    Dim r As Long
    Dim unitKey As String
    Dim codeKey As String
    Dim demKey As String
    Dim approvedMin As Double
    Dim approvedMax As Double
    Dim firstDateSerial As Double
    Dim lastUpdateSerial As Double
    Dim minRow As Long
    Dim maxRow As Long
    Dim key As String

    For r = 2 To lastRow
        unitKey = NormUnit(ws.Cells(r, 2).Value)
        codeKey = NormCode(ws.Cells(r, 3).Value)
        demKey = UCase$(Trim(CStr(ws.Cells(r, 4).Value)))
        If demKey = "" Then demKey = "NO_DEM"

        approvedMin = ReadPositiveDouble(ws.Cells(r, 5).Value)
        approvedMax = ReadPositiveDouble(ws.Cells(r, 6).Value)

        If unitKey <> "" And codeKey <> "" And approvedMin > 0 And approvedMax > 0 Then
            key = GetGesnExceptionKey(unitKey, codeKey, demKey)
            firstDateSerial = ReadDateSerial(ws.Cells(r, 7).Value)
            lastUpdateSerial = ReadDateSerial(ws.Cells(r, 8).Value)
            minRow = CLng(ReadPositiveDouble(ws.Cells(r, 9).Value))
            maxRow = CLng(ReadPositiveDouble(ws.Cells(r, 10).Value))

            ws.Cells(r, 1).Value = key
            d(key) = Array(approvedMin, approvedMax, firstDateSerial, lastUpdateSerial, minRow, maxRow, r)
        End If
    Next r

    Set BuildGesnExceptionsDict = d
End Function

Public Function EnsureGesnExceptionsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GESN_EXCEPTIONS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = GESN_EXCEPTIONS_SHEET
    End If

    Dim headers As Variant
    headers = Array("ExceptionKey", _
                    "NormUnit", _
                    "NormGESN", _
                    "DemFlag", _
                    "ApprovedMinPrice", _
                    "ApprovedMaxPrice", _
                    "FirstExceptionDate", _
                    "LastRangeUpdateDate", _
                    "ApprovedMinCatalogRow", _
                    "ApprovedMaxCatalogRow", _
                    "Comment")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If Trim(CStr(ws.Cells(1, i + 1).Value)) = "" Then
            ws.Cells(1, i + 1).Value = headers(i)
        End If
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
    End With

    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).AutoFilter
    ws.Columns("A:K").AutoFit
    On Error GoTo 0

    Set EnsureGesnExceptionsSheet = ws
End Function

Public Sub InitGesnExceptionsSheet()
    Dim ws As Worksheet
    Set ws = EnsureGesnExceptionsSheet()
    If ws Is Nothing Then
        MsgBox "Ne udalos sozdat' list GESN_Exceptions.", vbCritical
    Else
        MsgBox "List GESN_Exceptions gotov.", vbInformation
    End If
End Sub

Public Sub ApproveCurrentPriceException()
    Dim wsLog As Worksheet
    Set wsLog = GetPriceCheckLogSheet()
    If wsLog Is Nothing Then Exit Sub

    If ActiveSheet.Name <> wsLog.Name Then
        MsgBox "Perejdite na Price_Check_Log i vyberite yachejku v stroke dlya utverzhdeniya.", vbExclamation
        Exit Sub
    End If

    If ActiveCell.Row < 2 Then
        MsgBox "Vyberite stroku dannyh, a ne zagolovok.", vbExclamation
        Exit Sub
    End If

    Dim updatedCount As Long
    updatedCount = ApprovePriceLogRow(wsLog, ActiveCell.Row)

    If updatedCount > 0 Then
        MsgBox "Diapazon v GESN_Exceptions obnovlen.", vbInformation
    End If
End Sub

Public Sub ApproveSelectedPriceException()
    ' Backward-compatible name: accepts the current active row.
    Call ApproveCurrentPriceException
End Sub

Public Sub ApproveMarkedPriceExceptions()
    Dim wsLog As Worksheet
    Set wsLog = GetPriceCheckLogSheet()
    If wsLog Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, LOG_COL_NORM_GESN).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "V Price_Check_Log net strok dlya obrabotki.", vbInformation
        Exit Sub
    End If

    Dim r As Long
    Dim cnt As Long
    For r = 2 To lastRow
        If IsApproveMark(wsLog.Cells(r, LOG_COL_APPROVE).Value) Then
            cnt = cnt + ApprovePriceLogRow(wsLog, r)
        End If
    Next r

    MsgBox "Obrabotano otmechennyh strok: " & CStr(cnt) & ".", vbInformation
End Sub

Private Function ApprovePriceLogRow(ByVal wsLog As Worksheet, ByVal r As Long) As Long
    ApprovePriceLogRow = 0
    If r < 2 Then Exit Function

    Dim unitKey As String
    Dim codeKey As String
    Dim demKey As String
    Dim suggestedMin As Double
    Dim suggestedMax As Double
    Dim minCatalogRow As Long
    Dim maxCatalogRow As Long
    Dim newCatalogRow As Long
    Dim key As String

    unitKey = NormUnit(wsLog.Cells(r, LOG_COL_NORM_UNIT).Value)
    codeKey = NormCode(wsLog.Cells(r, LOG_COL_NORM_GESN).Value)
    demKey = UCase$(Trim(CStr(wsLog.Cells(r, LOG_COL_DEM_FLAG).Value)))
    If demKey = "" Then demKey = "NO_DEM"

    suggestedMin = ReadPositiveDouble(wsLog.Cells(r, LOG_COL_SUGGESTED_MIN).Value)
    suggestedMax = ReadPositiveDouble(wsLog.Cells(r, LOG_COL_SUGGESTED_MAX).Value)
    minCatalogRow = CLng(ReadPositiveDouble(wsLog.Cells(r, LOG_COL_MIN_CATALOG_ROW).Value))
    maxCatalogRow = CLng(ReadPositiveDouble(wsLog.Cells(r, LOG_COL_MAX_CATALOG_ROW).Value))
    newCatalogRow = CLng(ReadPositiveDouble(wsLog.Cells(r, LOG_COL_NEW_CATALOG_ROW).Value))

    If suggestedMin <= 0 Or suggestedMax <= 0 Then
        MsgBox "V stroke " & CStr(r) & " ne hvataet SuggestedNewMin/SuggestedNewMax.", vbExclamation
        Exit Function
    End If

    If unitKey = "" Or codeKey = "" Then
        MsgBox "V stroke " & CStr(r) & " ne hvataet NormUnit/NormGESN.", vbExclamation
        Exit Function
    End If

    key = GetGesnExceptionKey(unitKey, codeKey, demKey)

    Dim wsEx As Worksheet
    Set wsEx = EnsureGesnExceptionsSheet()
    If wsEx Is Nothing Then Exit Function

    Dim foundRow As Long
    foundRow = FindExceptionRow(wsEx, key)

    Dim outRow As Long
    Dim oldMin As Double
    Dim oldMax As Double
    Dim newMin As Double
    Dim newMax As Double

    If foundRow = 0 Then
        outRow = wsEx.Cells(wsEx.Rows.Count, 1).End(xlUp).Row + 1
        If outRow < 2 Then outRow = 2

        wsEx.Cells(outRow, 1).Value = key
        wsEx.Cells(outRow, 2).Value = unitKey
        wsEx.Cells(outRow, 3).Value = codeKey
        wsEx.Cells(outRow, 4).Value = demKey
        wsEx.Cells(outRow, 5).Value = suggestedMin
        wsEx.Cells(outRow, 6).Value = suggestedMax
        wsEx.Cells(outRow, 7).Value = Date
        wsEx.Cells(outRow, 8).Value = Date
        wsEx.Cells(outRow, 9).Value = ResolveMinCatalogRow(minCatalogRow, newCatalogRow)
        wsEx.Cells(outRow, 10).Value = ResolveMaxCatalogRow(maxCatalogRow, newCatalogRow)
        wsEx.Cells(outRow, 11).Value = "Approved from Price_Check_Log"
    Else
        outRow = foundRow
        oldMin = ReadPositiveDouble(wsEx.Cells(outRow, 5).Value)
        oldMax = ReadPositiveDouble(wsEx.Cells(outRow, 6).Value)

        newMin = oldMin
        newMax = oldMax
        If newMin <= 0 Or suggestedMin < newMin Then newMin = suggestedMin
        If newMax <= 0 Or suggestedMax > newMax Then newMax = suggestedMax

        wsEx.Cells(outRow, 5).Value = newMin
        wsEx.Cells(outRow, 6).Value = newMax
        If Trim(CStr(wsEx.Cells(outRow, 7).Value)) = "" Then wsEx.Cells(outRow, 7).Value = Date
        wsEx.Cells(outRow, 8).Value = Date

        If oldMin <= 0 Or suggestedMin < oldMin Then
            wsEx.Cells(outRow, 9).Value = ResolveMinCatalogRow(minCatalogRow, newCatalogRow)
        End If
        If oldMax <= 0 Or suggestedMax > oldMax Then
            wsEx.Cells(outRow, 10).Value = ResolveMaxCatalogRow(maxCatalogRow, newCatalogRow)
        End If
    End If

    wsLog.Cells(r, LOG_COL_STATUS).Value = "APPROVED_RANGE_EXPANDED"
    wsLog.Cells(r, LOG_COL_APPROVE).ClearContents
    wsLog.Rows(r).Interior.Color = RGB(226, 239, 218)
    wsEx.Columns("A:K").AutoFit

    ApprovePriceLogRow = 1
End Function

Public Sub StampMissingCatalogAddedDates()
    Call LoadSettings

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If gSettings.catAddedDateCol < 1 Then
        MsgBox "Ne zadana kolonka daty dobavleniya kataloga na Instrument.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, gSettings.catTaskCol).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, gSettings.catCodeCol).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, gSettings.catCodeCol).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, gSettings.catPriceCol).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, gSettings.catPriceCol).End(xlUp).Row

    If lastRow < 4 Then
        MsgBox "Ne najdeny stroki kataloga na aktivnom liste.", vbExclamation
        Exit Sub
    End If

    If Trim(CStr(ws.Cells(3, gSettings.catAddedDateCol).Value)) = "" Then
        ws.Cells(3, gSettings.catAddedDateCol).Value = "CatalogAddedDate"
        ws.Cells(3, gSettings.catAddedDateCol).Font.Bold = True
    End If

    Dim r As Long
    Dim cnt As Long
    For r = 4 To lastRow
        If Trim(CStr(ws.Cells(r, gSettings.catTaskCol).Value)) <> "" Or _
           Trim(CStr(ws.Cells(r, gSettings.catCodeCol).Value)) <> "" Or _
           Trim(CStr(ws.Cells(r, gSettings.catPriceCol).Value)) <> "" Then

            If Trim(CStr(ws.Cells(r, gSettings.catAddedDateCol).Value)) = "" Then
                ws.Cells(r, gSettings.catAddedDateCol).Value = Date
                ws.Cells(r, gSettings.catAddedDateCol).NumberFormat = "dd.mm.yyyy"
                cnt = cnt + 1
            End If
        End If
    Next r

    MsgBox "Zapolneno dat dobavleniya: " & CStr(cnt) & ".", vbInformation
End Sub

Private Function GetPriceCheckLogSheet() As Worksheet
    On Error Resume Next
    Set GetPriceCheckLogSheet = ThisWorkbook.Worksheets(PRICE_CHECK_LOG_SHEET)
    On Error GoTo 0

    If GetPriceCheckLogSheet Is Nothing Then
        MsgBox "List Price_Check_Log ne najden.", vbExclamation
    End If
End Function

Private Function IsApproveMark(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim(CStr(v)))

    IsApproveMark = (s = "1" Or s = "DA" Or s = "YES" Or s = "Y" Or s = "OK" Or _
                     s = ChrW(1044) & ChrW(1040))
End Function

Private Function ResolveMinCatalogRow(ByVal minCatalogRow As Long, ByVal newCatalogRow As Long) As Long
    If minCatalogRow > 0 Then
        ResolveMinCatalogRow = minCatalogRow
    Else
        ResolveMinCatalogRow = newCatalogRow
    End If
End Function

Private Function ResolveMaxCatalogRow(ByVal maxCatalogRow As Long, ByVal newCatalogRow As Long) As Long
    If maxCatalogRow > 0 Then
        ResolveMaxCatalogRow = maxCatalogRow
    Else
        ResolveMaxCatalogRow = newCatalogRow
    End If
End Function

Private Function FindExceptionRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim r As Long
    For r = 2 To lastRow
        If UCase$(Trim(CStr(ws.Cells(r, 1).Value))) = UCase$(Trim(key)) Then
            FindExceptionRow = r
            Exit Function
        End If
    Next r
End Function

Private Function ReadPositiveDouble(ByVal v As Variant) As Double
    On Error GoTo BadNumber

    ReadPositiveDouble = 0
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then ReadPositiveDouble = CDbl(v)
        Exit Function
    End If

    Dim s As String
    s = Trim(CStr(v))
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    If s = "" Then Exit Function

    s = Replace(s, ".", Application.DecimalSeparator)
    s = Replace(s, ",", Application.DecimalSeparator)
    If IsNumeric(s) Then
        If CDbl(s) > 0 Then ReadPositiveDouble = CDbl(s)
    End If
    Exit Function

BadNumber:
    ReadPositiveDouble = 0
End Function

Private Function ReadDateSerial(ByVal v As Variant) As Double
    On Error GoTo BadDate
    ReadDateSerial = 0
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Trim(CStr(v)) = "" Then Exit Function
    If IsDate(v) Then ReadDateSerial = CDbl(DateValue(CDate(v)))
    Exit Function
BadDate:
    ReadDateSerial = 0
End Function

Attribute VB_Name = "Module4"
' ============================================================
'  MODULE 3: OBRABOTKA SMETY v6.7
'  Ispravleniya v6.1:
'  - dobavlena proverka razbrosa cen analogov Max/Min
'  - problemnye analogi krasjatsya svetlo-krasnym
'  - polnye stroki kataloga pishutsya v Price_Check_Log
'  - vyvodimye ceny analogov umnozhayutsya na regionalnyj koefficient
'  - usilena ochistka starogo bloka analogov po fakticheski ispolzuemomu diapazonu
'  - poisk analogov utochnen po svyazke Ed.Izm + GESN/FER/Perechen
'  - filtr demontazha mozhno vklyuchat/otklyuchat na liste Instrument
' ============================================================
Option Explicit

Public Sub ProcessSmeta(ws As Worksheet, _
                         catalog As Object, _
                         headerRow As Long, _
                         dataStart As Long, _
                         analogStartCol As Long, _
                         sectionCol As Long, _
                         ByRef logLines() As String)

    Dim subStage As String
    Dim r As Long
    Dim mi As Long

    On Error GoTo ErrHandler

    subStage = "start ProcessSmeta"

    If ws.ProtectContents Then
        Err.Raise 1004, "ProcessSmeta", _
                  "List smety zashchishchen. Snimite zashchitu lista i zapustite makros povtorno."
    End If

    Dim secDict As Object: Set secDict = BuildSectionDict()
    Dim KR As String: KR = " /" & ChrW(1050) & ChrW(1056)

    If analogStartCol < 1 Then
        Err.Raise 1004, "ProcessSmeta", "analogStartCol = 0. Proverite nastrojki Instrument."
    End If

    ' === SHAG 1: Sobrat' rabochie stroki =======================
    subStage = "Shag 1: skanirovanie rabochih strok"
    Application.StatusBar = "[3/4] Shag 1: Skaniruyu stroki..."

    Dim lastRow As Long
    ' Vazhno: v nekotoryh smetah stolbec A v tele tablitsy pustoj.
    ' Poetomu poslednyuyu stroku opredelyaem ne po A, a po vazhnym stolbcam smety:
    ' kod GESN/FER/Perechen, bazovaya cena i naimenovanie rabot.
    lastRow = ws.Cells(ws.Rows.Count, gSettings.colSearch).End(xlUp).Row

    Dim lrTmp As Long
    lrTmp = ws.Cells(ws.Rows.Count, gSettings.colF).End(xlUp).Row
    If lrTmp > lastRow Then lastRow = lrTmp

    lrTmp = ws.Cells(ws.Rows.Count, gSettings.colSmetaWorkName).End(xlUp).Row
    If lrTmp > lastRow Then lastRow = lrTmp

    lrTmp = ws.Cells(ws.Rows.Count, gSettings.colSmetaUnit).End(xlUp).Row
    If lrTmp > lastRow Then lastRow = lrTmp

    Dim mainRows() As Long: ReDim mainRows(1 To lastRow)
    Dim mCount As Long: mCount = 0

    Dim nameExcludedRows As Object
    Set nameExcludedRows = CreateObject("Scripting.Dictionary")
    nameExcludedRows.CompareMode = 1

    ' Vazhnaya logika:
    ' ran'she obrabatyvalis' tol'ko stroki, gde v stolbce A byl celyj nomer
    ' (1, 2, 3...). Iz-za etogo stroki 1.1, 2.1, 3.15 i t.p. propuskalis'.
    ' Teper' obrabatyvaem kazhduyu stroku, gde est' kod v zadannom stolbce
    ' poiska, est' edinica izmereniya i est' bazovaya cena v zadannom stolbce bazy.
    For r = dataStart To lastRow
        Dim rowCode As String
        rowCode = NormCode(ws.Cells(r, gSettings.colSearch).Value)

        If rowCode <> "" Then
            Dim rowUnit As String
            rowUnit = NormUnit(ws.Cells(r, gSettings.colSmetaUnit).Value)

            If rowUnit <> "" Then
                Dim baseVal As Variant
                baseVal = ws.Cells(r, gSettings.colF).Value

                If Not IsEmpty(baseVal) And IsNumeric(baseVal) Then
                    If CDbl(baseVal) > 0 Then
                        mCount = mCount + 1
                        mainRows(mCount) = r

                        If IsNameExcluded(ws.Cells(r, gSettings.colSmetaWorkName).Value, "SMETA") Then
                            nameExcludedRows(CStr(r)) = True
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If mCount = 0 Then
        ReDim logLines(0)
        logLines(0) = "No data rows with code, unit and base price."
        Exit Sub
    End If
    ReDim Preserve mainRows(1 To mCount)

    ' === SHAG 2: Sobrat' analogi i poryadok stolbcov ============
    subStage = "Shag 2: podbor analogov i poryadok stolbcov"
    Application.StatusBar = "[3/4] Shag 2: Selecting analogs..."

    ' colDefs(ci, 0) = nomer zadachi
    ' colDefs(ci, 1) = region dlya etoj pozicii ceny
    ' colDefs(ci, 2) = nomer ceny vnutri odnoj zadachi: 1, 2, 3...
    Dim maxOutCols As Long
    maxOutCols = ws.Columns.Count - analogStartCol + 1
    If maxOutCols < 1 Then
        Err.Raise 1004, "ProcessSmeta", "Pervyj stolbec analogov vyshe dopustimogo chisla stolbcov Excel."
    End If

    Dim colDefs() As Variant
    ReDim colDefs(1 To maxOutCols, 0 To 2)

    Dim colCount As Long: colCount = 0
    Dim colKeyIdx As Object
    Set colKeyIdx = CreateObject("Scripting.Dictionary")
    colKeyIdx.CompareMode = 1

    Dim rowAnalogs() As Object: ReDim rowAnalogs(1 To mCount)

    Dim taskOrder As Collection: Set taskOrder = New Collection
    Dim taskMaxPi As Object: Set taskMaxPi = CreateObject("Scripting.Dictionary")
    Dim taskRegionByPos As Object: Set taskRegionByPos = CreateObject("Scripting.Dictionary")
    taskMaxPi.CompareMode = 1
    taskRegionByPos.CompareMode = 1

    For mi = 1 To mCount
        r = mainRows(mi)
        Dim normC As String: normC = NormCode(ws.Cells(r, gSettings.colSearch).Value)
        Dim lookupKey As String
        lookupKey = AnalogSearchKey(ws.Cells(r, gSettings.colSmetaUnit).Value, ws.Cells(r, gSettings.colSearch).Value)

        If nameExcludedRows.Exists(CStr(r)) Then GoTo NextMi

        If lookupKey <> "" And catalog.Exists(lookupKey) Then
            Dim sourceHasDem As Boolean
            sourceHasDem = HasDemontazh(ws.Cells(r, gSettings.colSmetaWorkName).Value)

            If gSettings.demontazhFilterEnabled Then
                Set rowAnalogs(mi) = FilterAnalogsByDemontazh(catalog(lookupKey), sourceHasDem)
            Else
                ' Filtr demontazha otklyuchen na Instrument: berem vse analogi po klyuchu Ed.Izm + Perechen.
                Set rowAnalogs(mi) = catalog(lookupKey)
            End If

            If rowAnalogs(mi).Count > 0 Then
                Dim tKey As Variant
                For Each tKey In rowAnalogs(mi).Keys
                    Dim tid As String: tid = CStr(tKey)
                    Dim entries As Collection
                    Set entries = rowAnalogs(mi)(tid)

                Dim priceCount As Long: priceCount = CLng(entries.Count)

                If Not taskMaxPi.Exists(tid) Then
                    taskOrder.Add tid
                    taskMaxPi.Add tid, priceCount
                Else
                    If priceCount > CLng(taskMaxPi(tid)) Then taskMaxPi(tid) = priceCount
                End If

                Dim epi As Long
                For epi = 1 To priceCount
                    Dim regKey As String: regKey = tid & "|" & CStr(epi)
                    Dim regVal As String: regVal = ""
                    On Error Resume Next
                    regVal = CStr(entries(epi)(1))
                    On Error GoTo ErrHandler

                    If Not taskRegionByPos.Exists(regKey) Then
                        taskRegionByPos.Add regKey, regVal
                    ElseIf CStr(taskRegionByPos(regKey)) = "" And regVal <> "" Then
                        taskRegionByPos(regKey) = regVal
                    End If
                Next epi
            Next tKey
        End If
    End If
NextMi:
Next mi

    Dim tOrd As Variant
    For Each tOrd In taskOrder
        Dim tidOrd As String: tidOrd = CStr(tOrd)
        Dim pi As Long
        For pi = 1 To CLng(taskMaxPi(tidOrd))
            colCount = colCount + 1
            If colCount > UBound(colDefs, 1) Then
                colCount = colCount - 1
                Exit For
            End If

            Dim posKey As String: posKey = tidOrd & "|" & CStr(pi)
            Dim posRegion As String: posRegion = ""
            If taskRegionByPos.Exists(posKey) Then posRegion = CStr(taskRegionByPos(posKey))

            colDefs(colCount, 0) = tidOrd
            colDefs(colCount, 1) = posRegion
            colDefs(colCount, 2) = pi
            If Not colKeyIdx.Exists(tidOrd & "|" & pi) Then colKeyIdx.Add tidOrd & "|" & pi, colCount
        Next pi
    Next tOrd

    ' === SHAG 3: Ochistit' starye analogi =======================
    subStage = "Shag 3: ochistka staryh analogovyh stolbcov"
    Application.StatusBar = "[3/4] Shag 3: Ochistka stolbcov..."

    Dim lastCol As Long
    ' Ran'she ochistka orientirovalas' tol'ko na poslednij zagolovok v headerRow.
    ' Iz-za etogo starie analogi/hvosty mogli ostavat'sya nizhe ili pravee.
    ' Teper' berem maksimalnyj stolbec iz zagolovka i fakticheski ispolzuemogo diapazona lista.
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim usedLastCol As Long
    usedLastCol = GetLastUsedColumn(ws)
    If usedLastCol > lastCol Then lastCol = usedLastCol

    If lastCol >= analogStartCol Then
        SafeClearAnalogBlock ws, headerRow, lastRow, analogStartCol, lastCol
    End If

    ' === SHAG 4: Zagolovki zadach/regionov ======================
    subStage = "Shag 4: zapis zagolovkov analogov"
    Dim hFillBlue As Long: hFillBlue = RGB(217, 225, 242)
    Dim taskFillBlue As Long: taskFillBlue = RGB(221, 235, 247)
    Dim dupFill   As Long: dupFill = RGB(217, 217, 217)
    Dim problemFill As Long: problemFill = RGB(255, 199, 206)

    If colCount > 0 Then
        Application.StatusBar = "[3/4] Shag 4: Pyshu zagolovki..."

        Dim ci As Long
        For ci = 1 To colCount
            Dim hCol As Long: hCol = analogStartCol + ci - 1

            PrepareCellForWrite ws.Cells(headerRow, hCol)
            With ws.Cells(headerRow, hCol)
                .Value = CStr(colDefs(ci, 0))
                .Font.Bold = True
                .Font.Size = 9
                .Interior.Color = hFillBlue
                .HorizontalAlignment = xlCenter
                .WrapText = True
            End With

            PrepareCellForWrite ws.Cells(headerRow + 1, hCol)
            With ws.Cells(headerRow + 1, hCol)
                .Value = CStr(colDefs(ci, 1))
                .Font.Italic = True
                .Font.Size = 9
                .Interior.Color = hFillBlue
                .HorizontalAlignment = xlCenter
                .WrapText = True
            End With

            On Error Resume Next
            ws.Columns(hCol).ColumnWidth = 16
            On Error GoTo ErrHandler

            If IsTaskMarkedBlue(colDefs(ci, 0)) Then
                ws.Range(ws.Cells(headerRow, hCol), ws.Cells(lastRow, hCol)).Interior.Color = taskFillBlue

                ws.Cells(headerRow, hCol).Interior.Color = hFillBlue
                ws.Cells(headerRow + 1, hCol).Interior.Color = hFillBlue
            End If
        Next ci
    End If

    Dim lastACol As Long
    If colCount > 0 Then
        lastACol = analogStartCol + colCount - 1
    Else
        lastACol = 0
    End If

    ' Regionalnyj koefficient berem iz yachejki obrabatyvaemogo lista smety.
    ' Adres yachejki zadaetsya na liste Instrument v fajle makrosa.
    Dim regionalCoef As Double
    regionalCoef = GetRegionalCoefficient(ws)

    ' Reestr utverzhdennyh diapazonov MIN/MAX po svyazke Ed.Izm + GESN + DemFlag.
    Dim gesnExceptions As Object
    Set gesnExceptions = BuildGesnExceptionsDict()

    ' === SHAG 5-9: Obrabotka kazhdoj stroki =====================
    subStage = "Shag 5: postroenie strochnyh formul i analogov"
    ReDim logLines(1 To mCount + 2)
    logLines(1) = "Stroka|Kod|EdIzm|Zadach|Cen|Razdel|Formula|PriceSpread|RegionalCoef"
    Dim logIdx As Long: logIdx = 1

    ' Ne dubliruem zapis' odnoj i toj zhe problemnoj pary MIN/MAX v Price_Check_Log
    ' v ramkah odnogo zapuska. Esli v smete odin i tot zhe kod+ed.izm.
    ' povtoryaetsya mnogo raz, v log popadet tolko pervaya para MIN/MAX.
    Dim priceLogSeen As Object
    Set priceLogSeen = CreateObject("Scripting.Dictionary")
    priceLogSeen.CompareMode = 1

    For mi = 1 To mCount
        r = mainRows(mi)

        Dim nfs As String
        nfs = NormCode(ws.Cells(r, gSettings.colSearch).Value)

        Dim isNameExcludedRow As Boolean
        isNameExcludedRow = nameExcludedRows.Exists(CStr(r))

        If mi Mod 10 = 0 Then
            Application.StatusBar = "[3/4] Shag 5: stroka " & mi & " iz " & mCount
        End If

        ' --- 6: Zapicat' ceny analogov -------------------------
        subStage = "Shag 6: zapis cen analogov, stroka " & r
        Dim totalP As Long: totalP = 0
        Dim hasAn  As Boolean: hasAn = False

        Dim isPriceProblem As Boolean
        Dim minAnalogPrice As Double
        Dim maxAnalogPrice As Double
        Dim priceRatio As Double
        Dim ratioProblem As Boolean

        ratioProblem = IsProblemPriceGroup(rowAnalogs(mi), _
                                           gSettings.priceSpreadLimit, _
                                           minAnalogPrice, _
                                           maxAnalogPrice, _
                                           priceRatio)

        Dim normUForRow As String
        normUForRow = NormUnit(ws.Cells(r, gSettings.colSmetaUnit).Value)

        Dim sourceHasDemForRow As Boolean
        sourceHasDemForRow = HasDemontazh(ws.Cells(r, gSettings.colSmetaWorkName).Value)

        Dim demKeyForRow As String
        demKeyForRow = GetDemKey(sourceHasDemForRow, gSettings.demontazhFilterEnabled)

        Dim exceptionKey As String
        exceptionKey = GetGesnExceptionKey(normUForRow, nfs, demKeyForRow)

        Dim hasApprovedException As Boolean
        hasApprovedException = False

        Dim approvedMinPrice As Double
        Dim approvedMaxPrice As Double
        Dim exceptionDateSerial As Double
        approvedMinPrice = 0
        approvedMaxPrice = 0
        exceptionDateSerial = 0

        Dim outOfRangeRows As Object
        Set outOfRangeRows = CreateObject("Scripting.Dictionary")
        outOfRangeRows.CompareMode = 1

        If Not gesnExceptions Is Nothing Then
            If gesnExceptions.Exists(exceptionKey) Then
                Dim exceptionRec As Variant
                exceptionRec = gesnExceptions(exceptionKey)
                approvedMinPrice = CDbl(exceptionRec(0))
                approvedMaxPrice = CDbl(exceptionRec(1))
                exceptionDateSerial = CDbl(exceptionRec(2))
                hasApprovedException = True

                MarkOutOfApprovedRange rowAnalogs(mi), _
                                       approvedMinPrice, _
                                       approvedMaxPrice, _
                                       exceptionDateSerial, _
                                       outOfRangeRows

                isPriceProblem = (outOfRangeRows.Count > 0)
            Else
                isPriceProblem = ratioProblem
            End If
        Else
            isPriceProblem = ratioProblem
        End If

        If Not rowAnalogs(mi) Is Nothing Then
            Dim tKey2 As Variant
            For Each tKey2 In rowAnalogs(mi).Keys
                Dim tid2 As String: tid2 = CStr(tKey2)
                Dim entries2 As Collection
                Set entries2 = rowAnalogs(mi)(tid2)

                Dim pi2 As Long
                For pi2 = 1 To entries2.Count
                    Dim ck2 As String: ck2 = tid2 & "|" & CStr(pi2)
                    If colKeyIdx.Exists(ck2) Then
                        Dim dCol As Long
                        dCol = analogStartCol + CLng(colKeyIdx(ck2)) - 1

                        Dim priceVal As Double
                        priceVal = CDbl(entries2(pi2)(0)) * regionalCoef

                        PrepareCellForWrite ws.Cells(r, dCol)
                        With ws.Cells(r, dCol)
                            .Value = Round(priceVal, 2)
                            .NumberFormat = "#,##0.00"
                            If IsTaskMarkedBlue(tid2) Then
                                .Interior.Color = taskFillBlue
                            ElseIf ShouldColorAnalog(entries2(pi2), isPriceProblem, hasApprovedException, outOfRangeRows) Then
                                .Interior.Color = problemFill
                            ElseIf pi2 > 1 Then
                                .Interior.Color = dupFill
                            End If
                        End With

                        totalP = totalP + 1
                        hasAn = True
                    End If
                Next pi2
            Next tKey2
        End If

        If isPriceProblem Then
            subStage = "Shag 6b: log problemnyh analogov po cenam, stroka " & r
            If hasApprovedException Then
                AppendOutOfRangePriceLog rowAnalogs(mi), _
                                         ws.Name, _
                                         r, _
                                         nfs, _
                                         normUForRow, _
                                         demKeyForRow, _
                                         gSettings.priceSpreadLimit, _
                                         minAnalogPrice, _
                                         maxAnalogPrice, _
                                         priceRatio, _
                                         approvedMinPrice, _
                                         approvedMaxPrice, _
                                         outOfRangeRows, _
                                         priceLogSeen, _
                                         regionalCoef
            Else
                AppendPriceCheckLog rowAnalogs(mi), _
                                    ws.Name, _
                                    r, _
                                    nfs, _
                                    normUForRow, _
                                    demKeyForRow, _
                                    gSettings.priceSpreadLimit, _
                                    minAnalogPrice, _
                                    maxAnalogPrice, _
                                    priceRatio, _
                                    priceLogSeen, _
                                    regionalCoef
            End If
        End If

        ' --- 7: Kod razdela - zapisyvaem kak tekst -------------
        subStage = "Shag 7: zapis koda razdela, stroka " & r
        Dim sc  As String: sc = ""
        If nfs <> "" Then
            sc = ResolveSectionCode(nfs, ws.Cells(r, gSettings.colSmetaWorkName).Value, secDict)
        End If
        If sc <> "" Then
            PrepareCellForWrite ws.Cells(r, sectionCol)
            With ws.Cells(r, sectionCol)
                .NumberFormat = "@"
                .Value = sc
            End With
        End If

        ' --- 8: Dobavit' kod GESN + /KR v zadannyj stolbec esli est' analog --
        subStage = "Shag 8: dobavlenie koda GESN + /KR, stroka " & r
        If hasAn Then
            PrepareCellForWrite ws.Cells(r, gSettings.colKR)
            Dim krEnd As String: krEnd = "/" & ChrW(1050) & ChrW(1056)
            Dim krBase As String
            krBase = Trim(CStr(ws.Cells(r, gSettings.colSearch).Value))
            krBase = Replace(krBase, vbCr, " " )
            krBase = Replace(krBase, vbLf, " " )
            krBase = Replace(krBase, vbTab, " " )
            krBase = Replace(krBase, Chr(160), " " )
            Do While InStr(1, krBase, "  ", vbBinaryCompare) > 0
                krBase = Replace(krBase, "  ", " ")
            Loop
            krBase = Trim(krBase)

            With ws.Cells(r, gSettings.colKR)
                .NumberFormat = "@"
                If krBase = "" Then
                    .Value = krEnd
                ElseIf Right(UCase$(krBase), Len(krEnd)) = krEnd Then
                    .Value = krBase
                Else
                    .Value = krBase & KR
                End If
            End With
        End If

        ' --- 9: Stolbec srednej/itogovoj ceny = FORMULA --------
        subStage = "Shag 9: zapis formuly srednej ceny, stroka " & r
        Dim gCell As Range: Set gCell = ws.Cells(r, gSettings.colAvg)
        PrepareCellForWrite gCell

        If lastACol >= analogStartCol Then
            If Not PutAverageFormulaR1C1(gCell, gSettings.colF, analogStartCol, lastACol) Then
                Err.Raise 1004, "ProcessSmeta", _
                          "Ne udalos zapisat formulu srednej ceny v " & gCell.Address(False, False)
            End If
        Else
            If Not PutSimpleFormulaR1C1(gCell, gSettings.colF) Then
                Err.Raise 1004, "ProcessSmeta", _
                          "Ne udalos zapisat prostuyu formulu v " & gCell.Address(False, False)
            End If
        End If

        With gCell
            .NumberFormat = "#,##0.00"
            .Font.Bold = True
            .Font.Color = RGB(0, 112, 192)
            .Font.Size = 9
        End With

        logIdx = logIdx + 1
        Dim zadCnt As String
        If rowAnalogs(mi) Is Nothing Then
            zadCnt = "0"
        Else
            zadCnt = CStr(rowAnalogs(mi).Count)
        End If
        Dim priceSpreadInfo As String
        If isNameExcludedRow Then
            priceSpreadInfo = "SKIPPED_BY_NAME_EXCLUSION"
        ElseIf isPriceProblem Then
            If hasApprovedException Then
                priceSpreadInfo = "OUT_OF_APPROVED_RANGE " & Format(priceRatio, "0.00")
            Else
                priceSpreadInfo = "PROBLEM " & Format(priceRatio, "0.00")
            End If
        ElseIf hasApprovedException Then
            priceSpreadInfo = "APPROVED_RANGE " & Format(priceRatio, "0.00")
        ElseIf totalP >= 2 And minAnalogPrice > 0 Then
            priceSpreadInfo = Format(priceRatio, "0.00")
        Else
            priceSpreadInfo = ""
        End If

        logLines(logIdx) = r & "|" & nfs & "|" & normUForRow & "|" & zadCnt & _
                           "|" & totalP & "|" & sc & "|" & gCell.formulaR1C1 & _
                           "|" & priceSpreadInfo & "|" & Format(regionalCoef, "0.0000")
    Next mi

    ReDim Preserve logLines(1 To logIdx)
    Exit Sub

ErrHandler:
    Dim msg As String
    msg = Err.Description
    If Len(msg) = 0 Then msg = "Application-defined or object-defined error"
    Err.Raise Err.Number, "ProcessSmeta", _
              msg & vbCrLf & _
              "Podetap: " & subStage & vbCrLf & _
              "Excel-stroka: " & CStr(r) & ", indeks: " & CStr(mi)
End Sub

Private Function GetRegionalCoefficient(ByVal ws As Worksheet) As Double
    Dim addr As String
    addr = Trim(gSettings.regionalCoefCellAddress)

    ' Po TZ: esli adres pustoj ili koefficient ne najden / ne prochitan,
    ' makros ne dolzhen ostanavlivat'sya. V etom sluchae koefficient = 1.
    GetRegionalCoefficient = 1
    If addr = "" Then Exit Function

    On Error GoTo UseDefaultCoef

    addr = Replace(addr, "$", "")

    Dim v As Variant
    v = ws.Range(addr).Value

    If IsError(v) Or IsEmpty(v) Or Trim(CStr(v)) = "" Then
        GetRegionalCoefficient = 1
        Exit Function
    End If

    Dim coef As Double
    If Not TryReadDouble(v, coef) Then
        GetRegionalCoefficient = 1
        Exit Function
    End If

    If coef <= 0 Then
        GetRegionalCoefficient = 1
        Exit Function
    End If

    GetRegionalCoefficient = coef
    Exit Function

UseDefaultCoef:
    GetRegionalCoefficient = 1
End Function

Private Function TryReadDouble(ByVal v As Variant, ByRef result As Double) As Boolean
    On Error GoTo BadNumber

    If IsNumeric(v) Then
        result = CDbl(v)
        TryReadDouble = True
        Exit Function
    End If

    Dim s As String
    s = Trim(CStr(v))
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")

    If s = "" Then GoTo BadNumber

    ' Podderzhka i zapyatoj, i tochki kak decimalnogo razdelitelya.
    s = Replace(s, ".", Application.DecimalSeparator)
    s = Replace(s, ",", Application.DecimalSeparator)

    If IsNumeric(s) Then
        result = CDbl(s)
        TryReadDouble = True
    Else
        TryReadDouble = False
    End If
    Exit Function

BadNumber:
    TryReadDouble = False
End Function

Private Function IsProblemPriceGroup(ByVal rowTasks As Object, _
                                     ByVal threshold As Double, _
                                     ByRef minPrice As Double, _
                                     ByRef maxPrice As Double, _
                                     ByRef ratio As Double) As Boolean
    minPrice = 0
    maxPrice = 0
    ratio = 0
    IsProblemPriceGroup = False

    If rowTasks Is Nothing Then Exit Function
    If threshold <= 0 Then Exit Function

    Dim total As Long: total = 0
    Dim tKey As Variant
    Dim entries As Collection
    Dim ent As Variant
    Dim p As Double

    For Each tKey In rowTasks.Keys
        Set entries = rowTasks(tKey)

        For Each ent In entries
            If IsNumeric(ent(0)) Then
                p = CDbl(ent(0))

                If p > 0 Then
                    total = total + 1

                    If minPrice = 0 Or p < minPrice Then minPrice = p
                    If p > maxPrice Then maxPrice = p
                End If
            End If
        Next ent
    Next tKey

    If total < 2 Then Exit Function
    If minPrice <= 0 Then Exit Function

    ratio = maxPrice / minPrice
    IsProblemPriceGroup = (ratio >= threshold)
End Function

Private Sub MarkOutOfApprovedRange(ByVal rowTasks As Object, _
                                  ByVal approvedMinPrice As Double, _
                                  ByVal approvedMaxPrice As Double, _
                                  ByVal exceptionDateSerial As Double, _
                                  ByVal outOfRangeRows As Object)
    If rowTasks Is Nothing Then Exit Sub
    If outOfRangeRows Is Nothing Then Exit Sub
    If approvedMinPrice <= 0 And approvedMaxPrice <= 0 Then Exit Sub

    Dim tKey As Variant
    Dim entries As Collection
    Dim ent As Variant
    Dim p As Double
    Dim addedSerial As Double

    For Each tKey In rowTasks.Keys
        Set entries = rowTasks(tKey)
        For Each ent In entries
            p = CDbl(ent(0))
            addedSerial = EntAddedDateSerial(ent)

            ' Esli data ne zapolnena, ne schitaem stroku novoj, chtoby ne dat' lozhnyj flag.
            If addedSerial > 0 Then
                If exceptionDateSerial <= 0 Or addedSerial > exceptionDateSerial Then
                    If (approvedMinPrice > 0 And p < approvedMinPrice) Or _
                       (approvedMaxPrice > 0 And p > approvedMaxPrice) Then
                        outOfRangeRows(CStr(CLng(ent(3)))) = True
                    End If
                End If
            End If
        Next ent
    Next tKey
End Sub

Private Function ShouldColorAnalog(ByVal ent As Variant, _
                                   ByVal isPriceProblem As Boolean, _
                                   ByVal hasApprovedException As Boolean, _
                                   ByVal outOfRangeRows As Object) As Boolean
    ShouldColorAnalog = False
    If Not isPriceProblem Then Exit Function

    If hasApprovedException Then
        If outOfRangeRows Is Nothing Then Exit Function
        ShouldColorAnalog = outOfRangeRows.Exists(CStr(CLng(ent(3))))
    Else
        ShouldColorAnalog = True
    End If
End Function

Private Function EntAddedDateSerial(ByVal ent As Variant) As Double
    On Error GoTo BadDate
    EntAddedDateSerial = 0
    If IsArray(ent) Then
        If UBound(ent) >= 8 Then
            If IsNumeric(ent(8)) Then EntAddedDateSerial = CDbl(ent(8))
        End If
    End If
    Exit Function
BadDate:
    EntAddedDateSerial = 0
End Function

Private Sub AppendPriceCheckLog(ByVal rowTasks As Object, _
                                ByVal smetaSheetName As String, _
                                ByVal smetaRow As Long, _
                                ByVal normCode As String, _
                                ByVal normUnit As String, _
                                ByVal demKey As String, _
                                ByVal threshold As Double, _
                                ByVal minPrice As Double, _
                                ByVal maxPrice As Double, _
                                ByVal ratio As Double, _
                                ByVal logSeen As Object, _
                                ByVal regionalCoef As Double)
    If rowTasks Is Nothing Then Exit Sub

    Dim minEnt As Variant
    Dim maxEnt As Variant
    Dim hasMin As Boolean: hasMin = False
    Dim hasMax As Boolean: hasMax = False

    Dim tKey As Variant
    Dim entries As Collection
    Dim ent As Variant
    Dim p As Double

    For Each tKey In rowTasks.Keys
        Set entries = rowTasks(tKey)

        For Each ent In entries
            If IsNumeric(ent(0)) Then
                p = CDbl(ent(0))

                If p > 0 Then
                    If Not hasMin Then
                        minEnt = ent
                        hasMin = True
                    ElseIf p < CDbl(minEnt(0)) Then
                        minEnt = ent
                    End If

                    If Not hasMax Then
                        maxEnt = ent
                        hasMax = True
                    ElseIf p > CDbl(maxEnt(0)) Then
                        maxEnt = ent
                    End If
                End If
            End If
        Next ent
    Next tKey

    If Not hasMin Or Not hasMax Then Exit Sub

    Dim pairKey As String
    pairKey = "RATIO_EXCEEDED||" & CStr(normUnit) & "||" & CStr(normCode) & "||" & _
              CStr(demKey) & "||" & _
              CStr(CLng(minEnt(3))) & "||" & CStr(CLng(maxEnt(3)))

    If Not logSeen Is Nothing Then
        If logSeen.Exists(pairKey) Then Exit Sub
        logSeen.Add pairKey, True
    End If

    Dim wsLog As Worksheet
    Set wsLog = GetOrCreatePriceCheckLog(rowTasks)
    If wsLog Is Nothing Then Exit Sub

    EnsurePriceLogHeaders wsLog

    Dim outRow As Long
    outRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If outRow < 2 Then outRow = 2

    WriteRatioPriceCheckLogRow wsLog, outRow, smetaSheetName, smetaRow, _
                               normUnit, normCode, demKey, threshold, _
                               minPrice, maxPrice, ratio, minEnt, maxEnt, regionalCoef

    wsLog.Columns("A:AK").AutoFit
End Sub

Private Sub AppendOutOfRangePriceLog(ByVal rowTasks As Object, _
                                     ByVal smetaSheetName As String, _
                                     ByVal smetaRow As Long, _
                                     ByVal normCode As String, _
                                     ByVal normUnit As String, _
                                     ByVal demKey As String, _
                                     ByVal threshold As Double, _
                                     ByVal minPrice As Double, _
                                     ByVal maxPrice As Double, _
                                     ByVal ratio As Double, _
                                     ByVal approvedMinPrice As Double, _
                                     ByVal approvedMaxPrice As Double, _
                                     ByVal outOfRangeRows As Object, _
                                     ByVal logSeen As Object, _
                                     ByVal regionalCoef As Double)
    If rowTasks Is Nothing Then Exit Sub
    If outOfRangeRows Is Nothing Then Exit Sub
    If outOfRangeRows.Count = 0 Then Exit Sub

    Dim wsLog As Worksheet
    Set wsLog = GetOrCreatePriceCheckLog(rowTasks)
    If wsLog Is Nothing Then Exit Sub

    EnsurePriceLogHeaders wsLog

    Dim tKey As Variant
    Dim entries As Collection
    Dim ent As Variant
    Dim rowKey As String
    Dim problemKey As String
    Dim p As Double
    Dim suggestedMin As Double
    Dim suggestedMax As Double
    Dim outRow As Long

    For Each tKey In rowTasks.Keys
        Set entries = rowTasks(tKey)
        For Each ent In entries
            rowKey = CStr(CLng(ent(3)))
            If outOfRangeRows.Exists(rowKey) Then
                p = CDbl(ent(0))
                problemKey = "OUT_OF_APPROVED_RANGE||" & GetGesnExceptionKey(normUnit, normCode, demKey) & "||" & rowKey & "||" & Format(p, "0.0000")

                If Not logSeen Is Nothing Then
                    If logSeen.Exists(problemKey) Then GoTo NextEnt
                    logSeen.Add problemKey, True
                End If

                suggestedMin = approvedMinPrice
                suggestedMax = approvedMaxPrice
                If suggestedMin <= 0 Or p < suggestedMin Then suggestedMin = p
                If suggestedMax <= 0 Or p > suggestedMax Then suggestedMax = p

                outRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
                If outRow < 2 Then outRow = 2

                WriteOutOfRangePriceCheckLogRow wsLog, outRow, smetaSheetName, smetaRow, _
                                                normUnit, normCode, demKey, threshold, _
                                                minPrice, maxPrice, ratio, approvedMinPrice, _
                                                approvedMaxPrice, suggestedMin, suggestedMax, ent, regionalCoef
            End If
NextEnt:
        Next ent
    Next tKey

    wsLog.Columns("A:AK").AutoFit
End Sub

Private Function GetOrCreatePriceCheckLog(ByVal rowTasks As Object) As Worksheet
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("Price_Check_Log")
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "Price_Check_Log"
    End If

    Set GetOrCreatePriceCheckLog = wsLog
End Function

Private Sub WriteRatioPriceCheckLogRow(ByVal wsLog As Worksheet, _
                                       ByVal outRow As Long, _
                                       ByVal smetaSheetName As String, _
                                       ByVal smetaRow As Long, _
                                       ByVal normUnit As String, _
                                       ByVal normCode As String, _
                                       ByVal demKey As String, _
                                       ByVal threshold As Double, _
                                       ByVal minPrice As Double, _
                                       ByVal maxPrice As Double, _
                                       ByVal ratio As Double, _
                                       ByVal minEnt As Variant, _
                                       ByVal maxEnt As Variant, _
                                       ByVal regionalCoef As Double)
    WriteBasePriceLogFields wsLog, outRow, "RATIO_EXCEEDED", smetaSheetName, smetaRow, _
                            normUnit, normCode, demKey, threshold, minPrice, maxPrice, ratio, _
                            0, 0, minPrice, maxPrice

    WriteMinMaxInfo wsLog, outRow, minEnt, maxEnt
    WriteOutputPriceInfo wsLog, outRow, regionalCoef, minPrice, maxPrice, 0
    wsLog.Rows(outRow).Interior.Color = RGB(255, 235, 238)
End Sub

Private Sub WriteOutOfRangePriceCheckLogRow(ByVal wsLog As Worksheet, _
                                            ByVal outRow As Long, _
                                            ByVal smetaSheetName As String, _
                                            ByVal smetaRow As Long, _
                                            ByVal normUnit As String, _
                                            ByVal normCode As String, _
                                            ByVal demKey As String, _
                                            ByVal threshold As Double, _
                                            ByVal minPrice As Double, _
                                            ByVal maxPrice As Double, _
                                            ByVal ratio As Double, _
                                            ByVal approvedMinPrice As Double, _
                                            ByVal approvedMaxPrice As Double, _
                                            ByVal suggestedMinPrice As Double, _
                                            ByVal suggestedMaxPrice As Double, _
                                            ByVal ent As Variant, _
                                            ByVal regionalCoef As Double)
    WriteBasePriceLogFields wsLog, outRow, "OUT_OF_APPROVED_RANGE", smetaSheetName, smetaRow, _
                            normUnit, normCode, demKey, threshold, minPrice, maxPrice, ratio, _
                            approvedMinPrice, approvedMaxPrice, suggestedMinPrice, suggestedMaxPrice

    WriteNewAnalogInfo wsLog, outRow, ent
    WriteOutputPriceInfo wsLog, outRow, regionalCoef, minPrice, maxPrice, CDbl(ent(0))
    wsLog.Rows(outRow).Interior.Color = RGB(255, 235, 238)
End Sub

Private Sub WriteBasePriceLogFields(ByVal wsLog As Worksheet, _
                                    ByVal outRow As Long, _
                                    ByVal checkReason As String, _
                                    ByVal smetaSheetName As String, _
                                    ByVal smetaRow As Long, _
                                    ByVal normUnit As String, _
                                    ByVal normCode As String, _
                                    ByVal demKey As String, _
                                    ByVal threshold As Double, _
                                    ByVal minPrice As Double, _
                                    ByVal maxPrice As Double, _
                                    ByVal ratio As Double, _
                                    ByVal approvedMinPrice As Double, _
                                    ByVal approvedMaxPrice As Double, _
                                    ByVal suggestedMinPrice As Double, _
                                    ByVal suggestedMaxPrice As Double)
    wsLog.Cells(outRow, 1).Value = Now
    wsLog.Cells(outRow, 2).Value = checkReason
    wsLog.Cells(outRow, 3).Value = "NEW"
    wsLog.Cells(outRow, 4).Value = ""
    wsLog.Cells(outRow, 5).Value = smetaSheetName
    wsLog.Cells(outRow, 6).Value = smetaRow
    wsLog.Cells(outRow, 7).Value = normUnit
    wsLog.Cells(outRow, 8).Value = normCode
    wsLog.Cells(outRow, 9).Value = demKey
    wsLog.Cells(outRow, 10).Value = threshold
    If minPrice > 0 Then wsLog.Cells(outRow, 11).Value = Round(minPrice, 2)
    If maxPrice > 0 Then wsLog.Cells(outRow, 12).Value = Round(maxPrice, 2)
    If ratio > 0 Then wsLog.Cells(outRow, 13).Value = Round(ratio, 4)
    If approvedMinPrice > 0 Then wsLog.Cells(outRow, 14).Value = Round(approvedMinPrice, 2)
    If approvedMaxPrice > 0 Then wsLog.Cells(outRow, 15).Value = Round(approvedMaxPrice, 2)
    If suggestedMinPrice > 0 Then wsLog.Cells(outRow, 16).Value = Round(suggestedMinPrice, 2)
    If suggestedMaxPrice > 0 Then wsLog.Cells(outRow, 17).Value = Round(suggestedMaxPrice, 2)
End Sub

Private Sub WriteMinMaxInfo(ByVal wsLog As Worksheet, _
                            ByVal outRow As Long, _
                            ByVal minEnt As Variant, _
                            ByVal maxEnt As Variant)
    wsLog.Cells(outRow, 18).Value = CLng(minEnt(3))
    wsLog.Cells(outRow, 19).Value = CStr(minEnt(5))
    wsLog.Cells(outRow, 20).Value = CStr(minEnt(1))
    wsLog.Cells(outRow, 21).Value = CDbl(minEnt(0))

    wsLog.Cells(outRow, 22).Value = CLng(maxEnt(3))
    wsLog.Cells(outRow, 23).Value = CStr(maxEnt(5))
    wsLog.Cells(outRow, 24).Value = CStr(maxEnt(1))
    wsLog.Cells(outRow, 25).Value = CDbl(maxEnt(0))
End Sub

Private Sub WriteNewAnalogInfo(ByVal wsLog As Worksheet, _
                               ByVal outRow As Long, _
                               ByVal ent As Variant)
    Dim addedSerial As Double
    addedSerial = EntAddedDateSerial(ent)

    wsLog.Cells(outRow, 26).Value = CLng(ent(3))
    wsLog.Cells(outRow, 27).Value = CStr(ent(5))
    wsLog.Cells(outRow, 28).Value = CStr(ent(1))
    wsLog.Cells(outRow, 29).Value = CDbl(ent(0))
    If addedSerial > 0 Then wsLog.Cells(outRow, 30).Value = CDate(addedSerial)
    wsLog.Cells(outRow, 31).Value = CStr(ent(6))
    If UBound(ent) >= 7 Then wsLog.Cells(outRow, 32).Value = CStr(ent(7))
    wsLog.Cells(outRow, 33).Value = IIf(CBool(ent(2)), "DEM", "NO_DEM")
End Sub

Private Sub WriteOutputPriceInfo(ByVal wsLog As Worksheet, _
                                 ByVal outRow As Long, _
                                 ByVal regionalCoef As Double, _
                                 ByVal minPrice As Double, _
                                 ByVal maxPrice As Double, _
                                 ByVal newPrice As Double)
    If regionalCoef <= 0 Then regionalCoef = 1

    wsLog.Cells(outRow, 34).Value = Round(regionalCoef, 4)
    If minPrice > 0 Then wsLog.Cells(outRow, 35).Value = Round(minPrice * regionalCoef, 2)
    If maxPrice > 0 Then wsLog.Cells(outRow, 36).Value = Round(maxPrice * regionalCoef, 2)
    If newPrice > 0 Then wsLog.Cells(outRow, 37).Value = Round(newPrice * regionalCoef, 2)
End Sub

Private Sub EnsurePriceLogHeaders(ByVal wsLog As Worksheet)
    Dim headers(1 To 37) As String
    headers(1) = "RunDate"
    headers(2) = "CheckReason"
    headers(3) = "Status"
    headers(4) = "Approve"
    headers(5) = "SmetaSheet"
    headers(6) = "SmetaRow"
    headers(7) = "NormUnit"
    headers(8) = "NormGESN"
    headers(9) = "DemFlag"
    headers(10) = "Threshold"
    headers(11) = "MinAnalogPrice"
    headers(12) = "MaxAnalogPrice"
    headers(13) = "MaxMinRatio"
    headers(14) = "ApprovedMinPrice"
    headers(15) = "ApprovedMaxPrice"
    headers(16) = "SuggestedNewMin"
    headers(17) = "SuggestedNewMax"
    headers(18) = "MinCatalogRow"
    headers(19) = "MinTaskId"
    headers(20) = "MinRegion"
    headers(21) = "MinPrice"
    headers(22) = "MaxCatalogRow"
    headers(23) = "MaxTaskId"
    headers(24) = "MaxRegion"
    headers(25) = "MaxPrice"
    headers(26) = "NewCatalogRow"
    headers(27) = "NewTaskId"
    headers(28) = "NewRegion"
    headers(29) = "NewPrice"
    headers(30) = "CatalogAddedDate"
    headers(31) = "CatalogNormCode"
    headers(32) = "CatalogNormUnit"
    headers(33) = "CatalogDemFlag"
    headers(34) = "SmetaRegionalCoef"
    headers(35) = "MinOutputPrice"
    headers(36) = "MaxOutputPrice"
    headers(37) = "NewOutputPrice"

    Dim i As Long

    For i = LBound(headers) To UBound(headers)
        wsLog.Cells(1, i).Value = headers(i)
    Next i

    With wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(1, UBound(headers)))
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        On Error Resume Next
        .AutoFilter
        On Error GoTo 0
    End With
End Sub


Private Function FilterAnalogsByDemontazh(srcTasks As Object, sourceHasDem As Boolean) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = 1

    Dim tKey As Variant
    Dim srcEntries As Collection
    Dim keptEntries As Collection
    Dim ent As Variant
    Dim analogHasDem As Boolean

    For Each tKey In srcTasks.Keys
        Set srcEntries = srcTasks(tKey)
        Set keptEntries = New Collection

        For Each ent In srcEntries
            analogHasDem = False

            On Error Resume Next
            analogHasDem = CBool(ent(2))
            On Error GoTo 0

            ' Demontazh filter is strict:
            ' - source with demont -> only analogs with demont
            ' - source without demont -> only analogs without demont
            ' Word "montazh" is neutral and does not count as demontazh.
            If sourceHasDem Then
                If analogHasDem Then keptEntries.Add ent
            Else
                If Not analogHasDem Then keptEntries.Add ent
            End If
        Next ent

        If keptEntries.Count > 0 Then
            res.Add CStr(tKey), keptEntries
        End If
    Next tKey

    Set FilterAnalogsByDemontazh = res
End Function
Private Sub SafeClearAnalogBlock(ws As Worksheet, _
                                 ByVal headerRow As Long, _
                                 ByVal lastRow As Long, _
                                 ByVal firstCol As Long, _
                                 ByVal lastCol As Long)
    Dim clearLastRow As Long
    clearLastRow = lastRow

    ' Vazhno: starie analogi mogut byt' nizhe raschetnogo lastRow,
    ' naprimer posle proshlogo zapuska s drugoj logikoj opredeleniya konca tablitsy.
    ' Poetomu ochishchaem do poslednej fakticheski zapolnennoj stroki lista.
    Dim usedLastRow As Long
    usedLastRow = GetLastUsedRow(ws)
    If usedLastRow > clearLastRow Then clearLastRow = usedLastRow
    If clearLastRow < headerRow Then clearLastRow = headerRow

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(clearLastRow + 5, lastCol))

    ' Starie vyvodnye stolbcy mogut soderzhat merge-yachejki.
    ' Esli ih ne razobedinit, zapis zagolovkov/chisel mozhet dat oshibku 1004.
    On Error Resume Next
    rng.UnMerge
    rng.ClearContents
    rng.ClearFormats
    On Error GoTo 0
End Sub

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", _
                          After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    On Error GoTo 0

    If f Is Nothing Then
        GetLastUsedRow = 1
    Else
        GetLastUsedRow = f.Row
    End If
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", _
                          After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    On Error GoTo 0

    If f Is Nothing Then
        GetLastUsedColumn = 1
    Else
        GetLastUsedColumn = f.Column
    End If
End Function

Private Sub PrepareCellForWrite(ByVal target As Range)
    On Error Resume Next
    If target.MergeCells Then target.MergeArea.UnMerge
    On Error GoTo 0
End Sub

Private Function R1C1Ref(ByVal colOffset As Long) As String
    If colOffset = 0 Then
        R1C1Ref = "RC"
    Else
        R1C1Ref = "RC[" & CStr(colOffset) & "]"
    End If
End Function

Private Function PutSimpleFormulaR1C1(target As Range, ByVal baseCol As Long) As Boolean
    On Error GoTo Fail
    Dim baseOffset As Long
    baseOffset = baseCol - target.Column
    target.formulaR1C1 = "=" & R1C1Ref(baseOffset)
    PutSimpleFormulaR1C1 = True
    Exit Function
Fail:
    PutSimpleFormulaR1C1 = False
End Function

Private Function PutAverageFormulaR1C1(target As Range, _
                                       ByVal baseCol As Long, _
                                       ByVal analogStartCol As Long, _
                                       ByVal analogEndCol As Long) As Boolean
    On Error GoTo Fail

    Dim baseRef As String
    Dim analogStartRef As String
    Dim analogEndRef As String
    Dim formulaR1C1 As String

    baseRef = R1C1Ref(baseCol - target.Column)
    analogStartRef = R1C1Ref(analogStartCol - target.Column)
    analogEndRef = R1C1Ref(analogEndCol - target.Column)

    formulaR1C1 = "=MAX(" & baseRef & ",IFERROR(AVERAGE(" & _
                  baseRef & "," & analogStartRef & ":" & analogEndRef & ")," & _
                  baseRef & "))"

    target.formulaR1C1 = formulaR1C1
    PutAverageFormulaR1C1 = True
    Exit Function

Fail:
    PutAverageFormulaR1C1 = False
End Function

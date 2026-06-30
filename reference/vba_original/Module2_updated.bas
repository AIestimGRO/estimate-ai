Attribute VB_Name = "Module2"
' ============================================================
'  MODULE 1: GLAVNYY - TOCHKA VHODA v4.1
'  Ispravleniya v4:
'  - zakrytie knig cherez SafeCloseWorkbook, bez oshibki pri povtornom Close
'  - posle zakrytiya wbCat/wbSm ssylka srazu obnulyaetsya
'  - v soobshenie ob oshibke dobavlen etap, gde ona voznikla
'  - dobavleno otklyuchenie EnableEvents na vremya raboty
' ============================================================
Option Explicit

Public Const DEDUP_PCT As Double = 0.04

Public Sub RunAnalogSearch()

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation

    Dim stage As String
    stage = "Start"

    ' --- Shag 0: Zagruzit' nastrojki iz lista Instrument ---
    Call LoadSettings
    If Not ValidateSettingsForRun() Then Exit Sub

    ' Reload name exclusion rules on every run, because user may edit the sheet.
    Call ResetNameExclusionRules

    ' --- Zagolovki dialogov ---
    Dim s1 As String, s2 As String
    Dim s3 As String, s4 As String, s5 As String
    s1 = ChrW(1042) & ChrW(1099) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1077)
    s2 = ChrW(1092) & ChrW(1072) & ChrW(1081) & ChrW(1083)
    s3 = ChrW(1050) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1086) & ChrW(1075)
    s4 = ChrW(1056) & ChrW(1053) & ChrW(1052) & ChrW(1062)
    s5 = ChrW(1086) & ChrW(1073) & ChrW(1088) & ChrW(1072) & ChrW(1073) & ChrW(1072) & ChrW(1090)
    s5 = s5 & ChrW(1099) & ChrW(1074) & ChrW(1072) & ChrW(1077) & ChrW(1084) & ChrW(1099) & ChrW(1081)
    Dim t1 As String: t1 = s1 & " " & s2 & " - " & s3 & " " & s4
    Dim t2 As String: t2 = s1 & " " & s5 & " " & s2

    ' --- Shag 1: Vybrat' fajly ---
    Dim f1Path As String, smPath As String
    f1Path = PickFile(t1)
    If f1Path = "" Then MsgBox "Otmena.", vbExclamation: Exit Sub

    smPath = PickFile(t2)
    If smPath = "" Then MsgBox "Otmena.", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim t0 As Double: t0 = Timer
    Dim wbCat As Workbook, wbSm As Workbook

    On Error GoTo ErrHandler

    ' --- Shag 2: Otkryt' fajly ---
    stage = "Otkrytie kataloga RNMC"
    Application.StatusBar = "[1/4] Zagruzhayu katalog..."
    Set wbCat = Workbooks.Open(f1Path, ReadOnly:=True, UpdateLinks:=False)

    stage = "Otkrytie obrabatyvaemogo fajla"
    Set wbSm = Workbooks.Open(smPath, ReadOnly:=False, UpdateLinks:=False)

    ' Najti list smety (soderzhit "OS" v nazvanii). Esli ne najden - pervyj list.
    stage = "Poisk lista smety"
    Dim wsSm As Worksheet
    Dim wsName As String
    wsName = FindSheetPart(wbSm, ChrW(1054) & ChrW(1057))
    If wsName = "" Then wsName = wbSm.Worksheets(1).Name
    Set wsSm = wbSm.Worksheets(wsName)

    ' --- Shag 3: Opredelit' strukturu smetnoj tablicy ---
    stage = "Opredelenie struktury smety"
    Dim headerRow As Long, dataStart As Long
    Dim analogStartCol As Long, sectionCol As Long
    If Not DetectLayout(wsSm, headerRow, dataStart, analogStartCol, sectionCol) Then
        MsgBox "Struktura lista smety ne opredelena." & vbCrLf & _
               "Zagolovok 'GESN/FER/Perechen' ne najden v stolbce " & _
               gSettings.colSearch & " (" & ColNumToLetter(gSettings.colSearch) & ")." & vbCrLf & _
               "Proverite nastrojku 'Stolbec GESN/FER' na liste Instrument.", vbCritical
        SafeCloseWorkbook wbCat, False
        SafeCloseWorkbook wbSm, False
        GoTo Cleanup
    End If

    ' --- Shag 4: Postroit' katalog ---
    stage = "Chtenie i sborka kataloga"
    Application.StatusBar = "[2/4] Chitayu katalog..."
    Dim catalog As Object
    Set catalog = BuildCatalog(wbCat)
    SafeCloseWorkbook wbCat, False

    If catalog.Count = 0 Then
        MsgBox "Katalog pust. Proverite fajl F1 i nastrojki stolbcov kataloga na liste Instrument.", vbCritical
        SafeCloseWorkbook wbSm, False
        GoTo Cleanup
    End If

    ' --- Shag 5: Podobrat' analogi ---
    stage = "Podbor i vyvod analogov"
    Application.StatusBar = "[3/4] Podbirayu analogi..."
    Dim logLines() As String
    ProcessSmeta wsSm, catalog, headerRow, dataStart, analogStartCol, sectionCol, logLines

    ' --- Shag 6: Log i sohranenie ---
    stage = "Zapis' loga"
    Application.StatusBar = "[4/4] Sohranyayu log..."
    WriteLog logLines, smPath, Timer - t0

    stage = "Sohranenie logov v fajle makrosa"
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    On Error GoTo ErrHandler

    stage = "Sohranenie obrabatyvaemogo fajla"
    Application.StatusBar = "[4/4] Sohranyayu fajl..."
    wbSm.Save

    stage = "Zakrytie obrabatyvaemogo fajla"
    SafeCloseWorkbook wbSm, False

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents

    If Err.Number = 0 Then
        MsgBox "Gotovo! " & Format(Timer - t0, "0.0") & " sek." & vbCrLf & _
               "Podrobnosti na liste Log.", vbInformation
    End If
    Exit Sub

ErrHandler:
    Dim en As Long: en = Err.Number
    Dim ed As String: ed = Err.Description

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents

    MsgBox "Oshibka #" & en & ": " & ed & vbCrLf & _
           "Etap: " & stage, vbCritical

    SafeCloseWorkbook wbCat, False
    SafeCloseWorkbook wbSm, False
End Sub

Public Function PickFile(title As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = title
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Public Function FindSheetPart(wb As Workbook, part As String) As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, part, vbTextCompare) > 0 Then
            FindSheetPart = ws.Name: Exit Function
        End If
    Next ws
End Function

Public Sub SafeCloseWorkbook(ByRef wb As Workbook, Optional ByVal saveChanges As Boolean = False)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close saveChanges:=saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub

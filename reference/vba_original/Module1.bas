Attribute VB_Name = "Module1"
' ============================================================
'  MODULE 0: NASTROJKI STOLBCOV v3.5
'  Vse vazhnye stolbcy zadayutsya na liste "Instrument".
'  Stolbec C - znachenie, stolbec D - bukva stolbca/esli primenimo.
'
'  Pervyj zapusk: InitSettingsBlock
'  Osnovnoj zapusk: RunAnalogSearch
' ============================================================
Option Explicit

Public Type TSettings
    ' --- obrabatyvaemyj fajl smety ---
    colSearch        As Long   ' stolbec s GESN/FER/Perechnem v smete
    colAvg           As Long   ' stolbec itogovoj / srednej ceny v smete
    colKR            As Long   ' stolbec, kuda dobavlyaetsya /KR
    colF             As Long   ' stolbec bazovoj ceny v smete
    colSection       As Long   ' stolbec koda razdela v smete
    colAnalogStart   As Long   ' pervyj stolbec dlya analogov v smete
    colSmetaWorkName As Long   ' stolbec naimenovaniya rabot v smete dlya filtra demontazha
    colSmetaUnit     As Long   ' stolbec edinicy izmereniya v smete dlya tochnogo poiska

    ' --- katalog RNMC ---
    catTaskCol       As Long   ' katalog: nomer zadachi
    catPriceCol      As Long   ' katalog: cena edinicy rabot bez NDS
    catCodeCol       As Long   ' katalog: GESN/FER/Perechen
    catRegionCol     As Long   ' katalog: region
    catWorkNameCol   As Long   ' katalog: naimenovanie rabot dlya filtra demontazha
    catUnitCol       As Long   ' katalog: edinica izmereniya dlya tochnogo poiska
    catAddedDateCol  As Long   ' katalog: data dobavleniya stroki v katalog

    ' --- proverka cen analogov ---
    priceSpreadLimit As Double ' porog MaxPrice / MinPrice, naprimer 2

    ' --- regionalnyj koefficient ---
    regionalCoefCellAddress As String ' adres yachejki s regionalnym koefficientom v liste smety, naprimer F12

    ' --- filtr demontazha ---
    demontazhFilterEnabled As Boolean ' True = filtr demontazha vklyuchen, False = otklyuchen
End Type

Public gSettings As TSettings

Private Const SETTINGS_SHEET As String = "Instrument"
Private Const ROW_FIRST      As Long = 34
Private Const COL_VAL        As Long = 3

' --- Znacheniya po umolchaniyu ---
Public Sub SetDefaultSettings()
    ' Smeta / obrabatyvaemyj fajl
    gSettings.colSearch = 14          ' N
    gSettings.colAvg = 7              ' G
    gSettings.colKR = 14              ' N
    gSettings.colF = 6                ' F
    gSettings.colSection = 15         ' O
    gSettings.colAnalogStart = 16     ' P
    gSettings.colSmetaWorkName = 3    ' C
    gSettings.colSmetaUnit = 4        ' D

    ' Katalog RNMC
    gSettings.catTaskCol = 2          ' B
    gSettings.catPriceCol = 7         ' G
    gSettings.catCodeCol = 14         ' N
    gSettings.catRegionCol = 16       ' P
    gSettings.catWorkNameCol = 3      ' C
    gSettings.catUnitCol = 4          ' D
    gSettings.catAddedDateCol = 17    ' Q

    ' Proverka razbrosa cen analogov
    gSettings.priceSpreadLimit = 2

    ' Regionalnyj koefficient
    ' Pustoe znachenie = koefficient 1, chtoby starye shablony ne lomalis'.
    gSettings.regionalCoefCellAddress = ""

    ' Filtr montazh/demontazh po umolchaniyu vklyuchen, chtoby staroe povedenie ne izmenilos'.
    gSettings.demontazhFilterEnabled = True
End Sub

Private Function ReadSettingLong(ws As Worksheet, rw As Long, defaultValue As Long, ByRef badCount As Long) As Long
    Dim v As Variant
    v = ws.Cells(rw, COL_VAL).Value

    If IsEmpty(v) Or Trim(CStr(v)) = "" Then
        ReadSettingLong = defaultValue
    ElseIf IsNumeric(v) And CLng(v) >= 1 Then
        ReadSettingLong = CLng(v)
    Else
        badCount = badCount + 1
        ReadSettingLong = defaultValue
    End If
End Function

Private Function ReadSettingDouble(ws As Worksheet, rw As Long, defaultValue As Double, ByRef badCount As Long) As Double
    Dim v As Variant
    v = ws.Cells(rw, COL_VAL).Value

    If IsEmpty(v) Or Trim(CStr(v)) = "" Then
        ReadSettingDouble = defaultValue
    ElseIf IsNumeric(v) And CDbl(v) > 0 Then
        ReadSettingDouble = CDbl(v)
    Else
        badCount = badCount + 1
        ReadSettingDouble = defaultValue
    End If
End Function

Private Function ReadSettingFlag01(ws As Worksheet, rw As Long, defaultValue As Boolean, ByRef badCount As Long) As Boolean
    Dim v As Variant
    v = ws.Cells(rw, COL_VAL).Value

    If IsEmpty(v) Or Trim(CStr(v)) = "" Then
        ReadSettingFlag01 = defaultValue
    ElseIf IsNumeric(v) Then
        If CDbl(v) = 0 Then
            ReadSettingFlag01 = False
        ElseIf CDbl(v) = 1 Then
            ReadSettingFlag01 = True
        Else
            badCount = badCount + 1
            ReadSettingFlag01 = defaultValue
        End If
    Else
        badCount = badCount + 1
        ReadSettingFlag01 = defaultValue
    End If
End Function

Private Function ReadSettingText(ws As Worksheet, rw As Long, defaultValue As String) As String
    Dim v As Variant
    v = ws.Cells(rw, COL_VAL).Value

    If IsEmpty(v) Then
        ReadSettingText = defaultValue
    Else
        ReadSettingText = Trim(CStr(v))
        If ReadSettingText = "" Then ReadSettingText = defaultValue
    End If
End Function

' --- Prochitat' nastrojki iz lista Instrument ---
Public Function LoadSettings() As Boolean
    Call SetDefaultSettings

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        LoadSettings = True
        Exit Function
    End If

    Dim badCount As Long: badCount = 0

    ' Staruyu raskladku strok ne lomayem: stroki 34-43 ostayutsya kak byli.
    gSettings.colSearch = ReadSettingLong(ws, ROW_FIRST + 0, gSettings.colSearch, badCount)
    gSettings.colAvg = ReadSettingLong(ws, ROW_FIRST + 1, gSettings.colAvg, badCount)
    gSettings.colKR = ReadSettingLong(ws, ROW_FIRST + 2, gSettings.colKR, badCount)
    gSettings.colF = ReadSettingLong(ws, ROW_FIRST + 3, gSettings.colF, badCount)
    gSettings.colSection = ReadSettingLong(ws, ROW_FIRST + 4, gSettings.colSection, badCount)
    gSettings.colAnalogStart = ReadSettingLong(ws, ROW_FIRST + 5, gSettings.colAnalogStart, badCount)

    gSettings.catTaskCol = ReadSettingLong(ws, ROW_FIRST + 6, gSettings.catTaskCol, badCount)
    gSettings.catPriceCol = ReadSettingLong(ws, ROW_FIRST + 7, gSettings.catPriceCol, badCount)
    gSettings.catCodeCol = ReadSettingLong(ws, ROW_FIRST + 8, gSettings.catCodeCol, badCount)
    gSettings.catRegionCol = ReadSettingLong(ws, ROW_FIRST + 9, gSettings.catRegionCol, badCount)

    ' Novye nastrojki dobavleny nizhe starogo bloka.
    gSettings.colSmetaWorkName = ReadSettingLong(ws, ROW_FIRST + 10, gSettings.colSmetaWorkName, badCount)
    gSettings.catWorkNameCol = ReadSettingLong(ws, ROW_FIRST + 11, gSettings.catWorkNameCol, badCount)
    gSettings.priceSpreadLimit = ReadSettingDouble(ws, ROW_FIRST + 12, gSettings.priceSpreadLimit, badCount)

    ' Adres yachejki s regionalnym koefficientom v obrabatyvaemom liste smety.
    ' Esli pustoj, v ProcessSmeta budet ispol'zovan koefficient 1.
    gSettings.regionalCoefCellAddress = ReadSettingText(ws, ROW_FIRST + 13, gSettings.regionalCoefCellAddress)

    ' Stolbcy edinic izmereniya dlya tochnogo poiska po svyazke Ed.Izm + Perechen.
    gSettings.colSmetaUnit = ReadSettingLong(ws, ROW_FIRST + 14, gSettings.colSmetaUnit, badCount)
    gSettings.catUnitCol = ReadSettingLong(ws, ROW_FIRST + 15, gSettings.catUnitCol, badCount)

    ' 1 = filtr demontazha vklyuchen, 0 = otklyuchen.
    gSettings.demontazhFilterEnabled = ReadSettingFlag01(ws, ROW_FIRST + 16, gSettings.demontazhFilterEnabled, badCount)

    ' Stolbec daty dobavleniya stroki v katalog RNMC.
    gSettings.catAddedDateCol = ReadSettingLong(ws, ROW_FIRST + 17, gSettings.catAddedDateCol, badCount)

    If badCount > 0 Then
        MsgBox "V bloke nastroek est' nekorrektnye znacheniya: " & badCount & "." & vbCrLf & _
               "Dlya nih ispol'zovany znacheniya po umolchaniyu.", vbExclamation
    End If

    LoadSettings = True
End Function

' --- Proverka kriticheskih nastroek pered zapuskom ---
Public Function ValidateSettingsForRun() As Boolean
    ValidateSettingsForRun = False

    If gSettings.colSearch < 1 Or gSettings.colAvg < 1 Or _
       gSettings.colKR < 1 Or gSettings.colF < 1 Or _
       gSettings.colSection < 1 Or gSettings.colAnalogStart < 1 Or _
       gSettings.colSmetaWorkName < 1 Or gSettings.colSmetaUnit < 1 Or _
       gSettings.catTaskCol < 1 Or gSettings.catPriceCol < 1 Or _
       gSettings.catCodeCol < 1 Or gSettings.catRegionCol < 1 Or _
       gSettings.catWorkNameCol < 1 Or gSettings.catUnitCol < 1 Or _
       gSettings.catAddedDateCol < 1 Then
        MsgBox "V nastrojkah stolbcov est' znachenie menshe 1." & vbCrLf & _
               "Proverite zheltye yachejki v bloke Instrument, stroki 34-51.", vbCritical
        Exit Function
    End If

    If gSettings.priceSpreadLimit <= 0 Then
        MsgBox "Nekorrektnyj porog proverki cen analogov." & vbCrLf & _
               "Na liste Instrument porog MaxPrice / MinPrice dolzhen byt' bolshe 0.", vbCritical
        Exit Function
    End If

    If Trim(gSettings.regionalCoefCellAddress) <> "" Then
        If Not IsValidA1CellAddress(gSettings.regionalCoefCellAddress) Then
            MsgBox "Nekorrektnyj adres yachejki regionalnogo koefficienta: " & _
                   gSettings.regionalCoefCellAddress & vbCrLf & _
                   "Ukazhite adres v formate A1, naprimer F12 ili $F$12.", vbCritical
            Exit Function
        End If
    End If

    ' Pervyj stolbec analogov dolzhen byt' pravee osnovnyh stolbcov smety,
    ' inache ochistka analogov mozhet steret ishodye dannye ili formulu srednej ceny.
    Dim maxMain As Long
    maxMain = gSettings.colSearch
    If gSettings.colAvg > maxMain Then maxMain = gSettings.colAvg
    If gSettings.colKR > maxMain Then maxMain = gSettings.colKR
    If gSettings.colF > maxMain Then maxMain = gSettings.colF
    If gSettings.colSection > maxMain Then maxMain = gSettings.colSection
    If gSettings.colSmetaWorkName > maxMain Then maxMain = gSettings.colSmetaWorkName
    If gSettings.colSmetaUnit > maxMain Then maxMain = gSettings.colSmetaUnit

    If gSettings.colAnalogStart <= maxMain Then
        MsgBox "Nekorrektnaya nastrojka: pervyj stolbec analogov = " & _
               gSettings.colAnalogStart & " (" & ColNumToLetter(gSettings.colAnalogStart) & ")." & vbCrLf & _
               "On dolzhen byt' pravee osnovnyh stolbcov smety." & vbCrLf & _
               "Minimal'no dopustimo: " & (maxMain + 1) & " (" & ColNumToLetter(maxMain + 1) & ").", vbCritical
        Exit Function
    End If

    ValidateSettingsForRun = True
End Function

Private Function IsValidA1CellAddress(ByVal addr As String) As Boolean
    On Error GoTo BadAddress

    addr = Replace(Trim(addr), "$", "")
    If addr = "" Then
        IsValidA1CellAddress = False
        Exit Function
    End If

    Dim tmp As Range
    Set tmp = ThisWorkbook.Worksheets(SETTINGS_SHEET).Range(addr)
    IsValidA1CellAddress = (tmp.Cells.CountLarge = 1)
    Exit Function

BadAddress:
    IsValidA1CellAddress = False
End Function

' --- Sozdat' blok nastroek na liste Instrument ---
Public Sub InitSettingsBlock()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "List '" & SETTINGS_SHEET & "' ne najden.", vbExclamation
        Exit Sub
    End If

    ' Vazhno: ne sbrosit' uzhe nastroyennye pol'zovatelem znacheniya.
    ' Esli blok nastroek uzhe byl, zdes' sohranyaem tekushchie zheltye yachejki
    ' i nizhe vozvrashchaem ih obratno vmesto znachenij po umolchaniyu.
    Dim oldVals(1 To 18) As Variant
    Dim oldI As Long
    On Error Resume Next
    For oldI = 1 To 18
        oldVals(oldI) = ws.Cells(ROW_FIRST + oldI - 1, COL_VAL).Value
    Next oldI
    On Error GoTo 0

    Call SetDefaultSettings

    Dim rw As Long
    ws.Rows(32).RowHeight = 8

    rw = 33
    ws.Range("B" & rw & ":D" & rw).UnMerge
    ws.Range("B" & rw & ":D" & rw).Merge
    With ws.Range("B" & rw)
        .Value = "NASTROJKA STOLBCOV I PROVEROK DLYA ANALIZA"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With

    Dim labels(1 To 18) As String
    Dim defaults(1 To 18) As Variant

    labels(1) = "SMETA: stolbec GESN/FER/Perechen dlya poiska"
    labels(2) = "SMETA: stolbec srednej/itogovoj ceny"
    labels(3) = "SMETA: stolbec dlya dobavleniya /KR"
    labels(4) = "SMETA: stolbec bazovoj ceny"
    labels(5) = "SMETA: stolbec koda razdela"
    labels(6) = "SMETA: pervyj stolbec dlya analogov"
    labels(7) = "KATALOG: stolbec nomera zadachi"
    labels(8) = "KATALOG: stolbec ceny bez NDS"
    labels(9) = "KATALOG: stolbec GESN/FER/Perechen"
    labels(10) = "KATALOG: stolbec regiona"
    labels(11) = "SMETA: stolbec naimenovaniya rabot"
    labels(12) = "KATALOG: stolbec naimenovaniya rabot"
    labels(13) = "PROVERKA: porog MaxPrice / MinPrice"
    labels(14) = "SMETA: adres yachejki regionalnogo koefficienta"
    labels(15) = "SMETA: stolbec edinicy izmereniya"
    labels(16) = "KATALOG: stolbec edinicy izmereniya"
    labels(17) = "PROVERKA: filtr demontazha vklyuchen (1-da, 0-net)"
    labels(18) = "KATALOG: stolbec daty dobavleniya stroki"

    defaults(1) = gSettings.colSearch
    defaults(2) = gSettings.colAvg
    defaults(3) = gSettings.colKR
    defaults(4) = gSettings.colF
    defaults(5) = gSettings.colSection
    defaults(6) = gSettings.colAnalogStart
    defaults(7) = gSettings.catTaskCol
    defaults(8) = gSettings.catPriceCol
    defaults(9) = gSettings.catCodeCol
    defaults(10) = gSettings.catRegionCol
    defaults(11) = gSettings.colSmetaWorkName
    defaults(12) = gSettings.catWorkNameCol
    defaults(13) = gSettings.priceSpreadLimit
    defaults(14) = gSettings.regionalCoefCellAddress
    defaults(15) = gSettings.colSmetaUnit
    defaults(16) = gSettings.catUnitCol
    defaults(17) = IIf(gSettings.demontazhFilterEnabled, 1, 0)
    defaults(18) = gSettings.catAddedDateCol

    Dim fillEven As Long: fillEven = RGB(235, 241, 250)
    Dim fillOdd  As Long: fillOdd = RGB(255, 255, 255)
    Dim fillEdit As Long: fillEdit = RGB(255, 255, 204)

    Dim i As Long
    For i = 1 To 18
        rw = ROW_FIRST + i - 1
        ws.Range("B" & rw & ":D" & rw).UnMerge
        ws.Rows(rw).RowHeight = 20

        ws.Cells(rw, 2).Value = labels(i)
        ws.Cells(rw, 2).Font.Size = 9
        ws.Cells(rw, 2).Interior.Color = IIf(i Mod 2 = 0, fillEven, fillOdd)

        If Trim(CStr(oldVals(i))) <> "" Then
            ws.Cells(rw, COL_VAL).Value = oldVals(i)
        Else
            ws.Cells(rw, COL_VAL).Value = defaults(i)
        End If
        ws.Cells(rw, COL_VAL).Font.Bold = True
        ws.Cells(rw, COL_VAL).Font.Size = 10
        ws.Cells(rw, COL_VAL).Interior.Color = fillEdit
        ws.Cells(rw, COL_VAL).HorizontalAlignment = xlCenter
        If i = 13 Then
            ws.Cells(rw, COL_VAL).NumberFormat = "0.00"
        ElseIf i = 14 Then
            ws.Cells(rw, COL_VAL).NumberFormat = "@"
            ws.Cells(rw, COL_VAL).HorizontalAlignment = xlLeft
        Else
            ws.Cells(rw, COL_VAL).NumberFormat = "0"
        End If

        If i <= 12 Or i = 15 Or i = 16 Or i = 18 Then
            ws.Cells(rw, 4).Formula = "=IF(C" & rw & ">0," & _
                "SUBSTITUTE(ADDRESS(1,C" & rw & ",4),""1"",""""),"""")"
        ElseIf i = 14 Then
            ws.Cells(rw, 4).Value = "adres A1"
        ElseIf i = 17 Then
            ws.Cells(rw, 4).Value = "1=vkl, 0=vykl"
        Else
            ws.Cells(rw, 4).Value = "-"
        End If
        ws.Cells(rw, 4).Font.Size = 9
        ws.Cells(rw, 4).Font.Color = RGB(100, 100, 100)
        ws.Cells(rw, 4).HorizontalAlignment = xlCenter
        ws.Cells(rw, 4).Interior.Color = IIf(i Mod 2 = 0, fillEven, fillOdd)

        With ws.Range(ws.Cells(rw, 2), ws.Cells(rw, 4)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
    Next i

    rw = ROW_FIRST + 18
    ws.Range("B" & rw & ":D" & rw).UnMerge
    ws.Range("B" & rw & ":D" & rw).Merge
    With ws.Range("B" & rw)
        .Value = "* Menyayte znacheniya v zheltyh yachejkah. Dlya koefficienta ukazhite adres yachejki, naprimer F12. Ed. izm. uchityvaetsya v klyuche poiska. Filtr demontazha: 1 = vklyuchen, 0 = otklyuchen. Stolbec daty kataloga nuzhen dlya umnyh isklyuchenij GESN."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
    End With

    MsgBox "Blok nastroek gotov na liste Instrument. Tekushchie znacheniya sohraneny, novye stroki dobavleny.", vbInformation
End Sub

' --- Bukva stolbca po nomeru ---
Public Function ColNumToLetter(n As Long) As String
    Dim s As String: s = ""
    Do While n > 0
        s = Chr(64 + ((n - 1) Mod 26 + 1)) & s
        n = (n - 1) \ 26
    Loop
    ColNumToLetter = s
End Function

Attribute VB_Name = "Module3"
' ============================================================
'  MODULE 2: KATALOG v3.2
'  - DetectLayout    - opredelit' strukturu lista smety
'  - NormCode        - normalizaciya koda GESN
'  - GESnPrefix      - izvlech' prefiks dlya slovarya
'  - BuildSectionDict- slovar' razdelov EKR-GESN
'  - BuildCatalog    - chitat' katalog F1 s deduplicaciej
' ============================================================
Option Explicit

' --- Normalizaciya koda GESN ----------------------------------
Public Function NormCode(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        NormCode = ""
        Exit Function
    End If

    Dim s As String
    s = CStr(v)

    ' Normalize separators before trimming. Some catalog cells contain code + line break + /KR.
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")

    s = UCase$(Trim$(s))

    If s = "" Then
        NormCode = ""
        Exit Function
    End If

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Do While InStr(s, " /") > 0
        s = Replace(s, " /", "/")
    Loop

    Do While InStr(s, "/ ") > 0
        s = Replace(s, "/ ", "/")
    Loop

    ' Remove /KR suffix before catalog search.
    Dim KR As String
    KR = "/" & ChrW(1050) & ChrW(1056)

    If Right$(s, Len(KR)) = KR Then
        s = Left$(s, Len(s) - Len(KR))
    End If

    NormCode = Trim$(s)
End Function

' --- Normalizaciya edinicy izmereniya dlya tochnogo poiska -----
Public Function NormUnit(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or v = "" Then
        NormUnit = ""
        Exit Function
    End If

    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(1105), ChrW(1077))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, ChrW(178), "2")
    s = Replace(s, ChrW(179), "3")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    ' U edinic izmereniya chasto otlichayutsya tolko probely/tochki.
    ' Dlya kljucha poiska delaem kompaktnyj variant.
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    s = Replace(s, "^", "")

    NormUnit = s
End Function

' --- Klyuch poiska analogov: edinica izmereniya + perechen GESN ---
Public Function AnalogSearchKey(unitVal As Variant, codeVal As Variant) As String
    Dim u As String
    Dim c As String

    u = NormUnit(unitVal)
    c = NormCode(codeVal)

    ' Esli net edinicy ili koda, ne sozdaem klyuch, chtoby ne bylo sluchajnyh sovpadenij.
    If u = "" Or c = "" Then
        AnalogSearchKey = ""
    Else
        AnalogSearchKey = u & "||" & c
    End If
End Function

Public Function HasDemontazh(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        HasDemontazh = False
        Exit Function
    End If

    Dim s As String
    s = LCase$(CStr(v))
    s = Replace(s, ChrW(1105), ChrW(1077))
    s = Replace(s, Chr(160), " ")

    Dim ch As Variant
    For Each ch In Array(".", ",", ";", ":", "(", ")", "[", "]", "{", "}", "/", "\", "-", "_", "+", "=", vbTab, vbCr, vbLf)
        s = Replace(s, CStr(ch), " ")
    Next ch

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Dim demRoot As String
    demRoot = ChrW(1076) & ChrW(1077) & ChrW(1084) & ChrW(1086) & ChrW(1085) & ChrW(1090)

    Dim parts() As String
    parts = Split(Trim$(s), " ")

    Dim i As Long
    Dim word As String

    For i = LBound(parts) To UBound(parts)
        word = parts(i)

        ' Lovim tolko koren demont. Obychnoe slovo montazh ne podhodit.
        If Left$(word, Len(demRoot)) = demRoot Then
            HasDemontazh = True
            Exit Function
        End If
    Next i

    HasDemontazh = False
End Function

' --- Prefiks GESN (napr "GESN26" iz "GESN26-01-003-01") -------
Public Function GESnPrefix(code As String) As String
    Dim s As String: s = UCase(code)
    Dim G As String: G = ChrW(1043) & ChrW(1069) & ChrW(1057) & ChrW(1053)
    Dim pos As Long: pos = InStr(s, G)
    If pos = 0 Then GESnPrefix = "": Exit Function
    Dim tail As String: tail = Mid(s, pos)
    Dim prefix As String: prefix = G
    Dim idx As Long: idx = 5
    ' Proverit' optsional'nuyu buqvu posle GESN (M, P, R = M, P, R)
    If idx <= Len(tail) Then
        Dim ch As String: ch = Mid(tail, idx, 1)
        Dim ac As Integer: ac = Asc(ch)
        If (ac >= 192 And ac <= 223) Or ac = 168 Then
            prefix = prefix & ch: idx = idx + 1
        End If
    End If
    ' Dve cyfry
    If idx + 1 <= Len(tail) Then
        Dim d1 As String: d1 = Mid(tail, idx, 1)
        Dim d2 As String: d2 = Mid(tail, idx + 1, 1)
        If d1 >= "0" And d1 <= "9" And d2 >= "0" And d2 <= "9" Then
            prefix = prefix & d1 & d2
        End If
    End If
    GESnPrefix = prefix
End Function

' --- Opredelit' strukturu lista smety ------------------------
Public Function DetectLayout(ws As Worksheet, _
                              ByRef headerRow As Long, _
                              ByRef dataStart As Long, _
                              ByRef analogStartCol As Long, _
                              ByRef sectionCol As Long) As Boolean
    Dim r As Long, c As Long
    ' Iskuem stroku shapki po stolbcu s GESN/FER/Perechnem iz nastroek
    Dim gesn As String: gesn = ChrW(1043) & ChrW(1069) & ChrW(1057) & ChrW(1053)
    Dim fer  As String: fer = ChrW(1060) & ChrW(1045) & ChrW(1056)
    Dim per  As String: per = ChrW(1055) & ChrW(1077) & ChrW(1088)

    For r = 1 To 50
        Dim hvSearch As String: hvSearch = CStr(ws.Cells(r, gSettings.colSearch).Value)
        If InStr(1, hvSearch, gesn, vbTextCompare) > 0 Or _
           InStr(1, hvSearch, fer, vbTextCompare) > 0 Or _
           InStr(1, hvSearch, per, vbTextCompare) > 0 Then

            headerRow = r
            dataStart = r + 2

            ' Kolonki, vazhnye dlya analiza, berem iz nastroek.
            ' Esli po kakoj-to prichine analogStart ne zadan, nachinaem posle koda razdela.
            sectionCol = gSettings.colSection
            If gSettings.colAnalogStart > 0 Then
                analogStartCol = gSettings.colAnalogStart
            Else
                analogStartCol = sectionCol + 1
            End If

            DetectLayout = True
            Exit Function
        End If
    Next r

    DetectLayout = False
End Function

' --- Slovar' razdelov EKR-GESN --------------------------------

Private Function CatalogDateSerial(ByVal v As Variant) As Double
    On Error GoTo BadDate

    CatalogDateSerial = 0
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Trim(CStr(v)) = "" Then Exit Function

    If IsDate(v) Then
        CatalogDateSerial = CDbl(DateValue(CDate(v)))
    End If
    Exit Function

BadDate:
    CatalogDateSerial = 0
End Function

Public Function BuildSectionDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim G As String: G = ChrW(1043) & ChrW(1069) & ChrW(1057) & ChrW(1053)
    Dim M As String: M = G & ChrW(1052)
    Dim P As String: P = G & ChrW(1055)
    Dim r As String: r = G & ChrW(1056)
    d(G & "01") = "01"
    d(G & "04") = "01"
    d(G & "05") = "02"
    d(G & "06") = "02"
    d(G & "07") = "02"
    d(G & "08") = "03"
    d(G & "10") = "03"
    d(G & "11") = "03"
    d(G & "12") = "03"
    d(G & "15") = "03"
    d(G & "09") = "04"
    d(G & "39") = "04"
    d(M & "38") = "04"
    d(G & "27") = "05"
    d(G & "28") = "05"
    d(G & "47") = "05"
    d(M & "20") = "05"
    d(M & "03") = "06"
    d(M & "06") = "06"
    d(M & "07") = "06"
    d(M & "13") = "06"
    d(M & "18") = "06"
    d(M & "19") = "06"
    d(M & "22") = "06"
    d(M & "37") = "06"
    d(P & "07") = "06"
    d(G & "13") = "07"
    d(G & "26") = "07"
    d(G & "45") = "07"
    d(G & "46") = "08"
    d(r & "67") = "08"
    d(r & "51") = "09"
    d(r & "52") = "09"
    d(r & "53") = "09"
    d(r & "55") = "09"
    d(r & "61") = "09"
    d(r & "63") = "09"
    d(r & "65") = "09"
    d(r & "66") = "09"
    d(r & "68") = "09"
    d(r & "69") = "09"
    d(G & "16") = "10"
    d(G & "17") = "10"
    d(G & "18") = "10"
    d(G & "22") = "11"
    d(G & "23") = "11"
    d(G & "24") = "11"
    d(G & "25") = "11"
    d(M & "12") = "11"
    d(G & "20") = "12"
    d(P & "03") = "12"
    d(M & "39") = "13"
    d(G & "34") = "14"
    d(M & "10") = "14"
    d(M & "11") = "15"
    d(G & "33") = "16"
    d(M & "08") = "16"
    d(P & "01") = "16"
    Set BuildSectionDict = d
End Function

Public Function ResolveSectionCode(ByVal normCode As String, _
                                   ByVal workName As Variant, _
                                   ByVal secDict As Object) As String
    ResolveSectionCode = ""

    Dim pfx As String
    pfx = GESnPrefix(normCode)
    If pfx = "" Then Exit Function

    If IsSection08PriorityPrefix(pfx) Then
        If HasDemontazh(workName) Then
            ResolveSectionCode = "08"
        Else
            ResolveSectionCode = PreferredNonDemSectionForPrefix(pfx, secDict)
            If ResolveSectionCode = "" Then ResolveSectionCode = "08"
        End If
    Else
        If Not secDict Is Nothing Then
            If secDict.Exists(pfx) Then ResolveSectionCode = CStr(secDict(pfx))
        End If
    End If
End Function

Private Function IsSection08PriorityPrefix(ByVal pfx As String) As Boolean
    Dim G As String
    Dim R As String
    G = ChrW(1043) & ChrW(1069) & ChrW(1057) & ChrW(1053)
    R = G & ChrW(1056)

    Select Case UCase$(Trim(CStr(pfx)))
        Case G & "09", G & "27", G & "28", G & "46", R & "67"
            IsSection08PriorityPrefix = True
        Case Else
            IsSection08PriorityPrefix = False
    End Select
End Function

Private Function PreferredNonDemSectionForPrefix(ByVal pfx As String, ByVal secDict As Object) As String
    PreferredNonDemSectionForPrefix = ""

    If secDict Is Nothing Then Exit Function
    If Not secDict.Exists(pfx) Then Exit Function

    Dim sectionCode As String
    sectionCode = Trim(CStr(secDict(pfx)))

    If sectionCode <> "" And sectionCode <> "08" Then
        PreferredNonDemSectionForPrefix = sectionCode
    End If
End Function

' --- Prochitat' katalog F1 i postroit' slovar' s dedup 4% -----
Public Function BuildCatalog(wbCat As Workbook) As Object
    Dim cat As Object
    Set cat = CreateObject("Scripting.Dictionary")
    cat.CompareMode = 1

    ' Najti list "Katalog"
    Dim wsName As String
    wsName = FindSheetPart(wbCat, ChrW(1050) & ChrW(1072) & ChrW(1090))
    If wsName = "" Then wsName = wbCat.Worksheets(1).Name
    Dim wsCat As Worksheet: Set wsCat = wbCat.Worksheets(wsName)

    Dim lastRow As Long
    lastRow = wsCat.Cells(wsCat.Rows.Count, gSettings.catTaskCol).End(xlUp).Row
    If wsCat.Cells(wsCat.Rows.Count, gSettings.catCodeCol).End(xlUp).Row > lastRow Then _
        lastRow = wsCat.Cells(wsCat.Rows.Count, gSettings.catCodeCol).End(xlUp).Row
    If wsCat.Cells(wsCat.Rows.Count, gSettings.catPriceCol).End(xlUp).Row > lastRow Then _
        lastRow = wsCat.Cells(wsCat.Rows.Count, gSettings.catPriceCol).End(xlUp).Row
    If wsCat.Cells(wsCat.Rows.Count, gSettings.catUnitCol).End(xlUp).Row > lastRow Then _
        lastRow = wsCat.Cells(wsCat.Rows.Count, gSettings.catUnitCol).End(xlUp).Row
    If lastRow < 5 Then Set BuildCatalog = cat: Exit Function

    ' Chitaem ne tolko vazhnye stolbcy, a vsyu ispolzuemuyu shirinu kataloga,
    ' chtoby potom skopirovat polnuyu ishodyuyu stroku v Price_Check_Log.
    Dim maxCol As Long
    Dim lastUsedCell As Range
    On Error Resume Next
    Set lastUsedCell = wsCat.Cells.Find(What:="*", _
                                        After:=wsCat.Cells(1, 1), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlPrevious, _
                                        MatchCase:=False)
    On Error GoTo 0

    If lastUsedCell Is Nothing Then
        maxCol = gSettings.catRegionCol
    Else
        maxCol = lastUsedCell.Column
    End If

    If gSettings.catTaskCol > maxCol Then maxCol = gSettings.catTaskCol
    If gSettings.catPriceCol > maxCol Then maxCol = gSettings.catPriceCol
    If gSettings.catCodeCol > maxCol Then maxCol = gSettings.catCodeCol
    If gSettings.catRegionCol > maxCol Then maxCol = gSettings.catRegionCol
    If gSettings.catWorkNameCol > maxCol Then maxCol = gSettings.catWorkNameCol
    If gSettings.catUnitCol > maxCol Then maxCol = gSettings.catUnitCol
    If gSettings.catAddedDateCol > maxCol Then maxCol = gSettings.catAddedDateCol

    Dim arr As Variant
    arr = wsCat.Range(wsCat.Cells(4, 1), wsCat.Cells(lastRow, maxCol)).Value

    Dim i As Long
    For i = 1 To UBound(arr, 1)
        ' KATALOG: nomer zadachi
        Dim taskId As String
        taskId = Trim(CStr(arr(i, gSettings.catTaskCol)))
        If taskId = "" Then GoTo NextRow

        ' KATALOG: task color-list is not a stop-list anymore.
        ' Rows from marked tasks remain available as analogs.

        ' KATALOG: cena bez NDS
        If IsEmpty(arr(i, gSettings.catPriceCol)) Or _
           Not IsNumeric(arr(i, gSettings.catPriceCol)) Then GoTo NextRow
        Dim price As Double: price = CDbl(arr(i, gSettings.catPriceCol))
        If price <= 0 Then GoTo NextRow

        ' KATALOG: kod GESN/FER/Perechen + edinica izmereniya
        Dim normC As String
        normC = NormCode(arr(i, gSettings.catCodeCol))
        If normC = "" Then GoTo NextRow

        Dim normU As String
        normU = NormUnit(arr(i, gSettings.catUnitCol))
        If normU = "" Then GoTo NextRow

        Dim searchKey As String
        searchKey = AnalogSearchKey(arr(i, gSettings.catUnitCol), arr(i, gSettings.catCodeCol))
        If searchKey = "" Then GoTo NextRow

        ' KATALOG: region
        Dim region As String
        region = Trim(CStr(arr(i, gSettings.catRegionCol)))

        ' KATALOG: exclude analog rows by risky work-name phrases.
        If IsNameExcluded(arr(i, gSettings.catWorkNameCol), "CATALOG") Then GoTo NextRow
        
        ' KATALOG: priznak demontazha
        Dim isDem As Boolean
        isDem = HasDemontazh(arr(i, gSettings.catWorkNameCol))

        ' Data dobavleniya stroki v katalog. Esli pustaya/ne data, hranim 0.
        Dim addedDateSerial As Double
        addedDateSerial = CatalogDateSerial(arr(i, gSettings.catAddedDateCol))

        ' Polnaya kopiya ishodnoj stroki kataloga dlya loga proverki cen.
        Dim rowCopy() As Variant
        ReDim rowCopy(1 To maxCol)

        Dim cc As Long
        For cc = 1 To maxCol
            rowCopy(cc) = arr(i, cc)
        Next cc

        ' Struktura kataloga:
        ' cat(edinica_izmereniya + kod)(zadacha) = Collection
        ' element Collection = Array(cena, region, isDem, catalogRowNumber, rowCopy, taskId, normC, normU, addedDateSerial)
        If Not cat.Exists(searchKey) Then
            cat.Add searchKey, CreateObject("Scripting.Dictionary")
            cat(searchKey).CompareMode = 1
        End If

        If Not cat(searchKey).Exists(taskId) Then
            Dim entries As Collection
            Set entries = New Collection
            cat(searchKey).Add taskId, entries
        End If

        cat(searchKey)(taskId).Add Array(price, region, isDem, CLng(i + 3), rowCopy, taskId, normC, normU, addedDateSerial)
        
NextRow:
    Next i

    ' Deduplicaciya 4% vnutri svyazki edinica izmereniya + kod + zadacha.
    ' Esli cena otlichaetsya bolee chem na 4%, ona ostayotsya kak otdelnyj analog.
    Dim ck As Variant
    For Each ck In cat.Keys
        Dim tk As Variant
        For Each tk In cat(ck).Keys
            Dim srcEntries As Collection: Set srcEntries = cat(ck)(tk)
            Dim keptEntries As Collection: Set keptEntries = New Collection
            Dim ent As Variant
            For Each ent In srcEntries
                Dim ok As Boolean: ok = True
                Dim keptEnt As Variant
                For Each keptEnt In keptEntries
                    If CDbl(keptEnt(0)) > 0 Then
                        ' Ne skhlepyvaem demontazhnyj i nedemontazhnyj analog mezhdu soboj.
                        ' Inache pri blizkih cenah mozhno poteryat' nuzhnyj analog posle filtra demontazha.
                        If CBool(ent(2)) = CBool(keptEnt(2)) Then
                            If Abs(CDbl(ent(0)) - CDbl(keptEnt(0))) / CDbl(keptEnt(0)) <= DEDUP_PCT Then
                                ok = False
                                Exit For
                            End If
                        End If
                    End If
                Next keptEnt
                If ok Then keptEntries.Add ent
            Next ent
            Set cat(ck)(tk) = keptEntries
        Next tk
    Next ck

    Set BuildCatalog = cat
End Function

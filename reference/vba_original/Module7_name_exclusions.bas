Attribute VB_Name = "Module7"
' ============================================================
'  MODULE 7: EXCLUSION RULES v1.1
'  Risky work-name phrases and task color-list are stored on Excel sheet Name_Exclusions.
'  VBA keeps only ASCII technical names and builds default Russian text via ChrW.
' ============================================================
Option Explicit

Public Const NAME_EXCLUSIONS_SHEET As String = "Name_Exclusions"

Private gNameRules As Collection
Private gNameRulesLoaded As Boolean
Private gTaskColorList As Object

Public Sub ResetNameExclusionRules()
    Set gNameRules = Nothing
    Set gTaskColorList = Nothing
    gNameRulesLoaded = False
End Sub

Public Sub InitNameExclusionsSheet()
    Dim ws As Worksheet
    Set ws = EnsureNameExclusionsSheet(True)

    If ws Is Nothing Then
        MsgBox "Name_Exclusions sheet was not created.", vbExclamation
    Else
        MsgBox "Name_Exclusions sheet is ready.", vbInformation
    End If
End Sub

Public Function IsNameExcluded(ByVal sourceText As Variant, ByVal scopeName As String) As Boolean
    IsNameExcluded = False

    If Not gNameRulesLoaded Then LoadNameExclusionRules
    If gNameRules Is Nothing Then Exit Function
    If gNameRules.Count = 0 Then Exit Function

    Dim textKey As String
    textKey = NormalizeNameForRules(sourceText)
    If textKey = "" Then Exit Function

    Dim scopeKey As String
    scopeKey = UCase$(Trim(CStr(scopeName)))
    If scopeKey = "" Then scopeKey = "BOTH"

    Dim rule As Variant
    Dim ruleScope As String
    Dim modeKey As String
    Dim pattern As String

    For Each rule In gNameRules
        ruleScope = UCase$(Trim(CStr(rule(0))))
        modeKey = UCase$(Trim(CStr(rule(1))))
        pattern = NormalizeNameForRules(rule(2))

        If pattern <> "" Then
            If ruleScope = "BOTH" Or ruleScope = scopeKey Then
                If RuleMatches(textKey, modeKey, pattern) Then
                    IsNameExcluded = True
                    Exit Function
                End If
            End If
        End If
    Next rule
End Function

Public Function IsTaskMarkedBlue(ByVal taskNumber As Variant) As Boolean
    IsTaskMarkedBlue = False

    If Not gNameRulesLoaded Then LoadNameExclusionRules
    If gTaskColorList Is Nothing Then Exit Function

    Dim taskKey As String
    taskKey = NormalizeTaskKey(taskNumber)
    If taskKey = "" Then Exit Function

    IsTaskMarkedBlue = gTaskColorList.Exists(taskKey)
End Function

Public Function IsTaskStopped(ByVal taskNumber As Variant) As Boolean
    ' Compatibility wrapper: task list no longer blocks analog selection.
    IsTaskStopped = False
End Function

Public Sub LoadNameExclusionRules()
    Set gNameRules = New Collection
    Set gTaskColorList = CreateObject("Scripting.Dictionary")
    gTaskColorList.CompareMode = 1
    gNameRulesLoaded = True

    Dim ws As Worksheet
    Set ws = EnsureNameExclusionsSheet(False)
    If ws Is Nothing Then Exit Sub

    LoadNameRulesFromSheet ws
    LoadTaskColorListFromSheet ws
End Sub

Private Sub LoadNameRulesFromSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 4).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    Dim enabledFlag As Variant
    Dim enabled As Boolean
    Dim scopeKey As String
    Dim modeKey As String
    Dim pattern As String
    Dim groupKey As String

    For r = 2 To lastRow
        enabledFlag = ws.Cells(r, 1).Value
        enabled = IsEnabledValue(enabledFlag)

        If enabled Then
            scopeKey = UCase$(Trim(CStr(ws.Cells(r, 2).Value)))
            modeKey = UCase$(Trim(CStr(ws.Cells(r, 3).Value)))
            pattern = Trim(CStr(ws.Cells(r, 4).Value))
            groupKey = Trim(CStr(ws.Cells(r, 5).Value))

            If scopeKey = "" Then scopeKey = "BOTH"
            If modeKey = "" Then modeKey = "ALL_WORDS"
            If pattern <> "" Then gNameRules.Add Array(scopeKey, modeKey, pattern, groupKey)
        End If
    Next r
End Sub

Private Sub LoadTaskColorListFromSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 9).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    Dim enabled As Boolean
    Dim taskKey As String

    For r = 2 To lastRow
        enabled = IsEnabledValue(ws.Cells(r, 8).Value)
        If enabled Then
            taskKey = NormalizeTaskKey(ws.Cells(r, 9).Value)
            If taskKey <> "" Then
                If Not gTaskColorList.Exists(taskKey) Then gTaskColorList.Add taskKey, True
            End If
        End If
    Next r
End Sub

Private Function EnsureNameExclusionsSheet(ByVal forceDefaultRows As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NAME_EXCLUSIONS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = NAME_EXCLUSIONS_SHEET
        forceDefaultRows = True
    End If

    EnsureNameExclusionHeaders ws

    Dim hasNameData As Boolean
    hasNameData = (Trim(CStr(ws.Cells(2, 4).Value)) <> "")
    If forceDefaultRows Or Not hasNameData Then AddDefaultNameExclusionRows ws

    Set EnsureNameExclusionsSheet = ws
End Function

Private Sub EnsureNameExclusionHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "Enabled"
    ws.Cells(1, 2).Value = "Scope"
    ws.Cells(1, 3).Value = "MatchMode"
    ws.Cells(1, 4).Value = "Pattern"
    ws.Cells(1, 5).Value = "Group"
    ws.Cells(1, 6).Value = "Comment"

    ws.Cells(1, 8).Value = "Enabled"
    ws.Cells(1, 9).Value = "TaskNumber"
    ws.Cells(1, 10).Value = "Reason"
    ws.Cells(1, 11).Value = "Comment"

    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range("H1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(112, 48, 160)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:A").ColumnWidth = 10
    ws.Columns("B:C").ColumnWidth = 14
    ws.Columns("D:D").ColumnWidth = 34
    ws.Columns("E:E").ColumnWidth = 22
    ws.Columns("F:F").ColumnWidth = 60

    ws.Columns("H:H").ColumnWidth = 10
    ws.Columns("I:I").ColumnWidth = 18
    ws.Columns("J:J").ColumnWidth = 22
    ws.Columns("K:K").ColumnWidth = 50
End Sub

Private Sub AddDefaultNameExclusionRows(ByVal ws As Worksheet)
    ws.Range("A2:F100").ClearContents

    WriteNameRule ws, 2, 1, "BOTH", "ALL_WORDS", Ru(1082, 1072, 1078, 1076, "|", 1087, 1086, 1089, 1083, 1077, 1076, 1091, 1102, 1097), "EACH_NEXT", "Each/next work variants."
    WriteNameRule ws, 3, 1, "BOTH", "ALL_WORDS", Ru(1089, 1084, "|", 1080, 1079, 1084, 1077, 1085, 1077, 1085), "CM_CHANGE", "Cm thickness/depth change variants."
    WriteNameRule ws, 4, 1, "BOTH", "ALL_WORDS", Ru(1084, 1084, "|", 1080, 1079, 1084, 1077, 1085, 1077, 1085), "MM_CHANGE", "Mm thickness/depth change variants."
    WriteNameRule ws, 5, 1, "BOTH", "ALL_WORDS", Ru(1076, 1086, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100, 1085, "|", 1097, 1080, 1090, 1082), "EXTRA_SHIELD", "Extra shield variants."

    WriteNameRule ws, 6, 0, "BOTH", "ALL_WORDS", Ru(1089, 1074, 1077, 1088, 1093, "|", 1076, 1086, 1073, 1072, 1074), "ABOVE_LIMIT_ADD", "Optional broad rule. Enable after review."
    WriteNameRule ws, 7, 0, "BOTH", "ALL_WORDS", Ru(1076, 1086, 1073, 1072, 1074, "|", 1080, 1089, 1082, 1083, 1102, 1095), "ADD_OR_EXCLUDE", "Optional broad rule. Enable after review."

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:F7").AutoFilter
    On Error GoTo 0
End Sub

Private Sub WriteNameRule(ByVal ws As Worksheet, _
                          ByVal r As Long, _
                          ByVal enabled As Long, _
                          ByVal scopeKey As String, _
                          ByVal modeKey As String, _
                          ByVal pattern As String, _
                          ByVal groupKey As String, _
                          ByVal commentText As String)
    ws.Cells(r, 1).Value = enabled
    ws.Cells(r, 2).Value = scopeKey
    ws.Cells(r, 3).Value = modeKey
    ws.Cells(r, 4).Value = pattern
    ws.Cells(r, 5).Value = groupKey
    ws.Cells(r, 6).Value = commentText
End Sub

Private Function RuleMatches(ByVal textKey As String, ByVal modeKey As String, ByVal pattern As String) As Boolean
    RuleMatches = False

    If modeKey = "CONTAINS" Then
        RuleMatches = (InStr(1, textKey, pattern, vbTextCompare) > 0)
    Else
        RuleMatches = MatchAllTokens(textKey, pattern)
    End If
End Function

Private Function MatchAllTokens(ByVal textKey As String, ByVal pattern As String) As Boolean
    Dim tokens() As String
    tokens = Split(pattern, "|")

    Dim i As Long
    Dim token As String

    MatchAllTokens = True
    For i = LBound(tokens) To UBound(tokens)
        token = Trim(CStr(tokens(i)))
        If token <> "" Then
            If InStr(1, textKey, token, vbTextCompare) = 0 Then
                MatchAllTokens = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NormalizeNameForRules(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(CStr(v))

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeNameForRules = Trim(s)
End Function

Private Function NormalizeTaskKey(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")

    NormalizeTaskKey = Trim(s)
End Function

Private Function IsEnabledValue(ByVal v As Variant) As Boolean
    IsEnabledValue = False

    If IsNumeric(v) Then
        IsEnabledValue = (CDbl(v) = 1)
    ElseIf UCase$(Trim(CStr(v))) = "TRUE" Then
        IsEnabledValue = True
    End If
End Function

Private Function Ru(ParamArray parts() As Variant) As String
    Dim i As Long
    Dim s As String

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            s = s & ChrW(CLng(parts(i)))
        Else
            s = s & CStr(parts(i))
        End If
    Next i

    Ru = s
End Function

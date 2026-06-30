Attribute VB_Name = "Module6"
Option Explicit

' =========================
' ENTRY
' =========================
Sub RNMC_AppendFiles_ToCatalog()

    Dim wsCat As Worksheet, wsLog As Worksheet
    Dim rootPath As String
    Dim dictLogged As Object

    Set wsCat = ThisWorkbook.Worksheets(SH_CATALOG())
    Set wsLog = ThisWorkbook.Worksheets(SH_FILELOG())

    rootPath = PickFolder()
    If rootPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.StatusBar = "Starting..."

    ' Build the logged-files dictionary ONCE here and pass it down
    Set dictLogged = BuildLoggedFilesDict(wsLog)

    ProcessFolderRecursive rootPath, wsCat, wsLog, dictLogged

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Done", vbInformation
End Sub

' =========================
' FOLDER WALK (recursive)
' dictLogged is built once in the entry point and passed through
' =========================
Private Sub ProcessFolderRecursive(ByVal folderPath As String, ByVal wsCat As Worksheet, _
                                    ByVal wsLog As Worksheet, ByVal dictLogged As Object)

    Dim fso As Object, fld As Object, subFld As Object, fil As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folderPath)

    For Each fil In fld.Files
        If IsExcelFile(fil.Name) Then
            If Not dictLogged.Exists(LCase$(Trim$(fil.Name))) Then
                Application.StatusBar = "Processing: " & fil.Name
                ProcessOneFile CStr(fil.path), CStr(fil.ParentFolder.path), wsCat, wsLog, dictLogged
            End If
        End If
    Next fil

    For Each subFld In fld.SubFolders
        ProcessFolderRecursive CStr(subFld.path), wsCat, wsLog, dictLogged
    Next subFld
End Sub

Private Function IsExcelFile(ByVal fileName As String) As Boolean
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then
        IsExcelFile = False
        Exit Function
    End If
    ext = LCase$(Mid$(fileName, dotPos + 1))
    IsExcelFile = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls")
End Function

' =========================
' PROCESS ONE FILE
' =========================
Private Sub ProcessOneFile(ByVal fullPath As String, ByVal folderFullPath As String, _
                            ByVal wsCat As Worksheet, ByVal wsLog As Worksheet, _
                            ByVal dictLogged As Object)

    Dim wb As Workbook
    Dim addedTotal As Long
    Dim regionName As String
    Dim taskNo As String
    Dim fileNameOnly As String
    Dim errDesc As String

    fileNameOnly = Dir$(fullPath)
    regionName = GetLastFolderName(folderFullPath)

    On Error GoTo SafeLogOnly

    Set wb = Workbooks.Open(fullPath, ReadOnly:=True)

    taskNo = ExtractTaskNumber(wb)

    addedTotal = ImportFromWorkbookToCatalog(wb, wsCat, taskNo, fileNameOnly, regionName)

    wb.Close SaveChanges:=False

    ' Mark as processed in dictionary and log
    dictLogged(LCase$(Trim$(fileNameOnly))) = 1
    AppendFileLog wsLog, folderFullPath, fileNameOnly, addedTotal, ""
    Exit Sub

SafeLogOnly:
    errDesc = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    ' Log with error info in column 4 so user can see what failed
    AppendFileLog wsLog, folderFullPath, fileNameOnly, 0, errDesc
    dictLogged(LCase$(Trim$(fileNameOnly))) = 1
End Sub

' =========================
' IMPORT FROM WORKBOOK -> CATALOG
' stop after first successful import (>0 rows) to avoid duplicates
' =========================
Private Function ImportFromWorkbookToCatalog(ByVal wbSrc As Workbook, ByVal wsCat As Worksheet, _
                                             ByVal taskNo As String, ByVal srcFile As String, _
                                             ByVal regionName As String) As Long
    Dim mapCat As Object
    Dim ws As Worksheet
    Dim hdrRow As Long
    Dim colName As Long, colUnit As Long, colQty As Long
    Dim mapSrc As Object
    Dim added As Long

    Set mapCat = BuildCatalogHeaderMap(wsCat, 3) ' headers on row 3
    added = 0

    For Each ws In wbSrc.Worksheets
        If FindHeaderRow(ws, hdrRow, colName, colUnit, colQty) Then
            Set mapSrc = BuildSourceHeaderMap(ws, hdrRow)

            added = added + AppendTableRows(ws, hdrRow, colName, colUnit, colQty, _
                                            mapSrc, mapCat, wsCat, taskNo, srcFile, regionName)

            If added > 0 Then Exit For
        End If
    Next ws

    ImportFromWorkbookToCatalog = added
End Function

' =========================
' APPEND ROWS
' End condition: 3 consecutive fully blank rows in (num + name + unit + qty)
' Skip if BOTH unit and qty are blank
' Values only (Value2)
' =========================
Private Function AppendTableRows( _
    ByVal wsSrc As Worksheet, _
    ByVal hdrRow As Long, _
    ByVal colName As Long, _
    ByVal colUnit As Long, _
    ByVal colQty As Long, _
    ByVal mapSrc As Object, _
    ByVal mapCat As Object, _
    ByVal wsCat As Worksheet, _
    ByVal taskNo As String, _
    ByVal srcFile As String, _
    ByVal regionName As String _
) As Long

    Dim r As Long
    Dim outRow As Long
    Dim nextNum As Long
    Dim added As Long
    Dim numCol As Long
    Dim started As Boolean
    Dim blankStreak As Long
    Dim isEndBlankRow As Boolean
    Dim key As Variant
    Dim srcCol As Long
    Dim catCol As Long
    Dim maxRow As Long

    nextNum = GetNextCatalogNumber(wsCat)

    numCol = FindNumberingCol(mapSrc)
    If numCol = 0 Then numCol = 1

    ' Safety ceiling: never go past UsedRange
    maxRow = wsSrc.UsedRange.row + wsSrc.UsedRange.Rows.count - 1

    r = hdrRow + 1
    started = False
    blankStreak = 0
    added = 0

    Do While r <= maxRow

        ' Skip leading empty rows until data begins
        If Not started Then
            If IsEmptyOrBlank(wsSrc.Cells(r, numCol).Value2) And _
               IsEmptyOrBlank(wsSrc.Cells(r, colUnit).Value2) And _
               IsEmptyOrBlank(wsSrc.Cells(r, colQty).Value2) Then
                r = r + 1
                GoTo ContinueLoop
            Else
                started = True
            End If
        End If

        isEndBlankRow = IsEmptyOrBlank(wsSrc.Cells(r, numCol).Value2) And _
                        IsEmptyOrBlank(wsSrc.Cells(r, colName).Value2) And _
                        IsEmptyOrBlank(wsSrc.Cells(r, colUnit).Value2) And _
                        IsEmptyOrBlank(wsSrc.Cells(r, colQty).Value2)

        If isEndBlankRow Then
            blankStreak = blankStreak + 1
        Else
            blankStreak = 0
        End If

        If blankStreak >= 3 Then Exit Do

        If Not (IsEmptyOrBlank(wsSrc.Cells(r, colUnit).Value2) And _
                IsEmptyOrBlank(wsSrc.Cells(r, colQty).Value2)) Then

            outRow = NextFreeRowAny(wsCat, 4)

            ' Special columns in Catalog
            wsCat.Cells(outRow, 1).Value2 = nextNum
            wsCat.Cells(outRow, 2).Value2 = taskNo
            wsCat.Cells(outRow, 15).Value2 = srcFile
            wsCat.Cells(outRow, 16).Value2 = regionName

            nextNum = nextNum + 1

            ' Map headers: only those that exist in Catalog
            For Each key In mapSrc.Keys
                If mapCat.Exists(key) Then
                    catCol = CLng(mapCat(key))
                    If catCol <> 1 And catCol <> 2 And catCol <> 15 And catCol <> 16 Then
                        srcCol = CLng(mapSrc(key))
                        wsCat.Cells(outRow, catCol).Value2 = wsSrc.Cells(r, srcCol).Value2
                    End If
                End If
            Next key

            added = added + 1
        End If

        r = r + 1

ContinueLoop:
    Loop

    AppendTableRows = added
End Function

' =========================
' HEADER ROW DETECTION
' Required: name + unit + qty (tolerant)
' =========================
Private Function FindHeaderRow(ByVal ws As Worksheet, ByRef hdrRow As Long, _
                                ByRef colName As Long, ByRef colUnit As Long, ByRef colQty As Long) As Boolean
    Dim lastR As Long, lastC As Long
    Dim r As Long, c As Long
    Dim s As String

    If ws.UsedRange.Cells.count = 1 And Trim$(CStr(ws.UsedRange.Cells(1, 1).Value2)) = "" Then
        FindHeaderRow = False
        Exit Function
    End If

    lastR = ws.UsedRange.row + ws.UsedRange.Rows.count - 1
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.count - 1

    If lastR > 400 Then lastR = 400
    If lastC > 150 Then lastC = 150

    For r = 1 To lastR
        colName = 0: colUnit = 0: colQty = 0

        For c = 1 To lastC
            s = NormalizeText(ws.Cells(r, c).Value2)

            If IsNameHeader(s) Then colName = c
            If IsUnitHeader(s) Then colUnit = c
            If IsQtyHeader(s) Then colQty = c
        Next c

        If colName > 0 And colUnit > 0 And colQty > 0 Then
            hdrRow = r
            FindHeaderRow = True
            Exit Function
        End If
    Next r

    FindHeaderRow = False
End Function

Private Function IsNameHeader(ByVal s As String) As Boolean
    Dim n1 As String, n2 As String
    n1 = NormalizeText(BuildHeader_Name())
    n2 = NormalizeText(BuildHeader_Name2())
    IsNameHeader = (s = n1 Or s = n2 Or Left$(s, Len(n2)) = n2)
End Function

Private Function IsUnitHeader(ByVal s As String) As Boolean
    Dim u As String
    u = NormalizeText(BuildHeader_Unit())
    IsUnitHeader = (Left$(s, Len(u)) = u Or InStr(1, s, u, vbTextCompare) > 0)
End Function

Private Function IsQtyHeader(ByVal s As String) As Boolean
    Dim q1 As String, q2 As String
    q1 = NormalizeText(BuildHeader_Qty())
    q2 = NormalizeText(BuildHeader_Qty2())
    IsQtyHeader = (InStr(1, s, q1, vbTextCompare) > 0 Or InStr(1, s, q2, vbTextCompare) > 0)
End Function

' =========================
' HEADER MAPS
' =========================
Private Function BuildCatalogHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Dim lastC As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastC = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastC
        h = Trim$(CStr(ws.Cells(headerRow, c).Value2))
        If h <> "" Then dict(NormalizeText(h)) = c
    Next c

    Set BuildCatalogHeaderMap = dict
End Function

Private Function BuildSourceHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Dim lastC As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastC = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastC
        h = Trim$(CStr(ws.Cells(headerRow, c).Value2))
        If h <> "" Then dict(NormalizeText(h)) = c
    Next c

    Set BuildSourceHeaderMap = dict
End Function

' =========================
' NUMBERING COLUMN DETECTION (no unicode literals)
' Tightened: require at least 2 chars to avoid false matches on "n"
' =========================
Private Function FindNumberingCol(ByVal mapSrc As Object) As Long
    Dim key As Variant
    Dim s As String
    Dim numSign As String

    numSign = ChrW(8470) ' №

    For Each key In mapSrc.Keys
        s = LCase$(CStr(key))

        If InStr(1, s, LCase$(numSign), vbTextCompare) > 0 _
           Or Left$(s, 2) = "no" _
           Or InStr(1, s, "pp", vbTextCompare) > 0 _
           Or InStr(1, s, "p/p", vbTextCompare) > 0 _
           Or InStr(1, s, "p-p", vbTextCompare) > 0 Then

            FindNumberingCol = CLng(mapSrc(key))
            Exit Function
        End If
    Next key

    FindNumberingCol = 0
End Function

' =========================
' TASK NUMBER (search top area)
' =========================
Private Function ExtractTaskNumber(ByVal wb As Workbook) As String

    Dim ws As Worksheet
    Dim lbl1 As String, lbl2 As String
    Dim r As Long, c As Long
    Dim txt As String, tail As String
    Dim p As Long

    lbl1 = BuildTaskLabel()
    lbl2 = BuildTaskLabelShort()

    For Each ws In wb.Worksheets
        For r = 1 To 50
            For c = 1 To 20

                txt = CStr(ws.Cells(r, c).Value2)
                If Len(txt) > 0 Then

                    p = InStr(1, txt, lbl1, vbTextCompare)
                    If p > 0 Then
                        tail = Trim$(Mid$(txt, p + Len(lbl1)))
                        tail = CleanupTaskTail(tail)
                        If tail <> "" Then
                            ExtractTaskNumber = tail
                            Exit Function
                        End If
                        tail = NeighborTaskValue(ws, r, c)
                        If tail <> "" Then
                            ExtractTaskNumber = tail
                            Exit Function
                        End If
                    End If

                    p = InStr(1, txt, lbl2, vbTextCompare)
                    If p > 0 Then
                        tail = Trim$(Mid$(txt, p + Len(lbl2)))
                        tail = CleanupTaskTail(tail)
                        If tail <> "" Then
                            ExtractTaskNumber = tail
                            Exit Function
                        End If
                        tail = NeighborTaskValue(ws, r, c)
                        If tail <> "" Then
                            ExtractTaskNumber = tail
                            Exit Function
                        End If
                    End If

                End If

            Next c
        Next r
    Next ws

    ExtractTaskNumber = ""
End Function

Private Function CleanupTaskTail(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)

    t = Replace(t, ":", "")
    t = Replace(t, "#", "")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)

    CleanupTaskTail = t
End Function

Private Function NeighborTaskValue(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim k As Long
    Dim v As String

    For k = 1 To 3
        v = Trim$(CStr(ws.Cells(r, c + k).Value2))
        If v <> "" Then
            NeighborTaskValue = v
            Exit Function
        End If
    Next k

    NeighborTaskValue = ""
End Function

' =========================
' FILELOG
' =========================
Private Function BuildLoggedFilesDict(ByVal wsLog As Worksheet) As Object
    Dim dict As Object
    Dim lastR As Long, r As Long
    Dim fn As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastR = wsLog.Cells(wsLog.Rows.count, 2).End(xlUp).row
    If lastR < 1 Then lastR = 1

    For r = 1 To lastR
        fn = Trim$(CStr(wsLog.Cells(r, 2).Value2))
        If fn <> "" Then dict(LCase$(fn)) = 1
    Next r

    Set BuildLoggedFilesDict = dict
End Function

' Removed IsFileLogged — use dictLogged.Exists() everywhere instead

Private Sub AppendFileLog(ByVal wsLog As Worksheet, ByVal folderFullPath As String, _
                           ByVal fileName As String, ByVal addedCount As Long, _
                           ByVal errText As String)
    Dim r As Long
    r = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).row
    If r < 1 Then r = 1

    If Trim$(CStr(wsLog.Cells(r, 1).Value2)) <> "" Or Trim$(CStr(wsLog.Cells(r, 2).Value2)) <> "" Then
        r = r + 1
    End If

    wsLog.Cells(r, 1).Value2 = folderFullPath
    wsLog.Cells(r, 2).Value2 = fileName
    wsLog.Cells(r, 3).Value2 = addedCount
    wsLog.Cells(r, 4).Value2 = Now()               ' timestamp
    If errText <> "" Then
        wsLog.Cells(r, 5).Value2 = errText          ' error description if any
    End If
End Sub

' =========================
' CATALOG HELPERS
' =========================
Private Function NextFreeRowAny(ByVal ws As Worksheet, ByVal minRow As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        NextFreeRowAny = minRow
    Else
        If lastCell.row < minRow Then
            NextFreeRowAny = minRow
        Else
            NextFreeRowAny = lastCell.row + 1
        End If
    End If
End Function

Private Function GetNextCatalogNumber(ByVal wsCat As Worksheet) As Long
    Dim lastR As Long
    Dim mx As Variant

    lastR = wsCat.Cells(wsCat.Rows.count, 1).End(xlUp).row
    If lastR < 4 Then
        GetNextCatalogNumber = 1
        Exit Function
    End If

    On Error Resume Next
    mx = Application.WorksheetFunction.Max(wsCat.Range("A4:A" & lastR))
    On Error GoTo 0

    If IsError(mx) Or mx = "" Then
        GetNextCatalogNumber = 1
    Else
        GetNextCatalogNumber = CLng(mx) + 1
    End If
End Function

' =========================
' TEXT / PATH UTIL
' =========================
Private Function NormalizeText(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    NormalizeText = s
End Function

Private Function IsEmptyOrBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyOrBlank = True
    Else
        IsEmptyOrBlank = (Trim$(CStr(v)) = "")
    End If
End Function

Private Function PickFolder() As String
    With Application.FileDialog(4)
        .Title = "Select root folder with RNMC files (can include subfolders)"
        ' .Show returns -1 on OK, 0 on Cancel — fixed from original
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Private Function GetLastFolderName(ByVal folderFullPath As String) As String
    Dim p As Long
    Dim s As String
    s = folderFullPath
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    p = InStrRev(s, "\")
    If p > 0 Then
        GetLastFolderName = Mid$(s, p + 1)
    Else
        GetLastFolderName = s
    End If
End Function

' =========================
' SHEET NAMES (no unicode literals)
' =========================
Private Function SH_CATALOG() As String
    ' "Каталог"
    SH_CATALOG = ChrW(1050) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1086) & ChrW(1075)
End Function

Private Function SH_FILELOG() As String
    ' "FileLog"
    SH_FILELOG = "FileLog"
End Function

' =========================
' REQUIRED HEADERS / LABELS (no unicode literals)
' =========================
Private Function BuildHeader_Name() As String
    ' "Наименование работ"
    BuildHeader_Name = ChrW(1053) & ChrW(1072) & ChrW(1080) & ChrW(1084) & ChrW(1077) & ChrW(1085) & _
                       ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077) & _
                       ChrW(32) & ChrW(1088) & ChrW(1072) & ChrW(1073) & ChrW(1086) & ChrW(1090)
End Function

Private Function BuildHeader_Name2() As String
    ' "Наименование"
    BuildHeader_Name2 = ChrW(1053) & ChrW(1072) & ChrW(1080) & ChrW(1084) & ChrW(1077) & ChrW(1085) & _
                        ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077)
End Function

Private Function BuildHeader_Unit() As String
    ' "Ед.изм."
    BuildHeader_Unit = ChrW(1045) & ChrW(1076) & ChrW(46) & ChrW(1080) & ChrW(1079) & ChrW(1084) & ChrW(46)
End Function

Private Function BuildHeader_Qty() As String
    ' "Кол-во"
    BuildHeader_Qty = ChrW(1050) & ChrW(1086) & ChrW(1083) & ChrW(45) & ChrW(1074) & ChrW(1086)
End Function

Private Function BuildHeader_Qty2() As String
    ' "Количество"
    BuildHeader_Qty2 = ChrW(1050) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & _
                       ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086)
End Function

Private Function BuildTaskLabel() As String
    ' "№ задачи 1Ф"
    BuildTaskLabel = ChrW(8470) & ChrW(32) & _
                     ChrW(1079) & ChrW(1072) & ChrW(1076) & ChrW(1072) & ChrW(1095) & ChrW(1080) & _
                     ChrW(32) & ChrW(49) & ChrW(1060)
End Function

Private Function BuildTaskLabelShort() As String
    ' "№ задачи"
    BuildTaskLabelShort = ChrW(8470) & ChrW(32) & _
                           ChrW(1079) & ChrW(1072) & ChrW(1076) & ChrW(1072) & ChrW(1095) & ChrW(1080)
End Function



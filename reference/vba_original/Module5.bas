Attribute VB_Name = "Module5"
' ============================================================
'  MODULE 4: LOG v3.0
'  - WriteLog  - zapisat' statistiku na list "Log"
' ============================================================
Option Explicit

Public Sub WriteLog(logLines() As String, smPath As String, elapsed As Double)

    ' Nazvanie lista Loga (kirillica cherez Chr)
    Dim logName As String
    logName = ChrW(1051) & ChrW(1086) & ChrW(1075)

    ' Najti ili sozdat' list
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(logName)
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = logName
    End If

    wsLog.Cells.Clear

    ' --- Shapka ---
    With wsLog.Cells(1, 1)
        .Value = "Log: " & smPath
        .Font.Bold = True: .Font.Size = 11
    End With
    wsLog.Cells(2, 1).Value = Format(Now, "dd.mm.yyyy hh:mm")
    wsLog.Cells(3, 1).Value = "Vremya: " & Format(elapsed, "0.0") & " sek."
    wsLog.Cells(4, 1).Value = "Deduplicaciya: " & DEDUP_PCT * 100 & "%"

    If LBound(logLines) > UBound(logLines) Then GoTo Finish

    ' --- Zagolovjki tablicy (stroka 6) ---
    Dim headers() As String
    headers = Split(logLines(LBound(logLines)), "|")
    Dim hc As Long
    For hc = 0 To UBound(headers)
        With wsLog.Cells(6, hc + 1)
            .Value = headers(hc)
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next hc

    ' --- Stroki dannyh ---
    Dim outRow As Long: outRow = 7
    Dim li As Long
    For li = LBound(logLines) + 1 To UBound(logLines)
        Dim cols() As String: cols = Split(logLines(li), "|")
        Dim cc As Long
        For cc = 0 To UBound(cols)
            wsLog.Cells(outRow, cc + 1).Value = cols(cc)
        Next cc
        If outRow Mod 2 = 0 Then
            wsLog.Rows(outRow).Interior.Color = RGB(242, 242, 242)
        End If
        outRow = outRow + 1
    Next li

    ' --- Itog ---
    Dim totalWithAnalogs As Long: totalWithAnalogs = 0
    For li = LBound(logLines) + 1 To UBound(logLines)
        Dim c4 As String
        c4 = Split(logLines(li), "|")(3)  ' stolbec "Cen"
        If IsNumeric(c4) And CLng(c4) > 0 Then
            totalWithAnalogs = totalWithAnalogs + 1
        End If
    Next li

    With wsLog.Cells(outRow + 1, 1)
        .Value = "Vsego strok obrabotano: " & (outRow - 7)
        .Font.Bold = True
    End With
    With wsLog.Cells(outRow + 2, 1)
        .Value = "Iz nih s analogami: " & totalWithAnalogs
        .Font.Bold = True
    End With

    wsLog.Columns("A:G").AutoFit

Finish:
    wsLog.Activate
End Sub

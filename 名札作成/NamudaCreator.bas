Attribute VB_Name = "NamudaCreator"
' Nagasaki Kita Tokushukai Hospital - Employee Badge Creator
' All Japanese text uses ChrW() so the .bas file imports correctly
' regardless of file encoding.
Option Explicit

Private Const BADGE_ROWS     As Integer = 14
Private Const BADGE_COLS     As Integer = 11
Private Const BADGES_PER_ROW As Integer = 2
Private Const COL_GAP        As Integer = 1
Private Const ROW_GAP        As Integer = 1

Private Const COL_ID   As Integer = 1
Private Const COL_NAME As Integer = 2
Private Const COL_KANA As Integer = 3
Private Const COL_DEPT As Integer = 4
Private Const COL_JOB  As Integer = 5
Private Const COL_POS  As Integer = 6
Private Const COL_JOIN As Integer = 7
Private Const COL_MADE As Integer = 8

'--------------------------------------------------------------
' Japanese string helpers (ChrW avoids encoding issues)
'--------------------------------------------------------------
Private Function SH_MASTER() As String
    ' 職員マスター
    SH_MASTER = ChrW(32887) & ChrW(21729) & ChrW(12510) & ChrW(12473) & ChrW(12479) & ChrW(12540)
End Function

Private Function SH_BADGE() As String
    ' 名札印刷
    SH_BADGE = ChrW(21517) & ChrW(26413) & ChrW(21360) & ChrW(21047)
End Function

Private Function FN_GOTHIC() As String
    ' MS Pゴシック
    FN_GOTHIC = "MS P" & ChrW(12468) & ChrW(12471) & ChrW(12483) & ChrW(12463)
End Function

Private Function FN_MINCHO() As String
    ' MS P明朝
    FN_MINCHO = "MS P" & ChrW(26126) & ChrW(26397)
End Function

Private Function STR_CORP() As String
    ' 医療法人　徳洲会
    STR_CORP = ChrW(21307) & ChrW(30274) & ChrW(27861) & ChrW(20154) & ChrW(12288) & ChrW(24499) & ChrW(27954) & ChrW(20250)
End Function

Private Function STR_HOSPITAL() As String
    ' 長崎北徳洲会病院
    STR_HOSPITAL = ChrW(38263) & ChrW(23822) & ChrW(21271) & ChrW(24499) & ChrW(27954) & ChrW(20250) & ChrW(30149) & ChrW(38498)
End Function

'--------------------------------------------------------------
' Employee data type
'--------------------------------------------------------------
Private Type EmpData
    ID      As String
    Name    As String
    Kana    As String
    Dept    As String
    JobType As String
    Pos     As String
End Type

Private Function ReadEmpRow(ws As Worksheet, r As Long) As EmpData
    Dim e As EmpData
    e.ID      = CStr(ws.Cells(r, COL_ID).Value)
    e.Name    = CStr(ws.Cells(r, COL_NAME).Value)
    e.Kana    = CStr(ws.Cells(r, COL_KANA).Value)
    e.Dept    = CStr(ws.Cells(r, COL_DEPT).Value)
    e.JobType = CStr(ws.Cells(r, COL_JOB).Value)
    e.Pos     = CStr(ws.Cells(r, COL_POS).Value)
    ReadEmpRow = e
End Function

'==============================================================
' PUBLIC SUBS  (assign to buttons on the sheet)
'==============================================================

' Initial setup: run once to create the two worksheets
Public Sub Setup()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call CreateMasterSheet
    Call CreateBadgeSheet
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Worksheets(SH_MASTER()).Activate
    Worksheets(SH_MASTER()).Cells(2, 1).Select
    MsgBox "Setup complete.", vbInformation
End Sub

' Create badges for rows selected in the master sheet
Public Sub CreateBadgesSelected()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub
    If ActiveSheet.Name <> mn Then
        MsgBox "Select rows in the " & mn & " sheet first.", vbExclamation: Exit Sub
    End If

    Dim selRows() As Long, cnt As Integer: cnt = 0
    Dim cel As Range
    For Each cel In Selection
        If cel.Row > 1 Then
            Dim dup As Boolean: dup = False
            Dim j As Integer
            For j = 0 To cnt - 1
                If selRows(j) = cel.Row Then dup = True: Exit For
            Next j
            If Not dup Then
                ReDim Preserve selRows(cnt)
                selRows(cnt) = cel.Row
                cnt = cnt + 1
            End If
        End If
    Next cel
    If cnt = 0 Then MsgBox "Select data rows (row 2+).", vbInformation: Exit Sub

    Dim ws As Worksheet: Set ws = Worksheets(mn)
    Dim empList() As EmpData: ReDim empList(cnt - 1)
    Dim validCnt As Integer: validCnt = 0
    Dim i As Integer
    For i = 0 To cnt - 1
        Dim r As Long: r = selRows(i)
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            empList(validCnt) = ReadEmpRow(ws, r)
            ws.Cells(r, COL_MADE).Value = Date
            validCnt = validCnt + 1
        End If
    Next i
    If validCnt = 0 Then MsgBox "No valid ID found.", vbInformation: Exit Sub
    Call GenerateBadges(empList, validCnt)
    MsgBox validCnt & " badge(s) created.", vbInformation
End Sub

' Create badges for all employees
Public Sub CreateBadgesAll()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(mn)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No data found.", vbInformation: Exit Sub

    Dim empList() As EmpData: ReDim empList(lastRow - 2)
    Dim cnt As Integer: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            empList(cnt) = ReadEmpRow(ws, r)
            cnt = cnt + 1
        End If
    Next r
    If cnt = 0 Then MsgBox "No data found.", vbInformation: Exit Sub
    If MsgBox("Create badges for all " & cnt & " employee(s)?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then ws.Cells(r, COL_MADE).Value = Date
    Next r
    Call GenerateBadges(empList, cnt)
    MsgBox cnt & " badge(s) created.", vbInformation
End Sub

' Create badges filtered by join year/month  e.g. April batch
Public Sub CreateBadgesByMonth()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub

    Dim inputVal As String
    inputVal = InputBox("Enter join year/month (e.g. 2026/4):", "Filter by Month", Format(Date, "yyyy/m"))
    If inputVal = "" Then Exit Sub
    inputVal = Replace(inputVal, "-", "/")
    Dim parts() As String: parts = Split(inputVal, "/")
    If UBound(parts) < 1 Then MsgBox "Use format: 2026/4", vbExclamation: Exit Sub
    Dim targetYear As Integer:  targetYear  = CInt(parts(0))
    Dim targetMonth As Integer: targetMonth = CInt(parts(1))

    Dim ws As Worksheet: Set ws = Worksheets(mn)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    Dim empList() As EmpData: ReDim empList(lastRow - 2)
    Dim cnt As Integer: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) = "" Then GoTo NextRow
        Dim jd As Variant: jd = ws.Cells(r, COL_JOIN).Value
        If Not IsDate(jd) Then GoTo NextRow
        If Year(CDate(jd)) = targetYear And Month(CDate(jd)) = targetMonth Then
            empList(cnt) = ReadEmpRow(ws, r)
            ws.Cells(r, COL_MADE).Value = Date
            cnt = cnt + 1
        End If
NextRow:
    Next r
    If cnt = 0 Then MsgBox "No employees found for " & targetYear & "/" & targetMonth, vbInformation: Exit Sub
    If MsgBox("Create " & cnt & " badge(s) for " & targetYear & "/" & targetMonth & "?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call GenerateBadges(empList, cnt)
    MsgBox cnt & " badge(s) created.", vbInformation
End Sub

' Print preview
Public Sub BadgePrintPreview()
    If SheetExists(SH_BADGE()) Then
        Worksheets(SH_BADGE()).PrintPreview
    Else
        MsgBox "No badge sheet found. Create badges first.", vbExclamation
    End If
End Sub

'==============================================================
' PRIVATE: badge generation
'==============================================================

Private Sub GenerateBadges(empList() As EmpData, cnt As Integer)
    Call ResetBadgeSheet
    Dim wb As Worksheet: Set wb = Worksheets(SH_BADGE())
    Call SetAllDimensions(wb, cnt)
    Application.ScreenUpdating = False
    Dim i As Integer
    For i = 0 To cnt - 1
        Call DrawOneBadge(wb, i, empList(i))
    Next i
    Application.ScreenUpdating = True
    wb.Activate
    wb.Cells(1, 1).Select
End Sub

Private Sub SetAllDimensions(ws As Worksheet, cnt As Integer)
    Dim rowCount As Integer: rowCount = ((cnt - 1) \ BADGES_PER_ROW) + 1
    Dim rh(13) As Single
    rh(0)  = 2:  rh(1)  = 13: rh(2)  = 13: rh(3)  = 11
    rh(4)  = 15: rh(5)  = 15: rh(6)  = 15: rh(7)  = 15
    rh(8)  = 12: rh(9)  = 5
    rh(10) = 11: rh(11) = 11: rh(12) = 11: rh(13) = 2
    Dim rb As Integer, ri As Integer, br As Long
    For rb = 0 To rowCount - 1
        br = rb * (BADGE_ROWS + ROW_GAP) + 1
        For ri = 0 To BADGE_ROWS - 1
            ws.Rows(br + ri).RowHeight = rh(ri)
        Next ri
        If rb < rowCount - 1 Then ws.Rows(br + BADGE_ROWS).RowHeight = 6
    Next rb
    Dim cw(10) As Single
    cw(0) = 1:   cw(1) = 3.5: cw(2) = 3.5: cw(3) = 3.5
    cw(4) = 3.5: cw(5) = 3.5: cw(6) = 3
    cw(7) = 0.5: cw(8) = 3:   cw(9) = 3:   cw(10) = 1.5
    Dim ci As Integer
    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + ci).ColumnWidth = cw(ci)
    Next ci
    ws.Columns(1 + BADGE_COLS).ColumnWidth = 1
    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + BADGE_COLS + COL_GAP + ci).ColumnWidth = cw(ci)
    Next ci
End Sub

Private Sub DrawOneBadge(ws As Worksheet, badgeIdx As Integer, emp As EmpData)
    Dim sRow As Long: sRow = (badgeIdx \ BADGES_PER_ROW) * (BADGE_ROWS + ROW_GAP) + 1
    Dim sCol As Long: sCol = (badgeIdx Mod BADGES_PER_ROW) * (BADGE_COLS + COL_GAP) + 1
    Call DrawBadgeContent(ws, sRow, sCol, emp)
End Sub

Private Sub DrawBadgeContent(ws As Worksheet, sRow As Long, sCol As Long, emp As EmpData)
    Dim fr As Range
    Set fr = ws.Range(ws.Cells(sRow, sCol), ws.Cells(sRow + BADGE_ROWS - 1, sCol + BADGE_COLS - 1))
    fr.UnMerge: fr.ClearContents: fr.ClearFormats
    fr.Interior.Color = RGB(255, 255, 255)

    Dim fc As Long: fc = RGB(80, 80, 160)
    fr.Borders(xlEdgeLeft).LineStyle   = xlContinuous: fr.Borders(xlEdgeLeft).Weight   = xlMedium: fr.Borders(xlEdgeLeft).Color   = fc
    fr.Borders(xlEdgeRight).LineStyle  = xlContinuous: fr.Borders(xlEdgeRight).Weight  = xlMedium: fr.Borders(xlEdgeRight).Color  = fc
    fr.Borders(xlEdgeTop).LineStyle    = xlContinuous: fr.Borders(xlEdgeTop).Weight    = xlMedium: fr.Borders(xlEdgeTop).Color    = fc
    fr.Borders(xlEdgeBottom).LineStyle = xlContinuous: fr.Borders(xlEdgeBottom).Weight = xlMedium: fr.Borders(xlEdgeBottom).Color = fc

    ' Left area
    Call MC(ws, sRow+1, sCol, 2, 7, emp.Dept, 14, True,  FN_GOTHIC(), RGB(0,0,0),   RGB(255,255,255), xlLeft,   xlCenter)

    Dim jobPos As String: jobPos = emp.JobType
    If Trim(emp.Pos) <> "" Then jobPos = jobPos & ChrW(12288) & emp.Pos
    Call MC(ws, sRow+3, sCol, 1, 7, jobPos,   10, False, FN_GOTHIC(), RGB(0,0,0),   RGB(255,255,255), xlLeft,   xlCenter)
    Call MC(ws, sRow+4, sCol, 4, 7, emp.Kana, 32, True,  FN_GOTHIC(), RGB(0,0,0),   RGB(255,255,255), xlCenter, xlCenter)
    Call MC(ws, sRow+8, sCol, 1, 7, emp.Name, 10, False, FN_MINCHO(), RGB(50,50,50), RGB(255,255,255), xlCenter, xlCenter)

    ' Right area: blank top, barcode bottom
    Dim RC As Long: RC = sCol + 7
    ws.Range(ws.Cells(sRow,   RC), ws.Cells(sRow+3, sCol+BADGE_COLS-1)).Merge
    Dim bc As Range
    Set bc = ws.Range(ws.Cells(sRow+4, RC), ws.Cells(sRow+8, sCol+BADGE_COLS-1))
    bc.Merge
    bc.Value = "| || ||| || | ||| || | ||"
    bc.Font.Size = 7: bc.Font.Name = "Courier New": bc.Font.Color = RGB(0, 0, 0)
    bc.Interior.Color = RGB(255, 255, 255)
    bc.HorizontalAlignment = xlCenter: bc.VerticalAlignment = xlCenter

    ' Blue bar + ID
    Dim bar As Range
    Set bar = ws.Range(ws.Cells(sRow+9, sCol), ws.Cells(sRow+9, sCol+8))
    bar.Merge: bar.Interior.Color = RGB(42, 107, 183)
    Dim idr As Range
    Set idr = ws.Range(ws.Cells(sRow+9, sCol+9), ws.Cells(sRow+9, sCol+BADGE_COLS-1))
    idr.Merge
    idr.Value = emp.ID: idr.Font.Size = 8: idr.Font.Name = "Arial"
    idr.Font.Color = RGB(0,0,0): idr.Interior.Color = RGB(255,255,255)
    idr.HorizontalAlignment = xlRight: idr.VerticalAlignment = xlCenter

    ' Footer
    Dim fbg As Long: fbg = RGB(214, 234, 248)
    Dim fa As Range
    Set fa = ws.Range(ws.Cells(sRow+10, sCol), ws.Cells(sRow+12, sCol+BADGE_COLS-1))
    fa.Interior.Color = fbg
    fa.Borders(xlEdgeTop).LineStyle = xlContinuous: fa.Borders(xlEdgeTop).Color = RGB(80, 80, 160)
    Dim lr As Range
    Set lr = ws.Range(ws.Cells(sRow+10, sCol), ws.Cells(sRow+12, sCol+2))
    lr.Merge: lr.Interior.Color = fbg
    lr.Borders(xlEdgeRight).LineStyle = xlContinuous: lr.Borders(xlEdgeRight).Color = RGB(150, 150, 200)
    Call MC(ws, sRow+10, sCol+3, 1, 8, STR_CORP(),     8,  False, FN_MINCHO(), RGB(0,0,0), fbg, xlLeft, xlCenter)
    Call MC(ws, sRow+11, sCol+3, 2, 8, STR_HOSPITAL(), 13, True,  FN_MINCHO(), RGB(0,0,0), fbg, xlLeft, xlCenter)

    Call TryInsertLogo(ws, sRow, sCol)
End Sub

Private Sub MC(ws As Worksheet, sr As Long, sc As Long, nr As Integer, nc As Integer, _
               val As String, sz As Integer, bold As Boolean, fn As String, _
               fc As Long, bg As Long, ha As XlHAlign, va As XlVAlign)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(sr, sc), ws.Cells(sr+nr-1, sc+nc-1))
    rng.Merge: rng.Value = val
    rng.Font.Size = sz: rng.Font.Bold = bold: rng.Font.Name = fn
    rng.Font.Color = fc: rng.Interior.Color = bg
    rng.HorizontalAlignment = ha: rng.VerticalAlignment = va
End Sub

Private Sub TryInsertLogo(ws As Worksheet, sRow As Long, sCol As Long)
    Dim lp As String: lp = ThisWorkbook.Path & "\tokushukai_logo.png"
    If Dir(lp) = "" Then Exit Sub
    Dim ar As Range
    Set ar = ws.Range(ws.Cells(sRow+10, sCol), ws.Cells(sRow+12, sCol+2))
    On Error GoTo Ex
    Dim p As Shape
    Set p = ws.Shapes.AddPicture(Filename:=lp, LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, Left:=ar.Left+2, Top:=ar.Top+2, Width:=-1, Height:=ar.Height-4)
    p.LockAspectRatio = msoTrue
    p.Name = "Logo_R" & sRow & "C" & sCol
Ex: On Error GoTo 0
End Sub

'==============================================================
' PRIVATE: sheet creation
'==============================================================

Private Sub CreateMasterSheet()
    On Error Resume Next: Worksheets(SH_MASTER()).Delete: On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets.Add(Before:=Worksheets(1))
    ws.Name = SH_MASTER()

    ' Column headers
    ws.Cells(1,1).Value = ChrW(32887) & ChrW(21729) & "ID"  ' 職員ID
    ws.Cells(1,2).Value = ChrW(27663) & ChrW(21517) & _     ' 氏名（漢字）
                          ChrW(65288) & ChrW(28450) & ChrW(23383) & ChrW(65289)
    ws.Cells(1,3).Value = ChrW(33495) & ChrW(23383) & _     ' 苗字ひらがな
                          ChrW(12402) & ChrW(12425) & ChrW(12364) & ChrW(12394)
    ws.Cells(1,4).Value = ChrW(37096) & ChrW(32626) & ChrW(21517)          ' 部署名
    ws.Cells(1,5).Value = ChrW(32887) & ChrW(31278)                         ' 職種
    ws.Cells(1,6).Value = ChrW(24441) & ChrW(32887)                         ' 役職
    ws.Cells(1,7).Value = ChrW(20837) & ChrW(32887) & ChrW(26085)           ' 入職日
    ws.Cells(1,8).Value = ChrW(21517) & ChrW(26413) & _                    ' 名札作成日
                          ChrW(20316) & ChrW(25104) & ChrW(26085)

    With ws.Range("A1:H1")
        .Font.Bold = True: .Font.Color = RGB(255,255,255)
        .Interior.Color = RGB(42,107,183): .HorizontalAlignment = xlCenter: .RowHeight = 22
    End With
    ws.Columns("A").ColumnWidth = 10: ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 16: ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 16: ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 12: ws.Columns("H").ColumnWidth = 12
    ws.Columns("G").NumberFormat = "yyyy/m/d"
    ws.Columns("H").NumberFormat = "yyyy/m/d"

    ' Sample row (delete after confirming)
    ws.Cells(2,COL_ID).Value   = "108699"
    ws.Cells(2,COL_NAME).Value = ChrW(26647) & ChrW(21407) & ChrW(12288) & ChrW(21083)  ' 栗原　剛
    ws.Cells(2,COL_KANA).Value = ChrW(12367) & ChrW(12426) & ChrW(12399) & ChrW(12425)  ' くりはら
    ws.Cells(2,COL_DEPT).Value = ChrW(32207) & ChrW(21209) & ChrW(35506)                ' 総務課
    ws.Cells(2,COL_JOB).Value  = ChrW(20107) & ChrW(21209) & ChrW(21729)                ' 事務員
    ws.Cells(2,COL_POS).Value  = ChrW(20027) & ChrW(20219)                              ' 主任
    ws.Cells(2,COL_JOIN).Value = DateSerial(2026, 4, 1)

    ws.Activate
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub CreateBadgeSheet()
    On Error Resume Next: Worksheets(SH_BADGE()).Delete: On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = SH_BADGE()
    With ws.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlPortrait
        .LeftMargin   = Application.CentimetersToPoints(0.8)
        .RightMargin  = Application.CentimetersToPoints(0.8)
        .TopMargin    = Application.CentimetersToPoints(1.0)
        .BottomMargin = Application.CentimetersToPoints(1.0)
        .CenterHorizontally = True: .PrintGridlines = False
    End With
    ws.Cells(1, 1).Select
End Sub

Private Sub ResetBadgeSheet()
    Application.DisplayAlerts = False
    On Error Resume Next: Worksheets(SH_BADGE()).Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    Call CreateBadgeSheet
End Sub

Private Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = Worksheets(sName): On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

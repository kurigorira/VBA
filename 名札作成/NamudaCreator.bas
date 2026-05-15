Attribute VB_Name = "NamudaCreator"
' Nagasaki Kita Tokushukai Hospital - Employee Badge Creator
' All Japanese text uses ChrW() so the .bas file imports correctly.
' Badge: 90mm wide x 55mm tall (landscape), Meiryo font, NW7 barcode at top
Option Explicit

'==============================================================
' Constants  (module-level declarations FIRST)
'==============================================================
Private Const BADGE_ROWS     As Integer = 14   ' ~55mm tall
Private Const BADGE_COLS     As Integer = 14   ' ~90mm wide
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

' Employee data array field indices
Private Const F_ID   As Integer = 0
Private Const F_NAME As Integer = 1
Private Const F_KANA As Integer = 2
Private Const F_DEPT As Integer = 3
Private Const F_JOB  As Integer = 4
Private Const F_POS  As Integer = 5
Private Const F_COLS As Integer = 6

'==============================================================
' PUBLIC SUBS  (assign to buttons on the sheet)
'==============================================================

' Run once to create the two worksheets
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

' -------------------------------------------------------
' Run this first if the barcode appears as plain text.
' It scans common NW-7 font name candidates and reports
' which one is actually installed on this PC.
' Then update FN_NW7() at the bottom of this module.
' -------------------------------------------------------
Public Sub DetectBarcodeFont()
    If Not SheetExists(SH_BADGE()) Then Call CreateBadgeSheet
    Dim ws As Worksheet: Set ws = Worksheets(SH_BADGE())
    Dim tmp As Range: Set tmp = ws.Range("A1")
    Dim orig As String: orig = tmp.Font.Name

    Dim cands() As String
    cands = Split( _
        "NW7,NW-7,NW 7,CodaBar,Codabar,CODABAR," & _
        "NW7Std,NW7 Std,Free NW7,NW7Bar," & _
        "NW-7 Barcode,NW7 Barcode,BarNW7," & _
        "Code NW7,IDAutomationSMNW7", ",")

    Dim found As String: found = ""
    Dim i As Integer
    For i = 0 To UBound(cands)
        tmp.Font.Name = cands(i)
        If LCase(Trim(tmp.Font.Name)) = LCase(Trim(cands(i))) Then
            found = found & "  " & cands(i) & vbCrLf
        End If
    Next i
    tmp.Font.Name = orig

    If found = "" Then
        MsgBox "NW-7フォントが見つかりません。" & vbCrLf & vbCrLf & _
               "確認方法:「ホーム」タブ　→「フォント」ドロップダウンを開き" & vbCrLf & _
               "『nw』または『barcode』で検索してフォント名を確認してください。" & vbCrLf & vbCrLf & _
               "その後このモジュール末尾の FN_NW7() の返り値を正しいフォント名に変更してください。", _
               vbExclamation, "NW7フォントが見つかりません"
    Else
        MsgBox "検出されたNW-7フォント:「" & vbCrLf & found & vbCrLf & _
               "FN_NW7()の返り値と一致している場合はそのままです。" & vbCrLf & _
               "違う場合は FN_NW7() を上記の名前に変更してください。", _
               vbInformation, "NW7フォント検出結果"
    End If
End Sub

' Create badges for rows selected in the master sheet
Public Sub CreateBadgesSelected()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub
    If ActiveSheet.Name <> mn Then
        MsgBox "Select rows in the " & mn & " sheet first.", vbExclamation: Exit Sub
    End If

    Dim selRows() As Long
    Dim cnt As Integer: cnt = 0
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
    Dim empArr() As String
    ReDim empArr(cnt - 1, F_COLS - 1)
    Dim validCnt As Integer: validCnt = 0
    Dim i As Integer
    For i = 0 To cnt - 1
        Dim r As Long: r = selRows(i)
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            Call ReadRow(ws, r, empArr, validCnt)
            ws.Cells(r, COL_MADE).Value = Date
            validCnt = validCnt + 1
        End If
    Next i
    If validCnt = 0 Then MsgBox "No valid ID found.", vbInformation: Exit Sub
    Call GenerateBadges(empArr, validCnt)
    MsgBox validCnt & " badge(s) created.", vbInformation
End Sub

' Create badges for all employees
Public Sub CreateBadgesAll()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(mn)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No data found.", vbInformation: Exit Sub

    Dim empArr() As String
    ReDim empArr(lastRow - 2, F_COLS - 1)
    Dim cnt As Integer: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            Call ReadRow(ws, r, empArr, cnt)
            cnt = cnt + 1
        End If
    Next r
    If cnt = 0 Then MsgBox "No data found.", vbInformation: Exit Sub
    If MsgBox("Create badges for all " & cnt & " employee(s)?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then ws.Cells(r, COL_MADE).Value = Date
    Next r
    Call GenerateBadges(empArr, cnt)
    MsgBox cnt & " badge(s) created.", vbInformation
End Sub

' Create badges filtered by join year/month (e.g. April batch)
Public Sub CreateBadgesByMonth()
    Dim mn As String: mn = SH_MASTER()
    If Not SheetExists(mn) Then MsgBox "Run Setup first.", vbExclamation: Exit Sub

    Dim inputVal As String
    inputVal = InputBox("Enter join year/month (e.g. 2026/4):", "Filter by Month", Format(Date, "yyyy/m"))
    If inputVal = "" Then Exit Sub
    inputVal = Replace(inputVal, "-", "/")
    Dim parts() As String: parts = Split(inputVal, "/")
    If UBound(parts) < 1 Then MsgBox "Use format: 2026/4", vbExclamation: Exit Sub
    Dim targetYear  As Integer: targetYear  = CInt(parts(0))
    Dim targetMonth As Integer: targetMonth = CInt(parts(1))

    Dim ws As Worksheet: Set ws = Worksheets(mn)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    Dim empArr() As String
    ReDim empArr(lastRow - 2, F_COLS - 1)
    Dim cnt As Integer: cnt = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) = "" Then GoTo NextRow
        Dim jd As Variant: jd = ws.Cells(r, COL_JOIN).Value
        If Not IsDate(jd) Then GoTo NextRow
        If Year(CDate(jd)) = targetYear And Month(CDate(jd)) = targetMonth Then
            Call ReadRow(ws, r, empArr, cnt)
            ws.Cells(r, COL_MADE).Value = Date
            cnt = cnt + 1
        End If
NextRow:
    Next r
    If cnt = 0 Then
        MsgBox "No employees for " & targetYear & "/" & targetMonth, vbInformation: Exit Sub
    End If
    If MsgBox("Create " & cnt & " badge(s) for " & targetYear & "/" & targetMonth & "?", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call GenerateBadges(empArr, cnt)
    MsgBox cnt & " badge(s) created.", vbInformation
End Sub

' Print preview
Public Sub BadgePrintPreview()
    If SheetExists(SH_BADGE()) Then
        Worksheets(SH_BADGE()).PrintPreview
    Else
        MsgBox "No badge sheet. Create badges first.", vbExclamation
    End If
End Sub

'==============================================================
' PRIVATE: core logic
'==============================================================

Private Sub ReadRow(ws As Worksheet, r As Long, empArr() As String, idx As Integer)
    empArr(idx, F_ID)   = CStr(ws.Cells(r, COL_ID).Value)
    empArr(idx, F_NAME) = CStr(ws.Cells(r, COL_NAME).Value)
    empArr(idx, F_KANA) = CStr(ws.Cells(r, COL_KANA).Value)
    empArr(idx, F_DEPT) = CStr(ws.Cells(r, COL_DEPT).Value)
    empArr(idx, F_JOB)  = CStr(ws.Cells(r, COL_JOB).Value)
    empArr(idx, F_POS)  = CStr(ws.Cells(r, COL_POS).Value)
End Sub

Private Sub GenerateBadges(empArr() As String, cnt As Integer)
    ' Verify NW7 font before drawing
    If Not FontExists(FN_NW7()) Then
        If MsgBox("NW-7バーコードフォント [" & FN_NW7() & "] が見つかりません。" & vbCrLf & vbCrLf & _
                  "「DetectBarcodeFont」マクロを実行して正しいフォント名を確認することをお勧めします。" & vbCrLf & vbCrLf & _
                  "そのまま続行しますか？", _
                  vbYesNo + vbExclamation, "NW7フォント未検出") = vbNo Then Exit Sub
    End If

    Call ResetBadgeSheet
    Dim wb As Worksheet: Set wb = Worksheets(SH_BADGE())
    Call SetAllDimensions(wb, cnt)
    Application.ScreenUpdating = False
    Dim i As Integer
    For i = 0 To cnt - 1
        Call DrawOneBadge(wb, i, _
            empArr(i, F_ID),   empArr(i, F_NAME), empArr(i, F_KANA), _
            empArr(i, F_DEPT), empArr(i, F_JOB),  empArr(i, F_POS))
    Next i
    Application.ScreenUpdating = True
    wb.Activate
    wb.Cells(1, 1).Select
End Sub

' Returns True if the named font is installed and recognised by Excel
Private Function FontExists(fontName As String) As Boolean
    If Not SheetExists(SH_BADGE()) Then Call CreateBadgeSheet
    Dim tmp As Range: Set tmp = Worksheets(SH_BADGE()).Range("A1")
    Dim orig As String: orig = tmp.Font.Name
    tmp.Font.Name = fontName
    FontExists = (LCase(Trim(tmp.Font.Name)) = LCase(Trim(fontName)))
    tmp.Font.Name = orig
End Function

' Set row heights and column widths
' Row  0 : top pad          2pt
' Row  1 : ID text         10pt
' Row 2-3 : NW7 barcode    13pt x2
' Row  4 : spacer           4pt
' Row 5-6 : dept            12pt x2
' Row  7 : job/pos         11pt
' Row 8-10: kana (large)   15pt x3
' Row 11 : kanji name      11pt
' Row12-13: footer         12pt x2  => total 157pt ~= 55mm
Private Sub SetAllDimensions(ws As Worksheet, cnt As Integer)
    Dim rowCount As Integer: rowCount = ((cnt - 1) \ BADGES_PER_ROW) + 1

    Dim rh(13) As Single
    rh(0)=2:   rh(1)=10:  rh(2)=13:  rh(3)=13:  rh(4)=4
    rh(5)=12:  rh(6)=12:  rh(7)=11:  rh(8)=15:  rh(9)=15
    rh(10)=15: rh(11)=11: rh(12)=12: rh(13)=12

    Dim rb As Integer, ri As Integer, br As Long
    For rb = 0 To rowCount - 1
        br = rb * (BADGE_ROWS + ROW_GAP) + 1
        For ri = 0 To BADGE_ROWS - 1
            ws.Rows(br + ri).RowHeight = rh(ri)
        Next ri
        If rb < rowCount - 1 Then ws.Rows(br + BADGE_ROWS).RowHeight = 4
    Next rb

    ' col 0=0.8, col 1-12=2.5 each, col 13=1.0  => ~90mm
    Dim cw(13) As Single
    cw(0) = 0.8: cw(13) = 1.0
    Dim ci As Integer
    For ci = 1 To 12
        cw(ci) = 2.5
    Next ci

    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + ci).ColumnWidth = cw(ci)
    Next ci
    ws.Columns(1 + BADGE_COLS).ColumnWidth = 1
    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + BADGE_COLS + COL_GAP + ci).ColumnWidth = cw(ci)
    Next ci
End Sub

Private Sub DrawOneBadge(ws As Worksheet, badgeIdx As Integer, _
                          empID As String, empName As String, empKana As String, _
                          dept  As String, jobType As String, pos    As String)
    Dim sRow As Long: sRow = (badgeIdx \ BADGES_PER_ROW) * (BADGE_ROWS + ROW_GAP) + 1
    Dim sCol As Long: sCol = (badgeIdx Mod BADGES_PER_ROW) * (BADGE_COLS + COL_GAP) + 1
    Call DrawBadgeContent(ws, sRow, sCol, empID, empName, empKana, dept, jobType, pos)
End Sub

' Badge layout (90mm wide x 55mm tall, landscape):
'  +--------------------------------------------------+
'  | 108699 (right, small)                   row +1  |
'  | ======= NW7 barcode (full width) =====  rows+2,3|
'  | 部署名                               rows+5,6  |
'  | 職種　役職                           row  +7  |
'  | くりはら (large)                   rows+8~10 |
'  | 栗原 剛                              row  +11  |
'  +--[logo]--医療法人　徳洲会----------rows+12,13|
'  +--------------------------------------------------+
Private Sub DrawBadgeContent(ws As Worksheet, sRow As Long, sCol As Long, _
                              empID As String, empName As String, empKana As String, _
                              dept  As String, jobType As String, pos    As String)
    Dim fr As Range
    Set fr = ws.Range(ws.Cells(sRow, sCol), ws.Cells(sRow + BADGE_ROWS - 1, sCol + BADGE_COLS - 1))
    fr.UnMerge: fr.ClearContents: fr.ClearFormats
    fr.Interior.Color = RGB(255, 255, 255)

    Dim fc As Long: fc = RGB(80, 80, 160)
    With fr.Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: .Color = fc: End With
    With fr.Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: .Color = fc: End With
    With fr.Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: .Color = fc: End With
    With fr.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = fc: End With

    Dim fn As String: fn = FN_MEIRYO()

    ' --- Row +1 : ID number (human-readable, right-aligned) ---
    Call MC(ws, sRow+1, sCol, 1, BADGE_COLS, empID, 8, False, fn, RGB(80,80,160), RGB(255,255,255), xlRight, xlCenter)

    ' --- Rows +2,+3 : NW7 barcode (full width) ---
    ' NW-7 (Codabar) format: start-char A + digits + stop-char A
    ' If barcode still shows as text, run DetectBarcodeFont() to find the correct font name,
    ' then update FN_NW7() at the bottom of this module.
    Dim bcR As Range
    Set bcR = ws.Range(ws.Cells(sRow+2, sCol), ws.Cells(sRow+3, sCol+BADGE_COLS-1))
    bcR.Merge
    bcR.Value = "A" & empID & "A"
    bcR.Font.Name = FN_NW7()
    bcR.Font.Size = 20
    bcR.Font.Color = RGB(0, 0, 0): bcR.Interior.Color = RGB(255, 255, 255)
    bcR.HorizontalAlignment = xlCenter: bcR.VerticalAlignment = xlCenter
    bcR.WrapText = False

    ' --- Main content ---
    Call MC(ws, sRow+5, sCol, 2, BADGE_COLS, dept,    13, True,  fn, RGB(0,0,0),    RGB(255,255,255), xlLeft,   xlCenter)

    Dim jp As String: jp = jobType
    If Trim(pos) <> "" Then jp = jp & ChrW(12288) & pos
    Call MC(ws, sRow+7, sCol, 1, BADGE_COLS, jp,      10, False, fn, RGB(0,0,0),    RGB(255,255,255), xlLeft,   xlCenter)

    Call MC(ws, sRow+8, sCol, 3, BADGE_COLS, empKana, 24, True,  fn, RGB(0,0,0),    RGB(255,255,255), xlCenter, xlCenter)
    Call MC(ws, sRow+11, sCol, 1, BADGE_COLS, empName, 10, False, fn, RGB(50,50,50), RGB(255,255,255), xlCenter, xlCenter)

    ' --- Footer (rows +12,+13) ---
    Dim fbg As Long: fbg = RGB(214, 234, 248)
    Dim fa As Range
    Set fa = ws.Range(ws.Cells(sRow+12, sCol), ws.Cells(sRow+13, sCol+BADGE_COLS-1))
    fa.Interior.Color = fbg
    With fa.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Color = RGB(80, 80, 160): End With

    Dim lr As Range
    Set lr = ws.Range(ws.Cells(sRow+12, sCol), ws.Cells(sRow+13, sCol+2))
    lr.Merge: lr.Interior.Color = fbg
    With lr.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Color = RGB(150, 150, 200): End With

    Call MC(ws, sRow+12, sCol+3, 1, BADGE_COLS-3, STR_CORP(),     8,  False, fn, RGB(0,0,0), fbg, xlLeft, xlCenter)
    Call MC(ws, sRow+13, sCol+3, 1, BADGE_COLS-3, STR_HOSPITAL(), 11, True,  fn, RGB(0,0,0), fbg, xlLeft, xlCenter)

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
    Set ar = ws.Range(ws.Cells(sRow+12, sCol), ws.Cells(sRow+13, sCol+2))
    On Error GoTo Ex
    Dim p As Shape
    Set p = ws.Shapes.AddPicture(Filename:=lp, LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, Left:=ar.Left+2, Top:=ar.Top+2, Width:=-1, Height:=ar.Height-4)
    p.LockAspectRatio = msoTrue
    p.Name = "Logo_R" & sRow & "C" & sCol
Ex: On Error GoTo 0
End Sub

'==============================================================
' PRIVATE: sheet creation / helpers
'==============================================================

Private Sub CreateMasterSheet()
    On Error Resume Next: Worksheets(SH_MASTER()).Delete: On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets.Add(Before:=Worksheets(1))
    ws.Name = SH_MASTER()

    ws.Cells(1,1).Value = ChrW(32887) & ChrW(21729) & "ID"
    ws.Cells(1,2).Value = ChrW(27663) & ChrW(21517) & ChrW(65288) & ChrW(28450) & ChrW(23383) & ChrW(65289)
    ws.Cells(1,3).Value = ChrW(33495) & ChrW(23383) & ChrW(12402) & ChrW(12425) & ChrW(12364) & ChrW(12394)
    ws.Cells(1,4).Value = ChrW(37096) & ChrW(32626) & ChrW(21517)
    ws.Cells(1,5).Value = ChrW(32887) & ChrW(31278)
    ws.Cells(1,6).Value = ChrW(24441) & ChrW(32887)
    ws.Cells(1,7).Value = ChrW(20837) & ChrW(32887) & ChrW(26085)
    ws.Cells(1,8).Value = ChrW(21517) & ChrW(26413) & ChrW(20316) & ChrW(25104) & ChrW(26085)

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

    ws.Cells(2,COL_ID).Value   = "108699"
    ws.Cells(2,COL_NAME).Value = ChrW(26647) & ChrW(21407) & ChrW(12288) & ChrW(21083)
    ws.Cells(2,COL_KANA).Value = ChrW(12367) & ChrW(12426) & ChrW(12399) & ChrW(12425)
    ws.Cells(2,COL_DEPT).Value = ChrW(32207) & ChrW(21209) & ChrW(35506)
    ws.Cells(2,COL_JOB).Value  = ChrW(20107) & ChrW(21209) & ChrW(21729)
    ws.Cells(2,COL_POS).Value  = ChrW(20027) & ChrW(20219)
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

'--------------------------------------------------------------
' Japanese string helpers  (ChrW = encoding-independent)
'--------------------------------------------------------------
Private Function SH_MASTER() As String
    SH_MASTER = ChrW(32887) & ChrW(21729) & ChrW(12510) & ChrW(12473) & ChrW(12479) & ChrW(12540)
End Function
Private Function SH_BADGE() As String
    SH_BADGE = ChrW(21517) & ChrW(26413) & ChrW(21360) & ChrW(21047)
End Function
Private Function FN_MEIRYO() As String
    FN_MEIRYO = "Meiryo"
End Function

' *** NW-7 barcode font name ***
' If the barcode shows as plain text (e.g. A108699A), run DetectBarcodeFont()
' to find the correct name, then change the return value below.
Private Function FN_NW7() As String
    FN_NW7 = "NW7"
End Function

Private Function STR_CORP() As String
    STR_CORP = ChrW(21307) & ChrW(30274) & ChrW(27861) & ChrW(20154) & ChrW(12288) & ChrW(24499) & ChrW(27954) & ChrW(20250)
End Function
Private Function STR_HOSPITAL() As String
    STR_HOSPITAL = ChrW(38263) & ChrW(23822) & ChrW(21271) & ChrW(24499) & ChrW(27954) & ChrW(20250) & ChrW(30149) & ChrW(38498)
End Function

Attribute VB_Name = "NamudaCreator"
'==============================================================
' 長崎北徳洲会病院 職員名札作成システム v2.2
' レイアウト：横型 / A4縦に2枚幅
'==============================================================
Option Explicit

' ---- シート名 ------------------------------------------------
Private Const MASTER_SHEET   As String = "職員マスター"
Private Const BADGE_SHEET    As String = "名札印刷"

' ---- 名札1枚のセル数 ------------------------------------
Private Const BADGE_ROWS     As Integer = 14   ' 縦 14行 ≈ 54mm
Private Const BADGE_COLS     As Integer = 11   ' 横 11列 ≈ 86mm
Private Const BADGES_PER_ROW As Integer = 2    ' A4に横2枚並べる
Private Const COL_GAP        As Integer = 1    ' 名札間の列ギャップ
Private Const ROW_GAP        As Integer = 1    ' 名札間の行ギャップ

' ---- マスターシート列番号 ------------------------------
Private Const COL_ID   As Integer = 1
Private Const COL_NAME As Integer = 2
Private Const COL_KANA As Integer = 3
Private Const COL_DEPT As Integer = 4
Private Const COL_JOB  As Integer = 5
Private Const COL_POS  As Integer = 6
Private Const COL_JOIN As Integer = 7
Private Const COL_MADE As Integer = 8

'==============================================================
' [公開] 初回セットアップ
'==============================================================
Public Sub ワークブック初期化()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call CreateMasterSheet
    Call CreateBadgeSheet
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Worksheets(MASTER_SHEET).Activate
    Worksheets(MASTER_SHEET).Cells(2, 1).Select
    MsgBox "初期化完了。" & Chr(10) & _
           MASTER_SHEET & " シートに職員情報を入力してください。", vbInformation
End Sub

'==============================================================
' [公開] 選択した行の職員の名札を作成
' 操作：職員マスターシートで対象行（1行または複数行）を選択→実行
'==============================================================
Public Sub 名札作成_選択職員()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に[ワークブック初期化]を実行してください。", vbExclamation: Exit Sub
    End If
    If ActiveSheet.Name <> MASTER_SHEET Then
        MsgBox MASTER_SHEET & " シートで対象行を選択してから実行してください。", vbExclamation: Exit Sub
    End If

    ' 選択行を重複なしで収集
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

    If cnt = 0 Then
        MsgBox "2行目以降のデータ行を選択してから実行してください。", vbInformation: Exit Sub
    End If

    Dim ws As Worksheet: Set ws = Worksheets(MASTER_SHEET)
    Dim empList() As EmpData
    ReDim empList(cnt - 1)
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

    If validCnt = 0 Then
        MsgBox "職員IDが入力されている行がありません。", vbInformation: Exit Sub
    End If

    Call GenerateBadges(empList, validCnt)
    MsgBox validCnt & " 枚の名札を作成しました。", vbInformation
End Sub

'==============================================================
' [公開] 全職員の名札を一括作成
'==============================================================
Public Sub 名札作成_全職員()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に[ワークブック初期化]を実行してください。", vbExclamation: Exit Sub
    End If
    Dim ws As Worksheet: Set ws = Worksheets(MASTER_SHEET)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "職員データがありません。", vbInformation: Exit Sub
    End If

    Dim empList() As EmpData
    Dim cnt As Integer: cnt = 0
    ReDim empList(lastRow - 2)
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            empList(cnt) = ReadEmpRow(ws, r)
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then
        MsgBox "職員データがありません。", vbInformation: Exit Sub
    End If
    If MsgBox(cnt & " 名全員の名札を作成しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' 作成日を一括記録
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            ws.Cells(r, COL_MADE).Value = Date
        End If
    Next r

    Call GenerateBadges(empList, cnt)
    MsgBox cnt & " 枚の名札を作成しました。", vbInformation
End Sub

'==============================================================
' [公開] 入職年月で絞り込んで作成（4月一括などに便利）
'==============================================================
Public Sub 名札作成_入職年月指定()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に[ワークブック初期化]を実行してください。", vbExclamation: Exit Sub
    End If

    ' 年月入力ダイアログ
    Dim inputVal As String
    inputVal = InputBox("入職年月を入力してください。" & Chr(10) & _
                       "例） 2026/4  または 2026-04", _
                       "入職年月小指定", _
                       Format(Date, "yyyy/m"))
    If inputVal = "" Then Exit Sub

    ' 年・月を解析
    inputVal = Replace(inputVal, "-", "/")
    Dim parts() As String: parts = Split(inputVal, "/")
    If UBound(parts) < 1 Then
        MsgBox "入力形式が正しくありません。例） 2026/4", vbExclamation: Exit Sub
    End If
    Dim targetYear  As Integer: targetYear  = CInt(parts(0))
    Dim targetMonth As Integer: targetMonth = CInt(parts(1))

    Dim ws As Worksheet: Set ws = Worksheets(MASTER_SHEET)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    Dim empList() As EmpData
    Dim cnt As Integer: cnt = 0
    ReDim empList(lastRow - 2)
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) = "" Then GoTo NextRow
        Dim joinDate As Variant: joinDate = ws.Cells(r, COL_JOIN).Value
        If Not IsDate(joinDate) Then GoTo NextRow
        If Year(CDate(joinDate)) = targetYear And Month(CDate(joinDate)) = targetMonth Then
            empList(cnt) = ReadEmpRow(ws, r)
            ws.Cells(r, COL_MADE).Value = Date
            cnt = cnt + 1
        End If
NextRow:
    Next r

    If cnt = 0 Then
        MsgBox targetYear & "年" & targetMonth & "月入職の職員が見つかりません。", vbInformation: Exit Sub
    End If
    If MsgBox(targetYear & "年" & targetMonth & "月入職　" & cnt & " 名分の名札を作成しますか？", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Call GenerateBadges(empList, cnt)
    MsgBox cnt & " 枚の名札を作成しました。", vbInformation
End Sub

'==============================================================
' [公開] 印刷プレビュー
'==============================================================
Public Sub 印刷プレビュー()
    If SheetExists(BADGE_SHEET) Then
        Worksheets(BADGE_SHEET).PrintPreview
    Else
        MsgBox "名札印刷シートがありません。先に名札を作成してください。", vbExclamation
    End If
End Sub

'==============================================================
' 職員データが扱いやすいようデータ型でまとめる
'==============================================================
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
' 名札シートをリセットして複数人分の名札を一括生成
'==============================================================
Private Sub GenerateBadges(empList() As EmpData, cnt As Integer)
    Call ResetBadgeSheet
    Dim wb As Worksheet: Set wb = Worksheets(BADGE_SHEET)

    ' 列幅・行高は全名札共通で先に一括設定
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

'==============================================================
' 全準敗の行高・列幅を一括設定（内容描画前に呼ぶ）
'==============================================================
Private Sub SetAllDimensions(ws As Worksheet, cnt As Integer)
    Dim rowCount As Integer: rowCount = (((cnt - 1) \ BADGES_PER_ROW) + 1)
    Dim totalRows As Long
    totalRows = rowCount * (BADGE_ROWS + ROW_GAP)

    ' 行高パターン
    Dim rh(13) As Single
    rh(0)  = 2 : rh(1)  = 13 : rh(2)  = 13 : rh(3)  = 11
    rh(4)  = 15 : rh(5)  = 15 : rh(6)  = 15 : rh(7)  = 15
    rh(8)  = 12 : rh(9)  = 5
    rh(10) = 11 : rh(11) = 11 : rh(12) = 11 : rh(13) = 2

    Dim rowBlock As Integer, ri As Integer
    Dim baseRow As Long
    For rowBlock = 0 To rowCount - 1
        baseRow = rowBlock * (BADGE_ROWS + ROW_GAP) + 1
        For ri = 0 To BADGE_ROWS - 1
            ws.Rows(baseRow + ri).RowHeight = rh(ri)
        Next ri
        ' ギャップ行
        If rowBlock < rowCount - 1 Then
            ws.Rows(baseRow + BADGE_ROWS).RowHeight = 6
        End If
    Next rowBlock

    ' 列幅：左エリア、0ガップ7　1右エリア　9
    ' 左列（名札1枚目）
    Dim cw(10) As Single
    cw(0) = 1 : cw(1) = 3.5 : cw(2) = 3.5 : cw(3) = 3.5
    cw(4) = 3.5 : cw(5) = 3.5 : cw(6) = 3
    cw(7) = 0.5 : cw(8) = 3 : cw(9) = 3 : cw(10) = 1.5

    Dim ci As Integer
    ' 左列ブロック (cols 1〜11)
    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + ci).ColumnWidth = cw(ci)
    Next ci
    ' 左右のギャップ列 (col 12)
    ws.Columns(1 + BADGE_COLS).ColumnWidth = 1
    ' 右列ブロック (cols 13〜23)
    For ci = 0 To BADGE_COLS - 1
        ws.Columns(1 + BADGE_COLS + COL_GAP + ci).ColumnWidth = cw(ci)
    Next ci
End Sub

'==============================================================
' 1枚の名札を描画
'==============================================================
Private Sub DrawOneBadge(ws As Worksheet, badgeIdx As Integer, emp As EmpData)
    Dim colIdx As Integer: colIdx = badgeIdx Mod BADGES_PER_ROW
    Dim rowIdx As Integer: rowIdx = badgeIdx \ BADGES_PER_ROW
    Dim sRow As Long: sRow = rowIdx * (BADGE_ROWS + ROW_GAP) + 1
    Dim sCol As Long: sCol = colIdx * (BADGE_COLS + COL_GAP) + 1

    Call DrawBadgeContent(ws, sRow, sCol, emp)
End Sub

'--------------------------------------------------------------
' 名札の内容を描画
'--------------------------------------------------------------
Private Sub DrawBadgeContent(ws As Worksheet, sRow As Long, sCol As Long, emp As EmpData)
    ' 全体クリア
    Dim fullRng As Range
    Set fullRng = ws.Range(ws.Cells(sRow, sCol), _
                            ws.Cells(sRow + BADGE_ROWS - 1, sCol + BADGE_COLS - 1))
    fullRng.UnMerge
    fullRng.ClearContents
    fullRng.ClearFormats
    fullRng.Interior.Color = RGB(255, 255, 255)

    ' 外枠（青紫色）
    Dim fc As Long: fc = RGB(80, 80, 160)
    With fullRng
        .Borders(xlEdgeLeft).LineStyle    = xlContinuous
        .Borders(xlEdgeLeft).Weight       = xlMedium
        .Borders(xlEdgeLeft).Color        = fc
        .Borders(xlEdgeRight).LineStyle   = xlContinuous
        .Borders(xlEdgeRight).Weight      = xlMedium
        .Borders(xlEdgeRight).Color       = fc
        .Borders(xlEdgeTop).LineStyle     = xlContinuous
        .Borders(xlEdgeTop).Weight        = xlMedium
        .Borders(xlEdgeTop).Color         = fc
        .Borders(xlEdgeBottom).LineStyle  = xlContinuous
        .Borders(xlEdgeBottom).Weight     = xlMedium
        .Borders(xlEdgeBottom).Color      = fc
    End With

    ' ===== 左エリア (cols 0-6) ======================

    ' 1. 部署名（行+1〜+2）
    Call MC(ws, sRow + 1, sCol, 2, 7, emp.Dept, 14, True, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlLeft, xlCenter)

    ' 2. 職種・役職（行+3）
    Dim jobPos As String
    jobPos = emp.JobType
    If Trim(emp.Pos) <> "" Then jobPos = jobPos & "　" & emp.Pos
    Call MC(ws, sRow + 3, sCol, 1, 7, jobPos, 10, False, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlLeft, xlCenter)

    ' 3. 苗字ひらがな（行+4〜+7）
    Call MC(ws, sRow + 4, sCol, 4, 7, emp.Kana, 32, True, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlCenter, xlCenter)

    ' 4. 氏名（漢字）（行+8）
    Call MC(ws, sRow + 8, sCol, 1, 7, emp.Name, 10, False, "MS P明朝", _
            RGB(50, 50, 50), RGB(255, 255, 255), xlCenter, xlCenter)

    ' ===== 右エリア: 空白（行+0〜+3）+ バーコード（行+4〜+8）=====
    Dim RC As Long: RC = sCol + 7

    ' 右上: 空白
    ws.Range(ws.Cells(sRow, RC), _
             ws.Cells(sRow + 3, sCol + BADGE_COLS - 1)).Merge

    ' 右下: バーコード
    Dim bcRng As Range
    Set bcRng = ws.Range(ws.Cells(sRow + 4, RC), _
                          ws.Cells(sRow + 8, sCol + BADGE_COLS - 1))
    bcRng.Merge
    With bcRng
        .Interior.Color      = RGB(255, 255, 255)
        .Value               = "| || ||| || | ||| || | ||"
        .Font.Size           = 7
        .Font.Name           = "Courier New"
        .Font.Color          = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    ' ===== 青バー + ID（行+9） ======================

    ' 青バー（cols 0-8）
    Dim barRng As Range
    Set barRng = ws.Range(ws.Cells(sRow + 9, sCol), _
                           ws.Cells(sRow + 9, sCol + 8))
    barRng.Merge
    barRng.Interior.Color = RGB(42, 107, 183)

    ' ID番号（cols 9-10）
    Dim idRng As Range
    Set idRng = ws.Range(ws.Cells(sRow + 9, sCol + 9), _
                          ws.Cells(sRow + 9, sCol + BADGE_COLS - 1))
    idRng.Merge
    With idRng
        .Value               = emp.ID
        .Font.Size           = 8
        .Font.Name           = "Arial"
        .Font.Color          = RGB(0, 0, 0)
        .Interior.Color      = RGB(255, 255, 255)
        .HorizontalAlignment = xlRight
        .VerticalAlignment   = xlCenter
    End With

    ' ===== フッター（行+10〜+12） =======================
    Dim footBG As Long: footBG = RGB(214, 234, 248)

    Dim footAll As Range
    Set footAll = ws.Range(ws.Cells(sRow + 10, sCol), _
                            ws.Cells(sRow + 12, sCol + BADGE_COLS - 1))
    footAll.Interior.Color = footBG
    footAll.Borders(xlEdgeTop).LineStyle = xlContinuous
    footAll.Borders(xlEdgeTop).Color     = RGB(80, 80, 160)

    ' ロゴエリア（左 3列）
    Dim logoRng As Range
    Set logoRng = ws.Range(ws.Cells(sRow + 10, sCol), _
                            ws.Cells(sRow + 12, sCol + 2))
    logoRng.Merge
    logoRng.Interior.Color = footBG
    logoRng.Borders(xlEdgeRight).LineStyle = xlContinuous
    logoRng.Borders(xlEdgeRight).Color     = RGB(150, 150, 200)

    ' 医療法人徳洲会
    Call MC(ws, sRow + 10, sCol + 3, 1, 8, "医療法人　徳洲会", 8, False, "MS P明朝", _
            RGB(0, 0, 0), footBG, xlLeft, xlCenter)

    ' 長崎北徳洲会病院
    Call MC(ws, sRow + 11, sCol + 3, 2, 8, "長崎北徳洲会病院", 13, True, "MS P明朝", _
            RGB(0, 0, 0), footBG, xlLeft, xlCenter)

    ' ロゴ画像挿入（tokushukai_logo.png が同フォルダにあれば）
    Call TryInsertLogo(ws, sRow, sCol)
End Sub

'--------------------------------------------------------------
' セルマージ・書式設定のヘルパー
'--------------------------------------------------------------
Private Sub MC(ws As Worksheet, _
               sr As Long, sc As Long, nr As Integer, nc As Integer, _
               val As String, sz As Integer, bold As Boolean, fn As String, _
               fc As Long, bg As Long, ha As XlHAlign, va As XlVAlign)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(sr, sc), ws.Cells(sr + nr - 1, sc + nc - 1))
    rng.Merge
    rng.Value               = val
    rng.Font.Size           = sz
    rng.Font.Bold           = bold
    rng.Font.Name           = fn
    rng.Font.Color          = fc
    rng.Interior.Color      = bg
    rng.HorizontalAlignment = ha
    rng.VerticalAlignment   = va
End Sub

'--------------------------------------------------------------
' 徳洲会ロゴ画像をフッターに挿入
' 「ブックと同じフォルダ」に tokushukai_logo.png を配置すると自動挿入
'--------------------------------------------------------------
Private Sub TryInsertLogo(ws As Worksheet, sRow As Long, sCol As Long)
    Dim logoPath As String
    logoPath = ThisWorkbook.Path & "\tokushukai_logo.png"
    If Dir(logoPath) = "" Then Exit Sub

    Dim ar As Range
    Set ar = ws.Range(ws.Cells(sRow + 10, sCol), ws.Cells(sRow + 12, sCol + 2))
    On Error GoTo ErrExit
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=logoPath, LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, _
        Left:=ar.Left + 2, Top:=ar.Top + 2, _
        Width:=-1, Height:=ar.Height - 4)
    pic.LockAspectRatio = msoTrue
    pic.Name = "Logo_R" & sRow & "C" & sCol
ErrExit:
    On Error GoTo 0
End Sub

'==============================================================
' 職員マスターシートを新規作成
'==============================================================
Private Sub CreateMasterSheet()
    On Error Resume Next
    Worksheets(MASTER_SHEET).Delete
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = Worksheets.Add(Before:=Worksheets(1))
    ws.Name = MASTER_SHEET

    Dim hdr As Variant
    hdr = Array("職員ID", "氏名（漢字）", "苗字ひらがな", "部署名", "職種", "役職", "入職日", "名札作成日")
    Dim i As Integer
    For i = 0 To UBound(hdr)
        ws.Cells(1, i + 1).Value = hdr(i)
    Next i
    With ws.Range("A1:H1")
        .Font.Bold           = True
        .Font.Color          = RGB(255, 255, 255)
        .Interior.Color      = RGB(42, 107, 183)
        .HorizontalAlignment = xlCenter
        .RowHeight           = 22
    End With
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("G").NumberFormat = "yyyy/m/d"
    ws.Columns("H").NumberFormat = "yyyy/m/d"

    ' サンプル
    ws.Cells(2, COL_ID).Value   = "108699"
    ws.Cells(2, COL_NAME).Value = "栗原　剛"
    ws.Cells(2, COL_KANA).Value = "くりはら"
    ws.Cells(2, COL_DEPT).Value = "総務課"
    ws.Cells(2, COL_JOB).Value  = "事務員"
    ws.Cells(2, COL_POS).Value  = "主任"
    ws.Cells(2, COL_JOIN).Value = DateSerial(2026, 4, 1)

    ws.Activate
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

'==============================================================
' 名札印刷シートを新規作成
'==============================================================
Private Sub CreateBadgeSheet()
    On Error Resume Next
    Worksheets(BADGE_SHEET).Delete
    On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = BADGE_SHEET
    With ws.PageSetup
        .PaperSize          = xlPaperA4
        .Orientation        = xlPortrait
        .LeftMargin         = Application.CentimetersToPoints(0.8)
        .RightMargin        = Application.CentimetersToPoints(0.8)
        .TopMargin          = Application.CentimetersToPoints(1.0)
        .BottomMargin       = Application.CentimetersToPoints(1.0)
        .CenterHorizontally = True
        .PrintGridlines     = False
    End With
    ws.Cells(1, 1).Select
End Sub

'==============================================================
' 名札印刷シートをリセット
'==============================================================
Private Sub ResetBadgeSheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(BADGE_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Call CreateBadgeSheet
End Sub

'==============================================================
' シートの存在確認
'==============================================================
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

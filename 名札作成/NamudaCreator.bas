Attribute VB_Name = "NamudaCreator"
'==============================================================
' モジュール名 : NamudaCreator
' 説明         : 長崎北徳洲会病院 職員名札作成システム
' 対象         : Excel 2016以降
' 更新         : 2026-05-15
'==============================================================

Option Explicit

'--------------------------------------------------------------
' 設定定数（必要に応じて変更してください）
'--------------------------------------------------------------
Private Const MASTER_SHEET  As String = "職員マスター"
Private Const BADGE_SHEET   As String = "名札印刷"

' 名札1枚あたりのセル数
Private Const BADGE_ROWS    As Integer = 22   ' 縦22行
Private Const BADGE_COLS    As Integer = 6    ' 横 6列
Private Const BADGES_PER_ROW As Integer = 2   ' 横に並べる枚数
Private Const COL_GAP       As Integer = 1    ' 名札間の列ギャップ
Private Const ROW_GAP       As Integer = 1    ' 名札間の行ギャップ

' 徳洲会カラー（緑系）
Private Const COL_MAIN      As Long = 4423780 ' RGB(68,164,100)  ヘッダー・枠
Private Const COL_DARK      As Long = 2277173 ' RGB(53,130,34)   区切り線
Private Const COL_BG_LIGHT  As Long = 15859179 ' RGB(235,250,241) フッター背景

' マスターシート列定義
Private Const COL_ID        As Integer = 1   ' 職員ID
Private Const COL_NAME      As Integer = 2   ' 氏名（漢字）
Private Const COL_KANA      As Integer = 3   ' 苗字ひらがな
Private Const COL_DEPT      As Integer = 4   ' 部署名
Private Const COL_JOB       As Integer = 5   ' 職種
Private Const COL_POS       As Integer = 6   ' 役職
Private Const COL_JOIN      As Integer = 7   ' 入職日
Private Const COL_MADE      As Integer = 8   ' 名札作成日

'==============================================================
' 【公開】ワークブック初期化（初回のみ実行）
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

    MsgBox "初期化が完了しました。" & Chr(10) & _
           MASTER_SHEET & " シートに職員情報を入力してください。", vbInformation
End Sub

'==============================================================
' 【公開】選択した職員行の名札を作成
'         マスターシートで対象行を選択してから実行
'==============================================================
Public Sub 名札作成_選択職員()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に「ワークブック初期化」を実行してください。", vbExclamation
        Exit Sub
    End If
    If ActiveSheet.Name <> MASTER_SHEET Then
        MsgBox MASTER_SHEET & " シートで対象行を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Dim ws      As Worksheet
    Dim wsBadge As Worksheet
    Set ws = Worksheets(MASTER_SHEET)

    ' 選択行を収集（ヘッダー除外・重複除去）
    Dim selRows() As Long
    Dim cnt       As Integer
    cnt = 0

    Dim cel As Range
    For Each cel In Selection
        If cel.Row > 1 Then
            Dim dup As Boolean
            dup = False
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
        MsgBox "データ行（2行目以降）を選択してから実行してください。", vbInformation
        Exit Sub
    End If

    Call ResetBadgeSheet
    Set wsBadge = Worksheets(BADGE_SHEET)

    Application.ScreenUpdating = False

    Dim i As Integer
    For i = 0 To cnt - 1
        Dim r As Long
        r = selRows(i)
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) = "" Then GoTo NextEmp

        Call DrawOneBadge(wsBadge, i, _
            CStr(ws.Cells(r, COL_ID).Value), _
            CStr(ws.Cells(r, COL_NAME).Value), _
            CStr(ws.Cells(r, COL_KANA).Value), _
            CStr(ws.Cells(r, COL_DEPT).Value), _
            CStr(ws.Cells(r, COL_JOB).Value), _
            CStr(ws.Cells(r, COL_POS).Value))

        ws.Cells(r, COL_MADE).Value = Date   ' 作成日を記録

NextEmp:
    Next i

    Application.ScreenUpdating = True
    wsBadge.Activate
    wsBadge.Cells(1, 1).Select

    MsgBox cnt & " 枚の名札を作成しました。" & Chr(10) & _
           BADGE_SHEET & " シートを確認してください。", vbInformation
End Sub

'==============================================================
' 【公開】全職員の名札を一括作成
'==============================================================
Public Sub 名札作成_全職員()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に「ワークブック初期化」を実行してください。", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Worksheets(MASTER_SHEET)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "職員データがありません。", vbInformation
        Exit Sub
    End If

    If MsgBox(lastRow - 1 & " 名全員の名札を作成しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Call ResetBadgeSheet
    Dim wsBadge As Worksheet
    Set wsBadge = Worksheets(BADGE_SHEET)

    Application.ScreenUpdating = False

    Dim badgeNum As Integer
    badgeNum = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) = "" Then GoTo Skip

        Call DrawOneBadge(wsBadge, badgeNum, _
            CStr(ws.Cells(r, COL_ID).Value), _
            CStr(ws.Cells(r, COL_NAME).Value), _
            CStr(ws.Cells(r, COL_KANA).Value), _
            CStr(ws.Cells(r, COL_DEPT).Value), _
            CStr(ws.Cells(r, COL_JOB).Value), _
            CStr(ws.Cells(r, COL_POS).Value))

        ws.Cells(r, COL_MADE).Value = Date
        badgeNum = badgeNum + 1
Skip:
    Next r

    Application.ScreenUpdating = True
    wsBadge.Activate
    wsBadge.Cells(1, 1).Select

    MsgBox badgeNum & " 枚の名札を作成しました。", vbInformation
End Sub

'==============================================================
' 【公開】印刷プレビュー
'==============================================================
Public Sub 印刷プレビュー()
    If SheetExists(BADGE_SHEET) Then
        Worksheets(BADGE_SHEET).PrintPreview
    Else
        MsgBox "名札印刷シートがありません。先に名札を作成してください。", vbExclamation
    End If
End Sub

'==============================================================
' 1枚の名札を描画
' badgeIdx : 0始まりのインデックス（位置計算に使用）
'==============================================================
Private Sub DrawOneBadge(ws As Worksheet, badgeIdx As Integer, _
                         empID   As String, empName  As String, empKana As String, _
                         dept    As String, jobType  As String, pos     As String)

    ' A4縦2列配置：左右の開始列を交互に割り当て
    Dim colIdx As Integer  ' 0=左, 1=右
    Dim rowIdx As Integer  ' 0,1,2...
    colIdx = badgeIdx Mod BADGES_PER_ROW
    rowIdx = badgeIdx \ BADGES_PER_ROW

    Dim sRow As Long
    Dim sCol As Long
    sRow = rowIdx * (BADGE_ROWS + ROW_GAP) + 1
    sCol = colIdx * (BADGE_COLS + COL_GAP) + 1

    Call SetCellLayout(ws, sRow, sCol)
    Call DrawBadgeContent(ws, sRow, sCol, empID, empName, empKana, dept, jobType, pos)
End Sub

'--------------------------------------------------------------
' 行の高さ・列の幅を設定（名札1枚分）
'--------------------------------------------------------------
Private Sub SetCellLayout(ws As Worksheet, sRow As Long, sCol As Long)
    ' 行の高さ（合計 ≈ 86mm）
    ' インデックス 0-21 で BADGE_ROWS=22 行
    Dim rh(21) As Single
    rh(0)  = 14 : rh(1)  = 14  ' ヘッダーバー（2行）
    rh(2)  = 18 : rh(3)  = 18 : rh(4) = 18  ' 部署名（3行）
    rh(5)  = 3                               ' 区切り線
    rh(6)  = 18                              ' 職種
    rh(7)  = 16                              ' 役職
    rh(8)  = 24 : rh(9)  = 24 : rh(10) = 24 ' ひらがな（7行）
    rh(11) = 24 : rh(12) = 24 : rh(13) = 24
    rh(14) = 24
    rh(15) = 15 : rh(16) = 15               ' 氏名（2行）
    rh(17) = 12                              ' ID
    rh(18) = 3                               ' 区切り線
    rh(19) = 14 : rh(20) = 14 : rh(21) = 10 ' ロゴ・法人名（3行）

    Dim i As Integer
    For i = 0 To BADGE_ROWS - 1
        ws.Rows(sRow + i).RowHeight = rh(i)
    Next i

    ' 列幅（合計 ≈ 54mm）インデックス 0-5 で BADGE_COLS=6 列
    Dim cw(5) As Single
    cw(0) = 2 : cw(1) = 10 : cw(2) = 16
    cw(3) = 10 : cw(4) = 8 : cw(5) = 8

    For i = 0 To BADGE_COLS - 1
        ws.Columns(sCol + i).ColumnWidth = cw(i)
    Next i
End Sub

'--------------------------------------------------------------
' 名札のセルコンテンツを描画
'--------------------------------------------------------------
Private Sub DrawBadgeContent(ws As Worksheet, sRow As Long, sCol As Long, _
                              empID   As String, empName  As String, empKana As String, _
                              dept    As String, jobType  As String, pos     As String)

    Dim totalRange As Range
    Set totalRange = ws.Range(ws.Cells(sRow, sCol), _
                              ws.Cells(sRow + BADGE_ROWS - 1, sCol + BADGE_COLS - 1))

    ' 既存内容をクリア
    totalRange.UnMerge
    totalRange.ClearContents
    totalRange.ClearFormats

    ' 背景を白に
    totalRange.Interior.Color = RGB(255, 255, 255)

    ' 外枠
    With totalRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous : .Weight = xlThick : .Color = COL_MAIN
    End With
    With totalRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous : .Weight = xlThick : .Color = COL_MAIN
    End With
    With totalRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous : .Weight = xlThick : .Color = COL_MAIN
    End With
    With totalRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous : .Weight = xlThick : .Color = COL_MAIN
    End With

    '--- 1. ヘッダーバー「長崎北徳洲会病院」(行0-1, 2行) ---
    Call SetMergedCell(ws, sRow + 0, sCol, 2, BADGE_COLS, _
        "長崎北徳洲会病院", 10, True, "MS P明朝", _
        RGB(255, 255, 255), COL_MAIN, xlCenter, xlCenter)

    '--- 2. 部署名 (行2-4, 3行) ---
    Call SetMergedCell(ws, sRow + 2, sCol, 3, BADGE_COLS, _
        dept, 13, True, "MS Pゴシック", _
        COL_MAIN, RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 3. 区切り線 (行5) ---
    Call FillSeparator(ws, sRow + 5, sCol, BADGE_COLS, COL_DARK)

    '--- 4. 職種 (行6, 1行) ---
    Call SetMergedCell(ws, sRow + 6, sCol, 1, BADGE_COLS, _
        jobType, 10, False, "MS Pゴシック", _
        RGB(50, 50, 50), RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 5. 役職 (行7, 1行) ---
    Dim posDisplay As String
    posDisplay = IIf(pos = "", "　", pos)  ' 空欄でも高さを保つ
    Call SetMergedCell(ws, sRow + 7, sCol, 1, BADGE_COLS, _
        posDisplay, 9, False, "MS Pゴシック", _
        RGB(90, 90, 90), RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 6. 苗字ひらがな（大） (行8-14, 7行) ---
    Call SetMergedCell(ws, sRow + 8, sCol, 7, BADGE_COLS, _
        empKana, 34, True, "MS Pゴシック", _
        RGB(0, 0, 0), RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 7. 氏名（漢字） (行15-16, 2行) ---
    Call SetMergedCell(ws, sRow + 15, sCol, 2, BADGE_COLS, _
        empName, 10, False, "MS P明朝", _
        RGB(50, 50, 50), RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 8. ID (行17, 1行) ---
    Call SetMergedCell(ws, sRow + 17, sCol, 1, BADGE_COLS, _
        "ID：" & empID, 8, False, "MS Pゴシック", _
        RGB(120, 120, 120), RGB(255, 255, 255), xlCenter, xlCenter)

    '--- 9. 区切り線 (行18) ---
    Call FillSeparator(ws, sRow + 18, sCol, BADGE_COLS, COL_DARK)

    '--- 10. フッター：徳洲会ロゴ＋法人名 (行19-21, 3行) ---
    Dim footerText As String
    footerText = "医療法人　徳洲会" & Chr(10) & "長崎北徳洲会病院"

    Call SetMergedCell(ws, sRow + 19, sCol, 3, BADGE_COLS, _
        footerText, 7, False, "MS P明朝", _
        COL_MAIN, COL_BG_LIGHT, xlCenter, xlCenter)

    ' フッターは複数行テキストを折り返す
    ws.Range(ws.Cells(sRow + 19, sCol), _
             ws.Cells(sRow + 21, sCol + BADGE_COLS - 1)).WrapText = True

    ' ロゴ画像を挿入（画像ファイルが存在する場合のみ）
    Call TryInsertLogo(ws, sRow, sCol)
End Sub

'--------------------------------------------------------------
' セルをマージして書式・値を設定するヘルパー
'--------------------------------------------------------------
Private Sub SetMergedCell(ws As Worksheet, _
                           startRow As Long, startCol As Long, _
                           numRows  As Integer, numCols As Integer, _
                           cellVal  As String, fSize   As Integer, _
                           isBold   As Boolean, fName  As String, _
                           fColor   As Long, bgColor As Long, _
                           hAlign   As XlHAlign, vAlign As XlVAlign)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, startCol), _
                       ws.Cells(startRow + numRows - 1, startCol + numCols - 1))
    rng.Merge
    rng.Value            = cellVal
    rng.Font.Size        = fSize
    rng.Font.Bold        = isBold
    rng.Font.Name        = fName
    rng.Font.Color       = fColor
    rng.Interior.Color   = bgColor
    rng.HorizontalAlignment = hAlign
    rng.VerticalAlignment   = vAlign
End Sub

'--------------------------------------------------------------
' 区切り線（塗りつぶしセル）を描画
'--------------------------------------------------------------
Private Sub FillSeparator(ws As Worksheet, _
                           sepRow As Long, startCol As Long, _
                           numCols As Integer, sepColor As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(sepRow, startCol), _
                       ws.Cells(sepRow, startCol + numCols - 1))
    rng.Merge
    rng.Interior.Color = sepColor
    rng.RowHeight = 3
End Sub

'--------------------------------------------------------------
' 徳洲会ロゴ画像の挿入
' このブックと同じフォルダに "tokushukai_logo.png" を置いてください
' 画像がない場合は何もしません
'--------------------------------------------------------------
Private Sub TryInsertLogo(ws As Worksheet, sRow As Long, sCol As Long)
    Dim logoPath As String
    logoPath = ThisWorkbook.Path & "\tokushukai_logo.png"

    If Dir(logoPath) = "" Then Exit Sub  ' ファイルなしはスキップ

    ' フッターエリアの位置・サイズを取得
    Dim footerRange As Range
    Set footerRange = ws.Range(ws.Cells(sRow + 19, sCol), _
                               ws.Cells(sRow + 21, sCol + BADGE_COLS - 1))

    Dim logoLeft   As Double : logoLeft   = footerRange.Left + 3
    Dim logoTop    As Double : logoTop    = footerRange.Top + 2
    Dim logoHeight As Double : logoHeight = footerRange.Height - 4

    On Error GoTo InsertFail
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=logoPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, _
        Left:=logoLeft, Top:=logoTop, _
        Width:=-1, Height:=logoHeight)
    pic.LockAspectRatio = msoTrue
InsertFail:
    On Error GoTo 0
End Sub

'==============================================================
' マスターシートを新規作成
'==============================================================
Private Sub CreateMasterSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Worksheets(MASTER_SHEET).Delete
    On Error GoTo 0

    Set ws = Worksheets.Add(Before:=Worksheets(1))
    ws.Name = MASTER_SHEET

    ' ヘッダー列名
    Dim headers As Variant
    headers = Array("職員ID", "氏名（漢字）", "苗字ひらがな", "部署名", "職種", "役職", "入職日", "名札作成日")
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    ' ヘッダー書式
    With ws.Range("A1:H1")
        .Font.Bold          = True
        .Font.Color         = RGB(255, 255, 255)
        .Interior.Color     = COL_MAIN
        .HorizontalAlignment = xlCenter
        .RowHeight          = 22
    End With

    ' 列幅
    ws.Columns("A").ColumnWidth = 10   ' 職員ID
    ws.Columns("B").ColumnWidth = 14   ' 氏名
    ws.Columns("C").ColumnWidth = 16   ' 苗字ひらがな
    ws.Columns("D").ColumnWidth = 22   ' 部署名
    ws.Columns("E").ColumnWidth = 16   ' 職種
    ws.Columns("F").ColumnWidth = 14   ' 役職
    ws.Columns("G").ColumnWidth = 12   ' 入職日
    ws.Columns("H").ColumnWidth = 12   ' 名札作成日

    ' 入職日・作成日列を日付書式に
    ws.Columns("G").NumberFormat = "yyyy/m/d"
    ws.Columns("H").NumberFormat = "yyyy/m/d"

    ' ウィンドウ枠の固定
    ws.Activate
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True

    ' 入力例（削除して使用）
    ws.Cells(2, COL_ID).Value    = "00001"
    ws.Cells(2, COL_NAME).Value  = "山田 太郎"
    ws.Cells(2, COL_KANA).Value  = "やまだ"
    ws.Cells(2, COL_DEPT).Value  = "内科病棟"
    ws.Cells(2, COL_JOB).Value   = "看護師"
    ws.Cells(2, COL_POS).Value   = "副師長"
    ws.Cells(2, COL_JOIN).Value  = DateSerial(2026, 4, 1)
End Sub

'==============================================================
' 名札印刷シートを新規作成
'==============================================================
Private Sub CreateBadgeSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Worksheets(BADGE_SHEET).Delete
    On Error GoTo 0

    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = BADGE_SHEET

    ' A4縦・印刷設定
    With ws.PageSetup
        .PaperSize            = xlPaperA4
        .Orientation          = xlPortrait
        .LeftMargin           = Application.CentimetersToPoints(0.8)
        .RightMargin          = Application.CentimetersToPoints(0.8)
        .TopMargin            = Application.CentimetersToPoints(1.0)
        .BottomMargin         = Application.CentimetersToPoints(1.0)
        .CenterHorizontally   = True
        .CenterVertically     = False
        .PrintGridlines       = False
    End With

    ws.Cells(1, 1).Select
End Sub

'==============================================================
' 名札印刷シートをリセット（削除→再作成）
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
' ユーティリティ：シートの存在確認
'==============================================================
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

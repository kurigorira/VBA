Attribute VB_Name = "NamudaCreator"
'==============================================================
' 長崎北徳洲会病院 職員名札作成システム v2.1
' レイアウト：横型（実物名札に準拠）
'
' 名札構成：
'   上左： 部署名、職種・役職
'   中央： 苗字ひらがな（大）、氏名漢字
'   右上： 空白（写真なし）
'   右下： バーコードエリア
'   中段： 青いバー＋ID番号（右端）
'   下部： 徳洲会ロゴ（左）＋医療法人徳洲会・長崎北徳洲会病院（右）
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
' [公開] 選択職員の名札を作成
'              職員マスターシートで対象行選択後に実行
'==============================================================
Public Sub 名札作成_選択職員()
    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "先に[ワークブック初期化]を実行してください。", vbExclamation: Exit Sub
    End If
    If ActiveSheet.Name <> MASTER_SHEET Then
        MsgBox MASTER_SHEET & " シートで対象行を選択してから実行してください。", vbExclamation: Exit Sub
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

    If cnt = 0 Then
        MsgBox "2行目以降のデータ行を選択してから実行してください。", vbInformation: Exit Sub
    End If

    Call ResetBadgeSheet
    Dim ws  As Worksheet: Set ws  = Worksheets(MASTER_SHEET)
    Dim wb  As Worksheet: Set wb  = Worksheets(BADGE_SHEET)
    Application.ScreenUpdating = False

    Dim i As Integer
    For i = 0 To cnt - 1
        Dim r As Long: r = selRows(i)
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            Call DrawOneBadge(wb, i, _
                CStr(ws.Cells(r, COL_ID).Value), _
                CStr(ws.Cells(r, COL_NAME).Value), _
                CStr(ws.Cells(r, COL_KANA).Value), _
                CStr(ws.Cells(r, COL_DEPT).Value), _
                CStr(ws.Cells(r, COL_JOB).Value), _
                CStr(ws.Cells(r, COL_POS).Value))
            ws.Cells(r, COL_MADE).Value = Date
        End If
    Next i

    Application.ScreenUpdating = True
    wb.Activate
    wb.Cells(1, 1).Select
    MsgBox cnt & " 枚の名札を作成しました。", vbInformation
End Sub

'==============================================================
' [公開] 全職員一括作成
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
    If MsgBox(lastRow - 1 & " 名全員の名札を作成しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Call ResetBadgeSheet
    Dim wb As Worksheet: Set wb = Worksheets(BADGE_SHEET)
    Application.ScreenUpdating = False

    Dim badgeNum As Integer: badgeNum = 0
    Dim r As Long
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, COL_ID).Value)) <> "" Then
            Call DrawOneBadge(wb, badgeNum, _
                CStr(ws.Cells(r, COL_ID).Value), _
                CStr(ws.Cells(r, COL_NAME).Value), _
                CStr(ws.Cells(r, COL_KANA).Value), _
                CStr(ws.Cells(r, COL_DEPT).Value), _
                CStr(ws.Cells(r, COL_JOB).Value), _
                CStr(ws.Cells(r, COL_POS).Value))
            ws.Cells(r, COL_MADE).Value = Date
            badgeNum = badgeNum + 1
        End If
    Next r

    Application.ScreenUpdating = True
    wb.Activate
    wb.Cells(1, 1).Select
    MsgBox badgeNum & " 枚の名札を作成しました。", vbInformation
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
' 1枚の名札を描画
'==============================================================
Private Sub DrawOneBadge(ws As Worksheet, badgeIdx As Integer, _
                          empID   As String, empName  As String, empKana As String, _
                          dept    As String, jobType  As String, pos     As String)
    Dim colIdx As Integer: colIdx = badgeIdx Mod BADGES_PER_ROW
    Dim rowIdx As Integer: rowIdx = badgeIdx \ BADGES_PER_ROW
    Dim sRow As Long: sRow = rowIdx * (BADGE_ROWS + ROW_GAP) + 1
    Dim sCol As Long: sCol = colIdx * (BADGE_COLS + COL_GAP) + 1

    Call SetCellDimensions(ws, sRow, sCol)
    Call DrawBadgeContent(ws, sRow, sCol, empID, empName, empKana, dept, jobType, pos)
End Sub

'--------------------------------------------------------------
' 行の高さ・列幅を設定
' 合計目標: 縦≈148pt(≈ 52mm)、横≈253pt(≈ 89mm)
'--------------------------------------------------------------
Private Sub SetCellDimensions(ws As Worksheet, sRow As Long, sCol As Long)
    Dim rh(13) As Single
    rh(0)  = 2   ' 上パディング
    rh(1)  = 13  ' 部署名 上
    rh(2)  = 13  ' 部署名 下
    rh(3)  = 11  ' 職種・役職
    rh(4)  = 15  ' ひらがな 1
    rh(5)  = 15  ' ひらがな 2
    rh(6)  = 15  ' ひらがな 3
    rh(7)  = 15  ' ひらがな 4
    rh(8)  = 12  ' 氏名（漢字）
    rh(9)  = 5   ' 青バー + ID
    rh(10) = 11  ' フッター 1
    rh(11) = 11  ' フッター 2
    rh(12) = 11  ' フッター 3
    rh(13) = 2   ' 下パディング
    Dim i As Integer
    For i = 0 To BADGE_ROWS - 1
        ws.Rows(sRow + i).RowHeight = rh(i)
    Next i

    ' 列幅：左エリア(col 0-6)、中間ギャップ(col 7)、右エリア(col 8-10)
    Dim cw(10) As Single
    cw(0) = 1    ' 左パディング
    cw(1) = 3.5
    cw(2) = 3.5
    cw(3) = 3.5
    cw(4) = 3.5
    cw(5) = 3.5
    cw(6) = 3    ' 左エリア右端
    cw(7) = 0.5  ' 中間ギャップ
    cw(8) = 3    ' 右エリア（空白・バーコード）
    cw(9) = 3
    cw(10) = 1.5 ' 右パディング
    For i = 0 To BADGE_COLS - 1
        ws.Columns(sCol + i).ColumnWidth = cw(i)
    Next i
End Sub

'--------------------------------------------------------------
' 名札の内容を描画（実物名札の横型レイアウト）
'--------------------------------------------------------------
Private Sub DrawBadgeContent(ws As Worksheet, sRow As Long, sCol As Long, _
                              empID   As String, empName  As String, empKana As String, _
                              dept    As String, jobType  As String, pos     As String)
    ' -- 全体クリア --
    Dim fullRng As Range
    Set fullRng = ws.Range(ws.Cells(sRow, sCol), _
                            ws.Cells(sRow + BADGE_ROWS - 1, sCol + BADGE_COLS - 1))
    fullRng.UnMerge
    fullRng.ClearContents
    fullRng.ClearFormats
    fullRng.Interior.Color = RGB(255, 255, 255)

    ' -- 外枠（青紫色） --
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

    ' 1. 部署名（行+1〜+2、左エリア）
    Call MC(ws, sRow + 1, sCol, 2, 7, dept, 14, True, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlLeft, xlCenter)

    ' 2. 職種・役職（行+3）
    Dim jobPos As String
    If Trim(pos) <> "" Then
        jobPos = jobType & "　" & pos
    Else
        jobPos = jobType
    End If
    Call MC(ws, sRow + 3, sCol, 1, 7, jobPos, 10, False, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlLeft, xlCenter)

    ' 3. 苗字ひらがな（行+4〜+7）大フォント
    Call MC(ws, sRow + 4, sCol, 4, 7, empKana, 32, True, "MS Pゴシック", _
            RGB(0, 0, 0), RGB(255, 255, 255), xlCenter, xlCenter)

    ' 4. 氏名（漢字）（行+8）
    Call MC(ws, sRow + 8, sCol, 1, 7, empName, 10, False, "MS P明朝", _
            RGB(50, 50, 50), RGB(255, 255, 255), xlCenter, xlCenter)

    ' ===== 右エリア (cols 7-10): 写真なし→空白 ============
    Dim RC As Long: RC = sCol + 7
    Dim RW As Integer: RW = 4   ' 列数

    ' 5. 上半分（行+0〜+3）: 空白
    Dim blankTop As Range
    Set blankTop = ws.Range(ws.Cells(sRow, RC), _
                             ws.Cells(sRow + 3, sCol + BADGE_COLS - 1))
    blankTop.Merge
    blankTop.Interior.Color = RGB(255, 255, 255)

    ' 6. 下半分（行+4〜+8）: バーコードエリア
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
        .WrapText            = False
    End With

    ' ===== 青いバー + ID（行+9） ======================

    ' 7. 青バー（cols 0-8）
    Dim barRng As Range
    Set barRng = ws.Range(ws.Cells(sRow + 9, sCol), _
                           ws.Cells(sRow + 9, sCol + 8))
    barRng.Merge
    barRng.Interior.Color = RGB(42, 107, 183)

    ' 8. ID番号（cols 9-10、青バーの右端）
    Dim idRng As Range
    Set idRng = ws.Range(ws.Cells(sRow + 9, sCol + 9), _
                          ws.Cells(sRow + 9, sCol + BADGE_COLS - 1))
    idRng.Merge
    With idRng
        .Value               = empID
        .Font.Size           = 8
        .Font.Name           = "Arial"
        .Font.Color          = RGB(0, 0, 0)
        .Interior.Color      = RGB(255, 255, 255)
        .HorizontalAlignment = xlRight
        .VerticalAlignment   = xlCenter
    End With

    ' ===== フッター（行+10〜+12） =======================
    Dim footBG As Long: footBG = RGB(214, 234, 248)

    ' 9. フッター全体背景
    Dim footAll As Range
    Set footAll = ws.Range(ws.Cells(sRow + 10, sCol), _
                            ws.Cells(sRow + 12, sCol + BADGE_COLS - 1))
    footAll.Interior.Color = footBG
    footAll.Borders(xlEdgeTop).LineStyle = xlContinuous
    footAll.Borders(xlEdgeTop).Color     = RGB(80, 80, 160)

    ' 10. ロゴエリア（左 3列）
    Dim logoRng As Range
    Set logoRng = ws.Range(ws.Cells(sRow + 10, sCol), _
                            ws.Cells(sRow + 12, sCol + 2))
    logoRng.Merge
    With logoRng
        .Interior.Color      = footBG
        .Value               = ""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With
    logoRng.Borders(xlEdgeRight).LineStyle = xlContinuous
    logoRng.Borders(xlEdgeRight).Color     = RGB(150, 150, 200)

    ' 11. 医療法人徳洲会（行+10、右 8列）
    Call MC(ws, sRow + 10, sCol + 3, 1, 8, "医療法人　徳洲会", 8, False, "MS P明朝", _
            RGB(0, 0, 0), footBG, xlLeft, xlCenter)

    ' 12. 長崎北徳洲会病院（行+11〜+12、右 8列）
    Call MC(ws, sRow + 11, sCol + 3, 2, 8, "長崎北徳洲会病院", 13, True, "MS P明朝", _
            RGB(0, 0, 0), footBG, xlLeft, xlCenter)

    ' ===== ロゴ画像の挿入 (tokushukai_logo.png が同フォルダにあれば) ===
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
' 徳洲会ロゴ画像をフッターロゴエリアに挿入
' 「ブックと同じフォルダ」に tokushukai_logo.png を配置してください
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

    ws.Columns("A").ColumnWidth = 10  ' 職員ID
    ws.Columns("B").ColumnWidth = 14  ' 氏名
    ws.Columns("C").ColumnWidth = 16  ' 苗字ひらがな
    ws.Columns("D").ColumnWidth = 22  ' 部署名
    ws.Columns("E").ColumnWidth = 16  ' 職種
    ws.Columns("F").ColumnWidth = 14  ' 役職
    ws.Columns("G").ColumnWidth = 12  ' 入職日
    ws.Columns("H").ColumnWidth = 12  ' 名札作成日
    ws.Columns("G").NumberFormat = "yyyy/m/d"
    ws.Columns("H").NumberFormat = "yyyy/m/d"

    ws.Cells(2, COL_ID).Value   = "108699"
    ws.Cells(2, COL_NAME).Value = "栗原　剛"
    ws.Cells(2, COL_KANA).Value = "くりはら"
    ws.Cells(2, COL_DEPT).Value = "総務課"
    ws.Cells(2, COL_JOB).Value  = "事務員"
    ws.Cells(2, COL_POS).Value  = "主任"
    ws.Cells(2, COL_JOIN).Value = DateSerial(2020, 4, 1)

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

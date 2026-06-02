Attribute VB_Name = "PatientPeriodReport"
Option Explicit

' ============================================================
'  入院患者 期間管理レポート作成マクロ
'  対象CSV列: B=病棟, D=患者氏名, G=患者コード, O=入院日,
'             AP=入院期間区分, AU-AY=DPC関連期間情報
' ============================================================

' --- 元CSVの列番号（Excel上で開いたとき、1始まり） ---
Private Const CSV_B  As Integer = 2   ' 病棟
Private Const CSV_D  As Integer = 4   ' 患者氏名
Private Const CSV_G  As Integer = 7   ' 患者コード
Private Const CSV_O  As Integer = 15  ' 入院日
Private Const CSV_AP As Integer = 42  ' 入院期間区分
Private Const CSV_AU As Integer = 47  ' AU列（手術等予定①）
Private Const CSV_AV As Integer = 48  ' AV列（手術等予定②）
Private Const CSV_AW As Integer = 49  ' AW列（DPC期間①  ※残り日数含む可）
Private Const CSV_AX As Integer = 50  ' AX列（DPC期間②  ※残り日数含む可）
Private Const CSV_AY As Integer = 51  ' AY列（DPC期間③  ※残り日数含む可）

' --- 入院期間③ 残り日数しきい値 ---
Private Const THRESHOLD_CRITICAL As Integer = 7   ' 赤：残り7日以内
Private Const THRESHOLD_WARNING  As Integer = 14  ' 黄：残り14日以内

' --- レポート出力列（1始まり） ---
Private Const R_BYOTO   As Integer = 1  ' 病棟
Private Const R_NAME    As Integer = 2  ' 患者氏名
Private Const R_PATNO   As Integer = 3  ' 患者コード
Private Const R_NYUIN   As Integer = 4  ' 入院日
Private Const R_PERIOD  As Integer = 5  ' 入院期間区分
Private Const R_REMAIN  As Integer = 6  ' 残り日数（数値）
Private Const R_AU      As Integer = 7  ' AU列
Private Const R_AV      As Integer = 8  ' AV列
Private Const R_AW      As Integer = 9  ' AW列
Private Const R_AX      As Integer = 10 ' AX列
Private Const R_AY      As Integer = 11 ' AY列

Private Const R_MAX_COL As Integer = 11 ' 最終出力列数

' --- 患者データ型 ---
Private Type PatientData
    byoto     As String  ' 病棟
    name      As String  ' 患者氏名
    patNo     As String  ' 患者コード
    nyuinDate As String  ' 入院日
    period    As String  ' 入院期間区分
    remainDays As Long   ' 残り日数（-1=対象外）
    colAU     As String
    colAV     As String
    colAW     As String
    colAX     As String
    colAY     As String
End Type

' =============================================================
' メインエントリポイント
' =============================================================
Public Sub CreatePatientReport()
    Dim fd As FileDialog
    Dim csvPath As String

    ' ---- CSVファイル選択 ----
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "入院患者CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "ファイルが選択されませんでした。処理を中断します。", vbExclamation
            Exit Sub
        End If
        csvPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' ---- CSVをワークブックとして開く ----
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet

    ' OriginにShift-JIS（932）を指定して開く
    Set wbCSV = Workbooks.Open( _
        Filename:=csvPath, _
        Format:=6, _
        Local:=True, _
        Origin:=xlWindows)
    Set wsCSV = wbCSV.Sheets(1)

    ' ---- データ読み込み ----
    Dim lastRow As Long
    lastRow = wsCSV.Cells(wsCSV.Rows.Count, 1).End(xlUp).Row

    Dim patients() As PatientData
    ReDim patients(1 To lastRow - 1)  ' ヘッダー除く
    Dim patCount As Long
    patCount = 0

    Dim i As Long
    For i = 2 To lastRow
        ' 空行スキップ（A列が空）
        If Trim(CStr(wsCSV.Cells(i, 1).Value)) = "" Then GoTo NextRow

        Dim p As PatientData
        p.byoto     = SafeStr(wsCSV.Cells(i, CSV_B))
        p.name      = SafeStr(wsCSV.Cells(i, CSV_D))
        p.patNo     = SafeStr(wsCSV.Cells(i, CSV_G))
        p.nyuinDate = SafeStr(wsCSV.Cells(i, CSV_O))
        p.period    = SafeStr(wsCSV.Cells(i, CSV_AP))
        p.colAU     = SafeStr(wsCSV.Cells(i, CSV_AU))
        p.colAV     = SafeStr(wsCSV.Cells(i, CSV_AV))
        p.colAW     = SafeStr(wsCSV.Cells(i, CSV_AW))
        p.colAX     = SafeStr(wsCSV.Cells(i, CSV_AX))
        p.colAY     = SafeStr(wsCSV.Cells(i, CSV_AY))

        ' 残り日数を AW→AX→AY の順に探す（各行に必ず1つ存在）
        p.remainDays = -1
        Dim rd As Long
        rd = ExtractRemainingDays(p.colAW)
        If rd >= 0 Then
            p.remainDays = rd
        Else
            rd = ExtractRemainingDays(p.colAX)
            If rd >= 0 Then
                p.remainDays = rd
            Else
                rd = ExtractRemainingDays(p.colAY)
                If rd >= 0 Then p.remainDays = rd
            End If
        End If

        patCount = patCount + 1
        patients(patCount) = p

NextRow:
    Next i

    ' CSV ブックを保存せずに閉じる
    wbCSV.Close SaveChanges:=False

    If patCount = 0 Then
        MsgBox "データが見つかりませんでした。", vbExclamation
        GoTo Cleanup
    End If

    ' ---- レポートシート生成 ----
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("患者入院期間レポート").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Dim wsRep As Worksheet
    Set wsRep = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsRep.Name = "患者入院期間レポート"

    ' ---- レポート構築 ----
    Call BuildReport(wsRep, patients, patCount)

    wsRep.Activate
    wsRep.Cells(1, 1).Select

    MsgBox "レポートを作成しました！" & vbCrLf & _
           "シート名：患者入院期間レポート", vbInformation, "完了"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
End Sub

' =============================================================
' レポート本体の構築
' =============================================================
Private Sub BuildReport(ws As Worksheet, patients() As PatientData, patCount As Long)
    Dim curRow As Long
    curRow = 1

    ' 列ヘッダー定義
    Dim hdrs(1 To R_MAX_COL) As String
    hdrs(R_BYOTO)  = "病棟"
    hdrs(R_NAME)   = "患者氏名"
    hdrs(R_PATNO)  = "患者コード"
    hdrs(R_NYUIN)  = "入院日"
    hdrs(R_PERIOD) = "入院期間区分"
    hdrs(R_REMAIN) = "残り日数"
    hdrs(R_AU)     = "手術等予定①"
    hdrs(R_AV)     = "手術等予定②"
    hdrs(R_AW)     = "DPC期間①"
    hdrs(R_AX)     = "DPC期間②"
    hdrs(R_AY)     = "DPC期間③"

    ' ==========================================
    ' レポートタイトル
    ' ==========================================
    With ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, R_MAX_COL))
        .Merge
        .Value = "入院患者 期間管理レポート　　作成日：" & Format(Now, "yyyy/mm/dd")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlLeft
        .RowHeight = 28
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(0, 0, 0)
        End With
    End With
    curRow = curRow + 1

    ' 凡例行
    curRow = WriteLegend(ws, curRow)
    curRow = curRow + 1

    ' ==========================================
    ' SECTION 1: 入院期間③ 残り日数アラート上位
    ' ==========================================
    ' 対象行インデックスを収集
    Dim alertIdx() As Long
    Dim alertCnt As Long
    alertCnt = 0
    ReDim alertIdx(1 To patCount)

    Dim j As Long
    For j = 1 To patCount
        If IsInPeriod3(patients(j).period) And patients(j).remainDays >= 0 Then
            alertCnt = alertCnt + 1
            alertIdx(alertCnt) = j
        End If
    Next j

    ' 残り日数で昇順ソート（バブルソート）
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To alertCnt - 1
        For b = a + 1 To alertCnt
            If patients(alertIdx(a)).remainDays > patients(alertIdx(b)).remainDays Then
                tmp = alertIdx(a)
                alertIdx(a) = alertIdx(b)
                alertIdx(b) = tmp
            End If
        Next b
    Next a

    If alertCnt > 0 Then
        ' セクションバナー
        curRow = WriteSectionBanner(ws, curRow, _
            "★ 優先確認  |  入院期間③ 残り日数少ない患者（昇順）　　計 " & alertCnt & " 名", _
            RGB(180, 0, 0), RGB(255, 255, 255))
        curRow = WriteHeader(ws, curRow, hdrs)
        For a = 1 To alertCnt
            curRow = WriteDataRow(ws, curRow, patients(alertIdx(a)))
        Next a
        curRow = curRow + 1
    End If

    ' ==========================================
    ' SECTION 2: 期間超え出来高
    ' ==========================================
    Dim overIdx() As Long
    Dim overCnt As Long
    overCnt = 0
    ReDim overIdx(1 To patCount)

    For j = 1 To patCount
        If InStr(patients(j).period, "期間超え出来高") > 0 Then
            overCnt = overCnt + 1
            overIdx(overCnt) = j
        End If
    Next j

    If overCnt > 0 Then
        curRow = WriteSectionBanner(ws, curRow, _
            "✕ 要対応  |  期間超え出来高 患者一覧　　計 " & overCnt & " 名", _
            RGB(64, 64, 64), RGB(255, 255, 255))
        curRow = WriteHeader(ws, curRow, hdrs)
        For a = 1 To overCnt
            curRow = WriteDataRow(ws, curRow, patients(overIdx(a)))
        Next a
        curRow = curRow + 1
    End If

    ' ==========================================
    ' SECTION 3: 全患者データ一覧
    ' ==========================================
    curRow = WriteSectionBanner(ws, curRow, _
        "一覧  |  全患者データ　　計 " & patCount & " 名", _
        RGB(0, 70, 127), RGB(255, 255, 255))
    curRow = WriteHeader(ws, curRow, hdrs)
    For j = 1 To patCount
        curRow = WriteDataRow(ws, curRow, patients(j))
    Next j

    ' ==========================================
    ' 書式仕上げ
    ' ==========================================
    ws.Columns("A:K").AutoFit
    ' 最低列幅調整
    ws.Columns(R_NAME).ColumnWidth = IIf(ws.Columns(R_NAME).ColumnWidth < 12, 12, ws.Columns(R_NAME).ColumnWidth)
    ws.Columns(R_PERIOD).ColumnWidth = IIf(ws.Columns(R_PERIOD).ColumnWidth < 16, 16, ws.Columns(R_PERIOD).ColumnWidth)

    ' 印刷設定
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ""
    End With
End Sub

' =============================================================
' セクションバナー行の書き込み
' =============================================================
Private Function WriteSectionBanner(ws As Worksheet, startRow As Long, _
                                    title As String, bgColor As Long, fgColor As Long) As Long
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))
        .Merge
        .Value = title
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = fgColor
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft
        .RowHeight = 22
    End With
    WriteSectionBanner = startRow + 1
End Function

' =============================================================
' ヘッダー行の書き込み
' =============================================================
Private Function WriteHeader(ws As Worksheet, startRow As Long, hdrs() As String) As Long
    Dim c As Integer
    For c = 1 To R_MAX_COL
        With ws.Cells(startRow, c)
            .Value = hdrs(c)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))
        .RowHeight = 18
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(31, 73, 125)
    End With
    WriteHeader = startRow + 1
End Function

' =============================================================
' データ行の書き込み＆色付け
' =============================================================
Private Function WriteDataRow(ws As Worksheet, startRow As Long, p As PatientData) As Long
    ' データ書き込み
    ws.Cells(startRow, R_BYOTO).Value  = p.byoto
    ws.Cells(startRow, R_NAME).Value   = p.name
    ws.Cells(startRow, R_PATNO).Value  = p.patNo
    ws.Cells(startRow, R_NYUIN).Value  = p.nyuinDate
    ws.Cells(startRow, R_PERIOD).Value = p.period
    ws.Cells(startRow, R_AU).Value     = p.colAU
    ws.Cells(startRow, R_AV).Value     = p.colAV
    ws.Cells(startRow, R_AW).Value     = p.colAW
    ws.Cells(startRow, R_AX).Value     = p.colAX
    ws.Cells(startRow, R_AY).Value     = p.colAY

    If p.remainDays >= 0 Then
        ws.Cells(startRow, R_REMAIN).Value = p.remainDays
        ws.Cells(startRow, R_REMAIN).NumberFormat = "0""日"""
        ws.Cells(startRow, R_REMAIN).HorizontalAlignment = xlCenter
    End If

    ' 行範囲
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))

    ' 行タイプ判定
    Dim rowType As String
    rowType = GetRowType(p)

    ' ---- 背景色・フォント色の設定 ----
    Select Case rowType

        Case "CRIT3"
            ' 入院期間③ 残り7日以内 → 強い赤背景
            rng.Interior.Color = RGB(255, 199, 206)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Bold = True
                .Font.Color = RGB(192, 0, 0)
            End With
            With ws.Cells(startRow, R_REMAIN)
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(192, 0, 0)
            End With
            ' 残り日数セルに枠強調
            With ws.Cells(startRow, R_REMAIN).Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(192, 0, 0)
            End With

        Case "WARN3"
            ' 入院期間③ 残り14日以内 → 黄色背景
            rng.Interior.Color = RGB(255, 242, 204)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Bold = True
                .Font.Color = RGB(156, 101, 0)
            End With
            With ws.Cells(startRow, R_REMAIN)
                .Font.Bold = True
                .Font.Color = RGB(156, 101, 0)
            End With

        Case "SAFE3"
            ' 入院期間③ 余裕あり → 薄緑
            rng.Interior.Color = RGB(226, 239, 218)
            ws.Cells(startRow, R_PERIOD).Font.Color = RGB(55, 126, 34)

        Case "PERIOD2"
            ' 入院期間② → 薄青
            rng.Interior.Color = RGB(221, 235, 247)
            ws.Cells(startRow, R_PERIOD).Font.Color = RGB(31, 73, 125)

        Case "PERIOD1"
            ' 入院期間① → ほぼ白
            rng.Interior.Color = RGB(248, 248, 255)

        Case "OVER"
            ' 期間超え出来高 → グレー＋打消し線＋赤太字
            rng.Interior.Color = RGB(217, 217, 217)
            rng.Font.Color = RGB(120, 120, 120)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Strikethrough = True
                .Font.Bold = True
                .Font.Color = RGB(192, 0, 0)
            End With

        Case Else
            ' その他
            rng.Interior.Color = RGB(250, 250, 250)

    End Select

    ' 下罫線（薄いグレー）
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).Color = RGB(210, 210, 210)

    WriteDataRow = startRow + 1
End Function

' =============================================================
' 凡例行の書き込み
' =============================================================
Private Function WriteLegend(ws As Worksheet, startRow As Long) As Long
    ' 凡例ラベルと色のペア
    Dim legends(1 To 6, 1 To 2) As String  ' (1)ラベル (2)色コード
    legends(1, 1) = "  入院期間③ 残り7日以内（要即対応）  " : legends(1, 2) = "CRIT3"
    legends(2, 1) = "  入院期間③ 残り14日以内（注意）  "   : legends(2, 2) = "WARN3"
    legends(3, 1) = "  入院期間③ 余裕あり  "              : legends(3, 2) = "SAFE3"
    legends(4, 1) = "  入院期間②  "                       : legends(4, 2) = "PERIOD2"
    legends(5, 1) = "  入院期間①  "                       : legends(5, 2) = "PERIOD1"
    legends(6, 1) = "  期間超え出来高（要対応）  "          : legends(6, 2) = "OVER"

    ws.Cells(startRow, 1).Value = "■ 凡例："
    ws.Cells(startRow, 1).Font.Bold = True

    Dim col As Integer
    col = 2
    Dim k As Integer
    For k = 1 To 6
        With ws.Cells(startRow, col)
            .Value = legends(k, 1)
            .Font.Bold = (legends(k, 2) = "CRIT3" Or legends(k, 2) = "OVER")
            Select Case legends(k, 2)
                Case "CRIT3"  : .Interior.Color = RGB(255, 199, 206) : .Font.Color = RGB(192, 0, 0)
                Case "WARN3"  : .Interior.Color = RGB(255, 242, 204) : .Font.Color = RGB(156, 101, 0)
                Case "SAFE3"  : .Interior.Color = RGB(226, 239, 218) : .Font.Color = RGB(55, 126, 34)
                Case "PERIOD2": .Interior.Color = RGB(221, 235, 247) : .Font.Color = RGB(31, 73, 125)
                Case "PERIOD1": .Interior.Color = RGB(248, 248, 255)
                Case "OVER"
                    .Interior.Color = RGB(217, 217, 217)
                    .Font.Color = RGB(192, 0, 0)
                    .Font.Strikethrough = True
            End Select
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
        col = col + 1
    Next k
    WriteLegend = startRow + 1
End Function

' =============================================================
' ヘルパー関数群
' =============================================================

' セル値を安全に文字列化
Private Function SafeStr(cell As Range) As String
    On Error Resume Next
    SafeStr = Trim(CStr(cell.Value))
    On Error GoTo 0
End Function

' 「(残り**日)」から日数を数値で抽出（見つからない場合は -1）
Private Function ExtractRemainingDays(cellVal As String) As Long
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(cellVal, "残り")
    If pos1 = 0 Then
        ExtractRemainingDays = -1
        Exit Function
    End If
    pos2 = InStr(pos1, cellVal, "日")
    If pos2 = 0 Then
        ExtractRemainingDays = -1
        Exit Function
    End If
    Dim numStr As String
    numStr = Mid(cellVal, pos1 + 2, pos2 - pos1 - 2)
    If IsNumeric(numStr) Then
        ExtractRemainingDays = CLng(numStr)
    Else
        ExtractRemainingDays = -1
    End If
End Function

' 入院期間③ かどうか判定
Private Function IsInPeriod3(periodStr As String) As Boolean
    IsInPeriod3 = (InStr(periodStr, "入院期間③") > 0 Or InStr(periodStr, "入院期間?") > 0)
End Function

' 行タイプ文字列を返す
Private Function GetRowType(p As PatientData) As String
    If InStr(p.period, "期間超え出来高") > 0 Then
        GetRowType = "OVER"
    ElseIf IsInPeriod3(p.period) Then
        If p.remainDays >= 0 And p.remainDays <= THRESHOLD_CRITICAL Then
            GetRowType = "CRIT3"
        ElseIf p.remainDays >= 0 And p.remainDays <= THRESHOLD_WARNING Then
            GetRowType = "WARN3"
        Else
            GetRowType = "SAFE3"
        End If
    ElseIf InStr(p.period, "入院期間②") > 0 Then
        GetRowType = "PERIOD2"
    ElseIf InStr(p.period, "入院期間①") > 0 Then
        GetRowType = "PERIOD1"
    Else
        GetRowType = "OTHER"
    End If
End Function

Attribute VB_Name = "PatientPeriodReport"
Option Explicit

' ============================================================
'  入院患者 期間管理レポート作成マクロ
'  対象CSV列: B=病棟, D=患者氏名, G=患者コード, O=入院日,
'             AP=入院期間区分, AU-AY=DPC関連期間情報
'
'  【使い方】
'   1. このマクロが入った .xlsm を開く
'   2. Alt+F8 → CreatePatientReport を実行（またはボタンをクリック）
'   3. CSVファイルを選択するだけで「レポート」シートが生成される
'   4. 内容確認後、ファイル→名前を付けて保存 で日付別に保存する
'
'  【ボタン設置方法】
'   開発タブ→挿入→フォームコントロール→ボタン を配置し
'   CreatePatientReport を割り当てる
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

' 前回使用フォルダを記憶するセル位置（設定シート）
Private Const SETTING_SHEET As String = "設定"
Private Const SETTING_LASTFOLDER_ROW As Long = 1
Private Const SETTING_LASTFOLDER_COL As Long = 2

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
' メインエントリポイント（ボタンまたはマクロ実行で呼ぶ）
' =============================================================
Public Sub CreatePatientReport()
    Dim fd As FileDialog
    Dim csvPath As String
    Dim initFolder As String

    ' 前回フォルダを取得
    initFolder = GetLastFolder()

    ' ---- CSVファイル選択 ----
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "入院患者CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If initFolder <> "" Then .InitialFileName = initFolder & "\"
        If .Show <> -1 Then
            MsgBox "ファイルが選択されませんでした。処理を中断します。", vbExclamation
            Exit Sub
        End If
        csvPath = .SelectedItems(1)
    End With

    ' フォルダを記憶
    SaveLastFolder Left(csvPath, InStrRev(csvPath, "\") - 1)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' ---- CSVをワークブックとして開く（Shift-JIS対応） ----
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet
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
    ReDim patients(1 To lastRow - 1)
    Dim patCount As Long
    patCount = 0

    Dim i As Long
    For i = 2 To lastRow
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

    ' CSVブックを閉じる
    wbCSV.Close SaveChanges:=False

    If patCount = 0 Then
        MsgBox "データが見つかりませんでした。", vbExclamation
        GoTo Cleanup
    End If

    ' ---- レポートシート生成 ----
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' 日付つきシート名（例: レポート_20260602）
    Dim sheetName As String
    sheetName = "レポート_" & Format(Now, "yyyymmdd")

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Dim wsRep As Worksheet
    Set wsRep = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsRep.Name = sheetName

    ' ---- レポート構築 ----
    Call BuildReport(wsRep, patients, patCount, csvPath)

    wsRep.Activate
    wsRep.Cells(1, 1).Select

    ' ---- 別名保存の案内 ----
    Dim saveName As String
    saveName = Left(csvPath, InStrRev(csvPath, "\")) & _
               "入院期間レポート_" & Format(Now, "yyyymmdd") & ".xlsm"

    Dim ans As VbMsgBoxResult
    ans = MsgBox("レポートを作成しました！" & vbCrLf & vbCrLf & _
                 "【シート名】" & sheetName & vbCrLf & vbCrLf & _
                 "別名で保存しますか？" & vbCrLf & _
                 "(はい → 同じフォルダに自動保存)", _
                 vbQuestion + vbYesNo, "完了")

    If ans = vbYes Then
        Application.DisplayAlerts = False
        wb.SaveAs Filename:=saveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        MsgBox "保存しました：" & vbCrLf & saveName, vbInformation
    End If

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
Private Sub BuildReport(ws As Worksheet, patients() As PatientData, _
                        patCount As Long, csvPath As String)
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

    ' --------------------------------------------------
    ' タイトル行
    ' --------------------------------------------------
    With ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, R_MAX_COL))
        .Merge
        .Value = "入院患者 期間管理レポート　　" & _
                 Format(Now, "yyyy年mm月dd日") & "　取込元: " & Mid(csvPath, InStrRev(csvPath, "\") + 1)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlLeft
        .RowHeight = 26
    End With
    curRow = curRow + 1

    ' 凡例行
    curRow = WriteLegend(ws, curRow)
    curRow = curRow + 1

    ' --------------------------------------------------
    ' SECTION 1: 入院期間③ 残り日数アラート（昇順）
    ' --------------------------------------------------
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

    ' 残り日数で昇順ソート
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
        curRow = WriteSectionBanner(ws, curRow, _
            "★ 優先確認  |  入院期間③ 残り日数少ない順  (" & alertCnt & " 名)", _
            RGB(180, 0, 0), RGB(255, 255, 255))
        curRow = WriteHeader(ws, curRow, hdrs)
        For a = 1 To alertCnt
            curRow = WriteDataRow(ws, curRow, patients(alertIdx(a)))
        Next a
        curRow = curRow + 1
    End If

    ' --------------------------------------------------
    ' SECTION 2: 期間超え出来高
    ' --------------------------------------------------
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
            "✕ 要対応  |  期間超え出来高  (" & overCnt & " 名)", _
            RGB(64, 64, 64), RGB(255, 255, 255))
        curRow = WriteHeader(ws, curRow, hdrs)
        For a = 1 To overCnt
            curRow = WriteDataRow(ws, curRow, patients(overIdx(a)))
        Next a
        curRow = curRow + 1
    End If

    ' --------------------------------------------------
    ' SECTION 3: 全患者データ一覧
    ' --------------------------------------------------
    curRow = WriteSectionBanner(ws, curRow, _
        "一覧  |  全患者データ  (" & patCount & " 名)", _
        RGB(0, 70, 127), RGB(255, 255, 255))
    curRow = WriteHeader(ws, curRow, hdrs)
    For j = 1 To patCount
        curRow = WriteDataRow(ws, curRow, patients(j))
    Next j

    ' --------------------------------------------------
    ' 書式仕上げ
    ' --------------------------------------------------
    ws.Columns("A:K").AutoFit
    ws.Columns(R_NAME).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_NAME).ColumnWidth, 12)
    ws.Columns(R_PERIOD).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_PERIOD).ColumnWidth, 16)
    ws.Columns(R_AW).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_AW).ColumnWidth, 20)
    ws.Columns(R_AX).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_AX).ColumnWidth, 20)
    ws.Columns(R_AY).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_AY).ColumnWidth, 20)

    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub

' =============================================================
' セクションバナー行
' =============================================================
Private Function WriteSectionBanner(ws As Worksheet, startRow As Long, _
                                    title As String, bgColor As Long, fgColor As Long) As Long
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))
        .Merge
        .Value = title
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = fgColor
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft
        .RowHeight = 20
    End With
    WriteSectionBanner = startRow + 1
End Function

' =============================================================
' ヘッダー行
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

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))

    Select Case GetRowType(p)

        Case "CRIT3"
            rng.Interior.Color = RGB(255, 199, 206)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Bold = True : .Font.Color = RGB(192, 0, 0)
            End With
            With ws.Cells(startRow, R_REMAIN)
                .Font.Bold = True : .Font.Size = 11 : .Font.Color = RGB(192, 0, 0)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlMedium
                .Borders.Color = RGB(192, 0, 0)
            End With

        Case "WARN3"
            rng.Interior.Color = RGB(255, 242, 204)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Bold = True : .Font.Color = RGB(156, 101, 0)
            End With
            With ws.Cells(startRow, R_REMAIN)
                .Font.Bold = True : .Font.Color = RGB(156, 101, 0)
            End With

        Case "SAFE3"
            rng.Interior.Color = RGB(226, 239, 218)
            ws.Cells(startRow, R_PERIOD).Font.Color = RGB(55, 126, 34)

        Case "PERIOD2"
            rng.Interior.Color = RGB(221, 235, 247)
            ws.Cells(startRow, R_PERIOD).Font.Color = RGB(31, 73, 125)

        Case "PERIOD1"
            rng.Interior.Color = RGB(248, 248, 255)

        Case "OVER"
            rng.Interior.Color = RGB(217, 217, 217)
            rng.Font.Color = RGB(120, 120, 120)
            With ws.Cells(startRow, R_PERIOD)
                .Font.Strikethrough = True
                .Font.Bold = True
                .Font.Color = RGB(192, 0, 0)
            End With

        Case Else
            rng.Interior.Color = RGB(250, 250, 250)

    End Select

    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).Color = RGB(210, 210, 210)

    WriteDataRow = startRow + 1
End Function

' =============================================================
' 凡例行
' =============================================================
Private Function WriteLegend(ws As Worksheet, startRow As Long) As Long
    ws.Cells(startRow, 1).Value = "■ 凡例："
    ws.Cells(startRow, 1).Font.Bold = True

    Dim legends(1 To 6, 1 To 3) As Variant  ' label, bgColor, fgColor
    legends(1, 1) = "  入院期間③ 残り7日以内（要即対応）  "
    legends(1, 2) = RGB(255, 199, 206) : legends(1, 3) = RGB(192, 0, 0)

    legends(2, 1) = "  入院期間③ 残り14日以内（注意）  "
    legends(2, 2) = RGB(255, 242, 204) : legends(2, 3) = RGB(156, 101, 0)

    legends(3, 1) = "  入院期間③ 余裕あり  "
    legends(3, 2) = RGB(226, 239, 218) : legends(3, 3) = RGB(55, 126, 34)

    legends(4, 1) = "  入院期間②  "
    legends(4, 2) = RGB(221, 235, 247) : legends(4, 3) = RGB(31, 73, 125)

    legends(5, 1) = "  入院期間①  "
    legends(5, 2) = RGB(248, 248, 255) : legends(5, 3) = RGB(0, 0, 0)

    legends(6, 1) = "  期間超え出来高（要対応）  "
    legends(6, 2) = RGB(217, 217, 217) : legends(6, 3) = RGB(192, 0, 0)

    Dim k As Integer
    For k = 1 To 6
        With ws.Cells(startRow, k + 1)
            .Value = legends(k, 1)
            .Font.Bold = (k = 1 Or k = 6)
            .Font.Color = legends(k, 3)
            .Interior.Color = legends(k, 2)
            If k = 6 Then .Font.Strikethrough = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
    Next k
    WriteLegend = startRow + 1
End Function

' =============================================================
' ヘルパー：前回フォルダの記憶・読み出し
' =============================================================
Private Function GetLastFolder() As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTING_SHEET)
    If ws Is Nothing Then
        GetLastFolder = ""
    Else
        GetLastFolder = CStr(ws.Cells(SETTING_LASTFOLDER_ROW, SETTING_LASTFOLDER_COL).Value)
    End If
    On Error GoTo 0
End Function

Private Sub SaveLastFolder(folderPath As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTING_SHEET)
    If ws Is Nothing Then
        ' 設定シートを作成（非表示）
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SETTING_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Cells(SETTING_LASTFOLDER_ROW, 1).Value = "最終使用フォルダ"
    End If
    ws.Cells(SETTING_LASTFOLDER_ROW, SETTING_LASTFOLDER_COL).Value = folderPath
    On Error GoTo 0
End Sub

' =============================================================
' ヘルパー関数群
' =============================================================

Private Function SafeStr(cell As Range) As String
    On Error Resume Next
    SafeStr = Trim(CStr(cell.Value))
    On Error GoTo 0
End Function

' 「(残り**日)」から日数を数値で抽出（-1=見つからない）
Private Function ExtractRemainingDays(cellVal As String) As Long
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(cellVal, "残り")
    If pos1 = 0 Then ExtractRemainingDays = -1 : Exit Function
    pos2 = InStr(pos1, cellVal, "日")
    If pos2 = 0 Then ExtractRemainingDays = -1 : Exit Function
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
    IsInPeriod3 = (InStr(periodStr, "入院期間③") > 0 Or _
                   InStr(periodStr, "入院期間?") > 0)
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

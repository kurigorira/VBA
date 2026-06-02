Attribute VB_Name = "PatientPeriodReport"
Option Explicit

' ============================================================
'  入院患者 期間管理レポート作成マクロ
'  【CSV列の正確なマッピング】
'   B=部屋(病棟), D=科(診療科), G=患者コード, H=患者氏名,
'   J=性別, K=年齢, O=入院日, AB=退院日,
'   AP=DPC入院期間, AU=診断群分類, AV=DPC決定病名,
'   AW=DPC入院期間1, AX=DPC入院期間2, AY=包括終了日
'
'  【使い方】
'   1. このマクロが入った .xlsm を開く
'   2. Alt+F8 → CreatePatientReport を実行（またはボタンをクリック）
'   3. CSVファイルを選択するだけでレポートシートが生成される
'   4. 内容確認後、名前を付けて保存で日付別に保存する
'
'  【ボタン設置方法】
'   開発タブ→挿入→フォームコントロール→ボタンを配置し
'   CreatePatientReport を割り当てる
' ============================================================

' --- 元CSVの列番号（Excelで開いたとき、1始まり） ---
Private Const CSV_B  As Integer = 2   ' 部屋（病棟）
Private Const CSV_G  As Integer = 7   ' 患者コード
Private Const CSV_H  As Integer = 8   ' 患者氏名
Private Const CSV_J  As Integer = 10  ' 性別
Private Const CSV_K  As Integer = 11  ' 年齢
Private Const CSV_O  As Integer = 15  ' 入院日
Private Const CSV_AB As Integer = 28  ' 退院日
Private Const CSV_AP As Integer = 42  ' DPC入院期間（入院期間区分）
Private Const CSV_AU As Integer = 47  ' 診断群分類
Private Const CSV_AV As Integer = 48  ' DPC決定病名
Private Const CSV_AW As Integer = 49  ' DPC入院期間1（残り日数含む可）
Private Const CSV_AX As Integer = 50  ' DPC入院期間2（残り日数含む可）
Private Const CSV_AY As Integer = 51  ' 包括終了日（残り日数含む可）

' --- 入院期間③ 残り日数しきい値 ---
Private Const THRESHOLD_CRITICAL As Integer = 7   ' 赤：残り7日以内
Private Const THRESHOLD_WARNING  As Integer = 14  ' 黄：残り14日以内

' --- レポート出力列（1始まり） ---
Private Const R_BYOTO   As Integer = 1  ' 部屋（病棟）
Private Const R_NAME    As Integer = 2  ' 患者氏名
Private Const R_PATNO   As Integer = 3  ' 患者コード
Private Const R_SEX     As Integer = 4  ' 性別
Private Const R_AGE     As Integer = 5  ' 年齢
Private Const R_NYUIN   As Integer = 6  ' 入院日
Private Const R_TAIIN   As Integer = 7  ' 退院日
Private Const R_PERIOD  As Integer = 8  ' DPC入院期間区分
Private Const R_REMAIN  As Integer = 9  ' 残り日数（数値）
Private Const R_AU      As Integer = 10 ' 診断群分類
Private Const R_AV      As Integer = 11 ' DPC決定病名
Private Const R_AW      As Integer = 12 ' DPC入院期間1
Private Const R_AX      As Integer = 13 ' DPC入院期間2
Private Const R_AY      As Integer = 14 ' 包括終了日

Private Const R_MAX_COL As Integer = 14 ' 最終出力列数

' 前回使用フォルダを記憶するセル位置（設定シート）
Private Const SETTING_SHEET As String = "設定"
Private Const SETTING_LASTFOLDER_ROW As Long = 1
Private Const SETTING_LASTFOLDER_COL As Long = 2

' --- 患者データ型 ---
Private Type PatientData
    byoto         As String  ' 部屋（病棟）
    name          As String  ' 患者氏名
    patNo         As String  ' 患者コード
    sex           As String  ' 性別
    age           As String  ' 年齢
    nyuinDate     As String  ' 入院日
    dischargeDate As String  ' 退院日（確定含む）
    period        As String  ' DPC入院期間区分
    remainDays    As Long    ' 残り日数（-1=対象外）
    colAU         As String  ' 診断群分類
    colAV         As String  ' DPC決定病名
    colAW         As String  ' DPC入院期間1
    colAX         As String  ' DPC入院期間2
    colAY         As String  ' 包括終了日
End Type

' =============================================================
' メインエントリポイント
' =============================================================
Public Sub CreatePatientReport()
    Dim fd As FileDialog
    Dim csvPath As String
    Dim initFolder As String

    initFolder = GetLastFolder()

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "入院患者CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If initFolder <> "" Then .InitialFileName = initFolder & "\"
        If .Show <> -1 Then
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
        csvPath = .SelectedItems(1)
    End With

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
        p.byoto         = SafeStr(wsCSV.Cells(i, CSV_B))
        p.name          = SafeStr(wsCSV.Cells(i, CSV_H))
        p.patNo         = SafeStr(wsCSV.Cells(i, CSV_G))
        p.sex           = SafeStr(wsCSV.Cells(i, CSV_J))
        p.age           = SafeStr(wsCSV.Cells(i, CSV_K))
        p.nyuinDate     = SafeStr(wsCSV.Cells(i, CSV_O))
        p.dischargeDate = SafeStr(wsCSV.Cells(i, CSV_AB))
        p.period        = SafeStr(wsCSV.Cells(i, CSV_AP))
        p.colAU         = SafeStr(wsCSV.Cells(i, CSV_AU))
        p.colAV         = SafeStr(wsCSV.Cells(i, CSV_AV))
        p.colAW         = SafeStr(wsCSV.Cells(i, CSV_AW))
        p.colAX         = SafeStr(wsCSV.Cells(i, CSV_AX))
        p.colAY         = SafeStr(wsCSV.Cells(i, CSV_AY))

        ' 残り日数を AW→AX→AY の順に探す
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

    wbCSV.Close SaveChanges:=False

    If patCount = 0 Then
        MsgBox "データが見つかりませんでした。", vbExclamation
        GoTo Cleanup
    End If

    ' ---- レポートシート生成 ----
    Dim wb As Workbook
    Set wb = ThisWorkbook

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

    Call BuildReport(wsRep, patients, patCount, csvPath)

    wsRep.Activate
    wsRep.Cells(1, 1).Select

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

    Dim hdrs(1 To R_MAX_COL) As String
    hdrs(R_BYOTO)  = "病室"
    hdrs(R_NAME)   = "患者氏名"
    hdrs(R_PATNO)  = "患者コード"
    hdrs(R_SEX)    = "性別"
    hdrs(R_AGE)    = "年齢"
    hdrs(R_NYUIN)  = "入院日"
    hdrs(R_TAIIN)  = "退院日"
    hdrs(R_PERIOD) = "DPC入院期間"
    hdrs(R_REMAIN) = "残り日数"
    hdrs(R_AU)     = "診断群分類"
    hdrs(R_AV)     = "DPC決定病名"
    hdrs(R_AW)     = "DPC期間①"
    hdrs(R_AX)     = "DPC期間②"
    hdrs(R_AY)     = "包括終了日"

    ' タイトル行
    With ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, R_MAX_COL))
        .Merge
        .Value = "入院患者 期間管理レポート　　" & _
                 Format(Now, "yyyy年mm月dd日") & "　取込元: " & Mid(csvPath, InStrRev(csvPath, "\") + 1)
        .Font.Bold = True : .Font.Size = 13 : .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlLeft : .RowHeight = 26
    End With
    curRow = curRow + 1

    curRow = WriteLegend(ws, curRow)
    curRow = curRow + 1

    ' --------------------------------------------------
    ' SECTION 0: 本日退院患者
    ' --------------------------------------------------
    Dim todayIdx() As Long
    Dim todayCnt As Long
    todayCnt = 0
    ReDim todayIdx(1 To patCount)

    Dim j As Long
    For j = 1 To patCount
        If IsTodayDischarge(patients(j).dischargeDate) Then
            todayCnt = todayCnt + 1
            todayIdx(todayCnt) = j
        End If
    Next j

    If todayCnt > 0 Then
        curRow = WriteSectionBanner(ws, curRow, _
            "◎ 本日退院患者　(" & todayCnt & " 名)　　" & Format(Now, "yyyy/mm/dd"), _
            RGB(70, 130, 40), RGB(255, 255, 255))
        curRow = WriteHeader(ws, curRow, hdrs)
        For j = 1 To todayCnt
            curRow = WriteDataRow(ws, curRow, patients(todayIdx(j)))
        Next j
        curRow = curRow + 1
    Else
        curRow = WriteSectionBanner(ws, curRow, _
            "◎ 本日退院患者　(0 名)　　" & Format(Now, "yyyy/mm/dd"), _
            RGB(150, 150, 150), RGB(255, 255, 255))
        curRow = curRow + 1
    End If

    ' --------------------------------------------------
    ' SECTION 1: 入院期間③ 残り日数アラート（昇順）
    ' --------------------------------------------------
    Dim alertIdx() As Long
    Dim alertCnt As Long
    alertCnt = 0
    ReDim alertIdx(1 To patCount)

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
                tmp = alertIdx(a) : alertIdx(a) = alertIdx(b) : alertIdx(b) = tmp
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
            "? 要対応  |  期間超え出来高  (" & overCnt & " 名)", _
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

    ' 書式仕上げ
    ws.Columns("A:N").AutoFit
    ws.Columns(R_NAME).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_NAME).ColumnWidth, 12)
    ws.Columns(R_PERIOD).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_PERIOD).ColumnWidth, 16)
    ws.Columns(R_AV).ColumnWidth = WorksheetFunction.Max(ws.Columns(R_AV).ColumnWidth, 20)

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
        .Font.Bold = True : .Font.Size = 11 : .Font.Color = fgColor
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft : .RowHeight = 20
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
            .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
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
    ws.Cells(startRow, R_SEX).Value    = p.sex
    ws.Cells(startRow, R_AGE).Value    = p.age
    ws.Cells(startRow, R_NYUIN).Value  = p.nyuinDate
    ws.Cells(startRow, R_TAIIN).Value  = p.dischargeDate
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

    ' 性別・年齢・退院日はセンタリング
    ws.Cells(startRow, R_SEX).HorizontalAlignment = xlCenter
    ws.Cells(startRow, R_AGE).HorizontalAlignment = xlCenter

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, R_MAX_COL))

    ' 本日退院患者
    If IsTodayDischarge(p.dischargeDate) Then
        rng.Interior.Color = RGB(198, 239, 206)
        ws.Cells(startRow, R_NAME).Font.Bold = True
        ws.Cells(startRow, R_NAME).Font.Color = RGB(0, 97, 0)
        ws.Cells(startRow, R_TAIIN).Font.Bold = True
        ws.Cells(startRow, R_TAIIN).Font.Color = RGB(0, 97, 0)
    Else
        Select Case GetRowType(p)
            Case "CRIT3"
                rng.Interior.Color = RGB(255, 199, 206)
                With ws.Cells(startRow, R_PERIOD)
                    .Font.Bold = True : .Font.Color = RGB(192, 0, 0)
                End With
                With ws.Cells(startRow, R_REMAIN)
                    .Font.Bold = True : .Font.Size = 11 : .Font.Color = RGB(192, 0, 0)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlMedium : .Borders.Color = RGB(192, 0, 0)
                End With
            Case "WARN3"
                rng.Interior.Color = RGB(255, 242, 204)
                With ws.Cells(startRow, R_PERIOD)
                    .Font.Bold = True : .Font.Color = RGB(156, 101, 0)
                End With
                ws.Cells(startRow, R_REMAIN).Font.Bold = True
                ws.Cells(startRow, R_REMAIN).Font.Color = RGB(156, 101, 0)
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
                    .Font.Bold = True : .Font.Color = RGB(192, 0, 0)
                End With
            Case Else
                rng.Interior.Color = RGB(250, 250, 250)
        End Select
    End If

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

    Dim legends(1 To 7, 1 To 3) As Variant
    legends(1, 1) = "  本日退院患者  "
    legends(1, 2) = RGB(198, 239, 206) : legends(1, 3) = RGB(0, 97, 0)

    legends(2, 1) = "  入院期間③ 残り7日以内（要即対応）  "
    legends(2, 2) = RGB(255, 199, 206) : legends(2, 3) = RGB(192, 0, 0)

    legends(3, 1) = "  入院期間③ 残り14日以内（注意）  "
    legends(3, 2) = RGB(255, 242, 204) : legends(3, 3) = RGB(156, 101, 0)

    legends(4, 1) = "  入院期間③ 余裕あり  "
    legends(4, 2) = RGB(226, 239, 218) : legends(4, 3) = RGB(55, 126, 34)

    legends(5, 1) = "  入院期間②  "
    legends(5, 2) = RGB(221, 235, 247) : legends(5, 3) = RGB(31, 73, 125)

    legends(6, 1) = "  入院期間①  "
    legends(6, 2) = RGB(248, 248, 255) : legends(6, 3) = RGB(0, 0, 0)

    legends(7, 1) = "  期間超え出来高（要対応）  "
    legends(7, 2) = RGB(217, 217, 217) : legends(7, 3) = RGB(192, 0, 0)

    Dim k As Integer
    For k = 1 To 7
        With ws.Cells(startRow, k + 1)
            .Value = legends(k, 1)
            .Font.Bold = (k = 1 Or k = 7)
            .Font.Color = legends(k, 3)
            .Interior.Color = legends(k, 2)
            If k = 7 Then .Font.Strikethrough = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
    Next k
    WriteLegend = startRow + 1
End Function

' =============================================================
' ヘルパー関数群
' =============================================================

Private Function SafeStr(cell As Range) As String
    On Error Resume Next
    SafeStr = Trim(CStr(cell.Value))
    On Error GoTo 0
End Function

' 本日退院か判定：「yyyy/m/d」と「yyyy/m/d（確定）」両対応
Private Function IsTodayDischarge(dischStr As String) As Boolean
    If dischStr = "" Then
        IsTodayDischarge = False
        Exit Function
    End If
    ' 日付部分だけ抽出（「（確定）」などの接尾語を除去）
    Dim dateStr As String
    Dim parenPos As Long
    parenPos = InStr(dischStr, "(")
    If parenPos = 0 Then parenPos = InStr(dischStr, "（")
    If parenPos > 0 Then
        dateStr = Trim(Left(dischStr, parenPos - 1))
    Else
        dateStr = Trim(dischStr)
    End If

    ' 本日と比較
    Dim todayStr As String
    todayStr = Format(Now, "yyyy/m/d")
    IsTodayDischarge = (dateStr = todayStr)
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
    IsInPeriod3 = (InStr(periodStr, "入院期間?") > 0 Or _
                   InStr(periodStr, "3") > 0 And InStr(periodStr, "入院期間") > 0)
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
    ElseIf InStr(p.period, "入院期間") > 0 And InStr(p.period, "2") > 0 Then
        GetRowType = "PERIOD2"
    ElseIf InStr(p.period, "入院期間") > 0 And InStr(p.period, "1") > 0 Then
        GetRowType = "PERIOD1"
    Else
        GetRowType = "OTHER"
    End If
End Function

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
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SETTING_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Cells(SETTING_LASTFOLDER_ROW, 1).Value = "最終使用フォルダ"
    End If
    ws.Cells(SETTING_LASTFOLDER_ROW, SETTING_LASTFOLDER_COL).Value = folderPath
    On Error GoTo 0
End Sub

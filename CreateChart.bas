Attribute VB_Name = "CreateChart"
Option Explicit

'===================================================
' 各シートに棒グラフを作成するマクロ
' 横軸: 郵送日グループ (1〜15日 / 16〜31日)
' 系列: 有効期限切れ(ピンク) / 間に合った(緑)
'===================================================

' 全シートにグラフを作成
Sub CreateChartAllSheets()
    Dim ws As Worksheet
    Dim createdCount As Long
    createdCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If HasData(ws) Then
            Call CreateChartForSheet(ws)
            createdCount = createdCount + 1
        End If
    Next ws

    MsgBox createdCount & " 枚のシートにグラフを作成しました。", vbInformation, "完了"
End Sub

' アクティブシートのみにグラフを作成
Sub CreateChartActiveSheet()
    If HasData(ActiveSheet) Then
        Call CreateChartForSheet(ActiveSheet)
        MsgBox "グラフを作成しました。", vbInformation, "完了"
    Else
        MsgBox "データが見つかりません。", vbExclamation, "エラー"
    End If
End Sub

' シートにデータがあるか確認
Private Function HasData(ws As Worksheet) As Boolean
    ' E列のヘッダーが「申請書郵送日」またはデータがあるか確認
    HasData = (ws.Cells(ws.Rows.Count, "E").End(xlUp).Row >= 3)
End Function

' 指定シートにグラフを作成
Private Sub CreateChartForSheet(ws As Worksheet)

    '--- 定数定義 ---
    Const DATA_START_ROW As Long = 3    ' データ開始行 (行2=ヘッダー, 行3=データ)
    Const COL_MAILING_DATE As String = "E"  ' 申請書郵送日
    Const COL_EXPIRED As String = "I"       ' 有効期限切れ
    Const CHART_WIDTH  As Double = 400
    Const CHART_HEIGHT As Double = 300

    '--- データ集計 ---
    ' counts(グループ, ステータス)
    '   グループ: 0=郵送日1〜15日, 1=郵送日16〜31日
    '   ステータス: 0=有効期限切れ, 1=間に合った
    Dim counts(1, 1) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim mailingDay As Integer
    Dim expiredStatus As String
    Dim groupIdx As Integer

    lastRow = ws.Cells(ws.Rows.Count, COL_MAILING_DATE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub

    For i = DATA_START_ROW To lastRow
        Dim cellVal As String
        cellVal = CStr(ws.Cells(i, COL_MAILING_DATE).Value)

        If cellVal <> "" And IsDate(cellVal) Then
            mailingDay = Day(CDate(cellVal))
            expiredStatus = CStr(ws.Cells(i, COL_EXPIRED).Value)

            If mailingDay >= 1 And mailingDay <= 15 Then
                groupIdx = 0
            ElseIf mailingDay >= 16 And mailingDay <= 31 Then
                groupIdx = 1
            Else
                GoTo ContinueLoop
            End If

            If expiredStatus = "間に合った" Then
                counts(groupIdx, 1) = counts(groupIdx, 1) + 1
            Else
                counts(groupIdx, 0) = counts(groupIdx, 0) + 1
            End If
        End If
ContinueLoop:
    Next i

    '--- 既存グラフを削除 ---
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    '--- グラフ配置位置の決定 (データの右側) ---
    Dim chartLeft As Double
    Dim chartTop  As Double
    chartLeft = ws.Columns("M").Left
    chartTop  = ws.Rows(2).Top

    '--- グラフオブジェクトを作成 ---
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(chartLeft, chartTop, CHART_WIDTH, CHART_HEIGHT)

    Dim cht As Chart
    Set cht = chartObj.Chart

    ' 集合縦棒グラフ
    cht.ChartType = xlColumnClustered

    ' 既定の系列を削除
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    '--- 系列1: 有効期限切れ (ピンク) ---
    Dim s1 As Series
    Set s1 = cht.SeriesCollection.NewSeries()
    With s1
        .Name   = "有効期限切れ"
        .Values  = Array(counts(0, 0), counts(1, 0))
        .XValues = Array("郵送日が1〜15日", "郵送日が16〜31日")
        .Interior.Color = RGB(255, 182, 193)   ' ライトピンク
        .Border.LineStyle = xlNone
    End With

    '--- 系列2: 間に合った (緑) ---
    Dim s2 As Series
    Set s2 = cht.SeriesCollection.NewSeries()
    With s2
        .Name   = "間に合った"
        .Values  = Array(counts(0, 1), counts(1, 1))
        .Interior.Color = RGB(144, 238, 144)   ' ライトグリーン
        .Border.LineStyle = xlNone
    End With

    '--- 縦軸の設定 ---
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "人数"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = False
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With

    '--- 横軸の設定 ---
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "期限何日前に申請"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = False
    End With

    '--- グラフタイトル非表示 ---
    cht.HasTitle = False

    '--- 凡例を下部に表示 ---
    With cht.Legend
        .Position = xlLegendPositionBottom
    End With

    '--- 書式整理 ---
    With cht.PlotArea
        .Border.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With

    With cht.ChartArea
        .Border.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
    End With

    With chartObj.Border
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With

    Set cht = Nothing
    Set chartObj = Nothing

End Sub

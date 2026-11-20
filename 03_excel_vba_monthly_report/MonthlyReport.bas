Option Explicit

' =============================================================
'  月別・店舗別の売上集計マクロ
'
'  ・A1セルが「日付」になっているシートを、すべて店舗のデータとして読み込む
'    （列：A=日付, B=商品, C=数量, D=金額）
'  ・「集計」シートを作り直し、次の3つを作る
'      1. 月 × 店舗 の売上表（合計行・合計列付き）
'      2. 月別・店舗別の棒グラフ
'      3. 商品別の売上ランキング
'  ・日付や金額が正しく入っていない行はスキップし、件数を最後に表示する
' =============================================================

Private Const SUMMARY_SHEET As String = "集計"
Private Const HEADER_COLOR As Long = 16247773   ' RGB(221, 235, 247)

Public Sub CreateMonthlySummary()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim totals As Object      ' "月|店舗" -> 金額
    Dim months As Object      ' 月 -> True
    Dim products As Object    ' 商品 -> Array(数量, 金額)
    Dim stores As Collection
    Dim data As Variant
    Dim tmp As Variant
    Dim lastRow As Long, i As Long, j As Long
    Dim monthKey As String, key As String, productName As String
    Dim usedRows As Long, skipped As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wb = ThisWorkbook
    Set totals = CreateObject("Scripting.Dictionary")
    Set months = CreateObject("Scripting.Dictionary")
    Set products = CreateObject("Scripting.Dictionary")
    Set stores = New Collection

    ' ---- 1. 各店舗シートのデータを読み込む ----
    For Each ws In wb.Worksheets
        If ws.Name <> SUMMARY_SHEET And Trim$(CStr(ws.Range("A1").Value)) = "日付" Then
            stores.Add ws.Name
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                data = ws.Range("A2:D" & lastRow).Value   ' 一括で配列に読み込む（速い）
                For i = 1 To UBound(data, 1)
                    If IsEmpty(data(i, 1)) And IsEmpty(data(i, 4)) Then
                        ' 完全な空行は数えない
                    ElseIf IsValidRow(data(i, 1), data(i, 3), data(i, 4)) Then
                        monthKey = Format$(CDate(data(i, 1)), "yyyy-mm")
                        key = monthKey & "|" & ws.Name
                        totals(key) = totals(key) + CDbl(data(i, 4))
                        months(monthKey) = True

                        productName = Trim$(CStr(data(i, 2)))
                        If Not products.Exists(productName) Then products.Add productName, Array(0#, 0#)
                        tmp = products(productName)
                        tmp(0) = tmp(0) + CDbl(data(i, 3))
                        tmp(1) = tmp(1) + CDbl(data(i, 4))
                        products(productName) = tmp
                        usedRows = usedRows + 1
                    Else
                        skipped = skipped + 1
                    End If
                Next i
            End If
        End If
    Next ws

    If stores.Count = 0 Or months.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "集計できるデータが見つかりませんでした。" & vbCrLf & _
               "A1セルが「日付」のシートに、日付・商品・数量・金額を入力してください。", vbExclamation
        Exit Sub
    End If

    ' ---- 2. 「集計」シートを作り直す ----
    On Error Resume Next
    Set sh = wb.Worksheets(SUMMARY_SHEET)
    On Error GoTo ErrHandler
    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        sh.Delete
        Application.DisplayAlerts = True
    End If
    Set sh = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    sh.Name = SUMMARY_SHEET

    ' ---- 3. 月 × 店舗 の表 ----
    Dim monthArr As Variant
    Dim nM As Long, nS As Long
    Dim outArr() As Variant
    Dim v As Double, rowSum As Double

    monthArr = SortedKeys(months)
    nM = UBound(monthArr) - LBound(monthArr) + 1
    nS = stores.Count
    ReDim outArr(1 To nM + 2, 1 To nS + 2)

    outArr(1, 1) = "月"
    For j = 1 To nS
        outArr(1, j + 1) = stores(j)
    Next j
    outArr(1, nS + 2) = "合計"

    For i = 1 To nM
        outArr(i + 1, 1) = monthArr(LBound(monthArr) + i - 1)
        rowSum = 0
        For j = 1 To nS
            key = outArr(i + 1, 1) & "|" & stores(j)
            If totals.Exists(key) Then v = totals(key) Else v = 0
            outArr(i + 1, j + 1) = v
            rowSum = rowSum + v
        Next j
        outArr(i + 1, nS + 2) = rowSum
    Next i

    outArr(nM + 2, 1) = "合計"
    For j = 2 To nS + 2
        v = 0
        For i = 2 To nM + 1
            v = v + outArr(i, j)
        Next i
        outArr(nM + 2, j) = v
    Next j

    sh.Range("A2").Resize(nM, 1).NumberFormat = "@"   ' "2026-07" が日付に変換されないようにする
    sh.Range("A1").Resize(nM + 2, nS + 2).Value = outArr
    FormatTable sh.Range("A1").Resize(nM + 2, nS + 2), True

    ' ---- 4. 棒グラフ ----
    Dim co As ChartObject
    Set co = sh.ChartObjects.Add( _
        Left:=sh.Cells(1, nS + 4).Left, Top:=sh.Range("A1").Top, Width:=440, Height:=260)
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For j = 1 To nS
            With .SeriesCollection.NewSeries
                .Name = stores(j)
                .Values = sh.Range("A2").Offset(0, j).Resize(nM, 1)
                .XValues = sh.Range("A2").Resize(nM, 1)
            End With
        Next j
        .HasTitle = True
        .ChartTitle.Text = "月別・店舗別 売上"
        .HasLegend = True
    End With

    ' ---- 5. 商品別ランキング ----
    Dim startRow As Long, nP As Long
    Dim productArr As Variant
    Dim prodOut() As Variant

    productArr = KeysSortedByAmount(products)
    nP = UBound(productArr) - LBound(productArr) + 1
    startRow = nM + 5
    ReDim prodOut(1 To nP + 1, 1 To 3)
    prodOut(1, 1) = "商品"
    prodOut(1, 2) = "数量"
    prodOut(1, 3) = "金額"
    For i = 1 To nP
        productName = productArr(LBound(productArr) + i - 1)
        tmp = products(productName)
        prodOut(i + 1, 1) = productName
        prodOut(i + 1, 2) = tmp(0)
        prodOut(i + 1, 3) = tmp(1)
    Next i
    sh.Cells(startRow - 1, 1).Value = "商品別ランキング"
    sh.Cells(startRow - 1, 1).Font.Bold = True
    sh.Cells(startRow, 1).Resize(nP + 1, 3).Value = prodOut
    FormatTable sh.Cells(startRow, 1).Resize(nP + 1, 3), False

    sh.Range("A1").Resize(1, nS + 2).EntireColumn.AutoFit
    sh.Activate
    sh.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "集計が完了しました。" & vbCrLf & vbCrLf & _
           "対象シート：" & nS & " 店舗" & vbCrLf & _
           "集計した行：" & usedRows & " 行" & vbCrLf & _
           "スキップした行：" & skipped & " 行（日付・数量・金額が正しくない行）", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
End Sub

' 日付・数量・金額が集計に使える値かどうか
Private Function IsValidRow(ByVal d As Variant, ByVal qty As Variant, ByVal amount As Variant) As Boolean
    IsValidRow = False
    If IsEmpty(d) Or IsEmpty(qty) Or IsEmpty(amount) Then Exit Function
    If Not IsDate(d) Then Exit Function
    If Not IsNumeric(qty) Or Not IsNumeric(amount) Then Exit Function
    IsValidRow = True
End Function

' Dictionary のキーを昇順に並べた配列を返す
Private Function SortedKeys(ByVal dict As Object) As Variant
    Dim arr As Variant, i As Long, j As Long, t As Variant
    arr = dict.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    SortedKeys = arr
End Function

' 商品を金額の大きい順に並べた配列を返す
Private Function KeysSortedByAmount(ByVal dict As Object) As Variant
    Dim arr As Variant, i As Long, j As Long, t As Variant
    Dim a As Variant, b As Variant
    arr = dict.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            a = dict(arr(i))
            b = dict(arr(j))
            If b(1) > a(1) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    KeysSortedByAmount = arr
End Function

' 表の見た目を整える（見出し行の色、数値の桁区切り、合計行の太字）
Private Sub FormatTable(ByVal rng As Range, ByVal hasTotalRow As Boolean)
    With rng.Rows(1)
        .Font.Bold = True
        .Interior.Color = HEADER_COLOR
    End With
    rng.Offset(1, 1).Resize(rng.Rows.Count - 1, rng.Columns.Count - 1).NumberFormat = "#,##0"
    rng.Borders.LineStyle = xlContinuous
    If hasTotalRow Then
        rng.Rows(rng.Rows.Count).Font.Bold = True
        rng.Columns(rng.Columns.Count).Font.Bold = True
    End If
End Sub

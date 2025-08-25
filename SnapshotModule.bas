Option Explicit

' ===========================
' Snapshot Builder for SinglePane (SP.FINANCIALS / SP.FINANCIALS_AGG)
' Creates dynamic Snapshot report grouped by Fund with Fund subtotals and Portfolio totals.
' Requirements:
'   - Sheets: "My Properties" and "Usali Reference" with spilled arrays at A1#
'   - SinglePane add-in installed and logged in
'   - Creates/uses: "Snapshot" (report), "USALI Map" (metric mapping), "Helper" (hidden code lists)
'   - Only user input: Month (1–12) on Snapshot!B1; Year on Snapshot!B2 (defaults to YEAR(TODAY()))
' ===========================

Private Const SH_PROPS As String = "My Properties"
Private Const SH_USALI As String = "Usali Reference"
Private Const SH_SNAP As String = "Snapshot"
Private Const SH_HELP As String = "Helper"
Private Const SH_MAP As String = "USALI Map"

Private Const NAME_USALI_DISPLAY As String = "UsaliMap_Display"
Private Const NAME_USALI_CODE As String = "UsaliMap_Code"
Private Const NAME_STR_DISPLAY As String = "StrMap_Display"
Private Const NAME_STR_CODE As String = "StrMap_Code"
Private Const NAME_MONTHNUM As String = "MonthNum"
Private Const NAME_YEARNUM As String = "YearNum"
Private Const NF_CURR As String = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
Private Const NF_CURR_K As String = "_($* #,##0,_) ;_($* (#,##0,)_);_($* ""-""_) ;_(@_)"
Private Const NF_PCT As String = "0.0%"
Private Const NF_PCT_WHOLE As String = "0.0""%"""
Private Const NF_PCT_SIGN As String = "+0.0%;-0.0%;0.0%"
Private Const NF_DEC As String = "0.0"


Private Const FUND_EXCLUDE As String = "Stonebridge Legacy"  ' put this Fund last + exclude from Managed portfolio total

' Track STR fund table layout for later reference
Private fundDataFirstRow As Long, fundDataLastRow As Long
Private fundTableStartCol As Long, fundTableLastCol As Long
Private fundSubtotalRows As Object
Private Function MetricsList() As Variant
    MetricsList = Array("Occ", "ADR", "RevPAR", "Total Rev (000's)", "NOI (000's)", "NOI Margin")
End Function
Private Sub EnsureSnapHeaderNames(ws As Worksheet)
    ' Ensure Input names exist
    EnsureNamesForInput  ' MonthText, MonthNum, YearNum on Input

    ' Kill any sheet-scoped duplicates that might shadow workbook names
    KillAllSheetScoped "Snap_MonthFull"
    KillAllSheetScoped "Snap_YearNum"
    KillAllSheetScoped "Snap_MonthNum"
    KillAllSheetScoped "Snap_MonthMMM"
     KillAllSheetScoped "Snap_MonthText"


    ' Link Snapshot header cells to Input via formulas (dynamic, not values)
    With ws
        .Range("G2").Formula = "=MonthText" ' full month, e.g., "June"
        .Range("H2").Formula = "=YearNum"   ' four-digit year
    End With

    ' Bind workbook-scoped names to these cells
    AddOrReplaceName "Snap_MonthFull", ws.Range("G2")  ' MMMM (display)
    AddOrReplaceName "Snap_YearNum", ws.Range("H2")    ' year number

    ' Month number from full month (1..12)
    ' Month number from full month text (handles "June" or "Jun", ignores extra spaces)
On Error Resume Next: ThisWorkbook.Names("Snap_MonthNum").Delete: On Error GoTo 0
ThisWorkbook.Names.Add Name:="Snap_MonthNum", _
    RefersTo:="=MONTH(DATEVALUE(""1 ""&TRIM(Snap_MonthFull)))"

' Three-letter token for SP (e.g., "Jun") built from the parsed month + the chosen year
On Error Resume Next: ThisWorkbook.Names("Snap_MonthMMM").Delete: On Error GoTo 0
ThisWorkbook.Names.Add Name:="Snap_MonthMMM", _
    RefersTo:="=TEXT(DATE(Snap_YearNum,Snap_MonthNum,1),""MMM"")"

End Sub


Private Function MetricIsPercent(metric As String) As Boolean
    MetricIsPercent = (UCase$(metric) = "OCC" Or UCase$(metric) = "NOI MARGIN")
End Function

Private Function MetricIsThousands(metric As String) As Boolean
    MetricIsThousands = (UCase$(metric) = "TOTAL REV (000'S)" Or UCase$(metric) = "NOI (000'S)")
End Function

Private Function StrMetricIsPercent(metric As String) As Boolean
    StrMetricIsPercent = (UCase$(metric) = "OCC" Or InStr(1, metric, "%", vbTextCompare) > 0)
End Function

Private Function StrMetricIsCurrency(metric As String) As Boolean
    StrMetricIsCurrency = (UCase$(metric) = "ADR" Or UCase$(metric) = "REVPAR")
End Function

Private Sub SortVariantStringArray(ByRef arr As Variant)
    If Not IsArray(arr) Then Exit Sub
    QuickSortVar arr, LBound(arr), UBound(arr)
End Sub
Private Function ColorHex&(hex As String)
    ' hex like "E03C31" (no #)
    ColorHex = RGB(val("&H" & Mid$(hex, 1, 2)), val("&H" & Mid$(hex, 3, 2)), val("&H" & Mid$(hex, 5, 2)))
End Function
Private Sub WriteTwoRowHeader(ws As Worksheet, topRow As Long, mode As String, startCol As Long, lastCol As Long)
    Const RED_HEX As String = "E03C31"
    Dim red&: red = ColorHex(RED_HEX)

    Dim hdr1 As Range, hdr2 As Range
    Set hdr1 = ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow, lastCol))
    Set hdr2 = ws.Range(ws.Cells(topRow + 1, 2), ws.Cells(topRow + 1, lastCol))

    ' Fill/Font/Heights
    With hdr1
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13
        .RowHeight = 18
    End With
    With hdr2
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13
        .RowHeight = 35
        .WrapText = True
    End With

    ' First three columns merged over two rows (left aligned)
    ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 2)).Merge
    ws.Range(ws.Cells(topRow, 3), ws.Cells(topRow + 1, 3)).Merge
    ws.Range(ws.Cells(topRow, 4), ws.Cells(topRow + 1, 4)).Merge
    ws.Cells(topRow, 2).Value = "Hotel"
    ws.Cells(topRow, 3).Value = "Rooms"
    ws.Cells(topRow, 4).Value = "Manager"
    With ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 4))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Build band labels & metric labels
    Dim bandLabels As Variant, verLabels As Variant
    If UCase$(mode) = "FY" Then
        bandLabels = Array("FY Forecast+Actual", "FY Budget", "FY Var vs Bud", "FY LY", "FY Var vs LY")
        verLabels = Array("FORECAST", "BUDGET", "VAR_BUD", "LY", "VAR_LY")
    Else
        bandLabels = Array(mode & " Actual", mode & " Budget", mode & " Var vs Bud", mode & " LY", mode & " Var vs LY")
        verLabels = Array("ACTUAL", "BUDGET", "VAR_BUD", "LY", "VAR_LY")
    End If

    Dim metrics As Variant: metrics = MetricsList()
    Dim mLB&, mUB&, b&, j&, c1&, c2&
    mLB = LBound(metrics): mUB = UBound(metrics)

    For b = LBound(bandLabels) To UBound(bandLabels)
        c1 = startCol + (b - LBound(bandLabels)) * (mUB - mLB + 1)
        c2 = c1 + (mUB - mLB)

        ' Clear any existing text in the span, put label only in the leftmost cell
        ws.Range(ws.Cells(topRow, c1), ws.Cells(topRow, c2)).ClearContents
        ws.Cells(topRow, c1).Value = bandLabels(b)

        With ws.Range(ws.Cells(topRow, c1), ws.Cells(topRow, c2))
            .MergeCells = False
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With

        ' Metric labels per band (wrap / centered)
        For j = mLB To mUB
            ws.Cells(topRow + 1, c1 + (j - mLB)).Value = metrics(j)
            With ws.Cells(topRow + 1, c1 + (j - mLB))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next j
    Next b

    ' Column widths
    Dim metricsPerBand&: metricsPerBand = (mUB - mLB + 1)
    For b = LBound(bandLabels) To UBound(bandLabels)
        c1 = startCol + (b - LBound(bandLabels)) * metricsPerBand
        Dim isVar As Boolean
        isVar = ((b - LBound(bandLabels)) = 2) Or ((b - LBound(bandLabels)) = 4)
        For j = mLB To mUB
            Dim colIdx&: colIdx = c1 + (j - mLB)
            Dim metricName As String: metricName = UCase$(CStr(metrics(j)))
            If (metricName = "TOTAL REV (000'S)" Or metricName = "NOI (000'S)") And Not isVar Then
                ws.Columns(colIdx).ColumnWidth = 14
            Else
                ws.Columns(colIdx).ColumnWidth = 11.5
            End If
        Next j
    Next b

    ' First three columns widths
    ws.Columns(2).ColumnWidth = 28 ' Hotel
    ws.Columns(3).ColumnWidth = 9  ' Rooms
    ws.Columns(4).ColumnWidth = 18 ' Manager
    ws.Columns(5).Hidden = True    ' Code (hidden)
End Sub

Private Sub WriteStrHeader(ws As Worksheet, topRow As Long, startCol As Long, lastCol As Long, Optional secLabelsOverride As Variant, Optional secMetricsOverride As Variant, Optional firstColLabel As String = "Hotel", Optional includeRoomsAndManagerCols As Boolean = True)
    Const RED_HEX As String = "E03C31"
    Dim red&: red = ColorHex(RED_HEX)

    Dim hdr1 As Range, hdr2 As Range
    Set hdr1 = ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow, lastCol))
    Set hdr2 = ws.Range(ws.Cells(topRow + 1, 2), ws.Cells(topRow + 1, lastCol))

    With hdr1
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13
        .RowHeight = 18
    End With
    With hdr2
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13
        .RowHeight = 35: .WrapText = True
    End With

    ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 2)).Merge
    ws.Cells(topRow, 2).Value = firstColLabel
    If includeRoomsAndManagerCols Then
        ws.Range(ws.Cells(topRow, 3), ws.Cells(topRow + 1, 3)).Merge
        ws.Range(ws.Cells(topRow, 4), ws.Cells(topRow + 1, 4)).Merge
        ws.Cells(topRow, 3).Value = "Rooms"
        ws.Cells(topRow, 4).Value = "Manager"
        With ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 4))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    Else
        ws.Range(ws.Cells(topRow, 3), ws.Cells(topRow + 1, 4)).ClearContents
        With ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 2))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    End If

    Dim secLabels As Variant, secMetrics As Variant
    If IsMissing(secLabelsOverride) Then
        secLabels = Array("MTD STR Data", "YTD", "Running 3 Month", "Running 12 Month")
        secMetrics = Array( _
            Array("Occ", "Occ Index", "% Change", "ADR", "ADR Index", "% Change", "RevPAR", "RevPAR Index", "% Change"), _
            Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"), _
            Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"), _
            Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"))
    Else
        secLabels = secLabelsOverride
        secMetrics = secMetricsOverride
    End If

    Dim i&, start&, j&
    start = startCol
    For i = LBound(secLabels) To UBound(secLabels)
        Dim metrics As Variant: metrics = secMetrics(i)
        Dim mUB&, mLB&: mLB = LBound(metrics): mUB = UBound(metrics)
        Dim secStart&, secEnd&
        secStart = start
        secEnd = start + (mUB - mLB)
        ws.Range(ws.Cells(topRow, secStart), ws.Cells(topRow, secEnd)).ClearContents
        ws.Cells(topRow, secStart).Value = secLabels(i)
        With ws.Range(ws.Cells(topRow, secStart), ws.Cells(topRow, secEnd))
            .MergeCells = False
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        For j = mLB To mUB
            ws.Cells(topRow + 1, secStart + (j - mLB)).Value = metrics(j)
            With ws.Cells(topRow + 1, secStart + (j - mLB))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next j
        start = secEnd + 1
    Next i

    ws.Columns(2).ColumnWidth = 28
    ws.Columns(3).ColumnWidth = 9
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).Hidden = True
End Sub

Private Sub WriteSTRFormulas(ws As Worksheet, r As Long, codeRef As String, startCol As Long)
    Dim secCodes As Variant
    secCodes = Array( _
        Array("Occ", "MPI", "MPI % Chg", "ADR", "ARI", "ARI % Chg", "RevPAR", "RGI", "RGI % Chg"), _
        Array("MPI", "MPI % Chg", "ARI", "ARI % Chg", "RGI", "RGI % Chg"), _
        Array("MPI", "MPI % Chg", "ARI", "ARI % Chg", "RGI", "RGI % Chg"), _
        Array("MPI", "MPI % Chg", "ARI", "ARI % Chg", "RGI", "RGI % Chg"))

    Dim secAgg As Variant
    secAgg = Array("month", "yearToDate", "running3Month", "running12Month")

    Dim dateExpr As String
    dateExpr = "EOMONTH(DATE(Snap_YearNum,Snap_MonthNum,1),0)"

    Dim s&, j&, col&
    col = startCol
    For s = LBound(secCodes) To UBound(secCodes)
        Dim arr As Variant: arr = secCodes(s)
        For j = LBound(arr) To UBound(arr)
            Dim metric As String: metric = arr(j)
            Dim fml As String
            fml = "=SP.STR(" & codeRef & "," & dateExpr & ",""" & secAgg(s) & """,""" & metric & """,""Subject"",""Total"")"
            ws.Cells(r, col).Formula = fml
            If StrMetricIsCurrency(metric) Then
                ws.Cells(r, col).NumberFormat = NF_CURR
            ElseIf StrMetricIsPercent(metric) Then
                If InStr(1, metric, "% Chg", vbTextCompare) > 0 Or _
                   InStr(1, metric, "% Change", vbTextCompare) > 0 Then
                    ws.Cells(r, col).NumberFormat = NF_PCT_WHOLE
                Else
                    ws.Cells(r, col).NumberFormat = NF_PCT
                End If
            Else
                ws.Cells(r, col).NumberFormat = NF_DEC
            End If
            col = col + 1
        Next j
    Next s
End Sub

Private Sub WriteStrAvgFromRange(ws As Worksheet, targetRow As Long, startCol As Long, firstRow As Long, lastRow As Long)
    Dim secStart As Long
    ' MTD section – blank actual stats
    ws.Cells(targetRow, startCol).ClearContents
    ws.Cells(targetRow, startCol + 3).ClearContents
    ws.Cells(targetRow, startCol + 6).ClearContents

    Dim offsets As Variant
    offsets = Array(1, 4, 7)
    Dim i As Long, col As Long
    For i = LBound(offsets) To UBound(offsets)
        col = startCol + offsets(i)
        ws.Cells(targetRow, col).Formula = "=AVERAGE(" & ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Address(False, False) & ")"
        ws.Cells(targetRow, col).NumberFormat = NF_DEC
        ws.Cells(targetRow, col + 1).Formula = "=AVERAGE(" & ws.Range(ws.Cells(firstRow, col + 1), ws.Cells(lastRow, col + 1)).Address(False, False) & ")"
        ws.Cells(targetRow, col + 1).NumberFormat = NF_PCT_WHOLE
    Next i

    ' Remaining sections: YTD, Running 3, Running 12
    secStart = startCol + 9
    Dim s As Long
    For s = 1 To 3
        For i = 0 To 2
            col = secStart + i * 2
            ws.Cells(targetRow, col).Formula = "=AVERAGE(" & ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Address(False, False) & ")"
            ws.Cells(targetRow, col).NumberFormat = NF_DEC
            ws.Cells(targetRow, col + 1).Formula = "=AVERAGE(" & ws.Range(ws.Cells(firstRow, col + 1), ws.Cells(lastRow, col + 1)).Address(False, False) & ")"
            ws.Cells(targetRow, col + 1).NumberFormat = NF_PCT_WHOLE
        Next i
        secStart = secStart + 6
    Next s
End Sub

Private Sub WriteStrAvgAcrossRows(ws As Worksheet, targetRow As Long, startCol As Long, rowsArr As Variant)
    Dim i As Long, col As Long, secStart As Long
    ws.Cells(targetRow, startCol).ClearContents
    ws.Cells(targetRow, startCol + 3).ClearContents
    ws.Cells(targetRow, startCol + 6).ClearContents

    Dim offsets As Variant
    offsets = Array(1, 4, 7)
    For i = LBound(offsets) To UBound(offsets)
        col = startCol + offsets(i)
        ws.Cells(targetRow, col).Formula = BuildAvgList(ws, rowsArr, col)
        ws.Cells(targetRow, col).NumberFormat = NF_DEC
        ws.Cells(targetRow, col + 1).Formula = BuildAvgList(ws, rowsArr, col + 1)
        ws.Cells(targetRow, col + 1).NumberFormat = NF_PCT_WHOLE
    Next i

    secStart = startCol + 9
    Dim s As Long
    For s = 1 To 3
        For i = 0 To 2
            col = secStart + i * 2
            ws.Cells(targetRow, col).Formula = BuildAvgList(ws, rowsArr, col)
            ws.Cells(targetRow, col).NumberFormat = NF_DEC
            ws.Cells(targetRow, col + 1).Formula = BuildAvgList(ws, rowsArr, col + 1)
            ws.Cells(targetRow, col + 1).NumberFormat = NF_PCT_WHOLE
        Next i
        secStart = secStart + 6
    Next s
End Sub

Private Function BuildAvgList(ws As Worksheet, rowsArr As Variant, col As Long) As String
    Dim i As Long, listStr As String
    For i = LBound(rowsArr) To UBound(rowsArr)
        listStr = listStr & "," & ws.Cells(rowsArr(i), col).Address(False, False)
    Next i
    If Len(listStr) > 0 Then listStr = Mid$(listStr, 2)
    BuildAvgList = "=AVERAGE(" & listStr & ")"
End Function

Private Sub WriteStrManagerAverages(ws As Worksheet, r As Long, startCol As Long)
    Dim mgrRange As String
    mgrRange = ws.Range(ws.Cells(fundDataFirstRow, 4), ws.Cells(fundDataLastRow, 4)).Address(False, False)
    Dim colOut As Long: colOut = startCol
    Dim periods As Variant
    periods = Array( _
        Array(1, 2, 4, 5, 7, 8), _
        Array(9, 10, 11, 12, 13, 14), _
        Array(15, 16, 17, 18, 19, 20), _
        Array(21, 22, 23, 24, 25, 26))
    Dim p As Long, o As Long, fundCol As Long, addr As String
    For p = 0 To UBound(periods)
        Dim arrOff As Variant: arrOff = periods(p)
        For o = 0 To UBound(arrOff)
            fundCol = fundTableStartCol + arrOff(o)
            addr = ws.Range(ws.Cells(fundDataFirstRow, fundCol), ws.Cells(fundDataLastRow, fundCol)).Address(False, False)
            ws.Cells(r, colOut).Formula = "=AVERAGEIF(" & mgrRange & "," & ws.Cells(r, 2).Address(False, False) & "," & addr & ")"
            If o Mod 2 = 0 Then
                ws.Cells(r, colOut).NumberFormat = NF_DEC
            Else
                ws.Cells(r, colOut).NumberFormat = NF_PCT_WHOLE
            End If
            colOut = colOut + 1
        Next o
    Next p
End Sub

Private Function ShortManagerName(ByVal s As String) As String
    s = Trim$(CStr(s))
    If Len(s) = 0 Then ShortManagerName = "": Exit Function
    If LCase$(Left$(s, 10)) = "great wolf" Then
        ShortManagerName = "Great Wolf"
    Else
        ShortManagerName = Split(s, " ")(0)
    End If
End Function



Private Function VarianceAsPctOfBase(ByVal metric As String) As Boolean
    Select Case UCase$(metric)
        Case "ADR", "REVPAR", "TOTAL REV (000'S)", "NOI (000'S)"
            VarianceAsPctOfBase = True
        Case Else
            VarianceAsPctOfBase = False   ' Occ and NOI Margin ? simple subtraction (p.p.)
    End Select
End Function

Private Sub WriteMetricFormulas(ws As Worksheet, r As Long, mode As String, isAgg As Boolean, codeOrName As String, headerTopRow As Long, startCol As Long)
    Dim metrics As Variant: metrics = MetricsList()
    Dim bands As Variant
    If UCase$(mode) = "FY" Then
        bands = Array("FORECAST", "BUDGET", "VAR_BUD", "LY", "VAR_LY")
    Else
        bands = Array("ACTUAL", "BUDGET", "VAR_BUD", "LY", "VAR_LY")
    End If

    Dim mLB&, mUB&, bLB&, bUB&, b&, j&, bandCol&, col&, hdrAddr$, mName$
    mLB = LBound(metrics): mUB = UBound(metrics)
    bLB = LBound(bands):   bUB = UBound(bands)

    For b = bLB To bUB
        bandCol = startCol + (b - bLB) * (mUB - mLB + 1)
        For j = mLB To mUB
            col = bandCol + (j - mLB)
            hdrAddr = ws.Cells(headerTopRow + 1, col).Address(False, False)
            mName = CStr(metrics(j))
            Dim tgt As Range: Set tgt = ws.Cells(r, col)

            ' >>> Manual NOI Margin for aggregate rows: margin = NOI / Total Rev (per same band)
            If isAgg And UCase$(mName) = "NOI MARGIN" Then
                Select Case bands(b)
                    Case "ACTUAL", "BUDGET", "LY", "FORECAST"
                        Dim idxRev&, idxNOI&, colRev&, colNOI&, offRev&, offNOI&
                        idxRev = metricIndex("Total Rev (000's)")
                        idxNOI = metricIndex("NOI (000's)")
                        colRev = bandCol + idxRev
                        colNOI = bandCol + idxNOI
                        offRev = colRev - col
                        offNOI = colNOI - col
                        tgt.FormulaR1C1 = "=IFERROR(R[0]C[" & offNOI & "]/R[0]C[" & offRev & "],"""")"
                        tgt.NumberFormat = NF_PCT
                    Case "VAR_BUD", "VAR_LY"
                        ' Variances will be computed below using the normal PutVariance call
                        PutVariance tgt, ws, r, mName, startCol, _
                                   IIf(bands(b) = "VAR_BUD", "ACTUAL", "ACTUAL"), _
                                   IIf(bands(b) = "VAR_BUD", "BUDGET", "LY")
                End Select
                GoTo NextMetricCell
            End If

            ' Normal path
            Select Case bands(b)
                Case "ACTUAL"
                    PutMetric tgt, mode, "Actual", mName, isAgg, codeOrName, hdrAddr
                Case "BUDGET"
                    PutMetric tgt, mode, "Budget", mName, isAgg, codeOrName, hdrAddr
                Case "LY"
                    PutMetric tgt, mode, "LY_Actual", mName, isAgg, codeOrName, hdrAddr
                Case "FORECAST"
                    PutMetric tgt, "FY", "Forecast", mName, isAgg, codeOrName, hdrAddr
                Case "VAR_BUD"
                    PutVariance tgt, ws, r, mName, startCol, "ACTUAL", "BUDGET"
                Case "VAR_LY"
                    PutVariance tgt, ws, r, mName, startCol, "ACTUAL", "LY"
            End Select

NextMetricCell:
        Next j
    Next b
End Sub


Private Sub QuickSortVar(ByRef a As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As Variant
    i = lo: j = hi
    pivot = LCase$(CStr(a((lo + hi) \ 2)))
    Do While i <= j
        Do While LCase$(CStr(a(i))) < pivot: i = i + 1: Loop
        Do While LCase$(CStr(a(j))) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortVar a, lo, j
    If i < hi Then QuickSortVar a, i, hi
End Sub

Private Function MoveValueToEndVariant(ByVal arr As Variant, ByVal val As String) As Variant
    If Not IsArray(arr) Then
        MoveValueToEndVariant = arr
        Exit Function
    End If
    Dim tmp As Collection: Set tmp = New Collection
    Dim i As Long, seen As Boolean
    For i = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(i)), val, vbTextCompare) <> 0 Then
            tmp.Add CStr(arr(i))
        Else
            seen = True
        End If
    Next
    If seen Then tmp.Add val
    Dim out() As String: ReDim out(0 To tmp.Count - 1)
    For i = 1 To tmp.Count: out(i - 1) = tmp(i): Next
    MoveValueToEndVariant = out
End Function
Public Sub BuildSnapshot()
    Dim calcState As XlCalculation, scrn As Boolean, enEvents As Boolean
    calcState = Application.Calculation
    scrn = Application.ScreenUpdating
    enEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
     EnsureSheetsAndNames      ' creates Snapshot/Helper if missing
    EnsureInputSheet          ' creates Input if missing
    EnsureNamesForInput       ' binds MonthText/MonthNum/YearNum


  
    Dim wsSnap As Worksheet: Set wsSnap = Worksheets(SH_SNAP)
    wsSnap.Cells.Clear

    EnsureSnapHeaderNames wsSnap
    EnsureUsaliMap
    EnsureNamedRanges
    
    Dim propsRng As Range
    Set propsRng = SpillOrRegion(Worksheets(SH_PROPS))
    If propsRng Is Nothing Then Err.Raise vbObjectError + 1, , "'My Properties' spill (A1#) not found."
    
    ' Identify key columns in My Properties spill
    Dim cCode&, cHotel&, cMgmt&, cRooms&, cFund&, cMarket&
    cCode = FindHeaderCol(propsRng, "Code")
    cHotel = FindHeaderCol(propsRng, "HotelName")
    cMgmt = FindHeaderCol(propsRng, "ManagementCompany")
    cRooms = FindHeaderCol(propsRng, "Rooms")
    cFund = FindHeaderCol(propsRng, "Fund")
    cMarket = FindHeaderCol(propsRng, "Market")
    If cCode * cHotel * cMgmt * cRooms * cFund * cMarket = 0 Then
        Err.Raise vbObjectError + 2, , "Missing required columns in 'My Properties' spill: Code, HotelName, ManagementCompany, Rooms, Fund, Market."
    End If
    
     
    
    ' Build Fund -> Hotels mapping, Manager -> Hotels mapping, and Market -> Hotels mapping
    Dim fundDict As Object: Set fundDict = CreateObject("Scripting.Dictionary")
    Dim mgrDict As Object: Set mgrDict = CreateObject("Scripting.Dictionary")
    Dim marketDict As Object: Set marketDict = CreateObject("Scripting.Dictionary")
    
    Dim r&, lastR&: lastR = propsRng.Rows.Count
    For r = 2 To lastR ' skip headers
        Dim fund$, hotel$, code$, mgmt$, roomsVal As Variant, market$
        fund = Nz(propsRng.Cells(r, cFund).Value)
        If Len(fund) = 0 Then fund = "(Unassigned)"
        hotel = Nz(propsRng.Cells(r, cHotel).Value)
        code = Nz(propsRng.Cells(r, cCode).Value)
        mgmt = Nz(propsRng.Cells(r, cMgmt).Value)
        roomsVal = propsRng.Cells(r, cRooms).Value
        market = Nz(propsRng.Cells(r, cMarket).Value)
        If Len(market) = 0 Then market = "(Unassigned)"

        If Len(hotel) > 0 And Len(code) > 0 Then
            If Not fundDict.exists(fund) Then fundDict.Add fund, New Collection
            If Not mgrDict.exists(mgmt) Then mgrDict.Add mgmt, New Collection
            If Not marketDict.exists(market) Then marketDict.Add market, New Collection
            Dim rec As Variant
            rec = Array(hotel, code, mgmt, roomsVal)
            fundDict(fund).Add rec
            mgrDict(mgmt).Add rec
            marketDict(market).Add rec
        End If
    Next
    
    ' Sort hotels by name within each fund & sort funds alpha with FUND_EXCLUDE last
   Dim funds As Variant
   funds = fundDict.keys
   If IsArray(funds) Then
        SortVariantStringArray funds
        funds = MoveValueToEndVariant(funds, FUND_EXCLUDE)
   Else
        funds = Array()
   End If

   Dim mgrs As Variant
   mgrs = mgrDict.keys
   If IsArray(mgrs) Then SortVariantStringArray mgrs Else mgrs = Array()

   Dim markets As Variant
   markets = marketDict.keys
   If IsArray(markets) Then SortVariantStringArray markets Else markets = Array()

    PrepareHelperCodeLists fundDict, funds, mgrDict, mgrs, marketDict, markets

  

    
    ' Create the three stacked tables
    Dim rowPtr&: rowPtr = 4
    rowPtr = BuildOneTable(wsSnap, fundDict, funds, rowPtr, "MTD")
    rowPtr = rowPtr + 2
    rowPtr = BuildOneTable(wsSnap, fundDict, funds, rowPtr, "YTD")
    rowPtr = rowPtr + 2
    rowPtr = BuildOneTable(wsSnap, fundDict, funds, rowPtr, "FY")
    rowPtr = rowPtr + 2
    rowPtr = BuildStrFundTable(wsSnap, fundDict, funds, rowPtr)
    rowPtr = rowPtr + 1
    rowPtr = BuildStrManagerTable(wsSnap, mgrs, rowPtr)
    rowPtr = rowPtr + 1
    rowPtr = BuildStrMarketTable(wsSnap, markets, rowPtr)
    
' Format overall sheet (no .Select calls)
With wsSnap
    .Columns(5).Hidden = True   ' Code column (E) hidden
End With


    
    
CleanExit:
    Application.Calculation = calcState
    Application.ScreenUpdating = scrn
    Application.EnableEvents = enEvents
    Exit Sub
CleanFail:
    MsgBox "BuildSnapshot failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' ============== Core table builder ==============

Private Function BuildOneTable(ws As Worksheet, fundDict As Object, funds As Variant, startRow As Long, mode As String) As Long
    ' mode: "MTD", "YTD", "FY"
    Dim rowPtr&: rowPtr = startRow

    ' Header rows (two rows), starting at rowPtr
    Dim startCol&: startCol = 6   ' metrics begin at column F; B=Hotel, C=Rooms, D=Manager, E=Code(hidden)
    Dim metrics As Variant: metrics = MetricsList()
    Dim bandsCount&: bandsCount = 5
    Dim lastCol&: lastCol = startCol + bandsCount * (UBound(metrics) - LBound(metrics) + 1) - 1

    WriteTwoRowHeader ws, rowPtr, mode, startCol, lastCol
    Dim dataFirstRow&: dataFirstRow = rowPtr + 2   ' data starts two rows below header
    rowPtr = dataFirstRow
    Dim headerTopRow&: headerTopRow = dataFirstRow - 2

    Dim propsArea As Range
Set propsArea = SpillOrRegion(Worksheets(SH_PROPS))
Dim propsRef As String
propsRef = "'" & SH_PROPS & "'!" & propsArea.Address(True, True)


    ' Ensure funds array is usable
    If Not IsArray(funds) Then funds = Array()
    Dim hasFunds As Boolean
    On Error Resume Next: hasFunds = (UBound(funds) >= LBound(funds)): On Error GoTo 0

    ' Alternating shading flag
    Dim shadeToggle As Boolean: shadeToggle = False

    ' Loop funds (no fund header rows; assets flow until subtotal)
    Dim f As Long, fund As String, coll As Collection
    If hasFunds Then
        For f = LBound(funds) To UBound(funds)
            On Error Resume Next
            fund = Trim$(CStr(funds(f))): If Len(fund) = 0 Then fund = "(Unassigned)"
            On Error GoTo 0
            If Not fundDict.exists(fund) Then GoTo NextFund

            Set coll = fundDict(fund)

            ' Sort hotels by name
            Dim arr(), i&, j&
            ReDim arr(1 To coll.Count)
            For i = 1 To coll.Count: arr(i) = coll(i): Next
            QuickSortByIndex arr, 0 ' by HotelName

            ' Asset rows
            For i = LBound(arr) To UBound(arr)
                Dim hotel$, code$, mgr$, roomsVal As Variant
                hotel = arr(i)(0): code = arr(i)(1): mgr = arr(i)(2): roomsVal = arr(i)(3)

                ' Basic columns
                ws.Cells(rowPtr, 2).Value = hotel
               ' Get spilled anchor once

' Rooms (col C)
ws.Cells(rowPtr, 3).Value = roomsVal   ' Rooms
    ws.Cells(rowPtr, 4).Value = ShortManagerName(mgr)        ' Manager
    ws.Cells(rowPtr, 5).Value = code
WriteMetricFormulas ws, rowPtr, mode, False, ws.Cells(rowPtr, 5).Address(False, False), headerTopRow, startCol

    

                ' Alternate shading
                If shadeToggle Then ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol)).Interior.Color = RGB(232, 232, 232)
                shadeToggle = Not shadeToggle

                rowPtr = rowPtr + 1
            Next i

      ' Fund subtotal (AGG over fund)
            Dim codesName$: codesName = "Codes_" & SanitizeName(fund)
            ws.Cells(rowPtr, 2).Value = fund & " — Subtotal"
            ws.Cells(rowPtr, 2).Font.Bold = True
            WriteMetricFormulas ws, rowPtr, mode, True, codesName, headerTopRow, startCol

            ' Medium outline around the subtotal row (B:lastCol), no internal grid
            With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
                .Font.Bold = True
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
            End With

            rowPtr = rowPtr + 1

NextFund:
        Next f
    End If

    ' Spacer row after the last fund (Stonebridge Legacy should be last)
    ws.Rows(rowPtr).RowHeight = 10
    rowPtr = rowPtr + 1

    ' Portfolio totals
    ws.Cells(rowPtr, 2).Value = "Total Managed Portfolio (ex. " & FUND_EXCLUDE & ")"
    ws.Cells(rowPtr, 2).Font.Bold = True
    
    ' Total Managed Portfolio
    WriteMetricFormulas ws, rowPtr, mode, True, "Codes_TotalManaged", headerTopRow, startCol

    ' Make all VALUES (data rows only) 12pt
    Dim dataTop As Long: dataTop = headerTopRow + 2   ' row right under the 2 header rows
    If rowPtr - 1 >= dataTop Then
        ws.Range(ws.Cells(dataTop, 2), ws.Cells(rowPtr - 1, lastCol)).Font.Size = 12
    End If

    ' Medium outline + larger font + taller row for Total Managed
    With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
        .Font.Size = 13
          .Font.Bold = True
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
    End With
    ws.Rows(rowPtr).RowHeight = 20
    rowPtr = rowPtr + 1

    ' Total Portfolio row
    ws.Cells(rowPtr, 2).Value = "Total Portfolio"
    ws.Cells(rowPtr, 2).Font.Bold = True
    WriteMetricFormulas ws, rowPtr, mode, True, "Codes_TotalPortfolio", headerTopRow, startCol

    ' Medium outline + larger font + taller row for Total Portfolio
    With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
        .Font.Size = 13
          .Font.Bold = True
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
    End With
    ws.Rows(rowPtr).RowHeight = 20
    rowPtr = rowPtr + 1

    ' Outline border around the whole table (B:lastCol, header+all rows)
 Dim tableTop&: tableTop = dataFirstRow - 2
ApplyBandSeparators ws, tableTop, rowPtr - 1, startCol, metrics


    BuildOneTable = rowPtr
End Function

Private Function BuildStrFundTable(ws As Worksheet, fundDict As Object, funds As Variant, startRow As Long) As Long
    Dim rowPtr&: rowPtr = startRow
    Dim startCol&: startCol = 6
    Dim sec1Cols&: sec1Cols = 9
    Dim sec2Cols&: sec2Cols = 6
    Dim totalCols&: totalCols = sec1Cols + sec2Cols * 3
    Dim lastCol&: lastCol = startCol + totalCols - 1

    WriteStrHeader ws, rowPtr, startCol, lastCol
    Dim dataFirstRow&: dataFirstRow = rowPtr + 2
    rowPtr = dataFirstRow
    fundDataFirstRow = dataFirstRow
    fundTableStartCol = startCol
    fundTableLastCol = lastCol
    Set fundSubtotalRows = CreateObject("Scripting.Dictionary")

    Dim shadeToggle As Boolean: shadeToggle = False
    Dim f As Long, fund As String, coll As Collection
    If IsArray(funds) Then
        For f = LBound(funds) To UBound(funds)
            fund = CStr(funds(f))
            If Not fundDict.exists(fund) Then GoTo NextFund
            Set coll = fundDict(fund)
            Dim arr(), i&, fundStartRow&
            fundStartRow = rowPtr
            ReDim arr(1 To coll.Count)
            For i = 1 To coll.Count: arr(i) = coll(i): Next
            QuickSortByIndex arr, 0
            For i = LBound(arr) To UBound(arr)
                Dim hotel$, code$, mgr$, roomsVal As Variant
                hotel = arr(i)(0): code = arr(i)(1): mgr = arr(i)(2): roomsVal = arr(i)(3)
                ws.Cells(rowPtr, 2).Value = hotel
                ws.Cells(rowPtr, 3).Value = roomsVal
                ws.Cells(rowPtr, 4).Value = ShortManagerName(mgr)
                ws.Cells(rowPtr, 5).Value = code
                WriteSTRFormulas ws, rowPtr, ws.Cells(rowPtr, 5).Address(False, False), startCol
                If shadeToggle Then ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol)).Interior.Color = RGB(232, 232, 232)
                shadeToggle = Not shadeToggle
                rowPtr = rowPtr + 1
            Next i

            ws.Cells(rowPtr, 2).Value = fund & "  Subtotal"
            ws.Cells(rowPtr, 2).Font.Bold = True
            WriteStrAvgFromRange ws, rowPtr, startCol, fundStartRow, rowPtr - 1
            fundSubtotalRows.Add fund, rowPtr
            With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
                .Font.Bold = True
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
                With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
            End With
            rowPtr = rowPtr + 1
NextFund:
        Next f
    End If

    fundDataLastRow = rowPtr - 1

    ws.Rows(rowPtr).RowHeight = 10
    rowPtr = rowPtr + 1

    ws.Cells(rowPtr, 2).Value = "Total Managed Portfolio (ex. " & FUND_EXCLUDE & ")"
    ws.Cells(rowPtr, 2).Font.Bold = True
    Dim managedRows() As Variant, k As Variant, idx As Long
    ReDim managedRows(0 To fundSubtotalRows.Count - 1)
    idx = 0
    For Each k In fundSubtotalRows.keys
        If k <> FUND_EXCLUDE Then
            managedRows(idx) = fundSubtotalRows(k)
            idx = idx + 1
        End If
    Next k
    ReDim Preserve managedRows(0 To idx - 1)
    WriteStrAvgAcrossRows ws, rowPtr, startCol, managedRows
    With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
        .Font.Size = 12
        .Font.Bold = True
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
    End With
    ws.Rows(rowPtr).RowHeight = 20
    rowPtr = rowPtr + 1

    ws.Cells(rowPtr, 2).Value = "Total Portfolio"
    ws.Cells(rowPtr, 2).Font.Bold = True
    Dim allRows() As Variant
    ReDim allRows(0 To fundSubtotalRows.Count - 1)
    idx = 0
    For Each k In fundSubtotalRows.keys
        allRows(idx) = fundSubtotalRows(k)
        idx = idx + 1
    Next k
    WriteStrAvgAcrossRows ws, rowPtr, startCol, allRows
    With ws.Range(ws.Cells(rowPtr, 2), ws.Cells(rowPtr, lastCol))
        .Font.Size = 12
        .Font.Bold = True
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: End With
    End With
    ws.Rows(rowPtr).RowHeight = 20
    rowPtr = rowPtr + 1

    ws.Range(ws.Cells(dataFirstRow, 2), ws.Cells(rowPtr - 1, lastCol)).Font.Size = 12

    Dim tableTop&: tableTop = dataFirstRow - 2
    ApplyStrSeparators ws, tableTop, rowPtr - 1, startCol, Array(sec1Cols, sec2Cols, sec2Cols, sec2Cols), True

    fundTableLastCol = lastCol
    BuildStrFundTable = rowPtr
End Function

Private Function BuildStrManagerTable(ws As Worksheet, mgrs As Variant, startRow As Long) As Long
    Dim rowPtr&: rowPtr = startRow
    Dim startCol&: startCol = 6
    Dim secCols&: secCols = 6
    Dim totalCols&: totalCols = secCols * 4
    Dim lastCol&: lastCol = startCol + totalCols - 1

    Dim secLabels As Variant, secMetrics As Variant
    secLabels = Array("MTD STR Data", "YTD", "Running 3 Month", "Running 12 Month")
    secMetrics = Array( _
        Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"), _
        Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"), _
        Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"), _
        Array("Occ Index", "% Change", "ADR Index", "% Change", "RevPAR Index", "% Change"))
    WriteStrHeader ws, rowPtr, startCol, lastCol, secLabels, secMetrics, "Manager", False
    Dim dataFirstRow&: dataFirstRow = rowPtr + 2
    rowPtr = dataFirstRow

    Dim i As Long
    If IsArray(mgrs) Then
        For i = LBound(mgrs) To UBound(mgrs)
            ws.Cells(rowPtr, 2).Value = ShortManagerName(CStr(mgrs(i)))
            ws.Cells(rowPtr, 2).Font.Bold = True
            WriteStrManagerAverages ws, rowPtr, startCol
            rowPtr = rowPtr + 1
        Next i
    End If

    ws.Range(ws.Cells(dataFirstRow, 2), ws.Cells(rowPtr - 1, lastCol)).Font.Size = 12

    Dim tableTop&: tableTop = dataFirstRow - 2
    ApplyStrSeparators ws, tableTop, rowPtr - 1, startCol, Array(secCols, secCols, secCols, secCols)
    ApplyStrDotDividers ws, tableTop, rowPtr - 1, startCol, 4, secCols

    BuildStrManagerTable = rowPtr
End Function


Private Sub WriteStrMarketHeader(ws As Worksheet, topRow As Long, startCol As Long, segments As Variant, metrics As Variant)
    Const RED_HEX As String = "E03C31"
    Dim red&: red = ColorHex(RED_HEX)
    Dim metricsPerSeg&: metricsPerSeg = UBound(metrics) - LBound(metrics) + 1
    Dim lastCol&: lastCol = startCol + (UBound(segments) - LBound(segments) + 1) * metricsPerSeg - 1

    Dim hdr1 As Range, hdr2 As Range
    Set hdr1 = ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow, lastCol))
    Set hdr2 = ws.Range(ws.Cells(topRow + 1, 2), ws.Cells(topRow + 1, lastCol))

    With hdr1
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13: .RowHeight = 18
    End With
    With hdr2
        .Interior.Color = red: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 13: .RowHeight = 35: .WrapText = True
    End With

    ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 2)).Merge
    ws.Cells(topRow, 2).Value = "Market"
    With ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 1, 2))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    Dim s&, secStart&, secEnd&, m&
    secStart = startCol
    For s = LBound(segments) To UBound(segments)
        secEnd = secStart + metricsPerSeg - 1
        ws.Range(ws.Cells(topRow, secStart), ws.Cells(topRow, secEnd)).ClearContents
        ws.Cells(topRow, secStart).Value = segments(s)
        With ws.Range(ws.Cells(topRow, secStart), ws.Cells(topRow, secEnd))
            .MergeCells = False
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        For m = LBound(metrics) To UBound(metrics)
            ws.Cells(topRow + 1, secStart + (m - LBound(metrics))).Value = metrics(m)
            With ws.Cells(topRow + 1, secStart + (m - LBound(metrics)))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next m
        secStart = secEnd + 1
    Next s

    ws.Columns(2).ColumnWidth = 28
End Sub

Private Sub WriteSTRMarketFormulas(ws As Worksheet, r As Long, codeOrName As String, startCol As Long, segments As Variant, metrics As Variant)
    Dim dateExpr As String
    dateExpr = "EOMONTH(DATE(Snap_YearNum,Snap_MonthNum,1),0)"

    Dim s&, m&, col&
    col = startCol
    For s = LBound(segments) To UBound(segments)
        For m = LBound(metrics) To UBound(metrics)
            Dim metric As String: metric = metrics(m)
            Dim fml As String
            fml = "=SP.STR(INDEX(" & codeOrName & ",1)," & dateExpr & ",""month"",""" & metric & """,""Market Scale"",""" & segments(s) & """)"
            ws.Cells(r, col).Formula = fml
            If StrMetricIsCurrency(metric) Then
                ws.Cells(r, col).NumberFormat = NF_CURR
            ElseIf StrMetricIsPercent(metric) Then
                If InStr(1, metric, "% Chg", vbTextCompare) > 0 Or _
                   InStr(1, metric, "% Change", vbTextCompare) > 0 Then
                    ws.Cells(r, col).NumberFormat = NF_PCT_WHOLE
                Else
                    ws.Cells(r, col).NumberFormat = NF_PCT
                End If
            Else
                ws.Cells(r, col).NumberFormat = NF_DEC
            End If
            col = col + 1
        Next m
    Next s
End Sub

Private Function BuildStrMarketTable(ws As Worksheet, markets As Variant, startRow As Long) As Long
    Dim rowPtr&: rowPtr = startRow
    Dim startCol&: startCol = 6
    Dim segments As Variant: segments = Array("Total", "Transient", "Group", "Contract")
    Dim metrics As Variant: metrics = Array("Occ", "Occ % Chg", "ADR", "ADR % Chg", "RevPAR", "RevPAR % Chg")
    Dim secCols&: secCols = UBound(metrics) - LBound(metrics) + 1
    Dim lastCol&: lastCol = startCol + (UBound(segments) - LBound(segments) + 1) * secCols - 1

    WriteStrMarketHeader ws, rowPtr, startCol, segments, metrics
    Dim dataFirstRow&: dataFirstRow = rowPtr + 2
    rowPtr = dataFirstRow

    Dim i As Long
    If IsArray(markets) Then
        For i = LBound(markets) To UBound(markets)
            Dim market As String: market = CStr(markets(i))
            Dim codesName$: codesName = "CodesMarket_" & SanitizeName(market)
            ws.Cells(rowPtr, 2).Value = market
            ws.Cells(rowPtr, 2).Font.Bold = True
            WriteSTRMarketFormulas ws, rowPtr, codesName, startCol, segments, metrics
            rowPtr = rowPtr + 1
        Next i
    End If

    ws.Range(ws.Cells(dataFirstRow, 2), ws.Cells(rowPtr - 1, lastCol)).Font.Size = 12

    Dim tableTop&: tableTop = dataFirstRow - 2
    ApplyStrSeparators ws, tableTop, rowPtr - 1, startCol, Array(secCols, secCols, secCols, secCols)
    ApplyStrDotDividers ws, tableTop, rowPtr - 1, startCol, _
        UBound(segments) - LBound(segments) + 1, secCols

    BuildStrMarketTable = rowPtr
End Function


Private Sub PutVariance(tgt As Range, ws As Worksheet, r As Long, ByVal metric As String, startCol As Long, leftBand As String, rightBand As String)
    Dim mIdx As Long: mIdx = metricIndex(metric)
    Dim bandIndexLeft As Long, bandIndexRight As Long
    bandIndexLeft = BandIndex(leftBand)
    bandIndexRight = BandIndex(rightBand)

    Dim metrics As Variant: metrics = MetricsList()
    Dim metricsPerBand As Long
    metricsPerBand = UBound(metrics) - LBound(metrics) + 1

    Dim colLeft As Long, colRight As Long, offLeft As Long, offRight As Long
    colLeft = startCol + bandIndexLeft * metricsPerBand + mIdx
    colRight = startCol + bandIndexRight * metricsPerBand + mIdx
    offLeft = colLeft - tgt.Column
    offRight = colRight - tgt.Column

    Dim rLeft As String, rRight As String
    rLeft = "R[0]C[" & offLeft & "]"
    rRight = "R[0]C[" & offRight & "]"

    If VarianceAsPctOfBase(metric) Then
        tgt.FormulaR1C1 = "=IFERROR(IF(ABS(" & rRight & ")>0,(" & rLeft & "-" & rRight & ")/ABS(" & rRight & "),""""),"""")"
    Else
        tgt.FormulaR1C1 = "=IFERROR(" & rLeft & "-" & rRight & ","""")"
    End If

    tgt.NumberFormat = "+0.0%;-0.0%;0.0%"
End Sub




Private Sub PutMetric(tgt As Range, mode As String, version As String, ByVal metric As String, _
                      isAgg As Boolean, codeOrName As String, ByVal metricHeaderAddr As String)
    ' Month token to SP: MTD -> MMM; YTD -> MMMYTD; FY -> "Total Year"
    Dim tokenExpr As String
    Select Case UCase$(mode)
        Case "MTD": tokenExpr = "Snap_MonthMMM"
        Case "YTD": tokenExpr = "Snap_MonthMMM&""YTD"""
        Case "FY":  tokenExpr = """Total Year"""
        Case Else:  tokenExpr = """Total Year"""
    End Select

    ' Version literal. For FY Forecast, we want ForecastN, where N = month number.
    Dim ver As String
    Select Case UCase$(version)
        Case "ACTUAL":    ver = """Actual"""
        Case "BUDGET":    ver = """Budget"""
        Case "FORECAST"
            If UCase$(mode) = "FY" Then
                ver = """Forecast""&Snap_MonthNum"   ' e.g., Forecast6
            Else
                ver = """Forecast"""
            End If
        Case "LY_ACTUAL": ver = """LY_Actual"""
        Case Else:        ver = """" & version & """" ' pass-through if you add more
    End Select

    ' Map from header cell text to USALI code
    Dim usaliLookup As String
    usaliLookup = "IFERROR(XLOOKUP(" & metricHeaderAddr & "," & NAME_USALI_DISPLAY & "," & NAME_USALI_CODE & "),"""")"

    ' Build the SP formula
    Dim fml As String
    If isAgg Then
        fml = "=SP.FINANCIALS_AGG(" & codeOrName & "," & usaliLookup & "," & tokenExpr & ",Snap_YearNum," & ver & ")"
    Else
        fml = "=SP.FINANCIALS(" & codeOrName & "," & usaliLookup & "," & tokenExpr & ",Snap_YearNum," & ver & ")"
    End If
    tgt.Formula = fml

    ' Number formats
    If MetricIsPercent(metric) Then
        tgt.NumberFormat = NF_PCT
    ElseIf MetricIsThousands(metric) Then
        tgt.NumberFormat = NF_CURR_K
    Else
        tgt.NumberFormat = NF_CURR
    End If
End Sub






Private Function BandIndex(band As String) As Long
    Select Case UCase$(band)
        Case "ACTUAL": BandIndex = 0
        Case "BUDGET": BandIndex = 1
        Case "VAR_BUD": BandIndex = 2
        Case "LY": BandIndex = 3
        Case "VAR_LY": BandIndex = 4
        Case Else: BandIndex = 0
    End Select
End Function

Private Function metricIndex(metric As String) As Long
    Select Case UCase$(metric)
        Case "OCC":                    metricIndex = 0
        Case "ADR":                    metricIndex = 1
        Case "REVPAR":                 metricIndex = 2
        Case "TOTAL REV (000'S)":      metricIndex = 3
        Case "NOI (000'S)":            metricIndex = 4
        Case "NOI MARGIN":             metricIndex = 5
        Case Else:                     metricIndex = 0
    End Select
End Function


' ============== Helper code lists for SP.FINANCIALS_AGG ==============

Private Sub PrepareHelperCodeLists(fundDict As Object, funds As Variant, mgrDict As Object, mgrs As Variant, marketDict As Object, markets As Variant)
    Dim wsH As Worksheet: Set wsH = Worksheets(SH_HELP)
    wsH.Cells.Clear

    ' Validate funds
    Dim hasFunds As Boolean
    On Error Resume Next
    hasFunds = (IsArray(funds) And UBound(funds) >= LBound(funds))
    On Error GoTo 0

    ' Dictionaries for uniqueness
    Dim allSet As Object, managedSet As Object
    Set allSet = CreateObject("Scripting.Dictionary")
    Set managedSet = CreateObject("Scripting.Dictionary")

    Dim r As Long: r = 1
    Dim f As Long
    If hasFunds Then
        For f = LBound(funds) To UBound(funds)
            Dim fund As String
            On Error Resume Next
            fund = Trim$(CStr(funds(f)))
            On Error GoTo 0
            If Len(fund) = 0 Then fund = "(Unassigned)"

            If fundDict.exists(fund) Then
                Dim coll As Collection: Set coll = fundDict(fund)
                wsH.Cells(r, 1).Value = "Fund: " & fund
                wsH.Cells(r, 1).Font.Bold = True
                r = r + 1

                Dim startR As Long: startR = r
                Dim i As Long
                For i = 1 To coll.Count
                    Dim code As String
                    code = Trim$(CStr(coll(i)(1)))
                    If Len(code) > 0 Then
                        If Not allSet.exists(code) Then allSet.Add code, True
                        If StrComp(fund, FUND_EXCLUDE, vbTextCompare) <> 0 Then
                            If Not managedSet.exists(code) Then managedSet.Add code, True
                        End If
                        wsH.Cells(r, 1).Value = code
                        r = r + 1
                    End If
                Next i

                ' Name the fund block
                Dim nm As String: nm = "Codes_" & SanitizeName(fund)
                On Error Resume Next: ThisWorkbook.Names(nm).Delete: On Error GoTo 0
                If r > startR Then
                    ThisWorkbook.Names.Add Name:=nm, RefersTo:=wsH.Range(wsH.Cells(startR, 1), wsH.Cells(r - 1, 1))
                Else
                    ThisWorkbook.Names.Add Name:=nm, RefersTo:=wsH.Range("A1:A1")
                End If

                r = r + 1
            End If
        Next f
    End If

    ' Managed total
    Dim startM As Long: startM = r
    Dim key As Variant
    For Each key In managedSet.keys
        wsH.Cells(r, 1).Value = key
        r = r + 1
    Next key
    On Error Resume Next: ThisWorkbook.Names("Codes_TotalManaged").Delete: On Error GoTo 0
    If managedSet.Count > 0 Then
        ThisWorkbook.Names.Add Name:="Codes_TotalManaged", RefersTo:=wsH.Range(wsH.Cells(startM, 1), wsH.Cells(r - 1, 1))
    Else
        ThisWorkbook.Names.Add Name:="Codes_TotalManaged", RefersTo:=wsH.Range("A1:A1")
    End If

    ' Spacer
    r = r + 1

    ' Portfolio total
    Dim startP As Long: startP = r
    For Each key In allSet.keys
        wsH.Cells(r, 1).Value = key
        r = r + 1
    Next key
    On Error Resume Next: ThisWorkbook.Names("Codes_TotalPortfolio").Delete: On Error GoTo 0
    If allSet.Count > 0 Then
        ThisWorkbook.Names.Add Name:="Codes_TotalPortfolio", RefersTo:=wsH.Range(wsH.Cells(startP, 1), wsH.Cells(r - 1, 1))
    Else
        ThisWorkbook.Names.Add Name:="Codes_TotalPortfolio", RefersTo:=wsH.Range("A1:A1")
    End If

    ' Manager code lists
    Dim hasMgrs As Boolean
    On Error Resume Next: hasMgrs = (IsArray(mgrs) And UBound(mgrs) >= LBound(mgrs)): On Error GoTo 0
    If hasMgrs Then
        Dim m As Long
        For m = LBound(mgrs) To UBound(mgrs)
            Dim mgr As String: mgr = CStr(mgrs(m))
            If mgrDict.exists(mgr) Then
                Dim collM As Collection: Set collM = mgrDict(mgr)
                wsH.Cells(r, 1).Value = "Mgr: " & mgr
                wsH.Cells(r, 1).Font.Bold = True
                r = r + 1
                Dim startR2 As Long: startR2 = r
                Dim i2 As Long
                For i2 = 1 To collM.Count
                    Dim codeM As String
                    codeM = Trim$(CStr(collM(i2)(1)))
                    If Len(codeM) > 0 Then
                        wsH.Cells(r, 1).Value = codeM
                        r = r + 1
                    End If
                Next i2
                Dim nmMgr As String: nmMgr = "CodesMgr_" & SanitizeName(mgr)
                On Error Resume Next: ThisWorkbook.Names(nmMgr).Delete: On Error GoTo 0
                If r > startR2 Then
                    ThisWorkbook.Names.Add Name:=nmMgr, RefersTo:=wsH.Range(wsH.Cells(startR2, 1), wsH.Cells(r - 1, 1))
                Else
                    ThisWorkbook.Names.Add Name:=nmMgr, RefersTo:=wsH.Range("A1:A1")
                End If
                r = r + 1
            End If
        Next m
    End If

    ' Market code lists
    Dim hasMarkets As Boolean
    On Error Resume Next: hasMarkets = (IsArray(markets) And UBound(markets) >= LBound(markets)): On Error GoTo 0
    If hasMarkets Then
        Dim mk As Long
        For mk = LBound(markets) To UBound(markets)
            Dim market As String: market = CStr(markets(mk))
            If marketDict.exists(market) Then
                Dim collMk As Collection: Set collMk = marketDict(market)
                wsH.Cells(r, 1).Value = "Market: " & market
                wsH.Cells(r, 1).Font.Bold = True
                r = r + 1
                Dim startMk As Long: startMk = r
                Dim iMk As Long
                For iMk = 1 To collMk.Count
                    Dim codeMk As String
                    codeMk = Trim$(CStr(collMk(iMk)(1)))
                    If Len(codeMk) > 0 Then
                        wsH.Cells(r, 1).Value = codeMk
                        r = r + 1
                    End If
                Next iMk
                Dim nmMarket As String: nmMarket = "CodesMarket_" & SanitizeName(market)
                On Error Resume Next
                ThisWorkbook.Names(nmMarket).Delete
                On Error GoTo 0
                If r > startMk Then
                    ThisWorkbook.Names.Add Name:=nmMarket, RefersTo:=wsH.Range(wsH.Cells(startMk, 1), wsH.Cells(r - 1, 1))
                Else
                    ThisWorkbook.Names.Add Name:=nmMarket, RefersTo:=wsH.Range("A1:A1")
                End If
                r = r + 1
            End If
        Next mk
    End If

    wsH.Visible = xlSheetHidden
End Sub



' ============== Setup helpers ==============

Private Sub EnsureSheetsAndNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(SH_PROPS): On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 10, , "'My Properties' sheet not found."
    Set ws = Nothing
    
    On Error Resume Next
    Set ws = Worksheets(SH_USALI): On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 11, , "'Usali Reference' sheet not found."
    Set ws = Nothing
    
    If Not SheetExists(SH_SNAP) Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = SH_SNAP
    If Not SheetExists(SH_HELP) Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = SH_HELP
    
    ' Keep Helper hidden
    Worksheets(SH_HELP).Visible = xlSheetHidden
    
    ' Named cells for Month/Year inputs will be added after Snapshot exists

End Sub

Private Sub EnsureUsaliMap()
    ' Auto-build "USALI Map" from the tenant's "Usali Reference" spill
    Dim wsMap As Worksheet, wsRef As Worksheet, refRng As Range
    Dim headers As Range, usaliCol As Long
    
    ' Ensure the reference exists
    If Not SheetExists(SH_USALI) Then Err.Raise vbObjectError + 500, , "'" & SH_USALI & "' sheet not found."
    Set wsRef = Worksheets(SH_USALI)
    Set refRng = SpillOrRegion(wsRef)                ' uses your helper; targets A1# if spilled
    
    ' Find the USALI text column by header
    usaliCol = FindHeaderCol(refRng, "usali")
    If usaliCol = 0 Then Err.Raise vbObjectError + 501, , "'" & SH_USALI & "' missing 'usali' header."
    
    ' Create/prepare USALI Map sheet
    If Not SheetExists(SH_MAP) Then
        Set wsMap = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsMap.name = SH_MAP
    Else
        Set wsMap = Worksheets(SH_MAP)
    End If
    wsMap.Cells.Clear
    
    ' Headers
    wsMap.Range("A1").Value = "DisplayMetric"
    wsMap.Range("B1").Value = "USALI"
    wsMap.Range("C1").Value = "Notes"
    wsMap.Range("E1").Value = "STR_Display"
    wsMap.Range("F1").Value = "STR_Code"
    
    ' --- Curated mapping (DisplayMetric -> exact USALI strings from your tenant) ---
    ' You can add/remove lines here later; the code below will keep only those that exist.
    Dim pairs As Variant
    pairs = Array( _
        Array("Occ", "Total Occ % - 100"), _
        Array("ADR", "Total ADR - 100"), _
        Array("RevPAR", "RevPAR - 100"), _
        Array("Total Rev (000's)", "Total Revenue - 000"), _
        Array("NOI (000's)", "Total Net Operating Income - 000"), _
        Array("NOI Margin", "Total Net Operating Income % - 000"), _
        Array("GOP (000's)", "Total Hotel GOP - 000"), _
        Array("GOP Margin", "Total Hotel GOP % - 000"), _
        Array("EBITDA (000's)", "Total EBITDA - 000"), _
        Array("EBITDA Margin", "Total EBITDA % - 000"), _
        Array("IBNO (000's)", "Total Income Before Non-Operating - 000"), _
        Array("Hotel NOI (000's)", "Hotel Net Operating Income - 000"), _
        Array("Paid Occ %", "Paid Occ % - 100"), _
        Array("Total Arrivals", "Total Arrivals - 100"), _
        Array("Total Hours Worked (000's)", "Total Hours Worked - 000") _
    )
    
    ' Build a fast look set of reference USALI strings
    Dim refSet As Object: Set refSet = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To refRng.Rows.Count
        Dim v As Variant: v = refRng.Cells(r, usaliCol).Value
        If Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 Then refSet(Trim$(CStr(v))) = True
        End If
    Next r
    
    ' Write rows that exist in the reference, flag those that don't
    Dim outRow As Long: outRow = 2
    Dim i As Long, disp As String, usali As String
    For i = LBound(pairs) To UBound(pairs)
        disp = CStr(pairs(i)(0))
        usali = CStr(pairs(i)(1))
        wsMap.Cells(outRow, 1).Value = disp
        wsMap.Cells(outRow, 2).Value = usali
        If Not refSet.exists(usali) Then
            wsMap.Cells(outRow, 3).Value = "NOT FOUND in Usali Reference"
            wsMap.Rows(outRow).Interior.Color = RGB(255, 245, 238) ' light highlight so you can spot it
        End If
        outRow = outRow + 1
    Next i

    ' STR mapping
    Dim strPairs As Variant
    strPairs = Array( _
        Array("Occ", "Occ"), _
        Array("ADR", "ADR"), _
        Array("RevPAR", "RevPAR"), _
        Array("Occ Index", "MPI"), _
        Array("ADR Index", "ARI"), _
        Array("RevPAR Index", "RGI"), _
        Array("% Change", "MPI % Chg"), _
        Array("% Change", "ARI % Chg"), _
        Array("% Change", "RGI % Chg") _
    )
    Dim outRow2 As Long: outRow2 = 2
    For i = LBound(strPairs) To UBound(strPairs)
        wsMap.Cells(outRow2, 5).Value = strPairs(i)(0)
        wsMap.Cells(outRow2, 6).Value = strPairs(i)(1)
        outRow2 = outRow2 + 1
    Next i

    ' (Re)bind the named ranges used by formulas
    AddOrReplaceName NAME_USALI_DISPLAY, wsMap.Range("A:A")
    AddOrReplaceName NAME_USALI_CODE, wsMap.Range("B:B")
    AddOrReplaceName NAME_STR_DISPLAY, wsMap.Range("E:E")
    AddOrReplaceName NAME_STR_CODE, wsMap.Range("F:F")
End Sub



Private Sub EnsureNamedRanges()
    ' Named ranges for USALI Map columns
    AddOrReplaceName NAME_USALI_DISPLAY, Worksheets(SH_MAP).Range("A:A")
    AddOrReplaceName NAME_USALI_CODE, Worksheets(SH_MAP).Range("B:B")
    AddOrReplaceName NAME_STR_DISPLAY, Worksheets(SH_MAP).Range("E:E")
    AddOrReplaceName NAME_STR_CODE, Worksheets(SH_MAP).Range("F:F")
End Sub

Private Sub AddMonthValidation(tgt As Range)
    With tgt.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="12"
        .ErrorTitle = "Enter a month from 1 to 12"
        .InputTitle = "Month"
        .ErrorMessage = "Please enter a whole number 1–12."
        .InputMessage = "Type 1–12"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Function SpillOrRegion(ws As Worksheet) As Range
    On Error Resume Next
    Set SpillOrRegion = ws.Range("A1#")
    If SpillOrRegion Is Nothing Then
        On Error GoTo 0
        Set SpillOrRegion = ws.Range("A1").CurrentRegion
    End If
End Function

Private Function FindHeaderCol(rng As Range, headerText As String) As Long
    Dim c As Range
    For Each c In rng.Rows(1).Cells
        If Trim$(LCase$(c.Value)) = Trim$(LCase$(headerText)) Then
            FindHeaderCol = c.Column - rng.Column + 1
            Exit Function
        End If
    Next
    FindHeaderCol = 0
End Function

Private Sub ApplyNumberFormats(ws As Worksheet, r1 As Long, r2 As Long)
    If r2 < r1 Then Exit Sub
    Dim lastCol&: lastCol = ws.Cells(r1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(r1 - 2, 1), ws.Cells(r2, lastCol)).Borders.LineStyle = xlContinuous
End Sub


' ============== Utilities ==============

Private Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Sub AddOrReplaceName(nm As String, tgt As Range)
   
    ' Remove any sheet-scoped duplicates everywhere
    KillAllSheetScoped nm
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=tgt
End Sub


Private Function SanitizeName(s As String) As String
    Dim t As String
    Dim i As Long, ch As String

    ' Replace any character that Excel would reject in a Name with an underscore
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            t = t & ch
        Else
            t = t & "_"
        End If
    Next i

    ' Excel names cannot be empty or start with a number
    If Len(t) = 0 Then t = "NA"
    If Not (Left$(t, 1) Like "[A-Za-z_]") Then t = "_" & t

    SanitizeName = t
End Function

Private Function Nz(v As Variant, Optional dflt As String = "") As String
    If IsError(v) Then
        Nz = dflt
        Exit Function
    End If
    On Error GoTo Clean
    Nz = Trim$(CStr(v))
    If Len(Nz) = 0 Then Nz = dflt
    Exit Function
Clean:
    Nz = dflt
End Function


Private Sub QuickSortText(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first: j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While LCase$(arr(i)) < LCase$(pivot): i = i + 1: Loop
        Do While LCase$(arr(j)) > LCase$(pivot): j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortText arr, first, j
    If i < last Then QuickSortText arr, i, last
End Sub

Private Function MoveValueToEnd(arr() As String, val As String) As String()
    If (Not Not arr) = 0 Then
        MoveValueToEnd = arr
        Exit Function
    End If
    Dim tmp As Collection: Set tmp = New Collection
    Dim i&, seen As Boolean
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), val, vbTextCompare) <> 0 Then tmp.Add arr(i) Else seen = True
    Next
    If seen Then tmp.Add val
    Dim out() As String: ReDim out(0 To tmp.Count - 1)
    For i = 1 To tmp.Count: out(i - 1) = tmp(i): Next
    MoveValueToEnd = out
End Function

Private Sub QuickSortByIndex(a As Variant, idx As Long)
    ' a is 1-based collection-like array of Variant() records; sort by a(i)(idx) ascending (text)
    QuickSortRec a, LBound(a), UBound(a), idx
End Sub

Private Sub QuickSortRec(a As Variant, ByVal lo As Long, ByVal hi As Long, idx As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As Variant
    i = lo: j = hi
    pivot = LCase$(CStr(a((lo + hi) \ 2)(idx)))
    Do While i <= j
        Do While LCase$(CStr(a(i)(idx))) < pivot: i = i + 1: Loop
        Do While LCase$(CStr(a(j)(idx))) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortRec a, lo, j, idx
    If i < hi Then QuickSortRec a, i, hi, idx
End Sub
' ======== Input sheet & names ========
Public Sub AutoSetupOnOpen()
    EnsureInputSheet
    EnsureNamesForInput
End Sub

Private Sub EnsureInputButtons(ws As Worksheet)
    Dim btn As Shape

    ' Report generation button
    On Error Resume Next
    Set btn = ws.Shapes("btnBuildSnapshot")
    On Error GoTo 0
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddFormControl( _
            Type:=xlButtonControl, _
            Left:=ws.Range("D1").Left, _
            Top:=ws.Range("D1").Top, _
            Width:=140, Height:=28)
        btn.Name = "btnBuildSnapshot"
    End If
    With btn
        .OnAction = "BuildFormatRun"
        .TextFrame.Characters.Text = "Generate Report"
        .TextFrame.Characters.Font.Size = 11
        .Placement = xlMove
    End With

    ' PDF export button
    On Error Resume Next
    Set btn = ws.Shapes("btnExportSnapshotPDF")
    On Error GoTo 0
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddFormControl( _
            Type:=xlButtonControl, _
            Left:=ws.Range("D2").Left, _
            Top:=ws.Range("D2").Top, _
            Width:=140, Height:=28)
        btn.Name = "btnExportSnapshotPDF"
    End If
    With btn
        .OnAction = "BuildSnapshotReportPDF"
        .TextFrame.Characters.Text = "Export to PDF"
        .TextFrame.Characters.Font.Size = 11
        .Placement = xlMove
    End With
End Sub

Private Sub EnsureInputSheet()
    Const SH_INPUT As String = "Input"
    Dim ws As Worksheet

    If SheetExists(SH_INPUT) Then
        ' Sheet already exists: do NOT recreate or clear.
        Set ws = Worksheets(SH_INPUT)

        EnsureInputButtons ws
        ws.Visible = xlSheetVisible
        Exit Sub
    End If

    ' Create the Input sheet for the first time
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = SH_INPUT
    ws.Visible = xlSheetVisible

    With ws
        .Range("A1").Value = "Select Month:"
        .Range("A2").Value = "Select Year:"
        .Range("B1").Value = Format(Date, "MMMM")
        .Range("B2").Value = Year(Date)

        ' Month dropdown (full month names)
        With .Range("B1").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="January,February,March,April,May,June,July,August,September,October,November,December"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ErrorMessage = "Pick a month name"
        End With

        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 16
        .Range("A1:A2").Font.Bold = True
    End With
    EnsureInputButtons ws
End Sub


Private Sub EnsureNamesForInput()

    ' Nuke any sheet-scoped duplicates first
    KillAllSheetScoped "MonthText"
    KillAllSheetScoped "YearNum"
    KillAllSheetScoped "MonthNum"

    ' Point MonthText to Input!B1 (full month); YearNum to Input!B2
    AddOrReplaceName "MonthText", Worksheets("Input").Range("B1")
    AddOrReplaceName "YearNum", Worksheets("Input").Range("B2")

    ' MonthNum as formula: MATCH() over an inline month list
    On Error Resume Next: ThisWorkbook.Names("MonthNum").Delete: On Error GoTo 0
    ThisWorkbook.Names.Add Name:="MonthNum", _
        RefersTo:="=MATCH(Input!B1,{""January"",""February"",""March"",""April"",""May"",""June"",""July"",""August"",""September"",""October"",""November"",""December""},0)"
End Sub
Private Sub KillAllSheetScoped(ByVal nm As String)
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        sh.Names(nm).Delete
        On Error GoTo 0
    Next sh
End Sub




Private Sub ApplyBandSeparators(ws As Worksheet, topRow As Long, bottomRow As Long, startCol As Long, metrics As Variant)
    Dim metricsPerBand&: metricsPerBand = (UBound(metrics) - LBound(metrics) + 1)

     With ws.Range(ws.Cells(topRow, 3), ws.Cells(bottomRow, 3)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium
    End With

   
    ' Vertical separator on the LEFT of the first value column (F): left edge of column startCol
    With ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, startCol)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    ' Band boundaries: after band 1–4
    Dim b As Long, boundaryCol As Long
    For b = 1 To 4
        boundaryCol = startCol + b * metricsPerBand - 1
        With ws.Range(ws.Cells(topRow, boundaryCol), ws.Cells(bottomRow, boundaryCol)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlMedium
        End With
    Next b

    ' Thick outline around full table
    ws.Range(ws.Cells(topRow, 2), ws.Cells(bottomRow, startCol + 5 * metricsPerBand - 1)).BorderAround Weight:=xlThick
End Sub

Private Sub ApplyStrSeparators(ws As Worksheet, topRow As Long, bottomRow As Long, startCol As Long, secWidths As Variant, Optional includeRoomsMgrDivider As Boolean = False)
    Dim lastCol&, i&
    lastCol = startCol - 1
    For i = LBound(secWidths) To UBound(secWidths)
        lastCol = lastCol + CLng(secWidths(i))
    Next i

    If includeRoomsMgrDivider Then
        With ws.Range(ws.Cells(topRow, 3), ws.Cells(bottomRow, 3)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    End If

    With ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, startCol)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    Dim boundaryCol&: boundaryCol = startCol
    For i = LBound(secWidths) To UBound(secWidths) - 1
        boundaryCol = boundaryCol + secWidths(i) - 1
        With ws.Range(ws.Cells(topRow, boundaryCol), ws.Cells(bottomRow, boundaryCol)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        boundaryCol = boundaryCol + 1
    Next i

    ws.Range(ws.Cells(topRow, 2), ws.Cells(bottomRow, lastCol)).BorderAround Weight:=xlThick
End Sub

Private Sub ApplyStrDotDividers(ws As Worksheet, topRow As Long, bottomRow As Long, startCol As Long, sectionCount As Long, secWidth As Long)
    Dim s As Long, baseCol As Long
    For s = 0 To sectionCount - 1
        baseCol = startCol + s * secWidth
        With ws.Range(ws.Cells(topRow, baseCol + 1), ws.Cells(bottomRow, baseCol + 1)).Borders(xlEdgeRight)
            .LineStyle = xlDot
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(topRow, baseCol + 3), ws.Cells(bottomRow, baseCol + 3)).Borders(xlEdgeRight)
            .LineStyle = xlDot
            .Weight = xlThin
        End With
    Next s
End Sub

Public Sub BuildFormatRun()
    AutoSetupOnOpen         ' make sure Input and names exist
    BuildSnapshot           ' your existing builder (now uses names)
    FormatSnapshotShell     ' header bar + band header finalize + spacing + gridlines
End Sub

Private Function GetNameVal(ByVal nm As String) As Variant
    Dim n As name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then
        GetNameVal = CVErr(xlErrName)
    Else
        GetNameVal = Evaluate(n.refersTo) ' works for ranges or formulas
    End If
End Function


Private Sub FormatSnapshotShell()
    Const RED_HEX As String = "E03C31"
     Dim red&: red = ColorHex(RED_HEX)
    Dim ws As Worksheet: Set ws = Worksheets(SH_SNAP)

    ' Unfreeze panes & clear splits
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
    End With

    ' Row 1 blank, Row 3 exactly one blank spacer
    ws.Rows(1).RowHeight = ws.StandardHeight
    ws.Rows(3).RowHeight = ws.StandardHeight

    ' Find the rightmost used column across the sheet
    Dim lastCol As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not lastCell Is Nothing Then
        lastCol = lastCell.Column
    Else
        lastCol = 20
    End If

    ' Red header bar B2 : lastCol
    With ws.Range(ws.Cells(2, 2), ws.Cells(2, lastCol))
        .Interior.Color = red
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 32
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        .Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    End With

    ' Pull Month/Year from names on Input sheet
    Dim yr As Long, mo As Long
    yr = CLng(GetNameVal("YearNum"))
    mo = CLng(GetNameVal("MonthNum"))
    Dim monthText As String: monthText = Format(DateSerial(yr, mo, 1), "MMMM")

    ' Title (left)
 
With ws.Cells(2, 2)
    .Value = "Hospitality Portfolio Snapshot"
    .Font.Size = 16
End With

' "As of:" label
ws.Cells(2, 6).Value = "As of:"
ws.Cells(2, 6).Font.Size = 14

' DO NOT set values in G2/H2 here.
' G2/H2 are already formulas via EnsureSnapHeaderNames:
'   G2: =MonthText   (full month from Input, e.g., "June")
'   H2: =YearNum     (e.g., 2025)
ws.Cells(2, 7).Font.Size = 14   ' G2
ws.Cells(2, 8).Font.Size = 14   ' H2


    ' Created (far right)
    With ws.Cells(2, lastCol)
        .Value = "Created: " & Format(Now, "mmm d, yyyy h:mm AM/PM")
        .HorizontalAlignment = xlRight
        .Font.Size = 12
    End With

    ' Left block widths + keep Code hidden
    ws.Columns(2).ColumnWidth = 28 ' Hotel
    ws.Columns(3).ColumnWidth = 9  ' Rooms
    ws.Columns(4).ColumnWidth = 18 ' Manager
    ws.Columns(5).Hidden = True    ' Code
End Sub

Public Sub HardResetSnapshotConfig()
    ' 1) Remove any old inputs on Snapshot and kill validations
    RemoveLegacySnapshotInputs

    ' 2) Delete any sheet-scoped duplicates of MonthNum/YearNum/MonthText
    RemoveSheetScopedNames Array("MonthNum", "YearNum", "MonthText", "Snap_MonthText", "Snap_YearNum")

    ' 3) Ensure Input sheet + workbook-scoped names exist
    EnsureInputSheet
    EnsureNamesForInput

    MsgBox "Reset complete. Now click 'Generate Snapshot' on the Input sheet.", vbInformation
End Sub

Private Sub RemoveLegacySnapshotInputs()
    On Error Resume Next
    If SheetExists("Snapshot") Then
        With Worksheets("Snapshot")
            .Range("A1:B2").Validation.Delete
            .Range("A1:B2").ClearContents
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub RemoveSheetScopedNames(nameList As Variant)
    Dim sh As Worksheet, nm As Variant
    For Each sh In ThisWorkbook.Worksheets
        For Each nm In nameList
            On Error Resume Next
            sh.Names(CStr(nm)).Delete
            On Error GoTo 0
        Next nm
    Next sh
End Sub
Public Sub FixNamesNow()
    ' Kill all sheet-scoped duplicates
    KillAllSheetScoped "MonthText"
    KillAllSheetScoped "YearNum"
    KillAllSheetScoped "MonthNum"
    KillAllSheetScoped "UsaliMap_Display"
    KillAllSheetScoped "UsaliMap_Code"
    KillAllSheetScoped "Snap_MonthText"
    KillAllSheetScoped "Snap_YearNum"

    ' Recreate input names cleanly
    EnsureNamesForInput

    ' Rebind USALI map names cleanly
    EnsureUsaliMap
    EnsureNamedRanges

    ' Quick sanity checks
    Debug.Print "MonthText -> "; GetNameVal("MonthText")
    Debug.Print "YearNum   -> "; GetNameVal("YearNum")
    Debug.Print "MonthNum  -> "; GetNameVal("MonthNum")
End Sub


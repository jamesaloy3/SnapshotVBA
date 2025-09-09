Option Explicit

Private Const SH_SNAP As String = "Snapshot"
Private Const LOGO_SHAPE As String = "Logo"
Private Const FOOTER_TEXT As String = "Confidential - For internal use only"

' Main entry point called by button on Input sheet
Public Sub BuildSnapshotReportPDF()
    ' Use the snapshot already visible without triggering a rebuild
    Dim ws As Worksheet
    Set ws = Worksheets(SH_SNAP)

    ' Apply page setup and capture logo path for later cleanup
    Dim logoPath As String
    logoPath = ApplyReportPageSetup(ws)
    ApplyTablePageBreaks ws
    DoEvents

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "SnapshotReport_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, OpenAfterPublish:=False
    ' clean up temporary logo file
    If Len(logoPath) > 0 Then On Error Resume Next: Kill logoPath: On Error GoTo 0
    MsgBox "Snapshot report exported to:" & vbCrLf & pdfPath, vbInformation
End Sub

' Configure header/footer, orientation, paper size, logo etc.
Private Function ApplyReportPageSetup(ws As Worksheet) As String
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperTabloid
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .CenterFooter = FOOTER_TEXT
        .PrintTitleRows = "$1:$3"
        ApplyReportPageSetup = InsertLogo(.Parent)
    End With
End Function

' Insert logo on every page via header picture
Private Function InsertLogo(ws As Worksheet) As String
    On Error GoTo NoLogo
    Dim shp As Shape
    Set shp = ws.Shapes(LOGO_SHAPE)
    Dim tmpPath As String
    tmpPath = Environ("TEMP") & Application.PathSeparator & "snapLogo.png"
    shp.Export Filename:=tmpPath, FilterName:="PNG"
    With ws.PageSetup
        .LeftHeaderPicture.Filename = tmpPath
        .LeftHeader = "&G"
    End With
    InsertLogo = tmpPath
    Exit Function
NoLogo:
    ' Silently ignore if logo not found
    InsertLogo = ""
End Function

' Insert a manual break after the second table and provide spacing between sections
Private Sub ApplyTablePageBreaks(ws As Worksheet)
    ws.ResetAllPageBreaks
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim r As Long, headerCount As Long
    headerCount = 0
    For r = 4 To lastRow
        Select Case UCase$(Trim$(ws.Cells(r, 2).Value))
            Case "HOTEL", "MANAGER", "MARKET"
                ' add a little space before each table header
                If r > 1 Then
                    ws.Rows(r - 1).RowHeight = ws.StandardHeight
                    ws.Rows(r - 1).Borders.LineStyle = xlLineStyleNone
                End If
                headerCount = headerCount + 1
                ' first two tables on page 1, break before third table
                If headerCount = 3 Then
                    ws.HPageBreaks.Add Before:=ws.Rows(r)
                End If
        End Select
    Next r
End Sub


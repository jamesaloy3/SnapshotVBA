Option Explicit

Private Const SH_SNAP As String = "Snapshot"
Private Const LOGO_SHAPE As String = "Logo"
Private Const FOOTER_TEXT As String = "Confidential - For internal use only"

' Main entry point called by button on Input sheet
Public Sub BuildSnapshotReportPDF()
    ' Ensure snapshot content is up to date
    BuildSnapshot

    Dim ws As Worksheet
    Set ws = Worksheets(SH_SNAP)

    ApplyReportPageSetup ws
    ApplyTablePageBreaks ws

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "SnapshotReport_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "Snapshot report exported to:" & vbCrLf & pdfPath, vbInformation
End Sub

' Configure header/footer, orientation, paper size, logo etc.
Private Sub ApplyReportPageSetup(ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperTabloid
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .CenterFooter = FOOTER_TEXT
        .PrintTitleRows = "$1:$3"
        InsertLogo .Parent
    End With
End Sub

' Insert logo on every page via header picture
Private Sub InsertLogo(ws As Worksheet)
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
    Kill tmpPath
    Exit Sub
NoLogo:
    ' Silently ignore if logo not found
End Sub

' Ensure page breaks occur before each major table so headers repeat cleanly
Private Sub ApplyTablePageBreaks(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim r As Long, firstTable As Boolean
    firstTable = False
    For r = 4 To lastRow
        Select Case UCase$(Trim$(ws.Cells(r, 2).Value))
            Case "HOTEL", "MANAGER", "MARKET"
                If firstTable Then
                    ws.HPageBreaks.Add Before:=ws.Rows(r)
                Else
                    firstTable = True
                End If
        End Select
    Next r
End Sub


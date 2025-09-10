Option Explicit

Private Const FOOTER_TEXT As String = "Confidential - For internal use only"

' Build a consolidated PDF for all Cash Forecast Variance reports
Public Sub BuildCashForecastVariancePDF()
    Dim ws As Worksheet, cfvSheets As Collection
    Set cfvSheets = New Collection
    
    ' Collect all sheets that appear to be CFV reports (look for HotelName named range)
    For Each ws In ThisWorkbook.Worksheets
        If SheetIsCashForecastVariance(ws) Then cfvSheets.Add ws
    Next ws
    
    If cfvSheets.Count = 0 Then
        MsgBox "No Cash Forecast Variance reports found", vbInformation
        Exit Sub
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To cfvSheets.Count)
    Dim i As Long
    For i = 1 To cfvSheets.Count
        arr(i) = cfvSheets(i).Name
        ApplyCfvPageSetup cfvSheets(i)
    Next i
    
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "CashForecastVariance_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    
    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet
    ThisWorkbook.Sheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                                    Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=False
    prevSheet.Select
    MsgBox "Cash Forecast Variance PDF exported to:" & vbCrLf & pdfPath, vbInformation
End Sub

' Determine if a worksheet is a CFV report by checking for a local HotelName name
Private Function SheetIsCashForecastVariance(ws As Worksheet) As Boolean
    On Error Resume Next
    SheetIsCashForecastVariance = Not ws.Range("HotelName") Is Nothing
    On Error GoTo 0
End Function

' Ensure each report prints on a single page with consistent settings
Private Sub ApplyCfvPageSetup(ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterFooter = FOOTER_TEXT
    End With
End Sub


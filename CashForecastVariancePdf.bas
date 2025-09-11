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
    Dim firstSheet As Worksheet
    Set firstSheet = cfvSheets(1)
    Dim yr As Long, mo As Long
    Dim monthName As String
    On Error Resume Next
    yr = CLng(firstSheet.Range("RYear_YYYY").Value)
    monthName = CStr(firstSheet.Range("Month_MMMM").Value)
    On Error GoTo 0
    If Len(monthName) > 0 Then mo = Month(DateValue("1 " & monthName & " " & yr))
    Dim prefix As String
    If yr > 0 And mo > 0 Then
        prefix = Format(DateSerial(yr, mo, 1), "mmYY")
    Else
        prefix = Format(Now, "mmYY")
    End If

    Dim folderPath As String
    folderPath = ThisWorkbook.Path & Application.PathSeparator & "CashForecastVariance"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Dim pdfPath As String
    pdfPath = folderPath & Application.PathSeparator & prefix & "_CashForecastVariance.pdf"

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


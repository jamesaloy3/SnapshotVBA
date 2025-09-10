Option Explicit

' Constants for sheet and name references
Private Const SH_PROPS As String = "My Properties"
Private Const SH_USALI As String = "Usali Reference"
Private Const SH_MAP As String = "USALI Map"

Private Const NAME_USALI_DISPLAY As String = "UsaliMap_Display"
Private Const NAME_USALI_CODE As String = "UsaliMap_Code"

' Main entry point
Public Sub BuildCashForecastVariance()
    Dim calcState As XlCalculation
    Dim scrn As Boolean, enEvents As Boolean
    calcState = Application.Calculation
    scrn = Application.ScreenUpdating
    enEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    EnsureUsaliMap

    Dim props As Collection
    Set props = GetHotelList()
    If props Is Nothing Then GoTo CleanFail

    Dim tplPath As String
    tplPath = ThisWorkbook.Path & Application.PathSeparator & "CashForecastVariance_Template.xlsx"
    Dim tplWB As Workbook
    Set tplWB = Workbooks.Open(Filename:=tplPath, ReadOnly:=True)
    Dim tplSheet As Worksheet
    Set tplSheet = tplWB.Worksheets(1)

    Dim prop As Variant, ws As Worksheet
    For Each prop In props
        tplSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ws.Name = Left$(SanitizeName(CStr(prop(0))), 31)
        LocalizeCFVNames ws
        ws.Range("HotelName").Value = CStr(prop(0))
        ws.Range("PropCode").Value = CStr(prop(1))
        ws.Range("TimeAgg").Value = "Total Year"
        ws.Range("RYear_YYYY").Value = Year(Date)
        FillMetric ws, "Metric1_DisplayName", "Metric1_Values"
        FillMetric ws, "Metric2_DisplayName", "Metric2_Values"
        FillMetric ws, "Metric3_DisplayName", "Metric3_Values"
    Next prop

    tplWB.Close SaveChanges:=False

CleanExit:
    Application.Calculation = calcState
    Application.ScreenUpdating = scrn
    Application.EnableEvents = enEvents
    Exit Sub

CleanFail:
    MsgBox "Failed to build Cash Forecast Variance report", vbExclamation
    On Error Resume Next
    tplWB.Close SaveChanges:=False
    GoTo CleanExit
End Sub

' Fill a metric value range with SP.FINANCIALS formulas
Private Sub FillMetric(ws As Worksheet, displayName As String, valueName As String)
    Dim dispCell As Range, valRng As Range, verRng As Range
    Set dispCell = ws.Range(displayName)
    Set valRng = ws.Range(valueName)
    Set verRng = ws.Range("Version")

    Dim baseFormula As String
    baseFormula = "=SP.FINANCIALS(PropCode,INDEX(" & NAME_USALI_CODE & ",MATCH(" & dispCell.Address(False, False) & "," & NAME_USALI_DISPLAY & ",0)),TimeAgg,RYear_YYYY,"
    Dim i As Long
    For i = 1 To valRng.Rows.Count
        valRng.Cells(i, 1).Formula = baseFormula & verRng.Cells(i, 1).Address(False, False) & ")"
    Next i
End Sub

' Convert copied template names to sheet-scoped names
Private Sub LocalizeCFVNames(ws As Worksheet)
    Dim nmList As Variant
    nmList = Array("ReportArea", "HotelName", "PropCode", _
                    "Metric1_DisplayName", "Metric1_Values", _
                    "Metric2_DisplayName", "Metric2_Values", _
                    "Metric3_DisplayName", "Metric3_Values", _
                    "Version", "TimeAgg", "RYear_YYYY", "Month_MMMM")
    Dim nmName As Variant, nmObj As Name
    For Each nmName In nmList
        On Error Resume Next
        Set nmObj = ThisWorkbook.Names(CStr(nmName))
        On Error GoTo 0
        If Not nmObj Is Nothing Then
            Dim refSheet As String
            refSheet = Split(Split(nmObj.RefersTo, "'")(1), "'")(0)
            ws.Names.Add Name:=CStr(nmName), RefersTo:=Replace(nmObj.RefersTo, "'" & refSheet & "'!", "'" & ws.Name & "'!")
            nmObj.Delete
        End If
    Next nmName
End Sub

' Build list of hotels from My Properties
Private Function GetHotelList() As Collection
    Dim ws As Worksheet
    If Not SheetExists(SH_PROPS) Then
        MsgBox "My Properties sheet not found", vbExclamation
        Exit Function
    End If
    Set ws = Worksheets(SH_PROPS)

    Dim propsRng As Range
    Set propsRng = SpillOrRegion(ws)
    If propsRng Is Nothing Then Exit Function

    Dim cCode&, cHotel&, cMgmt&
    cCode = FindHeaderCol(propsRng, "Code")
    cHotel = FindHeaderCol(propsRng, "HotelName")
    cMgmt = FindHeaderCol(propsRng, "ManagementCompany")
    If cCode * cHotel * cMgmt = 0 Then Exit Function

    Dim col As New Collection
    Dim r As Long, lastR As Long
    lastR = propsRng.Rows.Count
    For r = 2 To lastR
        Dim nm$, cd$, mgmt$
        nm = Nz(propsRng.Cells(r, cHotel).Value)
        cd = Nz(propsRng.Cells(r, cCode).Value)
        mgmt = Nz(propsRng.Cells(r, cMgmt).Value)
        If Len(nm) > 0 And Len(cd) > 0 Then
            If StrComp(mgmt, "Stonebridge Legacy", vbTextCompare) <> 0 Then
                col.Add Array(nm, cd)
            End If
        End If
    Next r
    Set GetHotelList = col
End Function

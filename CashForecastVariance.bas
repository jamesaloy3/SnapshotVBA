Option Explicit

' Constants for sheet and name references
Private Const SH_PROPS As String = "My Properties"
Private Const SH_USALI As String = "Usali Reference"
Private Const SH_MAP As String = "USALI Map"
Private Const SH_INPUT As String = "CFV Input"

' Expose USALI mapping names globally so other modules can access them
Public Const NAME_USALI_DISPLAY As String = "UsaliMap_Display"
Public Const NAME_USALI_CFV_DISPLAY As String = "UsaliMap_CFV_Display"
Public Const NAME_USALI_CODE As String = "UsaliMap_Code"

Private Const NAME_CFV_MONTH As String = "CFV_Month"
Private Const NAME_CFV_MET1 As String = "CFV_Metric1"
Private Const NAME_CFV_MET2 As String = "CFV_Metric2"
Private Const NAME_CFV_MET3 As String = "CFV_Metric3"

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
    EnsureCfvInputSheet

    Dim monthName As String
    monthName = Nz(Range(NAME_CFV_MONTH).Value)
    Dim timeAgg As String, displayMonth As String
    If StrComp(monthName, "Total Year", vbTextCompare) = 0 Then
        timeAgg = "Total Year"
        displayMonth = Format(DateAdd("m", -1, Date), "MMMM")
    Else
        timeAgg = "Month"
        displayMonth = monthName
    End If

    Dim mCode(1 To 3) As String, mDisp(1 To 3) As String
    mCode(1) = Nz(Range(NAME_CFV_MET1).Value)
    mCode(2) = Nz(Range(NAME_CFV_MET2).Value)
    mCode(3) = Nz(Range(NAME_CFV_MET3).Value)

    Dim i As Long
    For i = 1 To 3
        mDisp(i) = CfvDisplayFromCode(mCode(i))
    Next i

    Dim props As Collection
    Set props = GetHotelList()
    If props Is Nothing Then GoTo CleanFail

    ' Remove any previously generated report sheets to avoid name collisions
    DeleteExistingCfvReports props

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
        ws.Range("TimeAgg").Value = timeAgg
        ws.Range("RYear_YYYY").Value = Year(Date)
        ws.Range("Month_MMMM").Value = displayMonth
        ws.Range("Metric1_DisplayName").Value = mDisp(1)
        ws.Range("Metric2_DisplayName").Value = mDisp(2)
        ws.Range("Metric3_DisplayName").Value = mDisp(3)
        ws.Columns(3).Hidden = True
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


' Convert copied template names to sheet-scoped names
Private Sub LocalizeCFVNames(ws As Worksheet)
    Dim nmObj As Name, nmName As Variant
    Dim toProcess As New Collection

    ' Collect workbook-level names referring to this sheet
    For Each nmObj In ThisWorkbook.Names
        If InStr(1, nmObj.RefersTo, "'" & ws.Name & "'!", vbTextCompare) > 0 Then
            toProcess.Add nmObj.Name
        End If
    Next nmObj

    ' Localize each name to the worksheet scope
    For Each nmName In toProcess
        Set nmObj = ThisWorkbook.Names(CStr(nmName))
        ws.Names.Add Name:=nmObj.Name, RefersTo:=nmObj.RefersTo
        nmObj.Delete
    Next nmName
End Sub

' Delete any existing report sheets matching the property names
Private Sub DeleteExistingCfvReports(props As Collection)
    Dim prop As Variant, shName As String
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For Each prop In props
        shName = Left$(SanitizeName(CStr(prop(0))), 31)
        If SheetExists(shName) Then
            Worksheets(shName).Delete
        End If
    Next prop

    Application.DisplayAlerts = prevAlerts
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

Private Sub EnsureCfvInputSheet()
    Dim ws As Worksheet
    If Not SheetExists(SH_INPUT) Then
        Set ws = Worksheets.Add(Before:=Worksheets(1))
        ws.Name = SH_INPUT
        ws.Range("A1").Value = "Month"
        ws.Range("A2").Value = "Metric 1 (USALI)"
        ws.Range("A3").Value = "Metric 2 (USALI)"
        ws.Range("A4").Value = "Metric 3 (USALI)"
    Else
        Set ws = Worksheets(SH_INPUT)
    End If

    With ws.Range("B1")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=Join(Array("January","February","March","April","May","June","July","August","September","October","November","December","Total Year"), ",")
    End With

    With ws.Range("B2:B4")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:="=" & NAME_USALI_CODE
    End With

    AddOrReplaceName NAME_CFV_MONTH, ws.Range("B1")
    AddOrReplaceName NAME_CFV_MET1, ws.Range("B2")
    AddOrReplaceName NAME_CFV_MET2, ws.Range("B3")
    AddOrReplaceName NAME_CFV_MET3, ws.Range("B4")

    Dim btn As Button, buildBtn As Button
    For Each btn In ws.Buttons
        If InStr(1, btn.OnAction, "BuildCashForecastVariance", vbTextCompare) > 0 Then
            If buildBtn Is Nothing Then
                Set buildBtn = btn
                buildBtn.Name = "btnCfvGenerateReport"
            Else
                btn.Delete
            End If
        End If
    Next btn
    If buildBtn Is Nothing Then
        Set buildBtn = ws.Buttons.Add(ws.Range("A6").Left, ws.Range("A6").Top, 150, 30)
    End If
    With buildBtn
        .Caption = "Generate Report"
        .OnAction = "BuildCashForecastVariance"
        .Name = "btnCfvGenerateReport"
    End With
End Sub

Private Function UsaliDisplayFromCode(code As String) As String
    On Error GoTo Clean
    Dim codes As Range, displays As Range, idx As Variant
    Set codes = Range(NAME_USALI_CODE)
    Set displays = Range(NAME_USALI_DISPLAY)
    idx = Application.Match(code, codes, 0)
    If Not IsError(idx) Then
        UsaliDisplayFromCode = displays.Cells(CLng(idx), 1).Value
    Else
        UsaliDisplayFromCode = code
    End If
    Exit Function
Clean:
    UsaliDisplayFromCode = code
End Function

Private Function CfvDisplayFromCode(code As String) As String
    On Error GoTo Clean
    Dim codes As Range, displays As Range, idx As Variant
    Set codes = Range(NAME_USALI_CODE)
    Set displays = Range(NAME_USALI_CFV_DISPLAY)
    idx = Application.Match(code, codes, 0)
    If Not IsError(idx) Then
        CfvDisplayFromCode = displays.Cells(CLng(idx), 1).Value
    Else
        CfvDisplayFromCode = code
    End If
    Exit Function
Clean:
    CfvDisplayFromCode = code
End Function

Private Sub EnsureUsaliMap()
    Dim wsMap As Worksheet, wsRef As Worksheet, refRng As Range
    Dim headers As Range, usaliCol As Long

    If Not SheetExists(SH_USALI) Then Err.Raise vbObjectError + 500, , "'" & SH_USALI & "' sheet not found."
    Set wsRef = Worksheets(SH_USALI)
    Set refRng = SpillOrRegion(wsRef)
    usaliCol = FindHeaderCol(refRng, "usali")
    If usaliCol = 0 Then Err.Raise vbObjectError + 501, , "'" & SH_USALI & "' missing 'usali' header."

    If Not SheetExists(SH_MAP) Then
        Set wsMap = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsMap.Name = SH_MAP
    Else
        Set wsMap = Worksheets(SH_MAP)
    End If
    wsMap.Cells.Clear

    wsMap.Range("A1").Value = "DisplayMetric"
    wsMap.Range("B1").Value = "CFV_DisplayMetric"
    wsMap.Range("C1").Value = "USALI"
    wsMap.Range("D1").Value = "Notes"

    Dim pairs As Variant
    pairs = Array( _
        Array("Occ", "Occ", "Total Rooms Sold - 100"), _
        Array("ADR", "ADR", "Total ADR - 100"), _
        Array("RevPAR", "RevPAR", "RevPAR - 100"), _
        Array("Total Rev (000's)", "Total Rev", "Total Revenue - 000"), _
        Array("NOI (000's)", "NOI", "Total EBITDA Less Reserves - 000"), _
        Array("NOI Margin", "NOI Margin", "Total Net Operating Income % - 000"), _
        Array("GOP (000's)", "GOP", "Total Hotel GOP - 000"), _
        Array("GOP Margin", "GOP Margin", "Total Hotel GOP % - 000"), _
        Array("EBITDA (000's)", "EBITDA", "Total EBITDA - 000"), _
        Array("EBITDA Margin", "EBITDA Margin", "Total EBITDA % - 000"), _
        Array("IBNO (000's)", "IBNO", "Total Income Before Non-Operating - 000"), _
        Array("Hotel NOI (000's)", "Hotel NOI", "Hotel Net Operating Income - 000"), _
        Array("Paid Occ %", "Paid Occ %", "Paid Occ % - 100"), _
        Array("Total Arrivals", "Total Arrivals", "Total Arrivals - 100"), _
        Array("Total Hours Worked (000's)", "Total Hours Worked", "Total Hours Worked - 000") _
    )

    Dim refSet As Object
    Set refSet = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To refRng.Rows.Count
        Dim v As Variant: v = refRng.Cells(r, usaliCol).Value
        If Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 Then refSet(Trim$(CStr(v))) = True
        End If
    Next r

    Dim outRow As Long: outRow = 2
    Dim i As Long, disp As String, cfvDisp As String, usali As String
    For i = LBound(pairs) To UBound(pairs)
        disp = pairs(i)(0)
        cfvDisp = pairs(i)(1)
        usali = pairs(i)(2)
        wsMap.Cells(outRow, 1).Value = disp
        wsMap.Cells(outRow, 2).Value = cfvDisp
        wsMap.Cells(outRow, 3).Value = usali
        If Not refSet.exists(usali) Then
            wsMap.Cells(outRow, 4).Value = "NOT FOUND in Usali Reference"
        End If
        outRow = outRow + 1
    Next i

    AddOrReplaceName NAME_USALI_DISPLAY, wsMap.Range("A:A")
    AddOrReplaceName NAME_USALI_CFV_DISPLAY, wsMap.Range("B:B")
    AddOrReplaceName NAME_USALI_CODE, wsMap.Range("C:C")
End Sub


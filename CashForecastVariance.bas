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

' ==== Helpers copied from Snapshot module ====

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

Private Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Function SanitizeName(s As String) As String
    Dim t As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            t = t & ch
        Else
            t = t & "_"
        End If
    Next i
    If Len(t) = 0 Then t = "NA"
    If Not (Left$(t, 1) Like "[A-Za-z_]") Then t = "_" & t
    SanitizeName = t
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
    wsMap.Range("B1").Value = "USALI"
    wsMap.Range("C1").Value = "Notes"

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
    Dim i As Long, disp As String, usali As String
    For i = LBound(pairs) To UBound(pairs)
        disp = pairs(i)(0)
        usali = pairs(i)(1)
        wsMap.Cells(outRow, 1).Value = disp
        wsMap.Cells(outRow, 2).Value = usali
        If Not refSet.exists(usali) Then
            wsMap.Cells(outRow, 3).Value = "NOT FOUND in Usali Reference"
        End If
        outRow = outRow + 1
    Next i

    AddOrReplaceName NAME_USALI_DISPLAY, wsMap.Range("A:A")
    AddOrReplaceName NAME_USALI_CODE, wsMap.Range("B:B")
End Sub

Private Sub AddOrReplaceName(nm As String, tgt As Range)
    KillAllSheetScoped nm
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=tgt
End Sub

Private Sub KillAllSheetScoped(ByVal nm As String)
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        sh.Names(nm).Delete
        On Error GoTo 0
    Next sh
End Sub


Attribute VB_Name = "Pivot"
Option Explicit

'======================== SETTINGS ========================
Private Const SRC_SHEET As String = "Helper Sheet"
Private Const FLAT_SHEET As String = "Pivot_Source_Data" ' The hidden auto-prep sheet
Private Const PIVOT_SHEET As String = "Pivot - Summary"
Private Const PT_NAME As String = "pt_Summary"
Private Const FLD_SPECIFIC_TRADE As String = "Specific Trade Breakdown"

' Source header names (must match exactly)
Private Const FLD_BUILDING As String = "Building"
Private Const FLD_WHO As String = "Who can perform work"
Private Const FLD_TOTAL_HOURS As String = "Total Hours"
Private Const FLD_TOTAL_MATCOST As String = "Annual PM Material Cost"
Private Const FLD_ZONE As String = "Zone"

' Captions assigned to the two data fields in the Pivot
Private Const CAP_HOURS As String = "Sum of Total Hours"
Private Const CAP_COST As String = "Sum of PM Material Cost"

' Optional default Zone selection: "" = All
Private Const DEFAULT_ZONE As String = "" ' e.g., "North"
'====================== /SETTINGS =========================

Public Sub PivotBuilder_Rebuild()
    Dim WsSrc As Worksheet, wsPvt As Worksheet, wsFlat As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range, origRange As Range
    Dim pt As PivotTable
    Dim panelStart As Range
    Dim startCol As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. FLATTEN DATA SILENTLY
    Call AutoFlattenData

    ' 2. SOURCE VALIDATION
    On Error Resume Next
    Set wsFlat = ThisWorkbook.Worksheets(FLAT_SHEET)
    On Error GoTo 0

    ' Stop using UsedRange. Calculate exact bounds.
    lastCol = wsFlat.Cells(1, wsFlat.Columns.count).End(xlToLeft).Column
    lastRow = wsFlat.Cells(wsFlat.Rows.count, 1).End(xlUp).Row

    ' Fail-safe: If the sheet is entirely empty, give the pivot at least one blank row to read
    If lastRow < 2 Then lastRow = 2

    Set srcRange = wsFlat.Range(wsFlat.Cells(1, 1), wsFlat.Cells(lastRow, lastCol))

    ' 3. CLEAN UP OLD NAMED RANGE (To permanently kill the cache bug)
    On Error Resume Next
    ThisWorkbook.names("TempPivotSource").Delete
    On Error GoTo 0

    ' 4. RECREATE PIVOT SHEET
    Set wsPvt = RecreatePivotSheet_safe(PIVOT_SHEET)

    ' 5. BUILD PIVOT
    Set pt = BuildVerticalPivot_safe(srcRange, wsPvt, PT_NAME, wsPvt.Range("A6"))

    ' 6. WRITE TOTALS PANEL
    Set WsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    lastRow = WsSrc.Cells(WsSrc.Rows.count, "A").End(xlUp).Row
    lastCol = WsSrc.Cells(1, WsSrc.Columns.count).End(xlToLeft).Column
    Set origRange = WsSrc.Range(WsSrc.Cells(1, 1), WsSrc.Cells(lastRow, lastCol))

    startCol = pt.TableRange2.Column + pt.TableRange2.Columns.count + 2
    If startCol < 4 Then startCol = 4
    Set panelStart = wsPvt.Cells(6, startCol)

    WriteTotalsPanel_FromSource _
        pt:=pt, _
        WsOut:=wsPvt, _
        StartCell:=panelStart, _
        WsSrc:=WsSrc, _
        SrcData:=origRange, _
        FldZone:=FLD_ZONE, _
        FldBuilding:=FLD_BUILDING, _
        FldHours:=FLD_TOTAL_HOURS, _
        FldCost:=FLD_TOTAL_MATCOST

     ' 7. CLEAN UP
    ' We no longer activate anything here to prevent screen jerking
    ' On Error Resume Next
    ' currentScreen.Activate
    ' If Err.Number <> 0 Then
    '    wsPvt.Activate
    ' End If
    ' On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub PivotBuilder_Refresh()
    Call PivotBuilder_Rebuild
End Sub

' =========================================================================
' The invisible data prep engine.
' =========================================================================
Private Sub AutoFlattenData()
    Dim WsSrc As Worksheet, wsDest As Worksheet
    Dim arrData As Variant, arrOut() As Variant
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, outRow As Long
    Dim colWho As Long, colTotal As Long, colMat As Long
    Dim colBld As Long, colZone As Long
    Dim colMulti1 As Long, colMulti2 As Long

    Set WsSrc = ThisWorkbook.Sheets(SRC_SHEET)

    lastRow = WsSrc.Cells(WsSrc.Rows.count, "A").End(xlUp).Row
    lastCol = WsSrc.Cells(1, WsSrc.Columns.count).End(xlToLeft).Column

    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    ' --- CRITICAL FIX: Completely Nuke the Hidden Sheet to Destroy Ghost Data ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(FLAT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsDest = ThisWorkbook.Sheets.Add(After:=WsSrc)
    wsDest.Name = FLAT_SHEET
    wsDest.Visible = xlSheetHidden

    ' Quietly return user to where they were so Add doesn't hijack focus
    currentSheet.Activate

    If lastRow < 2 Then
        ' Copy headers only and exit so pivot creates empty
        WsSrc.Rows(1).Copy wsDest.Rows(1)
        wsDest.Cells(1, lastCol + 1).Value = FLD_SPECIFIC_TRADE
        Exit Sub
    End If

    arrData = WsSrc.Range(WsSrc.Cells(1, 1), WsSrc.Cells(lastRow, lastCol)).Value

    For j = 1 To lastCol
        If LCase(Trim(arrData(1, j))) = LCase(Trim(FLD_WHO)) Then colWho = j
        If LCase(Trim(arrData(1, j))) = LCase(Trim(FLD_TOTAL_HOURS)) Then colTotal = j
        If LCase(Trim(arrData(1, j))) = LCase(Trim(FLD_TOTAL_MATCOST)) Then colMat = j
        If LCase(Trim(arrData(1, j))) = LCase(Trim(FLD_BUILDING)) Then colBld = j
        If LCase(Trim(arrData(1, j))) = LCase(Trim(FLD_ZONE)) Then colZone = j
        If LCase(Trim(arrData(1, j))) = LCase("Trade 1 PM Hours/Reactive hours included") Then colMulti1 = j
        If LCase(Trim(arrData(1, j))) = LCase("Trade 2 PM Hours/Reactive hours included") Then colMulti2 = j
    Next j

    If colWho = 0 Or colTotal = 0 Or colMat = 0 Or colBld = 0 Or colZone = 0 Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "PIVOT ERROR: Missing Headers!", vbCritical
        End
    End If

    ReDim arrOut(1 To lastRow * 3, 1 To lastCol + 1)

    For j = 1 To lastCol: arrOut(1, j) = arrData(1, j): Next j
    arrOut(1, lastCol + 1) = FLD_SPECIFIC_TRADE
    outRow = 2

    Dim tradeRaw As String, tradesArr() As String
    Dim matCost As Double, hrs1 As Double, hrs2 As Double, tot As Double
    Dim bldRaw As String

    For i = 2 To lastRow
        ' --- CRITICAL FIX: Aggressive string cleaning for hidden spaces and zeroes ---
        bldRaw = Application.WorksheetFunction.clean(Trim$(SafeStr(arrData(i, colBld))))
        bldRaw = Replace(bldRaw, Chr(160), "") ' Removes HTML spaces

        ' Filter out blanks AND literal zeroes returning from empty formulas
        If Len(bldRaw) > 0 And LCase$(bldRaw) <> "(blank)" And bldRaw <> "0" Then

            tradeRaw = SafeStr(arrData(i, colWho))
            matCost = NzDbl(arrData(i, colMat))
            tot = NzDbl(arrData(i, colTotal))

            If InStr(tradeRaw, "/") > 0 Then
                tradesArr = Split(tradeRaw, "/")

                If colMulti1 > 0 Then hrs1 = NzDbl(arrData(i, colMulti1)) Else hrs1 = 0
                If colMulti2 > 0 Then hrs2 = NzDbl(arrData(i, colMulti2)) Else hrs2 = 0

                ' ROW A: First Trade
                For j = 1 To lastCol: arrOut(outRow, j) = arrData(i, j): Next j
                arrOut(outRow, colWho) = tradeRaw
                arrOut(outRow, lastCol + 1) = Trim(tradesArr(0))
                arrOut(outRow, colTotal) = hrs1
                arrOut(outRow, colMat) = 0
                outRow = outRow + 1

                ' ROW B: Second Trade
                For j = 1 To lastCol: arrOut(outRow, j) = arrData(i, j): Next j
                arrOut(outRow, colWho) = tradeRaw
                arrOut(outRow, lastCol + 1) = Trim(tradesArr(UBound(tradesArr)))
                arrOut(outRow, colTotal) = hrs2
                arrOut(outRow, colMat) = 0
                outRow = outRow + 1

                ' ROW C: Shared Materials
                If matCost <> 0 Then
                    For j = 1 To lastCol: arrOut(outRow, j) = arrData(i, j): Next j
                    arrOut(outRow, colWho) = tradeRaw
                    arrOut(outRow, lastCol + 1) = "Shared Materials"
                    arrOut(outRow, colTotal) = 0
                    arrOut(outRow, colMat) = matCost
                    outRow = outRow + 1
                End If
            Else
                For j = 1 To lastCol: arrOut(outRow, j) = arrData(i, j): Next j
                arrOut(outRow, lastCol + 1) = tradeRaw
                outRow = outRow + 1
            End If
        End If ' End of Building Check
    Next i

    ' Write output (Ensures we don't crash if the sheet is 100% empty)
    If outRow > 2 Then
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(outRow - 1, lastCol + 1)).Value = arrOut
    Else
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, lastCol + 1)).Value = arrOut
    End If
End Sub
Private Function CheckHeaders(ByVal src As Range, ByVal names As Variant) As String
    Dim f As Variant, found As Range
    For Each f In names
        Set found = src.Rows(1).Find(What:=CStr(f), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If found Is Nothing Then
            CheckHeaders = "Missing header: " & CStr(f)
            Exit Function
        End If
    Next f
    CheckHeaders = ""
End Function

Private Function RecreatePivotSheet_safe(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim currentSheet As Worksheet

    ' Remember the currently active sheet to prevent jumping when Add activates the new sheet
    Set currentSheet = ActiveSheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo 0

    With ws
        .Range("H1").Value = "Building x Who (Vertical) - Total Hours & Material Cost"
        .Range("H1").Font.Bold = True
        .Range("H1").Font.Size = 14
        .Range("H2").Value = "Refreshed on: " & Format(Now, "yyyy-mm-dd hh:nn")
        .Range("H2").Font.Italic = True
    End With

    ' Quietly return the user to where they were
    currentSheet.Activate

    Set RecreatePivotSheet_safe = ws
End Function

Private Function BuildVerticalPivot_safe( _
    ByVal srcRange As Range, _
    ByVal wsTarget As Worksheet, _
    ByVal ptName As String, _
    ByVal destCell As Range) As PivotTable

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim uniquePtName As String
    Dim destDataStr As String
    Dim srcDataStr As String
    Dim c As Long

    ' 1. CRITICAL: Prevent 1004 by ensuring no blank headers exist in the source data!
    For c = 1 To srcRange.Columns.count
        If Trim(srcRange.Cells(1, c).Value) = "" Then
            srcRange.Cells(1, c).Value = "Column_" & c
        End If
    Next c

    ' 2. Generate a guaranteed unique Pivot Table name
    uniquePtName = ptName & "_" & Format(Now, "hhmmss")

    ' 3. CRITICAL FIX: Create explicitly unique source string to bypass Excel Cache Recycling
    srcDataStr = "'" & srcRange.Worksheet.Name & "'!" & srcRange.Address(ReferenceStyle:=xlR1C1)

    ' 4. Format destination explicitly
    destDataStr = "'" & wsTarget.Name & "'!" & destCell.Address(ReferenceStyle:=xlR1C1)

    ' 5. Create Cache and Table directly from the raw string
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcDataStr)
    pc.MissingItemsLimit = xlMissingItemsNone ' Force drop of old ghost data

    Set pt = pc.CreatePivotTable(TableDestination:=destDataStr, TableName:=uniquePtName)

    pt.ManualUpdate = True
    With pt
        .HasAutoFormat = False
        .NullString = "-"

        With .PivotFields(FLD_BUILDING)
            .Orientation = xlRowField
            .Position = 1
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With

        With .PivotFields(FLD_WHO)
            .Orientation = xlRowField
            .Position = 2
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With

        With .PivotFields(FLD_SPECIFIC_TRADE)
            .Orientation = xlRowField
            .Position = 3
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With

        With .PivotFields(FLD_TOTAL_HOURS)
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Caption = CAP_HOURS
        End With

        With .PivotFields(FLD_TOTAL_MATCOST)
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "$#,##0.00"
            .Caption = CAP_COST
        End With

        .DataPivotField.Orientation = xlRowField
        .DataPivotField.Position = 3
        .ColumnGrand = False
        .RowGrand = True
    End With

    pt.ManualUpdate = False
    pt.PivotCache.Refresh ' Double tap to guarantee clean data

    On Error Resume Next
    pt.PivotFields(FLD_BUILDING).PivotItems("(blank)").Visible = False
    pt.PivotFields(FLD_WHO).PivotItems("(blank)").Visible = False
    On Error GoTo 0

    pt.TableRange2.Columns.AutoFit
    Set BuildVerticalPivot_safe = pt
End Function

Private Sub WriteTotalsPanel_FromSource(ByVal pt As PivotTable, ByVal WsOut As Worksheet, ByVal StartCell As Range, ByVal WsSrc As Worksheet, ByVal SrcData As Range, ByVal FldZone As String, ByVal FldBuilding As String, ByVal FldHours As String, ByVal FldCost As String)
    Dim r As Long, c1 As Long, c2 As Long, c3 As Long
    Dim zoneFilter As String
    Dim dict As Object, tradeDict As Object
    Dim i As Long
    Dim colZone As Long, colBld As Long, colH As Long, colC As Long, colWho As Long
    Dim colM1 As Long, colM2 As Long
    Dim key As String, v As Variant
    Dim grandH As Double, grandC As Double

    On Error Resume Next
    zoneFilter = pt.PivotFields(FldZone).CurrentPage
    If Err.Number <> 0 Then zoneFilter = ""
    On Error GoTo 0

    If LCase$(zoneFilter) = "(all)" Then zoneFilter = ""

    colZone = FindHeaderCol(SrcData, FldZone)
    colBld = FindHeaderCol(SrcData, FldBuilding)
    colH = FindHeaderCol(SrcData, FldHours)
    colC = FindHeaderCol(SrcData, FldCost)
    colWho = FindHeaderCol(SrcData, FLD_WHO)
    colM1 = FindHeaderCol(SrcData, "Trade 1 PM Hours/Reactive hours included")
    colM2 = FindHeaderCol(SrcData, "Trade 2 PM Hours/Reactive hours included")

    If colBld = 0 Or colH = 0 Or colC = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Set tradeDict = CreateObject("Scripting.Dictionary")
    tradeDict.CompareMode = 1

    For i = 2 To SrcData.Rows.count
        Dim z As String, b As String
        Dim h As Double, c As Double

        If colZone > 0 Then z = SafeStr(SrcData.Cells(i, colZone).Value) Else z = ""

        If (zoneFilter = "") Or (StrComp(z, zoneFilter, vbTextCompare) = 0) Then
            b = Trim$(SafeStr(SrcData.Cells(i, colBld).Value))
            h = NzDbl(SrcData.Cells(i, colH).Value)
            c = NzDbl(SrcData.Cells(i, colC).Value)

            If Len(b) > 0 And LCase$(b) <> "(blank)" Then

                ' --- TALLY BUILDING TOTALS ---
                If Not dict.Exists(b) Then dict.Add b, Array(0#, 0#)
                v = dict(b)
                v(0) = v(0) + h
                v(1) = v(1) + c
                dict(b) = v

                grandH = grandH + h
                grandC = grandC + c

                ' --- TALLY INDIVIDUAL TRADE HOURS ---
                If colWho > 0 Then
                    Dim tradeRaw As String
                    tradeRaw = Trim$(SafeStr(SrcData.Cells(i, colWho).Value))

                    If Len(tradeRaw) = 0 Or LCase$(tradeRaw) = "(blank)" Then
                        tradeRaw = "Unassigned (Blank Trade)"
                    End If

                    If InStr(tradeRaw, "/") > 0 Then
                        Dim tArr() As String
                        Dim t1 As String, t2 As String
                        Dim h1 As Double, h2 As Double

                        tArr = Split(tradeRaw, "/")
                        t1 = Trim$(tArr(0))
                        t2 = Trim$(tArr(UBound(tArr)))

                        ' Pull the exact numbers from Columns Q and S (which already contain all the math)
                        If colM1 > 0 Then h1 = NzDbl(SrcData.Cells(i, colM1).Value) Else h1 = 0
                        If colM2 > 0 Then h2 = NzDbl(SrcData.Cells(i, colM2).Value) Else h2 = 0

                        If Not tradeDict.Exists(t1) Then tradeDict.Add t1, 0#
                        If Not tradeDict.Exists(t2) Then tradeDict.Add t2, 0#

                        tradeDict(t1) = tradeDict(t1) + h1
                        tradeDict(t2) = tradeDict(t2) + h2
                    Else
                        If Not tradeDict.Exists(tradeRaw) Then tradeDict.Add tradeRaw, 0#
                        tradeDict(tradeRaw) = tradeDict(tradeRaw) + h
                    End If
                End If
            End If
        End If
    Next i

    c1 = StartCell.Column
    c2 = c1 + 1
    c3 = c1 + 2
    r = StartCell.Row

    With WsOut
        .Cells(r, c1).Value = "Totals (by Building)"
        .Cells(r, c1).Font.Bold = True
        r = r + 1

        .Cells(r, c1).Value = "Building"
        .Cells(r, c2).Value = "Total Hours"
        .Cells(r, c3).Value = "Total PM Material Cost"
        .Range(.Cells(r, c1), .Cells(r, c3)).Font.Bold = True
        .Range(.Cells(r, c1), .Cells(r, c3)).Interior.Color = RGB(242, 242, 242)
        r = r + 1
    End With

    Dim k As Variant, keys As Variant
    keys = dict.keys
    If dict.count > 1 Then QuickSortText keys, LBound(keys), UBound(keys)

    For Each k In keys
        v = dict(k)
        With WsOut
            .Cells(r, c1).Value = CStr(k)
            .Cells(r, c2).Value = v(0): .Cells(r, c2).NumberFormat = "#,##0.00"
            .Cells(r, c3).Value = v(1): .Cells(r, c3).NumberFormat = "$#,##0.00"
            .Range(.Cells(r, c1), .Cells(r, c3)).Font.Bold = True
            .Range(.Cells(r, c1), .Cells(r, c3)).Interior.Color = RGB(242, 242, 242)
        End With
        r = r + 1
    Next k

    r = r + 1
    With WsOut
        .Cells(r, c1).Value = "GRAND TOTAL (All Visible)"
        .Cells(r, c2).Value = grandH: .Cells(r, c2).NumberFormat = "#,##0.00"
        .Cells(r, c3).Value = grandC: .Cells(r, c3).NumberFormat = "$#,##0.00"
        .Range(.Cells(r, c1), .Cells(r, c3)).Font.Bold = True
        .Range(.Cells(r, c1), .Cells(r, c3)).Interior.Color = RGB(217, 217, 217)
        With .Range(.Cells(r, c1), .Cells(r, c3)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .weight = xlThick
            .Color = RGB(150, 150, 150)
        End With
    End With

    r = r + 2

    If tradeDict.count > 0 Then
        With WsOut
            .Cells(r, c1).Value = "Individual Trade Breakdown"
            .Cells(r, c1).Font.Bold = True
            r = r + 1

            .Cells(r, c1).Value = "Trade"
            .Cells(r, c2).Value = "Total Hours"
            .Cells(r, c3).Value = "Estimated FTEs"
            .Range(.Cells(r, c1), .Cells(r, c3)).Font.Bold = True
            .Range(.Cells(r, c1), .Cells(r, c3)).Interior.Color = RGB(242, 242, 242)
            r = r + 1
        End With

        Dim tKeys As Variant, tKey As Variant
        tKeys = tradeDict.keys
        If tradeDict.count > 1 Then QuickSortText tKeys, LBound(tKeys), UBound(tKeys)

        For Each tKey In tKeys
            With WsOut
                .Cells(r, c1).Value = CStr(tKey)
                .Cells(r, c2).Value = tradeDict(tKey)
                .Cells(r, c2).NumberFormat = "#,##0.00"
                .Cells(r, c3).Value = tradeDict(tKey) / 2080 ' FTE Calculation
                .Cells(r, c3).NumberFormat = "#,##0.00"
                .Range(.Cells(r, c1), .Cells(r, c3)).Interior.Color = RGB(242, 242, 242)
            End With
            r = r + 1
        Next tKey

        WsOut.Range(WsOut.Cells(StartCell.Row, c1), WsOut.Cells(r, c3)).Columns.AutoFit
    End If
End Sub

Private Function FindHeaderCol(ByVal SrcData As Range, ByVal HeaderText As String) As Long
    Dim f As Range
    Set f = SrcData.Rows(1).Find(What:=HeaderText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If f Is Nothing Then
        FindHeaderCol = 0
    Else
        FindHeaderCol = f.Column - SrcData.Column + 1
    End If
End Function

Private Function NzDbl(ByVal v As Variant) As Double
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
    On Error GoTo 0
End Function

Private Sub QuickSortText(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, tmp As Variant

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While CStr(arr(i)) < CStr(pivot): i = i + 1: Loop
        Do While CStr(arr(j)) > CStr(pivot): j = j - 1: Loop

        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If first < j Then QuickSortText arr, first, j
    If i < last Then QuickSortText arr, i, last
End Sub

Private Function CheckColumnData(ByVal src As Range, ByVal headerName As String) As String
    Dim f As Range
    Dim colIndex As Long
    Dim dataRange As Range

    Set f = src.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If f Is Nothing Then
        CheckColumnData = "Missing header: " & headerName
        Exit Function
    End If

    colIndex = f.Column - src.Column + 1

    If src.Rows.count < 2 Then
        CheckColumnData = ""
        Exit Function
    End If

    Set dataRange = src.Columns(colIndex).offset(1, 0).Resize(src.Rows.count - 1, 1)

    If Application.WorksheetFunction.CountA(dataRange) = 0 Then
        CheckColumnData = "The '" & headerName & "' column is empty."
    Else
        CheckColumnData = ""
    End If
End Function

Private Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
    On Error GoTo 0
End Function

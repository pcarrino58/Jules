' ==============================================================================
' 1. INITIAL COMPARISON: Cross-references PMs against Job Plans
' ==============================================================================
Sub CompareJobPlans()
    Dim wb As Workbook
    Dim wsPM As Worksheet, wsJP As Worksheet, wsReport As Worksheet
    Dim colPM As Long, colJP As Long
    Dim lastRowPM As Long, lastRowJP As Long
    Dim i As Long
    Dim dictPM As Object, dictJP As Object
    Dim key As Variant
    Dim reportRow As Long
    Dim cell As Range

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set wsPM = wb.Sheets("List of PMs")
    Set wsJP = wb.Sheets("List of Job Plans")
    On Error GoTo 0

    If wsPM Is Nothing Or wsJP Is Nothing Then
        MsgBox "Required sheets ('List of PMs' and/or 'List of Job Plans') are missing in this workbook.", vbCritical
        Exit Sub
    End If

    If wsPM.AutoFilterMode Then wsPM.AutoFilterMode = False
    If wsJP.AutoFilterMode Then wsJP.AutoFilterMode = False

    colPM = 0: colJP = 0

    For Each cell In wsPM.Rows(1).Cells
        If Trim(cell.Value) = "Next JOB Plan" And colPM = 0 Then colPM = cell.Column
    Next cell
    For Each cell In wsJP.Rows(1).Cells
        If Trim(cell.Value) = "Job Plan" And colJP = 0 Then colJP = cell.Column
    Next cell

    If colPM = 0 Or colJP = 0 Then Exit Sub

    Set dictPM = CreateObject("Scripting.Dictionary")
    dictPM.CompareMode = vbTextCompare
    Set dictJP = CreateObject("Scripting.Dictionary")
    dictJP.CompareMode = vbTextCompare

    lastRowPM = wsPM.Cells(wsPM.Rows.Count, colPM).End(xlUp).Row
    lastRowJP = wsJP.Cells(wsJP.Rows.Count, colJP).End(xlUp).Row

    Dim arrPM As Variant
    If lastRowPM = 2 Then
        ReDim arrPM(1 To 1, 1 To 1)
        arrPM(1, 1) = wsPM.Cells(2, colPM).Value
    ElseIf lastRowPM > 2 Then
        arrPM = wsPM.Range(wsPM.Cells(2, colPM), wsPM.Cells(lastRowPM, colPM)).Value
    End If

    If IsArray(arrPM) Then
        For i = 1 To UBound(arrPM, 1)
            Dim valPM As String
            valPM = Trim(arrPM(i, 1))
            If valPM <> "" And Not dictPM.exists(valPM) Then dictPM.Add valPM, "List of PMs"
        Next i
    End If

    Dim arrJP As Variant
    If lastRowJP = 2 Then
        ReDim arrJP(1 To 1, 1 To 1)
        arrJP(1, 1) = wsJP.Cells(2, colJP).Value
    ElseIf lastRowJP > 2 Then
        arrJP = wsJP.Range(wsJP.Cells(2, colJP), wsJP.Cells(lastRowJP, colJP)).Value
    End If

    If IsArray(arrJP) Then
        For i = 1 To UBound(arrJP, 1)
            Dim valJP As String
            valJP = Trim(arrJP(i, 1))
            If valJP <> "" And Not dictJP.exists(valJP) Then dictJP.Add valJP, "List of Job Plans"
        Next i
    End If

    On Error Resume Next
    Set wsReport = wb.Sheets("Job Plan Comparison")
    On Error GoTo 0
    If wsReport Is Nothing Then
        Set wsReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsReport.Name = "Job Plan Comparison"
    Else
        If wsReport.AutoFilterMode Then wsReport.AutoFilterMode = False
        wsReport.Cells.Clear
    End If

    Dim totalKeys As Long
    totalKeys = dictPM.Count + dictJP.Count
    If totalKeys = 0 Then Exit Sub

    Dim arrReport() As Variant
    ReDim arrReport(1 To totalKeys + 1, 1 To 3)

    arrReport(1, 1) = "Job Plan Identifier"
    arrReport(1, 2) = "Match Status"
    arrReport(1, 3) = "Location"

    reportRow = 2
    For Each key In dictPM.Keys
        arrReport(reportRow, 1) = key
        If dictJP.exists(key) Then
            arrReport(reportRow, 2) = "Match"
            arrReport(reportRow, 3) = "Found in Both"
            dictJP.Remove key
        Else
            arrReport(reportRow, 2) = "Mismatch (Missing)"
            arrReport(reportRow, 3) = "'List of PMs' Only"
        End If
        reportRow = reportRow + 1
    Next key
    For Each key In dictJP.Keys
        arrReport(reportRow, 1) = key
        arrReport(reportRow, 2) = "Mismatch (Unused)"
        arrReport(reportRow, 3) = "'List of Job Plans' Only"
        reportRow = reportRow + 1
    Next key

    wsReport.Range("A1").Resize(reportRow - 1, 3).Value = arrReport

    With wsReport.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    Dim rngStatus As Range
    Set rngStatus = wsReport.Range("B2:B" & reportRow - 1)
    rngStatus.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Mismatch (Missing)"""
    rngStatus.FormatConditions(rngStatus.FormatConditions.Count).Font.Color = vbRed
    rngStatus.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Mismatch (Unused)"""
    rngStatus.FormatConditions(rngStatus.FormatConditions.Count).Font.Color = RGB(255, 165, 0)

    wsReport.Columns("A:C").AutoFit
    wsReport.Range("A1:C" & reportRow - 1).AutoFilter
End Sub

' ==============================================================================
' 2. FREQUENCY EXTRACTION
' ==============================================================================
Sub IdentifyJobPlanFrequencies()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim colID As Long, colFreq As Long
    Dim cell As Range, jpID As String, freq As String
    Static regEx As Object
    Dim matches As Object
    Dim freqNum As String, freqCode As String

    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Job Plan Comparison")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    colID = 0
    For Each cell In ws.Rows(1).Cells
        If Trim(cell.Value) = "Job Plan Identifier" And colID = 0 Then colID = cell.Column
    Next cell
    If colID = 0 Then Exit Sub

    colFreq = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(1, colFreq).Value <> "Frequency" Then
        colFreq = colFreq + 1
        With ws.Cells(1, colFreq)
            .Value = "Frequency"
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
        End With
    End If

    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = False: regEx.IgnoreCase = True: regEx.Pattern = "(\d*)([DWMY])\s*$"
    End If

    lastRow = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim arrData As Variant
    If lastRow = 2 Then
        ReDim arrData(1 To 1, 1 To 1)
        arrData(1, 1) = ws.Cells(2, colID).Value
    Else
        arrData = ws.Range(ws.Cells(2, colID), ws.Cells(lastRow, colID)).Value
    End If

    Dim arrFreq() As String
    ReDim arrFreq(1 To UBound(arrData, 1), 1 To 1)

    For i = 1 To UBound(arrData, 1)
        jpID = Trim(arrData(i, 1))
        jpID = Replace(jpID, Chr(160), "")
        freq = "Unknown / Custom"

        If regEx.Test(jpID) Then
            Set matches = regEx.Execute(jpID)
            freqNum = matches(0).SubMatches(0)
            freqCode = UCase(matches(0).SubMatches(1))

            Select Case freqCode
                Case "D"
                    If freqNum = "" Or freqNum = "1" Then freq = "Daily" Else freq = "Every " & freqNum & " Days"
                Case "W"
                    If freqNum = "" Or freqNum = "1" Then freq = "Weekly" Else freq = "Every " & freqNum & " Weeks"
                Case "M"
                    If freqNum = "" Or freqNum = "1" Then
                        freq = "Monthly"
                    ElseIf freqNum = "3" Then
                        freq = "Quarterly"
                    ElseIf freqNum = "6" Then
                        freq = "Semi-Annual"
                    Else
                        freq = "Every " & freqNum & " Months"
                    End If
                Case "Y"
                    If freqNum = "" Or freqNum = "1" Then freq = "Annual" Else freq = "Every " & freqNum & " Years"
            End Select
        End If
        arrFreq(i, 1) = freq
    Next i

    ws.Range(ws.Cells(2, colFreq), ws.Cells(lastRow, colFreq)).Value = arrFreq
    ws.Columns(colFreq).AutoFit
End Sub

' ==============================================================================
' 3. MAP DESCRIPTIONS
' ==============================================================================
Sub AddJobPlanDescriptions()
    Dim wb As Workbook
    Dim wsComp As Worksheet, wsJP As Worksheet
    Dim colCompID As Long, colCompDesc As Long
    Dim colJPID As Long, colJPDesc As Long
    Dim lastRowJP As Long, lastRowComp As Long
    Dim i As Long, dictDesc As Object, cell As Range, jpID As String

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsComp = wb.Sheets("Job Plan Comparison")
    Set wsJP = wb.Sheets("List of Job Plans")
    On Error GoTo 0
    If wsComp Is Nothing Or wsJP Is Nothing Then Exit Sub

    colJPID = 0: colJPDesc = 0
    For Each cell In wsJP.Rows(1).Cells
        Dim headTxt As String
        headTxt = Trim(UCase(cell.Value))
        If (headTxt = "JOB PLAN" Or headTxt = "JOB PLAN IDENTIFIER") And colJPID = 0 Then colJPID = cell.Column
        If (headTxt = "DESCRIPTION" Or headTxt = "JOB PLAN DESCRIPTION") And colJPDesc = 0 Then colJPDesc = cell.Column
    Next cell
    If colJPID = 0 Or colJPDesc = 0 Then Exit Sub

    Set dictDesc = CreateObject("Scripting.Dictionary")
    dictDesc.CompareMode = vbTextCompare

    lastRowJP = wsJP.Cells(wsJP.Rows.Count, colJPID).End(xlUp).Row
    If lastRowJP > 1 Then
        Dim arrJP As Variant
        arrJP = wsJP.Range(wsJP.Cells(1, 1), wsJP.Cells(lastRowJP, Application.Max(colJPID, colJPDesc))).Value
        For i = 2 To UBound(arrJP, 1)
            jpID = Trim(arrJP(i, colJPID))
            If jpID <> "" And Not dictDesc.exists(jpID) Then dictDesc.Add jpID, Trim(arrJP(i, colJPDesc))
        Next i
    End If

    colCompID = 0
    For Each cell In wsComp.Rows(1).Cells
        If Trim(cell.Value) = "Job Plan Identifier" And colCompID = 0 Then colCompID = cell.Column
    Next cell
    If colCompID = 0 Then Exit Sub

    colCompDesc = wsComp.Cells(1, wsComp.Columns.Count).End(xlToLeft).Column
    If wsComp.Cells(1, colCompDesc).Value <> "Job Plan Description" Then
        colCompDesc = colCompDesc + 1
        With wsComp.Cells(1, colCompDesc)
            .Value = "Job Plan Description"
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
        End With
    End If

    lastRowComp = wsComp.Cells(wsComp.Rows.Count, colCompID).End(xlUp).Row
    If lastRowComp < 2 Then Exit Sub

    Dim arrCompID As Variant
    If lastRowComp = 2 Then
        ReDim arrCompID(1 To 1, 1 To 1)
        arrCompID(1, 1) = wsComp.Cells(2, colCompID).Value
    Else
        arrCompID = wsComp.Range(wsComp.Cells(2, colCompID), wsComp.Cells(lastRowComp, colCompID)).Value
    End If

    Dim arrCompDesc() As Variant
    ReDim arrCompDesc(1 To UBound(arrCompID, 1), 1 To 1)

    For i = 1 To UBound(arrCompID, 1)
        jpID = Trim(arrCompID(i, 1))
        If dictDesc.exists(jpID) Then arrCompDesc(i, 1) = dictDesc(jpID) Else arrCompDesc(i, 1) = "No Description Found"
    Next i

    wsComp.Range(wsComp.Cells(2, colCompDesc), wsComp.Cells(lastRowComp, colCompDesc)).Value = arrCompDesc
    wsComp.Columns(colCompDesc).AutoFit
End Sub

' ==============================================================================
' 4. SCRUB DESCRIPTIONS
' ==============================================================================
Sub ScrubDescriptionFrequencies()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, colDesc As Long
    Dim cell As Range, descText As String
    Static regEx As Object
    Static regExPunct As Object

    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Job Plan Comparison")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    colDesc = 0
    For Each cell In ws.Rows(1).Cells
        If Trim(cell.Value) = "Job Plan Description" And colDesc = 0 Then colDesc = cell.Column
    Next cell
    If colDesc = 0 Then Exit Sub

    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True: regEx.IgnoreCase = True
        regEx.Pattern = "\b(Daily|Weekly|Bi-Weekly|Monthly|Quarterly|Semi-Annual|Semi Annual|Annually|Annual|Yearly|Bi-Annual|Every \d+ (Days?|Weeks?|Months?|Years?)|\d+ (Day|Week|Month|Year)s?)\b"
    End If

    If regExPunct Is Nothing Then
        Set regExPunct = CreateObject("VBScript.RegExp")
        regExPunct.Global = True
        regExPunct.Pattern = "^[\s\-_,]+|[\s\-_,]+$"
    End If

    lastRow = ws.Cells(ws.Rows.Count, colDesc).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim arrDesc As Variant
    If lastRow = 2 Then
        ReDim arrDesc(1 To 1, 1 To 1)
        arrDesc(1, 1) = ws.Cells(2, colDesc).Value
    Else
        arrDesc = ws.Range(ws.Cells(2, colDesc), ws.Cells(lastRow, colDesc)).Value
    End If

    For i = 1 To UBound(arrDesc, 1)
        descText = Trim(arrDesc(i, 1))
        If Len(descText) > 0 And descText <> "No Description Found" Then
            descText = regEx.Replace(descText, "")
            descText = Replace(descText, "()", "")
            descText = Replace(descText, "( )", "")
            Do While InStr(descText, "  ") > 0
                descText = Replace(descText, "  ", " ")
            Loop

            descText = regExPunct.Replace(descText, "")
            arrDesc(i, 1) = Trim(descText)
        End If
    Next i

    ws.Range(ws.Cells(2, colDesc), ws.Cells(lastRow, colDesc)).Value = arrDesc
    ws.Columns(colDesc).AutoFit
End Sub

' ==============================================================================
' 5. HELPER FUNCTION: ASSET ID VALIDATOR (YOUR SOLUTION)
' ==============================================================================
Function IsValidAssetID(ByVal aID As String) As Boolean
    aID = UCase(Trim(aID))

    ' Block blank or impossibly tiny IDs
    If Len(aID) < 3 Then
        IsValidAssetID = False
        Exit Function
    End If

    ' Block known generic placeholders to prevent False Positive Exact Matches
    Select Case aID
        Case "N/A", "TBD", "NONE", "UNKNOWN", "VARIOUS", "000", "---", "NULL", "VARIES"
            IsValidAssetID = False
        Case Else
            IsValidAssetID = True
    End Select
End Function

' ==============================================================================
' 6. MASTER GROUPING & 3-TIER MULTI-PASS MATCHING ENGINE
' ==============================================================================
Sub GroupAssetsByJobPlanDescription()
    Dim wb As Workbook
    Dim wsComp As Worksheet, wsAsset As Worksheet, wsPM As Worksheet, wsLog As Worksheet
    Dim colDesc As Long, lastRowComp As Long, logIndex As Long
    Dim cell As Range, i As Long, rA As Long
    Dim sheetCounter As Long, totalSheets As Long

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsComp = wb.Sheets("Job Plan Comparison")
    Set wsPM = wb.Sheets("List of PMs")
    On Error GoTo 0
    If wsComp Is Nothing Or wsPM Is Nothing Then Exit Sub

    colDesc = 0
    For Each cell In wsComp.Rows(1).Cells
        If Trim(cell.Value) = "Job Plan Description" And colDesc = 0 Then colDesc = cell.Column
    Next cell
    If colDesc = 0 Then Exit Sub

    Dim assetSheets As Collection
    Set assetSheets = New Collection
    For Each wsAsset In wb.Worksheets
        If InStr(1, wsAsset.Name, "Assets", vbTextCompare) > 0 And wsAsset.Name <> "Job Plan Comparison" And wsAsset.Name <> "Master Asset Data" Then assetSheets.Add wsAsset
    Next wsAsset
    If assetSheets.Count = 0 Then Exit Sub

    totalSheets = assetSheets.Count

    If wsComp.Columns.Count > colDesc Then wsComp.Range(wsComp.Cells(1, colDesc + 1), wsComp.Cells(wsComp.Rows.Count, wsComp.Columns.Count)).Clear

    On Error Resume Next
    Set wsLog = wb.Sheets("Audit Drilldown Data")
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsLog.Name = "Audit Drilldown Data"
        wsLog.Visible = xlSheetHidden
    Else
        wsLog.Cells.Clear
    End If
    wsLog.Range("A1:E1").Value = Array("Campus", "Job Plan Description", "Match Type", "Asset ID", "Asset Details")
    wsLog.Range("A1:E1").Font.Bold = True

    Dim logArr() As Variant
    ReDim logArr(1 To 500000, 1 To 5)
    logIndex = 1

    Dim startCol As Long, c As Long
    startCol = colDesc + 2: c = startCol

    Dim dictTotalPlans As Object, dictUniqueCampusAssets As Object, dictCampusInventory As Object
    Dim dictTotalExact As Object, dictTotalTier2 As Object, dictTotalTier3 As Object
    Dim dictCellExact As Object, dictCellT2 As Object, dictCellT3 As Object
    Dim dictNativeIDString As Object, dictNativeDesc As Object
    Dim dictExactJP As Object, dictT2JP As Object, dictT3JP As Object, dictJPDesc As Object
    Dim dictSheetAssetToRow As Object

    Set dictTotalPlans = CreateObject("Scripting.Dictionary")
    Set dictUniqueCampusAssets = CreateObject("Scripting.Dictionary")
    Set dictCampusInventory = CreateObject("Scripting.Dictionary")
    Set dictTotalExact = CreateObject("Scripting.Dictionary")
    Set dictTotalTier2 = CreateObject("Scripting.Dictionary")
    Set dictTotalTier3 = CreateObject("Scripting.Dictionary")
    Set dictCellExact = CreateObject("Scripting.Dictionary")
    Set dictCellT2 = CreateObject("Scripting.Dictionary")
    Set dictCellT3 = CreateObject("Scripting.Dictionary")
    Set dictNativeIDString = CreateObject("Scripting.Dictionary")
    Set dictNativeDesc = CreateObject("Scripting.Dictionary")
    Set dictExactJP = CreateObject("Scripting.Dictionary")
    Set dictT2JP = CreateObject("Scripting.Dictionary")
    Set dictT3JP = CreateObject("Scripting.Dictionary")
    Set dictJPDesc = CreateObject("Scripting.Dictionary")
    Set dictSheetAssetToRow = CreateObject("Scripting.Dictionary")

    sheetCounter = 1
    For Each wsAsset In assetSheets
        Application.StatusBar = "Pre-loading Data: " & wsAsset.Name & " (" & sheetCounter & " of " & totalSheets & ")..."
        DoEvents

        If wsAsset.AutoFilterMode Then wsAsset.AutoFilterMode = False

        wsComp.Cells(1, c).Value = wsAsset.Name
        wsComp.Cells(1, c).Font.Bold = True
        wsComp.Cells(1, c).Interior.Color = RGB(226, 239, 218)

        dictTotalPlans.Add wsAsset.Name, 0
        dictTotalExact.Add wsAsset.Name, 0
        dictTotalTier2.Add wsAsset.Name, 0
        dictTotalTier3.Add wsAsset.Name, 0

        dictUniqueCampusAssets.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictCampusInventory.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictSheetAssetToRow.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictSheetAssetToRow(wsAsset.Name).CompareMode = vbTextCompare

        dictExactJP.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictT2JP.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictT3JP.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictJPDesc.Add wsAsset.Name, CreateObject("Scripting.Dictionary")

        Dim colCampAsset As Long, colCampType As Long, colCampSubType As Long
        colCampAsset = 0: colCampType = 0: colCampSubType = 0

        For Each cell In wsAsset.Rows(1).Cells
            Dim headTxt As String
            headTxt = Trim(UCase(cell.Value))
            If (headTxt = "ASSET ID" Or headTxt = "ASSET #" Or headTxt = "ASSET" Or headTxt = "EQUIPMENT ID") And colCampAsset = 0 Then
                colCampAsset = cell.Column
            ElseIf headTxt = "ASSET TYPE" And colCampType = 0 Then
                colCampType = cell.Column
            ElseIf (headTxt = "ASSET SUBTYPE" Or headTxt = "SUB TYPE" Or headTxt = "SUBTYPE" Or headTxt = "ASSET SUB TYPE") And colCampSubType = 0 Then
                colCampSubType = cell.Column
            End If
        Next cell

        Dim lrA As Long, rngLast As Range
        Set rngLast = wsAsset.Cells.Find(What:="*", After:=wsAsset.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not rngLast Is Nothing Then lrA = rngLast.Row Else lrA = 1

        If lrA > 1 Then
            Dim maxCol As Long
            maxCol = wsAsset.Cells(1, wsAsset.Columns.Count).End(xlToLeft).Column
            If colCampAsset > maxCol Then maxCol = colCampAsset
            If colCampType > maxCol Then maxCol = colCampType
            If colCampSubType > maxCol Then maxCol = colCampSubType

            Dim arrInv As Variant
            arrInv = wsAsset.Range(wsAsset.Cells(1, 1), wsAsset.Cells(lrA, maxCol)).Value

            For rA = 2 To UBound(arrInv, 1)
                Dim aID As String, aType As String, aSub As String
                aID = "": aType = "": aSub = ""

                If colCampAsset > 0 Then aID = Trim(arrInv(rA, colCampAsset))
                If colCampType > 0 Then aType = CleanFuzzyString(CStr(arrInv(rA, colCampType)))
                If colCampSubType > 0 Then aSub = CleanFuzzyString(CStr(arrInv(rA, colCampSubType)))

                dictCampusInventory(wsAsset.Name).Add CStr(rA), aID & "||" & aType & "||" & aSub

                ' === NEW VALIDATION LOGIC INJECTED HERE ===
                If IsValidAssetID(aID) Then
                    If Not dictSheetAssetToRow(wsAsset.Name).exists(aID) Then
                        Set dictSheetAssetToRow(wsAsset.Name)(aID) = New Collection
                    End If
                    dictSheetAssetToRow(wsAsset.Name)(aID).Add CStr(rA)
                End If
            Next rA
        End If
        c = c + 1
        sheetCounter = sheetCounter + 1
    Next wsAsset

    Application.StatusBar = "Caching Dictionary Descriptions..."
    DoEvents

    Dim dictAllDesc As Object
    Set dictAllDesc = CreateObject("Scripting.Dictionary")
    dictAllDesc.CompareMode = vbTextCompare

    lastRowComp = wsComp.Cells(wsComp.Rows.Count, 1).End(xlUp).Row
    Dim arrCompMaster As Variant
    arrCompMaster = wsComp.Range(wsComp.Cells(1, 1), wsComp.Cells(lastRowComp, colDesc)).Value

    Dim jpID As String, jpDesc As String
    For i = 2 To UBound(arrCompMaster, 1)
        jpID = Trim(arrCompMaster(i, 1))
        jpDesc = CleanFuzzyString(CStr(arrCompMaster(i, colDesc)))
        If jpDesc <> "" And jpDesc <> "No Description Found" Then
            If Not dictAllDesc.exists(jpDesc) Then dictAllDesc.Add jpDesc, CreateObject("Scripting.Dictionary")
            If Not dictAllDesc(jpDesc).exists(jpID) Then dictAllDesc(jpDesc).Add jpID, 1
        End If
    Next i

    Dim colPMJP As Long, colPMAsset As Long
    colPMJP = 0: colPMAsset = 0
    For Each cell In wsPM.Rows(1).Cells
        Dim pHead As String
        pHead = Trim(UCase(cell.Value))
        If pHead = "NEXT JOB PLAN" And colPMJP = 0 Then colPMJP = cell.Column
        If (pHead = "ASSET" Or pHead = "ASSET ID" Or pHead = "ASSET #" Or pHead = "EQUIPMENT ID") And colPMAsset = 0 Then colPMAsset = cell.Column
    Next cell
    If colPMJP = 0 Or colPMAsset = 0 Then Exit Sub

    Dim dictPMAssets As Object
    Set dictPMAssets = CreateObject("Scripting.Dictionary")
    dictPMAssets.CompareMode = vbTextCompare

    Dim lastRowPM As Long, pID As String, pAsset As String
    lastRowPM = wsPM.Cells(wsPM.Rows.Count, colPMJP).End(xlUp).Row
    If lastRowPM > 1 Then
        Dim maxPMCol As Long
        maxPMCol = Application.Max(colPMJP, colPMAsset)
        Dim arrPMData As Variant
        arrPMData = wsPM.Range(wsPM.Cells(1, 1), wsPM.Cells(lastRowPM, maxPMCol)).Value

        For i = 2 To UBound(arrPMData, 1)
            pID = Trim(arrPMData(i, colPMJP))
            pAsset = Trim(arrPMData(i, colPMAsset))

            ' === NEW VALIDATION LOGIC INJECTED HERE ===
            If pID <> "" And pAsset <> "" Then
                If IsValidAssetID(pAsset) Then
                    If Not dictPMAssets.exists(pID) Then
                        dictPMAssets.Add pID, CreateObject("Scripting.Dictionary")
                        dictPMAssets(pID).CompareMode = vbTextCompare
                    End If
                    If Not dictPMAssets(pID).exists(pAsset) Then dictPMAssets(pID).Add pAsset, 1
                End If
            End If
        Next i
    End If

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False: regEx.IgnoreCase = True: regEx.Pattern = "(\d*)([DWMY])\s*$"

    Dim dictSheetPrefixes As Object
    Set dictSheetPrefixes = CreateObject("Scripting.Dictionary")
    For Each wsAsset In assetSheets
        dictSheetPrefixes.Add wsAsset.Name, Split(GetSheetPrefixes(wsAsset.Name), ",")
    Next wsAsset

    Dim dictScrubbedIDs As Object
    Set dictScrubbedIDs = CreateObject("Scripting.Dictionary")
    dictScrubbedIDs.CompareMode = vbTextCompare

    Dim vDesc As Variant, vID As Variant
    For Each vDesc In dictAllDesc.Keys
        For Each vID In dictAllDesc(vDesc).Keys
            If Not dictScrubbedIDs.exists(vID) Then
                Dim tmpID As String
                tmpID = Trim(Replace(vID, Chr(160), ""))
                If regEx.Test(tmpID) Then
                    tmpID = regEx.Replace(tmpID, "")
                    Do While Right(tmpID, 1) = "-" Or Right(tmpID, 1) = "_" Or Right(tmpID, 1) = " "
                        tmpID = Left(tmpID, Len(tmpID) - 1)
                    Loop
                End If
                dictScrubbedIDs.Add vID, tmpID
            End If
        Next vID
    Next vDesc

    ' ---------------------------------------------------------
    ' PASS 1: TIER 1 EXACT MATCHES & NATIVE CAMPUS DISCOVERY
    ' ---------------------------------------------------------
    Dim jpDescCount As Long, totalDescKeys As Long
    totalDescKeys = dictAllDesc.Count
    jpDescCount = 1

    Dim jpDescKey As Variant
    For Each jpDescKey In dictAllDesc.Keys
        If jpDescCount Mod 1000 = 0 Or jpDescCount = 1 Then
            Application.StatusBar = "Pass 1: Exact Matches (" & jpDescCount & " of " & totalDescKeys & " desc)..."
            DoEvents
        End If

        jpDesc = CStr(jpDescKey)

        For Each wsAsset In assetSheets
            Dim nativeKey As String
            nativeKey = wsAsset.Name & "|" & jpDesc

            Dim prefixArray() As String
            prefixArray = dictSheetPrefixes(wsAsset.Name)

            Dim matchedIDs As Object
            Set matchedIDs = CreateObject("Scripting.Dictionary")

            Dim kID As Variant
            For Each kID In dictAllDesc(jpDesc).Keys
                Dim isMatch As Boolean, p As Long
                isMatch = False
                For p = LBound(prefixArray) To UBound(prefixArray)
                    If UCase(Left(kID, Len(Trim(prefixArray(p))))) = Trim(prefixArray(p)) Then isMatch = True: Exit For
                Next p

                If isMatch Then
                    dictNativeDesc(nativeKey) = 1

                    Dim scrubbedID As String
                    scrubbedID = dictScrubbedIDs(kID)
                    matchedIDs(scrubbedID) = 1

                    If dictPMAssets.exists(kID) Then
                        Dim aKey As Variant
                        For Each aKey In dictPMAssets(kID).Keys
                            If dictSheetAssetToRow(wsAsset.Name).exists(aKey) Then

                                Dim matchedRow As Variant
                                For Each matchedRow In dictSheetAssetToRow(wsAsset.Name)(aKey)
                                    dictCellExact(nativeKey) = dictCellExact(nativeKey) + 1

                                    If Not dictUniqueCampusAssets(wsAsset.Name).exists(matchedRow) Then
                                        dictUniqueCampusAssets(wsAsset.Name).Add matchedRow, "Exact"
                                        dictTotalExact(wsAsset.Name) = dictTotalExact(wsAsset.Name) + 1

                                        logArr(logIndex, 1) = wsAsset.Name
                                        logArr(logIndex, 2) = jpDesc
                                        logArr(logIndex, 3) = "Exact"
                                        logArr(logIndex, 4) = aKey

                                        Dim exData() As String, exType As String, exSub As String
                                        exData = Split(dictCampusInventory(wsAsset.Name)(matchedRow), "||")
                                        exType = exData(1): exSub = exData(2)

                                        If exType <> "" And exSub <> "" Then
                                            logArr(logIndex, 5) = exType & " - " & exSub
                                        Else
                                            logArr(logIndex, 5) = exType & exSub
                                        End If
                                        logIndex = logIndex + 1
                                    End If

                                    If Not dictExactJP(wsAsset.Name).exists(matchedRow) Then dictExactJP(wsAsset.Name).Add matchedRow, kID Else If InStr(dictExactJP(wsAsset.Name)(matchedRow), kID) = 0 Then dictExactJP(wsAsset.Name)(matchedRow) = dictExactJP(wsAsset.Name)(matchedRow) & ", " & kID
                                    If Not dictJPDesc(wsAsset.Name).exists(matchedRow) Then dictJPDesc(wsAsset.Name).Add matchedRow, jpDesc Else If InStr(dictJPDesc(wsAsset.Name)(matchedRow), jpDesc) = 0 Then dictJPDesc(wsAsset.Name)(matchedRow) = dictJPDesc(wsAsset.Name)(matchedRow) & " | " & jpDesc
                                Next matchedRow

                            End If
                        Next aKey
                    End If
                End If
            Next kID

            Dim idString As String, numPlans As Long, mID As Variant
            idString = "": numPlans = 0
            For Each mID In matchedIDs.Keys
                If idString = "" Then idString = mID Else idString = idString & ", " & mID
                numPlans = numPlans + 1
            Next mID

            dictNativeIDString(nativeKey) = idString
            dictTotalPlans(wsAsset.Name) = dictTotalPlans(wsAsset.Name) + numPlans
        Next wsAsset
        jpDescCount = jpDescCount + 1
    Next jpDescKey

    Dim dictCampusNativeDesc As Object, dictCampusNonNativeDesc As Object
    Set dictCampusNativeDesc = CreateObject("Scripting.Dictionary")
    Set dictCampusNonNativeDesc = CreateObject("Scripting.Dictionary")

    For Each wsAsset In assetSheets
        dictCampusNativeDesc.Add wsAsset.Name, CreateObject("Scripting.Dictionary")
        dictCampusNonNativeDesc.Add wsAsset.Name, CreateObject("Scripting.Dictionary")

        Dim uDesc As Variant
        For Each uDesc In dictAllDesc.Keys
            If dictNativeDesc.exists(wsAsset.Name & "|" & uDesc) Then
                dictCampusNativeDesc(wsAsset.Name).Add uDesc, 1
            ElseIf uDesc <> "No Description Found" And uDesc <> "" Then
                dictCampusNonNativeDesc(wsAsset.Name).Add uDesc, 1
            End If
        Next uDesc
    Next wsAsset

    ' ---------------------------------------------------------
    ' PASS 2: TIER 2 FUZZY MATCHES (Native Campus Only)
    ' ---------------------------------------------------------
    sheetCounter = 1
    For Each wsAsset In assetSheets
        Application.StatusBar = "Pass 2: Fuzzy Native Matches (" & wsAsset.Name & " - " & sheetCounter & " of " & totalSheets & ")..."
        DoEvents

        Dim dictT2TypeCache As Object
        Set dictT2TypeCache = CreateObject("Scripting.Dictionary")
        dictT2TypeCache.CompareMode = vbTextCompare

        Dim rKey As Variant
        For Each rKey In dictCampusInventory(wsAsset.Name).Keys
            If Not dictUniqueCampusAssets(wsAsset.Name).exists(rKey) Then

                Dim invData2() As String
                invData2 = Split(dictCampusInventory(wsAsset.Name)(rKey), "||")
                Dim typePairKey As String, aTypeStr As String, aSubStr As String
                aTypeStr = invData2(1): aSubStr = invData2(2)
                typePairKey = aTypeStr & "|" & aSubStr

                Dim matchedT2Desc As String
                matchedT2Desc = ""

                If dictT2TypeCache.exists(typePairKey) Then
                    matchedT2Desc = dictT2TypeCache(typePairKey)
                Else
                    If Len(aTypeStr) > 0 Or Len(aSubStr) > 0 Then
                        matchedT2Desc = GetBestJobPlanMatch(aTypeStr, aSubStr, dictCampusNativeDesc(wsAsset.Name))
                    End If
                    dictT2TypeCache.Add typePairKey, matchedT2Desc
                End If

                If matchedT2Desc <> "" Then
                    Dim nativeKey2 As String
                    nativeKey2 = wsAsset.Name & "|" & matchedT2Desc

                    dictCellT2(nativeKey2) = dictCellT2(nativeKey2) + 1

                    dictUniqueCampusAssets(wsAsset.Name).Add rKey, "Tier 2 Fuzzy"
                    dictTotalTier2(wsAsset.Name) = dictTotalTier2(wsAsset.Name) + 1

                    logArr(logIndex, 1) = wsAsset.Name
                    logArr(logIndex, 2) = matchedT2Desc
                    logArr(logIndex, 3) = "Tier 2 Fuzzy"
                    logArr(logIndex, 4) = IIf(invData2(0) = "", "Row " & rKey, invData2(0))

                    If aTypeStr <> "" And aSubStr <> "" Then
                        logArr(logIndex, 5) = aTypeStr & " - " & aSubStr
                    Else
                        logArr(logIndex, 5) = aTypeStr & aSubStr
                    End If
                    logIndex = logIndex + 1

                    Dim fuzID2 As String
                    fuzID2 = dictNativeIDString(nativeKey2)
                    If fuzID2 = "" Then fuzID2 = "Tier 2 - No Explicit Job Plan"
                    dictT2JP(wsAsset.Name).Add rKey, fuzID2
                    dictJPDesc(wsAsset.Name).Add rKey, matchedT2Desc
                End If
            End If
        Next rKey
        sheetCounter = sheetCounter + 1
    Next wsAsset

    ' ---------------------------------------------------------
    ' PASS 3: TIER 3 FUZZY MATCHES (Cross-Campus Orphans)
    ' ---------------------------------------------------------
    sheetCounter = 1
    For Each wsAsset In assetSheets
        Application.StatusBar = "Pass 3: Cross-Campus Orphan Match (" & wsAsset.Name & " - " & sheetCounter & " of " & totalSheets & ")..."
        DoEvents

        Dim dictT3TypeCache As Object
        Set dictT3TypeCache = CreateObject("Scripting.Dictionary")
        dictT3TypeCache.CompareMode = vbTextCompare

        Dim rKey3 As Variant
        For Each rKey3 In dictCampusInventory(wsAsset.Name).Keys
            If Not dictUniqueCampusAssets(wsAsset.Name).exists(rKey3) Then

                Dim invData3() As String
                invData3 = Split(dictCampusInventory(wsAsset.Name)(rKey3), "||")
                Dim typePairKey3 As String, aTypeStr3 As String, aSubStr3 As String
                aTypeStr3 = invData3(1): aSubStr3 = invData3(2)
                typePairKey3 = aTypeStr3 & "|" & aSubStr3

                Dim matchedT3Desc As String
                matchedT3Desc = ""

                If dictT3TypeCache.exists(typePairKey3) Then
                    matchedT3Desc = dictT3TypeCache(typePairKey3)
                Else
                    If Len(aTypeStr3) > 0 Or Len(aSubStr3) > 0 Then
                        matchedT3Desc = GetBestJobPlanMatch(aTypeStr3, aSubStr3, dictCampusNonNativeDesc(wsAsset.Name))
                    End If
                    dictT3TypeCache.Add typePairKey3, matchedT3Desc
                End If

                If matchedT3Desc <> "" Then
                    Dim nativeKey3 As String
                    nativeKey3 = wsAsset.Name & "|" & matchedT3Desc

                    dictCellT3(nativeKey3) = dictCellT3(nativeKey3) + 1

                    dictUniqueCampusAssets(wsAsset.Name).Add rKey3, "Tier 3 Fuzzy"
                    dictTotalTier3(wsAsset.Name) = dictTotalTier3(wsAsset.Name) + 1

                    logArr(logIndex, 1) = wsAsset.Name
                    logArr(logIndex, 2) = matchedT3Desc
                    logArr(logIndex, 3) = "Tier 3 Fuzzy"
                    logArr(logIndex, 4) = IIf(invData3(0) = "", "Row " & rKey3, invData3(0))

                    If aTypeStr3 <> "" And aSubStr3 <> "" Then
                        logArr(logIndex, 5) = aTypeStr3 & " - " & aSubStr3
                    Else
                        logArr(logIndex, 5) = aTypeStr3 & aSubStr3
                    End If
                    logIndex = logIndex + 1

                    dictT3JP(wsAsset.Name).Add rKey3, "Tier 3 - Cross-Campus Match"
                    dictJPDesc(wsAsset.Name).Add rKey3, matchedT3Desc
                End If
            End If
        Next rKey3
        sheetCounter = sheetCounter + 1
    Next wsAsset

    ' Sweep remaining absolute orphans for the Drilldown Log
    For Each wsAsset In assetSheets
        Dim uKey As Variant
        For Each uKey In dictCampusInventory(wsAsset.Name).Keys
            If Not dictUniqueCampusAssets(wsAsset.Name).exists(uKey) Then

                Dim uArr() As String
                uArr = Split(dictCampusInventory(wsAsset.Name)(uKey), "||")

                logArr(logIndex, 1) = wsAsset.Name
                logArr(logIndex, 2) = "Unmatched Asset"
                logArr(logIndex, 3) = "Unmatched"
                logArr(logIndex, 4) = IIf(uArr(0) = "", "Row " & uKey, uArr(0))

                If uArr(1) <> "" And uArr(2) <> "" Then
                    logArr(logIndex, 5) = uArr(1) & " - " & uArr(2)
                Else
                    logArr(logIndex, 5) = uArr(1) & uArr(2)
                End If
                logIndex = logIndex + 1
            End If
        Next uKey
    Next wsAsset

    If logIndex > 1 Then
        wsLog.Range("A2").Resize(logIndex - 1, 5).Value = logArr
    End If

    ' ---------------------------------------------------------
    ' PASS 4: COMPILE OUTPUT MATRIX
    ' ---------------------------------------------------------
    Application.StatusBar = "Compiling Output Matrix..."
    DoEvents

    Dim dictProcessedDesc As Object
    Set dictProcessedDesc = CreateObject("Scripting.Dictionary")

    Dim outArray() As Variant
    ReDim outArray(1 To UBound(arrCompMaster, 1) - 1, 1 To assetSheets.Count)

    For i = 2 To UBound(arrCompMaster, 1)
        jpDesc = CleanFuzzyString(CStr(arrCompMaster(i, colDesc)))

        If jpDesc <> "" And jpDesc <> "No Description Found" And dictAllDesc.exists(jpDesc) Then
            If Not dictProcessedDesc.exists(jpDesc) Then
                dictProcessedDesc.Add jpDesc, 1
                Dim colOffset As Long
                colOffset = 1

                For Each wsAsset In assetSheets
                    Dim nKey As String
                    nKey = wsAsset.Name & "|" & jpDesc

                    Dim eC As Long, t2C As Long, t3C As Long
                    eC = 0: t2C = 0: t3C = 0
                    If dictCellExact.exists(nKey) Then eC = dictCellExact(nKey)
                    If dictCellT2.exists(nKey) Then t2C = dictCellT2(nKey)
                    If dictCellT3.exists(nKey) Then t3C = dictCellT3(nKey)

                    Dim idStr As String
                    If dictNativeIDString.exists(nKey) Then idStr = dictNativeIDString(nKey)

                    If idStr <> "" Or eC > 0 Or t2C > 0 Or t3C > 0 Then
                        Dim outText As String
                        If idStr = "" Then outText = "No Campus Plan" Else outText = idStr

                        Dim countStr As String
                        countStr = " (Exact: " & eC
                        If t2C > 0 Then countStr = countStr & " | T2: " & t2C
                        If t3C > 0 Then countStr = countStr & " | T3: " & t3C
                        countStr = countStr & ")"

                        outArray(i - 1, colOffset) = outText & countStr
                    Else
                        outArray(i - 1, colOffset) = ""
                    End If
                    colOffset = colOffset + 1
                Next wsAsset
            End If
        End If
    Next i

    wsComp.Cells(2, startCol).Resize(UBound(arrCompMaster, 1) - 1, assetSheets.Count).Value = outArray

    Dim totalRow As Long, sheetTotalRow As Long
    totalRow = lastRowComp + 2: sheetTotalRow = lastRowComp + 3

    wsComp.Cells(totalRow, colDesc).Value = "MATCHED TOTALS:"
    wsComp.Cells(totalRow, colDesc).Font.Bold = True
    wsComp.Cells(totalRow, colDesc).HorizontalAlignment = xlRight

    wsComp.Cells(sheetTotalRow, colDesc).Value = "TOTAL SHEET ASSETS:"
    wsComp.Cells(sheetTotalRow, colDesc).Font.Bold = True
    wsComp.Cells(sheetTotalRow, colDesc).HorizontalAlignment = xlRight

    c = startCol
    For Each wsAsset In assetSheets
        wsComp.Cells(totalRow, c).Value = "Total Plans: " & dictTotalPlans(wsAsset.Name) & " | Exact: " & dictTotalExact(wsAsset.Name) & " | T2 Fuzzy: " & dictTotalTier2(wsAsset.Name) & " | T3 Fuzzy: " & dictTotalTier3(wsAsset.Name) & " | Total Matched: " & dictUniqueCampusAssets(wsAsset.Name).Count
        wsComp.Cells(totalRow, c).Font.Bold = True
        wsComp.Cells(totalRow, c).Interior.Color = RGB(255, 242, 204)

        Dim rngLastFinal As Range, lastA As Long, rawCount As Long
        Set rngLastFinal = wsAsset.Cells.Find(What:="*", After:=wsAsset.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not rngLastFinal Is Nothing Then
            lastA = rngLastFinal.Row
            rawCount = lastA - 1
            If rawCount < 0 Then rawCount = 0
        Else
            rawCount = 0
        End If

        wsComp.Cells(sheetTotalRow, c).Value = "Raw Asset Count: " & rawCount
        wsComp.Cells(sheetTotalRow, c).Font.Bold = True
        wsComp.Cells(sheetTotalRow, c).Interior.Color = RGB(221, 235, 247)
        c = c + 1
    Next wsAsset

    Dim lastCol As Long
    lastCol = startCol + assetSheets.Count - 1
    wsComp.Columns(startCol).Resize(, assetSheets.Count).AutoFit
    wsComp.Activate
    If wsComp.AutoFilterMode Then wsComp.AutoFilterMode = False
    wsComp.Range(wsComp.Cells(1, 1), wsComp.Cells(1, lastCol)).AutoFilter
    ActiveWindow.FreezePanes = False
    wsComp.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    wsComp.Range("A1").Select

    ' ---------------------------------------------------------
    ' 7. ASSET SHEET WRITE-BACK LOGIC (ROW INDEX DRIVEN)
    ' ---------------------------------------------------------
    sheetCounter = 1
    For Each wsAsset In assetSheets
        Application.StatusBar = "Writing Data Back to Sheets (" & wsAsset.Name & " - " & sheetCounter & " of " & totalSheets & ")..."
        DoEvents

        If wsAsset.AutoFilterMode Then wsAsset.AutoFilterMode = False

        Dim colMainAsset As Long
        colMainAsset = 0
        For Each cell In wsAsset.Rows(1).Cells
            Dim writeHead As String
            writeHead = Trim(UCase(cell.Value))
            If (writeHead = "ASSET ID" Or writeHead = "ASSET #" Or writeHead = "ASSET" Or writeHead = "EQUIPMENT ID") And colMainAsset = 0 Then
                colMainAsset = cell.Column: Exit For
            End If
        Next cell

        If colMainAsset > 0 Then
            Dim colExact As Long, colFuzzy2 As Long, colFuzzy3 As Long, colDescMatch As Long
            colExact = 0: colFuzzy2 = 0: colFuzzy3 = 0: colDescMatch = 0

            Dim lc As Long
            lc = wsAsset.Cells(1, wsAsset.Columns.Count).End(xlToLeft).Column

            Dim cIter As Long
            For cIter = 1 To lc
                If Trim(wsAsset.Cells(1, cIter).Value) = "Exact Job Plan Match" Then colExact = cIter
                If Trim(wsAsset.Cells(1, cIter).Value) = "Tier 2 Fuzzy Match" Then colFuzzy2 = cIter
                If Trim(wsAsset.Cells(1, cIter).Value) = "Tier 3 Fuzzy Match" Then colFuzzy3 = cIter
                If Trim(wsAsset.Cells(1, cIter).Value) = "Matched Description" Then colDescMatch = cIter
            Next cIter

            If colExact = 0 Then
                lc = lc + 1: colExact = lc
                wsAsset.Cells(1, colExact).Value = "Exact Job Plan Match"
                wsAsset.Cells(1, colExact).Font.Bold = True
                wsAsset.Cells(1, colExact).Interior.Color = RGB(226, 239, 218)
            End If
            If colFuzzy2 = 0 Then
                lc = lc + 1: colFuzzy2 = lc
                wsAsset.Cells(1, colFuzzy2).Value = "Tier 2 Fuzzy Match"
                wsAsset.Cells(1, colFuzzy2).Font.Bold = True
                wsAsset.Cells(1, colFuzzy2).Interior.Color = RGB(255, 242, 204)
            End If
            If colFuzzy3 = 0 Then
                lc = lc + 1: colFuzzy3 = lc
                wsAsset.Cells(1, colFuzzy3).Value = "Tier 3 Fuzzy Match"
                wsAsset.Cells(1, colFuzzy3).Font.Bold = True
                wsAsset.Cells(1, colFuzzy3).Interior.Color = RGB(255, 204, 204)
            End If
            If colDescMatch = 0 Then
                lc = lc + 1: colDescMatch = lc
                wsAsset.Cells(1, colDescMatch).Value = "Matched Description"
                wsAsset.Cells(1, colDescMatch).Font.Bold = True
                wsAsset.Cells(1, colDescMatch).Interior.Color = RGB(221, 235, 247)
            End If

            Dim writeLastRow As Long
            Set rngLastFinal = wsAsset.Cells.Find(What:="*", After:=wsAsset.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not rngLastFinal Is Nothing Then writeLastRow = rngLastFinal.Row Else writeLastRow = 1

            If writeLastRow > 1 Then
                wsAsset.Range(wsAsset.Cells(2, colExact), wsAsset.Cells(writeLastRow, colExact)).ClearContents
                wsAsset.Range(wsAsset.Cells(2, colFuzzy2), wsAsset.Cells(writeLastRow, colFuzzy2)).ClearContents
                wsAsset.Range(wsAsset.Cells(2, colFuzzy3), wsAsset.Cells(writeLastRow, colFuzzy3)).ClearContents
                wsAsset.Range(wsAsset.Cells(2, colDescMatch), wsAsset.Cells(writeLastRow, colDescMatch)).ClearContents

                Dim arrExact(), arrF2(), arrF3(), arrDescOut()
                ReDim arrExact(1 To writeLastRow - 1, 1 To 1)
                ReDim arrF2(1 To writeLastRow - 1, 1 To 1)
                ReDim arrF3(1 To writeLastRow - 1, 1 To 1)
                ReDim arrDescOut(1 To writeLastRow - 1, 1 To 1)

                Dim writeRow As Long
                For writeRow = 2 To writeLastRow
                    Dim strRow As String
                    strRow = CStr(writeRow)
                    If dictExactJP(wsAsset.Name).exists(strRow) Then arrExact(writeRow - 1, 1) = dictExactJP(wsAsset.Name)(strRow)
                    If dictT2JP(wsAsset.Name).exists(strRow) Then arrF2(writeRow - 1, 1) = dictT2JP(wsAsset.Name)(strRow)
                    If dictT3JP(wsAsset.Name).exists(strRow) Then arrF3(writeRow - 1, 1) = dictT3JP(wsAsset.Name)(strRow)
                    If dictJPDesc(wsAsset.Name).exists(strRow) Then arrDescOut(writeRow - 1, 1) = dictJPDesc(wsAsset.Name)(strRow)
                Next writeRow

                wsAsset.Cells(2, colExact).Resize(UBound(arrExact, 1), 1).Value = arrExact
                wsAsset.Cells(2, colFuzzy2).Resize(UBound(arrF2, 1), 1).Value = arrF2
                wsAsset.Cells(2, colFuzzy3).Resize(UBound(arrF3, 1), 1).Value = arrF3
                wsAsset.Cells(2, colDescMatch).Resize(UBound(arrDescOut, 1), 1).Value = arrDescOut
            End If
            wsAsset.Columns(colExact).AutoFit
            wsAsset.Columns(colFuzzy2).AutoFit
            wsAsset.Columns(colFuzzy3).AutoFit
            wsAsset.Columns(colDescMatch).AutoFit
        End If
        sheetCounter = sheetCounter + 1
    Next wsAsset

    Application.StatusBar = False
End Sub

' ==============================================================================
' 8. HELPER FUNCTIONS: Prefix Locator & String Sanitizer
' ==============================================================================
Function GetSheetPrefixes(sheetName As String) As String
    Dim sName As String
    sName = UCase(sheetName)

    If InStr(sName, "16 Y") > 0 Or InStr(sName, "16Y") > 0 Then GetSheetPrefixes = "16Y": Exit Function
    If InStr(sName, "160") > 0 Then GetSheetPrefixes = "FSD": Exit Function
    If InStr(sName, "DON") > 0 Then GetSheetPrefixes = "DON": Exit Function
    If InStr(sName, "MLS") > 0 Then GetSheetPrefixes = "MLS": Exit Function
    If InStr(sName, "RBCC") > 0 Or InStr(sName, "RBC") > 0 Then GetSheetPrefixes = "AHHRBC,ASHRBC,RBC": Exit Function
    If InStr(sName, "SPL") > 0 Then GetSheetPrefixes = "SPL": Exit Function
    If InStr(sName, "TDC") > 0 Then GetSheetPrefixes = "ASHT1,TDC": Exit Function
    If InStr(sName, "VWP") > 0 Then GetSheetPrefixes = "VWP": Exit Function

    If InStr(sName, "GRA") > 0 Then GetSheetPrefixes = "GRA": Exit Function
    If InStr(sName, "PWC") > 0 Then GetSheetPrefixes = "PWC": Exit Function
    If InStr(sName, "STA") > 0 Then GetSheetPrefixes = "STA": Exit Function
    If InStr(sName, "WAT") > 0 Then GetSheetPrefixes = "WAT": Exit Function
    If InStr(sName, "BAY") > 0 Then GetSheetPrefixes = "BAY": Exit Function

    GetSheetPrefixes = Split(sName, " ")(0)
End Function

Function CleanFuzzyString(ByVal txt As String) As String
    If Len(txt) = 0 Then Exit Function
    txt = Replace(txt, Chr(160), " ")
    txt = Replace(txt, Chr(9), " ")
    txt = Replace(txt, Chr(10), " ")
    txt = Replace(txt, Chr(13), " ")
    txt = Trim(txt)
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    CleanFuzzyString = txt
End Function

' ==============================================================================
' 9. PIVOT GENERATOR
' ==============================================================================
Sub CreateAssetPivotTable()
    Dim wb As Workbook
    Dim wsLog As Worksheet, wsPivot As Worksheet
    Dim pc As PivotCache, pt As PivotTable
    Dim pvtRange As Range, lastRow As Long

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsLog = wb.Sheets("Audit Drilldown Data")
    On Error GoTo 0
    If wsLog Is Nothing Then Exit Sub

    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Set pvtRange = wsLog.Range("A1:E" & lastRow)

    On Error Resume Next
    Set wsPivot = wb.Sheets("Asset Drilldown Pivot")
    On Error GoTo 0

    If wsPivot Is Nothing Then
        Set wsPivot = wb.Sheets.Add(After:=wb.Sheets("Job Plan Comparison"))
        wsPivot.Name = "Asset Drilldown Pivot"
    Else
        Dim ptOld As PivotTable
        For Each ptOld In wsPivot.PivotTables
            ptOld.TableRange2.Clear
        Next ptOld
        wsPivot.Cells.Clear
    End If

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pvtRange)
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("B3"), TableName:="DrilldownPivot")

    With pt
        .PivotFields("Campus").Orientation = xlRowField
        .PivotFields("Campus").Position = 1
        .PivotFields("Job Plan Description").Orientation = xlRowField
        .PivotFields("Job Plan Description").Position = 2
        .PivotFields("Match Type").Orientation = xlColumnField
        .PivotFields("Match Type").Position = 1
        .AddDataField .PivotFields("Asset ID"), "Count of Assets", xlCount
        .RowAxisLayout xlOutlineRow
        .TableStyle2 = "PivotStyleMedium9"
    End With

    wsPivot.Columns("B:E").AutoFit
    wsPivot.Activate
    wsPivot.Range("A1").Select
End Sub

' ==============================================================================
' 10. CREATE MASTER CONSOLIDATED DATA SHEET
' ==============================================================================
Sub CreateMasterDataSheet()
    Dim wb As Workbook
    Dim wsMaster As Worksheet, wsAsset As Worksheet
    Dim lastRowMaster As Long, lastRowAsset As Long, lastColAsset As Long
    Dim headersCopied As Boolean
    Dim rngLast As Range, dataRange As Range
    Dim objTable As ListObject

    Set wb = ActiveWorkbook

    Application.StatusBar = "Building Master Data Sheet..."
    DoEvents

    On Error Resume Next
    Set wsMaster = wb.Sheets("Master Asset Data")
    On Error GoTo 0

    If wsMaster Is Nothing Then
        Set wsMaster = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsMaster.Name = "Master Asset Data"
    Else
        wsMaster.Cells.Clear
        For Each objTable In wsMaster.ListObjects
            objTable.Unlist
        Next objTable
        wsMaster.Cells.Clear
    End If

    headersCopied = False
    lastRowMaster = 1

    For Each wsAsset In wb.Worksheets
        If InStr(1, wsAsset.Name, "Assets", vbTextCompare) > 0 And wsAsset.Name <> "Job Plan Comparison" And wsAsset.Name <> "Master Asset Data" Then

            If wsAsset.AutoFilterMode Then wsAsset.AutoFilterMode = False

            Set rngLast = wsAsset.Cells.Find(What:="*", After:=wsAsset.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not rngLast Is Nothing Then lastRowAsset = rngLast.Row Else lastRowAsset = 1

            lastColAsset = wsAsset.Cells(1, wsAsset.Columns.Count).End(xlToLeft).Column

            If lastRowAsset > 1 Then
                If Not headersCopied Then
                    wsMaster.Cells(1, 1).Value = "Campus Prefix"
                    wsMaster.Cells(1, 2).Value = "Source Sheet"
                    wsAsset.Range(wsAsset.Cells(1, 1), wsAsset.Cells(1, lastColAsset)).Copy Destination:=wsMaster.Cells(1, 3)
                    headersCopied = True
                End If

                Set dataRange = wsAsset.Range(wsAsset.Cells(2, 1), wsAsset.Cells(lastRowAsset, lastColAsset))

                lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 3).End(xlUp).Row + 1

                wsMaster.Cells(lastRowMaster, 3).Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

                Dim campusPrefix As String
                campusPrefix = Split(GetSheetPrefixes(wsAsset.Name), ",")(0)

                wsMaster.Range(wsMaster.Cells(lastRowMaster, 1), wsMaster.Cells(lastRowMaster + dataRange.Rows.Count - 1, 1)).Value = campusPrefix
                wsMaster.Range(wsMaster.Cells(lastRowMaster, 2), wsMaster.Cells(lastRowMaster + dataRange.Rows.Count - 1, 2)).Value = wsAsset.Name
            End If
        End If
    Next wsAsset

    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 3).End(xlUp).Row
    lastColAsset = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column

    If lastRowMaster > 1 Then
        Dim tblRange As Range
        Set tblRange = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(lastRowMaster, lastColAsset))

        Set objTable = wsMaster.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        objTable.Name = "MasterDataTable"
        objTable.TableStyle = "TableStyleMedium2"

        With objTable.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsMaster.Range("A2:A" & lastRowMaster), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With

        wsMaster.Columns.AutoFit
    End If

    wsMaster.Activate
    wsMaster.Range("A1").Select
End Sub

' ==============================================================================
' 11. MASTER CONTROL SCRIPT
' ==============================================================================
Sub UpdateAllJobPlanData()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Call CompareJobPlans
    Call IdentifyJobPlanFrequencies
    Call AddJobPlanDescriptions
    Call ScrubDescriptionFrequencies
    Call GroupAssetsByJobPlanDescription
    Call CreateAssetPivotTable
    Call CreateMasterDataSheet

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "Audit Complete! Matrix, Asset Sheets, Master Data, and Pivot Table have been successfully updated.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    MsgBox "An error occurred during the update: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 12. UN-NESTED WATERFALL ENGINE (Strict Hierarchy)
' ==============================================================================
Function NormalizeDesc(ByVal txt As String) As String
    Dim jNorm As String
    jNorm = UCase(Trim(txt))
    jNorm = Replace(Replace(jNorm, "SWITCH GEAR", "SWITCHGEAR"), "TANKS", "TANK")
    jNorm = Replace(Replace(jNorm, "PUMPS", "PUMP"), "VALVES", "VALVE")
    jNorm = Replace(Replace(jNorm, "DAMPERS", "DAMPER"), "WATER HEATERS", "WATER HEATER")
    jNorm = Replace(jNorm, "AIR TERMINALS", "AIR TERMINAL")
    NormalizeDesc = jNorm
End Function

Function GetBestJobPlanMatch(ByVal aType As String, ByVal aSub As String, ByVal dictDesc As Object) As String
    Dim jpDesc As String, jNorm As String
    Dim key As Variant

    aType = UCase(Trim(aType))
    aSub = UCase(Trim(aSub))

    ' Clean up Plurals and common formatting in Asset strings
    aType = Replace(Replace(aType, "SWITCH GEAR", "SWITCHGEAR"), "TANKS", "TANK")
    aType = Replace(Replace(aType, "PUMPS", "PUMP"), "VALVES", "VALVE")
    aSub = Replace(Replace(aSub, "SWITCH GEAR", "SWITCHGEAR"), "TANKS", "TANK")
    aSub = Replace(Replace(aSub, "PUMPS", "PUMP"), "VALVES", "VALVE")

    ' Ignore dimensions masquerading as subtypes
    If UCase(aSub) Like "SIZE*" Or UCase(aSub) Like "*#X#*" Or UCase(aSub) Like "*LB*" Then aSub = ""

    Dim aFull As String, aFullRev As String
    aFull = Trim(aType & " " & aSub)
    aFullRev = Trim(aSub & " " & aType)

    ' =========================================================================
    ' STEP 1: EXACT MATCHES
    ' =========================================================================
    If aFull <> "" Then
        For Each key In dictDesc.Keys
            jNorm = NormalizeDesc(CStr(key))
            If aFull = jNorm Then GetBestJobPlanMatch = CStr(key): Exit Function
        Next key
    End If

    If aFullRev <> "" Then
        For Each key In dictDesc.Keys
            jNorm = NormalizeDesc(CStr(key))
            If aFullRev = jNorm Then GetBestJobPlanMatch = CStr(key): Exit Function
        Next key
    End If

    If aSub <> "" Then
        For Each key In dictDesc.Keys
            jNorm = NormalizeDesc(CStr(key))
            If aSub = jNorm Then GetBestJobPlanMatch = CStr(key): Exit Function
        Next key
    End If

    If aType <> "" Then
        For Each key In dictDesc.Keys
            jNorm = NormalizeDesc(CStr(key))
            If aType = jNorm Then GetBestJobPlanMatch = CStr(key): Exit Function
        Next key
    End If

    ' =========================================================================
    ' STEP 2: SPECIFIC MATCHING PHRASES (Listed Rules)
    ' =========================================================================
    For Each key In dictDesc.Keys
        jpDesc = CStr(key)
        jNorm = NormalizeDesc(jpDesc)

        If (aType = "VAV" Or aSub = "VAV" Or aType = "VAV BOX") And InStr(jNorm, "AIR TERMINAL") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "PUMP" And InStr(aSub, "CENTRIFUGAL") > 0 And InStr(jNorm, "PUMP") > 0 And InStr(jNorm, "FIRE") = 0 And InStr(jNorm, "SUMP") = 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "PUMP" And aSub = "" And InStr(jNorm, "PUMP") > 0 And InStr(jNorm, "FIRE") = 0 And InStr(jNorm, "SUMP") = 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "SUMP PUMP" And InStr(jNorm, "SUMP PUMP") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "TANK" And InStr(aSub, "EXPANSION") > 0 And InStr(jNorm, "EXPANSION") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "TANK" And InStr(aSub, "CONDENSATE") > 0 And InStr(jNorm, "CONDENSATE") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf InStr(aType, "FUEL") > 0 And InStr(aType, "TANK") > 0 And InStr(jNorm, "DIESEL") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "SWITCHGEAR" And InStr(jNorm, "SWITCHGEAR") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "TRANSFORMER" And InStr(jNorm, "TRANSFORMER") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf (InStr(aType, "WATER HEATER") > 0 Or InStr(aType, "WATER TANK") > 0) And InStr(jNorm, "WATER HEATER") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "VFD" And InStr(jNorm, "VFD") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "SPRINKLER" And InStr(jNorm, "SPRINKLER") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "CARD READER" And InStr(jNorm, "SECURITY SYSTEM") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "MOTOR CONTROL CENTER" And (jNorm = "MCC" Or InStr(jNorm, "MOTOR CONTROL") > 0) Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "DAMPER" And InStr(jNorm, "DAMPER") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf (InStr(aType, "BACK FLOW") > 0 Or InStr(aType, "BACKFLOW") > 0) And InStr(jNorm, "BACKFLOW") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf InStr(aFull, "ELEVATOR") > 0 And InStr(aFull, "TRACTION") > 0 And InStr(jNorm, "TRACTION ELEVATOR") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        ElseIf aType = "FIRE EXTINGUISHER" And InStr(jNorm, "EXTINGUISHER") > 0 Then
            GetBestJobPlanMatch = jpDesc: Exit Function
        End If
    Next key

    ' =========================================================================
    ' STEP 3: WORD PATTERNS FROM BOTH (Safe Multi-Word Substrings Un-Nested)
    ' =========================================================================
    If aType <> "" And aSub <> "" Then
        For Each key In dictDesc.Keys
            jpDesc = CStr(key)
            jNorm = NormalizeDesc(jpDesc)
            If InStr(jNorm, aType) > 0 And InStr(jNorm, aSub) > 0 Then
                GetBestJobPlanMatch = jpDesc: Exit Function
            End If
        Next key
    End If

    If aFull <> "" And InStr(aFull, " ") > 0 Then
        For Each key In dictDesc.Keys
            jpDesc = CStr(key)
            jNorm = NormalizeDesc(jpDesc)
            If InStr(jNorm, aFull) > 0 Then GetBestJobPlanMatch = jpDesc: Exit Function
        Next key
    End If

    If aType <> "" And InStr(aType, " ") > 0 Then
        For Each key In dictDesc.Keys
            jpDesc = CStr(key)
            jNorm = NormalizeDesc(jpDesc)
            If InStr(jNorm, aType) > 0 Then GetBestJobPlanMatch = jpDesc: Exit Function
        Next key
    End If

    If aSub <> "" And InStr(aSub, " ") > 0 Then
        For Each key In dictDesc.Keys
            jpDesc = CStr(key)
            jNorm = NormalizeDesc(jpDesc)
            If InStr(jNorm, aSub) > 0 Then GetBestJobPlanMatch = jpDesc: Exit Function
        Next key
    End If

    GetBestJobPlanMatch = ""
End Function
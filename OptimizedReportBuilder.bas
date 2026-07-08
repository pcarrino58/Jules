Attribute VB_Name = "OptimizedReportBuilder"
Option Explicit

'=========================================================
' Dynamic Non-Base Cost Report Builder (Optimized)
'
' Creates:
'   1) Non Base Cost Data MMDD
'   2) Non Base Cost By Month MMDD       - real PivotTable
'   3) Non Base Cost By [Client] PO MMDD - real PivotTable
'
' Raw source sheets expected:
'   Invoice Data (Dynamically imported from Downloads)
'   Labour Data (Dynamically imported from Downloads)
'   Labour Rates / Labour Rates 2026-0202 (Static)
'   M# Lookup (Static)
'   Commodity (Static)
'
' Main macro to run:
'   BuildNonBaseCostReport
'=========================================================

Private Const REPORT_FY As String = "FY26"

Private Const EXACT_EXISTING_LABOUR_DATE_LOGIC As Boolean = False
'If set to False, labour rows will use Labour Date instead of matching Invoice Data Approval Date.

Public Sub BuildNonBaseCostReport()

    Dim wb As Workbook
    Dim wsInv As Worksheet
    Dim wsLab As Worksheet
    Dim wsRates As Worksheet
    Dim wsM As Worksheet
    Dim wsComm As Worksheet
    Dim wsData As Worksheet
    Dim wsMonth As Worksheet
    Dim wsPO As Worksheet

    Dim outSuffix As String
    Dim downloadsPath As String
    Dim invFilePath As String
    Dim labFilePath As String

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean

    Dim errMsg As String

    On Error GoTo CleanFail

    Set wb = ThisWorkbook

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Determine dynamic suffix (e.g., " 0630" for current month/day)
    outSuffix = " " & Format(Date, "mmdd")

    ' --- Dynamic Folder Picker ---
    With Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
        .Title = "Select the folder containing the Ainsworth CSV files"
        .AllowMultiSelect = False
        If .Show = -1 Then ' If the user clicks OK
            downloadsPath = .SelectedItems(1) & "\"
        Else
            MsgBox "Folder selection canceled. Macro stopped.", vbExclamation
            Application.ScreenUpdating = oldScreenUpdating
            Application.DisplayAlerts = oldDisplayAlerts
            Application.Calculation = oldCalculation
            Application.EnableEvents = oldEnableEvents
            Exit Sub
        End If
    End With

    ' --- Dynamic Client Selection ---
    Dim clientDict As Object
    Dim selectedClient As String
    Dim clientKeys As Variant
    Dim promptMsg As String
    Dim userChoice As String
    Dim choiceIndex As Integer
    Dim i As Integer

    Set clientDict = GetAvailableClients(downloadsPath)

    If clientDict.Count = 0 Then
        Err.Raise vbObjectError + 110, , "No client files found in the selected folder."
    ElseIf clientDict.Count = 1 Then
        ' Auto-select if only one client exists in the folder
        selectedClient = clientDict.Keys()(0)
    Else
        ' Prompt user if multiple clients are present
        clientKeys = clientDict.Keys()
        promptMsg = "Multiple clients found. Enter the number for the client you want to process:" & vbCrLf & vbCrLf
        For i = LBound(clientKeys) To UBound(clientKeys)
            promptMsg = promptMsg & (i + 1) & ". " & clientKeys(i) & vbCrLf
        Next i

        userChoice = InputBox(promptMsg, "Select Client")

        If userChoice = "" Or Not IsNumeric(userChoice) Then
            Err.Raise vbObjectError + 111, , "Invalid or canceled client selection."
        End If

        choiceIndex = CInt(userChoice) - 1
        If choiceIndex < LBound(clientKeys) Or choiceIndex > UBound(clientKeys) Then
            Err.Raise vbObjectError + 112, , "Invalid client number selected."
        End If

        selectedClient = clientKeys(choiceIndex)
    End If

    ' --- Find the newest matching files using wildcard patterns ---
    invFilePath = GetNewestFile(downloadsPath, "*[" & selectedClient & "]*Invoices*.csv")
    labFilePath = GetNewestFile(downloadsPath, "*[" & selectedClient & "]*Labour*.csv")

    If invFilePath = "" Then Err.Raise vbObjectError + 105, , "Could not find Invoices source file in Downloads."
    If labFilePath = "" Then Err.Raise vbObjectError + 106, , "Could not find Labour source file in Downloads."

    ' Import the new dynamic data sheets from the downloaded files
    Set wsInv = ImportDataSheet(wb, invFilePath, "Invoices Data" & outSuffix)
    Set wsLab = ImportDataSheet(wb, labFilePath, "Labour Data" & outSuffix)

    SplitInvoiceAccountData wsInv

    ' Find static sheets
    Set wsRates = FindSheetLike(wb, Array("labour", "rate"), Array("data", "non base"))
    If wsRates Is Nothing Then Set wsRates = FindSheetLike(wb, Array("labor", "rate"), Array("data", "non base"))

    Set wsM = FindSheetLike(wb, Array("m", "lookup"), Array())
    If wsM Is Nothing Then Set wsM = FindSheetLike(wb, Array("lookup"), Array())

    Set wsComm = FindSheetLike(wb, Array("commodity"), Array())

    If wsRates Is Nothing Then Err.Raise vbObjectError + 102, , "Labour Rates sheet was not found."
    If wsM Is Nothing Then Err.Raise vbObjectError + 103, , "M# Lookup sheet was not found."
    If wsComm Is Nothing Then Err.Raise vbObjectError + 104, , "Commodity sheet was not found."

    FillLabourRatesAndCosts wsLab, wsRates

    ' Recreate output sheets.
    Set wsData = RecreateSheet(wb, "Non Base Cost Data" & outSuffix)
    Set wsMonth = RecreateSheet(wb, "Non Base Cost By Month" & outSuffix)

    ' Safe Sheet Naming to prevent Error 1004 (31 char limit)
    Dim poSheetName As String
    poSheetName = "PO Cost " & selectedClient & outSuffix
    If Len(poSheetName) > 31 Then
        poSheetName = Left("PO Cost " & selectedClient, 31 - Len(outSuffix)) & outSuffix
    End If
    Set wsPO = RecreateSheet(wb, poSheetName)

    'Build source data and real PivotTables.
    BuildNonBaseData wsInv, wsLab, wsM, wsComm, wsData
    CreateNormalPivotTables wsData, wsMonth, wsPO, REPORT_FY

    wsInv.Cells.EntireColumn.AutoFit
    wsInv.Cells.EntireRow.AutoFit

    wsLab.Cells.EntireColumn.AutoFit
    wsLab.Cells.EntireRow.AutoFit

    wsData.Cells.EntireColumn.AutoFit
    wsData.Cells.EntireRow.AutoFit

    wsMonth.Cells.EntireColumn.AutoFit
    wsMonth.Cells.EntireRow.AutoFit

    wsPO.Cells.EntireColumn.AutoFit
    wsPO.Cells.EntireRow.AutoFit

CleanExit:

    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents

    If Err.Number = 0 Then
        MsgBox selectedClient & " Non-Base Cost data and PivotTables have been rebuilt for " & REPORT_FY & " using suffix " & outSuffix & ".", vbInformation
    End If

    Exit Sub

CleanFail:
    errMsg = "Macro stopped." & vbCrLf & vbCrLf & _
             "Error " & Err.Number & ": " & Err.Description

    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0

    MsgBox errMsg, vbExclamation

End Sub

 '=========================================================
' Pre-Processing: Split Invoice Account String (Optimized)
'=========================================================

Private Sub SplitInvoiceAccountData(ByVal wsInv As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim rawVal As String
    Dim parts() As String

    ' Set Headers in L (12), M (13), N (14)
    wsInv.Cells(1, 12).Value = "M#"
    wsInv.Cells(1, 13).Value = "Account"
    wsInv.Cells(1, 14).Value = "GL"

    lastRow = LastUsedRow(wsInv)
    If lastRow < 2 Then Exit Sub

    ' Load Column C (3) into memory
    Dim srcData As Variant
    srcData = wsInv.Range(wsInv.Cells(2, 3), wsInv.Cells(lastRow, 3)).Value

    Dim arrSize As Long
    If lastRow = 2 Then
        arrSize = 1
    Else
        arrSize = UBound(srcData, 1)
    End If

    Dim outData() As Variant
    ReDim outData(1 To arrSize, 1 To 3)

    ' Loop through rows and split Column C (3) by period
    For r = 1 To arrSize
        If lastRow = 2 Then
            rawVal = CStr(srcData)
        Else
            rawVal = CStr(srcData(r, 1))
        End If

        If Len(rawVal) > 0 Then
            parts = Split(rawVal, ".")

            ' Write the pieces to L, M, and N arrays
            If UBound(parts) >= 0 Then outData(r, 1) = parts(0)
            If UBound(parts) >= 1 Then outData(r, 2) = parts(1)
            If UBound(parts) >= 2 Then outData(r, 3) = parts(2)
        End If
    Next r

    ' Write back to sheet in one bulk operation
    wsInv.Range(wsInv.Cells(2, 12), wsInv.Cells(lastRow, 14)).Value = outData

    ' Format columns for readability
    wsInv.Columns("L:N").AutoFit
End Sub

'=========================================================
' File system utilities
'=========================================================
Private Function GetNewestFile(ByVal folderPath As String, ByVal filePattern As String) As String
    Dim fileName As String
    Dim newestFile As String
    Dim maxDate As Date
    Dim fileDate As Date

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & filePattern)

    Do While fileName <> ""
        fileDate = FileDateTime(folderPath & fileName)
        If fileDate > maxDate Then
            maxDate = fileDate
            newestFile = folderPath & fileName
        End If
        fileName = Dir()
    Loop

    GetNewestFile = newestFile
End Function

Private Function ImportDataSheet(ByVal targetWb As Workbook, ByVal filePath As String, ByVal newSheetName As String) As Worksheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim newWs As Worksheet

    'Delete if it already exists
    On Error Resume Next
    targetWb.Worksheets(newSheetName).Delete
    On Error GoTo 0

    'Open the source CSV file silently, forcing local date/number formats
    Set srcWb = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=False, Local:=True)
    Set srcWs = srcWb.Worksheets(1) 'Assume data is on the first sheet

    'Copy the sheet into the target workbook
    srcWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)
    Set newWs = targetWb.Worksheets(targetWb.Worksheets.Count)
    newWs.Name = newSheetName

    'Close the source file
    srcWb.Close SaveChanges:=False

    Set ImportDataSheet = newWs
End Function

'=========================================================
' Step 1: Fill Labour Rate and Labour Cost (Optimized)
'=========================================================

Private Sub FillLabourRatesAndCosts(ByVal wsLab As Worksheet, ByVal wsRates As Worksheet)

    Dim cCreated As Long
    Dim cLabContact As Long
    Dim cLabFirst As Long
    Dim cLabLast As Long
    Dim cLabHours As Long
    Dim cLabRateType As Long
    Dim cLabRate As Long
    Dim cLabCost As Long
    Dim lastColLab As Long

    Dim cRatesContact As Long
    Dim cRatesType As Long
    Dim cRatesValue As Long

    Dim lastLab As Long
    Dim lastRates As Long

    Dim r As Long

    Dim rateDict As Object
    Dim key As String
    Dim rateVal As Variant
    Dim hoursVal As Variant
    Dim contactName As String

    ' Added "Created On" to catch new export format
    cCreated = HeaderColAny(wsLab, Array("CreatedOnDate", "Created On Date", "Created On"), False)
    If cCreated > 0 Then
        wsLab.Columns(cCreated).Delete
    End If

    Set rateDict = CreateObject("Scripting.Dictionary")

    ' Labour export formats vary. New Nissan labour exports do not always include a clean Contact field.
    ' Use c_firstname + c_lastname when available; fall back to Contact/Target only if needed.
    cLabFirst = HeaderColAny(wsLab, Array("First Name", "Firstname", "c_firstname"), False)
    cLabLast = HeaderColAny(wsLab, Array("Last Name", "Lastname", "c_lastname"), False)
    cLabContact = HeaderColAny(wsLab, Array("Contact", "Target"), False)

    If cLabFirst = 0 Or cLabLast = 0 Then
        If cLabContact = 0 Then
            Err.Raise vbObjectError + 201, , "Missing required labour name headers on sheet '" & wsLab.Name & "'. Expected c_firstname/c_lastname or Contact/Target."
        End If
    End If

    cLabHours = HeaderColAny(wsLab, Array("Labor Hours", "Labour Hours", "Quantity"))
    cLabRateType = HeaderColAny(wsLab, Array("Rate Type"))

    ' Dynamically find the end of the sheet so we don't overwrite data
    lastColLab = wsLab.Cells(1, wsLab.Columns.Count).End(xlToLeft).Column
    cLabRate = lastColLab + 1
    cLabCost = lastColLab + 2
    wsLab.Cells(1, cLabRate).Value = "Labour Rate"
    wsLab.Cells(1, cLabCost).Value = "Labour Cost"

    cRatesContact = HeaderColAny(wsRates, Array("Contact"))
    cRatesType = HeaderColAny(wsRates, Array("Rate Type"))
    cRatesValue = HeaderColAny(wsRates, Array("Value", "Rate", "Labour Rate", "Labor Rate"))

    lastRates = LastUsedRow(wsRates)

    ' Load Rates into array
    If lastRates >= 2 Then
        Dim ratesData As Variant
        ratesData = wsRates.Range(wsRates.Cells(1, 1), wsRates.Cells(lastRates, wsRates.Cells(1, wsRates.Columns.Count).End(xlToLeft).Column)).Value

        For r = 2 To lastRates
            key = MakeRateKey(ratesData(r, cRatesContact), ratesData(r, cRatesType))
            If Len(key) > 1 Then
                If IsNumeric(ratesData(r, cRatesValue)) Then
                    rateDict(key) = CDbl(ratesData(r, cRatesValue))
                End If
            End If
        Next r
    End If

    lastLab = LastUsedRow(wsLab)

    If lastLab >= 2 Then
        Dim labData As Variant
        labData = wsLab.Range(wsLab.Cells(1, 1), wsLab.Cells(lastLab, lastColLab)).Value

        Dim outRates() As Variant
        ReDim outRates(1 To lastLab - 1, 1 To 2) ' Rate and Cost

        For r = 2 To lastLab
            If cLabFirst > 0 And cLabLast > 0 Then
                contactName = Trim$(CStr(labData(r, cLabFirst)) & " " & CStr(labData(r, cLabLast)))
            Else
                contactName = CStr(labData(r, cLabContact))
            End If

            key = MakeRateKey(contactName, labData(r, cLabRateType))

            If rateDict.Exists(key) Then
                outRates(r - 1, 1) = rateDict(key)
            End If

            rateVal = outRates(r - 1, 1)
            hoursVal = labData(r, cLabHours)

            If IsNumeric(rateVal) And IsNumeric(hoursVal) Then
                outRates(r - 1, 2) = CDbl(rateVal) * CDbl(hoursVal)
            End If
        Next r

        ' Write back to sheet in one bulk operation
        wsLab.Range(wsLab.Cells(2, cLabRate), wsLab.Cells(lastLab, cLabCost)).Value = outRates
    End If

    wsLab.Columns(cLabRate).NumberFormat = "0.00"
    wsLab.Columns(cLabCost).NumberFormat = "0.00"

End Sub

'=========================================================
' Step 2: Build combined Non Base Cost Data (Optimized)
'=========================================================

Private Sub BuildNonBaseData( _
    ByVal wsInv As Worksheet, _
    ByVal wsLab As Worksheet, _
    ByVal wsM As Worksheet, _
    ByVal wsComm As Worksheet, _
    ByVal wsData As Worksheet)

    Dim mDict As Object
    Dim commDict As Object

    Dim outRow As Long
    Dim r As Long
    Dim lastInv As Long
    Dim lastLab As Long

    Dim cInvBuilding As Long
    Dim cInvReq As Long
    Dim cInvAcctNo As Long
    Dim cInvAmount As Long
    Dim cInvWO As Long
    Dim cInvType As Long
    Dim cInvPO As Long
    Dim cInvBrief As Long
    Dim cInvWorkDesc As Long
    Dim cInvComments As Long
    Dim cInvDate As Long
    Dim cInvM As Long
    Dim cInvAcct As Long
    Dim cInvGL As Long

    Dim cLabBuilding As Long
    Dim cLabWO As Long
    Dim cLabType As Long
    Dim cLabPO As Long
    Dim cLabBrief As Long
    Dim cLabWorkDesc As Long
    Dim cLabComments As Long
    Dim cLabDate As Long
    Dim cLabCost As Long

    Dim dtVal As Variant
    Dim glVal As Variant
    Dim buildingName As String

    Set mDict = BuildMLookup(wsM)
    Set commDict = BuildCommodityLookup(wsComm)

    With wsData
        .Range("A1:S1").Value = Array( _
            "Source", _
            "Building Name", _
            "Requisition ID", _
            "Account Number", _
            "Invoice Amount", _
            "Work Order No.", _
            "WO Type", _
            "Customer PO #", _
            "Brief Description", _
            "Work Description", _
            "Technician / Employee Comments", _
            "Invoice Approval Date", _
            "M#", _
            "Account", _
            "GL", _
            "Commodity Category", _
            "Nissan FY", _
            "Invoice Month", _
            "Building Filter")
    End With

    ' Updated Arrays for Invoice Data just in case their format changed too
    cInvBuilding = HeaderColAny(wsInv, Array("Building Name", "c_buildingname"))
    cInvReq = HeaderColAny(wsInv, Array("Requisition ID", "c_requisitionid"))
    cInvAcctNo = HeaderColAny(wsInv, Array("Account Number", "c_accountno"))
    cInvAmount = HeaderColAny(wsInv, Array("Invoice Amount"))
    cInvWO = HeaderColAny(wsInv, Array("Work Order No.", "Work Order No", "c_workorderno"))
    cInvType = HeaderColAny(wsInv, Array("WO Type", "c_wotype"))
    cInvPO = HeaderColAny(wsInv, Array("Customer PO #", "Customer PO", "c_customerpo"))
    cInvBrief = HeaderColAny(wsInv, Array("Brief Description", "Description"))
    cInvWorkDesc = HeaderColAny(wsInv, Array("Work Description", "details"))
    cInvComments = HeaderColAny(wsInv, Array("Technician / Employee Comments", "c_comments"))
    cInvDate = HeaderColAny(wsInv, Array("Invoice Approval Date"))
    cInvM = HeaderColAny(wsInv, Array("M#"))
    cInvAcct = HeaderColAny(wsInv, Array("Account"))
    cInvGL = HeaderColAny(wsInv, Array("GL"))

    lastInv = LastUsedRow(wsInv)
    Dim invData As Variant
    If lastInv >= 2 Then
        invData = wsInv.Range(wsInv.Cells(1, 1), wsInv.Cells(lastInv, wsInv.Cells(1, wsInv.Columns.Count).End(xlToLeft).Column)).Value
    End If

    ' Updated Arrays for Labour Data based on the new CSV format
    cLabBuilding = HeaderColAny(wsLab, Array("Building Name", "c_buildingname"))
    cLabWO = HeaderColAny(wsLab, Array("Work Order No.", "Work Order No", "c_workorderno"))
    cLabType = HeaderColAny(wsLab, Array("WO Type", "c_wotype"))
    cLabPO = HeaderColAny(wsLab, Array("Customer PO #", "Customer PO", "c_customerpo"))
    cLabBrief = HeaderColAny(wsLab, Array("Brief Description", "Description"), False)
    cLabWorkDesc = HeaderColAny(wsLab, Array("Work Description", "details"), False)
    cLabComments = HeaderColAny(wsLab, Array("Technician / Employee Comments", "c_comments"), False)
    cLabDate = HeaderColAny(wsLab, Array("Labour Date", "Labor Date"))
    cLabCost = HeaderColAny(wsLab, Array("Labour Cost", "Labor Cost"))

    lastLab = LastUsedRow(wsLab)
    Dim labData As Variant
    If lastLab >= 2 Then
        labData = wsLab.Range(wsLab.Cells(1, 1), wsLab.Cells(lastLab, wsLab.Cells(1, wsLab.Columns.Count).End(xlToLeft).Column)).Value
    End If

    Dim maxOutRows As Long
    maxOutRows = IIf(lastInv >= 2, lastInv, 0) + IIf(lastLab >= 2, lastLab, 0)

    If maxOutRows = 0 Then GoTo FormatSheet

    Dim outArray() As Variant
    ReDim outArray(1 To maxOutRows, 1 To 19)
    outRow = 1

    If lastInv >= 2 Then
        For r = 2 To lastInv
            If Len(Trim$(CStr(invData(r, cInvType)))) > 0 Then
                glVal = invData(r, cInvGL)
                dtVal = invData(r, cInvDate)
                buildingName = CStr(invData(r, cInvBuilding))

                outArray(outRow, 1) = "Invoices"
                outArray(outRow, 2) = buildingName
                outArray(outRow, 3) = invData(r, cInvReq)
                outArray(outRow, 4) = invData(r, cInvAcctNo)
                outArray(outRow, 5) = NzD(invData(r, cInvAmount))
                outArray(outRow, 6) = invData(r, cInvWO)
                outArray(outRow, 7) = invData(r, cInvType)
                outArray(outRow, 8) = CleanPOValue(invData(r, cInvPO))
                outArray(outRow, 9) = invData(r, cInvBrief)
                outArray(outRow, 10) = invData(r, cInvWorkDesc)
                outArray(outRow, 11) = invData(r, cInvComments)
                outArray(outRow, 12) = dtVal
                outArray(outRow, 13) = invData(r, cInvM)
                outArray(outRow, 14) = invData(r, cInvAcct)
                outArray(outRow, 15) = glVal
                outArray(outRow, 16) = LookupDict(commDict, CStr(glVal))
                outArray(outRow, 17) = FiscalYearLabel(dtVal)
                outArray(outRow, 18) = MonthStartDate(dtVal)
                outArray(outRow, 19) = buildingName

                outRow = outRow + 1
            End If
        Next r
    End If

    If lastLab >= 2 Then
        For r = 2 To lastLab
            buildingName = CStr(labData(r, cLabBuilding))

            If EXACT_EXISTING_LABOUR_DATE_LOGIC Then
                If lastInv >= 2 And r <= lastInv Then
                    dtVal = invData(r, cInvDate)
                Else
                    dtVal = labData(r, cLabDate)
                End If
            Else
                dtVal = labData(r, cLabDate)
            End If

            outArray(outRow, 1) = "Labour"
            outArray(outRow, 2) = buildingName
            outArray(outRow, 3) = Empty
            outArray(outRow, 4) = Empty
            outArray(outRow, 5) = NzD(labData(r, cLabCost))
            outArray(outRow, 6) = labData(r, cLabWO)
            outArray(outRow, 7) = labData(r, cLabType)
            outArray(outRow, 8) = CleanPOValue(labData(r, cLabPO))
            If cLabBrief > 0 Then outArray(outRow, 9) = labData(r, cLabBrief) Else outArray(outRow, 9) = Empty
            If cLabWorkDesc > 0 Then outArray(outRow, 10) = labData(r, cLabWorkDesc) Else outArray(outRow, 10) = Empty
            If cLabComments > 0 Then outArray(outRow, 11) = labData(r, cLabComments) Else outArray(outRow, 11) = Empty
            outArray(outRow, 12) = dtVal
            outArray(outRow, 13) = LookupDict(mDict, NormalizeText(buildingName))
            outArray(outRow, 14) = 1511
            outArray(outRow, 15) = 4150
            outArray(outRow, 16) = LookupDict(commDict, "4150")
            outArray(outRow, 17) = FiscalYearLabel(dtVal)
            outArray(outRow, 18) = MonthStartDate(dtVal)
            outArray(outRow, 19) = buildingName

            outRow = outRow + 1
        Next r
    End If

    If outRow > 1 Then
        wsData.Range("A2").Resize(outRow - 1, 19).Value = outArray
    End If

FormatSheet:
    FormatDataSheet wsData

End Sub

'=========================================================
' Step 3: Create real PivotTables
'=========================================================

Private Sub CreateNormalPivotTables( _
    ByVal wsData As Worksheet, _
    ByVal wsMonth As Worksheet, _
    ByVal wsPO As Worksheet, _
    ByVal reportFY As String)

    Dim wb As Workbook
    Dim srcRange As Range
    Dim lo As ListObject
    Dim pcMonth As PivotCache
    Dim pcPO As PivotCache
    Dim ptMonth As PivotTable
    Dim ptPO As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim srcAddress As String

    Set wb = wsData.Parent

    lastRow = LastUsedRow(wsData)
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    Set srcRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    On Error Resume Next
    Set lo = wsData.ListObjects("tblNonBaseCostData")
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = wsData.ListObjects.Add(xlSrcRange, srcRange, , xlYes)
        lo.Name = "tblNonBaseCostData"
    Else
        lo.Resize srcRange
    End If

    On Error Resume Next
    lo.TableStyle = "TableStyleMedium2"
    On Error GoTo 0

    wsMonth.Cells.Clear
    wsPO.Cells.Clear

    srcAddress = lo.Range.Address(ReferenceStyle:=xlR1C1, External:=True)

    Set pcMonth = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcAddress)
    Set pcPO = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcAddress)

    On Error Resume Next
    pcMonth.MissingItemsLimit = xlMissingItemsNone
    pcPO.MissingItemsLimit = xlMissingItemsNone
    On Error GoTo 0

    Set ptMonth = pcMonth.CreatePivotTable( _
        TableDestination:=wsMonth.Range("A1"), _
        TableName:="ptNonBaseCostByMonth")

    BuildMonthPivot ptMonth, reportFY

    Set ptPO = pcPO.CreatePivotTable( _
        TableDestination:=wsPO.Range("A1"), _
        TableName:="ptNonBaseCostByPO")

    BuildPOPivot ptPO, reportFY

End Sub

'=========================================================
' Pivot 1: Month / Commodity Pivot
'=========================================================

Private Sub BuildMonthPivot(ByVal pt As PivotTable, ByVal reportFY As String)

    Dim df As PivotField

    With pt

        .ManualUpdate = True

        On Error Resume Next
        .ClearAllFilters
        .RowAxisLayout xlCompactRow
        .RepeatAllLabels xlDoNotRepeatLabels
        .DisplayFieldCaptions = True
        .ShowDrillIndicators = True
        .HasAutoFormat = True
        .TableStyle2 = "PivotStyleMedium9"
        .PreserveFormatting = True
        On Error GoTo 0

        With .PivotFields("Nissan FY")
            .Orientation = xlPageField
            .Position = 1
            .ClearAllFilters
            On Error Resume Next
            .CurrentPage = reportFY
            On Error GoTo 0
        End With

        With .PivotFields("Building Filter")
            .Orientation = xlPageField
            .Position = 2
            .ClearAllFilters
        End With

        With .PivotFields("Commodity Category")
            .Orientation = xlRowField
            .Position = 1
        End With
        SetFieldSubtotal .PivotFields("Commodity Category"), True

        With .PivotFields("WO Type")
            .Orientation = xlRowField
            .Position = 2
        End With
        SetFieldSubtotal .PivotFields("WO Type"), True

        With .PivotFields("Invoice Month")
            .Orientation = xlRowField
            .Position = 3
            .AutoSort xlAscending, "Invoice Month" ' Added chronological sorting
            On Error Resume Next
            .NumberFormat = "mmm"
            On Error GoTo 0
        End With
        SetFieldSubtotal .PivotFields("Invoice Month"), True

        With .PivotFields("Building Name")
            .Orientation = xlRowField
            .Position = 4
        End With
        SetFieldSubtotal .PivotFields("Building Name"), True

        With .PivotFields("Work Order No.")
            .Orientation = xlRowField
            .Position = 5
        End With
        SetFieldSubtotal .PivotFields("Work Order No."), False

        Set df = .AddDataField(.PivotFields("Invoice Amount"), "Sum of Invoice Amount", xlSum)
        df.NumberFormat = "$#,##0.00"

        .ManualUpdate = False
        .RefreshTable

    End With

    FormatPivotLikeCurrent pt
    CollapsePivotToMonthLevel pt, True

End Sub

'=========================================================
' Pivot 2: Nissan PO Pivot
'=========================================================

Private Sub BuildPOPivot(ByVal pt As PivotTable, ByVal reportFY As String)

    Dim df As PivotField

    With pt

        .ManualUpdate = True

        On Error Resume Next
        .ClearAllFilters
        .RowAxisLayout xlCompactRow
        .RepeatAllLabels xlDoNotRepeatLabels
        .DisplayFieldCaptions = True
        .ShowDrillIndicators = True
        .HasAutoFormat = True
        .TableStyle2 = "PivotStyleMedium9"
        .PreserveFormatting = True
        On Error GoTo 0

        With .PivotFields("Nissan FY")
            .Orientation = xlPageField
            .Position = 1
            .ClearAllFilters
            On Error Resume Next
            .CurrentPage = reportFY
            On Error GoTo 0
        End With

        With .PivotFields("Customer PO #")
            .Orientation = xlRowField
            .Position = 1
        End With
        SetFieldSubtotal .PivotFields("Customer PO #"), True

        With .PivotFields("WO Type")
            .Orientation = xlRowField
            .Position = 2
        End With
        SetFieldSubtotal .PivotFields("WO Type"), True

        With .PivotFields("Invoice Month")
            .Orientation = xlRowField
            .Position = 3
            .AutoSort xlAscending, "Invoice Month" ' Added chronological sorting
            On Error Resume Next
            .NumberFormat = "mmm"
            On Error GoTo 0
        End With
        SetFieldSubtotal .PivotFields("Invoice Month"), True

        With .PivotFields("Building Name")
            .Orientation = xlRowField
            .Position = 4
        End With
        SetFieldSubtotal .PivotFields("Building Name"), True

        With .PivotFields("Work Order No.")
            .Orientation = xlRowField
            .Position = 5
        End With
        SetFieldSubtotal .PivotFields("Work Order No."), False

        Set df = .AddDataField(.PivotFields("Invoice Amount"), "Sum of Invoice Amount", xlSum)
        df.NumberFormat = "$#,##0.00"

        .ManualUpdate = False
        .RefreshTable

    End With

    FormatPivotLikeCurrent pt
    CollapsePivotToMonthLevel pt, True

End Sub

'=========================================================
' Pivot formatting helpers
'=========================================================

Private Sub FormatPivotLikeCurrent(ByVal pt As PivotTable)

    Dim ws As Worksheet
    Dim df As PivotField

    Set ws = pt.Parent

    On Error Resume Next

    With pt
        .ColumnGrand = True
        .RowGrand = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .NullString = ""
        .PreserveFormatting = True
        .ShowDrillIndicators = True
        .TableStyle2 = "PivotStyleMedium9"
        .CompactLayoutRowHeader = "Row Labels"
    End With

    For Each df In pt.DataFields
        df.NumberFormat = "$#,##0.00"
    Next df

    On Error GoTo 0

    With ws
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
    End With

End Sub

Private Sub CollapsePivotToMonthLevel(ByVal pt As PivotTable, Optional ByVal showBuildingsUnderMonths As Boolean = True)

    Dim pf As PivotField
    Dim pi As PivotItem

    On Error Resume Next

    Set pf = pt.RowFields(1)
    For Each pi In pf.PivotItems
        pi.ShowDetail = True
    Next pi

    Set pf = pt.RowFields(2)
    For Each pi In pf.PivotItems
        pi.ShowDetail = True
    Next pi

    Set pf = pt.RowFields(3)
    For Each pi In pf.PivotItems
        pi.ShowDetail = showBuildingsUnderMonths
    Next pi

    If showBuildingsUnderMonths Then
        Set pf = pt.RowFields(4)
        For Each pi In pf.PivotItems
            pi.ShowDetail = False
        Next pi
    End If

    On Error GoTo 0

End Sub

Private Sub SetFieldSubtotal(ByVal pf As PivotField, ByVal showAutomaticSubtotal As Boolean)

    Dim i As Long

    On Error Resume Next

    For i = 1 To 12
        pf.Subtotals(i) = False
    Next i

    If showAutomaticSubtotal Then
        pf.Subtotals(1) = True
    End If

    On Error GoTo 0

End Sub

'=========================================================
' Lookup builders
'=========================================================

Private Function BuildMLookup(ByVal wsM As Worksheet) As Object

    Dim d As Object
    Dim cM As Long
    Dim cSite As Long
    Dim r As Long
    Dim lastRow As Long
    Dim siteName As String
    Dim srcData As Variant

    Set d = CreateObject("Scripting.Dictionary")

    cM = HeaderColAny(wsM, Array("JDE M#", "M#"), False)
    If cM = 0 Then cM = 1

    cSite = HeaderColAny(wsM, Array("Site Name", "Building Name"), False)
    If cSite = 0 Then cSite = 2

    lastRow = LastUsedRow(wsM)
    If lastRow >= 2 Then
        srcData = wsM.Range(wsM.Cells(1, 1), wsM.Cells(lastRow, wsM.Cells(1, wsM.Columns.Count).End(xlToLeft).Column)).Value
        For r = 2 To lastRow
            siteName = NormalizeText(srcData(r, cSite))
            If Len(siteName) > 0 Then
                d(siteName) = srcData(r, cM)
            End If
        Next r
    End If

    Set BuildMLookup = d

End Function

Private Function BuildCommodityLookup(ByVal wsComm As Worksheet) As Object

    Dim d As Object
    Dim r As Long
    Dim lastRow As Long
    Dim glKey As String
    Dim srcData As Variant

    Set d = CreateObject("Scripting.Dictionary")

    lastRow = LastUsedRow(wsComm)
    If lastRow >= 2 Then
        srcData = wsComm.Range(wsComm.Cells(1, 1), wsComm.Cells(lastRow, 2)).Value
        For r = 2 To lastRow
            glKey = Trim$(CStr(srcData(r, 1)))
            If Len(glKey) > 0 Then
                d(glKey) = srcData(r, 2)
            End If
        Next r
    End If

    Set BuildCommodityLookup = d

End Function

'=========================================================
' Data formatting
'=========================================================

Private Sub FormatDataSheet(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = LastUsedRow(ws)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    With ws

        .Rows(1).Font.Bold = True

        If .AutoFilterMode Then .AutoFilterMode = False
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).AutoFilter

        .Columns("E:E").NumberFormat = "0.00"
        .Columns("H:H").NumberFormat = "@"
        .Columns("L:L").NumberFormat = "m/d/yyyy h:mm"
        .Columns("N:O").NumberFormat = "0"
        .Columns("R:R").NumberFormat = "mmm"
        .Columns("A:S").EntireColumn.AutoFit

    End With

End Sub

'=========================================================
' Generic utilities
'=========================================================

Private Function FindSheetLike(ByVal wb As Workbook, ByVal includeWords As Variant, ByVal excludeWords As Variant) As Worksheet

    Dim ws As Worksheet
    Dim nm As String
    Dim i As Long
    Dim ok As Boolean

    For Each ws In wb.Worksheets

        nm = LCase$(ws.Name)
        ok = True

        For i = LBound(includeWords) To UBound(includeWords)
            If InStr(1, nm, LCase$(CStr(includeWords(i))), vbTextCompare) = 0 Then
                ok = False
                Exit For
            End If
        Next i

        If ok Then
            If IsArrayAllocated(excludeWords) Then
                For i = LBound(excludeWords) To UBound(excludeWords)
                    If Len(CStr(excludeWords(i))) > 0 Then
                        If InStr(1, nm, LCase$(CStr(excludeWords(i))), vbTextCompare) > 0 Then
                            ok = False
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If

        If ok Then
            Set FindSheetLike = ws
            Exit Function
        End If

    Next ws

End Function

Private Function RecreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Delete
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName

    Set RecreateSheet = ws

End Function

Private Function HeaderColAny(ByVal ws As Worksheet, ByVal possibleHeaders As Variant, Optional ByVal required As Boolean = True) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim i As Long
    Dim cellText As String
    Dim headerText As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        cellText = CleanHeader(ws.Cells(1, c).Value)

        For i = LBound(possibleHeaders) To UBound(possibleHeaders)
            headerText = CleanHeader(CStr(possibleHeaders(i)))

            If cellText = headerText Then
                HeaderColAny = c
                Exit Function
            End If

        Next i

    Next c

    If required Then
        Err.Raise vbObjectError + 200, , "Missing required header on sheet '" & ws.Name & "': " & JoinVariant(possibleHeaders, " / ")
    End If

End Function

Private Function EnsureHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long

    Dim c As Long
    Dim lastCol As Long

    c = HeaderColAny(ws, Array(headerName), False)

    If c > 0 Then
        EnsureHeader = c
    Else
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        EnsureHeader = lastCol + 1
        ws.Cells(1, EnsureHeader).Value = headerName
    End If

End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = f.Row
    End If

End Function

Private Function CleanHeader(ByVal v As Variant) As String

    CleanHeader = LCase$(Trim$(Replace(CStr(v), Chr$(160), " ")))

End Function

Private Function NormalizeText(ByVal v As Variant) As String

    Dim s As String

    s = CStr(v)
    s = Replace(s, Chr$(160), " ")
    s = Application.WorksheetFunction.Trim(s)

    NormalizeText = UCase$(s)

End Function

Private Function MakeRateKey(ByVal contactVal As Variant, ByVal rateTypeVal As Variant) As String

    MakeRateKey = NormalizeText(contactVal) & "|" & NormalizeText(rateTypeVal)

End Function

Private Function LookupDict(ByVal d As Object, ByVal key As String) As Variant

    If d.Exists(key) Then
        LookupDict = d(key)
    Else
        LookupDict = vbNullString
    End If

End Function

Private Function FiscalYearLabel(ByVal dtVal As Variant) As String

    Dim d As Date
    Dim fy As Long

    If Not IsDate(dtVal) Then
        FiscalYearLabel = vbNullString
        Exit Function
    End If

    d = CDate(dtVal)

    If Month(d) >= 4 Then
        fy = Year(d)
    Else
        fy = Year(d) - 1
    End If

    FiscalYearLabel = "FY" & Right$(CStr(fy), 2)

End Function

Private Function MonthStartDate(ByVal dtVal As Variant) As Variant

    Dim d As Date

    If Not IsDate(dtVal) Then
        MonthStartDate = vbNullString
        Exit Function
    End If

    d = CDate(dtVal)
    MonthStartDate = DateSerial(Year(d), Month(d), 1)

End Function

Private Function CleanPOValue(ByVal v As Variant) As Variant

    If IsError(v) Then
        CleanPOValue = vbNullString
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        CleanPOValue = vbNullString
    Else
        CleanPOValue = CStr(v)
    End If

End Function

Private Function NzD(ByVal v As Variant) As Double

    If IsError(v) Then
        NzD = 0#
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        NzD = 0#
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0#
    End If

End Function

Private Function JoinVariant(ByVal v As Variant, ByVal delimiter As String) As String

    Dim i As Long
    Dim s As String

    For i = LBound(v) To UBound(v)
        If Len(s) > 0 Then s = s & delimiter
        s = s & CStr(v(i))
    Next i

    JoinVariant = s

End Function

Private Function IsArrayAllocated(ByVal v As Variant) As Boolean

    On Error GoTo NotAllocated

    If IsArray(v) Then
        Dim lb As Long
        Dim ub As Long

        lb = LBound(v)
        ub = UBound(v)

        IsArrayAllocated = True
    End If

    Exit Function

NotAllocated:
    IsArrayAllocated = False

End Function

Private Function GetAvailableClients(ByVal folderPath As String) As Object
    Dim dict As Object
    Dim fileName As String
    Dim startPos As Long
    Dim endPos As Long
    Dim clientName As String

    Set dict = CreateObject("Scripting.Dictionary")
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Scan for any Ainsworth project cost files, allowing any text around brackets and at the end
    fileName = Dir(folderPath & "*[*]*Project*Cost*.csv")

    Do While fileName <> ""
        startPos = InStr(1, fileName, "[")
        endPos = InStr(startPos, fileName, "]")

        If startPos > 0 And endPos > startPos Then
            clientName = Mid$(fileName, startPos + 1, endPos - startPos - 1)
            If Not dict.Exists(clientName) Then
                dict.Add clientName, clientName
            End If
        End If
        fileName = Dir()
    Loop

    Set GetAvailableClients = dict
End Function

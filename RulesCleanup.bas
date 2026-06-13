Attribute VB_Name = "RulesCleanup"

Option Explicit

Public Sub OptimizeAndCleanRules()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim key As String
    Dim dict As Object
    Dim rowsToDelete As Range
    
    Set ws = EnsureRulesSheet()
    Set dict = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    ' =========================================================================
    ' 1. SYSTEM-WIDE FILTERS
    ' =========================================================================
    ' strip_hash: Removes words starting with # (e.g. #3, #4)
    UpsertRule ws, "strip_hash", "", "TRUE"
    ' strip_alnum: Removes alphanumeric tags (e.g. B32, P-101)
    UpsertRule ws, "strip_alnum", "", "TRUE"

    ' =========================================================================
    ' 2. PRIMARY ASSET HIERARCHY (The "Anchor" Logic)
    ' =========================================================================
    ' Defining these as 'core' ensures the matcher ONLY looks at items
    ' containing these words if they appear in your input.
    UpsertRule ws, "core", "pump", "", "Anchor Asset"
    UpsertRule ws, "core", "vfd", "", "Anchor Asset"
    UpsertRule ws, "core", "tank", "", "Anchor Asset"
    UpsertRule ws, "core", "boiler", "", "Anchor Asset"
    UpsertRule ws, "core", "chiller", "", "Anchor Asset"
    UpsertRule ws, "core", "fan", "", "Anchor Asset"

    ' =========================================================================
    ' 3. LOGICAL MODIFIERS (The "Filtering" Logic)
    ' =========================================================================
    ' These 'boost' rules add scoring weight to specific pairings.
    ' Weight values (0.5 - 2.0) help differentiate the best match.
    UpsertRule ws, "boost", "glycol", "", "1.5"
    UpsertRule ws, "boost", "condensing", "", "1.2"
    UpsertRule ws, "boost", "hydronic", "", "1.0"
    UpsertRule ws, "boost", "domestic", "", "1.0"

    ' =========================================================================
    ' 4. CONTEXT CLEANUP (Standardizing & Noise Reduction)
    ' =========================================================================
    UpsertRule ws, "alias", "w/", "with"
    UpsertRule ws, "alias", "ahu", "air handling unit"
    UpsertRule ws, "alias", "blr", "boiler"
    
    ' Strip context-heavy words that create scoring noise in mechanical names
    UpsertRule ws, "alias", "mech", "", "Noise Reduction"
    UpsertRule ws, "alias", "side", "", "Noise Reduction"
    UpsertRule ws, "alias", "system", "", "Noise Reduction"

    ' =========================================================================
    ' 5. DUPLICATE CLEANUP
    ' =========================================================================
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For r = lastRow To 2 Step -1
        Dim ruleType As String, inputPhrase As String
        ruleType = LCase(Trim(ws.Cells(r, "A").Value))
        inputPhrase = LCase(Trim(ws.Cells(r, "B").Value))
        key = ruleType & "|" & inputPhrase
        
        If dict.Exists(key) Then
            If rowsToDelete Is Nothing Then Set rowsToDelete = ws.Rows(r) Else Set rowsToDelete = Union(rowsToDelete, ws.Rows(r))
        Else
            dict(key) = True
        End If
    Next r
    If Not rowsToDelete Is Nothing Then rowsToDelete.Delete

    ws.Columns("A:D").AutoFit
    Application.ScreenUpdating = True
    MsgBox "Rules Sheet optimized with Asset Hierarchy logic!"
End Sub

Private Function EnsureRulesSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Rules Sheet")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Rules Sheet"
        ws.Range("A1:D1").Value = Array("Type", "InputPhrase", "OutputData", "Weight/Notes")
    End If
    Set EnsureRulesSheet = ws
End Function

Private Sub UpsertRule(ws As Worksheet, ruleType As String, phrase As String, output As String, Optional weight As Variant = "")
    Dim lastRow As Long, r As Long
    Dim cleanType As String, cleanPhrase As String
    cleanType = LCase(Application.Trim(ruleType))
    cleanPhrase = LCase(Application.Trim(phrase))
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If LCase(Trim(ws.Cells(r, "A").Value)) = cleanType And LCase(Trim(ws.Cells(r, "B").Value)) = cleanPhrase Then
            ws.Cells(r, "C").Value = output
            If weight <> "" Then ws.Cells(r, "D").Value = weight
            Exit Sub
        End If
    Next r
    ws.Cells(lastRow + 1, "A").Value = cleanType
    ws.Cells(lastRow + 1, "B").Value = cleanPhrase
    ws.Cells(lastRow + 1, "C").Value = output
    ws.Cells(lastRow + 1, "D").Value = weight
End Sub

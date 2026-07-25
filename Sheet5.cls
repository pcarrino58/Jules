VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


' ============================================================
' MAIN EVENT LISTENER (Worksheet_Change) - PYTHON OPTIMIZED
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim rngMatchInput As Range, rngMatchResult As Range
    Dim rngSchedFreq As Range, rngSchedDate As Range
    Dim hitMatchA As Range, hitMatchB As Range
    Dim hitSched As Range
    Dim r As Range, rowNum As Long
    Dim cell As Range
    Dim isSearching As Boolean
    isSearching = False

    Set rngMatchInput = Me.Range("A2:A" & Me.Rows.count)
    Set rngMatchResult = Me.Range("B2:B" & Me.Rows.count)
    Set rngSchedFreq = Me.Range("T2:T" & Me.Rows.count)
    Set rngSchedDate = Me.Range("X2:X" & Me.Rows.count)

    Set hitMatchA = Intersect(Target, rngMatchInput)
    Set hitMatchB = Intersect(Target, rngMatchResult)
    Set hitSched = Intersect(Target, Union(rngSchedFreq, rngSchedDate))

    If hitMatchA Is Nothing And hitMatchB Is Nothing And hitSched Is Nothing Then GoTo CleanExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ============================================================
    ' DELETION / CLEAR BOARD LOGIC
    ' ============================================================
    If Target.Cells.count > 20 And Not hitMatchA Is Nothing Then
        Dim actualHitA As Range
        Set actualHitA = Intersect(hitMatchA, Me.UsedRange)
        If actualHitA Is Nothing Then GoTo CleanExit

        If Application.WorksheetFunction.CountA(actualHitA) = 0 Then
            actualHitA.offset(0, 1).ClearContents
            actualHitA.offset(0, 2).ClearContents
            actualHitA.offset(0, 3).ClearContents
            actualHitA.offset(0, 1).Resize(, 2).Interior.Color = RGB(252, 186, 186)

            actualHitA.offset(0, 22).ClearContents
            actualHitA.offset(0, 22).Interior.Color = RGB(201, 231, 167)
            actualHitA.offset(0, 23).ClearContents
            actualHitA.offset(0, 24).Resize(, 12).ClearContents
            actualHitA.offset(0, 36).ClearContents

            On Error Resume Next
            actualHitA.offset(0, 1).Validation.Delete
            Call Matcher.CleanupOrphanedPickLists

            Dim wsStaging As Worksheet
            Set wsStaging = ThisWorkbook.Worksheets("AI Staging")
            If Not wsStaging Is Nothing Then
                Dim lastS As Long
                lastS = wsStaging.Cells(wsStaging.Rows.count, "A").End(xlUp).Row
                If lastS >= 2 Then
                    wsStaging.Range("A2:C" & lastS).ClearContents
                    wsStaging.Range("A2:C" & lastS).Interior.ColorIndex = xlNone
                End If
            End If

            On Error GoTo ErrorHandler
        End If
        
        EnsureNoGoodMatchHighlight
        GoTo CleanExit
    End If

    ' ============================================================
    ' INPUT LOGIC (Column A)
    ' ============================================================
    If Not hitMatchA Is Nothing Then
        Dim rngBlank As Range, c As Range
        
        For Each c In hitMatchA.Cells
            If Len(Trim$(CStr(c.Value2))) = 0 Then
                If rngBlank Is Nothing Then Set rngBlank = c Else Set rngBlank = Union(rngBlank, c)
            End If
        Next c

        If Not rngBlank Is Nothing Then
            For Each r In rngBlank.Cells
                rowNum = r.Row
                Me.Cells(rowNum, "B").ClearContents
                Me.Cells(rowNum, "C").ClearContents
                Me.Range("B" & rowNum & ":C" & rowNum).Interior.Color = RGB(252, 186, 186)
                
                On Error Resume Next
                Me.Cells(rowNum, "B").Validation.Delete
                ThisWorkbook.names("PickRow_" & rowNum).Delete
                On Error GoTo ErrorHandler
            Next r
        End If
        
        EnsureNoGoodMatchHighlight
    End If

    ' ============================================================
    ' DROPDOWN RESOLUTION & AI LEARNING LOGIC (Column B)
    ' ============================================================
    If Not hitMatchB Is Nothing Then
        If Target.CountLarge = 1 Then
            Dim rowKey As Long: rowKey = Target.Row
            Dim valToLearn As String: valToLearn = Matcher.CleanPastedOutput(Trim$(CStr(Target.Value2)))
            
            If Len(valToLearn) > 0 And valToLearn <> "Reject AI Guess" Then
                
                ' --- DYNAMIC SEARCH-ON-ENTER (Strict Exact Match) ---
                Dim isValid As Boolean
                isValid = False
                
                ' 1. Absolute strict test: Does this exact phrase exist in the Uniformat lookup?
                On Error Resume Next
                isValid = Matcher.IsValidOutputList(valToLearn)
                On Error GoTo ErrorHandler
                
                ' 2. If it is NOT a standalone asset, did they explicitly click an exact dropdown item?
                If Not isValid Then
                    Dim vList As String
                    On Error Resume Next
                    vList = Me.Cells(rowKey, "B").Validation.Formula1
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                    If Len(vList) > 0 Then
                        If Left(vList, 1) = "=" Then
                            ' It's a named range list
                            Dim rList As Variant
                            On Error Resume Next
                            rList = Evaluate(vList)
                            Err.Clear
                            On Error GoTo ErrorHandler
                            
                            If Not IsEmpty(rList) Then
                                Dim rItem As Variant
                                For Each rItem In rList
                                    If Not IsError(rItem) Then
                                        If LCase(Trim$(CStr(rItem))) = LCase(valToLearn) Then
                                            isValid = True
                                            Exit For
                                        End If
                                    End If
                                Next rItem
                            End If
                        Else
                            ' It's a comma-separated list
                            Dim splitVals() As String
                            Dim vItem As Variant
                            splitVals = Split(vList, ",")
                            For Each vItem In splitVals
                                If LCase(Trim$(CStr(vItem))) = LCase(valToLearn) Then
                                    isValid = True
                                    Exit For
                                End If
                            Next vItem
                        End If
                    End If
                End If
                
                ' 3. The Search Engine Trigger
                If isValid = False Then
                    If Len(valToLearn) >= 3 And InStr(1, valToLearn, "Click to", vbTextCompare) = 0 Then
                        On Error Resume Next
                        
                        Dim searchSuccess As Boolean
                        searchSuccess = Matcher.BuildSearchDropdown(Me, rowKey, valToLearn)
                        
                        If searchSuccess Then
                            isSearching = True
                            
                            ' Force UI to wake up so SendKeys registers visually
                            Application.ScreenUpdating = True
                            Me.Cells(rowKey, "B").Select
                            Application.SendKeys "%{DOWN}"
                        Else
                            Me.Cells(rowKey, "B").Value = "No matches found"
                        End If
                        
                        On Error GoTo ErrorHandler
                    End If
                    GoTo SkipLearning
                End If
                ' --------------------------------------------
                
                ' 1. Strip the yellow paint and restore Pink
                Me.Range("B" & rowKey & ":C" & rowKey).Interior.Color = RGB(252, 186, 186)
                
                ' 2. Delete the dropdown so it becomes a static locked answer
                On Error Resume Next
                Target.Validation.Delete
                ThisWorkbook.names("PickRow_" & rowKey).Delete
                Err.Clear
                On Error GoTo ErrorHandler
                
                ' 3. Update Column C's Status
                Me.Cells(rowKey, "C").Value2 = "USER_LEARNED"
                
' 4. Delegate Learning, Python API, and Auto-Propagation to the Master Engine
                Call Matcher.RecordLearnedOverride(rowKey)
                
            End If
        End If
    End If

SkipLearning:

    ' ============================================================
    ' SCHEDULE LOGIC
    ' ============================================================
    If Not hitSched Is Nothing Then
        Dim schedRows As Object: Set schedRows = CreateObject("Scripting.Dictionary")
        For Each r In hitSched.Cells: schedRows(r.Row) = True: Next r
        Dim schedKey As Variant
        For Each schedKey In schedRows.keys: UpdateScheduleForRow CLng(schedKey): Next schedKey
    End If

CleanExit:
    On Error Resume Next
    ' Only run the pivot refresh if we are NOT currently trying to open a search dropdown
    If Target.Cells.count < 50 And isSearching = False Then Call PivotBuilder_Refresh
    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isRunning = False
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isRunning = False
    MsgBox "Macro Crash Detected in Worksheet_Change!" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "VBA Error Tracker"
End Sub
' ============================================================
' HELPER: THE VBA GENERALIZER
' ============================================================
Function get_generalized_signature(text As String) As String
    ' A perfectly mirrored VBA clone of your Python Generalizer
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.pattern = "\s*#?\s*\d+[a-zA-Z]*$"
    End If
    
    Dim prevText As String
    text = Trim(text)
    
    ' Aggressively strip spaces, #, and trailing alphanumeric codes (e.g., " #1", "2B")
    text = Trim(regEx.Replace(text, ""))
    
    get_generalized_signature = LCase(text)
End Function
Private Sub UpdateScheduleForRow(ByVal r As Long)
    Dim ws As Worksheet: Set ws = Me
    Dim freqStr As String, startDate As Variant, startMonth As Long
    Dim schedule(1 To 12) As String, m As Long, idx As Long, off As Variant
    Dim hasM As Boolean, hasQ As Boolean, hasS As Boolean, hasA As Boolean

    freqStr = CStr(ws.Cells(r, "T").Value)
    startDate = ws.Cells(r, "X").Value
    ws.Range(ws.Cells(r, 25), ws.Cells(r, 36)).ClearContents

    If Len(Trim$(freqStr)) = 0 Or Not IsDateValid(startDate) Then Exit Sub
    If IsDate(startDate) Then startMonth = Month(CDate(startDate)) Else startMonth = CLng(startDate)

    ParseCodes freqStr, hasM, hasQ, hasS, hasA
    If hasA Then schedule(startMonth) = "A"
    If hasS Then
        If schedule(startMonth) = "" Then schedule(startMonth) = "S"
        idx = MonthFromOffset(startMonth, 6): If schedule(idx) = "" Then schedule(idx) = "S"
    End If
    If hasQ Then
        For Each off In Array(0, 3, 6, 9)
            idx = MonthFromOffset(startMonth, CLng(off)): If schedule(idx) = "" Then schedule(idx) = "Q"
        Next off
    End If
    If hasM Then
        For m = 1 To 12
            If schedule(m) = "" Then schedule(m) = "M"
        Next m
    End If
    For m = 1 To 12
        If schedule(m) <> "" Then ws.Cells(r, 24 + m).Value = schedule(m)
    Next m
End Sub

Private Function MonthFromOffset(startMonth As Long, offset As Long) As Long
    Dim m As Long: m = startMonth + offset
    Do While m > 12: m = m - 12: Loop
    Do While m < 1: m = m + 12: Loop
    MonthFromOffset = m
End Function

Private Sub ParseCodes(ByVal freqStr As String, ByRef hasM As Boolean, ByRef hasQ As Boolean, ByRef hasS As Boolean, ByRef hasA As Boolean)
    Dim cleaned As String, i As Long, tokens() As String, t As Variant
    hasM = False: hasQ = False: hasS = False: hasA = False
    cleaned = UCase$(freqStr)
    cleaned = Replace(cleaned, ",", " "): cleaned = Replace(cleaned, "/", " "): cleaned = Replace(cleaned, "-", " ")
    For i = 1 To Len(cleaned)
        If Mid$(cleaned, i, 1) < "A" Or Mid$(cleaned, i, 1) > "Z" Then Mid$(cleaned, i, 1) = " "
    Next i
    cleaned = Trim$(cleaned)
    Do While InStr(cleaned, "  ") > 0: cleaned = Replace(cleaned, "  ", " "): Loop
    If Len(cleaned) = 0 Then Exit Sub
    tokens = Split(cleaned, " ")
    For Each t In tokens
        Select Case t: Case "M": hasM = True: Case "Q": hasQ = True: Case "S": hasS = True: Case "A": hasA = True: End Select
    Next t
End Sub

Private Function IsDateValid(val As Variant) As Boolean
    If IsDate(val) Then IsDateValid = True: Exit Function
    If IsNumeric(val) Then If val >= 1 And val <= 12 Then IsDateValid = True: Exit Function
    IsDateValid = False
End Function

Private Sub EnsureNoGoodMatchHighlight()
    Dim lastRow As Long, applies As Range, i As Long

    lastRow = Me.Cells(Me.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    Set applies = Me.Range("B2:C" & lastRow)

    ' Remove legacy rules without changing worksheet focus
    For i = Me.Columns("B:C").FormatConditions.count To 1 Step -1
        If Me.Columns("B:C").FormatConditions(i).Type = 2 Then
            If InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$B", vbTextCompare) > 0 Or _
               InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$C", vbTextCompare) > 0 Then
                Me.Columns("B:C").FormatConditions(i).Delete
            End If
        End If
    Next i

    Dim formulaOriginal1 As String, formulaOriginal2 As String, formulaNew As String
    Dim f1 As String, f2 As String, f3 As String

    formulaOriginal1 = "=OR($B2=""No good match"", $B2=""No matching words"", $B2=""Click to Review Guess..."", $C2=""Requires AI"", $C2=""Check DB"")"
    formulaOriginal2 = "=OR($C2=""AI Failed to Resolve"", AND(LEFT($C2,14)=""Multiple Rules"", ISERR(SEARCH(""Resolved"",$C2))), ISNUMBER(SEARCH(""Review AI Guess"",$C2)))"
    formulaNew = "=OR($B2=""Subtype missing input"", $B2=""Missing input phrase"", ISNUMBER(SEARCH(""Click dropdown to view"", $B2)), $C2=""Requires human"", $C2=""Search Results"", $C2=""Invalid Entry"")"

    ' Apply the formula relative to the ActiveCell by converting A1 -> R1C1 (absolute to row 2) -> A1 (relative to ActiveCell)
    Dim active As Range
    Set active = ActiveCell
    If active Is Nothing Then Set active = applies.Cells(1, 1)

    f1 = Application.ConvertFormula(Application.ConvertFormula(formulaOriginal1, xlA1, xlR1C1, , applies.Cells(1, 1)), xlR1C1, xlA1, , active)
    f2 = Application.ConvertFormula(Application.ConvertFormula(formulaOriginal2, xlA1, xlR1C1, , applies.Cells(1, 1)), xlR1C1, xlA1, , active)
    f3 = Application.ConvertFormula(Application.ConvertFormula(formulaNew, xlA1, xlR1C1, , applies.Cells(1, 1)), xlR1C1, xlA1, , active)

    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=f1)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With

    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=f2)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With

    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=f3)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With
End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Intersect(Target, Me.Range("B:C")) Is Nothing Or Target.Row < 2 Or Target.Cells.count > 1 Then Exit Sub

    Dim inputPhrase As String: inputPhrase = Trim(Me.Cells(Target.Row, "A").Value2)
    If inputPhrase = "" Then Exit Sub

    ' 1. Stop Excel's default double-click jump
    Cancel = True

    ' 2. THE SHIELD: Turn off events so typing "Thinking" doesn't trigger the Pivot Refresh!
    Application.EnableEvents = False
    Me.Cells(Target.Row, "C").Value = "AI Thinking..."
    Application.EnableEvents = True ' Turn it right back on

    DoEvents
    Call Matcher.RunAISniper(Me, Target.Row, inputPhrase)
End Sub




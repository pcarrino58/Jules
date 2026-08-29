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
    If Target.Cells.count > 1 And Not hitMatchA Is Nothing Then
        Dim actualHitA As Range
        Set actualHitA = Intersect(hitMatchA, Me.UsedRange)
        If Not actualHitA Is Nothing Then
    If Application.WorksheetFunction.CountA(actualHitA) = 0 Then
                actualHitA.offset(0, 1).ClearContents ' Clears Column B
                actualHitA.offset(0, 2).ClearContents ' Clears Column C
                actualHitA.offset(0, 3).ClearContents ' Clears Column D
                actualHitA.offset(0, 4).ClearContents ' Clears Column E (PM Trigger)
                actualHitA.offset(0, 5).ClearContents ' Clears Column F (PM Feasibility)
                ' --- UPDATED COLOR REVERT LOGIC ---
                actualHitA.Interior.Color = RGB(201, 231, 167) ' Reverts Column A to Green
                actualHitA.offset(0, 1).Resize(, 2).Interior.Color = RGB(252, 186, 186) ' Reverts B:C to Pink
                ' ----------------------------------

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

                Dim wsScrub As Worksheet
                Set wsScrub = ThisWorkbook.Worksheets("Scrubbed Inputs")
                If Not wsScrub Is Nothing Then
                    wsScrub.Cells.Clear
                End If

                On Error GoTo ErrorHandler
            End If
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
                Me.Cells(rowNum, "D").ClearContents
                Me.Cells(rowNum, "E").ClearContents
                Me.Cells(rowNum, "F").ClearContents
                
                Me.Range("A" & rowNum).Interior.Color = RGB(201, 231, 167)
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
        Dim updateCell As Range
        Dim isMultiPaste As Boolean
        isMultiPaste = (hitMatchB.Cells.count > 1) ' Detects if you pasted multiple cells
        
        ' Loop through every pasted cell instead of restricting to Target.CountLarge = 1
        For Each updateCell In hitMatchB.Cells
            Dim rowKey As Long: rowKey = updateCell.Row
            Dim valToLearn As String: valToLearn = Matcher.CleanPastedOutput(Trim$(CStr(updateCell.Value2)))
            
            If Len(valToLearn) > 0 And valToLearn <> "Reject AI Guess" Then
                Dim isValid As Boolean
                isValid = False
                
                On Error Resume Next
                isValid = Matcher.IsValidOutputList(valToLearn)
                On Error GoTo ErrorHandler
                
                If Not isValid Then
                    Dim vList As String
                    On Error Resume Next
                    vList = Me.Cells(rowKey, "B").Validation.Formula1
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                    If Len(vList) > 0 Then
                        If Left(vList, 1) = "=" Then
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
                
                If isValid = False Then
                    If Len(valToLearn) >= 3 And InStr(1, valToLearn, "Click to", vbTextCompare) = 0 Then
                        On Error Resume Next
                        Dim searchSuccess As Boolean
                        searchSuccess = Matcher.BuildSearchDropdown(Me, rowKey, valToLearn)
                        
                        If searchSuccess Then
                            isSearching = True
                            Application.ScreenUpdating = True
                            Me.Cells(rowKey, "B").Select
                            Application.SendKeys "%{DOWN}"
                        Else
                            Me.Cells(rowKey, "B").Value = "No matches found"
                        End If
                        On Error GoTo ErrorHandler
                    End If
                    GoTo SkipLearningCell
                End If
                
                ' --- COLOR REVERT LOGIC ---
                Me.Range("A" & rowKey).Interior.Color = RGB(201, 231, 167)
                Me.Range("B" & rowKey & ":C" & rowKey).Interior.Color = RGB(252, 186, 186)
                
                On Error Resume Next
                updateCell.Validation.Delete
                ThisWorkbook.names("PickRow_" & rowKey).Delete
                Err.Clear
                On Error GoTo ErrorHandler
                
                ' >>> FIXED ORDER: Let the Matcher read Column C BEFORE we overwrite it! <<<
                Call Matcher.RecordLearnedOverride(rowKey, isMultiPaste)
                
                ' Now it is safe to overwrite Column C
                Me.Cells(rowKey, "C").Value2 = "USER_LEARNED"
            End If
SkipLearningCell:
        Next updateCell
        
' Save the hard drive file exactly once at the end of the paste
        If isMultiPaste Then
            Call Matcher.FlushBatchLearningsToPython
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
    If isSearching = False Then
        ' Only recalculate PM rules if it was a single cell edit.
        ' If pasting 100 rows, skip this so it doesn't freeze the screen.
        If Not isMultiPaste Then
            Call Matcher.CalculatePMRules(Me)
        End If
        
        ' REMOVED: Call PivotBuilder_Refresh
        ' Pivot tables should NOT be refreshed on every keystroke.
        ' Move this to a "Refresh Dashboard" button on your Pivot sheet.
    End If
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
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.pattern = "\s*#?\s*\d+[a-zA-Z]*$"
    End If
    
    Dim prevText As String
    text = Trim(text)
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

    For i = Me.Columns("B:C").FormatConditions.count To 1 Step -1
        If Me.Columns("B:C").FormatConditions(i).Type = 2 Then
            If InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$B", vbTextCompare) > 0 Or _
               InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$C", vbTextCompare) > 0 Then
                Me.Columns("B:C").FormatConditions(i).Delete
            End If
        End If
    Next i

    Dim formulaOriginal1 As String, formulaOriginal2 As String, formulaNew As String
    Dim formulaGuess As String

    ' Formulas MUST be written relative to Row 2, which is the top of the 'applies' range.
    formulaOriginal1 = "=OR($B2=""No good match"", $B2=""No matching words"", $B2=""Click to Review Guess..."", $C2=""Requires AI"", $C2=""Check DB"")"
    
    ' Revert original formula back to standard errors
    formulaOriginal2 = "=OR($C2=""AI Failed to Resolve"", AND(LEFT($C2,14)=""Multiple Rules"", ISERR(SEARCH(""Resolved"",$C2))), ISNUMBER(SEARCH(""Review AI Guess"",$C2)))"
    
    formulaNew = "=OR($B2=""Subtype missing input"", $B2=""Missing input phrase"", ISNUMBER(SEARCH(""Click dropdown to view"", $B2)), $C2=""Requires human"", $C2=""Search Results"", $C2=""Invalid Entry"")"
    
    ' NEW: Dedicated formula for AI guesses
    formulaGuess = "=OR($C2=""AI_BEST_GUESS"", $C2=""EXPERT_RULE (Auto-Resolved)"")"

    ' Apply standard yellow rules...
    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaOriginal1)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With

    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaOriginal2)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With

    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaNew)
        .Interior.Color = 65535
        .StopIfTrue = False
    End With

    ' Apply new #FFC000 rule...
    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaGuess)
        .Interior.Color = RGB(255, 192, 0)
        .StopIfTrue = False
    End With
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Intersect(Target, Me.Range("B:C")) Is Nothing Or Target.Row < 2 Or Target.Cells.count > 1 Then Exit Sub

    Dim inputPhrase As String: inputPhrase = Trim(Me.Cells(Target.Row, "A").Value2)
    If inputPhrase = "" Then Exit Sub

    Cancel = True
    Application.EnableEvents = False
    Me.Cells(Target.Row, "C").Value = "AI Thinking..."
    Application.EnableEvents = True

    DoEvents
    Call Matcher.RunAISniper(Me, Target.Row, inputPhrase)
End Sub


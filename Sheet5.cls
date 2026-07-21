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
    ' ?? 1. SILENCER REMOVED: Now routes to a real ErrorHandler
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

    Set rngMatchInput = Me.Range("A2:A" & Me.Rows.count)
    Set rngMatchResult = Me.Range("B2:B" & Me.Rows.count)
    Set rngSchedFreq = Me.Range("T2:T" & Me.Rows.count)
    Set rngSchedDate = Me.Range("X2:X" & Me.Rows.count)

    Set hitMatchA = Intersect(Target, rngMatchInput)
    Set hitMatchB = Intersect(Target, rngMatchResult)
    Set hitSched = Intersect(Target, Union(rngSchedFreq, rngSchedDate))

    If hitMatchA Is Nothing And hitMatchB Is Nothing And hitSched Is Nothing Then GoTo CleanExit

    ' ============================================================
    ' ??? THE "DO NOT DISTURB" SHIELD (Massive Performance Boost)
    ' ============================================================
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ============================================================
    ' DELETION / CLEAR BOARD LOGIC (Bulk Deletes)
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
    ' INPUT LOGIC (Column A) - STRIPPED DOWN FOR DELETIONS ONLY
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
            Dim valToLearn As String: valToLearn = Trim$(CStr(Target.Value2))
            
            If Len(valToLearn) > 0 And valToLearn <> "Reject AI Guess" Then
                
                ' --- DYNAMIC SEARCH-ON-ENTER (Safeguarded) ---
                Dim isValid As Boolean
                isValid = True ' Default to True to prevent freezing if Matcher crashes
                
                On Error Resume Next
                isValid = Matcher.IsValidOutputList(valToLearn)
                If Not isValid Then
                    ' Fix: check if it came from the validation list (dropdown)
                    Dim vList As String
                    On Error Resume Next
                    vList = Me.Cells(rowKey, "B").Validation.Formula1
                    On Error GoTo ErrorHandler

                    If Len(vList) > 0 Then
                        If InStr(1, vList, valToLearn, vbTextCompare) > 0 Then
                            isValid = True
                        End If
                    End If

                    ' Fallback: re-evaluate the original input to see if this was a valid multiple rules choice
                    If Not isValid Then
                        Dim rawA As String, outA As String, confA As String
                        rawA = Trim$(CStr(Me.Cells(rowKey, "A").Value2))
                        If Len(rawA) > 0 Then
                            Call Matcher.GetBestMatchForInput(rawA, outA, confA)
                            If InStr(1, confA, "Multiple Rules", vbTextCompare) > 0 Or confA = "AI_HYBRID_MATCH" Then
                                If InStr(1, outA, valToLearn, vbTextCompare) > 0 Then
                                    isValid = True
                                End If
                            End If
                        End If
                    End If
                End If
                On Error GoTo ErrorHandler
                
                If isValid = False Then
                    If Len(valToLearn) >= 3 And InStr(1, valToLearn, "Click to", vbTextCompare) = 0 Then
                        On Error Resume Next ' Safeguard the dropdown builder
                        Call Matcher.BuildSearchDropdown(Me, rowKey, valToLearn)
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
                On Error GoTo ErrorHandler
                
                ' 3. Update Column C's Status
                Me.Cells(rowKey, "C").Value2 = "USER_LEARNED"
                
                ' 4. Ping Python to memorize this rule permanently
                Dim rawInput As String: rawInput = Trim$(CStr(Me.Cells(rowKey, "A").Value2))
                If Len(rawInput) > 0 Then
                    Dim httpLearn As Object
                    Set httpLearn = CreateObject("MSXML2.XMLHTTP")
                    Dim jsonLearn As String
                    jsonLearn = "{""phrase"": """ & Replace(Replace(rawInput, "\", "\\"), """", "\""") & """, ""match"": """ & Replace(Replace(valToLearn, "\", "\\"), """", "\""") & """}"
                    
                    On Error Resume Next
                    httpLearn.Open "POST", "http://127.0.0.1:5000/learn", False
                    httpLearn.setRequestHeader "Content-Type", "application/json"
                    httpLearn.send jsonLearn
                    On Error GoTo ErrorHandler
                    
                    ' ?? 5. FIXED AUTO-PROPAGATE! (Native VBA text match)
                    Dim propR As Long
                    Dim propLast As Long
                    propLast = Me.Cells(Me.Rows.count, 1).End(xlUp).Row

                    If propLast >= 2 Then
                        Dim arrData As Variant
                        arrData = Me.Range("A1:C" & propLast).Value2

                        Dim rowsToUpdate As Object
                        Set rowsToUpdate = CreateObject("Scripting.Dictionary")
                        
                        ' Clean the target phrase to compare against
                        Dim cleanTarget As String
                        cleanTarget = LCase(Trim(rawInput))

                        For propR = 2 To propLast
                            If propR <> rowKey Then
                                Dim candidateInput As String
                                candidateInput = LCase(Trim(CStr(arrData(propR, 1))))
                                
                                ' If it is an identical text match, auto-fill it instantly
                                If Len(candidateInput) > 0 And candidateInput = cleanTarget Then
                                    arrData(propR, 2) = valToLearn
                                    arrData(propR, 3) = "USER_LEARNED"
                                    rowsToUpdate(propR) = True
                                End If
                            End If
                        Next propR

                        ' Instantly dump the updated data back to the sheet
                        Me.Range("A1:C" & propLast).Value2 = arrData

                        ' Clean up the formatting for newly propagated rows
                        If rowsToUpdate.count > 0 Then
                            Dim k As Variant
                            For Each k In rowsToUpdate.keys
                                Me.Range("B" & CLng(k) & ":C" & CLng(k)).Interior.Color = RGB(252, 186, 186)
                                On Error Resume Next
                                Me.Cells(CLng(k), 2).Validation.Delete
                                On Error GoTo ErrorHandler
                            Next k
                        End If
                    End If
                End If
                
            End If
        End If
    End If

SkipLearning:

    ' ============================================================
    ' SCHEDULE LOGIC (Kept Active)
    ' ============================================================
    If Not hitSched Is Nothing Then
        Dim schedRows As Object: Set schedRows = CreateObject("Scripting.Dictionary")
        For Each r In hitSched.Cells: schedRows(r.Row) = True: Next r
        Dim schedKey As Variant
        For Each schedKey In schedRows.keys: UpdateScheduleForRow CLng(schedKey): Next schedKey
    End If

CleanExit:
    On Error Resume Next

    ' ??? TURNED OFF AUTO-REFRESH TO STOP SCREEN JUMPING
    If Target.Cells.count < 50 Then Call PivotBuilder_Refresh
    On Error GoTo 0

    ' ??? REMOVE THE SHIELD
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isRunning = False
    Exit Sub

ErrorHandler:
    ' ?? NEW: Real Error Tracker to expose hidden crashes
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
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    Dim prevText As String
    text = Trim(text)
    
    Do
        prevText = text
        ' Aggressively strip spaces, #, and trailing alphanumeric codes (e.g., " #1", "2B")
        regEx.pattern = "\s*#?\s*\d+[a-zA-Z]*$"
        text = Trim(regEx.Replace(text, ""))
    Loop Until text = prevText
    
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
    Dim previousSelection As Range

    lastRow = Me.Cells(Me.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ' Target the formatting on BOTH Column B and Column C
    Set applies = Me.Range("B2:C" & lastRow)

    ' =========================================================
    ' NEW: The "ActiveCell" Anchor Fix
    ' Excel shifts conditional formatting formulas if your cursor isn't on the top row!
    ' This parks the cursor on B2 just while writing the rule.
    ' =========================================================
    Set previousSelection = Selection
    Me.Range("B2").Select

    ' 1. Scrape away the old formatting rules from Columns B and C
    For i = Me.Columns("B:C").FormatConditions.count To 1 Step -1
        If Me.Columns("B:C").FormatConditions(i).Type = 2 Then ' xlExpression
            ' We look for $B or $C in case Excel previously shifted the row numbers
            If InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$B", vbTextCompare) > 0 Or _
               InStr(1, Me.Columns("B:C").FormatConditions(i).Formula1, "$C", vbTextCompare) > 0 Then
                Me.Columns("B:C").FormatConditions(i).Delete
            End If
        End If
    Next i

    ' 2. The Watchdog Formulas
    Dim formulaOriginal As String
    Dim formulaNew As String

    ' Your exact original conditions (Untouched)
    formulaOriginal = "=OR($B2=""No good match"", $B2=""No matching words"", $B2=""Click to Review Guess..."", $C2=""Requires AI"", $C2=""Check DB"", $C2=""AI Failed to Resolve"", AND(LEFT($C2,14)=""Multiple Rules"", ISERR(SEARCH(""Resolved"",$C2))), ISNUMBER(SEARCH(""Review AI Guess"",$C2)))"

    ' The new conditions for Subtypes, Manual Fixes, and Search Results
    formulaNew = "=OR($B2=""Subtype missing input"", $B2=""Missing input phrase"", ISNUMBER(SEARCH(""Click dropdown to view"", $B2)), $C2=""Requires human"", $C2=""Search Results"", $C2=""Invalid Entry"")"

    ' Apply the original formula
    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaOriginal)
        .Interior.Color = 65535 ' Solid Yellow
        .StopIfTrue = False
    End With

    ' Apply the new formula
    With applies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaNew)
        .Interior.Color = 65535 ' Solid Yellow
        .StopIfTrue = False
    End With

    ' Put the cursor back exactly where the user left it!
    previousSelection.Select
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




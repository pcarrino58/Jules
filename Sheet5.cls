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
' MAIN EVENT LISTENER (Worksheet_Change) - BATCH OPTIMIZED
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanFail

    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim rngMatchInput As Range, rngMatchResult As Range
    Dim rngSchedFreq As Range, rngSchedDate As Range
    Dim hitMatchA As Range, hitMatchB As Range
    Dim hitSched As Range
    Dim r As Range, rowNum As Long

    Set rngMatchInput = Me.Range("A2:A" & Me.Rows.count)
    Set rngMatchResult = Me.Range("B2:B" & Me.Rows.count)
    Set rngSchedFreq = Me.Range("Q2:Q" & Me.Rows.count)
    Set rngSchedDate = Me.Range("T2:T" & Me.Rows.count)

    Set hitMatchA = Intersect(Target, rngMatchInput)
    Set hitMatchB = Intersect(Target, rngMatchResult)
    Set hitSched = Intersect(Target, Union(rngSchedFreq, rngSchedDate))

    If hitMatchA Is Nothing And hitMatchB Is Nothing And hitSched Is Nothing Then GoTo CleanExit

    ' ============================================================
    ' ?? THE "DO NOT DISTURB" SHIELD (Massive Performance Boost)
    ' ============================================================
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual '<-- Stops formula lag

    ' ============================================================
    ' DELETION / CLEAR BOARD LOGIC
    ' ============================================================
    If Target.Cells.count > 20 And Not hitMatchA Is Nothing Then
        If Application.WorksheetFunction.CountA(hitMatchA) = 0 Then
            hitMatchA.offset(0, 1).ClearContents
            hitMatchA.offset(0, 2).ClearContents
            On Error Resume Next
            hitMatchA.offset(0, 1).Validation.Delete
            Call Matcher.CleanupOrphanedPickLists
            
            ' --- NEW: SURGICAL CLEAR FOR AI STAGING ---
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
            ' ------------------------------------------
            
            On Error GoTo CleanFail
        Else
            Call Matcher.MatchOnlyChangedRows(hitMatchA)
        End If
        GoTo CleanExit
    End If
    
    ' ============================================================
    ' INPUT LOGIC (Column A)
    ' ============================================================
    If Not hitMatchA Is Nothing Then
        Dim rngNonBlank As Range, rngBlank As Range, c As Range
        For Each c In hitMatchA.Cells
            If Len(Trim$(CStr(c.Value2))) > 0 Then
                Dim skipAutoMatch As Boolean
                skipAutoMatch = False
                If Not hitMatchB Is Nothing Then
                    If Not Intersect(c.EntireRow, hitMatchB) Is Nothing Then
                        If Len(Trim$(CStr(Me.Cells(c.Row, "B").Value2))) > 0 Then
                            skipAutoMatch = True
                        End If
                    End If
                End If
                If Not skipAutoMatch Then
                    If rngNonBlank Is Nothing Then Set rngNonBlank = c Else Set rngNonBlank = Union(rngNonBlank, c)
                End If
            Else
                If rngBlank Is Nothing Then Set rngBlank = c Else Set rngBlank = Union(rngBlank, c)
            End If
        Next c

        If Not rngBlank Is Nothing Then
            For Each r In rngBlank.Cells
                rowNum = r.Row
                Me.Cells(rowNum, "B").ClearContents
                Me.Cells(rowNum, "C").ClearContents
                On Error Resume Next
                Me.Cells(rowNum, "B").Validation.Delete
                ThisWorkbook.names("PickRow_" & rowNum).Delete
                On Error GoTo CleanFail
            Next r
        End If
        If Not rngNonBlank Is Nothing Then Matcher.MatchOnlyChangedRows rngNonBlank
        EnsureNoGoodMatchHighlight
    End If

    ' ============================================================
    ' LEARNING & CONFIRMATION LOGIC (Column B)
    ' ============================================================
    If Not hitMatchB Is Nothing Then
        
        Dim uniqueRows As Object: Set uniqueRows = CreateObject("Scripting.Dictionary")
        For Each r In hitMatchB.Cells: uniqueRows(r.Row) = True: Next r
        
        Dim key As Variant
        Dim cellInput As Range, cellOutput As Range, cellConf As Range
        
        For Each key In uniqueRows.keys
            Set cellInput = Me.Cells(key, "A")
            Set cellOutput = Me.Cells(key, "B")
            Set cellConf = Me.Cells(key, "C")
            
            Dim valToLearn As String
            valToLearn = Matcher.CleanPastedOutput(CStr(cellOutput.Value2))
            
            ' -------------------------------------------------------------------------
            ' NEW FIX: We MUST check if this came from a dropdown BEFORE we delete it!
            ' -------------------------------------------------------------------------
            Dim skipLearning As Boolean
            skipLearning = Matcher.IsValueFromDropdown(Me, CLng(key), valToLearn)
            
            ' NOW we can safely delete the validation if it's a single cell change
            If Target.CountLarge = 1 And Target.Row = key Then
                On Error Resume Next
                If cellOutput.Validation.Type = xlValidateList Then cellOutput.Validation.Delete
                Dim nm As Name: Set nm = ThisWorkbook.names("PickRow_" & key)
                If Not nm Is Nothing Then nm.Delete
                On Error GoTo CleanFail
            End If
            
            cellOutput.Value2 = valToLearn
            
            If Len(Trim$(CStr(cellInput.Value2))) > 0 And Len(valToLearn) > 0 Then
                If Matcher.IsValidOutputList(valToLearn) Then
                    
                    ' If they picked a dropdown value, we DO NOT LEARN!
                    If skipLearning Then
                        cellConf.Value = "Multiple Rules (Resolved)"
                        
                        ' --- NEW: We still need to manually trigger propagation! ---
                        Dim norm As String, sig As String
                        norm = Matcher.NormalizeAndAlias(Trim$(CStr(cellInput.Value2)), Matcher.mAliasDict)
                        sig = Matcher.BuildLearnSignature(norm)
                        Call Matcher.PropagateLearnedChoice(Me, CLng(key), sig, norm, valToLearn)
                        ' ------------------------------------------------------------
                        
                    Else
                        ' Trigger internal Matcher learning AND Python AI Learning (which also propagates)
                        Matcher.RecordLearnedOverride CLng(key)
                        cellConf.Value = "High"
                    End If
                    
                Else
                    cellConf.Value = "Check DB"
                End If
            End If
        Next key
        EnsureNoGoodMatchHighlight
    End If

    ' ============================================================
    ' SCHEDULE LOGIC
    ' ============================================================
    If Not hitSched Is Nothing Then
        Dim schedRows As Object: Set schedRows = CreateObject("Scripting.Dictionary")
        For Each r In hitSched.Cells: schedRows(r.Row) = True: Next r
        For Each key In schedRows.keys: UpdateScheduleForRow CLng(key): Next key
    End If

CleanExit:
    On Error Resume Next
    
    ' ?? TURNED OFF AUTO-REFRESH TO STOP SCREEN JUMPING
    ' If Target.Cells.count < 50 Then Call PivotBuilder_Refresh
    On Error GoTo 0
    
    ' ?? REMOVE THE SHIELD
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isRunning = False
    Exit Sub
    
CleanFail:
    Resume CleanExit
End Sub

Private Sub UpdateScheduleForRow(ByVal r As Long)
    Dim ws As Worksheet: Set ws = Me
    Dim freqStr As String, startDate As Variant, startMonth As Long
    Dim schedule(1 To 12) As String, m As Long, idx As Long, off As Variant
    Dim hasM As Boolean, hasQ As Boolean, hasS As Boolean, hasA As Boolean

    freqStr = CStr(ws.Cells(r, "Q").Value)
    startDate = ws.Cells(r, "T").Value
    ws.Range(ws.Cells(r, 21), ws.Cells(r, 32)).ClearContents

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
        If schedule(m) <> "" Then ws.Cells(r, 20 + m).Value = schedule(m)
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
    Set applies = Me.Range("B2:B" & lastRow)
    For i = Me.Columns("B").FormatConditions.count To 1 Step -1
        If Me.Columns("B").FormatConditions(i).Type = 2 Then
            If InStr(LCase$(Me.Columns("B").FormatConditions(i).Formula1), "no good match") > 0 Then Me.Columns("B").FormatConditions(i).Delete
        End If
    Next i
    With applies.FormatConditions.Add(Type:=2, Formula1:="=OR($B2=""No good match"", $B2=""No matching words"")")
        .Interior.Color = vbYellow: .StopIfTrue = False
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

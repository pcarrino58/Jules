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
' MAIN EVENT LISTENER (Worksheet_Change)
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

    If Target.Cells.count > 20 And Not hitMatchA Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        If Application.WorksheetFunction.CountA(hitMatchA) = 0 Then
            hitMatchA.offset(0, 1).ClearContents
            hitMatchA.offset(0, 2).ClearContents
            On Error Resume Next
            hitMatchA.offset(0, 1).Validation.Delete
            Call Matcher.CleanupOrphanedPickLists
            On Error GoTo CleanFail
        Else
            Call Matcher.MatchInputToLookup_Top3Matches_Vertical_Fast
        End If
        GoTo CleanExit
    End If
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Not hitMatchA Is Nothing Then
        Dim rngNonBlank As Range, rngBlank As Range, c As Range
        For Each c In hitMatchA.Cells
            If Len(Trim$(CStr(c.Value2))) > 0 Then
                If rngNonBlank Is Nothing Then Set rngNonBlank = c Else Set rngNonBlank = Union(rngNonBlank, c)
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
    ' Learning Logic (Column B)
    ' ============================================================
    If Not hitMatchB Is Nothing Then
        
        ' 1. Clean up dropdown validations natively if they exist
        If Target.CountLarge = 1 And Not Intersect(Target, rngMatchResult) Is Nothing Then
            On Error Resume Next
            If Target.Validation.Type = xlValidateList Then
                Target.Validation.Delete
            End If
            
            ' Exorcise the "Ghost" dropdown names
            Dim nm As Name
            Set nm = ThisWorkbook.names("PickRow_" & Target.Row)
            If Not nm Is Nothing Then nm.Delete
            On Error GoTo CleanFail
        End If
        
        Dim cellInput As Range, cellOutput As Range, cellConf As Range
        Dim k As Range, uniqueRows As Object
        Set uniqueRows = CreateObject("Scripting.Dictionary")
        For Each k In hitMatchB.Cells: uniqueRows(k.Row) = True: Next k
        
        Dim key As Variant
        For Each key In uniqueRows.keys
            Set cellInput = Me.Cells(key, "A")
            Set cellOutput = Me.Cells(key, "B")
            Set cellConf = Me.Cells(key, "C")
            
            ' 2. NEW SAFETY CATCH: Was this row handled by the AI?
            Dim wasAIRow As Boolean
            wasAIRow = False
            If InStr(1, CStr(cellConf.Value2), "AI", vbTextCompare) > 0 Then wasAIRow = True
            
            Dim valToLearn As String
            valToLearn = Trim$(Replace(CStr(cellOutput.Value2), Chr(160), " "))
            
            If Len(Trim$(CStr(cellInput.Value2))) > 0 And Len(valToLearn) > 0 Then
                If Matcher.IsValidOutputList(valToLearn) Then
                    
                    ' 3. Trigger native Excel memory and auto-propagation
                    Matcher.RecordLearnedOverride CLng(key)
                    
                    ' 4. ONLY teach the AI if this was confirming or correcting an AI guess
                    If wasAIRow Then
                        Call TeachAI(Trim$(CStr(cellInput.Value2)), valToLearn)
                        cellConf.Value = "AI Learned!" ' Visual confirmation on your sheet
                    Else
                        cellConf.Value = "High" ' Normal manual entries stay as High
                    End If
                    
                Else
                    cellConf.Value = "Check DB"
                End If
            End If
        Next key
        EnsureNoGoodMatchHighlight
    End If

    If Not hitSched Is Nothing Then
        Dim schedRows As Object
        Set schedRows = CreateObject("Scripting.Dictionary")
        For Each r In hitSched.Cells
            schedRows(r.Row) = True
        Next r
        Dim vRow As Variant
        For Each vRow In schedRows.keys
            UpdateScheduleForRow CLng(vRow)
        Next vRow
    End If

CleanExit:
    Call PivotBuilder_Refresh
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isRunning = False
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

Private Sub UpdateScheduleForRow(ByVal r As Long)
    Dim ws As Worksheet
    Set ws = Me
    Dim freqStr As String, startDate As Variant, startMonth As Long
    Dim schedule(1 To 12) As String, m As Long, idx As Long, off As Variant
    Dim hasM As Boolean, hasQ As Boolean, hasS As Boolean, hasA As Boolean

    freqStr = CStr(ws.Cells(r, "Q").Value)
    startDate = ws.Cells(r, "T").Value
    ws.Range(ws.Cells(r, 21), ws.Cells(r, 32)).ClearContents

    If Len(Trim$(freqStr)) = 0 Or Not IsDateValid(startDate) Then Exit Sub
    
    If IsDate(startDate) Then
        startMonth = Month(CDate(startDate))
    Else
        startMonth = CLng(startDate)
    End If
    
    ParseCodes freqStr, hasM, hasQ, hasS, hasA
    
    If hasA Then schedule(startMonth) = "A"
    
    If hasS Then
        If schedule(startMonth) = "" Then schedule(startMonth) = "S"
        idx = MonthFromOffset(startMonth, 6)
        If schedule(idx) = "" Then schedule(idx) = "S"
    End If

    If hasQ Then
        For Each off In Array(0, 3, 6, 9)
            idx = MonthFromOffset(startMonth, CLng(off))
            If schedule(idx) = "" Then schedule(idx) = "Q"
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
    Dim m As Long
    m = startMonth + offset
    Do While m > 12: m = m - 12: Loop
    Do While m < 1: m = m + 12: Loop
    MonthFromOffset = m
End Function

Private Sub ParseCodes(ByVal freqStr As String, ByRef hasM As Boolean, ByRef hasQ As Boolean, ByRef hasS As Boolean, ByRef hasA As Boolean)
    Dim cleaned As String, tokens() As String, t As Variant, i As Long
    hasM = False: hasQ = False: hasS = False: hasA = False
    cleaned = UCase$(freqStr)
    cleaned = Replace(cleaned, ",", " "): cleaned = Replace(cleaned, "/", " "): cleaned = Replace(cleaned, "-", " ")
    
    For i = 1 To Len(cleaned)
        If Mid$(cleaned, i, 1) < "A" Or Mid$(cleaned, i, 1) > "Z" Then Mid$(cleaned, i, 1) = " "
    Next i
    cleaned = Trim$(cleaned)
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop
    If Len(cleaned) = 0 Then Exit Sub
    tokens = Split(cleaned, " ")

    For Each t In tokens
        Select Case t
            Case "M": hasM = True
            Case "Q": hasQ = True
            Case "S": hasS = True
            Case "A": hasA = True
        End Select
    Next t
End Sub

Private Function IsDateValid(val As Variant) As Boolean
    If IsDate(val) Then IsDateValid = True: Exit Function
    If IsNumeric(val) Then
        If val >= 1 And val <= 12 Then IsDateValid = True: Exit Function
    End If
    IsDateValid = False
End Function

Private Sub EnsureNoGoodMatchHighlight()
    Dim lastRow As Long, applies As Range
    Dim i As Long
    
    lastRow = Me.Cells(Me.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    Set applies = Me.Range("B2:B" & lastRow)
    
    ' Loop backwards to safely find and delete any old versions of this formatting rule
    For i = Me.Columns("B").FormatConditions.count To 1 Step -1
        If Me.Columns("B").FormatConditions(i).Type = 2 Then
            Dim fText As String
            fText = LCase$(Me.Columns("B").FormatConditions(i).Formula1)
            If InStr(fText, "no good match") > 0 Or InStr(fText, "no matching words") > 0 Then
                Me.Columns("B").FormatConditions(i).Delete
            End If
        End If
    Next i
    
    ' Apply the brand new unified rule that catches BOTH phrases
    With applies.FormatConditions.Add(Type:=2, Formula1:="=OR($B2=""No good match"", $B2=""No matching words"")")
        .Interior.Color = vbYellow
        .StopIfTrue = False
    End With
End Sub

' ============================================================
' AI SNIPER EVENT (Worksheet_BeforeDoubleClick)
' ============================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Intersect(Target, Me.Range("B:C")) Is Nothing Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Cells.count > 1 Then Exit Sub
    
    Dim rowNum As Long
    rowNum = Target.Row
    
    Dim inputPhrase As String
    inputPhrase = Trim(Me.Cells(rowNum, "A").Value2)
    
    If inputPhrase = "" Then Exit Sub
    
    Cancel = True
    Me.Cells(rowNum, "C").Value = "AI Thinking..."
    DoEvents
    
    Call Matcher.RunAISniper(Me, rowNum, inputPhrase)
End Sub

' ============================================================
' AI TEACHER (Syncs manual overrides to Python AI Vault)
' ============================================================
Private Sub TeachAI(messyText As String, correctResult As String)
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim jsonPayload As String
    messyText = Replace(messyText, """", "\""")
    correctResult = Replace(correctResult, """", "\""")
    jsonPayload = "{""messy_phrase"": """ & messyText & """, ""official_name"": """ & correctResult & """}"
    http.Open "POST", "http://127.0.0.1:5000/learn", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload
    Set http = Nothing
End Sub

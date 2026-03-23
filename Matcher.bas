Attribute VB_Name = "Matcher"
Option Explicit
          
' ==========================================================
' MASTER MATCHER MODULE - PART 1 of 3
' ==========================================================
          
Private mLookupPhraseSet As Object
Private mAliasDict As Object
Private mStopWords As Object
Private mBoostDict As Object
Private mCoreSet As Object
Private mQualSet As Object
Private mPhraseMap As Object
Private mForcedOutputDict As Object
Private mProtectSet As Object
Private mAliasPrefix As Object
Private mAliasSuffix As Object
Private mStripNumeric As Boolean
Private mStripHash As Boolean
Private mStripAlphaNum As Boolean
Private mLookupPhrases() As String
Private mLookupWords() As Variant
Private mLookupDicts() As Object
Private mLookupSubjects() As String
Private mRowCategory() As String
Private mLookupCount As Long
Private mLookupWordCounts() As Long
Private mTokenIndex As Object
Private mLookupVocab As Object
Private mTokenWeights As Object
Private mTokenCatFreq As Object
Private mSignatureDict As Object
          
Private MIN_MATCH_THRESHOLD As Double
Private Const DOMINANCE_MARGIN As Double = 0.3
Private Const MIN_REQUIRED_OVERLAP_DEFAULT As Long = 2
Private Const PICK_SHEET_NAME As String = "PickLists"
Private Const LEARNED_SHEET As String = "LearnedMappings"
          
' Category prepass
Private mCategoryDict As Object
Private mCategoryKeys() As String
Private mCategoryCount As Long
Private mCategoryAliasIndex As Object
Private mRegEx As Object
Private mExternalFileCache As Object
          
' Learned Overrides
Private mLearnedExact As Object
Private mLearnedSigBest As Object
Private mLearnedSigCount As Object
Private mLearnedVocabSigBest As Object
Private mLearnedLoaded As Boolean
Private mLearnedSubstKeys() As String
          
Public Sub MatchInputToLookup_Top3Matches_Vertical_Fast()
    DoMatchingAndAttachDropDowns
End Sub
          
Public Sub MatchOnlyChangedRows(ByVal rngChanged As Range)
    Dim wsInput As Worksheet, wsLookup As Worksheet
    Dim lastInputRow As Long
    Dim rowsDict As Object, c As Range
    Dim outVal As String, confVal As String
              
    On Error GoTo CleanFail
    Set wsInput = rngChanged.Worksheet
    Set wsLookup = GetLookupSheet()
    If wsLookup Is Nothing Then Exit Sub
              
    Set rowsDict = CreateObject("Scripting.Dictionary")
              
    lastInputRow = wsInput.Cells(wsInput.Rows.count, "A").End(xlUp).Row
              
    For Each c In rngChanged
        If c.Column = 1 Then
            If c.Row >= 2 And c.Row <= lastInputRow Then rowsDict(c.Row) = True
        End If
    Next c
    If rowsDict.count = 0 Then Exit Sub
          
    InitializeMatcher wsLookup
              
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
          
    Dim r As Variant, inputPhrase As String
    Dim cellVal As Variant
          
    For Each r In rowsDict.keys
        cellVal = wsInput.Cells(r, "A").Value
        If Not IsError(cellVal) Then
            inputPhrase = CStr(cellVal)
            
            ' 1. The Sniper: Run your strict VBA Matcher first
            GetBestMatchForInput inputPhrase, outVal, confVal
            
            ' ============================================================
            ' ROBUST FAILURE CHECK
            ' ============================================================
            Dim vbaFailed As Boolean
            vbaFailed = False
            
            If Trim(outVal) = "" Then vbaFailed = True
            If InStr(1, LCase(outVal), "no good") > 0 Then vbaFailed = True
            If InStr(1, LCase(outVal), "no match") > 0 Then vbaFailed = True
            If InStr(1, LCase(confVal), "fuzzy") > 0 Then vbaFailed = True
            
            ' ============================================================
            ' 1.5 THE EXCEL MEMORY BYPASS
            ' ============================================================
            If vbaFailed Then
                Dim wsMem As Worksheet
                Dim memFound As Range
                
                On Error Resume Next
                Set wsMem = ThisWorkbook.Sheets("LearnedMappings")
                On Error GoTo CleanFail
                
                If Not wsMem Is Nothing Then
                    Set memFound = wsMem.Columns(1).Find(What:=Trim(inputPhrase), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    If Not memFound Is Nothing Then
                        outVal = Trim(memFound.offset(0, 1).Value)
                        confVal = "Local Memory"
                        vbaFailed = False ' Turn off failure flag
                    End If
                End If
            End If
            
            ' ============================================================
            ' 2. THE AI SAFETY NET
            ' ============================================================
            If vbaFailed Then
                Dim aiGuess As String
                aiGuess = GetAIMatch(inputPhrase)
                
                If aiGuess <> "UNKNOWN" And aiGuess <> "AI Offline" And aiGuess <> "" Then
                    outVal = aiGuess
                    confVal = "AI Guess - Review Needed"
                End If
            End If
                      
            ' 3. Write final winner to the sheet
            wsInput.Cells(r, "B").Value = outVal
            wsInput.Cells(r, "C").Value = confVal
            wsInput.Cells(r, "B").WrapText = True
            
            ' ============================================================
            ' THE VISUAL FLAG
            ' ============================================================
            If confVal = "AI Guess - Review Needed" Then
                wsInput.Cells(r, "C").Interior.Color = RGB(255, 192, 0)
                wsInput.Cells(r, "C").Font.Bold = True
            ElseIf InStr(1, LCase(outVal), "no match") > 0 Or InStr(1, LCase(outVal), "no good") > 0 Or Trim(outVal) = "UNKNOWN" Then
                wsInput.Cells(r, "C").Interior.Color = RGB(255, 153, 153)
                wsInput.Cells(r, "C").Font.Bold = True
            ElseIf confVal = "Local Memory" Then
                wsInput.Cells(r, "C").Interior.Color = RGB(252, 186, 186)
                wsInput.Cells(r, "C").Font.Bold = False
            End If
                      
            CreateSelectionListForSingleRow wsInput, CLng(r)
        Else
            wsInput.Cells(r, "C").Value = "Input Error"
            wsInput.Cells(r, "C").Interior.Color = RGB(255, 0, 0)
        End If
    Next r
          
CleanExit:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub
CleanFail:
    MsgBox "Error in MatchOnlyChangedRows (Row " & r & "): " & Err.Description, vbCritical
    Resume CleanExit
End Sub
          
Public Sub CleanupOrphanedPickLists()
    Dim nm As Name
    On Error Resume Next
    For Each nm In ThisWorkbook.names
        If nm.Name Like "PickRow_*" Then nm.Delete
    Next nm
    On Error GoTo 0
End Sub
          
Public Sub RecordLearnedOverride(ByVal rowNum As Long)
    Dim wsInput As Worksheet, wsLearn As Worksheet
    Dim rawInput As String, chosen As String
    Dim normalized As String, sig As String
    Dim lastRow As Long, f As Range
    Dim wsLookup As Worksheet
          
    If rowNum < 2 Then Exit Sub
              
    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets("Helper Sheet")
    On Error GoTo 0
              
    If wsInput Is Nothing Then Exit Sub
          
    rawInput = CStr(wsInput.Cells(rowNum, "A").Value2)
    chosen = CStr(wsInput.Cells(rowNum, "B").Value2)
              
    chosen = SanitizeLearnedOutput(chosen)
    If Not IsValidOutputList(chosen) Then Exit Sub
    If Len(Trim$(rawInput)) = 0 Then Exit Sub
    If Len(Trim$(chosen)) = 0 Then Exit Sub
          
    If mAliasDict Is Nothing Then EnsureDictionaries
    LoadLearnedOverrides
          
    normalized = NormalizeAndAlias(rawInput, mAliasDict)
    If Len(normalized) = 0 Then Exit Sub
          
    sig = BuildLearnSignature(normalized)
    EnsureLearnedSheet
    Set wsLearn = ThisWorkbook.Worksheets(LEARNED_SHEET)
          
    mLearnedExact(normalized) = chosen
    If Len(sig) > 0 Then
        mLearnedSigBest(sig) = chosen
        If Not mLearnedSigCount.Exists(sig) Then mLearnedSigCount(sig) = 0
        mLearnedSigCount(sig) = CLng(mLearnedSigCount(sig)) + 1
    End If
          
    lastRow = wsLearn.Cells(wsLearn.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
              
    Set f = wsLearn.Range("A2:A" & lastRow).Find(What:=normalized, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        wsLearn.Cells(f.Row, "B").Value2 = chosen
        wsLearn.Cells(f.Row, "C").Value2 = sig
        wsLearn.Cells(f.Row, "D").Value2 = CLng(val(wsLearn.Cells(f.Row, "D").Value2)) + 1
        wsLearn.Cells(f.Row, "E").Value2 = Now
    Else
        wsLearn.Cells(lastRow + 1, "A").Value2 = normalized
        wsLearn.Cells(lastRow + 1, "B").Value2 = chosen
        wsLearn.Cells(lastRow + 1, "C").Value2 = sig
        wsLearn.Cells(lastRow + 1, "D").Value2 = 1
        wsLearn.Cells(lastRow + 1, "E").Value2 = Now
    End If
              
    AppendToExternalFile normalized, chosen, sig
     
    ' --- NEW: Trigger Auto-Propagation across the sheet ---
    PropagateLearnedChoice wsInput, rowNum, sig, chosen
End Sub
          
Public Function SanitizeLearnedOutput(ByVal outP As String) As String
    If LCase$(Trim$(outP)) = "d of the" Then SanitizeLearnedOutput = "": Exit Function
    SanitizeLearnedOutput = outP
End Function
          
Public Function IsValidOutputList(ByVal outP As String) As Boolean
    Dim parts() As String, i As Long, s As String
              
    If mLookupPhraseSet Is Nothing Then
        Dim wsL As Worksheet
        Set wsL = GetLookupSheet()
        If Not wsL Is Nothing Then InitializeMatcher wsL
    End If
              
    If mLookupPhraseSet Is Nothing Then Exit Function
          
    outP = Replace(outP, vbCrLf, vbLf)
    outP = Replace(outP, vbCr, vbLf)
    outP = Trim$(outP)
              
    If Len(outP) = 0 Then Exit Function
          
    parts = Split(outP, vbLf)
              
    For i = LBound(parts) To UBound(parts)
        s = GetStandardKey(parts(i))
        If Len(s) > 0 Then
            If Not mLookupPhraseSet.Exists(s) Then Exit Function
        End If
    Next i
              
    IsValidOutputList = True
End Function
          
Private Sub DoMatchingAndAttachDropDowns()
    Dim wsInput As Worksheet, wsLookup As Worksheet
    Dim lastInputRow As Long
    Dim inputArr As Variant
    Dim n As Long, i As Long
    Dim outVals() As Variant, confVals() As Variant
    Dim inputPhrase As String
    Dim outVal As String, confVal As String
          
    On Error GoTo CleanFail
    If Not SheetExists("Helper Sheet") Then MsgBox "Helper Sheet not found!", vbCritical: Exit Sub
    Set wsInput = ThisWorkbook.Worksheets("Helper Sheet")
    Set wsLookup = GetLookupSheet()
    If wsLookup Is Nothing Then MsgBox "Lookup Sheet not found!": Exit Sub
          
    lastInputRow = wsInput.Cells(wsInput.Rows.count, "A").End(xlUp).Row
    If lastInputRow < 2 Then MsgBox "No data found.", vbInformation: Exit Sub
          
    CleanupOrphanedPickLists
          
    inputArr = wsInput.Range("A2:A" & lastInputRow).Value
    If Not IsArray(inputArr) Then
        Dim tempArr() As Variant
        ReDim tempArr(1 To 1, 1 To 1)
        tempArr(1, 1) = inputArr
        inputArr = tempArr
    End If
    InitializeMatcher wsLookup
          
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
          
    n = UBound(inputArr, 1)
    ReDim outVals(1 To n, 1 To 1)
    ReDim confVals(1 To n, 1 To 1)
          
    For i = 1 To n
        If Not IsError(inputArr(i, 1)) Then
            inputPhrase = CStr(inputArr(i, 1))
            GetBestMatchForInput inputPhrase, outVal, confVal
            outVals(i, 1) = outVal
            confVals(i, 1) = confVal
        Else
            outVals(i, 1) = "Source Error in Col A"
            confVals(i, 1) = "Error"
        End If
    Next i
          
    With wsInput
        .Range("B2").Resize(n, 1).Value = outVals
        .Range("C2").Resize(n, 1).Value = confVals
        .Range("B2").Resize(n, 1).WrapText = True
    End With
          
    CreateSelectionListsForColumnB wsInput, n
          
CleanExit:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    MsgBox "Matching complete!", vbInformation
    Exit Sub
CleanFail:
    MsgBox "Error in DoMatchingAndAttachDropDowns (Row " & i & "): " & Err.Description, vbCritical
    Resume CleanExit
End Sub
          
Private Function GetLookupSheet() As Worksheet
    On Error Resume Next
    Set GetLookupSheet = ThisWorkbook.Worksheets("Uniformat RS Means Lookup")
    On Error GoTo 0
End Function
          
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
          
Private Sub GetBestMatchForInput(ByVal rawInput As String, ByRef outVal As String, ByRef confVal As String)
    Dim normalizedInput As String
    Dim filteredInput As String
    Dim cleanRaw As String
    Dim sNorm As String
    Dim inputWords() As String
    Dim inputSig As String
              
    outVal = "": confVal = ""
              
    LoadLearnedOverrides
              
    cleanRaw = GetStandardKey(rawInput)
    If Len(cleanRaw) = 0 Then
        outVal = "Blank Input": confVal = "none"
        Exit Sub
    End If
              
    normalizedInput = NormalizeAndAlias(rawInput, mAliasDict)
    filteredInput = GetFilteredInput(normalizedInput)
      
    If Len(filteredInput) = 0 Then
        outVal = "No matching words": confVal = "none"
        Exit Sub
    End If
  
    normalizedInput = SubstituteLearnedPhrases(normalizedInput)
    sNorm = " " & LCase$(normalizedInput) & " "
              
    If IsNonAssetInput(sNorm) Then
        outVal = "": confVal = "excluded"
        Exit Sub
    End If
 
    ' -------------------------------------------------------------------------
    ' PRIORITY 1: RULES SHEET (Forced Lists) - MOVED TO TOP!
    ' -------------------------------------------------------------------------
    If CheckRulesMatch(cleanRaw, normalizedInput, outVal, confVal) Then Exit Sub
 
    ' -------------------------------------------------------------------------
    ' PRIORITY 2: EXACT LOOKUP MATCH (Restored)
    ' -------------------------------------------------------------------------
    If CheckExactLookupMatch(rawInput, normalizedInput, outVal, confVal) Then Exit Sub

    ' -------------------------------------------------------------------------
    ' PRIORITY 0: PUMP DEFAULT LOGIC
    ' -------------------------------------------------------------------------
    Dim sPad As String
    sPad = " " & normalizedInput & " "
      
    If InStr(1, sPad, " pump ", vbTextCompare) > 0 Then
        Dim pumpTypes As Variant, pType As Variant, isSpecificPump As Boolean
          
        pumpTypes = Array("sump", "submersible", "fuel", "condensate", "hydraulic", _
                          "fire", "well", "ejector", "metering", "vacuum", _
                          "sewage", "lift", "jockey", "booster", "rotary", _
                          "centrifugal", "circ", "circulation", "dosing", "chem")
          
        isSpecificPump = False
        For Each pType In pumpTypes
            If InStr(1, sPad, " " & pType & " ", vbTextCompare) > 0 Then
                isSpecificPump = True
                Exit For
            End If
        Next pType
          
        If Not isSpecificPump Then
            outVal = "Centrifugal Pump"
            confVal = "Default Logic"
            Exit Sub
        End If
    End If
 
    ' -------------------------------------------------------------------------
    ' PRIORITY 3: LEARNED OVERRIDES
    ' -------------------------------------------------------------------------
    If TryLearnedOverride(normalizedInput, outVal, confVal) Then Exit Sub
          
    ' -------------------------------------------------------------------------
    ' PRIORITY 4: MULTI-WORD PHRASE MATCH
    ' -------------------------------------------------------------------------
    If CheckMultiWordPhraseMatch(filteredInput, outVal, confVal) Then Exit Sub
          
    ' -------------------------------------------------------------------------
    ' PRIORITY 4.5 & 4.6: SIGNATURE MATCHES
    ' -------------------------------------------------------------------------
    inputSig = GetCanonicalSignature(normalizedInput)
    If Len(inputSig) > 0 Then
        If mSignatureDict.Exists(inputSig) Then
            outVal = mSignatureDict(inputSig)
            confVal = "Pattern Match"
            Exit Sub
        End If
    End If
             
    inputSig = GetVocabSignature(normalizedInput)
    If Len(inputSig) > 0 Then
        If mSignatureDict.Exists(inputSig) Then
            outVal = mSignatureDict(inputSig)
            confVal = "Pattern Match (Filtered)"
            Exit Sub
        End If
    End If
          
    ' -------------------------------------------------------------------------
    ' PRIORITY 5: ONE-WORD PHRASE MATCH
    ' -------------------------------------------------------------------------
    If CheckOneWordPhraseMatch(filteredInput, outVal, confVal) Then Exit Sub
              
    ' -------------------------------------------------------------------------
    ' PRIORITY 6: FUZZY SCORING (Best Guess)
    ' -------------------------------------------------------------------------
    inputWords = Tokenize(normalizedInput)
    If IsEmptyArray(inputWords) Then
        outVal = "No input": confVal = "none"
    Else
        MatchOneRow normalizedInput, inputWords, outVal, confVal
    End If
End Sub

          
Private Sub MatchOneRow(ByVal normalizedInput As String, ByRef inputWords() As String, _
                        ByRef outVal As String, ByRef confVal As String)
    Dim bestScore As Double, confidence As String
    Dim topPhrase() As String
    Dim topScore() As Double
    ReDim topPhrase(1 To 3)
    ReDim topScore(1 To 3)
              
    Dim i As Long
    Dim constructedWords() As String
    Dim constructedPhrase As String
    Dim cwCount As Long
              
    cwCount = 0
    ReDim constructedWords(0 To UBound(inputWords))
          
    ' =========================================================
    ' PATTERN CONSTRUCTION LOGIC
    ' =========================================================
    For i = LBound(inputWords) To UBound(inputWords)
        Dim w As String: w = LCase$(Trim$(inputWords(i)))
                
        If mTokenIndex.Exists(w) Then
            constructedWords(cwCount) = w
            cwCount = cwCount + 1
        ' Else
            ' TYPO WATCHDOG REMOVED HERE
            ' Unrecognized words are permanently discarded
        End If
    Next i
              
    Dim finalInputWords() As String
    If cwCount > 0 Then
        ReDim Preserve constructedWords(0 To cwCount - 1)
        finalInputWords = constructedWords
        constructedPhrase = Join(constructedWords, " ")
    Else
        finalInputWords = inputWords
        constructedPhrase = normalizedInput
    End If
          
    ' =========================================================
    ' FILTERING & TOKENIZATION (UPDATED: UNIQUE ONLY)
    ' =========================================================
    Dim inputSet As Object: Set inputSet = CreateObject("Scripting.Dictionary")
    Dim filteredWords() As String
    Dim fwCount As Long, ww As String
              
    ReDim filteredWords(0 To UBound(finalInputWords))
    fwCount = 0
      
    ' NEW: De-Duplication Dictionary
    Dim seenDict As Object
    Set seenDict = CreateObject("Scripting.Dictionary")
      
    For i = LBound(finalInputWords) To UBound(finalInputWords)
        ww = LCase$(Trim$(finalInputWords(i)))
        If ww <> "" Then
            ' Only process if we haven't seen this word yet in this row
            If Not seenDict.Exists(ww) Then
                If IsMeaningfulToken(ww, mStopWords) Then
                    If IsAllAlphaNum(ww) Then
                        If Not IsTagLike(ww) Then
                              
                             ' Apply the Strict Filter we discussed earlier
                             If IsAllowedScoringToken(ww) Then
                                filteredWords(fwCount) = ww
                                fwCount = fwCount + 1
                                seenDict(ww) = True ' Mark as seen
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
      
    If fwCount = 0 Then
        outVal = "No good match": confVal = "none"
        Exit Sub
    Else
        ReDim Preserve filteredWords(0 To fwCount - 1)
    End If
          
    Dim inputTotalWeight As Double: inputTotalWeight = 0#
    Dim iw As Double
    For i = LBound(filteredWords) To UBound(filteredWords)
        inputSet(filteredWords(i)) = True
        iw = 1#
        If Not mTokenWeights Is Nothing Then
             If mTokenWeights.Exists(filteredWords(i)) Then iw = CDbl(mTokenWeights(filteredWords(i)))
        End If
        inputTotalWeight = inputTotalWeight + iw
    Next i
          
    Dim mustTokens As Object: Set mustTokens = ExtractCoreTokens(filteredWords, mCoreSet)
    Dim qualTokens As Object: Set qualTokens = ExtractQualifierTokens(filteredWords, mQualSet)
    Dim forceCategory As String: forceCategory = InferForcedCategory(constructedPhrase)
      
    ' Infer Subject and Category for Semantic Weighting
    Dim inputSubject As String: inputSubject = GetInputSubject(filteredWords)
    Dim inputCategory As String: inputCategory = ""
    If fwCount > 0 Then
        inputCategory = GetBestCategory(filteredWords)
    End If
          
    ' =========================================================
    ' CANDIDATE SEARCH
    ' =========================================================
    Dim candIDs As Object: Set candIDs = CreateObject("Scripting.Dictionary")
    For i = LBound(filteredWords) To UBound(filteredWords)
        Dim w2 As String: w2 = filteredWords(i)
        If mTokenIndex.Exists(w2) Then
            Dim coll As Collection: Set coll = mTokenIndex(w2)
            Dim v As Variant
            For Each v In coll
                candIDs(v) = True
            Next v
        End If
    Next i
       
    If candIDs.count = 0 Then
        For i = 1 To mLookupCount: candIDs(i) = True: Next i
    End If
       
    bestScore = 0#: confidence = ""
    Dim id As Variant
              
    ' =========================================================
    ' SCORING LOOP
    ' =========================================================
    For Each id In candIDs.keys
        Dim d As Object: Set d = mLookupDicts(id)
                  
        If Not ContainsAllTokens(d, mustTokens) Then GoTo NextCand
        ' PROPOSED (Bonus-based):
        ' Don't skip candidates just because they miss a qualifier.
        ' Instead, calculate the penalty later in the scoring section.
        Dim qualMatchCount As Long
        qualMatchCount = GetMeaningfulOverlapCount(candIDs, qualTokens, mStopWords)
         
        If Len(forceCategory) > 0 Then
            If Not CandidateContainsCategory(d, forceCategory) Then GoTo NextCand
        End If
          
        Dim matchedCnt As Long
        Dim s As Double
           
        ' F1 Overlap Score
        s = CalculateOverlapScoreF1(inputSet, mLookupWords(id), mStopWords, mTokenWeights, inputTotalWeight, matchedCnt)
           
        If s > 0 Then
            ' Bigram Bonus
            Dim bigramSim As Double
            bigramSim = GetBigramSimilarity(filteredWords, mLookupWords(id))
            If bigramSim > 0 Then s = s * (1 + (bigramSim * 0.2))
               
            ' SEMANTIC LOGIC (Category and Subject alignment)
            Dim semBoost As Double: semBoost = 0#
               
            ' 1. Category Match
            If Len(inputCategory) > 0 Then
                If LCase$(mRowCategory(id)) = inputCategory Then semBoost = semBoost + 0.1
            End If
               
            ' 2. Subject Match (Head Noun) - UPDATED WITH PENALTY
            If Len(inputSubject) > 0 Then
                Dim lookupSubj As String: lookupSubj = mLookupSubjects(id)
                If Len(lookupSubj) > 0 Then
                    ' Check if the primary asset subjects match
                    If IsFuzzyMatch(inputSubject, lookupSubj) Then
                        semBoost = semBoost + 0.2 ' Reward alignment
                    Else
                        ' NEW LOGIC: Only penalize if the lookup doesn't even contain our input subject!
                        ' This allows "sump pump" to match "Pump, sump" without penalty.
                        If Not d.Exists(inputSubject) Then
                            s = s * 0.4
                        End If
                    End If
                End If
            End If
            s = s + semBoost
        End If
          
        ' Check Match Thresholds
        If s >= MIN_MATCH_THRESHOLD Then
            Dim ov As Long: ov = GetMeaningfulOverlapCount(inputSet, mLookupWords(id), mStopWords)
            If ov < MIN_REQUIRED_OVERLAP_DEFAULT Then
                Dim permitSingle As Boolean: permitSingle = False
                If ov = 1 Then
                    For i = LBound(mLookupWords(id)) To UBound(mLookupWords(id))
                        Dim w3 As String: w3 = LCase$(Trim$(mLookupWords(id)(i)))
                        If IsMeaningfulToken(w3, mStopWords) Then
                            If inputSet.Exists(w3) Then
                                If IsStrongSingleToken(w3) Then permitSingle = True
                                Exit For
                            End If
                        End If
                    Next i
                End If
                If Not permitSingle Then GoTo NextCand
            End If
            InsertTop3 topPhrase, topScore, mLookupPhrases(id), s
        End If
           
        If s > bestScore Then
            bestScore = s
            confidence = "Fuzzy (" & Format(s, "0.0") & ")"
        End If
NextCand:
    Next id
              
    ' =========================================================
    ' FINAL OUTPUT GENERATION
    ' =========================================================
    Dim collapseToSingle As Boolean
    If CountNonEmpty(topPhrase) = 1 Then
        collapseToSingle = True
    Else
        collapseToSingle = ShouldCollapseToSingle(topScore)
    End If
              
    Dim resultStr As String: resultStr = JoinNonEmpty(topPhrase, vbLf)
    If resultStr = "" Then
        Dim fallbackCat As String
        fallbackCat = GetBestCategory(filteredWords)
         
        ' NEW: Only show the category list if it contains 6 or fewer items
        If Len(fallbackCat) > 0 Then
            If mCategoryDict(fallbackCat).count > 6 Then
                outVal = "No good match": confVal = "none"
            Else
                outVal = "No good match. Did you mean: " & UCase(fallbackCat) & "?" & vbLf & GetCategoryItemsList(fallbackCat)
                confVal = "category"
            End If
        Else
            outVal = "No good match": confVal = "none"
        End If
    Else
        If collapseToSingle Then outVal = topPhrase(1) Else outVal = resultStr
        confVal = confidence
    End If
End Sub
' ==========================================================
' MASTER MATCHER MODULE - PART 2 of 3
' ==========================================================
          
' NEW SCORING FUNCTION: F1 Score (Harmonic Mean)
Public Function CalculateOverlapScoreF1(ByVal inputSet As Object, ByVal lookupWords As Variant, ByVal stopWords As Object, ByVal tokenWeights As Object, ByVal inputTotalWeight As Double, Optional ByRef matchedCountOut As Long) As Double
    Dim i As Long, matchedCount As Long, totalCount As Long, w As String
    Dim inputKey As Variant
    Dim matchedWeight As Double, targetTotalWeight As Double
    Dim tw As Double
              
    matchedCountOut = 0
    If IsEmptyArray(lookupWords) Then CalculateOverlapScoreF1 = 0#: Exit Function
    If inputTotalWeight <= 0 Then CalculateOverlapScoreF1 = 0#: Exit Function
          
    On Error GoTo EH
    matchedCount = 0: totalCount = 0
    matchedWeight = 0#
    targetTotalWeight = 0#
              
    For i = LBound(lookupWords) To UBound(lookupWords)
        w = Trim$(lookupWords(i))
        If IsMeaningfulToken(w, stopWords) Then
            totalCount = totalCount + 1
                      
            tw = 1#
            If Not tokenWeights Is Nothing Then
                If tokenWeights.Exists(w) Then tw = CDbl(tokenWeights(w))
            End If
                      
            targetTotalWeight = targetTotalWeight + tw
                      
            Dim found As Boolean: found = False
            Dim sim As Double: sim = 0#
            Dim bestSim As Double: bestSim = 0#
                      
            If inputSet.Exists(w) Then
                found = True
                sim = 1#
            Else
                For Each inputKey In inputSet.keys
                    Dim s As Double
                    s = GetFuzzySimilarity(w, CStr(inputKey))
                    If s > bestSim Then bestSim = s
                Next inputKey
                sim = bestSim
                If sim > 0 Then found = True
            End If
                      
            If found Then
                matchedCount = matchedCount + 1
                matchedWeight = matchedWeight + (tw * sim)
            End If
        End If
    Next i
              
    matchedCountOut = matchedCount
       
    If matchedWeight = 0 Then CalculateOverlapScoreF1 = 0#: Exit Function
       
    ' F1 Score Calculation
    Dim precision As Double
    Dim recall As Double
       
    precision = matchedWeight / inputTotalWeight
    recall = matchedWeight / targetTotalWeight
       
    If (precision + recall) > 0 Then
        CalculateOverlapScoreF1 = 2 * (precision * recall) / (precision + recall)
    Else
        CalculateOverlapScoreF1 = 0#
    End If
       
    Exit Function
EH: CalculateOverlapScoreF1 = 0#
End Function
          
Private Sub InitializeMatcher(wsLookup As Worksheet)
    MIN_MATCH_THRESHOLD = 0.5
    mLearnedLoaded = False ' Force reload of learned data
    If mRegEx Is Nothing Then
        Set mRegEx = CreateObject("VBScript.RegExp")
        With mRegEx
            .Global = True
            .IgnoreCase = True
            .Pattern = "(\d)x(\d)"
        End With
    End If
    EnsureDictionaries
    BuildLookupCache wsLookup
End Sub
          
Private Function CheckExactLookupMatch(ByVal rawInput As String, ByVal normInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    If mLookupPhraseSet Is Nothing Then Exit Function
    Dim scrubbedRaw As String, scrubbedNorm As String
    scrubbedRaw = GetStandardKey(rawInput)
    scrubbedNorm = GetStandardKey(normInput)
    If mLookupPhraseSet.Exists(scrubbedRaw) Then
        outVal = Trim$(rawInput)
        confVal = "Exact Match"
        CheckExactLookupMatch = True
        Exit Function
    End If
    If mLookupPhraseSet.Exists(scrubbedNorm) Then
        outVal = normInput
        confVal = "Exact Match"
        CheckExactLookupMatch = True
        Exit Function
    End If
End Function
          
Private Function CheckRulesMatch(ByVal rawInput As String, ByVal normInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    Dim sNorm As String, sRaw As String
    Dim key As Variant, otherKey As Variant
    Dim matchKeys As Object, finalResults As Object
    Dim isSubset As Boolean
              
    If mForcedOutputDict Is Nothing Then
        CheckRulesMatch = False
        Exit Function
    End If
              
    ' --- NEW: Priority Check for Exact Learned Overrides ---
    ' Before we apply generic rules (like "Pump" -> "Centrifugal"),
    ' check if the user has explicitly overridden this exact phrase.
    If Not mLearnedExact Is Nothing Then
        If mLearnedExact.Exists(normInput) Then
            outVal = SanitizeLearnedOutput(mLearnedExact(normInput))
            confVal = "Learned (Override Rule)"
            CheckRulesMatch = True
            Exit Function
        End If
    End If
    ' -----------------------------------------------------

    Set matchKeys = CreateObject("Scripting.Dictionary")
    Set finalResults = CreateObject("Scripting.Dictionary")
              
    ' Strip hyphens and slashes so tags like "DHWT-2" become "DHWT 2"
    sRaw = " " & GetStandardKey(rawInput) & " "
    sRaw = Replace(sRaw, "-", " ")
    sRaw = Replace(sRaw, "_", " ")
    sRaw = Replace(sRaw, "/", " ")
    sRaw = Application.WorksheetFunction.Trim(sRaw)
    sRaw = " " & sRaw & " "
     
    sNorm = " " & normInput & " "
              
    ' 1. Find ALL potential rule keys that exist in the input
    For Each key In mForcedOutputDict.keys
        If InStr(1, sNorm, " " & key & " ", vbTextCompare) > 0 Or _
           InStr(1, sRaw, " " & key & " ", vbTextCompare) > 0 Then
            matchKeys(CStr(key)) = True
        End If
    Next key
 
    ' 2. Filter out shorter matches that are actually part of longer matches
    For Each key In matchKeys.keys
        isSubset = False
     
        For Each otherKey In matchKeys.keys
            If key <> otherKey Then
                If InStr(1, " " & otherKey & " ", " " & key & " ", vbTextCompare) > 0 Then
                    isSubset = True
                    Exit For
                End If
            End If
        Next otherKey
                  
        If Not isSubset Then
            Dim parts() As String, i As Long
            Dim ruleOutput As String
               
            If Not IsNull(mForcedOutputDict(key)) Then
                ruleOutput = CStr(mForcedOutputDict(key))
                parts = Split(ruleOutput, vbLf)
                        
                For i = LBound(parts) To UBound(parts)
                    If Len(Trim$(parts(i))) > 0 Then finalResults(Trim$(parts(i))) = True
                Next i
            End If
        End If
    Next key
       
    ' 3. Output the combined unique results
    If finalResults.count > 0 Then
         
        ' --- NEW: Memory Check Before Applying Generic Rule List ---
        Dim ruleSig As String
        ruleSig = BuildLearnSignature(normInput)
        If Len(ruleSig) > 0 Then
            If Not mLearnedSigBest Is Nothing Then
                If mLearnedSigBest.Exists(ruleSig) Then
                    outVal = SanitizeLearnedOutput(mLearnedSigBest(ruleSig))
                    confVal = "Rule Hit (Learned Override)"
                    CheckRulesMatch = True
                    Exit Function
                End If
            End If
        End If
        ' -----------------------------------------------------------
         
        outVal = Join(finalResults.keys, vbLf)
        If finalResults.count = 1 Then
            confVal = "Rule Hit"
        Else
            confVal = "Multiple Rules (" & finalResults.count & ")"
        End If
        CheckRulesMatch = True
        Exit Function
    End If
          
    CheckRulesMatch = False
End Function
Private Function CheckCategoryMatch(ByVal normInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    If mCategoryDict Is Nothing Then Exit Function
    If mCategoryDict.Exists(normInput) Then
        outVal = "Category Match: " & UCase(normInput) & vbLf & GetCategoryItemsList(normInput)
        confVal = "category"
        CheckCategoryMatch = True
        Exit Function
    End If
    Dim i As Long, cat As String
    Dim sInput As String
    sInput = " " & normInput & " "
    If mCategoryCount > 0 Then
        For i = 1 To mCategoryCount
            cat = mCategoryKeys(i)
            If InStr(1, sInput, " " & cat & " ", vbTextCompare) > 0 Then
                outVal = "Category Match: " & UCase(cat) & vbLf & GetCategoryItemsList(cat)
                confVal = "category"
                CheckCategoryMatch = True
                Exit Function
            End If
        Next i
    End If
    CheckCategoryMatch = False
End Function
          
Private Function CheckOneWordPhraseMatch(ByVal filteredInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    Dim j As Long
    Dim sInput As String
    sInput = " " & filteredInput & " "
    For j = 1 To mLookupCount
        If mLookupWordCounts(j) = 1 Then
            Dim phrase As String
            Dim toks As Variant
            toks = mLookupWords(j)
            If Not IsEmptyArray(toks) Then
                phrase = toks(0)
                If InStr(1, sInput, " " & phrase & " ", vbTextCompare) > 0 Then
                    outVal = mLookupPhrases(j)
                    confVal = "Single Word"
                    CheckOneWordPhraseMatch = True
                    Exit Function
                End If
            End If
        End If
    Next j
    CheckOneWordPhraseMatch = False
End Function
          
Private Function CheckMultiWordPhraseMatch(ByVal filteredInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    Dim j As Long
    Dim sInput As String
    Dim bestLen As Long, bestIdx As Long
    sInput = " " & filteredInput & " "
    bestLen = 0
    bestIdx = 0
    For j = 1 To mLookupCount
        If mLookupWordCounts(j) >= 2 Then
            Dim phrase As String
            Dim toks As Variant
            toks = mLookupWords(j)
            phrase = Join(toks, " ")
            If Len(phrase) > 0 Then
                 If InStr(1, sInput, " " & phrase & " ", vbTextCompare) > 0 Then
                    If mLookupWordCounts(j) > bestLen Then
                        bestLen = mLookupWordCounts(j)
                        bestIdx = j
                    End If
                 End If
            End If
        End If
    Next j
    If bestIdx > 0 Then
        outVal = mLookupPhrases(bestIdx)
        confVal = "Phrase Match"
        CheckMultiWordPhraseMatch = True
        Exit Function
    End If
    CheckMultiWordPhraseMatch = False
End Function
          
Private Function GetBestCategory(ByRef inputWords() As String) As String
    Dim i As Long, t As String
    Dim catScores As Object
    Dim bestCat As String
    Dim bestScore As Double
    Dim cat As Variant, totalFreq As Double, prob As Double
    Set catScores = CreateObject("Scripting.Dictionary")
    bestScore = 0#
    bestCat = ""
    If IsEmptyArray(inputWords) Then Exit Function
    If mTokenCatFreq Is Nothing Then Exit Function
    For i = LBound(inputWords) To UBound(inputWords)
        t = inputWords(i)
        If mTokenCatFreq.Exists(t) Then
            Dim d As Object: Set d = mTokenCatFreq(t)
            totalFreq = 0
            For Each cat In d.keys
                totalFreq = totalFreq + d(cat)
            Next cat
            If totalFreq > 0 Then
                For Each cat In d.keys
                    prob = d(cat) / totalFreq
                    catScores(cat) = catScores(cat) + prob
                Next cat
            End If
        End If
    Next i
    For Each cat In catScores.keys
        If catScores(cat) > bestScore Then
            bestScore = catScores(cat)
            bestCat = cat
        End If
    Next cat
    If bestScore >= 0.5 Then GetBestCategory = bestCat
End Function
          
Private Function IsNonAssetInput(ByVal sNorm As String) As Boolean
    Dim badWords As Variant, bw As Variant
    ' Expanded list to include procedural actions that aren't physical assets
    badWords = Array(" report ", " reports ", " survey ", " surveys ", " audit ", " audits ", _
                     " maintenance plan ", " emergency response ", " invoice ", " permit ", _
                     " warranty ", " fee ", " tax ", " meeting ", " verify ", " check ", _
                     " inspect ", " travel ", " quote ", " discuss ", " review ")
       
    For Each bw In badWords
        If InStr(1, sNorm, CStr(bw), vbTextCompare) > 0 Then
            IsNonAssetInput = True
            Exit Function
        End If
    Next bw
    IsNonAssetInput = False
End Function
          
Public Function GetStandardKey(ByVal txt As String) As String
    Dim s As String
    s = LCase$(txt)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, Chr(9), " ")
    s = Replace(s, Chr(10), " ")
    s = Replace(s, Chr(13), " ")
    s = Replace(s, Chr(173), "")
    s = Replace(s, vbTab, " ")
    s = Application.WorksheetFunction.Trim(s)
    GetStandardKey = s
End Function
          
Public Function GetFilteredInput(ByVal normInput As String) As String
    Dim tokens() As String, out() As String
    Dim i As Long, count As Long
    Dim tok As String
    tokens = Tokenize(normInput)
    If IsEmptyArray(tokens) Then GetFilteredInput = "": Exit Function
    ReDim out(0 To UBound(tokens))
    count = 0
    If mLookupVocab Is Nothing Then
        GetFilteredInput = normInput
        Exit Function
    End If
    For i = LBound(tokens) To UBound(tokens)
        tok = tokens(i)
        If mLookupVocab.Exists(tok) Then
            out(count) = tok
            count = count + 1
        End If
    Next i
    If count = 0 Then
        GetFilteredInput = ""
    Else
        ReDim Preserve out(0 To count - 1)
        GetFilteredInput = Join(out, " ")
    End If
End Function
          
Public Function Tokenize(ByVal text As String) As String()
    Dim parts() As String, tmp() As String, i As Long, count As Long
    parts = Split(text, " ")
    If UBound(parts) = -1 Then
        Tokenize = Split(vbNullString, " ")
        Exit Function
    End If
    ReDim tmp(0 To UBound(parts))
    count = 0
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then tmp(count) = Trim$(parts(i)): count = count + 1
    Next i
    If count = 0 Then
        ReDim tmp(0)
        Tokenize = Split("", " ")
    Else
        ReDim Preserve tmp(0 To count - 1)
        Tokenize = tmp
    End If
End Function
          
Public Function NormalizeAndAlias(ByVal text As String, ByVal aliasDict As Object) As String
    Dim t As String, key As Variant
    t = LCase$(text)
    t = Replace$(t, vbCrLf, " ")
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace(t, "(", " ")
    t = Replace(t, ")", " ")
    t = Replace(t, "[", " ")
    t = Replace(t, "]", " ")
    t = Replace(t, "{", " ")
    t = Replace(t, "}", " ")
             
    If mRegEx Is Nothing Then
        Set mRegEx = CreateObject("VBScript.RegExp")
        With mRegEx
            .Global = True
            .IgnoreCase = True
        End With
    End If
             
    With mRegEx
        .Pattern = "(\d)x(\d)"
        If .Test(t) Then t = .Replace(t, "$1 x $2")
        .Pattern = "([a-z])(\d)"
        If .Test(t) Then t = .Replace(t, "$1 $2")
        .Pattern = "(\d)([a-z])"
        If .Test(t) Then t = .Replace(t, "$1 $2")
    End With
             
    t = Replace(t, "/", " / ")
    t = Replace(t, ",", " ")
    t = Replace(t, ";", " ")
    t = Replace(t, "-", " ")
    t = Replace(t, "_", " ")
    t = Application.WorksheetFunction.Trim(t)
       
    If Not mPhraseMap Is Nothing Then
        For Each key In mPhraseMap.keys
            t = ReplaceWordish(t, CStr(key), CStr(mPhraseMap(key)))
        Next key
        t = Application.WorksheetFunction.Trim(t)
    End If
   
    ' --- NEW: SUFFIX STEMMING LOGIC ---
    Dim stemArr() As String, sIdx As Long
    stemArr = Split(t, " ")
    For sIdx = LBound(stemArr) To UBound(stemArr)
        Dim wBody As String: wBody = stemArr(sIdx)
        If Len(wBody) > 4 Then
            If Right$(wBody, 3) = "ing" Then wBody = Left$(wBody, Len(wBody) - 3)
            If Right$(wBody, 2) = "ed" Then wBody = Left$(wBody, Len(wBody) - 2)
            If Right$(wBody, 2) = "es" Then wBody = Left$(wBody, Len(wBody) - 2)
            If Right$(wBody, 1) = "s" And Right$(wBody, 2) <> "ss" Then wBody = Left$(wBody, Len(wBody) - 1)
        End If
        stemArr(sIdx) = wBody
    Next sIdx
    t = Join(stemArr, " ")
    ' ----------------------------------
   
    ' Standard replacements moved to EnsureDictionaries (mAliasDict defaults)
   
    If Not aliasDict Is Nothing Then
        For Each key In aliasDict.keys
            t = ReplaceWordish(t, CStr(key), CStr(aliasDict(key)))
        Next key
        t = Application.WorksheetFunction.Trim(t)
    End If
        
    ' STRICT ALPHA-ONLY FILTER
    With mRegEx
        .Pattern = "[^a-z ]"
        t = .Replace(t, " ")
    End With
    t = Application.WorksheetFunction.Trim(t)
        
    Dim parts() As String, out() As String, i As Long, tok As String, n As Long
    parts = Split(t, " ")
    ReDim out(0 To UBound(parts))
    n = 0
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If tok <> "" Then
            If Not mProtectSet.Exists(tok) Then
                If mStripHash And Left$(tok, 1) = "#" Then GoTo NextToken
                If mStripNumeric And IsNumeric(tok) Then GoTo NextToken
                If mStripAlphaNum And IsAlphaNumTag(tok) Then GoTo NextToken
            End If
            ' Alias processing
            out(n) = tok: n = n + 1
        End If
NextToken:
    Next i
       
    If n = 0 Then NormalizeAndAlias = "" Else ReDim Preserve out(0 To n - 1): NormalizeAndAlias = Join(out, " ")
End Function
          
Private Function ReplaceWordish(ByVal text As String, ByVal findWhat As String, ByVal replaceWith As String) As String
    Dim padded As String
    padded = " " & text & " "
    padded = Replace$(padded, " " & findWhat & " ", IIf(replaceWith = "", " ", " " & replaceWith & " "))
    ReplaceWordish = Application.WorksheetFunction.Trim(padded)
End Function
          
Private Function Levenshtein(ByVal s1 As String, ByVal s2 As String) As Long
    Dim i As Long, j As Long
    Dim l1 As Long, l2 As Long
    Dim d() As Long
    Dim min1 As Long, min2 As Long, min3 As Long
    l1 = Len(s1)
    l2 = Len(s2)
    If l1 = 0 Then Levenshtein = l2: Exit Function
    If l2 = 0 Then Levenshtein = l1: Exit Function
    ReDim d(l1, l2)
    For i = 0 To l1: d(i, 0) = i: Next i
    For j = 0 To l2: d(0, j) = j: Next j
    For i = 1 To l1
        For j = 1 To l2
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                min3 = d(i - 1, j - 1) + 1
                If min1 < min2 Then
                    If min1 < min3 Then d(i, j) = min1 Else d(i, j) = min3
                Else
                    If min2 < min3 Then d(i, j) = min2 Else d(i, j) = min3
                End If
            End If
        Next j
    Next i
    Levenshtein = d(l1, l2)
End Function
          
Private Function IsFuzzyMatch(ByVal w1 As String, ByVal w2 As String) As Boolean
    If w1 = w2 Then IsFuzzyMatch = True: Exit Function
    Dim l1 As Long, l2 As Long
    l1 = Len(w1): l2 = Len(w2)
    If Abs(l1 - l2) > 2 Then Exit Function
    If l1 < 4 Or l2 < 4 Then Exit Function ' INCREASED FROM 3 TO 4
    Dim dist As Long
    dist = Levenshtein(w1, w2)
    If l1 <= 5 Then
        If dist <= 1 Then IsFuzzyMatch = True
    Else
        If dist <= 2 Then IsFuzzyMatch = True
    End If
End Function
         
Private Function GetFuzzySimilarity(ByVal w1 As String, ByVal w2 As String) As Double
    If w1 = w2 Then GetFuzzySimilarity = 1#: Exit Function
             
    Dim l1 As Long, l2 As Long
    l1 = Len(w1): l2 = Len(w2)
             
    If l1 < 3 Or l2 < 3 Then GetFuzzySimilarity = 0#: Exit Function
             
    Dim maxDist As Long
    If l1 <= 5 Then maxDist = 1 Else maxDist = 2
             
    If Abs(l1 - l2) > maxDist Then GetFuzzySimilarity = 0#: Exit Function
             
    ' Heuristic: First char match for short words
    If (l1 = 3 Or l2 = 3) And Left$(w1, 1) <> Left$(w2, 1) Then GetFuzzySimilarity = 0#: Exit Function
         
    Dim dist As Long
    dist = Levenshtein(w1, w2)
             
    If dist <= maxDist Then
        Dim maxLen As Long
        If l1 > l2 Then maxLen = l1 Else maxLen = l2
        GetFuzzySimilarity = 1# - (CDbl(dist) / CDbl(maxLen))
    Else
        GetFuzzySimilarity = 0#
    End If
End Function
          
Private Function GetConsecutiveMatchCount(ByVal inputWords As Variant, ByVal lookupWords As Variant) As Long
    Dim maxLen As Long
    Dim currentLen As Long
    Dim i As Long, j As Long
    Dim p1 As Long, p2 As Long
    If IsEmptyArray(inputWords) Or IsEmptyArray(lookupWords) Then
        GetConsecutiveMatchCount = 0
        Exit Function
    End If
    maxLen = 0
    For i = LBound(inputWords) To UBound(inputWords)
        For j = LBound(lookupWords) To UBound(lookupWords)
            If inputWords(i) = lookupWords(j) Then
                currentLen = 1
                p1 = i + 1
                p2 = j + 1
                Do While p1 <= UBound(inputWords) And p2 <= UBound(lookupWords)
                    If inputWords(p1) = lookupWords(p2) Then
                        currentLen = currentLen + 1
                        p1 = p1 + 1
                        p2 = p2 + 1
                    Else
                        Exit Do
                    End If
                Loop
                If currentLen > maxLen Then maxLen = currentLen
            End If
        Next j
    Next i
    GetConsecutiveMatchCount = maxLen
End Function
          
Private Function IsMeaningfulToken(ByVal w As String, ByVal stopWords As Object) As Boolean
    w = LCase$(Trim$(w))
    If w = "" Then IsMeaningfulToken = False: Exit Function
    If stopWords.Exists(w) Then IsMeaningfulToken = False: Exit Function
    If Len(w) <= 1 Then IsMeaningfulToken = False: Exit Function
    IsMeaningfulToken = True
End Function
          
Private Function IsAllowedLearnToken(ByVal tok As String) As Boolean
    If mLookupVocab Is Nothing Then IsAllowedLearnToken = False: Exit Function
    IsAllowedLearnToken = mLookupVocab.Exists(tok)
End Function
          
Private Function IsAllowedScoringToken(ByVal tok As String) As Boolean
    ' STRICT MODE:
    ' Only allow words that explicitly exist in the lookup vocabulary.
    ' If "Sector" is not in the lookup data, it returns False and is discarded.
      
    If mLookupVocab Is Nothing Then
        IsAllowedScoringToken = False
        Exit Function
    End If
  
    If mLookupVocab.Exists(tok) Then
        IsAllowedScoringToken = True
    Else
        ' This is the change: Previously, we allowed IsAllAlphaNum() here.
        ' Now, we strictly return False for anything unknown.
        IsAllowedScoringToken = False
    End If
End Function
          
Private Function IsAllAlphaNum(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    s = LCase$(Trim$(s))
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If Not ((ch >= 97 And ch <= 122) Or (ch >= 48 And ch <= 57)) Then IsAllAlphaNum = False: Exit Function
    Next i
    IsAllAlphaNum = True
End Function
          
Private Function IsTagLike(ByVal tok As String) As Boolean
    Dim t As String: t = LCase$(Trim$(tok))
    If t = "" Then Exit Function
    If Left$(t, 1) = "@" Then IsTagLike = True: Exit Function
    If HasAnyDigit(t) Then
        If Right$(t, 1) <> "v" And Right$(t, 2) <> "kv" Then
            If t Like "[a-z]#[#]*" Or t Like "[a-z][a-z]#[#]*" Or t Like "[a-z][a-z][a-z]#[#]*" Then IsTagLike = True: Exit Function
        End If
    End If
End Function
          
Private Function HasAnyDigit(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch >= 48 And ch <= 57 Then HasAnyDigit = True: Exit Function
    Next i
End Function
          
Private Function IsAlphaNumTag(ByVal tok As String) As Boolean
    If tok Like "[a-z][0-9]*" Then IsAlphaNumTag = True: Exit Function
    If tok Like "[a-z].[0-9]*" Then IsAlphaNumTag = True: Exit Function
    If tok Like "[a-z]-[0-9]*" Then IsAlphaNumTag = True: Exit Function
End Function
          
Private Function GetMeaningfulOverlapCount(ByVal inputSet As Object, ByVal lookupWords As Variant, ByVal stopWords As Object) As Long
    Dim i As Long, w As String, c As Long
    If IsEmptyArray(lookupWords) Then GetMeaningfulOverlapCount = 0: Exit Function
    For i = LBound(lookupWords) To UBound(lookupWords)
        w = Trim$(lookupWords(i))
        If IsMeaningfulToken(w, stopWords) Then If inputSet.Exists(w) Then c = c + 1
    Next i
    GetMeaningfulOverlapCount = c
End Function
          
Private Function IsStrongSingleToken(ByVal w As String) As Boolean
    If mCoreSet.Exists(w) Then IsStrongSingleToken = True: Exit Function
    Select Case w
        Case "pump", "fan", "boiler", "furnace", "transformer", "compressor", "chiller", "motor", "sump", "sprinkler"
            IsStrongSingleToken = True: Exit Function
    End Select
    IsStrongSingleToken = False
End Function
          
Public Function ExtractCoreTokens(ByRef words() As String, ByVal coreSet As Object) As Object
    Dim req As Object: Set req = CreateObject("Scripting.Dictionary")
    Dim i As Long, w As String
    If IsEmptyArray(words) Then Set ExtractCoreTokens = req: Exit Function
    For i = LBound(words) To UBound(words)
        w = words(i): If w <> "" Then If coreSet.Exists(w) Then req(w) = True
    Next i
    Set ExtractCoreTokens = req
End Function
          
Public Function ExtractQualifierTokens(ByRef words() As String, ByVal qualSet As Object) As Object
    Dim req As Object: Set req = CreateObject("Scripting.Dictionary")
    Dim i As Long, w As String
    If IsEmptyArray(words) Then Set ExtractQualifierTokens = req: Exit Function
    For i = LBound(words) To UBound(words)
        w = words(i): If w <> "" Then If qualSet.Exists(w) Then req(w) = True
    Next i
    Set ExtractQualifierTokens = req
End Function
          
Public Function ContainsAllTokens(ByVal candDict As Object, ByVal required As Object) As Boolean
    Dim key As Variant
    If required Is Nothing Then ContainsAllTokens = True: Exit Function
    If required.count = 0 Then ContainsAllTokens = True: Exit Function
    For Each key In required.keys
        If Not candDict.Exists(CStr(key)) Then ContainsAllTokens = False: Exit Function
    Next key
    ContainsAllTokens = True
End Function
          
Public Function IsEmptyArray(arr As Variant) As Boolean
    On Error GoTo EH
    Dim lb As Long, ub As Long
    lb = LBound(arr): ub = UBound(arr)
    IsEmptyArray = (ub < lb)
    Exit Function
EH: IsEmptyArray = True
End Function
          
Private Function IsDescriptorToken(ByVal tok As String) As Boolean
    Select Case tok
        Case "north", "south", "east", "west", "wing", "wings", "side", "area", "zone", _
             "room", "rm", "suite", "level", "floor", "fl", "serving", "served", _
             "corridor", "hall", "hallway", "lobby", "mezz", "mezzanine", _
            "upper", "lower", "basement", "penthouse", "ph", "mechanical"
            IsDescriptorToken = True
    End Select
End Function
          
Private Function ShouldCollapseToSingle(ByRef topScore() As Double) As Boolean
    Dim s1 As Double, s2 As Double
    s1 = topScore(1): s2 = topScore(2)
    If s2 = 0 Then ShouldCollapseToSingle = True: Exit Function
    If s1 - s2 >= DOMINANCE_MARGIN Then ShouldCollapseToSingle = True: Exit Function
    ShouldCollapseToSingle = False
End Function
          
Private Sub InsertTop3(ByRef topPhrase() As String, ByRef topScore() As Double, ByVal phrase As String, ByVal score As Double)
    Dim i As Long, p As Long
    If PhraseExists(topPhrase, phrase) Then Exit Sub
    For i = 1 To 3
        If score > topScore(i) Then
            For p = 3 To i + 1 Step -1
                topPhrase(p) = topPhrase(p - 1)
                topScore(p) = topScore(p - 1)
            Next p
            topPhrase(i) = phrase
            topScore(i) = score
            Exit Sub
        End If
    Next i
End Sub
          
Private Function PhraseExists(ByRef arr() As String, ByVal s As String) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then If arr(i) = s Then PhraseExists = True: Exit Function
    Next i
    PhraseExists = False
End Function
          
Private Function CountNonEmpty(ByRef arr() As String) As Long
    Dim i As Long, c As Long
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then c = c + 1
    Next i
    CountNonEmpty = c
End Function
          
Private Function JoinNonEmpty(ByRef s() As String, ByVal sep As String) As String
    Dim i As Long, res As String
    For i = LBound(s) To UBound(s)
        If Len(s(i)) > 0 Then If Len(res) = 0 Then res = s(i) Else res = res & sep & s(i)
    Next i
    JoinNonEmpty = res
End Function
          
Private Sub QuickSortStringsAsc(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, temp As String
    i = first: j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortStringsAsc arr, first, j
    If i < last Then QuickSortStringsAsc arr, i, last
End Sub
          
Private Sub QuickSortStringsLengthDesc(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, temp As String
    i = first: j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While Len(arr(i)) > Len(pivot): i = i + 1: Loop
        Do While Len(arr(j)) < Len(pivot): j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortStringsLengthDesc arr, first, j
    If i < last Then QuickSortStringsLengthDesc arr, i, last
End Sub
          
Private Sub BuildCategoryDictionaryFromLookup(ByVal wsLookup As Worksheet)
    Dim lastRow As Long
    lastRow = wsLookup.Cells(wsLookup.Rows.count, "A").End(xlUp).Row
    Set mCategoryDict = CreateObject("Scripting.Dictionary")
    Set mTokenCatFreq = CreateObject("Scripting.Dictionary")
    Dim keys As New Collection
    If lastRow < 2 Then mCategoryCount = 0: Exit Sub
    Dim r As Long, cat As String, phrase As String
    Dim tokens() As String, t As String, k As Long
    For r = 2 To lastRow
        phrase = CStr(wsLookup.Cells(r, "A").Value)
        cat = LCase$(Trim$(CStr(wsLookup.Cells(r, "C").Value)))
        If Len(cat) > 0 Then
            If Not mCategoryDict.Exists(cat) Then
                mCategoryDict.add cat, CreateObject("Scripting.Dictionary")
                keys.add cat
            End If
            mCategoryDict(cat)(phrase) = True
            Dim norm As String
            norm = NormalizeAndAlias(phrase, mAliasDict)
            tokens = Tokenize(norm)
            If Not IsEmptyArray(tokens) Then
                For k = LBound(tokens) To UBound(tokens)
                    t = tokens(k)
                    If IsMeaningfulToken(t, mStopWords) Then
                        If Not mTokenCatFreq.Exists(t) Then
                             mTokenCatFreq.add t, CreateObject("Scripting.Dictionary")
                        End If
                        Dim d As Object: Set d = mTokenCatFreq(t)
                        d(cat) = d(cat) + 1
                    End If
                Next k
            End If
        End If
    Next r
    mCategoryCount = keys.count
    If mCategoryCount > 0 Then
        ReDim mCategoryKeys(1 To mCategoryCount)
        Dim i As Long
        For i = 1 To mCategoryCount
            mCategoryKeys(i) = CStr(keys(i))
        Next i
    Else
        Erase mCategoryKeys
    End If
    Set mCategoryAliasIndex = CreateObject("Scripting.Dictionary")
    Dim keyMap As Variant
    If Not mPhraseMap Is Nothing Then
        For Each keyMap In mPhraseMap.keys
            Dim mapped As String: mapped = LCase$(CStr(mPhraseMap(keyMap)))
            If Len(mapped) > 0 Then mCategoryAliasIndex(keyMap) = mapped
        Next keyMap
    End If
End Sub
          
Private Function InferForcedCategory(ByVal normalizedInput As String) As String
    Dim sNorm As String: sNorm = " " & LCase$(normalizedInput) & " "
    If mLookupVocab Is Nothing Then If Not mTokenIndex Is Nothing Then BuildLookupVocabulary
    If InStr(1, sNorm, " pump ", vbTextCompare) > 0 Then InferForcedCategory = "pump": Exit Function
    If InStr(1, sNorm, " boiler ", vbTextCompare) > 0 Or InStr(1, sNorm, " blr ", vbTextCompare) > 0 Then InferForcedCategory = "boiler": Exit Function
    If InStr(1, sNorm, " furnace ", vbTextCompare) > 0 Or InStr(1, sNorm, " forced air heater ", vbTextCompare) > 0 Then InferForcedCategory = "furnace": Exit Function
    If InStr(1, sNorm, " fan ", vbTextCompare) > 0 Or InStr(1, sNorm, " exhaust ", vbTextCompare) > 0 Then InferForcedCategory = "fan": Exit Function
    InferForcedCategory = ""
End Function
          
Private Function CandidateContainsCategory(ByVal candDict As Object, ByVal categoryName As String) As Boolean
    CandidateContainsCategory = False
    If candDict Is Nothing Then Exit Function
    Dim catTok As String: catTok = LCase$(Trim$(categoryName))
    If Len(catTok) = 0 Then Exit Function
    If candDict.Exists(catTok) Then CandidateContainsCategory = True
End Function
          
Private Function GetCategoryItemsList(ByVal catName As String) As String
    Dim res As String
    Dim count As Long
    Dim key As Variant
    If mCategoryDict Is Nothing Then Exit Function
    If Not mCategoryDict.Exists(catName) Then Exit Function
    Dim d As Object: Set d = mCategoryDict(catName)
    count = 0
    For Each key In d.keys
        If count < 30 Then
            If res = "" Then res = key Else res = res & vbLf & key
            count = count + 1
        Else
            res = res & vbLf & "... (More items in category)"
            Exit For
        End If
    Next key
    GetCategoryItemsList = res
End Function
          
Private Function GetDominantCategory(ByVal token As String) As String
    If mTokenCatFreq Is Nothing Then Exit Function
    If Not mTokenCatFreq.Exists(token) Then Exit Function
    Dim d As Object: Set d = mTokenCatFreq(token)
    Dim cat As Variant, total As Double, count As Double
    Dim bestCat As String, bestProp As Double
    total = 0
    For Each cat In d.keys
        total = total + d(cat)
    Next cat
    If total = 0 Then Exit Function
    bestProp = 0
    For Each cat In d.keys
        count = d(cat)
        If count / total > bestProp Then
            bestProp = count / total
            bestCat = cat
        End If
    Next cat
    Dim threshold As Double: threshold = 0.7
    Select Case token
        Case "truck", "vehicle", "car", "van": threshold = 0.4
    End Select
    If bestProp >= threshold Then GetDominantCategory = bestCat
End Function
          
Private Function GetCanonicalSignature(ByVal text As String) As String
    Dim tokens() As String
    Dim i As Long, j As Long, temp As String
    tokens = Tokenize(NormalizeAndAlias(text, mAliasDict))
    If IsEmptyArray(tokens) Then GetCanonicalSignature = "": Exit Function
    Dim clean() As String
    Dim count As Long: count = 0
    ReDim clean(0 To UBound(tokens))
     
    ' NEW: Remove duplicates from the signature
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
     
    For i = LBound(tokens) To UBound(tokens)
        If IsMeaningfulToken(tokens(i), mStopWords) Then
            If Not seen.Exists(tokens(i)) Then
                clean(count) = tokens(i)
                count = count + 1
                seen(tokens(i)) = True
            End If
        End If
    Next i
    If count = 0 Then GetCanonicalSignature = "": Exit Function
    ReDim Preserve clean(0 To count - 1)
    For i = LBound(clean) To UBound(clean) - 1
        For j = i + 1 To UBound(clean)
            If clean(i) > clean(j) Then
                temp = clean(i)
                clean(i) = clean(j)
                clean(j) = temp
            End If
        Next j
    Next i
    GetCanonicalSignature = Join(clean, "|")
End Function
         
Private Function GetVocabSignature(ByVal text As String) As String
    Dim tokens() As String
    Dim i As Long, j As Long, temp As String
    tokens = Tokenize(NormalizeAndAlias(text, mAliasDict))
    If IsEmptyArray(tokens) Then GetVocabSignature = "": Exit Function
             
    Dim clean() As String
    Dim count As Long: count = 0
    ReDim clean(0 To UBound(tokens))
             
    If mLookupVocab Is Nothing Then Exit Function
     
    ' NEW: Remove duplicates from the signature
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
             
    For i = LBound(tokens) To UBound(tokens)
        Dim tok As String: tok = tokens(i)
        If IsMeaningfulToken(tok, mStopWords) Then
            If mLookupVocab.Exists(tok) Then
                If Not seen.Exists(tok) Then
                    clean(count) = tok
                    count = count + 1
                    seen(tok) = True
                End If
            End If
        End If
    Next i
             
    If count = 0 Then GetVocabSignature = "": Exit Function
    ReDim Preserve clean(0 To count - 1)
             
    ' Simple Bubble Sort (for short arrays usually < 10 items)
    For i = LBound(clean) To UBound(clean) - 1
        For j = i + 1 To UBound(clean)
            If clean(i) > clean(j) Then
                temp = clean(i)
                clean(i) = clean(j)
                clean(j) = temp
            End If
        Next j
    Next i
             
    GetVocabSignature = Join(clean, "|")
End Function
         
' ==========================================================
' MASTER MATCHER MODULE - PART 3 of 3
' ==========================================================
          
Private Function BuildLearnSignature(ByVal normalizedInput As String) As String
    Dim words() As String, keep() As String
    Dim i As Long, n As Long, tok As String
    Dim wsL As Worksheet
          
    If mLookupVocab Is Nothing Then
        Set wsL = GetLookupSheet()
        If Not wsL Is Nothing Then
            If mAliasDict Is Nothing Then EnsureDictionaries
            BuildLookupCache wsL
        End If
    End If
          
    words = Tokenize(normalizedInput)
    If IsEmptyArray(words) Then BuildLearnSignature = "": Exit Function
              
    ReDim keep(0 To UBound(words))
    n = 0
    For i = LBound(words) To UBound(words)
        tok = LCase$(Trim$(words(i)))
        If tok = "" Then GoTo NextTok
        If Not IsMeaningfulToken(tok, mStopWords) Then GoTo NextTok
        If IsDescriptorToken(tok) Then GoTo NextTok
        ' If Not IsAllAlphaNum(tok) Then GoTo NextTok ' Relaxed to allow hyphenated tags
        If IsTagLike(tok) Then GoTo NextTok
        
        ' Allow token if it is in vocab OR if it is numeric (to distinguish Pump 1 from Pump 2)
        If Not IsAllowedLearnToken(tok) Then
            If Not IsNumeric(tok) Then GoTo NextTok
        End If
                  
        keep(n) = tok
        n = n + 1
NextTok:
    Next i
    If n = 0 Then BuildLearnSignature = "": Exit Function
    ReDim Preserve keep(0 To n - 1)
    Call QuickSortStringsAsc(keep, 0, UBound(keep))
    BuildLearnSignature = Join(keep, "|")
End Function
          
Private Sub EnsureDictionaries()
    Set mAliasDict = CreateObject("Scripting.Dictionary")
    Set mStopWords = CreateObject("Scripting.Dictionary")
    Set mBoostDict = CreateObject("Scripting.Dictionary")
    Set mCoreSet = CreateObject("Scripting.Dictionary")
    Set mQualSet = CreateObject("Scripting.Dictionary")
    Set mPhraseMap = CreateObject("Scripting.Dictionary")
    Set mForcedOutputDict = CreateObject("Scripting.Dictionary")
    Set mProtectSet = CreateObject("Scripting.Dictionary")
    Set mAliasPrefix = CreateObject("Scripting.Dictionary")
    Set mAliasSuffix = CreateObject("Scripting.Dictionary")
          
    ' Default Stop Words
    mStopWords("a") = True
    mStopWords("an") = True
    mStopWords("the") = True
    mStopWords("and") = True
    mStopWords("of") = True
    mStopWords("in") = True
    mStopWords("with") = True
    mStopWords("for") = True
    mStopWords("to") = True
    mStopWords("is") = True
    mStopWords("are") = True
    mStopWords("be") = True
    mStopWords("was") = True
    mStopWords("were") = True
    mStopWords("at") = True
    mStopWords("by") = True
    mStopWords("from") = True
    mStopWords("on") = True
    mStopWords("as") = True
    mStopWords("this") = True
    mStopWords("that") = True
    mStopWords("it") = True
    mStopWords("or") = True
             
    ' Default Synonyms
    ' mAliasDict("water") = "hydronic"
    mAliasDict("hp") = "horsepower"
    mAliasDict("kw") = "kilowatts"
    mAliasDict("kva") = "kilovoltampere"
    mAliasDict("v") = "volts"
    mAliasDict("a") = "amps"
    mAliasDict("gpm") = "gallons per minute"
    
    mAliasDict("h2o") = "hydronic"
    mAliasDict("restroom") = "lavatories"
    mAliasDict("restrooms") = "lavatories"
    mAliasDict("bathroom") = "lavatories"
    mAliasDict("bathrooms") = "lavatories"
    mAliasDict("womans") = "women"
    mAliasDict("lighting") = "light"
    mAliasDict("monitoring") = "monitor"
    mAliasDict("make") = "maker"
    mAliasDict("roll up") = "rollup"
    mAliasDict("rollup") = "rollup"
    mAliasDict("water meter") = "water flow meter"
    mAliasDict("air conditioning systems equipment") = "terminal and package units"
    mAliasDict("refrigerant gas") = "refrigerant"
    mAliasDict("ahu") = "air handling unit"
    mAliasDict("rtu") = "roof top unit"
    mAliasDict("vfd") = "variable frequency drive"
    mAliasDict("mau") = "makeup air unit"
    mAliasDict("fcu") = "fan coil unit"
    mAliasDict("vav") = "variable air volume"
    mAliasDict("dwh") = "domestic water heater"
          
    mStripNumeric = False: mStripHash = False: mStripAlphaNum = False
          
    ' Load rules from external sheet
    SyncRulesWithExternal ThisWorkbook, "Rules Sheet"
    LoadRulesFromSheet ThisWorkbook, "Rules Sheet", _
        mAliasDict, mPhraseMap, mStopWords, mBoostDict, mCoreSet, mQualSet, mForcedOutputDict, _
        mProtectSet, mAliasPrefix, mAliasSuffix, _
        mStripNumeric, mStripHash, mStripAlphaNum
           
' Populate mCoreSet with defaults if empty
         If mCoreSet.count = 0 Then
             Dim defaults As Variant
             Dim def As Variant
             ' OLD LINE:
             ' defaults = Array("pump", "fan", "boiler", "furnace", "transformer", "compressor", "chiller", "motor", "sump", "sprinkler", "valve")
               
             ' NEW LINE (Added "hydrant"):
             defaults = Array("pump", "fan", "boiler", "furnace", "transformer", _
                              "compressor", "chiller", "motor", "sump", "sprinkler", _
                              "valve", "hydrant", "tower", "exchanger", "tank")
               
             For Each def In defaults
                 mCoreSet(def) = True
             Next def
        End If
End Sub
          
Private Sub BuildLookupCache(ByVal wsLookup As Worksheet)
    Dim lastLookupRow As Long
    lastLookupRow = wsLookup.Cells(wsLookup.Rows.count, "A").End(xlUp).Row
    If lastLookupRow < 2 Then mLookupCount = 0: Exit Sub
        
    ' 1. LOAD THE MAIN LOOKUP DATA
    Dim lookupArr As Variant
    lookupArr = wsLookup.Range("A2:C" & lastLookupRow).Value
        
    Dim j As Long, k As Long
    mLookupCount = UBound(lookupArr, 1)
            
    ReDim mLookupPhrases(1 To mLookupCount)
    ReDim mLookupWords(1 To mLookupCount)
    ReDim mLookupDicts(1 To mLookupCount)
    ReDim mRowCategory(1 To mLookupCount)
    ReDim mLookupWordCounts(1 To mLookupCount)
    ReDim mLookupSubjects(1 To mLookupCount)
        
    Set mTokenIndex = CreateObject("Scripting.Dictionary")
    Set mLookupPhraseSet = CreateObject("Scripting.Dictionary")
    Set mSignatureDict = CreateObject("Scripting.Dictionary")
    mLookupPhraseSet.CompareMode = vbTextCompare
    mSignatureDict.CompareMode = vbTextCompare
            
    Dim raw As String, normalized As String, tokens() As String, d As Object, tok As String
    Dim scrubbedRaw As String, sig As String
    Dim filteredToks() As String
    Dim ftCount As Long
        
    For j = 1 To mLookupCount
        If Not IsError(lookupArr(j, 3)) Then mRowCategory(j) = LCase$(Trim$(CStr(lookupArr(j, 3))))
        If Not IsError(lookupArr(j, 1)) Then
            raw = CStr(lookupArr(j, 1))
            mLookupPhrases(j) = raw
                  
            scrubbedRaw = GetStandardKey(raw)
            If Len(scrubbedRaw) > 0 Then mLookupPhraseSet(scrubbedRaw) = True
                  
            ' Normalize the Name
            normalized = NormalizeAndAlias(raw, mAliasDict)
                  
            ' Tokenize directly without Data Enrichment bloat
            tokens = Tokenize(normalized)
                  
            ' Create Canonical Signature
            sig = GetCanonicalSignature(raw)
            If Len(sig) > 0 Then
                If Not mSignatureDict.Exists(sig) Then mSignatureDict(sig) = raw
            End If
                   
            Set d = CreateObject("Scripting.Dictionary")
            ReDim filteredToks(0 To UBound(tokens))
            ftCount = 0
                    
            If Not IsEmptyArray(tokens) Then
                For k = LBound(tokens) To UBound(tokens)
                    tok = tokens(k)
                    If Len(tok) > 0 Then
                        d(tok) = True
                        filteredToks(ftCount) = tok
                        ftCount = ftCount + 1
                                
                        Dim coll As Collection
                        If mTokenIndex.Exists(tok) Then
                            Set coll = mTokenIndex(tok)
                        Else
                            Set coll = New Collection
                            mTokenIndex.add tok, coll
                        End If
                        coll.add j
                    End If
                Next k
            End If
            Set mLookupDicts(j) = d
                    
            If ftCount > 0 Then
                ReDim Preserve filteredToks(0 To ftCount - 1)
                mLookupWords(j) = filteredToks
                mLookupWordCounts(j) = ftCount
                ' Use Improved Subject Detection during build time
                mLookupSubjects(j) = GetInputSubject(filteredToks)
            Else
                mLookupWords(j) = Array()
                mLookupWordCounts(j) = 0
                mLookupSubjects(j) = ""
            End If
        Else
            mLookupPhrases(j) = ""
            mLookupWords(j) = Array()
            Set mLookupDicts(j) = CreateObject("Scripting.Dictionary")
            mLookupWordCounts(j) = 0
        End If
    Next j
        
    BuildCategoryDictionaryFromLookup wsLookup
    BuildTokenWeights
    BuildLookupVocabulary
End Sub
      
          
Private Sub BuildTokenWeights()
    Dim k As Variant, count As Long
    Dim w As Double
    Set mTokenWeights = CreateObject("Scripting.Dictionary")
              
    If mTokenIndex Is Nothing Then Exit Sub
    If mLookupCount = 0 Then Exit Sub
              
    For Each k In mTokenIndex.keys
        count = mTokenIndex(k).count
        w = 1 + Log(mLookupCount / (count + 1))
           
        ' Manual Boost for Core Tokens (prevents Rare Word Dominance)
        If mCoreSet.Exists(k) Then w = w * 2.5
           
        If w < 0.1 Then w = 0.1
        mTokenWeights(CStr(k)) = w
    Next k
End Sub
          
Private Sub BuildLookupVocabulary()
    Dim k As Variant
    Set mLookupVocab = CreateObject("Scripting.Dictionary")
    If mTokenIndex Is Nothing Then Exit Sub
    For Each k In mTokenIndex.keys
        mLookupVocab(CStr(k)) = True
    Next k
End Sub
          
Private Sub CreateSelectionListForSingleRow(ByVal wsInput As Worksheet, ByVal rowNum As Long)
    Dim wsList As Worksheet, cell As Range
    Dim parts() As String, raw As String
    Dim listName As String, itemCount As Long
    Dim rngList As Range
          
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(PICK_SHEET_NAME)
    On Error GoTo 0
    If wsList Is Nothing Then
        Set wsList = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsList.Name = PICK_SHEET_NAME
    End If
    wsList.Visible = xlSheetHidden
          
    Set cell = wsInput.Cells(rowNum, "B")
    raw = CStr(cell.Value2)
    If Len(raw) = 0 Then
        On Error Resume Next: cell.Validation.Delete: On Error GoTo 0
        Exit Sub
    End If
              
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    parts = Split(raw, vbLf)
    itemCount = UBound(parts) - LBound(parts) + 1
          
    listName = "PickRow_" & rowNum
    If itemCount >= 2 Then
        Set rngList = GetOrCreateListRange(wsList, listName, itemCount)
        Dim i As Long, p As Long: p = LBound(parts)
        For i = 1 To itemCount
            rngList.Cells(i, 1).Value = Trim$(parts(p))
            p = p + 1
        Next i
                  
        With cell.Validation
            .Delete
            .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
        End With
    Else
        On Error Resume Next: cell.Validation.Delete: On Error GoTo 0
        On Error Resume Next: ThisWorkbook.names(listName).Delete: On Error GoTo 0
    End If
End Sub
          
Private Function GetOrCreateListRange(ByVal wsList As Worksheet, ByVal listName As String, ByVal itemCount As Long) As Range
    Dim nm As Name, rng As Range
    On Error Resume Next
    Set nm = ThisWorkbook.names(listName)
    On Error GoTo 0
              
    If Not nm Is Nothing Then
        On Error Resume Next: Set rng = nm.RefersToRange: On Error GoTo 0
        If Not rng Is Nothing Then
             If rng.Parent.Name = wsList.Name And rng.Rows.count = itemCount Then
                Set GetOrCreateListRange = rng
                Exit Function
             End If
        End If
    End If
          
    Dim startRow As Long, endRow As Long
    startRow = wsList.Cells(wsList.Rows.count, "A").End(xlUp).Row
    If startRow < 1 Then startRow = 1
    If Len(wsList.Cells(startRow, "A").Value) > 0 Then startRow = startRow + 1
    endRow = startRow + itemCount - 1
          
    Set rng = wsList.Range(wsList.Cells(startRow, "A"), wsList.Cells(endRow, "A"))
              
    On Error Resume Next: ThisWorkbook.names(listName).Delete: On Error GoTo 0
    ThisWorkbook.names.add Name:=listName, RefersTo:=rng
    Set GetOrCreateListRange = rng
End Function
          
Private Sub CreateSelectionListsForColumnB(ByVal wsInput As Worksheet, ByVal n As Long)
    Dim wsList As Worksheet
    Dim i As Long, cell As Range, parts() As String
    Dim listStartRow As Long, listEndRow As Long
    Dim rngList As Range, listName As String
    Dim raw As String
          
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(PICK_SHEET_NAME)
    On Error GoTo 0
    If wsList Is Nothing Then
        Set wsList = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsList.Name = PICK_SHEET_NAME
    End If
    wsList.Visible = xlSheetHidden
    wsList.Cells.Clear
          
    listStartRow = 1
    For i = 2 To (n + 1)
        Set cell = wsInput.Cells(i, "B")
        raw = CStr(cell.Value2)
        If Len(raw) > 0 Then
            raw = Replace(raw, vbCrLf, vbLf)
            raw = Replace(raw, vbCr, vbLf)
            parts = Split(raw, vbLf)
            If (UBound(parts) - LBound(parts) + 1) >= 2 Then
                listEndRow = listStartRow + (UBound(parts) - LBound(parts))
                Dim r As Long, p As Long
                p = LBound(parts)
                For r = listStartRow To listEndRow
                    wsList.Cells(r, "A").Value = Trim$(parts(p))
                    p = p + 1
                Next r
                          
                listName = "PickRow_" & i
                On Error Resume Next: ThisWorkbook.names(listName).Delete: On Error GoTo 0
                Set rngList = wsList.Range(wsList.Cells(listStartRow, "A"), wsList.Cells(listEndRow, "A"))
                ThisWorkbook.names.add Name:=listName, RefersTo:=rngList
                          
                With cell.Validation
                    .Delete
                    .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowError = True
                End With
                listStartRow = listEndRow + 1
            Else
                On Error Resume Next: cell.Validation.Delete: On Error GoTo 0
            End If
        Else
            On Error Resume Next: cell.Validation.Delete: On Error GoTo 0
        End If
    Next i
End Sub
          
Private Sub LoadLearnedOverrides()
    Dim ws As Worksheet, lastRow As Long, arr As Variant
    Dim i As Long, normIn As String, outP As String, sig As String
    Dim times As Long
    Dim vSig As String
          
    If mLearnedLoaded Then Exit Sub
    EnsureLearnedSheet
    Set ws = ThisWorkbook.Worksheets(LEARNED_SHEET)
          
    Set mLearnedExact = CreateObject("Scripting.Dictionary")
    Set mLearnedSigBest = CreateObject("Scripting.Dictionary")
    Set mLearnedSigCount = CreateObject("Scripting.Dictionary")
    Set mLearnedVocabSigBest = CreateObject("Scripting.Dictionary")
          
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        mLearnedLoaded = True
        Exit Sub
    End If
          
    arr = ws.Range("A2:E" & lastRow).Value2
    For i = 1 To UBound(arr, 1)
        normIn = CStr(arr(i, 1))
        outP = CStr(arr(i, 2))
        sig = CStr(arr(i, 3))
        times = CLng(val(arr(i, 4)))
          
        If Len(normIn) > 0 And Len(outP) > 0 Then
            mLearnedExact(normIn) = outP
                     
            vSig = GetVocabSignature(normIn)
            If Len(vSig) > 0 Then
                mLearnedVocabSigBest(vSig) = outP
            End If
        End If
        If Len(sig) > 0 And Len(outP) > 0 Then
            If Not mLearnedSigBest.Exists(sig) Then
                mLearnedSigBest(sig) = outP
                mLearnedSigCount(sig) = times
            Else
                 If times > CLng(mLearnedSigCount(sig)) Then
                    mLearnedSigBest(sig) = outP
                    mLearnedSigCount(sig) = times
                End If
            End If
        End If
    Next i
    MigrateLocalLearningsToExternal
    ImportExternalData
        
    If mLearnedExact.count > 0 Then
        Dim k As Long, key As Variant
        ReDim mLearnedSubstKeys(0 To mLearnedExact.count - 1)
        k = 0
        For Each key In mLearnedExact.keys
            mLearnedSubstKeys(k) = CStr(key)
            k = k + 1
        Next key
        QuickSortStringsLengthDesc mLearnedSubstKeys, LBound(mLearnedSubstKeys), UBound(mLearnedSubstKeys)
    Else
        Erase mLearnedSubstKeys
    End If
        
    mLearnedLoaded = True
End Sub
          
Private Sub EnsureLearnedSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEARNED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = LEARNED_SHEET
        ws.Range("A1:E1").Value = Array("NormalizedInput", "OutputPhrase", "Signature", "TimesUsed", "LastUpdated")
        ws.Columns("A:E").EntireColumn.AutoFit
    End If
    ws.Visible = xlSheetHidden
End Sub
          
Private Function GetExternalFilePath() As String
    GetExternalFilePath = Environ("APPDATA") & "\Uniformat_Learned.txt"
End Function
          
Private Sub AppendToExternalFile(ByVal normIn As String, ByVal outP As String, ByVal sig As String)
    ' We no longer append. We redirect all save requests to the clean sync method.
    MigrateLocalLearningsToExternal
End Sub
          
Private Sub MigrateLocalLearningsToExternal()
    Dim ws As Worksheet, lastRow As Long, arr As Variant
    Dim i As Long, normIn As String, outP As String, sig As String
    Dim p As String, fNum As Integer
    Dim lineStr As String, parts() As String
    Dim diskMap As Object, key As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEARNED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    p = GetExternalFilePath()
    Set diskMap = CreateObject("Scripting.Dictionary")
    
    ' 1. Read existing external file into memory (if exists)
    If Dir(p) <> "" Then
        fNum = FreeFile
        Open p For Input As #fNum
        Do While Not EOF(fNum)
            Line Input #fNum, lineStr
            If Len(lineStr) > 0 Then
                parts = Split(lineStr, "|")
                If UBound(parts) >= 1 Then
                    ' Key = NormalizedInput
                    key = parts(0)
                    diskMap(key) = lineStr
                End If
            End If
        Loop
        Close #fNum
    End If
    
    ' 2. Merge local data (overwrites external if collision, or we could check timestamps)
    ' Current logic: Local wins (since user just made a choice)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        arr = ws.Range("A2:C" & lastRow).Value2
        For i = 1 To UBound(arr, 1)
            normIn = CStr(arr(i, 1))
            outP = CStr(arr(i, 2))
            sig = CStr(arr(i, 3))
            If Len(normIn) > 0 And Len(outP) > 0 Then
                key = normIn
                diskMap(key) = normIn & "|" & outP & "|" & sig & "|" & Now
            End If
        Next i
    End If
    
    ' 3. Write back to external file
    fNum = FreeFile
    Open p For Output As #fNum
    Dim k As Variant
    For Each k In diskMap.keys
        Print #fNum, diskMap(k)
    Next k
    Close #fNum
End Sub
          
Private Sub ImportExternalData()
    Dim p As String, fNum As Integer
    Dim lineStr As String, parts() As String
    Dim normIn As String, outP As String, sig As String
    Dim cacheKey As String
    If Not mExternalFileCache Is Nothing Then Exit Sub
    Set mExternalFileCache = CreateObject("Scripting.Dictionary")
    p = GetExternalFilePath()
    If Dir(p) = "" Then Exit Sub
    If mLearnedExact Is Nothing Then Set mLearnedExact = CreateObject("Scripting.Dictionary")
    If mLearnedSigBest Is Nothing Then Set mLearnedSigBest = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    fNum = FreeFile
    Open p For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, lineStr
        If Len(lineStr) > 0 Then
            parts = Split(lineStr, "|")
            If UBound(parts) >= 2 Then
                normIn = parts(0)
                outP = parts(1)
                sig = parts(2)
                cacheKey = normIn & "|" & outP & "|" & sig
                mExternalFileCache(cacheKey) = True
                If Len(normIn) > 0 And Len(outP) > 0 Then
                    mLearnedExact(normIn) = outP
                End If
                If Len(sig) > 0 And Len(outP) > 0 Then
                    mLearnedSigBest(sig) = outP
                End If
            End If
        End If
    Loop
    Close #fNum
    On Error GoTo 0
End Sub
          
Private Function GetExternalRulesFilePath() As String
    GetExternalRulesFilePath = Environ("APPDATA") & "\Uniformat_Rules.txt"
End Function
          
Private Sub SyncRulesWithExternal(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim p As String, lineStr As String
    Dim fNum As Integer
    Dim r As Long, lastRow As Long
     
    ' 1. Locate the Rules Sheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
     
    ' 2. Get the file path for the backup text file
    p = GetExternalRulesFilePath()
     
    ' 3. Find the last row of data in the Excel sheet
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
     
    ' 4. Open the text file to overwrite it entirely (Output mode clears the file)
    fNum = FreeFile
    Open p For Output As #fNum
     
    ' 5. Loop through the Excel sheet and write every rule to the text file
    If lastRow >= 2 Then
        For r = 2 To lastRow
            lineStr = Trim$(CStr(ws.Cells(r, "A").Value)) & vbTab & _
                      Trim$(CStr(ws.Cells(r, "B").Value)) & vbTab & _
                      Trim$(CStr(ws.Cells(r, "C").Value)) & vbTab & _
                      Trim$(CStr(ws.Cells(r, "D").Value))
            Print #fNum, lineStr
        Next r
    End If
     
    ' 6. Close and release the file
    Close #fNum
End Sub
          
Private Sub LoadRulesFromSheet(ByVal wb As Workbook, ByVal sheetName As String, _
    ByVal aliasDict As Object, ByVal phraseMap As Object, _
    ByVal stopWords As Object, ByVal boostDict As Object, _
    ByVal coreSet As Object, ByVal qualSet As Object, _
    ByVal forcedOut As Object, ByVal protectSet As Object, _
    ByVal aliasPrefix As Object, ByVal aliasSuffix As Object, _
    ByRef stripNumeric As Boolean, ByRef stripHash As Boolean, ByRef stripAlnum As Boolean)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(sheetName): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long, r As Long, typ As String
    Dim fromText As String, toText As String, wt As Variant
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    For r = 2 To lastRow
        typ = LCase$(Application.WorksheetFunction.Trim(CStr(ws.Cells(r, "A").Value)))
        fromText = GetStandardKey(CStr(ws.Cells(r, "B").Value))
        toText = CStr(ws.Cells(r, "C").Value)
        wt = ws.Cells(r, "D").Value
        If typ <> "" Then
            Select Case typ
                Case "alias": aliasDict(fromText) = LCase$(Trim$(toText))
                Case "strip": aliasDict(fromText) = ""
                Case "phrase": phraseMap(fromText) = LCase$(Trim$(toText))
                Case "stop": stopWords(fromText) = True
                Case "boost": If IsNumeric(wt) Then boostDict(fromText) = CDbl(wt)
                Case "core": coreSet(fromText) = True
                Case "qual": qualSet(fromText) = True
                Case "list": forcedOut(fromText) = NormalizeNewlines(toText)
                Case "protect": protectSet(fromText) = True
                Case "alias_prefix": aliasPrefix(fromText) = LCase$(Trim$(toText))
                Case "alias_suffix": aliasSuffix(fromText) = LCase$(Trim$(toText))
                Case "strip_numeric": stripNumeric = True
                Case "strip_hash": stripHash = True
                Case "strip_alnum": stripAlnum = True
            End Select
        End If
    Next r
End Sub
          
Private Function NormalizeNewlines(ByVal s As String) As String
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    NormalizeNewlines = s
End Function
          
' ==========================================================
' MISSING FUNCTION: TryLearnedOverride (Paste at Bottom)
' ==========================================================
          
Private Function SubstituteLearnedPhrases(ByVal normInput As String) As String
    Dim i As Long
    Dim key As String, outVal As String
    Dim res As String
        
    res = normInput
        
    If IsEmptyArray(mLearnedSubstKeys) Then
        SubstituteLearnedPhrases = res
        Exit Function
    End If
        
    For i = LBound(mLearnedSubstKeys) To UBound(mLearnedSubstKeys)
        key = mLearnedSubstKeys(i)
        ' Optimization: Check if key exists before trying replace
        If InStr(1, res, key, vbTextCompare) > 0 Then
            outVal = mLearnedExact(key)
            ' We use ReplaceWordish to ensure we don't match parts of words (e.g. "he" in "the")
            ' But we must be careful: ReplaceWordish expects clean input.
            res = ReplaceWordish(res, key, outVal)
        End If
    Next i
        
    SubstituteLearnedPhrases = res
End Function
          
Private Function TryLearnedOverride(ByVal normalizedInput As String, ByRef outVal As String, ByRef confVal As String) As Boolean
    Dim sig As String
              
    ' Ensure data is loaded
    LoadLearnedOverrides
          
    ' 1. Check Exact Match
    If Not mLearnedExact Is Nothing Then
        If mLearnedExact.Exists(normalizedInput) Then
            outVal = SanitizeLearnedOutput(mLearnedExact(normalizedInput))
            If IsValidOutputList(outVal) Then
               confVal = "Learned"
                TryLearnedOverride = True
                Exit Function
            End If
        End If
    End If
          
    ' 2. Check Signature Match (Scrambled words)
    sig = BuildLearnSignature(normalizedInput)
    If Len(sig) > 0 Then
        If Not mLearnedSigBest Is Nothing Then
            If mLearnedSigBest.Exists(sig) Then
                outVal = SanitizeLearnedOutput(mLearnedSigBest(sig))
                If IsValidOutputList(outVal) Then
                    confVal = "Learned"
                    TryLearnedOverride = True
                    Exit Function
                End If
            End If
        End If
    End If
             
    ' 3. Check Vocab Signature Match (Learned)
    If Not mLearnedVocabSigBest Is Nothing Then
        Dim vSig As String
        vSig = GetVocabSignature(normalizedInput)
        If Len(vSig) > 0 Then
            If mLearnedVocabSigBest.Exists(vSig) Then
                outVal = SanitizeLearnedOutput(mLearnedVocabSigBest(vSig))
                If IsValidOutputList(outVal) Then
                    confVal = "Learned (Vocab)"
                    TryLearnedOverride = True
                    Exit Function
                End If
            End If
        End If
    End If
              
    TryLearnedOverride = False
End Function
       
Public Sub DiagnoseMatcher()
    Dim ws As Worksheet
    Dim msg As String
           
    ' 1. Set the worksheet (Verify name matches your tab exactly)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Uniformat RS Means Lookup")
    On Error GoTo 0
           
    If ws Is Nothing Then
        MsgBox "CRITICAL ERROR: Sheet 'Uniformat RS Means Lookup' not found!"
        Exit Sub
    End If
           
    ' 2. Initialize the Dictionary (This builds mTokenIndex)
    ' This calls the Private Sub inside this module
    Call InitializeMatcher(ws)
           
    ' 3. Check if the dictionary is actually alive
    If mTokenIndex Is Nothing Then
        MsgBox "CRITICAL ERROR: mTokenIndex failed to initialize."
        Exit Sub
    End If
           
    ' 4. Test specific words to see if the Matcher knows them
    msg = "Dictionary Status:" & vbCrLf
    msg = msg & "Total Words Known: " & mTokenIndex.count & vbCrLf
    msg = msg & "--------------------------------" & vbCrLf
    msg = msg & "Knows 'tank'? " & mTokenIndex.Exists("tank") & vbCrLf
    msg = msg & "Knows 'pump'? " & mTokenIndex.Exists("pump") & vbCrLf
    msg = msg & "Knows 'water'? " & mTokenIndex.Exists("water") & vbCrLf
    msg = msg & "Knows 'barrie'? " & mTokenIndex.Exists("barrie") & vbCrLf
    msg = msg & "Knows 'dom'? " & mTokenIndex.Exists("dom") & " (If False, aliases needed)"
           
    MsgBox msg
End Sub
   
  
' ==========================================================
' MISSING FUNCTION: FindClosestVocabMatch (Added)
' ==========================================================
  
Private Function FindClosestVocabMatch(ByVal token As String) As String
    ' Returns the closest match from mLookupVocab if it meets a similarity threshold
    ' This is a "Typo Watchdog" helper
      
    If mLookupVocab Is Nothing Then Exit Function
    If mLookupVocab.Exists(token) Then
        FindClosestVocabMatch = token
        Exit Function
    End If
      
    ' Optimization: Don't scan if token is too short
    If Len(token) < 3 Then Exit Function
      
    Dim bestMatch As String
    Dim bestScore As Double
    Dim key As Variant
    Dim score As Double
    Dim sKey As String
      
    bestScore = 0#
      
    For Each key In mLookupVocab.keys
        sKey = CStr(key)
        ' Optimization: Only check words of similar length (+/- 2 chars)
        If Abs(Len(sKey) - Len(token)) <= 2 Then
            ' Use the existing GetFuzzySimilarity function
            score = GetFuzzySimilarity(token, sKey)
            If score > bestScore Then
                bestScore = score
                bestMatch = sKey
            End If
        End If
    Next key
      
    ' Threshold for correction: Must be a very strong match (e.g. > 0.8) to risk auto-correction
    If bestScore >= 0.8 Then
        FindClosestVocabMatch = bestMatch
    End If
End Function
  
  
' NEW HELPER FUNCTIONS
' ==========================================================
   
' Improved Subject Detection (Looks for last CORE Token)
Private Function GetInputSubject(ByVal tokens As Variant) As String
    If IsEmptyArray(tokens) Then Exit Function
    If mCoreSet Is Nothing Then
        GetInputSubject = tokens(UBound(tokens))
        Exit Function
    End If
       
    Dim i As Long
    ' Scan backwards
    For i = UBound(tokens) To LBound(tokens) Step -1
        Dim t As String: t = tokens(i)
        If mCoreSet.Exists(t) Then
            GetInputSubject = t
            Exit Function
        End If
    Next i
       
    ' Fallback to last word if no core token found
    GetInputSubject = tokens(UBound(tokens))
End Function
   
' Bi-gram Similarity
Private Function GetBigramSimilarity(ByVal inputToks As Variant, ByVal targetToks As Variant) As Double
    If UBound(inputToks) < 1 Or UBound(targetToks) < 1 Then Exit Function
       
    Dim inputGrams As Object: Set inputGrams = CreateObject("Scripting.Dictionary")
    Dim targetGrams As Object: Set targetGrams = CreateObject("Scripting.Dictionary")
       
    Dim i As Long, s As String
    For i = LBound(inputToks) To UBound(inputToks) - 1
        s = inputToks(i) & "|" & inputToks(i + 1)
        inputGrams(s) = True
    Next i
    For i = LBound(targetToks) To UBound(targetToks) - 1
        s = targetToks(i) & "|" & targetToks(i + 1)
        targetGrams(s) = True
    Next i
       
    Dim matchCount As Long: matchCount = 0
    Dim k As Variant
    For Each k In inputGrams.keys
        If targetGrams.Exists(k) Then matchCount = matchCount + 1
    Next k
       
    If matchCount = 0 Then Exit Function
       
    ' Dice Coefficient for Bigrams
    Dim totalGrams As Long
    totalGrams = inputGrams.count + targetGrams.count
       
    If totalGrams > 0 Then
        GetBigramSimilarity = (2 * matchCount) / totalGrams
    End If
End Function
Public Sub PropagateLearnedChoice(ByVal ws As Worksheet, ByVal sourceRow As Long, ByVal targetSig As String, ByVal chosenOut As String)
    Dim lastRow As Long, r As Long
    Dim raw As String, norm As String, rowSig As String
     
    If Len(targetSig) = 0 Or Len(chosenOut) = 0 Then Exit Sub
     
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
     
    On Error GoTo SafeExit ' Ensure we re-enable events
    Application.ScreenUpdating = False
    Application.EnableEvents = False ' Pauses events so we don't trigger an infinite loop
     
    For r = 2 To lastRow
        If r <> sourceRow Then
            raw = CStr(ws.Cells(r, "A").Value2)
            If Len(Trim$(raw)) > 0 Then
                ' Generate the signature to see if it has the "same meaning"
                norm = NormalizeAndAlias(raw, mAliasDict)
                rowSig = BuildLearnSignature(norm)
                 
                ' If the signatures match (e.g., DHWT-1 and DHWT-2), update it!
                If rowSig = targetSig Then
                    ws.Cells(r, "B").Value2 = chosenOut
                    ws.Cells(r, "C").Value2 = "Auto-Propagated"
                     
                    ' Delete the dropdown list since you've made the final decision
                    On Error Resume Next
                    ws.Cells(r, "B").Validation.Delete
                    On Error GoTo 0
                End If
            End If
        End If
    Next r
     
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then MsgBox "Error propagating choice: " & Err.Description
End Sub



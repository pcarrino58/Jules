Attribute VB_Name = "AIStaging"
' ============================================================
' AI STAGING & PRE-FLIGHT SUITE
' ============================================================

Public Sub PrepAndRunAIPreFlight()
    Dim wsHelper As Worksheet, wsStaging As Worksheet
    Dim lastRowHelper As Long
    
    On Error Resume Next
    Set wsHelper = ThisWorkbook.Worksheets("Helper Sheet")
    Set wsStaging = ThisWorkbook.Worksheets("AI Staging")
    On Error GoTo 0
    
    If wsHelper Is Nothing Then Exit Sub
    
    ' 1. Create/Reset Staging Sheet safely
    If wsStaging Is Nothing Then
        Set wsStaging = ThisWorkbook.Worksheets.Add(After:=wsHelper)
        wsStaging.Name = "AI Staging"
    End If
    
    Application.ScreenUpdating = False
    
    ' 2. Clear data but keep the button (Surgical Clear)
    Dim lastS As Long
    lastS = wsStaging.Cells(wsStaging.Rows.count, "A").End(xlUp).Row
    If lastS >= 2 Then wsStaging.Range("A2:C" & lastS).ClearContents: wsStaging.Range("A2:C" & lastS).Interior.ColorIndex = xlNone
    
    wsStaging.Cells(1, 1).Value = "Messy Input": wsStaging.Cells(1, 2).Value = "Verified Output": wsStaging.Cells(1, 3).Value = "AI Safety Analysis"
    wsStaging.Range("A1:C1").Font.Bold = True
    
    ' 3. Pull Data from Helper Sheet
    lastRowHelper = wsHelper.Cells(wsHelper.Rows.count, "A").End(xlUp).Row
    If lastRowHelper >= 2 Then
        wsHelper.Range("A2:B" & lastRowHelper).Copy
        wsStaging.Range("A2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    
    ' 4. Run Analysis
    Call RunStagingAnalysis(wsStaging)
    
    wsStaging.Activate
    wsStaging.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub RunStagingAnalysis(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim inText As String, outText As String
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        inText = Trim(CStr(ws.Cells(i, 1).Value))
        outText = Trim(CStr(ws.Cells(i, 2).Value))
        
        ' Flag: Single Word is too vague for AI training
        If InStr(1, inText, " ") = 0 Then
            ws.Cells(i, 3).Value = "[WARNING] Single word (Vague)"
            ws.Cells(i, 3).Interior.Color = 65535 ' Yellow
        Else
            ws.Cells(i, 3).Value = "[SAFE] Ready to Teach"
            ws.Cells(i, 3).Interior.Color = 13565855 ' Green
        End If
    Next i
End Sub

Public Sub ExportCleanDataForAI()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("AI Staging")
    Dim lastRow As Long, i As Long, fNum As Integer, csvPath As String
    Dim uniqueEntries As Object: Set uniqueEntries = CreateObject("Scripting.Dictionary")
    Dim key As String, count As Long
    
    csvPath = Environ("APPDATA") & "\AI_Bulk_Training_Data.csv"
    fNum = FreeFile
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' 1. Use a Dictionary to filter out duplicates BEFORE exporting
    For i = 2 To lastRow
        ' ONLY export rows that contain the word "[SAFE]"
        If InStr(1, UCase(ws.Cells(i, 3).Value), "[SAFE]") > 0 Then
            ' Create a unique key combining the input and output
            key = Trim(ws.Cells(i, 1).Value) & "|" & Trim(ws.Cells(i, 2).Value)
            uniqueEntries(key) = True
        End If
    Next i
    
    ' 2. Write ONLY the unique pairs to the CSV
    Open csvPath For Output As #fNum
    Print #fNum, "messy_phrase,official_name"
    
    Dim entry As Variant, parts() As String
    count = 0
    For Each entry In uniqueEntries.keys
        parts = Split(entry, "|")
        Print #fNum, """" & parts(0) & """,""" & parts(1) & """"
        count = count + 1
    Next entry
    
    Close #fNum
    MsgBox "AI Training set compressed! Exported " & count & " unique clean pairs.", vbInformation
End Sub

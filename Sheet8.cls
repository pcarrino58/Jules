VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ===== Module: RulesHelpers =====
Option Explicit

Public Function EnsureRulesSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Rules Sheet")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Rules Sheet"
        ws.Range("A1:D1").Value = Array("Type", "From", "To", "Weight")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:D").ColumnWidth = 28
    End If
    Set EnsureRulesSheet = ws
End Function

' Upsert by (Type, From). If row exists, updates; else appends.
Public Sub UpsertRule(ByVal ws As Worksheet, _
                      ByVal ruleType As String, _
                      ByVal fromText As String, _
                      ByVal toText As String, _
                      Optional ByVal weight As Variant)
    Dim lastRow As Long, r As Long
    Dim t As String, f As String
    Dim updated As Boolean: updated = False

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For r = 2 To lastRow
        t = LCase$(Trim$(CStr(ws.Cells(r, "A").Value)))
        f = LCase$(Trim$(CStr(ws.Cells(r, "B").Value)))
        If t = LCase$(ruleType) And f = LCase$(fromText) Then
            ' Update existing
            ws.Cells(r, "C").Value = toText
            If Not IsMissing(weight) Then
                ws.Cells(r, "D").Value = weight
            Else
                ws.Cells(r, "D").ClearContents
            End If
            updated = True
            Exit For
        End If
    Next r

    If Not updated Then
        r = lastRow + 1
        ws.Cells(r, "A").Value = ruleType
        ws.Cells(r, "B").Value = fromText
        ws.Cells(r, "C").Value = toText
        If Not IsMissing(weight) Then ws.Cells(r, "D").Value = weight
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' This ensures that if any rule definition changes, it forces a text backup sync
    On Error GoTo ErrorHandler
    
    ' Temporarily disable events to prevent cascade loops
    Application.EnableEvents = False
    
    ' Run the public sync tool from your Matcher module
    Call Matcher.SyncRulesWithExternal(ThisWorkbook, "Rules Sheet")

CleanExit:
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    MsgBox "Failed to auto-save Rules to AppData: " & Err.Description, vbCritical
    Resume CleanExit
End Sub




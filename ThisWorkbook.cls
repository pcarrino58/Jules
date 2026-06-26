VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'========================================================
' Rebuild the Pivot AFTER a successful save.
'========================================================
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Not Success Then Exit Sub
    On Error GoTo Done
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Call the builder in the standard module
    Call PivotBuilder_Rebuild

Done:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'========================================================
' Rebuild the Pivot ANYTIME data changes in the helper sheet.
'========================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Only react if the change is on the source sheet
    If Sh.Name = "Helper Sheet" Then
        On Error GoTo DoneHandler
        
        Dim oldActive As Object
        Set oldActive = ActiveSheet
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        
        ' Smart refresh
        Call PivotBuilder_Refresh
        
        ' Restore focus to the sheet user was editing (Helper Sheet)
        If Not oldActive Is Nothing Then
            On Error Resume Next
            oldActive.Activate
            On Error GoTo DoneHandler
        End If

DoneHandler:
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    ElseIf Sh.Name = "Rules Sheet" Then
        ' Sync dynamically
        Matcher.SyncRulesWithExternal Me, "Rules Sheet"
    ElseIf Sh.Name = "LearnedMappings" Then
        ' Sync dynamically
        Matcher.SyncLearningsWithExternal Me, "LearnedMappings"
    End If
End Sub


'========================================================
' Cleanup on Open
'========================================================
Private Sub Workbook_Open()
    ' Automatically clean up temporary dropdown ranges on open
    Matcher.CleanupOrphanedPickLists

    ' Perform Static Startup Sync
    Matcher.SyncRulesWithExternal Me, "Rules Sheet"
    Matcher.SyncLearningsWithExternal Me, "LearnedMappings"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Perform Static Shutdown Sync
    Matcher.SyncRulesWithExternal Me, "Rules Sheet"
    Matcher.SyncLearningsWithExternal Me, "LearnedMappings"
End Sub

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
    End If
End Sub


'========================================================
' Cleanup on Open
'========================================================
Private Sub Workbook_Open()
    Call ImportAppDataToTemplate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Do you want to sync your latest rules and learned mappings to your AI Master Vault before closing?", _
                      vbYesNo + vbQuestion, "AI Auto-Sync")
    
    If response = vbYes Then
        ' Turn on the hourglass
        Application.Cursor = xlWait
        
        ' Step 1
        Application.StatusBar = "Syncing AI Vault (1/3): Saving Lookups..."
        DoEvents
        Call SyncLookupsWithExternal(ThisWorkbook, "Uniformat RS Means Lookup")
        
        ' Step 2
        Application.StatusBar = "Syncing AI Vault (2/3): Saving Rules..."
        DoEvents
        Call SyncRulesWithExternal(ThisWorkbook, "Rules Sheet")
        
        ' Step 3
        Application.StatusBar = "Syncing AI Vault (3/3): Saving Learned Mappings..."
        DoEvents
        Call OverwriteLearnedFromSheet
        
        ' Reset everything
        Application.StatusBar = False
        Application.Cursor = xlDefault
        
        MsgBox "AI Master Vault successfully updated!", vbInformation, "Sync Complete"
    End If
End Sub



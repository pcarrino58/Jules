VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    ' Prevents the screen from flickering while navigating to the tab
    Application.ScreenUpdating = False
    
    ' Triggers the refresh engine located in Pivot.bas
    Call PivotBuilder_Refresh
    
    Application.ScreenUpdating = True
End Sub

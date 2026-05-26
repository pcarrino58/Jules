Attribute VB_Name = "MatcherUtils"
' This is a dummy implementation to support compilation of the test suite.

Public Sub InitializeMatcher(ws As Worksheet)
    ' Dummy implementation
    Debug.Print "Matcher initialized with lookup sheet: " & ws.Name
End Sub

Public Sub GetBestMatchForInput(ByVal inputPhrase As String, ByRef outResult As String, ByRef outConf As Double)
    ' Dummy implementation
    outResult = "Found: " & inputPhrase
    outConf = 0.95 ' Return a Double value
End Sub

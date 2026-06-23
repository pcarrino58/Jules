VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    Dim watch As Range
    Dim hit As Range
    Dim rowsDict As Object
    Dim changedRows As Object
    Dim area As Range
    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rngA As Range
    Dim c As Range
    Dim eVal As Variant
    Dim arr As Variant
    Dim cnt As Long
    Dim lastVal As Variant
    Dim j As Long
    Dim v As Variant
    Dim addrEList As String
    Dim key As Variant
    Dim re As Object
    Dim txt As String
    Dim prevCalc As XlCalculation

    Set watch = Union(Me.Columns("A"), Me.Columns("E"), Me.Range("AZ:BI"))
    Set hit = Intersect(Target, watch)
    If hit Is Nothing Then Exit Sub

    prevCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If prevCalc <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic

    Set changedRows = CreateObject("Scripting.Dictionary")
    Set rngA = Intersect(Target, Me.Columns("A"))

    If Not rngA Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        ' whole-word match, singular or plural, for common synonyms
        re.Pattern = "\b(restroom|bathroom|toilet|lavatory)s?\b"
        For Each c In rngA.Cells
            txt = CStr(c.Value)
            If Len(txt) > 0 Then
                txt = Replace(txt, Chr(160), " ")
                ' replace matched words only, keep surrounding text
                If re.Test(txt) Then
                    c.Value = re.Replace(txt, "Washrooms")
                    If Not changedRows.Exists(c.Row) Then changedRows.Add c.Row, True

                    With Me.Cells(c.Row, "E")
                        If .HasArray Then
                            .FormulaArray = .FormulaArray
                        ElseIf Len(.formula) > 0 Then
                            .formula = .formula
                        Else
                            .Calculate
                        End If
                    End With
                End If
            End If
        Next c
    End If

    Set rowsDict = CreateObject("Scripting.Dictionary")
    For Each area In hit.Areas
        firstRow = area.Row
        lastRow = area.Row + area.Rows.count - 1
        If lastRow >= 2 Then
            If firstRow < 2 Then firstRow = 2
            For r = firstRow To lastRow
                If Not rowsDict.Exists(r) Then rowsDict.Add r, vbNullString
            Next r
        End If
    Next area

    addrEList = ""
    For Each key In rowsDict.keys
        r = CLng(key)

        With Me.Cells(r, "E")
            If .WrapText = False Then .WrapText = True
        End With

        eVal = Me.Cells(r, "E").Value2
        If LenB(Trim$(CStr(eVal))) = 0 Then
            Me.Cells(r, "C").ClearContents
        Else
            arr = Me.Range("AZ" & r & ":BI" & r).Value2
            cnt = 0
            lastVal = Empty

            For j = LBound(arr, 2) To UBound(arr, 2)
                v = arr(1, j)
                If Not IsError(v) Then
                    If LenB(Trim$(CStr(v))) > 0 Then
                        cnt = cnt + 1
                        lastVal = v
                        If cnt > 1 Then Exit For
                    End If
                End If
            Next j

            If cnt = 1 Then
                If Me.Cells(r, "C").Value2 <> lastVal Then
                    Me.Cells(r, "C").Value2 = lastVal
                End If
            Else
                If Not IsEmpty(Me.Cells(r, "C").Value2) Then
                    Me.Cells(r, "C").ClearContents
                End If
            End If
        End If

        If Len(addrEList) = 0 Then
            addrEList = Me.Cells(r, "E").Address
        Else
            addrEList = addrEList & "," & Me.Cells(r, "E").Address
        End If
    Next key

    If changedRows.count > 0 Then
        For Each key In changedRows.keys
            r = CLng(key)
            If InStr(1, addrEList, Me.Cells(r, "E").Address, vbTextCompare) = 0 Then
                If Len(addrEList) = 0 Then
                    addrEList = Me.Cells(r, "E").Address
                Else
                    addrEList = addrEList & "," & Me.Cells(r, "E").Address
                End If
            End If
        Next key
    End If

    If Len(addrEList) > 0 Then
        Me.Range(addrEList).Rows.AutoFit
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
End Sub


Attribute VB_Name = "Module5"
Option Explicit

Sub FilePicker_Button2_Click()

    Call FilePicker(ws5, "E5", "J5")

End Sub

Sub Search_Button2_Click()

    Call Search(ws5, "E5", "J5")

End Sub

Sub Clear_Button4_Click()

    Call ClearRange(ws5, 10, "D", "G", "D")
    
    With ws5
        .Range("M10:M21").Value = 0
    End With
    
End Sub

Sub ForecastChecker(targetSheet As Worksheet, ForecastTracker As Workbook, lastRow As Long, sizeByRow As Long)

    Dim currentWR As String
    Dim currentYear As String
    Dim rgFound As Range
    Dim row As Long
    Dim innerCol As Integer
    Dim lowerBound As Integer
    Dim janTotal    As Integer: janTotal = 0
    Dim febTotal    As Integer: febTotal = 0
    Dim marTotal    As Integer: marTotal = 0
    Dim aprTotal    As Integer: aprTotal = 0
    Dim mayTotal    As Integer: mayTotal = 0
    Dim junTotal    As Integer: junTotal = 0
    Dim julTotal    As Integer: julTotal = 0
    Dim augTotal    As Integer: augTotal = 0
    Dim sepTotal    As Integer: sepTotal = 0
    Dim octTotal    As Integer: octTotal = 0
    Dim novTotal    As Integer: novTotal = 0
    Dim decTotal    As Integer: decTotal = 0
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim match As Object
    
    With targetSheet
        .Range("M10:M21").Value = 0
    End With
    
    ' getting year
    With regEx
        .Global = False
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "([0-9]{4})"
    End With
    
    If regEx.test(ForecastTracker.Name) Then
        Set match = regEx.Execute(ForecastTracker.Name)
        currentYear = match.Value
    End If
    
    ' Find lower bound
    With ForecastTracker.Worksheets(1).UsedRange
        Set rgFound = .Cells.Find(What:=currentYear & " Project Forecast", After:=.Range("A1"), LookIn:=xlValues, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
        On Error GoTo NoYear
        lowerBound = rgFound.row
        On Error GoTo 0
        Set rgFound = Nothing
    End With
    
    ' Programmatically checking the master tracker for the user-inputted WRs
    With ForecastTracker.Worksheets(1)
    
        For row = 10 To lastRow
            
            currentWR = Trim(CStr(targetSheet.Range("D" & row).Value))
            Application.StatusBar = "Checking " & currentWR & ", " & CStr(Round(row / sizeByRow, 2) * 100) & "% Complete"
            
            While Left(currentWR, 5) <> "HBCBS" And (row - 10) < sizeByRow - 1
                row = row + 1
            Wend
            
            'For Each wSheet In Workbooks(nameOfDoc).Worksheets
            With .UsedRange
                Set rgFound = .Cells.Find(What:=currentWR, After:=.Range("A1"), LookIn:=xlValues, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False) ' .Activate
            End With
                
            If rgFound Is Nothing And Len(Trim(CStr(targetSheet.Range("E" & row).Value))) = 0 Then
                targetSheet.Range("E" & row).Value = "Not on tracker"
                targetSheet.Range("E" & row).Interior.Color = RGB(255, 150, 150)
                
            ElseIf Not (rgFound Is Nothing) Then
                If rgFound.row < lowerBound Then
                    targetSheet.Range("E" & row).Value = "On Tracker"
                    
                    If Len(Trim(CStr(.Range("C" & rgFound.row).Value))) > 0 Then
                        targetSheet.Range("F" & row).Value = "Jan = " & .Range("C" & rgFound.row).Value
                        janTotal = janTotal + .Range("C" & rgFound.row).Value
                    End If
                    
                    For innerCol = 4 To 14
                        If Len(Trim(CStr(.Cells(rgFound.row, innerCol).Value))) > 0 Then
                            targetSheet.Range("F" & row).Value = targetSheet.Range("F" & row).Value _
                                    & vbCrLf _
                                    & .Cells(2, innerCol).Value _
                                    & " = " _
                                    & .Cells(rgFound.row, innerCol).Value
                        End If
                        If Len(.Cells(rgFound.row, innerCol).Value) > 0 Then
                            Select Case innerCol
                                Case 4
                                    febTotal = febTotal + .Range("D" & rgFound.row).Value
                                Case 5
                                    marTotal = marTotal + .Range("E" & rgFound.row).Value
                                Case 6
                                    aprTotal = aprTotal + .Range("F" & rgFound.row).Value
                                Case 7
                                    mayTotal = mayTotal + .Range("G" & rgFound.row).Value
                                Case 8
                                    junTotal = junTotal + .Range("H" & rgFound.row).Value
                                Case 9
                                    julTotal = julTotal + .Range("I" & rgFound.row).Value
                                Case 10
                                    augTotal = augTotal + .Range("J" & rgFound.row).Value
                                Case 11
                                    sepTotal = sepTotal + .Range("K" & rgFound.row).Value
                                Case 12
                                    octTotal = octTotal + .Range("L" & rgFound.row).Value
                                Case 13
                                    novTotal = novTotal + .Range("M" & rgFound.row).Value
                                Case 14
                                    decTotal = decTotal + .Range("N" & rgFound.row).Value
                            End Select
                        End If
                    Next
                End If
            End If
                
            targetSheet.Range("F" & row).HorizontalAlignment = xlLeft
        Next row
    End With
    
    With targetSheet
        .Range("M10").Value = janTotal
        .Range("M11").Value = febTotal
        .Range("M12").Value = marTotal
        .Range("M13").Value = aprTotal
        .Range("M14").Value = mayTotal
        .Range("M15").Value = junTotal
        .Range("M16").Value = julTotal
        .Range("M17").Value = augTotal
        .Range("M18").Value = sepTotal
        .Range("M19").Value = octTotal
        .Range("M20").Value = novTotal
        .Range("M21").Value = decTotal
    End With
    
    ' Closing sequence for this routine
    Application.StatusBar = ""
    ThisWorkbook.Activate
    
    Exit Sub
    
NoYear:
    MsgBox "Can't find cell with value: [Year] Project Forecast on first worksheet of selected document."

End Sub







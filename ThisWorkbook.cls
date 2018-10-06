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

Sub Workbook_Open()
    
    
    Dim monthArray As Variant
    Dim monthArray2 As Variant
    Dim spot As Integer
    
    monthArray = Array("None", "1 - January", "2 - February", "3 - March", "4 - April", _
                    "5 - May", "6 - June", "7 - July", "8 - August", "9 - September", _
                    "10 - October", "11 - November", "12 - December")
                
    monthArray2 = Array("None", "January", "February", "March", "April", "May", "June", "July", "August", _
                    "September", "October", "November", "December")
    
    ' Setting global worksheet variables - NOT working at the moment
    Call Declarations
    
    ' cleaning up sheets
    CleanUpToolSheet ws1, 10, "A", "G", pwSpot:="C6"
    CleanUpToolSheet ws2, 10, "A", "C", pwSpot:="C6"
    CleanUpToolSheet ws3, 10, "D", "I", attachSpot:="J5"
    CleanUpToolSheet ws4, 10, "A", "H", pwSpot:="C6"
    CleanUpToolSheet ws5, 10, "D", "F", attachSpot:="J5"
    CleanUpToolSheet ws6, 10, "A", "N", pwSpot:="C6"
    CleanUpToolSheet ws7, 10, "D", "F", attachSpot:="J5"
    CleanUpToolSheet ws8, 10, "A", "F", pwSpot:="C6"
    CleanUpToolSheet ws9, 10, "B", "O"
    CleanUpToolSheet ws11, 12, "A", "H"
    ' CleanUpToolSheet ws10, 10, "A", "F" - write tool, not ready yet
    
    For spot = LBound(monthArray) To UBound(monthArray)
        ws1.ComboBox1.AddItem monthArray(spot)
        ws4.ComboBoxAR.AddItem monthArray(spot)
        ws8.ComboBox1BA.AddItem monthArray(spot)
    Next spot
    
    For spot = LBound(monthArray2) To UBound(monthArray2)
        ws7.ComboBoxMWS7.AddItem monthArray2(spot)
    Next
    
    With ws1
        .ComboBox2.AddItem "Off"
        .ComboBox2.AddItem "On"
        .ComboBox3.AddItem "Off"
        .ComboBox3.AddItem "On"
    End With
    
    With ws7
        .ComboBoxAU.AddItem "Off"
        .ComboBoxAU.AddItem "On"
        .ComboBoxAI.AddItem "Off"
        .ComboBoxAI.AddItem "On"
    End With
    
    With ws8
        .ComboBox2BA.AddItem "Off"
        .ComboBox2BA.AddItem "On"
        .ComboBox3BA.AddItem "Intake"
        .ComboBox3BA.AddItem "Master Tracker"
        .ComboBox3BA.AddItem "Funding CC"
        .ComboBox3BA.AddItem "UAT Hours"
        .ComboBox3BA.AddItem "Planning Tab"
    End With
    
    With ws9
        .ComboBox2FC.AddItem "Off"
        .ComboBox2FC.AddItem "All"
        .ComboBox2FC.AddItem "Bharath's > Yours"
        .ComboBoxAILI.AddItem "Off"
        .ComboBoxAILI.AddItem "On"
        .ComboBoxMonth.AddItem "Please select"
        .ComboBoxMonth.AddItem "January"
        .ComboBoxMonth.AddItem "February"
        .ComboBoxMonth.AddItem "March"
        .ComboBoxMonth.AddItem "April"
        .ComboBoxMonth.AddItem "May"
        .ComboBoxMonth.AddItem "June"
        .ComboBoxMonth.AddItem "July"
        .ComboBoxMonth.AddItem "August"
        .ComboBoxMonth.AddItem "September"
        .ComboBoxMonth.AddItem "October"
        .ComboBoxMonth.AddItem "November"
        .ComboBoxMonth.AddItem "December"
    End With
    
    With ws10.ComboBoxMonth
        .AddItem "Filter for a month"
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With
        
    ' setting list indexes
    ws1.ComboBox1.ListIndex = 0
    ws1.ComboBox2.ListIndex = 0
    ws1.ComboBox3.ListIndex = 0
    ws4.ComboBoxAR.ListIndex = 0
    ws7.ComboBoxAU.ListIndex = 0
    ws7.ComboBoxAI.ListIndex = 0
    ws7.ComboBoxMWS7.ListIndex = 0
    ws8.ComboBox1BA.ListIndex = 0
    ws8.ComboBox2BA.ListIndex = 0
    ws8.ComboBox3BA.ListIndex = 0
    ws9.ComboBox2FC.ListIndex = 0
    ws9.ComboBoxAILI.ListIndex = 0
    ws9.ComboBoxMonth.ListIndex = 0
    ws10.ComboBoxMonth.ListIndex = 0
 
End Sub


Sub CleanUpToolSheet(ws As Worksheet, startRow As Integer, startCol As String, endCol As String, Optional attachSpot As Variant, Optional pwSpot As Variant)

    Dim lastRow As Long: lastRow = 0
    Dim col As Variant
    
    With ws
        For Each col In .Range(.Cells(startRow, startCol), .Cells(startRow, endCol)).Columns
            If lastRow < .Cells(.Rows.Count, col.Column).End(xlUp).row Then
                lastRow = .Cells(.Rows.Count, col.Column).End(xlUp).row
            End If
        Next
        
        If lastRow >= startRow Then
            .Range(.Cells(startRow, startCol), .Cells(lastRow, endCol)).UnMerge
            .Range(.Cells(startRow, startCol), .Cells(lastRow, endCol)).Borders.LineStyle = xlNone
            .Range(.Cells(startRow, startCol), .Cells(lastRow, endCol)).ClearContents
            .Range(.Cells(startRow, startCol), .Cells(lastRow, endCol)).Interior.ColorIndex = 0
            .Range(.Cells(startRow - 1, startCol), .Cells(startRow - 1, endCol)).Borders.LineStyle = xlContinuous
        End If
        
        If IsMissing(attachSpot) = False Then
            .Range(attachSpot).Value = "No File Attached"
            .Range(attachSpot).Font.Color = vbWhite
            .Range(attachSpot).Interior.Color = RGB(192, 0, 0)
        End If
        
        If IsMissing(pwSpot) = False Then
            .Range(pwSpot).Value = ""
        End If
    End With

End Sub


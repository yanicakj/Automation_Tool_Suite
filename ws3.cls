VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ws3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

    Dim lastRow As Long
    Dim cellInfo As String
    
    With ws3
    
        lastRow = .Cells(.Rows.Count, "D").End(xlUp).row
        cellInfo = Trim(CStr(.Range("E5").Value))
        
        If lastRow = 9 Then
            .Range("A1").Activate
            .Range("E5").Activate
        End If
        
        If BookOpen(returnFileName(cellInfo, "\")) Then
            .Range("J5").Value = "Master Tracker Attached"
            .Range("J5").Interior.Color = vbGreen
            .Range("J5").Font.Color = vbBlack
        Else
            .Range("J5").Value = "No File Attached"
            .Range("J5").Font.Color = vbWhite
            .Range("J5").Interior.Color = RGB(192, 0, 0)
        End If
        
    End With

End Sub

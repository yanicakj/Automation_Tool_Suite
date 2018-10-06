VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ws6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

    Dim lastRow As Long
    
    lastRow = ws6.Cells(ws6.Rows.Count, "A").End(xlUp).row
    
    If lastRow = 9 Then
        ws6.Range("A1").Activate
        ws6.Range("C7").Activate
    End If

End Sub


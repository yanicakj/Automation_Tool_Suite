Attribute VB_Name = "Module6"
Option Explicit

Sub Run_Button4_Click()

    Call CQscrub(ws6, "A")

End Sub

Sub LooperHour(counter As Integer, sizeByRow As Integer, ie As InternetExplorer, searchStringObj As Object, ObjElement As Object, jumpBack As Integer, targetSheet As Worksheet, wrCol As String)
        
    Dim identification  As String
    Dim closeObj        As Object
    Dim impactObj       As Object
    Dim tabObj          As Object
    Dim indLabel        As IHTMLElement
    Dim LabelCollection As IHTMLElementCollection
    Dim releaseCheck   As String
    Dim firstEntry As Boolean: firstEntry = True
    Dim hoursGrid As Object
    Dim tableCollection As Object
    Dim indTable As Object
    Dim tdCollection As Object
    Dim indTd As Object
    Dim colSpot As Integer
    
    With targetSheet

        On Error Resume Next

        ' Updating the status bar programmatically
        Application.StatusBar = "Checking " & CStr(.Range("A" & CStr(counter + 10)).Value) & ", " & CStr(Round(counter / sizeByRow, 2) * 100) & "% Complete"
        
        searchStringObj.Value = CStr(.Range("A" & CStr(counter + 10)).Value)
        ObjElement.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:06"))
    
        ' Clicking on WR tab to focus screen
        identification = "..." & CStr((counter - jumpBack))
        Do While tabObj Is Nothing
            Set tabObj = ie.document.getElementById(identification)
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        Application.Wait (Now + TimeValue("0:00:01"))
        tabObj.Click
    
        ' Clicking Estimates Tab
        identification = "..." & CStr((counter - jumpBack) + 1) & "..." & CStr(9 + ((counter - jumpBack) * 18))
        Set impactObj = ie.document.getElementById(identification)
        impactObj.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:03"))
    
        ' Finding Estimate Line Items
        identification = "..." & (2 + ((counter - jumpBack) * 3)) & "..."
        Do While hoursGrid Is Nothing
            Set hoursGrid = ie.document.getElementById(identification)
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        Application.Wait (Now + TimeValue("0:00:02"))

        Set tableCollection = hoursGrid.getElementsByTagName("...")
        For Each indTable In tableCollection
            colSpot = 2
            Set tdCollection = indTable.getElementsByTagName("...")
            For Each indTd In tdCollection
                If firstEntry = True Then
                    .Cells((10 + counter), colSpot).Value = indTd.innerText
                Else
                    .Cells((10 + counter), colSpot).Value = .Cells((10 + counter), colSpot).Value & vbCrLf & indTd.innerText
                End If
                colSpot = colSpot + 1
            Next
            firstEntry = False
        Next
    
        ' Closing tab
        Set closeObj = ie.document.getElementsByClassName(". . .")(0)
        closeObj.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop

    End With
    
    On Error GoTo 0

End Sub



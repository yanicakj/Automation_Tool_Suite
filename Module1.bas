Attribute VB_Name = "Module1"
Option Explicit

Sub Run_Button1_Click()

    Call CQscrub(ws1, "A")

End Sub


Sub CQscrub(targetSheet As Worksheet, wrCol As String, Optional Caller As String)

    Dim ie As InternetExplorer ' Variable for Internet Explorer instance
    Dim counter As Integer: counter = 0
    Dim i As Long: i = 0
    Dim row As Integer
    Dim pwObj As Object
    Dim password As String
    Dim searchStringObj As Object
    Dim logoutObj As Object
    Dim loginObj As Object
    Dim ObjElement As Object
    Dim objCollection As Object
    Dim errorDivObj As Object
    Dim closeObj As Object
    Dim jumpBack As Integer: jumpBack = 0
    Dim lastRow As Integer
    Dim sizeByRow As Integer
    Dim matchFoundIndex As Long
    Dim Dupe As Boolean: Dupe = False
    Dim rowSpot As Integer
    
    ' Setup
    With targetSheet
    
        ' Checks if password cell is empty, prompts user and ends program if empty
        If Len(CStr(Trim(.Range("C6").Value))) = 0 Then
            .Range("C6").Activate
            MsgBox "Please enter your password!"
            Exit Sub
        Else
            password = Trim(CStr(.Range("C6").Value))
        End If
    
        ' clean & reset
        CleanupWRs targetSheet, 10, wrCol
        
        If targetSheet Is ws6 Then
            ClearRange targetSheet, 10, "B", "O", wrCol
        ElseIf targetSheet Is ws1 Then
            ClearFrontEndScrub targetSheet, wrCol, "G"
        ElseIf targetSheet Is ws2 Then
            ClearFrontEndScrub targetSheet, wrCol, "C"
        ElseIf targetSheet Is ws4 Then
            ClearFrontEndScrub targetSheet, wrCol, "H"
        End If
        
        ' Finding last row with empties and calculating size of list of WRs
        lastRow = .Cells(.Rows.Count, wrCol).End(xlUp).row
        sizeByRow = lastRow - 9
    
    End With

    ' Creating ie instance
    On Error GoTo ErrHandler
    Set ie = CreateObject("InternetExplorer.Application")
    ' uncomment below for testing
    'ie.Height = 1000
    'ie.Width = 1000
    ie.Visible = False
    
    ' Navigation
    On Error Resume Next
    ie.navigate "http://....."
    Application.StatusBar = "Loading http://...."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:07"))
    Application.StatusBar = "Please wait..."

    ' first checking if user is logged in
    Set objCollection = ie.document.getElementsByTagName("....")
    Do While i < objCollection.Length
        If objCollection(i).Name = "passwordId" Then
            objCollection(i).Value = password
            Set loginObj = ie.document.getElementById("...")
            loginObj.Click
            Application.Wait (Now + TimeValue("0:00:05"))
            Exit Do
        End If
        i = i + 1
    Loop: i = 0

    ' searching for "search" bar and saving as object
    Do While i < objCollection.Length
        If objCollection(i).Name = "..." Then
            Set searchStringObj = objCollection(i)
            Exit Do
        End If
        i = i + 1
    Loop
    Set ObjElement = ie.document.getElementById("...")

    ' main loop
    With targetSheet
        Do While counter < sizeByRow + 1
            If counter <> 0 Then
                For rowSpot = 10 + counter To 10 Step -1
                    If .Range(wrCol & CStr(counter + 10)).Value = .Range(wrCol & CStr(rowSpot - 1)).Value Then
                        .Range("B" & CStr(counter + 10)).Value = "dupe"
                        .Range("C" & CStr(counter + 10)).Value = "dupe"
                        .Range("B" & CStr(counter + 10)).Interior.Color = RGB(150, 255, 255)
                        .Range("C" & CStr(counter + 10)).Interior.Color = RGB(150, 255, 255)
                        If targetSheet.CodeName = "ws1" Or targetSheet.CodeName = "ws4" Then
                            .Range("D" & CStr(counter + 10)).Value = "dupe"
                            .Range("E" & CStr(counter + 10)).Value = "dupe"
                            .Range("F" & CStr(counter + 10)).Value = "dupe"
                            .Range("D" & CStr(counter + 10)).Interior.Color = RGB(150, 255, 255)
                            .Range("E" & CStr(counter + 10)).Interior.Color = RGB(150, 255, 255)
                            .Range("F" & CStr(counter + 10)).Interior.Color = RGB(150, 255, 255)
                        End If
                        
                        Dupe = True
                        Exit For
                    End If
                Next
            End If
    
            If Left(Trim(CStr(.Range(wrCol & CStr(counter + 10)))), 5) <> "HBCBS" Or Dupe Then
                counter = counter + 1
                jumpBack = jumpBack + 1
            Else
                If targetSheet.CodeName = "ws1" Then
                    Looper counter, sizeByRow, ie, searchStringObj, ObjElement, jumpBack, targetSheet, wrCol
                ElseIf targetSheet.CodeName = "ws2" Then
                    Looper2 counter, sizeByRow, ie, searchStringObj, ObjElement, jumpBack, targetSheet, wrCol
                ElseIf targetSheet.CodeName = "ws4" Then
                    Looper counter, sizeByRow, ie, searchStringObj, ObjElement, jumpBack, targetSheet, wrCol
                ElseIf targetSheet.CodeName = "ws6" Then
                    LooperHour counter, sizeByRow, ie, searchStringObj, ObjElement, jumpBack, targetSheet, wrCol
                End If
                
                counter = counter + 1
            End If
    
            Dupe = False
        Loop
    End With


    ' Closing final open tab
    Application.StatusBar = "Finishing Process..."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Set closeObj = ie.document.getElementsByClassName("...")(0)
    closeObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Set logoutObj = ie.document.getElementById("...")
    logoutObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:03"))

    ' Closing down program
    Set pwObj = ie.document.getElementById("...")
    pwObj.Value = password
    Set loginObj = ie.document.getElementById("...")
    loginObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:06"))
    ie.Quit
    
    ' cleanup
    Set ie = Nothing
    Set ObjElement = Nothing
    Set objCollection = Nothing
    Application.StatusBar = ""
    If Caller <> "Rohini" Then MsgBox "Done!"
    
    Exit Sub

ErrHandler:

    MsgBox "Unable to open Internet Exlporer! Exiting!"

End Sub

Sub Looper(counter As Integer, sizeByRow As Integer, ie As InternetExplorer, searchStringObj As Object, ObjElement As Object, jumpBack As Integer, targetSheet As Worksheet, wrCol As String)

    Dim i As Integer
    Dim inv As String
    Dim identification As String
    Dim closeObj As Object
    Dim yourObj As Object
    Dim stateObj As Object
    Dim invObj As Object
    Dim impactObj As Object
    Dim tabObj As Object
    Dim ele As IHTMLElement
    Dim id As String
    Dim LabelCollection As IHTMLElementCollection
    Dim releaseCheckStr As String
    Dim releaseObj As Object
    Dim headlineObj As Object
    Dim systemsObj As Object
    Dim leadCol As String
    Dim dateCol As String
    Dim statCol As String
    Dim uinvCol As String
    Dim headCol As String
    Dim systCol As String
    Dim projNameBool As Boolean: projNameBool = False
    
    If wrCol = "A" Then
        leadCol = "B": dateCol = "E": statCol = "C": uinvCol = "D": headCol = "F": systCol = "G"
    ElseIf wrCol = "C" Then
        leadCol = "H": dateCol = "D": statCol = "F": uinvCol = "G": headCol = "B": systCol = "I"
        projNameBool = True
    End If

    On Error Resume Next

    With targetSheet
        ' Updating the status bar programmatically
        Application.StatusBar = "Checking " & CStr(.Range(wrCol & CStr(counter + 10)).Value) & ", " & CStr(Round(counter / sizeByRow, 2) * 100) & "% Complete"
        searchStringObj.Value = CStr(.Range(wrCol & CStr(counter + 10)).Value)
        ObjElement.Click
        
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:06"))

        ' Finding the IT Lead
        identification = "..." & CStr(9 + ((counter - jumpBack) * 49))
        Set yourObj = ie.document.getElementById(identification)
        .Range(leadCol & (10 + counter)).Value = yourObj.Value
            
        ' Finding WR State
        identification = "..." & CStr(2 + ((counter - jumpBack) * 3))
        Set stateObj = ie.document.getElementById(identification)
        .Range(statCol & (10 + counter)).Value = stateObj.Value
        If stateObj.Value = "Withdrawn" Or stateObj.Value = "Closed" Or stateObj.Value = "Deferred" Or stateObj.Value = "Rejected" Then
            .Range(statCol & (10 + counter)).Interior.Color = RGB(255, 150, 150)
        End If
    
        ' Finding Release Date
        For i = 4 To 1 Step -1
            releaseCheckStr = "..." & CStr(4 + ((counter - jumpBack) * 8))
            Set releaseObj = ie.document.getElementById(releaseCheckStr)
            If releaseObj.Value <> "" Then
                Exit For
            End If
        Next
        .Range(dateCol & (10 + counter)).Value = releaseObj.Value
        
        If .CodeName = "ws1" Then
            
            ' checking release
            If .OLEObjects("ComboBox1").Object.Text <> "None" Then
                If Trim(Left(.OLEObjects("ComboBox1").Object.Text, 2)) <> Mid(releaseObj.Value, 1, InStr(releaseObj.Value, "/") - 1) Then
                    .Range(dateCol & (10 + counter)).Interior.Color = RGB(255, 255, 150)
                End If
            End If
    
            ' headline
            If .OLEObjects("ComboBox2").Object.Text = "On" Then
                identification = "..." & CStr(0 + ((counter - jumpBack) * 8))
                Set headlineObj = ie.document.getElementById(identification)
                .Range(headCol & (10 + counter)).Value = headlineObj.Value
            End If
        
            ' systems to be tested
            If .OLEObjects("ComboBox3").Object.Text = "On" Then
                identification = "..." & CStr(1 + ((counter - jumpBack) * 11))
                Set systemsObj = ie.document.getElementById(identification)
                .Range(systCol & (counter + 10)).Value = systemsObj.innerText
            End If
            
        ElseIf projNameBool Then
        
            identification = "..." & CStr(0 + ((counter - jumpBack) * 8))
            Set headlineObj = ie.document.getElementById(identification)
            .Range(headCol & (10 + counter)).Value = headlineObj.Value
            
            ' checking release
            If .OLEObjects("ComboBoxAR").Object.Text <> "None" Then
                If Trim(Left(.OLEObjects("ComboBoxAR").Object.Text, 2)) <> Mid(releaseObj.Value, 1, InStr(releaseObj.Value, "/") - 1) Then
                    .Range(dateCol & (10 + counter)).Interior.Color = RGB(255, 255, 150)
                End If
            End If
        
        End If
    
        ' Clicking on WR tab to focus screen
        identification = "top-..." & CStr((counter - jumpBack))
        Set tabObj = ie.document.getElementById(identification)
        tabObj.Click
    
        ' Clicking Impact Analysis Tab
        identification = "..." & CStr((counter - jumpBack) + 1) & "..." & CStr(5 + ((counter - jumpBack) * 18))
        Set impactObj = ie.document.getElementById(identification)
        impactObj.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:01"))
    
        ' Finding UAT Involvement
        Set LabelCollection = ie.document.getElementsByClassName("...")
        For Each ele In LabelCollection
            If ele.innerText = "UAT CoE" Then
                id = Replace(ele.id, "cap_", "")
            End If
        Next
        
        Set LabelCollection = ie.document.getElementsByClassName("...")
        For Each ele In LabelCollection
            If InStr(ele.innerHTML, id) > 0 Then
                If InStr(ele.innerText, "Yes") > 0 Then
                    inv = "Yes"
                ElseIf InStr(ele.innerText, "No") > 0 Then
                    inv = "No"
                Else
                    inv = "Blank"
                End If
            End If
        Next
        .Range(uinvCol & (10 + counter)).Value = inv
        
        ' Closing tab
        Set closeObj = ie.document.getElementsByClassName("...")(0)
        closeObj.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
    
    End With
    
    If wrCol = "A" Then
        FormatRow targetSheet, counter + 10, "G"
    ElseIf wrCol = "C" Then
        FormatRow targetSheet, counter + 10, "H"
    End If

    On Error GoTo 0

End Sub

Sub ClearFrontEndScrub(targetSheet As Worksheet, wrCol As String, endCol As String)

    Dim lastRow As Integer
    
    With targetSheet
        lastRow = .Cells(.Rows.Count, wrCol).End(xlUp).row
        
        If lastRow = 9 Then: Exit Sub
        
        If wrCol = "A" Then
            .Range("B10:G" & lastRow).ClearContents
            .Range("B10:G" & lastRow).Interior.ColorIndex = 0
        ElseIf wrCol = "C" Then
            .Range("A10:B" & lastRow).ClearContents
            .Range("D10:H" & lastRow).ClearContents
            .Range("A10:B" & lastRow).Interior.ColorIndex = 0
            .Range("D10:H" & lastRow).Interior.ColorIndex = 0
        End If
        
        .Range("A10:G" & lastRow).Borders.LineStyle = xlNone
        .Range("A9:" & endCol & "9").Borders.LineStyle = xlContinous
        
    End With
    

End Sub

Sub Clear_Button1_Click()

    Call ClearFrontEndScrub(ws1, "A", "G")

End Sub

Sub Clear_Button2_Click()

    Call ClearFrontEndScrub(ws2, "A", "C")

End Sub

Sub Clear_Button3_Click()

    Call ClearFrontEndScrub(ws4, "C", "H")

End Sub

Sub FormatRow(targetSheet As Worksheet, row As Integer, col As String)

    With targetSheet
    
        .Range("A" & row & ":" & col & row).VerticalAlignment = xlCenter
        .Range("A" & row & ":H" & row).HorizontalAlignment = xlCenter
        
        If col = "G" Then
            .Range("F" & row & ":G" & row).HorizontalAlignment = xlLeft
        End If
    
    End With

End Sub

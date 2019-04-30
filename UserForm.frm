'Brandon Skar, 4/23/2018, Move Form for Expeditors

'___________________________________________________________________________________________________________________
'                           SECTION 1: OK BUTTON, CANCEL BUTTON, CLEAR BUTTON
'___________________________________________________________________________________________________________________
'This function is called when the 'OK' button is clicked on the user form or if the "Enter" key is pressed in any textbox.
'The function will determine if the user wants to make make a "Door to Yard", "Yard to Door", "Two Trailer Moves At Once" or a "Double Drop."
'There is a moveComplete boolean to determine if a move was completed successfully, letting the user know if an error occured during the move.
Private Sub okButton_Click()
    Dim moveComplete As Boolean 'false by default
    Application.CutCopyMode = False
    Application.ScreenUpdating = False
    
    '__________INVALID MOVE____________
    If Label3 = "INVALID" Or Label8 = "INVALID" Then
        MsgBox "Error. Your move was not completed."
        Exit Sub
    End If
    
    '_____________ONLY DOOR TO YARD MOVE__________________
    If Label3 <> "" And Label4 = "AVAILABLE" And TextBox3 = "" And TextBox5 = "" Then
        doorToYard
        moveComplete = True
    End If
    
    '_____________ONLY YARD TO DOOR MOVE__________________
    If Label11 <> "" And Label8 = "AVAILABLE" And TextBox1 = "" And TextBox2 = "" Then
        yardToDoor
        moveComplete = True
    End If
    
    '___________TWO TRAILER MOVES AT ONCE_________________
    If Label3 <> "" And Label4 <> "" And Label11 <> "" And Label8 <> "" And (Label4 <> "DOUBLE DROP" Or Label8 <> "DOUBLE DROP") Then
        doorToYard
        yardToDoor
        moveComplete = True
    End If
    
    '___________SWAP TRAILERS (DOUBLE DROP)________________
    If Label4 = "DOUBLE DROP" And Label8 = "DOUBLE DROP" And Label3 <> "" And Label11 <> "" Then
        swapTrailers
        moveComplete = True
    End If
    
    'Error message if there was no move completed.
    If Not moveComplete Then
        MsgBox "Error. Your move was not completed."
        Exit Sub
    End If
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = True
    Unload Me 'Close user form
End Sub

'This function is called when the 'Cancel' button is clicked.
'The function closes the user form. No moves are made.
Private Sub cancelButton_Click()
    Unload Me 'Close user form
End Sub

'This function is called when the 'Clear' button is clicked.
'The function sets all the textboxes on the user form to an empty string.
Private Sub clearButton_Click()
    TextBox1 = ""
    Label3 = ""
    TextBox2 = ""
    TextBox3 = ""
    Label11 = ""
    TextBox5 = ""
    Label8 = ""
    TextBox4 = ""
    Label9 = ""
    Label4 = ""
End Sub

'___________________________________________________________________________________________________________________
'                           SECTION 2: TEXTBOX CHANGE EVENT AND AFTER UPDATE EVENT
'___________________________________________________________________________________________________________________
'This function is called when there is any change in TextBox1 (Move From Door to Yard).
'The function searches for an exact match in TextBox1 IGNORING CASE SENSITIVITY from Cell C1 down to the last used cell in column C.
'When a match is found Label3 will be filled with the value of cell one space to the right of the matched value (Trailer Number).
Private Sub TextBox1_Change()
    Dim iLastRow As Integer
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row

    If UCase(TextBox1) = UCase(TextBox5) And UCase(TextBox2) = UCase(TextBox3) And TextBox1 <> "" And TextBox2 <> "" And TextBox3 <> "" And TextBox5 <> "" Then
        Label4 = "DOUBLE DROP"
        Label4.ForeColor = RGB(0, 0, 0)
        Label8 = "DOUBLE DROP"
        Label8.ForeColor = RGB(0, 0, 0)
        Exit Sub
    ElseIf UCase(TextBox1) = UCase(TextBox2) And TextBox1 <> "" Then
        Label4 = "SAME LOCATION"
        Label4.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    If UCase(TextBox1) = UCase(TextBox5) And TextBox1 <> "" Then
        Label8 = "AVAILABLE"
        Label8.ForeColor = RGB(30, 255, 120)
    End If
    For i = 1 To iLastRow
        If UCase(TextBox1) <> UCase(TextBox5) And TextBox5 <> "" _
            And (CStr(Cells(i + 1, 2).Value) = UCase(TextBox5) And CStr(Cells(i + 1, 3).Value) <> "") Then
                Label8 = "OCCUPIED"
                Label8.ForeColor = RGB(255, 0, 0)
        ElseIf UCase(TextBox1) = CStr(Cells(i + 1, 2)) Then
            Label3 = Cells(i + 1, 3).Value
            Exit For
        ElseIf TextBox1 = "" Then
            Label3 = ""
            Exit For
        Else
            Label3 = ""
        End If
    Next i
    Me.Repaint
End Sub

'This function is called when there is any change in TextBox2 (Move To Door to Yard).
'The function searches for an exact match in TextBox1 IGNORING CASE SENSITIVITY from Cell C1 down to the last used cell in column C.
'When a match is found the program will examine the cell one space to the right of the match and determine if the cell is empty or not.
'   if the cell is empty then a green "AVAILABLE" will be put into Label4, if it has any value in it then it will be filled with a red "OCCUPIED"
Private Sub TextBox2_Change()
    Dim iLastRow As Integer
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    If UCase(TextBox1) = UCase(TextBox5) And UCase(TextBox2) = UCase(TextBox3) And TextBox1 <> "" And TextBox2 <> "" And TextBox3 <> "" And TextBox5 <> "" Then
        Label4 = "DOUBLE DROP"
        Label4.ForeColor = RGB(0, 0, 0)
        Label8 = "DOUBLE DROP"
        Label8.ForeColor = RGB(0, 0, 0)
        Exit Sub
    ElseIf UCase(TextBox2) = UCase(TextBox5) And TextBox5 <> "" Then
        Label8 = "OCCUPIED"
        Label8.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    For i = 1 To iLastRow
        If UCase(TextBox1) = UCase(TextBox2) And TextBox2 <> "" Then
            Label4 = "SAME LOCATION"
            Label4.ForeColor = RGB(255, 0, 0)
            Exit For
        ElseIf UCase(TextBox2) = CStr(Cells(i + 1, 2)) And Cells(i + 1, 3).Value = "" Then
            Label4 = "AVAILABLE"
            Label4.ForeColor = RGB(30, 255, 120)
            Exit For
        ElseIf UCase(TextBox2) = CStr(Cells(i + 1, 2)) And Cells(i + 1, 3).Value <> "" Then
            Label4 = "OCCUPIED"
            Label4.ForeColor = RGB(255, 0, 0)
            Exit For
        ElseIf UCase(TextBox2) = "" Then
            Label4 = ""
            Exit For
        Else
            Label4 = ""
        End If
    Next i
    Me.Repaint
End Sub

'This function is called when there is any change in TextBox3 (Move From Yard to Door).
'The function searches for an exact match in TextBox1 IGNORING CASE SENSITIVITY from Cell C1 down to the last used cell in column C.
'When a match is found Label11 will be filled with the value of cell one space to the right of the matched value (Trailer Number).
Private Sub TextBox3_Change()
    Dim iLastRow As Integer
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    If UCase(TextBox1) = UCase(TextBox5) And UCase(TextBox2) = UCase(TextBox3) And TextBox1 <> "" And TextBox2 <> "" And TextBox3 <> "" And TextBox5 <> "" Then
        Label4 = "DOUBLE DROP"
        Label4.ForeColor = RGB(0, 0, 0)
        Label8 = "DOUBLE DROP"
        Label8.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    For i = 1 To iLastRow
        If UCase(TextBox3) = UCase(TextBox5) And TextBox3 <> "" Then
            Label8 = "SAME LOCATION"
            Label8.ForeColor = RGB(255, 0, 0)
            Exit For
        ElseIf UCase(TextBox3) = CStr(Cells(i + 1, 2)) Then
            Label11 = Cells(i + 1, 3).Value
            Exit For
        ElseIf TextBox3 = "" Then
            Label11 = ""
            Exit For
        Else
            Label11 = ""
        End If
    Next i
    Me.Repaint
End Sub

'This function is called when there is any change in TextBox4 (DC Yard to Door).
'If the set of characters entered matches any string in the array at any time then the first found string will be put into Label9.
Private Sub TextBox4_Change()
    Dim arrDCList() As Variant
    arrDCList = Array("BLOOMFIELD", "BRIDGEWATER", "BROWNSBURG", _
        "CHARLOTTE", "COMPTON", "DECATUR", "EVANSVILLE", "JEFFERSON", "LAS VEGAS", _
        "MEMPHIS", "PHILADELPHIA", "PITTSTON", "PHOENIX", _
        "SAN PEDRO", "TUCSON", "WOBURN", "WORCESTER")
    For i = 0 To UBound(arrDCList) - LBound(arrDCList)
        If UCase(arrDCList(i)) Like UCase(TextBox4) & "*" Then
            Label9 = arrDCList(i)
            Exit For
        Else
            Label9 = ""
        End If
    Next i
    If TextBox4 = "" Then
        Label9 = ""
    End If
    Me.Repaint
End Sub

'This function is called when there is any change in TextBox2 (Move To Door to Yard).
'The function searches for an exact match in TextBox1 IGNORING CASE SENSITIVITY from Cell C1 down to the last used cell in column C.
'When a match is found the program will examine the cell one space to the right of the match and determine if the cell is empty or not.
'   if the cell is empty then a green "AVAILABLE" will be put into Label8, if it has any value in it then it will be filled with a red "OCCUPIED"
Private Sub TextBox5_Change()
    Dim iLastRow As Integer
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    If UCase(TextBox1) = UCase(TextBox5) And UCase(TextBox2) = UCase(TextBox3) And TextBox1 <> "" And TextBox2 <> "" And TextBox3 <> "" And TextBox5 <> "" Then
        Label4 = "DOUBLE DROP"
        Label4.ForeColor = RGB(0, 0, 0)
        Label8 = "DOUBLE DROP"
        Label8.ForeColor = RGB(0, 0, 0)
        Exit Sub
    ElseIf UCase(TextBox5) = UCase(TextBox1) And TextBox1 <> "" Then
        Label8 = "AVAILABLE"
        Label8.ForeColor = RGB(30, 255, 120)
        Exit Sub
    ElseIf UCase(TextBox2) = UCase(TextBox5) And TextBox2 <> "" Then
        Label8 = "OCCUPIED"
        Label8.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    For i = 1 To iLastRow
        If UCase(TextBox3) = UCase(TextBox5) And UCase(TextBox5) <> "" Then
            Label8 = "SAME LOCATION"
            Label8.ForeColor = RGB(255, 0, 0)
            Exit For
        ElseIf UCase(TextBox5) = CStr(Cells(i + 1, 2)) And Cells(i + 1, 3).Value = "" Then
            Label8 = "AVAILABLE"
            Label8.ForeColor = RGB(30, 255, 120)
            Exit For
        ElseIf UCase(TextBox5) = CStr(Cells(i + 1, 2)) And Cells(i + 1, 3).Value <> "" Then
            Label8 = "OCCUPIED"
            Label8.ForeColor = RGB(255, 0, 0)
            Exit For
        ElseIf UCase(TextBox5) = "" Then
            Label8 = ""
            Exit For
        Else
            Label8 = ""
        End If
    Next i
    Me.Repaint
End Sub

'This function is called after leaving TextBox1. The function prevents users from entering invalid door numbers and causing errors.
'The function gets the string value of TextBox1 and compares it to all the elements in the array IGNORING CASE SENSITIVITY.
'   When a match is found the boolean validDoor is set to true.
'   If no match is found then Label3 is set to "INVALID"
Private Sub TextBox1_AfterUpdate()
    Dim validDoor As Boolean
    validDoor = False
    Dim doorNumbers() As Variant
    doorNumbers = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", _
                "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", _
                "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", _
                "41", "42", "43", "44", "45", "DY-02", "DY-03", "DY-04", "DY-05", "DY-06", "DY-07", _
                "DY-08", "DY-09", "DY-10", "DY-11", "DY-12", "DY-13", "DY-14", "DY-15", "DY-16", _
                "DY-17", "DY-18", "DY-19", "DY-20", "DY-21", "DY-22", "DY-23", "DY-24")
    For Each e In doorNumbers
        If UCase(TextBox1) = e Then
            validDoor = True
            Exit For
        End If
    Next e
    If TextBox1 <> "" And Not validDoor Then
        Label3 = "INVALID"
    End If
End Sub

'This function is called after leaving TextBox5. The function prevents users from entering invalid door numbers and causing errors.
'The function gets the string value of TextBox1 and compares it to all the elements in the array IGNORING CASE SENSITIVITY.
'   When a match is found the boolean validDoor is set to true.
'   If no match is found then Label3 is set to a black "INVALID"
Private Sub TextBox5_AfterUpdate()
    Dim validDoor As Boolean
    validDoor = False
    Dim doorNumbers() As Variant
    doorNumbers = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", _
                "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", _
                "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", _
                "41", "42", "43", "44", "45", "DY-02", "DY-03", "DY-04", "DY-05", "DY-06", "DY-07", _
                "DY-08", "DY-09", "DY-10", "DY-11", "DY-12", "DY-13", "DY-14", "DY-15", "DY-16", _
                "DY-17", "DY-18", "DY-19", "DY-20", "DY-21", "DY-22", "DY-23", "DY-24")
    For Each e In doorNumbers
        If UCase(TextBox5) = e Then
            validDoor = True
            Exit For
        End If
    Next e
    If TextBox5 <> "" And Not validDoor Then
        Label8.ForeColor = RGB(0, 0, 0)
        Label8 = "INVALID"
    End If
End Sub

'___________________________________________________________________________________________________________________
'                           SECTION 3: MOVE FUNCTIONS
'___________________________________________________________________________________________________________________
'This function is called when Label3 is not empty and Label4 says "AVAILABLE" and TextBox3 is EMPTY and TextBox5 is EMPTY
'   If Label3 <> "" And Label4 = "AVAILABLE" And TextBox3 = "" And TextBox5 = ""
'The function logs a before move copy into the 'MOVE LOG' sheet, changes the status's of the trailer then copies
'   the necessary cells and pastes them into the destination then logs an after move copy into the 'MOVE LOG' sheet.
Private Sub doorToYard()
    Dim trailer As String
    Dim logCopy As Range
    Dim rng As Range
    Dim dest As String
    Dim dest2 As String
    Dim origin As String
    Dim origin2 As String
    Dim iLastRow As Integer
    trailer = Label3
    Set rng = Nothing
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    'Get destination location
    For i = 1 To iLastRow
        If UCase(TextBox2) = CStr(Cells(i, 2).Value) Then
            dest = Cells(i, 3).Address
            dest2 = Range(dest).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Get original location
    For i = 1 To iLastRow
        If UCase(TextBox1) = CStr(Cells(i, 2).Value) Then
            origin = Cells(i, 3).Address
            origin2 = Range(origin).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Store copy before moves to keep track of data
    Set logCopy = Range(Range(origin).Address & ":" & origin2)
    Call logBeforeDtoY(logCopy)
    
    'If status is "FULL" change to EMPTY with yellow background otherwise do nothing
    If Range(origin).Offset(0, 8).Value = "FULL" Then
        Range(origin).Offset(0, 8).Value = "EMPTY"
        Range(origin).Offset(0, 8).Interior.Color = RGB(255, 255, 0)
    End If

    'If status contains '%' then change to LOADED with green background
    If Range(origin).Offset(0, 8).Text Like "*%" Then
        Range(origin).Offset(0, 8).Value = "LOADED"
        Range(origin).Offset(0, 8).Interior.Color = RGB(146, 208, 80)
    End If
    
    'Put date stamp in H column
    Range(origin).Offset(0, 5).Value = Date
    
    'If M value is "" then M = PENDING / if N value is "" then N = PENDING
    If Range(origin).Offset(0, 10).Value = "" And Range(origin).Offset(0, 8).Value <> "EMPTY" Then
        Range(origin).Offset(0, 10).Value = "PENDING"
    End If
    If Range(origin).Offset(0, 11).Value = "" And Range(origin).Offset(0, 8).Value <> "EMPTY" Then
        Range(origin).Offset(0, 11).Value = "PENDING"
    End If
    
    'Select col C - N, copy and paste, go back to original loc and clear contents
    For i = 1 To iLastRow
        If trailer = CStr(Cells(i, 3).Value) Then
            Set rng = Worksheets("YARD").Range("C" & i & ":" & "N" & i)
            rng.Copy
            Range(dest).PasteSpecial xlPasteAll
            Range("C" & i & ":" & "N" & i).Interior.Color = RGB(255, 255, 255)
            Range("C" & i & ":" & "N" & i).Select
            Selection.ClearContents
            Exit For
        End If
    Next i
    
    'Store copy after moves to keep track of data
    Set logCopy = Range(Range(dest).Address & ":" & dest2)
    Call logAfterDtoY(logCopy)
    Range(dest).Select
End Sub

'This function is called when Label11 is not empty and Label8 says "AVAILABLE" and TextBox1 is EMPTY and TextBox2 is EMPTY
'   If Label11 <> "" And Label8 = "AVAILABLE" And TextBox1 = "" And TextBox2 = ""
'The function logs a before move copy into the 'MOVE LOG' sheet, changes the status's of the trailer then copies
'   the necessary cells and pastes them into the destination then logs an after move copy into the 'MOVE LOG' sheet.
Private Sub yardToDoor()
    Dim trailer As String
    Dim logCopy As Range
    Dim rng As Range
    Dim dest As String
    Dim dest2 As String
    Dim origin As String
    Dim origin2 As String
    Dim iLastRow As Integer
    trailer = Label11
    Set rng = Nothing
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    'Get destination location
    For i = 1 To iLastRow
        If UCase(TextBox5) = CStr(Cells(i, 2).Value) Then
            dest = Cells(i, 3).Address
            dest2 = Range(dest).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Get original location
    For i = 1 To iLastRow
        If trailer = CStr(Cells(i, 3).Value) Then
            origin = Cells(i, 3).Address
            origin2 = Range(origin).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Store copy before moves to keep track of data
    Set logCopy = Range(Range(origin).Address & ":" & origin2)
    Call logBeforeYtoD(logCopy)
    
    'If status is "DROP" change to 0%
    If Range(origin).Offset(0, 8).Value = "DROP" Then
        Range(origin).Offset(0, 8).Value = "0"
    End If
    
    'Put date stamp in G column
    Range(origin).Offset(0, 4).Value = Date
    
    'Put DC into cell
    If Range(origin).Offset(0, 8).Text Like "*%" Then
        Range(origin).Offset(0, 9).Value = Label9
    End If
    'Select col C - N, copy and paste, go back to original loc and clear contents
    For i = 1 To iLastRow
        If trailer = CStr(Cells(i, 3).Value) Then
            Set rng = Worksheets("YARD").Range("C" & i & ":" & "N" & i)
            rng.Copy
            Range(dest).PasteSpecial xlPasteAll
            Range("C" & i & ":" & "N" & i).Interior.Color = RGB(255, 255, 255)
            Range("C" & i & ":" & "N" & i).Select
            Selection.ClearContents
            Exit For
        End If
    Next i
    
    'Store copy after moves to keep track of data
    Set logCopy = Range(Range(dest).Address & ":" & dest2)
    Call logAfterYtoD(logCopy)
    Range(dest).Select
End Sub

'This function is called when Label4 says "DOUBLE DROP" and Label8 says "DOUBLE DROP" and Label3 is not EMPTY and Label11 is not EMPTY
'   Label4 = "DOUBLE DROP" And Label8 = "DOUBLE DROP" And Label3 <> "" And Label11 <> ""
'The function logs a before move copy of the trailer at the door, then changes the status's and logs an after move copy of the trailer
'   and holds the cell range of the trailer in a holder variable.
'It then logs a before move copy of the trailer in the yard , then changes the status's and logs an after move copy of the trailer
'   and holds the cell range of the trailer in a holder variable.
'It then pastes the data into the necessary ranges of the Yard Check.
Private Sub swapTrailers()
    Dim trailer1 As String
    Dim trailer2 As String
    Dim yardCopy As Range
    Dim doorCopy As Range
    Dim logCopy As Range
    Dim holder(2) As Variant
    Dim yardLoc As String
    Dim yardLoc2 As String
    Dim doorLoc As String
    Dim doorLoc2 As String
    Dim iLastRow As Integer
    trailer1 = Label3
    trailer2 = Label11
    iLastRow = Worksheets("YARD").Cells(Rows.Count, 2).End(xlUp).Row
    
    'Get destination Cell for trailer1 YARD LOCATION
    For i = 1 To iLastRow
        If trailer2 = CStr(Cells(i, 3).Value) Then
            yardLoc = Cells(i, 3).Address
            yardLoc2 = Range(yardLoc).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Get original Cell for trailer2 DOOR LOC
    For i = 1 To iLastRow
        If trailer1 = CStr(Cells(i, 3).Value) Then
            doorLoc = Cells(i, 3).Address
            doorLoc2 = Range(doorLoc).Offset(0, 11).Address
            Exit For
        End If
    Next i
    
    'Store copy before moves to keep track of data
    Set logCopy = Range(Range(doorLoc).Address & ":" & doorLoc2)
    Call logBeforeDtoY(logCopy)
    
    'If status contains '%' then change to LOADED and put green background
    If Range(doorLoc).Offset(0, 8).Text Like "*%" Then
        Range(doorLoc).Offset(0, 8).Value = "LOADED"
        Range(doorLoc).Offset(0, 8).Interior.Color = RGB(146, 208, 80)
    End If
    
    'Put date stamp in H column
    Range(doorLoc).Offset(0, 5).Value = Date
    
    'If M value is "" then M = PENDING / if N value is "" then N = PENDING
    If Range(doorLoc).Offset(0, 10).Value = "" And Range(doorLoc).Offset(0, 8).Value = "LOADED" Then
        Range(doorLoc).Offset(0, 10).Value = "PENDING"
    End If
    If Range(doorLoc).Offset(0, 11).Value = "" And Range(doorLoc).Offset(0, 8).Value = "LOADED" Then
        Range(doorLoc).Offset(0, 11).Value = "PENDING"
    End If
    
    'If status is "FULL" change to EMPTY with yellow background at yard location and make door location white background
    If Range(doorLoc).Offset(0, 8).Value = "FULL" Then
        Range(doorLoc).Offset(0, 8).Value = "EMPTY"
        Range(doorLoc).Offset(0, 8).Interior.Color = RGB(255, 255, 0)
    End If
    
    'Store copy after moves to keep track of data
    Set logCopy = Range(Range(doorLoc).Address & ":" & doorLoc2)
    Call logAfterDtoY(logCopy)
    
    ' Make interior color of door location white
    Range(doorLoc & ":" & doorLoc2).Interior.Color = RGB(255, 255, 255)
    
    'Put DC into cell of yard trailer
    Range(yardLoc).Offset(0, 9).Value = Label9
    
    'Now get all the door trailer information and paste into yard
    
    'Store copy before moves to keep track of data
    Set logCopy = Range(Range(yardLoc).Address & ":" & yardLoc2)
    Call logBeforeYtoD(logCopy)
    
    'If status is "DROP" change to 0% of yard trailer
    If Range(yardLoc).Offset(0, 8).Value = "DROP" Then
        Range(yardLoc).Offset(0, 8).Value = "0"
    End If
    
    'Put date stamp in G column of yard trailer
    Range(yardLoc).Offset(0, 4).Value = Date
    
    'Store copy after moves to keep track of data
    Set logCopy = Range(Range(yardLoc).Address & ":" & yardLoc2)
    Call logAfterYtoD(logCopy)
    
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    Set yardCopy = Range(yardLoc & ":" & yardLoc2)
    Set doorCopy = Range(doorLoc & ":" & doorLoc2)
    
    holder(0) = yardCopy
    holder(1) = doorCopy
    yardCopy = holder(1)
    doorCopy = holder(0)
    
    'If STATUS is loaded or full make green, if EMPTY, make yellow, else make white
    If Range(yardLoc).Offset(0, 8).Value = "LOADED" Then
        Range(yardLoc).Offset(0, 8).Interior.Color = RGB(146, 208, 80)
    ElseIf Range(yardLoc).Offset(0, 8).Value = "EMPTY" Then
        Range(yardLoc).Offset(0, 8).Interior.Color = RGB(255, 255, 0)
        Range(yardLoc).Offset(0, 10).Value = ""
        Range(yardLoc).Offset(0, 11).Value = ""
    Else
        Range(yardLoc).Offset(0, 8).Interior.Color = RGB(255, 255, 255)
    End If
    
    'If STATUS is full make green, else make white
    If Range(doorLoc).Offset(0, 8).Value = "FULL" Then
        Range(doorLoc).Offset(0, 8).Interior.Color = RGB(146, 208, 80)
    Else
        Range(doorLoc).Offset(0, 8).Interior.Color = RGB(255, 255, 255)
    End If

    Range(doorLoc).Select
End Sub

'___________________________________________________________________________________________________________________
'                           SECTION 4: LOGGING THE MOVES IN THE MOVE LOG
'___________________________________________________________________________________________________________________
'This function is called in the doorToYard function.
'The function creates a copy of the trailer and its status's before any changes in the 'MOVE LOG' sheet.
Private Sub logBeforeDtoY(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Rows("2:3").EntireRow.Insert
        .Range("2:3").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("2:3").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("2:3").Borders.LineStyle = xlContinuous
        .Range("2:3").Borders.Weight = xlThin
        .Range("A2").Value = "BEFORE"
        .Range("B2").Value = UCase(TextBox1)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'This function is called in the doorToYard function.
'The function creates a copy of the trailer and its status's after all changes in the 'MOVE LOG' sheet.
Private Sub logAfterDtoY(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Range("A2").EntireRow.Insert
        .Range("A2").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("A2").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("A2").Borders.LineStyle = xlContinuous
        .Range("A2").Borders.Weight = xlThin
        .Range("A2").Value = "AFTER"
        .Range("B2").Value = UCase(TextBox2)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'This function is called in the yardToDoor function.
'The function creates a copy of the trailer and its status's before any changes in the 'MOVE LOG' sheet.
Private Sub logBeforeYtoD(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Rows("2:3").EntireRow.Insert
        .Range("2:3").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("2:3").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("2:3").Borders.LineStyle = xlContinuous
        .Range("2:3").Borders.Weight = xlThin
        .Range("A2").Value = "BEFORE"
        .Range("B2").Value = UCase(TextBox3)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'This function is called in the yardToDoor function.
'The function creates a copy of the trailer and its status's after all changes in the 'MOVE LOG' sheet.
Private Sub logAfterYtoD(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Range("A2").EntireRow.Insert
        .Range("A2").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("A2").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("A2").Borders.LineStyle = xlContinuous
        .Range("A2").Borders.Weight = xlThin
        .Range("A2").Value = "AFTER"
        .Range("B2").Value = UCase(TextBox5)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'This function is called in the swapTrailers function.
'The function creates a copy of the trailer and its status's before any changes in the 'MOVE LOG' sheet.
Private Sub logBeforeSwap(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Rows("2:3").EntireRow.Insert
        .Range("2:3").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("2:3").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("2:3").Borders.LineStyle = xlContinuous
        .Range("2:3").Borders.Weight = xlThin
        .Range("A2").Value = "BEFORE"
        .Range("B2").Value = UCase(TextBox1)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'This function is called in the swapTrailers function.
'The function creates a copy of the trailer and its status's after all changes in the 'MOVE LOG' sheet.
Private Sub logAfterSwap(rng As Range)
    Dim wkbLog As Worksheet
    Set wkbLog = Worksheets("MOVE LOG")
    
    With wkbLog
        .Range("A2").EntireRow.Insert
        .Range("A2").EntireRow.Interior.Color = RGB(255, 255, 255)
        .Range("A2").EntireRow.Font.Color = RGB(0, 0, 0)
        .Range("A2").Borders.LineStyle = xlContinuous
        .Range("A2").Borders.Weight = xlThin
        .Range("A2").Value = "AFTER"
        .Range("B2").Value = UCase(TextBox2)
        rng.Copy
        .Range("C2").PasteSpecial xlPasteAll
        .Range("O2") = Time
    End With
End Sub

'___________________________________________________________________________________________________________________
'                           SECTION 5: KEYDOWN FUNCTIONS (ENTER KEY)
'___________________________________________________________________________________________________________________
'This function is called when the 'Enter' button is pressed while inside TextBox1
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        okButton_Click
    End If
End Sub

'This function is called when the 'Enter' button is pressed while inside TextBox2
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        okButton_Click
    End If
End Sub

'This function is called when the 'Enter' button is pressed while inside TextBox3
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        okButton_Click
    End If
End Sub

'This function is called when the 'Enter' button is pressed while inside TextBox4
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        okButton_Click
    End If
End Sub

'This function is called when the 'Enter' button is pressed while inside TextBox5
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
       okButton_Click
    End If
End Sub



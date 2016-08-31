Option Explicit

Sub load_data_errors()

Dim counter1 As Long
Dim counter2 As Long
Dim opsarr() As String
Dim opsuniq() As String
Dim x As Integer, y As Integer

counter1 = 1
counter2 = 1

ReDim opsarr(counter1) As String
ReDim opsuniq(counter2) As String

For x = 1 To 1000
With Sheet1
    
    If InStr(.Cells(x, 1).Value, " ") > 0 And Left(.Cells(x, 1).Value, 6) <> "PICKER" And .Cells(x, 1).Value <> "" Then
    opsarr(counter1) = .Cells(x, 1).Value
    
        If counter1 = 1 Then
        
            opsuniq(counter2) = opsarr(counter1)
            counter2 = counter2 + 1
            ReDim Preserve opsuniq(counter2) As String
            
        ElseIf counter1 > 1 And opsarr(counter1) <> opsarr(counter1 - 1) Then
        
            opsuniq(counter2) = opsarr(counter1)
            counter2 = counter2 + 1
            ReDim Preserve opsuniq(counter2) As String
            
        End If
    
    counter1 = counter1 + 1
    ReDim Preserve opsarr(counter1) As String
    End If
    
End With
Next x

For y = 1 To counter2
    
    usf_errors.lsb_ops.AddItem opsuniq(y)

Next y

usf_errors.cmb_weekday.AddItem "Sunday"
usf_errors.cmb_weekday.AddItem "Monday"
usf_errors.cmb_weekday.AddItem "Tuesday"
usf_errors.cmb_weekday.AddItem "Wednesday"
usf_errors.cmb_weekday.AddItem "Thursday"
usf_errors.cmb_weekday.AddItem "Friday"

usf_errors.cmb_pick_errors.AddItem "Local"
usf_errors.cmb_pick_errors.AddItem "Remote"

End Sub
Sub load_data()

Dim counter1 As Long
Dim counter2 As Long
Dim opsarr() As String
Dim opsuniq() As String
Dim x As Integer, y As Integer

counter1 = 1
counter2 = 1

ReDim opsarr(counter1) As String
ReDim opsuniq(counter2) As String

For x = 1 To 1000
With Sheet1
    
    If InStr(.Cells(x, 1).Value, " ") > 0 And Left(.Cells(x, 1).Value, 6) <> "PICKER" And .Cells(x, 1).Value <> "" Then
    opsarr(counter1) = .Cells(x, 1).Value
    
        If counter1 = 1 Then
        
            opsuniq(counter2) = opsarr(counter1)
            counter2 = counter2 + 1
            ReDim Preserve opsuniq(counter2) As String
            
        ElseIf counter1 > 1 And opsarr(counter1) <> opsarr(counter1 - 1) Then
        
            opsuniq(counter2) = opsarr(counter1)
            counter2 = counter2 + 1
            ReDim Preserve opsuniq(counter2) As String
            
        End If
    
    counter1 = counter1 + 1
    ReDim Preserve opsarr(counter1) As String
    End If
    
End With
Next x

For y = 1 To counter2
    
    usf_main.lsb_ops.AddItem opsuniq(y)

Next y

usf_main.cmb_weekday.AddItem "Sunday"
usf_main.cmb_weekday.AddItem "Monday"
usf_main.cmb_weekday.AddItem "Tuesday"
usf_main.cmb_weekday.AddItem "Wednesday"
usf_main.cmb_weekday.AddItem "Thursday"
usf_main.cmb_weekday.AddItem "Friday"

usf_main.cmb_pick.AddItem "Local"
usf_main.cmb_pick.AddItem "Remote"

End Sub


Sub picktype()

If usf_main.cmb_pick.Value = "Local" Then
usf_main.cmb_picktype.Clear
usf_main.cmb_picktype.AddItem "Core Ambient"
usf_main.cmb_picktype.AddItem "Core Chill"
usf_main.cmb_picktype.AddItem "Core Freezer"
usf_main.cmb_picktype.AddItem "TRG Ambient"
usf_main.cmb_picktype.AddItem "TRG Chill"
usf_main.cmb_picktype.AddItem "TRG Freezer"
Else
usf_main.cmb_picktype.Clear
usf_main.cmb_picktype.AddItem "Ambient"
usf_main.cmb_picktype.AddItem "Veg"
usf_main.cmb_picktype.AddItem "Grid D1D3"
usf_main.cmb_picktype.AddItem "Grid D1D2"
usf_main.cmb_picktype.AddItem "Freezer"
End If

End Sub
Sub picktype_errors()

If usf_errors.cmb_pick_errors.Value = "Local" Then
usf_errors.cmb_picktype.Clear
usf_errors.cmb_picktype.AddItem "Core Ambient"
usf_errors.cmb_picktype.AddItem "Core Chill"
usf_errors.cmb_picktype.AddItem "Core Freezer"
usf_errors.cmb_picktype.AddItem "TRG Ambient"
usf_errors.cmb_picktype.AddItem "TRG Chill"
usf_errors.cmb_picktype.AddItem "TRG Freezer"
Else
usf_errors.cmb_picktype.Clear
usf_errors.cmb_picktype.AddItem "Ambient"
usf_errors.cmb_picktype.AddItem "Veg"
usf_errors.cmb_picktype.AddItem "Grid D1D3"
usf_errors.cmb_picktype.AddItem "Grid D1D2"
usf_errors.cmb_picktype.AddItem "Freezer"
End If

End Sub

Sub formstart()
usf_main.Show
End Sub

Sub form_erros()
usf_errors.Show
End Sub

Sub voicedownload() 'for night shift only, will need slight alteration in check for how many days run for day shift

Dim counter1 As Long, counter2 As Long, counter3 As Long
Dim x As Long, y As Long, z As Long, q As Long
Dim firstname As String, lastname As String
Dim finder As Range
Dim tempholder As String
Dim pickerarr() As Variant
Dim findday As String
Dim dayplusone As String
Dim weekd1 As Integer, weekd2 As Integer
Dim startdate As Date, finishdate As Date
Dim who As String
Dim where As String
Dim cases As Long
Dim hours As Double

With Sheet6
        If .Cells(2, 1).Value = "" Or Left(.Cells(2, 1).Value, 4) <> "Pick" Then
        MsgBox "Please paste InfoCentre download to cell A1 and re-run."
        Exit Sub
        End If
        
        dayplusone = Left(Right(.Cells(2, 1).Value, 17), 11)
        findday = Right(Left(.Cells(2, 1).Value, 29), 11)
        weekd1 = weekday(CDate(findday))
        weekd2 = weekday(CDate(dayplusone))
        
        'in case info centre was run for more than 1 day [night shift specific], if used for days this need to be reduced to same day
        If CInt(Left(dayplusone, 2)) - CInt(Left(findday, 2)) > 1 Or weekd2 - weekd1 > 1 Then
        MsgBox "Please re-run InfoCentre report. Current report was run for more than 1 night."
        Exit Sub
        End If
                
        Select Case weekd1
            Case 1
            findday = "Sunday"
            Case 2
            findday = "Monday"
            Case 3
            findday = "Tuesday"
            Case 4
            findday = "Wednesday"
            Case 5
            findday = "Thursday"
            Case 6
            findday = "Friday"
            Case 7
            findday = "Saturday"
        End Select
        
        For x = 1 To 1000
            If Left(.Cells(x, 1).Value, 4) <> "Pick" _
            And Left(.Cells(x, 1).Value, 4) <> "Zone" _
            And Left(.Cells(x, 1).Value, 8) <> "Employee" _
            And Left(.Cells(x, 1).Offset(0, 1).Value, 4) = "Area" _
            And .Cells(x, 1).Offset(0, 1).Value <> "" _
            Then
                If .Cells(x, 1).Value <> "" _
                Then
                    tempholder = .Cells(x, 1).Value
                Else
                    .Cells(x, 1).Value = tempholder
                End If
            counter1 = counter1 + 1
            End If
        Next x
        
        ReDim pickerarr(1 To counter1, 1 To 4) As Variant
        
        counter2 = 1
        
        For y = 1 To 1000
        
            If Left(.Cells(y, 1).Value, 4) <> "Pick" _
            And Left(.Cells(y, 1).Value, 4) <> "Zone" _
            And Left(.Cells(y, 1).Value, 8) <> "Employee" _
            And Left(.Cells(y, 1).Offset(0, 1).Value, 4) = "Area" _
            And .Cells(y, 1).Offset(0, 2).Value <> "" _
            And .Cells(y, 1).Offset(0, 2).Value <> 0 _
            And .Cells(y, 1).Offset(0, 4).Value <> "" _
            And .Cells(y, 1).Offset(0, 5).Value <> "" _
            And .Cells(y, 1).Offset(0, 6).Value <> "" _
            Then
            
            ' Grab name, capitalize and invert order (surname & name)
            tempholder = .Cells(y, 1).Value
            firstname = Left(tempholder, InStr(tempholder, " ") - 1)
            firstname = UCase(Left(firstname, 1)) + Right(firstname, Len(firstname) - 1)
            lastname = Right(tempholder, Len(tempholder) - InStr(tempholder, " "))
            lastname = UCase(Left(lastname, 1)) + Right(lastname, Len(lastname) - 1)
            
            pickerarr(counter2, 1) = lastname + " " + firstname
            
            tempholder = Right(.Cells(y, 2).Value, 4)
            Select Case tempholder
            Case 1001
            pickerarr(counter2, 2) = "Local Core Freezer"
            Case 1002
            pickerarr(counter2, 2) = "Local TRG Freezer"
            Case 8031
            pickerarr(counter2, 2) = "Local Core Chill"
            Case 8032
            pickerarr(counter2, 2) = "Local TRG Chill"
            Case 9001
            pickerarr(counter2, 2) = "Local Core Ambient"
            Case 9002
            pickerarr(counter2, 2) = "Local TRG Ambient"
            End Select
            ' Total time picked (without downtime)
            pickerarr(counter2, 3) = Round(CInt(Left(.Cells(y, 3).Value, 3)) + Round((CInt(Left(Right(.Cells(y, 3).Value, 5), 2)) / 60) + Round(CInt(Right(.Cells(y, 3).Value, 2)) / 3600, 2), 2), 2)
            '+ downtime
            'pickerarr(counter2, 3) = Round(CInt(Left(.Cells(y, 13).Value, 3)) + Round((CInt(Left(Right(.Cells(y, 13).Value, 5), 2)) / 60) + Round(CInt(Right(.Cells(y, 13).Value, 2)) / 3600, 2), 2), 2)
            ' Total cases picked
            pickerarr(counter2, 4) = .Cells(y, 7).Value
            counter2 = counter2 + 1
            
            End If
            
        Next y
        
        counter3 = 2
End With


With Sheet7
            .Range(.Cells(5, 5), .Cells(117, 5)).ClearContents
        
        
        For z = 1 To counter1
        
            If z = 1 Then
                .Cells(z + 4, 5).Value = pickerarr(z, 1)
            Else
            Set finder = .Range(.Cells(5, 5), .Cells(z + 4, 5)).Find(pickerarr(z, 1), LookIn:=xlValues, lookat:=xlWhole)
                If Not finder Is Nothing Then
                GoTo skipthis
                Else
                .Cells(counter3 + 4, 5).Value = pickerarr(z, 1)
                counter3 = counter3 + 1
                End If
            End If
skipthis:
        Next z
End With

For q = 1 To counter1


    
    With Sheet7
    
        If pickerarr(q, 3) = 0 Or pickerarr(q, 4) = 0 Then
        GoTo skipthis2
        Else
        who = pickerarr(q, 1)
        where = pickerarr(q, 2)
        cases = pickerarr(q, 4)
        hours = pickerarr(q, 3)
    
        Call the_great_finder(who, findday, where, hours, cases)
        
        'test printing output
        '.Cells(q, 7).Value = pickerarr(q, 1)
        '.Cells(q, 8).Value = pickerarr(q, 2)
        '.Cells(q, 9).Value = pickerarr(q, 3)
        '.Cells(q, 10).Value = pickerarr(q, 4)
        End If
        
    End With
skipthis2:
Next q


Erase pickerarr

End Sub

Sub the_great_finder(ByVal name As String, whenpicked As String, wherepicked As String, howlong As Double, howmuch As Long)

'test run data
'name = "Bennett Sam"
'wherepicked = "Local Core Chill"
'whenpicked = "SUNDAY"
'howlong = 3.51
'howmuch = 939

Dim tempvalue As Long
Dim columncounter As Long
Dim x As Long, y As Long, z As Long, counter1 As Long
Dim rowtargeter1 As Range, rowtargeter2 As Range, coltargeter As Range
Dim lastrow As Long
Dim tempdouble As Double

With Sheet1

        lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
        
        If name = "Smoulch Leslek" Then
        name = "Smoluch Leszek"
        End If
        
                
        If name = "Coctinho Nelson" Then
        name = "Coutinho Nelson"
        End If
        
        If name = "Lewandowski Kristof" Then
        name = "Lewandowski Krzysztof"
        End If
        
        If name = "Bodyova Veronika" Then
        name = "Body Veronica"
        End If
        
        
        Set rowtargeter1 = .Range(.Cells(1, 1), .Cells(lastrow, 1)).Find(name, LookIn:=xlValues, lookat:=xlWhole)
        
        'in case the name is not on the sheet
            If rowtargeter1 Is Nothing Then
            MsgBox "Picker's name " + name + " could not be recognised."
            usf_main.tb_picker.Value = ""
            Exit Sub
            End If
        
        tempvalue = rowtargeter1.Row
        
        counter1 = 1
        
        'in case of end of the list of names in column A (there won't be more than a 100 pick types I don't think)
        Do Until rowtargeter1.Offset(1, 0).Value <> rowtargeter1.Value Or counter1 > 11
            counter1 = counter1 + 1
            Set rowtargeter1 = rowtargeter1.Offset(1, 0)
        Loop
        
        Set rowtargeter2 = .Range(.Cells(tempvalue, 2), .Cells(tempvalue - 1 + counter1, 2)).Find(wherepicked, LookIn:=xlValues, lookat:=xlWhole)
        
        'in case pick type was altered
            If rowtargeter2 Is Nothing Then
            MsgBox "Pick type was not recognised."
            usf_main.cmb_picktype.Value = ""
            Exit Sub
            End If
            
        x = rowtargeter2.Row

        Set coltargeter = .Range(.Cells(1, 1), .Cells(1, 100)).Find(whenpicked, LookIn:=xlValues, lookat:=xlPart)
        
        'in case of mistyped day of the week
            If coltargeter Is Nothing Then
            MsgBox "Day of the week could not be recognised."
            usf_main.cmb_weekday = ""
            Exit Sub
            End If
            
        columncounter = coltargeter.Column - 1
        tempdouble = howmuch / howlong
        .Cells(x, columncounter).Offset(0, 1).Value = howmuch
        .Cells(x, columncounter).Offset(0, 1).NumberFormat = "0"
        .Cells(x, columncounter).Offset(0, 2).Value = howlong
        .Cells(x, columncounter).Offset(0, 2).NumberFormat = "0.00"
        .Cells(x, columncounter).Offset(0, 3).Value = tempdouble
        .Cells(x, columncounter).Offset(0, 3).NumberFormat = "0.00"
        .Cells(x, columncounter).Offset(0, 4).Value = prod(tempdouble, wherepicked)
        .Cells(x, columncounter).Offset(0, 4).NumberFormat = "0.00%"
        
        If .Cells(x, columncounter).Offset(0, 5).Value = "" Then
        .Cells(x, columncounter).Offset(0, 5).Value = 0
        .Cells(x, columncounter).Offset(0, 6).Value = 1
        .Cells(x, columncounter).Offset(0, 6).NumberFormat = "0.00%"
        Else
        tempdouble = (howmuch - .Cells(x, columncounter).Offset(0, 5).Value) / howmuch
        tempdouble = Round(tempdouble, 2)
        .Cells(x, columncounter).Offset(0, 6).Value = tempdouble
        .Cells(x, columncounter).Offset(0, 6).NumberFormat = "0.00%"
        End If
        
End With

Set rowtargeter1 = Nothing
Set rowtargeter2 = Nothing


End Sub

Sub the_error_finder(ByVal name As String, whenpicked As String, wherepicked As String, howmuch As Long)

'test run data
'name = "Bennett Sam"
'wherepicked = "Local Core Chill"
'whenpicked = "SUNDAY"
'howmuch = 9

Dim tempvalue As Long
Dim columncounter As Long
Dim x As Long, y As Long, z As Long, counter1 As Long
Dim rowtargeter1 As Range, rowtargeter2 As Range, coltargeter As Range
Dim lastrow As Long
Dim tempdouble As Double

With Sheet1

        lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
        
        Set rowtargeter1 = .Range(.Cells(1, 1), .Cells(lastrow, 1)).Find(name, LookIn:=xlValues, lookat:=xlWhole)
        
        'in case the name is not on the sheet
            If rowtargeter1 Is Nothing Then
            MsgBox "Picker's name " + name + " could not be recognised."
            usf_main.tb_picker.Value = ""
            Exit Sub
            End If
        
        tempvalue = rowtargeter1.Row
        
        counter1 = 1
        
        'in case of end of the list of names in column A (there won't be more than a 100 pick types I don't think)
        Do Until rowtargeter1.Offset(1, 0).Value <> rowtargeter1.Value Or counter1 > 11
            counter1 = counter1 + 1
            Set rowtargeter1 = rowtargeter1.Offset(1, 0)
        Loop
        
        Set rowtargeter2 = .Range(.Cells(tempvalue, 2), .Cells(tempvalue - 1 + counter1, 2)).Find(wherepicked, LookIn:=xlValues, lookat:=xlWhole)
        
        'in case pick type was altered
            If rowtargeter2 Is Nothing Then
            MsgBox "Pick type was not recognised."
            usf_main.cmb_picktype.Value = ""
            Exit Sub
            End If
            
        x = rowtargeter2.Row

        Set coltargeter = .Range(.Cells(1, 1), .Cells(1, 100)).Find(whenpicked, LookIn:=xlValues, lookat:=xlPart)
        
        'in case of mistyped day of the week
            If coltargeter Is Nothing Then
            MsgBox "Day of the week could not be recognised."
            usf_main.cmb_weekday = ""
            Exit Sub
            End If
            
        columncounter = coltargeter.Column - 1
       
        .Cells(x, columncounter).Offset(0, 5).Value = howmuch
        
        If .Cells(x, columncounter).Offset(0, 1).Value = "" Or .Cells(x, columncounter).Offset(0, 1).Value = 0 Then
        .Cells(x, columncounter).Offset(0, 6).Value = ""
        .Cells(x, columncounter).Offset(0, 6).NumberFormat = "0.00%"
        Else
        tempdouble = (.Cells(x, columncounter).Offset(0, 1).Value - howmuch) / .Cells(x, columncounter).Offset(0, 1).Value
        tempdouble = Round(tempdouble, 2)
        .Cells(x, columncounter).Offset(0, 6).Value = tempdouble
        .Cells(x, columncounter).Offset(0, 6).NumberFormat = "0.00%"
        End If
               
End With

Set rowtargeter1 = Nothing
Set rowtargeter2 = Nothing


End Sub

Public Function ifNumeric(passedvalue) As Boolean

Dim tempholder As Variant
Dim nonumbersarray(10) As Variant
Dim r As Long
Dim matchcounter As Long

nonumbers = "q,w,e,r,t,y,u,i,o,p,[,],{,},a,s,d,f,g,h,j,k,l,;,:,',@,#,~,z,x,c,v,b,n,m,<,>,/,?,!,"",£,$,%,^,&,*,(,),_,-,=,+,Q,W,E,R,T,Y,U,I,O,P,A,S,D,F,G,H,J,K,L,Z,X,C,V,B,N,M"
tempholder = Split(nonumbers, ",")
matchcounter = 0

For r = 0 To UBound(tempholder)
    If InStr(passedvalue, tempholder(r)) <> 0 Then matchcounter = matchcounter + 1
    End If
Next r

If matchcounter > 0 Then ifNumeric = False
Else: ifNumeric = True
End If

End Function

Public Function prod(rate As Double, picktype As String) As Double

Dim x As Integer
Dim y As Double

    For x = 5 To 33
    
        With Sheet3
            
            If .Cells(x, 5).Value = picktype Then
            y = rate / .Cells(x, 6).Value
            End If
        
        End With
    
        prod = y
    
    Next x

End Function

Sub protection()

With Sheet1
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

With Sheet3
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

With Sheet4
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

With Sheet7
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

With Sheet8
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

With Sheet9
.Protect Password:="hellcat666", UserInterfaceOnly:=True
End With

End Sub

Sub leaderboards_load()

Dim pickerboard() As Variant
Dim z As Long, x As Long, y As Long
Dim lastrow As Long
Dim counter1 As Long, rowcounter As Long, counter2 As Long
Dim task As String

rowcounter = 0
counter2 = 0

With Sheet1
    
    lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For z = 3 To lastrow
    
        If .Cells(z, 39).Value <> "" And .Cells(z, 39).Value > 0 Then
        rowcounter = rowcounter + 1
        End If
    
    Next z
    
    
    ReDim pickerboard(1 To rowcounter, 1 To 8) As Variant

    For x = 3 To lastrow
    
        If .Cells(x, 39).Value > 0 And .Cells(x, 39).Value <> "" Then
            counter2 = counter2 + 1
            pickerboard(counter2, 1) = .Cells(x, 1).Value
            pickerboard(counter2, 2) = .Cells(x, 2).Value
            pickerboard(counter2, 3) = .Cells(x, 39).Value
            pickerboard(counter2, 4) = .Cells(x, 40).Value
            pickerboard(counter2, 5) = .Cells(x, 41).Value
            pickerboard(counter2, 6) = .Cells(x, 42).Value
            pickerboard(counter2, 7) = .Cells(x, 43).Value
            pickerboard(counter2, 8) = .Cells(x, 44).Value
        
        End If
    
    Next x

End With

counter1 = 1


With Sheet4
    
        .Range(.Cells(3, 3), .Cells(202, 10)).ClearContents
        
For y = 1 To counter2
        
        task = .Cells(1, 9).Value
        If pickerboard(y, 2) = task Or task = "All chambers" Then
        
            .Cells(counter1 + 2, 3).Value = pickerboard(y, 1)
            .Cells(counter1 + 2, 4).Value = pickerboard(y, 2)
            .Cells(counter1 + 2, 5).Value = pickerboard(y, 3)
            .Cells(counter1 + 2, 6).Value = pickerboard(y, 4)
            .Cells(counter1 + 2, 6).NumberFormat = "0.00"
            .Cells(counter1 + 2, 7).Value = pickerboard(y, 5)
            .Cells(counter1 + 2, 7).NumberFormat = "0.00"
            .Cells(counter1 + 2, 8).Value = pickerboard(y, 6)
            .Cells(counter1 + 2, 8).NumberFormat = "0.00%"
            .Cells(counter1 + 2, 9).Value = pickerboard(y, 7)
            .Cells(counter1 + 2, 10).Value = pickerboard(y, 8)
            .Cells(counter1 + 2, 10).NumberFormat = "0.00%"
            counter1 = counter1 + 1
            
        End If
    
Next y

End With

End Sub

Sub sort_name()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("C3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_cases()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("E3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_time()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("F3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_prod()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("H3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_errors()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("I3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_accuracy()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("J3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub
Sub sort_cph()

    Application.ScreenUpdating = False
    Range("C2:J201").Select
    Selection.Sort Key1:=Range("G3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("C3").Select
    Application.ScreenUpdating = True
    
End Sub

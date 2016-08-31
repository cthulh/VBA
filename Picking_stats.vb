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
        
            pickerarr(counter2, 3) = Round(CInt(Left(.Cells(y, 3).Value, 3)) + Round((CInt(Left(Right(.Cells(y, 3).Value, 5), 2)) / 60) + Round(CInt(Right(.Cells(y, 3).Value, 2)) / 3600, 2), 2), 2)
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

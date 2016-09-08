Sub delrow()
    Selection.Delete Shift:=xlUp
End Sub
Sub addrow()
    Selection.Insert Shift:=xlDown
End Sub

Sub forecast()

Dim lastrow As Long, lastcol As Long, x As Long, y As Long, z As Long, j As Long, k As Long, l As Long
Dim datafrom As Long, expirycriterion As Long
Dim expiredarr() As Variant
Dim comparisondate As Date, expirydate As Date, expiryperiod As String
Dim titlestr As String

With Sheet5
    If IsDate(.Cells(4, 8)) Then
        comparisondate = DateAdd("m", 3, .Cells(4, 8))
    Else
        MsgBox ("Invalid date format. Please use format DD/MM/YYY and re run macro.")
        .Cells(4, 8).Value = ""
        GoTo finish
    End If
End With

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

With Sheet1

    lastrow = .Cells(.Rows.Count, 6).End(xlUp).Row
    lastcol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    datafrom = .Range(.Cells(1, 5), .Cells(20, 5)).Find("Job Title", LookIn:=xlValues).Row + 1
    expirycriterion = .Range(.Cells(1, 1), .Cells(5, lastrow)).Find("Once", LookIn:=xlValues).Row
    ReDim expiredarr(1 To lastrow * lastcol + 1, 1 To 4) As Variant
    
    j = 1
    'array row counter for below loop
    
    For y = 6 To lastcol
        
        If .Cells(expirycriterion, y).Value <> "Once" Then
        
            For x = datafrom To lastrow
                
                If IsDate(.Cells(x, y).Value) Then
                
                    expiryperiod = .Cells(expirycriterion, y).Value
    
                    Select Case expiryperiod
                    'check what is the validity period
                    Case "1 Y"
                    expirydate = DateAdd("yyyy", 1, .Cells(x, y))
                    Case "2 Y"
                    expirydate = DateAdd("yyyy", 2, .Cells(x, y))
                    Case "3 Y"
                    expirydate = DateAdd("yyyy", 3, .Cells(x, y))
                    Case "4 Y"
                    expirydate = DateAdd("yyyy", 4, .Cells(x, y))
                    Case "5 Y"
                    expirydate = DateAdd("yyyy", 5, .Cells(x, y))
                    Case Else
                    MsgBox ("Expiry criterion not mathcing existing profiles of: 1 Y, 2 Y, 3 Y, 4 Y, 5 Y.")
                    End Select
                
                    If expirydate <= comparisondate Then
                    'if training is expired add the data to the array

                    expiredarr(j, 1) = .Cells(x, 1).Value
                    expiredarr(j, 2) = .Cells(x, 5).Value
                    expiredarr(j, 3) = .Cells(datafrom - 1, y).Value
                    expiredarr(j, 4) = DateValue(Month(expirydate) & "/" & Day(expirydate) & "/" & Year(expirydate))
                    'above = converting the date from UK format to US format
                    
                    j = j + 1
                
                    End If
                
                End If
    
            Next x
        
        End If

    Next y

End With

k = 1
'array row counter for below loop

With Sheet5

    'clear the list and formatting
    
    .Range(.Cells(5, 2), .Cells(lastrow * lastcol, 5)).ClearContents
    .Cells(1, 1).Select

    For z = 5 To UBound(expiredarr, 1)
    
        If expiredarr(k, 1) = Empty Then
        
            GoTo skip
            
        Else
            On Error GoTo nexxt
            .Cells(z, 2).Value = expiredarr(k, 1)
            .Cells(z, 3).Value = expiredarr(k, 2)
            .Cells(z, 4).Value = expiredarr(k, 3)
            .Cells(z, 5).Value = CStr(expiredarr(k, 4))
            
            'applying report formating per row of data
            .Rows(z).RowHeight = 15
            .Range(.Cells(z, 2), .Cells(z, 5)).HorizontalAlignment = xlCenter
            .Range(.Cells(z, 2), .Cells(z, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(z, 2), .Cells(z, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(z, 2), .Cells(z, 5)).Borders(xlEdgeBottom).Weight = xlThin
nexxt:
            k = k + 1
        
        End If
    
    Next z
    
skip:

titlestr = "TRAINING REQUIREMENTS FOR PERIOD: " & CStr(Cells(4, 8).Value) & " - " & CStr(comparisondate)
Cells(1, 1).Value = titlestr

End With

lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'sort results by date of expiry
    Range("B5:E" & lastrow).Select
    Selection.Sort Key1:=Range("E5"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    Range("A1").Select
    
finish:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub sort_by_name()

Dim lastrow As Long, lastcol As Long, freq As Long, datafrom As Long

freq = Range(Cells(1, 1), Cells(5, 20)).Find("Once", LookIn:=xlValues).Row
datafrom = Range(Cells(1, 1), Cells(5, 20)).Find("Full name", LookIn:=xlValues).Row
lastrow = Application.WorksheetFunction.Match("X", Range(Cells(1, 1), Cells(1000, 1)), 0) - 1
lastcol = Cells(freq, Columns.Count).End(xlToLeft).Column

Application.ScreenUpdating = False

    Range(Cells(datafrom, 1), Cells(lastrow, lastcol)).Select
    Selection.Sort Key1:=Range("A" & datafrom + 1), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("A1").Select
    
Application.ScreenUpdating = True
    
End Sub

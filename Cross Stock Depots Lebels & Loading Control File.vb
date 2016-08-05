Option Explicit

Sub main_procedure()

Dim customers() As Variant
Dim counter As Long, arr_counter As Long, lastrow As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'1. Scoop all customer data from a ZL18 download (data is sorted by route)
With Sheet1
    ' Establish doc length
    lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
    ' Dimention the customer array to rows = lastrow, columns = 5
    ' Columns: ROUTE / DROP / CUSTOMER / FROZEN VOL / CHILL & AMB VOL
    ReDim customers(1 To lastrow, 1 To 5) As Variant
    arr_counter = 1
    'Iterate through data, report starts from row 8
    For counter = 8 To lastrow
        If Left(.Cells(counter, 1).Value, 2) = "RG" Then
            ' Route number
            customers(arr_counter, 1) = .Cells(counter, 1).Value
            ' Drop number
            customers(arr_counter, 2) = .Cells(counter, 3).Value
            ' Customer name
            customers(arr_counter, 3) = .Cells(counter, 10).Value
            ' Frozen volume
            customers(arr_counter, 4) = .Cells(counter, 17).Value
            ' Chilled and ambient volume
            customers(arr_counter, 5) = .Cells(counter, 20).Value
            arr_counter = arr_counter + 1
        End If
    Next counter
End With

'2. Write data into tab "ALL"
With Sheet2
    ' Report starts from row 3
    counter = 3
    ' Clear report
    .Range(.Cells(3, 1), .Cells(501, 2)).ClearContents
    .Range(.Cells(3, 4), .Cells(501, 6)).ClearContents
    ' Write data from array to report
    For arr_counter = 1 To UBound(customers, 1)
        If customers(arr_counter, 1) = Empty Then Exit For
        ' Route
        .Cells(counter, 1).Value = Trim(customers(arr_counter, 1))
        ' Drop
        .Cells(counter, 2).Value = customers(arr_counter, 2)
        ' Skip 1 column for a formula finding the account number
        ' Then Customer name
        .Cells(counter, 4).Value = Trim(customers(arr_counter, 3))
        ' Frozen volume
        .Cells(counter, 5).Value = customers(arr_counter, 4)
        ' Chilled & ambient volume
        .Cells(counter, 6).Value = customers(arr_counter, 5)
        counter = counter + 1
    Next arr_counter
End With

' Garbage clearout
Erase customers

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub

Sub generate_labels()

Dim f_labels() As Variant, ca_labels() As Variant
Dim counter As Long, arr_counter As Long, lastrow As Long, f_label_count As Integer, f_label_counter As Integer, ca_label_count As Integer, ca_label_counter As Integer, x As Long, y As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' 1. Generate a pallet label entry per a calculated pallet, frozen and chilled & ambient separately
With Sheet2
    ' Establish doc length
    lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
    ' Set max dimentions of label counts equal to total customers * 10
    ' 4 columns: ROUTE / DROP / ACCOUNT / CUSTOMER
    ReDim f_labels(1 To lastrow * 10, 1 To 4) As Variant
    ReDim ca_labels(1 To lastrow * 10, 1 To 4) As Variant
    f_label_counter = 1
    ca_label_counter = 1
    ' Report starts from row 3
    For counter = 3 To lastrow
        ' Generate a label for each frozen calculated pallet
        If .Cells(counter, 7).Value > 0 Then
            f_label_count = .Cells(counter, 7).Value
            For x = 1 To f_label_count
                ' ROUTE
                f_labels(f_label_counter, 1) = .Cells(counter, 1).Value
                ' DROP
                f_labels(f_label_counter, 2) = .Cells(counter, 2).Value
                ' ACCOUNT
                f_labels(f_label_counter, 3) = .Cells(counter, 3).Value
                ' CUSTOMER NAME
                f_labels(f_label_counter, 4) = .Cells(counter, 4).Value
                ' NEXT LABEL NUMBER
                f_label_counter = f_label_counter + 1
            Next x
        End If
        ' Generate a label for each chilled & ambient calculated pallet
        If .Cells(counter, 8).Value > 0 Then
            ca_label_count = .Cells(counter, 8).Value
            For x = 1 To ca_label_count
                ' ROUTE
                ca_labels(ca_label_counter, 1) = .Cells(counter, 1).Value
                ' DROP
                ca_labels(ca_label_counter, 2) = .Cells(counter, 2).Value
                ' ACCOUNT
                ca_labels(ca_label_counter, 3) = .Cells(counter, 3).Value
                ' CUSTOMER NAME
                ca_labels(ca_label_counter, 4) = .Cells(counter, 4).Value
                ' NEXT LABEL NUMBER
                ca_label_counter = ca_label_counter + 1
            Next x
        End If
    Next counter
End With
    
' 2. Write Frozen labels to "FRZ Labels List" tab
With Sheet4
    ' Clear area for new data
    .Range(.Cells(1, 2), .Cells(1000, 5)).ClearContents
    .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
    For f_label_counter = 1 To UBound(f_labels, 1)
        ' When on empty array entry exit loop
        If f_labels(f_label_counter, 1) = Empty Then Exit For
        .Cells(f_label_counter, 2).Value = f_labels(f_label_counter, 1)
        .Cells(f_label_counter, 3).Value = f_labels(f_label_counter, 2)
        .Cells(f_label_counter, 4).Value = f_labels(f_label_counter, 3)
        .Cells(f_label_counter, 5).Value = f_labels(f_label_counter, 4)
    Next f_label_counter
End With

' 3. Write Chill & ambient labels to "CH&AMB Label List" tab
With Sheet8
    ' Clear area for new data
    .Range(.Cells(1, 2), .Cells(1000, 5)).ClearContents
    .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
    For ca_label_counter = 1 To UBound(ca_labels, 1)
        ' When on empty array entry exit loop
        If ca_labels(ca_label_counter, 1) = Empty Then Exit For
        .Cells(ca_label_counter, 2).Value = ca_labels(ca_label_counter, 1)
        .Cells(ca_label_counter, 3).Value = ca_labels(ca_label_counter, 2)
        .Cells(ca_label_counter, 4).Value = ca_labels(ca_label_counter, 3)
        .Cells(ca_label_counter, 5).Value = ca_labels(ca_label_counter, 4)
    Next ca_label_counter
End With

' Garbage clearout
Erase ca_labels
Erase f_labels

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
End Sub

Sub test_print_area()

Worksheets("Sheet1").PageSetup.PrintArea = "$A$1:$C$5"

End Sub

Sub selected_route_range_xdock(ByVal depot As String, ByVal temp As String)

Dim counter As Long, lastrow As Long, arr_counter As Long
Dim data() As Variant

arr_counter = 1

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

If temp = "FRZ" Then
    With Sheet4
        ' Establish row number of the last entry
        lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
        ' Set max dimentions for data array lastrow x 4
        ReDim data(1 To lastrow, 1 To 4) As Variant
        'Collect all customer data matching passed in depot
        For counter = 1 To lastrow
            If Left(.Cells(counter, 2).Value, 2) = depot Then
                data(arr_counter, 1) = .Cells(counter, 2).Value
                data(arr_counter, 2) = .Cells(counter, 3).Value
                data(arr_counter, 3) = .Cells(counter, 4).Value
                data(arr_counter, 4) = .Cells(counter, 5).Value
                arr_counter = arr_counter + 1
            End If
        Next counter
        ' Clear area for new data
        .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
        .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
        ' Print all data into second table
        For counter = 1 To UBound(data, 1)
            If data(counter, 1) = Empty Then Exit For
            .Cells(counter, 7).Value = data(counter, 1)
            .Cells(counter, 8).Value = data(counter, 2)
            .Cells(counter, 9).Value = data(counter, 3)
            .Cells(counter, 10).Value = data(counter, 4)
        Next counter
    End With
ElseIf temp = "CA" Then
    With Sheet8
        ' Establish row number of the last entry
        lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
        ' Set max dimentions for data array lastrow x 4
        ReDim data(1 To lastrow, 1 To 4) As Variant
        'Collect all customer data matching passed in depot
        For counter = 1 To lastrow
            If Left(.Cells(counter, 2).Value, 2) = depot Then
                data(arr_counter, 1) = .Cells(counter, 2).Value
                data(arr_counter, 2) = .Cells(counter, 3).Value
                data(arr_counter, 3) = .Cells(counter, 4).Value
                data(arr_counter, 4) = .Cells(counter, 5).Value
                arr_counter = arr_counter + 1
            End If
        Next counter
        ' Clear area for new data
        .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
        .Range(.Cells(1, 7), .Cells(1000, 10)).ClearContents
        ' Print all data into second table
        For counter = 1 To UBound(data, 1)
            If data(counter, 1) = Empty Then Exit For
            .Cells(counter, 7).Value = data(counter, 1)
            .Cells(counter, 8).Value = data(counter, 2)
            .Cells(counter, 9).Value = data(counter, 3)
            .Cells(counter, 10).Value = data(counter, 4)
        Next counter
    End With
Else
    ' If wrong temp selected
    MsgBox "Wrong temperature selected."
    Exit Sub
End If


Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

' Grabage clearout
Erase data

End Sub

Sub AL_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "AL"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub AL_CA()

Dim depot As String, chamber As String, route_range As String

depot = "AL"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub TW_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "TW"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub TW_CA()

Dim depot As String, chamber As String, route_range As String

depot = "TW"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub WG_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "WG"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub WG_CA()

Dim depot As String, chamber As String, route_range As String

depot = "WG"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub YT_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "YT"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub YT_CA()

Dim depot As String, chamber As String, route_range As String

depot = "YT"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub GT_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "GT"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub GT_CA()

Dim depot As String, chamber As String, route_range As String

depot = "GT"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub PL_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "PL"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub PL_CA()

Dim depot As String, chamber As String, route_range As String

depot = "PL"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub TF_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "TF"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub TF_CA()

Dim depot As String, chamber As String, route_range As String

depot = "TF"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub NE_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "NE"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub NE_CA()

Dim depot As String, chamber As String, route_range As String

depot = "NE"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub DH_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "DH"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub DH_CA()

Dim depot As String, chamber As String, route_range As String

depot = "DH"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub RG_FRZ()

Dim depot As String, chamber As String, route_range As String

depot = "RG"
chamber = "FRZ"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub RG_CA()

Dim depot As String, chamber As String, route_range As String

depot = "RG"
chamber = "CA"
route_range = "SELECTED"

Call selected_route_range_xdock(depot, chamber)
Call print_labels(route_range, chamber)

End Sub

Sub ALL_ROUTES_FRZ()

Dim chamber As String, route_range As String

chamber = "FRZ"
route_range = "ALL"

Call print_labels(route_range, chamber)

End Sub

Sub ALL_ROUTES_CA()

Dim chamber As String, route_range As String

chamber = "CA"
route_range = "ALL"

Call print_labels(route_range, chamber)

End Sub

Sub print_labels(ByVal selection As String, temp As String)

Dim last_printable_row As Long, lastrow As Long

' Each label is 64 rows long
last_printable_row = 64

If selection = "ALL" Then
    If temp = "FRZ" Then
        
        With Sheet4
            ' Find the number of labels
            lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
            ' Calculate last printable row on labels tab for print area setting
            last_printable_row = last_printable_row * lastrow
        End With
        
        ' Print all frozen labels
        With Sheet5
            ' Set print area
            .PageSetup.PrintArea = "$A$1:$J$" & last_printable_row
            ' Print sheet
            .PrintOut
        End With
        
    ElseIf temp = "CA" Then
        With Sheet8
            ' Find the number of labels
            lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
            ' Calculate last printable row on labels tab for print area setting
            last_printable_row = last_printable_row * lastrow
        End With
        
        ' Print all chilled & ambient labels
        With Sheet7
            ' Set print area
            .PageSetup.PrintArea = "$A$1:$J$" & last_printable_row
            ' Print sheet
            .PrintOut
        End With
        
    Else
        MsgBox "Wrong temperature selected."
        Exit Sub
    End If
    
ElseIf selection = "SELECTED" Then
    If temp = "FRZ" Then
    
        With Sheet4
            ' Find the number of labels
            lastrow = .Cells(Rows.Count, 7).End(xlUp).Row
            ' Calculate last printable row on labels tab for print area setting
            last_printable_row = last_printable_row * lastrow
        End With
        
        With Sheet6
            ' Set print area
            .PageSetup.PrintArea = "$A$1:$J$" & last_printable_row
            ' Print sheet
            .PrintOut
        End With
    
    ElseIf temp = "CA" Then
    
        With Sheet8
            ' Find the number of labels
            lastrow = .Cells(Rows.Count, 7).End(xlUp).Row
            ' Calculate last printable row on labels tab for print area setting
            last_printable_row = last_printable_row * lastrow
        End With
        
        With Sheet10
            ' Set print area
            .PageSetup.PrintArea = "$A$1:$J$" & last_printable_row
            ' Print sheet
            .PrintOut
        End With
        
    Else
        MsgBox "Wrong temperature selected."
        Exit Sub
    End If
Else
    MsgBox "Wrong print selection."
    Exit Sub
End If

End Sub

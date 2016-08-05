Option Explicit

Sub main_procedure()

Dim customers() As Variant
Dim counter As Long, arr_counter As Long, lastrow As Long

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

End Sub
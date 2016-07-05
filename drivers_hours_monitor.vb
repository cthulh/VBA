Function Total_Hours(hours_array As Range)

 Dim x As Long, y As Long, z As Long
 Dim cell As Range
 Dim hours As Integer, minutes As Integer
 Dim minutes_format As String
 Dim tempval As String
 x = 1
 
 For Each cell In hours_array
  tempval = cell.Value
  If Len(tempval) = 4 Then
   tempval = "0" & tempval
  End If
  
  hours = hours + CInt(Left(tempval, 2))
  minutes = minutes + CInt(Right(tempval, 2))

  x = x + 1
  
 Next cell

 hours = hours + minutes \ 60
 minutes = minutes Mod 60
 If Len(CStr(minutes)) < 2 Then
    minutes_format = "0" & minutes
 Else
    minutes_format = CStr(minutes)
 End If
 
 Total_Hours = (hours & ":" & minutes_format)

End Function

Function Average_Hours(hours_array As Range)

 Dim x As Long, y As Long, z As Long
 Dim cell As Range
 Dim hours As Long, minutes As Long
 Dim minutes_format As String
 Dim tempval As String
 x = 1
 
 For Each cell In hours_array
  tempval = cell.Value
  If Len(tempval) = 4 Then
   tempval = "0" & tempval
  End If
  
  hours = hours + CInt(Left(tempval, 2))
  minutes = minutes + CInt(Right(tempval, 2))

  x = x + 1
  
 Next cell

 minutes = ((hours * 60) + minutes) / 17
 hours = minutes \ 60
 minutes = minutes Mod 60
 
 If Len(CStr(minutes)) = 1 Then
  minutes_format = "0" & minutes
 Else
  minutes_format = CStr(minutes)
 End If
 
 Average_Hours = hours & ": " & minutes_format

End Function

Sub test_Total_Hours()

 Dim x As Long, y As Long, z As Long
 Dim cell As Range, hours_array As Range
 Set hours_array = Range("C3:S3")
 Dim hours As Integer, minutes As Integer
 Dim minutes_format As String
 Dim tempval As String
 x = 1
 
 For Each cell In hours_array
  tempval = cell.Value
  If Len(tempval) = 4 Then
   tempval = "0" & tempval
  End If
  
  hours = hours + CInt(Left(tempval, 2))
  minutes = minutes + CInt(Right(tempval, 2))

  x = x + 1
  
 Next cell

 hours = hours + minutes \ 60
 minutes = minutes Mod 60
 If Len(CStr(minutes)) < 2 Then
    minutes_format = "0" & minutes
 Else
    minutes_format = CStr(minutes)
 End If
 
 MsgBox (hours & ":" & minutes_format)

End Sub

Sub test_Average_Hours()

 Dim x As Long, y As Long, z As Long
 Dim cell As Range, hours_array As Range
 Set hours_array = Range("C3:S3")
 Dim hours As Long, minutes As Long
 Dim minutes_format As String
 Dim tempval As String
 x = 1
 
 For Each cell In hours_array
  tempval = cell.Value
  If Len(tempval) = 4 Then
   tempval = "0" & tempval
  End If
  
  hours = hours + CInt(Left(tempval, 2))
  minutes = minutes + CInt(Right(tempval, 2))

  x = x + 1
  
 Next cell

 minutes = ((hours * 60) + minutes) / 17
 hours = minutes \ 60
 minutes = minutes Mod 60
 
 If Len(CStr(minutes)) = 1 Then
  minutes_format = "0" & minutes
 Else
  minutes_format = CStr(minutes)
 End If
 
 MsgBox (hours & ":" & minutes_format)

End Sub

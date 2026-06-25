Function met2utc(met As Double) As String
    ' Ensures the formula recalculates when sheet data changes
    Application.Volatile 

    Dim ws As Worksheet
    Dim xRange As Range, yRange As Range, configRange As Range
    Dim idx As Variant
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim interpolatedY As Double
    Dim offsetVal As Double
    Dim scaleFactor As Double
    Dim finalYDayRaw As Double
    Dim finalYDate As Date
    Dim doy As Long
    Dim totalSeconds As Double
    Dim fracSeconds As Double
    Dim fracString As String

    ' 1. Set worksheet
    Set ws = ThisWorkbook.Sheets("sclk_data")
    
    ' 2. Dynamically set ranges based on text definitions in A2, B2, and C2
    On Error GoTo RangeError
    Set xRange = ws.Range(ws.Range("A2").Text)
    Set yRange = ws.Range(ws.Range("B2").Text)
    Set configRange = ws.Range(ws.Range("C2").Text)
    On Error GoTo 0

    ' 3. Retrieve Date/Time Offset and Scale Factor from the evaluated config range
    offsetVal = CDbl(configRange.Cells(1).Value) 
    scaleFactor = configRange.Cells(2).Value

    ' Check for division by zero on the scale factor
    If scaleFactor = 0 Then
        met2utc = "Error: Scale Factor is 0"
        Exit Function
    End If

    ' 4. Find the lower bound index using Excel's MATCH equivalent
    On Error Resume Next
    idx = Application.Match(met, xRange, 1)
    On Error GoTo 0

    ' Handle errors or out-of-bounds cases
    If IsError(idx) Then
        met2utc = "Error: Below Range"
        Exit Function
    ElseIf idx >= xRange.Cells.Count Then
        met2utc = "Error: Above Range"
        Exit Function
    End If

    ' 5. Extract bounding points
    x1 = xRange.Cells(idx).Value
    x2 = xRange.Cells(idx + 1).Value
    y1 = yRange.Cells(idx).Value
    y2 = yRange.Cells(idx + 1).Value

    ' Check for division by zero
    If x2 = x1 Then
        met2utc = "Error: Div by 0"
        Exit Function
    End If

    ' 6. Perform Piecewise Linear Interpolation
    interpolatedY = y1 + ((met - x1) / (x2 - x1)) * (y2 - y1)

    ' 7. DIVIDE by Scale Factor to convert seconds BEFORE adding the Day Offset
    finalYDayRaw = (interpolatedY / scaleFactor) + offsetVal
    finalYDate = CDate(finalYDayRaw)

    ' 8. Calculate Day of Year (DOY)
    doy = Int(finalYDayRaw) - DateSerial(Year(finalYDate), 1, 1) + 1

    ' 9. Isolate fractional seconds to 6 decimal places
    totalSeconds = (finalYDayRaw - Int(finalYDayRaw)) * 86400#
    fracSeconds = totalSeconds - Int(totalSeconds)
    fracString = Format(fracSeconds, ".000000")

    ' 10. Construct the final precise string
    met2utc = Format(finalYDate, "yy-") & _
              Format(doy, "000") & _
              Format(finalYDate, "/hh:mm:ss") & _
              fracString
              
    Exit Function

RangeError:
    met2utc = "Error: Invalid Range Definition"
End Function

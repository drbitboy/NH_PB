REM  *****  BASIC  *****

Function met2utc(met As Double) As String
    Dim oDoc As Object
    Dim oSheet As Object
    Dim xRange As Object, yRange As Object, configRange As Object
    Dim oFunc As Object
    Dim idx As Variant
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim interpolatedY As Double
    Dim offsetVal As Double
    Dim scaleFactor As Double
    Dim finalYDayRaw As Double
    Dim finalYDate As Double
    Dim doy As Long
    Dim totalSeconds As Double
    Dim fracSeconds As Double
    Dim fracString As String
    Dim sYear As String, sDoy As String, sTime As String
    Dim oOffsetCell As Object, oScaleCell As Object
    
    ' Universal error handler setup for LibreOffice Basic
    On Local Error GoTo RangeError

    ' 1. Access the document and the specific sheet
    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByName("sclk_data")
    
    ' 2. Dynamically set ranges based on the text definitions in A2, B2, and C2
    xRange = oSheet.getCellRangeByName(oSheet.getCellRangeByName("A2").String)
    yRange = oSheet.getCellRangeByName(oSheet.getCellRangeByName("B2").String)
    configRange = oSheet.getCellRangeByName(oSheet.getCellRangeByName("C2").String)
    On Local Error GoTo 0 ' Clear error handler for normal logic

    ' 3. Initialize Calc functions service
    oFunc = CreateUnoService("com.sun.star.sheet.FunctionAccess")

    ' Retrieve Offset and Scale Factor with type safety fallbacks
    oOffsetCell = configRange.getCellByPosition(0, 0)
    oScaleCell = configRange.getCellByPosition(0, 1)

    If oOffsetCell.Value = 0 And oOffsetCell.String <> "" Then
        offsetVal = CDbl(oFunc.callFunction("VALUE", Array(oOffsetCell.String)))
    Else
        offsetVal = oOffsetCell.Value
    End If
    
    scaleFactor = oScaleCell.Value
    
    ' Check for division by zero on the scale factor
    If scaleFactor = 0 Then
        met2utc = "Error: Scale Factor is 0"
        Exit Function
    End If
    
    ' 4. Use Calc's MATCH function
    On Local Error Resume Next
    idx = oFunc.callFunction("MATCH", Array(met, xRange, 1))
    On Local Error GoTo 0
    
    ' Handle errors or out-of-bounds cases
    If IsEmpty(idx) Or IsNull(idx) Or idx < 1 Then
        met2utc = "Error: Below Range"
        Exit Function
    ElseIf idx >= xRange.Rows.Count Then
        met2utc = "Error: Above Range"
        Exit Function
    End If
    
    ' 5. Extract bounding points (Rows are 0-indexed in API)
    x1 = xRange.getCellByPosition(0, idx - 1).Value
    x2 = xRange.getCellByPosition(0, idx).Value
    y1 = yRange.getCellByPosition(0, idx - 1).Value
    y2 = yRange.getCellByPosition(0, idx).Value
    
    ' Check for division by zero
    If x2 = x1 Then
        met2utc = "Error: Div by 0"
        Exit Function
    End If
    
    ' 6. Perform Piecewise Linear Interpolation
    interpolatedY = y1 + ((met - x1) / (x2 - x1)) * (y2 - y1)
    
    ' 7. DIVIDE by Scale Factor to convert seconds BEFORE adding the Day Offset
    finalYDayRaw = (interpolatedY / scaleFactor) + offsetVal
    finalYDate = Int(finalYDayRaw)
    
    ' 8. Calculate Day of Year (DOY) using Calc function access
    sYear = oFunc.callFunction("TEXT", Array(finalYDate, "yyyy"))
    Dim firstDayOfYear As Double
    firstDayOfYear = oFunc.callFunction("DATE", Array(CInt(oFunc.callFunction("YEAR", Array(finalYDate))), 1, 1))
    doy = Clng(finalYDate - firstDayOfYear + 1)
    
    ' 9. Isolate fractional seconds to 6 decimal places
    totalSeconds = (finalYDayRaw - Int(finalYDayRaw)) * 86400.0
    fracSeconds = totalSeconds - Int(totalSeconds)
    
    ' Format components
    fracString = oFunc.callFunction("TEXT", Array(fracSeconds, ".000000"))
    sDoy = oFunc.callFunction("TEXT", Array(doy, "000"))
    sTime = oFunc.callFunction("TEXT", Array(finalYDayRaw, "hh:mm:ss"))
    
    ' 10. Construct final string
    met2utc = sYear & "-" & sDoy & "/" & sTime & fracString
    Exit Function

RangeError:
    met2utc = "Error: Invalid Range Definition"
End Function

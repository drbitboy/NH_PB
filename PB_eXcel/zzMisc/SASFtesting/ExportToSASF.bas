Attribute VB_Name = "ExportToSASF"
Sub ExportToSASF()
    Dim ws As Worksheet, r As Long, currentYear As String, trackID As String, lastTrackID As String
    Dim colO_Val As String, startText As String, timePart As String, startMET As String, endMET As String
    Dim calculatedBytes As String, commentStr As String, keyType As String, booleanFields(1 To 15) As String
    Dim i As Integer, activityIndex As Integer, selectedArea As Range, startRow As Long, endRow As Long
    Dim anyEtoI As Boolean, anyJorK As Boolean, cellVal As String, colLetter As String, userChoice As VbMsgBoxResult
    Dim fileNum As Integer, filePath As String, targetCols As Variant, associatedIndex As Variant, boolStr As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of data rows on the worksheet before running this macro.", vbExclamation, "No Rows Selected"
        Exit Sub
    End If
    
    filePath = Application.GetSaveAsFilename(InitialFileName:="output.sasf", FileFilter:="SASF Files (*.sasf), *.sasf")
    If filePath = "False" Then Exit Sub
    
    Set ws = ActiveSheet
    currentYear = Format(Date, "yyyy")
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    lastTrackID = ""
    activityIndex = 1
    targetCols = Array("E", "K", "F", "G", "H", "I", "J")
    associatedIndex = Array(8, 9, 10, 11, 12, 13, 14)
    
    For Each selectedArea In Selection.Areas
        startRow = selectedArea.Row
        endRow = selectedArea.Row + selectedArea.Rows.Count - 1
        For r = startRow To endRow
            If r < 2 Then
                GoTo NextRow
            End If
            
            anyEtoI = (Application.CountA(ws.Range("E" & r & ":I" & r)) > 0)
            anyJorK = (Application.CountA(ws.Range("J" & r & ":K" & r)) > 0)
            If Not (anyEtoI Xor anyJorK) Then
                GoTo NextRow
            End If
            
            colO_Val = Trim(ws.Cells(r, "O").Value)
            trackID = Trim(ws.Cells(r, "P").Value)
            
            If colO_Val = "" And trackID = "" Then
                GoTo NextRow
            End If
            If trackID = "" And colO_Val <> "" Then
                trackID = lastTrackID
            End If
            If InStr(trackID, "_") = 0 Then
                GoTo NextRow
            End If
            
            ' --- REFACTORED INLINE ERROR HANDLING FOR MATHEMATICAL ROUNDING AND CLng ---
            On Error GoTo SkipDueToCLngError
            startMET = CStr(CLng(Application.Round(ws.Cells(r, "Q").Value, 0)))
            endMET = CStr(CLng(Application.Round(ws.Cells(r, "R").Value, 0)))
            calculatedBytes = CStr(CLng(Application.Round(ws.Cells(r, "V").Value, 0)))
            On Error GoTo 0 ' Resets standard error checking immediately
            ' ---------------------------------------------------------------------------
            
            If startMET = "" Or endMET = "" Or calculatedBytes = "" Then
                GoTo NextRow
            End If
            If (startMET Like "*[!0-9]*") Or (endMET Like "*[!0-9]*") Or (calculatedBytes Like "*[!0-9]*") Then
                GoTo NextRow
            End If
            
            For i = 1 To 15
                booleanFields(i) = "FALSE"
            Next i
            
            If LCase(Trim(ws.Cells(r, "E").Value)) = "hsk3" Then booleanFields(8) = "TRUE"
            If LCase(Trim(ws.Cells(r, "K").Value)) = "hsk4" Then booleanFields(9) = "TRUE"
            If LCase(Trim(ws.Cells(r, "F").Value)) = "dt1" Then booleanFields(10) = "TRUE"
            If LCase(Trim(ws.Cells(r, "G").Value)) = "dt2" Then booleanFields(11) = "TRUE"
            If LCase(Trim(ws.Cells(r, "H").Value)) = "dt3" Then booleanFields(12) = "TRUE"
            If LCase(Trim(ws.Cells(r, "I").Value)) = "dt4" Then booleanFields(13) = "TRUE"
            If LCase(Trim(ws.Cells(r, "J").Value)) = "sdt1" Then booleanFields(14) = "TRUE"
            
            For i = 0 To 6
                colLetter = targetCols(i)
                cellVal = Trim(ws.Cells(r, colLetter).Value)
                If booleanFields(associatedIndex(i)) = "FALSE" And cellVal <> "" Then
                    userChoice = MsgBox("Row: " & r & vbCrLf & "Column: " & colLetter & vbCrLf & "Value: """ & cellVal & """" & vbCrLf & vbCrLf & _
                                        "This cell has a bad value and the row will be excluded from the SASF." & vbCrLf & vbCrLf & _
                                        "Click [OK] to skip this row and continue processing." & vbCrLf & "Click [Cancel] to abort.", vbOKCancel + vbCritical, "Bad Value")
                    If userChoice = vbOK Then
                        GoTo NextRow
                    End If
                    If userChoice = vbCancel Then
                        Close #fileNum
                        MsgBox "SASF compilation aborted by user.", vbExclamation
                        Exit Sub
                    End If
                End If
            Next i
            
            If InStr(LCase(trackID), "pepssi") > 0 Then
                keyType = "pepssi"
            Else
                keyType = "swap"
            End If
            
            If InStr(colO_Val, "/") > 0 Then
                timePart = Mid(colO_Val, InStr(colO_Val, "/") + 1)
                startText = currentYear & "-" & Left(colO_Val, InStr(colO_Val, "/") - 1) & "T" & timePart & ":00,"
            Else
                startText = currentYear & "-000T00:00:00,"
            End If
            
            commentStr = ",\SWAP " & Trim(ws.Cells(r, "N").Value) & " " & Trim(ws.Cells(r, "L").Value) & " to " & Trim(ws.Cells(r, "M").Value) & " \"
            
            If trackID <> lastTrackID Then
                If lastTrackID <> "" Then
                    Print #fileNum, "),"
                    Print #fileNum, "end;"
                    Print #fileNum, ""
                End If
                Print #fileNum, "request(PBACK_TRK_" & trackID & "_" & keyType & ","
                Print #fileNum, "    START_TIME, " & startText
                Print #fileNum, "    REQUESTOR, ""aph"","
                Print #fileNum, "    PROCESSOR, ""IEM1"","
                Print #fileNum, "    KEY, ""SSR"")"
                Print #fileNum, ""
                activityIndex = 1
            Else
                Print #fileNum, "    ),"
            End If
            
            Print #fileNum, "    activity(" & activityIndex & ","
            If activityIndex = 1 Then
                Print #fileNum, "        SCHEDULED_TIME, \+0:01:05\,FROM_PREVIOUS_START,"
            Else
                Print #fileNum, "        SCHEDULED_TIME, \+0:00:05\,FROM_PREVIOUS_START,"
            End If
            Print #fileNum, "        COMMENT" & commentStr
            
            boolStr = Join(booleanFields, ",")
            Print #fileNum, "        SSR(CAS_SSR_PLAYBACK_LOWSPEED," & startMET & "," & endMET & "," & boolStr & "," & calculatedBytes & ")"
            lastTrackID = trackID
            activityIndex = activityIndex + 1
            
            GoTo NextRow
            
SkipDueToCLngError:
            Err.Clear
            On Error GoTo 0
            GoTo NextRow
            
NextRow:
        Next r
    Next selectedArea
    
    If lastTrackID <> "" Then
        Print #fileNum, "    )"
        Print #fileNum, "end;"
    End If
    Close #fileNum
    MsgBox "SASF file creation completed with updated boolean parameters!", vbInformation
End Sub



Attribute VB_Name = "ExportToSASF"
Sub ExportToSASF()
    Dim ws As Worksheet
    Dim r As Long
    Dim currentYear As String
    Dim trackID As String
    Dim lastTrackID As String
    Const validTrackID As String = "[0-3][0-9][0-9]_[0-9A-F]"
    Dim startTime_Val As String
    Const validStartTime As String = "[0-3][0-9][0-9]/[0-2][0-9]:[0-5][0-9]"
    Dim startDOY As String
    Dim lastStartDOY As String
    Dim startText As String
    Dim timePart As String
    Dim startMET As String
    Dim endMET As String
    Dim calculatedBytes As String
    Dim commentStr As String
    Dim keyType As String
    Dim booleanFields(1 To 15) As String
    Dim i As Integer
    Dim activityIndex As Integer
    Dim selectedArea As Range
    Dim startRow As Long, endRow As Long
    Dim anyEtoI As Boolean
    Dim anyJorK As Boolean
    Dim cellVal As String
    Dim colLetter As String
    Dim userChoice As VbMsgBoxResult
    Dim fileNum As Integer
    Dim filePath As String
    Dim targetCols As Variant
    Dim associatedIndex As Variant
    Dim boolStr As String
    
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
    
    ' Write the header:  terminate lines with linefeed; semicolon prevents Print from appending vbCrLf
    Print #fileNum, SASFHeader() & vbLf;
    
    lastTrackID = ""
    activityIndex = 1
    targetCols = Array("E", "K", "F", "G", "H", "I", "J")
    associatedIndex = Array(8, 9, 10, 11, 12, 13, 14)
    
    ' ensure year does not change on first valid start time
    lastStartDOY = "000"
    
    For Each selectedArea In Selection.Areas
        startRow = selectedArea.Row
        endRow = selectedArea.Row + selectedArea.Rows.Count - 1
        For r = startRow To endRow
            If r < 2 Then
                GoTo NextRow
            End If
            
            ' PEPSSI XOR SWAP PB data types must be present for this row to generate an activity
            anyEtoI = (Application.CountA(ws.Range("E" & r & ":I" & r)) > 0)
            anyJorK = (Application.CountA(ws.Range("J" & r & ":K" & r)) > 0)
            If Not (anyEtoI Xor anyJorK) Then GoTo NextRow
            
            startTime_Val = Trim(ws.Cells(r, "O").Text)   ' Track/request time from .sci_dldv file, DOY/HH:MM
            trackID = Trim(ws.Cells(r, "P").Text)    ' Track ID from .sci_dldv file, DOY_#, or blank
            
            ' If both request START_TIME (Column P) and track ID are blank, skip to next row
            If startTime_Val = "" And trackID = "" Then GoTo NextRow
            
            ' Check START_TIME when non-blank track ID value is valid (DOY_#)
            If trackID Like validTrackID Then
            
                ' A valid row's start time must be of the form "DOY/HH:MM"
                If startTime_Val Like validStartTime Then
                
                    ' Increment year if start time's DOY decreases, indicating the load crossed a new-year boundary
                    startDOY = Left(startTime_Val, 3)
                    If startDOY < lastStartDOY Then
                        currentYear = Trim(CLng(currentYear) + 1)
                    End If
                    lastStartDOY = startDOY

                ' IF track ID is of the form DOY/HH:MM, AND start time is not of the form DOY_#, THEN skip this row
                Else
                    GoTo NextRow
                End If
    
            ' If track ID is blank and START_TIME (Column O) is non-blank, then use previous track ID if valid
            ElseIf trackID = "" And startTime_Val <> "" And lastTrackID Like validTrackID Then
                trackID = lastTrackID

            ' Otherwise, this is not a valid playback row
            Else
                GoTo NextRow
            End If
            
            startMET = ""
            endMET = ""
            calculatedBytes = ""
            On Error GoTo SkipDueToCLngError
            startMET = CStr(CLng(Application.Round(ws.Cells(r, "Q").Value, 0)))
            endMET = CStr(CLng(Application.Round(ws.Cells(r, "R").Value, 0)))
            calculatedBytes = CStr(CLng(Application.Round(ws.Cells(r, "V").Value, 0)))
            On Error GoTo 0 ' Resets standard error checking immediately
            ' ---------------------------------------------------------------------------
            
            If startMET = "" Or endMET = "" Or calculatedBytes = "" Then GoTo NextRow

            If (startMET Like "*[!0-9]*") Or (endMET Like "*[!0-9]*") Or (calculatedBytes Like "*[!0-9]*") Then GoTo NextRow
            
            ' Assume all PlayBack (PB) data types cell on this row are blank
            For i = 1 To 15
                booleanFields(i) = "FALSE"
            Next i
            
            ' Detect standard PB data types in columns E-K; enforce no whitespace; enforce case-sensitive
            If ws.Cells(r, "E").Text = "HSK3" Then booleanFields(8) = "TRUE"
            If ws.Cells(r, "K").Text = "HSK4" Then booleanFields(9) = "TRUE"
            If ws.Cells(r, "F").Text = "DT1" Then booleanFields(10) = "TRUE"
            If ws.Cells(r, "G").Text = "DT2" Then booleanFields(11) = "TRUE"
            If ws.Cells(r, "H").Text = "DT3" Then booleanFields(12) = "TRUE"
            If ws.Cells(r, "I").Text = "DT4" Then booleanFields(13) = "TRUE"
            If ws.Cells(r, "J").Text = "SDT1" Then booleanFields(14) = "TRUE"
            
            ' Ignore rows where PB data types are non-blank and non-standard
            For i = 0 To 6
                colLetter = targetCols(i)
                cellVal = CStr(ws.Cells(r, colLetter).Value)
                If booleanFields(associatedIndex(i)) = "FALSE" And cellVal <> "" Then
                    userChoice = MsgBox("Cell:  " & ws.Cells(r, colLetter).Address(False, False) & vbCrLf & "Value: """ & cellVal & """" _
                                        & vbCrLf & vbCrLf _
                                        & "This cell has a bad value and the row will be excluded from the SASF." & vbCrLf & vbCrLf _
                                        & "Click [OK] to skip this row and continue processing." & vbCrLf _
                                        & "Click [Cancel] to abort." _
                                        , vbOKCancel + vbCritical, "Bad Value")
                    
                    ' Skip row, continue with next row
                    If userChoice = vbOK Then GoTo NextRow
                    
                    ' Cancel SASF writing
                    If userChoice = vbCancel Then
                        Close #fileNum
                        MsgBox "SASF compilation aborted by user.", vbExclamation
                        Exit Sub
                    End If
                End If
            Next i
            
            ' Determine if request and/or activity is PEPSSI or SWAP
            If booleanFields(9) = "FALSE" And booleanFields(14) = "FALSE" Then
                keyType = "pepssi"
            Else
                keyType = "swap"
            End If
            
            ' Build request start time:  DOY/HH:MM becomes YYYY-DOYTHH:MM:00
            If InStr(startTime_Val, "/") > 0 Then
                timePart = Mid(startTime_Val, InStr(startTime_Val, "/") + 1)
                startText = currentYear & "-" & Left(startTime_Val, InStr(startTime_Val, "/") - 1) & "T" & timePart & ":00,"
            Else
                startText = currentYear & "-000T00:00:00,"
            End If
            
            
            ' New request for new track ID; to here, track ID will always be of form DOY_#
            If trackID <> lastTrackID Then
            
                ' Finish any previous activity and previous request
                If lastTrackID <> "" Then
                    Print #fileNum, vbTab & ")," & vbLf;
                    Print #fileNum, "end;" & vbLf;
                    Print #fileNum, "" & vbLf;
                End If
                
                ' New request
                Print #fileNum, "request(PBACK_TRK_" & trackID & "_" & keyType & "," & vbLf;
                Print #fileNum, vbTab & vbTab & "START_TIME, " & startText & vbLf;
                Print #fileNum, vbTab & vbTab & "REQUESTOR, ""aph""," & vbLf;
                Print #fileNum, vbTab & vbTab & "PROCESSOR, ""IEM1""," & vbLf;
                Print #fileNum, vbTab & vbTab & "KEY, ""SSR"")" & vbLf;
                Print #fileNum, "" & vbLf;
                
                ' this activity is the first of the request
                activityIndex = 1
            
            ' Finish the previous activity
            Else
                Print #fileNum, vbTab & ")," & vbLf;
            End If
            
            ' Output the activity
            ' - activity line and time
            Print #fileNum, vbTab & "activity(" & activityIndex & "," & vbLf;
            If activityIndex = 1 Then
                Print #fileNum, vbTab & vbTab & "SCHEDULED_TIME,\+0:00:05\,FROM_PREVIOUS_START," & vbLf;
            Else
                'Print #fileNum, vbTab & vbTab & "SCHEDULED_TIME,\+0:01:05\,FROM_PREVIOUS_START," & vbLf;
                Print #fileNum, vbTab & vbTab & "SCHEDULED_TIME,\+0:00:05\,FROM_PREVIOUS_START," & vbLf;
            End If
            
            ' - activity comment line
            commentStr = ",\" & UCase(keyType) & " " & Trim(ws.Cells(r, "N").Text) _
                    & " " & Trim(ws.Cells(r, "L").Text) & " to " & Trim(ws.Cells(r, "M").Text) _
                    & " \,"
            Print #fileNum, vbTab & vbTab & "COMMENT" & commentStr & vbLf;
            
            ' Playback line
            boolStr = Join(booleanFields, ",")
            Print #fileNum, vbTab & vbTab & "SSR(CAS_SSR_PLAYBACK_LOWSPEED," & startMET & "," & endMET & "," & boolStr & "," & calculatedBytes & ")" & vbLf;
            lastTrackID = trackID
            
            ' Increment activity number within request
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
        Print #fileNum, vbTab & ")," & vbLf;
        Print #fileNum, "end;" & vbLf;
    End If
    Print #fileNum, vbLf;
    Print #fileNum, "$$EOF" & vbLf;
    Close #fileNum
    MsgBox "SASF file creation completed with updated boolean parameters!", vbInformation
End Sub

' Boilerplate header for SASF
Function SASFHeader() As String
    Dim headerA() As Variant
    Dim headerB() As Variant
    Dim headerC() As Variant
    
    ' Join lines with vbLf line termination
    ' - VBA limits the number of continuations to 25:
    '   - there are 61 lines; break them into three groups of 21 or less
    ' - first dozen or so line end with CRLF;
    headerA = Array _
    ("CCSD3ZF0000100000001NJPL3KS0L015$$MARK$$;" & vbCr _
    , "MISSION_NAME = new_horizons;" & vbCr _
    , "SPACECRAFT_NAME = new_horizons;" & vbCr _
    , "DATA_SET_ID = SPACECRAFT_ACTIVITY_SEQUENCE;" & vbCr _
    , "FILE_NAME = 26222_pepssi_swap_pb_v1.sasf;" & vbCr _
    , "APPLICABLE_START_TIME = 2026-222T02:09:00.000;" & vbCr _
    , "APPLICABLE_STOP_TIME = 2026-243T05:10:00.000;" & vbCr _
    , "PRODUCT_CREATION_TIME = 2026-190T16:48:57;" & vbCr _
    , "PRODUCER_ID = whittke1;" & vbCr _
    , "SEQ_ID = 26222;" & vbCr _
    , "HOST_ID = bond;" & vbCr _
    , "CCSD3RE00000$$MARK$$NJPL3IF0M01300000001;" & vbCr _
    , "$$new_horiz SPACECRAFT ACTIVITY SEQUENCE FILE" _
    , "************************************************************" _
    , "*PROJECT    new_horizons" _
    , "*SPACECRAFT 98" _
    , "*OPERATOR   seqgen User" _
    , "*FILE_CMPLT TRUE" _
    , "*DATE       Thu Jul  9 16:48:57 2026" _
    , "*SEQ_GEN    V27.0 Wed Aug 6 08:35:42 PDT 2003" _
    , "*BEGIN      2026-222T02:09:00.000" _
    )

    ' Note tabs
    headerB = Array _
    ("*CUTOFF     2026-243T05:10:00.000" _
    , "*TITLE      26222" _
    , "*EPOCHS_DEF " _
    , "*EPOCH_2014_MU69_CA," & vbTab & "2019-001T05:33:48.000" _
    , "*EPOCH_MU69_CA," & vbTab & "2019-001T07:00:00.000" _
    , "*EPOCH_Pluto_CA," & vbTab & "2015-195T11:47:00.000" _
    , "*EPOCHS_END " _
    , "*Input files used:" _
    , "*File Type  Last modified             File name" _
    , "*CONTEXT    Tue Nov 12 17:30:03 2013  /disks/d13/project/new_horizons/mops/prod/seq/cvf/GC_PARAMS.cvf" _
    , "*CONTEXT    Wed Apr  2 15:41:39 2025  /disks/d13/project/new_horizons/mops/prod/seq/cvf/APIDS.cvf" _
    , "*CONTEXT    Tue Aug  6 16:10:15 2019  /disks/d13/project/new_horizons/mops/prod/seq/cvf/dsn_trx.cvf" _
    , "*CONTEXT    Fri Nov 15 17:14:08 2019  /disks/d13/project/new_horizons/mops/prod/seq/cvf/INT_RATIO.cvf" _
    , "*SC_MODEL   Thu Jun 11 14:26:36 2026  /disks/d13/project/new_horizons/mops/prod/seq/smf/new_horizons.smf" _
    , "*CATALOG    Thu May  7 11:09:46 2026  /disks/d13/project/new_horizons/mops/prod/seq/satf/new_horizons.satf" _
    , "*LIGHTTIME  Tue Jul 23 15:15:53 2024  /disks/d13/project/new_horizons/mops/prod/seq/support/new_horizons.ltf" _
    , "*RULES      Thu Mar 13 14:13:22 2025  /disks/d13/project/new_horizons/mops/prod/seq/fmrf/new_horizons.fmrf" _
    , "*CLOCK      Mon Jul  6 13:54:23 2026  /disks/d13/project/new_horizons/mops/prod/seq/support/new_horizons_current.coeff" _
    , "*SCRIPT     Mon Jan 23 20:09:19 2006  /disks/d13/project/new_horizons/mops/prod/seq/support/new_horizons.script" _
    , "*LEGENDS    Sat Jan 21 14:58:25 2006  /disks/d13/project/new_horizons/mops/prod/seq/support/new_horizons.legend" _
    , "*SEQUENCE   Thu Jul  9 16:49:01 2026  /homes/harchap9/seq_dev/gen/2026/26222/pb/26222_pepssi_swap_pb_v1.sasf" _
    )

    headerC = Array _
    ("*CONDITIONS Wed Jun 24 17:41:55 2026  /disks/d13/project/new_horizons/mops/prod/seq/fincon/26201.fincon" _
    , "*ALLOCATION " _
    , "*BG_SEQUENCE" _
    , "*DEFINITION " _
    , "*DEP_CONTEXT" _
    , "*EVENTS     " _
    , "*MASK       " _
    , "*OPTG_FD    " _
    , "*GEOMETRY   " _
    , "*REDUNDANT  " _
    , "*REQUESTS   " _
    , "*RESOLUTION " _
    , "*TELEMETRY  " _
    , "*TYPEDEF    " _
    , "*VIEWPERIOD " _
    , "*VIEW_FD    " _
    , "************************************************************" _
    , "$$EOH" _
    , "$$EOD" _
    )
        
    SASFHeader = Join(headerA, vbLf) & vbLf & Join(headerB, vbLf) & vbLf & Join(headerC, vbLf)

End Function

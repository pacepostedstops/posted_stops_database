Attribute VB_Name = "STOP_SUMMARY"
Sub MergeMultiWkshts()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim CopyRng As Range

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Delete the summary sheet if it exists.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("Stops_Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Delete the Legend sheet if it exists.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("Legend").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.name = "Stops_Summary"
    Range("A1").Value = "WORKSHEET"
    Range("B1").Value = "STOP_SEQ"
    Range("C1").Value = "STOP_ID"
    Range("D1").Value = "CURRENT_STOP_NAME"
    Range("E1").Value = "ON_STREET"
    Range("F1").Value = "CROSS_STREET"
    Range("G1").Value = "PROPOSED_STOP_NAME"
    Range("H1").Value = "ROUTE_DIRECTION"
    Range("I1").Value = "BUS_DIRECTION"
    Range("J1").Value = "EXISTING_STOP_LOC"
    Range("K1").Value = "EXISTING_STOP_TYPE"
    Range("L1").Value = "PROPOSED_STOP_LOC"
    Range("M1").Value = "SIGN_TEXT"
    
    ' Extra fields for parsing stop location
    '''''
    Range("N1").Value = "STOP LOCATION"
    Range("O1").Value = "POSITION"
    Range("P1").Value = "CORNER"
    '''''

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If sh.name <> DestSh.name And sh.name <> "Communications" Then

            ' Find the last row with data on the summary worksheet.
            Last = DestSh.Cells.SpecialCells(xlCellTypeLastCell).Row

            ' Specify the range to place the data.
            LastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row
            Set CopyRng = sh.Range("A5:K" & LastRow)
            Set CopyRngRouteNo = sh.Range("N5:N" & LastRow)
            
            ' Test to see whether there are enough rows in the summary
            ' worksheet to copy all the data.
            If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                MsgBox "There are not enough rows in the " & _
                   "summary worksheet to place the data."
                GoTo ExitTheSub
            End If

            ' This statement copies values from each
            ' worksheet.
            CopyRng.Copy
            With DestSh.Cells(Last + 1, "B")
                .PasteSpecial xlPasteValues
                .PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End With
            
            CopyRngRouteNo.Copy
            With DestSh.Cells(Last + 1, "M")
                .PasteSpecial xlPasteValues
                .PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End With
            
            'Find last row after copy
            LastPostCopy = DestSh.Cells.SpecialCells(xlCellTypeLastCell).Row
                        
            ' This statement will copy the sheet
            ' name in the A column.
            DestSh.Cells(Last + 1, "A").Resize(CopyRng.Rows.Count).Value = sh.name

        End If
    Next

ExitTheSub:

    Application.Goto DestSh.Cells(1)
        
    ' Populate right-most columns with formulas to parse stop location.
    Range("N2:N" & LastPostCopy).FormulaR1C1 = "=IF(ISNUMBER((SEARCH(""Retain"",RC12))), RC10,RC12)"
    Range("O2:O" & LastPostCopy).FormulaR1C1 = _
       "=IF(ISNUMBER(SEARCH(""NS"",RC[-1]))=TRUE,""Nearside"",IF(ISNUMBER(SEARCH(""FS"",RC[-1]))=TRUE,""Farside"",IF(ISNUMBER(SEARCH(""MB"",RC[-1]))=TRUE,""Midblock"",IF(ISNUMBER(SEARCH(""Parking"",RC[-1]))=TRUE,""Parking lot"",IF(ISNUMBER(SEARCH(""Driveway"",RC[-1]))=TRUE,""Driveway"",IF(ISNUMBER(SEARCH(""Eliminate"",RC[-1]))=TRUE,""Eliminate"",IF(ISNUMBER(SEARCH(""TC"",RC" & _
        "[-1]))=TRUE,""TC"",IF(ISNUMBER(SEARCH(""Terminal"",RC[-1]))=TRUE,""Terminal"",IF(ISNUMBER(SEARCH(""Bay"",RC[-1]))=TRUE,""Bay"",RC[-1])))))))))"
    Range("P2:P" & LastPostCopy).FormulaR1C1 = _
       "=IF(OR(RC[-1]=""Nearside"",RC[-1]=""Farside""),IF(ISNUMBER(SEARCH(""NW"",RC[-2]))=TRUE,""NW"",IF(ISNUMBER(SEARCH(""SW"",RC[-2]))=TRUE,""SW"",IF(ISNUMBER(SEARCH(""SE"",RC[-2]))=TRUE,""SE"",IF(ISNUMBER(SEARCH(""NE"",RC[-2]))=TRUE,""NE"","""")))),RC[-1])"
                         
    ' Remove extraneous formatting
    Range("A1:P" & LastPostCopy).WrapText = False
    Range("A1:P" & LastPostCopy).Font.name = "Arial"
    Range("A1:P" & LastPostCopy).Font.Size = 10
    Range("A1:P" & LastPostCopy).Borders.LineStyle = xlContinuous
    Rows("1:1").Font.Bold = True
    
    ' AutoFit the column width in the summary sheet.
    DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


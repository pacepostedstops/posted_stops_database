''' <summary>
''' Visual Basic script for merging route file/stop lists into data 
''' format compatible for import into the Posted Stops Database.
''' </summary>

Attribute VB_Name = "MergeRouteFiles"
Function Exists(name As String)
     For Each sh In ActiveWorkbook.Worksheets
         If name = sh.name Then
             Set SummarySheet = ActiveWorkbook.Worksheets(name)
             Exit Function
         End If
     Next sh
     Set SummarySheet = ActiveWorkbook.Worksheets.Add
     SummarySheet.name = name
 End Function
    
Sub MergeAllWorkbooks()
    Dim SummarySheet As Worksheet
    Dim Last As Long
    Dim FolderPath As String
    Dim FileName As String
    Dim WorkBk As Workbook
    Dim CopyRng As Range
    

    ' Calls function to determine if merge worksheet exists
    Exists ("Merge")

    ' Names field headings from merge
    Range("A1").Value = "ROUTE"
    Range("B1").Value = "WORKSHEET"
    Range("C1").Value = "STOP_SEQ"
    Range("D1").Value = "STOP_ID"
    Range("E1").Value = "ORIGINAL_STOP_NAME"
    Range("F1").Value = "ON_STREET"
    Range("G1").Value = "CROSS_STREET"
    Range("H1").Value = "PROPOSED_STOP_NAME"
    Range("I1").Value = "ROUTE_DIR"
    Range("J1").Value = "BUS_DIR"
    Range("K1").Value = "EXISTING_STOP_LOC"
    Range("L1").Value = "EXISTING_STOP_TYPE"
    Range("M1").Value = "PROPOSED_STOP_LOC"
    Range("N1").Value = "STOP_AMENITIES"
    Range("O1").Value = "PAX_ACCESSIBILITY"
    Range("P1").Value = "FIELD_NOTES"
    
    ' Extra fields for parsing stop location
    Range("Q1").Value = "STOP_LOCATION"
    Range("R1").Value = "POSITION"
    Range("S1").Value = "CORNER"
	
    ' If replacing the sheet, then set SummarySheet to a new added sheet
    Set SummarySheet = ActiveWorkbook.Worksheets("Merge")
    
    ' Modify this folder path to point to the files you want to use.
    FolderPath = "N:\Sherwin\Test\Database\Import\"
    
    ' Call Dir the first time, pointing it to all Excel files in the folder path.
    FileName = Dir(FolderPath & "*.xl*")
    
    ' Loop until Dir returns an empty string.
    Do While FileName <> ""
	
        ' Open a workbook in the folder
        Set WorkBk = Workbooks.Open(FolderPath & FileName)
        
        ' Loop through all worksheets and copy the data to the
        ' summary worksheet in new workbook.
        For Each sh In WorkBk.Worksheets
            If sh.name <> "Communications" And sh.name <> "Legend" And sh.name <> "Pr_stp_name_Do not Delete" Then
    
                ' Find the last row with data on the summary worksheet.
                Last = SummarySheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
                ' Specify the range of data to copy.
                LastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row
                Set CopyRng = sh.Range("A5:K" & LastRow)
                                
                ' Find the correct column
                colTracker = 1
                For Each col In sh.Range("A3:Z3")
                    If InStr(1, col.Value, "Amenities") > 0 Then
                        AmenitiesCol = colTracker
                    End If
                    If InStr(1, col.Value, "Passenger Accessibility") > 0 Then
                        PaxAccessCol = colTracker
                    End If
                    If InStr(1, col.Value, "Notes") > 0 Then
                        NotesCol = colTracker
                    End If
                    colTracker = colTracker + 1
                Next
                
                Debug.Print FileName, sh.name
                                
                
                ' Specify range of 'Amenities' and 'Notes' to copy
                With sh
                    Set CopyRngAmenities = Range(.Cells(5, AmenitiesCol), .Cells(LastRow, AmenitiesCol))
                    Set CopyRngPaxAccess = Range(.Cells(5, PaxAccessCol), .Cells(LastRow, PaxAccessCol))
                    Set CopyRngNotes = Range(.Cells(5, NotesCol), .Cells(LastRow, NotesCol))
                End With
                                
                ' Copies range of data to wksht.
                CopyRng.Copy
                With SummarySheet.Cells(Last + 1, "C")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With
                                
                ' Copies range of 'Amenities' column to wksht.
                CopyRngAmenities.Copy
                With SummarySheet.Cells(Last + 1, "N")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With
                
                ' Copies range of 'Pax Accessibility' column to wksht.
                CopyRngPaxAccess.Copy
                With SummarySheet.Cells(Last + 1, "O")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With
                                
                ' Copies range of 'Notes' column to wksht.
                CopyRngNotes.Copy
                With SummarySheet.Cells(Last + 1, "P")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With
                
                ' Find last row after copy
                LastPostCopy = SummarySheet.Cells.SpecialCells(xlCellTypeLastCell).Row

                ' Parse route number
                RouteNo = Left(FileName, 3)
                                
                ' Copy sheet/pattern name and route number to columns A & B, respectively
                SummarySheet.Range("B" & Last + 1 & ":B" & LastPostCopy).Value = sh.name
                SummarySheet.Range("A" & Last + 1 & ":A" & LastPostCopy).Value = RouteNo
                
            End If
        Next
        
        ' Close the source workbook without saving changes.
        WorkBk.Close savechanges:=False
        
        ' Use Dir to get the next file name.
        FileName = Dir()
    Loop
        
        
    ' Populate right-most columns with formulas to parse stop location.
    Range("Q2:Q" & LastPostCopy).FormulaR1C1 = "=IF(ISNUMBER((SEARCH(""Retain"",RC13))), RC11,RC13)"
    Range("R2:R" & LastPostCopy).FormulaR1C1 = _
       "=IF(ISNUMBER(SEARCH(""NS"",RC[-1]))=TRUE,""Nearside"",IF(ISNUMBER(SEARCH(""FS"",RC[-1]))=TRUE,""Farside"",IF(ISNUMBER(SEARCH(""MB"",RC[-1]))=TRUE,""Midblock"",IF(ISNUMBER(SEARCH(""Parking"",RC[-1]))=TRUE,""Parking lot"",IF(ISNUMBER(SEARCH(""Driveway"",RC[-1]))=TRUE,""Driveway"",IF(ISNUMBER(SEARCH(""Eliminate"",RC[-1]))=TRUE,""Eliminate"",IF(ISNUMBER(SEARCH(""TC"",RC" & _
        "[-1]))=TRUE,""TC"",IF(ISNUMBER(SEARCH(""Terminal"",RC[-1]))=TRUE,""Terminal"",IF(ISNUMBER(SEARCH(""Bay"",RC[-1]))=TRUE,""Bay"",RC[-1])))))))))"
    Range("S2:S" & LastPostCopy).FormulaR1C1 = _
       "=IF(OR(RC[-1]=""Nearside"",RC[-1]=""Farside""),IF(ISNUMBER(SEARCH(""NW"",RC[-2]))=TRUE,""NW"",IF(ISNUMBER(SEARCH(""SW"",RC[-2]))=TRUE,""SW"",IF(ISNUMBER(SEARCH(""SE"",RC[-2]))=TRUE,""SE"",IF(ISNUMBER(SEARCH(""NE"",RC[-2]))=TRUE,""NE"","""")))),RC[-1])"

    ' Remove extraneous formatting
    Range("A1:S" & LastPostCopy).WrapText = False
    Range("A1:S" & LastPostCopy).Font.name = "Arial"
    Range("A1:S" & LastPostCopy).Font.Size = 10
    Range("A1:S" & LastPostCopy).Borders.LineStyle = xlContinuous
    Range("A1:S" & LastPostCopy).HorizontalAlignment = xlLeft
    Rows("1:1").Font.Bold = True
        
    ' Call AutoFit on the destination sheet so that all
    ' data is readable.
    SummarySheet.Columns.AutoFit
    
    MsgBox LastPostCopy & " rows imported."	

End Sub
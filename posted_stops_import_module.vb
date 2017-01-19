''' <summary>
''' Visual Basic script for importing merged route file/stop list 
''' data into the POSTED_STOPS_LOCATIONS table of the Posted Stops 
''' Database.
''' </summary>

Attribute VB_Name = "ImportStops"
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Function to return network login name
Function fOSUserName() As String
Dim lngLen As Long, lngX As Long
Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        fOSUserName = Left$(strUserName, lngLen - 1)
    Else
        fOSUserName = vbNullString
    End If
End Function

Sub Import()

    Dim sumsh As Excel.Worksheet
    Dim xlWrkBk As Excel.Workbook
    Dim existing As Recordset
    Dim ConvDateMsg, ConvDateTitle As String
    Dim ConvDate As Variant
    
    ' Retrieve import workbook
    Set xlWrkBk = GetObject("N:\Sherwin\Test\Database\database_merge.xlsm")
    ' Retrieve import worksheet
    Set sumsh = xlWrkBk.Worksheets("Merge")

    ' Track column number of import worksheet
    colTracker = 1
    For Each col In sumsh.Range("A1:Z1")
        If col.Value = "ROUTE" Then
            routeCol = colTracker
        End If
        If InStr(1, col.Value, "WORKSHEET") > 0 Then
            patternCol = colTracker
        End If
        If InStr(1, col.Value, "STOP_ID") > 0 Then
            stopIdcol = colTracker
        End If
        If InStr(1, col.Value, "ON_STREET") > 0 Then
            mainStCol = colTracker
        End If
        If InStr(1, col.Value, "CROSS_STREET") > 0 Then
            crossStCol = colTracker
        End If
        If InStr(1, col.Value, "BUS_DIR") > 0 Then
            busdirCol = colTracker
        End If
        If InStr(1, col.Value, "STOP_AMENITIES") > 0 Then
            amenitiesCol = colTracker
        End If
        If InStr(1, col.Value, "PAX_ACCESSIBILITY") > 0 Then
            paxAccessCol = colTracker
        End If
        If InStr(1, col.Value, "FIELD_NOTES") > 0 Then
            notesCol = colTracker
        End If
        If InStr(1, col.Value, "POSITION") > 0 Then
            posCol = colTracker
        End If
        If InStr(1, col.Value, "CORNER") > 0 Then
            cornerCol = colTracker
        End If
        colTracker = colTracker + 1
    Next
    
    ' Find last row of data in import worksheet
    sumLastRow = sumsh.Cells(sumsh.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize row counter
    rows_counter = 0

    ' Create input box for setting conversion date
    ConvDateMsg = "Enter the date of conversion to posted-stops-only for these stops."
    ConvDateTitle = "Posted Stops Conversion Date"
    ConvDate = InputBox(ConvDateMsg, ConvDateTitle)
    
    ' Loop through each row of import worksheet
    For i = 2 To sumLastRow
        ' Copy values of each cell
        Set CopyRoute = sumsh.Cells(i, routeCol)
        Set CopyPattern = sumsh.Cells(i, patternCol)
        Set CopyStopID = sumsh.Cells(i, stopIdcol)
        Set CopyMainSt = sumsh.Cells(i, mainStCol)
        Set CopyCrossSt = sumsh.Cells(i, crossStCol)
        Set CopyBusDir = sumsh.Cells(i, busdirCol)
        Set CopyPos = sumsh.Cells(i, posCol)
        Set CopyCorner = sumsh.Cells(i, cornerCol)
        Set CopyAmenities = sumsh.Cells(i, amenitiesCol)
        Set CopyAccess = sumsh.Cells(i, paxAccessCol)
        Set CopyNotes = sumsh.Cells(i, notesCol)
        
        ' Create recordset of existing DB records
        rsQuery = "SELECT STOP_ID, BUS_DIR, POSITION FROM POSTED_STOPS_LOCATIONS"
        Set existing = CurrentDb.OpenRecordset(rsQuery)
        
        ' Searches if current stop exists in DB records
        existing.FindFirst "[STOP_ID] = '" & CopyStopID & "'"
                
            ' Stop does not exist in database
            If existing.NoMatch Then
                        
                ' Reads in all non-existing stops that are not non-IBS stops or "removed" stops
                If CopyStopID <> "" And CopyPos <> "Remove" Then
                                
                    ' Read in new stops
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO POSTED_STOPS_LOCATIONS(STOP_ID, ON_STREET, CROSS_STREET, BUS_DIR, POSITION, CORNER, " _
                                 & "AMENITIES, PAX_ACCESSIBILITY, FIELD_NOTES, ACTIVE, STOP_CHANGE, NOTES, DATE_IMPORTED, REQUESTER, " _
								 & "IMPORTED_BY, EFFECTIVE_DATE) " _
                                 & "VALUES('" & CopyStopID & "','" & CopyMainSt & "','" & CopyCrossSt & "','" & CopyBusDir & "','" & CopyPos & "'" _
                                 & ",'" & CopyCorner & "','" & CopyAmenities & "','" & CopyAccess & "','" & Replace(CopyNotes, "'", "''") _
                                 & "', TRUE, FALSE, 'Initial Import for Posted Stops.','" & Date & "','Posted Stops','" & fOSUserName & "','" & ConvDate & "');"
                    DoCmd.SetWarnings True
                    rows_counter = rows_counter + 1
                                        
                End If
                                
            ' Stop does exist in database - stop is not read in
            Else
                        
                ' Searches if existing stop matches stop ID, bus direction, and position
                existing.FindFirst "[STOP_ID] = '" & CopyStopID & "' AND [BUS_DIR] = '" & CopyBusDir & "' AND [POSITION] = '" & CopyPos & "'"
                                
                    ' Existing stop has one key field changed - followup required
                    If existing.NoMatch Then
                        Debug.Print "Stop not imported - key field changed. " & "Route: " & CopyRoute & ", Pattern: " & _
                        CopyPattern & ", Stop ID: " & CopyStopID & ", Bus Dir: " & CopyBusDir & ", Position: " & CopyPos
                                                
                    ' Existing stop has no key fields changed
                    ' Else
                        ' Debug.Print "Existing stop (no rev.):", CopyStopID, CopyBusDir, CopyPos
                                                
                    End If
            End If
    Next
    
    ' Print import summary
    Debug.Print rows_counter & " of " & sumLastRow - 1 & " rows imported"
    MsgBox "Import complete. Number of rows imported: " & rows_counter & " of " & sumLastRow - 1 & ""
    
End Sub

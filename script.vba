Sub FilterAndSplitData()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim driverName As String
    Dim cell As Range
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headers As Variant
    Dim driverDict: Set driverDict = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    Dim currentSheetName As String
    Dim headerRow As Long
    Dim targetRow As Long
    Dim i As Long
    Dim rowOffset As Long
    Dim providerCol As Long
    Dim dataCols As Range

    On Error GoTo ErrorHandler

    ' Ensure we are working with the active workbook and sheet
    Set dataSheet = ActiveWorkbook.ActiveSheet

    ' Get the current sheet name
    currentSheetName = dataSheet.Name

    ' Define headers to keep (excluding Driver Name)
    headers = Array("Stop Number", "Visit Name", "Address", "Phone", "Notes", "Provider")

    ' Find the last row and column
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    lastCol = dataSheet.Cells(1, dataSheet.Columns.Count).End(xlToLeft).Column

    ' Loop through the data and add rows to the dictionary by driver name
    For Each cell In dataSheet.Range("A2:A" & lastRow)
        driverName = cell.Value
        If cell.Offset(0, 2).Value <> 0 Then ' Check if Stop Number is not 0
            If Not driverDict.exists(driverName) Then
                driverDict.Add driverName, New Collection
            End If
            driverDict(driverName).Add cell.EntireRow
        End If
    Next cell

    ' Find the column number for "Provider"
    providerCol = Application.Match("Provider", dataSheet.Rows(1), 0)

    ' Loop through the dictionary and create new sheets
    For Each key In driverDict.keys
        Set newWs = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        newWs.Name = key
        
        ' Set font size to 10 for the entire sheet
        newWs.Cells.Font.Size = 10

        ' Add headers and rename "Stop Number" to "#"
        For i = LBound(headers) To UBound(headers)
            If headers(i) = "Stop Number" Then
                newWs.Cells(1, i + 1).Value = "#"
            Else
                newWs.Cells(1, i + 1).Value = headers(i)
            End If
            ' Set the background color of the header cells to #A6A6A6 and make the text bold
            With newWs.Cells(1, i + 1)
                .Interior.Color = RGB(166, 166, 166)
                .Font.Bold = True
            End With
        Next i

        ' Add rows and shift them up to below the header row
        rowOffset = 1
        For Each cell In driverDict(key)
            rowOffset = rowOffset + 1
            For i = LBound(headers) To UBound(headers)
                newWs.Cells(rowOffset, i + 1).Value = cell.Cells(1, Application.Match(headers(i), dataSheet.Rows(1), 0)).Value
            Next i

            ' Change background color if the provider is not "Traditional Kitchen"
            If cell.Cells(1, providerCol).Value <> "Traditional Kitchen" Then
                Set dataCols = newWs.Range(newWs.Cells(rowOffset, 1), newWs.Cells(rowOffset, UBound(headers) + 1))
                dataCols.Interior.Color = RGB(217, 217, 217)
            End If
        Next cell

        ' Apply all borders to all cell data
        With newWs.Range(newWs.Cells(1, 1), newWs.Cells(rowOffset, UBound(headers) + 1)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        ' Format columns as per the requirements
        With newWs
            .Columns("A:A").ColumnWidth = 2 ' Stop Number (renamed to #)
            .Columns("B:B").ColumnWidth = 12 ' Visit Name
            .Columns("C:C").ColumnWidth = 40 ' Address
            .Columns("D:D").ColumnWidth = 11 ' Phone
            .Columns("E:E").ColumnWidth = 35 ' Notes
            .Columns("F:F").ColumnWidth = 18 ' Provider
            
            ' Center all cells horizontally and vertically
            .Cells.HorizontalAlignment = xlCenter
            .Cells.VerticalAlignment = xlCenter
            
            ' Wrap text in all cells
            .Cells.WrapText = True
            
            ' Set the print settings to landscape and fit all columns on one page, and set the sheet name as a header
            With .PageSetup
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .CenterHeader = "&""Arial,Bold""&16" & newWs.Name
            End With
        End With
    Next key

    ' Remove all columns except specified headers in dataSheet
    For i = lastCol To 1 Step -1
        If IsError(Application.Match(dataSheet.Cells(1, i).Value, headers, 0)) Then
            dataSheet.Columns(i).Delete
        End If
    Next i

    ' Remove rows where Stop Number is 0 in dataSheet
    For i = lastRow To 2 Step -1
        If dataSheet.Cells(i, Application.Match("Stop Number", dataSheet.Rows(1), 0)).Value = 0 Then
            dataSheet.Rows(i).Delete
        End If
    Next i

    ' Shift all rows with data up, to the row below the row containing the column names
    headerRow = 1
    targetRow = headerRow + 1

    For i = targetRow To lastRow
        If Application.WorksheetFunction.CountA(dataSheet.Rows(i)) > 0 Then
            dataSheet.Rows(i).Cut Destination:=dataSheet.Rows(targetRow)
            targetRow = targetRow + 1
        End If
    Next i

    ' Delete the initial sheet
    Application.DisplayAlerts = False
    dataSheet.Delete
    Application.DisplayAlerts = True

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
    Application.DisplayAlerts = True
End Sub

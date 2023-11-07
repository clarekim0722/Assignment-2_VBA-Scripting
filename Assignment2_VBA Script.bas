Attribute VB_Name = "Module1"
Sub ExtractDataToColumnJ()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    Dim i As Long, j As Long

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    j = 2

    For i = 2 To lastRow Step 251
        ws.Cells(j, "J").Value = ws.Cells(i, "C").Value
        j = j + 1
    Next i
End Sub

Sub ExtractDataToColumnK()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    Dim i As Long, j As Long

    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    j = 2 ' Start at row 2 in Column K

    For i = 252 To lastRow Step 251
        ws.Cells(j, "K").Value = ws.Cells(i, "F").Value
        j = j + 1
    Next i
End Sub

Sub SubtractAndColorize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    Dim i As Long
    Dim result As Double

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        result = ws.Cells(i, "K").Value - ws.Cells(i, "J").Value
        ws.Cells(i, "L").Value = result

        ' Apply conditional formatting to cells in Column L
        If result < 0 Then
            ws.Cells(i, "L").Interior.Color = RGB(255, 0, 0) ' Red
        ElseIf result > 0 Then
            ws.Cells(i, "L").Interior.Color = RGB(0, 255, 0) ' Green
        Else
            ' Reset cell color if result is zero
            ws.Cells(i, "L").Interior.ColorIndex = xlNone
            
    
        End If
    Next i
End Sub


Sub CalculatePercentageChangeAndRound()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    Dim i As Long
    Dim percentageChange As Double

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        Dim oldValue As Double
        Dim newValue As Double
        oldValue = ws.Cells(i, "J").Value
        newValue = ws.Cells(i, "K").Value
        
        ' Calculate the percentage change and round to two decimal places
        If oldValue <> 0 Then
            percentageChange = Round(((newValue - oldValue) / Abs(oldValue)) * 100, 2)
        Else
            percentageChange = 0
        End If

        ' Place the result in Column M
        ws.Cells(i, "M").Value = percentageChange
    Next i
End Sub

Sub SumEveryOther250Rows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018") ' Change "Sheet1" to your sheet's name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    Dim i As Long
    Dim sum As Double
    Dim j As Long
    j = 2 ' Start at row 2 in Column N

    For i = 2 To lastRow Step 250
        sum = Application.WorksheetFunction.sum(ws.Range("G" & i & ":G" & (i + 250 - 1)))
        ws.Cells(j, "N").Value = sum
        j = j + 1
    Next i
End Sub

Sub FindGreatestPositivePercentageIncrease()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    Dim maxPositivePercentageIncrease As Double
    Dim stockName As String

    ' Initialize maxPositivePercentageIncrease to a very small number
    maxPositivePercentageIncrease = -1E+20

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        If ws.Cells(i, "M").Value > 0 And ws.Cells(i, "M").Value > maxPositivePercentageIncrease Then
            maxPositivePercentageIncrease = ws.Cells(i, "M").Value
            stockName = ws.Cells(i, 1).Value ' Assuming the stock name is in Column A
        End If

    Next i

    ' Display the stock with the greatest positive percentage increase in Cell(2, 18)
    If maxPositivePercentageIncrease > -1E+20 Then
        ws.Cells(2, 18).Value = "Stock: " & stockName & " (" & maxPositivePercentageIncrease & "%)"
    Else
        ws.Cells(2, 18).Value = "No positive values found in Column M"
    End If
End Sub


Sub FindUniqueValuesInColumnA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018") ' Change "Sheet1" to your sheet's name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim uniqueValues As Collection
    Set uniqueValues = New Collection

    Dim cell As Range
    Dim i As Long

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        Set cell = ws.Cells(i, "A")
        On Error Resume Next
        uniqueValues.Add cell.Value, CStr(cell.Value)
        On Error GoTo 0
    Next i

    ' Paste unique values to Column I starting from row 2
    For i = 2 To uniqueValues.Count + 1
        ws.Cells(i, "I").Value = uniqueValues.Item(i)
    Next i
End Sub


Sub FindGreatestPercentageIncreaseWithStockName()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    Dim maxPositivePercentageIncrease As Double
    Dim stockName As String

    ' Initialize maxPositivePercentageIncrease to a very small number
    maxPositivePercentageIncrease = -1E+20

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        If ws.Cells(i, "M").Value > 0 And ws.Cells(i, "M").Value > maxPositivePercentageIncrease Then
            maxPositivePercentageIncrease = ws.Cells(i, "M").Value
            stockName = ws.Cells(i, 9).Value ' Assuming the stock name is in Column I
        End If
    Next i

    ' Display the stock with the greatest positive percentage increase and the stock name in Cell(2, 18)
    If maxPositivePercentageIncrease > -1E+20 Then
        ws.Cells(2, 18).Value = "Stock: " & stockName & " (" & maxPositivePercentageIncrease & "%)"
    Else
        ws.Cells(2, 18).Value = "No positive values found in Column M"
    End If
End Sub


Sub FindGreatestPercentageDecreaseWithStockName()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    Dim maxNegativePercentageDecrease As Double
    Dim stockName As String

    ' Initialize maxNegativePercentageDecrease to a very large positive number
    maxNegativePercentageDecrease = 1E+20

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        If ws.Cells(i, "M").Value < 0 And ws.Cells(i, "M").Value < maxNegativePercentageDecrease Then
            maxNegativePercentageDecrease = ws.Cells(i, "M").Value
            stockName = ws.Cells(i, 9).Value ' Assuming the stock name is in Column I
        End If
    Next i

    ' Display the stock with the greatest negative percentage decrease and the stock name in Cell(3, 18)
    
   If maxNegativePercentageDecrease < 1E+20 Then
        ws.Cells(3, 18).Value = "Stock: " & stockName & " (" & maxNegativePercentageDecrease & "%)"
    Else
        ws.Cells(3, 18).Value = "No negative values found in Column M"
    End If
    
End Sub

Sub FindGreatestTotalVolumeWithStockName()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("2018")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    Dim maxTotalVolume As Double
    Dim stockName As String

    ' Initialize maxTotalVolume to a very small number
    maxTotalVolume = -1E+20

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        If ws.Cells(i, "N").Value > maxTotalVolume Then
            maxTotalVolume = ws.Cells(i, "N").Value
            stockName = ws.Cells(i, 9).Value ' Assuming the stock name is in Column I
        End If
    Next i

    ' Display the stock with the greatest total volume and the stock name in Cell(4, 18)
    If maxTotalVolume > -1E+20 Then
        ws.Cells(4, 18).Value = "Stock: " & stockName & " (" & maxTotalVolume & ")"
    Else
        ws.Cells(4, 18).Value = "No values found in Column N"
    End If
End Sub

Sub ProcessWorksheets()
    Dim ws As Worksheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the worksheet should be processed (e.g., based on its name)
        If ws.Name = "Sheet1" Or ws.Name = "2019" Or ws.Name = "2020" Then
            ' Call the subroutine to process the current worksheet
            ProcessWorksheet ws
        End If
    Next ws
End Sub

Sub ProcessWorksheet(ws As Worksheet)
    ' Put your existing VBA script for processing a single worksheet here
    ' For example, the script to find the stock with the greatest total volume

    Dim lastRow As Long
    Dim maxTotalVolume As Double
    Dim stockName As String

    ' Initialize maxTotalVolume to a very small number
    maxTotalVolume = -1E+20

    ' Find the last used row in Column N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    ' Loop through rows starting from row 2
    For i = 2 To lastRow
        If ws.Cells(i, "N").Value > maxTotalVolume Then
            maxTotalVolume = ws.Cells(i, "N").Value
            stockName = ws.Cells(i, 9).Value ' Assuming the stock name is in Column I
        End If
    Next i

    ' Display the stock with the greatest total volume and the stock name in Cell(4, 18)
    If maxTotalVolume > -1E+20 Then
        ws.Cells(4, 18).Value = "Stock: " & stockName & " (" & maxTotalVolume & ")"
    Else
        ws.Cells(4, 18).Value = "No values found in Column N"
    End If
End Sub




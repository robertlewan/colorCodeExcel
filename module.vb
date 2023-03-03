Sub ColorCellsByValue()
    Dim lastRow As Long
    Dim valueCount As Long
    Dim valueDict As Object
    Dim valueArray As Variant
    Dim i As Long
    Dim rng As Range
    Dim cell As Range
    Dim color As Long
    
    ' Get the last row of data in column C
    lastRow = Cells(Rows.Count, "C").End(xlUp).Row
    
    ' Create a dictionary to store the values and their counts
    Set valueDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the cells in column C and count the values
    For i = 2 To lastRow 'assuming row 1 is header row
        If Not IsEmpty(Cells(i, "C")) Then
            If valueDict.Exists(Cells(i, "C").Value) Then
                valueDict(Cells(i, "C").Value) = valueDict(Cells(i, "C").Value) + 1
            Else
                valueDict.Add Cells(i, "C").Value, 1
            End If
        End If
    Next i
    
    ' Get the unique values from the dictionary and store them in an array
    valueArray = valueDict.Keys
    
    ' Loop through the array of values and assign a random color to each one
    For i = LBound(valueArray) To UBound(valueArray)
        color = Int((16 ^ 6 - 1 + 1) * Rnd()) ' Generate a random color
        Set rng = Range("C:C").Find(What:=valueArray(i), LookIn:=xlValues, LookAt:=xlWhole) ' Find the range of cells containing the value
        If Not rng Is Nothing Then
            firstCell = rng.Address ' Store the address of the first cell found
            Do
                ' Loop through the cells in the range and fill them with the random color
                For Each cell In rng
                    If cell.Value = valueArray(i) Then
                        cell.Interior.Color = color
                    End If
                Next cell
                Set rng = Range("C:C").FindNext(rng) ' Find the next cell with the search value
            Loop Until rng.Address = firstCell ' Exit the loop if we've returned to the first cell
        End If
    Next i
End Sub

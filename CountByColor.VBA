    'Created function that sorts by color
    Function CountByColor(rng As Range, clr As Range) As Long
    Dim cell As Range
    For Each cell In rng
        If cell.Interior.Color = clr.Interior.Color Then
            CountByColor = CountByColor + 1
        End If
    Next cell
    End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rowNum As Long
    
    ' Loop through each cell in the target range
    For Each cell In Target
        rowNum = cell.Row
        ' SENDER INFO
        ' If the change is in column C (e.g., C5, C6, C7, ...)
        If cell.Column = 3 Then
            If Range("D" & rowNum).Value <> "" Then
                Range("D" & rowNum).Value = ""
            End If
            If Range("E" & rowNum).Value <> "" Then
                Range("E" & rowNum).Value = ""
            End If
        End If
        
        ' If the change is in column D (e.g., D5, D6, D7, ...)
        If cell.Column = 4 Then
            If Range("E" & rowNum).Value <> "" Then
                Range("E" & rowNum).Value = ""
            End If
        End If
        
        ' RECEIVER INFO
        ' If the change is in column I (e.g., I5, I6, I7, ...)
        If cell.Column = 9 Then
            If Range("J" & rowNum).Value <> "" Then
                Range("J" & rowNum).Value = ""
            End If
            If Range("K" & rowNum).Value <> "" Then
                Range("K" & rowNum).Value = ""
            End If
        End If
        
        ' If the change is in column J (e.g., J5, J6, J7, ...)
        If cell.Column = 10 Then
            If Range("K" & rowNum).Value <> "" Then
                Range("K" & rowNum).Value = ""
            End If
        End If
    Next cell
End Sub


# Healthcare-performance-analysis


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watchRange As Range
    Dim cell As Range
    Dim dataLogSheet As Worksheet
    Dim rowOffset As Long
    Dim colOffset As Long

    Set dataLogSheet = ThisWorkbook.Sheets("Data Log")
    Set watchRange = Me.Range("F2:S50") ' Expanded tracking range (14 columns)

    If Not Intersect(Target, watchRange) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Intersect(Target, watchRange)
            If cell.Value = "Yes" Or cell.Value = "No" Then
                rowOffset = cell.Row
                colOffset = cell.Column - 2 ' Shifting F?D, G?E, ..., S?Q
                If dataLogSheet.Cells(rowOffset, colOffset).Value = "" Then
                    dataLogSheet.Cells(rowOffset, colOffset).Value = Date
                End If
            ElseIf cell.Value = "" Then
                dataLogSheet.Cells(cell.Row, cell.Column - 2).Value = ""
            End If
        Next cell
        Application.EnableEvents = True
    End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim watchRange As Range
    Dim cell As Range
    Dim dataLogSheet As Worksheet
    Dim rowOffset As Long
    Dim colOffset As Long

    Set dataLogSheet = ThisWorkbook.Sheets("Data Log")
    Set watchRange = Me.Range("F2:T50") ' 17 columns: F ? T

    If Not Intersect(Target, watchRange) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Intersect(Target, watchRange)
            If cell.Value = "Yes" Or cell.Value = "No" Then
                rowOffset = cell.Row
                colOffset = cell.Column - 2 ' F?D, G?E, ..., T?R
                If dataLogSheet.Cells(rowOffset, colOffset).Value = "" Then
                    dataLogSheet.Cells(rowOffset, colOffset).Value = Date
                End If
            ElseIf cell.Value = "" Then
                dataLogSheet.Cells(cell.Row, cell.Column - 2).Value = ""
            End If
        Next cell
        Application.EnableEvents = True
    End If
End Sub

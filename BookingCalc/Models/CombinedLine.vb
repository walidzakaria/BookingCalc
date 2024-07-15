Public Class CombinedLine
    Public Property Details As String
    Public Property StartDate As Date
    Public Property EndDate As Date
    Public Property Value As Double

    Public Sub New(details As String, startDate As Date, endDate As Date, value As Double)
        Me.Details = details
        Me.StartDate = startDate
        Me.EndDate = endDate
        Me.Value = value
    End Sub

    Public Shared Sub AddLine(lines As List(Of CombinedLine), newLine As CalculationLine)
        Dim existing As Boolean = False
        For Each line In lines
            If line.Details = newLine.Details AndAlso line.Value = newLine.Value AndAlso line.EndDate.AddDays(1) = newLine.TravelDate Then
                existing = True
                line.EndDate = newLine.TravelDate
                Exit For
            End If
        Next
        If Not existing Then
            lines.Add(New CombinedLine(newLine.Details, newLine.TravelDate, newLine.TravelDate, newLine.Value))
        End If

    End Sub
End Class

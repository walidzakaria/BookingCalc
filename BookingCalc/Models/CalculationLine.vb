Public Class CalculationLine
    Public Property Details As String
    Public Property TravelDate As Date
    Public Property Value As Double?

    Public Sub New(details As String, travelDate As Date, value As Double?)
        Me.Details = details
        Me.TravelDate = travelDate
        Me.Value = value
    End Sub
End Class

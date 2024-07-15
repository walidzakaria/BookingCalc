Public Class Person
    Public Property Name As String
    Public Property Age As Integer?
    Public Property PersonType As PersonType
    Public Property BaseRate As List(Of CalculationLine)
    Public Property MealSupplement As List(Of CalculationLine)
    Public Property Discount As List(Of CalculationLine)

    Public Sub New(name As String, age As Integer, personType As PersonType)
        Me.Name = name
        Me.Age = age
        Me.PersonType = personType
    End Sub

    Public Function GetCalcTotal() As Double
        Dim baseRateValue = BaseRate.Sum(Function(c) If(c.Value, 0))
        Dim mealSupplementValue = MealSupplement.Sum(Function(c) If(c.Value, 0))
        Dim discountValue = Discount.Sum(Function(c) If(c.Value, 0))

        Return baseRateValue + mealSupplementValue + discountValue
    End Function
End Class

Public Class Booking
    Public Property BookingDate As Date
    Public Property TravelFrom As Date
    Public Property TravelTo As Date
    Public Property RoomType As Integer
    Public Property MealPlan As Integer
    Public Property Pax As List(Of Person)
    Public Property Calculation As List(Of CombinedLine)
    Public Property CalculationTotal As Double

    Public Sub New(bookingDate As Date, travelFrom As Date, travelTo As Date,
                   roomType As Integer, mealPlan As Integer, pax As List(Of Person))
        Me.BookingDate = bookingDate
        Me.TravelFrom = travelFrom
        Me.TravelTo = travelTo
        Me.RoomType = roomType
        Me.MealPlan = mealPlan
        Me.Pax = pax
    End Sub

    Public Sub Calculate(contractId As Integer)
        CalcAdultRate(contractId)
        CalcChildRate(contractId)
        CalcAdultBoard(contractId)
        CalcChildBoard(contractId)
        CalcAdultDiscounts(contractId)
        ShowCalc()
        ShowTotal()
    End Sub

    Public Sub ShowTotal()
        CalculationTotal = 0
        For Each person In Pax
            CalculationTotal += person.GetCalcTotal()
        Next
    End Sub
    Public Sub ShowCalc()
        Calculation = New List(Of CombinedLine)
        Dim adultCounter As Integer
        Dim childCounter As Integer
        Dim prefix As String
        For Each person In Pax
            If person.PersonType = PersonType.Adult Then
                adultCounter += 1
                prefix = $"Adult {adultCounter}"
            Else
                childCounter += 1
                prefix = $"Child {adultCounter}"
            End If
            For Each baseRate In person.BaseRate
                baseRate.Details = $"{prefix} {baseRate.Details}"
                CombinedLine.AddLine(Calculation, baseRate)
            Next
            For Each supplement In person.MealSupplement
                supplement.Details = $"{prefix} {supplement.Details}"
                CombinedLine.AddLine(Calculation, supplement)
            Next
            For Each discount In person.Discount
                discount.Details = $"{prefix} {discount.Details}"
                CombinedLine.AddLine(Calculation, discount)
            Next
        Next
    End Sub
    Public Function GetCalcDetails() As List(Of CalculationLine)
        Dim result = New List(Of CalculationLine)
        For Each person In Pax
            result.AddRange(person.BaseRate)
        Next
        Return result
    End Function

    Public Sub CalcAdultRate(contractId As Integer)
        Dim adultsRateQuery As String = $"
            SELECT Csg.SeasonFrom, Csg.SeasonTo, Csg.SeasonGroup, Ra.AdultRate, Ra.ExtraAdultRate, Ra.SGLSuppl, Ra.ApplyAs
            FROM tblhbContractRate Ra
            JOIN tblhbContractSeasonGroup Csg ON Ra.SeasonGroup = Csg.SeasonGroup
            WHERE Ra.ContractID = {contractId} AND Ra.RoomCategoryID = {RoomType};
        "
        Dim adultsRateTable = Utils.GetData(adultsRateQuery)
        Dim stdPax = 2  ' This needs to be retrieved from the database
        ' Check if the room is single
        Dim singleRoom As Boolean = Pax.Where(Function(p) p.PersonType = PersonType.Adult).Count() = 1
        If singleRoom Then
            ' Calculate the SGL Person
            Dim stdAdultRate = GetAdultRate(adultsRateTable, "AdultRate", $"Rate Adult (1) Rate", CalcType.Value, Nothing)
            Pax(0).BaseRate = stdAdultRate
            Dim calculationType As CalcType
            If CInt(adultsRateTable.Rows(0)("ApplyAs")) < 3 Then
                calculationType = CalcType.Value
            Else
                calculationType = CalcType.Percent
            End If
            Dim singleRates = GetAdultRate(adultsRateTable, "SGLSuppl", $"SGL Suppl. (1) Rate", calculationType, CalculationBase.Supplement)
            Pax(0).BaseRate.AddRange(singleRates)
        Else
            ' Calculate all adults
            For x As Integer = 1 To Pax.Count
                If Pax(x - 1).PersonType = PersonType.Child Then Continue For
                Dim stdAdultRates = GetAdultRate(adultsRateTable, "AdultRate", $"Rate Adult ({x}) Rate", CalcType.Value, Nothing)
                Pax(x - 1).BaseRate = stdAdultRates
            Next

            ' Add Extra adults reductions
            Dim extraCalcType As CalcType
            If CInt(adultsRateTable.Rows(0)("ApplyAs")) Mod 2 = 0 Then
                extraCalcType = CalcType.Percent
            Else
                extraCalcType = CalcType.Value
            End If
            For x As Integer = stdPax + 1 To Pax.Count
                If Pax(x - 1).PersonType = PersonType.Child Then Continue For
                Dim extraAdultReduction = GetAdultRate(adultsRateTable, "ExtraAdultRate", $"Extra Adult ({x}) Reduction", extraCalcType, CalculationBase.Reduction)
                Pax(x - 1).BaseRate.AddRange(extraAdultReduction)
            Next
        End If
    End Sub

    Public Sub CalcChildRate(contractId As Integer)
        Dim adultsRateQuery As String = $"
            SELECT Csg.SeasonFrom, Csg.SeasonTo, Csg.SeasonGroup, Cra.ChildString, Chp.AgeFrom, Chp.AgeTo,
                Ra.AdultRate, ChB.ChildValue1, ChB.ChildValue2, ChB.ChildValue3, ChB.ChildValue4, ChB.ChildValue5,
                ChB.ApplyAs
            FROM tblhbContractRate Ra
            JOIN tblhbContractSeasonGroup Csg ON Ra.SeasonGroup = Csg.SeasonGroup
            JOIN tblhbContractChildRoomRate Cra ON Ra.hbContractRateID = Cra.hbContractRateID
            JOIN tblhbContractChildPolicy Chp ON Cra.hbContractPolicyChildID = Chp.hbContractPolicyChildID
            JOIN tblhbContractChildBoard ChB ON Chp.hbContractPolicyChildID = ChB.hbContractPolicyChildID
            WHERE Ra.ContractID = {contractId}
                AND Ra.RoomCategoryID = {RoomType}
                AND ChB.hbContractAdultBoardID = {MealPlan} AND ChB.hbContractRoomDetID = {RoomType};
        "
        Dim childRatesTable = Utils.GetData(adultsRateQuery)
        Dim childType As CalcType
        If CInt(childRatesTable.Rows(0)("ApplyAs")) = 1 Then
            childType = CalcType.Value
        Else
            childType = CalcType.Percent
        End If

        Dim childCounter As Integer = 0

        For Each child In Pax
            If child.PersonType = PersonType.Adult Then Continue For
            childCounter += 1
            Dim childRates = GetChildRate(childRatesTable, child, childType, childCounter, "Rate")
            child.BaseRate = childRates
        Next

    End Sub
    Public Function GetAdultRate(rates As DataTable, rateHeader As String, details As String _
                                 , calcType As CalcType, calculationBase As CalculationBase?) _
                                 As List(Of CalculationLine)

        Dim result = New List(Of CalculationLine)
        Dim travelDate As Date = TravelFrom
        Dim description As String
        Dim rateValue As Double
        While travelDate <= TravelFrom
            Dim rate = rates.AsEnumerable().Where(
                Function(r) travelDate >= CDate(r("SeasonFrom")) _
                AndAlso travelDate >= CDate(r("SeasonTo"))).First()
            If calculationBase Is Nothing Then
                description = $"{details} value"
                rateValue = CDbl(rate(rateHeader))
            ElseIf calculationBase.Value = BookingCalc.CalculationBase.Reduction Then
                If calcType = CalcType.Value Then
                    description = $"{details} value reduction"
                    rateValue = -CDbl(rate(rateHeader))
                Else
                    rateValue = CDbl(rate("AdultRate"))
                    Dim percentValue = CDbl(rate(rateHeader))
                    rateValue = rateValue * (100 - (percentValue / 100))
                    description = $"{details} reduction {percentValue}%"
                End If
            Else
                If calcType = CalcType.Value Then
                    description = $"{details} value"
                    rateValue = CDbl(rate(rateHeader))
                Else
                    rateValue = CDbl(rate("AdultRate"))
                    Dim percentValue = CDbl(rate(rateHeader))
                    rateValue = rateValue * (1 + (percentValue / 100))
                    description = $"{details} {percentValue}%"
                End If
            End If

            Dim value = New CalculationLine(description, travelDate, rateValue)

            travelDate.AddDays(1)
        End While
        Return result
    End Function

    Public Function GetChildRate(rates As DataTable, child As Person _
                                 , calcType As CalcType, childOrder As Short, details As String) _
                                 As List(Of CalculationLine)

        Dim result = New List(Of CalculationLine)
        Dim travelDate As Date = TravelFrom
        Dim description As String
        Dim rateValue As Double
        While travelDate <= TravelFrom
            Dim rate = rates.AsEnumerable().Where(
                Function(r) travelDate >= CDate(r("SeasonFrom")) _
                AndAlso travelDate >= CDate(r("SeasonTo")) _
                AndAlso child.Age >= CInt(r("AgeFrom")) _
                AndAlso child.Age <= CInt(r("AgeTo"))).First()

            If calcType = CalcType.Value Then
                description = $"{details} Child {childOrder} value: {travelDate:yyyy-MM-dd}"
                rateValue = CDbl(rate($"ChildValue{childOrder}"))
            Else
                rateValue = CDbl(rate("AdultRate"))
                Dim percentValue = CDbl(rate($"ChildValue{childOrder}"))
                rateValue = rateValue * (100 - (percentValue / 100))
                description = $"{details} Child ({childOrder}) reduction {percentValue}%: {travelDate:yyyy-MM-dd}"
            End If

            Dim value = New CalculationLine(description, travelDate, rateValue)

            travelDate.AddDays(1)
        End While
        Return result
    End Function

    Public Sub CalcAdultBoard(contractId As Integer)
        Dim adultsRateQuery As String = $"
            SELECT Csg.SeasonFrom, Csg.SeasonTo, Csg.SeasonGroup, Ra.AdultRate, Ra.ExtraAdultRate, Ra.ApplyAs
            FROM tblhbContractAdultBoard Ra
            JOIN tblhbContractSeasonGroup Csg ON Ra.SeasonGroup = Csg.SeasonGroup
            WHERE Ra.ContractID = {contractId} AND Ra.BoardID = {MealPlan} AND Ra.RoomCategoryID = {RoomType};
        "
        Dim adultsRateTable = Utils.GetData(adultsRateQuery)

        ' Calculate all adults
        Dim allAdults As Integer = Pax.Where(Function(p) p.PersonType = PersonType.Adult).Count()
        For x As Integer = 1 To allAdults
            Dim stdAdultRates = GetAdultRate(adultsRateTable, "AdultRate", $"Board Adult ({x}) Rate", CalcType.Value, Nothing)
            Pax(x - 1).MealSupplement = stdAdultRates
        Next
    End Sub
    Public Sub CalcChildBoard(contractId As Integer)
        Dim adultsRateQuery As String = $"
            SELECT Csg.SeasonFrom, Csg.SeasonTo, Csg.SeasonGroup, Cra.ChildString, Chp.AgeFrom, Chp.AgeTo,
                Ra.AdultRate, ChB.ChildValue1, ChB.ChildValue2, ChB.ChildValue3, ChB.ChildValue4, ChB.ChildValue5,
                ChB.ApplyAs
            FROM tblhbContractRate Ra
            JOIN tblhbContractSeasonGroup Csg ON Ra.SeasonGroup = Csg.SeasonGroup
            JOIN tblhbContractChildRoomRate Cra ON Ra.hbContractRateID = Cra.hbContractRateID
            JOIN tblhbContractChildPolicy Chp ON Cra.hbContractPolicyChildID = Chp.hbContractPolicyChildID
            JOIN tblhbContractChildBoard ChB ON Chp.hbContractPolicyChildID = ChB.hbContractPolicyChildID
            WHERE Ra.ContractID = {contractId}
                AND ChB.hbContractAdultBoardID = {MealPlan}
                AND ChB.hbContractRoomDetID = {RoomType};
        "
        Dim childRateTable = Utils.GetData(adultsRateQuery)

        Dim childType As CalcType
        If CInt(childRateTable.Rows(0)("ApplyAs")) = 1 Then
            childType = CalcType.Value
        Else
            childType = CalcType.Percent
        End If

        Dim childCounter As Integer = 0

        For Each child In Pax
            If child.PersonType = PersonType.Adult Then Continue For
            childCounter += 1
            Dim childBoard = GetChildRate(childRateTable, child, childType, childCounter, "Board Child")
            child.MealSupplement = childBoard
        Next

    End Sub

    Public Sub CalcAdultDiscounts(contractId As Integer)
        ' @TODO: fix the query
        Dim adultsRateQuery As String = $"
            SELECT *
            FROM tblhbContractOffers Ofh
            WHERE Ra.ContractID = {contractId} AND Ra.BoardID = {MealPlan} AND Ra.RoomCategoryID = {RoomType};
        "
        Dim adultsDiscountTable = Utils.GetData(adultsRateQuery)

        ' Calculate all adults
        Dim allAdults As Integer = Pax.Where(Function(p) p.PersonType = PersonType.Adult).Count()
        For x As Integer = 1 To allAdults
            Dim adultDiscounts = GetReductionRate(adultsDiscountTable, "Adult Reduction", Pax(x - 1))
            Pax(x - 1).Discount = adultDiscounts
        Next
    End Sub

    Public Function GetReductionRate(discounts As DataTable, details As String _
                                 , person As Person) _
                                 As List(Of CalculationLine)

        Dim result = New List(Of CalculationLine)
        Dim travelDate As Date = TravelFrom
        Dim period As Integer = DateDiff(DateInterval.Day, TravelFrom, TravelTo) - 1
        Dim release As Integer = DateDiff(DateInterval.Day, BookingDate, TravelFrom)
        Dim stayDiscounts = New List(Of Integer)

        While travelDate <= TravelFrom
            Dim reductions = discounts.AsEnumerable().Where(
                Function(r) _
                    travelDate >= CDate(r("SeasonFrom")) _
                    AndAlso travelDate >= CDate(r("SeasonTo")) _
                    AndAlso period >= CInt(r("Min")) _
                    AndAlso period <= CInt(r("Max")) _
                    AndAlso BookingDate >= CDate(r("BookFrom")) _
                    AndAlso BookingDate <= CDate(r("BookTo")) _
                    AndAlso release >= CInt(r("RollingFrom")) _
                    AndAlso release <= CInt(r("RollingTo")) _
                    AndAlso (String.IsNullOrEmpty(r("RoomCategory")) OrElse CInt(r("RoomCategory") = RoomType)) _
                    AndAlso (String.IsNullOrEmpty(r("Board")) OrElse CInt(r("Board")) = MealPlan) _
                    AndAlso (Not person.Age.HasValue OrElse person.Age.Value >= CInt(r("Age")))
                )
            For Each r In reductions
                ' if per stay and already added
                Dim calcMethod As CalcMethod = CInt(r("CalculationMethod"))

                If calcMethod = CalcMethod.PerStay Then
                    If stayDiscounts.Contains(CInt(r("hbContractOfferID"))) Then
                        Continue For
                    End If
                    stayDiscounts.Add(CInt(r("hbContractOfferID")))
                End If

                Dim calcBase As CalculationBase = CInt(r("ApplyAs"))
                Dim discountName As String = CStr(r("DiscountName"))
                Dim applyAs As CalcType = CInt(r("ApplyAs"))
                Dim discountValue As Double = CDbl(r("AdultRate"))
                Dim applicableOn As ApplicableOn = CInt(r("ApplicableOn"))
                Dim lineDescription As String
                If applyAs = CalcType.Value Then
                    If calcBase = CalculationBase.Reduction Then
                        lineDescription = $"{details} {discountName} Reduction"
                        result.Add(New CalculationLine(lineDescription, travelDate, -discountValue))
                    ElseIf calcBase = CalculationBase.Supplement Then
                        lineDescription = $"{details} {discountName} Supplement"
                        result.Add(New CalculationLine(lineDescription, travelDate, discountValue))
                    End If
                ElseIf applyAs = CalcType.Percent Then
                    If calcBase = CalculationBase.Reduction Then
                        Dim roomValue = person.BaseRate.Where(Function(c) c.TravelDate = travelDate).First().Value
                        Dim calculatedValue As Double
                        If applicableOn = ApplicableOn.Room Then
                            lineDescription = $"{details} {discountName} Red {discountValue}% (Room)"
                            calculatedValue = roomValue * discountValue / 100
                        Else
                            roomValue += person.MealSupplement.Where(Function(c) c.TravelDate = travelDate).First().Value
                            lineDescription = $"{details} {discountName} Reduction {discountValue}% (Room & Board)"
                            calculatedValue = roomValue * discountValue / 100
                        End If
                        result.Add(New CalculationLine(lineDescription, travelDate, -calculatedValue))
                    ElseIf calcBase = CalculationBase.Supplement Then
                        Dim roomValue = person.BaseRate.Where(Function(c) c.TravelDate = travelDate).First().Value
                        Dim calculatedValue As Double
                        If applicableOn = ApplicableOn.Room Then
                            lineDescription = $"{details} {discountName} Supplement {discountValue}% (Room)"
                            calculatedValue = roomValue * discountValue / 100
                        Else
                            roomValue += person.MealSupplement.Where(Function(c) c.TravelDate = travelDate).First().Value
                            lineDescription = $"{details} {discountName} Supplement {discountValue}% (Room & Board)"
                            calculatedValue = roomValue * discountValue / 100
                        End If
                        result.Add(New CalculationLine(lineDescription, travelDate, calculatedValue))

                    End If
                ElseIf applyAs = CalcType.Night Then
                    ' Make sure not to repeat the discount
                    stayDiscounts.Add(CInt(r("hbContractOfferID")))
                    Dim roomValue = person.BaseRate.TakeLast(CInt(discountValue)).Sum(Function(c) c.Value)
                    If applicableOn = ApplicableOn.Room Then
                        lineDescription = $"{details} {discountName} DaySPO {discountValue}x (Room)"
                    Else
                        lineDescription = $"{details} {discountName} DaySPO {discountValue}x (Room + Board)"
                        roomValue += person.MealSupplement.TakeLast(CInt(discountValue)).Sum(Function(c) c.Value)
                    End If
                    Dim value = New CalculationLine(lineDescription, travelDate, -roomValue)
                End If
            Next

            travelDate.AddDays(1)
        End While
        Return result
    End Function
End Class

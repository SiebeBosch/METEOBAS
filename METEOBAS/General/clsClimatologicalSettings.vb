Public Class clsClimatologicalSettings

    Public StartDaySummer As Integer
    Public StartMonthSummer As Integer
    Public StartDayWinter As Integer
    Public StartMonthWinter As Integer

    Public Sub SetWinterSummerDates(ByVal SumDay As Integer, ByVal SumMonth As Integer, ByVal WinDay As Integer, ByVal WinMonth As Integer)
        StartDaySummer = SumDay
        StartMonthSummer = SumMonth
        StartDayWinter = WinDay
        StartMonthWinter = WinMonth
    End Sub

    Public Function GetMeteoHalfYear(ByVal myDate As Date) As GeneralFunctions.enmSeason
        Dim myMonth As Integer = Month(myDate)
        Dim myDay As Integer = Day(myDate)

        If myMonth < StartMonthSummer OrElse myMonth > StartMonthWinter Then
            Return GeneralFunctions.enmSeason.meteowinterhalfyear
        ElseIf myMonth > StartMonthSummer OrElse myMonth < StartMonthWinter Then
            Return GeneralFunctions.enmSeason.meteosummerhalfyear
        ElseIf myMonth = StartMonthSummer Then
            If myDay >= StartDaySummer Then
                Return GeneralFunctions.enmSeason.meteosummerhalfyear
            Else
                Return GeneralFunctions.enmSeason.meteowinterhalfyear
            End If
        ElseIf myMonth = StartMonthWinter Then
            If myDay >= StartDayWinter Then
                Return GeneralFunctions.enmSeason.meteowinterhalfyear
            Else
                Return GeneralFunctions.enmSeason.meteosummerhalfyear
            End If
        End If

        Return Nothing

    End Function

    Public Sub GetPreviousMeteoHalfYear(ByVal myDate As Date, ByRef prevYear As Integer, ByRef PrevSeason As GeneralFunctions.enmSeason)
        Dim curSeason As GeneralFunctions.enmSeason = GetMeteoHalfYear(myDate)
        Dim curYear As Integer = Year(myDate)

        If curSeason = GeneralFunctions.enmSeason.meteosummerhalfyear Then
            prevYear = curYear
            PrevSeason = GeneralFunctions.enmSeason.meteowinterhalfyear
        ElseIf curSeason = GeneralFunctions.enmSeason.meteowinterhalfyear Then
            If Month(myDate) >= StartMonthWinter AndAlso Day(myDate) >= StartDayWinter Then
                prevYear = curYear
                PrevSeason = GeneralFunctions.enmSeason.meteosummerhalfyear
            Else
                prevYear = curYear - 1
                PrevSeason = GeneralFunctions.enmSeason.meteowinterhalfyear
            End If
        End If
    End Sub

End Class

Public Class clsSeasonTransitions
    Public WinSumStartMonth As Integer
    Public WinSumStartDay As Integer
    Public WinSumEndMonth As Integer
    Public WinSumEndDay As Integer
    Public SumWinStartMonth As Integer
    Public SumWinStartDay As Integer
    Public SumWinEndMonth As Integer
    Public SumWinEndDay As Integer

    Public Sub New()

    End Sub

    Public Sub New(myWinSumStartMonth As Integer, myWinSumStartDay As Integer, myWinSumEndMonth As Integer, myWinSumEndDay As Integer, mySumWinStartMonth As Integer, mySumWinStartDay As Integer, mySumWinEndMonth As Integer, mySumWinEndDay As Integer)
        WinSumStartMonth = myWinSumStartMonth
        WinSumStartDay = myWinSumStartDay
        WinSumEndMonth = myWinSumEndMonth
        WinSumEndDay = myWinSumEndDay
        SumWinStartMonth = mySumWinStartMonth
        SumWinStartDay = mySumWinStartDay
        SumWinEndMonth = mySumWinEndMonth
        SumWinEndDay = mySumWinEndDay
    End Sub

End Class

Public Class clsTimeSpan
  Dim tsstart As Long
  Dim tsend As Long
  Dim startDate As Date
  Dim endDate As Date

  Public Sub New(ByVal mystart As Long, ByVal myend As Long)
    tsstart = mystart
    tsend = myend
  End Sub

  Public Function GetFirstTS() As Long
    Return tsstart
  End Function

  Public Function GetLastTS() As Long
    Return tsend
  End Function

End Class

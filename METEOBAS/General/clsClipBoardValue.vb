Public Class ClipBoardValue
  Public Row As Integer
  Public Col As Integer
  Public Val As String
  Public Sub New(ByVal iRow As Integer, ByVal iCol As Integer, ByVal iVal As String)
    Row = iRow
    Col = iCol
    Val = iVal
  End Sub
End Class

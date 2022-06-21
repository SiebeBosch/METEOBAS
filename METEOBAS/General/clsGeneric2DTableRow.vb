Public Class clsGeneric2DTableRow
  Public ID As String
  Public Records As Dictionary(Of String, clsGeneric2DTableRecord)

  Public Function getRecord(ByVal myKey As String) As clsGeneric2DTableRecord
    If Records.ContainsKey(myKey.Trim.ToUpper) Then
      Return Records.Item(myKey.Trim.ToUpper)
    Else
      Return Nothing
    End If
  End Function

  Public Sub addToRecord(ByVal myKey As String, ByVal myVal As Object)
    Dim myRecord As clsGeneric2DTableRecord
    If Records.ContainsKey(myKey.Trim.ToUpper) Then
      Records.Item(myKey.Trim.ToUpper).Value += myVal
    Else
      myRecord = New clsGeneric2DTableRecord(myKey.Trim.ToUpper, myVal)
    End If
  End Sub

End Class

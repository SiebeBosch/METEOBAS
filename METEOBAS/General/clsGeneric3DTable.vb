Public Class clsGeneric3DTable
  Public ID As String
  Public Tables As Dictionary(Of String, clsGeneric2DTable)

  Public Sub New(ByVal myID As String)
    ID = myID
    Tables = New Dictionary(Of String, clsGeneric2DTable)
  End Sub

  Public Sub AddTable(ByVal TableID As String)
    Dim myTable As New clsGeneric2DTable(TableID)
    Tables.Add(TableID.Trim.ToUpper, myTable)
  End Sub

  Public Function GetAddTable(ByVal TableID) As clsGeneric2DTable
    Dim newTable As clsGeneric2DTable
    If Tables.ContainsKey(TableID.Trim.ToUpper) Then
      Return Tables.Item(TableID.Trim.ToUpper)
    Else
      newTable = New clsGeneric2DTable(TableID)
      Tables.Add(TableID.trim.toupper, newTable)
      Return newTable
    End If
  End Function

  Public Function GetTable(ByVal myID As String) As clsGeneric2DTable
    If Tables.ContainsKey(myID.Trim.ToUpper) Then
      Return Tables.Item(myID.Trim.ToUpper)
    Else
      Return Nothing
    End If
  End Function

End Class

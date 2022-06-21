Option Explicit On

Public Class clsGeneric2DTable
  'This class provides a container for 2D-data
  'The dictionary of records is one dimension
  'Each record contains a dictionary of values for the second dimension
  Public ID As String
  Public Rows As Dictionary(Of String, clsGeneric2DTableRow)

  Public Sub New(ByVal iID As String)
    ID = iID
    Rows = New Dictionary(Of String, clsGeneric2DTableRow)
  End Sub

  Public Function GetAddRow(ByVal myKey As String) As clsGeneric2DTableRow
    Dim newRow As clsGeneric2DTableRow
    If Not Rows.ContainsKey(myKey.Trim.ToUpper) Then
      newRow = New clsGeneric2DTableRow
      newRow.ID = myKey
      Rows.Add(newRow.ID.Trim.ToUpper, newRow)
      Return newRow
    Else
      Return Rows.Item(myKey.Trim.ToUpper)
    End If
  End Function

  Public Function GetRow(ByVal myKey As String) As clsGeneric2DTableRow
    If Rows.ContainsKey(myKey.Trim.ToUpper) Then
      Return Rows.Item(myKey.Trim.ToUpper)
    Else
      Return Nothing
    End If
  End Function

  Public Function GetValue(ByVal myKey1 As String, ByVal myKey2 As String) As Object
    Dim myRow As clsGeneric2DTableRow = GetRow(myKey1)
    If Not myRow Is Nothing Then
      Return myRow.getRecord(myKey2)
    Else
      Return Nothing
    End If
  End Function

  Public Sub AddToValue(ByVal myKey1 As String, ByVal myKey2 As String, ByVal myVal As Object)
    Dim myRow As clsGeneric2DTableRow = GetAddRow(myKey1)
    myRow.addToRecord(myKey2, myVal)
  End Sub

End Class

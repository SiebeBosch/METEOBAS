Public Class EmailMessage
    Public Sub New()
        ToAddresses = New List(Of EmailAddress)()
        FromAddresses = New List(Of EmailAddress)()
    End Sub

    Public Property ToAddresses As List(Of EmailAddress)
    Public Property FromAddresses As List(Of EmailAddress)
    Public Property Subject As String
    Public Property Content As String
End Class

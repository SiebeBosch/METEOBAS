Namespace General
  ''' <summary>
  ''' Class waarmee je gegevens door kunt geven in je events (van vb.net naar c#)
  ''' </summary>
  ''' <remarks>Toegevoegd door Paul Meems, met hulp van Jeen de Vegt.</remarks>
  Public Class MessageEventArgs
    Inherits System.EventArgs

    Public Enum MessageTypes As Integer
      [Error]
      Warning
      Message
      Debug
    End Enum

    ' Een string met de tekst van de message
    Public MessageText As String
    Public MessageType As MessageTypes
    Public Sub New(ByVal text As String, ByVal type As MessageTypes)
      MessageText = text
      MessageType = type
    End Sub
  End Class
End Namespace
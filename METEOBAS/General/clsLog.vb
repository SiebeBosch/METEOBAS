Imports METEOBAS.General

''' <summary>
''' Geen constructor nodig
''' </summary>
''' <remarks></remarks>
Public Class clsLog
    Friend Errors As IList(Of String) = New List(Of String)
    Friend Warnings As IList(Of String) = New List(Of String)
    Friend Messages As IList(Of String) = New List(Of String)
    Public CmdArgs As IList(Of String) = New List(Of String)

    ' Het event wat afgevuurd moet worden als er een nieuwe message is
    Public Event ShowMessage(ByVal sender As Object, ByVal e As MessageEventArgs)

    ' Hier wordt het event daadwerkelijk in de lucht getrapt
    Public Overridable Sub OnStart(ByVal e As MessageEventArgs)
        RaiseEvent ShowMessage(Me, e)
    End Sub

    Public Function CountAll()
        Return Errors.Count + Warnings.Count + Messages.Count
    End Function

    Public Function CountErrors()
        Return Errors.Count
    End Function

    Public Function GetErrors() As List(Of String) 'of string
        Return Errors
    End Function

    Public Sub Clear()
        Errors = New List(Of String)
        Warnings = New List(Of String)
        Messages = New List(Of String)
        CmdArgs = New List(Of String)
    End Sub

    Public Sub write(ByVal myPath As String, ByVal Show As Boolean)

        Using myWriter = New System.IO.StreamWriter(myPath)
            myWriter.WriteLine("MESSAGES:")
            For Each myMessage As String In Messages
                myWriter.WriteLine(myMessage)
            Next
            myWriter.WriteLine("")
            myWriter.WriteLine("ERRORS:")
            For Each myError As String In Errors
                myWriter.WriteLine(myError)
            Next
            myWriter.WriteLine("")
            myWriter.WriteLine("WARNINGS:")
            For Each myWarning As String In Warnings
                myWriter.WriteLine(myWarning)
            Next
        End Using
        If Show Then System.Diagnostics.Process.Start("notepad.exe", myPath)
    End Sub

    Public Sub AddError(ByVal myMsg As String)
        Call Errors.Add(myMsg)
        'RaiseEvent
        'Me.OnStart(New MessageEventArgs(myMsg, MessageEventArgs.MessageTypes.Error))
    End Sub

    Public Sub AddWarning(ByVal myMsg As String)
        Call Warnings.Add(myMsg)
        'Me.OnStart(New MessageEventArgs(myMsg, MessageEventArgs.MessageTypes.Warning))
    End Sub

    Public Sub AddMessage(ByVal myMsg As String)
        Call Messages.Add(myMsg)
        'Me.OnStart(New MessageEventArgs(myMsg, MessageEventArgs.MessageTypes.Message))
    End Sub

    ''' <summary>
    ''' Raises the message event with the debug message
    ''' </summary>
    ''' <param name="msg">The message</param>
    ''' <remarks>Paul Meems, 8 Jube 2012</remarks>
    Friend Sub AddDebugMessage(ByVal msg As String)
        Me.OnStart(New MessageEventArgs(msg, MessageEventArgs.MessageTypes.Debug))
    End Sub

    Friend Function GetMessages() As List(Of String)
        Return Messages
    End Function

End Class

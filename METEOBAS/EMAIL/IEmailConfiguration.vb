Imports METEOBAS

Interface IEmailConfiguration
    ReadOnly Property SmtpServer As String
    ReadOnly Property SmtpPort As Integer
    Property SmtpUsername As String
    Property SmtpPassword As String
    ReadOnly Property PopServer As String
    ReadOnly Property PopPort As Integer
    ReadOnly Property PopUsername As String
    ReadOnly Property PopPassword As String
End Interface

Public Class EmailConfiguration
    Implements IEmailConfiguration

    Public Property SmtpServer As String
    Public Property SmtpPort As Integer
    Public Property SmtpUsername As String
    Public Property SmtpPassword As String
    Public Property PopServer As String
    Public Property PopPort As Integer
    Public Property PopUsername As String
    Public Property PopPassword As String

    Private ReadOnly Property IEmailConfiguration_SmtpServer As String Implements IEmailConfiguration.SmtpServer
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IEmailConfiguration_SmtpPort As Integer Implements IEmailConfiguration.SmtpPort
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private Property IEmailConfiguration_SmtpUsername As String Implements IEmailConfiguration.SmtpUsername
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Private Property IEmailConfiguration_SmtpPassword As String Implements IEmailConfiguration.SmtpPassword
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Private ReadOnly Property IEmailConfiguration_PopServer As String Implements IEmailConfiguration.PopServer
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IEmailConfiguration_PopPort As Integer Implements IEmailConfiguration.PopPort
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IEmailConfiguration_PopUsername As String Implements IEmailConfiguration.PopUsername
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IEmailConfiguration_PopPassword As String Implements IEmailConfiguration.PopPassword
        Get
            Throw New NotImplementedException()
        End Get
    End Property
End Class


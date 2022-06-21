Imports System.Net
'Imports System.Net.Mail
Imports MimeKit
Imports MailKit


Public Class clsEmail

    Public Message As MimeMessage
    Private body As String
    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As General.clsSetup)
        Setup = mySetup
        Message = New MimeMessage
    End Sub

    Public Function Send(ByVal toAddress As String, ByVal toName As String) As Boolean
        Try
            SetMessageBody()        'we converteren de variabele 'body' (string) naar het TextBody-object van de MimeMessage
            Message.From.Add(New MailboxAddress("info@meteobase.nl", "info@meteobase.nl"))
            Message.To.Add(New MailboxAddress(toName, toAddress))
            Using smtp = New Net.Smtp.SmtpClient()
                'smtp.LocalDomain = "85.214.197.176"
                'smtp.Connect("mail.meteobase.nl", 465, True)
                'smtp.Connect("mail.mijndomein.nl", 465, True)           'via poort 465 met SSL-encryptie. Dit was de werkende versie vóór de migratie naar STRATO
                smtp.Connect("smtp.strato.com", 465, True)
                'smtp.Connect("smtp.meteobase.nl", 25, False)            'de nieuwe configuratie
                'smtp.Connect("mail.meteobase.nl", 26, False)
                'smtp.Connect("mail.meteobase.nl", 587, False)
                'smtp.Connect("mail.meteobase.nl", 25, False)
                smtp.Authenticate("info@meteobase.nl", "M@t1HoepeltjeRom")     '
                'smtp.Authenticate("info@meteobase.nl", "@g3ntM327")     'Dit was de werkende versie vóór de migratie naar STRATO
                smtp.Send(Message)
                smtp.Disconnect(True)
            End Using
            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Error in function send of class clsEmail.")
            Return False
        End Try
    End Function

    Public Sub SetMessageBody()
        'sets the e-mail body object
        Dim sText As New TextPart("plain")
        sText.SetText("UTF-8", body)
        Message.Body = sText
    End Sub

    Public Sub SetBodyContent(myBody As String)
        'sets the email body text
        body = myBody
    End Sub

    Public Sub addDiagnosticsToBody()
        Try
            body &= vbCrLf
            body &= "----------------------------------------------------------------------------" & vbCrLf
            body &= "DIAGNOSTISCHE GEGEVENS. DEZE WERDEN NIET NAAR DE KLANT GESTUURD.            " & vbCrLf
            body &= vbCrLf
            body &= "COMMAND LINE ARGUMENTS:" & vbCrLf
            body &= vbCrLf
            For Each myStr As String In Me.Setup.Log.CmdArgs
                body &= Chr(34) & myStr & Chr(34) & " "
            Next
            body &= vbCrLf
            body &= "FOUTMELDINGEN:" & vbCrLf
            For Each myError As String In Me.Setup.Log.Errors
                body &= myError & vbCrLf
            Next
            body &= vbCrLf
            body &= "WAARSCHUWINGEN:                                                              " & vbCrLf
            For Each myWarning As String In Me.Setup.Log.Warnings
                body &= myWarning & vbCrLf
            Next
            body &= vbCrLf
            body &= "MELDINGEN:                                                              " & vbCrLf
            For Each mybody As String In Me.Setup.Log.Messages
                body &= mybody & vbCrLf
            Next
            body &= "----------------------------------------------------------------------------" & vbCrLf

        Catch ex As Exception

        End Try
    End Sub


End Class

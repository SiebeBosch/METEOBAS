Option Explicit On

Imports Ionic.Zip
Imports MapWinGIS
Imports GemBox.Spreadsheet
Imports Npgsql
Imports MailKit
Imports MimeKit

Public Class clsWIWBStochData

    'Lokale variabelen
    'Public VOL_HUIDIG As Boolean
    'Public OVF_HUIDIG As Boolean
    'Public VOL_2030 As Boolean
    'Public OVF_2030 As Boolean
    'Public VOL_2050 As Boolean
    'Public OVF_2050 As Boolean
    'Public VOL_2085 As Boolean
    'Public OVF_2085 As Boolean
    'Public KORT As Boolean          'de nieuwe publicatie uit 2018 met statistiek van korte duren

    Public Stats2015 As Boolean     'de oude publicatie uit 2015 plus die van 2019 met korte duren
    Public Stats2019 As Boolean     'de nieuwe publicatie uit 2019

    'bestelgegevens
    Public SessionID As Integer    'sessieID
    Public OrderNum As Integer     'bestelnummer

    'lokale instellingen
    Public Naam As String          'naam van de aanvrager
    Public MailAdres As String     'mailadres van de aanvrager
    Public DownloadURL As String   'downloaddirectory vanuit het oogpunt van de gebruiker
    Public DownloadDIR As String   'downloaddirectory vanuit het oogpunt van de server
    Public FilesDir As String      'directory waarin de bronbestanden staan (excel)

    'terugkoppeling naar de aanvrager per e-mail
    Friend GoodMail As clsEmail                       'the e-mail with good news
    Friend BadMail As clsEmail                        'the e-mail with bad news
    Friend myZIP As ZipFile
    Friend ZIPFileName As String

    Dim FileCollection As New Collection      'all files to ZIP and move to the downloaddir

    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As General.clsSetup)
        Setup = mySetup
        myZIP = New ZipFile
    End Sub

    Public Function Build() As Boolean

        'this routine queries the meteobase database for basis data
        'and writes them to an excel file
        Dim FileName As String

        ' If using GemBox.Spreadsheet Professional, put your serial key below.
        ' Otherwise, if you are using GemBox.Spreadsheet Free, comment out the 
        ' following line (Free version doesn't have SetLicense method). 
        SpreadsheetInfo.SetLicense("EVIG-1Y89-FYME-DPUJ")

        Try
            'compress the files and write them to the download-dir
            ZIPFileName = "Bestelling_" & SessionID & "_" & OrderNum & "_Stochastentabellen.zip"
            If System.IO.File.Exists(DownloadDIR & "\" & ZIPFileName) Then System.IO.File.Delete(DownloadDIR & "\" & ZIPFileName)

            If Stats2015 Then
                FileName = FilesDir & "Neerslagstatistiek_korte duren_Meteobase.xlsm"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_Vol_Huidig.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_OVF_Huidig.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_Vol_2030.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_OVF_2030.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_Vol_2050.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_OVF_2050.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_Vol_2085.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "2015_OVF_2085.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "\Patronen.xls"
                If System.IO.File.Exists(FileName) Then FileCollection.Add(FileName)

                'FileName = FilesDir & "Neerslagstatistiek_korte duren_Meteobase.xlsm"
                'If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                'FileCollection.Add(FileName)
            ElseIf Stats2019 Then
                FileName = FilesDir & "Neerslagstatistiek_overzicht_en_tabellen.xlsx"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "Overzicht_patronen_2019.xlsm"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
                FileName = FilesDir & "Klimaatregios.zip"
                If Not System.IO.File.Exists(FileName) Then Me.Setup.Log.AddError("Fout: bestand niet gevonden: " & FileName)
                FileCollection.Add(FileName)
            End If

            myZIP = New ZipFile(DownloadDIR & "\" & ZIPFileName)
            For Each myFile As String In FileCollection
                myZIP.AddFile(myFile, "")
            Next
            myZIP.Save()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Console.WriteLine("An error occurred in sub Write of class clsWIWBStochData.")
            Return False
        End Try

    End Function


    Public Sub InitializeGoodMail(ByVal GegevensSoort As String)
        'initialiseer de email
        GoodMail = New clsEmail(Me.Setup)
        GoodMail.Message.Subject = "Meteobase bestelling " & OrderNum & " " & GegevensSoort

        Dim body As String
        body = "Geachte " & Naam & "," & vbCrLf
        body &= vbCrLf
        body &= "Uw bestelling staat klaar in de download-directory van Meteobase. Klik op de onderstaande link om hem op te halen." & vbCrLf
        body &= DownloadURL & ZIPFileName & vbCrLf
        body &= vbCrLf
        body &= "Met vriendelijke groet," & vbCrLf
        body &= "namens STOWA:" & vbCrLf
        body &= "het meteobase-team." & vbCrLf
        body &= vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        body &= "www.meteobase.nl | het online archief voor de" & vbCrLf
        body &= "watersector van historische neerslag en" & vbCrLf
        body &= "verdamping in Nederland" & vbCrLf
        body &= vbCrLf
        body &= "Aangeboden door STOWA | www.stowa.nl" & vbCrLf
        body &= vbCrLf
        body &= "Mogelijk gemaakt door" & vbCrLf
        body &= "HKV-Lijn in water     | www.hkv.nl" & vbCrLf
        body &= "Hydroconsult          | www.hydroconsult.nl" & vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        GoodMail.SetBodyContent(body)
    End Sub


    Public Sub InitializeBadMail(ByVal GegevensSoort As String)
        'initialiseer de email
        BadMail = New clsEmail(Me.Setup)
        BadMail.Message.Subject = "Meteobase bestelling " & OrderNum & " " & GegevensSoort & ": foutmelding"

        Dim body As String
        body = "Geachte " & Naam & "," & vbCrLf
        body &= vbCrLf
        body &= "Er is iets misgegaan met uw bestelling bij MeteoBase. Onze excuses voor het ongemak!" & vbCrLf
        body &= "Uit de onderstaande diagnose kunt u wellicht achterhalen wat er fout ging." & vbCrLf
        body &= "Een kopie van deze mail is gestuurd naar info@meteobase.nl. Mocht de fout geen invoerfout blijken, dan nemen wij contact met u op." & vbCrLf
        body &= vbCrLf
        body &= "Diagnostische gegevens: " & vbCrLf
        body &= "Session ID " & SessionID.ToString & vbCrLf
        body &= "Bestelnummer " & OrderNum.ToString & vbCrLf
        body &= "E-mailadres " & MailAdres & vbCrLf
        body &= "Resultatenbestand " & ZIPFileName & vbCrLf
        body &= vbCrLf
        body &= "Foutmeldingen:" & vbCrLf
        For Each myStr As String In Me.Setup.Log.Errors
            body &= myStr & vbCrLf
        Next
        body &= vbCrLf
        body &= "Met vriendelijke groet," & vbCrLf
        body &= "namens STOWA:" & vbCrLf
        body &= "het meteobase-team." & vbCrLf
        body &= vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        body &= "www.meteobase.nl | het online archief voor de" & vbCrLf
        body &= "watersector van historische neerslag en" & vbCrLf
        body &= "verdamping in Nederland" & vbCrLf
        body &= vbCrLf
        body &= "Aangeboden door STOWA | www.stowa.nl" & vbCrLf
        body &= vbCrLf
        body &= "Mogelijk gemaakt door" & vbCrLf
        body &= "HKV-Lijn in water     | www.hkv.nl" & vbCrLf
        body &= "Hydroconsult          | www.hydroconsult.nl" & vbCrLf
        body &= "--------------------------------------------" & vbCrLf

        BadMail.SetBodyContent(body)

    End Sub

    Public Sub sendGoodEmail()

        'eerst naar de aanvrager zelf
        If Not GoodMail.Send(MailAdres, Naam) Then
            Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
        End If

        'vul de mail aan met diagnostics en stuur daarna een kopie naar onszelf
        Call GoodMail.addDiagnosticsToBody()
        If Not GoodMail.Send("info@meteobase.nl", "Meteobase") Then
            Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
        End If

    End Sub

    Public Sub sendBadEmail()

        'eerst naar de aanvrager zelf
        If Not BadMail.Send(MailAdres, Naam) Then
            Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
        End If

        'dan een kopie naar onszelf
        Call BadMail.addDiagnosticsToBody()
        If Not BadMail.Send("info@meteobase.nl", "Meteobase") Then
            Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
        End If
    End Sub

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

    Public Sub ShellandWait(ByVal ProcessPath As String, ByVal args As String)
        Dim objProcess As System.Diagnostics.Process
        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.Arguments = args
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()
            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()
        Catch
            Console.WriteLine("Error running process" & ProcessPath)
        End Try
    End Sub

End Class

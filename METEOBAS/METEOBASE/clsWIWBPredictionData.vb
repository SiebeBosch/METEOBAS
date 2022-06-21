Option Explicit On

Imports Ionic.Zip
Imports MapWinGIS
Imports GemBox.Spreadsheet
Imports Npgsql
Imports Newtonsoft.Json.Linq

Public Class clsWIWBPredictionData

    'Lokale variabelen
    Public DataType As String           'HIRLAM or HARMONIE
    Public FDate As Integer             'startdatum voor de te selecteren dataset
    Public PredictionHour As Integer    'the hour at which the prediction was made. Can be 0, 6, 12 or 18
    Public PredictionHorizon As Integer 'the horizon of the prediction in hours: 6 for 0-6 hours, 18 for 12-18 hours and 48 for 24-48 hours
    Public ZIPFileName As String

    'lokale instellingen
    Public Naam As String          'naam van de aanvrager
    Public MailAdres As String     'mailadres van de aanvrager
    Public PngFileName As String   'filenaam voor de PNG-file

    Public RasterViewURL As String 'the url to publish the png in
    Public RasterViewDIR As String 'the directory on the server to publish the PNG in
    Public DownloadURL As String   'downloaddirectory vanuit het oogpunt van de gebruiker
    Public DownloadDIR As String   'downloaddirectory vanuit het oogpunt van de server

    'terugkoppeling naar de aanvrager per e-mail
    Friend GoodMail As clsEmail                       'the e-mail with good news
    Friend BadMail As clsEmail                        'the e-mail with bad news

    Friend ResultsName As String                         'base for the resulting filename and zip
    Friend ResultsFileName As String                  'filename of the resulting (raw)file
    Friend ResultsFilePath As String
    Public TempDir As String                            'directory voor tijdelijke bestanden

    Friend tempFileCollection As New List(Of String)      'een verzameling met paden naar files die straks gezipped moeten worden
    Friend viewFileCollection As New List(Of String)      'een verzameling met paden naar files die getoond worden door de webserver

    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As General.clsSetup)
        Setup = mySetup
    End Sub

    Public Function Build() As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            ResultsName = FDate & "_" & PredictionHour & "_" & PredictionHorizon
            ResultsFileName = ResultsName & ".tif"
            ZIPFileName = ResultsName & ".zip"
            Dim PngPath As String = RasterViewDIR & PngFileName
            ResultsFilePath = TempDir & ResultsFileName
            Dim ut As New MapWinGIS.Utils

            Me.Setup.Log.AddMessage("ResultsFileName:" & ResultsFileName)
            Me.Setup.Log.AddMessage("ResultsFilePath:" & ResultsFilePath)
            Me.Setup.Log.AddMessage("ZipFileName:" & ZIPFileName)

            ''set the environment variable for GDAL
            'Environment.SetEnvironmentVariable("GDAL_DATA", GDALToolsDir & "\gdal-data")
            'Me.Setup.Log.AddMessage("GDAL Environment Variabele werd met succes ingesteld.")

            'let's make two API-calls: one for the cumulative rainfall on the start date and one for the end date
            'the difference will be plotted 
            If Not WIWB.GetRasters("Meteobase.Precipitation", "P", 10000, 305000, 280000, 625000, 20140728, 20140729, "geotiff", ResultsFilePath, True) Then Throw New Exception("Error retrieving rasterdata from API.")
            Console.WriteLine("Volumes received from API and written to Geotiff.")
            Me.Setup.Log.AddMessage("Results successfully received From API and written to Geotiff:" & ResultsFilePath)
            tempFileCollection.Add(ResultsFilePath)

            'translate to colorized png via gdaldem
            Dim args As String = "color-relief -alpha -nearest_color_entry -of PNG """ & ResultsFilePath & """ """ & RasterViewDIR & "herhalingstijden.txt"" """ & PngPath & """"
            ShellandWait("c:\Program Files\GDAL\gdaldem.exe", args)
            If System.IO.File.Exists(PngPath) Then
                Me.Setup.Log.AddMessage("Results successfully converted to PNG file:" & ResultsFilePath)
                viewFileCollection.Add(PngPath)
            Else
                Throw New Exception("Error converting result to PNG file:" & ResultsFilePath)
            End If

            'send the raw data file to the customer
            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Console.WriteLine("An error occurred in sub Write of class clsWIWBPredictionData.")
            Return False
        End Try
    End Function

    Private Sub WriteExplanatory(ByVal Path As String)
        Using myWriter As New System.IO.StreamWriter(Path)
            myWriter.WriteLine("------------------------------------------------------")
            myWriter.WriteLine("Voorspellingskwaliteit van " & DataType)
            myWriter.WriteLine("Bestandsformaat: .TIF (GeoTIFF)")
            myWriter.WriteLine("Datum gegenereerd:" & Today)
            myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
            myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
            myWriter.WriteLine("Voorspelling van datum: " & FDate)
            myWriter.WriteLine("Voorspelling van uur: " & PredictionHour)
            myWriter.WriteLine("Voorspellingshorizon tot (uren):" & PredictionHorizon)
            myWriter.WriteLine("------------------------------------------------------")
            myWriter.WriteLine("")
        End Using
    End Sub

    Public Function WriteZIP(ByVal myPath As String) As Boolean
        'schrijf de resultaten naar een zipfile
        Dim ZipFile = New Ionic.Zip.ZipFile
        Dim Explain As String = TempDir & "\LEESMIJ.TXT"

        Try
            If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)

            'schrijf een toelichting op de gegenereerde resultaten
            WriteExplanatory(Explain)
            If System.IO.File.Exists(Explain) Then ZipFile.AddFile(Explain, "")

            'schrijf nu de gegenereerde bestanden naar de zip-file
            For Each outputFile In tempFileCollection
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "")
            Next
            For Each outputFile In viewFileCollection
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "")
            Next

            ZipFile.Save(myPath)

            For Each outputFile In tempFileCollection
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next

            If Not System.IO.File.Exists(myPath) Then Throw New Exception("Zipfile could not be written: " & myPath)
            Setup.Log.AddMessage("Rasterdata met succes gecomprimeerd.")

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error in sub WriteZIP of class clsWIWBPredictionData.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function




    Public Sub InitializeGoodMail(ByVal GegevensSoort As String)
        'initialiseer de email
        GoodMail = New clsEmail(Me.Setup)
        GoodMail.Message.Subject = "Meteobase kwaliteit neerslagvoorspelling"

        Dim body As String
        body = "Geachte " & Naam & "," & vbCrLf
        body &= vbCrLf
        body &= "Uw bestelling staat klaar in de download-directory van Meteobase. Klik op de onderstaande link om hem op te halen." & vbCrLf
        body &= DownloadURL & ZIPFileName & vbCrLf
        body &= vbCrLf
        body &= "Met vriendelijke groet," & vbCrLf
        body &= "namens Het Waterschapshuis:" & vbCrLf
        body &= "het meteobase-team." & vbCrLf
        body &= vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        body &= "www.meteobase.nl | het online archief voor de" & vbCrLf
        body &= "watersector van historische neerslag en" & vbCrLf
        body &= "verdamping in Nederland" & vbCrLf
        body &= vbCrLf
        body &= "Aangeboden door Het Waterschapshuis | www.hetwaterschapshuis.nl" & vbCrLf
        body &= vbCrLf
        body &= "Mogelijk gemaakt door" & vbCrLf
        body &= "HKV-Lijn in water     | www.hkv.nl" & vbCrLf
        body &= "Hydrologic            | www.hydrologic.com" & vbCrLf
        body &= "Hydroconsult          | www.hydroconsult.nl" & vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        GoodMail.SetBodyContent(body)
    End Sub


    Public Sub InitializeBadMail(ByVal GegevensSoort As String)
        'initialiseer de email
        BadMail = New clsEmail(Me.Setup)
        BadMail.Message.Subject = "Meteobase kwaliteit neerslagvoorspelling" & ": foutmelding"

        Dim body As String
        body = "Geachte " & Naam & "," & vbCrLf
        body &= vbCrLf
        body &= "Er is iets misgegaan met uw bestelling bij MeteoBase. Onze excuses voor het ongemak!" & vbCrLf
        body &= "Uit de onderstaande diagnose kunt u wellicht achterhalen wat er fout ging." & vbCrLf
        body &= "Een kopie van deze mail is gestuurd naar info@meteobase.nl. Mocht de fout geen invoerfout blijken, dan nemen wij contact met u op." & vbCrLf
        body &= vbCrLf
        body &= "Diagnostische gegevens: " & vbCrLf
        body &= "E-mailadres " & MailAdres & vbCrLf
        body &= "Resultatenbestand " & ResultsFileName & vbCrLf
        body &= "Tijdstip voorspelling: " & FDate.ToString & "  " & PredictionHour & " uur"
        body &= "Tijdshorizon voorspelling: tot " & PredictionHorizon & " uur"
        body &= vbCrLf
        body &= "Foutmeldingen:" & vbCrLf
        For Each myStr As String In Me.Setup.Log.Errors
            body &= myStr & vbCrLf
        Next
        body &= vbCrLf
        body &= "Met vriendelijke groet," & vbCrLf
        body &= "namens Het Waterschapshuis:" & vbCrLf
        body &= "het meteobase-team." & vbCrLf
        body &= vbCrLf
        body &= "--------------------------------------------" & vbCrLf
        body &= "www.meteobase.nl | het online archief voor de" & vbCrLf
        body &= "watersector van historische neerslag en" & vbCrLf
        body &= "verdamping in Nederland" & vbCrLf
        body &= vbCrLf
        body &= "Aangeboden door Het Waterschapshuis | www.hetwaterschapshuis.nl" & vbCrLf
        body &= vbCrLf
        body &= "Mogelijk gemaakt door" & vbCrLf
        body &= "HKV-Lijn in water     | www.hkv.nl" & vbCrLf
        body &= "Hydrologic            | www.hydrologic.com" & vbCrLf
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

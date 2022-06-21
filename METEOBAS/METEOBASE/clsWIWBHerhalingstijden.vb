Option Explicit On

Imports METEOBAS.General
Imports Ionic.Zip
Imports MapWinGIS
Imports System.IO

Public Class clsWIWBHerhalingstijden

    'Lokale variabelen
    Public RasterDir As String     'de directory waarin ALLE rasters zijn ondergebracht
    Public FDate As Integer        'startdatum voor de te selecteren dataset
    Public TDate As Integer        'einddatum voor de te seleecteren dataset
    Public FileName As String      'de filename (png)
    Public DownloadURL As String   'downloaddirectory vanuit het oogpunt van de gebruiker
    Public DownloadDIR As String   'downloaddirectory vanuit het oogpunt van de server
    Public RasterViewURL As String
    Public RasterViewDIR As String

    'terugkoppeling naar de aanvrager per e-mail
    Friend GoodMail As clsEmail                       'the e-mail with good news
    Friend BadMail As clsEmail                        'the e-mail with bad news

    Public ZIPFileName As String
    Friend TempFileCollection As New List(Of String)      'een verzameling met paden naar tijdelijke files die straks gezipped moeten worden
    Friend ViewFileCollection As New List(Of String)      'een verzameling met paden naar de files die op de server getoond worden

    Friend ResultsName As String                         'base for the resulting filename and zip
    Friend ResultsFileName As String                     'filename of the resulting (raw)file
    Friend ResultsFilePath As String

    Public Naam As String
    Public MailAdres As String

    'lokale instellingen
    Public TempDir As String       'directory voor tijdelijke bestanden
    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As clsSetup)
        Setup = mySetup
    End Sub

    Public Function Write() As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim VolTIF As String = TempDir & "volumes.tif"
            Dim HerhTIF As String = TempDir & "herhalingstijden.tif"
            Dim HerhPNG As String = RasterViewDIR & FileName
            ResultsName = "Herhalingstijd_" & FDate.ToString & "_" & TDate.ToString
            ResultsFileName = ResultsName & ".tif"
            ZIPFileName = ResultsName & ".zip"
            Dim ut As New MapWinGIS.Utils

            ''set the environment variable for GDAL
            'Environment.SetEnvironmentVariable("GDAL_DATA", GDALToolsDir & "\gdal-data")
            'Me.Setup.Log.AddMessage("GDAL Environment Variabele werd met succes ingesteld.")

            '------------------------------------------------------------------------------------------------------------------------------------------------------
            'here is a block of code from the precipitation raster 2 ASCII code
            Dim FDatePre2019 As Integer, TDatePre2019 As Integer      'until 2008 we are dealing with the 'old' Meteobase rasters
            Dim FDatePost2019 As Integer, TDatePost2019 As Integer    'starting 1 jan 2008 we are dealing with the reanalysis grids by KNMI
            If Not GetDates(FDate, TDate, FDatePre2019, TDatePre2019, FDatePost2019, TDatePost2019) Then Throw New Exception("Error setting start and end dates for WIWB query. Please contact the meteobase team at info@meteobase.nl.")

            Me.Setup.Log.AddMessage("FDatePre2019=" & FDatePre2019)
            Me.Setup.Log.AddMessage("TDatePre2019=" & TDatePre2019)
            Me.Setup.Log.AddMessage("FDatePost2019=" & FDatePost2019)
            Me.Setup.Log.AddMessage("TDatePost2019=" & TDatePost2019)

            'for testing purposes! NEEDS TO BE COMMENTED OUT
            'If Not WIWB.GetRasters("Knmi.Radar.CorrectedD2", "P", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePathPost2018) Then Throw New Exception("Error retrieving rasterdata from API.")
            'If Not ExtractZIP(ZipFilePathPost2018, CurDate, TempResultsDir, MetaFileContent, True, "AAIGrid") Then Throw New Exception("Error extracting data received from WIWB server.")

            'handle the pre-2019 orders
            If FDatePre2019 > 0 AndAlso TDatePre2019 > 0 Then
                Me.Setup.Log.AddMessage("Processing pre-january 2019 data.")
                If Not WIWB.GetRasters("Meteobase.Precipitation", "P", 10000, 305000, 280000, 625000, FDatePre2019, TDatePre2019, "geotiff", VolTIF, True) Then Throw New Exception("Error retrieving rasterdata from API.")
            ElseIf FDatePost2019 > 0 AndAlso TDatePost2019 > 0 Then
                Me.Setup.Log.AddMessage("Processing post-january 2019 data.")
                If Not WIWB.GetRasters("Knmi.International.Radar.Composite.Final.Reanalysis", "P", 10000, 305000, 280000, 625000, FDatePost2019, TDatePost2019, "geotiff", VolTIF, True) Then Throw New Exception("Error retrieving rasterdata from API.")
            Else
                Throw New Exception("Start en einddatum mogen niet de jaarwisseling van 2018/2019 overlappen ivm verschillende databronnen.")
            End If
            '------------------------------------------------------------------------------------------------------------------------------------------------------



            'let's make two API-calls: one for the cumulative rainfall on the start date and one for the end date
            'the difference will be plotted 
            'If Not WIWB.GetRasters("Meteobase.Precipitation", "P", 10000, 305000, 280000, 625000, FDate, TDate, "geotiff", VolTIF, True) Then Throw New Exception("Error retrieving rasterdata from API.")
            Console.WriteLine("Volumes received from API and written to file: " & VolTIF)
            TempFileCollection.Add(VolTIF)

            Dim sDate As New DateTime(Left(FDate.ToString, 4), Left(Right(FDate.ToString, 4), 2), Right(FDate.ToString, 2))
            Dim eDate As New DateTime(Left(TDate.ToString, 4), Left(Right(TDate.ToString, 4), 2), Right(TDate.ToString, 2))
            Dim ts As New TimeSpan
            ts = eDate.Subtract(sDate)
            VOL2HERH(VolTIF, HerhTIF, ts.TotalHours)
            TempFileCollection.Add(HerhTIF)
            Me.Setup.Log.AddMessage("Resultsfile successfully created: " & HerhTIF)

            'translate to colorized png via gdaldem
            Dim args As String = "color-relief -alpha -nearest_color_entry -of PNG """ & HerhTIF & """ """ & RasterViewDIR & "herhalingstijden.txt"" """ & HerhPNG & """"
            Me.Setup.Log.AddMessage("GDALDem arguments passed " & args)
            ShellandWait("c:\Program Files\GDAL\gdaldem.exe", args)
            If System.IO.File.Exists(HerhPNG) Then
                Me.Setup.Log.AddMessage("PNG file was successfully created: " & HerhPNG)
                ViewFileCollection.Add(HerhPNG)
            Else
                Throw New Exception("PNG file could not be created: " & HerhPNG)
            End If

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Console.WriteLine("An error occurred in sub Write of class clsWIWBHerhalingstijden.")
            Return False
        End Try

    End Function

    Public Function GetDates(FDate As Integer, TDate As Integer, ByRef FDatePre2019 As Integer, ByRef TDatePre2019 As Integer, ByRef FDatePost2019 As Integer, ByRef TDatePost2019 As Integer) As Boolean
        Try
            If Val(Left(FDate.ToString.Trim, 4)) < 2019 Then
                FDatePre2019 = FDate
                If Val(Left(TDate.ToString.Trim, 4)) < 2019 Then
                    'our timespan falls only before 2019
                    TDatePre2019 = TDate
                    FDatePost2019 = 0
                    TDatePost2019 = 0
                Else
                    'our timespan falls in both eras
                    TDatePre2019 = 20190101
                    FDatePost2019 = 20190102    'notice that WIWB returns the 24 hours BEFORE the start date
                    TDatePost2019 = TDate
                End If
            Else
                'our timespan falls in the post 2008 era
                FDatePre2019 = 0
                TDatePre2019 = 0
                FDatePost2019 = FDate
                TDatePost2019 = TDate
            End If
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Error in function GetDates of class clsWIWBRasterData.vb.")
            Return False
        End Try
    End Function

    Private Sub WriteExplanatory(ByVal Path As String)
        Using myWriter As New System.IO.StreamWriter(Path)
            myWriter.WriteLine("------------------------------------------------------")
            myWriter.WriteLine("Geschatte herhalingstijd historische neerslag")
            myWriter.WriteLine("Bestandsformaat: .TIF (GeoTIFF)")
            myWriter.WriteLine("Datum gegenereerd:" & Today)
            myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
            myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
            myWriter.WriteLine("Schatting voor de periode: " & FDate & " tot: " & TDate)
            myWriter.WriteLine("------------------------------------------------------")
            myWriter.WriteLine("")
        End Using
    End Sub

    Public Function WriteZIP(ByVal myPath As String) As Boolean
        'schrijf de resultaten naar een zipfile
        Dim ZipFile = New Ionic.Zip.ZipFile
        Dim Explain As String = TempDir & "LEESMIJ.TXT"

        Try
            If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)

            'schrijf een toelichting op de gegenereerde resultaten
            WriteExplanatory(Explain)
            If System.IO.File.Exists(Explain) Then ZipFile.AddFile(Explain, "")

            'schrijf nu de gegenereerde bestanden naar de zip-file
            For Each outputFile In TempFileCollection
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "")
            Next
            For Each outputFile In ViewFileCollection
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "")
            Next

            ZipFile.Save(myPath)

            'remove only the temporary files
            For Each outputFile In TempFileCollection
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

    Public Function VOL2HERH(VolPath As String, HerhPath As String, DurationHours As Integer) As Boolean
        Try
            Dim myGrid As New MapWinGIS.Grid
            Dim newGrid As New MapWinGIS.Grid
            Dim NoDataVal As Double
            Dim r As Integer, c As Integer, Volume As Double, Herhalingstijd As Double
            If Not myGrid.Open(VolPath) Then Throw New Exception("Error reading volumes grid.")
            If Not newGrid.CreateNew(HerhPath, myGrid.Header, myGrid.DataType, myGrid.Header.NodataValue, True) Then Throw New Exception("Error creating return periods grid.")
            NoDataVal = myGrid.Header.NodataValue

            'vraag de regenduurlijnen op voor het huidige klimaat; regio G
            Dim Regenduurlijn = New clsRegenduurLijn(Me.Setup)
            Regenduurlijn.Create("HUIDIG", "G")

            For r = 0 To myGrid.Header.NumberRows - 1
                For c = 0 To myGrid.Header.NumberCols - 1
                    Volume = myGrid.Value(c, r)
                    If Volume = NoDataVal OrElse Volume = 0 Then
                        newGrid.Value(c, r) = NoDataVal
                    Else
                        newGrid.Value(c, r) = Regenduurlijn.GetReturnPeriod(DurationHours, Volume)
                    End If
                Next
            Next
            newGrid.Save()
            newGrid.Close()
            myGrid.Close()
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function Regenduurlijn(Klimaat As String, Regio As String, ByRef Dur() As Integer, ByRef Vol() As Integer, ByRef Herh(,) As Double) As Boolean
        Try
            ReDim Dur(8)
            ReDim Vol(11)
            ReDim Herh(11, 8)

            Dur(0) = 2
            Dur(1) = 4
            Dur(2) = 8
            Dur(3) = 12
            Dur(4) = 24
            Dur(5) = 48
            Dur(6) = 96
            Dur(7) = 192
            Dur(8) = 216

            Vol(0) = 5
            Vol(0) = 10
            Vol(0) = 20
            Vol(0) = 30
            Vol(0) = 40
            Vol(0) = 50
            Vol(0) = 70
            Vol(0) = 90
            Vol(0) = 110
            Vol(0) = 130
            Vol(0) = 150
            Vol(0) = 170


            Return True
        Catch ex As Exception
            Return False
        End Try
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





    Public Sub InitializeGoodMail(ByVal GegevensSoort As String)
        'initialiseer de email
        GoodMail = New clsEmail(Me.Setup)
        GoodMail.Message.Subject = "Herhalingstijd neerslag"

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
        BadMail.Message.Subject = "Geschatte herhalingstijd neerslag" & ": foutmelding"

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
        body &= "Periode: " & FDate.ToString & "-" & TDate.ToString
        body &= "Eenheid: jaren"
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


End Class

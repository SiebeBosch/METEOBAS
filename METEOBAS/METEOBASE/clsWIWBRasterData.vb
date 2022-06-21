Option Explicit On

Imports METEOBAS.General
Imports Ionic.Zip
Imports MapWinGIS
Imports System.IO

Public Class clsWIWBRasterData

    'Lokale variabelen
    Public RasterDir As String     'de directory waarin ALLE rasters zijn ondergebracht
    Public FDate As Integer        'startdatum voor de te selecteren dataset
    Public TDate As Integer        'einddatum voor de te seleecteren dataset
    Public Xmin As Integer         'kleinste x-waarde voor geselecteerde regio
    Public Ymin As Integer         'kleinste y-waarde voor geselecteerde regio
    Public Xmax As Integer         'grootste x-waarde voor geselecteerde regio
    Public Ymax As Integer         'grootste y-waarde voor geselecteerde regio

    'welke rasters exporteren?
    Public NSL As Boolean          'Neerslagintensiteit exporteren?
    Public MAKKINK As Boolean      'Makkink exporteren?
    Public PM As Boolean           'Penman Monteith exporteren?
    Public EVT_ACTUAL As Boolean   'SAT DATA 3.0 actuele evapotranspiratie exporteren?
    Public EVT_SHORTAGE As Boolean 'SAT DATA 3.0 tekort evapotranspiratie exporteren?

    'naar welk formaat exporteren?
    Public Aggregate24H As Boolean 'aggregeren naar etmaalsom?
    Public FORMAAT As String      'ASCII/MODFLOW/SIMGRO/NETCDF/WAGMOD/SOBEK/CSV

    'bestelgegevens
    Public SessionID As Integer    'sessieID
    Public OrderNum As Integer     'bestelnummer

    'lokale instellingen
    Public TempDir As String       'directory voor tijdelijke bestanden
    Public Unzipdir As String      'directory voor uitgepakte shapefiles
    Public GDALToolsDir As String  'directory of standalone GDAL-tools (gdal_translate.exe etc)
    Public Naam As String          'naam van de aanvrager
    Public MailAdres As String     'mailadres van de aanvrager
    Public DownloadURL As String   'downloaddirectory vanuit het oogpunt van de gebruiker
    Public DownloadDIR As String   'downloaddirectory vanuit het oogpunt van de server
    Public ZipFile As String       'de ZIP-file met alle resultaten die uiteindelijk gedownload wordt
    Public ShapeFileZIP As String  'de Shapefile die gebruikt wordt voor conversie naar SOBEK
    Public ShapeField As String    'de naam van het veld in de shapefile dat het ID van de polygoon bevat
    Public ShapeFieldIdx As Integer 'indexnummer van het shapeveld voor gebiedsID

    'terugkoppeling naar de aanvrager per e-mail
    Friend GoodMail As clsEmail                       'the e-mail with good news
    Friend BadMail As clsEmail                        'the e-mail with bad news

    Friend FileCollectionNSL As New List(Of String)   'een verzameling met paden naar files die straks gezipped moeten worden
    Friend FileCollectionPEN As New List(Of String)   'een verzameling met paden naar files die straks gezipped moeten worden
    Friend FileCollectionMAK As New List(Of String)   'een verzameling met paden naar files die straks gezipped moeten worden
    Friend FileCollectionEVT As New List(Of String)   'een verzameling met paden naar files die staks gezipped moeten worden: SAT DATA Actual Evapotranspiration
    Friend FileCollectionSHO As New List(Of String)   'een verzameling met paden naar files die staks gezipped moeten worden: SAT DATA Evaportranspiration Shortage
    Friend FileCollectionMETA As New List(Of String)  'een verzameling met paden voor meta-bestanden die ook mee moeten in de zip
    Friend tmpFileCollection As New List(Of String)   'een tijdelijke verzameling met paden naar files

    Friend MetaFileContent As String
    Friend MetaFilePath As String

    Friend ShapeFile As String                        'pad naar de shapefile
    Friend NSLDir As String                           'directory met neerslagrasters
    Friend MAKDir As String                           'directory met makkink-rasters
    Friend PMDir As String                            'directory met penman-rasters

    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As clsSetup)
        Setup = mySetup
    End Sub

    Public Function Write() As Boolean
        Dim i As Long

        Try
            ''set the environment variable for GDAL
            'Environment.SetEnvironmentVariable("GDAL_DATA", GDALToolsDir & "\gdal-data")
            'Me.Setup.Log.AddMessage("GDAL Environment Variabele werd met succes ingesteld.")

            '--------------------------------------------------------------------------------------------------------------------------------------------
            '   VALIDATION OF THE SPATIAL AND TEMPORAL EXTENT
            '--------------------------------------------------------------------------------------------------------------------------------------------
            'round the co-ordinates to whole kilometres
            If FORMAAT = "ASCII" OrElse FORMAAT = "NETCDF" OrElse FORMAAT = "MODFLOW" OrElse FORMAAT = "SIMGRO" OrElse FORMAAT = "HDF5" Then
                Xmin = Setup.GeneralFunctions.RoundUD(Xmin / 1000, 0, False) * 1000
                Xmax = Setup.GeneralFunctions.RoundUD(Xmax / 1000, 0, True) * 1000
                Ymin = Setup.GeneralFunctions.RoundUD(Ymin / 1000, 0, False) * 1000
                Ymax = Setup.GeneralFunctions.RoundUD(Ymax / 1000, 0, True) * 1000
                If Xmin >= Xmax Or Ymin >= Ymax Then Throw New Exception("Ongeldige selectierechthoek op de kaart aangetroffen.")
                If Xmin < -20000 OrElse Xmax > 315000 OrElse Ymin < 265000 OrElse Ymax > 650000 Then Throw New Exception("Selectierechtoek te ver buiten de landsgrenzen.")

                'bereken het aantal tijdstappen * x * y en bepaal of de selectie acceptabel is
                Dim nt As Integer = 0
                For i = FDate To TDate
                    If Setup.GeneralFunctions.DateIntIsValid(i) Then
                        nt += 1
                    End If
                Next
                Me.Setup.Log.AddMessage("Ruimtelijke begrenzing werd met succes vastgelegd.")

                'foutafhandeling. Let op: is enigszins redundant want wordt al in het Javascript afgehandeld
                If nt > (366 * 2) Then
                    Throw New Exception("Uw tijdselectie was te groot. Selecteer maximaal twee jaren en probeer het opnieuw.")
                Else
                    Me.Setup.Log.AddMessage("De opgegeven ruimtelijke en temporele selecties werden geaccepteerd.")
                End If
                Me.Setup.Log.AddMessage("De gebruikerslimitaties werden succesvol afgehandeld.")
            End If
            '--------------------------------------------------------------------------------------------------------------------------------------------

            '--------------------------------------------------------------------------------------------------------------------------------------------
            ' PROCESSING THE ORDER
            '--------------------------------------------------------------------------------------------------------------------------------------------
            If NSL Then
                If Not writeNSL() Then
                    Me.Setup.Log.AddError("Er is een fout opgetreden bij het wegschrijven van de neerslag")
                Else
                    Me.Setup.Log.AddMessage("Neerslag werd met succes weggeschreven.")
                End If
            End If
            If MAKKINK Then
                If Not writeMAK() Then
                    Me.Setup.Log.AddError("Er is een fout opgetreden bij het wegschrijven van de Makkink-verdamping")
                Else
                    Me.Setup.Log.AddMessage("Makkink-verdamping werd met succes weggeschreven.")
                End If
            End If
            If PM Then
                If Not WritePM() Then
                    Me.Setup.Log.AddError("Er is een fout opgetreden bij het wegschrijven van de Penman-verdamping")
                Else
                    Me.Setup.Log.AddMessage("Penman-verdamping werd met succes weggeschreven.")
                End If
            End If
            If EVT_ACTUAL Then
                Dim StartDate As New Date(Left(FDate.ToString.Trim, 4), Left(Right(FDate.ToString.Trim, 4), 2), Right(FDate.ToString.Trim, 2))
                Dim EndDate As New Date(Left(TDate.ToString.Trim, 4), Left(Right(TDate.ToString.Trim, 4), 2), Right(TDate.ToString.Trim, 2))
                If Year(StartDate) < 2020 Then Me.Setup.Log.AddError("Startdatum ligt vóór de eerst beschikbare actuele verdampingsgegevens uit SATDATA 3.0. Beschikbaarheid begint op 1 januari 2020.")
                If EndDate > Now.AddDays(-92) Then Me.Setup.Log.AddWarning("Einddatum is mogelijk later dan de meest recente actuele verdampingsgegevens uit SATDATA 3.0. De verwerkingstijd bedraagt doorgaans 3 maanden.")
                If Not WriteEVT_ACT() Then
                    Me.Setup.Log.AddError("Er is een fout opgetreden bij het wegschrijven van de SATDATA 3.0 actuele evapotranspiratie.")
                Else
                    Me.Setup.Log.AddMessage("SATDATA 3.0 actuele evapotranspiratie werd met succes weggeschreven.")
                End If
            End If
            If EVT_SHORTAGE Then
                Dim StartDate As New Date(Left(FDate.ToString.Trim, 4), Left(Right(FDate.ToString.Trim, 4), 2), Right(FDate.ToString.Trim, 2))
                Dim EndDate As New Date(Left(TDate.ToString.Trim, 4), Left(Right(TDate.ToString.Trim, 4), 2), Right(TDate.ToString.Trim, 2))
                If Year(StartDate) < 2020 Then Me.Setup.Log.AddError("Startdatum ligt vóór de eerst beschikbare verdampingstekorten uit SATDATA 3.0. Beschikbaarheid begint op 1 januari 2020.")
                If EndDate > Now.AddDays(-92) Then Me.Setup.Log.AddWarning("Einddatum is mogelijk later dan de meest recente verdampingstekorten uit SATDATA 3.0. De verwerkingstijd bedraagt doorgaans 3 maanden.")

                If Not WriteEVT_SHO() Then
                    Me.Setup.Log.AddError("Er is een fout opgetreden bij het wegschrijven van het SATDATA 3.0 evapotranspiratietekort.")
                Else
                    Me.Setup.Log.AddMessage("SATDATA 3.0 evapotranspiratietekort werd met succes weggeschreven.")
                End If
            End If
            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Console.WriteLine("An error occurred in sub Write of class clsMBRasterData.")
            Return False
        End Try

    End Function

    Public Function UnZipShapeFile() As Boolean
        'Author: Siebe Bosch
        'Date: 23-5-2013
        'Description: this function unzips the contents of a compressed shapefile and places it in the temporary upload-dir
        'It also stores the path to the shp-file
        Dim myZip As New ZipFile, SHP As Boolean, SHX As Boolean, DBF As Boolean

        Try
            'clear the uploaddir
            Setup.GeneralFunctions.DeleteFilesInDir(Unzipdir, "*.*", System.IO.SearchOption.TopDirectoryOnly)

            'if the zip-file exists
            If System.IO.File.Exists(ShapeFileZIP) Then
                myZip = New ZipFile(ShapeFileZIP)
                If myZip.CheckZip(ShapeFileZIP) Then
                    myZip.ExtractAll(Unzipdir, ExtractExistingFileAction.OverwriteSilently)

                    Dim di = New IO.DirectoryInfo(Unzipdir)
                    Dim diar1 As IO.FileInfo() = di.GetFiles()
                    Dim dra As IO.FileInfo

                    'list the names of all files in the specified directory
                    For Each dra In diar1
                        If dra.Extension.Trim.ToUpper = ".SHP" Then
                            ShapeFile = dra.FullName
                            SHP = True
                        End If
                        If dra.Extension.Trim.ToUpper = ".SHX" Then SHX = True
                        If dra.Extension.Trim.ToUpper = ".DBF" Then DBF = True
                    Next

                    'is the shapefile complete?
                    If SHP AndAlso SHX AndAlso DBF Then
                        Return True
                    Else
                        Throw New Exception("Fout: shapefile was niet compleet. De ZIP-file moet een .shp, .shx en .dbf-bestand bevatten.")
                    End If
                Else
                    Throw New Exception("Fout: aangeleverde zip-bestand is niet geldig of leesbaar.")
                End If
            Else
                Throw New Exception("Fout: kon gecomprimeerd bestand met shapefile " & ShapeFileZIP & " niet vinden op de server.")
            End If
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function InitializeGoodMail(ByVal GegevensSoort As String) As Boolean
        Try
            'initialiseer de email
            GoodMail = New clsEmail(Me.Setup)
            GoodMail.Message.Subject = "Meteobase bestelling " & OrderNum & " " & GegevensSoort

            Dim body As String
            body = "Geachte " & Naam & "," & vbCrLf
            body &= vbCrLf

            If Me.Setup.Log.CountErrors > 0 Then
                body &= vbCrLf
                body &= "Uw bestelling is afgehandeld, echter, met de volgende foutmeldingen:" & vbCrLf
                body &= vbCrLf
                For i = 0 To Me.Setup.Log.CountErrors - 1
                    body &= "*" & vbTab & Me.Setup.Log.Errors(i) & vbCrLf
                Next
                body &= vbCrLf
                body &= "Uw bestanden staan klaar in de download-directory van Meteobase. Klik op de onderstaande link om ze op te halen." & vbCrLf
                body &= DownloadURL & ZipFile & vbCrLf
            Else
                body &= "Uw bestelling is met succes afgehandeld en staat klaar in de download-directory van Meteobase. Klik op de onderstaande link om hem op te halen." & vbCrLf
                body &= DownloadURL & ZipFile & vbCrLf
            End If


            body &= vbCrLf
            body &= "Diagnostische gegevens: " & vbCrLf
            body &= "Session ID " & SessionID.ToString & vbCrLf
            body &= "Bestelnummer " & OrderNum.ToString & vbCrLf
            body &= "E-mailadres " & MailAdres & vbCrLf
            body &= "Resultatenbestand " & ZipFile & vbCrLf
            body &= "Coördinaten: xmin=" & Xmin.ToString & " xmax=" & Xmax.ToString & " ymin=" & Ymin.ToString & " ymax=" & Ymax.ToString & vbCrLf
            body &= "Tijdsspanne: van=" & FDate.ToString & " tot=" & TDate.ToString & vbCrLf
            body &= vbCrLf

            If Me.Setup.Log.Warnings.Count > 0 Then
                body &= vbCrLf
                body &= "Eventuele waarschuwingen:" & vbCrLf
                For Each myStr As String In Me.Setup.Log.Warnings
                    body &= myStr & vbCrLf
                Next
            End If

            If Me.Setup.Log.Errors.Count > 0 Then
                body &= vbCrLf
                body &= "Eventuele foutmeldingen:" & vbCrLf
                For Each myStr As String In Me.Setup.Log.Errors
                    body &= myStr & vbCrLf
                Next
            End If

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
            body &= "Aangeboden door Het Waterschapshuis | www.hetwaterschapshuis.nl" & vbCrLf
            body &= vbCrLf
            body &= "Mogelijk gemaakt door" & vbCrLf
            body &= "Hydrologic            | www.hydrologic.nl" & vbCrLf
            body &= "HKV-Lijn in water     | www.hkv.nl" & vbCrLf
            body &= "Hydroconsult          | www.hydroconsult.nl" & vbCrLf
            body &= "--------------------------------------------" & vbCrLf

            GoodMail.SetBodyContent(body)

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error initializing Good Email.")
            Return False
        End Try
    End Function


    Public Function InitializeBadMail(ByVal GegevensSoort As String) As Boolean
        Try
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
            body &= "Resultatenbestand " & ZipFile & vbCrLf
            body &= "Coördinaten: xmin=" & Xmin.ToString & " xmax=" & Xmax.ToString & " ymin=" & Ymin.ToString & " ymax=" & Ymax.ToString & vbCrLf
            body &= "Tijdsspanne: van=" & FDate.ToString & " tot=" & TDate.ToString & vbCrLf
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
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function sendGoodEmail() As Boolean
        Try
            'eerst naar de aanvrager zelf
            If Not GoodMail.Send(MailAdres, Naam) Then Throw New Exception("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")

            'vul de mail aan met diagnostics en stuur daarna een kopie naar onszelf
            Call GoodMail.addDiagnosticsToBody()
            If Not GoodMail.Send("info@meteobase.nl", "Meteobase") Then Throw New Exception("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function sendBadEmail() As Boolean
        Try
            'eerst naar de aanvrager zelf
            If Not BadMail.Send(MailAdres, Naam) Then
                Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
            End If

            'dan een kopie naar onszelf
            Call BadMail.addDiagnosticsToBody()
            If Not BadMail.Send("info@meteobase.nl", "Meteobase") Then
                Me.Setup.Log.AddError("Verzenden e-mail is niet gelukt. Neem a.u.b. contact met ons op via info@meteobase.nl.")
            End If
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function WriteMAK_NC()
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts Makkink grids to ASC-format
        Dim myYear As Integer, myPath As String, outputFile As String, myArgs As String
        Dim i As Long, k As Long
        'Dim XMLFile As String = ""
        Dim ResultsZIPFile As String 'zipfile met het resultaat
        Dim MAKDir As String = RasterDir & "\MAK_RD"
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim OutputDir As String = TempDir & "\MAK"
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        ResultsZIPFile = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-Apache\Php\apache\www\meteobase\downloads\" & ZipFile

        If Not System.IO.Directory.Exists(OutputDir) Then
            System.IO.Directory.CreateDirectory(OutputDir)
        End If

        'makkink
        For i = FDate To TDate
            If Setup.GeneralFunctions.DateIntIsValid(i) Then
                myYear = Left(i, 4)
                myPath = MAKDir & "\" & myYear & "\EVAP_MK_" & Str(i).Trim & ".nc"
                If System.IO.File.Exists(myPath) Then
                    myArgs = ""

                    outputFile = OutputDir & "\" & Str(OrderNum).Trim & "_MAK_" & Str(i).Trim & ".nc"
                    'XMLFile = Replace(outputFile, ".nc", ".nc.aux.xml")
                    FileCollectionMAK.Add(outputFile)
                    'FileCollectionMAK.Add(XMLFile)
                    'hier volgen de argumenten. 
                    myArgs = "-of NetCDF -a_nodata -999 -b 1 " & "-projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile
                    ProcessCollection.Add(Process.Start(GDALExePath, myArgs))

                    'wachten zolang er nog te veel incomplete shells openstaan!
                    Dim mySW As New System.Diagnostics.Stopwatch
                    mySW.Start()
                    Dim nUnfinished As Integer = 5
                    While nUnfinished > 4
                        nUnfinished = 0
                        For k = 0 To ProcessCollection.Count - 1
                            If Not ProcessCollection(k).HasExited Then nUnfinished += 1
                        Next
                        If mySW.ElapsedMilliseconds > 30000 Then nUnfinished = 0 'veligheidsklep
                    End While

                End If
            Else
            End If
        Next

        'ga door tot alle files geschreven zijn
        Dim SW As New System.Diagnostics.Stopwatch
        SW.Reset()
        SW.Start()
        Dim Done As Boolean = False
        While Not Done
            Done = True
            For i = 0 To ProcessCollection.Count - 1
                If Not ProcessCollection.Item(i).HasExited Then Done = False
            Next
            If SW.ElapsedMilliseconds > 600000 Then Done = True '10 minutes processing time is the maximum
        End While

        SW.Stop()
        SW.Reset()
        Return True
    End Function


    Public Function WriteMAK_ASC()
        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim unzipFile As String = TempDir & "\tmp.tif"
            Dim TempResultsDir As String = TempDir & "\MAK"
            Dim ZipFilePath As String = TempDir & "\MAK_RAW.zip"
            Dim ut As New MapWinGIS.Utils
            Dim FileName As String = ""
            Dim CurDate As New DateTime(Left(FDate, 4), Left(Right(FDate, 4), 2), Right(FDate, 2))

            If Not System.IO.Directory.Exists(TempResultsDir) Then System.IO.Directory.CreateDirectory(TempResultsDir)
            If Not WIWB.DownloadRasters("Meteobase.Evaporation.Makkink", "Evaporation", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")

            'now that the zipfile has been downloaded, read each individual file from the file and warp & translate it
            Dim zipMS As New MemoryStream()
            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    ZipEntry.Extract(zipMS)
                    CurDate = CurDate.AddDays(1)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'create a filename for evaporation
                    FileName = MakeFileName("MAK", CurDate, False, "ASC")

                    'use the gdal drivers inside SOBEK utilities to translate to ASCII
                    'notice that warping is not necessary since the evaporation rasters are already in RD coordinates
                    ut.TranslateRaster(unzipFile, TempResultsDir & "\" & FileName, "-of AAIGrid")

                    'finally add the newly created file to a collection of paths, for later zipping
                    FileCollectionMAK.Add(TempResultsDir & "\" & FileName)
                Next
            End Using
            ut = Nothing

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing Makkink evaporation grids.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function


    Public Function writeNSL() As Boolean
        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        Try
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If

            Select Case FORMAAT.Trim.ToUpper
                Case Is = "NETCDF"
                    If Not WriteNSL_NC() Then Return False
                Case Is = "HDF5"
                    If Not WriteNSL_HDF5() Then Return False
                Case Is = "ASCII"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "MODFLOW"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "SIMGRO"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "SOBEK"
                    If Not WriteNSL_POLY("BUI") Then Return False
                Case Is = "CSV"
                    If Not WriteNSL_POLY("csv") Then Return False
            End Select

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WritePM() As Boolean

        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        If Not System.IO.Directory.Exists(TempDir) Then
            System.IO.Directory.CreateDirectory(TempDir)
        End If

        Select Case FORMAAT.Trim.ToUpper
            Case Is = "ASCII"
                If Not Write_ASC("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
            Case Is = "NETCDF"
                Me.Setup.Log.AddError("Error: export naar NetCDF wordt niet langer ondersteund.")
            Case Is = "HDF5"
                If Not Write_HDF5("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
            Case Is = "MODFLOW"
                If Not Write_ASC("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
            Case Is = "SIMGRO"
                If Not Write_ASC("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
            Case Is = "SOBEK"
                If Not WriteEVP_POLY("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
            Case Is = "CSV"
                If Not WriteEVP_POLY("Meteobase.Evaporation.PennmanMonteith", "Evaporation", "MB_PM", FileCollectionPEN) Then Return False
        End Select

        Return True

    End Function

    Public Function WriteEVT_ACT() As Boolean

        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        If Not System.IO.Directory.Exists(TempDir) Then
            System.IO.Directory.CreateDirectory(TempDir)
        End If

        Select Case FORMAAT.Trim.ToUpper
            Case Is = "ASCII"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
            Case Is = "NETCDF"
                Me.Setup.Log.AddError("Error: export naar NetCDF wordt niet langer ondersteund.")
            Case Is = "HDF5"
                If Not Write_HDF5("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
            Case Is = "MODFLOW"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
            Case Is = "SIMGRO"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
            Case Is = "SOBEK"
                If Not WriteEVP_POLY("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
            Case Is = "CSV"
                If Not WriteEVP_POLY("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationActual", "EVT_ACT", FileCollectionEVT) Then Return False
        End Select

        Return True

    End Function

    Public Function WriteEVT_SHO() As Boolean

        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        If Not System.IO.Directory.Exists(TempDir) Then
            System.IO.Directory.CreateDirectory(TempDir)
        End If

        Select Case FORMAAT.Trim.ToUpper
            Case Is = "ASCII"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
            Case Is = "NETCDF"
                'If Not WritePM_NC() Then Return False
            Case Is = "HDF5"
                If Not Write_HDF5("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
            Case Is = "MODFLOW"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
            Case Is = "SIMGRO"
                If Not Write_ASC("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
            Case Is = "SOBEK"
                If Not WriteEVP_POLY("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
            Case Is = "CSV"
                If Not WriteEVP_POLY("Satdata.Evapotranspiration.Reanalysis", "EvapotranspirationShortage", "EVT_SHO", FileCollectionSHO) Then Return False
        End Select

        Return True

    End Function

    Public Function writeMAK() As Boolean
        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        If Not System.IO.Directory.Exists(TempDir) Then
            System.IO.Directory.CreateDirectory(TempDir)
        End If

        'corrigeer de datum voor jaartallen die mogelijk nog niet beschikbaar zijn.
        'Siebe: let op: jaarlijks updaten!!!
        Dim LastMakkinkDate As Integer = My.Settings.LastMakkinkDate

        If TDate > LastMakkinkDate Then
            Setup.Log.AddError("Ongeldige datumselectie. Verdamping volgens Makkink uitsluitend beschikbaar tot " & LastMakkinkDate & ". Einddatum werd automatisch gecorrigeerd.")
            TDate = LastMakkinkDate
        End If
        If FDate > LastMakkinkDate Then
            Throw New Exception("Ongeldige datumselectie. Verdamping volgens Makkink uitsluitend beschikbaar tot " & LastMakkinkDate & ". Bestelling werd afgebroken.")
        End If

        Select Case FORMAAT.Trim.ToUpper
            Case Is = "ASCII"
                If Not Write_ASC("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
            Case Is = "NETCDF"
                Me.Setup.Log.AddError("Error: export naar NetCDF wordt niet langer ondersteund.")
            Case Is = "HDF5"
                If Not Write_HDF5("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
            Case Is = "MODFLOW"
                If Not Write_ASC("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
            Case Is = "SIMGRO"
                If Not Write_ASC("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
            Case Is = "SOBEK"
                If Not WriteEVP_POLY("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
            Case Is = "CSV"
                If Not WriteEVP_POLY("Meteobase.Evaporation.Makkink", "Evaporation", "MB_MAK", FileCollectionMAK) Then Return False
        End Select

        Return True

    End Function



    Public Function WritePM_NC()
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts Penman-Monteith grids to ASC-format
        Dim myYear As Integer, myPath As String, outputFile As String, myArgs As String
        Dim i As Long, k As Long
        'Dim XMLFile As String = ""
        Dim ResultsZIPFile As String 'zipfile met het resultaat
        Dim PMDir As String = RasterDir & "\PM_RD"
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim OutputDir As String = TempDir & "\PM"
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        ResultsZIPFile = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-Apache\Php\apache\www\meteobase\downloads\" & ZipFile

        If Not System.IO.Directory.Exists(OutputDir) Then
            System.IO.Directory.CreateDirectory(OutputDir)
        End If

        'penman monteith
        For i = FDate To TDate
            If Setup.GeneralFunctions.DateIntIsValid(i) Then
                myYear = Left(i, 4)
                myPath = PMDir & "\" & myYear & "\EVAP_PM_" & Str(i).Trim & ".nc"
                If System.IO.File.Exists(myPath) Then
                    myArgs = ""

                    outputFile = OutputDir & "\" & Str(OrderNum).Trim & "_PM_" & Str(i).Trim & ".nc"
                    'XMLFile = Replace(outputFile, ".nc", ".nc.aux.xml")
                    FileCollectionPEN.Add(outputFile)
                    'FileCollectionPEN.Add(XMLFile)
                    'hier volgen de argumenten.
                    myArgs = "-of NetCDF -a_nodata -999 -b 1 " & "-projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile
                    ProcessCollection.Add(Process.Start(GDALExePath, myArgs))

                    'wachten zolang er nog te veel incomplete shells openstaan!
                    Dim mySW As New System.Diagnostics.Stopwatch
                    mySW.Start()
                    Dim nUnfinished As Integer = 5
                    While nUnfinished > 4
                        nUnfinished = 0
                        For k = 0 To ProcessCollection.Count - 1
                            If Not ProcessCollection(k).HasExited Then nUnfinished += 1
                        Next
                        If mySW.ElapsedMilliseconds > 30000 Then nUnfinished = 0 'veligheidsklep
                    End While

                End If
            Else
            End If
        Next

        'ga door tot alle files geschreven zijn
        Dim SW As New System.Diagnostics.Stopwatch
        SW.Reset()
        SW.Start()
        Dim Done As Boolean = False
        While Not Done
            Done = True
            For i = 0 To ProcessCollection.Count - 1
                If Not ProcessCollection.Item(i).HasExited Then Done = False
            Next
            If SW.ElapsedMilliseconds > 600000 Then Done = True '10 minutes processing time is the maximum
        End While

        SW.Stop()
        SW.Reset()
        Return True
    End Function

    Public Function Write_HDF5(DataSource As String, Parameter As String, FileNameBase As String, ByRef FileCollection As List(Of String))
        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim ZipFilePath As String = TempDir & "\" & FileNameBase & ".Zip"
            If Not WIWB.DownloadRasters(DataSource, Parameter, Xmin, Ymin, Xmax, Ymax, FDate, TDate, "HDF5", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")
            FileCollection.Add(ZipFilePath)
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing evaporation grids.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function



    Public Function Write_ASC(DataSource As String, DataParameter As String, FileNameBase As String, ByRef FileCollection As List(Of String))

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim unzipFile As String = TempDir & "\tmp.tif"
            Dim TempResultsDir As String = TempDir & "\" & FileNameBase
            Dim ZipFilePath As String = TempDir & "\WIWB_RAW.zip"
            Dim ut As New MapWinGIS.Utils
            Dim FileName As String = ""
            Dim CurDate As New DateTime(Left(FDate, 4), Left(Right(FDate, 4), 2), Right(FDate, 2))

            If Not System.IO.Directory.Exists(TempResultsDir) Then System.IO.Directory.CreateDirectory(TempResultsDir)
            If Not WIWB.DownloadRasters(DataSource, DataParameter, Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath,, Aggregate24H) Then Throw New Exception("Error retrieving rasterdata from API.")

            'now that the zipfile has been downloaded, read each individual file from the file and warp & translate it
            Dim zipMS As New MemoryStream()
            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    ZipEntry.Extract(zipMS)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'create a filename for evaporation
                    FileName = MakeFileName(FileNameBase, CurDate, False, "ASC")

                    'use the gdal drivers inside SOBEK utilities to translate to ASCII
                    'notice that warping is not necessary since the evaporation rasters are already in RD coordinates
                    ut.TranslateRaster(unzipFile, TempResultsDir & "\" & FileName, "-Of AAIGrid")

                    'finally add the newly created file to a collection of paths, for later zipping
                    FileCollection.Add(TempResultsDir & "\" & FileName)
                    CurDate = CurDate.AddDays(1)
                Next
            End Using
            ut = Nothing

            '------------------------------------------------------------------------------------------------------------------
            'eenheidsconversie t.b.v. SIMGRO en MODFLOW
            '------------------------------------------------------------------------------------------------------------------
            'MetaFilePath = TempDir & "\Mete_grid.inp"
            'If FORMAAT.Trim.ToUpper = "SIMGRO" Then
            '    Using simWriter As New System.IO.StreamWriter(MetaFilePath)
            '        simWriter.Write(MetaFileContent)
            '    End Using
            '    If Not FileCollectionMETA.Contains(MetaFilePath) Then
            '        FileCollectionMETA.Add(MetaFilePath)
            '    End If
            'End If

            'if MODFLOW or SIMGRO, the grid values need adjustment (unit conversion)
            Dim myRaster As clsASCIIGrid
            If FORMAAT.Trim.ToUpper = "MODFLOW" OrElse FORMAAT.Trim.ToUpper = "SIMGRO" Then
                Me.Setup.Log.AddMessage("Exportformaat = " & FORMAAT & ". conversie van eenheden wordt uitgevoerd.")
                For Each myPath In FileCollection
                    myRaster = New clsASCIIGrid(Me.Setup)
                    If Not myRaster.Read(myPath) Then
                        Me.Setup.Log.AddError("Kon raster niet lezen voor eenheidsconversie: " & myPath & " eenheid blijft 0.1 mm/etmaal")
                    Else
                        If FORMAAT.Trim.ToUpper = "MODFLOW" Then
                            For r = 0 To myRaster.rows - 1
                                For c = 0 To myRaster.cols - 1
                                    'If Not myRaster.cells(r, c) = myRaster.nodataval Then
                                    'myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                    'End If
                                Next
                            Next
                        ElseIf FORMAAT.Trim.ToUpper = "SIMGRO" Then
                            For r = 0 To myRaster.rows - 1
                                For c = 0 To myRaster.cols - 1
                                    If myRaster.cells(r, c) = myRaster.nodataval Then
                                        myRaster.cells(r, c) = 0
                                        'Else
                                        'myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                    End If
                                Next
                            Next
                        End If
                    End If
                    If Not myRaster.Write(myPath) Then Me.Setup.Log.AddError("Kon raster niet schrijven na eenheidsconversie: " & myPath)
                Next
            End If


            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing Penman-Monteith evaporation grids.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WritePM_ASC()

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim unzipFile As String = TempDir & "\tmp.tif"
            Dim TempResultsDir As String = TempDir & "\PEN"
            Dim ZipFilePath As String = TempDir & "\PEN_RAW.zip"
            Dim ut As New MapWinGIS.Utils
            Dim FileName As String = ""
            Dim CurDate As New DateTime(Left(FDate, 4), Left(Right(FDate, 4), 2), Right(FDate, 2))

            If Not System.IO.Directory.Exists(TempResultsDir) Then System.IO.Directory.CreateDirectory(TempResultsDir)
            FileCollectionPEN = New List(Of String)
            If Not WIWB.DownloadRasters("Meteobase.Evaporation.PennmanMonteith", "Evaporation", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")

            'now that the zipfile has been downloaded, read each individual file from the file and warp & translate it
            Dim zipMS As New MemoryStream()
            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    ZipEntry.Extract(zipMS)
                    CurDate = CurDate.AddDays(1)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'create a filename for evaporation
                    FileName = MakeFileName("PEN", CurDate, False, "ASC")

                    'use the gdal drivers inside SOBEK utilities to translate to ASCII
                    'notice that warping is not necessary since the evaporation rasters are already in RD coordinates
                    ut.TranslateRaster(unzipFile, TempResultsDir & "\" & FileName, "-of AAIGrid")

                    'finally add the newly created file to a collection of paths, for later zipping
                    FileCollectionPEN.Add(TempResultsDir & "\" & FileName)
                Next
            End Using
            ut = Nothing

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing Penman-Monteith evaporation grids.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WriteNSL_NC() As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim unzipFile As String = TempDir & "\tmp.tif"
            Dim TempResultsDir As String = TempDir & "\NSL"
            Dim ZipFilePath As String = TempDir & "\NSL_RAW.zip"
            Dim ut As New MapWinGIS.Utils

            If Not System.IO.Directory.Exists(TempResultsDir) Then System.IO.Directory.CreateDirectory(TempResultsDir)
            FileCollectionNSL = New List(Of String)
            'If Not WIWB.GetRasters("KNMI.Radar.Uncorrected", "P", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")
            If Not WIWB.DownloadRasters("Meteobase.Precipitation", "P", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")

            'now that the zipfile has been downloaded, read each individual file from the file and warp & translate it
            Dim zipMS As New MemoryStream()
            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    ZipEntry.Extract(zipMS)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    If System.IO.File.Exists(unzipFile) Then System.IO.File.Delete(unzipFile)
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'use the gdal drivers inside SOBEK utilities to 1) warp to RD and write to NetCDF
                    'ut.GDALWarp(unzipFile, TempResultsDir & "\" & Replace(ZipEntry.FileName, ".tif", ".nc"), " -of NetCDF -s_srs " & Chr(34) & "+proj=stere +lat_0=90 +lat_ts=60 +lon_0=0 +k=1 +x_0=0 +y_0=0 +a=6378.14 +b=6356.75" & Chr(34) & " -t_srs EPSG:28992")

                    'notice that Meteobase rasters are already in RD, so no warping needed
                    ut.TranslateRaster(unzipFile, TempResultsDir & "\" & Replace(ZipEntry.FileName, ".tif", ".nc"), " -of NetCDF" & " -t_srs EPSG:28992")

                    'finally add the newly created file to a collection of paths, for later zipping
                    FileCollectionNSL.Add(TempResultsDir & "\" & Replace(ZipEntry.FileName, ".tif", ".nc"))
                Next
            End Using

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing daily precipitation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try



    End Function


    Public Function WriteNSL_HDF5() As Boolean
        'hdf5 is the native file format for data residing on the WIWB database
        'this means that there is NO warping and NO translation required. We will simply zip the received zipfile and provide that directly to the customer
        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim ZipFilePath As String = TempDir & "\NSL_HDF5.zip"

            FileCollectionNSL = New List(Of String)
            If Not WIWB.DownloadRasters("Meteobase.Precipitation", "P", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "HDF5", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")
            FileCollectionNSL.Add(ZipFilePath)

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing daily precipitation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function AggregateBandsAndExport() As Boolean
        'Author: Siebe Bosch
        'Date: 13-8-2013
        'Description: This function aggregates all bands inside a netcdf-file and exports the result
        Dim inputFile As String, outputFile As String
        Dim mbGrid As MapWinGIS.Grid, myGrid As MapWinGIS.Grid
        Dim i As Integer, r As Integer, c As Integer, Vals(,) As Single
        Dim maxrowidx As Integer, maxcolidx As Integer
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim myCollection As New List(Of String)
        Dim FileName As String

        Try

            ' make a reference to a directory
            Dim di As New IO.DirectoryInfo(RasterDir)
            Dim diar1 As IO.FileInfo() = di.GetFiles()
            Dim dra As IO.FileInfo

            'list the names of all files in the specified directory
            For Each dra In diar1
                myCollection.Add(dra.FullName)
            Next

            'All we'll have to do now is walk through the results of WriteNSL_NC
            For Each inputFile In myCollection
                FileName = Me.Setup.GeneralFunctions.FileNameFromPath(inputFile)
                outputFile = Me.Setup.Settings.ExportDirRoot & "\" & Replace(FileName, ".nc", ".asc")

                mbGrid = New MapWinGIS.Grid
                If Not mbGrid.Open(inputFile, GridDataType.UnknownDataType, True) Then Throw New Exception("Kon neerslaggrid niet openen.")

                If Not mbGrid.OpenBand(3) Then Throw New Exception("Kon neerslagband 1 niet vinden in netcdf-file.")

                myGrid = New MapWinGIS.Grid
                myGrid.CreateNew(outputFile, mbGrid.Header, GridDataType.FloatDataType, 0, True, GridFileType.Ascii)
                maxrowidx = myGrid.Header.NumberRows - 1
                maxcolidx = myGrid.Header.NumberCols - 1

                For i = 1 To mbGrid.NumBands
                    mbGrid.OpenBand(i)
                    Vals = Me.Setup.GeneralFunctions.ArrayFromMapWindowGrid(mbGrid)
                    For r = 0 To maxrowidx
                        For c = 0 To maxcolidx
                            myGrid.Value(c, r) += Vals(c, r) / 100
                        Next
                    Next
                Next

                If Not myGrid.Save() Then
                    Throw New Exception("Kon geaggregeerd neerslagraster niet schrijven.")
                End If
            Next

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
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

    Public Function ExtractZIP(ZipFilePath As String, ByRef CurDate As DateTime, TempResultsDir As String, ByRef MetaFileContent As String, ByVal WarpFromKNMI As Boolean, GDALOutputFormat As String) As Boolean
        Try
            'now that the zipfile has been downloaded, read each individual file from the file and warp & translate it
            Dim i As Integer = 0
            Dim unzipFile As String = TempDir & "\tmp.tif"
            Dim zipMS As New MemoryStream()
            Dim ut As New MapWinGIS.Utils
            Dim FileName As String = ""
            Dim DayNum As Integer = 0

            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                Me.Setup.Log.AddMessage("Processing " & zipReceived.Entries.Count & " files in " & ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    ZipEntry.Extract(zipMS)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    If System.IO.File.Exists(unzipFile) Then System.IO.File.Delete(unzipFile)
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'create a filename for precipitation
                    FileName = MakeFileName("NSL", CurDate, True, "ASC")

                    If WarpFromKNMI Then
                        Dim args As String = "-overwrite -s_srs " & Chr(34) & "+proj=stere +lat_0=90 +lat_ts=60 +lon_0=0 +k=1 +x_0=0 +y_0=0 +a=6378.14 +b=6356.75" & Chr(34) & " -t_srs EPSG:28992"
                        Me.Setup.Log.AddMessage("Warping raster " & unzipFile & " from KNMI coordinates, using gdalwarp with arguments: " & args)
                        If Not ut.GDALWarp(unzipFile, unzipFile & "_RD", args) Then Throw New Exception("Unable to warp precipitation grid to RD coordinates: " & unzipFile)
                        Me.Setup.Log.AddMessage("Translating raster " & unzipFile & "_RD" & " to " & GDALOutputFormat & ": " & FileName)
                        If Not System.IO.File.Exists(unzipFile & "_RD") Then Me.Setup.Log.AddError("Error: warped file does not exist: " & unzipFile & "_RD")
                        If Not ut.TranslateRaster(unzipFile & "_RD", TempResultsDir & "\" & FileName, "-of " & GDALOutputFormat) Then Throw New Exception("Could not translate precipitation grid to the required file format.")
                    Else
                        Me.Setup.Log.AddMessage("Translating raster to ASCII format: " & unzipFile)
                        ut.TranslateRaster(unzipFile, TempResultsDir & "\" & FileName, "-of AAIGrid")
                    End If


                    'finally add the newly created file to a collection of paths, for later zipping
                    FileCollectionNSL.Add(TempResultsDir & "\" & FileName)

                    'hier schrijven we de metadata weg t.b.v. SIMGRO. Merk op dat we ook maar meteen een verwijzing naar verdamping maken, al weten we natuurlijk niet of die ook wordt opgevraagd.
                    'let op: het kan zijn dat SIMGRO eigenlijk _24 wil voor wat 00h moet zijn. In dat geval moeten we dit nog ombouwen
                    Dim NSLFile As String = Setup.GeneralFunctions.FileNameFromPath(FileCollectionNSL.Item(FileCollectionNSL.Count - 1))
                    Dim MAKFile As String = MakeFileName("MAK", CurDate, False, "ASC")
                    MetaFileContent &= Format(DayNum + i / 24, "0.0000000000") & ", " & Format(CurDate.Year, "0000") & "," & Chr(34) & NSLFile & Chr(34) & "," & Chr(34) & MAKFile & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & vbCrLf

                    'keep track of the time proceeding
                    CurDate = CurDate.AddHours(1)
                    i += 1
                    If i = 24 Then
                        i = 0
                        DayNum += 1
                    End If
                Next
            End Using
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function WriteNSL_ASC() As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim TempResultsDir As String = TempDir & "\NSL"
            Dim warpFile As String = TempDir & "\tmp_warped.tif"
            Dim MetaFilePath As String = TempDir & "\Mete_grid.inp"
            Dim ZipFilePathPre2019 As String = TempDir & "\NSL_pre2019.zip"
            Dim ZipFilePathPost2019 As String = TempDir & "\NSL_post2019.zip"
            Dim MetaFileContent As String = ""
            Dim myRaster As clsASCIIGrid
            Dim CurDate As New DateTime(Left(FDate, 4), Left(Right(FDate, 4), 2), Right(FDate, 2))

            If Not System.IO.Directory.Exists(TempResultsDir) Then System.IO.Directory.CreateDirectory(TempResultsDir)
            FileCollectionNSL = New List(Of String)
            FileCollectionMETA = New List(Of String)

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
                If Not WIWB.DownloadRasters("Meteobase.Precipitation", "P", Xmin, Ymin, Xmax, Ymax, FDatePre2019, TDatePre2019, "geotiff", ZipFilePathPre2019) Then Throw New Exception("Error retrieving rasterdata from API.")
                If Not ExtractZIP(ZipFilePathPre2019, CurDate, TempResultsDir, MetaFileContent, False, "AAIGrid") Then Throw New Exception("Error extracting data received from WIWB server.")
            End If

            'handle the post-2019 orders
            If FDatePost2019 > 0 AndAlso TDatePost2019 > 0 Then
                Me.Setup.Log.AddMessage("Processing post-january 2019 data.")
                If Not WIWB.DownloadRasters("Knmi.International.Radar.Composite.Final.Reanalysis", "P", Xmin, Ymin, Xmax, Ymax, FDatePost2019, TDatePost2019, "geotiff", ZipFilePathPost2019) Then Throw New Exception("Error retrieving rasterdata from API.")
                If Not ExtractZIP(ZipFilePathPost2019, CurDate, TempResultsDir, MetaFileContent, True, "AAIGrid") Then Throw New Exception("Error extracting data received from WIWB server.")
            End If

            'as a final action, write the SIMGRO meta file (if required)
            If FORMAAT.Trim.ToUpper = "SIMGRO" Then
                Using simWriter As New System.IO.StreamWriter(MetaFilePath)
                    simWriter.Write(MetaFileContent)
                End Using
                If Not FileCollectionMETA.Contains(MetaFilePath) Then
                    FileCollectionMETA.Add(MetaFilePath)
                End If

                'every grid for SIMGRO must contain values in the unit mm/d, even if they are hourly grids
                Me.Setup.Log.AddMessage("Exportformaat = " & FORMAAT & ". conversie van eenheden wordt uitgevoerd.")
                For Each myPath In FileCollectionNSL
                    myRaster = New clsASCIIGrid(Me.Setup)
                    If Not myRaster.Read(myPath) Then
                        Me.Setup.Log.AddError("Kon rasterbestand niet openen voor bewerking. Eenheid van de rasterdata blijft 1 mm/h")
                        Exit For
                    Else
                        For r = 0 To myRaster.rows - 1
                            For c = 0 To myRaster.cols - 1
                                If myRaster.cells(r, c) = myRaster.nodataval Then
                                    myRaster.cells(r, c) = 0
                                Else
                                    myRaster.cells(r, c) = myRaster.cells(r, c) * 24
                                End If
                            Next
                        Next
                        If Not myRaster.Write(myPath) Then Me.Setup.Log.AddError("Could not write grid file after unit conversion: " & myPath)
                    End If
                Next

            End If

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing daily precipitation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try


    End Function

    Public Function MakeFileName(DataType As String, CurDate As DateTime, IncludeHour As Boolean, FileExtension As String) As String
        If IncludeHour Then
            Return DataType & "_" & Format(CurDate.Year, "0000") & Format(CurDate.Month, "00") & Format(CurDate.Day, "00") & "_" & Format(CurDate.Hour, "00") & "." & FileExtension
        Else
            Return DataType & "_" & Format(CurDate.Year, "0000") & Format(CurDate.Month, "00") & Format(CurDate.Day, "00") & "." & FileExtension
        End If
    End Function

    Public Function WriteNSL_POLY(ByVal FileExtension As String) As Boolean
        'Author: Siebe Bosch
        'Date: 22 april 2013
        'Description: writes a SOBEK .bui file from Meteobase rasterdata
        'Exportformat can either be BUI or CSV
        'Dim sf As New MapWinGIS.Shapefile
        'Dim NewProjection As MapWinGIS.GeoProjection
        Dim WIWB As New clsWIWB_API(Me.Setup)
        Dim CurProjection As MapWinGIS.GeoProjection
        Dim outputFile As String
        Dim ResultsBUIFile As String
        Dim NSLDir As String = RasterDir & "\NSL_RD"
        Dim BuiFile As New clsMeteoFile(Me.Setup)
        Dim EraIdx As Integer
        Dim iShape As Integer, tsIdx As Long = -1
        Dim myStation As clsMeteoStation = Nothing
        Dim myShape As Shape
        Dim mean As Double, min As Double, max As Double
        Dim StationName As String
        Dim ZipFilePathPre2019 As String = TempDir & "\NSL_pre2019.zip"
        Dim ZipFilePathPost2019 As String = TempDir & "\NSL_post2019.zip"
        Dim unzipFile As String = TempDir & "\tmp.tif"
        Dim TempResultsDir As String = TempDir & "\NSL"
        Dim MetaFileContent As String = ""
        Dim ut As New MapWinGIS.Utils

        Try
            Me.Setup.GISData.SubcatchmentShapeFile.Initialize()
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf = New MapWinGIS.Shapefile

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
                Me.Setup.Log.AddMessage("Map voor tussentijds resultaat is met succes aangemaakt.")
            End If

            'create a path for the resulting bui-file
            ResultsBUIFile = TempDir & "\MB_NSL." & FileExtension
            If File.Exists(ResultsBUIFile) Then File.Delete(ResultsBUIFile)

            'open the shapefile and set its projection by default to Amersfoort New
            If Not Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Open(ShapeFile) Then Throw New Exception("Fout: kon de shapefile niet openen.")
            Me.Setup.Log.AddMessage("Shapefile kon met succes worden geopend")

            If Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.GeoProjection.Name = "" Then
                Me.Setup.Log.AddMessage("Shapefile had geen projectie; RD (Amersfoort) werd aangenomen.")
            ElseIf Not Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.GeoProjection.Name = "RD_New" Then
                Me.Setup.Log.AddMessage("Projectie shapefile werd niet herkend; RD (Amersfoort) werd aangenomen.")
            End If

            'assign RD_New projection anyhow because some dialects are not recognized by MapWindow
            CurProjection = New MapWinGIS.GeoProjection
            CurProjection.ImportFromProj4("+proj=sterea +lat_0=52.15616055555555 +lon_0=5.38763888888889 +k=0.9999079 +x_0=155000 +y_0=463000 +ellps=bessel +units=m +no_defs")
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.GeoProjection = CurProjection

            If Not Me.Setup.GISData.SubcatchmentShapeFile.findAreaIDFieldIdx(ShapeField) Then Throw New Exception("Fout: kon veld met gebiedsID " & ShapeField & " niet vinden in shapefile.")
            Me.Setup.Log.AddMessage("Geldig Shapeveld werd gevonden.")
            ''If Not Me.Setup.GISData.AreaShapeFile.IDsUnique Then Throw New Exception("Fout: dubbele ID's in shapefile aangetroffen. Iedere polygoon moet een uniek gebiedsID hebben.")

            'every shape gets its own meteorological station
            For iShape = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(iShape)
                StationName = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.CellValue(Me.Setup.GISData.SubcatchmentShapeFile.SubcatchmentIDFieldIdx, iShape).ToString
                BuiFile.GetAddMeteoStation(StationName.Trim.ToUpper, StationName, iShape, myShape.Area)
                If myShape.Extents.xMax < 10 Then
                    'deze shapefile staat waarschijnlijk niet in RD-coordinaten!
                    Throw New Exception("Error in functie WriteNSL_POLY. X-coordinaat van de shape valt buiten het verwachte bereik. Controleer of de shapefile wel RD-projectie heeft.")
                End If
            Next
            Me.Setup.Log.AddMessage("Neerslagstations behorende bij shapes zijn met succes weggeschreven.")

            outputFile = TempDir & "\" & "MB_NSL." & FileExtension

            Xmin = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.xMin
            Xmax = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.xMax
            Ymin = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.yMin
            Ymax = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.yMax


            'initialize bui-file
            Dim StartDate As New DateTime(Left(FDate, 4), Right(Left(FDate, 6), 2), Right(Left(FDate, 8), 2))
            Dim EndDate As New DateTime(Left(TDate, 4), Right(Left(TDate, 6), 2), Right(Left(TDate, 8), 2))
            EndDate = EndDate.AddDays(1) 'adds one day to the end date because the last date is INCLUDED
            Dim ts As New TimeSpan(1, 0, 0)
            BuiFile.InitializeRecords(StartDate, EndDate, ts)

            'decide the pre- and post 2018 dates
            Dim FDatePre2019 As Integer, TDatePre2019 As Integer      'until 2019 we are dealing with the 'old' Meteobase rasters
            Dim FDatePost2019 As Integer, TDatePost2019 As Integer    'starting 1 jan 2019 we are dealing with the reanalysis grids by KNMI
            If Not GetDates(FDate, TDate, FDatePre2019, TDatePre2019, FDatePost2019, TDatePost2019) Then Throw New Exception("Error setting start and end dates for WIWB query. Please contact the meteobase team at info@meteobase.nl.")

            'now construct the API call, using the shapefile's extent
            'If Not WIWB.GetRasters("Meteobase.Precipitation", "P", Xmin, Ymin, Xmax, Ymax, FDate, TDate, "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving rasterdata from API.")

            If FDatePre2019 > 0 AndAlso TDatePre2019 > 0 Then
                Me.Setup.Log.AddMessage("Processing pre-january 2019 data.")
                If Not WIWB.DownloadRasters("Meteobase.Precipitation", "P", Xmin, Ymin, Xmax, Ymax, FDatePre2019, TDatePre2019, "geotiff", ZipFilePathPre2019) Then Throw New Exception("Error retrieving rasterdata from API.")
                'If Not ExtractZIP(ZipFilePathPre2019, CurDate, TempResultsDir, MetaFileContent, False, "AAIGrid") Then Throw New Exception("Error extracting data received from WIWB server.")
            End If

            If FDatePost2019 > 0 AndAlso TDatePost2019 > 0 Then
                Me.Setup.Log.AddMessage("Processing post-january 2019 data.")
                If Not WIWB.DownloadRasters("Knmi.International.Radar.Composite.Final.Reanalysis", "P", Xmin, Ymin, Xmax, Ymax, FDatePost2019, TDatePost2019, "geotiff", ZipFilePathPost2019) Then Throw New Exception("Error retrieving rasterdata from API.")
                'If Not ExtractZIP(ZipFilePathPost2019, CurDate, TempResultsDir, MetaFileContent, True, "AAIGrid") Then Throw New Exception("Error extracting data received from WIWB server.")
            End If

            'now that the zipfile has been downloaded, read each individual file
            Dim ZipFilePath As String
            Dim WarpFromKNMI As Boolean
            Dim i As Integer = -1

            For EraIdx = 1 To 2
                If EraIdx = 1 Then
                    ZipFilePath = ZipFilePathPre2019
                    WarpFromKNMI = False                 'the pre-2019 grids are already in RD-coordinates
                Else
                    ZipFilePath = ZipFilePathPost2019
                    WarpFromKNMI = True                  'the grids need to be warped from KNMI to RD before aggregating by polygon
                    'Dim sfWarped As New MapWinGIS.Shapefile
                    'Dim KNMIProj As New GeoProjection
                    'Dim reprojectedCount As Integer
                    'KNMIProj.ImportFromProj4("+proj=stere +lat_0=90 +lat_ts=60 +lon_0=0 +k=1 +x_0=0 +y_0=0 +a=6378.14 +b=6356.75")
                    'sfWarped = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Reproject(KNMIProj, reprojectedCount)
                End If

                Me.Setup.Log.AddMessage("Processing content of " & ZipFilePath)
                If System.IO.File.Exists(ZipFilePath) Then
                    Dim zipMS As New MemoryStream()
                    Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                        Me.Setup.Log.AddMessage(ZipFilePath & " contains " & zipReceived.Entries.Count & " entries.")

                        For Each ZipEntry As ZipEntry In zipReceived.Entries
                            i += 1
                            ZipEntry.Extract(zipMS)
                            'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                            If System.IO.File.Exists(unzipFile) Then System.IO.File.Delete(unzipFile)
                            Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                            zipMS.WriteTo(file)
                            file.Close()
                            zipMS.Seek(0, SeekOrigin.Begin)

                            'if necessary warp the grid to RD co-ordinates
                            Dim unzipFileRD As String = Replace(unzipFile, ".tif", "_RD.tif",,, CompareMethod.Text)
                            If System.IO.File.Exists(unzipFileRD) Then System.IO.File.Delete(unzipFileRD)
                            If WarpFromKNMI Then
                                Me.Setup.Log.AddMessage("Transforming " & unzipFile & " to " & unzipFileRD)
                                Dim args As String = "-overwrite -s_srs " & Chr(34) & "+proj=stere +lat_0=90 +lat_ts=60 +lon_0=0 +k=1 +x_0=0 +y_0=0 +a=6378.14 +b=6356.75" & Chr(34) & " -t_srs EPSG:28992"
                                Me.Setup.Log.AddMessage("Warping raster " & unzipFile & " from KNMI coordinates, using gdalwarp with arguments: " & args)
                                ut = New MapWinGIS.Utils
                                If Not ut.GDALWarp(unzipFile, unzipFileRD, args) Then Throw New Exception("Unable to warp precipitation grid to RD coordinates: " & unzipFile)
                            Else
                                Me.Setup.Log.AddMessage("Copying " & unzipFile & " to " & unzipFileRD)
                                System.IO.File.Copy(unzipFile, unzipFileRD)
                            End If

                            'now let's do some gridpolystatistics
                            Dim myGrid As New MapWinGIS.Grid
                            ut = New MapWinGIS.Utils
                            If Not myGrid.Open(unzipFileRD) Then
                                Throw New Exception("Could not open grid file.")
                            Else
                                'walk through each meteo station and calculate the mean evapotranspiration value
                                For j = 0 To BuiFile.MeteoStations.MeteoStations.Count - 1
                                    myStation = BuiFile.MeteoStations.MeteoStations.Values(j)

                                    'retrieve the mean value for each of the shapes that are represented by this station
                                    'using those mean values we will derive the composite mean value for the station
                                    Dim TotalArea As Double = 0
                                    Dim TotalValue As Double = 0

                                    For k = 0 To myStation.ShapeIndices.Count - 1
                                        myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(myStation.ShapeIndices(k))
                                        TotalArea += myShape.Area
                                        If ut.GridStatisticsForPolygon(myGrid, myGrid.Header, myGrid.Extents, myShape, myGrid.Header.NodataValue, mean, min, max) Then
                                            TotalValue += mean * myShape.Area
                                        End If
                                    Next

                                    'calculate the mean value for all polygons involved
                                    If TotalArea > 0 Then
                                        BuiFile.Values(i, j) = Math.Round(TotalValue / TotalArea, 2) 'the weighted mean value for all underlying shapes
                                    Else
                                        BuiFile.Values(i, j) = 0
                                    End If
                                Next
                            End If
                            ut = Nothing
                            myGrid.Close()

                        Next
                    End Using
                End If
            Next

            Me.Setup.Log.AddMessage("Neerslagbestand werd met succes gevuld met rasterdata.")

            If FileExtension.Trim.ToUpper = "BUI" Then
                BuiFile.WriteBUI(ResultsBUIFile)
                Me.Setup.Log.AddMessage("Bui-file voor SOBEK werd met succes geschreven.")
                FileCollectionNSL.Add(ResultsBUIFile)
                Me.Setup.Log.AddMessage("Bui-file voor SOBEK werd toegevoegd aan de bestandscollectie.")
            ElseIf FileExtension.Trim.ToUpper = "CSV" Then
                BuiFile.WriteAsCSV(ResultsBUIFile)
                Me.Setup.Log.AddMessage("CSV-file werd met succes geschreven.")
                FileCollectionNSL.Add(ResultsBUIFile)
                Me.Setup.Log.AddMessage("CSV-file werd toegevoegd aan de bestandscollectie.")
            End If
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Close()

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Er is een fout opgetreden in WriteNSL_POLY van de klasse clsWIWBRasterData.")
            Return False
        End Try

    End Function



    Public Function WriteEVP_POLY(WIWBDataSourceCode As String, WIWBVariableCode As String, FileNameBase As String, ByRef FileCollection As List(Of String)) As Boolean
        'Author: Siebe Bosch
        'Date: 10 july 2020
        'Description: writes file from Meteobase rasterdata, containing SATDATA 3.0 actual evapotranspiration
        Dim sf As New MapWinGIS.Shapefile
        Dim WIWB As New clsWIWB_API(Me.Setup)
        Dim CurProjection As MapWinGIS.GeoProjection
        Dim ResultsEVPFile As String
        Dim EVPDir As String = RasterDir & "\FileNameBase"
        Dim MeteoFile As New clsMeteoFile(Me.Setup)
        Dim myStation As clsMeteoStation = Nothing
        Dim StationName As String
        Dim myShape As Shape
        Dim i As Long = -1
        Dim CurDate As DateTime
        Dim mean As Double, min As Double, max As Double
        Dim ZipFilePath As String = TempDir & "\WIWB_RAW.zip"
        Dim unzipFile As String = TempDir & "\tmp.tif"
        Dim FileExtension As String = ""

        Try
            'initialize the shapefile
            Me.Setup.GISData.SubcatchmentShapeFile.Initialize()
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf = New MapWinGIS.Shapefile

            'set the output format
            Select Case FORMAAT
                Case Is = "CSV"
                    FileExtension = "csv"
                Case Is = "SOBEK"
                    FileExtension = "evp"
            End Select

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
                Me.Setup.Log.AddMessage("Map voor tussentijdse bestanden is met succes aangemaakt.")
            End If

            'create a path for the resulting evp-file
            ResultsEVPFile = TempDir & "\" & FileNameBase & "." & FileExtension
            If File.Exists(ResultsEVPFile) Then File.Delete(ResultsEVPFile)

            'merge all shapes into one big shape since SOBEK only supports one uniform evaporation value
            If Not sf.Open(ShapeFile) Then Throw New Exception("Kon de shapefile niet openen.")
            Me.Setup.PassAreaShape(sf) 'passes on the area shapefile to 
            'Me.Setup.GISData.SubcatchmentShapeFile.MergeAllShapes()    'siebe: disabled this on 10-7-2020
            Me.Setup.Log.AddMessage("Shapefile kon met succes worden geopend")

            If Not Me.Setup.GISData.SubcatchmentShapeFile.findAreaIDFieldIdx(ShapeField) Then Throw New Exception("Fout: kon veld met gebiedsID " & ShapeField & " niet vinden in shapefile.")
            Me.Setup.Log.AddMessage("Geldig Shapeveld werd gevonden.")

            'every shape gets its own meteorological station
            For iShape = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(iShape)
                StationName = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.CellValue(Me.Setup.GISData.SubcatchmentShapeFile.SubcatchmentIDFieldIdx, iShape).ToString
                'Debug.Print("Station name of shape " & iShape & " is " & StationName)
                MeteoFile.GetAddMeteoStation(StationName.Trim.ToUpper, StationName, iShape, myShape.Area)
            Next
            Me.Setup.Log.AddMessage("Neerslagstations behorende bij shapes zijn met succes weggeschreven.")


            'assign RD_New projection anyhow because some dialects are not recognized by MapWindow
            CurProjection = New MapWinGIS.GeoProjection
            CurProjection.ImportFromProj4("+proj=sterea +lat_0=52.15616055555555 +lon_0=5.38763888888889 +k=0.9999079 +x_0=155000 +y_0=463000 +ellps=bessel +units=m +no_defs")
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.GeoProjection = CurProjection

            'set the min/max extents for our query. This will save a lot of bandwidth
            Xmin = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.xMin
            Xmax = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.xMax
            Ymin = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.yMin
            Ymax = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Extents.yMax

            'initialize the Meteo file
            Dim StartDate As New Date(Left(FDate, 4), Right(Left(FDate, 6), 2), Right(Left(FDate, 8), 2)) 'note: if the message Cannot provide the value: host value not found appears when debuggin, try adding .ToString
            Dim EndDate As New Date(Left(TDate, 4), Right(Left(TDate, 6), 2), Right(Left(TDate, 8), 2))
            EndDate = EndDate.AddDays(1)                'we add one day since our data set includes both the from AND to date. If we don't do this we won't get results for the last timestep
            Dim ts As New TimeSpan(24, 0, 0)            'evaporation comes with 24h increments
            MeteoFile.SetStartDate(StartDate)
            MeteoFile.SetEndDate(EndDate)
            MeteoFile.SetTimestepSize(ts)
            Dim nTs As Integer = MeteoFile.CountTimesteps
            ReDim MeteoFile.Values(0 To nTs - 1, 0 To MeteoFile.MeteoStations.MeteoStations.Count - 1)

            'now construct the API call, using the shapefile's extent
            If Not WIWB.DownloadRasters(WIWBDataSourceCode, WIWBVariableCode, Xmin, Ymin, Xmax, Ymax, Convert.ToInt64(Format(StartDate, "yyyyMMdd")), Convert.ToInt64(Format(EndDate, "yyyyMMdd")), "geotiff", ZipFilePath) Then Throw New Exception("Error retrieving SATDATA rasterdata from API.")

            'now that the zipfile has been downloaded, read each individual file
            Dim zipMS As New MemoryStream()
            Using zipReceived As ZipFile = Ionic.Zip.ZipFile.Read(ZipFilePath)
                For Each ZipEntry As ZipEntry In zipReceived.Entries
                    i += 1
                    CurDate = StartDate.AddDays(i)           'this is a little tricky! Better would be to extract the date/time from the file
                    ZipEntry.Extract(zipMS)

                    'write the zipfile entry to a temporary file on the local drive, inside a subdirectory that represents the ordernumber
                    If System.IO.File.Exists(unzipFile) Then System.IO.File.Delete(unzipFile)
                    Dim file As New FileStream(unzipFile, FileMode.Create, FileAccess.Write)
                    zipMS.WriteTo(file)
                    file.Close()
                    zipMS.Seek(0, SeekOrigin.Begin)

                    'now let's do some gridpolystatistics
                    Dim myGrid As New MapWinGIS.Grid
                    Dim ut As New MapWinGIS.Utils
                    If Not myGrid.Open(unzipFile) Then
                        Throw New Exception("Could not open grid file.")
                    Else
                        'walk through each meteo station and calculate the mean evapotranspiration value
                        For j = 0 To MeteoFile.MeteoStations.MeteoStations.Count - 1
                            myStation = MeteoFile.MeteoStations.MeteoStations.Values(j)

                            'retrieve the mean value for each of the shapes that are represented by this station
                            'using those mean values we will derive the composite mean value for the station
                            Dim TotalArea As Double = 0
                            Dim TotalValue As Double = 0

                            For k = 0 To myStation.ShapeIndices.Count - 1
                                myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(myStation.ShapeIndices(k))
                                TotalArea += myShape.Area
                                If ut.GridStatisticsForPolygon(myGrid, myGrid.Header, myGrid.Extents, myShape, myGrid.Header.NodataValue, mean, min, max) Then
                                    TotalValue += mean * myShape.Area
                                End If
                            Next

                            'calculate the mean value for all polygons involved
                            If TotalArea > 0 Then
                                MeteoFile.Values(i, j) = Math.Round(TotalValue / TotalArea, 2) 'the weighted mean value for all underlying 
                            Else
                                MeteoFile.Values(i, j) = 0
                            End If
                        Next
                    End If
                    ut = Nothing
                    myGrid.Close()
                Next
            End Using

            Me.Setup.Log.AddMessage("Verdampingsbestand werd met succes gevuld met SATDATA Actuele Evapotranspiration-rasterdata.")

            If FORMAAT = "CSV" Then
                'in CSV format we can really write the result for each polygon to a separate column
                MeteoFile.WriteAsCSV(ResultsEVPFile)
                Me.Setup.Log.AddMessage("Verdampingsfile SAT DATA 3.0 werd met succes geschreven.")
                FileCollection.Add(ResultsEVPFile)
                Me.Setup.Log.AddMessage("EVP-file met SAT DATA 3.0 in CSV-formaat werd toegevoegd aan de bestandscollectie.")
            ElseIf FORMAAT = "SOBEK" Then
                'in SOBEK only one timeseries for evaporation is accepted. Therefore we must aggregate by weighing factor
                MeteoFile.WriteEVP(ResultsEVPFile)
                Me.Setup.Log.AddMessage("Verdampingsfile SATDATA werd met succes geschreven.")
                FileCollection.Add(ResultsEVPFile)
                Me.Setup.Log.AddMessage("EVP-file met SATDATA voor SOBEK werd toegevoegd aan de bestandscollectie.")
            End If

            sf.Close()
            Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Close()

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Er is een fout opgetreden in WriteEVT_ACT_POLY van de klasse clsWIWBRasterData.")
            Return False
        End Try

    End Function



    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

    Private Sub WriteExplanatory(ByVal Path As String)
        Using myWriter As New StreamWriter(Path)
            If FORMAAT.Trim.ToUpper = "MODFLOW" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens geschikt voor import in MODFLOW.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: mm/h")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink, geschikt voor import in MODFLOW.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith, geschikt voor import in MODFLOW.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT.Trim.ToUpper = "SIMGRO" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens geschikt voor import in SIMGRO.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT.Trim.ToUpper = "ASCII" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens in Arc/Info-formaat.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Met gebruikmaking van de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/uur")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Met gebruikmaking van de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Met gebruikmaking van de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Met gebruikmaking van de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .ASC (Arc/Info-raster)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Met gebruikmaking van de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT.Trim.ToUpper = "NETCDF" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens in NetCDF-formaat.")
                    myWriter.WriteLine("Bestandsformaat: .nc")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/uur")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink.")
                    myWriter.WriteLine("Bestandsformaat: .nc (NetCDF)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith.")
                    myWriter.WriteLine("Bestandsformaat: .nc (NetCDF)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .nc (NetCDF)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .nc (NetCDF)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT.Trim.ToUpper = "HDF5" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens in HDF5-formaat.")
                    myWriter.WriteLine("Bestandsformaat: .h5 (HDF5)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/uur")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink.")
                    myWriter.WriteLine("Bestandsformaat: .h5 (HDF5)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith.")
                    myWriter.WriteLine("Bestandsformaat: .h5 (HDF5)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .h5 (HDF5)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0.")
                    myWriter.WriteLine("Bestandsformaat: .h5 (HDF5)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl via de API van WIWB")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte rasters: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: 1 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT = "SOBEK" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens in SOBEK-formaat.")
                    myWriter.WriteLine("Bestandsformaat: .BUI (SOBEK-neerslagbestand)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: mm/uur")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink, geschikt voor SOBEK")
                    myWriter.WriteLine("Bestandsformaat: .EVP (SOBEK-verdampingsbestand)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith, geschikt voor SOBEK")
                    myWriter.WriteLine("Bestandsformaat: .EVP (SOBEK-verdampingsbestand)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0, geschikt voor SOBEK")
                    myWriter.WriteLine("Bestandsformaat: .EVP (SOBEK-verdampingsbestand)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0")
                    myWriter.WriteLine("Bestandsformaat: .EVP (SOBEK-verdampingsbestand)")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            ElseIf FORMAAT = "CSV" Then
                If NSL Then
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens, geaggregeerd naar polygonen, in CSV-formaat.")
                    myWriter.WriteLine("Bestandsformaat: .CSV (generiek tekstbestand) met ; als scheidingsteken")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 uur")
                    myWriter.WriteLine("Eenheid gegevens: mm/uur")
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If MAKKINK Then
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Makkink")
                    myWriter.WriteLine("Bestandsformaat: .CSV (generiek tekstbestand) met ; als scheidingsteken")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If PM Then
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingsgegevens volgens Penman-Monteith")
                    myWriter.WriteLine("Bestandsformaat: .CSV (generiek tekstbestand) met ; als scheidingsteken")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_ACTUAL Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Actuele evapotranspiratie volgens SATDATA 3.0")
                    myWriter.WriteLine("Bestandsformaat: .CSV (generiek tekstbestand) met ; als scheidingsteken")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
                If EVT_SHORTAGE Then
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Author: VanderSat, 2020                                                                                 ")
                    myWriter.WriteLine("License: Creative Commons license CC BY-NC-SA www.creativecommons.org/licenses/by-nc-sa/4.0/legalcode.nl")
                    myWriter.WriteLine("Acknowledgement: contains modified ESA Copernicus Sentinel data 2015-2020                               ")
                    myWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
                    myWriter.WriteLine("Verdampingstekort (Epot - Eact) volgens SATDATA 3.0")
                    myWriter.WriteLine("Bestandsformaat: .CSV (generiek tekstbestand) met ; als scheidingsteken")
                    myWriter.WriteLine("Datum gegenereerd:" & Today)
                    myWriter.WriteLine("Gegenereerd door: www.meteobase.nl")
                    myWriter.WriteLine("Projectie: RD new (Amersfoort, rijksdriehoekstelsel)")
                    myWriter.WriteLine("Gegevens van datum: " & FDate)
                    myWriter.WriteLine("Gegevens tot datum: " & TDate)
                    myWriter.WriteLine("X-coordinaat linksonder: " & Xmin)
                    myWriter.WriteLine("Y-coordinaat linksonder: " & Ymin)
                    myWriter.WriteLine("X-coordinaat rechtsboven: " & Xmax)
                    myWriter.WriteLine("Y-coordinaat rechtsboven: " & Ymax)
                    myWriter.WriteLine("Tijdstapgrootte reeksen: 1 etmaal")
                    myWriter.WriteLine("Eenheid gegevens: mm/etmaal")
                    myWriter.WriteLine("-----------------------------------------------------------------------")
                    myWriter.WriteLine("")
                End If
            End If
        End Using
    End Sub

    Public Function WriteZIP(ByVal myPath As String) As Boolean
        'schrijf de resultaten naar een zipfile
        Dim ZipFile = New Ionic.Zip.ZipFile
        Dim Explain As String = TempDir & "\LEESMIJ.TXT"

        Try
            If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
            If FileCollectionNSL.Count + FileCollectionMAK.Count + FileCollectionPEN.Count + FileCollectionEVT.Count + FileCollectionSHO.Count > 65000 Then
                Me.Setup.Log.AddWarning("Waarschuwing: meer dan 65000 bestanden in de ZIP-file. Dit kan leiden tot een ongeldig bestand.")
            Else
                Me.Setup.Log.AddMessage("Writing " & FileCollectionNSL.Count & " precipitation files, " & FileCollectionMAK.Count & " Makkink files, " & FileCollectionPEN.Count & " Penman files, " & FileCollectionEVT.Count & " evapotranspiration-files and " & FileCollectionSHO.Count & " evaporation shortage files to zipfile.")
            End If

            'schrijf een toelichting op de gegenereerde resultaten
            WriteExplanatory(Explain)
            If System.IO.File.Exists(Explain) Then ZipFile.AddFile(Explain, "")

            'schrijf nu de gegenereerde bestanden naar de zip-file
            For Each metaFile In FileCollectionMETA
                If System.IO.File.Exists(metaFile) Then ZipFile.AddFile(metaFile, "")
            Next
            If FileCollectionMETA.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionMETA.Count & " meta-bestanden geschreven naar zip-file.")

            For Each outputFile In FileCollectionNSL
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "NEERSLAG")
            Next
            If FileCollectionNSL.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionNSL.Count & " neerslagbestanden geschreven naar zip-file.")

            For Each outputFile In FileCollectionMAK
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "MAKKINK")
            Next
            If FileCollectionMAK.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionMAK.Count & " Makkink verdampingsbestanden geschreven naar zip-file.")

            For Each outputFile In FileCollectionPEN
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "PENMAN")
            Next
            If FileCollectionPEN.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionPEN.Count & " Penman verdampingsbestanden geschreven naar zip-file.")

            For Each outputFile In FileCollectionEVT
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "EVAPOTRANSPIRATIE")
            Next
            If FileCollectionEVT.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionPEN.Count & " bestanden met SAT DATA actuele verdamping geschreven naar zip-file.")

            For Each outputFile In FileCollectionSHO
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile, "VERDAMPINGSTEKORT")
            Next
            If FileCollectionSHO.Count > 0 Then Me.Setup.Log.AddMessage(FileCollectionPEN.Count & " bestanden met SAT DATA verdampingstekort geschreven naar zip-file.")

            Me.Setup.Log.AddMessage("Bezig met opslaan zip-file naar " & myPath)
            ZipFile.Save(myPath)

            For Each outputFile In FileCollectionNSL
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next
            For Each outputFile In FileCollectionMAK
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next
            For Each outputFile In FileCollectionPEN
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next
            For Each outputFile In FileCollectionEVT
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next
            For Each outputFile In FileCollectionSHO
                If System.IO.File.Exists(outputFile) Then System.IO.File.Delete(outputFile)
            Next

            If Not System.IO.File.Exists(myPath) Then Throw New Exception("Zipfile could not be written: " & myPath)
            Setup.Log.AddMessage("Rasterdata met succes gecomprimeerd.")

            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error in sub WriteZIP of class clsMBRasterData.")
            Me.Setup.Log.AddError(ex.Message)
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

End Class

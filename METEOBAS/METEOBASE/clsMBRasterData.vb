Option Explicit On

Imports METEOBAS.General
Imports Ionic.Zip
Imports MapWinGIS
Imports System.IO

Public Class clsMBRasterData

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

    'naar welk formaat exporteren?
    Public Aggregate24H As Boolean 'aggregeren naar etmaalsom?
    Public FORMAAT As String      'ASCII/MODFLOW/SIMGRO/NETCDF/WAGMOD/SOBEK/CSV

    'bestelgegevens
    Public SessionID As Integer    'sessieID
    Public OrderNum As Integer     'bestelnummer

    'lokale instellingen
    Public TempDir As String       'directory voor tijdelijke bestanden
    Public UploadDir As String     'directory voor uitgepakte shapefiles
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
    Friend FileCollectionMETA As New List(Of String)  'een verzameling met paden voor meta-bestanden die ook mee moeten in de zip
    Friend tmpFileCollection As New List(Of String)   'een tijdelijke verzameling met paden naar files

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

            'set the environment variable for GDAL
            Environment.SetEnvironmentVariable("GDAL_DATA", GDALToolsDir & "\gdal-data")
            Me.Setup.Log.AddMessage("GDAL Environment Variabele werd met succes ingesteld.")

            'round the co-ordinates to whole kilometres
            If FORMAAT = "ASCII" OrElse FORMAAT = "NETCDF" OrElse FORMAAT = "MODFLOW" OrElse FORMAAT = "SIMGRO" Then
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

            'En hier gaan we onder de motorkap de bestelling verwerken
            If NSL Then
                If Not writeNSL() Then Throw New Exception("Er is een fout opgetreden in de subroutine writeNSL")
                Me.Setup.Log.AddMessage("Neerslag werd met succes weggeschreven.")
            End If
            If PM Then
                If Not WritePM() Then Throw New Exception("Er is een fout opgetreden in de subroutine writePM")
                Me.Setup.Log.AddMessage("Penman-evaporatie werd met succes weggeschreven.")
            End If
            If MAKKINK Then
                If Not writeMAK() Then Throw New Exception("Er is een fout opgetreden in de subroutine writeMAK")
                Me.Setup.Log.AddMessage("Makkinkevaporatie werd met succes weggeschreven.")
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

            'if the zip-file exists
            If System.IO.File.Exists(ShapeFileZIP) Then
                myZip = New ZipFile(ShapeFileZIP)
                If myZip.CheckZip(ShapeFileZIP) Then
                    myZip.ExtractAll(UploadDir, ExtractExistingFileAction.OverwriteSilently)

                    Dim di = New IO.DirectoryInfo(UploadDir)
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
            body &= "Uw bestelling staat klaar in de download-directory van Meteobase. Klik op de onderstaande link om hem op te halen." & vbCrLf
            body &= DownloadURL & ZipFile & vbCrLf

            If Me.Setup.Log.Warnings.Count > 0 Then
                body &= vbCrLf
                body &= "Eventuele waarschuwingen:" & vbCrLf
                For Each myStr As String In Me.Setup.Log.Warnings
                    body &= myStr & vbCrLf
                Next

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
            body &= "Aangeboden door STOWA | www.stowa.nl" & vbCrLf
            body &= vbCrLf
            body &= "Mogelijk gemaakt door" & vbCrLf
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
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts Makkink grids to ASC-format
        Dim myYear As Integer, myPath As String, outputFile As String, myArgs As String
        Dim i As Long, k As Long, r As Long, c As Long
        Dim XMLFile As String = ""
        Dim ResultsZIPFile As String 'zipfile met het resultaat
        Dim MAKDir As String = RasterDir & "\MAK_RD"
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim OutputDir As String
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        Dim myRaster As clsASCIIGrid
        ResultsZIPFile = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-Apache\Php\apache\www\meteobase\downloads\" & ZipFile

        OutputDir = TempDir & "\MAK"
        If Not System.IO.Directory.Exists(OutputDir) Then
            System.IO.Directory.CreateDirectory(OutputDir)
        End If

        'walk through all dates
        For i = FDate To TDate
            If Setup.GeneralFunctions.DateIntIsValid(i) Then
                myYear = Left(i, 4)
                myPath = MAKDir & "\" & myYear & "\EVAP_MK_" & Str(i).Trim & ".nc"
                If System.IO.File.Exists(myPath) Then
                    myArgs = ""

                    outputFile = OutputDir & "\MAK_" & Str(i).Trim & ".asc"         'set the output filename
                    FileCollectionMAK.Add(outputFile)                                                         'add the filename to the collection
                    myArgs = "-of AAIGrid -a_nodata -999 -b 1 " & "-projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile
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

        'unit conversion. original PM-data is in 0.01 mm/d
        'MODFLOW requires mm/d
        'SIMGRO requires mm/d PLUS: does not accept nodata-values. Replace by 0
        If FORMAAT.Trim.ToUpper = "MODFLOW" OrElse FORMAAT.Trim.ToUpper = "SIMGRO" Then
            Me.Setup.Log.AddMessage("Exportformaat = " & FORMAAT & ". conversie van eenheden wordt uitgevoerd.")
            For Each myPath In FileCollectionMAK
                myRaster = New clsASCIIGrid(Me.Setup)
                If Not myRaster.Read(myPath) Then
                    Me.Setup.Log.AddError("Kon Makkink-raster niet lezen voor eenheidsconversie: " & myPath & " eenheid blijft 0.1 mm/etmaal")
                Else
                    If FORMAAT.Trim.ToUpper = "MODFLOW" Then
                        For r = 0 To myRaster.rows - 1
                            For c = 0 To myRaster.cols - 1
                                If Not myRaster.cells(r, c) = myRaster.nodataval Then
                                    myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                End If
                            Next
                        Next
                    ElseIf FORMAAT.Trim.ToUpper = "SIMGRO" Then
                        For r = 0 To myRaster.rows - 1
                            For c = 0 To myRaster.cols - 1
                                If myRaster.cells(r, c) = myRaster.nodataval Then
                                    myRaster.cells(r, c) = 0
                                Else
                                    myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                End If
                            Next
                        Next
                    End If
                End If
                If Not myRaster.Write(myPath) Then Me.Setup.Log.AddError("Kon Makkink-raster niet schrijven na eenheidsconversie: " & myPath)
            Next
        End If

        Return True
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
                Case Is = "ASCII"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "MODFLOW"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "SIMGRO"
                    If Not WriteNSL_ASC() Then Return False
                Case Is = "SOBEK"
                    If Not WriteNSL_POLY("BUI") Then Return False
                Case Is = "CSV"
                    If Not WriteNSL_POLY("CSV") Then Return False
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
                If Not WritePM_ASC() Then Return False
            Case Is = "NETCDF"
                If Not WritePM_NC() Then Return False
            Case Is = "MODFLOW"
                If Not WritePM_ASC() Then Return False
            Case Is = "SIMGRO"
                If Not WritePM_ASC() Then Return False
            Case Is = "SOBEK"
                If Not WritePM_SBK() Then Return False
            Case Is = "CSV"
                If Not WritePM_CSV() Then Return False
        End Select

        Return True

    End Function


    Public Function writeMAK() As Boolean
        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)

        If Not System.IO.Directory.Exists(TempDir) Then
            System.IO.Directory.CreateDirectory(TempDir)
        End If

        Select Case FORMAAT.Trim.ToUpper
            Case Is = "ASCII"
                If Not WriteMAK_ASC() Then Return False
            Case Is = "NETCDF"
                If Not WriteMAK_NC() Then Return False
            Case Is = "MODFLOW"
                If Not WriteMAK_ASC() Then Return False
            Case Is = "SIMGRO"
                If Not WriteMAK_ASC() Then Return False
            Case Is = "SOBEK"
                If Not WriteMAK_SBK() Then Return False
            Case Is = "CSV"
                If Not WriteMAK_CSV() Then Return False
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

    Public Function WritePM_ASC()
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts Penman-Monteith grids to ASC-format
        Dim myYear As Integer, myPath As String, outputFile As String, myArgs As String
        Dim i As Long, k As Long, r As Long, c As Long
        Dim XMLFile As String = ""
        Dim ResultsZIPFile As String 'zipfile met het resultaat
        Dim PMDir As String = RasterDir & "\PM_RD"
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim OutputDir As String
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        Dim myRaster As clsASCIIGrid
        ResultsZIPFile = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-Apache\Php\apache\www\meteobase\downloads\" & ZipFile

        OutputDir = TempDir & "\PM"
        If Not System.IO.Directory.Exists(OutputDir) Then
            System.IO.Directory.CreateDirectory(OutputDir)
        End If

        'walk through all dates
        For i = FDate To TDate
            If Setup.GeneralFunctions.DateIntIsValid(i) Then
                myYear = Left(i, 4)
                myPath = PMDir & "\" & myYear & "\EVAP_PM_" & Str(i).Trim & ".nc"
                If System.IO.File.Exists(myPath) Then
                    myArgs = ""

                    outputFile = OutputDir & "\PM_" & Str(i).Trim & ".asc"          'set the output filename
                    FileCollectionPEN.Add(outputFile)                                                             'add the filename to the collection
                    myArgs = "-of AAIGrid -a_nodata -999 -b 1 " & "-projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile
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
                Else
                    Me.Setup.Log.AddError("Bestand niet gevonden: " & myPath)
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

        'unit conversion. original PM-data is in 0.01 mm/d
        'MODFLOW requires mm/d
        'SIMGRO requires mm/d PLUS: does not accept nodata-values. Replace by 0
        If FORMAAT.Trim.ToUpper = "MODFLOW" OrElse FORMAAT.Trim.ToUpper = "SIMGRO" Then
            Me.Setup.Log.AddMessage("Exportformaat = " & FORMAAT & ". conversie van eenheden wordt uitgevoerd.")
            For Each myPath In FileCollectionPEN
                myRaster = New clsASCIIGrid(Me.Setup)
                If Not myRaster.Read(myPath) Then
                    Me.Setup.Log.AddError("Kon Penman-raster niet lezen voor eenheidsconversie: " & myPath & " eenheid blijft 0.1 mm/etmaal")
                Else
                    If FORMAAT.Trim.ToUpper = "MODFLOW" Then
                        For r = 0 To myRaster.rows - 1
                            For c = 0 To myRaster.cols - 1
                                If Not myRaster.cells(r, c) = myRaster.nodataval Then
                                    myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                End If
                            Next
                        Next
                    ElseIf FORMAAT.Trim.ToUpper = "SIMGRO" Then
                        For r = 0 To myRaster.rows - 1
                            For c = 0 To myRaster.cols - 1
                                If myRaster.cells(r, c) = myRaster.nodataval Then
                                    myRaster.cells(r, c) = 0
                                Else
                                    myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                End If
                            Next
                        Next
                    End If
                End If
                If Not myRaster.Write(myPath) Then Me.Setup.Log.AddError("Kon Penman-raster niet schrijven na eenheidsconversie: " & myPath)
            Next
        End If

        Return True
    End Function

    Public Function WriteNSL_NC() As Boolean
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts precipitation grids to NetCDF-format in RD-coordinates
        Dim myYear As Integer, myPath As String, outputFile As String, myArgs As String
        Dim i As Long, k As Long
        Dim XMLFile As String = ""
        Dim NSLDir As String = RasterDir & "\NSL_RD"
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim OutputDir As String = TempDir & "\NSL"
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        Dim myUtils As New MapWinGIS.Utils

        Try
            If Not System.IO.Directory.Exists(OutputDir) Then
                System.IO.Directory.CreateDirectory(OutputDir)
            End If

            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then
                    myYear = Left(i, 4)
                    myPath = NSLDir & "\" & myYear & "\NDB2_" & Str(i).Trim & ".nc"
                    If System.IO.File.Exists(myPath) Then
                        myArgs = ""
                        outputFile = OutputDir & "\" & Str(OrderNum).Trim & "_NSL_" & Str(i).Trim & ".nc"

                        'gdal_translate will export each band as a separate file (unfortunately)
                        'For j = 1 To 24
                        '   FileCollectionNSL.Add(Left(outputFile, outputFile.Length - 3) & "_" & Format(j, "00") & ".nc")
                        'Next

                        'EDIT 1-10-2017 by SIEBE BOSCH: it seems like gdal_translate has changed into writing multiple bands
                        'when performing a translate (yeeeyy!!!)
                        FileCollectionNSL.Add(outputFile)

                        myArgs = "-of NetCDF -sds -projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile  'exports all 24 subdatasets at once
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
                If SW.ElapsedMilliseconds > 600000 Then Done = True '10 minutes processing time is the maximume
            End While

            SW.Stop()
            SW.Reset()
            Return True
        Catch ex As Exception
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


    Public Function WriteNSL_ASC() As Boolean
        'Author: Siebe Bosch
        'Date: 22-5-2013
        'Description: This function converts precipitation grids to ASC-format
        Dim myPath As String, outputFile As String, myArgs As String
        Dim i As Long, j As Long, k As Long, DayNum As Integer = 0
        Dim NSLDir As String = RasterDir & "\NSL_RD"
        Dim OutputDir As String = TempDir & "\NSL"
        Dim XMLFile As String = ""
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim myRaster As clsASCIIGrid
        Dim myYear As Integer
        Dim myUtils As New MapWinGIS.Utils
        Dim r As Long, c As Long
        Dim nUnfinished As Integer = 5
        Dim mySW As New System.Diagnostics.Stopwatch
        Dim GDALExePath As String = GDALToolsDir & "\gdal_translate.exe"
        Dim MetaFilePath As String = TempDir & "\Mete_grid.inp"
        Dim MetaFileContent As String = ""

        Try

            If Not System.IO.File.Exists(GDALExePath) Then
                Throw New Exception("Executable does not exist: " & GDALExePath)
            End If

            If Not System.IO.Directory.Exists(OutputDir) Then
                System.IO.Directory.CreateDirectory(OutputDir)
            End If



            '---------------OLD OLD OLD OLD OLD OLD--------------------------------'
            'dit gedeelte is van voor de tijd dat we gedurende de bestelling warpten.
            'Dim NSLDir As String = RasterDir & "\NSL_KNMI"
            'First we'll have to warp the NetCDF-files to the RD co-ordinate system
            'at the same time we'll crop them to the right co-ordinates
            'If Not WriteNSL_NC() Then Throw New Exception("Fout in routine WriteNSL_NC. Kon neerslagrasters niet bijsnijden.")
            '---------------OLD OLD OLD OLD OLD OLD--------------------------------'
            'siebe()
            'now use the results of WriteNSL_NC as input for this function
            tmpFileCollection = New List(Of String) 'FileCollectionNSL OLD OLD OLD
            FileCollectionNSL = New List(Of String)
            FileCollectionMETA = New List(Of String)

            '---------------NEW NEW NEW NEW NEW NEW--------------------------------'
            'this routine only collects the valid input files for further processing
            'and then export all 24 bands in one go (option -sds)
            DayNum = -1
            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then
                    DayNum += 1
                    myYear = Left(i, 4)
                    myPath = NSLDir & "\" & myYear & "\NDB2_" & Str(i).Trim & ".nc"
                    Me.Setup.Log.AddMessage("Converting from NetCDF to ASCII: " & myPath)
                    If System.IO.File.Exists(myPath) Then

                        For j = 1 To 24
                            outputFile = OutputDir & "\NSL_" & Str(i).Trim & "_" & Format(j, "00") & ".asc"
                            myArgs = "-of AAIGrid -b " & j & " -projwin " & Xmin & " " & Ymax & " " & Xmax & " " & Ymin & " " & myPath & " " & outputFile  'exports all 24 subdatasets at once
                            ProcessCollection.Add(Process.Start(GDALExePath, myArgs))
                            FileCollectionNSL.Add(outputFile)

                            'hier schrijven we de metadata weg t.b.v. SIMGRO. Merk op dat we ook maar meteen een verwijzing naar verdamping maken, al weten we natuurlijk niet of die ook wordt opgevraagd.
                            Dim NSLFile As String = Setup.GeneralFunctions.FileNameFromPath(outputFile)
                            Dim MAKFile As String = Replace(Left(NSLFile, NSLFile.Length - 7) & ".ASC", "NSL", "MAK")
                            MetaFileContent &= Format(DayNum + (j - 1) / 24, "0.0000000000") & ", " & Format(myYear, "0000") & "," & Chr(34) & NSLFile & Chr(34) & "," & Chr(34) & MAKFile & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & "," & Chr(34) & "NoValue" & Chr(34) & vbCrLf

                            'niet meer dan vier gelijktijdige shells open hebben
                            mySW = New System.Diagnostics.Stopwatch
                            mySW.Start()
                            nUnfinished = 5
                            While nUnfinished > 4
                                nUnfinished = 0
                                For k = 0 To ProcessCollection.Count - 1
                                    If Not ProcessCollection(k).HasExited Then nUnfinished += 1
                                Next
                                If mySW.ElapsedMilliseconds > 30000 Then nUnfinished = 0 'veligheidsklep
                            End While
                        Next
                    Else
                        Me.Setup.Log.AddError("File does not exist: " & myPath)
                    End If
                End If
            Next

            'before we start renaming the files make sure all processes have been completed
            mySW.Reset()
            mySW.Start()
            nUnfinished = 5
            While nUnfinished > 0
                nUnfinished = 0
                For k = 0 To ProcessCollection.Count - 1
                    If Not ProcessCollection(k).HasExited Then nUnfinished += 1
                Next
                If mySW.ElapsedMilliseconds > 3600000 Then Throw New Exception("Error in subroutine WriteNSL_ASC. Waiting for all raster conversions to complete took too long.")
            End While
            mySW.Stop()

            'also wait until the write protection on all files has been released
            For Each myFile As String In FileCollectionNSL
                i = 0
                While Me.Setup.GeneralFunctions.FileInUse(myFile)
                    i += 1
                    Threading.Thread.Sleep(1000)
                    If i > 600 Then Throw New Exception("Error in subroutine WriteNSL_ASC. Waiting for the write protection to be revoked took too long: " & myFile)
                End While
            Next

            If FORMAAT.Trim.ToUpper = "SIMGRO" Then
                Using simWriter As New System.IO.StreamWriter(MetaFilePath)
                    simWriter.Write(MetaFileContent)
                End Using
                FileCollectionMETA.Add(MetaFilePath)
            End If

            'if MODFLOW or SIMGRO, the grid values need adjustment (unit conversion)
            'Original rasters: 0.01 mm/h
            'MODFLOW: 1/100 mm/h to mm/h
            'SIMGRO: 1/100 mm/h to mm/d. Plus: nodata-values not allowed
            If FORMAAT.Trim.ToUpper = "MODFLOW" OrElse FORMAAT.Trim.ToUpper = "SIMGRO" Then
                Me.Setup.Log.AddMessage("Exportformaat = " & FORMAAT & ". conversie van eenheden wordt uitgevoerd.")
                For Each myPath In FileCollectionNSL
                    myRaster = New clsASCIIGrid(Me.Setup)
                    If Not myRaster.Read(myPath) Then
                        Me.Setup.Log.AddError("Kon rasterbestand niet openen voor bewerking. Eenheid van de rasterdata blijft 1/10 mm/h")
                        Exit For
                    Else
                        Select Case FORMAAT.Trim.ToUpper
                            Case Is = "MODFLOW"
                                For r = 0 To myRaster.rows - 1
                                    For c = 0 To myRaster.cols - 1
                                        If Not myRaster.cells(r, c) = myRaster.nodataval Then
                                            myRaster.cells(r, c) = myRaster.cells(r, c) / 100
                                        End If
                                    Next
                                Next
                            Case Is = "SIMGRO"
                                For r = 0 To myRaster.rows - 1
                                    For c = 0 To myRaster.cols - 1
                                        If myRaster.cells(r, c) = myRaster.nodataval Then
                                            myRaster.cells(r, c) = 0
                                        Else
                                            myRaster.cells(r, c) = myRaster.cells(r, c) / 100 * 24
                                        End If
                                    Next
                                Next
                        End Select
                        If Not myRaster.Write(myPath) Then Me.Setup.Log.AddError("Could not write grid file after unit conversion: " & myPath)
                    End If
                Next
            End If

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WriteNSL_POLY(ByVal FileExtension As String) As Boolean
        'Author: Siebe Bosch
        'Date: 22 april 2013
        'Description: writes a SOBEK .bui file from Meteobase rasterdata
        'Exportformat can either be BUI or CSV
        'Dim sf As New MapWinGIS.Shapefile
        'Dim NewProjection As MapWinGIS.GeoProjection
        Dim CurProjection As MapWinGIS.GeoProjection
        Dim outputFile As String
        Dim ResultsBUIFile As String, InputFile As String
        Dim NSLDir As String = RasterDir & "\NSL_RD"
        Dim myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer
        Dim BuiFile As New clsMeteoFile(Me.Setup)
        Dim iShape As Integer, k As Long = -1, i As Long, j As Long
        Dim myStation As clsMeteoStation = Nothing
        Dim myShape As Shape
        Dim header As GridHeader
        Dim extents As Extents
        Dim noData As Double
        Dim mean As Double, min As Double, max As Double
        Dim StationName As String

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

            'every unique shape ID gets its own meteorological station
            For iShape = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(iShape)
                StationName = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.CellValue(Me.Setup.GISData.SubcatchmentShapeFile.SubcatchmentIDFieldIdx, iShape).ToString
                BuiFile.GetAddMeteoStation(StationName.Trim.ToUpper, StationName, iShape, myShape.Area)
            Next
            Me.Setup.Log.AddMessage("Neerslagstations behorende bij shapes zijn met succes weggeschreven.")

            outputFile = TempDir & "\" & "MB_NSL." & FileExtension

            'now construct the API call, using the shapefile's extent
            For i = FDate To TDate

                If Setup.GeneralFunctions.DateIntIsValid(i) Then
                    myYear = Left(i, 4)
                    myMonth = Left(Right(i, 4), 2)
                    myDay = Right(i, 2)

                    InputFile = NSLDir & "\" & myYear & "\NDB2_" & Str(i).Trim & ".nc"

                    Dim mygrid = New MapWinGIS.Grid
                    If Not mygrid.Open(InputFile, GridDataType.UnknownDataType, False) Then
                        Me.Setup.Log.AddError("Kon neerslagbestand " & InputFile & " niet openen.")
                    Else
                        header = mygrid.Header
                        extents = mygrid.Extents
                        noData = CDbl(header.NodataValue)
                        mean = InlineAssignHelper(min, InlineAssignHelper(max, 0.0))
                        For j = 1 To mygrid.NumBands

                            k += 1
                            myHour = j - 1

                            'initialize bui-file
                            If i = FDate AndAlso j = 1 Then
                                Dim StartDate As New DateTime(Left(FDate, 4), Right(Left(FDate, 6), 2), Right(Left(FDate, 8), 2))
                                Dim EndDate As New DateTime(Left(TDate, 4), Right(Left(TDate, 6), 2), Right(Left(TDate, 8), 2))
                                EndDate = EndDate.AddDays(1) 'adds one day to the end date because the last date is INCLUDED
                                Dim ts As New TimeSpan(mygrid.NumBands / 24, 0, 0)
                                BuiFile.InitializeRecords(StartDate, EndDate, ts)
                            End If

                            If mygrid.OpenBand(j) Then
                                For iShape = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                                    myShape = Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(iShape)
                                    Dim ut = New MapWinGIS.Utils()
                                    If ut.GridStatisticsForPolygon(mygrid, header, extents, myShape, noData, mean, min, max) Then
                                        BuiFile.Values(k, iShape) = Math.Round(mean / 100, 2)
                                    Else
                                        BuiFile.Values(k, iShape) = Math.Round(-0.01, 2) 'error clipping raster by polygon (probably all nodata values, so apply -0.10 precip)
                                    End If
                                    ut = Nothing
                                Next
                            Else
                                Throw New Exception("Fout: kon band " & j & " van neerslagraster " & mygrid.Filename & " niet openen.")
                            End If
                        Next
                    End If
                    mygrid.Close()
                    mygrid = Nothing
                    'GC.Collect() 'force garbage collection after every grid
                    If i / 5 = Math.Round(i / 5, 0) Then GC.Collect() 'force garbage collection every 5 grids
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
            Me.Setup.Log.AddError("Er is een fout opgetreden in WriteNSL_POLY van de klasse clsMBRasterData.")
            Return False
        End Try

    End Function

    Public Function WriteMAK_SBK() As Boolean
        'Author: Siebe Bosch
        'Date: 22 may 2013
        'Description: writes a SOBEK .evp file containing Makkink Evaporation data from Meteobase rasterdata
        Dim sf As New MapWinGIS.Shapefile
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim ResultsEVPFile As String, InputFile As String
        Dim MAKDir As String = RasterDir & "\MAK_RD"
        Dim myYear As Integer, myMonth As Integer, myDay As Integer
        Dim EvpFile As New clsMeteoFile(Me.Setup)
        Dim myDate As Date
        Dim header As GridHeader
        Dim extents As Extents
        Dim noData As Double
        Dim ut As New MapWinGIS.Utils()
        Dim mean As Double, min As Double, max As Double
        Dim mygrid As Grid
        Dim i As Long

        Try

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If

            'create a path for the resulting bui-file
            ResultsEVPFile = TempDir & "\MB_MAK.EVP"

            'merge all shapes into one big shape since SOBEK only supports one uniform evaporation value
            If Not sf.Open(ShapeFile) Then Throw New Exception("Kon de shapefile niet openen.")
            Me.Setup.PassAreaShape(sf)
            Me.Setup.GISData.SubcatchmentShapeFile.MergeAllShapes()

            'walk through all dates and process the corresponding grid files
            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then

                    myYear = Left(i, 4)
                    myMonth = Left(Right(i, 4), 2)
                    myDay = Right(i, 2)
                    myDate = New DateTime(myYear, myMonth, myDay, 0, 0, 0)

                    InputFile = MAKDir & "\" & myYear & "\EVAP_MK_" & Str(i).Trim & ".nc"
                    If System.IO.File.Exists(InputFile) Then

                        mygrid = New Grid
                        If Not mygrid.Open(InputFile, GridDataType.UnknownDataType, False) Then
                            Dim settings As New GlobalSettings
                            Throw New Exception("Kon neerslagbestand " & InputFile & " niet openen.")
                        Else
                            header = mygrid.Header
                            extents = mygrid.Extents
                            noData = CDbl(header.NodataValue)
                            mean = InlineAssignHelper(min, InlineAssignHelper(max, 0.0))

                            ut = New MapWinGIS.Utils
                            For j = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                                If ut.GridStatisticsForPolygon(mygrid, header, extents, Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(j), noData, mean, min, max) Then
                                    EvpFile.Values(i, j) = mean / 100
                                End If
                            Next

                            ut = Nothing
                            mygrid.Close()
                        End If
                    End If
                End If
            Next

            'write the resulting .EVP file and add it to the file collection for later compression
            EvpFile.WriteEVP(ResultsEVPFile)
            FileCollectionMAK.Add(ResultsEVPFile)
            sf.Close()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WriteMAK_CSV() As Boolean
        'Author: Siebe Bosch
        'Date: 22 may 2013
        'Description: writes a CSV file containing Makkink Evaporation data from Meteobase rasterdata
        'aggregated by polygon
        Dim sf As New MapWinGIS.Shapefile
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim ResultsCSVFile As String, InputFile As String
        Dim MAKDir As String = RasterDir & "\MAK_RD"
        Dim myYear As Integer, myMonth As Integer, myDay As Integer
        Dim EvpFile As New clsMeteoFile(Me.Setup)
        Dim myDate As Date
        Dim header As GridHeader
        Dim extents As Extents
        Dim noData As Double
        Dim ut As MapWinGIS.Utils
        Dim mean As Double, min As Double, max As Double
        Dim mygrid As Grid
        Dim i As Long

        Try

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If

            'create a path for the resulting bui-file
            ResultsCSVFile = TempDir & "\MB_MAK.CSV"

            'merge all shapes into one big shape since SOBEK only supports one uniform evaporation value
            If Not sf.Open(ShapeFile) Then Throw New Exception("Kon de shapefile niet openen.")
            Me.Setup.PassAreaShape(sf)
            Me.Setup.GISData.SubcatchmentShapeFile.MergeAllShapes()

            'walk through all dates and process the corresponding grid files
            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then

                    myYear = Left(i, 4)
                    myMonth = Left(Right(i, 4), 2)
                    myDay = Right(i, 2)
                    myDate = New DateTime(myYear, myMonth, myDay, 0, 0, 0)

                    InputFile = MAKDir & "\" & myYear & "\EVAP_MK_" & Str(i).Trim & ".nc"
                    If System.IO.File.Exists(InputFile) Then

                        mygrid = New Grid
                        If Not mygrid.Open(InputFile, GridDataType.UnknownDataType, False) Then
                            Dim settings As New GlobalSettings
                            Throw New Exception("Kon verdampingsbestand " & InputFile & " niet openen.")
                        Else
                            header = mygrid.Header
                            extents = mygrid.Extents
                            noData = CDbl(header.NodataValue)
                            mean = InlineAssignHelper(min, InlineAssignHelper(max, 0.0))

                            ut = New MapWinGIS.Utils
                            For j = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                                If ut.GridStatisticsForPolygon(mygrid, header, extents, Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(j), noData, mean, min, max) Then
                                    EvpFile.Values(i, j) = mean / 100
                                Else
                                    EvpFile.Values(i, j) = -999
                                End If
                            Next
                            ut = Nothing

                            mygrid.Close()
                        End If
                    End If
                End If
            Next

            'write the resulting .CSV file and add it to the file collection for later compression
            EvpFile.WriteAsCSV(ResultsCSVFile)
            FileCollectionMAK.Add(ResultsCSVFile)
            sf.Close()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function
    Public Function WritePM_SBK() As Boolean
        'Author: Siebe Bosch
        'Date: 22 may 2013
        'Description: writes a SOBEK .evp file containing Penman Monteith Evaporation data from Meteobase rasterdata
        Dim sf As New MapWinGIS.Shapefile
        Dim ProcessCollection As New List(Of System.Diagnostics.Process)
        Dim ResultsEVPFile As String, InputFile As String
        Dim PMDir As String = RasterDir & "\PM_RD"
        Dim myYear As Integer, myMonth As Integer, myDay As Integer
        Dim EvpFile As New clsMeteoFile(Me.Setup)
        Dim myDate As Date
        Dim header As GridHeader
        Dim extents As Extents
        Dim noData As Double
        Dim ut As New MapWinGIS.Utils()
        Dim mean As Double, min As Double, max As Double
        Dim mygrid As Grid
        Dim i As Long

        Try

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If

            'create a path for the resulting bui-file
            ResultsEVPFile = TempDir & "\MB_PM.EVP"

            'merge all shapes into one big shape since SOBEK only supports one uniform evaporation value
            If Not sf.Open(ShapeFile) Then Throw New Exception("Kon de shapefile niet openen.")
            Me.Setup.PassAreaShape(sf)
            Me.Setup.GISData.SubcatchmentShapeFile.MergeAllShapes()

            'walk through all dates and process the corresponding grid files
            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then

                    myYear = Left(i, 4)
                    myMonth = Left(Right(i, 4), 2)
                    myDay = Right(i, 2)
                    myDate = New DateTime(myYear, myMonth, myDay, 0, 0, 0)

                    InputFile = PMDir & "\" & myYear & "\EVAP_PM_" & Str(i).Trim & ".nc"
                    If System.IO.File.Exists(InputFile) Then

                        mygrid = New Grid
                        If Not mygrid.Open(InputFile, GridDataType.UnknownDataType, False) Then
                            Dim settings As New GlobalSettings
                            Throw New Exception("Kon verdampingsbestand " & InputFile & " niet openen.")
                        Else
                            header = mygrid.Header
                            extents = mygrid.Extents
                            noData = CDbl(header.NodataValue)
                            mean = InlineAssignHelper(min, InlineAssignHelper(max, 0.0))

                            ut = New MapWinGIS.Utils
                            For j = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                                If ut.GridStatisticsForPolygon(mygrid, header, extents, Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(j), noData, mean, min, max) Then
                                    EvpFile.Values(i, j) = mean / 100
                                End If
                            Next
                            mygrid.Close()
                        End If
                    End If
                End If
            Next

            'write the resulting .EVP file and add it to the file collection for later compression
            EvpFile.WriteEVP(ResultsEVPFile)
            FileCollectionPEN.Add(ResultsEVPFile)
            sf.Close()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function WritePM_CSV() As Boolean
        'Author: Siebe Bosch
        'Date: 22 may 2013
        'Description: writes a SOBEK .evp file containing Penman Monteith Evaporation data from Meteobase rasterdata
        Dim sf As New MapWinGIS.Shapefile
        Dim ResultsCSVFile As String, InputFile As String
        Dim PMDir As String = RasterDir & "\PM_RD"
        Dim myYear As Integer, myMonth As Integer, myDay As Integer
        Dim EvpFile As New clsMeteoFile(Me.Setup)
        Dim myDate As Date
        Dim header As GridHeader
        Dim extents As Extents
        Dim noData As Double
        Dim ut As MapWinGIS.Utils
        Dim mean As Double, min As Double, max As Double
        Dim mygrid As Grid
        Dim i As Long

        Try

            'set the temporary output directory or create it
            If Not System.IO.Directory.Exists(TempDir) Then
                System.IO.Directory.CreateDirectory(TempDir)
            End If

            'create a path for the resulting bui-file
            ResultsCSVFile = TempDir & "\MB_PM.CSV"

            'merge all shapes into one big shape since SOBEK only supports one uniform evaporation value
            If Not sf.Open(ShapeFile) Then Throw New Exception("Kon de shapefile niet openen.")
            Me.Setup.PassAreaShape(sf)
            Me.Setup.GISData.SubcatchmentShapeFile.MergeAllShapes()

            'walk through all dates and process the corresponding grid files
            For i = FDate To TDate
                If Setup.GeneralFunctions.DateIntIsValid(i) Then

                    myYear = Left(i, 4)
                    myMonth = Left(Right(i, 4), 2)
                    myDay = Right(i, 2)
                    myDate = New DateTime(myYear, myMonth, myDay, 0, 0, 0)

                    InputFile = PMDir & "\" & myYear & "\EVAP_PM_" & Str(i).Trim & ".nc"
                    If System.IO.File.Exists(InputFile) Then

                        mygrid = New Grid
                        If Not mygrid.Open(InputFile, GridDataType.UnknownDataType, False) Then
                            Dim settings As New GlobalSettings
                            Throw New Exception("Kon verdampingsbestand " & InputFile & " niet openen.")
                        Else
                            header = mygrid.Header
                            extents = mygrid.Extents
                            noData = CDbl(header.NodataValue)
                            mean = InlineAssignHelper(min, InlineAssignHelper(max, 0.0))

                            ut = New MapWinGIS.Utils
                            For j = 0 To Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.NumShapes - 1
                                If ut.GridStatisticsForPolygon(mygrid, header, extents, Me.Setup.GISData.SubcatchmentShapeFile.PolySF.sf.Shape(j), noData, mean, min, max) Then
                                    EvpFile.Values(i, j) = mean / 100
                                Else
                                    EvpFile.Values(i, j) = -999
                                End If
                            Next
                            ut = Nothing

                            mygrid.Close()
                            mygrid = Nothing
                        End If
                    End If
                End If
            Next

            'write the resulting .EVP file and add it to the file collection for later compression
            EvpFile.WriteAsCSV(ResultsCSVFile)
            FileCollectionPEN.Add(ResultsCSVFile)
            sf.Close()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

    Private Sub WriteExplanatory(ByVal Path As String)
        Using myWriter As New StreamWriter(Path)
            If FORMAAT = "MODFLOW" Then
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

            ElseIf FORMAAT = "SIMGRO" Then
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
            ElseIf FORMAAT = "ASCII" Then
                If NSL Then
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("Neerslagradargegevens in Arc/Info-formaat.")
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
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/uur")
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
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
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
                    myWriter.WriteLine("Eenheid gegevens: 0.01 mm/etmaal")
                    myWriter.WriteLine("------------------------------------------------------")
                    myWriter.WriteLine("")
                End If

            ElseIf FORMAAT = "NETCDF" Then
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
            End If
        End Using
    End Sub

    Public Function WriteZIP(ByVal myPath As String) As Boolean
        'schrijf de resultaten naar een zipfile
        Dim ZipFile = New Ionic.Zip.ZipFile
        Dim Explain As String = TempDir & "\LEESMIJ.TXT"

        Try
            If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
            If FileCollectionNSL.Count + FileCollectionMAK.Count + FileCollectionPEN.Count > 65000 Then
                Me.Setup.Log.AddWarning("Waarschuwing: meer dan 65000 bestanden in de ZIP-file. Dit kan leiden tot een ongeldig bestand.")
            End If

            'schrijf een toelichting op de gegenereerde resultaten
            WriteExplanatory(Explain)
            If System.IO.File.Exists(Explain) Then ZipFile.AddFile(Explain)

            'schrijf nu de gegenereerde bestanden naar de zip-file
            For Each metaFile In FileCollectionMETA
                If System.IO.File.Exists(metaFile) Then ZipFile.AddFile(metaFile)
            Next
            For Each outputFile In FileCollectionNSL
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile)
            Next
            For Each outputFile In FileCollectionMAK
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile)
            Next
            For Each outputFile In FileCollectionPEN
                If System.IO.File.Exists(outputFile) Then ZipFile.AddFile(outputFile)
            Next
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

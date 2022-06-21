Option Explicit On

Imports Ionic.Zip
Imports MapWinGIS
Imports GemBox.Spreadsheet
Imports Npgsql
Imports Newtonsoft.Json.Linq

Public Class clsWIWBBasisData

    'Lokale variabelen
    Public FDate As Integer        'startdatum voor de te selecteren dataset
    Public TDate As Integer        'einddatum voor de te seleecteren dataset
    Public Etmaal As Boolean       'tijdbasis (Etmaal/Uur)
    Public Stations As New clsMeteoStations(Me.Setup)

    'welke data exporteren?
    Public NSL As Boolean          'Neerslagintensiteit exporteren?
    Public MAKKINK As Boolean      'Makkink exporteren?

    'bestelgegevens
    Public SessionID As Integer    'sessieID
    Public OrderNum As Integer     'bestelnummer

    'lokale instellingen
    Public Naam As String          'naam van de aanvrager
    Public MailAdres As String     'mailadres van de aanvrager
    Public DownloadURL As String   'downloaddirectory vanuit het oogpunt van de gebruiker
    Public DownloadDIR As String   'downloaddirectory vanuit het oogpunt van de server

    'terugkoppeling naar de aanvrager per e-mail
    Friend GoodMail As clsEmail                       'the e-mail with good news
    Friend BadMail As clsEmail                        'the e-mail with bad news
    Friend ExcelFile As String                        'filename of the resulting Excel-file

    Private Setup As General.clsSetup

    Public Sub New(ByRef mySetup As General.clsSetup)
        Setup = mySetup
    End Sub

    Public Sub getHourStationNames()
        Dim conn As NpgsqlConnection
        Dim comm As NpgsqlCommand
        Dim myMS As clsMeteoStation
        Dim myNummer As Integer, myName As String

        conn = New NpgsqlConnection
        'standard: User ID=root;Password=myPassword;Host=localhost;Port=5432;Database=myDataBase; Pooling=true;Min Pool Size=0;Max Pool Size=100;Connection Lifetime=0;
        conn.ConnectionString = "User ID=postgres;Password=sch1aap;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
        conn.Open()

        comm = New NpgsqlCommand("SELECT * FROM data.stations WHERE timevalue = 'uur'", conn)
        Dim dr As Npgsql.NpgsqlDataReader
        dr = comm.ExecuteReader()

        While dr.Read()
            myNummer = dr(2)
            myName = dr(1)

            myMS = GetStationByNumber(myNummer)

            If Not myMS Is Nothing Then
                myMS.Name = myName
            End If

        End While
        dr.Dispose()

        'sluit de verbinding met de database
        conn.Close()
        If conn.State = System.Data.ConnectionState.Open Then conn.Close()
        conn.Dispose()

    End Sub

    Public Sub getStationNames(ByVal Etmaal As Boolean)
        '-----------------------------------------------------------------------------------
        ' 7-12-2017
        ' gets a list of all stations from the WIWB API for a given timestep value
        '----------------------------------------------------------------------------------


        Dim conn As NpgsqlConnection
        Dim comm As NpgsqlCommand
        Dim myMS As clsMeteoStation
        Dim myNummer As Integer, myName As String

        conn = New NpgsqlConnection
        'standard: User ID=root;Password=myPassword;Host=localhost;Port=5432;Database=myDataBase; Pooling=true;Min Pool Size=0;Max Pool Size=100;Connection Lifetime=0;
        conn.ConnectionString = "User ID=postgres;Password=sch1aap;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
        conn.Open()

        If Etmaal Then
            comm = New NpgsqlCommand("SELECT * FROM data.stations WHERE timevalue = 'dag'", conn)
        Else
            comm = New NpgsqlCommand("SELECT * FROM data.stations WHERE timevalue = 'uur'", conn)
        End If
        Dim dr As Npgsql.NpgsqlDataReader
        dr = comm.ExecuteReader()

        While dr.Read()
            myNummer = dr(2)
            myName = dr(1)

            myMS = GetStationByNumber(myNummer)
            If Not myMS Is Nothing Then
                myMS.Name = myName
            End If

        End While
        dr.Dispose()

        'sluit de verbinding met de database
        conn.Close()
        If conn.State = System.Data.ConnectionState.Open Then conn.Close()
        conn.Dispose()

    End Sub


    Public Function GetStationByNumber(ByVal myNum As Integer) As clsMeteoStation

        'look in the existing stations and see if it's already there
        Dim myStation As clsMeteoStation
        For Each myStation In Stations.MeteoStations.Values
            If myStation.Number = myNum Then
                Return myStation
            End If
        Next

        'not found so return nothing
        Return Nothing

    End Function

    Public Function GetAddStationByNumber(ByVal myNum As String) As clsMeteoStation

        'look in the existing stations and see if it's already there
        Dim myStation As clsMeteoStation
        For Each myStation In Stations.MeteoStations.Values
            If myStation.Number = myNum.Trim Then
                Return myStation
            End If
        Next

        'not found, so add it and return it
        myStation = New clsMeteoStation(Me.Setup)
        myStation.Number = myNum.Trim
        Stations.MeteoStations.Add(myStation.Number.Trim, myStation)
        Return myStation

    End Function

    Public Function Build() As Boolean

        'this routine queries the meteobase database for basis data
        'and writes them to an excel file

        ' If using GemBox.Spreadsheet Professional, put your serial key below.
        ' Otherwise, if you are using GemBox.Spreadsheet Free, comment out the 
        ' following line (Free version doesn't have SetLicense method). 
        SpreadsheetInfo.SetLicense("EVIG-1Y89-FYME-DPUJ")
        Dim oExcel As ExcelFile = New ExcelFile
        Dim worksheets As ExcelWorksheetCollection = oExcel.Worksheets

        Try

            If Etmaal Then
                'verwerkt een bestelling voor etmaalstations
                If Not processNeerslagBasisDaily(worksheets) Then
                    Throw New Exception("Error processing Daily precipitation.")
                Else
                    Me.Setup.Log.AddMessage("Neerslag etmaalstations is met succes weggeschreven.")
                    ExcelFile = "Bestelling_" & Trim(SessionID) & "_" & Str(OrderNum).Trim & "_etmaalstations.xlsx"
                    oExcel.Save(DownloadDIR & "\" & ExcelFile)
                End If
            Else
                'verwerkt een bestelling voor uurstations
                ExcelFile = "Bestelling_" & Trim(SessionID) & "_" & Str(OrderNum).Trim & "_uurstations.xlsx"
                If NSL Then
                    Call processNeerslagBasisHourly(worksheets)
                    Me.Setup.Log.AddMessage("Neerslag uurstations is met succes weggeschreven.")
                End If
                If MAKKINK Then
                    Call processMakkinkBasisDaily(worksheets)
                    Me.Setup.Log.AddMessage("Verdamping uurstations is met succes weggeschreven.")
                End If
                oExcel.Save(DownloadDIR & ExcelFile, SaveOptions.XlsxDefault)
            End If
            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Console.WriteLine("An error occurred in sub Write of class clsMBBasisData.")
            Return False
        End Try

    End Function

    Public Function processNeerslagBasisDaily(ByVal worksheets As ExcelWorksheetCollection) As Boolean
        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim myResponse As List(Of String) = WIWB.GetTimeSeries("Knmi.IrisValidated", "P", Stations, FDate, TDate, "hydronet.csv", 1440)
            Dim myRecord As String, i As Long, X As Double, Y As Double
            Dim myMS As clsMeteoStation, myMeteoVal As clsMeteoValue
            Dim DateStr As String, LocStr As String, NumStr As String, ParStr As String, Value As Double
            Dim InData As Boolean = False
            Dim InLocations As Boolean = False

            For i = 0 To myResponse.Count - 1
                myRecord = myResponse(i)
                If Not InLocations AndAlso InStr(myRecord, "[Locations]") Then
                    InLocations = True
                    i += 1
                ElseIf InLocations Then
                    myRecord = myResponse(i)
                    If myRecord = "" Then
                        InLocations = False
                    Else
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        LocStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        X = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Y = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMS = Stations.GetByNumber(NumStr.Trim.ToUpper)
                        If Not myMS Is Nothing Then
                            myMS.ID = NumStr
                            myMS.Number = NumStr
                            myMS.Name = LocStr
                            myMS.X = X
                            myMS.Y = Y
                        End If
                    End If
                End If


                If Not InData AndAlso InStr(myRecord, "Date time") > 0 Then
                    InData = True
                ElseIf InData Then
                    myMeteoVal = New clsMeteoValue
                    DateStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                    If DateStr.Length >= 8 Then
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        ParStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Value = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMeteoVal.DateTimeVal = New DateTime(Left(DateStr, 4), Right(Left(DateStr, 7), 2), Right(Left(DateStr, 10), 2), Right(Left(DateStr, 13), 2), Right(Left(DateStr, 16), 2), Right(Left(DateStr, 19), 2))
                        myMeteoVal.ValueAdjusted = Value
                        myMeteoVal.ValueObserved = Value
                        myMeteoVal.ValueCorrected = Value
                        myMS = Stations.GetByKey(NumStr)
                        If Not myMS Is Nothing Then myMS.PrecipitationDaily.Add(Format(myMeteoVal.DateTimeVal, "yyyyMMddHHmmss"), myMeteoVal)
                    End If
                End If
            Next

            Me.Setup.Log.AddMessage("Processing daily precipitation complete.")
            Call writeEtmaalNeerslagToExcel(worksheets)
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing daily precipitation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try



    End Function

    Public Function processMakkinkBasisDaily(ByVal worksheets As ExcelWorksheetCollection) As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            Dim myResponse As List(Of String) = WIWB.GetTimeSeries("Knmi.Evaporation", "Evaporation", Stations, FDate, TDate, "hydronet.csv", 1440)
            Dim myRecord As String, i As Long, X As Double, Y As Double
            Dim myMS As clsMeteoStation, myMeteoVal As clsMeteoValue
            Dim DateStr As String, LocStr As String, NumStr As String, ParStr As String, Value As Double
            Dim InData As Boolean = False
            Dim InLocations As Boolean = False

            For i = 0 To myResponse.Count - 1
                myRecord = myResponse(i)
                If Not InLocations AndAlso InStr(myRecord, "[Locations]") Then
                    InLocations = True
                    i += 1
                ElseIf InLocations Then
                    myRecord = myResponse(i)
                    If myRecord = "" Then
                        InLocations = False
                    Else
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        LocStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        X = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Y = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMS = Stations.GetByNumber(NumStr.Trim.ToUpper)
                        If Not myMS Is Nothing Then
                            myMS.ID = NumStr
                            myMS.Number = NumStr
                            myMS.Name = LocStr
                            myMS.X = X
                            myMS.Y = Y
                        End If
                    End If
                End If


                If Not InData AndAlso InStr(myRecord, "Date time") > 0 Then
                    InData = True
                ElseIf InData Then
                    myMeteoVal = New clsMeteoValue
                    DateStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                    If DateStr.Length >= 8 Then
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        ParStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Value = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMeteoVal.DateTimeVal = New DateTime(Left(DateStr, 4), Right(Left(DateStr, 7), 2), Right(Left(DateStr, 10), 2), Right(Left(DateStr, 13), 2), Right(Left(DateStr, 16), 2), Right(Left(DateStr, 19), 2))
                        myMeteoVal.ValueAdjusted = Value
                        myMeteoVal.ValueObserved = Value
                        myMeteoVal.ValueCorrected = Value
                        myMS = Stations.GetByKey(NumStr)
                        If Not myMS Is Nothing Then myMS.EvaporationDaily.Add(Format(myMeteoVal.DateTimeVal, "yyyyMMddHHmmss"), myMeteoVal)
                    End If
                End If
            Next

            Me.Setup.Log.AddMessage("Processing evaporation complete.")
            Call writeEtmaalMakkinkToExcel(worksheets, "Basis.Makkink.Etmaal")
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing evaporation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function processNeerslagBasisHourly(ByVal worksheets As ExcelWorksheetCollection) As Boolean

        Try
            Dim WIWB As New clsWIWB_API(Me.Setup)
            'Dim myResponse As List(Of String) = WIWB.GetTimeSeries("Knmi.AwsTenMinutes", "P", Stations, FDate, TDate, "hydronet.csv", 60)
            Dim myResponse As List(Of String) = WIWB.GetTimeSeries("Knmi.Synops", "P", Stations, FDate, TDate, "hydronet.csv", 60)
            Dim myRecord As String, i As Long, X As Double, Y As Double
            Dim myMS As clsMeteoStation, myMeteoVal As clsMeteoValue
            Dim DateStr As String, LocStr As String, NumStr As String, ParStr As String, Value As Double
            Dim InData As Boolean = False
            Dim InLocations As Boolean = False

            For i = 0 To myResponse.Count - 1
                myRecord = myResponse(i)
                If Not InLocations AndAlso InStr(myRecord, "[Locations]") Then
                    InLocations = True
                    i += 1
                ElseIf InLocations Then
                    myRecord = myResponse(i)
                    If myRecord = "" Then
                        InLocations = False
                    Else
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        LocStr = Setup.GeneralFunctions.ParseString(myRecord, ",").Trim
                        X = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Y = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMS = Stations.GetByNumber(NumStr.Trim.ToUpper)
                        If Not myMS Is Nothing Then
                            myMS.ID = NumStr
                            myMS.Number = NumStr
                            myMS.Name = LocStr
                            myMS.X = X
                            myMS.Y = Y
                        End If
                    End If
                End If


                If Not InData AndAlso InStr(myRecord, "Date time") > 0 Then
                    InData = True
                ElseIf InData Then
                    myMeteoVal = New clsMeteoValue
                    DateStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                    If DateStr.Length >= 8 Then
                        NumStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        ParStr = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        Value = Setup.GeneralFunctions.ParseString(myRecord, ",")
                        myMeteoVal.DateTimeVal = New DateTime(Left(DateStr, 4), Right(Left(DateStr, 7), 2), Right(Left(DateStr, 10), 2), Right(Left(DateStr, 13), 2), Right(Left(DateStr, 16), 2), Right(Left(DateStr, 19), 2))
                        myMeteoVal.ValueAdjusted = Value
                        myMeteoVal.ValueObserved = Value
                        myMeteoVal.ValueCorrected = Value
                        myMS = Stations.GetByKey(NumStr)
                        If Not myMS Is Nothing Then myMS.PrecipitationHourly.Add(Format(myMeteoVal.DateTimeVal, "yyyyMMddHHmmss"), myMeteoVal)
                    End If
                End If
            Next

            Me.Setup.Log.AddMessage("Processing hourly precipitation complete.")
            Call writeBasisNeerslagUurToExcel(worksheets, "Basis.Neerslag.Uur")
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error processing hourly precipitation.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try


    End Function

    Public Sub writeEtmaalMakkinkToExcel(ByRef Worksheets As ExcelWorksheetCollection, ByVal SheetName As String)
        Dim r As Long, c As Long
        Dim ws As ExcelWorksheet
        ws = Worksheets.Add(SheetName)

        c = -4
        For Each myMS In Stations.MeteoStations.Values

            r = -1
            c += 4

            r += 1
            ws.Cells(r, c).Value = "Data ontsloten door:"
            ws.Cells(r, c + 1).Value = "www.meteobase.nl"
            ws.Cells(r, c + 2).Value = "via de API van WIWB"
            r += 1
            ws.Cells(r, c).Value = "Herkomst brongegevens:"
            ws.Cells(r, c + 1).Value = "KNMI, dataproduct nr.:"
            ws.Cells(r, c + 2).Value = "NL-CLIMATE-EV-DATA-D"
            r += 1
            ws.Cells(r, c).Value = "Naam station:"
            ws.Cells(r, c + 1).Value = myMS.Name
            r += 1
            ws.Cells(r, c).Value = "Nummer station:"
            ws.Cells(r, c + 1).Value = myMS.Number
            r += 1
            ws.Cells(r, c).Value = "Ligging station (RD):"
            ws.Cells(r, c + 1).Value = myMS.X
            ws.Cells(r, c + 2).Value = myMS.Y
            r += 1
            ws.Cells(r, c).Value = "Datumwaarde:"
            ws.Cells(r, c + 1).Value = "Meetwaarde [mm]:"

            For Each myMeteoVal In myMS.EvaporationDaily.Values
                r += 1
                ws.Cells(r, c).Value = myMeteoVal.DateTimeVal
                ws.Cells(r, c + 1).Value = myMeteoVal.ValueObserved
            Next
        Next

    End Sub

    Public Function writeEtmaalNeerslagToExcel(ByRef Worksheets As ExcelWorksheetCollection) As Boolean

        Try
            Dim r As Long, c As Long
            Dim ws As ExcelWorksheet
            ws = Worksheets.Add("Basis.Neerslag.Etmaal")
            If ws Is Nothing Then Throw New Exception("Error: Excel worksheet Basis.Neerslag.Etmaal could not be created.")

            c = -4
            For Each myMS In Stations.MeteoStations.Values

                r = -1
                c += 4

                r += 1
                ws.Cells(r, c).Value = "Data ontsloten door:"
                ws.Cells(r, c + 1).Value = "www.meteobase.nl"
                ws.Cells(r, c + 2).Value = "via de API van WIWB"
                r += 1
                ws.Cells(r, c).Value = "Herkomst brongegevens:"
                ws.Cells(r, c + 1).Value = "KNMI, dataproduct nr.:"
                ws.Cells(r, c + 2).Value = "NL-OBS-SURF-DECODED-1H"
                r += 1
                ws.Cells(r, c).Value = "Naam station:"
                ws.Cells(r, c + 1).Value = myMS.Name
                r += 1
                ws.Cells(r, c).Value = "Nummer station:"
                ws.Cells(r, c + 1).Value = myMS.Number
                r += 1
                ws.Cells(r, c).Value = "Ligging station (RD):"
                ws.Cells(r, c + 1).Value = myMS.X
                ws.Cells(r, c + 2).Value = myMS.Y
                r += 1
                ws.Cells(r, c).Value = "Datumwaarde:"
                ws.Cells(r, c + 1).Value = "Meetwaarde [mm]:"

                For Each myMeteoVal In myMS.PrecipitationDaily.Values
                    r += 1
                    ws.Cells(r, c).Value = myMeteoVal.DateTimeVal
                    ws.Cells(r, c + 1).Value = myMeteoVal.ValueObserved
                Next
            Next

            Me.Setup.Log.AddMessage("Neerslag etmaalstations is met succes naar Excel geschreven.")
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError("Error writing daily precipitation to Excel.")
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try


    End Function

    Public Function writeBasisNeerslagUurToExcel(ByRef WorkSheets As ExcelWorksheetCollection, ByVal SheetName As String) As Boolean
        Try
            Dim r As Long, c As Long

            Dim ws As ExcelWorksheet
            ws = WorkSheets.Add(SheetName)

            c = -5
            For Each myMS In Stations.MeteoStations.Values

                r = -1
                c += 5

                r += 1
                ws.Cells(r, c).Value = "Data ontsloten via:"
                ws.Cells(r, c + 1).Value = "www.meteobase.nl"
                ws.Cells(r, c + 2).Value = "via de API van WIWB"
                r += 1
                ws.Cells(r, c).Value = "Herkomst brongegevens:"
                ws.Cells(r, c + 1).Value = "KNMI, dataproduct nr.:"
                ws.Cells(r, c + 2).Value = "NL-OBS-SURF-10M-EXT"
                r += 1
                ws.Cells(r, c).Value = "Naam station:"
                ws.Cells(r, c + 1).Value = myMS.Name
                r += 1
                ws.Cells(r, c).Value = "Nummer station:"
                ws.Cells(r, c + 1).Value = myMS.Number
                r += 1
                ws.Cells(r, c).Value = "Ligging station (RD):"
                ws.Cells(r, c + 1).Value = myMS.X
                ws.Cells(r, c + 2).Value = myMS.Y
                r += 1
                ws.Cells(r, c).Value = "Datumwaarde:"
                ws.Cells(r, c + 1).Value = "Meetwaarde [mm]:"
                For Each myMeteoVal In myMS.PrecipitationHourly.Values
                    r += 1
                    ws.Cells(r, c).Value = myMeteoVal.DateTimeVal
                    ws.Cells(r, c + 1).Value = myMeteoVal.ValueObserved
                Next
            Next
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Error in function writeBasisNeerslagUurToExcel.")
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
        body &= DownloadURL & ExcelFile & vbCrLf
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
        body &= "Resultatenbestand " & ExcelFile & vbCrLf
        body &= "Tijdsspanne: van=" & FDate.ToString & " tot=" & TDate.ToString & vbCrLf
        body &= "Etmaalstations: " & Etmaal
        body &= "Neerslag: " & NSL
        body &= "Makkink: " & MAKKINK
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
        If Not GoodMail.Send("info@meteobase.nl.nl", "Meteobase") Then
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

Option Explicit On

Imports Ionic.Zip
Imports MapWinGIS
Imports GemBox.Spreadsheet
Imports Npgsql
Imports MailKit
Imports MimeKit

Public Class clsMBBasisData

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

    Public Function GetAddStationByNumber(ByVal myNum As Integer) As clsMeteoStation

        'look in the existing stations and see if it's already there
        Dim myStation As clsMeteoStation
        For Each myStation In Stations.MeteoStations.Values
            If myStation.Number = myNum Then
                Return myStation
            End If
        Next

        'not found, so add it and return it
        myStation = New clsMeteoStation(Me.Setup)
        myStation.Number = myNum
        Stations.MeteoStations.Add(myStation.Number.ToString.Trim, myStation)
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
                Call processNeerslagBasisDaily(worksheets)
                Me.Setup.Log.AddMessage("Neerslag etmaalstations is met succes weggeschreven.")
                ExcelFile = "Bestelling_" & Trim(SessionID) & "_" & Str(OrderNum).Trim & "_etmaalstations.xlsx"
                oExcel.Save(DownloadDIR & "\" & ExcelFile)
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

    Public Sub processNeerslagBasisDaily(ByVal worksheets As ExcelWorksheetCollection)
        Dim conn As NpgsqlConnection
        Dim comm As NpgsqlCommand
        Dim myMeteoVal As clsMeteoValue
        Dim myMS As clsMeteoStation

        conn = New NpgsqlConnection
        'standard: User ID=root;Password=myPassword;Host=localhost;Port=5432;Database=myDataBase; Pooling=true;Min Pool Size=0;Max Pool Size=100;Connection Lifetime=0;
        conn.ConnectionString = "User ID=postgres;Password=sch1aap;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
        conn.Open()

        For Each myMS In Stations.MeteoStations.Values
            comm = New NpgsqlCommand("SELECT * FROM data.precipitation_daily WHERE datumveld >= " & FDate & " AND datumveld <= " & TDate & " AND station = " & myMS.Number.ToString.Trim & " ORDER BY datumveld", conn)
            Dim dr As Npgsql.NpgsqlDataReader
            dr = comm.ExecuteReader()

            While dr.Read()
                myMeteoVal = New clsMeteoValue
                myMeteoVal.Datum = dr(9)
                myMeteoVal.Tijd = 0
                myMeteoVal.DateTimeVal = New DateTime(Left(myMeteoVal.Datum, 4), Left(Right(myMeteoVal.Datum, 4), 2), Right(myMeteoVal.Datum, 2))
                myMeteoVal.ValueObserved = Val(dr(6))
                If myMeteoVal.ValueObserved < 0 Then
                    myMeteoVal.ValueCorrected = 0
                Else
                    myMeteoVal.ValueCorrected = myMeteoVal.ValueObserved / 10
                End If
                myMeteoVal.ValueAdjusted = myMeteoVal.ValueCorrected
                If Not dr(6).ToString.Trim = "" Then myMS.PrecipitationDaily.Add(myMeteoVal.Datum, myMeteoVal)

            End While
            dr.Dispose()
        Next

        'Hier komt de area reduction factor!
        ' 

        'sluit de verbinding met de database
        conn.Close()
        If conn.State = System.Data.ConnectionState.Open Then conn.Close()
        conn.Dispose()

        Call writeEtmaalNeerslagToExcel(worksheets)

    End Sub

    Public Sub processMakkinkBasisDaily(ByVal worksheets As ExcelWorksheetCollection)
        Dim conn As NpgsqlConnection
        Dim comm As NpgsqlCommand
        Dim myMeteoVal As clsMeteoValue
        Dim myQuery As String

        For Each myMS In Stations.MeteoStations.Values

            'myQuery = "SELECT datumveld, st" & myMS.Number.ToString.Trim & " FROM data.makkinkraw WHERE datumveld >= " & FDate & " AND datumveld <= " & TDate & " ORDER BY datumveld"
            myQuery = "SELECT * FROM data.evaporation_daily WHERE station = " & myMS.Number.ToString.Trim & " AND datumveld >= " & FDate & " AND datumveld <= " & TDate & " ORDER BY datumveld"

            conn = New NpgsqlConnection
            'standard: User ID=root;Password=myPassword;Host=localhost;Port=5432;Database=myDataBase; Pooling=true;Min Pool Size=0;Max Pool Size=100;Connection Lifetime=0;
            'conn.ConnectionString = "User ID=arend;Password=arend;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
            conn.ConnectionString = "User ID=postgres;Password=sch1aap;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
            conn.Open()

            comm = New NpgsqlCommand(myQuery, conn)
            Dim dr As Npgsql.NpgsqlDataReader
            dr = comm.ExecuteReader()

            While dr.Read()
                myMeteoVal = New clsMeteoValue
                myMeteoVal.ValueObserved = Val(dr(3))
                myMeteoVal.Datum = dr(4)
                myMeteoVal.Tijd = 0
                myMeteoVal.DateTimeVal = New DateTime(Left(myMeteoVal.Datum, 4), Left(Right(myMeteoVal.Datum, 4), 2), Right(myMeteoVal.Datum, 2))
                myMeteoVal.ValueCorrected = myMeteoVal.ValueObserved / 10
                If Not dr(3).ToString.Trim = "" Then myMS.EvaporationDaily.Add(myMeteoVal.Datum, myMeteoVal)
            End While
            dr.Dispose()

            'sluit de verbinding met de database
            conn.Close()
            If conn.State = System.Data.ConnectionState.Open Then conn.Close()
            conn.Dispose()
        Next

        Call writeEtmaalMakkinkToExcel(worksheets, "Basis.Makkink.Etmaal")

    End Sub

    Public Sub processNeerslagBasisHourly(ByVal worksheets As ExcelWorksheetCollection)
        Dim conn As NpgsqlConnection
        Dim comm As NpgsqlCommand
        Dim myMeteoVal As clsMeteoValue
        Dim myMS As clsMeteoStation
        Dim myKey As String
        Dim nDouble As Long

        'HIER NUGET AANROEPEN



        conn = New NpgsqlConnection
        'standard: User ID=root;Password=myPassword;Host=localhost;Port=5432;Database=myDataBase; Pooling=true;Min Pool Size=0;Max Pool Size=100;Connection Lifetime=0;
        conn.ConnectionString = "User ID=postgres;Password=sch1aap;Host=localhost;Port=5432;Database=meteobase; Pooling=true"
        conn.Open()

        For Each myMS In Stations.MeteoStations.Values
            comm = New NpgsqlCommand("SELECT * FROM data.precipitation_hourly WHERE datumveld >= " & FDate & " AND datumveld <= " & TDate & " AND station = " & myMS.Number.ToString.Trim & " ORDER BY datumveld, tijdveld", conn)
            Dim dr As Npgsql.NpgsqlDataReader
            dr = comm.ExecuteReader()

            While dr.Read()
                myMeteoVal = New clsMeteoValue
                myMeteoVal.Datum = dr(6)
                myMeteoVal.Tijd = dr(7)
                If myMeteoVal.Tijd = 24 Then
                    myMeteoVal.DateTimeVal = New DateTime(Left(myMeteoVal.Datum, 4), Left(Right(myMeteoVal.Datum, 4), 2), Right(myMeteoVal.Datum, 2))
                    myMeteoVal.DateTimeVal = myMeteoVal.DateTimeVal.AddDays(1)
                Else
                    myMeteoVal.DateTimeVal = New DateTime(Left(myMeteoVal.Datum, 4), Left(Right(myMeteoVal.Datum, 4), 2), Right(myMeteoVal.Datum, 2), myMeteoVal.Tijd, 0, 0)
                End If
                myMeteoVal.ValueObserved = Val(dr(4))

                If myMeteoVal.ValueObserved < 0 Then
                    myMeteoVal.ValueCorrected = 0
                Else
                    myMeteoVal.ValueCorrected = myMeteoVal.ValueObserved / 10
                End If

                myKey = myMeteoVal.Datum & "_" & myMeteoVal.Tijd
                If Not myMS.PrecipitationHourly.ContainsKey(myKey) Then
                    myMS.PrecipitationHourly.Add(myKey, myMeteoVal)
                Else
                    nDouble += 1
                End If

            End While
            dr.Dispose()

            If nDouble > 0 Then
                Me.Setup.Log.AddWarning(nDouble & " instances of multiple values found for precipitation at station " & myMS.Name)
            End If

        Next

        'sluit de verbinding met de database
        conn.Close()
        If conn.State = System.Data.ConnectionState.Open Then conn.Close()
        conn.Dispose()

        Call writeBasisNeerslagUurToExcel(worksheets, "Basis.Neerslag.Uur")

    End Sub

    Public Sub writeEtmaalMakkinkToExcel(ByRef Worksheets As ExcelWorksheetCollection, ByVal SheetName As String)
        Dim r As Long, c As Long
        Dim ws As ExcelWorksheet
        ws = Worksheets.Add(SheetName)

        c = -4
        For Each myMS In Stations.MeteoStations.Values

            r = -1
            c += 4

            r += 1
            ws.Cells(r, c).Value = "Data ontsloten via:"
            ws.Cells(r, c + 1).Value = "www.meteobase.nl"
            r += 1
            ws.Cells(r, c).Value = "Herkomst brongegevens:"
            ws.Cells(r, c + 1).Value = "KNMI"
            r += 1
            ws.Cells(r, c).Value = "Naam station:"
            ws.Cells(r, c + 1).Value = myMS.Name
            r += 1
            ws.Cells(r, c).Value = "Nummer station:"
            ws.Cells(r, c + 1).Value = myMS.Number
            r += 1
            ws.Cells(r, c).Value = "Klimaatscenario:"
            ws.Cells(r, c + 1).Value = "HUIDIG"
            r += 1
            ws.Cells(r, c).Value = "Datum:"
            ws.Cells(r, c + 1).Value = "Datumwaarde:"
            ws.Cells(r, c + 2).Value = "Meetwaarde [0.1 mm]:"
            ws.Cells(r, c + 3).Value = "Meetwaarde bewerkt [mm]:"

            For Each myMeteoVal In myMS.EvaporationDaily.Values
                r += 1
                ws.Cells(r, c).Value = myMeteoVal.Datum
                ws.Cells(r, c + 1).Value = myMeteoVal.DateTimeVal
                ws.Cells(r, c + 2).Value = myMeteoVal.ValueObserved
                ws.Cells(r, c + 3).Value = myMeteoVal.ValueCorrected
            Next
        Next

    End Sub

    Public Sub writeEtmaalNeerslagToExcel(ByRef Worksheets As ExcelWorksheetCollection)
        Dim r As Long, c As Long
        Dim ws As ExcelWorksheet
        ws = Worksheets.Add("Basis.Neerslag.Etmaal")

        c = -4
        For Each myMS In Stations.MeteoStations.Values

            r = -1
            c += 4

            r += 1
            ws.Cells(r, c).Value = "Data ontsloten via:"
            ws.Cells(r, c + 1).Value = "www.meteobase.nl"
            r += 1
            ws.Cells(r, c).Value = "Herkomst brongegevens:"
            ws.Cells(r, c + 1).Value = "KNMI"
            r += 1
            ws.Cells(r, c).Value = "Waarde -1 staat voor: "
            ws.Cells(r, c + 1).Value = "< 0.05 mm."
            ws.Cells(r, c + 2).Value = "in kolom 'meetwaarde bewerkt' vervangen door:"
            ws.Cells(r, c + 3).Value = "0 mm (advies Klimaatdesk KNMI)"
            r += 1
            ws.Cells(r, c).Value = "Naam station:"
            ws.Cells(r, c + 1).Value = myMS.Name
            r += 1
            ws.Cells(r, c).Value = "Nummer station:"
            ws.Cells(r, c + 1).Value = myMS.Number
            r += 1
            ws.Cells(r, c).Value = "Datum:"
            ws.Cells(r, c + 1).Value = "Datumwaarde:"
            ws.Cells(r, c + 2).Value = "Meetwaarde [0.1 mm]:"
            ws.Cells(r, c + 3).Value = "Meetwaarde bewerkt [mm]:"

            For Each myMeteoVal In myMS.PrecipitationDaily.Values
                r += 1
                ws.Cells(r, c).Value = myMeteoVal.Datum
                ws.Cells(r, c + 1).Value = myMeteoVal.DateTimeVal
                ws.Cells(r, c + 2).Value = myMeteoVal.ValueObserved
                ws.Cells(r, c + 3).Value = myMeteoVal.ValueCorrected
            Next
        Next

    End Sub

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
                r += 1
                ws.Cells(r, c).Value = "Herkomst brongegevens:"
                ws.Cells(r, c + 1).Value = "KNMI"
                r += 1
                ws.Cells(r, c).Value = "Waarde -1 staat voor: "
                ws.Cells(r, c + 1).Value = "< 0.05 mm."
                ws.Cells(r, c + 2).Value = "in kolom 'meetwaarde bewerkt' vervangen door:"
                ws.Cells(r, c + 3).Value = "0 mm (advies Klimaatdesk KNMI)"
                r += 1
                ws.Cells(r, c).Value = "Naam station:"
                ws.Cells(r, c + 1).Value = myMS.Name
                r += 1
                ws.Cells(r, c).Value = "Nummer station:"
                ws.Cells(r, c + 1).Value = myMS.Number
                r += 1
                ws.Cells(r, c).Value = "Datum:"
                ws.Cells(r, c + 1).Value = "Tijd:"
                ws.Cells(r, c + 2).Value = "Datum/tijd:"
                ws.Cells(r, c + 3).Value = "Meetwaarde [0.1 mm]:"
                ws.Cells(r, c + 4).Value = "Meetwaarde bewerkt [mm]:"
                For Each myMeteoVal In myMS.PrecipitationHourly.Values
                    r += 1
                    ws.Cells(r, c).Value = myMeteoVal.Datum
                    ws.Cells(r, c + 1).Value = myMeteoVal.Tijd
                    ws.Cells(r, c + 2).Value = myMeteoVal.DateTimeVal
                    ws.Cells(r, c + 3).Value = myMeteoVal.ValueObserved
                    ws.Cells(r, c + 4).Value = myMeteoVal.ValueCorrected
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

        GoodMail = New clsEmail(Me.Setup)
        GoodMail.Message.Subject = "Meteobase bestelling " & OrderNum & " " & GegevensSoort

        'set the email's body text
        Dim sText As New TextPart("plain")
        sText.SetText("UTF-8", body)
        GoodMail.Message.Body = sText



    End Sub


    Public Sub InitializeBadMail(ByVal GegevensSoort As String)
        'initialiseer de email
        Dim body As String
        body = "Geachte " & Naam & ", " & vbCrLf
        body &= vbCrLf
        body &= "Er is iets misgegaan met uw bestelling bij MeteoBase. Onze excuses voor het ongemak!" & vbCrLf
        body &= "Uit de onderstaande diagnose kunt u wellicht achterhalen wat er fout ging." & vbCrLf
        body &= "Een kopie van deze mail is gestuurd naar info@meteobase.nl. Mocht de fout geen invoerfout blijken, dan nemen wij contact met u op." & vbCrLf
        body &= vbCrLf
        body &= "Diagnostische gegevens:  " & vbCrLf
        body &= "Session ID " & SessionID.ToString & vbCrLf
        body &= "Bestelnummer " & OrderNum.ToString & vbCrLf
        body &= "E-mailadres " & MailAdres & vbCrLf
        body &= "Resultatenbestand " & ExcelFile & vbCrLf
        body &= "Tijdsspanne: van = " & FDate.ToString & " tot=" & TDate.ToString & vbCrLf
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


        BadMail = New clsEmail(Me.Setup)
        BadMail.Message.Subject = "Meteobase bestelling " & OrderNum & " " & GegevensSoort & ": foutmelding"

        'set the email's body text
        Dim sText As New TextPart("plain")
        sText.SetText("UTF-8", body)
        GoodMail.Message.Body = sText


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

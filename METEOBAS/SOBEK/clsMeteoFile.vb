Option Explicit On
Imports METEOBAS.General
Imports System.IO

Public Class clsMeteoFile
    'This class was newly created by Siebe Bosch on 12 july 2020
    'It has been constructed from the components of the two separate classes clsBuiFile and clsEVPFile
    'Now, from within one single class both evaporation and precipitation can be written

    ''' <summary>
    ''' Geen constructor nodig
    ''' </summary>
    ''' <remarks></remarks>

    Friend StartDate As Date
    Friend EndDate As Date
    Friend TimeStep As TimeSpan                 'the (equidistant) timestep

    Friend MeteoStations As clsMeteoStations    'contains a dictionary of meteo stations. Also keeps trak of all polygon indices per station
    Friend Values(,) As Single                  'all the data in an array instead of a (memory consuming) clsTimeTable

    Private setup As clsSetup

    'SIEBE: op 23-7 heb ik een eerste aanzet gemaakt om af te stappen van de clsTimeTable (te memory consuming) en over te stappen op een 2D-array
    'het object TimeTable zal dus op termijn worden uitgefaseerd, ten faveure van Values(,)
    'omdat een .bui-file alleen equidistante data tijdstippen bevat is een array dates() niet nodig

    Public Sub New(ByVal mySetup As clsSetup)
        Me.setup = mySetup
        'TimeTable = New clsTimeTable(Me.setup)
        MeteoStations = New clsMeteoStations(Me.setup)
    End Sub

    Public Sub SetStartDate(myDate As Date)
        StartDate = myDate
    End Sub

    Public Sub SetEndDate(myDate As Date)
        EndDate = myDate
    End Sub

    Public Sub SetTimestepSize(myTs As TimeSpan)
        TimeStep = myTs
    End Sub

    Public Function CountTimesteps() As Integer
        Return GetTimeSpan.TotalMinutes / TimeStep.TotalMinutes
    End Function

    Public Function GetTimeSpan() As TimeSpan
        Return EndDate.Subtract(StartDate)
    End Function

    Public Function GetnRecords() As Long
        Return setup.GeneralFunctions.RoundUD(GetTimeSpan.TotalSeconds / TimeStep.TotalSeconds, 0, False)
    End Function

    Public Function InitializeRecords(myStartDate As DateTime, myEndDate As DateTime, ts As TimeSpan) As Boolean
        Try
            'this function initializes the records for this bui file
            If MeteoStations.MeteoStations.Count = 0 Then Throw New Exception("Error in function InitializeRecords of class clsBuiFile: Meteo stations not yet added.")

            'set startdate, enddate, total timespan and timestep
            StartDate = myStartDate
            EndDate = myEndDate
            TimeStep = ts
            Dim nRecords = GetnRecords()

            If nRecords > 0 Then
                ReDim Values(0 To nRecords - 1, 0 To MeteoStations.MeteoStations.Count - 1)
            Else
                Throw New Exception("Error initializing bui file. Number of records could not be computed.")
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function


    Public Sub WriteEVP(ByVal path As String)

        'deze functie schrijft een .evp-file voor SOBEK
        'let op: SOBEK ondersteunt slechts één meteostation. Dit betekent dat we de data moeten aggregeren
        Dim i As Long, myVal As Single

        Using evpWriter As New StreamWriter(path, False)
            'doorloop alle areas en schrijf de 'bui' weg
            evpWriter.WriteLine("*Name of this file: " & Me.setup.Settings.ExportDirRoot & "\model.evp")
            evpWriter.WriteLine("*Date and time of construction: " & Now & ".")
            evpWriter.WriteLine("*Evaporation file")
            evpWriter.WriteLine("*Meteo data: evaporation intensity in mm/day")
            evpWriter.WriteLine("*First record: start date, data in mm/day")
            evpWriter.WriteLine("*Datum (year month day), verdamping (mm/dag) voor elk weerstation")
            evpWriter.WriteLine("*jaar maand dag verdamping[mm]")

            'first calculate the summed up weight of all stations together
            Dim TotalWeight As Double = 0
            For Each myStation In MeteoStations.MeteoStations.Values
                TotalWeight += myStation.Factor
            Next

            Dim CurDate As Date = StartDate
            For i = 0 To UBound(Values, 1)

                'we need to aggregate the values from all underlying polygons since SOBEK only supports ONE evaporation station
                myVal = 0
                For j = 0 To MeteoStations.MeteoStations.Count - 1
                    myVal += MeteoStations.MeteoStations.Values(j).Factor / TotalWeight * Values(i, j)
                Next

                'write the resulting value to our EVP File
                evpWriter.WriteLine(" " & Year(CurDate) & "  " & Month(CurDate) & "  " & Day(CurDate) & " " & myVal)
                CurDate = CurDate.AddSeconds(TimeStep.TotalSeconds)
            Next
        End Using

    End Sub

    Friend Function GetAddMeteoStation(ByVal Key As String, ByVal Station As String) As clsMeteoStation
        If MeteoStations.MeteoStations.ContainsKey(Key) Then
            Return MeteoStations.MeteoStations.Item(Key)
        Else
            Dim myStation As New clsMeteoStation(Me.setup)
            myStation.ID = Station
            myStation.Factor = 1
            Call MeteoStations.MeteoStations.Add(Key, myStation)
            Return myStation
        End If
    End Function

    Friend Function GetAddMeteoStation(ByVal Key As String, ByVal Station As String, ShapeIdx As Integer, ByVal WeighingFactor As Double) As clsMeteoStation
        'if our station is already in the list, just add the weight factor
        If MeteoStations.MeteoStations.ContainsKey(Key) Then
            MeteoStations.MeteoStations.Item(Key).Factor += WeighingFactor
            MeteoStations.MeteoStations.Item(Key).ShapeIndices.Add(ShapeIdx)
            Return MeteoStations.MeteoStations.Item(Key)
        Else
            Dim myStation As New clsMeteoStation(Me.setup)
            myStation.Factor = WeighingFactor
            myStation.ID = Station
            myStation.ShapeIndices.Add(ShapeIdx)
            Call MeteoStations.MeteoStations.Add(Key, myStation)
            Return myStation
        End If

    End Function

    Public Function BuildSTOWATYPE_BUI(ByVal StationName As String, ByVal Vol As Double, ByVal ARF As Double, ByVal StartDate As DateTime, ByVal Fractie As Double(), ByVal Uitloop As Integer) As Boolean

        Dim i As Integer
        Dim ms As clsMeteoStation
        TimeStep = New TimeSpan(1, 0, 0)

        ms = GetAddMeteoStation(StationName.Trim.ToUpper, StationName)


        Call InitializeRecords(StartDate, StartDate.AddHours(Fractie.Count + Uitloop), New TimeSpan(1, 0, 0))

        'werkelijke bui
        For i = 0 To Fractie.Count - 1
            Values(i, 0) = Vol * ARF * Fractie(i)
        Next

        'uitloop
        For i = Fractie.Count To Fractie.Count + Uitloop - 1
            Values(i, 0) = 0
        Next

    End Function


    Public Sub BuildSTOWATYPE_EVP(ByVal Evap() As Double, StartDate As DateTime, ByVal UitloopHours As Integer)

        'NOTE: EVAPORATION FILES ALL HAVE DAILY VALUES!
        'werkelijke bui
        Dim i As Long
        ReDim Values(0 To UBound(Evap) + setup.GeneralFunctions.RoundUD(UitloopHours / 24, 0, True) - 1, 0)
        For i = 0 To Evap.Count - 1
            Values(i, 0) = Evap(i)
        Next

        'uitloop
        For i = Evap.Count To Evap.Count + setup.GeneralFunctions.RoundUD(UitloopHours / 24, 0, True) - 1
            Values(i, 0) = 0
        Next

    End Sub

    Public Function BuildLongTermEVAP(ByVal seizoen As GeneralFunctions.enmSeason, ByVal Duration As Integer, ByVal Uitloop As Integer) As Boolean

        Dim i As Integer
        Dim ms As clsMeteoStation
        TimeStep = New TimeSpan(1, 0, 0)
        Dim Evap As Double

        Select Case seizoen
            Case Is = GeneralFunctions.enmSeason.hydrosummerhalfyear
                StartDate = New DateTime(2000, 7, 1, 0, 0, 0)
                Evap = 4
            Case Is = GeneralFunctions.enmSeason.hydrowinterhalfyear
                StartDate = New DateTime(2000, 1, 1, 0, 0, 0)
                Evap = 2
            Case Is = GeneralFunctions.enmSeason.meteosummerhalfyear
                StartDate = New DateTime(2000, 7, 1, 0, 0, 0)
                Evap = 4
            Case Is = GeneralFunctions.enmSeason.meteowinterhalfyear
                StartDate = New DateTime(2000, 1, 1, 0, 0, 0)
                Evap = 2
            Case Is = GeneralFunctions.enmSeason.meteospringquarter
                StartDate = New DateTime(2000, 4, 1, 0, 0, 0)
                Evap = 3
            Case Is = GeneralFunctions.enmSeason.meteosummerquarter
                StartDate = New DateTime(2000, 7, 1, 0, 0, 0)
                Evap = 4
            Case Is = GeneralFunctions.enmSeason.meteoautumnquarter
                StartDate = New DateTime(2000, 11, 1, 0, 0, 0)
                Evap = 3
            Case Is = GeneralFunctions.enmSeason.meteowinterquarter
                StartDate = New DateTime(2000, 1, 1, 0, 0, 0)
                Evap = 2
            Case Else
                StartDate = New DateTime(2000, 1, 1, 0, 0, 0)
                Evap = 2
        End Select

        'verdamping
        ms = GetAddMeteoStation("NEERSLAG", "Neerslag")
        For i = 0 To Duration - 1
            Values(i, 0) = Evap
        Next

        'uitloop
        For i = 1 To Uitloop
            Values(i, 0) = 0
        Next

    End Function

    Public Sub WriteBUI(ByVal path As String, Optional ByVal nDigits As Integer = 2)

        Try
            'deze functie schrijft een .bui-file aan de hand van shapes in een shapefile plus een reeds gevulde lijst met stations
            Dim i As Integer, myStr As String, nTim As Long
            Dim nStations As Integer = MeteoStations.MeteoStations.Count
            Dim myStation As clsMeteoStation

            'determine the number formatting for the .bui file (e.g. "0.0000")
            Dim Formatting As String = "#.0"
            If nDigits > 1 Then
                For i = 2 To nDigits
                    Formatting &= "#"
                Next
            End If

            'zoek de informatie over tijdstappen op
            If GetnRecords() = 0 Then
                Throw New Exception("Error: no records in bui-file. Cannot write.")
            Else
                'let op: einddatum moet gecorrigeerd worden door er één tijdstap bij op te tellen (omdat die tijdstap zelf ook nog meetelt)
                Using buiWriter As New StreamWriter(path, False)
                    'doorloop alle areas en schrijf de 'bui' weg
                    buiWriter.WriteLine("*Name of this file: " & path)
                    buiWriter.WriteLine("*Date and time of construction: " & Now & ".")
                    buiWriter.WriteLine("1")
                    buiWriter.WriteLine("*Aantal stations")
                    buiWriter.WriteLine(MeteoStations.MeteoStations.Count)
                    buiWriter.WriteLine("*Namen van stations")
                    For Each myStation In MeteoStations.MeteoStations.Values
                        buiWriter.WriteLine(Chr(39) & myStation.ID & Chr(39)) 'de stations
                    Next
                    buiWriter.WriteLine("*Aantal gebeurtenissen (omdat het 1 bui betreft is dit altijd 1)")
                    buiWriter.WriteLine("*en het aantal seconden per waarnemingstijdstap")
                    buiWriter.WriteLine(" 1  " & TimeStep.TotalSeconds)
                    buiWriter.WriteLine("*Elke commentaarregel wordt begonnen met een * (asterisk).")
                    buiWriter.WriteLine("*Eerste record bevat startdatum en -tijd, lengte van de gebeurtenis in dd hh mm ss")
                    buiWriter.WriteLine("*Het format is: yyyymmdd:hhmmss:ddhhmmss")
                    buiWriter.WriteLine("*Daarna voor elk station de neerslag in mm per tijdstap.")

                    Dim dagen As Integer = GetTimeSpan.Days
                    Dim uren As Integer = GetTimeSpan.Hours
                    Dim seconds As Integer = GetTimeSpan.Seconds

                    'schrijf de instellingen voor datum/tijd en tijdstap naar de buifile. Zet het begin op de daadwerkelijke start van de resultaten
                    'en vul de tijdstappen met resultaten van een tijdstap verder
                    buiWriter.WriteLine(" " & StartDate.Year & " " & StartDate.Month & " " & StartDate.Day & " " & StartDate.Hour & " " & StartDate.Minute & " 0 " & dagen & " " & uren & " " & seconds & " 0 ")

                    'write the meteorological data
                    nTim = GetnRecords()

                    For i = 0 To nTim - 1

                        myStr = ""

                        For j = 0 To MeteoStations.MeteoStations.Count - 1
                            myStr &= (" " & Format(Values(i, j), Formatting))
                        Next
                        buiWriter.WriteLine(myStr)
                    Next

                    buiWriter.Close()
                End Using

            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Me.setup.Log.AddError("Error in sub Write of class clsBuiFile.")
        End Try

    End Sub

    Public Sub WriteAsCSV(ByVal path As String, Optional ByVal nDigits As Integer = 2)

        Try
            'deze functie schrijft een .csv-file aan de hand van shapes in een shapefile plus een reeds gevulde lijst met stations
            Dim i As Integer, j As Long, myStr As String, nTim As Long
            Dim myStation As clsMeteoStation
            Dim curDate As New DateTime

            'determine the number formatting for the .bui file (e.g. "0.0000")
            Dim Formatting As String = "#.0"
            If nDigits > 1 Then
                For i = 2 To nDigits
                    Formatting &= "#"
                Next
            End If


            'zoek de informatie over tijdstappen op
            If GetnRecords() = 0 Then
                Throw New Exception("Error: no records in timetable. Cannot write .csv file.")
            Else
                'let op: einddatum moet gecorrigeerd worden door er één tijdstap bij op te tellen (omdat die tijdstap zelf ook nog meetelt)
                Using csvWriter As New StreamWriter(path, False)

                    'doorloop alle areas en schrijf de 'csv' weg
                    Dim tmpStr As String = "Datum/Tijd"

                    For j = 0 To MeteoStations.MeteoStations.Values.Count - 1
                        myStation = MeteoStations.MeteoStations.Values(j)
                        tmpStr &= ";" & myStation.ID
                    Next
                    csvWriter.WriteLine(tmpStr)

                    nTim = GetnRecords()

                    For i = 0 To nTim - 1
                        myStr = ""
                        curDate = StartDate.AddSeconds(TimeStep.TotalSeconds * i)

                        tmpStr = Format(curDate, "yyyy/MM/dd HH:mm:ss")
                        For j = 0 To MeteoStations.MeteoStations.Values.Count - 1
                            tmpStr &= ";" & Format(Values(i, j), Formatting)
                        Next
                        csvWriter.WriteLine(tmpStr)
                    Next
                End Using

            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Me.setup.Log.AddError("Error in sub Write of class clsBuiFile.")
        End Try

    End Sub

End Class

Option Explicit On
Imports METEOBAS.General

Public Class clsMeteoStation
    Friend ID As String
    Friend CatchmentIdx As Integer    'index number of the catchment
    Friend Factor As Double           'discharge factor for this station
    Friend RepresentsCSO As Boolean   'Whether this station represents a Combined Sewer Overflow location
    Friend ShapeIndices As New List(Of Integer) 'contains the index numbers of all shapes/areas that are represented by this station
    Private setup As clsSetup

    Public Name As String
    Public Number As String
    Public Lat As Double
    Public Lon As Double
    Public X As Double
    Public Y As Double
    Public StationType As GeneralFunctions.enmMeteoStationType

    Public PrecipitationHourly As New Dictionary(Of String, clsMeteoValue)
    Public PrecipitationDaily As New Dictionary(Of String, clsMeteoValue)
    Public EvaporationDaily As New Dictionary(Of String, clsMeteoValue)

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
    End Sub

    Public Sub Clear()
        PrecipitationHourly = New Dictionary(Of String, clsMeteoValue)
        PrecipitationDaily = New Dictionary(Of String, clsMeteoValue)
        EvaporationDaily = New Dictionary(Of String, clsMeteoValue)
    End Sub

    Public Sub calcStats(ByVal Duration As Integer, ByVal Area As Double, ByRef mu As Double, ByRef alpha As Double, ByRef kappa As Double)
        'deze functie berekent de statistische parameters van de volumekansverdeling: 
        'mu = locatieparameter' alpha = schaalparameter, kappa = vormparameter
        Dim y As Double
        Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c As Double

        a1 = 17.92  'was 17.92 in 2009
        a2 = 0.225  'was 0.225 in 2009
        b1 = -3.57  'was -3.57 in 2009
        b2 = 0.427  'was 0.43 in 2009
        c = 0.128   'was 0.128 in 2009
        mu = a1 * Duration ^ a2 + b1 * Area ^ c + (b2 * Area ^ c) * Math.Log(Duration)

        a1 = 0.337  'Was 0.344 in 2009
        a2 = -0.018 'Was -0.025 in 2009
        b1 = -0.014 'Was -0.016 in 2009
        b2 = 0.0    'Was 0.0003 in 2009
        y = a1 + a2 * Math.Log(Duration) + b1 * Math.Log(Area) + b2 * Duration * Math.Log(Area)
        alpha = y * mu

        a1 = -0.206 'Was -0.206 in 2009
        b1 = 0.018  'Was 0.022 in 2009
        b2 = 0      'Was -0.004 in 2009
        kappa = a1 + b1 * Math.Log(Area) + b2 * Math.Log(Duration) * Math.Log(Area)

    End Sub

    Public Function calcARI(ByVal Volume As Double, ByVal mu As Double, ByVal alpha As Double, ByVal kappa As Double) As Double
        'berekent de herhalingstijd van een neerslagvolume, gegeven de kansverdeling (mu, alpha en kappa)

        'volledig uitgeschreven vergelijkingen (zit ergens een fout in):
        'overschrijdingskans F = Math.Exp(-1 * (1 - (kappa / alpha) * (Volume - mu)) ^ (1 / kappa))
        'overschrijdingsfreq F_jaar = -Math.Log(1 - F)

        Dim F_jaar As Double 'overschrijdingsfrequentie op jaarbasis
        F_jaar = (1 - kappa / alpha * (Volume - mu)) ^ (1 / kappa)
        Return 1 / F_jaar

        'onderstaand is een test of de terugrekening weer hetzelfde volume genereert
        'Dim myVol = calcVol(Area, Duration, ARI, mu, alpha, kappa)

    End Function

    Public Function calcVol(ByVal Area As Double, ByVal Duration As Integer, ByVal ARI As Double) As Double
        'Deze functie rekent terug. Gegeven duur, Oppervlak en overschrijdingskans
        'rekent hij het volume over een oppervlak groter dan puntneerslag uit
        Dim F_jaar As Double = 1 / ARI
        Dim mu As Double, alpha As Double, kappa As Double

        'bereken de parameterwaarden voor de kansdichtheidsfunctie
        'en vervolgens het bijbehorende neerslagvolume
        Call calcStats(Duration, Area, mu, alpha, kappa)
        Return mu + alpha / kappa * (1 - F_jaar ^ kappa)

    End Function

    Public Sub calcAreaReductionHourly(ByVal Area As Double)
        'in deze routine doorlopen we de hele tijdreeks en berekenen we de
        'oppervlaktegereduceerde neerslag
        'onderscheid eerst de buien en ken hun herhalingstijd eraan toe
        'begin met de korte buien. Als een langere bui de korte overlapt,
        'wordt de korte overschreden
        'Als alle herhalingstijden zijn berekend, kan de gereduceerde neerslag worden berekend
        Dim i As Integer, j As Integer, k As Integer, Duration As Integer
        Dim mySum As Double, myNextSum As Double
        Dim mu As Double, alpha As Double, kappa As Double
        Dim ARI As Double 'resp overschrijdingskans en herhalingstijd
        Dim myVal As clsMeteoValue, SkipEvent As Boolean

        For i = 1 To 2          '<8 is debugsetting. Standaard tot 8 laten lopen!
            Select Case i
                Case Is = 1
                    Duration = 4
                Case Is = 2
                    Duration = 8
                Case Is = 3
                    Duration = 12
                Case Is = 4
                    Duration = 24
                Case Is = 5
                    Duration = 48
                Case Is = 6
                    Duration = 96
                Case Is = 7
                    Duration = 192
                Case Is = 8
                    Duration = 216
            End Select

            'Bereken voor de huidige duur de kansverdeling van de neerslagvolumes
            'Voor een puntbron van neerslag wordt het minimumoppervlak van 6 km2 aangehouden
            Call calcStats(Duration, 6, mu, alpha, kappa)

            'doorloop de gecorrigeerde neerslagwaarden en onderscheid buien hierbinnen
            For j = 0 To PrecipitationHourly.Values.Count - 1
                mySum = getSumOfCorrected(j, Duration) 'haal de neerslagsom op
                myNextSum = getSumOfCorrected(j + 1, Duration)

                If myNextSum < mySum Then 'nu weten we dat we een losse bui te pakken hebben
                    'bereken de overschrijdingskans van deze puntneerslagsom en haal on the fly ook de bijbehorende Herhalingstijd binnen
                    ARI = calcARI(mySum, mu, alpha, kappa)

                    'alleen als de herhalingstijd > 1 jaar is, schrijven we hem weg
                    If ARI >= 1 Then

                        'doorloop eerst de lijst met herhalingstijden om te checken of hij al is toegekend
                        SkipEvent = False 'initialiseer SkipEvent
                        For k = j To j + Duration - 1
                            myVal = PrecipitationHourly.Values(k)
                            If myVal.ARI > ARI Then
                                'helaas, een gebeurtenis met kortere duur had al een grotere herhalingstijd. We skippen deze bui voor de huidige duur
                                SkipEvent = True
                                Exit For
                            End If
                        Next

                        'als deze gebeurtenis nog niet is overruled door een zeldzamer herhalingstijd bij kortere duur:
                        'leg de herhalingstijd vast!
                        If Not SkipEvent Then
                            For k = j To j + Duration - 1
                                myVal = PrecipitationHourly.Values(k)
                                myVal.ARI = ARI           'leg voor deze bui de herhalingstijd vast
                                myVal.EventSum = mySum    'leg voor deze bui de neerslagsom vast
                                myVal.Duration = Duration 'leg voor deze bui de neerslagduur vast
                            Next
                        End If
                        'Bui is afgehandeld, dus zet j aan het einde van de bui
                        j = j + Duration - 1
                    End If
                End If
            Next
        Next

        'LET OP: onderstaand is de implementatie van de gebiedsreductiefactor.
        'Wordt nu niet meer aangeboden, want in aparte XLS
        'we hebben alle neerslagduren doorlopen, en weten nu voor elke bui in de reeks
        'welke herhalingstijd (ARI), som (mySum) en duur (Duration) hij heeft
        'Daarom gaan we nu de hele reeks nogmaals doorlopen en de gecorrigeerde neerslag
        'berekenen en wegschrijven
        'For i = 0 To PrecipitationHourly.Values.Count - 1
        '  myVal = PrecipitationHourly.Values(i)
        '  If myVal.ARI >= 1 Then

        '    'bereken de kansverdelingsparameters voor het nieuwe oppervlak 
        '    Call calcStats(myVal.Duration, Area, mu, alpha, kappa)

        '    'bereken de oppervlaktegecorrigeerde som
        '    myCorrSum = calcVol(Area, myVal.Duration, myVal.ARI)
        '    For j = i To i + myVal.Duration - 1
        '      adjustVal = PrecipitationHourly.Values(j)
        '      adjustVal.ValueAdjusted = adjustVal.ValueCorrected * myCorrSum / myVal.EventSum
        '    Next
        '    i = i + myVal.Duration - 1 'nu we deze records bewerkt hebben, kunnen we de bui overslaan in de hoofdloop
        '  End If
        'Next
    End Sub

    Public Function getSumOfCorrected(ByVal StartIdx As Long, ByVal Duration As Integer) As Double
        Dim i As Long, EndIdx As Long
        Dim mySum As Double = 0

        EndIdx = Math.Min(StartIdx + Duration - 1, PrecipitationHourly.Values.Count - 1)
        For i = StartIdx To EndIdx
            mySum += PrecipitationHourly.Values(i).ValueCorrected
        Next
        Return mySum

    End Function

End Class

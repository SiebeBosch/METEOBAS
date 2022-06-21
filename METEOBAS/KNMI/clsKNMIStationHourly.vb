Imports System.IO
Imports METEOBAS.General

Public Class clsKNMIStationHourly
    Public Name As String
    Public Number As Integer
    Public X As Double
    Public Y As Double
    Public Lat As Double
    Public Lon As Double
    Public nRecords As Long

    Public DATEINT() As Long
    Public HOURINT() As Integer
    Public NSLRAW() As Single
    Public NSLCOR() As Single
    Public DURATION() As Integer
    Public EVENTSUM() As Double
    Public EVENTSUMCHECK() As Double
    Public ARI() As Single
    Public mu() As Single
    Public alpha() As Single
    Public kappa() As Single

    Private Setup As clsSetup

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.Setup = mySetup
    End Sub

    Public Function calcVol(ByVal Area As Double, ByVal Duration As Integer, ByVal ARI As Double, ByVal mu As Double, ByVal alpha As Double, ByVal kappa As Double) As Double
        'Deze functie rekent terug. Gegeven duur, Oppervlak en overschrijdingskans
        'rekent hij het volume over een oppervlak groter dan puntneerslag uit
        Dim F_jaar As Double = 1 / ARI
        Return mu + alpha / kappa * (1 - F_jaar ^ kappa)

    End Function

    Public Sub calcEVENTS(ByVal minARI As Single, ByVal Area As Double)
        'deze routine berekent voor de gehele dataset de herhalingstijd mits hij voldoet aan het gegeven minimum
        Dim i As Integer, j As Long, k As Long, Dur As Integer, mySum As Double, myNextSum As Double, H As Single
        Dim SkipEvent As Boolean
        Dim myMu As Double, myAlpha As Double, myKappa As Double 'probability function parameters

        For i = 1 To 8          '<8 is debugsetting. Standaard tot 8 laten lopen!
            Select Case i
                Case Is = 1
                    Dur = 4
                Case Is = 2
                    Dur = 8
                Case Is = 3
                    Dur = 12
                Case Is = 4
                    Dur = 24
                Case Is = 5
                    Dur = 48
                Case Is = 6
                    Dur = 96
                Case Is = 7
                    Dur = 192
                Case Is = 8
                    Dur = 216
            End Select


            'Bereken voor de huidige duur de kansverdeling van de neerslagvolumes
            'Voor een puntbron van neerslag wordt het minimumoppervlak van 6 km2 aangehouden
            Call calcStats(Dur, Area, myMu, myAlpha, myKappa)

            'doorloop de gecorrigeerde neerslagwaarden en onderscheid buien hierbinnen
            For j = 0 To NSLRAW.Count - 1
                mySum = getSumOfCorrected(j, Dur) 'haal de som van de (gecorrigeerde) neerslag
                myNextSum = getSumOfCorrected(j + 1, Dur)

                If myNextSum < mySum Then 'nu weten we dat we een losse bui te pakken hebben
                    'bereken de overschrijdingskans van deze puntneerslagsom en haal on the fly ook de bijbehorende Herhalingstijd binnen
                    H = calcARI(mySum, myMu, myAlpha, myKappa)

                    'alleen als de herhalingstijd > minimum is, schrijven we hem weg
                    If H >= minARI Then

                        'doorloop eerst de lijst met herhalingstijden om te checken of hij al is toegekend
                        SkipEvent = False 'initialiseer SkipEvent
                        For k = j To j + Dur - 1
                            If ARI(k) > H Then
                                'helaas, een gebeurtenis met kortere duur had al een grotere herhalingstijd. We skippen deze bui voor de huidige duur
                                SkipEvent = True
                                Exit For
                            End If
                        Next

                        'als deze gebeurtenis nog niet is overruled door een zeldzamer herhalingstijd bij kortere duur:
                        'leg de herhalingstijd vast!
                        If Not SkipEvent Then
                            For k = j To j + Dur - 1
                                ARI(k) = H                'leg voor deze bui de herhalingstijd vast
                                EVENTSUM(k) = mySum       'leg voor deze bui de neerslagsom vast
                                DURATION(k) = Dur       'leg voor deze bui de neerslagduur vast
                                mu(k) = myMu              'leg voor deze bui de mu (locatieparameter kansdichtheidsfunctie) vast
                                alpha(k) = myAlpha        'leg voor deze bui alpha vast
                                kappa(k) = myKappa        'leg voor deze bui kappa vast
                                EVENTSUMCHECK(k) = calcVol(Area, Dur, H, myMu, myAlpha, myKappa)
                            Next
                        End If
                        'Bui is afgehandeld, dus zet j aan het einde van de bui
                        j = j + Dur - 1
                    End If
                End If
            Next
        Next

    End Sub

    Public Sub calcStats(ByVal Duration As Integer, ByVal Area As Double, ByRef mu As Double, ByRef alpha As Double, ByRef kappa As Double)
        'deze functie berekent de statistische parameters van de volumekansverdeling: 
        'mu = locatieparameter' alpha = schaalparameter, kappa = vormparameter
        Dim y As Double
        Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c As Double

        a1 = 17.92
        a2 = 0.225
        b1 = -3.57
        b2 = 0.43
        c = 0.128
        mu = a1 * Duration ^ a2 + b1 * Area ^ c + (b2 * Area ^ c) * Math.Log(Duration)

        a1 = 0.344
        a2 = -0.025
        b1 = -0.016
        b2 = 0.0003
        y = a1 + a2 * Math.Log(Duration) + b1 * Math.Log(Area) + b2 * Duration * Math.Log(Area)
        alpha = y * mu

        a1 = -0.206
        b1 = 0.022
        b2 = -0.004
        kappa = a1 + b1 * Math.Log(Area) + b2 * Math.Log(Duration) * Math.Log(Area)

    End Sub

    Public Function calcARI(ByVal Volume As Double, ByVal mu As Double, ByVal alpha As Double, ByVal kappa As Double) As Double
        'berekent de herhalingstijd van een neerslagvolume, gegeven de kansverdeling (mu, alpha en kappa)

        Dim F_jaar As Double 'overschrijdingsfrequentie op jaarbasis
        F_jaar = (1 - kappa / alpha * (Volume - mu)) ^ (1 / kappa)
        Return 1 / F_jaar

        'onderstaand is een test of de terugrekening weer hetzelfde volume genereert
        'Dim myVol = calcVol(Area, Duration, ARI, mu, alpha, kappa)

    End Function


    Public Function getSumOfCorrected(ByVal StartIdx As Long, ByVal Duration As Integer) As Double
        Dim i As Long, EndIdx As Long
        Dim mySum As Double = 0

        EndIdx = Math.Min(StartIdx + Duration - 1, NSLCOR.Count - 1)
        For i = StartIdx To EndIdx
            mySum += NSLCOR(i)
        Next
        Return mySum

    End Function

    Public Sub write(ByVal myPath As String)
        Dim i As Long, mystr As String
        Using myWriter As New StreamWriter(myPath)
            myWriter.WriteLine("Station,Date,Hour,Precipitation raw, Precipitation corrected,Sum Event, Duration event,ARI,mu,alpha,kappa,Volume Check")
            For i = 0 To NSLCOR.Count - 1
                mystr = Number & "," & DATEINT(i) & "," & HOURINT(i) & "," & NSLRAW(i) & "," & NSLCOR(i) & "," & EVENTSUM(i) & "," & DURATION(i) & "," & ARI(i) & "," & mu(i) & "," & alpha(i) & "," & kappa(i) & "," & EVENTSUMCHECK(i)
                myWriter.WriteLine(mystr)
            Next
            myWriter.Close()
        End Using

    End Sub

End Class

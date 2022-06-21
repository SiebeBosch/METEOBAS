Imports METEOBAS.General

Public Class clsRegenduurLijn
    Dim Dur(8) As Integer
    Dim Vol(11) As Integer
    Dim Herh(11, 8) As Double
    Private Setup As clsSetup

    Public Sub New(ByRef mySetup As clsSetup)
        Setup = mySetup
    End Sub

    Public Sub Create(Klimaat As String, Regio As String)
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
        Vol(1) = 10
        Vol(2) = 20
        Vol(3) = 30
        Vol(4) = 40
        Vol(5) = 50
        Vol(6) = 70
        Vol(7) = 90
        Vol(8) = 110
        Vol(9) = 130
        Vol(10) = 150
        Vol(11) = 170


        Select Case Klimaat.Trim.ToUpper
            Case Is = "HUIDIG"
                Select Case Regio.Trim.ToUpper
                    Case Is = "G"
                        '5mm
                        Herh(0, 0) = 0.5
                        Herh(0, 1) = 0.5
                        Herh(0, 2) = 0.5
                        Herh(0, 3) = 0.5
                        Herh(0, 4) = 0.5
                        Herh(0, 5) = 0.5
                        Herh(0, 6) = 0.5
                        Herh(0, 7) = 0.5
                        Herh(0, 8) = 0.5

                        '10mm
                        Herh(1, 0) = 0.5
                        Herh(1, 1) = 0.5
                        Herh(1, 2) = 0.5
                        Herh(1, 3) = 0.5
                        Herh(1, 4) = 0.5
                        Herh(1, 5) = 0.5
                        Herh(1, 6) = 0.5
                        Herh(1, 7) = 0.5
                        Herh(1, 8) = 0.5

                        '20mm
                        Herh(2, 0) = 1.2
                        Herh(2, 1) = 0.7
                        Herh(2, 2) = 0.5
                        Herh(2, 3) = 0.5
                        Herh(2, 4) = 0.5
                        Herh(2, 5) = 0.5
                        Herh(2, 6) = 0.5
                        Herh(2, 7) = 0.5
                        Herh(2, 8) = 0.5

                        '30mm
                        Herh(3, 0) = 4.9
                        Herh(3, 1) = 3.0
                        Herh(3, 2) = 1.6
                        Herh(3, 3) = 1.1
                        Herh(3, 4) = 0.5
                        Herh(3, 5) = 0.5
                        Herh(3, 6) = 0.5
                        Herh(3, 7) = 0.5
                        Herh(3, 8) = 0.5

                        '40mm
                        Herh(4, 0) = 17.5
                        Herh(4, 1) = 10.4
                        Herh(4, 2) = 5.5
                        Herh(4, 3) = 3.6
                        Herh(4, 4) = 1.5
                        Herh(4, 5) = 0.6
                        Herh(4, 6) = 0.5
                        Herh(4, 7) = 0.5
                        Herh(4, 8) = 0.5

                        '50mm
                        Herh(5, 0) = 54.9
                        Herh(5, 1) = 32.2
                        Herh(5, 2) = 16.7
                        Herh(5, 3) = 10.6
                        Herh(5, 4) = 4.4
                        Herh(5, 5) = 1.6
                        Herh(5, 6) = 0.5
                        Herh(5, 7) = 0.5
                        Herh(5, 8) = 0.5

                        '70mm
                        Herh(6, 0) = 400.2
                        Herh(6, 1) = 232.4
                        Herh(6, 2) = 117.5
                        Herh(6, 3) = 73.2
                        Herh(6, 4) = 28.3
                        Herh(6, 5) = 9.1
                        Herh(6, 6) = 2.4
                        Herh(6, 7) = 0.6
                        Herh(6, 8) = 0.5

                        '90mm
                        Herh(7, 0) = 1000
                        Herh(7, 1) = 1000
                        Herh(7, 2) = 628.3
                        Herh(7, 3) = 388.3
                        Herh(7, 4) = 146.7
                        Herh(7, 5) = 44.5
                        Herh(7, 6) = 10.5
                        Herh(7, 7) = 2.0
                        Herh(7, 8) = 1.5

                        '110mm
                        Herh(8, 0) = 1000
                        Herh(8, 1) = 1000
                        Herh(8, 2) = 1000
                        Herh(8, 3) = 1000
                        Herh(8, 4) = 637.0
                        Herh(8, 5) = 192.2
                        Herh(8, 6) = 43.3
                        Herh(8, 7) = 6.8
                        Herh(8, 8) = 4.9

                        '130mm
                        Herh(9, 0) = 1000
                        Herh(9, 1) = 1000
                        Herh(9, 2) = 1000
                        Herh(9, 3) = 1000
                        Herh(9, 4) = 1000
                        Herh(9, 5) = 742.7
                        Herh(9, 6) = 171.2
                        Herh(9, 7) = 24.9
                        Herh(9, 8) = 17.0

                        '150mm
                        Herh(10, 0) = 1000
                        Herh(10, 1) = 1000
                        Herh(10, 2) = 1000
                        Herh(10, 3) = 1000
                        Herh(10, 4) = 1000
                        Herh(10, 5) = 1000
                        Herh(10, 6) = 650.3
                        Herh(10, 7) = 98.3
                        Herh(10, 8) = 65.8

                        '150mm
                        Herh(11, 0) = 1000
                        Herh(11, 1) = 1000
                        Herh(11, 2) = 1000
                        Herh(11, 3) = 1000
                        Herh(11, 4) = 1000
                        Herh(11, 5) = 1000
                        Herh(11, 6) = 1000
                        Herh(11, 7) = 426.0
                        Herh(11, 8) = 290.2

                End Select
        End Select

    End Sub

    Public Function GetReturnPeriod(DurationHours As Double, Volume As Double) As Double

        Try
            'here we will interpolate between the four surrounding points in the 2D array
            'first find the four corners
            Dim sDurIdx As Integer = 0, eDurIdx As Integer = 0
            Dim sVolIdx As Integer, eVolIdx As Integer

            If DurationHours <= Dur(0) Then
                sDurIdx = 0
                eDurIdx = 0
            ElseIf DurationHours >= Dur(Dur.Count - 1) Then
                sDurIdx = Dur.Count - 1
                eDurIdx = Dur.Count - 1
            Else
                For i = 0 To Dur.Count - 1
                    If Dur(i) <= DurationHours Then sDurIdx = i
                    If Dur(i) >= DurationHours Then
                        eDurIdx = i
                        Exit For
                    End If
                Next
            End If

            'get the surrounding volumes
            If Volume <= Vol(0) Then
                sVolIdx = 0
                eVolIdx = 0
            ElseIf Volume >= Vol(Vol.Count - 1) Then
                sVolIdx = Vol.Count - 1
                eVolIdx = Vol.Count - 1
            Else
                For i = 0 To Vol.Count - 1
                    If Vol(i) <= Volume Then sVolIdx = i
                    If Vol(i) >= Volume Then
                        eVolIdx = i
                        Exit For
                    End If
                Next
            End If

            'get the return periods for all four corner points
            Dim HerhsVolsDur As Double = Herh(sVolIdx, sDurIdx)
            Dim HerheVolsDur As Double = Herh(eVolIdx, sDurIdx)
            Dim HerhsVoleDur As Double = Herh(sVolIdx, eDurIdx)
            Dim HerheVoleDur As Double = Herh(eVolIdx, eDurIdx)

            'first interpolate between the volumes of the startDuration
            Dim HerhsDur = Me.Setup.GeneralFunctions.Interpolate(Vol(sVolIdx), Herh(sVolIdx, sDurIdx), Vol(eVolIdx), Herh(eVolIdx, sDurIdx), Volume)
            Dim HerheDur = Me.Setup.GeneralFunctions.Interpolate(Vol(sVolIdx), Herh(sVolIdx, eDurIdx), Vol(eVolIdx), Herh(eVolIdx, eDurIdx), Volume)

            'then interpolate along the durations
            Return Me.Setup.GeneralFunctions.Interpolate(Dur(sDurIdx), HerhsDur, Dur(eDurIdx), HerheDur, DurationHours)
        Catch ex As Exception
            Me.Setup.Log.AddError("Fout bij het interpoleren uit regenduurlijn. Kon herhalingstijd niet opvragen voor volume=" & Volume & " en duur=" & DurationHours)
            Return 0
        End Try


    End Function

End Class

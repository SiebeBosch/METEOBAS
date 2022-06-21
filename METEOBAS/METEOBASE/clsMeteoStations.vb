Option Explicit On
Imports METEOBAS.General
Imports System.IO

''' <summary>
''' Geen constructor nodig
''' </summary>
''' <remarks></remarks>
Public Class clsMeteoStations
    Friend MeteoStations As New Dictionary(Of String, clsMeteoStation)

    Private setup As clsSetup

    Public Sub New(ByVal mySetup As clsSetup)
        Me.setup = mySetup
    End Sub

    Public Sub Initialize()
        MeteoStations = New Dictionary(Of String, clsMeteoStation)
    End Sub

    Public Function GetByKey(ID As String) As clsMeteoStation
        If MeteoStations.ContainsKey(ID.Trim.ToUpper) Then
            Return MeteoStations.Item(ID.Trim.ToUpper)
        Else
            Return Nothing
        End If
    End Function



    Public Sub Add(ByVal Name As String, ByVal soort As String, ByVal ARF As Double)
        Dim myStation As New clsMeteoStation(Me.setup)
        myStation.ID = Name
        myStation.Name = Name
        myStation.Factor = ARF
        myStation.StationType = Me.setup.GeneralFunctions.getMeteoStationTypeFromString(soort)
        MeteoStations.Add(myStation.ID, myStation)
    End Sub

    Friend Function getStationIdx(ByVal ID As String) As Integer
        'Returns the index number of the requested meteo station
        Dim i As Integer, myStation As clsMeteoStation

        For i = 0 To MeteoStations.Count - 1
            myStation = MeteoStations.Values(i)
            If myStation.ID.Trim.ToUpper = ID.Trim.ToUpper Then Return i
        Next

        Return -1

    End Function

    Friend Function GetStation(ByRef ms As clsMeteoStation, ByVal ID As String) As Boolean
        If Not MeteoStations.ContainsKey(ID.Trim.ToUpper) Then
            Return False
        Else
            ms = MeteoStations.Item(ID.Trim.ToUpper)
            Return True
        End If
    End Function


    Friend Function GetAdd(ByRef ms As clsMeteoStation, ByVal ID As String) As clsMeteoStation
        If Not MeteoStations.ContainsKey(ID.Trim.ToUpper) Then
            MeteoStations.Add(ms.ID.Trim.ToUpper, ms)
        Else
            Return MeteoStations.Item(ID.Trim.ToUpper)
        End If
        Return MeteoStations(ID)

    End Function

    Friend Function GetByNumber(ByVal StationNumber As String) As clsMeteoStation
        For Each myMS As clsMeteoStation In MeteoStations.Values
            If myMS.Number = StationNumber.Trim Then
                Return myMS
            End If
        Next
        Return Nothing
    End Function


End Class

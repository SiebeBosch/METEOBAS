Imports System.IO
Imports METEOBAS.General

Public Class clsKNMIData
    Public MBTextFiles As New Collection 'of clsMeteoBaseTextFile
    Public StationsHourly As New Collection 'of clsKNMIStationHourly
    Private Setup As clsSetup

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.Setup = mySetup
    End Sub

    Public Function readMBTextFiles(ByVal myDir As String) As Boolean
        'reads datafiles for MeteoBase
        Dim myFiles() As String, myFile As String, i As Integer
        myFiles = Directory.GetFiles(myDir)
        For i = 0 To myFiles.Count - 1
            myFile = myFiles(i)
            Call readMBTextFile(myFile)
        Next
        Return True
    End Function

    Public Function calcHourlyEventStats(ByVal myArea As Double) As Boolean
        Dim i As Integer = 0
        For Each myStation As clsKNMIStationHourly In StationsHourly
            i += 1
            myStation.calcEVENTS(1, myArea)
        Next
        Return True
    End Function

    Public Function writeMBTextFiles(ByVal myDir As String) As Boolean
        For Each myStation As clsKNMIStationHourly In StationsHourly
            myStation.write(Me.Setup.importDir & "\" & myStation.Number & ".csv")
        Next
        Return True
    End Function

    Public Sub readMBTextFile(ByVal myFile As String)
        'reads a datafile for MeteoBase
        Dim entry As String
        Dim myStationNum As Integer, myDate As Integer, myHour As Integer
        Dim myStation As clsKNMIStationHourly

        Using myReader As New StreamReader(myFile)
            While Not myReader.EndOfStream
                entry = myReader.ReadLine.Trim
                If Not Left(entry, 1) = "#" Then
                    myStationNum = Setup.GeneralFunctions.ParseString(entry, ",")
                    If Not StationsHourly.Contains(myStationNum.ToString.Trim) Then
                        myStation = New clsKNMIStationHourly(Me.Setup)
                        myStation.Number = myStationNum
                        StationsHourly.Add(myStation, myStationNum.ToString.Trim)
                        ReDim myStation.DATEINT(120 * 365 * 24)
                        ReDim myStation.HOURINT(120 * 365 * 24)
                        ReDim myStation.NSLRAW(120 * 365 * 24)
                        ReDim myStation.NSLCOR(120 * 365 * 24)
                    Else
                        myStation = StationsHourly.Item(myStationNum.ToString.Trim)
                    End If

                    myDate = Setup.GeneralFunctions.ParseString(entry, ",")
                    myHour = Setup.GeneralFunctions.ParseString(entry, ",")
                    If Not entry.Trim = "" Then
                        myStation.nRecords += 1
                        myStation.DATEINT(myStation.nRecords - 1) = myDate
                        myStation.HOURINT(myStation.nRecords - 1) = myHour
                        myStation.NSLRAW(myStation.nRecords - 1) = Setup.GeneralFunctions.ParseString(entry, ",")
                        If myStation.NSLRAW(myStation.nRecords - 1) < 0 Then
                            myStation.NSLCOR(myStation.nRecords - 1) = 0
                        Else
                            myStation.NSLCOR(myStation.nRecords - 1) = myStation.NSLRAW(myStation.nRecords - 1) / 10
                        End If

                    End If
                End If
            End While
        End Using

        'herdimensioneer de arrays
        For Each myStation In StationsHourly
            ReDim Preserve myStation.DATEINT(myStation.nRecords - 1)
            ReDim Preserve myStation.HOURINT(myStation.nRecords - 1)
            ReDim Preserve myStation.NSLRAW(myStation.nRecords - 1)
            ReDim Preserve myStation.NSLCOR(myStation.nRecords - 1)

            'en dimensioneer de overige arrays!
            ReDim myStation.mu(myStation.nRecords - 1)
            ReDim myStation.alpha(myStation.nRecords - 1)
            ReDim myStation.kappa(myStation.nRecords - 1)
            ReDim myStation.DURATION(myStation.nRecords - 1)
            ReDim myStation.EVENTSUM(myStation.nRecords - 1)
            ReDim myStation.EVENTSUMCHECK(myStation.nRecords - 1)
            ReDim myStation.ARI(myStation.nRecords - 1)

        Next

    End Sub

End Class

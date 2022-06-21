Imports System.Net
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.IO
Imports METEOBAS.General


Public Class clsWIWB_API

    Private Setup As clsSetup

    Public Sub New(ByRef mySetup As clsSetup)
        Setup = mySetup
    End Sub

    Public Function GetDatasources(request As JObject) As JObject
        Try

            Dim baseUrl As String = "https://wiwb.hydronet.com"
            'Dim url__1 As String = Url.Combine(baseUrl, "entity", "dataSources", "get")
            Dim url__1 As String = baseUrl & "/api/entity/datasources/get"

            Dim webRequest As HttpWebRequest = TryCast(HttpWebRequest.Create(url__1), HttpWebRequest)
            webRequest.Method = "POST"
            webRequest.ContentType = "application/json"
            webRequest.Accept = "application/json"

            Dim authInfo As String = "siebe.bosch:" + "iY0Hofot3zaZWxyCOxPX"
            authInfo = Convert.ToBase64String(Encoding.[Default].GetBytes(authInfo))
            webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo

            'If request Is Nothing Then
            '    request = New TRequest()
            'End If

            Dim body As String = JsonConvert.SerializeObject(request)
            Using sw As New StreamWriter(webRequest.GetRequestStream())
                sw.Write(body)
            End Using
            Dim WebResponse As HttpWebResponse = DirectCast(webRequest.GetResponse(), HttpWebResponse)

            Using responseStream As New StreamReader(WebResponse.GetResponseStream())
                Dim response As String = responseStream.ReadToEnd()
                Dim responseObject As JObject = JsonConvert.DeserializeObject(Of JObject)(response)
                Return responseObject
            End Using

            Throw New Exception("Could not get datasources from API.")
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Throw New Exception("Error in function GetDatasources of class clsWIWB_API")
            Return Nothing
        End Try
    End Function

    Public Function GetTimeSeries(DataSourceCode As String, Parameter As String, ByRef Stations As clsMeteoStations, StartDate As Integer, EndDate As Integer, DataFormatCode As String, IntervalMinutes As Integer) As List(Of String) ' JObject

        '------------------------------------------------------------------------------------------------------------------------
        '  THIS FUNCTION REQUESTS TIMESERIES FOR GIVEN METEO STATIONS FROM THE WIWB API
        '------------------------------------------------------------------------------------------------------------------------

        Try
            Dim Results As New List(Of String)

            'start by composing a JSON object for the request
            Dim request As New JObject
            Dim Readers As New JArray
            Dim Exporter As New JObject
            Dim Settings As New JObject
            Dim Reader As New JObject
            Dim LocationCodes As New JArray
            Dim VariableCodes As New JArray
            Dim Interval As New JObject
            Dim ExporterSetting As New JObject

            'the JSON building blocks
            request("Readers") = Readers
            request("Exporter") = Exporter
            Reader = New JObject
            Reader("DataSourceCode") = DataSourceCode
            Readers.Add(Reader)
            Reader("Settings") = Settings

            'populate the JSON building blocks
            'Settings("StructureType") = "TimeSeries"           'dit is een alternatief voor de methode op de volgende regel
            Settings.Add("StructureType", "TimeSeries")
            For Each myStation As clsMeteoStation In Stations.MeteoStations.Values
                LocationCodes.Add(myStation.Number)
            Next
            Settings.Add("LocationCodes", LocationCodes)
            Settings.Add("ReadQuality", False)
            Settings.Add("ReadAvailability", True)
            Settings.Add("StartDate", StartDate.ToString & "000000")
            Settings.Add("EndDate", EndDate.ToString & "000000")
            VariableCodes.Add(Parameter)
            Settings.Add("VariableCodes", VariableCodes)

            Select Case IntervalMinutes
                Case Is = 60
                    Interval("Type") = "Hours"
                    Interval("Value") = 1
                Case Is = 60 * 24
                    Interval("Type") = "Days"
                    Interval("Value") = 1
                Case Else
                    Interval("Type") = "Minutes"
                    Interval("Value") = IntervalMinutes
            End Select
            Settings.Add("Interval", Interval)

            'Interval("Type") = "Hours"
            'Interval("Value") = 1
            'Interval("Type") = "Minutes10"
            'Interval("Value") = 1
            'Interval("Type") = "Days"
            'Interval("Value") = 1
            Settings.Add("ReadAccumulated", False)
            Settings.Add("CalculationType", "Default")
            ExporterSetting.Add("DigitsToRound", 2)
            Exporter.Add("DataFormatCode", DataFormatCode)
            Exporter.Add("Settings", ExporterSetting)

            Dim baseUrl As String = "https://wiwb.hydronet.com"
            'Dim url__1 As String = Url.Combine(baseUrl, "entity", "dataSources", "get")
            Dim url__1 As String = baseUrl & "/api/timeseries/get"

            Dim webRequest As HttpWebRequest = TryCast(HttpWebRequest.Create(url__1), HttpWebRequest)
            webRequest.Method = "POST"
            webRequest.ContentType = "application/json"
            webRequest.Accept = "application/json"

            Dim authInfo As String = "siebe.bosch:" + "iY0Hofot3zaZWxyCOxPX"
            authInfo = Convert.ToBase64String(Encoding.[Default].GetBytes(authInfo))
            webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo

            Dim body As String = JsonConvert.SerializeObject(request)
            Me.Setup.Log.AddMessage("JSON POST: " & body)
            Using sw As New StreamWriter(webRequest.GetRequestStream())
                sw.Write(body)
            End Using
            Dim WebResponse As HttpWebResponse = DirectCast(webRequest.GetResponse(), HttpWebResponse)

            Using responseStream As New StreamReader(WebResponse.GetResponseStream())
                While Not responseStream.EndOfStream
                    Results.Add(responseStream.ReadLine)
                End While
                'Dim response As String = responseStream.ReadToEnd()
                'Dim responseObject As JObject = JsonConvert.DeserializeObject(Of JObject)(response)
                'Return responseObject
                'Return response
                Return Results
            End Using

            Throw New Exception("Error retrieving timeseries through the WIWB API.")
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Throw New Exception("Could not get timeseries from API.")
            Return Nothing
        End Try
    End Function

    Public Function GetRasters(DataSourceCode As String, Parameter As String, XMin As Double, YMin As Double, XMax As Double, YMax As Double, StartDate As Integer, EndDate As Integer, DataFormatCode As String, ByRef Path As String, Optional ByVal Accumulated As Boolean = False) As Boolean ' JObject
        '------------------------------------------------------------------------------------------------------------------------
        '  THIS FUNCTION REQUESTS RASTERS FOR GIVEN SPATIAL EXTENT & TIMESPAN FROM THE WIWB API
        '  IT THEN WRITES A ZIPFILE CONTAINING ALL RESULTING GRIDS
        '------------------------------------------------------------------------------------------------------------------------

        Try
            Dim Results As New List(Of String)

            'start by composing a JSON object for the request
            Dim request As New JObject
            Dim Readers As New JArray
            Dim Exporter As New JObject
            Dim DownloadOptions As New JObject
            Dim Settings As New JObject
            Dim Reader As New JObject
            Dim LocationCodes As New JArray
            Dim VariableCodes As New JArray
            Dim Interval As New JObject
            Dim SpatialReference As New JObject
            Dim Extent As New JObject
            Dim ExporterSetting As New JObject

            'the JSON building blocks
            request("Readers") = Readers
            request("Exporter") = Exporter
            request("DataFlowTypeCode") = "Download"
            Reader = New JObject
            'Reader("DataFlowTypeCode") = "Download"
            Reader("DataSourceCode") = DataSourceCode
            Readers.Add(Reader)
            Reader("Settings") = Settings

            'populate the JSON building blocks
            Settings.Add("StructureType", "Grid")
            Settings.Add("ReadQuality", False)
            Settings.Add("ReadAvailability", True)
            Settings.Add("StartDate", StartDate.ToString & "000000")
            Settings.Add("EndDate", EndDate.ToString & "000000")
            VariableCodes.Add(Parameter)
            Settings.Add("VariableCodes", VariableCodes)
            SpatialReference("Epsg") = "28992"
            Extent("XLL") = XMin
            Extent("YLL") = YMin
            Extent("XUR") = XMax
            Extent("YUR") = YMax
            Extent.Add("SpatialReference", SpatialReference)
            Settings.Add("Extent", Extent)
            If Accumulated Then
                Interval("Type") = "Total"
            Else
                Interval("Type") = "Hours"
                Interval("Value") = 1
            End If
            Settings.Add("Interval", Interval)
            Settings.Add("ReadAccumulated", False)
            'Settings.Add("CalculationType", "Default")
            ExporterSetting.Add("DigitsToRound", 2)
            Exporter.Add("DataFormatCode", DataFormatCode)
            Exporter.Add("Settings", ExporterSetting)
            DownloadOptions.Add("DataFlowTypeCode", "Download")

            Dim baseUrl As String = "https://wiwb.hydronet.com"
            'Dim url__1 As String = Url.Combine(baseUrl, "entity", "dataSources", "get")
            Dim url__1 As String = baseUrl & "/api/grids/get"

            Dim webRequest As HttpWebRequest = TryCast(HttpWebRequest.Create(url__1), HttpWebRequest)
            webRequest.Method = "POST"
            webRequest.ContentType = "application/json"
            webRequest.Accept = "application/json"

            Dim authInfo As String = "siebe.bosch:" + "iY0Hofot3zaZWxyCOxPX"
            authInfo = Convert.ToBase64String(Encoding.[Default].GetBytes(authInfo))
            webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo
            webRequest.Timeout = 20 * 60 * 1000 'set the timeout at 20 minutes

            Dim body As String = JsonConvert.SerializeObject(request)
            Me.Setup.Log.AddMessage("JSON query: " & body)
            Using sw As New StreamWriter(webRequest.GetRequestStream())
                sw.Write(body)
            End Using

            Dim WebResponse As HttpWebResponse = DirectCast(webRequest.GetResponse(), HttpWebResponse)


            'write the response to the logfile. This should contain the 
            Me.Setup.Log.AddMessage("Webresponse received. Length= " & WebResponse.ContentLength)
            Me.Setup.Log.AddMessage("Webresponse received. Type= " & WebResponse.ContentType)
            Me.Setup.Log.AddMessage("Webresponse received. Encoding= " & WebResponse.ContentEncoding)

            Dim s As Stream = WebResponse.GetResponseStream()

            'Write to disk
            Dim fs As New FileStream(Path, FileMode.Create)

            Dim read As Byte() = New Byte(255) {}
            Dim count As Integer = s.Read(read, 0, read.Length)
            While count > 0
                fs.Write(read, 0, count)
                count = s.Read(read, 0, read.Length)
            End While

            'Close everything
            fs.Close()
            s.Close()
            WebResponse.Close()

            'For i = 1 To 10
            '    System.Threading.Thread.Sleep(10000)

            '    Using fs As New FileStream(Path, FileMode.Create)
            '        WebResponse.GetResponseStream.CopyTo(fs)
            '    End Using
            '    Return True
            'Next

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Throw New Exception("Could not get grids from API.")
            Return Nothing
        End Try
    End Function


    Public Function DownloadRasters(DataSourceCode As String, Parameter As String, XMin As Double, YMin As Double, XMax As Double, YMax As Double, StartDate As Integer, EndDate As Integer, DataFormatCode As String, ByRef Path As String, Optional ByVal Accumulated As Boolean = False, Optional ByVal Aggregate24H As Boolean = False) As Boolean ' JObject
        '------------------------------------------------------------------------------------------------------------------------
        '  THIS FUNCTION REQUESTS RASTERS FOR GIVEN SPATIAL EXTENT & TIMESPAN FROM THE WIWB API
        '  IT USES THE DOWNLOAD-OPTION, WHICH IS MEANT FOR LARGE ORDERS
        '------------------------------------------------------------------------------------------------------------------------

        Try
            Dim Results As New List(Of String)
            Dim DataFlowId As Integer, DataFlowIdStr As String = ""
            Dim DataFlowIds As New List(Of Integer)
            Dim body As String
            Dim WebResponse As HttpWebResponse
            Dim webRequest As HttpWebRequest

            'start by composing a JSON object for the request
            Dim request As New JObject
            Dim Readers As New JArray
            Dim Exporter As New JObject
            Dim DownloadOptions As New JObject
            Dim Settings As New JObject
            Dim Reader As New JObject
            Dim LocationCodes As New JArray
            Dim VariableCodes As New JArray
            Dim Interval As New JObject
            Dim SpatialReference As New JObject
            Dim Extent As New JObject
            Dim ExporterSetting As New JObject

            'the JSON building blocks
            request("Readers") = Readers
            request("Exporter") = Exporter
            request("DataFlowTypeCode") = "Download"
            request("DataSourceCode") = DataSourceCode

            Reader = New JObject
            'Reader("DataFlowTypeCode") = "Download"
            Reader("DataSourceCode") = DataSourceCode
            Readers.Add(Reader)
            Reader("Settings") = Settings

            'populate the JSON building blocks
            Settings.Add("StructureType", "Grid")
            Settings.Add("ReadQuality", False)
            Settings.Add("ReadAvailability", True)
            Settings.Add("StartDate", StartDate.ToString & "000000")
            Settings.Add("EndDate", EndDate.ToString & "000000")

            VariableCodes.Add(Parameter)
            Settings.Add("VariableCodes", VariableCodes)
            SpatialReference("Epsg") = "28992"
            Extent("XLL") = XMin
            Extent("YLL") = YMin
            Extent("XUR") = XMax
            Extent("YUR") = YMax
            Extent.Add("SpatialReference", SpatialReference)
            Settings.Add("Extent", Extent)
            If Accumulated Then
                Interval("Type") = "Total"
            Else
                If Aggregate24H Then
                    Interval("Type") = "Hours"
                    Interval("Value") = 24
                Else
                    Interval("Type") = "Hours"
                    Interval("Value") = 1
                End If
            End If
            Settings.Add("Interval", Interval)
            Settings.Add("ReadAccumulated", False)
            'Settings.Add("CalculationType", "Default")
            ExporterSetting.Add("DigitsToRound", 2)
            Exporter.Add("DataFormatCode", DataFormatCode)
            Exporter.Add("Settings", ExporterSetting)
            'DownloadOptions.Add("DataFlowTypeCode", "Download")
            'DownloadOptions.Add("DataSourceCode", DataSourceCode)

            body = JsonConvert.SerializeObject(request)
            Me.Setup.Log.AddMessage("JSON query: " & body)



            Dim baseUrl As String = "https://wiwb.hydronet.com"
            'Dim url__1 As String = Url.Combine(baseUrl, "entity", "dataSources", "get")
            'Dim url__1 As String = baseUrl & "/api/grids/get"
            'Dim url__1 = baseUrl & "/api/grids/createdownload"

            Dim createDownloadURL As String = baseUrl & "/api/grids/createdownload"
            Dim downloadURL As String = baseUrl & "/api/entity/dataflows/get"

            'set the authentication header
            Dim authInfo As String = "siebe.bosch:" + "iY0Hofot3zaZWxyCOxPX"
            authInfo = Convert.ToBase64String(Encoding.[Default].GetBytes(authInfo))

            'create a webrequest
            webRequest = TryCast(HttpWebRequest.Create(createDownloadURL), HttpWebRequest)
            webRequest.Method = "POST"
            webRequest.ContentType = "application/json"
            webRequest.Accept = "application/json"
            webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo
            webRequest.Timeout = 10 * 60 * 1000 'set the timeout at 10 minutes
            Using sw As New StreamWriter(webRequest.GetRequestStream())
                sw.Write(body)
            End Using

            WebResponse = DirectCast(webRequest.GetResponse(), HttpWebResponse)

            'write the response to the logfile. This should contain the 
            Me.Setup.Log.AddMessage("Webresponse received. Length= " & WebResponse.ContentLength)
            Me.Setup.Log.AddMessage("Webresponse received. Type= " & WebResponse.ContentType)
            Me.Setup.Log.AddMessage("Webresponse received. Encoding= " & WebResponse.ContentEncoding)
            Me.Setup.Log.AddMessage("Webresponse status Description: " & WebResponse.StatusDescription)

            If WebResponse.StatusCode = 200 Then
                Me.Setup.Log.AddMessage("StatusCode 200 (=OK) received from webresponse.")
                Dim dataStream As Stream = WebResponse.GetResponseStream
                Dim dataReader As New StreamReader(dataStream)
                Dim ResponseStr As String = dataReader.ReadToEnd
                Me.Setup.Log.AddMessage("Response string: " & ResponseStr)
                Dim ResponseJson As JObject = JObject.Parse(ResponseStr)
                DataFlowId = ResponseJson.Item("DataFlowId")
                DataFlowIdStr = Chr(34) & DataFlowId & Chr(34)
                DataFlowIds.Add(DataFlowId)
            Else
                Me.Setup.Log.AddMessage("StatusCode received from webresponse = " & WebResponse.StatusCode)
            End If

            request = New JObject
            request("DataFlowIds") = "[" & DataFlowId.ToString & "]"
            body = JsonConvert.SerializeObject(request)
            Me.Setup.Log.AddMessage("JSON query: " & body)


            'wait until the State of our response is "Finished"
            Dim Sleeping As Boolean = True
            Dim Slept As Integer = 0
            While Sleeping
                webRequest = TryCast(HttpWebRequest.Create(downloadURL), HttpWebRequest)
                webRequest.Method = "POST"
                webRequest.ContentType = "application/json"
                webRequest.Accept = "application/json"
                webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo
                webRequest.Timeout = 10 * 60 * 1000 'set the timeout at 10 minutes
                Using sw As New StreamWriter(webRequest.GetRequestStream())
                    sw.Write(body)
                End Using

                Dim FinalResponse As HttpWebResponse
                FinalResponse = DirectCast(webRequest.GetResponse(), HttpWebResponse)
                Dim dataStream As Stream = FinalResponse.GetResponseStream
                Dim dataReader As New StreamReader(dataStream)
                Dim ResponseStr As String = dataReader.ReadToEnd
                'Me.Setup.Log.AddMessage("JSON Response: " & ResponseStr)
                Dim ResponseJson As JObject = JObject.Parse(ResponseStr)                'item DataFlows hebben we nodig
                Dim DataFlows As JObject = ResponseJson("DataFlows")
                Dim DataFlow As JObject = DataFlows(DataFlowId.ToString)
                Dim State As String = DataFlow("State")
                If State.Trim.ToUpper = "FINISHED" Then
                    Sleeping = False                'as soon as the state is Finished, we'll leave the loop
                    Exit While
                ElseIf State.Trim.ToUpper = "ERROR" Then
                    Throw New Exception("Error while downloading WIWB Rasters. Server response: " & DataFlow.ToString)
                Else
                    Console.WriteLine("Waiting for downloadable file from WIWB. Waiting cycle " & Slept + 1 & ".")
                End If
                System.Threading.Thread.Sleep(5000)
                Slept += 1
                If Slept > 720 Then
                    Throw New Exception("Error: download request could not be completed in 60 minutes and was cancelled. Please check the size of your order.")
                    Sleeping = False 'set timeout at 60 minutes
                End If
            End While

            'finally download the file!
            Dim url As String = "https://wiwb.hydronet.com/api/grids/downloadfile?dataflowid=" & DataFlowId.ToString
            webRequest = TryCast(HttpWebRequest.Create(url), HttpWebRequest)
            webRequest.Method = "GET"
            webRequest.ContentType = "application/json"
            webRequest.Accept = "application/json"
            webRequest.Headers("Authorization") = Convert.ToString("Basic ") & authInfo
            webRequest.Timeout = 10 * 60 * 1000 'set the timeout at 10 minutes

            Dim response As HttpWebResponse = Nothing
            response = DirectCast(webRequest.GetResponse(), HttpWebResponse)
            Dim s As Stream = response.GetResponseStream()

            'Write to disk
            Dim fs As New FileStream(Path, FileMode.Create)

            Dim read As Byte() = New Byte(255) {}
            Dim count As Integer = s.Read(read, 0, read.Length)
            While count > 0
                fs.Write(read, 0, count)
                count = s.Read(read, 0, read.Length)
            End While

            'Close everything
            fs.Close()
            s.Close()
            response.Close()

            Return True

        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Throw New Exception("Could not retrieve grids from API.")
            Return Nothing
        End Try
    End Function


End Class

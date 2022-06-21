Option Explicit On
Imports METEOBAS.General
Imports METEOBAS.GeneralFunctions
Imports GemBox.Spreadsheet

Public Class clsSobekTable

    Friend ID As String
    Friend Multiplier As Double 'a factor that can be used (optionally) to aggregate the content of multiple tables to one
    Friend DateValStrings As New Dictionary(Of String, String)
    Friend Dates As New Dictionary(Of String, DateTime)
    Friend XValues As New Dictionary(Of String, Single) 'als het geen tijdtabel is
    Friend Values1 As New Dictionary(Of String, Single)
    Friend Values2 As New Dictionary(Of String, Single)
    Friend Values3 As New Dictionary(Of String, Single)
    Friend Values4 As New Dictionary(Of String, Single)
    Friend Values5 As New Dictionary(Of String, Single)
    Friend Values6 As New Dictionary(Of String, Single)
    Friend pdin1 As Integer '0 = continuous, 1 = block
    Friend pdin2 As Integer '0 = no return period, 1 = return period
    Friend PDINPeriod As String

    Public TimeStepSeconds As Integer

    'here comes a list of objects that is used for preprocessing of elevation data (collecting and sorting data)
    'Elevationcollection stores elevation values (= the KEY!) and the number of times they occur (= VALUE!)
    Friend ElevationCollection As New Dictionary(Of Single, Long)
    Friend SortedElevation As New Dictionary(Of Single, Long)

    Private setup As clsSetup

    Friend Sub AddPrefix(ByVal Prefix As String)
        ID = Prefix & ID
    End Sub

    Public Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
    End Sub


    Public Sub addToValues(ByVal FieldIdx As Integer, ByVal addValue As Double)
        Dim myDict As Dictionary(Of String, Single)
        Dim i As Long
        myDict = getDictionary(FieldIdx)
        For i = 0 To myDict.Count - 1
            myDict(i) += addValue
        Next
    End Sub

    Public Function IsIncreasing() As Boolean
        Dim i As Long
        For i = 0 To XValues.Values.Count - 2
            If XValues.Values(i + 1) < XValues.Values(i) Then Return False
        Next
        For i = 0 To Values1.Values.Count - 2
            If Values1.Values(i + 1) < Values1.Values(i) Then Return False
        Next
        Return True
    End Function

    Public Sub calcTimeStepSize()
        If Dates.Count > 1 Then TimeStepSeconds = Dates.Values(1).Subtract(Dates.Values(0)).TotalSeconds
    End Sub

    Public Function getnDataCols() As Integer
        Dim n As Integer = 0
        If Values1.Count > 0 Then n += 1
        If Values2.Count > 0 Then n += 1
        If Values3.Count > 0 Then n += 1
        If Values4.Count > 0 Then n += 1
        If Values5.Count > 0 Then n += 1
        If Values6.Count > 0 Then n += 1
        Return n
    End Function

    Private Function getDictionary(ByVal myNum As Integer) As Dictionary(Of String, Single)
        Select Case myNum
            Case Is = 0
                Return XValues
            Case Is = 1
                Return Values1
            Case Is = 2
                Return Values2
            Case Is = 3
                Return Values3
            Case Is = 4
                Return Values4
            Case Is = 5
                Return Values5
            Case Is = 6
                Return Values6
            Case Else
                Return Nothing
        End Select
    End Function

    Public Function IsSymmetric(ByVal DictNum As Long) As Boolean
        'finds out if the numbers in the given dictionary are symmetric
        'so walk from left to right and vice versa
        Dim i As Long, j As Long
        Dim myDict As Dictionary(Of String, Single)
        myDict = getDictionary(DictNum)

        For i = 0 To myDict.Values.Count - 1
            j = myDict.Values.Count - i - 1
            If myDict.Values(i) <> myDict.Values(j) Then
                Return False
            ElseIf i >= j Then
                Return True
            End If
        Next

        Return True

    End Function



    Public Function RemoveDuplicates(ByVal TableNum As Integer) As Boolean
        Dim myDict As Dictionary(Of String, Single) = getDictionary(TableNum)
        Dim i As Long, Done As Boolean, NoDoublesFound As Boolean = True

        While Not Done
            Done = True
            For i = 0 To myDict.Values.Count - 2
                If myDict.Values(i) = myDict.Values(i + 1) Then
                    If XValues.Count > 0 Then XValues.Remove(myDict.Keys(i))
                    If Values1.Count > 0 Then Values1.Remove(myDict.Keys(i))
                    If Values2.Count > 0 Then Values2.Remove(myDict.Keys(i))
                    If Values3.Count > 0 Then Values3.Remove(myDict.Keys(i))
                    If Values4.Count > 0 Then Values4.Remove(myDict.Keys(i))
                    If Values5.Count > 0 Then Values5.Remove(myDict.Keys(i))
                    If Values6.Count > 0 Then Values6.Remove(myDict.Keys(i))
                    NoDoublesFound = False
                    Done = False
                    Exit For
                End If
            Next
        End While
        Return NoDoublesFound
    End Function

    Public Function getYforGivenX(ByVal yDict As Integer, ByVal xDict As Integer, ByVal xVal As Double) As Double
        Dim xTable As Dictionary(Of String, Single) = getDictionary(xDict)
        Dim yTable As Dictionary(Of String, Single) = getDictionary(yDict)
        Dim i As Long
        Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double

        Try
            'date: 20-2-2015
            'author: Siebe Bosch
            'description: gives the corresponding value from another dictionary, given a value in the first dictionary
            'e.g. gets the elevation value that belongs to a given distance in an yz-table

            For i = 0 To xTable.Values.Count - 2
                x1 = xTable.Values(i)
                x2 = xTable.Values(i + 1)
                y1 = yTable.Values(i)
                y2 = yTable.Values(i + 1)

                If xVal >= x1 AndAlso xVal <= x2 Then
                    Return Me.setup.GeneralFunctions.Interpolate(x1, y1, x2, y2, xVal)
                End If
            Next

            'if we end up here, it's outside the range
            If xVal < xTable.Values(0) Then
                Return yTable.Values(0)
            ElseIf xVal > xTable.Values(xTable.Values.Count - 1) Then
                Return yTable.Values(yTable.Values.Count - 1)
            Else
                Throw New Exception("Error in function getYforGivenX in class clsSobekTable. Requested value not found in table.")
            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function getYForLowestX(ByVal yDict As Integer, ByVal xDict As Integer)
        Dim xTable As Dictionary(Of String, Single) = getDictionary(xDict)
        Dim yTable As Dictionary(Of String, Single) = getDictionary(yDict)
        Dim i As Long, minVal As Double = 9999999999999999, minIdx As Long = -1
        Dim nLowest As Integer, yVal As Double

        'date: 20-2-2015
        'author: Siebe Bosch
        'description: retrieves the corresponding value from another dictionary to the lowest value from one dictionary
        'e.g. gets the distance that belongs to the lowest elevation in an yz-table

        For i = 0 To xTable.Values.Count - 1
            If xTable.Values(i) < minVal Then
                minVal = xTable.Values(i)
                minIdx = i
            End If
        Next

        'lowest found. Store the corresponding y-value
        yVal = yTable.Values(minIdx)
        nLowest = 1

        'check if value to the left has the same minimum
        If minIdx > 0 AndAlso xTable.Values(minIdx - 1) = minVal Then
            nLowest += 1
            yVal += yTable.Values(minIdx - 1)
        End If

        'check if value to the right has the same minimum
        If minIdx < xTable.Values.Count - 1 AndAlso xTable.Values(minIdx + 1) = minVal Then
            nLowest += 1
            yVal += yTable.Values(minIdx + 1)
        End If

        'compute the final y-value for the minimum
        yVal = yVal / nLowest
        Return yVal

    End Function

    Public Function getValueForDate(ByVal myDate As Date, ByVal myNum As Integer, ByVal Extrapolate As Boolean) As Double
        Dim myDict As Dictionary(Of String, Single)
        Dim i As Long
        myDict = getDictionary(myNum)

        If myDate < Dates(0) Then
            If Extrapolate Then
                Return myDict(0)
            Else
                Return 0
            End If
        ElseIf myDate > Dates(Dates.Count - 1) Then
            If Extrapolate Then
                Return myDict(myDict.Count - 1)
            Else
                Return 0
            End If
        Else
            For i = 0 To Dates.Count - 1
                If Dates(i) >= myDate Then
                    Return myDict(i)
                End If
            Next
        End If

    End Function


    Public Sub PopulateFromDataTable(ByRef dt As DataTable, ByVal keyColumnIndex As Integer)
        Dim r As Long, key As String
        XValues.Clear()
        Values1.Clear()
        Values2.Clear()
        Values3.Clear()
        Values4.Clear()
        Values5.Clear()
        Values6.Clear()
        For r = 0 To dt.Rows.Count - 1
            key = Str(dt.Rows(r)(keyColumnIndex))  'creates a key for the tables based on a given column in the datatable
            XValues.Add(key, dt.Rows(r)(0))
            If dt.Columns.Count >= 2 Then Values1.Add(key, dt.Rows(r)(1))
            If dt.Columns.Count >= 3 Then Values2.Add(key, dt.Rows(r)(2))
            If dt.Columns.Count >= 4 Then Values3.Add(key, dt.Rows(r)(3))
            If dt.Columns.Count >= 5 Then Values4.Add(key, dt.Rows(r)(4))
            If dt.Columns.Count >= 6 Then Values5.Add(key, dt.Rows(r)(5))
            If dt.Columns.Count >= 7 Then Values6.Add(key, dt.Rows(r)(6))
        Next
    End Sub


    Public Sub DisaggregateTimeSeries(ByVal Divider As Integer, ByVal Val1 As Boolean, ByVal Val2 As Boolean, ByVal Val3 As Boolean, ByVal Val4 As Boolean, ByVal Val5 As Boolean, ByVal val6 As Boolean)
        Dim i As Long, j As Long, k As Long
        Dim newDates As New Dictionary(Of String, Date)
        Dim newValues1 As New Dictionary(Of String, Single)
        Dim newValues2 As New Dictionary(Of String, Single)
        Dim newValues3 As New Dictionary(Of String, Single)
        Dim newValues4 As New Dictionary(Of String, Single)
        Dim newValues5 As New Dictionary(Of String, Single)
        Dim newValues6 As New Dictionary(Of String, Single)

        k = -1
        For i = 0 To Dates.Count - 1
            For j = 0 To Divider - 1
                k += 1
                newDates.Add(Str(k).Trim, Dates(i).Add(TimeSpan.FromSeconds(j * TimeStepSeconds / Divider)))
                If Val1 Then newValues1.Add(Str(k).Trim, Values1(i))
                If Val2 Then newValues1.Add(Str(k).Trim, Values2(i))
                If Val3 Then newValues1.Add(Str(k).Trim, Values3(i))
                If Val4 Then newValues1.Add(Str(k).Trim, Values4(i))
                If Val5 Then newValues1.Add(Str(k).Trim, Values5(i))
                If val6 Then newValues1.Add(Str(k).Trim, Values6(i))
            Next
        Next

        Dates = newDates
        If Val1 Then Values1 = newValues1
        If Val2 Then Values2 = newValues2
        If Val3 Then Values3 = newValues3
        If Val4 Then Values4 = newValues4
        If Val5 Then Values5 = newValues5
        If val6 Then Values6 = newValues6

        Call calcTimeStepSize()

    End Sub

    Public Sub AggregateTimeSeries(ByVal Multiplier As Integer, ByVal Val1 As Boolean, ByVal Val2 As Boolean, ByVal Val3 As Boolean, ByVal Val4 As Boolean, ByVal Val5 As Boolean, ByVal val6 As Boolean)
        Dim i As Long, j As Long, k As Long
        Dim newDates As New Dictionary(Of String, Date)
        Dim newValues1 As New Dictionary(Of String, Single)
        Dim newValues2 As New Dictionary(Of String, Single)
        Dim newValues3 As New Dictionary(Of String, Single)
        Dim newValues4 As New Dictionary(Of String, Single)
        Dim newValues5 As New Dictionary(Of String, Single)
        Dim newValues6 As New Dictionary(Of String, Single)
        Dim Sum1 As Double, Sum2 As Double, Sum3 As Double, Sum4 As Double, Sum5 As Double, Sum6 As Double

        k = -1
        For i = 0 To Dates.Count - 1 Step Multiplier
            k += 1
            newDates.Add(Str(k).Trim, Dates(i))
            Sum1 = 0
            Sum2 = 0
            Sum3 = 0
            Sum4 = 0
            Sum5 = 0
            Sum6 = 0

            For j = 0 To Multiplier - 1
                If Val1 Then Sum1 += Values1(k + j)
                If Val2 Then Sum2 += Values2(k + j)
                If Val3 Then Sum3 += Values3(k + j)
                If Val4 Then Sum4 += Values4(k + j)
                If Val5 Then Sum5 += Values5(k + j)
                If val6 Then Sum6 += Values6(k + j)
            Next

            If Val1 Then newValues1.Add(Str(k).Trim, Val1 / Multiplier)
            If Val2 Then newValues2.Add(Str(k).Trim, Val2 / Multiplier)
            If Val3 Then newValues3.Add(Str(k).Trim, Val3 / Multiplier)
            If Val4 Then newValues4.Add(Str(k).Trim, Val4 / Multiplier)
            If Val5 Then newValues5.Add(Str(k).Trim, Val5 / Multiplier)
            If val6 Then newValues6.Add(Str(k).Trim, val6 / Multiplier)

        Next

        Dates = newDates
        If Val1 Then Values1 = newValues1
        If Val2 Then Values2 = newValues2
        If Val3 Then Values3 = newValues3
        If Val4 Then Values4 = newValues4
        If Val5 Then Values5 = newValues5
        If val6 Then Values6 = newValues6

        Call calcTimeStepSize()

    End Sub


    Public Sub FromArray(ByVal Vals() As Single, ByVal ValIdx As Integer)
        'puts the values of an array in the requested values list
        Dim Values As Dictionary(Of String, Single), i As Long
        Values = getDictionary(ValIdx)

        For i = 0 To Values.Count - 1
            Values(i) = Vals(i)
        Next

    End Sub

    Public Function getPercentageValue(ByVal myPerc As Integer, ByVal SearchListNum As Integer, ByVal ReturnListNum As Integer) As Single
        'this function searches a percentage in one list and returns the corresponding values from another list
        Dim SearchList As New Dictionary(Of String, Single)
        Dim ReturnList As New Dictionary(Of String, Single)

        SearchList = getDictionary(SearchListNum)
        ReturnList = getDictionary(ReturnListNum)

        Dim maxVal As Single = getMaxValue(SearchListNum)
        Dim minVal As Single = getMinValue(SearchListNum)
        Dim X3 As Single = myPerc / 100 * maxVal
        Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
        Dim i As Long

        For i = 0 To SearchList.Count - 2
            X1 = SearchList.Values(i)
            X2 = SearchList.Values(i + 1)
            Y1 = ReturnList.Values(i)
            Y2 = ReturnList.Values(i + 1)
            If X3 <= X2 AndAlso X3 >= X1 Then
                Return setup.GeneralFunctions.Interpolate(X1, Y1, X2, Y2, X3)
            End If
        Next

        If X3 <= minVal Then Return ReturnList.Values(0)
        If X3 >= maxVal Then Return ReturnList.Values(ReturnList.Values.Count - 1)

    End Function

    Public Function getMaxValue(ByVal ValuesListNumber As Integer) As Single
        Select Case ValuesListNumber
            Case Is = 0
                Return XValues.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 1
                Return Values1.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 2
                Return Values2.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 3
                Return Values3.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 4
                Return Values4.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 5
                Return Values5.Values(getMaxValueIdx(ValuesListNumber))
            Case Is = 6
                Return Values6.Values(getMaxValueIdx(ValuesListNumber))
        End Select

    End Function

    Public Function getMinValue(ByVal ValuesListNumber As Integer) As Single
        Select Case ValuesListNumber
            Case Is = 0
                Return XValues.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 1
                Return Values1.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 2
                Return Values2.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 3
                Return Values3.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 4
                Return Values4.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 5
                Return Values5.Values(GetMinValueIdx(ValuesListNumber))
            Case Is = 6
                Return Values6.Values(GetMinValueIdx(ValuesListNumber))
        End Select

    End Function

    Friend Sub ClearValues(ByVal ValIdx As Integer)

        Dim Values As New Dictionary(Of String, Single)
        Dim i As Long

        Select Case ValIdx
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        For i = 0 To Values.Count - 1
            Values(i) = 0
        Next

    End Sub

    Friend Sub MovingAverage(ByVal ValIdx As Long, ByVal nSteps As Long)
        Dim myAvg As New Dictionary(Of String, Single)
        Dim Values As New Dictionary(Of String, Single)
        Dim i As Long, j As Long, mySum As Single
        Dim radius As Integer = Me.setup.GeneralFunctions.RoundUD(nSteps / 2, 0, False)

        Select Case ValIdx
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        'calculate the moving average
        For i = 0 To radius - 1
            myAvg(i) = Values(i)
        Next
        For i = radius To Values.Count - 1 - radius
            mySum = 0
            For j = i - radius To i + radius
                mySum += Values(j)
            Next
            myAvg(i) = mySum / (2 * radius + 1)
        Next
        For i = Values.Count - radius To Values.Count - 1
            myAvg(i) = Values(i)
        Next

        'now copy the moving average back to the values
        For i = 0 To Values.Count - 1
            Values(i) = myAvg(i)
        Next


    End Sub

    Friend Function getavgValue(ByVal ValIdx As Long, Optional ByVal startTS As Long = 0, Optional ByVal EndTS As Long = 0) As Double
        Dim Values As New Dictionary(Of String, Single)
        Dim i As Long, Sum As Double

        Select Case ValIdx
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        If EndTS = 0 Then
            EndTS = Values.Count - 1
        ElseIf EndTS > Values.Count - 1 Then
            EndTS = Values.Count - 1
        End If

        For i = startTS To EndTS
            Sum += Values(i)
        Next

        If Values.Count > 0 Then
            Return Sum / Values.Count
        Else
            Return 0
        End If

    End Function

    Friend Function Clone() As clsSobekTable

        Dim newTable As New clsSobekTable(Me.setup)
        Dim myKey As String = ""
        Dim i As Long

        newTable.ID = ID
        newTable.pdin1 = pdin1
        newTable.pdin2 = pdin2
        newTable.PDINPeriod = PDINPeriod

        For i = 0 To DateValStrings.Count - 1
            myKey = DateValStrings.Keys(i)
            newTable.DateValStrings.Add(myKey, DateValStrings.Item(myKey))
        Next

        For i = 0 To Dates.Count - 1
            myKey = Dates.Keys(i)
            newTable.Dates.Add(myKey, Dates.Item(myKey))
        Next

        For i = 0 To XValues.Count - 1
            myKey = XValues.Keys(i)
            newTable.XValues.Add(myKey, XValues.Item(myKey))
        Next

        For i = 0 To Values1.Count - 1
            myKey = Values1.Keys(i)
            newTable.Values1.Add(myKey, Values1.Item(myKey))
        Next

        For i = 0 To Values2.Count - 1
            myKey = Values2.Keys(i)
            newTable.Values2.Add(myKey, Values2.Item(myKey))
        Next

        For i = 0 To Values3.Count - 1
            myKey = Values3.Keys(i)
            newTable.Values3.Add(myKey, Values3.Item(myKey))
        Next

        For i = 0 To Values4.Count - 1
            myKey = Values4.Keys(i)
            newTable.Values4.Add(myKey, Values4.Item(myKey))
        Next

        For i = 0 To Values5.Count - 1
            myKey = Values5.Keys(i)
            newTable.Values5.Add(myKey, Values5.Item(myKey))
        Next

        For i = 0 To Values6.Count - 1
            myKey = Values6.Keys(i)
            newTable.Values6.Add(myKey, Values6.Item(myKey))
        Next

        newTable.TimeStepSeconds = TimeStepSeconds
        Return newTable

    End Function

    Friend Sub AjustValuesByPercentage(ByVal Perc As Double, ByVal ValIdx As Integer, Optional ByVal mystartts As Long = -1, Optional ByVal myendts As Long = -1)
        'adjusts all values of a given collection by a given percentage
        Dim Values As New Dictionary(Of String, Single)
        Dim i As Long, startts As Long, endts As Long

        'select the values collection that needs adjustment
        Select Case ValIdx
            Case Is = 0
                Values = XValues
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        'set the starting timestep for the adjustments
        If mystartts < 0 Then
            startts = 0
        Else
            startts = mystartts
        End If

        'set the end timestep for the adjustments
        If myendts < 0 Then
            endts = Values.Count - 1
        Else
            endts = myendts
        End If

        'perform the actual adjustments
        For i = startts To endts
            Values(i) = Values(i) * (100 + Perc) / 100
        Next

    End Sub

    Friend Function Shift(ByVal ts As Integer, ByVal valIdx As Integer) As Boolean

        Dim i As Long, j As Long
        Dim tmpStor() As Single 'a container for the temporarily stored values that need shifting over the edges

        Dim Values As New Dictionary(Of String, Single)
        Select Case valIdx
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        If ts < 0 Then
            'first remember the part that will be cut off. it has to be glued back to the end after shifting
            ReDim tmpStor(-ts - 1)
            For i = 0 To -ts - 1
                tmpStor(i) = Values.Values(i)
            Next
            'now move all other items to the left
            For i = -ts To Values.Count - 1
                Values(i - -ts) = Values(i)
            Next
            'glue back the cutoff part
            j = -1
            For i = Values1.Count - -ts To Values1.Count - 1
                j += 1
                Values(i) = tmpStor(j)
            Next
            Return True
        ElseIf ts > 0 Then
            'first remember the part that will be cut off. it has to be glued back to the start after shifting
            ReDim tmpStor(ts - 1)
            j = -1
            For i = Values.Count - ts To Values.Count - 1
                j += 1
                tmpStor(j) = Values.Values(i)
            Next
            'now move all other items to the right
            For i = Values.Count - 1 To ts Step -1
                Values(i) = Values(i - ts)
            Next
            'glue back the cutoff part
            j = -1
            For i = 0 To ts - 1
                j += 1
                Values(i) = tmpStor(j)
            Next
            Return True
        End If

    End Function

    Public Function SortElevationData(ByVal ElevationMultiplier As Integer, ByVal CellSize As Double, Optional ByVal DivideByReachLength As Double = 1) As Boolean

        'this routine sorts the collected elevation data that has not yet been stored in the
        'standard objects Xvalues, Values1 etc., but that's still in the ElevationData dictionary
        'after sorting it will assign the results to the appropriate dictionaries
        Dim i As Long, nCum As Long



        If ElevationCollection.Count > 0 Then

            'sort the dictionary containing collected data by key
            Dim Sorted = From pair In ElevationCollection Order By pair.Key
            SortedElevation = Sorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

            'write all elevation data to a storage table, do this in the stepsize defined above
            For i = 0 To SortedElevation.Count - 1
                nCum += SortedElevation.Values(i)
                AddDataPair(2, SortedElevation.Keys(i) * ElevationMultiplier, nCum * CellSize * CellSize / DivideByReachLength)
            Next

            'de lijst met elevationdata mag worden leeggemaakt. Wordt hierna niet langer gebruikt
            ElevationCollection = Nothing
            SortedElevation = Nothing
            Return True
        Else
            Return False
        End If

    End Function

    Friend Function GetMinValueIdx(ByVal ValuesListNum As Integer, Optional ByVal StartIdx As Integer = -1, Optional ByVal EndIdx As Integer = -1) As Long
        Dim Values As New Dictionary(Of String, Single)
        Dim i As Long, minVal As Single = 99999999, minIdx As Long

        Values = getDictionary(ValuesListNum) 'get the dictionary of values that applies
        If StartIdx < 0 Then StartIdx = 0
        If EndIdx < 0 Then EndIdx = Values.Count - 1

        For i = StartIdx To EndIdx
            If Values.Values(i) < minVal Then
                minVal = Values.Values(i)
                minIdx = i
            End If
        Next
        Return minIdx
    End Function

    Public Function getMaxValueIdx(ByVal ValuesListNum As Integer, Optional ByVal StartIdx As Integer = -1, Optional ByVal EndIdx As Integer = -1) As Long

        Dim i As Long, myMax As Double = -99999999999, myIdx As Long = 0
        Dim Values As Dictionary(Of String, Single)

        Select Case ValuesListNum
            Case Is = 0
                Values = XValues
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
            Case Else
                Return 0
        End Select

        If StartIdx < 0 Then StartIdx = 0
        If EndIdx < 0 Then EndIdx = Values.Count - 1

        For i = StartIdx To EndIdx
            If Values.Values(i) > myMax Then
                myMax = Values.Values(i)
                myIdx = i
            End If
        Next
        Return myIdx

    End Function

    Friend Function getIntersectionXVal(ByVal ValIdx As Long, ByVal Upward As Boolean, ByVal Downward As Boolean, ByVal SearchValue As Double, ByVal StartIdx As Long, ByVal EndIdx As Long, ByRef XVal As Double) As Boolean
        'Searches the surrounding index points for a given Value
        Dim Values As New Dictionary(Of String, Single)
        Dim DeepestIdx As Integer, HighestIdx As Integer
        Dim DeepestVal As Double, HighestVal As Double
        Dim i As Long

        Select Case ValIdx
            Case Is = 1
                Values = Values1
            Case Is = 2
                Values = Values2
            Case Is = 3
                Values = Values3
            Case Is = 4
                Values = Values4
            Case Is = 5
                Values = Values5
            Case Is = 6
                Values = Values6
        End Select

        DeepestIdx = GetMinValueIdx(1, StartIdx, EndIdx)
        HighestIdx = getMaxValueIdx(1, StartIdx, EndIdx)
        DeepestVal = Values.Values(DeepestIdx)
        HighestVal = Values.Values(HighestIdx)

        If SearchValue > HighestVal Then
            Me.setup.Log.AddWarning("Could not interpolate in table " & ID & ": value " & SearchValue & " exceeds highest value in table.")
            Return False
        ElseIf SearchValue < DeepestVal Then
            Me.setup.Log.AddWarning("Could not interpolate in table " & ID & ": value " & SearchValue & " is lower than lowest value in table.")
            Return False
        End If

        If Upward Then
            For i = EndIdx - 1 To StartIdx Step -1 'move from right to left!
                If Values.Values(i) <= SearchValue And Values.Values(i + 1) >= SearchValue Then
                    XVal = setup.GeneralFunctions.Interpolate(Values.Values(i), XValues.Values(i), Values.Values(i + 1), XValues.Values(i + 1), SearchValue)
                    Return True
                End If
            Next
        End If

        If Downward Then
            For i = StartIdx To EndIdx - 1
                If Values.Values(i) >= SearchValue And Values.Values(i + 1) <= SearchValue Then
                    XVal = setup.GeneralFunctions.Interpolate(Values.Values(i), XValues.Values(i), Values.Values(i + 1), XValues.Values(i + 1), SearchValue)
                    Return True
                End If
            Next
        End If

        Return False

    End Function

    Public Function getSlopeFromValues1(ByVal fromIdx As Integer, ByVal toIdx As Integer) As Double
        Dim Val1 As Double = Values1.Values(fromIdx)
        Dim Val2 As Double = Values1.Values(toIdx)
        Dim X1 As Double = XValues.Values(fromIdx)
        Dim X2 As Double = XValues.Values(toIdx)

        If Not (X2 - X1) = 0 Then
            Return (Val2 - Val1) / (X2 - X1)
        Else
            Return 0
        End If

    End Function

    Public Function getTotalLength(ByVal FromIdx As Integer, ByVal ToIdx As Integer) As Double

        'Returns the total length of a segment as calculated by pythagoras
        'SQR(dX^2 + dY^2)
        Dim Val1 As Double = Values1.Values(FromIdx)
        Dim Val2 As Double = Values1.Values(ToIdx)
        Dim X1 As Double = XValues.Values(FromIdx)
        Dim X2 As Double = XValues.Values(ToIdx)

        Return Math.Sqrt((X2 - X1) ^ 2 + (Val2 - Val1) ^ 2)

    End Function

    Public Function getBaseLength(ByVal FromIdx As Integer, ByVal ToIdx As Integer) As Double
        'Returns the length of the base (XValues) between two given index points
        Dim X1 As Double = XValues.Values(FromIdx)
        Dim X2 As Double = XValues.Values(ToIdx)

        Return X2 - X1

    End Function

    Public Function getLastValues1Value() As Double
        Return Values1.Values(Values1.Count - 1)
    End Function


    Friend Sub buildDateValStrings(Optional ByVal nVals As Integer = 6, Optional ByVal nDecimals As Integer = 7)
        Dim i As Long
        DateValStrings = New Dictionary(Of String, String)
        Dim myStr As String = ""
        Dim myVal As String

        Select Case nDecimals
            Case Is = 1
                myVal = Format(Values1.Values(i), "0.0")
            Case Is = 2
                myVal = Format(Values1.Values(i), "0.00")
            Case Is = 3
                myVal = Format(Values1.Values(i), "0.000")
            Case Is = 4
                myVal = Format(Values1.Values(i), "0.0000")
            Case Is = 5
                myVal = Format(Values1.Values(i), "0.00000")
            Case Is = 6
                myVal = Format(Values1.Values(i), "0.000000")
            Case Else
                myVal = Format(Values1.Values(i), "0.0000000")
        End Select

        For i = 0 To Dates.Count - 1
            Select Case nVals
                Case Is = 1
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " <"
                Case Is = 2
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " " & Values2.Values(i) & " <"
                Case Is = 3
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " " & Values2.Values(i) & " " & Values3.Values(i) & " <"
                Case Is = 4
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " " & Values2.Values(i) & " " & Values3.Values(i) & " " & Values4.Values(i) & " <"
                Case Is = 5
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " " & Values2.Values(i) & " " & Values3.Values(i) & " " & Values4.Values(i) & " " & Values5.Values(i) & " <"
                Case Is = 6
                    myStr = "'" & Year(Dates.Values(i)) & "/" & Format(Month(Dates.Values(i)), "00") & "/" & Format(Day(Dates.Values(i)), "00") & ";" & Format(Hour(Dates.Values(i)), "00") & ":" & Format(Minute(Dates.Values(i)), "00") & ":" & Format(Second(Dates.Values(i)), "00") & "' " & myVal & " " & Values2.Values(i) & " " & Values3.Values(i) & " " & Values4.Values(i) & " " & Values5.Values(i) & " " & Values6.Values(i) & " <"
            End Select
            DateValStrings.Add(i.ToString, myStr)
        Next

    End Sub

    Friend Function InterpolateXValueFromValues(ByVal Val As Double, Optional ByVal ValIdx As Integer = 1) As Single
        Dim i As Integer
        If ValIdx < 0 Then ValIdx = 1
        If ValIdx > 6 Then ValIdx = 6

        'bepaal eerst uit welke tabel we waarden gaan teruggeven
        Dim SearchDict As Dictionary(Of String, Single) = Nothing
        Select Case ValIdx
            Case Is = 1
                SearchDict = Values1
            Case Is = 2
                SearchDict = Values2
            Case Is = 3
                SearchDict = Values3
            Case Is = 4
                SearchDict = Values4
            Case Is = 5
                SearchDict = Values5
            Case Is = 6
                SearchDict = Values6
        End Select

        If Val <= SearchDict.Values(0) Then
            Return XValues.Values(0)
        ElseIf Val >= SearchDict.Values(SearchDict.Count - 1) Then
            Return XValues.Values(XValues.Values.Count - 1)
        ElseIf SearchDict.Count < 2 Then
            Return setup.GeneralFunctions.Interpolate(SearchDict.Values(0), XValues.Values(0), SearchDict.Values(1), XValues.Values(1), Val)
        Else
            For i = 0 To SearchDict.Values.Count - 2
                If SearchDict.Values(i) <= Val AndAlso Val <= SearchDict.Values(i + 1) Then
                    'interpoleer lineair
                    Return setup.GeneralFunctions.Interpolate(SearchDict.Values(i), XValues.Values(i), SearchDict.Values(i + 1), XValues.Values(i + 1), Val)
                End If
            Next
        End If

    End Function

    Friend Function InterpolateFromDates(ByVal myDate As Date, Optional ByVal valIdx As Integer = 1) As Single
        Dim i As Integer
        If valIdx < 0 Then valIdx = 1
        If valIdx > 6 Then valIdx = 6

        'bepaal eerst uit welke tabel we waarden gaan teruggeven
        Dim ReturnDict As Dictionary(Of String, Single) = Nothing
        Select Case valIdx
            Case Is = 1
                ReturnDict = Values1
            Case Is = 2
                ReturnDict = Values2
            Case Is = 3
                ReturnDict = Values3
            Case Is = 4
                ReturnDict = Values4
            Case Is = 5
                ReturnDict = Values5
            Case Is = 6
                ReturnDict = Values6
        End Select

        If myDate <= Dates.Values(0) Then
            Return ReturnDict.Values(0)
        ElseIf myDate >= Dates.Values(Dates.Count - 1) Then
            Return ReturnDict.Values(ReturnDict.Values.Count - 1)
        ElseIf Dates.Count < 2 Then
            Return setup.GeneralFunctions.Interpolate(Dates.Values(0).ToOADate, ReturnDict.Values(0), Dates.Values(1).ToOADate, ReturnDict.Values(1), myDate.ToOADate)
        Else
            For i = 0 To Dates.Values.Count - 2
                If Dates.Values(i + 1) = myDate Then
                    Return ReturnDict.Values(i + 1)
                ElseIf Dates.Values(i) <= myDate AndAlso Dates.Values(i + 1) >= myDate Then
                    'interpoleer lineair
                    Return setup.GeneralFunctions.Interpolate(Dates.Values(i).ToOADate, ReturnDict.Values(i), Dates.Values(i + 1).ToOADate, ReturnDict.Values(i + 1), myDate.ToOADate)
                End If
            Next
        End If
    End Function

    Friend Function InterpolateFromXValues(ByVal Xval As Double, Optional ByVal ValIdx As Integer = 1) As Single
        Dim i As Integer
        If ValIdx < 0 Then ValIdx = 1
        If ValIdx > 6 Then ValIdx = 6

        'bepaal eerst uit welke tabel we waarden gaan teruggeven
        Dim ReturnDict As Dictionary(Of String, Single) = Nothing
        Select Case ValIdx
            Case Is = 1
                ReturnDict = Values1
            Case Is = 2
                ReturnDict = Values2
            Case Is = 3
                ReturnDict = Values3
            Case Is = 4
                ReturnDict = Values4
            Case Is = 5
                ReturnDict = Values5
            Case Is = 6
                ReturnDict = Values6
        End Select

        If XValues.ContainsKey(Str(Xval)) Then
            Return ReturnDict.Item(Str(Xval))
        ElseIf Xval <= XValues.Values(0) Then
            Return ReturnDict.Values(0)
        ElseIf Xval >= XValues.Values(XValues.Count - 1) Then
            Return ReturnDict.Values(ReturnDict.Values.Count - 1)
        ElseIf XValues.Count < 2 Then
            Return setup.GeneralFunctions.Interpolate(XValues.Values(0), ReturnDict.Values(0), XValues.Values(1), ReturnDict.Values(1), Xval)
        Else
            For i = 0 To XValues.Values.Count - 2
                If Xval >= XValues.Values(i) AndAlso Xval <= XValues.Values(i + 1) Then
                    If XValues.Values(i) = Xval Then
                        Return ReturnDict.Values(i)
                    ElseIf XValues.Values(i + 1) = Xval Then
                        Return ReturnDict.Values(i + 1)
                    ElseIf XValues.Values(i) < Xval AndAlso XValues.Values(i + 1) > Xval Then
                        Return setup.GeneralFunctions.Interpolate(XValues.Values(i), ReturnDict.Values(i), XValues.Values(i + 1), ReturnDict.Values(i + 1), Xval)
                    End If
                End If
            Next
        End If

    End Function

    Friend Sub Read(ByVal myTable As String)
        Dim myRecords() As String, tmp As String
        Dim myDate As DateTime

        'verwijder alle tokens
        myTable = Replace(myTable, "TBLE", "")
        myTable = Replace(myTable, "tble", "")

        myRecords = Split(myTable, "<")
        For i As Integer = 0 To UBound(myRecords) - 1
            Dim j As Integer = 0
            While Not myRecords(i) = ""
                tmp = Me.setup.GeneralFunctions.ParseString(myRecords(i))
                j += 1
                If j = 1 AndAlso InStr(tmp, "/") > 0 Then
                    myDate = Me.setup.GeneralFunctions.ConvertToDateTime(tmp, "yyyy/MM/dd;HH:mm:ss")
                    Me.Dates.Add(Str(i).Trim, myDate)
                ElseIf j = 1 Then
                    Me.XValues.Add(Str(i).Trim, tmp)
                End If

                If j = 2 Then Values1.Add(Str(i).Trim, tmp)
                If j = 3 AndAlso Not tmp = "<" Then Me.Values2.Add(Str(i).Trim, tmp)
                If j = 4 AndAlso Not tmp = "<" Then Me.Values3.Add(Str(i).Trim, tmp)
                If j = 5 AndAlso Not tmp = "<" Then Me.Values4.Add(Str(i).Trim, tmp)
                If j = 6 AndAlso Not tmp = "<" Then Me.Values5.Add(Str(i).Trim, tmp)
            End While
        Next i
    End Sub
    Friend Function GetPeriod() As Integer()
        Dim DateString As String
        DateString = PDINPeriod
        Dim Values(4) As Integer
        Values(1) = Me.setup.GeneralFunctions.ParseString(DateString, ";")
        Values(2) = Me.setup.GeneralFunctions.ParseString(DateString, ":")
        Values(3) = Me.setup.GeneralFunctions.ParseString(DateString, ":")
        Values(4) = Me.setup.GeneralFunctions.ParseString(DateString, ":")
        GetPeriod = Values
    End Function

    Friend Sub BuildFromTargetLevels(ByVal WP As Double, ByVal ZP As Double)

        pdin1 = 1
        pdin2 = 1
        PDINPeriod = "365;00:00:00"

        AddDatevalPair(New Date(2000, 1, 1), WP)
        AddDatevalPair(New Date(2000, 4, 15), ZP)
        AddDatevalPair(New Date(2000, 10, 15), WP)

    End Sub

    Friend Sub AddDataPairKeyByCount(ByVal nVals As Integer, ByVal xval As Double, ByVal val1 As Double, Optional ByVal val2 As Double = 0,
                         Optional ByVal val3 As Double = 0, Optional ByVal val4 As Double = 0,
                         Optional ByVal val5 As Double = 0, Optional ByVal val6 As Double = 0)
        Dim myStr As String = Str(XValues.Count)

        'voegt data aan de tabel toe waarbij de key telkens gelijk is aan het aantal items in de tabel
        'dit zoekt TRAAG, maar maakt het mogelijk om meerdere malen dezelfde X-waarde in de tabel te hebben
        If Not XValues.ContainsKey(myStr) AndAlso Not Values1.ContainsKey(myStr) Then
            If nVals >= 1 Then XValues.Add(myStr, xval)
            If nVals >= 2 Then Values1.Add(myStr, val1)
            If nVals >= 3 Then Values2.Add(myStr, val2)
            If nVals >= 4 Then Values3.Add(myStr, val3)
            If nVals >= 5 Then Values4.Add(myStr, val4)
            If nVals >= 6 Then Values5.Add(myStr, val5)
            If nVals >= 7 Then Values6.Add(myStr, val6)
        End If

    End Sub


    Friend Sub AddDataPair(ByVal nVals As Integer, ByVal xval As Double, ByVal val1 As Double, Optional ByVal val2 As Double = 0,
                           Optional ByVal val3 As Double = 0, Optional ByVal val4 As Double = 0,
                           Optional ByVal val5 As Double = 0, Optional ByVal val6 As Double = 0)
        'Dim myStr As String = Str(XValues.Count)
        Dim myStr As String = Str(Math.Round(xval, 10)).Trim
        'Dim myStr As String = Str(XValues.Count)

        '20121002 siebe: als key een string van de x-waarde ingezet om het zoeken in de tabel sneller te laten lopen
        '20130913 siebe: als key toch maar xvalues.count ingebracht omdat twee records met dezelfde x-waarden anders niet worden ondersteund
        '20131004 siebe: als key toch maar weer de x-waarde ingebracht omdat de zoeksnelheid anders een drama wordt
        If Not XValues.ContainsKey(myStr) AndAlso Not Values1.ContainsKey(myStr) Then
            If nVals >= 1 Then XValues.Add(myStr, xval)
            If nVals >= 2 Then Values1.Add(myStr, val1)
            If nVals >= 3 Then Values2.Add(myStr, val2)
            If nVals >= 4 Then Values3.Add(myStr, val3)
            If nVals >= 5 Then Values4.Add(myStr, val4)
            If nVals >= 6 Then Values5.Add(myStr, val5)
            If nVals >= 7 Then Values6.Add(myStr, val6)
        End If

    End Sub

    Friend Sub ShiftXValues(ByVal ShiftByVal As Double)
        For Each myXVal As Double In XValues.Values
            myXVal += ShiftByVal
        Next
    End Sub

    Friend Sub AddDate(ByVal myDate As DateTime)
        Dim mystr As String = Str(Dates.Count).Trim
        Dates.Add(mystr, myDate)
    End Sub

    Friend Sub AddValue1(ByVal myVal As Double)
        Dim myStr As String = Str(Values1.Count).Trim
        Values1.Add(myStr, myVal)
    End Sub

    Friend Sub AddValue2(ByVal myVal As Double)
        Dim myStr As String = Str(Values2.Count).Trim
        Values2.Add(myStr, myVal)
    End Sub

    Friend Sub AddDatevalPair(ByVal myDate As DateTime, ByVal val1 As Double, Optional ByVal val2 As Double = 0,
                              Optional ByVal val3 As Double = 0, Optional ByVal val4 As Double = 0,
                              Optional ByVal val5 As Double = 0, Optional ByVal val6 As Double = 0, Optional ByVal nVals As Integer = 1)
        Dim myStr As String = Str(Dates.Count).Trim
        Dates.Add(myStr, myDate)
        Select Case nVals
            Case Is = 1
                Values1.Add(myStr, val1)
            Case Is = 2
                Values1.Add(myStr, val1)
                Values2.Add(myStr, val2)
            Case Is = 3
                Values1.Add(myStr, val1)
                Values2.Add(myStr, val2)
                Values3.Add(myStr, val3)
            Case Is = 4
                Values1.Add(myStr, val1)
                Values2.Add(myStr, val2)
                Values3.Add(myStr, val3)
                Values4.Add(myStr, val4)
            Case Is = 5
                Values1.Add(myStr, val1)
                Values2.Add(myStr, val2)
                Values3.Add(myStr, val3)
                Values4.Add(myStr, val4)
                Values5.Add(myStr, val5)
            Case Is = 6
                Values1.Add(myStr, val1)
                Values2.Add(myStr, val2)
                Values3.Add(myStr, val3)
                Values4.Add(myStr, val4)
                Values5.Add(myStr, val5)
                Values6.Add(myStr, val6)
        End Select

    End Sub

    Friend Sub UpdateDateValPair(ByVal myDate As DateTime, ByVal val1 As Double, Optional ByVal val2 As Double = 0,
                              Optional ByVal val3 As Double = 0, Optional ByVal val4 As Double = 0,
                              Optional ByVal val5 As Double = 0, Optional ByVal val6 As Double = 0, Optional ByVal nVals As Integer = 1)
        Dim i As Long, mykey As String
        For i = 0 To Dates.Count - 1
            If Dates.Values(i) = myDate Then
                mykey = Dates.Keys(i)
                If nVals >= 1 Then Values1.Item(mykey) = val1
                If nVals >= 2 Then Values2.Item(mykey) = val2
                If nVals >= 3 Then Values3.Item(mykey) = val3
                If nVals >= 4 Then Values4.Item(mykey) = val4
                If nVals >= 5 Then Values5.Item(mykey) = val5
                If nVals >= 6 Then Values6.Item(mykey) = val6
                Exit Sub
            End If
        Next

    End Sub


    Friend Sub UpdateDateValPairByKey(ByVal myKey As String, ByVal val1 As Double, Optional ByVal val2 As Double = 0,
                              Optional ByVal val3 As Double = 0, Optional ByVal val4 As Double = 0,
                              Optional ByVal val5 As Double = 0, Optional ByVal val6 As Double = 0, Optional ByVal nVals As Integer = 1)
        If nVals >= 1 Then Values1.Item(myKey) = val1
        If nVals >= 2 Then Values2.Item(myKey) = val2
        If nVals >= 3 Then Values3.Item(myKey) = val3
        If nVals >= 4 Then Values4.Item(myKey) = val4
        If nVals >= 5 Then Values5.Item(myKey) = val5
        If nVals >= 6 Then Values6.Item(myKey) = val6
    End Sub

    Friend Function getValue1(ByVal XVal As Double) As Double
        'interpoleert een waarde uit de XValues/Values1-dataset
        Dim i As Integer
        Dim myStr As String = Str(Math.Round(XVal, 10)).Trim

        If Values1.ContainsKey(myStr) Then Return Values1.Item(myStr)

        If XVal < XValues.Values(0) Then
            Return Values1.Values(0)
        ElseIf XVal > XValues.Values(XValues.Count - 1) Then
            Return Values1.Values(Values1.Count - 1)
        End If

        For i = 0 To XValues.Count - 2
            If XValues.Values(i) = XVal Then
                Return Values1.Values(i)
            ElseIf XValues.Values(i + 1) = XVal Then
                Return Values1.Values(i + 1)
            ElseIf XValues.Values(i) < XVal AndAlso XValues.Values(i + 1) > XVal Then
                Return setup.GeneralFunctions.Interpolate(XValues.Values(i), Values1.Values(i), XValues.Values(i + 1), Values1.Values(i + 1), XVal)
            End If
        Next
    End Function

    Public Function getIdxFromValue(ByVal DictNum As Integer, ByVal myVal As Double) As Integer
        'Author: Siebe Bosch
        'Date: 17-6-2013
        'Description: Finds the indexnumber for the record that closest matches a given value in a given table
        Dim i As Integer, minDist As Double = 9999999999, Dist As Double
        Dim myIdx As Integer

        Dim myDict As Dictionary(Of String, Single)
        myDict = getDictionary(DictNum)

        For i = 0 To myDict.Count - 1
            Dist = Math.Abs(myDict.Values(i) - myVal)
            If Dist < minDist Then
                myIdx = i
                minDist = Dist
            End If
        Next

        Return myIdx

    End Function

    Public Function getNearestMaxFromValues1(ByVal StartIdx As Integer) As Integer
        'Author: Siebe Bosch
        'Date: 17-6-2013
        'Description: Finds the indexnumber for the maximumvalue from Values1 that lies closest to a given startindex
        Dim minDist As Double = 99999999
        Dim prevVal As Double, curVal As Double, nextVal As Double
        Dim curDist As Double, PeakIdx As Integer = StartIdx
        Dim i As Long

        For i = 1 To Values1.Values.Count - 2
            prevVal = Values1.Values(i - 1)
            curVal = Values1.Values(i)
            nextVal = Values1.Values(i + 1)
            If curVal > prevVal AndAlso curVal > nextVal Then
                curDist = Math.Abs(XValues.Values(i) - XValues.Values(StartIdx))
                If curDist < minDist Then
                    PeakIdx = i
                    minDist = curDist
                End If
            End If
        Next

        Return PeakIdx

    End Function

    Public Function getFrontValleyIdx(ByVal ListNumber As Integer, ByVal StartIdx As Integer, ByVal minDist As Double) As Integer
        Dim prevVal As Double, curVal As Double, nextVal As Double, curDist As Double
        Dim PeakIdx As Integer = StartIdx
        Dim myDict As New Dictionary(Of String, Single)
        Dim i As Long

        Select Case ListNumber
            Case Is = 0
                myDict = XValues
            Case Is = 1
                myDict = Values1
            Case Is = 2
                myDict = Values2
            Case Is = 3
                myDict = Values3
            Case Is = 4
                myDict = Values4
            Case Is = 5
                myDict = Values5
            Case Is = 6
                myDict = Values6
        End Select

        If StartIdx > 0 Then
            For i = StartIdx - 1 To 0 Step -1
                prevVal = myDict.Values(i + 1)
                curVal = myDict.Values(i)
                nextVal = myDict.Values(i - 1)
                curDist = Math.Abs(XValues.Values(StartIdx) - XValues.Values(i))
                If curVal <= prevVal AndAlso curVal < nextVal AndAlso curDist >= minDist Then
                    Return i
                End If
            Next
        End If
        Return 0

    End Function

    Public Function getBackValleyIdx(ByVal ListNumber As Integer, ByVal StartIdx As Integer, ByVal minDist As Double) As Integer
        Dim prevVal As Double, curVal As Double, nextVal As Double, curDist As Double
        Dim PeakIdx As Integer = StartIdx
        Dim myDict As New Dictionary(Of String, Single)
        Dim i As Long

        Select Case ListNumber
            Case Is = 0
                myDict = XValues
            Case Is = 1
                myDict = Values1
            Case Is = 2
                myDict = Values2
            Case Is = 3
                myDict = Values3
            Case Is = 4
                myDict = Values4
            Case Is = 5
                myDict = Values5
            Case Is = 6
                myDict = Values6
        End Select

        If StartIdx < XValues.Count - 1 Then
            For i = StartIdx + 1 To XValues.Count - 1
                prevVal = Values1.Values(i - 1)
                curVal = Values1.Values(i)
                nextVal = Values1.Values(i + 1)
                curDist = Math.Abs(XValues.Values(StartIdx) - XValues.Values(i))
                If curVal <= prevVal AndAlso curVal < nextVal AndAlso curDist >= minDist Then
                    Return i
                End If
            Next
        End If
        Return XValues.Count - 1

    End Function

    Public Function MovingAverageValues1() As clsSobekTable
        Dim newTable = New clsSobekTable(Me.setup)
        Dim myAvg As Double
        Dim i As Long

        'simply add the values that don't add up to complete moving averages
        For i = 0 To 0
            myAvg = Values1.Values(i)
            newTable.AddDataPair(2, XValues.Values(i), myAvg)
        Next

        'calculate the actual moving average
        For i = 1 To XValues.Count - 2
            myAvg = (Values1.Values(i - 1) + Values1.Values(i) + Values1.Values(i + 1)) / 3
            newTable.AddDataPair(2, XValues.Values(i), myAvg)
        Next

        'simply add the values that don't add up to complete moving averages
        For i = XValues.Count - 1 To XValues.Count - 1
            myAvg = Values1.Values(i)
            newTable.AddDataPair(2, XValues.Values(i), myAvg)
        Next

        Return newTable
    End Function

    Public Function GetLastValue(ByVal ValuesListNum As Integer) As Double
        Dim myValues As Dictionary(Of String, Single)
        Dim Idx As Long

        Select Case ValuesListNum
            Case Is = 0
                myValues = XValues
            Case Is = 1
                myValues = Values1
            Case Is = 2
                myValues = Values2
            Case Is = 3
                myValues = Values3
            Case Is = 4
                myValues = Values4
            Case Is = 5
                myValues = Values5
            Case Is = 6
                myValues = Values6
            Case Else
                Return 0
        End Select

        Idx = myValues.Count - 1
        Return myValues.Values(Idx)

    End Function

    Public Function getMaxValueIdxFromStartPoint(ByVal ValuesListNum As Integer, ByVal StartIdx As Long, ByVal Left As Boolean, ByVal Right As Boolean) As Long

        Dim i As Long, myMaxLeft As Double = -99999999999, myMaxRight As Double = -9999999999, maxLeftIdx As Long = 0, maxRightIdx As Long = 0
        Dim myValues As Dictionary(Of String, Single)

        Select Case ValuesListNum
            Case Is = 1
                myValues = Values1
            Case Is = 2
                myValues = Values2
            Case Is = 3
                myValues = Values3
            Case Is = 4
                myValues = Values4
            Case Is = 5
                myValues = Values5
            Case Is = 6
                myValues = Values6
            Case Else
                Return 0
        End Select

        If Left Then
            For i = StartIdx - 1 To 0 Step -1
                If myValues.Values(i) > myMaxLeft Then
                    myMaxLeft = myValues.Values(i)
                    maxLeftIdx = i
                End If
            Next
        End If

        If Right Then
            For i = StartIdx + 1 To myValues.Values.Count - 1
                If myValues.Values(i) > myMaxRight Then
                    myMaxRight = myValues.Values(i)
                    maxRightIdx = i
                End If
            Next
        End If

        If Left And Right Then
            If myMaxLeft > myMaxRight Then
                Return maxLeftIdx
            Else
                Return maxRightIdx
            End If
        ElseIf Left Then
            Return maxLeftIdx
        ElseIf Right Then
            Return maxRightIdx
        End If

    End Function

    Public Function getRightMinSlopeIdx(ByVal StartIdx As Long, ByVal XValuesDictionaryNumber As Integer, ByVal YValuesDictionaryNumber As Integer, ByRef mySlope As Double) As Long

        'returns the indexnumber for the location in the table with the minimum gradient ((y2-y1)/(x2-x1))
        'starting at a given indexpoint
        Dim XValues As New Dictionary(Of String, Single)
        Dim YValues As New Dictionary(Of String, Single)
        Dim Min As Double = 9999999999
        Dim minIdx As Integer
        Dim dX As Double, dY As Double
        Dim i As Long

        Select Case XValuesDictionaryNumber
            Case Is = 1
                XValues = Values1
            Case Is = 2
                XValues = Values2
            Case Is = 3
                XValues = Values3
            Case Is = 4
                XValues = Values4
            Case Is = 5
                XValues = Values5
            Case Is = 6
                XValues = Values6
        End Select

        Select Case YValuesDictionaryNumber
            Case Is = 1
                YValues = Values1
            Case Is = 2
                YValues = Values2
            Case Is = 3
                YValues = Values3
            Case Is = 4
                YValues = Values4
            Case Is = 5
                YValues = Values5
            Case Is = 6
                YValues = Values6
        End Select

        For i = StartIdx To XValues.Count - 2
            dX = XValues.Values(i + 1) - XValues.Values(i)
            dY = YValues.Values(i + 1) - YValues.Values(i)
            If Not dX = 0 Then
                If dY / dX > Min Then
                    Min = dY / dX
                    minIdx = i
                    mySlope = Min
                End If
            End If
        Next

        Return minIdx

    End Function


    Public Function getRightToeIdx(ByVal StartIdx As Long, ByVal XValuesDictionaryNumber As Integer, ByVal YValuesDictionaryNumber As Integer) As Long

        'specially for dike profiles
        'returns the indexnumber for the location in the table where the dike's toe is located (dy/dx > 0)
        'starting at a given indexpoint
        'if not found, the last value is returned
        Dim XValues As New Dictionary(Of String, Single)
        Dim YValues As New Dictionary(Of String, Single)
        Dim dX As Double, dY As Double
        Dim i As Long

        Select Case XValuesDictionaryNumber
            Case Is = 1
                XValues = Values1
            Case Is = 2
                XValues = Values2
            Case Is = 3
                XValues = Values3
            Case Is = 4
                XValues = Values4
            Case Is = 5
                XValues = Values5
            Case Is = 6
                XValues = Values6
        End Select

        Select Case YValuesDictionaryNumber
            Case Is = 1
                YValues = Values1
            Case Is = 2
                YValues = Values2
            Case Is = 3
                YValues = Values3
            Case Is = 4
                YValues = Values4
            Case Is = 5
                YValues = Values5
            Case Is = 6
                YValues = Values6
        End Select

        For i = StartIdx To XValues.Count - 2
            dX = XValues.Values(i + 1) - XValues.Values(i)
            dY = YValues.Values(i + 1) - YValues.Values(i)
            If Not dX = 0 Then
                If dY / dX > 0 Then
                    Return i
                End If
            End If
        Next

        'not found, so assume the last point represents the toe
        Return XValues.Count - 1

    End Function

    Public Sub writeToExcelWorkSheet(ByRef ws As ExcelWorksheet, ByVal r As Long, ByVal c As Long, ByVal XValName As String, ByVal Val1Name As String)

        Dim i As Long

        If XValues.Count > 0 Then
            ws.Cells(r, c).Value = XValName
            For i = 0 To XValues.Count - 1
                ws.Cells(r + i + 1, c).Value = XValues.Values(i)
            Next
        End If

        If Values1.Count > 0 Then
            c += 1
            ws.Cells(r, c).Value = Val1Name
            For i = 0 To Values1.Count - 1
                ws.Cells(r + i + 1, c).Value = Values1.Values(i)
            Next
        End If

    End Sub

    Public Sub writeTimeTableContentsToFile(ByRef myWriter As System.IO.StreamWriter, ByVal col1 As Boolean, ByVal col2 As Boolean, ByVal col3 As Boolean, ByVal col4 As Boolean, ByVal col5 As Boolean, ByVal col6 As Boolean)
        Dim i As Long, mystr As String
        For i = 0 To Dates.Count - 1
            mystr = "'" & Format(Dates.Item(i), "YYYY/MM/DD:HH:mm:SS")

            If col1 Then mystr &= " " & Values1.Values(i)
            If col2 Then mystr &= " " & Values2.Values(i)
            If col3 Then mystr &= " " & Values3.Values(i)
            If col4 Then mystr &= " " & Values4.Values(i)
            If col5 Then mystr &= " " & Values5.Values(i)
            If col6 Then mystr &= " " & Values6.Values(i)

            mystr &= " <"

        Next
    End Sub

End Class

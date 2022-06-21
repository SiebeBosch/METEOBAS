Imports System.IO
Imports System.Windows
Imports METEOBAS.General
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports GemBox.Spreadsheet
Imports System.Data.OleDb
Imports Microsoft.VisualBasic.DateAndTime
Imports CsvHelper.Configuration
Imports CsvHelper


Public Class GeneralFunctions
    Private setup As clsSetup
    Public unitConversion As clsUnitConversion

    Public Enum enmResultsType
        Timestep = 0
        Average = 1
        Minimum = 2
        Maximum = 3
        Percentile = 4
    End Enum

    Public Enum enmRasterType
        ASC = 0
        TIF = 1
        IMG = 2
    End Enum

    Public Enum enmBoundaryType
        H = 0
        Q = 1
    End Enum

    Public Enum enmCulvertController
        NONE = 0
        INLET = 1
        OUTLET = 2
        CALAMITY = 3
    End Enum
    Public Enum enmCulvertHandling
        REACHOBJECT = 0     'object op tak
        STRUCTURE_REACH = 1 'kunstwerkvak met profielen aan weerszijden
    End Enum

    Public Enum enmOrificeUsage
        OPEN = 0
        CLOSED = 1
        INLET = 2
        OUTLET = 3
        CALAMITY_OUTLET = 4
    End Enum

    Public Enum enmSiphonType
        SIPHON = 0
        INVERTED_SIPHON = 1
    End Enum

    Public Enum enmFlowDirection
        BOTH = 0
        POSITIVE = 1
        NEGATIVE = 2
        NONE = 3
    End Enum

    Public Enum enmPumpDirection
        PositiveUpstreamControl = 1
        PositiveDownstreamControl = 2
        PositiveBoth = 3
        NegativeUpstreamControl = -1
        NegativeDownstreamControl = -2
        NegativeBoth = -3
    End Enum

    Public Enum enmWeirCalamityControllerCategory
        NONE = 0
        DROP = 1
        RAISE = 2
    End Enum

    Public Enum enmControllerCategory
        NONE = 0
        TIME = 1
        HYDRAULIC = 2
        INTERVAL = 3
        PID = 4
    End Enum

    Public Enum enmModelResultsType
        H = 0
        Q = 1
        V = 2
        D = 3
    End Enum

    Public Enum enmStatisticalMethod
        Gumbel = 0
        Weibull = 1
        Gen_Pareto = 2
    End Enum

    Public Enum enmTimeSeriesProcessing
        none = 0
        cumulative = 1
        monthlycumulative = 2
        annualcumulative = 3
        monthlysum = 4
        percentile = 5
    End Enum

    Public Enum enmPumpType
        outlet = 0
        inlet = 1
    End Enum

    Public Enum enmControllerType
        level = 0
        time = 1
        'later add hydraulic, interval and pid
    End Enum

    Public Enum enmResampleMethod
        CellCenter = 0
        Bilinear = 1
    End Enum

    Public Enum enmTidalComponent
        VerhoogdLaagwater = 0
        VerhoogdeMiddenstand = 1
        VerhoogdHoogwater = 2
    End Enum

    Public Enum enmDataSource
        NONE = 0
        METEOBASE = 1
        HIRLAM = 2
        WALRUS = 3
        SOBEKRR = 4
        SOBEKCF = 5
    End Enum

    Public Enum enmDataType
        DATUM = 0             'if the date needs to be written in a textfile
        Q_OBSERVED = 1
        Q_SIMULATED = 2
        Q_BASEFLOW = 3
        Q_INTERFLOW = 4
        Q_INLET = 5
        MAKKINK = 6
        PENMAN = 7
        SEEPAGE = 8
        EFFLUENT = 9
        CSO = 10
        ZEROES = 11              'if a textfile column needs to be written, but the content is in fact irrelevant
        PRECIPITATION = 12
    End Enum

    Public Enum enmTimeStepSize
        SECOND = 0
        MINUTE = 1
        HOUR = 2
        DAY = 3
        MONTH = 4
        YEAR = 5
    End Enum

    Public Enum enmDataUnit
        NONE = 0
        MMperDAY = 1
        MMperHOUR = 2
        M3perMIN = 3
        M3perSEC = 4
    End Enum

    Public Enum enmScriptType
        IMPORTHIRLAM = 0
        WRITETIMESERIESTXT = 1
        IMPORTTIMESERIESTXT = 2
    End Enum

    Public Enum enmDSSDates
        STARTDATE = 0
        NOWDATE = 1
        ENDDATE = 2
    End Enum

    Public Enum enmSimulationModel
        SOBEK = 0
        WALRUS = 1
    End Enum

    Public Enum enmKlimaatScenario
        HUIDIG = 0
        KL2030 = 1 'KNMI '14 scenario 2030
        GL2050 = 2 'KNMI '14 scenario 2050 GL
        GH2050 = 3 'KNMI '14 scenario 2050 GH
        WL2050 = 4 'KNMI '14 scenario 2050 WL
        WH2050 = 5 'KNMI '14 scenario 2050 WH
        GL2085 = 6 'KNMI '14 scenario 2085 GL
        GH2085 = 7 'KNMI '14 scenario 2085 GH
        WL2085 = 8 'KNMI '14 scenario 2085 WL
        WH2085 = 9 'KNMI '14 scenario 2085 WH
    End Enum

    Public Enum enmNeerslagPatroon
        ONGECLASSIFICEERD = 0
        HOOG = 1
        MIDDELHOOG = 2
        MIDDELLAAG = 3
        LAAG = 4
        KORT = 5
        LANG = 6
        UNIFORM = 7
    End Enum

    Public Enum enmMeteoStationType
        precipitation = 0
        evaporation = 1
    End Enum

    Public Enum enmSeason
        yearround = 0
        meteowinterhalfyear = 1 'meteorological winterhalfyear: 1-9 to 1-3
        meteosummerhalfyear = 2 'meteorological summerhalfyear: 1-3 to 1-9
        meteosummerquarter = 3  'metorological summer: 1-6 to 1-9
        meteoautumnquarter = 4  'meteorological autumn: 1-9 to 1-12
        meteowinterquarter = 5  'meteorological winter: 1-12 to 1-3
        meteospringquarter = 6  'meteorological spring: 1-3 to 1-6
        hydrosummerhalfyear = 7 'hydrological summer: 15-4 to 15-10
        hydrowinterhalfyear = 8 'hydrological winter: 15-10 to 15-4
        marchthroughoctober = 9 'march through october
        novemberthroughfebruary = 10 'november through february
    End Enum


    Public Enum enmValueStatistic
        mean = 0
        min = 1
        max = 2
    End Enum

    Public Enum enmGridFormat
        AAIGrid = 0 'ASCI grid
        GRIB2 = 1   'GRIB2 as used in HIRLAM
        NetCDF = 2  'NetCDF rasters
        HFA = 3     'ERDAS IMG rasters
        GTiff = 4   'GeoTiff
    End Enum

    Public Enum enmApplicationArea
        GIS = 0
        GISminSBK = 1
        GISminSBKPaved = 2
        GISminSBkCSOPaved = 3
    End Enum

    Public Enum enmKNMI14Scenario
        GL2050 = 0
        GH2050 = 1
        WL2050 = 2
        WH2050 = 3
        GL2085 = 4
        GH2085 = 5
        WL2085 = 6
        WH2085 = 7
    End Enum

    Public Enum enmRainfallRunoffModel
        SOBEKRR = 0
        SACRAMENTO = 1
        HBV = 2
        WAGENINGENMODEL = 3
    End Enum

    Public Enum enmSobekStructureType
        RiverWeir = 0
        AdvancedWeir = 1
        GeneralStructure = 2
        RiverPump = 3
        DatabaseStructure = 4
        Weir = 6
        Orifice = 7
        Pump = 9
        Culvert = 10
        UniversalWeir = 11
        Bridge = 12
        BreachGrowth1DDamBreakNode = 13
        BreachGrowth2DDamBreakNode = 112
    End Enum

    Public Enum enmReachtype
        ReachCFChannel = 1
        ReachCFChannelWithLateral = 2
        ReachOFDamBreak = 3
        ReachSFPipe = 4
        ReachSFPipeWithRunoff = 5
        ReachSFDWAPipeWithRunoff = 6
        ReachSFRWAPipeWithRunoff = 7
        ReachSFPipeAndComb = 8
        ReachSFPipeAndMeas = 9
        ReachSFInternalWeir = 10
        ReachSFInternalOrifice = 11
        ReachSFInternalCulvert = 12
        ReachSFInternalPump = 13
        ReachOFLineBoundary = 14
        ReachOFLine1D2DBoundary = 15

    End Enum

    Public Enum enmNodetype
        NodeSFManhole = 1
        NodeSFManholeWithMeasurement = 2
        NodeSFManholeWithRunoff = 3
        NodeSFManholeWithLateralFlow = 4
        NodeSFManholeWithDischargeAndRunoff = 5
        NodeSFExternalPump = 11
        NodeCFConnectionNode = 12
        ConnNodeLatStor = 13
        NodeCFBoundary = 14
        NodeCFLinkage = 15
        NodeCFGridpoint = 16
        NodeCFGridpointFixed = 17
        MeasurementStation = 18
        NodeCFLateral = 19
        SBK_PROFILE = 20
        NodeCFWeir = 21
        NodeCFUniWeir = 22
        NodeCFOrifice = 23
        NodeCFCulvert = 24
        NodeCFBridge = 26
        NodeCFPump = 27
        NodeCFExtraResistance = 65
        NodeRRCFConnection = 34
        NodeRRCFConnectionNode = 35
        NodeRRPaved = 42
        NodeRRUnpaved = 43
        NodeRRGreenhouse = 44
        NodeRROpenWater = 45
        NodeRRBoundary = 46
        NodeRRPump = 47
        NodeRRIndustry = 48
        NodeRRSacramento = 54
        NodeRRWWTP = 56
        NodeRRWeir = 49
        NodeRROrifice = 50
        NodeRRFriction = 51
        NodeRRWageningenModel = 69

        'virtual node types for internal use
        Virtual = 777
        HBoundary = 888
        QBoundary = 999

    End Enum


    Public Enum enmProfileType
        tabulated = 0
        trapezium = 1
        opencircle = 2
        sedredge = 3
        closedcircle = 4
        eggshape = 6
        eggshape2 = 7
        closedrectangular = 8
        yztable = 10
        asymmetricaltrapezium = 11
    End Enum

    Public Enum enmMathOption
        Sum = 0
        Avg = 1
        Min = 2
        Max = 3
    End Enum

    Public Enum enmFrictionType
        Chezy = 0
        Manning = 1
        StricklerKN = 2
        StricklerKS = 3
        WhiteColebrook = 4
        BosBijkerk = 7
        GlobalFriction = 99
    End Enum


    '=======================================================
    'Service provided by Telerik (www.telerik.com)
    'Conversion powered by Refactoring Essentials.
    'Twitter: @telerik
    'Facebook: facebook.com/telerik
    '=======================================================


    Friend Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
        unitConversion = New clsUnitConversion
    End Sub

    Public Function SeasonFromString(myName As String) As enmSeason
        Select Case myName.ToLower
            Case Is = "zomer", "summer", "zomerhalfjaar", "summerhalfyear"
                Return enmSeason.hydrosummerhalfyear
            Case Is = "winter", "winterhalfjaar", "winterhalfyear"
                Return enmSeason.hydrowinterhalfyear
            Case Is = "meteowinter"
                Return enmSeason.meteowinterhalfyear
            Case Is = "meeteosummer", "meteozomer"
                Return enmSeason.meteosummerhalfyear
            Case Else
                Return Nothing
        End Select

    End Function

    Public Function ConsonantsInARow(myStr As String) As Integer
        'counts the largest row of consonants in succession
        Dim maxNum As Integer = 0
        Dim curNum As Integer = 0
        Dim myChar As String
        For i = 1 To myStr.Length
            myChar = Mid(myStr, i, 1)
            If IsConsonant(myChar) Then
                curNum += 1
                If curNum > maxNum Then maxNum = curNum
            Else
                'reset the current count
                curNum = 0
            End If
        Next
        Return maxNum
    End Function

    Public Function IsConsonant(myChar As String) As Boolean
        Select Case myChar.Trim.ToLower
            Case Is = "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "z"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Sub DeleteFilesInDir(directoryName As String, filter As String, RemovalOption As System.IO.SearchOption)
        For Each deleteFile In Directory.GetFiles(directoryName, filter, RemovalOption)
            File.Delete(deleteFile)
        Next
    End Sub

    Public Function VolumeToReturnPeriod(Volume As Double, DurationHours As Integer) As Double
        'this function estimates the return period for a given precipitation volume and duration
        'it uses the stochastic table for the current climate from meteobase in order to do so
        Dim Herh As New Dictionary(Of Integer, Dictionary(Of Integer, Double))
        Dim Vol As Dictionary(Of Integer, Double)

        '5 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 0.5)
        Vol.Add(4, 0.5)
        Vol.Add(8, 0.5)
        Vol.Add(12, 0.5)
        Vol.Add(24, 0.5)
        Vol.Add(48, 0.5)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(5, Vol)

        '10 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 0.5)
        Vol.Add(4, 0.5)
        Vol.Add(8, 0.5)
        Vol.Add(12, 0.5)
        Vol.Add(24, 0.5)
        Vol.Add(48, 0.5)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(5, Vol)

        '20 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1.2)
        Vol.Add(4, 0.7)
        Vol.Add(8, 0.5)
        Vol.Add(12, 0.5)
        Vol.Add(24, 0.5)
        Vol.Add(48, 0.5)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(10, Vol)

        '30 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 4.9)
        Vol.Add(4, 3.0)
        Vol.Add(8, 1.6)
        Vol.Add(12, 1.1)
        Vol.Add(24, 0.5)
        Vol.Add(48, 0.5)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(30, Vol)

        '40 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 17.5)
        Vol.Add(4, 10.4)
        Vol.Add(8, 5.5)
        Vol.Add(12, 3.6)
        Vol.Add(24, 1.5)
        Vol.Add(48, 0.6)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(40, Vol)


        '50 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 54.9)
        Vol.Add(4, 32.2)
        Vol.Add(8, 16.9)
        Vol.Add(12, 10.6)
        Vol.Add(24, 4.4)
        Vol.Add(48, 1.6)
        Vol.Add(96, 0.5)
        Vol.Add(192, 0.5)
        Vol.Add(216, 0.5)
        Herh.Add(50, Vol)


        '70 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 400.2)
        Vol.Add(4, 232.4)
        Vol.Add(8, 117.5)
        Vol.Add(12, 73.2)
        Vol.Add(24, 28.3)
        Vol.Add(48, 9.1)
        Vol.Add(96, 2.4)
        Vol.Add(192, 0.6)
        Vol.Add(216, 0.5)
        Herh.Add(70, Vol)


        '90 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1000)
        Vol.Add(4, 1000)
        Vol.Add(8, 628.3)
        Vol.Add(12, 388.3)
        Vol.Add(24, 146.7)
        Vol.Add(48, 44.5)
        Vol.Add(96, 10.5)
        Vol.Add(192, 2.0)
        Vol.Add(216, 1.5)
        Herh.Add(90, Vol)


        '110 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1000)
        Vol.Add(4, 1000)
        Vol.Add(8, 1000)
        Vol.Add(12, 1000)
        Vol.Add(24, 637)
        Vol.Add(48, 192.2)
        Vol.Add(96, 43.3)
        Vol.Add(192, 6.8)
        Vol.Add(216, 4.9)
        Herh.Add(110, Vol)

        '130 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1000)
        Vol.Add(4, 1000)
        Vol.Add(8, 1000)
        Vol.Add(12, 1000)
        Vol.Add(24, 1000)
        Vol.Add(48, 742.7)
        Vol.Add(96, 171.2)
        Vol.Add(192, 24.9)
        Vol.Add(216, 17.0)
        Herh.Add(130, Vol)

        '150 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1000)
        Vol.Add(4, 1000)
        Vol.Add(8, 1000)
        Vol.Add(12, 1000)
        Vol.Add(24, 1000)
        Vol.Add(48, 1000)
        Vol.Add(96, 650.3)
        Vol.Add(192, 98.3)
        Vol.Add(216, 65.8)
        Herh.Add(150, Vol)

        '170 mm
        Vol = New Dictionary(Of Integer, Double)  'key is duration in hours
        Vol.Add(2, 1000)
        Vol.Add(4, 1000)
        Vol.Add(8, 1000)
        Vol.Add(12, 1000)
        Vol.Add(24, 1000)
        Vol.Add(48, 1000)
        Vol.Add(96, 1000)
        Vol.Add(192, 426)
        Vol.Add(216, 290.2)
        Herh.Add(170, Vol)

        Return 100

    End Function

    Public Function AngleDifferenceDegrees(firstangle As Double, secondangle As Double) As Double
        Dim difference As Double = secondangle - firstangle
        Select Case difference
            Case Is < -180
                difference += 360
            Case Is > 180
                difference -= 360
        End Select
        If secondangle = firstangle Then
            Return 0
        Else
            Return (Math.Abs(difference))
        End If
    End Function

    Public Function ReplaceInvalidCharactersInPath(myPath As String, Optional ByVal ReplaceString As String = "_") As String
        While InStr(myPath, "/") > 0
            Me.setup.Log.AddWarning("File path " & myPath & " was slighly adjusted to prevent invalid characters in its path.")
            myPath = Replace(myPath, "/", ReplaceString)
        End While

        While InStr(3, myPath, ":") > 2
            Me.setup.Log.AddWarning("File path " & myPath & " was slightly adjusted to prevent invalid characters in its path.")
            myPath = Left(myPath, 2) & Replace(myPath, ":", ReplaceString, 3)
        End While

        Return myPath
    End Function


    Public Sub AddShapeFileToFileCollection(ByRef myCollection As Dictionary(Of String, String), ShpPath As String)
        Dim shxPath As String = Replace(ShpPath, ".shp", ".shx",,, CompareMethod.Text)
        Dim dbfPath As String = Replace(ShpPath, ".shp", ".dbf",,, CompareMethod.Text)
        If Not myCollection.ContainsKey(ShpPath.Trim.ToUpper) Then myCollection.Add(ShpPath.Trim.ToUpper, ShpPath)
        If Not myCollection.ContainsKey(shxPath.Trim.ToUpper) Then myCollection.Add(shxPath.Trim.ToUpper, shxPath)
        If Not myCollection.ContainsKey(dbfPath.Trim.ToUpper) Then myCollection.Add(dbfPath.Trim.ToUpper, dbfPath)
    End Sub


    Public Function datatableContainsString(ByRef dt As DataTable, Value As String, ColIdx As Integer) As Boolean
        Dim i As Long
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i)(ColIdx) = Value Then Return True
        Next
        Return False
    End Function

    Public Function dataTableCountStrings(ByRef dt As DataTable, Value As String, ColIdx As Integer) As Long

        Dim rows As DataRow()
        rows = dt.Select(dt.Columns(ColIdx).ColumnName & " = '" & Value & "'")
        Return rows.Length

        'Dim i As Long, n As Long = 0
        'For i = 0 To dt.Rows.Count - 1
        '    If dt.Rows(i)(ColIdx) = Value Then n += 1
        'Next
        'Return n
    End Function

    Public Function GetAddDataTableDoubleIdx(ByRef dt As DataTable, myVal As Double, colIdx As Integer) As Long
        '========================================================================================================
        ' this function searches a datatable for an existing record with a given value of type Double
        ' it will return the index number for the row where that date was found
        ' if not found, -1 is returned
        '========================================================================================================
        Dim i As Long

        'look for existing row
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i)(colIdx) = myVal Then
                Return i
            End If
        Next

        'not found, so add it
        Dim dtRow As DataRow
        dtRow = dt.NewRow()
        dt.Rows.Add(dtRow)
        dt.Rows(dt.Rows.Count - 1)(colIdx) = myVal
        Return dt.Rows.Count - 1

    End Function


    Public Function GetAddDataTableDateIdx(ByRef dt As DataTable, myDate As Date, colIdx As Integer) As Long
        '========================================================================================================
        ' this function searches a datatable for an existing record with a given date
        ' it will return the index number for the row where that date was found
        ' if not found, -1 is returned
        '========================================================================================================
        Dim i As Long

        'look for existing row
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i)(colIdx) = myDate Then
                Return i
            End If
        Next

        'not found, so add it
        Dim dtRow As DataRow
        dtRow = dt.NewRow()
        dt.Rows.Add(dtRow)
        dt.Rows(dt.Rows.Count - 1)(colIdx) = myDate
        Return dt.Rows.Count - 1

    End Function


    'Public Function DataTableStats(ByRef myTable As DataTable, query As String, ByRef myFirst As Double, ByRef myLast As Double, PercentileIncrement As Integer) As Dictionary(Of Integer, Double)
    '    Dim myStats As New Dictionary(Of Integer, Double)
    '    Dim rowIdx As Long, i As Long


    '    Dim foundRows() As DataRow
    '    foundRows = myTable.Select(query)

    '    'first retrieve the first and last values in the table
    '    myFirst = foundRows.GetValue(0)(1)
    '    myLast = foundRows(foundRows.Count - 1)(1)



    '    'Dim newList As New List(Of Double)
    '    'For i = 0 To myTable.Rows.Count - 1
    '    '    newList.Add(myTable.Rows(i)(ColIdx))
    '    'Next

    '    'For i = 0 To 100 Step PercentileIncrement
    '    '    rowIdx = i / 100 * (newList.Count - 1)
    '    '    If rowIdx < 0 Then rowIdx = 0
    '    '    If rowIdx > newList.Count - 1 Then rowIdx = newList.Count - 1
    '    '    myStats.Add(i, newList(rowIdx))
    '    'Next

    '    'create a derivative datatable that is sorted by the values in the column of interest
    '    'Dim ColName As String = myTable.Columns.Item(ColIdx).ColumnName
    '    'Dim dataView As New DataView(myTable)
    '    'dataView.Sort = ColName & " ASC"
    '    'Dim dataTable As DataTable = dataView.ToTable()

    '    'For i = 0 To 100 Step PercentileIncrement
    '    '    rowIdx = i / 100 * (dataTable.Rows.Count - 1)
    '    '    If rowIdx < 0 Then rowIdx = 0
    '    '    If rowIdx > dataTable.Rows.Count - 1 Then rowIdx = dataTable.Rows.Count - 1
    '    '    myStats.Add(i, dataTable.Rows(rowIdx)(ColIdx))
    '    'Next
    '    Return myStats
    'End Function

    Public Function DataTableStats(ByRef myTable As DataTable, StartIdx As Integer, EndIdx As Integer, ColIdx As Integer, ByRef myFirst As Double, ByRef myLast As Double, ByVal getMin As Boolean, ByRef myMin As Double, ByVal getMax As Double, ByRef myMax As Double, ByVal calcPercentiles As Boolean, ByVal PercentileIncrement As Integer) As Dictionary(Of Integer, Double)
        Dim myStats As New Dictionary(Of Integer, Double)
        Dim rowIdx As Long, i As Long

        StartIdx = Math.Max(0, StartIdx)
        EndIdx = Math.Min(EndIdx, myTable.Rows.Count - 1)

        'first retrieve the first and last values in the table
        myFirst = myTable.Rows(StartIdx)(ColIdx)
        myLast = myTable.Rows(EndIdx)(ColIdx)

        If getMin OrElse getMax OrElse calcPercentiles Then
            Dim newList As New List(Of Double)
            For i = StartIdx To EndIdx
                newList.Add(myTable.Rows(i)(ColIdx))
            Next

            newList.Sort()

            If getMin Then myMin = newList(0)
            If getMax Then myMax = newList(newList.Count - 1)

            If calcPercentiles Then
                If PercentileIncrement > 0 Then
                    For i = 0 To 100 Step PercentileIncrement
                        rowIdx = i / 100 * (newList.Count - 1)
                        If rowIdx < 0 Then rowIdx = 0
                        If rowIdx > newList.Count - 1 Then rowIdx = newList.Count - 1
                        myStats.Add(i, newList(rowIdx))
                    Next
                End If
            End If
        End If

        Return myStats
    End Function

    Public Sub DataTableMinMax(ByRef myTable As DataTable, Columns As List(Of Integer), ByRef minY As Double, ByRef maxY As Double)
        Dim i As Long, j As Integer
        minY = 9000000000.0
        maxY = -9000000000.0
        For i = 0 To myTable.Rows.Count - 1
            For j = 0 To Columns.Count - 1
                If myTable.Rows(i)(Columns(j)) < minY Then minY = myTable.Rows(i)(Columns(j))
                If myTable.Rows(i)(Columns(j)) > maxY Then maxY = myTable.Rows(i)(Columns(j))
            Next
        Next
    End Sub

    Public Function DataTableMax(ByRef myTable As DataTable, Columns As List(Of Integer)) As Double
        Dim i As Long, j As Integer
        Dim maxY As Double = -9000000000.0
        For i = 0 To myTable.Rows.Count - 1
            For j = 0 To Columns.Count - 1
                If myTable.Rows(i)(Columns(j)) > maxY Then maxY = myTable.Rows(i)(Columns(j))
            Next
        Next
        Return maxY
    End Function

    Public Function DataTableMin(ByRef myTable As DataTable, Columns As List(Of Integer)) As Double
        Dim i As Long, j As Integer
        Dim minY As Double = 9000000000.0
        For i = 0 To myTable.Rows.Count - 1
            For j = 0 To Columns.Count - 1
                If myTable.Rows(i)(Columns(j)) < minY Then minY = myTable.Rows(i)(Columns(j))
            Next
        Next
        Return minY
    End Function

    Public Function DataTableAvg(ByRef myTable As DataTable, Columns As List(Of Integer)) As Double
        Try
            Dim i As Long, j As Integer
            Dim n As Long, myValue As Double
            For i = 0 To myTable.Rows.Count - 1
                For j = 0 To Columns.Count - 1
                    n += 1
                    myValue += myTable.Rows(i)(Columns(j))
                Next
            Next
            If n > 0 Then
                Return myValue / n
            Else
                Throw New Exception("Error retrieving average value from datatable. Value of zero was returned.")
            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return 0
        End Try
    End Function

    Public Function DataTablePercentile(ByRef myTable As DataTable, Columns As List(Of Integer), Percentile As Double) As Double
        Try
            Dim i As Long, j As Integer
            Dim newTable As New List(Of Double)
            For i = 0 To myTable.Rows.Count - 1
                For j = 0 To Columns.Count - 1
                    newTable.Add(myTable.Rows(i)(Columns(j)))
                Next
            Next
            Return setup.GeneralFunctions.PercentileFromList(newTable, Percentile)
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return 0
        End Try
    End Function


    Public Function DataTableMinMaxAsList(myTable As DataTable, Columns As List(Of Integer), ByRef minList As List(Of Double), ByRef maxList As List(Of Double), OrderAscending As Boolean, Optional ByVal AcceptDoubleValues As Boolean = True) As Boolean
        'this function retrieves the minimum and maximum values from each specified column of a datatable
        'and sorts them in ascending order
        Try
            Dim r As Integer, c As Integer
            Dim myMin As Double, myMax As Double
            For c = 0 To Columns.Count - 1
                myMin = 9.0E+99
                myMax = -9.0E+99
                For r = 0 To myTable.Rows.Count - 1
                    If myTable.Rows(r)(Columns(c)) < myMin Then myMin = myTable.Rows(r)(Columns(c))
                    If myTable.Rows(r)(Columns(c)) > myMax Then myMax = myTable.Rows(r)(Columns(c))
                Next
                minList.Add(myMin)
                maxList.Add(myMax)
            Next

            If OrderAscending Then
                minList.Sort()
                maxList.Sort()
                'Else
                '    minList.Reverse()
                '    maxList.Reverse()
            End If
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function


    Public Sub DataTable2CSV(ByRef dtRes As DataTable, ByRef dtRuns As DataTable, path As String, delimiter As String)
        Dim i As Integer, j As Integer, k As Integer, myStr As String, RUNID As String

        path = ReplaceInvalidCharactersInPath(path, "_")
        Using csvWriter As New StreamWriter(path)

            'write the csv header
            myStr = ""
            For j = 0 To dtRes.Columns.Count - 1
                myStr &= dtRes.Columns(j).ColumnName & delimiter
            Next
            For j = 0 To dtRuns.Columns.Count - 2
                myStr &= dtRuns.Columns(j).ColumnName & delimiter
            Next
            myStr &= dtRuns.Columns(dtRuns.Columns.Count - 1).ColumnName & delimiter
            csvWriter.WriteLine(myStr)

            'write the csv data
            For i = 0 To dtRes.Rows.Count - 1
                'first the results
                myStr = ""
                For j = 0 To dtRes.Columns.Count - 1
                    myStr &= dtRes.Rows(i)(j) & delimiter
                Next

                'then the simulation information
                RUNID = dtRes.Rows(i)("RUNID")
                For j = 0 To dtRuns.Rows.Count - 1
                    If dtRuns.Rows(j)("RUNID") = RUNID Then
                        For k = 0 To dtRuns.Columns.Count - 2
                            myStr &= dtRuns.Rows(j)(k) & delimiter
                        Next
                        myStr &= dtRuns.Rows(j)(dtRuns.Columns.Count - 1)
                        Exit For
                    End If
                Next

                csvWriter.WriteLine(myStr)
            Next
        End Using
    End Sub


    Public Function AdjustDateForTimestep(ByVal myDate As DateTime, ByVal TimeStepSize As enmTimeStepSize) As DateTime
        'in some cases we're looking for daily values in hourly data. This means we'll have to match
        'e.g. 11/12/2012 18:00 with 11/12/2012

        'adjust the searched date in case of other timesteps e.g. day
        Select Case TimeStepSize
            Case Is = enmTimeStepSize.YEAR
                Return New DateTime(Year(myDate), 1, 1)
            Case Is = enmTimeStepSize.MONTH
                Return New DateTime(Year(myDate), Month(myDate), 1)
            Case Is = enmTimeStepSize.DAY
                Return New DateTime(Year(myDate), Month(myDate), Day(myDate))
            Case Is = enmTimeStepSize.HOUR
                Return New DateTime(Year(myDate), Month(myDate), Day(myDate), Hour(myDate), 0, 0)
            Case Is = enmTimeStepSize.MINUTE
                Return New DateTime(Year(myDate), Month(myDate), Day(myDate), Hour(myDate), Minute(myDate), 0)
            Case Is = enmTimeStepSize.SECOND
                Return myDate
            Case Else
                Return myDate
        End Select

    End Function

    Public Function MultiplierFromTimeStepConversion(ByVal FromTimestep As enmTimeStepSize, ByVal ToTimestep As enmTimeStepSize) As Double
        'this function returns a multiplier that converts one timestep to another
        If FromTimestep.ToString = ToTimestep.ToString Then Return 1

        Select Case FromTimestep
            Case Is = enmTimeStepSize.SECOND
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.MINUTE
                        Return 1 / 60
                    Case Is = enmTimeStepSize.HOUR
                        Return 1 / 3600
                    Case Is = enmTimeStepSize.DAY
                        Return 1 / (24 * 3600)
                    Case Is = enmTimeStepSize.MONTH
                        Return 1 / (30 * 24 * 3600)
                    Case Is = enmTimeStepSize.YEAR
                        Return 1 / (365.25 * 24 * 3600)
                End Select
            Case Is = enmTimeStepSize.MINUTE
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.SECOND
                        Return 60
                    Case Is = enmTimeStepSize.HOUR
                        Return 1 / 60
                    Case Is = enmTimeStepSize.DAY
                        Return 1 / (24 * 60)
                    Case Is = enmTimeStepSize.MONTH
                        Return 1 / (30 * 24 * 60)
                    Case Is = enmTimeStepSize.YEAR
                        Return 1 / (365.25 * 24 * 60)
                End Select
            Case Is = enmTimeStepSize.HOUR
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.SECOND
                        Return 3600
                    Case Is = enmTimeStepSize.MINUTE
                        Return 60
                    Case Is = enmTimeStepSize.DAY
                        Return 1 / 24
                    Case Is = enmTimeStepSize.MONTH
                        Return 1 / (30 * 24)
                    Case Is = enmTimeStepSize.YEAR
                        Return 1 / (365.25 * 24)
                End Select
            Case Is = enmTimeStepSize.DAY
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.SECOND
                        Return 3600 * 24
                    Case Is = enmTimeStepSize.MINUTE
                        Return 60 * 24
                    Case Is = enmTimeStepSize.HOUR
                        Return 24
                    Case Is = enmTimeStepSize.MONTH
                        Return 1 / 30
                    Case Is = enmTimeStepSize.YEAR
                        Return 1 / 365.25
                End Select
            Case Is = enmTimeStepSize.MONTH
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.SECOND
                        Return 3600 * 24 * 30
                    Case Is = enmTimeStepSize.MINUTE
                        Return 60 * 24 * 30
                    Case Is = enmTimeStepSize.HOUR
                        Return 24 * 30
                    Case Is = enmTimeStepSize.DAY
                        Return 30
                    Case Is = enmTimeStepSize.YEAR
                        Return 1 / 12
                End Select
            Case Is = enmTimeStepSize.YEAR
                Select Case ToTimestep
                    Case Is = enmTimeStepSize.SECOND
                        Return 3600 * 24 * 365.25
                    Case Is = enmTimeStepSize.MINUTE
                        Return 60 * 24 * 365.25
                    Case Is = enmTimeStepSize.HOUR
                        Return 24 * 365.25
                    Case Is = enmTimeStepSize.DAY
                        Return 365.25
                    Case Is = enmTimeStepSize.MONTH
                        Return 12
                End Select
        End Select
    End Function

    Public Function DateFormattingByTimeStep(ByVal TimeStepStr As String) As String
        'determine in which format to write to the database based on the timestep size
        'of a time series. e.g. timestep of 1 hour can be written to database as yyyy/MM/dd HH
        'timestep of 1 day can simply be written to database as yyyy/MM/dd
        Select Case TimeStepStr.Trim.ToLower
            Case Is = "hour"
                Return "yyyy/MM/dd HH"
            Case Is = "day"
                Return "yyyy/MM/dd"
            Case Is = "month"
                Return "yyyy/MM"
            Case Is = "year"
                Return "yyyy"
        End Select
        Return ""
    End Function

    Public Function getMeteoStationTypeFromString(ByVal myString As String) As enmMeteoStationType
        Select Case myString.Trim.ToUpper
            Case Is = "RAINFALL"
                Return enmMeteoStationType.precipitation
            Case Is = "NEERSLAG"
                Return enmMeteoStationType.precipitation
            Case Is = "REGEN"
                Return enmMeteoStationType.precipitation
            Case Is = "PRECIPITATION"
                Return enmMeteoStationType.precipitation
            Case Is = "EVAPORATION"
                Return enmMeteoStationType.evaporation
            Case Is = "EVAPOTRANSPIRATION"
                Return enmMeteoStationType.evaporation
            Case Is = "VERDAMPING"
                Return enmMeteoStationType.evaporation
            Case Is = "EVAPOTRANSPIRATIE"
                Return enmMeteoStationType.evaporation
        End Select
    End Function

    Public Function RemoveNumericPostfixes(ByVal NameBase As String, ByVal Delimiter As String) As String
        'this function removes multiple postfixes (preceeded by a delimiter) from a string
        'eg: REACH12_45_1 yields REACH12
        Dim Done As Boolean, PrevNameBase As String

        While Not Done
            PrevNameBase = NameBase
            NameBase = RemoveNumericPostFix(NameBase, Delimiter)
            If NameBase = PrevNameBase Then Return NameBase
        End While

        Return NameBase

    End Function

    Public Function RemoveNumericPostFix(ByVal NameBase As String, ByVal Delimiter As String) As String
        'this function removes a numeric postfix (preceeded by a delimiter) from a string
        'eg: REACH12_45 with delimiter "_" yields REACH12
        Dim i As Long
        For i = NameBase.Length To 1 Step -1
            If Mid(NameBase, i, 1) = Delimiter Then
                Return Left(NameBase, i - 1)
            ElseIf Not IsNumeric(Mid(NameBase, i, 1)) Then
                Exit For
            End If
        Next
        Return NameBase
    End Function

    Public Function AddToDate(ByVal RefDate As Date, ByVal AddNumber As Integer, ByVal AddUnits As String) As DateTime
        'set the absolute value for the start date
        Select Case AddUnits.Trim.ToLower
            Case Is = "uren", "hours"
                Return RefDate.AddHours(AddNumber)
            Case Is = "dagen", "days"
                Return RefDate.AddDays(AddNumber)
            Case Is = "weken", "weeks"
                Return RefDate.AddDays(7 * AddNumber)
            Case Is = "maanden", "months"
                Return RefDate.AddMonths(AddNumber)
            Case Is = "jaren", "years"
                Return RefDate.AddYears(AddNumber)
        End Select

        Return RefDate

    End Function


    Public Function LengthByTimeUnitConversion(ByVal myValue As Double, ByVal FromUnit As enmDataUnit, ByVal ToUnit As enmDataUnit) As Double
        If FromUnit = ToUnit Then Return myValue
        Select Case FromUnit
            Case Is = enmDataUnit.MMperDAY
                Select Case ToUnit
                    Case Is = enmDataUnit.MMperHOUR
                        Return myValue / 24
                End Select
            Case Is = enmDataUnit.MMperHOUR
                Select Case ToUnit
                    Case Is = enmDataUnit.MMperDAY
                        Return myValue * 24
                End Select
        End Select
    End Function

    Public Function BurnDateInStringByTemplate(ByRef myString As String, ByVal DateTemplate As String, ByVal mydate As DateTime) As Boolean
        'this function 'burns' a date inside an existing string, based on a given template for the formatting
        Try
            Dim DateString As String = DateTemplate

            If InStr(DateTemplate, "yyyy", CompareMethod.Binary) > 0 Then
                DateString = Replace(DateString, "yyyy", Year(mydate).ToString, , , CompareMethod.Text)
            End If

            If InStr(DateTemplate, "MM", CompareMethod.Binary) > 0 Then
                DateString = Replace(DateString, "MM", Format(Month(mydate), "00"), , , CompareMethod.Binary)
            End If

            If InStr(DateTemplate, "dd", CompareMethod.Text) > 0 Then
                DateString = Replace(DateString, "dd", Format(Day(mydate), "00"), , , CompareMethod.Text)
            End If

            If InStr(DateTemplate, "HH", CompareMethod.Text) > 0 Then
                DateString = Replace(DateString, "HH", Format(Hour(mydate), "00"), , , CompareMethod.Text)
            End If

            If InStr(DateTemplate, "mm", CompareMethod.Binary) > 0 Then
                DateString = Replace(DateString, "mm", Format(Minute(mydate), "00"), , , CompareMethod.Text)
            End If

            myString = Replace(myString, DateTemplate, DateString, , , CompareMethod.Text)

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function replaceFileNameInPath(path As String, NewFileName As String) As String
        Dim firstpos As Integer = 0, lastpos As Integer = 0

        'find the file extension
        While Not InStr(firstpos + 1, path, "\") = 0
            firstpos = InStr(firstpos + 1, path, "\")
        End While

        While Not InStr(lastpos + 1, path, ".") = 0
            lastpos = InStr(lastpos + 1, path, ".")
        End While

        If firstpos > 0 AndAlso lastpos > 0 Then
            Return Left(path, firstpos) & NewFileName & Right(path, path.Length - lastpos + 1)
        Else
            Return ""
        End If

    End Function

    Public Function getExtensionFromFileName(Path As String) As String
        Dim pointpos As Integer = 0
        While Not InStr(pointpos + 1, Path, ".") = 0
            pointpos = InStr(pointpos + 1, Path, ".")
        End While

        Return Right(Path, Path.Length - pointpos)

    End Function



    Public Function BurnIntInStringByTemplate(ByRef myString As String, ByVal ValTemplate As String, ByVal myValue As Long) As Boolean
        'this function 'burns' an integer inside an existing string, based on a given template for the formatting
        Try

            Dim FormattingString As String = ""
            Dim i As Integer
            For i = 1 To Len(ValTemplate)
                FormattingString &= "0"
            Next

            If InStr(myString, ValTemplate, CompareMethod.Text) > 0 Then
                myString = Replace(myString, ValTemplate, Format(myValue, FormattingString))
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function DateFromFormattedString(ByVal myString As String, ByVal DateFormatting As String, ByRef myDate As DateTime) As Boolean
        Dim myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMinute As Integer, mySecond As Integer

        Try
            'string contains a date formatting code. Find it and replace it with the actual date value
            Select Case DateFormatting.Trim.ToLower
                Case Is = "yyyymmddhhmmss"
                    myYear = Left(myString, 4)
                    myMonth = Left(Right(myString, 10), 2)
                    myDay = Left(Right(myString, 8), 2)
                    myHour = Left(Right(myString, 6), 2)
                    myMinute = Left(Right(myString, 4), 2)
                    mySecond = Right(myString, 2)
                Case Is = "yyyymmddhhmm"
                    myYear = Left(myString, 4)
                    myMonth = Left(Right(myString, 8), 2)
                    myDay = Left(Right(myString, 6), 2)
                    myHour = Left(Right(myString, 4), 2)
                    myMinute = Right(myString, 2)
                    myMinute = 0
                Case Is = "yyyymmddhh"
                    myYear = Left(myString, 4)
                    myMonth = Left(Right(myString, 6), 2)
                    myDay = Left(Right(myString, 4), 2)
                    myHour = Right(myString, 2)
                    myMinute = 0
                    mySecond = 0
                Case Is = "yyyymmdd"
                    myYear = Left(myString, 4)
                    myMonth = Left(Right(myString, 4), 2)
                    myDay = Right(myString, 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "yyyy/mm/dd"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = myString
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "yyyy/mm/dd hh"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ": ")
                    myMinute = 0
                    mySecond = 0
                Case Is = "yyyy/mm/dd hh"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ":")
                    myMinute = 0
                    mySecond = 0
                Case Is = "yyyy-mm-dd hh:mm"
                    myYear = ParseString(myString, "-")
                    myMonth = ParseString(myString, "-")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ":")
                    myMinute = myString
                    mySecond = 0
                Case Is = "yyyy/MM/dd HH:mm:ss"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ":")
                    myMinute = ParseString(myString, ":")
                    mySecond = myString
                Case Is = "yyyy/MM/dd hh:mm:ss"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ":")
                    myMinute = ParseString(myString, ":")
                    mySecond = myString
                Case Is = "yyyy-MM-dd hh:mm"
                    myYear = ParseString(myString, "/")
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, " ")
                    myHour = ParseString(myString, ":")
                    myMinute = ParseString(myString, ":")
                    mySecond = myString
                Case Is = "MM/dd/yyyy"
                    myMonth = ParseString(myString, "/")
                    myDay = ParseString(myString, "/")
                    myYear = myString
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Else
                    Throw New Exception("Error retrieving date from string. Probably the date formatting is not supported.")
            End Select

            myDate = New DateTime(myYear, myMonth, myDay, myHour, myMinute, mySecond)
            Return True

        Catch ex As Exception
            Me.setup.Log.AddError("Error converting date string to date/time: " & myString)
            Return False
        End Try


    End Function

    Public Function LongFromFormattedDateTimeString(ByVal myString As String, ByVal DateFormatting As String) As Long
        'this function returns a date/time string as a long
        'e.g. 2003/03/24 12:49:00 in yyyy/MM/dd hh:mm:ss-format becomes 20030224124900

        Dim pos As Integer
        Dim NewStr As String

        Try

            'get the year
            pos = InStr(DateFormatting, "yyyy", CompareMethod.Text)
            If pos > 0 Then
                NewStr = Mid(myString, pos, 4)
            Else
                Throw New Exception("Formatting for year not recognized in " & DateFormatting & ". Only 'yyyy' is supported.")
            End If

            'get the month
            pos = InStr(DateFormatting, "MM", CompareMethod.Binary)
            If pos > 0 Then
                NewStr &= Mid(myString, pos, 2)
            Else
                Throw New Exception("Formatting for year not recognized in " & DateFormatting & ". Only 'MM' is supported.")
            End If

            'get the day
            pos = InStr(DateFormatting, "dd", CompareMethod.Text)
            If pos > 0 Then
                NewStr &= Mid(myString, pos, 2)
            Else
                Throw New Exception("Formatting for year not recognized in " & DateFormatting & ". Only 'dd' is supported.")
            End If

            'get the hour
            pos = InStr(DateFormatting, "hh", CompareMethod.Text)
            If pos > 0 Then
                NewStr &= Mid(myString, pos, 2)
            Else
                'assume 0h
                NewStr &= "00"
            End If

            'get the minute
            pos = InStr(DateFormatting, "mm", CompareMethod.Binary)
            If pos > 0 Then
                NewStr &= Mid(myString, pos, 2)
            Else
                'assume 00
                NewStr &= "00"
            End If

            'get the second
            pos = InStr(DateFormatting, "ss", CompareMethod.Text)
            If pos > 0 Then
                NewStr &= Mid(myString, pos, 2)
            Else
                'assume 00
                NewStr &= "00"
            End If

            'return the outcome
            If IsNumeric(NewStr) Then
                Return Val(NewStr)
            Else
                Throw New Exception("Error converting date string to long: " & myString)
            End If

        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
        End Try

    End Function

    Public Function MapWinGeoProjectionFromString(ByVal ProjectionString As String, ByRef GeoProjection As MapWinGIS.GeoProjection) As Boolean

        ProjectionString = RemoveBoundingQuotes(ProjectionString)
        GeoProjection = New MapWinGIS.GeoProjection
        If InStr(ProjectionString, "EPSG:", CompareMethod.Text) > 0 Then
            ProjectionString = Replace(ProjectionString, "EPSG:", "", , , CompareMethod.Text)
            If Not GeoProjection.ImportFromEPSG(ProjectionString) Then Return False
        ElseIf InStr(ProjectionString, "Proj4", CompareMethod.Text) > 0 Then
            If Not GeoProjection.ImportFromProj4(RemoveBoundingQuotes(ProjectionString)) Then Return False
        Else
            Me.setup.Log.AddError("Error: projection string not recognized or supported: " & ProjectionString)
            Return False
        End If

        Return True

    End Function

    Public Function SetValueInStringByFormat(ByVal myString As String, ByVal ValueFormatting As String, ByVal myValue As Double) As String
        'replaces a given series of strings by a number, formatted according to the number of characters
        Dim ReplaceFormat As String = "", i As Long
        For i = 1 To ValueFormatting.Length
            ReplaceFormat &= "0"
        Next

        If InStr(myString, ValueFormatting) > 0 Then
            Return Replace(myString, ValueFormatting, Format(myValue, ReplaceFormat))
        Else
            Me.setup.Log.AddError("Error in function SetValueByFormatInString: formatting string was not found in " & myString)
            Return myString
        End If

    End Function


    Public Function ListUniqueValuesFromShapefile(ByVal Path As String, ByVal ShapeField As String, ByRef myList As List(Of String)) As Boolean
        Try
            Dim FieldIdx As Long, i As Long
            myList = New List(Of String)
            If Not System.IO.File.Exists(Path) Then Throw New Exception("Shapefile does not exist: " & Path)
            Dim mySf As New clsShapeFile(Me.setup, Path)
            If Not mySf.Open() Then Throw New Exception("Error reading shapefile: " & Path)
            FieldIdx = mySf.GetFieldIdx(ShapeField)
            If FieldIdx < 0 Then Throw New Exception("Field " & ShapeField & " not found in shapefile " & Path)

            For i = 0 To mySf.sf.NumShapes - 1
                If Not myList.Contains(mySf.sf.CellValue(FieldIdx, i)) Then myList.Add(mySf.sf.CellValue(FieldIdx, i))
            Next

            Return True
        Catch ex As Exception

            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Sub deleteShapeFile(ByVal ShpPath As String)

        Dim myPath As String = ShpPath
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(ShpPath, ".shp", ".shx", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(ShpPath, ".shp", ".dbf", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(ShpPath, ".shp", ".prj", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)

    End Sub

    Public Sub deleteGrid(ByVal GridPath As String)

        Dim myPath As String = GridPath
        Dim extention As String = Right(myPath, 4)

        'removes not only the grid file, but also the projection file, the hdr file (whatever that is) and the xml file.
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(GridPath, extention, ".prj", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(GridPath, extention, ".hdr", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)
        myPath = Replace(GridPath, extention, extention & ".aux.xml", , , CompareMethod.Text)
        If System.IO.File.Exists(myPath) Then System.IO.File.Delete(myPath)

    End Sub

    Public Function Percentile(ByVal Values() As Double, ByVal myPercentile As Double) As Double
        'this routine calculates the requested percentile from an array of values
        'first we'll need to sort the array
        Dim lower As Long, upper As Long
        Dim nSteps As Long = Values.Count - 1
        Array.Sort(Values)
        lower = RoundUD(nSteps * myPercentile, 0, False)
        upper = RoundUD(nSteps * myPercentile, 0, True)
        Return Interpolate(lower / nSteps, Values(lower), upper / nSteps, Values(upper), myPercentile, False)
    End Function

    Public Function PercentileFromList(ByVal Values As List(Of Double), ByVal myPercentile As Double) As Double
        'this routine calculates the requested percentile from an array of values
        'first we'll need to sort the array
        Dim lower As Long, upper As Long
        Dim nSteps As Long = Values.Count - 1
        Values.Sort()
        lower = RoundUD(nSteps * myPercentile, 0, False)
        upper = RoundUD(nSteps * myPercentile, 0, True)
        Return Interpolate(lower / nSteps, Values(lower), upper / nSteps, Values(upper), myPercentile, False)
    End Function

    Public Function CollectAllFilesInDir(ByVal path As String, ByVal SubDirs As Boolean, ByVal RightSidePartOfFileName As String, ByRef Paths As Collection) As Boolean
        Dim dirInfo As New IO.DirectoryInfo(path)
        Dim fileObject As FileSystemInfo

        If SubDirs = True Then
            For Each fileObject In dirInfo.GetFileSystemInfos()
                If System.IO.Directory.Exists(fileObject.FullName) Then
                    CollectAllFilesInDir(fileObject.FullName, SubDirs, RightSidePartOfFileName, Paths)
                Else
                    If Right(fileObject.FullName, RightSidePartOfFileName.Length).ToUpper = RightSidePartOfFileName.ToUpper Then
                        Paths.Add(fileObject.FullName)
                    End If
                End If
            Next
        Else
            For Each fileObject In dirInfo.GetFileSystemInfos()
                If Not System.IO.Directory.Exists(fileObject.FullName) Then
                    If Right(fileObject.FullName, RightSidePartOfFileName.Length).ToUpper = RightSidePartOfFileName.ToUpper Then
                        Paths.Add(fileObject.FullName)
                    End If
                End If
            Next
        End If
        Return True
    End Function


    Public Function StatsFromDataTable(ByRef dt As DataTable, FieldName As String, ByRef minVal As Double, ByRef maxVal As Double, ByRef avgVal As Double) As Boolean
        Try
            Dim myVal As Double, myMax As Double = -9.0E+99, myMin As Double = 9.0E+99, mySum As Double, n As Long
            For i = 0 To dt.Rows.Count - 1
                If Not IsDBNull(dt.Rows(i)(FieldName)) Then
                    myVal = dt.Rows(i)(FieldName)
                    If myVal > myMax Then myMax = myVal
                    If myVal < myMin Then myMin = myVal
                    mySum += myVal
                    n += 1
                End If
            Next

            If n > 0 Then
                minVal = myMin
                maxVal = myMax
                avgVal = mySum / n
            Else
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function PointToLineSnapping(x1 As Double, y1 As Double, x2 As Double, y2 As Double, pointx As Double, pointy As Double, SearchRadius As Double, ByRef Chainage As Double, ByRef SnapDist As Double) As Boolean
        '----------------------------------------------------------------------------------------------------------------------------------
        'This algorithm finds the snapping point of a point to a given line segment.
        'It does so by rotating the line segment and the point around the starting point of the line segment so that the line segment ligns with the x-axis
        'AND that the line segment has a positive direction
        'Author: Siebe Bosch
        'Date: 26 february 2017
        '----------------------------------------------------------------------------------------------------------------------------------
        Try
            Dim alpha As Double = LineAngleDegrees(x1, y1, x2, y2)              'angle of the line segment in degrees
            RotatePoint(x2, y2, x1, y1, 90 - alpha, x2, y2)                     'rotate the line so that it aligns with the x-axis
            RotatePoint(pointx, pointy, x1, y1, 90 - alpha, pointx, pointy)     'rotate the point-to-snap in a similar fashion

            If Math.Round(pointx, 2) >= Math.Round(x1, 2) AndAlso Math.Round(pointx, 2) <= Math.Round(x2, 2) AndAlso Math.Abs(pointy - y1) <= SearchRadius Then 'this is safe since we rotated the line segment so it has a positive direction
                Chainage = Math.Abs(pointx - x1)
                SnapDist = Math.Abs(pointy - y1)
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Me.setup.Log.AddError("Error in function PointToLineSnapping.")
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    'Public Function PointToLineSnapDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double, pointx As Double, pointy As Double, ByRef snapDistance As Double, ByRef snapChainage As Double, ByRef Xsect As Double, ByRef Ysect As Double) As Boolean
    '    'this routine find the snap distance for a point to a line segment, specified by two points
    '    'it does so by first calculating the perpendicular line and see whether they cross
    '    'note: TRUE is only returned when the snapping point is located ON the line segment between (x1,y1) and (x2,y2)

    '    'first define the line segment that we'll snap to
    '    Dim myLine As New clsLineDefinition(Me.setup)
    '    myLine = LineFromPoints(x1, y1, x2, y2)

    '    'then define a line perpendicular to the line segment, through the point to be snapped
    '    Dim perpLine As New clsLineDefinition(Me.setup)
    '    perpLine.a = -1 / myLine.a
    '    perpLine.b = pointy - perpLine.a * pointx

    '    'find out where both lines intersect
    '    Call LineIntersection(myLine.a, myLine.b, perpLine.a, perpLine.b, Xsect, Ysect)
    '    snapDistance = Pythagoras(pointx, pointy, Xsect, Ysect)
    '    snapChainage = Pythagoras(x1, y1, Xsect, Ysect)

    '    If Xsect >= Math.Min(x1, x2) AndAlso Xsect <= Math.Max(x1, x2) AndAlso Ysect >= Math.Min(y1, y2) AndAlso Ysect <= Math.Max(y1, y2) Then
    '        Return True
    '    Else
    '        Return False
    '    End If

    'End Function

    Public Function WidthOfCircle(r As Double, zCenter As Double, z As Double) As Double
        'returns the with of a circle at a given elevation
        If Math.Abs(z - zCenter) > r Then
            Return 0
        Else
            Return 2 * Math.Sqrt(r ^ 2 - (Math.Abs(z - zCenter)) ^ 2)
        End If
    End Function

    Public Function InterpolateFromDataTable(ByRef myTable As DataTable, ByVal SearchValue As Double, ByVal searchCol As Integer, ByVal returnCol As Integer) As Double
        Dim i As Long

        'decide if the table is ascending or descending
        If myTable.Rows(1)(searchCol) >= myTable.Rows(0)(searchCol) Then
            'in ascending tables
            If SearchValue <= myTable.Rows(0)(searchCol) Then
                Return myTable.Rows(0)(returnCol)
            ElseIf SearchValue >= myTable.Rows(myTable.Rows.Count - 1)(searchCol) Then
                Return myTable.Rows(myTable.Rows.Count - 1)(returnCol)
            Else
                For i = 0 To myTable.Rows.Count - 2
                    If myTable.Rows(i)(searchCol) <= SearchValue AndAlso myTable.Rows(i + 1)(searchCol) >= SearchValue Then
                        Return Interpolate(myTable.Rows(i)(searchCol), myTable.Rows(i)(returnCol), myTable.Rows(i + 1)(searchCol), myTable.Rows(i + 1)(returnCol), SearchValue)
                    End If
                Next
            End If
        Else
            If SearchValue >= myTable.Rows(0)(searchCol) Then
                Return myTable.Rows(0)(returnCol)
            ElseIf SearchValue <= myTable.Rows(myTable.Rows.Count - 1)(searchCol) Then
                Return myTable.Rows(myTable.Rows.Count - 1)(returnCol)
            Else
                'in descending tables
                For i = 0 To myTable.Rows.Count - 2
                    If myTable.Rows(i)(searchCol) >= SearchValue AndAlso myTable.Rows(i + 1)(searchCol) <= SearchValue Then
                        Return Interpolate(myTable.Rows(i)(searchCol), myTable.Rows(i)(returnCol), myTable.Rows(i + 1)(searchCol), myTable.Rows(i + 1)(returnCol), SearchValue)
                    End If
                Next
            End If
        End If
    End Function



    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        ' by making Generator static, we preserve the same instance '
        ' (i.e., do not create new instances with the same seed over and over) '
        ' between calls '
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function

    Public Function ParseDateString(ByVal DateStr As String, ByVal DateFormat As String, ByRef myDate As DateTime) As Boolean

        Dim myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMinute As Integer, mySecond As Integer

        Try

            Select Case DateFormat.ToUpper
                Case Is = "YYYYMMDD"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 4), 2)
                    myDay = Right(DateStr, 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "YYYYMMDDHH"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 6), 2)
                    myDay = Left(Right(DateStr, 4), 2)
                    myHour = Right(DateStr, 2)
                    myMinute = 0
                    mySecond = 0
                Case Is = "YYYYMMDDHHMM"
                    myYear = Left(DateStr, 4)
                    myMonth = Right(Left(DateStr, 6), 2)
                    myDay = Right(Left(DateStr, 8), 2)
                    myHour = Right(Left(DateStr, 10), 2)
                    myMinute = Right(Left(DateStr, 12), 2)
                    mySecond = 0
                Case Is = "YYYYMMDDHHMMSS"
                    myYear = Left(DateStr, 4)
                    myMonth = Right(Left(DateStr, 6), 2)
                    myDay = Right(Left(DateStr, 8), 2)
                    myHour = Right(Left(DateStr, 10), 2)
                    myMinute = Right(Left(DateStr, 12), 2)
                    mySecond = Right(DateStr, 2)
                Case Is = "YYYY/MM/DD", "YYYY-MM-DD"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 5), 2)
                    myDay = Right(DateStr, 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "YYYY/MM/DD HH: mm:SS", "YYYY-MM-DD HH:MM:SS"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 14), 2)
                    myDay = Left(Right(DateStr, 11), 2)
                    myHour = Left(Right(DateStr, 8), 2)
                    myMinute = Left(Right(DateStr, 5), 2)
                    mySecond = Right(DateStr, 2)
                Case Is = "YYYY/MM/DD HH:MM", "YYYY-MM-DD HH:MM"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 11), 2)
                    myDay = Left(Right(DateStr, 8), 2)
                    myHour = Left(Right(DateStr, 5), 2)
                    myMinute = Right(DateStr, 2)
                    mySecond = 0
                Case Is = "YYYY/MM/DD HH", "YYYY-MM-DD HH"
                    myYear = Left(DateStr, 4)
                    myMonth = Left(Right(DateStr, 8), 2)
                    myDay = Left(Right(DateStr, 5), 2)
                    myHour = Right(DateStr, 2)
                    myMinute = 0
                    mySecond = 0
                Case Is = "DD/MM/YYYY", "DD-MM-YYYY"
                    myYear = Right(DateStr, 4)
                    myMonth = Left(Right(DateStr, 7), 2)
                    myDay = Left(DateStr, 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "MM/DD/YYYY", "MM-DD-YYYY"
                    myYear = Right(DateStr, 4)
                    myMonth = Left(DateStr, 2)
                    myDay = Left(Right(DateStr, 7), 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "YYYY-MM-DD"
                    myYear = Left(DateStr, 4)
                    myMonth = Right(Left(DateStr, 7), 2)
                    myDay = Right(DateStr, 2)
                    myHour = 0
                    myMinute = 0
                    mySecond = 0
                Case Is = "YYYY-MM-DD:HH"
                    myYear = Left(DateStr, 4)
                    myMonth = Right(Left(DateStr, 7), 2)
                    myDay = Right(Left(DateStr, 10), 2)
                    myHour = Right(DateStr, 2)
                    myMinute = 0
                    mySecond = 0
            End Select
            myDate = New DateTime(myYear, myMonth, myDay, myHour, myMinute, mySecond)
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Date format not (yet) supported for parsing: " & DateFormat)
            Return False
        End Try





        'Dim yearstart As Integer, monthstart As Integer
        'Dim daystart As Integer, hourstart As Integer
        'Dim minstart As Integer, secstart As Integer

        'If (DateFormat = "" OrElse DateFormat.ToLower = "numeric") AndAlso IsNumeric(DateStr) Then
        '  Return Date.FromOADate(Val(DateStr))
        'Else
        '  yearstart = InStr(DateFormat, "y", CompareMethod.Text)
        '  If InStr(DateFormat, "yyyy", CompareMethod.Binary) > 0 Then
        '    myYear = Val(Mid(DateStr, yearstart, 4))
        '  Else
        '    myYear = Val("19" & Mid(DateStr, yearstart, 2))
        '  End If

        '  monthstart = InStr(DateFormat, "M", CompareMethod.Binary)
        '  If monthstart > 0 Then myMonth = Val(Mid(DateStr, monthstart, 2))

        '  daystart = InStr(DateFormat, "d", CompareMethod.Text)
        '  If daystart > 0 Then myDay = Val(Mid(DateStr, daystart, 2))

        '  hourstart = InStr(DateFormat, "h", CompareMethod.Text)
        '  If hourstart > 0 Then myHour = Val(Mid(DateStr, hourstart, 2))

        '  minstart = InStr(DateFormat, "M", CompareMethod.Binary)
        '  If minstart > 0 Then myMinute = Val(Mid(DateStr, minstart, 2))

        '  secstart = InStr(DateFormat, "s", CompareMethod.Text)
        '  If secstart > 0 Then mySecond = Val(Mid(DateStr, secstart, 2))

        '  myDate = New Date(myYear, myMonth, myDay, myHour, myMinute, mySecond)
        '  Return myDate
        'End If

    End Function

    Public Function AbsoluteToRelativePath(ByVal referencePath As [String], ByVal adjustPath As [String]) As [String]
        If [String].IsNullOrEmpty(referencePath) Then
            'Throw New ArgumentNullException("fromPath")
            Return adjustPath
        End If
        If [String].IsNullOrEmpty(adjustPath) Then
            Return adjustPath
            'Throw New ArgumentNullException("toPath")
        End If

        Dim fromUri As New Uri(referencePath)
        Dim toUri As New Uri(adjustPath)

        If fromUri.Scheme <> toUri.Scheme Then
            Return adjustPath
        End If
        ' path can't be made relative.
        Dim relativeUri As Uri = fromUri.MakeRelativeUri(toUri)
        Dim relativePath As [String] = Uri.UnescapeDataString(relativeUri.ToString())

        If toUri.Scheme.ToUpperInvariant() = "FILE" Then
            relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
        End If
        'relativePath = ".\" & relativePath

        Return relativePath
    End Function

    Public Function RelativeToAbsolutePath(ByVal myPath As String, ByVal myRootDir As String) As String

        Dim tmpPath As String
        If Mid(myPath.Trim, 2, 2) = ":\" Then
            'the given path is not relative but absolute. Do not change it!
            Return myPath
        Else
            While Left(myPath, 3) = "..\"
                myPath = Right(myPath, myPath.Length - 3)
                Dim myDirInfo As DirectoryInfo = Directory.GetParent(myRootDir)
                myRootDir = myDirInfo.FullName
            End While
            myPath = Replace(myPath, ".\", "")
            tmpPath = myRootDir & "\" & myPath
            tmpPath = Replace(tmpPath, "\\", "\")
        End If
        Return tmpPath

    End Function

    Public Function MileageOneUp(ByVal startNums As List(Of Integer), ByVal endNums As List(Of Integer), ByRef Stand As List(Of Integer)) As Boolean
        'werkt als een kilometerteller. Als het hectometergetal boven z'n maximum komt, springt hij terug naar nul
        'en gaat het getalletje ervoor een omhoog et cetera. Produceert TRUE bij succes
        'produceert FALSE als hij aan z'n eind is gekomen en niet verder kan ophogen
        Dim i As Integer, j As Integer
        Dim Done As Boolean, ThisIsTheEnd As Boolean

        'errorhandling
        If startNums.Count <> endNums.Count OrElse startNums.Count <> Stand.Count Then
            Me.setup.Log.AddError("Error in function MileageOneUp: the arrays must have the same size.")
            Return False
        End If

        'check whether the current state is possible. If not we'll assume it needs to be initialized and return the very first value
        For i = 0 To Stand.Count - 1
            If Stand(i) < startNums(i) OrElse Stand(i) > endNums(i) Then
                For j = 0 To Stand.Count - 1
                    Stand(j) = startNums(j)
                Next
                Return True
            End If
        Next

        'check whether the state is currently at its end. If so, return false
        ThisIsTheEnd = True
        For i = 0 To Stand.Count - 1
            If Stand(i) < endNums(i) Then
                ThisIsTheEnd = False
                Exit For
            End If
        Next
        If ThisIsTheEnd Then Return False

        'walk through the list of numbers and adjust the state
        Done = False
        i = Stand.Count
        While Not Done
            i -= 1
            If i < 0 Then
                Done = True
            ElseIf Stand(i) < endNums(i) Then
                Stand(i) += 1
                Done = True
            Else
                Stand(i) = startNums(i)
            End If
        End While
        Return True

    End Function

    Public Sub AddToString(ByRef myStr As String, NewText As String, AddDoubleQuotes As Boolean, AddLineBreak As Boolean)
        If AddDoubleQuotes Then myStr &= Chr(34) & NewText & Chr(34) Else myStr &= NewText
        If AddLineBreak Then myStr &= vbCrLf
    End Sub

    Public Function ParentDirFromDir(ByVal myDir As String) As String
        Dim ParentDir As String = ""
        Dim SlashPos As Integer, newPos As Integer
        If Right(myDir, 1) = "\" Then myDir = Left(myDir, myDir.Length - 1)

        'find the last slash
        SlashPos = 0
        newPos = InStr(myDir, "\")
        While newPos > SlashPos
            SlashPos = newPos
            newPos = InStr(SlashPos + 1, myDir, "\")
        End While

        Return Left(myDir, SlashPos - 1)

    End Function

    Function TransposeDataTable(ByVal dtOriginal As DataTable) As DataTable
        Dim dtReflection As New DataTable("Reflection")
        For i As Integer = 0 To dtOriginal.Rows.Count - 1
            dtReflection.Columns.Add(dtOriginal.Rows(i)(0))
        Next
        Dim row As DataRow
        For j As Integer = 1 To dtOriginal.Columns.Count - 1
            row = dtReflection.NewRow
            For k As Integer = 0 To dtOriginal.Rows.Count - 1
                row(k) = dtOriginal.Rows(k)(j)
            Next
            dtReflection.Rows.Add(row)
        Next
        Return dtReflection
    End Function

    Public Function BooleanFromText(ByVal myStr As String) As Boolean
        myStr = myStr.Trim.ToUpper

        Select Case myStr
            Case Is = "TRUE"
                Return True
            Case Is = "WAAR"
                Return True
            Case Is = "False"
                Return False
            Case Is = "ONWAAR"
                Return False
            Case Else
                Return False
        End Select

    End Function

    Public Function HasPrefix(ByVal myStr As String, ByVal myPrefix As String, ByVal CaseSensitive As Boolean) As Boolean

        myStr = myStr.Trim
        myPrefix = myPrefix.Trim

        If CaseSensitive Then
            If Left(myStr, myPrefix.Length) = myPrefix Then
                Return True
            Else
                Return False
            End If
        Else
            If Left(myStr.ToUpper, myPrefix.Length) = myPrefix.ToUpper Then
                Return True
            Else
                Return False
            End If
        End If

    End Function


    Public Function LineIntersection(ByVal a1 As Double, ByVal b1 As Double, ByVal a2 As Double, ByVal b2 As Double, ByRef X As Double, ByRef Y As Double) As Boolean
        'finds the intersection point for two lines, both defined as y = ax + b
        'and returns the X and Y-coordinate of that point

        'y = a1 * x + b1
        'y = a2 * x + b2
        'a1 * x + b1 = a2 * x + b2
        'x(a1 - a2) = b2 - b1
        'x = (b2-b1)/(a1-a2)

        If (a1 - a2) <> 0 Then
            X = (b2 - b1) / (a1 - a2)
            Y = a1 * X + b1
            Return True
        Else
            Return False
        End If

    End Function

    Public Function sendEmail(ByVal toAddress As String, ByVal Subject As String, ByVal Body As String)
        Try
            Dim message As System.Net.Mail.MailMessage
            Dim smtp As New System.Net.Mail.SmtpClient("mail.meteobase.nl")
            Dim fromMailAddress As System.Net.Mail.MailAddress
            Dim toMailAddress As System.Net.Mail.MailAddress

            smtp.Port = 26
            fromMailAddress = New System.Net.Mail.MailAddress("info@meteobase.nl")
            toMailAddress = New System.Net.Mail.MailAddress(toAddress)
            message = New System.Net.Mail.MailMessage()
            message.From = fromMailAddress
            message.To.Add(toMailAddress)
            message.Subject = Subject
            message.Body = Body

            smtp.EnableSsl = False
            smtp.UseDefaultCredentials = False
            smtp.Credentials = New System.Net.NetworkCredential("info@meteobase.nl", "@g3ntM327")
            smtp.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network

            smtp.Send(message)

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getSubDirFromPath(ByVal myPath As String, ByVal RootDir As String) As String
        Dim myDir As String = Path.GetDirectoryName(myPath)
        Return Replace(myDir, RootDir, "")
    End Function

    Public Function DirFromFileName(ByVal mypath As String) As String
        Return Path.GetDirectoryName(mypath)
    End Function

    ''' <summary>
    ''' This function validates the groundwater table provided
    ''' </summary>
    ''' <param name="myGT"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function ValidateGT(ByVal myGT As String, ByVal meanSL As Double, ByVal WP As Double) As String
        myGT = myGT.Trim.ToUpper
        myGT.Replace("*", "")
        Select Case myGT.Trim.ToUpper
            Case Is = "I"
                Return "I"
            Case Is = "II"
                Return "II"
            Case Is = "III"
                Return "III"
            Case Is = "IV"
                Return "IV"
            Case Is = "V"
                Return "V"
            Case Is = "VI"
                Return "VI"
            Case Is = "VII"
                Return "VII"
            Case Else
                'roughly estimate a ground water table by the difference between the mean surface level and the target level
                'in reality these are the groundwater tables:
                'GT:   GHG            GLG
                'I                < 50
                'II               50 - 80
                'III   < 40       80 - 120
                'IV    > 40       80 - 120
                'V     < 40       > 120
                'VI    40 - 80  > 120
                'VII   80       > 120
                If (meanSL - WP) <= 0.5 Then
                    Return "I"
                ElseIf (meanSL - WP) <= 1.2 Then
                    Return "III"
                ElseIf (meanSL - WP) <= 1.5 Then
                    Return "V"
                Else
                    Return "V"
                End If
        End Select
    End Function

    Friend Sub DischargeUnitConversion(ByRef Value As Double, ByVal DistFrom As String, ByVal DistTo As String, ByVal TimeFrom As String, ByVal TimeTo As String, Optional ByVal AreaM2 As Double = 0)
        'converteert de waarde van Value van de ene eenheid naar de andere
        'bijvoorbeeld m3/s naar mm/h
        If DistFrom = DistTo AndAlso TimeFrom = TimeTo Then Exit Sub

        Select Case DistFrom
            Case Is = "m3"
                Select Case DistTo
                    Case Is = "mm"
                        Value = Value / AreaM2 * 1000
                    Case Is = "m"
                        Value = Value / AreaM2
                End Select

            Case Is = "mm"
                Select Case DistTo
                    Case Is = "m3"
                        Value = Value / 1000 * AreaM2
                    Case Is = "m"
                        Value = Value / 1000
                End Select

            Case Is = "m"
                Select Case DistTo
                    Case Is = "m3"
                        Value = Value * AreaM2
                    Case Is = "mm"
                        Value = Value * 1000
                End Select

        End Select

        Select Case TimeFrom
            Case Is = "y"
                Select Case TimeTo
                    Case Is = "d"
                        Value = Value / 365
                    Case Is = "h"
                        Value = Value / 365 / 24
                    Case Is = "min"
                        Value = Value / 365 / 24 / 60
                    Case Is = "s"
                        Value = Value / 365 / 24 / 3600
                End Select

            Case Is = "d"
                Select Case TimeTo
                    Case Is = "y"
                        Value = Value * 365
                    Case Is = "h"
                        Value = Value / 24
                    Case Is = "min"
                        Value = Value / 24 / 60
                    Case Is = "s"
                        Value = Value / 24 / 3600
                End Select

            Case Is = "h"
                Select Case TimeTo
                    Case Is = "y"
                        Value = Value * 24 * 365
                    Case Is = "d"
                        Value = Value * 24
                    Case Is = "min"
                        Value = Value / 60
                    Case Is = "s"
                        Value = Value / 3600
                End Select

            Case Is = "s"
                Select Case TimeTo
                    Case Is = "y"
                        Value = Value * 3600 * 24 * 365
                    Case Is = "d"
                        Value = Value * 3600 * 24
                    Case Is = "h"
                        Value = Value * 3600
                    Case Is = "min"
                        Value = Value * 60
                End Select
        End Select
    End Sub


    Public Function RD2WGS84(ByVal X As Double, ByVal Y As Double, ByRef Lat As Double, ByRef Lon As Double) As Boolean
        'converteert RD-coordinaten naar Lat/Long (WGS84)
        'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
        Dim phi As Double
        Dim lambda As Double
        Dim PhiWGS As Double
        Dim LambdaWGS As Double

        Call RD2BESSEL(X, Y, phi, lambda)
        Call BESSEL2WGS84(phi, lambda, PhiWGS, LambdaWGS)
        Lat = PhiWGS
        Lon = LambdaWGS
        Return True

    End Function

    Public Function removePrefixFromID(ByVal ID As String, ByVal Prefixes As List(Of String)) As String
        Dim Prefix As String

        'Author: Siebe Bosch
        'Date: 11-7-2013
        'Description: removes a prefix from an ID and returns the result
        For Each Prefix In Prefixes
            If Left(ID.Trim.ToUpper, Prefix.Trim.Length) = Prefix.Trim.ToUpper Then
                Return Right(ID.Trim, ID.Trim.Length - Prefix.Trim.Length)
            End If
        Next
        Return ID.Trim 'prefix not found so return the original string

    End Function

    Public Function WGS842RD(ByVal Lat As Double, ByVal Lon As Double, Optional ByRef X As Double = 0, Optional ByRef y As Double = 0) As String
        'converteert WGS84-coordinaten (Lat/Long) naar RD
        'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
        Dim phiBes As Double
        Dim LambdaBes As Double
        Call WGS842BESSEL(Lat, Lon, phiBes, LambdaBes)
        Call BESSEL2RD(phiBes, LambdaBes, X, y)
        WGS842RD = X & "," & y

    End Function

    Public Function RemoveBoundingQuotes(ByVal myStr As String) As String
        If Left(myStr, 1) = Chr(34) OrElse Left(myStr, 1) = "'" Then myStr = Right(myStr, myStr.Length - 1)
        If Right(myStr, 1) = Chr(34) OrElse Right(myStr, 1) = "'" Then myStr = Left(myStr, myStr.Length - 1)
        Return myStr
    End Function

    Public Function GetBooleanFromString(ByVal myStr As String) As Boolean
        Dim tmpStr As String = myStr.Trim.ToUpper
        Select Case tmpStr
            Case Is = "TRUE"
                Return True
            Case Is = "WAAR"
                Return True
            Case Is = "FALSE"
                Return False
            Case Is = "UNTRUE"
                Return False
            Case Is = "ONWAAR"
                Return False
            Case Else
                Return False
        End Select
    End Function

    Public Sub RD2BESSEL(ByVal X As Double, ByVal y As Double, ByRef phi As Double, ByRef lambda As Double)

        'converteert RD-coordinaten naar phi en lambda voor een Bessel-functie
        'code is geheel gebaseerd op de routines van Ejo Schrama's software:
        'schrama@geo.tudelft.nl

        Dim x0 As Double
        Dim y0 As Double
        Dim k As Double
        Dim bigr As Double
        Dim m As Double
        Dim n As Double
        Dim lambda0 As Double
        Dim phi0 As Double
        Dim l0 As Double
        Dim b0 As Double
        Dim e As Double
        Dim a As Double

        Dim d_1 As Double, d_2 As Double, r As Double, sa As Double, ca As Double, psi As Double, cpsi As Double, spsi As Double
        Dim sb As Double, cb As Double, b As Double, sdl As Double, dl As Double, w As Double, q As Double
        Dim dq As Double, i As Long, pi As Double

        x0 = 155000
        y0 = 463000
        k = 0.9999079
        bigr = 6382644.571
        m = 0.003773953832
        n = 1.00047585668

        pi = Math.PI
        'pi = 3.14159265358979
        lambda0 = pi * 0.0299313271611111
        phi0 = pi * 0.289756447533333
        l0 = pi * 0.0299313271611111
        b0 = pi * 0.289561651383333

        e = 0.08169683122
        a = 6377397.155

        d_1 = X - x0
        d_2 = y - y0
        r = Math.Sqrt(d_1 ^ 2 + d_2 ^ 2)

        If r <> 0 Then
            sa = d_1 / r
            ca = d_2 / r
        Else
            sa = 0
            ca = 0
        End If

        psi = Math.Atan2(r, k * 2 * bigr) * 2
        cpsi = Math.Cos(psi)
        spsi = Math.Sin(psi)

        sb = ca * Math.Cos(b0) * spsi + Math.Sin(b0) * cpsi
        d_1 = sb
        cb = Math.Sqrt(1 - d_1 ^ 2)
        b = Math.Acos(cb)
        sdl = sa * spsi / cb
        dl = Math.Asin(sdl)
        lambda = dl / n + lambda0
        w = Math.Log(Math.Tan(b / 2 + pi / 4))
        q = (w - m) / n

        phi = Math.Atan(Math.Exp(1) ^ q) * 2 - pi / 2 'phi prime
        For i = 1 To 4
            dq = e / 2 * Math.Log((e * Math.Sin(phi) + 1) / (1 - e * Math.Sin(phi)))
            phi = Math.Atan(Math.Exp(1) ^ (q + dq)) * 2 - pi / 2
        Next

        lambda = lambda / pi * 180
        phi = phi / pi * 180

    End Sub

    Public Sub BESSEL2WGS84(ByVal phi As Double, ByVal lambda As Double, ByRef PhiWGS As Double, ByRef LamWGS As Double)
        Dim dphi As Double, dlam As Double, phicor As Double, lamcor As Double

        dphi = phi - 52
        dlam = lambda - 5
        phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
        lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
        PhiWGS = phi + phicor
        LamWGS = lambda + lamcor


    End Sub

    Public Sub WGS842BESSEL(ByVal PhiWGS As Double, ByVal LamWGS As Double, ByRef phi As Double, ByRef lambda As Double)
        Dim dphi As Double, dlam As Double, phicor As Double, lamcor As Double

        dphi = PhiWGS - 52
        dlam = LamWGS - 5
        phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
        lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
        phi = PhiWGS - phicor
        lambda = LamWGS - lamcor

    End Sub

    Public Sub BESSEL2RD(ByVal phiBes As Double, ByVal lamBes As Double, ByRef X As Double, ByRef y As Double)

        'converteert Lat/Long van een Bessel-functie naar X en Y in RD
        'code is geheel gebaseerd op de routines van Ejo Schrama's software:
        'schrama@geo.tudelft.nl

        Dim x0 As Double
        Dim y0 As Double
        Dim k As Double
        Dim bigr As Double
        Dim m As Double
        Dim n As Double
        Dim lambda0 As Double
        Dim phi0 As Double
        Dim l0 As Double
        Dim b0 As Double
        Dim e As Double
        Dim a As Double

        Dim d_1 As Double, d_2 As Double, r As Double, sa As Double, ca As Double, cpsi As Double, spsi As Double
        Dim b As Double, dl As Double, w As Double, q As Double
        Dim dq As Double, pi As Double, phi As Double, lambda As Double, s2psihalf As Double, cpsihalf As Double, spsihalf As Double
        Dim tpsihalf As Double

        x0 = 155000
        y0 = 463000
        k = 0.9999079
        bigr = 6382644.571
        m = 0.003773953832
        n = 1.00047585668

        pi = Math.PI
        'pi = 3.14159265358979
        lambda0 = pi * 0.0299313271611111
        phi0 = pi * 0.289756447533333
        l0 = pi * 0.0299313271611111
        b0 = pi * 0.289561651383333

        e = 0.08169683122
        a = 6377397.155

        phi = phiBes / 180 * pi
        lambda = lamBes / 180 * pi

        q = Math.Log(Math.Tan(phi / 2 + pi / 4))
        dq = e / 2 * Math.Log((e * Math.Sin(phi) + 1) / (1 - e * Math.Sin(phi)))
        q = q - dq
        w = n * q + m
        b = Math.Atan(Math.Exp(1) ^ w) * 2 - pi / 2
        dl = n * (lambda - lambda0)
        d_1 = Math.Sin((b - b0) / 2)
        d_2 = Math.Sin(dl / 2)
        s2psihalf = d_1 * d_1 + d_2 * d_2 * Math.Cos(b) * Math.Cos(b0)
        cpsihalf = Math.Sqrt(1 - s2psihalf)
        spsihalf = Math.Sqrt(s2psihalf)
        tpsihalf = spsihalf / cpsihalf
        spsi = spsihalf * 2 * cpsihalf
        cpsi = 1 - s2psihalf * 2
        sa = Math.Sin(dl) * Math.Cos(b) / spsi
        ca = (Math.Sin(b) - Math.Sin(b0) * cpsi) / (Math.Cos(b0) * spsi)
        r = k * 2 * bigr * tpsihalf
        X = Math.Round(r * sa + x0, 0)
        y = Math.Round(r * ca + y0, 0)

    End Sub

    Friend Function IDsSimilar(ByVal refID As String, ByVal myID As String) As Boolean

        'alles in bovenkast zetten
        refID = refID.Trim.ToUpper
        myID = myID.Trim.ToUpper

        'leading zeroes toevoegen om de ID's te uniformeren
        refID = AddLeadingZeroesToID(refID, "KGM", 5)
        refID = AddLeadingZeroesToID(refID, "KST", 5)
        refID = AddLeadingZeroesToID(refID, "KDU", 5)

        myID = AddLeadingZeroesToID(myID, "KGM", 5)
        myID = AddLeadingZeroesToID(myID, "KST", 5)
        myID = AddLeadingZeroesToID(myID, "KDU", 5)

        'compares two ID's and returns a boolean whether they are similar
        If refID = myID Then
            Return True
        ElseIf refID Like "*" & myID Then
            Return True
        ElseIf refID Like "*" & myID Then
            Return True
        Else
            Return False
        End If

    End Function

    Friend Function AddLeadingZeroesToID(ByVal ID, ByVal Identifier, ByVal nChars) As String
        'deze functie vult een ID zo aan dat je een vast aantal karakters achter de identifier krigjt
        'bijvoorbeeld: KGM001 met Identifier KGM en nChars=5 wordt KGM00001

        Dim iLoc As Integer = InStr(ID, Identifier)
        Dim rStr As String, lStr As String, i As Long

        If iLoc > 0 Then
            lStr = Left(ID, iLoc - 1)
            rStr = Right(ID, ID.length - Identifier.length + 1)
            For i = nChars To rStr.Length Step -1
                rStr = "0" & rStr
            Next
            Return lStr & Identifier & rStr
        Else
            Return ID
        End If
    End Function

    Friend Function Pythagoras(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
        Return Math.Sqrt((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
    End Function

    Friend Function GetDirFromPath(ByVal path As String) As String
        Try
            Return path.Substring(0, path.LastIndexOf("\") + 1)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Friend Function GetFileNameFromPath(ByVal path As String) As String
        Try
            Return path.Substring(path.LastIndexOf("\") + 1)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function MergeDataTablesFromDictionary(ByRef Tables As Dictionary(Of String, DataTable)) As DataTable
        Dim i As Integer, newTable As New DataTable, myTable As DataTable, newRow As DataRow
        Dim r As Long

        Try
            'use the first table as a basis. The title and data type for the first column will be extracte from that table
            newTable = New DataTable
            newTable.Columns.Add(Tables.Values(0).Columns(0).Caption, Tables.Values(0).Columns(0).DataType)

            'add the columns
            For i = 0 To Tables.Values.Count - 1
                myTable = Tables.Values(i)
                newTable.Columns.Add(Tables.Keys(i), myTable.Columns(1).DataType)
            Next

            'add the rows
            For r = 0 To Tables.Values(0).Rows.Count - 1
                newRow = newTable.NewRow()
                newRow(0) = Tables.Values(0).Rows(r)(0)

                For i = 0 To Tables.Values.Count - 1
                    myTable = Tables.Values(i)
                    newRow(i + 1) = Convert.ToDouble(myTable.Rows(r)(1))
                Next
                newTable.Rows.Add(newRow)
            Next

            Return newTable
        Catch ex As Exception
            Me.setup.Log.AddError("error merging datatables.")
            Return Nothing
        End Try


    End Function

    Public Function MDBnoQuery(ByRef con As OleDb.OleDbConnection, ByVal CommandText As String, Optional ByVal CloseAfterwards As Boolean = True) As Boolean
        'returns the number of rows affected
        Dim nAffected As Integer

        Try
            If Not con.State = ConnectionState.Open Then con.Open()
            Dim cmd As New OleDb.OleDbCommand
            cmd.Connection = con
            cmd.CommandText = CommandText
            nAffected = cmd.ExecuteNonQuery

            If nAffected < 1 Then Me.setup.Log.AddWarning("Query affected zero rows: " & CommandText)
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        Finally
            If CloseAfterwards Then con.Close()
        End Try
    End Function

    Public Function MDBQuery(ByRef con As System.Data.OleDb.OleDbConnection, ByVal myQuery As String, ByRef myTable As System.Data.DataTable, Optional ByVal CloseAfterwards As Boolean = True) As Boolean
        'queries an MBD (Access) database and returns the results in a datatable
        Try
            Dim da As OleDb.OleDbDataAdapter
            If Not con.State = ConnectionState.Open Then con.Open()

            da = New OleDb.OleDbDataAdapter(myQuery, con.ConnectionString)
            da.Fill(myTable)
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error executing query " & myQuery)
            Me.setup.Log.AddError(ex.Message)
            Return False
        Finally
            If CloseAfterwards Then con.Close()
        End Try

    End Function

    Public Function MDBRenameTable(ByRef con As OleDbConnection, OldName As String, NewName As String, CreateNewTableIfNotExists As Boolean) As Boolean
        'this function renames an Access table using SQL. It requires two commands
        Try

            If MDBTableExists(con, OldName) Then
                Dim query As String
                query = "SELECT * INTO " & NewName & " FROM " & OldName & ";"
                MDBnoQuery(con, query, False)
                query = "DROP TABLE " & OldName & ";"
                MDBnoQuery(con, query, True)
            Else
                If CreateNewTableIfNotExists Then
                    MDBCreateTable(con, NewName)
                End If
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function MDBTableExists(ByRef con As OleDbConnection, TableName As String) As Boolean
        'Dim schema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, TableName})
        'Return schema.Rows.OfType(Of DataRow)().Any(Function(r) r.ItemArray(2).ToString().ToLower() = TableName.ToLower())
        Dim myTable As Object, TableSchema As Object

        If Not con.State = ConnectionState.Open Then con.Open()
        TableSchema = con.GetSchema("TABLES")

        myTable = TableSchema.select("TABLE_NAME='" & TableName & "'")
        Return (myTable.length > 0)
    End Function

    Public Function MDBCreateTable(ByRef con As System.Data.OleDb.OleDbConnection, TableName As String) As Boolean
        Try
            Dim myTable As Object, TableSchema As Object
            Dim query As String

            If Not con.State = ConnectionState.Open Then con.Open()
            TableSchema = con.GetSchema("TABLES")

            myTable = TableSchema.select("TABLE_NAME='" & TableName & "'")
            If myTable.length = 0 Then
                query = "CREATE TABLE " & TableName & ";"
                setup.GeneralFunctions.MDBnoQuery(con, query, False)
            End If
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error creating table " & TableName & " in database.")
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function MDBCreateColumn(ByRef con As System.Data.OleDb.OleDbConnection, TableName As String, ColumnName As String, DataType As String, Optional ByVal COLINDEXNAME As String = "") As Boolean
        Try
            Dim TableSchema As Object, query As String
            Dim ColSchema As Object, col As Object

            If Not con.State = ConnectionState.Open Then con.Open()
            TableSchema = con.GetSchema("TABLES")
            ColSchema = con.GetSchema("COLUMNS")

            'create the column if it does not exist
            col = ColSchema.Select("TABLE_NAME='" & TableName & "' AND COLUMN_NAME ='" & ColumnName & "'")
            If col.Length = 0 Then
                query = "ALTER TABLE " & TableName & " ADD COLUMN " & ColumnName & " " & DataType & ";"
                setup.GeneralFunctions.MDBnoQuery(con, query, False)

                'column was created. Now set it as indexed, if required
                If COLINDEXNAME <> "" Then
                    query = "CREATE INDEX " & COLINDEXNAME & " ON " & TableName & " (" & ColumnName & ");"
                    setup.GeneralFunctions.MDBnoQuery(con, query, False)
                End If
            End If


            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error creating column " & ColumnName & " in database table " & TableName)
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function


    Public Function MDBDeleteColumn(ByRef con As System.Data.OleDb.OleDbConnection, TableName As String, ColumnName As String) As Boolean
        Try
            Dim TableSchema As Object, query As String
            Dim ColSchema As Object, col As Object

            If Not con.State = ConnectionState.Open Then con.Open()
            TableSchema = con.GetSchema("TABLES")
            ColSchema = con.GetSchema("COLUMNS")

            'create the column if it does not exist
            col = ColSchema.Select("TABLE_NAME='" & TableName & "' AND COLUMN_NAME ='" & ColumnName & "'")
            If col.Length > 0 Then
                query = "ALTER TABLE " & TableName & " DROP " & ColumnName & ";"
                setup.GeneralFunctions.MDBnoQuery(con, query, False)
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error removing table " & ColumnName & " from database table " & TableName)
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function MDBDropColumn(ByRef con As System.Data.OleDb.OleDbConnection, TableName As String, ColumnName As String) As Boolean
        Try
            Dim ColSchema As Object, col As Object
            Dim query As String

            If Not con.State = ConnectionState.Open Then con.Open()
            ColSchema = con.GetSchema("COLUMNS")

            col = ColSchema.select("TABLE_NAME='" & TableName & "' AND COLUMN_NAME='" & ColumnName & "'")
            If col.length > 0 Then
                query = "ALTER TABLE " & TableName & " DROP COLUMN " & ColumnName & ";"
                setup.GeneralFunctions.MDBnoQuery(con, query, False)
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error deleting column " & ColumnName & " from table " & TableName & " in database.")
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function MDBDropTable(ByRef con As System.Data.OleDb.OleDbConnection, TableName As String) As Boolean
        Try
            Dim TableSchema As Object, Table As Object
            Dim query As String

            If Not con.State = ConnectionState.Open Then con.Open()
            TableSchema = con.GetSchema("TABLES")

            Table = TableSchema.select("TABLE_NAME='" & TableName & "'")
            If Table.length > 0 Then
                query = "DROP TABLE " & TableName & ";"
                setup.GeneralFunctions.MDBnoQuery(con, query, False)
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Error deleting table " & TableName & " from database.")
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function ReplaceStringContentByToken(ByRef myContent As String, ByVal Token As String, ByVal newVal As Double) As Boolean
        Dim pos1 As Long, pos2 As Long
        Dim lines() As String, i As Long

        'aangepast op 7 nov 2015 door Siebe Bosch. het bewerken van de hele filecontent in een keer bleek te traag
        'daarom nu een split by harde return toegevoegd, en regel voor regel afhandelen

        Try
            lines = Split(myContent, vbCrLf)
            For i = 0 To lines.Count - 1
                While InStr(lines(i), Token, CompareMethod.Text) > 0
                    pos1 = InStr(lines(i), Token, CompareMethod.Text)
                    pos2 = InStr(pos1 + 1, lines(i), Token, CompareMethod.Text)
                    If pos1 > 0 AndAlso pos2 > pos1 Then
                        lines(i) = Left(lines(i), pos1 - 1) & newVal & Right(lines(i), lines(i).Length - pos2 - Token.Length + 1)
                    Else
                        Throw New Exception("Error replacing string content by token " & Token & ". Token is probably not symmetric around value to be adjusted.")
                    End If
                End While
            Next
            myContent = lines(0)
            For i = 1 To lines.Count - 1
                myContent &= vbCrLf & lines(i)
            Next

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function GetExponentFromNumber(ByVal myNumber As Double) As Integer
        Dim myPower As Integer = 0

        If myNumber < 1 Then
            While Not myNumber > 1
                myPower -= 1
                myNumber *= 10
            End While
        ElseIf myNumber >= 1 Then
            While Not myNumber / 10 < 1
                myPower += 1
                myNumber /= 10
            End While
        End If
        Return myPower

    End Function

    Public Function GetAxisValuesFromRange(ByVal min As Double, ByVal max As Double, ByRef chartMin As Double, ByRef chartMax As Double) As Boolean
        Dim myRange As Double
        Dim myExponent As Integer
        Dim n As Integer
        Try
            chartMin = min
            chartMax = max
            myRange = Math.Abs(max - min)

            If myRange = 0 Then
                If min = 0 Then
                    chartMin = min - 1
                    chartMax = min + 1
                    Return True
                Else
                    'min and max are equal so find a decent min and max around it
                    chartMin = setup.GeneralFunctions.RoundUD(Math.Min(min / 2, min * 2), 0, False)
                    chartMax = setup.GeneralFunctions.RoundUD(Math.Max(min / 2, min * 2), 0, True)
                    Return True
                End If
            End If

            myExponent = GetExponentFromNumber(myRange)
            If myExponent < 0 Then
                While RoundUD(min, 0, False) = RoundUD(max, 0, False)
                    min *= 10
                    max *= 10
                    n += 1
                End While
                chartMin = RoundUD(min, 0, False) / 10 ^ n
                chartMax = RoundUD(max, 0, True) / 10 ^ n
            ElseIf myExponent >= 0 Then
                chartMin = RoundUD(min / 10 ^ (myExponent - 1), 0, False) * 10 ^ (myExponent - 1)
                chartMax = RoundUD(max / 10 ^ (myExponent - 1), 0, True) * 10 ^ (myExponent - 1)
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function



    Friend Function GetMapWinGridCellCenterXY(ByVal XLLCenter As Double, ByVal YLLCenter As Double, ByVal DX As Double, ByVal DY As Double, ByVal nRows As Integer, ByVal nCols As Integer, ByVal rowIdx As Integer, ByVal colIdx As Integer, ByRef X As Double, ByRef Y As Double) As Boolean
        ' Siebe Bosch 8 July: adjusted
        ' MapWindow telt rijen van boven naar beneden (gechecked!)
        ' let op: gaat ervan uit dat de rowidx en colidx 0-based zijn
        X = XLLCenter + DX * colIdx
        Y = YLLCenter + ((nRows - 1) * DY) - (DY * rowIdx)
        Return True
    End Function

    Friend Shared Function ConvertoDate(ByVal dateString As String, ByRef result As DateTime) As Boolean
        Try
            Dim supportedFormats() As String = New String() {"MM/dd/yyyy", "MM/dd/yy", "ddMMMyyyy", "dMMMyyyy"}

            result = DateTime.ParseExact(dateString, supportedFormats, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    Friend Function ConvertToDateTime(ByVal datestring As String, ByVal formats As String) As DateTime
        Dim Year As Integer, Month As Integer, Day As Integer, Hour As Integer, Minute As Integer, Second As Integer
        Dim dateObject As DateTime

        If formats = "yyyy/MM/dd;HH:mm:ss" Then
            Year = Left(datestring, 4)
            Month = Mid(datestring, 6, 2)
            Day = Mid(datestring, 9, 2)
            Hour = Mid(datestring, 12, 2)
            Minute = Mid(datestring, 15, 2)
            Second = Right(datestring, 2)
        End If
        dateObject = New DateTime(Year, Month, Day, Hour, Minute, Second)
        Return dateObject

    End Function

    Public Function ParseDouble(ByRef myString As String, Optional ByVal Delimiter As String = " ") As Double
        Dim Done As Boolean, NewString As String = "", newChar As String
        myString = myString.Trim()
        While Not Done
            newChar = Left(myString, 1)
            myString = Right(myString, myString.Length - 1)
            If newChar = Delimiter Then
                Done = True
            Else
                NewString &= newChar
            End If
        End While
        Return Convert.ToDouble(NewString)

    End Function

    Public Function ParseString(ByRef myString As String, Optional ByVal Delimiter As String = " ",
                                Optional ByVal QuoteHandlingFlag As Integer = 1, Optional ByVal ResultMustBeNumeric As Boolean = False) As String

        Dim quoteEven As Boolean
        Dim tmpString As String = "", tmpChar As String = ""

        quoteEven = True

        'Quotehandlingflag: default = 1
        '0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
        '1 = items between single quotes are being treated as separate items (no parsing between single quotes)
        '2 = items between double quotes are being treated as separate items (no parsing between double quotes)

        Dim i As Integer
        For i = 1 To Len(myString)
            'snoep een karakter af van de string
            tmpChar = Left(myString, 1)

            If (tmpChar = "'" And QuoteHandlingFlag = 1) Or (tmpChar = Chr(34) And QuoteHandlingFlag = 2) Then
                If quoteEven = True Then
                    quoteEven = False
                    tmpString = tmpString & tmpChar
                    myString = Right(myString, myString.Length - 1)
                Else
                    quoteEven = True 'dit betekent dat we klaar zijn
                    tmpString = Right(tmpString, tmpString.Length - 1) 'laat bij het teruggeven meteen de quotes maar weg!
                    myString = Right(myString, myString.Length - 1).Trim
                    'Return tmpString
                    Exit For
                End If
            ElseIf tmpChar = Delimiter And quoteEven = True Then
                If Not tmpString = "" Then
                    myString = Right(myString, myString.Length - 1)
                    'Return tmpString
                    Exit For
                Else
                    myString = Right(myString, myString.Length - 1)
                End If
            Else
                myString = Right(myString, myString.Length - 1)
                tmpString = tmpString & tmpChar
            End If
        Next

        If ResultMustBeNumeric AndAlso Not IsNumeric(tmpString) Then
            myString = tmpString & Delimiter & myString
            Me.setup.Log.AddWarning("Numeric value expected after token " & tmpString & " in string " & myString & ". Value of 0 was returned.")
            Return 0
        Else
            Return tmpString
        End If


    End Function

    Public Function ParseStringToDictionary(ByVal myString As String, ByVal Delimiter As String) As Dictionary(Of String, String)
        'siebe bosch, 22-2-2015
        'parses a string and returns the results in a list of strings
        Dim newDict As New Dictionary(Of String, String)
        Dim myID As String
        While Not myString = ""
            myID = ParseString(myString, Delimiter)
            If Not newDict.ContainsKey(myID.Trim.ToUpper) Then
                newDict.Add(myID.Trim.ToUpper, myID.Trim)
            End If
        End While
        Return newDict
    End Function

    Public Function SplitTimeSpan(ByVal TimeSpan As clsTimeSpan, ByVal divider As Long) As Dictionary(Of Long, clsTimeSpan)
        'splits a given timespan e.g. 12 to 200404 into (almost) equal parts
        Dim SegLength As Long = Me.setup.GeneralFunctions.RoundUD((TimeSpan.GetLastTS - TimeSpan.GetFirstTS + 1) / divider, 0, False)
        Dim myDict As New Dictionary(Of Long, clsTimeSpan), mySpan As clsTimeSpan
        Dim i As Long, tsstart As Long, tsend As Long

        For i = 1 To divider
            tsstart = TimeSpan.GetFirstTS + (i - 1) * SegLength
            If i = divider Then
                tsend = TimeSpan.GetLastTS
            Else
                tsend = Math.Min(TimeSpan.GetFirstTS + i * SegLength - 1, TimeSpan.GetLastTS)
            End If
            mySpan = New clsTimeSpan(tsstart, tsend)
            myDict.Add(i, mySpan)
        Next
        Return myDict
    End Function

    Public Function IsEven(ByVal myNum As Long) As Boolean
        If (myNum And 1) = 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function Interpolate(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double,
                                ByVal X3 As Double, Optional ByVal BlockInterpolate As Boolean = False) As Double

        If X3 < X1 And X3 < X2 Then 'is niet interpoleren maar extrapoleren. Handhaaf de buitenste waarde
            Return Y1
        ElseIf X3 > X2 And X3 > X1 Then 'is niet interpoleren maar extrapoleren. Handhaaf de buitenste waarde
            Return Y2
        ElseIf X1 = X2 Then
            Return Y1
        Else
            If BlockInterpolate = True Then
                Return Y1
            Else
                Return Y1 + (Y2 - Y1) / (X2 - X1) * (X3 - X1)
            End If
        End If
    End Function

    Friend Function Extrapolate(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal X3 As Double) As Double
        'extrapolates linearly

        Dim Rico As Double = 0
        If X3 > X2 Then
            Rico = (Y2 - Y1) / (X2 - X1)
            Extrapolate = Y2 + (X3 - X2) * Rico
        ElseIf X3 < X1 Then
            Rico = (Y2 - Y1) / (X2 - X1)
            Extrapolate = Y1 - (X1 - X3) * Rico
        Else
            Extrapolate = -999
        End If

    End Function



    Public Function RoundUD(ByVal numD As Double, ByVal nDecimals As Integer, ByVal Up As Boolean) As Decimal

        Try
            Dim r As Double = 10 ^ (nDecimals)
            numD = r * numD
            Dim RoundedDown As Long = Math.Round(numD, 0)
            If RoundedDown > numD Then RoundedDown -= 1
            Dim Diff As Double = numD - RoundedDown

            If Diff > 0 Then
                If Up Then
                    Return (RoundedDown + 1) / r
                Else
                    Return (RoundedDown) / r
                End If
            Else
                Return (RoundedDown) / r
            End If
        Catch ex As Exception
            Me.setup.Log.AddError("Error in function RoundUD")
        End Try

    End Function

    Friend Function IDFROMSTRING(ByVal myStr As String, Optional ByVal Prefix As String = "", Optional ByVal CutoffString As String = "", Optional ByVal RemovePrefix As Boolean = True) As String
        Dim PrefixPos As Integer
        Dim tmpstr As String
        Dim ID As String = ""

        If Not CutoffString = "" And Not Prefix = "" Then
            'net zolang parsen tot we tegenkomen wat we nodig hebben
            While Not myStr = ""
                tmpstr = Me.ParseString(myStr, CutoffString)
                PrefixPos = InStr(1, tmpstr, Prefix, CompareMethod.Text)
                If PrefixPos > 0 Then
                    ID = Right(tmpstr, Len(tmpstr) - PrefixPos + 1)
                    If RemovePrefix Then
                        Return Right(ID, ID.Length - Prefix.Length)
                    Else
                        Return ID
                    End If

                End If
            End While
        Else
            Me.setup.Log.AddWarning("Functie zonder prefix of afbreekstring nog niet ondersteund.")
        End If

        Return String.Empty
    End Function

    Friend Function ParsestringNumeric(ByRef myString As String, Optional ByVal delimiter As String = " ") As Double
        Dim tmpString As String = "", tmpChar As String = ""
        Dim i As Long

        myString = myString.Trim
        For i = 1 To Len(myString)
            'snoep een karakter af van de string
            tmpChar = Mid(myString, i, 1)
            If tmpChar = delimiter Then
                If IsNumeric(tmpString) Then
                    myString = Right(myString, myString.Length - tmpString.Length - 1)
                    Return Val(tmpString)
                End If
            Else
                tmpString &= tmpChar
            End If
        Next i

        Return -999

    End Function

    Public Function DateIntIsValid(ByVal dateint As Integer) As Boolean
        Dim myYear As Integer
        Dim myMonth As Integer
        Dim myDay As Integer

        myYear = Left(dateint, 4)
        myMonth = Left(Right(dateint, 4), 2)
        myDay = Right(dateint, 2)

        If myYear < 1900 OrElse myYear > 2100 Then
            Return False
        ElseIf myMonth > 12 OrElse myMonth < 1 Then
            Return False
        ElseIf myDay < 1 Or myDay > 31 Then
            Return False
        ElseIf myDay > 30 AndAlso (myMonth = 2 Or myMonth = 4 Or myMonth = 6 Or myMonth = 9 Or myMonth = 11) Then
            Return False
        ElseIf myMonth = 2 And myDay > 29 Then
            Return False
        ElseIf myMonth = 2 And myDay > 28 And Not IsLeapYear(myYear) Then
            Return False
        End If
        Return True
    End Function

    Public Function IsLeapYear(ByVal myYear As Integer) As Boolean
        If myYear / 4 = Math.Round(myYear / 4, 0) Then
            Return True
        Else
            Return False
        End If
    End Function


    Friend Shared Function ConvertoDateTime(ByVal dateString As String, ByRef result As DateTime,
                                            Optional ByRef format As String = "dd/MM/yyyy") As Boolean
        Try
            result = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function FileNameFromPath(ByVal path) As String
        Dim done As Boolean
        Dim mypos As Integer, prevpos As Integer
        mypos = 0
        prevpos = 0

        While Not done
            prevpos = mypos
            mypos = InStr(mypos + 1, path, "\")
            If mypos = 0 Then
                done = True
            End If
        End While

        Return Right(path, Len(path) - prevpos)

    End Function

    ' Paul Meems: deze functie komt twee keer voor in deze classe
    Private Function ParsestringDubbel(ByRef myString As String, Optional ByVal Delimiter As String = " ",
                                Optional ByVal QuoteHandlingFlag As Integer = 1) As String

        Dim quoteEven As Boolean
        Dim tmpString As String = "", tmpChar As String = ""

        quoteEven = True

        'Quotehandlingflag: default = 1
        '0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
        '1 = items between single quotes are being treated as separate items (no parsing between single quotes)
        '2 = items between double quotes are being treated as separate items (no parsing between double quotes)

        Dim i As Integer
        For i = 1 To Len(myString)
            'snoep een karakter af van de string
            tmpChar = Left(myString, 1)

            If (tmpChar = "'" And QuoteHandlingFlag = 1) Or (tmpChar = Chr(34) And QuoteHandlingFlag = 2) Then
                If quoteEven = True Then
                    quoteEven = False
                    tmpString = tmpString & tmpChar
                    myString = Right(myString, myString.Length - 1)
                Else
                    quoteEven = True 'dit betekent dat we klaar zijn
                    tmpString = Right(tmpString, tmpString.Length - 1) 'laat bij het teruggeven meteen de quotes maar weg!
                    myString = Right(myString, myString.Length - 1).Trim
                    Return tmpString
                End If
            ElseIf tmpChar = Delimiter And quoteEven = True Then
                If Not tmpString = "" Then
                    myString = Right(myString, myString.Length - 1)
                    Return tmpString
                Else
                    myString = Right(myString, myString.Length - 1)
                End If
            Else
                myString = Right(myString, myString.Length - 1)
                tmpString = tmpString & tmpChar
            End If
        Next i

        Return tmpString

    End Function

    Public Function BNAString(ByVal ID As String, ByVal Name As String, ByVal X As Double, ByVal Y As Double) As String
        Return Chr(34) & ID & Chr(34) & "," & Chr(34) & Name & Chr(34) & ",1," & X & "," & Y
    End Function

    Public Function LpSpHa2M3pS(ByVal LpSpHa As Double, ByVal AreaM2 As Double) As Double
        Dim AreaHa As Double = AreaM2 / 10000
        Return LpSpHa * AreaHa / 1000
    End Function

    Public Function MMPD2M3PS(ByVal MMpD As Double, ByVal Opp As Double) As Double
        'converteert milimeters per dag naar kuubs per seconde
        Return (Opp * MMpD / 1000) / (24 * 3600)
    End Function

    Public Function mmph2m3ps(ByVal mm As Double, ByVal opp As Double) As Double
        'converteert milimeters per uur naar kuubs per seconde
        Return (opp * mm / 1000) / (3600)
    End Function

    Public Function m3ps2mmps(ByVal m3ps As Double, ByVal opp As Double) As Double
        'converteert m3/s naar mm/s
        If opp > 0 Then
            Return 1000 * m3ps / opp
        Else
            Return 0
        End If
    End Function

    Public Function m3ps2mmpd(ByVal m3ps As Double, ByVal opp As Double) As Double
        'converteert m3/s naar mm/d
        If opp > 0 Then
            Return 1000 * 3600 * 24 * m3ps / opp
        Else
            Return 0
        End If
    End Function

    Friend Function m3ps2mmph(ByVal m3ps As Double, ByVal opp As Double) As Double
        'converteert m3/s naar mm/h
        If opp > 0 Then
            Return 1000 * 3600 * m3ps / opp
        Else
            Return 0
        End If
    End Function

    Friend Function ParseTable(ByRef myString As String) As String
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(myString, "TBLE", CompareMethod.Binary)
        endPos = InStr(myString, "tble", CompareMethod.Binary)
        Dim table As String = ""

        If startPos > 0 AndAlso endPos > 0 Then
            table = Mid(myString, startPos, endPos - startPos + 4)
            myString = Right(myString, myString.Length - endPos - 3)
        End If
        Return table

    End Function

    Friend Function AutoCorrectPath(ByRef myPath As String) As Boolean
        If Right(myPath, 1) = "\" Then myPath = Left(myPath, myPath.Length - 1)
        If Directory.Exists(myPath) Then
            Return True
        Else
            Return False
        End If
    End Function

    Friend Function RemoveSurroundingQuotes(ByVal myString As String, SingleQuotes As Boolean, DoubleQuotes As Boolean) As String
        Dim tmpStr As String, Done As Boolean
        tmpStr = Trim(myString)

        While Not Done
            Done = True
            If SingleQuotes AndAlso Left(tmpStr, 1) = "'" Then
                tmpStr = Right(tmpStr, tmpStr.Length - 1)
                Done = False
            End If
            If SingleQuotes AndAlso Right(tmpStr, 1) = "'" Then
                tmpStr = Left(tmpStr, tmpStr.Length - 1)
                Done = False
            End If
            If DoubleQuotes AndAlso Left(tmpStr, 1) = Chr(34) Then
                tmpStr = Right(tmpStr, tmpStr.Length - 1)
                Done = False
            End If
            If DoubleQuotes AndAlso Right(tmpStr, 1) = Chr(34) Then
                tmpStr = Left(tmpStr, tmpStr.Length - 1)
                Done = False
            End If
        End While

        Return tmpStr
    End Function


    Friend Function RotatePoint(ByVal Xold As Double, ByVal Yold As Double, ByVal Xorigin As Double, ByVal Yorigin As Double,
                              ByVal Degrees As Double, ByRef Xnew As Double, ByRef Ynew As Double) As Boolean
        Dim r As Double, theta As Double, dy As Double, dx As Double
        'roteert een punt ten opzichte van zijn oorsprong

        dy = (Yold - Yorigin)
        dx = (Xold - Xorigin)
        r = Math.Sqrt(dx ^ 2 + dy ^ 2)

        'If dx = 0 Then dx = 0.00000000000001
        'theta = Math.Atan(dy / dx)


        Dim curangle As Double = LineAngleDegrees(Xorigin, Yorigin, Xold, Yold)
        Dim newangle As Double = curangle + Degrees

        dx = Math.Sin(D2R(newangle)) * r
        dy = Math.Cos(D2R(newangle)) * r

        Xnew = Xorigin + dx
        Ynew = Yorigin + dy


        ''Xnew = r * Math.Cos(theta - D2R(Degrees)) + Xorigin
        ''Ynew = r * Math.Sin(theta - D2R(Degrees)) + Yorigin

        'Xnew = r * Math.Sin(theta + D2R(Degrees)) + Xorigin
        'Ynew = r * Math.Cos(theta + D2R(Degrees)) + Yorigin

        Return True
    End Function

    Friend Function D2R(ByVal Angle As Double) As Double
        'graden naar radialen
        D2R = Angle / 180 * Math.PI
    End Function

    Friend Function R2D(ByVal Angle As Double) As Double
        'radialen naar graden
        R2D = Angle * 180 / Math.PI
    End Function

    Public Function LineAngleDegrees(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
        'berekent de hoek van een lijn tussen twee xy co-ordinaten
        Dim dX As Double, dY As Double

        dX = Math.Abs(X2 - X1)
        dY = Math.Abs(Y2 - Y1)

        If dX = 0 Then
            If dY = 0 Then
                Return 0
            ElseIf Y2 > Y1 Then
                Return 0
            ElseIf Y2 < Y1 Then
                Return 180
            End If
        ElseIf dY = 0 Then
            If X2 > X1 Then
                Return 90
            ElseIf X2 < X1 Then
                Return 270
            End If
        Else
            If X2 > X1 And Y2 > Y1 Then 'eerste kwadrant
                Return R2D(Math.Atan(dX / dY))
            ElseIf X2 > X1 And Y2 < Y1 Then 'tweede kwadrant
                Return 90 + R2D(Math.Atan(dY / dX))
            ElseIf X2 < X1 And Y2 < Y1 Then 'derde kwadrant
                Return 180 + R2D(Math.Atan(dX / dY))
            Else 'vierde kwadrant
                Return 270 + R2D(Math.Atan(dY / dX))
            End If
        End If

    End Function

    Friend Function NormalizeAngle(ByVal myAngle As Double) As Double

        'Me.setup.Log.AddDebugMessage("In NormalizeAngle. Angle: " & myAngle)

        If (Math.Abs(myAngle) > 360 * 10) Then
            Throw New ArgumentException("The input angle is much too big", "myAngle")
        End If

        While myAngle > 360
            myAngle -= 360
        End While

        While myAngle < 0
            myAngle += 360
        End While

        Return myAngle

    End Function

    Public Function QWEIR(ByVal Width As Double, ByVal DischCoef As Double, ByVal H1 As Double, ByVal H2 As Double, ByVal Z As Double, Optional ByVal LatContrCoef As Double = 1) As Double
        Dim Hup As Double, Hdown As Double, Multiplier As Double

        'Free flow: als h2 - z < 2/3 * (h1 -z)
        If H1 >= H2 Then
            Hup = H1
            Hdown = H2
            Multiplier = 1
        Else
            Hup = H2
            Hdown = H1
            Multiplier = -1
        End If

        If Hup <= Z Then
            Return 0
        ElseIf Hdown < Z Or (Hdown - Z) < 2 / 3 * (Hup - Z) Then
            'Free flow: Q = c * B * 2/3 * SQRT(2/3 * g) * (h1 - z)^1.5
            Return Multiplier * DischCoef * LatContrCoef * Width * 2 / 3 * Math.Sqrt(2 / 3 * 9.81) * (Hup - Z) ^ 1.5
        Else
            'Drowned flow: Q = c * B * (h2 -z) * SQRT(2 * g *(h1 - h2))
            Return Multiplier * DischCoef * LatContrCoef * Width * (Hdown - Z) * Math.Sqrt(2 * 9.81 * (Hup - Hdown))
        End If

    End Function


    Public Function QORIFICE(ByVal Z As Double, ByVal w As Double, ByVal gh As Double, ByVal mu As Double, ByVal cw As Double, ByVal H1 As Double, ByVal H2 As Double) As Double
        'Z = crest level
        'W = width
        'gh = gate height (openningshoogte)
        'mu = contraction coef (standaard 0.63)
        'cw = lateral contraction coef
        'h1 = waterstand bovenstrooms
        'h2 = waterstand benedenstrooms
        'ce = afvoercoefficient. standaard 1.5

        Dim Af As Double
        Dim ce As Double
        Dim g As Double
        Dim u As Double 'stroomsnelheid over de kruin. Moet eigenlijk iteratief worden bepaald maar ik zet hem even op 1
        u = 1
        ce = 1.5
        g = 9.81

        'bepaal of hij verdronken of vrij is
        If (H1 - Z) >= (3 / 2 * gh) Then   'orifice flow
            If H2 <= (Z + gh) Then 'free orifice flow
                Af = w * mu * gh
                Return cw * w * mu * gh * Math.Sqrt(2 * g * (H1 - (Z + mu * gh)))
            ElseIf H2 > (Z + gh) Then 'submerged orifice flow
                Af = w * mu * gh
                Return cw * w * mu * gh * Math.Sqrt(2 * g * (H1 - H2))
            End If
        ElseIf (H1 - Z) < (3 / 2 * gh) Then 'weir flow
            If (H1 - Z) > (3 / 2 * (H2 - Z)) Then 'free weir flow
                Af = w * 2 / 3 * (H1 - Z)
                Return cw * w * 2 / 3 * Math.Sqrt(2 / 3 * g * (H1 - Z) ^ 3 / 2)
            ElseIf (H1 - Z) <= (3 / 2 * (H2 - Z)) Then 'submerged weir flow
                Af = w * (H1 - Z - u ^ 2 / (2 * g))
                Return ce * cw * w * (H1 - Z - (u ^ 2 / (2 * g))) * Math.Sqrt(2 * g * (H1 - H2))
            End If
        Else
            MsgBox("Error: kon niet bepalen of orifice verdronken of vrij was.")
        End If


    End Function

    Public Function calcOrificeWidth(ByVal Q As Double, ByVal mu As Double, ByVal cw As Double, ByVal Z As Double, ByVal H1 As Double, ByVal gh As Double) As Double
        'calculates the desired dimensions for an orifice given a discharge and dH under free flow conditions
        'NOTE: assumes free orifice flow condition, thus (h1 - z) > 3/2*gh and h2 < (z + gh)
        'QORIFICE = cw * w * mu * gh * Math.Sqrt(2 * g * (H1 - (Z + mu * gh)))
        Dim w As Double, g As Double = 9.81

        w = Q / (cw * mu * gh * Math.Sqrt(2 * g * (H1 - (Z + mu * gh))))
        Return w

    End Function

    Public Function calcWeirWidth(ByVal Q As Double, ByVal ce As Double, ByVal sc As Double, ByVal overstorthoogte As Double) As Double
        'calculates the desired dimensions for a weir given a discharge and (h1 - z) under free flow conditions
        'QWEIR = ce * sc * w * 2/3 * Math.Sqrt(2/3 * g) * (h1 - z)^(3/2)
        'ce = discharge coef
        'cw = lateral contraction coef
        Dim w As Double, g As Double = 9.81

        w = Q / (ce * sc * (2 / 3) * Math.Sqrt(2 / 3 * g) * (overstorthoogte) ^ (3 / 2))
        Return w

    End Function


    Friend Sub SortCollectionOfObjects(ByRef col As Collection, ByVal psSortPropertyName As String,
                                       ByVal pbAscending As Boolean, Optional ByVal psKeyPropertyName As String = "")

        ' This is a cool function. I found this on Freevbcode.com. It was a VB6 function so I had to VB.NET-ify it. 
        ' Which was pretty simple. It sorts a collection by a property Ascending or Decending
        ' The Objects were originally declared as Variants. VB.Net has eliminated the Variant type so they must be declared as type Object. 
        ' Also Objects cannot be used with the Set keyword, so I had to remove the set keyword. Other than that I did not have to make any changes.

        Dim obj As Object, i As Integer, j As Integer
        Dim iMinMaxIndex As Integer, vMinMax As Object, vValue As Object
        Dim bSortCondition As Boolean, bUseKey As Boolean, sKey As String

        'als de propertyname leeg is, gebruiken we geen key
        bUseKey = (psKeyPropertyName <> "")

        'doorloop de collection
        For i = 1 To col.Count - 1
            obj = col(i)
            vMinMax = CallByName(obj, psSortPropertyName, vbGet)
            iMinMaxIndex = i

            'doorloop de collection vanaf i tot het eind nogmaals
            For j = i + 1 To col.Count
                obj = col(j)
                vValue = CallByName(obj, psSortPropertyName, vbGet)

                If (pbAscending) Then
                    bSortCondition = (vValue < vMinMax)
                Else
                    bSortCondition = (vValue > vMinMax)
                End If

                If (bSortCondition) Then
                    vMinMax = vValue
                    iMinMaxIndex = j
                End If

                obj = Nothing
            Next j

            If (iMinMaxIndex <> i) Then
                obj = col(iMinMaxIndex)

                col.Remove(iMinMaxIndex)
                If (bUseKey) Then
                    sKey = CStr(CallByName(obj, psKeyPropertyName, vbGet))
                    col.Add(obj, sKey, i)
                Else
                    col.Add(obj, , i)
                End If

                obj = Nothing
            End If

            obj = Nothing
        Next i

    End Sub

    Friend Sub SortCollectionOfDouble(ByRef col As Collection, ByVal psSortPropertyName As String,
                                      ByVal pbAscending As Boolean, Optional ByVal psKeyPropertyName As String = "")

        ' This is a cool function. I found this on Freevbcode.com. It was a VB6 function so I had to VB.NET-ify it. 
        ' Which was pretty simple. It sorts a collection by a property Ascending or Decending
        ' The Objects were originally declared as Variants. VB.Net has eliminated the Variant type so they must be declared as type Object. 
        ' Also Objects cannot be used with the Set keyword, so I had to remove the set keyword. Other than that I did not have to make any changes.

        Dim dbl As Double, i As Integer, j As Integer
        Dim iMinMaxIndex As Integer, vMinMax As Object, vValue As Object
        Dim bSortCondition As Boolean, bUseKey As Boolean

        'als de propertyname leeg is, gebruiken we geen key
        bUseKey = (psKeyPropertyName <> "")

        'doorloop de collection
        For i = 1 To col.Count - 1
            vMinMax = col(i)
            iMinMaxIndex = i

            'doorloop de collection vanaf i tot het eind nogmaals
            For j = i + 1 To col.Count
                vValue = col(j)

                If (pbAscending) Then
                    bSortCondition = (vValue < vMinMax)
                Else
                    bSortCondition = (vValue > vMinMax)
                End If

                If (bSortCondition) Then
                    vMinMax = vValue
                    iMinMaxIndex = j
                End If

            Next j

            If (iMinMaxIndex <> i) Then
                dbl = col(iMinMaxIndex)
                col.Remove(iMinMaxIndex)
                If (bUseKey) Then
                    col.Add(dbl, Nothing, i)
                Else
                    col.Add(dbl, Nothing, i)
                End If

            End If

        Next i

    End Sub

    Friend Function FormatI10(ByVal myVal As Long) As String
        Dim myStr As String = myVal.ToString.Trim
        Dim i As Long

        If myStr.Length > 10 Then
            Return Format(myVal, "0E00")
        Else
            For i = myStr.Length + 1 To 10
                myStr = " " & myStr
            Next
        End If
        Return myStr
    End Function


    Public Sub UpdateProgressBarConsole(ByVal i As Long, ByVal n As Long)
        Dim Cur As Integer = RoundUD(100 * (i / n), 0, True)
        Dim Nxt As Integer = RoundUD(100 * (i + 1) / n, 0, True)

        'we'll write 50 times # in total. Therefore only write if the next number is even
        If Cur <> Nxt AndAlso IsEven(Nxt) Then
            Console.Write(Chr(35))
        End If
    End Sub

    Public Sub FlushMemory()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub



    Friend Function Lgn5ToSobek(ByVal LGNCODE As Integer) As Integer

        '1 = grass
        '2 = corn
        '3 = potatoes
        '4 = sugarbeet
        '5 = grain
        '6 = miscellaneous
        '7 = non-arable land
        '8 = greenhouse area
        '9 = orchard
        '10 = bulbous plants
        '11 = foliage forest
        '12 = pine forest
        '13 = nature
        '14 = fallow
        '15 = vegetables
        '16 = flowers

        'zelf toegevoegd:
        '17 = water
        '18 = verhard
        '19 = glastuinbouw

        Select Case LGNCODE
            Case Is = 0 'bestaat eigenlijk niet, dus maak er maar gras van
                Return 1
            Case Is = 1 'gras
                Return 1
            Case Is = 2 'mais
                Return 2
            Case Is = 3 'aardappelen
                Return 3
            Case Is = 4 'suikerbiet
                Return 4
            Case Is = 5 'graan
                Return 5
            Case Is = 6 'overige landbouwgewassen
                Return 6
            Case Is = 8 'glastuinbouw
                Return 19
            Case Is = 9 'boomgaard
                Return 9
            Case Is = 10 'bollenteelt
                Return 10
            Case Is = 11 'loofbos
                Return 11
            Case Is = 12 'naaldbos
                Return 12
            Case Is = 16 'zoet water
                Return 17
            Case Is = 17 'zout water
                Return 17
            Case Is = 18 'stedelijk bebouwd
                Return 18
            Case Is = 19 'bebouwd buitengebied
                Return 18
            Case Is = 20 'loofbos in bebouwd gebied
                Return 1
            Case Is = 21 'naaldbos in bebouwd gebied
                Return 1
            Case Is = 22 'bos met dichte bebouwing
                Return 18
            Case Is = 23 'gras in bebouwd gebied
                Return 1
            Case Is = 24 'kale grond in bebouwd buitengebied
                Return 1
            Case Is = 25 'hoofdwegen en spoorwegen
                Return 18
            Case Is = 26 'bebouwing in agrarisch gebied
                Return 18
            Case Is = 30 'kwelders
                Return 13
            Case Is = 35 'open stuifzand
                Return 13
            Case Is = 36 'heide
                Return 13
            Case Is = 37 'matig vergraste heide
                Return 13
            Case Is = 38 'sterk vergraste heide
                Return 13
            Case Is = 39 'hoogveen
                Return 13
            Case Is = 40 'bos in hoogveen
                Return 13
            Case Is = 41 'overige moerasvegetatie
                Return 13
            Case Is = 42 'rietvegetatie
                Return 13
            Case Is = 43 'bos in moerasgebied
                Return 13
            Case Is = 45 'overig open begroeid natuurgebied
                Return 13
            Case Is = 46 'kale grond in natuurgebied
                Return 13
            Case Else
                Return 1
        End Select

    End Function

    Friend Function Lgn6ToSobek(ByVal LGNCODE As Integer) As Integer

        'legenda: http://www.wageningenur.nl/nl/Expertises-Dienstverlening/Onderzoeksinstituten/Alterra/Faciliteiten-Producten/Kaarten-en-GISbestanden/LGN-1/Bestanden/LGN6/LGN6-legenda.htm

        'Landgebruikstypen in SOBEK:
        '1 = grass
        '2 = corn
        '3 = potatoes
        '4 = sugarbeet
        '5 = grain
        '6 = miscellaneous
        '7 = non-arable land
        '8 = greenhouse area
        '9 = orchard
        '10 = bulbous plants
        '11 = foliage forest
        '12 = pine forest
        '13 = nature
        '14 = fallow
        '15 = vegetables
        '16 = flowers

        'zelf toegevoegd:
        '17 = water
        '18 = verhard
        '19 = glastuinbouw

        Select Case LGNCODE
            Case Is = 0 'bestaat eigenlijk niet, dus maak er maar gras van
                Return 1
            Case Is = 1 'agrarisch gras
                Return 1
            Case Is = 2 'mais
                Return 2
            Case Is = 3 'aardappelen
                Return 3
            Case Is = 4 'bieten
                Return 4
            Case Is = 5 'graan
                Return 5
            Case Is = 6 'overige landbouwgewassen
                Return 6
            Case Is = 8 'glastuinbouw
                Return 19
            Case Is = 9 'boomgaarden
                Return 9
            Case Is = 10 'bollenteelt
                Return 10
            Case Is = 11 'loofbos
                Return 11
            Case Is = 12 'naaldbos
                Return 12
            Case Is = 16 'zoet water
                Return 17
            Case Is = 17 'zout water
                Return 17
            Case Is = 18 'Bebouwing in primair bebouwd gebied
                Return 18
            Case Is = 19 'Bebouwing in secundair bebouwd gebied
                Return 18
            Case Is = 20 'Bos in primair bebouwd gebied
                Return 11
            Case Is = 22 'Bos in secundair bebouwd gebied
                Return 11
            Case Is = 23 'gras in primair bebouwd gebied
                Return 1
            Case Is = 24 'kale grond in bebouwd buitengebied
                Return 7
            Case Is = 25 'hoofdwegen en spoorwegen
                Return 18
            Case Is = 26 'Bebouwing in buitengebied
                Return 18
            Case Is = 28 'Gras in secundair bebouwd gebied
                Return 1
            Case Is = 30 'kwelders
                Return 13
            Case Is = 31 'Open zand in kustgebied
                Return 7
            Case Is = 32 'Duinen met een lage vegetatie
                Return 13
            Case Is = 32 'Duinen met een hoge vegetatie
                Return 13
            Case Is = 32 'Duinheid
                Return 13
            Case Is = 35 'open stuifzand en/of rivierzand
                Return 7
            Case Is = 36 'heide
                Return 13
            Case Is = 37 'matig vergraste heide
                Return 13
            Case Is = 38 'sterk vergraste heide
                Return 13
            Case Is = 39 'hoogveen
                Return 13
            Case Is = 40 'bos in hoogveengebied
                Return 13
            Case Is = 41 'overige moerasvegetatie
                Return 13
            Case Is = 42 'rietvegetatie
                Return 13
            Case Is = 43 'bos in moerasgebied
                Return 13
            Case Is = 45 'Natuurgraslanden
                Return 13
            Case Else
                Return 1
        End Select

    End Function


    Public Function LGN6TONBW(ByVal LGNCODE As Integer) As Integer

        'legenda: http://www.wageningenur.nl/nl/Expertises-Dienstverlening/Onderzoeksinstituten/Alterra/Faciliteiten-Producten/Kaarten-en-GISbestanden/LGN-1/Bestanden/LGN6/LGN6-legenda.htm

        Select Case LGNCODE
            Case Is = 1 'agrarisch gras
                Return 3
            Case Is = 2, 3, 4, 5, 6 'mais, aardappelen, bieten, graan, overige landbouwgewassen
                Return 1
            Case Is = 8, 9, 10 'glastuinbouw, boomgaarden, bollenteelt
                Return 2
            Case Is = 11, 12 'loofbos, naaldbos
                Return 4
            Case Is = 16, 17 'zoet water, zout water
                Return 0
            Case Is = 18, 19 'Bebouwing in primair bebouwd gebied, bebouwing in secundair bebouwd gebied
                Return 5
            Case Is = 20, 22 'Bos in primair bebouwd gebied, bos in secundair bebouwd gebied
                Return 4
            Case Is = 23 'gras in primair bebouwd gebied
                Return 3
            Case Is = 24, 25, 26 'kale grond in bebouwd buitengebied, hoofdwegen en spoorwegen, bebouwing in buitengebied
                Return 5
            Case Is = 28 'Gras in secundair bebouwd gebied
                Return 3
            Case Is = 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 41, 41, 42, 43, 44, 45 'kwelders, open zand in kustgebied, duinen met lage vegetatie, duinen met hoge vegetatie, duinheid, open stuifzand, heide, matig vergraste heide, sterk vergraste heide, hoogveen, bos in hoogveengebied, overige moerasvegetatie, rietvegetatie, bos in moerasgebied, natuurgraslanden
                Return 4
            Case Else
                Return 3
        End Select

    End Function

    Friend Function Bod2Capsim(ByVal EERSTE_BOD As String) As Integer

        'converteert bodemtypes uit de Bodemkaart Nederland naar het corresponderende CAPSIM bodemnummer in SOBEK
        'Veengronden: code V
        'Moerige gronden: code W
        'Podzolgronden: code H en Y
        'BrikGronden: code B
        'Dikke eerdgronden: code EZ EL en EK
        'Kalkloze zandgronden: code Z
        'Kalkhoudende zandgronden: code Z...A
        'Kalkhoudende bijzonder lutumarme gronden: code S...A
        'Niet gerijpte minerale gronden: code MO-zeeklei, RO-rivierklei
        'Zeekleigronden: code M
        'Rivierkleigronden: code R
        'Oude rivierkleigronden: code KR
        'Leemgronden: code L
        'Mariene afzettingen ouder dan Pleistoceen: code MA, MK, MZ
        'Fluviatiele afzttingen ouder dan Pleistoceen: code FG, FK
        'Kalksteenverweringsgronden: code KM, KK, KS
        'Ondiepe keileemgronden: code KX
        'Overige oude kleigronden: code KT
        'Grindgronden: code G

        If EERSTE_BOD = "|a GROEVE" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|b AFGRAV" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|c OPHOOG" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|d EGAL" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|e VERWERK" Then Return 113 'willekeurig bedenksel
        If EERSTE_BOD = "|f TERP" Then Return 116 'lichte klei, klopt wel enigszins met omgeving
        If EERSTE_BOD = "|g MOERAS" Then Return 101 'veengrond ligt voor de hand
        If EERSTE_BOD = "|g WATER" Then Return 0 'tsja
        If EERSTE_BOD = "|h BEBOUW" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|h DIJK" Then Return 119 'klei op zand aangenomen
        If EERSTE_BOD = "|i BOVLAND" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "|j MYNSTRT" Then Return 108 'podzol=leemarm grof zand
        If EERSTE_BOD = "AAKp" Then Return 119
        If EERSTE_BOD = "AAP" Then Return 105
        If EERSTE_BOD = "ABk" Then Return 119
        If EERSTE_BOD = "ABkt" Then Return 119
        If EERSTE_BOD = "ABl" Then Return 121
        If EERSTE_BOD = "ABv" Then Return 105
        If EERSTE_BOD = "ABvg" Then Return 105
        If EERSTE_BOD = "ABvt" Then Return 105
        If EERSTE_BOD = "ABvx" Then Return 105
        If EERSTE_BOD = "ABz" Then Return 113
        If EERSTE_BOD = "ABzt" Then Return 113
        If EERSTE_BOD = "AEk9" Then Return 116
        If EERSTE_BOD = "AEm5" Then Return 115
        If EERSTE_BOD = "AEm8" Then Return 116
        If EERSTE_BOD = "AEm9A" Then Return 116
        If EERSTE_BOD = "AEp6A" Then Return 116
        If EERSTE_BOD = "AEp7A" Then Return 116
        If EERSTE_BOD = "AFz" Then Return 113
        If EERSTE_BOD = "Aha" Then Return 121
        If EERSTE_BOD = "AHc" Then Return 121
        If EERSTE_BOD = "AHk" Then Return 121
        If EERSTE_BOD = "AHl" Then Return 121
        If EERSTE_BOD = "Ahs" Then Return 121
        If EERSTE_BOD = "AHt" Then Return 121
        If EERSTE_BOD = "AHv" Then Return 121
        If EERSTE_BOD = "AHz" Then Return 121
        If EERSTE_BOD = "AK" Then Return 119
        If EERSTE_BOD = "AKp" Then Return 119
        If EERSTE_BOD = "ALu" Then Return 116
        If EERSTE_BOD = "AM" Then Return 119
        If EERSTE_BOD = "AMm" Then Return 115
        If EERSTE_BOD = "AO" Then Return 119
        If EERSTE_BOD = "AOg" Then Return 119
        If EERSTE_BOD = "AOp" Then Return 119
        If EERSTE_BOD = "AOv" Then Return 119
        If EERSTE_BOD = "AP" Then Return 101
        If EERSTE_BOD = "App" Then Return 102
        If EERSTE_BOD = "AQ" Then Return 107
        If EERSTE_BOD = "AR" Then Return 119
        If EERSTE_BOD = "AS" Then Return 107
        If EERSTE_BOD = "aVc" Then Return 101
        If EERSTE_BOD = "AVk" Then Return 105
        If EERSTE_BOD = "AVo" Then Return 101
        If EERSTE_BOD = "aVp" Then Return 102
        If EERSTE_BOD = "aVpg" Then Return 102
        If EERSTE_BOD = "aVpx" Then Return 102
        If EERSTE_BOD = "aVs" Then Return 101
        If EERSTE_BOD = "aVz" Then Return 102
        If EERSTE_BOD = "aVzt" Then Return 102
        If EERSTE_BOD = "aVzx" Then Return 102
        If EERSTE_BOD = "AWg" Then Return 116
        If EERSTE_BOD = "AWo" Then Return 106
        If EERSTE_BOD = "AWv" Then Return 106
        If EERSTE_BOD = "AZ1" Then Return 114
        If EERSTE_BOD = "AZW0A" Then Return 107
        If EERSTE_BOD = "AZW0Al" Then Return 107
        If EERSTE_BOD = "AZW0Av" Then Return 107
        If EERSTE_BOD = "AZW1A" Then Return 119
        If EERSTE_BOD = "AZW1Ar" Then Return 119
        If EERSTE_BOD = "AZW1Aw" Then Return 119
        If EERSTE_BOD = "AZW5A" Then Return 119
        If EERSTE_BOD = "AZW6A" Then Return 119
        If EERSTE_BOD = "AZW6Al" Then Return 116
        If EERSTE_BOD = "AZW6Alv" Then Return 118
        If EERSTE_BOD = "AZW7Al" Then Return 116
        If EERSTE_BOD = "AZW7Alw" Then Return 116
        If EERSTE_BOD = "AZW7Alwp" Then Return 119
        If EERSTE_BOD = "AZW8A" Then Return 116
        If EERSTE_BOD = "AZW8Al" Then Return 116
        If EERSTE_BOD = "AZW8Alw" Then Return 116
        If EERSTE_BOD = "bEZ21" Then Return 112
        If EERSTE_BOD = "bEZ21g" Then Return 112
        If EERSTE_BOD = "bEZ21x" Then Return 112
        If EERSTE_BOD = "bEZ23" Then Return 112
        If EERSTE_BOD = "bEZ23g" Then Return 112
        If EERSTE_BOD = "bEZ23t" Then Return 112
        If EERSTE_BOD = "bEZ23x" Then Return 112
        If EERSTE_BOD = "bEZ30" Then Return 112
        If EERSTE_BOD = "bEZ30x" Then Return 112
        If EERSTE_BOD = "bgMn15C" Then Return 115
        If EERSTE_BOD = "bgMn25C" Then Return 115
        If EERSTE_BOD = "bgMn53C" Then Return 117
        If EERSTE_BOD = "BKd25" Then Return 115
        If EERSTE_BOD = "BKd25x" Then Return 115
        If EERSTE_BOD = "BKd26" Then Return 115
        If EERSTE_BOD = "BKh25" Then Return 115
        If EERSTE_BOD = "BKh25x" Then Return 115
        If EERSTE_BOD = "BKh26" Then Return 115
        If EERSTE_BOD = "BKh26x" Then Return 115
        If EERSTE_BOD = "BLb6" Then Return 121
        If EERSTE_BOD = "BLb6g" Then Return 121
        If EERSTE_BOD = "BLb6k" Then Return 121
        If EERSTE_BOD = "BLb6s" Then Return 121
        If EERSTE_BOD = "BLd5" Then Return 121
        If EERSTE_BOD = "BLd5g" Then Return 121
        If EERSTE_BOD = "BLd5t" Then Return 121
        If EERSTE_BOD = "BLd6" Then Return 121
        If EERSTE_BOD = "BLd6m" Then Return 121
        If EERSTE_BOD = "BLh5m" Then Return 121
        If EERSTE_BOD = "BLh6" Then Return 121
        If EERSTE_BOD = "BLh6g" Then Return 121
        If EERSTE_BOD = "BLh6m" Then Return 121
        If EERSTE_BOD = "BLh6s" Then Return 121
        If EERSTE_BOD = "BLn5m" Then Return 121
        If EERSTE_BOD = "BLn5t" Then Return 121
        If EERSTE_BOD = "BLn6" Then Return 121
        If EERSTE_BOD = "BLn6g" Then Return 121
        If EERSTE_BOD = "BLn6m" Then Return 121
        If EERSTE_BOD = "BLn6s" Then Return 121
        If EERSTE_BOD = "bMn15A" Then Return 115
        If EERSTE_BOD = "bMn15C" Then Return 115
        If EERSTE_BOD = "bMn25A" Then Return 115
        If EERSTE_BOD = "bMn25C" Then Return 115
        If EERSTE_BOD = "bMn35A" Then Return 116
        If EERSTE_BOD = "bMn45A" Then Return 117
        If EERSTE_BOD = "bMn56Cp" Then Return 119
        If EERSTE_BOD = "bMn85C" Then Return 116
        If EERSTE_BOD = "bMn86C" Then Return 117
        If EERSTE_BOD = "bRn46C" Then Return 117
        If EERSTE_BOD = "BZd23" Then Return 113
        If EERSTE_BOD = "BZd24" Then Return 113
        If EERSTE_BOD = "cHd21" Then Return 108
        If EERSTE_BOD = "cHd21g" Then Return 110
        If EERSTE_BOD = "cHd21x" Then Return 111
        If EERSTE_BOD = "cHd23" Then Return 113
        If EERSTE_BOD = "cHd23x" Then Return 111
        If EERSTE_BOD = "cHd30" Then Return 114
        If EERSTE_BOD = "cHn21" Then Return 109
        If EERSTE_BOD = "cHn21g" Then Return 110
        If EERSTE_BOD = "cHn21t" Then Return 111
        If EERSTE_BOD = "cHn21w" Then Return 111
        If EERSTE_BOD = "cHn21x" Then Return 111
        If EERSTE_BOD = "cHn23" Then Return 113
        If EERSTE_BOD = "cHn23g" Then Return 110
        If EERSTE_BOD = "cHn23t" Then Return 111
        If EERSTE_BOD = "cHn23wx" Then Return 111
        If EERSTE_BOD = "cHn23x" Then Return 111
        If EERSTE_BOD = "cHn30" Then Return 114
        If EERSTE_BOD = "cHn30g" Then Return 114
        If EERSTE_BOD = "cY21" Then Return 109
        If EERSTE_BOD = "cY21g" Then Return 110
        If EERSTE_BOD = "cY21x" Then Return 111
        If EERSTE_BOD = "cY23" Then Return 113
        If EERSTE_BOD = "cY23g" Then Return 113
        If EERSTE_BOD = "cY23x" Then Return 111
        If EERSTE_BOD = "cY30" Then Return 114
        If EERSTE_BOD = "cY30g" Then Return 114
        If EERSTE_BOD = "cZd21" Then Return 108
        If EERSTE_BOD = "cZd21g" Then Return 110
        If EERSTE_BOD = "cZd23" Then Return 113
        If EERSTE_BOD = "cZd30" Then Return 114
        If EERSTE_BOD = "dgMn58Cv" Then Return 117
        If EERSTE_BOD = "dgMn83C" Then Return 117
        If EERSTE_BOD = "dgMn88Cv" Then Return 117
        If EERSTE_BOD = "dhVb" Then Return 101
        If EERSTE_BOD = "dhVk" Then Return 106
        If EERSTE_BOD = "dhVr" Then Return 101
        If EERSTE_BOD = "dkVc" Then Return 103
        If EERSTE_BOD = "dMn86C" Then Return 117
        If EERSTE_BOD = "dMv41C" Then Return 118
        If EERSTE_BOD = "dMv61C" Then Return 118
        If EERSTE_BOD = "dpVc" Then Return 103
        If EERSTE_BOD = "dVc" Then Return 101
        If EERSTE_BOD = "dVd" Then Return 101
        If EERSTE_BOD = "dVk" Then Return 106
        If EERSTE_BOD = "dVr" Then Return 101
        If EERSTE_BOD = "dWo" Then Return 106
        If EERSTE_BOD = "dWol" Then Return 106
        If EERSTE_BOD = "EK19" Then Return 115
        If EERSTE_BOD = "EK19p" Then Return 119
        If EERSTE_BOD = "EK19x" Then Return 115
        If EERSTE_BOD = "EK76" Then Return 117
        If EERSTE_BOD = "EK79" Then Return 116
        If EERSTE_BOD = "EK79v" Then Return 116
        If EERSTE_BOD = "EK79w" Then Return 116
        If EERSTE_BOD = "EL5" Then Return 121
        If EERSTE_BOD = "eMn12Ap" Then Return 119
        If EERSTE_BOD = "eMn15A" Then Return 115
        If EERSTE_BOD = "eMn15Ap" Then Return 119
        If EERSTE_BOD = "eMn22A" Then Return 119
        If EERSTE_BOD = "eMn22Ap" Then Return 119
        If EERSTE_BOD = "eMn25A" Then Return 115
        If EERSTE_BOD = "eMn25Ap" Then Return 119
        If EERSTE_BOD = "eMn25Av" Then Return 118
        If EERSTE_BOD = "eMn35A" Then Return 116
        If EERSTE_BOD = "eMn35Ap" Then Return 119
        If EERSTE_BOD = "eMn35Av" Then Return 118
        If EERSTE_BOD = "eMn35Awp" Then Return 119
        If EERSTE_BOD = "eMn45A" Then Return 117
        If EERSTE_BOD = "eMn45Ap" Then Return 117
        If EERSTE_BOD = "eMn45Av" Then Return 118
        If EERSTE_BOD = "eMn52Cg" Then Return 119
        If EERSTE_BOD = "eMn52Cp" Then Return 119
        If EERSTE_BOD = "eMn52Cwp" Then Return 119
        If EERSTE_BOD = "eMn56Av" Then Return 118
        If EERSTE_BOD = "eMn82A" Then Return 119
        If EERSTE_BOD = "eMn82Ap" Then Return 119
        If EERSTE_BOD = "eMn82C" Then Return 119
        If EERSTE_BOD = "eMn82Cp" Then Return 119
        If EERSTE_BOD = "eMn86A" Then Return 117
        If EERSTE_BOD = "eMn86Av" Then Return 118
        If EERSTE_BOD = "eMn86C" Then Return 117
        If EERSTE_BOD = "eMn86Cv" Then Return 118
        If EERSTE_BOD = "eMn86Cw" Then Return 117
        If EERSTE_BOD = "eMo20A" Then Return 119
        If EERSTE_BOD = "eMo20Ap" Then Return 119
        If EERSTE_BOD = "eMo80A" Then Return 116
        If EERSTE_BOD = "eMo80Ap" Then Return 119
        If EERSTE_BOD = "eMo80C" Then Return 116
        If EERSTE_BOD = "eMo80Cv" Then Return 118
        If EERSTE_BOD = "eMOb72" Then Return 119
        If EERSTE_BOD = "eMOb75" Then Return 116
        If EERSTE_BOD = "eMOo05" Then Return 115
        If EERSTE_BOD = "eMv41C" Then Return 118
        If EERSTE_BOD = "eMv51A" Then Return 118
        If EERSTE_BOD = "eMv61C" Then Return 118
        If EERSTE_BOD = "eMv61Cp" Then Return 118
        If EERSTE_BOD = "eMv81A" Then Return 118
        If EERSTE_BOD = "eMv81Ap" Then Return 118
        If EERSTE_BOD = "epMn55A" Then Return 115
        If EERSTE_BOD = "epMn85A" Then Return 116
        If EERSTE_BOD = "epMo50" Then Return 115
        If EERSTE_BOD = "epMo80" Then Return 116
        If EERSTE_BOD = "epMv81" Then Return 118
        If EERSTE_BOD = "epRn56" Then Return 117
        If EERSTE_BOD = "epRn59" Then Return 119
        If EERSTE_BOD = "epRn86" Then Return 117
        If EERSTE_BOD = "eRn45A" Then Return 117
        If EERSTE_BOD = "eRn46A" Then Return 117
        If EERSTE_BOD = "eRn46Av" Then Return 118
        If EERSTE_BOD = "eRn47C" Then Return 117
        If EERSTE_BOD = "eRn52A" Then Return 119
        If EERSTE_BOD = "eRn66A" Then Return 117
        If EERSTE_BOD = "eRn66Av" Then Return 118
        If EERSTE_BOD = "eRn82A" Then Return 119
        If EERSTE_BOD = "eRn94C" Then Return 117
        If EERSTE_BOD = "eRn95A" Then Return 116
        If EERSTE_BOD = "eRn95Av" Then Return 118
        If EERSTE_BOD = "eRo40A" Then Return 117
        If EERSTE_BOD = "eRv01A" Then Return 118
        If EERSTE_BOD = "eRv01C" Then Return 118
        If EERSTE_BOD = "EZ50A" Then Return 107
        If EERSTE_BOD = "EZ50Av" Then Return 107
        If EERSTE_BOD = "EZg21" Then Return 112
        If EERSTE_BOD = "EZg21g" Then Return 112
        If EERSTE_BOD = "EZg21v" Then Return 112
        If EERSTE_BOD = "EZg21w" Then Return 112
        If EERSTE_BOD = "EZg23" Then Return 112
        If EERSTE_BOD = "EZg23g" Then Return 112
        If EERSTE_BOD = "EZg23t" Then Return 112
        If EERSTE_BOD = "EZg23tw" Then Return 112
        If EERSTE_BOD = "EZg23w" Then Return 112
        If EERSTE_BOD = "EZg23wg" Then Return 112
        If EERSTE_BOD = "EZg23wt" Then Return 112
        If EERSTE_BOD = "EZg30" Then Return 112
        If EERSTE_BOD = "EZg30g" Then Return 112
        If EERSTE_BOD = "EZg30v" Then Return 112
        If EERSTE_BOD = "fABk" Then Return 119
        If EERSTE_BOD = "fAFk" Then Return 119
        If EERSTE_BOD = "fAFz" Then Return 113
        If EERSTE_BOD = "faVc" Then Return 101
        If EERSTE_BOD = "faVz" Then Return 102
        If EERSTE_BOD = "faVzt" Then Return 102
        If EERSTE_BOD = "FG" Then Return 114
        If EERSTE_BOD = "fHn21" Then Return 109
        If EERSTE_BOD = "fhVc" Then Return 101
        If EERSTE_BOD = "fhVd" Then Return 101
        If EERSTE_BOD = "fhVz" Then Return 102
        If EERSTE_BOD = "fiVc" Then Return 105
        If EERSTE_BOD = "fiVz" Then Return 105
        If EERSTE_BOD = "fiWp" Then Return 105
        If EERSTE_BOD = "fiWz" Then Return 105
        If EERSTE_BOD = "FKk" Then Return 121
        If EERSTE_BOD = "fkpZg23" Then Return 119
        If EERSTE_BOD = "fkpZg23g" Then Return 120
        If EERSTE_BOD = "fkpZg23t" Then Return 119
        If EERSTE_BOD = "fKRn1" Then Return 119
        If EERSTE_BOD = "fKRn1g" Then Return 120
        If EERSTE_BOD = "fKRn2g" Then Return 120
        If EERSTE_BOD = "fKRn8" Then Return 119
        If EERSTE_BOD = "fKRn8g" Then Return 120
        If EERSTE_BOD = "fkVc" Then Return 103
        If EERSTE_BOD = "fkVs" Then Return 103
        If EERSTE_BOD = "fkVz" Then Return 104
        If EERSTE_BOD = "fkWz" Then Return 104
        If EERSTE_BOD = "fkWzg" Then Return 104
        If EERSTE_BOD = "fkZn21" Then Return 119
        If EERSTE_BOD = "fkZn23" Then Return 119
        If EERSTE_BOD = "fkZn23g" Then Return 120
        If EERSTE_BOD = "fkZn30" Then Return 120
        If EERSTE_BOD = "fMn56Cp" Then Return 119
        If EERSTE_BOD = "fMn56Cv" Then Return 118
        If EERSTE_BOD = "fpLn5" Then Return 121
        If EERSTE_BOD = "fpRn59" Then Return 119
        If EERSTE_BOD = "fpRn86" Then Return 117
        If EERSTE_BOD = "fpVc" Then Return 103
        If EERSTE_BOD = "fpVs" Then Return 103
        If EERSTE_BOD = "fpVz" Then Return 104
        If EERSTE_BOD = "fpZg21" Then Return 109
        If EERSTE_BOD = "fpZg21g" Then Return 110
        If EERSTE_BOD = "fpZg23" Then Return 113
        If EERSTE_BOD = "fpZg23g" Then Return 113
        If EERSTE_BOD = "fpZg23t" Then Return 111
        If EERSTE_BOD = "fpZg23x" Then Return 111
        If EERSTE_BOD = "fpZn21" Then Return 109
        If EERSTE_BOD = "fpZn23tg" Then Return 111
        If EERSTE_BOD = "fRn15C" Then Return 115
        If EERSTE_BOD = "fRn62C" Then Return 119
        If EERSTE_BOD = "fRn62Cg" Then Return 120
        If EERSTE_BOD = "fRn95C" Then Return 116
        If EERSTE_BOD = "fRo60C" Then Return 116
        If EERSTE_BOD = "fRv01C" Then Return 118
        If EERSTE_BOD = "fVc" Then Return 101
        If EERSTE_BOD = "fvWz" Then Return 102
        If EERSTE_BOD = "fvWzt" Then Return 102
        If EERSTE_BOD = "fvWztx" Then Return 102
        If EERSTE_BOD = "fVz" Then Return 102
        If EERSTE_BOD = "fZn21" Then Return 107
        If EERSTE_BOD = "fZn21g" Then Return 107
        If EERSTE_BOD = "fZn23" Then Return 113
        If EERSTE_BOD = "fZn23-F" Then Return 113
        If EERSTE_BOD = "fZn23g" Then Return 113
        If EERSTE_BOD = "fzVc" Then Return 105
        If EERSTE_BOD = "fzVz" Then Return 105
        If EERSTE_BOD = "fzVzt" Then Return 105
        If EERSTE_BOD = "fzWp" Then Return 105
        If EERSTE_BOD = "fzWz" Then Return 105
        If EERSTE_BOD = "fzWzt" Then Return 105
        If EERSTE_BOD = "gbEZ21" Then Return 112
        If EERSTE_BOD = "gbEZ30" Then Return 112
        If EERSTE_BOD = "gcHd30" Then Return 114
        If EERSTE_BOD = "gcHn21" Then Return 109
        If EERSTE_BOD = "gcHn30" Then Return 114
        If EERSTE_BOD = "gcY21" Then Return 109
        If EERSTE_BOD = "gcY23" Then Return 113
        If EERSTE_BOD = "gcY30" Then Return 114
        If EERSTE_BOD = "gcZd30" Then Return 114
        If EERSTE_BOD = "gHd21" Then Return 108
        If EERSTE_BOD = "gHd30" Then Return 114
        If EERSTE_BOD = "gHn21" Then Return 109
        If EERSTE_BOD = "gHn21t" Then Return 111
        If EERSTE_BOD = "gHn21x" Then Return 111
        If EERSTE_BOD = "gHn23" Then Return 113
        If EERSTE_BOD = "gHn23x" Then Return 111
        If EERSTE_BOD = "gHn30" Then Return 114
        If EERSTE_BOD = "gHn30t" Then Return 114
        If EERSTE_BOD = "gHn30x" Then Return 114
        If EERSTE_BOD = "gKRd1" Then Return 119
        If EERSTE_BOD = "gKRd7" Then Return 119
        If EERSTE_BOD = "gKRn1" Then Return 119
        If EERSTE_BOD = "gKRn2" Then Return 119
        If EERSTE_BOD = "gLd6" Then Return 121
        If EERSTE_BOD = "gLh6" Then Return 121
        If EERSTE_BOD = "gMK" Then Return 115
        If EERSTE_BOD = "gMn15C" Then Return 115
        If EERSTE_BOD = "gMn25C" Then Return 115
        If EERSTE_BOD = "gMn25Cv" Then Return 115
        If EERSTE_BOD = "gMn52C" Then Return 119
        If EERSTE_BOD = "gMn52Cp" Then Return 119
        If EERSTE_BOD = "gMn52Cw" Then Return 119
        If EERSTE_BOD = "gMn53C" Then Return 117
        If EERSTE_BOD = "gMn53Cp" Then Return 119
        If EERSTE_BOD = "gMn53Cpx" Then Return 119
        If EERSTE_BOD = "gMn53Cv" Then Return 118
        If EERSTE_BOD = "gMn53Cw" Then Return 117
        If EERSTE_BOD = "gMn53Cwp" Then Return 119
        If EERSTE_BOD = "gMn58C" Then Return 117
        If EERSTE_BOD = "gMn58Cv" Then Return 117
        If EERSTE_BOD = "nkZn50A" Then Return 119
        If EERSTE_BOD = "gMn82C" Then Return 119
        If EERSTE_BOD = "gMn83C" Then Return 117
        If EERSTE_BOD = "gMn83Cp" Then Return 117
        If EERSTE_BOD = "gMn83Cv" Then Return 118
        If EERSTE_BOD = "gMn83Cw" Then Return 117
        If EERSTE_BOD = "gMn83Cwp" Then Return 117
        If EERSTE_BOD = "gMn85C" Then Return 116
        If EERSTE_BOD = "gMn85Cv" Then Return 118
        If EERSTE_BOD = "gMn85Cwl" Then Return 116
        If EERSTE_BOD = "gMn88C" Then Return 117
        If EERSTE_BOD = "gMn88Cl" Then Return 117
        If EERSTE_BOD = "gMn88Clv" Then Return 118
        If EERSTE_BOD = "gMn88Cv" Then Return 118
        If EERSTE_BOD = "gMn88Cw" Then Return 117
        If EERSTE_BOD = "gpZg23x" Then Return 111
        If EERSTE_BOD = "gpZg30" Then Return 114
        If EERSTE_BOD = "gpZn21" Then Return 109
        If EERSTE_BOD = "gpZn21x" Then Return 111
        If EERSTE_BOD = "gpZn23x" Then Return 111
        If EERSTE_BOD = "gpZn30" Then Return 114
        If EERSTE_BOD = "gRd10A" Then Return 119
        If EERSTE_BOD = "gRn15A" Then Return 119
        If EERSTE_BOD = "gRn94Cv" Then Return 117
        If EERSTE_BOD = "gtZd30" Then Return 114
        If EERSTE_BOD = "gvWp" Then Return 102
        If EERSTE_BOD = "gY21" Then Return 109
        If EERSTE_BOD = "gY21g" Then Return 109
        If EERSTE_BOD = "gY23" Then Return 113
        If EERSTE_BOD = "gY30" Then Return 114
        If EERSTE_BOD = "gY30-F" Then Return 114
        If EERSTE_BOD = "gY30-G" Then Return 114
        If EERSTE_BOD = "gZb30" Then Return 114
        If EERSTE_BOD = "gZd21" Then Return 107
        If EERSTE_BOD = "gZd30" Then Return 114
        If EERSTE_BOD = "gzEZ21" Then Return 112
        If EERSTE_BOD = "gzEZ23" Then Return 112
        If EERSTE_BOD = "gzEZ30" Then Return 112
        If EERSTE_BOD = "gZn30" Then Return 114
        If EERSTE_BOD = "Hd21" Then Return 108
        If EERSTE_BOD = "Hd21g" Then Return 108
        If EERSTE_BOD = "Hd21x" Then Return 108
        If EERSTE_BOD = "Hd23" Then Return 113
        If EERSTE_BOD = "Hd23g" Then Return 110
        If EERSTE_BOD = "Hd23x" Then Return 111
        If EERSTE_BOD = "Hd30" Then Return 114
        If EERSTE_BOD = "Hd30g" Then Return 114
        If EERSTE_BOD = "hEV" Then Return 101
        If EERSTE_BOD = "Hn21" Then Return 109
        If EERSTE_BOD = "Hn21-F" Then Return 109
        If EERSTE_BOD = "Hn21g" Then Return 110
        If EERSTE_BOD = "Hn21gx" Then Return 110
        If EERSTE_BOD = "Hn21t" Then Return 111
        If EERSTE_BOD = "Hn21v" Then Return 109
        If EERSTE_BOD = "Hn21w" Then Return 109
        If EERSTE_BOD = "Hn21wg" Then Return 109
        If EERSTE_BOD = "Hn21x" Then Return 111
        If EERSTE_BOD = "Hn21x-F" Then Return 111
        If EERSTE_BOD = "Hn21xg" Then Return 111
        If EERSTE_BOD = "Hn23" Then Return 113
        If EERSTE_BOD = "Hn23-F" Then Return 113
        If EERSTE_BOD = "Hn23g" Then Return 110
        If EERSTE_BOD = "Hn23t" Then Return 111
        If EERSTE_BOD = "Hn23x" Then Return 111
        If EERSTE_BOD = "Hn23x-F" Then Return 111
        If EERSTE_BOD = "Hn23xg" Then Return 111
        If EERSTE_BOD = "Hn30" Then Return 114
        If EERSTE_BOD = "Hn30g" Then Return 114
        If EERSTE_BOD = "Hn30x" Then Return 114
        If EERSTE_BOD = "hRd10A" Then Return 119
        If EERSTE_BOD = "hRd10C" Then Return 119
        If EERSTE_BOD = "hRd90A" Then Return 116
        If EERSTE_BOD = "hVb" Then Return 101
        If EERSTE_BOD = "hVc" Then Return 101
        If EERSTE_BOD = "hVcc" Then Return 101
        If EERSTE_BOD = "hVd" Then Return 101
        If EERSTE_BOD = "hVk" Then Return 106
        If EERSTE_BOD = "hVkl" Then Return 106
        If EERSTE_BOD = "hVr" Then Return 101
        If EERSTE_BOD = "hVs" Then Return 101
        If EERSTE_BOD = "hVsc" Then Return 101
        If EERSTE_BOD = "hVz" Then Return 102
        If EERSTE_BOD = "hVzc" Then Return 102
        If EERSTE_BOD = "hVzg" Then Return 102
        If EERSTE_BOD = "hVzx" Then Return 102
        If EERSTE_BOD = "hZd20A" Then Return 107
        If EERSTE_BOD = "iVc" Then Return 105
        If EERSTE_BOD = "iVp" Then Return 105
        If EERSTE_BOD = "iVpc" Then Return 105
        If EERSTE_BOD = "iVpg" Then Return 105
        If EERSTE_BOD = "iVpt" Then Return 105
        If EERSTE_BOD = "iVpx" Then Return 105
        If EERSTE_BOD = "iVs" Then Return 105
        If EERSTE_BOD = "iVz" Then Return 105
        If EERSTE_BOD = "iVzg" Then Return 105
        If EERSTE_BOD = "iVzt" Then Return 105
        If EERSTE_BOD = "iVzx" Then Return 105
        If EERSTE_BOD = "iWp" Then Return 105
        If EERSTE_BOD = "iWpc" Then Return 105
        If EERSTE_BOD = "iWpg" Then Return 105
        If EERSTE_BOD = "iWpt" Then Return 105
        If EERSTE_BOD = "iWpx" Then Return 105
        If EERSTE_BOD = "iWz" Then Return 105
        If EERSTE_BOD = "iWzt" Then Return 105
        If EERSTE_BOD = "iWzx" Then Return 105
        If EERSTE_BOD = "kcHn21" Then Return 119
        If EERSTE_BOD = "kgpZg30" Then Return 120
        If EERSTE_BOD = "kHn21" Then Return 119
        If EERSTE_BOD = "kHn21g" Then Return 120
        If EERSTE_BOD = "kHn21x" Then Return 119
        If EERSTE_BOD = "kHn23" Then Return 119
        If EERSTE_BOD = "kHn23x" Then Return 119
        If EERSTE_BOD = "kHn30" Then Return 120
        If EERSTE_BOD = "KK" Then Return 121
        If EERSTE_BOD = "KM" Then Return 121
        If EERSTE_BOD = "kMn43C" Then Return 117
        If EERSTE_BOD = "kMn43Cp" Then Return 117
        If EERSTE_BOD = "kMn43Cpx" Then Return 117
        If EERSTE_BOD = "kMn43Cv" Then Return 118
        If EERSTE_BOD = "kMn43Cwp" Then Return 117
        If EERSTE_BOD = "kMn48C" Then Return 117
        If EERSTE_BOD = "kMn48Cl" Then Return 117
        If EERSTE_BOD = "kMn48Clv" Then Return 118
        If EERSTE_BOD = "kMn48Cv" Then Return 118
        If EERSTE_BOD = "kMn48Cvl" Then Return 118
        If EERSTE_BOD = "kMn48Cw" Then Return 117
        If EERSTE_BOD = "kMn63C" Then Return 117
        If EERSTE_BOD = "kMn63Cp" Then Return 119
        If EERSTE_BOD = "kMn63Cpx" Then Return 119
        If EERSTE_BOD = "kMn63Cv" Then Return 118
        If EERSTE_BOD = "kMn63Cwp" Then Return 119
        If EERSTE_BOD = "kMn68C" Then Return 117
        If EERSTE_BOD = "kMn68Cl" Then Return 117
        If EERSTE_BOD = "kMn68Cv" Then Return 118
        If EERSTE_BOD = "kpZg20A" Then Return 119
        If EERSTE_BOD = "kpZg21" Then Return 119
        If EERSTE_BOD = "kpZg21g" Then Return 120
        If EERSTE_BOD = "kpZg23" Then Return 119
        If EERSTE_BOD = "kpZg23g" Then Return 120
        If EERSTE_BOD = "kpZg23t" Then Return 119
        If EERSTE_BOD = "kpZg23x" Then Return 119
        If EERSTE_BOD = "kpZn21" Then Return 119
        If EERSTE_BOD = "kpZn21g" Then Return 120
        If EERSTE_BOD = "kpZn23" Then Return 119
        If EERSTE_BOD = "kpZn23x" Then Return 119
        If EERSTE_BOD = "KRd1" Then Return 119
        If EERSTE_BOD = "KRd1g" Then Return 120
        If EERSTE_BOD = "KRd7" Then Return 119
        If EERSTE_BOD = "KRd7g" Then Return 120
        If EERSTE_BOD = "KRn1" Then Return 119
        If EERSTE_BOD = "KRn1g" Then Return 120
        If EERSTE_BOD = "KRn2" Then Return 119
        If EERSTE_BOD = "KRn2g" Then Return 120
        If EERSTE_BOD = "KRn2w" Then Return 119
        If EERSTE_BOD = "KRn8" Then Return 119
        If EERSTE_BOD = "KRn8g" Then Return 120
        If EERSTE_BOD = "KS" Then Return 115
        If EERSTE_BOD = "kSn13A" Then Return 119
        If EERSTE_BOD = "kSn13Av" Then Return 119
        If EERSTE_BOD = "kSn13Aw" Then Return 119
        If EERSTE_BOD = "kSn14A" Then Return 119
        If EERSTE_BOD = "kSn14Ap" Then Return 119
        If EERSTE_BOD = "kSn14Av" Then Return 119
        If EERSTE_BOD = "kSn14Aw" Then Return 119
        If EERSTE_BOD = "kSn14Awp" Then Return 119
        If EERSTE_BOD = "KT" Then Return 115
        If EERSTE_BOD = "kVb" Then Return 103
        If EERSTE_BOD = "kVc" Then Return 103
        If EERSTE_BOD = "kVcc" Then Return 103
        If EERSTE_BOD = "kVd" Then Return 103
        If EERSTE_BOD = "kVk" Then Return 106
        If EERSTE_BOD = "kVr" Then Return 103
        If EERSTE_BOD = "kVs" Then Return 103
        If EERSTE_BOD = "kVsc" Then Return 103
        If EERSTE_BOD = "kVz" Then Return 104
        If EERSTE_BOD = "kVzc" Then Return 104
        If EERSTE_BOD = "kVzx" Then Return 104
        If EERSTE_BOD = "kWp" Then Return 104
        If EERSTE_BOD = "kWpg" Then Return 104
        If EERSTE_BOD = "kWpx" Then Return 104
        If EERSTE_BOD = "kWz" Then Return 104
        If EERSTE_BOD = "kWzg" Then Return 104
        If EERSTE_BOD = "kWzx" Then Return 104
        If EERSTE_BOD = "KX" Then Return 115
        If EERSTE_BOD = "kZb21" Then Return 119
        If EERSTE_BOD = "kZb23" Then Return 119
        If EERSTE_BOD = "kZn10A" Then Return 119
        If EERSTE_BOD = "kZn10Av" Then Return 119
        If EERSTE_BOD = "kZn21" Then Return 119
        If EERSTE_BOD = "kZn21g" Then Return 120
        If EERSTE_BOD = "kZn21p" Then Return 119
        If EERSTE_BOD = "kZn21r" Then Return 119
        If EERSTE_BOD = "kZn21w" Then Return 119
        If EERSTE_BOD = "kZn21x" Then Return 119
        If EERSTE_BOD = "kZn23" Then Return 119
        If EERSTE_BOD = "kZn30" Then Return 120
        If EERSTE_BOD = "kZn30A" Then Return 120
        If EERSTE_BOD = "kZn30Ar" Then Return 120
        If EERSTE_BOD = "kZn30x" Then Return 120
        If EERSTE_BOD = "kZn40A" Then Return 119
        If EERSTE_BOD = "kZn40Ap" Then Return 119
        If EERSTE_BOD = "kZn40Av" Then Return 119
        If EERSTE_BOD = "kZn50A" Then Return 119
        If EERSTE_BOD = "kZn50Ap" Then Return 119
        If EERSTE_BOD = "kZn50Ar" Then Return 119
        If EERSTE_BOD = "Ld5" Then Return 121
        If EERSTE_BOD = "Ld5g" Then Return 121
        If EERSTE_BOD = "Ld5m" Then Return 121
        If EERSTE_BOD = "Ld5t" Then Return 121
        If EERSTE_BOD = "Ld6" Then Return 121
        If EERSTE_BOD = "Ld6a" Then Return 121
        If EERSTE_BOD = "Ld6g" Then Return 121
        If EERSTE_BOD = "Ld6k" Then Return 121
        If EERSTE_BOD = "Ld6m" Then Return 121
        If EERSTE_BOD = "Ld6s" Then Return 121
        If EERSTE_BOD = "Ld6t" Then Return 121
        If EERSTE_BOD = "Ldd5" Then Return 121
        If EERSTE_BOD = "Ldd5g" Then Return 121
        If EERSTE_BOD = "Ldd6" Then Return 121
        If EERSTE_BOD = "Ldh5" Then Return 121
        If EERSTE_BOD = "Ldh5g" Then Return 121
        If EERSTE_BOD = "Ldh5t" Then Return 121
        If EERSTE_BOD = "Ldh6" Then Return 121
        If EERSTE_BOD = "Ldh6m" Then Return 121
        If EERSTE_BOD = "lFG" Then Return 114
        If EERSTE_BOD = "lFK" Then Return 121
        If EERSTE_BOD = "lFKk" Then Return 121
        If EERSTE_BOD = "Lh5" Then Return 121
        If EERSTE_BOD = "Lh5g" Then Return 121
        If EERSTE_BOD = "Lh6g" Then Return 121
        If EERSTE_BOD = "Lh6s" Then Return 121
        If EERSTE_BOD = "lKK" Then Return 116
        If EERSTE_BOD = "lKM" Then Return 116
        If EERSTE_BOD = "lKRd7" Then Return 119
        If EERSTE_BOD = "lKS" Then Return 121
        If EERSTE_BOD = "Ln5" Then Return 121
        If EERSTE_BOD = "Ln5g" Then Return 121
        If EERSTE_BOD = "Ln5m" Then Return 121
        If EERSTE_BOD = "Ln5t" Then Return 121
        If EERSTE_BOD = "Ln6a" Then Return 121
        If EERSTE_BOD = "Ln6m" Then Return 121
        If EERSTE_BOD = "Ln6t" Then Return 121
        If EERSTE_BOD = "Lnd5" Then Return 121
        If EERSTE_BOD = "Lnd5g" Then Return 121
        If EERSTE_BOD = "Lnd5m" Then Return 121
        If EERSTE_BOD = "Lnd5t" Then Return 121
        If EERSTE_BOD = "Lnd6" Then Return 121
        If EERSTE_BOD = "Lnd6v" Then Return 121
        If EERSTE_BOD = "Lnh6" Then Return 121
        If EERSTE_BOD = "MA" Then Return 116
        If EERSTE_BOD = "mcY23" Then Return 113
        If EERSTE_BOD = "mcY23x" Then Return 111
        If EERSTE_BOD = "mHd23" Then Return 113
        If EERSTE_BOD = "mHn21x" Then Return 111
        If EERSTE_BOD = "mHn23x" Then Return 111
        If EERSTE_BOD = "MK" Then Return 116
        If EERSTE_BOD = "mKK" Then Return 116
        If EERSTE_BOD = "mKRd7" Then Return 119
        If EERSTE_BOD = "mKX" Then Return 115
        If EERSTE_BOD = "mLd6s" Then Return 121
        If EERSTE_BOD = "mLh6s" Then Return 121
        If EERSTE_BOD = "Mn12A" Then Return 119
        If EERSTE_BOD = "Mn12Ap" Then Return 119
        If EERSTE_BOD = "Mn12Av" Then Return 119
        If EERSTE_BOD = "Mn12Awp" Then Return 119
        If EERSTE_BOD = "Mn15A" Then Return 115
        If EERSTE_BOD = "Mn15Ap" Then Return 119
        If EERSTE_BOD = "Mn15Av" Then Return 118
        If EERSTE_BOD = "Mn15Aw" Then Return 115
        If EERSTE_BOD = "Mn15Awp" Then Return 119
        If EERSTE_BOD = "Mn15C" Then Return 115
        If EERSTE_BOD = "Mn15Clv" Then Return 118
        If EERSTE_BOD = "Mn15Cv" Then Return 118
        If EERSTE_BOD = "Mn15Cw" Then Return 115
        If EERSTE_BOD = "Mn22A" Then Return 119
        If EERSTE_BOD = "Mn22Alv" Then Return 115
        If EERSTE_BOD = "Mn22Ap" Then Return 119
        If EERSTE_BOD = "Mn22Av" Then Return 115
        If EERSTE_BOD = "Mn22Aw" Then Return 119
        If EERSTE_BOD = "Mn22Awp" Then Return 119
        If EERSTE_BOD = "Mn22Ax" Then Return 119
        If EERSTE_BOD = "Mn25A" Then Return 115
        If EERSTE_BOD = "Mn25Alv" Then Return 115
        If EERSTE_BOD = "Mn25Ap" Then Return 119
        If EERSTE_BOD = "Mn25Av" Then Return 118
        If EERSTE_BOD = "Mn25Aw" Then Return 115
        If EERSTE_BOD = "Mn25Awp" Then Return 119
        If EERSTE_BOD = "Mn25C" Then Return 115
        If EERSTE_BOD = "Mn25Cp" Then Return 119
        If EERSTE_BOD = "Mn25Cv" Then Return 118
        If EERSTE_BOD = "Mn25Cw" Then Return 115
        If EERSTE_BOD = "Mn35A" Then Return 116
        If EERSTE_BOD = "Mn35Ap" Then Return 119
        If EERSTE_BOD = "Mn35Av" Then Return 118
        If EERSTE_BOD = "Mn35Aw" Then Return 116
        If EERSTE_BOD = "Mn35Awp" Then Return 119
        If EERSTE_BOD = "Mn35Ax" Then Return 116
        If EERSTE_BOD = "Mn45A" Then Return 117
        If EERSTE_BOD = "Mn45Ap" Then Return 119
        If EERSTE_BOD = "Mn45Av" Then Return 118
        If EERSTE_BOD = "Mn52C" Then Return 119
        If EERSTE_BOD = "Mn52Cp" Then Return 119
        If EERSTE_BOD = "Mn52Cpx" Then Return 119
        If EERSTE_BOD = "Mn52Cwp" Then Return 119
        If EERSTE_BOD = "Mn52Cx" Then Return 119
        If EERSTE_BOD = "Mn56A" Then Return 117
        If EERSTE_BOD = "Mn56Ap" Then Return 119
        If EERSTE_BOD = "Mn56Av" Then Return 118
        If EERSTE_BOD = "Mn56Aw" Then Return 117
        If EERSTE_BOD = "Mn56C" Then Return 117
        If EERSTE_BOD = "Mn56Cp" Then Return 119
        If EERSTE_BOD = "Mn56Cv" Then Return 118
        If EERSTE_BOD = "Mn56Cwp" Then Return 119
        If EERSTE_BOD = "Mn82A" Then Return 119
        If EERSTE_BOD = "Mn82Ap" Then Return 119
        If EERSTE_BOD = "Mn82C" Then Return 119
        If EERSTE_BOD = "Mn82Cp" Then Return 119
        If EERSTE_BOD = "Mn82Cpx" Then Return 119
        If EERSTE_BOD = "Mn82Cwp" Then Return 119
        If EERSTE_BOD = "Mn85C" Then Return 116
        If EERSTE_BOD = "Mn85Clwp" Then Return 119
        If EERSTE_BOD = "Mn85Cp" Then Return 119
        If EERSTE_BOD = "Mn85Cv" Then Return 118
        If EERSTE_BOD = "Mn85Cw" Then Return 116
        If EERSTE_BOD = "Mn85Cwp" Then Return 119
        If EERSTE_BOD = "Mn86A" Then Return 117
        If EERSTE_BOD = "Mn86Al" Then Return 117
        If EERSTE_BOD = "Mn86Av" Then Return 118
        If EERSTE_BOD = "Mn86Aw" Then Return 117
        If EERSTE_BOD = "Mn86C" Then Return 117
        If EERSTE_BOD = "Mn86Cl" Then Return 117
        If EERSTE_BOD = "Mn86Clv" Then Return 117
        If EERSTE_BOD = "Mn86Clw" Then Return 117
        If EERSTE_BOD = "Mn86Clwp" Then Return 119
        If EERSTE_BOD = "Mn86Cp" Then Return 119
        If EERSTE_BOD = "Mn86Cv" Then Return 118
        If EERSTE_BOD = "Mn86Cw" Then Return 117
        If EERSTE_BOD = "Mn86Cwp" Then Return 119
        If EERSTE_BOD = "Mo10A" Then Return 115
        If EERSTE_BOD = "Mo10Av" Then Return 115
        If EERSTE_BOD = "Mo20A" Then Return 115
        If EERSTE_BOD = "Mo20Av" Then Return 115
        If EERSTE_BOD = "Mo50C" Then Return 115
        If EERSTE_BOD = "Mo80A" Then Return 116
        If EERSTE_BOD = "Mo80Ap" Then Return 119
        If EERSTE_BOD = "Mo80Av" Then Return 118
        If EERSTE_BOD = "Mo80C" Then Return 116
        If EERSTE_BOD = "Mo80Cl" Then Return 116
        If EERSTE_BOD = "Mo80Cp" Then Return 119
        If EERSTE_BOD = "Mo80Cv" Then Return 118
        If EERSTE_BOD = "Mo80Cvl" Then Return 118
        If EERSTE_BOD = "Mo80Cw" Then Return 116
        If EERSTE_BOD = "Mo80Cwp" Then Return 119
        If EERSTE_BOD = "MOb12" Then Return 119
        If EERSTE_BOD = "MOb15" Then Return 115
        If EERSTE_BOD = "MOb72" Then Return 119
        If EERSTE_BOD = "MOb75" Then Return 116
        If EERSTE_BOD = "MOo02" Then Return 119
        If EERSTE_BOD = "MOo02v" Then Return 119
        If EERSTE_BOD = "MOo05" Then Return 115
        If EERSTE_BOD = "Mv41C" Then Return 118
        If EERSTE_BOD = "Mv41Cl" Then Return 118
        If EERSTE_BOD = "Mv41Cp" Then Return 118
        If EERSTE_BOD = "Mv41Cv" Then Return 118
        If EERSTE_BOD = "Mv51A" Then Return 118
        If EERSTE_BOD = "Mv51Al" Then Return 118
        If EERSTE_BOD = "Mv51Ap" Then Return 118
        If EERSTE_BOD = "Mv61C" Then Return 118
        If EERSTE_BOD = "Mv61Cl" Then Return 118
        If EERSTE_BOD = "Mv61Cp" Then Return 118
        If EERSTE_BOD = "Mv81A" Then Return 118
        If EERSTE_BOD = "Mv81Al" Then Return 118
        If EERSTE_BOD = "Mv81Ap" Then Return 118
        If EERSTE_BOD = "mY23" Then Return 113
        If EERSTE_BOD = "mY23x" Then Return 111
        If EERSTE_BOD = "mZb23x" Then Return 111
        If EERSTE_BOD = "MZk" Then Return 121
        If EERSTE_BOD = "MZz" Then Return 107
        If EERSTE_BOD = "nAO" Then Return 119
        If EERSTE_BOD = "nkZn21" Then Return 119
        If EERSTE_BOD = "nkZn50Ab" Then Return 119
        If EERSTE_BOD = "nMn15A" Then Return 115
        If EERSTE_BOD = "nMn15Av" Then Return 115
        If EERSTE_BOD = "nMo10A" Then Return 115
        If EERSTE_BOD = "nMo10Av" Then Return 118
        If EERSTE_BOD = "nMo80A" Then Return 116
        If EERSTE_BOD = "nMo80Aw" Then Return 116
        If EERSTE_BOD = "nMv61C" Then Return 118
        If EERSTE_BOD = "npMo50l" Then Return 115
        If EERSTE_BOD = "npMo80l" Then Return 116
        If EERSTE_BOD = "nSn13A" Then Return 113
        If EERSTE_BOD = "nSn13Av" Then Return 113
        If EERSTE_BOD = "nvWz" Then Return 102
        If EERSTE_BOD = "nZn21" Then Return 107
        If EERSTE_BOD = "nZn40A" Then Return 107
        If EERSTE_BOD = "nZn50A" Then Return 107
        If EERSTE_BOD = "nZn50Ab" Then Return 107
        If EERSTE_BOD = "ohVb" Then Return 101
        If EERSTE_BOD = "ohVc" Then Return 101
        If EERSTE_BOD = "ohVk" Then Return 106
        If EERSTE_BOD = "ohVs" Then Return 101
        If EERSTE_BOD = "opVb" Then Return 103
        If EERSTE_BOD = "opVc" Then Return 103
        If EERSTE_BOD = "opVk" Then Return 106
        If EERSTE_BOD = "opVs" Then Return 103
        If EERSTE_BOD = "pKRn1" Then Return 119
        If EERSTE_BOD = "pKRn1g" Then Return 120
        If EERSTE_BOD = "pKRn2" Then Return 119
        If EERSTE_BOD = "pKRn2g" Then Return 120
        If EERSTE_BOD = "pLn5" Then Return 121
        If EERSTE_BOD = "pLn5g" Then Return 121
        If EERSTE_BOD = "pMn52A" Then Return 119
        If EERSTE_BOD = "pMn52C" Then Return 119
        If EERSTE_BOD = "pMn52Cp" Then Return 119
        If EERSTE_BOD = "pMn55A" Then Return 115
        If EERSTE_BOD = "pMn55Av" Then Return 118
        If EERSTE_BOD = "pMn55Aw" Then Return 115
        If EERSTE_BOD = "pMn55C" Then Return 115
        If EERSTE_BOD = "pMn55Cp" Then Return 119
        If EERSTE_BOD = "pMn56C" Then Return 117
        If EERSTE_BOD = "pMn56Cl" Then Return 117
        If EERSTE_BOD = "pMn82A" Then Return 119
        If EERSTE_BOD = "pMn82C" Then Return 119
        If EERSTE_BOD = "pMn85A" Then Return 116
        If EERSTE_BOD = "pMn85Aw" Then Return 116
        If EERSTE_BOD = "pMn85C" Then Return 116
        If EERSTE_BOD = "pMn85Cv" Then Return 118
        If EERSTE_BOD = "pMn86C" Then Return 117
        If EERSTE_BOD = "pMn86Cl" Then Return 117
        If EERSTE_BOD = "pMn86Cv" Then Return 118
        If EERSTE_BOD = "pMn86Cw" Then Return 117
        If EERSTE_BOD = "pMn86Cwl" Then Return 117
        If EERSTE_BOD = "pMo50" Then Return 115
        If EERSTE_BOD = "pMo50l" Then Return 115
        If EERSTE_BOD = "pMo50w" Then Return 115
        If EERSTE_BOD = "pMo80" Then Return 116
        If EERSTE_BOD = "pMo80l" Then Return 116
        If EERSTE_BOD = "pMo80v" Then Return 118
        If EERSTE_BOD = "pMv51" Then Return 118
        If EERSTE_BOD = "pMv81" Then Return 118
        If EERSTE_BOD = "pMv81l" Then Return 118
        If EERSTE_BOD = "pMv81p" Then Return 118
        If EERSTE_BOD = "pRn56p" Then Return 119
        If EERSTE_BOD = "pRn56v" Then Return 118
        If EERSTE_BOD = "pRn56wp" Then Return 119
        If EERSTE_BOD = "pRn59" Then Return 119
        If EERSTE_BOD = "pRn59p" Then Return 119
        If EERSTE_BOD = "pRn59t" Then Return 119
        If EERSTE_BOD = "pRn59w" Then Return 119
        If EERSTE_BOD = "pRn86" Then Return 117
        If EERSTE_BOD = "pRn86p" Then Return 119
        If EERSTE_BOD = "pRn86t" Then Return 117
        If EERSTE_BOD = "pRn86v" Then Return 118
        If EERSTE_BOD = "pRn86w" Then Return 117
        If EERSTE_BOD = "pRn86wp" Then Return 119
        If EERSTE_BOD = "pRn89v" Then Return 118
        If EERSTE_BOD = "pRv81" Then Return 118
        If EERSTE_BOD = "pVb" Then Return 103
        If EERSTE_BOD = "pVc" Then Return 103
        If EERSTE_BOD = "pVcc" Then Return 103
        If EERSTE_BOD = "pVd" Then Return 103
        If EERSTE_BOD = "pVk" Then Return 106
        If EERSTE_BOD = "pVr" Then Return 103
        If EERSTE_BOD = "pVs" Then Return 103
        If EERSTE_BOD = "pVsc" Then Return 103
        If EERSTE_BOD = "pVsl" Then Return 103
        If EERSTE_BOD = "pVz" Then Return 104
        If EERSTE_BOD = "pVzx" Then Return 104
        If EERSTE_BOD = "pZg20A" Then Return 107
        If EERSTE_BOD = "pZg20Ar" Then Return 107
        If EERSTE_BOD = "pZg21" Then Return 109
        If EERSTE_BOD = "pZg21g" Then Return 110
        If EERSTE_BOD = "pZg21r" Then Return 111
        If EERSTE_BOD = "pZg21t" Then Return 111
        If EERSTE_BOD = "pZg21w" Then Return 109
        If EERSTE_BOD = "pZg21x" Then Return 111
        If EERSTE_BOD = "pZg23" Then Return 113
        If EERSTE_BOD = "pZg23g" Then Return 113
        If EERSTE_BOD = "pZg23r" Then Return 113
        If EERSTE_BOD = "pZg23t" Then Return 111
        If EERSTE_BOD = "pZg23w" Then Return 113
        If EERSTE_BOD = "pZg23x" Then Return 111
        If EERSTE_BOD = "pZg30" Then Return 114
        If EERSTE_BOD = "pZg30p" Then Return 114
        If EERSTE_BOD = "pZg30r" Then Return 114
        If EERSTE_BOD = "pZg30x" Then Return 114
        If EERSTE_BOD = "pZn21" Then Return 109
        If EERSTE_BOD = "pZn21g" Then Return 110
        If EERSTE_BOD = "pZn21t" Then Return 111
        If EERSTE_BOD = "pZn21tg" Then Return 109
        If EERSTE_BOD = "pZn21v" Then Return 109
        If EERSTE_BOD = "pZn21x" Then Return 111
        If EERSTE_BOD = "pZn23" Then Return 113
        If EERSTE_BOD = "pZn23g" Then Return 110
        If EERSTE_BOD = "pZn23gx" Then Return 110
        If EERSTE_BOD = "pZn23t" Then Return 111
        If EERSTE_BOD = "pZn23v" Then Return 113
        If EERSTE_BOD = "pZn23w" Then Return 113
        If EERSTE_BOD = "pZn23x" Then Return 111
        If EERSTE_BOD = "pZn23x-F" Then Return 111
        If EERSTE_BOD = "pZn30" Then Return 114
        If EERSTE_BOD = "pZn30g" Then Return 114
        If EERSTE_BOD = "pZn30r" Then Return 114
        If EERSTE_BOD = "pZn30w" Then Return 114
        If EERSTE_BOD = "pZn30x" Then Return 114
        If EERSTE_BOD = "Rd10A" Then Return 119
        If EERSTE_BOD = "Rd10Ag" Then Return 119
        If EERSTE_BOD = "Rd10C" Then Return 119
        If EERSTE_BOD = "Rd10Cg" Then Return 120
        If EERSTE_BOD = "Rd10Cm" Then Return 119
        If EERSTE_BOD = "Rd10Cp" Then Return 119
        If EERSTE_BOD = "Rd90A" Then Return 116
        If EERSTE_BOD = "Rd90C" Then Return 116
        If EERSTE_BOD = "Rd90Cg" Then Return 120
        If EERSTE_BOD = "Rd90Cm" Then Return 116
        If EERSTE_BOD = "Rd90Cp" Then Return 119
        If EERSTE_BOD = "Rn14C" Then Return 117
        If EERSTE_BOD = "Rn15A" Then Return 115
        If EERSTE_BOD = "Rn15C" Then Return 115
        If EERSTE_BOD = "Rn15Cg" Then Return 115
        If EERSTE_BOD = "Rn15Ct" Then Return 115
        If EERSTE_BOD = "Rn15Cw" Then Return 115
        If EERSTE_BOD = "Rn42Cg" Then Return 119
        If EERSTE_BOD = "Rn42Cp" Then Return 119
        If EERSTE_BOD = "Rn44C" Then Return 117
        If EERSTE_BOD = "Rn44Cv" Then Return 118
        If EERSTE_BOD = "Rn44Cw" Then Return 117
        If EERSTE_BOD = "Rn45A" Then Return 117
        If EERSTE_BOD = "Rn46A" Then Return 117
        If EERSTE_BOD = "Rn46Av" Then Return 118
        If EERSTE_BOD = "Rn46Aw" Then Return 117
        If EERSTE_BOD = "Rn47C" Then Return 117
        If EERSTE_BOD = "Rn47Cg" Then Return 120
        If EERSTE_BOD = "Rn47Cp" Then Return 119
        If EERSTE_BOD = "Rn47Cv" Then Return 118
        If EERSTE_BOD = "Rn47Cw" Then Return 117
        If EERSTE_BOD = "Rn47Cwp" Then Return 119
        If EERSTE_BOD = "Rn52A" Then Return 120
        If EERSTE_BOD = "Rn52Ag" Then Return 120
        If EERSTE_BOD = "Rn62C" Then Return 119
        If EERSTE_BOD = "Rn62Cg" Then Return 120
        If EERSTE_BOD = "Rn62Cp" Then Return 119
        If EERSTE_BOD = "Rn62Cwp" Then Return 119
        If EERSTE_BOD = "Rn66A" Then Return 117
        If EERSTE_BOD = "Rn66Av" Then Return 118
        If EERSTE_BOD = "Rn67C" Then Return 117
        If EERSTE_BOD = "Rn67Cg" Then Return 120
        If EERSTE_BOD = "Rn67Cp" Then Return 119
        If EERSTE_BOD = "Rn67Cv" Then Return 118
        If EERSTE_BOD = "Rn67Cwp" Then Return 119
        If EERSTE_BOD = "Rn82A" Then Return 119
        If EERSTE_BOD = "Rn82Ag" Then Return 120
        If EERSTE_BOD = "Rn94C" Then Return 117
        If EERSTE_BOD = "Rn94Cv" Then Return 118
        If EERSTE_BOD = "Rn95A" Then Return 116
        If EERSTE_BOD = "Rn95Av" Then Return 118
        If EERSTE_BOD = "Rn95C" Then Return 116
        If EERSTE_BOD = "Rn95Cg" Then Return 120
        If EERSTE_BOD = "Rn95Cm" Then Return 116
        If EERSTE_BOD = "Rn95Cp" Then Return 119
        If EERSTE_BOD = "Ro40A" Then Return 117
        If EERSTE_BOD = "Ro40Av" Then Return 118
        If EERSTE_BOD = "Ro40C" Then Return 117
        If EERSTE_BOD = "Ro40Cv" Then Return 118
        If EERSTE_BOD = "Ro40Cw" Then Return 117
        If EERSTE_BOD = "Ro60A" Then Return 116
        If EERSTE_BOD = "Ro60C" Then Return 116
        If EERSTE_BOD = "ROb72" Then Return 119
        If EERSTE_BOD = "ROb75" Then Return 116
        If EERSTE_BOD = "Rv01A" Then Return 118
        If EERSTE_BOD = "Rv01C" Then Return 118
        If EERSTE_BOD = "Rv01Cg" Then Return 118
        If EERSTE_BOD = "Rv01Cp" Then Return 118
        If EERSTE_BOD = "saVc" Then Return 101
        If EERSTE_BOD = "saVz" Then Return 102
        If EERSTE_BOD = "sHn21" Then Return 109
        If EERSTE_BOD = "shVz" Then Return 102
        If EERSTE_BOD = "skVc" Then Return 103
        If EERSTE_BOD = "skWz" Then Return 104
        If EERSTE_BOD = "Sn13A" Then Return 113
        If EERSTE_BOD = "Sn13Ap" Then Return 113
        If EERSTE_BOD = "Sn13Av" Then Return 113
        If EERSTE_BOD = "Sn13Aw" Then Return 113
        If EERSTE_BOD = "Sn13Awp" Then Return 113
        If EERSTE_BOD = "Sn14A" Then Return 113
        If EERSTE_BOD = "Sn14Ap" Then Return 113
        If EERSTE_BOD = "Sn14Av" Then Return 113
        If EERSTE_BOD = "spVc" Then Return 103
        If EERSTE_BOD = "spVz" Then Return 104
        If EERSTE_BOD = "sVc" Then Return 101
        If EERSTE_BOD = "sVk" Then Return 106
        If EERSTE_BOD = "sVp" Then Return 102
        If EERSTE_BOD = "sVs" Then Return 101
        If EERSTE_BOD = "svWp" Then Return 102
        If EERSTE_BOD = "svWz" Then Return 102
        If EERSTE_BOD = "svWzt" Then Return 102
        If EERSTE_BOD = "sVz" Then Return 102
        If EERSTE_BOD = "sVzt" Then Return 102
        If EERSTE_BOD = "sVzx" Then Return 102
        If EERSTE_BOD = "tZd21" Then Return 107
        If EERSTE_BOD = "tZd21g" Then Return 110
        If EERSTE_BOD = "tZd21v" Then Return 107
        If EERSTE_BOD = "tZd23" Then Return 113
        If EERSTE_BOD = "Vb" Then Return 101
        If EERSTE_BOD = "Vc" Then Return 101
        If EERSTE_BOD = "Vd" Then Return 101
        If EERSTE_BOD = "Vk" Then Return 106
        If EERSTE_BOD = "Vo" Then Return 101
        If EERSTE_BOD = "Vp" Then Return 102
        If EERSTE_BOD = "Vpx" Then Return 102
        If EERSTE_BOD = "Vr" Then Return 101
        If EERSTE_BOD = "Vs" Then Return 101
        If EERSTE_BOD = "Vsc" Then Return 101
        If EERSTE_BOD = "vWp" Then Return 102
        If EERSTE_BOD = "vWpg" Then Return 102
        If EERSTE_BOD = "vWpt" Then Return 102
        If EERSTE_BOD = "vWpx" Then Return 102
        If EERSTE_BOD = "vWz" Then Return 102
        If EERSTE_BOD = "vWzg" Then Return 102
        If EERSTE_BOD = "vWzr" Then Return 102
        If EERSTE_BOD = "vWzt" Then Return 102
        If EERSTE_BOD = "vWzx" Then Return 102
        If EERSTE_BOD = "Vz" Then Return 102
        If EERSTE_BOD = "Vzc" Then Return 102
        If EERSTE_BOD = "Vzg" Then Return 102
        If EERSTE_BOD = "Vzt" Then Return 102
        If EERSTE_BOD = "Vzx" Then Return 102
        If EERSTE_BOD = "Wg" Then Return 106
        If EERSTE_BOD = "Wgl" Then Return 106
        If EERSTE_BOD = "Wo" Then Return 106
        If EERSTE_BOD = "Wol" Then Return 106
        If EERSTE_BOD = "Wov" Then Return 106
        If EERSTE_BOD = "Y21" Then Return 109
        If EERSTE_BOD = "Y21g" Then Return 110
        If EERSTE_BOD = "Y21x" Then Return 111
        If EERSTE_BOD = "Y23" Then Return 113
        If EERSTE_BOD = "Y23b" Then Return 113
        If EERSTE_BOD = "Y23g" Then Return 110
        If EERSTE_BOD = "Y23x" Then Return 111
        If EERSTE_BOD = "Y30" Then Return 114
        If EERSTE_BOD = "Y30x" Then Return 114
        If EERSTE_BOD = "Zb20A" Then Return 107
        If EERSTE_BOD = "Zb21" Then Return 109
        If EERSTE_BOD = "Zb21g" Then Return 110
        If EERSTE_BOD = "Zb23" Then Return 113
        If EERSTE_BOD = "Zb23g" Then Return 113
        If EERSTE_BOD = "Zb23t" Then Return 111
        If EERSTE_BOD = "Zb23x" Then Return 111
        If EERSTE_BOD = "Zb30" Then Return 114
        If EERSTE_BOD = "Zb30A" Then Return 114
        If EERSTE_BOD = "Zb30g" Then Return 114
        If EERSTE_BOD = "Zd20A" Then Return 107
        If EERSTE_BOD = "Zd20Ab" Then Return 107
        If EERSTE_BOD = "Zd21" Then Return 107
        If EERSTE_BOD = "Zd21g" Then Return 107
        If EERSTE_BOD = "Zd23" Then Return 113
        If EERSTE_BOD = "Zd30" Then Return 114
        If EERSTE_BOD = "Zd30A" Then Return 114
        If EERSTE_BOD = "zEZ21" Then Return 112
        If EERSTE_BOD = "zEZ21g" Then Return 112
        If EERSTE_BOD = "zEZ21t" Then Return 112
        If EERSTE_BOD = "zEZ21w" Then Return 112
        If EERSTE_BOD = "zEZ21x" Then Return 112
        If EERSTE_BOD = "zEZ23" Then Return 112
        If EERSTE_BOD = "zEZ23g" Then Return 112
        If EERSTE_BOD = "zEZ23t" Then Return 112
        If EERSTE_BOD = "zEZ23w" Then Return 112
        If EERSTE_BOD = "zEZ23x" Then Return 112
        If EERSTE_BOD = "zEZ30" Then Return 112
        If EERSTE_BOD = "zEZ30g" Then Return 112
        If EERSTE_BOD = "zEZ30x" Then Return 112
        If EERSTE_BOD = "zgHd30" Then Return 114
        If EERSTE_BOD = "zgMn15C" Then Return 115
        If EERSTE_BOD = "zgMn88C" Then Return 117
        If EERSTE_BOD = "zgY30" Then Return 114
        If EERSTE_BOD = "zHd21" Then Return 108
        If EERSTE_BOD = "zHd21g" Then Return 108
        If EERSTE_BOD = "zHn21" Then Return 108
        If EERSTE_BOD = "zHn23" Then Return 109
        If EERSTE_BOD = "zhVk" Then Return 106
        If EERSTE_BOD = "zKRn1g" Then Return 120
        If EERSTE_BOD = "zKRn2" Then Return 119
        If EERSTE_BOD = "zkVc" Then Return 103
        If EERSTE_BOD = "zkWp" Then Return 104
        If EERSTE_BOD = "zMn15A" Then Return 115
        If EERSTE_BOD = "zMn22Ap" Then Return 119
        If EERSTE_BOD = "zMn25Ap" Then Return 119
        If EERSTE_BOD = "zMn56Cp" Then Return 117
        If EERSTE_BOD = "zMo10A" Then Return 115
        If EERSTE_BOD = "zMv41C" Then Return 118
        If EERSTE_BOD = "zMv61C" Then Return 118
        If EERSTE_BOD = "Zn10A" Then Return 107
        If EERSTE_BOD = "Zn10Ap" Then Return 107
        If EERSTE_BOD = "Zn10Av" Then Return 107
        If EERSTE_BOD = "Zn10Aw" Then Return 107
        If EERSTE_BOD = "Zn10Awp" Then Return 107
        If EERSTE_BOD = "Zn21" Then Return 107
        If EERSTE_BOD = "Zn21-F" Then Return 107
        If EERSTE_BOD = "Zn21g" Then Return 107
        If EERSTE_BOD = "Zn21-H" Then Return 107
        If EERSTE_BOD = "Zn21p" Then Return 107
        If EERSTE_BOD = "Zn21r" Then Return 107
        If EERSTE_BOD = "Zn21t" Then Return 107
        If EERSTE_BOD = "Zn21v" Then Return 107
        If EERSTE_BOD = "Zn21w" Then Return 107
        If EERSTE_BOD = "Zn21x" Then Return 107
        If EERSTE_BOD = "Zn21x-F" Then Return 107
        If EERSTE_BOD = "Zn23" Then Return 113
        If EERSTE_BOD = "Zn23-F" Then Return 113
        If EERSTE_BOD = "Zn23g" Then Return 113
        If EERSTE_BOD = "Zn23g-F" Then Return 113
        If EERSTE_BOD = "Zn23-H" Then Return 113
        If EERSTE_BOD = "Zn23p" Then Return 113
        If EERSTE_BOD = "Zn23r" Then Return 113
        If EERSTE_BOD = "Zn23t" Then Return 111
        If EERSTE_BOD = "Zn23x" Then Return 111
        If EERSTE_BOD = "Zn30" Then Return 114
        If EERSTE_BOD = "Zn30A" Then Return 114
        If EERSTE_BOD = "Zn30Ab" Then Return 114
        If EERSTE_BOD = "Zn30Ag" Then Return 114
        If EERSTE_BOD = "Zn30Ar" Then Return 114
        If EERSTE_BOD = "Zn30g" Then Return 114
        If EERSTE_BOD = "Zn30r" Then Return 114
        If EERSTE_BOD = "Zn30v" Then Return 114
        If EERSTE_BOD = "Zn30x" Then Return 114
        If EERSTE_BOD = "Zn40A" Then Return 107
        If EERSTE_BOD = "Zn40Ap" Then Return 107
        If EERSTE_BOD = "Zn40Ar" Then Return 107
        If EERSTE_BOD = "Zn40Av" Then Return 107
        If EERSTE_BOD = "Zn50A" Then Return 107
        If EERSTE_BOD = "Zn50Ab" Then Return 107
        If EERSTE_BOD = "Zn50Ap" Then Return 107
        If EERSTE_BOD = "Zn50Ar" Then Return 107
        If EERSTE_BOD = "Zn50Aw" Then Return 107
        If EERSTE_BOD = "zpZn23w" Then Return 113
        If EERSTE_BOD = "zRd10A" Then Return 119
        If EERSTE_BOD = "zRn15C" Then Return 115
        If EERSTE_BOD = "zRn47Cwp" Then Return 117
        If EERSTE_BOD = "zRn62C" Then Return 119
        If EERSTE_BOD = "zSn14A" Then Return 113
        If EERSTE_BOD = "zVc" Then Return 105
        If EERSTE_BOD = "zVp" Then Return 105
        If EERSTE_BOD = "zVpg" Then Return 105
        If EERSTE_BOD = "zVpt" Then Return 105
        If EERSTE_BOD = "zVpx" Then Return 105
        If EERSTE_BOD = "zVs" Then Return 105
        If EERSTE_BOD = "zVz" Then Return 105
        If EERSTE_BOD = "zVzg" Then Return 105
        If EERSTE_BOD = "zVzt" Then Return 105
        If EERSTE_BOD = "zVzx" Then Return 105
        If EERSTE_BOD = "zWp" Then Return 105
        If EERSTE_BOD = "zWpg" Then Return 105
        If EERSTE_BOD = "zWpt" Then Return 105
        If EERSTE_BOD = "zWpx" Then Return 105
        If EERSTE_BOD = "zWz" Then Return 105
        If EERSTE_BOD = "zWzg" Then Return 105
        If EERSTE_BOD = "zWzt" Then Return 105
        If EERSTE_BOD = "zWzx" Then Return 105
        If EERSTE_BOD = "zY21" Then Return 108
        If EERSTE_BOD = "zY21g" Then Return 108
        If EERSTE_BOD = "zY23" Then Return 109
        If EERSTE_BOD = "zY30" Then Return 114
        Return 0

    End Function

    Friend Function ParseSobekTable(ByRef myRecord As String) As clsSobekTable

        Dim tableString As String = Me.ParseTable(myRecord)

        Dim myTable As New clsSobekTable(Me.setup)
        myTable.Read(tableString)
        Return myTable

    End Function

    'Requires reference to MapWinGIS
    'mwSourceGrid is already instantiated and opened from an existing grid file
    'SourceGrid is dimensioned as Dim SourceGrid(MaxCol, MaxRow) as Float
    Public Function ArrayFromMapWindowGrid(ByRef mwSourceGrid As MapWinGIS.Grid) As Single(,)
        Dim m_mrow As Integer, m_mcol As Integer
        Dim row, col As Integer
        Dim vals() As Single
        m_mrow = mwSourceGrid.Header.NumberRows - 1
        m_mcol = mwSourceGrid.Header.NumberCols - 1
        Dim SourceGrid(m_mcol, m_mrow) As Single
        For row = 0 To m_mrow
            ReDim vals(m_mcol)
            mwSourceGrid.GetRow(row, vals(0))
            For col = 0 To m_mcol
                SourceGrid(col, row) = vals(col)
            Next
        Next
        Return SourceGrid
    End Function

    <SecurityPermissionAttribute(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.UnmanagedCode)>
    Public Sub ReleaseComObject(ByVal obj As Object, ByVal collect As Boolean)
        'Me.setup.Log.AddDebugMessage("Releasing " & obj.GetType.FullName & " Collect is " & collect)

        While True
            If (Marshal.ReleaseComObject(obj) <= 0) Then
                Exit While
            End If
        End While

        If collect Then
            'Me.setup.Log.AddMessage("Memory used before collection: " & GC.GetTotalMemory(False))
            GC.Collect()
        End If

        'Dim numBytes As Long = GC.GetTotalMemory(True)
        'Me.setup.Log.AddMessage("Memory used after releasing: " & Math.Round(numBytes / 1024 / 1024, 1).ToString())

    End Sub

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

    Public Function FileInUse(ByVal sFile As String) As Boolean
        Dim thisFileInUse As Boolean = False
        If System.IO.File.Exists(sFile) Then
            Try
                Using f As New IO.FileStream(sFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    ' thisFileInUse = False
                End Using
            Catch
                thisFileInUse = True
            End Try
        End If
        Return thisFileInUse
    End Function

    Public Sub ControlProcesses(ByRef ProcessCollection As List(Of System.Diagnostics.Process), ByVal MaxSimultaneously As Integer, ByVal WaitMilliseconds As Integer, ByVal SafetyValveMiliseconds As Integer)
        Dim mySW As New System.Diagnostics.Stopwatch
        Dim k As Long

        mySW.Start()
        Dim nUnfinished As Integer = 1000
        While nUnfinished > 0
            nUnfinished = 0
            For k = 0 To ProcessCollection.Count - 1
                If Not ProcessCollection(k).HasExited Then nUnfinished += 1
            Next
            If nUnfinished >= MaxSimultaneously Then System.Threading.Thread.Sleep(WaitMilliseconds)
            If mySW.ElapsedMilliseconds > SafetyValveMiliseconds Then nUnfinished = 0 'safety valve
        End While
        mySW.Stop()
    End Sub

    Public Function calChannelElevationPoint(ByVal ChanBedWidth As Double, ByVal ChanSlope As Double, ByVal ChanDepth As Double, ByVal mv As Double, ByVal Dist As Double, ByRef Inside As Boolean) As Double
        'returns the elevation level inside an ideal channel, given the distance from its center line
        'note: channelslope is defined as horizontal/vertical
        'the boolean Inside is meant to distinquish points that fall inside the channel (computed level < specified surface level) or outside
        Dim bl As Double = mv - ChanDepth 'bed level
        Dim cl As Double
        Dim OnSlopeDist As Double = Dist - ChanBedWidth / 2

        If OnSlopeDist <= 0 Then
            Inside = True
            Return bl
        Else
            'calculate the designed channel level and compare it to the specified surface level
            cl = bl + OnSlopeDist / ChanSlope
            If cl > mv Then
                'computed level exceeds specified surface level so we're outside the channel. return the surface level
                Inside = False
                Return mv
            Else
                'computed level underceeds specified surface level, so we're inside the channel. Return the computed value
                Inside = True
                Return cl
            End If
        End If

    End Function

    Public Shared Function LastDirFromDir(ByVal myPath As String)
        Dim Dirs As String()
        If Right(myPath, 1) = "\" Then myPath = Left(myPath, myPath.Length - 1)
        Dirs = Split(myPath, "\")
        Return Dirs(Dirs.Count - 1)
    End Function

    Public Shared Sub DeleteAllDirectoryContents(ByVal myDir As String)

        'first remove all files in the current dir
        Dim myFile As String
        For Each myFile In Directory.GetFiles(myDir)
            File.Delete(myFile)
        Next

        'then remove all files in the current dir
        Dim mySubDir As String
        For Each mySubDir In Directory.GetDirectories(myDir)
            Directory.Delete(mySubDir, True)
        Next

    End Sub

    Public Shared Sub DirectoryCopy(ByVal sourceDirName As String, ByVal destDirName As String, ByVal copySubDirs As Boolean)

        ' Get the subdirectories for the specified directory. 
        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        If Not dir.Exists Then
            Throw New DirectoryNotFoundException(
                "Source directory does not exist or could not be found: " _
                + sourceDirName)
        End If

        ' If the destination directory doesn't exist, create it. 
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        End If
        ' Get the files in the directory and copy them to the new location. 
        Dim files As FileInfo() = dir.GetFiles()
        For Each file In files
            Dim temppath As String = Path.Combine(destDirName, file.Name)
            file.CopyTo(temppath, False)
        Next file

        ' If copying subdirectories, copy them and their contents to new location. 
        If copySubDirs Then
            For Each subdir In dirs
                Dim temppath As String = Path.Combine(destDirName, subdir.Name)
                DirectoryCopy(subdir.FullName, temppath, copySubDirs)
            Next subdir
        End If
    End Sub

    Public Shared Function ReplaceSectionInFile(ByVal filename As String, ByVal HeadString As String, ByVal TailString As String, ByVal ReplaceString As String, ByVal CompareMethod As Microsoft.VisualBasic.CompareMethod) As Boolean
        Try
            Dim myContent As String
            Dim HeadPos As Long = 0, TailPos As Long = 1
            Dim LeftSection As String, RightSection As String

            'read the file content and replace the string
            Using myReader As New StreamReader(filename)
                myContent = myReader.ReadToEnd()

                While TailPos > HeadPos
                    HeadPos = InStr(HeadPos + 1, myContent, HeadString, CompareMethod)
                    TailPos = InStr(HeadPos + 1, myContent, TailString, CompareMethod)

                    If HeadPos = 0 OrElse TailPos = 0 Then
                        'we're done here. No match found
                        HeadPos = TailPos
                    ElseIf HeadPos < TailPos AndAlso HeadPos > 0 Then
                        'match found!
                        LeftSection = Left(myContent, HeadPos - 1)
                        RightSection = Right(myContent, myContent.Length - TailPos - TailString.Length)
                        myContent = LeftSection & ReplaceString & RightSection

                        'reset the search area for the next round
                        HeadPos = TailPos
                        TailPos = TailPos + 1
                    End If

                End While


            End Using

            'write the file
            Using myWriter As New StreamWriter(filename)
                myWriter.Write(myContent)
            End Using

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function ReplaceSectionInLine(ByVal filename As String, ByVal HeadString As String, ByVal TailString As String, ByVal ReplaceString As String, ByVal CompareMethod As Microsoft.VisualBasic.CompareMethod) As Boolean
        Try
            Dim myLine As String
            Dim HeadPos As Long = 0, TailPos As Long = 1
            Dim LeftSection As String, RightSection As String
            Dim newContent As String = ""

            'read the file content LINE BY LINE and replace the string
            Using myReader As New StreamReader(filename)
                myLine = myReader.ReadLine

                While TailPos > HeadPos
                    HeadPos = InStr(HeadPos + 1, myLine, HeadString, CompareMethod)
                    TailPos = InStr(HeadPos + 1, myLine, TailString, CompareMethod)
                    If HeadPos = 0 OrElse TailPos = 0 Then
                        'we're done here. No match found
                        HeadPos = TailPos
                    ElseIf HeadPos < TailPos AndAlso HeadPos > 0 Then
                        'match found!
                        LeftSection = Left(myLine, HeadPos - 1)
                        RightSection = Right(myLine, myLine.Length - TailPos - TailString.Length)
                        myLine = LeftSection & ReplaceString & RightSection

                        'reset the search area for the next round
                        HeadPos = TailPos
                        TailPos = TailPos + 1
                    End If
                End While

                'add the line to the new content
                If newContent = "" Then
                    newContent = myLine
                Else
                    newContent &= vbCrLf & myLine
                End If

            End Using

            'write the file
            Using myWriter As New StreamWriter(filename)
                myWriter.WriteLine(newContent)
            End Using

        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Shared Function ReplaceStringInFile(ByVal FileName As String, ByVal ReplaceStr As String, ByVal ReplaceByStr As String, ByVal CompareMethod As Microsoft.VisualBasic.CompareMethod) As Boolean
        Try
            Dim myContent As String

            'read the file content and replace the string
            Using myReader As New StreamReader(FileName)
                myContent = myReader.ReadToEnd()
                myContent = Replace(myContent, ReplaceStr, ReplaceByStr, , , CompareMethod)
            End Using

            'write the file
            Using myWriter As New StreamWriter(FileName)
                myWriter.Write(myContent)
            End Using

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function DirectoryRelativeToRoot(ByVal Path As String) As String
        Dim Root As String = Directory.GetDirectoryRoot(Path)
        Path = Replace(Path, Root, "")
        Return Path
    End Function


    Public Sub QueryAccessDataBase(ByVal myPath As String, ByVal myTable As String, ByVal queries As List(Of String))
        Dim cn As New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim myQuery As String


        cn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & myPath & ";"
        cn.Open()

        'execute the query

        'populate the grid containing summer volumes
        For Each myQuery In queries
            da = New OleDb.OleDbDataAdapter(myQuery, cn)
            da.Update(ds)
        Next

        cn.Close()
    End Sub


End Class


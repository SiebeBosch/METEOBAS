Imports METEOBAS.General
Imports METEOBAS.GeneralFunctions
Imports MapWinGIS

Public Class clsRaster
    Public Path As String
    Public ElevationMultiplier As Double 'used to convert units such as cm NAP to m NAP

    Public Grid As MapWinGIS.Grid   'optie 1: inlezen als Mapwindow-grid
    Friend ASCII As clsASCIIGrid    'optie 2: inlezen als 2D-array (sneller bewerken en zoeken)

    Public XLLCenter As Double, YLLCenter As Double
    Public XLLCorner As Double, YLLCorner As Double, XURCorner As Double, YURCorner As Double
    Public NoDataVal As Double
    Public dX As Double, dY As Double, CellArea As Double
    Public Selected(,) As Boolean
    Friend nCol As Integer, nRow As Integer
    Public Stats As clsGridStats

    Private setup As clsSetup

    Public Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
        Grid = New MapWinGIS.Grid
        ASCII = New clsASCIIGrid(Me.setup)
        Stats = New clsGridStats
    End Sub

    Public Sub New(ByRef mySetup As clsSetup, ByVal myPath As String)
        Me.setup = mySetup
        Grid = New MapWinGIS.Grid
        ASCII = New clsASCIIGrid(Me.setup)
        Stats = New clsGridStats
        Path = myPath
    End Sub

    Public Sub Initialize(ByVal myPath As String)
        Path = myPath
    End Sub

    Public Function ReadHeader(ByVal myType As MapWinGIS.GridDataType, ByVal inRAM As Boolean) As Boolean
        If Grid.Open(Path, myType, inRAM) Then
            Call CompleteMetaHeader()
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub GetCellBounds(ByRef r As Long, ByRef c As Long, ByRef Xcs As Double, ByRef Ycs As Double, ByRef Xce As Double, ByRef Yce As Double)
        Xcs = Grid.Header.XllCenter + (c - 0.5) * Grid.Header.dX
        Xce = Grid.Header.XllCenter + (c + 0.5) * Grid.Header.dX
        Ycs = Grid.Header.YllCenter + (Grid.Header.NumberRows - r - 0.5) * Grid.Header.dY
        Yce = Grid.Header.YllCenter + (Grid.Header.NumberRows - r - 1.5) * Grid.Header.dY
    End Sub

    Public Sub getGridCenter(ByRef X As Double, ByRef Y As Double)
        X = (XLLCorner + XURCorner) / 2
        Y = (YLLCorner + YURCorner) / 2
    End Sub

    Public Function setPath(ByVal myPath As String, Optional ByVal Multiplier As Double = 1) As Boolean
        If System.IO.File.Exists(myPath) Then
            Path = myPath
            ElevationMultiplier = Multiplier
            Return True
        Else
            Return False
        End If
    End Function

    Public Function PointInside(ByRef myPoint As MapWinGIS.Point) As Boolean
        If myPoint.x >= XLLCorner AndAlso myPoint.x < XURCorner Then
            If myPoint.y > YLLCorner AndAlso myPoint.y < YURCorner Then
                Return True
            End If
        End If
        Return False
    End Function


    Public Function ShapeInside(ByRef myShape As MapWinGIS.Shape) As Boolean
        If myShape.Extents.xMin >= XLLCorner AndAlso myShape.Extents.xMax <= XURCorner Then
            If myShape.Extents.yMin >= YLLCorner AndAlso myShape.Extents.yMax <= YURCorner Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Function ShapeOverlaps(ByRef myShape As MapWinGIS.Shape) As Boolean
        If myShape.Extents.xMin >= XLLCorner AndAlso myShape.Extents.xMin <= XURCorner AndAlso myShape.Extents.yMin >= YLLCorner AndAlso myShape.Extents.yMin <= YLLCorner Then Return True
        If myShape.Extents.xMin >= XLLCorner AndAlso myShape.Extents.xMin <= XURCorner AndAlso myShape.Extents.yMax >= YLLCorner AndAlso myShape.Extents.yMax <= YLLCorner Then Return True
        If myShape.Extents.xMax >= XLLCorner AndAlso myShape.Extents.xMax <= XURCorner AndAlso myShape.Extents.yMin >= YLLCorner AndAlso myShape.Extents.yMin <= YLLCorner Then Return True
        If myShape.Extents.xMax >= XLLCorner AndAlso myShape.Extents.xMax <= XURCorner AndAlso myShape.Extents.yMax >= YLLCorner AndAlso myShape.Extents.yMax <= YLLCorner Then Return True
        Return False
    End Function

    Public Sub Close()
        Grid.Close()
    End Sub

    Public Sub InitializeSelected()
        Dim Rows As Integer = Grid.Header.NumberRows
        Dim Cols As Integer = Grid.Header.NumberCols
        ReDim Selected(0 To Rows - 1, 0 To Cols - 1)
    End Sub

    Public Function PointInsideGrid(ByVal X As Double, ByVal Y As Double) As Boolean
        If X >= XLLCorner AndAlso X <= XLLCorner + dX * nCol AndAlso Y >= YLLCorner AndAlso Y <= YLLCorner + dY * nRow Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function PointInValidCell(ByVal X As Double, ByVal Y As Double) As Boolean
        Dim myValue As Single

        Try
            If Not GetCellValueFromXY(X, Y, myValue) Then
                Me.setup.Log.AddError("No grid value found for co-ordinate " & X & "," & Y & ". Location probably outside of grid.")
                Return False
            ElseIf Not myValue = NoDataVal Then
                Return True
            Else
                Return True
            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Friend Function GetCellValueFromXY(ByVal x As Double, ByVal y As Double, ByRef myValue As Single) As Boolean
        Dim r As Integer, c As Integer
        Try
            If Not GetRCFromXY(x, y, r, c) Then
                Me.setup.Log.AddError("Error getting grid row and column number for co-ordinate " & x & "," & y & ". Location probably outside grid.")
                Return False
            End If
            myValue = Grid.Value(c, r)
            Return True
        Catch ex As Exception
            myValue = 0
            Return False
        End Try
    End Function

    Public Function GetRCFromXY(ByVal X As Double, ByVal Y As Double, ByRef rowIdx As Integer, ByRef colIdx As Integer) As Boolean

        If X >= XLLCorner AndAlso X <= XURCorner AndAlso Y >= YLLCorner AndAlso Y <= YURCorner Then
            ' MapWindow telt rijen van boven naar beneden (gechecked!)
            colIdx = setup.GeneralFunctions.RoundUD((X - XLLCorner) / dX, 0, False)
            rowIdx = setup.GeneralFunctions.RoundUD((YURCorner - Y) / dY, 0, False)
            Return True
        Else
            Return False
        End If

    End Function

    Public Function getXYFromRC(ByVal r As Long, ByVal c As Long, ByRef X As Double, ByRef Y As Double) As Boolean
        'note: the row and column numbers are zero-based
        If r >= 0 AndAlso r <= (nRow - 1) AndAlso c >= 0 AndAlso c <= (nCol - 1) Then
            X = XLLCorner + (c + 0.5) * dX
            Y = YURCorner - (r + 0.5) * dY
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetRowColExtent(ByVal xMin As Double, ByVal xMax As Double, ByVal yMin As Double, ByVal yMax As Double,
                                    ByRef startCol As Integer, ByRef endCol As Integer, ByRef startRow As Integer, ByRef endRow As Integer) As Boolean

        ' Paul Meems, 5 June 2012: Added:
        If Grid Is Nothing Then Throw New Exception("The grid object is not set")

        'zoek het kolomnummer bij xMin
        If xMin > XURCorner OrElse xMax < XLLCorner OrElse yMin > YURCorner OrElse yMax < YLLCorner Then
            Me.setup.Log.AddDebugMessage("Shape valt in zijn geheel buiten het grid")
            Return False
        End If

        'let op: MapWindow telt rijen van boven naar beneden!!!!!
        startCol = Math.Max(0, Me.setup.GeneralFunctions.RoundUD((xMin - XLLCorner) / dX, 0, False))
        endCol = Math.Min(nCol - 1, Me.setup.GeneralFunctions.RoundUD((xMax - XLLCorner) / dX, 0, False) - 1)
        startRow = Math.Max(0, Me.setup.GeneralFunctions.RoundUD((YURCorner - yMax) / dY, 0, False))
        endRow = Math.Min(nRow - 1, Me.setup.GeneralFunctions.RoundUD((YURCorner - yMin) / dY, 0, False) - 1)

        Return True
    End Function

    Public Function GetCellCenter(ByVal rowIdx As Integer, ByVal colIdx As Integer, ByRef X As Double, ByRef Y As Double) As Boolean

        ' MapWindow telt rijen van boven naar beneden (gechecked!)
        X = XLLCenter + dX * colIdx
        Y = YLLCenter + ((nRow - 1) * dY) - (dY * rowIdx)
        Return True

    End Function

    Public Function GetLowestValue() As Double
        Dim r As Long, c As Long
        Dim minVal As Double = 9000000000.0

        For r = 0 To nRow - 1
            For c = 0 To nCol - 1
                If Grid.Value(c, r) <> Grid.Header.NodataValue AndAlso Grid.Value(c, r) < minVal Then
                    minVal = Grid.Value(c, r)
                End If
            Next
        Next
        Return minVal
    End Function

    Public Function CompleteMetaHeaderWithoutReading() As Boolean
        'completes the meta-header without actually reading the grid
        'this speeds up the reading process but BE CAREFUL! If the grid changes
        'during a process, the metadata is not automatically updated
        'Author: Siebe Bosch
        'Date: 7-2-2014
        Try

            If XLLCorner = 0 AndAlso dX > 0 AndAlso XLLCenter <> 0 Then
                XLLCorner = XLLCenter - dX / 2
                If nCol > 0 Then XURCorner = XLLCorner + nCol * dX
            ElseIf XURCorner = 0 AndAlso XLLCorner <> 0 AndAlso dX <> 0 AndAlso nCol > 0 Then
                XURCorner = XLLCorner + nCol * dX
            End If

            If YLLCorner = 0 AndAlso dY > 0 AndAlso YLLCenter <> 0 Then
                YLLCorner = YLLCenter - dY / 2
                If nRow > 0 Then YURCorner = YLLCorner + nRow * dY
            ElseIf YURCorner = 0 AndAlso YLLCorner <> 0 AndAlso dY <> 0 AndAlso nCol > 0 Then
                YURCorner = YLLCorner + nRow * dY
            End If

            If dX > 0 AndAlso dY > 0 Then
                CellArea = dX * dY
            End If

            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function CompleteMetaHeader() As Boolean
        'completes the meta-header using data from the actual header
        'Author: Siebe Bosch
        'Date: 7-2-2014

        Try
            If Not Grid.Header Is Nothing Then
                MetaDataFromHeader()
            Else
                If Not Grid.Open(Path) Then Throw New Exception("Could not read " & Path)
                MetaDataFromHeader()
                Grid.Close()
            End If
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Sub MetaDataFromHeader()

        XLLCenter = Grid.Header.XllCenter
        YLLCenter = Grid.Header.YllCenter
        nRow = Grid.Header.NumberRows
        nCol = Grid.Header.NumberCols
        dX = Grid.Header.dX
        dY = Grid.Header.dY
        NoDataVal = Grid.Header.NodataValue

        'afgeleide data
        XLLCorner = XLLCenter - dX / 2
        YLLCorner = YLLCenter - dY / 2
        YURCorner = YLLCorner + dY * nRow
        XURCorner = XLLCorner + dX * nCol
        CellArea = dX * dY

    End Sub

    Public Function Read(Optional ByVal InRAM As Boolean = True) As Boolean

        ' Paul Meems, 5 June 2012: Made several changes
        'If Not Me.Grid Is Nothing Then
        '  If Not String.IsNullOrEmpty(Me.Grid.Filename) Then
        '    Me.setup.Log.AddWarning("Elevation grid is already open!")
        '    Return False
        '  End If
        'End If
        If Not Grid.Open(Path, GridDataType.UnknownDataType, InRAM) Then
            Dim log As String = "Failed to open gridfile: " + Grid.ErrorMsg(Grid.LastErrorCode)
            Me.setup.Log.AddError(log)
            Throw New Exception(log)
        Else
            Call CompleteMetaHeader()
        End If

        Return True

    End Function

    Public Function GetNearestCell(ByVal r As Integer, ByVal c As Integer, ByRef Value As Double, ByRef Distance As Double, ByVal SkipZeroVals As Boolean) As Boolean
        'This function retrieves the nearest by valid grid cell. It does so by creating an increasing circle around the cell and
        'returning the fist valid cell value and its distance
        Dim CellRadius As Integer = 0
        Dim CircleCells As New Dictionary(Of String, clsRC)
        Dim myRC As clsRC, myVal As Double

        Do
            'create a collection of cells that overlap the circle around our center point
            CircleCells = CollectCellsFromCircle(c, r, CellRadius)

            If SkipZeroVals Then
                For Each myRC In CircleCells.Values
                    myVal = Grid.Value(myRC.C, myRC.R)
                    If myVal <> NoDataVal AndAlso myVal <> 0 AndAlso myVal > -3.0E+19 Then 'last one is to avoid nodata values of single type where nodat val is double
                        Distance = Me.setup.GeneralFunctions.Pythagoras(c, r, myRC.C, myRC.R) * dX    'distance (m) between the point found and the original cell
                        Value = myVal
                        Return True
                    End If
                Next
            Else
                For Each myRC In CircleCells.Values
                    myVal = Grid.Value(myRC.C, myRC.R)
                    If myVal <> NoDataVal AndAlso myVal > -3.0E+19 Then                    'last one is to avoid nodata values of single type where nodat val is double
                        Distance = Me.setup.GeneralFunctions.Pythagoras(c, r, myRC.C, myRC.R) * dX    'distance (m) between the point found and the original cell
                        Value = myVal
                        Return True
                    End If
                Next
            End If

            'safety valve
            CellRadius += 1
            If CellRadius > nRow Then Return False

        Loop
        Return False

    End Function

    Public Function GetNearestEdgeCell(ByVal StartR As Integer, ByVal StartC As Integer, ByRef EdgeR As Integer, ByRef EdgeC As Integer, ByVal MaxRadius As Double, ByRef Distance As Double) As Boolean
        'This function retrieves the nearest by valid grid cell. It does so by creating an increasing circle around the cell and
        'returning the fist valid cell value and its distance
        Dim CellRadius As Integer = 0
        Dim CircleCells As New Dictionary(Of String, clsRC)
        Dim myRC As clsRC


        Do
            'create a collection of cells that overlap the circle around our center point
            CircleCells = CollectCellsFromCircle(StartC, StartR, CellRadius)

            For Each myRC In CircleCells.Values
                If IsEdgeCell(myRC.R, myRC.C) Then
                    EdgeR = myRC.R
                    EdgeC = myRC.C
                    Distance = Me.setup.GeneralFunctions.Pythagoras(StartC, StartR, myRC.C, myRC.R) * dX    'distance (m) between the point found and the original cell
                    Return True
                End If
            Next

            'safety valve
            CellRadius += 1
            If CellRadius > nRow Then Return False
            If CellRadius * Grid.Header.dX > MaxRadius Then Return False
        Loop
        Return False

    End Function

    Public Function IsEdgeCell(r As Integer, c As Integer, Optional IncludeDiagonalCells As Boolean = True) As Boolean
        If Grid.Value(c, r) = Grid.Header.NodataValue Then Return False
        If r = 0 OrElse c = 0 OrElse r = Grid.Header.NumberRows - 1 OrElse c = Grid.Header.NumberCols - 1 Then Return True
        If Grid.Value(c, r - 1) = Grid.Header.NodataValue Then Return True
        If Grid.Value(c, r + 1) = Grid.Header.NodataValue Then Return True
        If Grid.Value(c - 1, r) = Grid.Header.NodataValue Then Return True
        If Grid.Value(c + 1, r) = Grid.Header.NodataValue Then Return True
        If IncludeDiagonalCells Then
            If Grid.Value(c - 1, r - 1) = Grid.Header.NodataValue Then Return True
            If Grid.Value(c - 1, r + 1) = Grid.Header.NodataValue Then Return True
            If Grid.Value(c + 1, r - 1) = Grid.Header.NodataValue Then Return True
            If Grid.Value(c + 1, r + 1) = Grid.Header.NodataValue Then Return True
        End If
        Return False
    End Function


    Public Function CollectCellsFromCircle(ByVal c0 As Integer, ByVal r0 As Integer, ByVal radius As Integer) As Dictionary(Of String, clsRC)

        'Date: 1-3-2014
        'Author: Siebe Bosch
        'Description: midpoint circle algorithm
        'c#-code from wikipedia: http://en.wikipedia.org/wiki/Midpoint_circle_algorithm
        'then converted to vb.net using http://www.developerfusion.com/tools/convert/csharp-to-vb/?batchId=dc6b40bd-ab00-42f9-bcee-a4c646033fd1
        'and added a collection of instances of the class clsXr to store the results in
        Dim c As Integer = radius, r As Integer = 0
        Dim radiusError As Integer = 1 - c
        Dim Points As New Dictionary(Of String, clsRC)
        Dim myKey As String

        While c >= r
            'note: row = y, col = x, to we've swapped the order in order to store results in clsRC
            myKey = Str(r + r0).Trim & "_" & Str(c + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(r + r0, c + c0) Then Points.Add(myKey, New clsRC(r + r0, c + c0))
            myKey = Str(c + r0).Trim & "_" & Str(r + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(c + r0, r + c0) Then Points.Add(myKey, New clsRC(c + r0, r + c0))
            myKey = Str(r + r0).Trim & "_" & Str(-c + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(r + r0, -c + c0) Then Points.Add(myKey, New clsRC(r + r0, -c + c0))
            myKey = Str(c + r0).Trim & "_" & Str(-r + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(c + r0, -r + c0) Then Points.Add(myKey, New clsRC(c + r0, -r + c0))
            myKey = Str(-r + r0).Trim & "_" & Str(-c + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(-r + r0, -c + c0) Then Points.Add(myKey, New clsRC(-r + r0, -c + c0))
            myKey = Str(-c + r0).Trim & "_" & Str(-r + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(-c + r0, -r + c0) Then Points.Add(myKey, New clsRC(-c + r0, -r + c0))
            myKey = Str(-r + r0).Trim & "_" & Str(c + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(-r + r0, c + c0) Then Points.Add(myKey, New clsRC(-r + r0, c + c0))
            myKey = Str(-c + r0).Trim & "_" & Str(r + c0).Trim
            If Not Points.ContainsKey(myKey) AndAlso CellValid(-c + r0, r + c0) Then Points.Add(myKey, New clsRC(-c + r0, r + c0))

            r += 1
            If radiusError < 0 Then
                radiusError += 2 * r + 1
            Else
                c -= 1
                radiusError += 2 * (r - c + 1)
            End If
        End While

        Return Points
    End Function

    Public Function CellValid(ByRef r As Integer, ByRef c As Integer) As Boolean
        If r >= 0 AndAlso r <= nRow - 1 Then
            If c >= 0 AndAlso c <= nCol - 1 Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Function CollectCellsFromLine(ByVal c0 As Integer, ByVal r0 As Integer, ByVal c1 As Integer, ByVal r1 As Integer) As Dictionary(Of String, clsRC)
        'Date: 1-3-2014
        'Author: Siebe Bosch
        'Description: returns a collection of cells that follow a straight line between (c0,r0) and (c1 and r1)
        'it uses the Bresenham's line algorithm to do so
        Dim Points As New Dictionary(Of String, clsRC)
        Dim myKey As String, myRC As clsRC
        Dim dx As Integer, dy As Integer
        Dim sx As Integer, sy As Integer
        Dim err As Integer, e2 As Integer

        dx = Math.Abs(c1 - c0)
        dy = Math.Abs(r1 - r0)
        If c0 < c1 Then sx = 1 Else sx = -1
        If r0 < r1 Then sy = 1 Else sy = -1
        err = dx - dy

        Do
            myRC = New clsRC(r0, c0)
            myKey = Str(r0).Trim & "_" & Str(c0).Trim
            If Not Points.ContainsKey(myKey) Then Points.Add(myKey, myRC)

            If c0 = c1 And r0 = r1 Then Exit Do
            e2 = 2 * err
            If e2 > -dy Then
                err = err - dy
                c0 = c0 + sx
            End If
            If c0 = c1 And r0 = r1 Then
                myRC = New clsRC(r0, c0)
                myKey = Str(r0).Trim & "_" & Str(c0).Trim
                If Not Points.ContainsKey(myKey) Then Points.Add(myKey, myRC)
                Exit Do
            End If
            If e2 < dx Then
                err = err + dx
                r0 = r0 + sy
            End If
        Loop

        Return Points
    End Function

End Class

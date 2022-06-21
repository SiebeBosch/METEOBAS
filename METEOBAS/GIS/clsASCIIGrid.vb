Imports System.IO
Imports METEOBAS.General

Public Class clsASCIIGrid

    Friend path As String
    Friend cols As Integer, rows As Integer
    Friend cellsize As Double, dx As Double, dy As Double
    Friend nodataval As Double
    Friend xllcorner As Double, yllcorner As Double
    Friend xllcenter As Double, yllcenter As Double
    Friend xurcorner As Double, yurcorner As Double

    Friend cells As Double(,)
    Private setup As clsSetup

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
    End Sub

    Friend Sub CreateNew(ByVal myPath As String, ByVal nCols As Integer, ByVal nRows As Integer, ByVal Xll As Double, ByVal yll As Double, ByVal CS As Double, ByVal Nodata As Double)
        cols = nCols
        rows = nRows
        cellsize = CS
        xllcorner = Xll
        yllcorner = yll
        nodataval = Nodata
        path = myPath
    End Sub

    Friend Function Write(ByVal myPath As String) As Boolean
        Dim r As Integer, c As Integer, mystr As String
        Using gridWriter As New StreamWriter(myPath)
            gridWriter.WriteLine("NCOLS " & cols)
            gridWriter.WriteLine("NROWS " & rows)
            gridWriter.WriteLine("XLLCORNER " & xllcorner)
            gridWriter.WriteLine("YLLCORNER " & yllcorner)
            gridWriter.WriteLine("CELLSIZE " & Math.Max(dx, cellsize))
            gridWriter.WriteLine("NODATA_VALUE " & nodataval)
            For r = 0 To rows - 1
                c = 0
                mystr = cells(r, c)
                For c = 1 To cols - 1
                    mystr = mystr & " " & cells(r, c)
                Next
                gridWriter.WriteLine(mystr)
            Next
        End Using
        Return True
    End Function

    Friend Function Read(ByVal myPath As String) As Boolean
        'Me.setup.Log.AddDebugMessage("In Read of clsASCIIGrid")
        path = myPath
        Dim row As Integer, col As Integer
        Dim tmpStr As String
        Dim entry As String
        Dim HeaderDataCollected As Boolean = False

        If Not File.Exists(path) Then
            Dim log As String = "Error reading ASCII Grid. Path doesn't exists"
            Me.setup.Log.AddError(log)
            Throw New ArgumentException(log, "path")
        End If


        'read the grid header
        Using ascreader As New StreamReader(path)
            While HeaderDataCollected = False
                entry = ascreader.ReadLine.Trim
                tmpStr = Me.setup.GeneralFunctions.ParseString(entry, " ")
                If tmpStr.Trim.ToLower = "ncols" Then
                    cols = entry
                ElseIf tmpStr.Trim.ToLower = "nrows" Then
                    rows = entry
                ElseIf tmpStr.Trim.ToLower = "xllcorner" Then
                    xllcorner = entry
                ElseIf tmpStr.Trim.ToLower = "yllcorner" Then
                    yllcorner = entry
                ElseIf tmpStr.Trim.ToLower = "xllcenter" Then
                    xllcenter = entry
                ElseIf tmpStr.Trim.ToLower = "yllcenter" Then
                    yllcenter = entry
                ElseIf tmpStr.Trim.ToLower = "cellsize" Then
                    cellsize = entry
                ElseIf tmpStr.Trim.ToLower = "dx" Then
                    dx = entry
                ElseIf tmpStr.Trim.ToLower = "dy" Then
                    dy = entry
                ElseIf tmpStr.Trim.ToLower = "nodata_value" Then
                    nodataval = entry
                    HeaderDataCollected = True
                End If
            End While

            If dx <> 0 Then cellsize = dx
            If xllcenter <> 0 Then xllcorner = xllcenter - cellsize / 2
            If yllcenter <> 0 Then yllcorner = yllcenter - cellsize / 2
            yurcorner = yllcorner + cellsize * rows
            xurcorner = xllcorner + cellsize * cols

            'header is ingelezen, dus ga nu door met de data zelf
            ReDim cells(rows - 1, cols - 1)


            'now read the griddata
            For row = 1 To rows
                entry = ascreader.ReadLine.Trim
                For col = 1 To cols
                    cells(row - 1, col - 1) = Me.setup.GeneralFunctions.ParseString(entry, " ")
                Next col
            Next row
        End Using
        Return True

    End Function

    Friend Function GetCellValueFromRC(ByVal r As Integer, ByVal c As Integer) As Double
        Return cells(r, c)
    End Function

    Friend Function GetCellValueFromXY(ByVal x As Double, ByVal y As Double) As Double
        Dim r As Integer, c As Integer
        Me.GetRCFromXY(x, y, r, c)
        Return cells(r, c)
    End Function

    Friend Sub GetRCFromXY(ByVal x As Double, ByVal y As Double, ByRef r As Integer, ByRef c As Integer)
        'wij (en SOBEK) tellen rijen van beneden naar boven
        c = Me.setup.GeneralFunctions.RoundUD((x - xllcorner) / cellsize, 0, False)
        r = Me.setup.GeneralFunctions.RoundUD((y - yllcorner) / cellsize, 0, False)
    End Sub

    Friend Function GetRCExtents(ByVal x_from As Double, ByVal x_to As Double, ByVal y_from As Double, ByVal y_to As Double,
                                 ByRef r_from As Integer, ByRef r_to As Integer, ByRef c_from As Integer, ByRef c_to As Integer) As Boolean
        'wij (en SOBEK) tellen rijen van beneden naar boven
        Call GetRCFromXY(x_from, y_from, r_from, c_from)
        Call GetRCFromXY(x_to, y_to, r_to, c_to)
        Return True
    End Function

    Friend Function GetXYfromRC(ByVal r As Integer, ByVal c As Integer, ByRef x As Double, ByRef y As Double) As Boolean
        x = xllcenter + (c + 0.5) * cellsize
        y = yllcenter + (r + 0.5) * cellsize
        Return True
    End Function
End Class

Imports METEOBAS.General

Public Class clsTiles
    'this class builds a dictionary of tiles, each able to contain a list of objects
    Dim Tiles As New Dictionary(Of String, List(Of Object))
    Dim Setup As clsSetup

    Dim xMin As Double, yMin As Double
    Dim xMax As Double, yMax As Double
    Dim cellSize As Double

    Public Sub New(ByRef mySetup As clssetup)
        Setup = mySetup
    End Sub


    Public Sub Create(myxmin As Double, myymin As Double, myxmax As Double, myymax As Double, mycellsize As Double)
        Dim r As Long, c As Long
        Dim myTile As List(Of Object)
        Dim myKey As String

        xMin = myxmin
        yMin = myymin
        xMax = myxmax
        yMax = myymax
        cellsize = mycellsize

        For c = 0 To Setup.GeneralFunctions.RoundUD((xMax - xMin) / cellsize, 0, False)
            For r = 0 To Setup.GeneralFunctions.RoundUD((yMax - yMin) / cellsize, 0, False)
                myKey = r.ToString.Trim & "_" & c.ToString.Trim
                myTile = New List(Of Object)
                Tiles.Add(myKey, myTile)
            Next
        Next

    End Sub

    Public Function AddPointObject(myObject As Object, x As Double, y As Double) As Boolean
        Try
            'add an object to its tile and returns the key
            Dim r As Long, c As Long, myKey As String
            r = Setup.GeneralFunctions.RoundUD((y - yMin) / cellSize, 0, False)
            c = Setup.GeneralFunctions.RoundUD((x - xMin) / cellSize, 0, False)
            myKey = r.ToString.Trim & "_" & c.ToString.Trim
            If Tiles.ContainsKey(myKey) Then
                Tiles.Item(myKey).Add(myObject)
            Else
                Throw New Exception("Error adding object to tile collection with x,y = " & x & "," & y)
            End If
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Error in function AddPointObject of class clsTiles.")
            Return False
        End Try
    End Function

    Public Function AddAreaObject(myObject As Object, objXMin As Double, objYMin As Double, objXMax As Double, objYMax As Double) As Boolean
        Try
            'add an object to its tile and returns the key
            Dim rMin As Long, cMin As Long, rMax As Long, cMax As Long, myKey As String
            Dim r As Long, c As Long

            'first find the range of rows and columns in this tiles object where our object overlaps
            rMin = Setup.GeneralFunctions.RoundUD((objYMin - yMin) / cellSize, 0, False)
            cMin = Setup.GeneralFunctions.RoundUD((objXMin - xMin) / cellSize, 0, False)
            rMax = Setup.GeneralFunctions.RoundUD((objYMax - yMin) / cellSize, 0, False)
            cMax = Setup.GeneralFunctions.RoundUD((objXMax - xMin) / cellSize, 0, False)

            For r = rMin To rMax
                For c = cMin To cMax
                    myKey = r.ToString.Trim & "_" & c.ToString.Trim
                    If Tiles.ContainsKey(myKey) Then
                        Tiles.Item(myKey).Add(myObject)
                    Else
                        Throw New Exception("Error adding object to tile collection with x,y-range = [" & objXMin & "-" & objXMax & "],[" & objYMin & "-" & objYMax & "]")
                    End If
                Next
            Next
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Me.Setup.Log.AddError("Error in function AddAreaObject of class clsTiles.")
            Return False
        End Try
    End Function


    Public Function getNearbyObjects(x As Double, y As Double, IncludeNeigbors As Boolean) As List(Of Object)
        Dim myList As New List(Of Object)
        Dim r As Long, c As Long, myKey As String
        Dim r1 As Long, c1 As Long
        r = Setup.GeneralFunctions.RoundUD((y - yMin) / cellSize, 0, False)
        c = Setup.GeneralFunctions.RoundUD((x - xMin) / cellSize, 0, False)
        myKey = r.ToString.Trim & "_" & c.ToString.Trim

        For i = 0 To Tiles.Item(myKey).Count - 1
            myList.Add(Tiles.Item(myKey).Item(i))
        Next

        If IncludeNeigbors Then
            For r1 = r - 1 To r + 1
                For c1 = c - 1 To c + 1
                    myKey = r1.ToString.Trim & "_" & c1.ToString.Trim
                    If Tiles.ContainsKey(myKey) AndAlso Not (r1 = r And c1 = 1) Then
                        For i = 0 To Tiles.Item(myKey).Count - 1
                            myList.Add(Tiles.Item(myKey).Item(i))
                        Next
                    End If
                Next
            Next
        End If

        Return myList

    End Function

End Class

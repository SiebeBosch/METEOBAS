Option Explicit On

Imports METEOBAS.General
Imports METEOBAS.GeneralFunctions
Imports System.IO
Imports MapWinGIS

Public Class clsGISData

    'Public GridCollection As clsGridCollection
    'Public SnappingPointsShapeFile As clsPointShapeFile     'a shapefile to log all snapping locations while building a model
    'Collections en Dictionaries
    'Public Catchments As clsCatchments
    Public SubcatchmentShapeFile As clsSubcatchmentShapeFile

    Private setup As clsSetup

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup

        'Shapefiles voor gebieden:
        'CatchmentShapeFile = New clsSubcatchmentShapeFile(Me.setup)             'hoogste orde gebieden (stroomgebieden)
        SubcatchmentShapeFile = New clsSubcatchmentShapeFile(Me.setup)              'middelste orde gebieden (peilgebieden)

        'Collections en Dictionaries
        'Catchments = New clsCatchments(Me.setup)                                        'Stroomgebieden
    End Sub

    Public Sub setAreaShapeFile(ByVal Path As String)
        SubcatchmentShapeFile = New clsSubcatchmentShapeFile(Me.setup, Path)
    End Sub

    Public Function ClipShapeFiles(ShapeFileToBeClipped As Shapefile, ShapeFileToClip As Shapefile) As Shapefile
        Try
            Dim newSf As MapWinGIS.Shapefile
            newSf = ShapeFileToBeClipped.Clip(False, ShapeFileToClip, False)
            If newSf Is Nothing Then Throw New Exception(ShapeFileToBeClipped.ErrorMsg(ShapeFileToBeClipped.LastErrorCode))
            Return newSf
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Me.setup.Log.AddError("Error clipping shapefile " & ShapeFileToBeClipped.Filename & " by shapefile " & ShapeFileToClip.Filename & ".")
            Return Nothing
        End Try
    End Function

    Public Function UnionShapeFiles(ShapeFile1 As Shapefile, ShapeFile2 As Shapefile) As Shapefile
        Return ShapeFile1.Union(False, ShapeFile2, False)
    End Function

    Public Function getShapeFieldIdxFromFileName(ByVal Path As String, ByVal FieldName As String) As Integer
        Dim sf = New MapWinGIS.Shapefile
        Dim i As Long

        Try
            If FieldName.Trim = "" Then
                Me.setup.Log.AddMessage("One or empty fieldnames specified for shapefile " & Path & ".")
                Return -1
            Else
                If System.IO.File.Exists(Path) Then
                    If sf.Open(Path) Then
                        For i = 0 To sf.NumFields - 1
                            If sf.Field(i).Name.Trim.ToLower = FieldName.Trim.ToLower Then
                                Return i
                            End If
                        Next
                        Me.setup.Log.AddWarning("Could not find shapefield " & FieldName & " in shapefile " & Path)
                        sf.Close()
                        Return -1
                    Else
                        Me.setup.Log.AddError("Could not read shapefile " & sf.Filename & ".")
                        Return -1
                    End If
                Else
                    Me.setup.Log.AddError("Shapefile does not exist: " & Path)
                    Return -1
                End If
            End If

        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return -1
        End Try

    End Function


    Public Function getShapeFieldTypeFromFileName(ByVal Path As String, ByVal FieldName As String) As MapWinGIS.FieldType
        Dim sf = New MapWinGIS.Shapefile
        Dim i As Long

        Try
            If System.IO.File.Exists(Path) Then
                If sf.Open(Path) Then
                    For i = 0 To sf.NumFields - 1
                        If sf.Field(i).Name.Trim.ToLower = FieldName.Trim.ToLower Then
                            Return sf.Field(i).Type
                        End If
                    Next
                    Me.setup.Log.AddError("Could not find shapefield " & FieldName & " in shapefile " & Path)
                    sf.Close()
                    Return Nothing
                Else
                    Me.setup.Log.AddError("Could not read shapefile " & sf.Filename & ".")
                    Return Nothing
                End If
            Else
                Me.setup.Log.AddError("Shapefile does not exist: " & Path)
                Return Nothing
            End If
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Function getShapeFieldIdxFromShapeFile(ByVal sf As MapWinGIS.Shapefile, ByVal FieldName As String) As Integer
        Dim i As Long

        Try
            For i = 0 To sf.NumFields - 1
                If sf.Field(i).Name.Trim.ToLower = FieldName.Trim.ToLower Then
                    Return i
                End If
            Next
            Me.setup.Log.AddError("Could not find shapefield " & FieldName & " in shapefile " & sf.Filename)
            Return -1
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return -1
        End Try

    End Function

    Public Function GetNearestShapeDistance(ByVal X As Double, ByVal Y As Double, ByRef mySF As MapWinGIS.Shapefile, ByRef ShapeIdx As Long, ByRef Chainage As Double, ByRef Distance As Double, ByVal Criterium As Double) As Boolean

        'calculates the closest distance from a point (X,Y) to the nearest shape in a shapefile
        'it does this by creating a circle and intersecting it with the shapefile
        'it adjusts the circle's radius until the smallest possible radius with intersection is reached
        Dim Done As Boolean = False
        Dim Radius As Double, Dist As Double, iShape As Long, myShape As MapWinGIS.Shape
        Dim ptShape As MapWinGIS.Shape, bufShape As MapWinGIS.Shape
        Dim MaxRadius As Double, MinRadius As Double, Intersects As Boolean

        'find the largest possible radius that intersects with the shapefile
        Radius = Me.setup.GeneralFunctions.Pythagoras(X, Y, mySF.Extents.xMin, mySF.Extents.yMin)
        Dist = Me.setup.GeneralFunctions.Pythagoras(X, Y, mySF.Extents.xMin, mySF.Extents.yMax)
        If Dist > Radius Then Radius = Dist
        Dist = Me.setup.GeneralFunctions.Pythagoras(X, Y, mySF.Extents.xMax, mySF.Extents.yMin)
        If Dist > Radius Then Radius = Dist
        Dist = Me.setup.GeneralFunctions.Pythagoras(X, Y, mySF.Extents.xMax, mySF.Extents.yMax)
        If Dist > Radius Then Radius = Dist

        'create a shape from the point
        ptShape = New MapWinGIS.Shape
        ptShape.Create(MapWinGIS.ShpfileType.SHP_POINT)
        ptShape.AddPoint(X, Y)

        'initialize lastRadius
        MaxRadius = Radius
        MinRadius = 0

        While Not Done

            'create a buffer around the shape
            bufShape = New MapWinGIS.Shape
            bufShape = ptShape.Buffer(Radius, 100)

            'check if it intersects 
            Intersects = False
            For iShape = 0 To mySF.NumShapes - 1
                myShape = mySF.Shape(iShape)
                If myShape.Intersects(bufShape) Then
                    Intersects = True
                    Exit For
                End If
            Next

            If Intersects Then
                MaxRadius = Radius
            Else
                MinRadius = Radius
            End If

            'calculate the new radius
            Radius = (MaxRadius + MinRadius) / 2
            If MaxRadius - MinRadius < Criterium Then Done = True

        End While

        Return Radius

    End Function

    Public Function MakeUniformIntField(ByVal shpPath As String, ByVal FieldName As String, ByVal FieldVal As Integer) As Integer
        'this function creates a new shapefile field of the Integer type 
        'and fills allrecords with the specified uniform value
        Dim sf As New MapWinGIS.Shapefile, iField As Long, iShape As Long, Found As Boolean = False
        Dim newField As New MapWinGIS.Field, FieldIdx As Long
        Try
            If System.IO.File.Exists(shpPath) Then

                If Not sf.Open(shpPath) Then Throw New Exception("Could not open shapefile " & shpPath)
                If Not sf.StartEditingTable Then Throw New Exception("Could not edit shapfile table " & shpPath)

                'remove the existing field of the same name (if found)
                For iField = 0 To sf.NumFields - 1
                    If sf.Field(iField).Name.ToLower = FieldName.ToLower Then
                        sf.EditDeleteField(iField)
                        Exit For
                    End If
                Next

                'create a new field of the required name
                FieldIdx = sf.EditAddField(FieldName, MapWinGIS.FieldType.INTEGER_FIELD, 10, 10)
                For iShape = 0 To sf.NumShapes - 1
                    sf.EditCellValue(FieldIdx, iShape, FieldVal)
                Next

                sf.StopEditingTable(True)
                sf.Close()
            End If

            Return FieldIdx

        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return 0
        End Try

    End Function


    Public Function getShapeFileType(ByVal myShapeFilePath As String, ByVal myType As MapWinGIS.ShpfileType) As Boolean
        Try
            Dim mw As New MapWinGIS.Shapefile
            If Not mw.Open(myShapeFilePath) Then Throw New Exception
            myType = mw.ShapefileType
            mw.Close()
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError("Shapefile could not be opened: " & myShapeFilePath)
            Return False
        End Try
    End Function

End Class


Option Explicit On

Imports METEOBAS.General
Imports System.IO
Imports MapWinGIS

Public Class clsSubcatchmentShapeFile
    Public PolySF As clsPolyShapeFile

    Public SubcatchmentIDField As String
    Public SubcatchmentIDFieldIdx As Integer = -1
    Public SubcatchmentNameField As String
    Public SubcatchmentNameFieldIdx As Integer = -1
    Public STRIDOutField As String

    'catchment and subcatchment data
    Public CatchmentIDField As String
    Public CatchmentNameField As String
    Public ZPField As String
    Public WPField As String
    Public InundationLevelField As String
    Public AfvCoefField As String
    Public EmergencyStopElevationField As String
    Public SelectionField As String

    Public CatchmentNameFieldIdx As Integer = -1
    Public STRIDOutFieldIdx As Integer = -1
    Public CatchmentIDFieldIdx As Integer = -1
    Public ZPFieldIdx As Integer = -1
    Public WPFieldIdx As Integer = -1
    Public InundationLevelFieldIdx As Integer = -1
    Public AfvCoefFieldIdx As Integer = -1
    Public EmergencyStopElevationFieldIdx As Integer = -1
    Public SelectionFieldIdx As Integer = -1

    'rainfall runoff data
    Public RunoffField As String
    Public InfiltrationField As String
    Public Alpha1Field As String
    Public Alpha2Field As String
    Public Alpha3Field As String
    Public Alpha4Field As String
    Public Depth1Field As String
    Public Depth2Field As String
    Public Depth3Field As String

    Public RunoffFieldIdx As Integer = -1
    Public InfiltrationFieldIdx As Integer = -1
    Public Alpha1FieldIdx As Integer = -1
    Public Alpha2FieldIdx As Integer = -1
    Public Alpha3FieldIdx As Integer = -1
    Public Alpha4FieldIdx As Integer = -1
    Public Depth1FieldIdx As Integer = -1
    Public Depth2FieldIdx As Integer = -1
    Public Depth3FieldIdx As Integer = -1

    'sewage area data
    Public SewageAreaIdField As String
    Public WWTPIDField As String
    Public StorageField As String
    Public POCField As String
    Public PavedAreaField As String 'the field that contains the (user defined) paved area as attribute value

    Public SewageAreaIdFieldIdx As Integer = -1
    Public WWTPIDFieldIdx As Integer = -1
    Public StorageFieldIdx As Integer = -1
    Public POCFieldIdx As Integer = -1
    Public PavedAreaFieldIdx As Integer = -1

    Friend TotalShape As New Shape                              'de hele shapefile samengevoegd tot één shape
    Private SelectionOperator As String
    Private SelectionOperand As Object
    Private setup As clsSetup

    Friend Sub New(ByRef mySetup As clsSetup, ByVal Path As String)
        Me.setup = mySetup
        PolySF = New clsPolyShapeFile(Me.setup, Path)
    End Sub

    Friend Sub New(ByRef mySetup As clsSetup)
        Me.setup = mySetup
        PolySF = New clsPolyShapeFile(Me.setup)
    End Sub

    Friend Sub Initialize(ByVal Path As String)
        PolySF = New clsPolyShapeFile(Me.setup, Path)
    End Sub

    Friend Sub Initialize()
        PolySF = New clsPolyShapeFile(Me.setup)
    End Sub

    Public Function Fix() As Boolean
        Dim fixedSf As New MapWinGIS.Shapefile
        If Not PolySF.sf.FixUpShapes(fixedSf) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function Fixup(newPath As String) As Boolean
        Try
            If Not PolySF.sf.Open(PolySF.Path) Then Throw New Exception("Error reading subcatchment shapefile")
            Dim fixSF As MapWinGIS.Shapefile = Nothing
            If Not PolySF.sf.FixUpShapes2(False, fixSF) Then Throw New Exception("Error fixing subcatchment shapefile: original shapefile still in use.")
            PolySF.sf = fixSF
            If Not fixSF.SaveAs(newPath) Then Throw New Exception("Error saving fixed subcatchment shapefile: original shapefile still in use.")
            PolySF.Path = newPath
            PolySF.Close()
            Return True
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function setPath(ByVal myPath As String) As Boolean
        If System.IO.File.Exists(myPath) Then
            PolySF.Path = myPath
            Return True
        Else
            Return False
        End If
    End Function

    Public Function getTargetLevelsByPoint(ByVal myPoint As MapWinGIS.Point, ByRef WP As Double, ByRef ZP As Double) As Double
        Dim Utils As New MapWinGIS.Utils, i As Long, myShape As MapWinGIS.Shape
        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If Utils.PointInPolygon(myShape, myPoint) Then
                ZP = PolySF.sf.CellValue(ZPFieldIdx, i)
                WP = PolySF.sf.CellValue(WPFieldIdx, i)
            End If
        Next
        Return Nothing
    End Function

    Public Function getSummerTargetLevelByPoint(ByVal myPoint As MapWinGIS.Point) As Double
        Dim Utils As New MapWinGIS.Utils, i As Long, myShape As MapWinGIS.Shape
        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If Utils.PointInPolygon(myShape, myPoint) Then
                Return PolySF.sf.CellValue(ZPFieldIdx, i)
            End If
        Next
        Return Nothing
    End Function

    Public Function getWinterTargetLevelByPoint(ByVal myPoint As MapWinGIS.Point) As Double
        Dim Utils As New MapWinGIS.Utils, i As Long, myShape As MapWinGIS.Shape
        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If Utils.PointInPolygon(myShape, myPoint) Then
                Return PolySF.sf.CellValue(WPFieldIdx, i)
            End If
        Next
        Return Nothing
    End Function

    Public Function getAreaIDByPoint(ByVal myPoint As MapWinGIS.Point) As String
        Dim Utils As New MapWinGIS.Utils, i As Long, myShape As MapWinGIS.Shape
        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If Utils.PointInPolygon(myShape, myPoint) Then
                Return PolySF.sf.CellValue(SubcatchmentIDFieldIdx, i)
            End If
        Next
        Return Nothing
    End Function

    Public Function getAreaIDByIdx(ByVal ShapeIdx As Long) As String
        Return PolySF.sf.CellValue(SubcatchmentIDFieldIdx, ShapeIdx)
    End Function

    Public Function getCatchmentIDByPoint(ByVal myPoint As MapWinGIS.Point) As String
        Dim Utils As New MapWinGIS.Utils, i As Long, myShape As MapWinGIS.Shape
        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If Utils.PointInPolygon(myShape, myPoint) Then
                Return PolySF.sf.CellValue(CatchmentIDFieldIdx, i)
            End If
        Next
        Return Nothing
    End Function

    Public Function setSubcatchmentIDField(ByVal FieldName As String) As Boolean
        SubcatchmentIDField = FieldName
        subcatchmentIDFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If subcatchmentIDFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setAreaNameField(ByVal FieldName As String) As Boolean
        SubcatchmentNameField = FieldName
        SubcatchmentNameFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If subcatchmentIDFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setCatchmentIDField(ByVal FieldName As String) As Boolean
        CatchmentIDField = FieldName
        CatchmentIDFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If CatchmentIDFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setCatchmentNameField(ByVal FieldName As String) As Boolean
        CatchmentNameField = FieldName
        CatchmentNameFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If CatchmentIDFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setKWKOutField(ByVal FieldName As String) As Boolean
        If FieldName.Trim = "" Then Return False
        STRIDOutField = FieldName
        STRIDOutFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If STRIDOutFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setWPField(ByVal FieldName As String) As Boolean
        WPField = FieldName
        WPFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If WPFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setInundationLevelField(ByVal FieldName As String) As Boolean
        InundationLevelField = FieldName
        InundationLevelFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If InundationLevelFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setSelectionField(ByVal FieldName As String) As Boolean
        SelectionField = FieldName
        SelectionFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If SelectionFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setSelectionOperator(ByVal myString As String) As Boolean
        SelectionOperator = myString
        Return True
    End Function

    Public Function setSelectionOperand(ByVal myOperand As Object) As Boolean
        SelectionOperand = myOperand
        Return True
    End Function

    Public Function setZPField(ByVal FieldName As String) As Boolean
        ZPField = FieldName
        ZPFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If ZPFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setWWTPIDField(ByVal FieldName As String) As Boolean
        WWTPIDField = FieldName
        WWTPIDFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If WWTPIDFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setRunoffField(ByVal FieldName As String) As Boolean
        RunoffField = FieldName
        RunoffFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If RunoffFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setInfiltrationField(ByVal FieldName As String) As Boolean
        InfiltrationField = FieldName
        InfiltrationFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If InfiltrationFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setAlpha1Field(ByVal FieldName As String) As Boolean
        Alpha1Field = FieldName
        Alpha1FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Alpha1FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setAlpha2Field(ByVal FieldName As String) As Boolean
        Alpha2Field = FieldName
        Alpha2FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Alpha2FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setAlpha3Field(ByVal FieldName As String) As Boolean
        Alpha3Field = FieldName
        Alpha3FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Alpha3FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setAlpha4Field(ByVal FieldName As String) As Boolean
        Alpha4Field = FieldName
        Alpha4FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Alpha4FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setDepth1Field(ByVal FieldName As String) As Boolean
        Depth1Field = FieldName
        Depth1FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Depth1FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setDepth2Field(ByVal FieldName As String) As Boolean
        Depth2Field = FieldName
        Depth2FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Depth2FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setDepth3Field(ByVal FieldName As String) As Boolean
        Depth3Field = FieldName
        Depth3FieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)
        If Depth3FieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Function setEmergencyStopElevationField(ByVal FieldName As String) As Boolean
        EmergencyStopElevationField = FieldName
        EmergencyStopElevationFieldIdx = setup.GISData.getShapeFieldIdxFromFileName(PolySF.Path, FieldName)

        If EmergencyStopElevationFieldIdx >= 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function setAreaIDFieldIdx(ByVal ShapeField As String) As Boolean
        Dim tmpField As String = ShapeField.Trim.ToUpper, i As Long
        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = ShapeField Then
                SubcatchmentIDFieldIdx = i
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetShapeByAreaID(ByVal ID As String) As MapWinGIS.Shape
        Dim i As Long

        Try
            If Not PolySF.Open() Then Throw New Exception("Error opening shapefile.")
            For i = 0 To PolySF.sf.NumShapes - 1
                If PolySF.sf.CellValue(SubcatchmentIDFieldIdx, i).ToString.Trim.ToUpper = ID.Trim.ToUpper Then
                    Return PolySF.sf.Shape(i)
                End If
            Next
            Return Nothing
            PolySF.Close()
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return Nothing
        End Try

    End Function



    Friend Function findAreaIDFieldIdx(ByVal ShapeField As String) As Boolean
        'omdat dit een vrij specifiek veld is, permitteren we ons wat vrijheid bij de zoektocht
        Dim tmpField As String = ShapeField.Trim.ToUpper
        Dim i As Long

        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = ShapeField.Trim.ToUpper Then
                SubcatchmentIDFieldIdx = i
                Return True
            End If
        Next

        'niet gevonden, dus zoek andere voordehandliggende opties
        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = "GPGIDENT" Then
                SubcatchmentIDFieldIdx = i
                Return True
            End If
        Next

        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = "GFEIDENT" Then
                SubcatchmentIDFieldIdx = i
                Return True
            End If
        Next

        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = "GAFIDENT" Then
                SubcatchmentIDFieldIdx = i
                Return True
            End If
        Next

        Return False

    End Function

    Public Function getInundationLevel(ByVal Shapeidx As Long) As Double
        Return Me.PolySF.sf.CellValue(InundationLevelFieldIdx, Shapeidx)
    End Function

    Public Function getSummerTargetLevel(ByVal shapeIdx As Long) As Double
        Return Me.PolySF.sf.CellValue(ZPFieldIdx, shapeIdx)
    End Function

    Public Function getWinterTargetLevel(ByVal shapeIdx As Long) As Double
        Return Me.PolySF.sf.CellValue(WPFieldIdx, shapeIdx)
    End Function

    Public Function getSelectedStatus(ByVal ShapeIdx As Long) As Boolean
        Dim myNum As Double = 0, myStr As String = "", Numeric As Boolean

        Try
            Numeric = IsNumeric(Me.PolySF.sf.CellValue(SelectionFieldIdx, ShapeIdx))
            If Numeric Then
                myNum = Me.PolySF.sf.CellValue(SelectionFieldIdx, ShapeIdx)
            Else
                myStr = Me.PolySF.sf.CellValue(SelectionFieldIdx, ShapeIdx).ToString.Trim.ToUpper
            End If

            If Numeric Then
                Select Case SelectionOperator
                    Case Is = ">"
                        Return myNum > SelectionOperand
                    Case Is = "<"
                        Return myNum < SelectionOperand
                    Case Is = ">="
                        Return myNum >= SelectionOperand
                    Case Is = "<="
                        Return myNum <= SelectionOperand
                    Case Is = "="
                        Return myNum = SelectionOperand
                    Case Else
                        Me.setup.Log.AddError("Error: invalid operator for numeric field selection " & SelectionOperator & ".")
                        Throw New Exception("Error in sub getSelectedStatus in class clsGebiedenShapeFile.")
                End Select
            Else
                Select Case SelectionOperator
                    Case Is = "IS"
                        Return myNum = SelectionOperand
                    Case Is = "NOT"
                        Return myNum <> SelectionOperand
                    Case Else
                        Me.setup.Log.AddError("Error: invalid operator for string field selection " & SelectionOperator & ".")
                        Throw New Exception("Error in sub getSelectedStatus in class clsGebiedenShapeFile.")
                End Select
            End If

            Me.setup.Log.AddError("Unsupported selection string encountered: " & SelectionOperator)
            Return False
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try

    End Function

    Public Function GetSelectedShapeIdx() As Long
        'finds out if a shape has been selected by the user and returns the corresponding index number
        'if no selected shape found, it returns -1
        Dim i As Long
        PolySF.Open()
        For i = 0 To PolySF.sf.NumShapes - 1
            If PolySF.sf.ShapeSelected(i) Then Return i
        Next
        Return -1
    End Function

    Public Sub ExportTotalShape(ByVal FileName As String)
        Dim ShapeIdx As Long, FieldIdx As Long
        Try
            Dim TotalSF As New MapWinGIS.Shapefile
            If Not TotalSF.CreateNew(FileName, MapWinGIS.ShpfileType.SHP_POLYGON) Then Throw New Exception("Could not create shapefile for merged shapes.")
            If Not TotalSF.StartEditingShapes(True) Then Throw New Exception("Could not start editing shapes.")
            FieldIdx = TotalSF.EditAddField("ID", MapWinGIS.FieldType.STRING_FIELD, 10, 10)
            ShapeIdx = TotalSF.EditAddShape(TotalShape)
            If Not TotalSF.StopEditingShapes(True, True) Then Throw New Exception("Could not stop editing merged shapefile.")
            TotalSF.Save()
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
        End Try

    End Sub


    Public Function IDsUnique() As Boolean

        'bepaalt of de Area Shapefile uitsluitend unieke Area ID's bevat
        Dim myIDs As New Dictionary(Of String, String)
        Dim myID As String, AllUnique As Boolean = True
        Dim i As Long

        If Open() Then
            For i = 0 To Me.PolySF.sf.NumShapes - 1
                myID = PolySF.sf.CellValue(SubcatchmentIDFieldIdx, i).ToString
                If myIDs.ContainsKey(myID.Trim.ToUpper) Then
                    Me.setup.Log.AddError("Multiple instances for Area ID " & myID & " in area shapefile.")
                    AllUnique = False
                Else
                    myIDs.Add(myID.Trim.ToUpper, myID)
                End If
            Next
        End If
        Return AllUnique
    End Function

    ''' <summary>
    ''' Deze subroutine voegt shapes behorende bij hetzelfde catchment samen tot één nieuwe shape per Catchment
    ''' </summary>
    ''' <remarks></remarks>

    Friend Function getUnderlyingShapeIdx(ByVal X As Double, ByVal Y As Double) As Integer
        'geeft voor een gegeven XY-coordinaat het indexnummer van de onderliggende shape terug
        Dim myUtils As New MapWinGIS.Utils
        Dim myShape As MapWinGIS.Shape
        Dim myPoint As New MapWinGIS.Point
        myPoint.x = X
        myPoint.y = Y
        Dim i As Long

        For i = 0 To PolySF.sf.NumShapes - 1
            myShape = PolySF.sf.Shape(i)
            If myUtils.PointInPolygon(myShape, myPoint) Then
                Return i
            End If
        Next
        Return 0

    End Function

    Public Function Open() As Boolean
        Try
            If Me.PolySF.Path = "" Then
                Me.setup.Log.AddWarning("Area shapefile is already open or path is empty.")
            Else
                If Not Me.PolySF.sf.Open(PolySF.Path) Then Throw New Exception("Could not open Area Shapefile.")
            End If
            Return True

        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Sub Close()
        Me.PolySF.sf.Close()
    End Sub

    Public Function GetShapeIdxByCoord(ByVal X As Double, ByVal Y As Double) As Long
        Try
            Dim i As Long, myUtils As New MapWinGIS.Utils
            Dim newPoint As New MapWinGIS.Point
            newPoint.x = X
            newPoint.y = Y
            If Not PolySF.Open Then Throw New Exception("Could not open shapefile: " & PolySF.sf.Filename)
            For i = 0 To PolySF.sf.NumShapes - 1
                If myUtils.PointInPolygon(PolySF.sf.Shape(i), newPoint) Then Return i
            Next
            PolySF.Close()
            Return -1
        Catch ex As Exception
            Me.setup.Log.AddError(ex.Message)
            Return -1
        End Try
    End Function

    Public Function MergeAllShapes() As Boolean
        'this routine merges all shapes inside a shapefile into one TotalShape
        Dim Shape1 As New MapWinGIS.Shape, Shape2 As New MapWinGIS.Shape
        Dim utils As New MapWinGIS.Utils()
        Dim i As Long

        Try
            'hanteer een UNION om alle shapes in één samen te voegen
            If Not Open() Then Throw New Exception("Could not open area shapefile.")

            'read the number of shapes in this shapefile
            Dim numShapes = Me.PolySF.sf.NumShapes
            If numShapes = 0 Then Throw New Exception("Shapefile was emtpy.")

            'read the first shape
            TotalShape = Me.PolySF.sf.Shape(0)
            If Not TotalShape.IsValid Then TotalShape.FixUp(TotalShape)
            If Not TotalShape.IsValid Then Throw New Exception("Error in shapefile: first shape is corrupt and could not be fixed.")

            'read the shapefile, starting with the second shape and merge with totalshape
            If numShapes > 1 Then
                For i = 1 To numShapes - 1
                    Console.WriteLine("Processing shape " & i + 1 & " from " & numShapes)

                    'set shape 1
                    Shape1 = TotalShape

                    'set shape 2 and attempt to fix if not valid
                    Shape2 = PolySF.sf.Shape(i)
                    If Not Shape2.IsValid Then Shape2.FixUp(Shape2)

                    'only if valid, merge with shape 1
                    If Shape2.IsValid Then
                        TotalShape = utils.ClipPolygon(MapWinGIS.PolygonOperation.UNION_OPERATION, Shape1, Shape2)
                        If Not TotalShape.IsValid Then
                            TotalShape.FixUp(TotalShape) 'als dit leidt tot een corrupte shape, probeer te fiksen
                            If Not TotalShape.IsValid Then TotalShape = Shape1 'als het fixen niet gelukt is, val terug op de oorspronkelijke totalshape, van voor de union
                        End If
                    End If

                Next i
            End If

            Return True
        Catch ex As Exception
            Dim log As String = "Error in MergeAllShapes"
            Me.setup.Log.AddError(log + ": " + ex.Message)
            Return False
        Finally
            'If Not Shape1 Is Nothing Then Me.setup.GeneralFunctions.ReleaseComObject(Shape1, False) releasing this one screws up the TotalShape
            If Not Shape2 Is Nothing Then Me.setup.GeneralFunctions.ReleaseComObject(Shape2, False)
            Me.setup.GeneralFunctions.ReleaseComObject(utils, True)
        End Try
    End Function


    Public Function GetCatchmentList() As Dictionary(Of String, String)
        'this function creates a list of all unique catchment ID's in the area shapefile
        Dim i As Long, myList As New Dictionary(Of String, String), myID As String
        For i = 0 To PolySF.sf.NumShapes
            myID = PolySF.sf.CellValue(CatchmentIDFieldIdx, i)
            If Not myList.ContainsKey(myID.Trim.ToUpper) Then
                myList.Add(myID.Trim.ToUpper, myID)
            End If
        Next
        Return myList
    End Function

    Public Function GetFieldIdx(ByVal NAME As String) As Integer
        Dim i As Long
        For i = 0 To PolySF.sf.NumFields - 1
            If PolySF.sf.Field(i).Name.Trim.ToUpper = NAME.Trim.ToUpper Then Return i
        Next
        Return -1
    End Function

End Class

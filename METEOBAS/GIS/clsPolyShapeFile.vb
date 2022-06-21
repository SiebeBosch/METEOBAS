Imports METEOBAS.General
Imports GemBox.Spreadsheet

Public Class clsPolyShapeFile

    Public Path As String
    Public sf As New MapWinGIS.Shapefile

    Public ValueField As String
    Public ValueFieldIdx As Integer = -1

    Private Setup As clsSetup

    Public Sub New(ByRef mySetup As clsSetup, ByVal myPath As String)
        Setup = mySetup
        Path = myPath
    End Sub

    Public Sub New(ByRef mySetup As clsSetup)
        Setup = mySetup
    End Sub

    Public Function SetPath(myPath As String) As Boolean
        Try
            Path = myPath
            Return True
        Catch ex As Exception
            Me.Setup.Log.AddError(ex.Message)
            Return False
        End Try
    End Function

    Public Function Open() As Boolean
        If sf.Open(Path) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Close() As Boolean
        If sf.Close Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function setValueField(ByVal FieldName As String) As Boolean
        ValueField = FieldName
        ValueFieldIdx = Setup.GISData.getShapeFieldIdxFromFileName(Path, FieldName)
        If ValueFieldIdx >= 0 Then Return True Else Return False
    End Function

    Public Sub ExportToTableau(ByVal FieldIdx As Integer, ByVal FileName As String)
        SpreadsheetInfo.SetLicense("EVIG-1Y89-FYME-DPUJ")

        Dim myShape As MapWinGIS.Shape, myPoint As MapWinGIS.Point
        Dim iShape As Long, iPoint As Long

        Dim oExcel As ExcelFile
        Dim worksheets As ExcelWorksheetCollection
        Dim ws As ExcelWorksheet, r As Long, c As Long

        'ieder stroomgebied krijgt z'n eigen Excel-file; dit omwille van het geheugengebruik
        'schrijf (omwille van de rekensnelheid) eerst de tijdstappen, en daarna pas de data
        oExcel = New ExcelFile
        worksheets = oExcel.Worksheets

        r = 0
        c = 0
        ws = worksheets.Add("Polygons")
        ws.Cells(r, c).Value = "ID"
        ws.Cells(r, c + 1).Value = "Lat"
        ws.Cells(r, c + 2).Value = "Lon"
        ws.Cells(r, c + 3).Value = "PointIdx"
        ws.Cells(r, c + 4).Value = "PolyIdx"

        For iShape = 0 To sf.NumShapes - 1
            myShape = sf.Shape(iShape)
            For iPoint = 0 To myShape.numPoints - 1
                myPoint = myShape.Point(iPoint)
                r += 1
                ws.Cells(r, c).Value = sf.CellValue(FieldIdx, iShape)
                ws.Cells(r, c + 1).Value = myPoint.y
                ws.Cells(r, c + 2).Value = myPoint.x
                ws.Cells(r, c + 3).Value = iPoint + 1
                ws.Cells(r, c + 4).Value = iShape + 1
            Next
        Next

        oExcel.Save(Me.Setup.Settings.ExportDirRoot & "\" & FileName, SaveOptions.XlsxDefault)

    End Sub

End Class

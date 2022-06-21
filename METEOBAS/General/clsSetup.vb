Option Explicit On
Imports System.Globalization
Imports MapWinGIS

Namespace General
    Public Class clsSetup
        Public GISData As clsGISData            'bevat alle shapefiles en rasters!
        Public KNMIData As clsKNMIData          'bevat alle data van het KNMI
        Public MBRasterData As clsMBRasterData  'bevat rasters uit Meteobase
        Public Settings As clsSettings          'de algemene instellingen

        Public GeneralFunctions As GeneralFunctions      'algemene functies

        Public ExcelFile As clsExcelBook        'een workbook in Excel voor resultaten
        Public Log As New clsLog                'de logfile
        Public EMail As clsEmail                'e-mail

        Friend importDir As String              'export directory voor resultaten van bewerkingen

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

            'MsgBox("Hoera, we zitten in de DLL (juiste versie)")

            ' Set dot as decimal point:
            Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture

            ' Initializeren classes:
            KNMIData = New clsKNMIData(Me)      'bevat alle data van het KNMI
            GISData = New clsGISData(Me)

            Settings = New clsSettings()      'de algemene instellingen
            ExcelFile = New clsExcelBook()
            GeneralFunctions = New GeneralFunctions(Me)  'algemene functies
            EMail = New clsEmail(Me)

        End Sub

        Public Sub PassAreaShape(ByVal sfArea As MapWinGIS.Shapefile)
            Me.GISData.SubcatchmentShapeFile.Initialize(sfArea.Filename)
            Me.GISData.SubcatchmentShapeFile.PolySF.sf = sfArea
        End Sub

        Public Sub SetImportDir(ByVal myImportDir As String)
            Me.importDir = myImportDir
        End Sub
        Public Sub addError(ByVal myError As String)
            Me.Log.AddError(myError)
        End Sub
        Public Sub addWarning(ByVal myWarning As String)
            Me.Log.AddWarning(myWarning)
        End Sub

        Public Function ReadMeteoBasePrecHourly(ByVal myPath As String) As Boolean
            KNMIData = New clsKNMIData(Me)
            Call KNMIData.readMBTextFile(myPath)
            Return True
        End Function

        Public Function ExportMeteoBasePrecHourly() As Boolean
            If KNMIData.writeMBTextFiles(Me.importDir) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function CalcMeteoBaseHourlyStats(ByVal myArea As Double) As Boolean
            If KNMIData.calcHourlyEventStats(myArea) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Sub WriteExcelFile(ByVal myPath As String)
            Me.ExcelFile.Path = myPath
            Call Me.ExcelFile.Save()
        End Sub

    End Class
End Namespace
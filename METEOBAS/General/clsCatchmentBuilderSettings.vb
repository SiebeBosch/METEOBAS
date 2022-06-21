Public Class clsCatchmentBuilderSettings

    Public ExportTemporaryGISFiles As Boolean          'a setting that allows you to save all intermediate (temporary) GIS files
    Public MinOwPerc As Double                         'minimum percentage openwater
    Public MinOWPercIncrease As Double                 'percentage openwater increase per meter
    Public MinRRNodeArea As Double                     'minimum area for any RR node
    Public SurfaceLevelPercentage As Integer           'the percentage of the surface level that will be used to determine the surface level in the rainfall runoff model
    Public OverruleEmergencyStopAtInundationPercentage 'the percentage of polder inundation at which an emergency stop will be overruled no matter what

    Public StorageFromLowestAreaOnly As Boolean = True 'als TRUE, wordt de bergingscapaciteit in een bemalen of gestuwd dummytakje alleen bepaald op basis van het laagste inliggende peilvak

    Public UseLGN As Boolean                            'keeps track if the LGN should be used for landuse analysis
    Public LGNVersion As Integer                        'the LGN version number (5 or 6)

    Public ErnstDefsFromShapefile As Boolean           'whether the ERNST drainage definitions are retrieved from the area shapefile (if TRUE) or from the definitions table (if FALSE)
    Public UseSnappingPointsShapeFile As Boolean       'whether or not to use a snapping points shapefile to snap the areas to their reach

    Public AddToExistingModel As Boolean = True        'debugsetting. Zet op true vóór compileren
    Public WriteStorageTable As Boolean = True         'debugsetting. Zet op true vóór compileren

    Public AllowDummyMerge As Boolean                  'allow dummyreaches to be merged in case they link to the same location on the backbone
    Public UpStreamMergeRadius As Integer              'radius to search in upstream direction for connection nodes to merge with
    Public DownStreamMergeRadius As Integer            'radius to search in downstream direction for connectoin nodes to merge with

    Public AllStructuresController As Boolean          'implement a hydraulic controller for overwriting by Delft FEWS
    Public SnappingInsideCatchment As Boolean          'prefer snapping to a merged shape covering an entire catchment over snapping to the closest sobek reach
    Public SnappingInsideShapeFile As Boolean

    Private Database As String                          'path to the access database that holds all catchment and area data

    Public WinSumStartDay As Integer
    Public WinSumStartMonth As Integer
    Public WinSumEndDay As Integer
    Public WinSumEndMonth As Integer
    Public SumWinStartDay As Integer
    Public SumWinStartMonth As Integer
    Public SumWinEndDay As Integer
    Public SumWinEndMonth As Integer

    Public InletCapacity As Double                    'inlet capacity in mm/d
    Public SummerFlushCap As Double                   'summer flush capacity in mm/d
    Public FlushStartDay As Integer
    Public FlushStartMonth As Integer
    Public FlushStopDay As Integer
    Public FlushStopMonth As Integer

    Public CSOPrefix As String                        'prefix for CSO-locations (riooloverstorten)

    Public ReachConnectionPrefix As String = "c"

    Public ProfileStorPrefix As String = "pStor"
    Public ProfileOutPrefix As String = "pOut"
    Public ProfileOut2Prefix As String = "pOut2"
    Public ProfileInPrefix As String = "pIn"
    Public ProfileIn2Prefix As String = "pIn2"

    Public ReachStorPrefix As String = "rStor"
    Public ReachOutPrefix As String = "rOut"
    Public ReachOut2Prefix As String = "rOut2"
    Public ReachInPrefix As String = "rIn"
    Public ReachIn2Prefix As String = "rIn2"

    Public ConnNodePrefix As String = "cn"
    Public StorNodePrefix As String = "stor"
    Public MeasPrefix As String = "meas"
    Public LatStorPrefix As String = "latstor"

    Public FlowLimiterStructurePrefix As String = "lim"

    Public MaalstopPrefix As String = "ms"
    Public ReachMSPrefix As String = "ms1"
    Public ReachMS2Prefix As String = "ms2"
    Public ProfileMS1Prefix As String = "pMs1"
    Public ProfileMS2Prefix As String = "pMs2"
    Public MS1Prefix As String = "ms1"
    Public MS2Prefix As String = "ms2"

    Public UnpavedPrefix As String = "up"
    Public PavedPrefix As String = "pv"
    Public GreenhousePrefix As String = "gr"
    Public RRCFPavedPrefix As String = "rrcfPv"
    Public RRCFUnpavedPrefix As String = "rrcfUp"
    Public RRCFGreenhousePrefix As String = "rrcfGr"
    Public rrLinkPrefix As String = "brch"
    Public rrSewageLinkPrefix As String = "rwzi"
    Public InlaatPrefix As String = "i"
    Public FlushPrefix As String = "f"


End Class

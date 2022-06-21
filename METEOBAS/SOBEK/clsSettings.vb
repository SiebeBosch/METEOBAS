Option Explicit On

''' <summary>
''' Geen constructor nodig
''' </summary>
''' <remarks></remarks>
Public Class clsSettings
    'TODO: Converteer naar struct
    Public ExportDirRoot As String
    Public ExportDirLogs As String

    Public Sub addNetworkSTRecord()
        Stop
    End Sub

    Public Sub SetExportDirs(ByVal myPath As String, ByVal MakeSobekSubDirs As Boolean, ByVal ClearAll As Boolean)
        'this sub sets all export directories for Topology, Flow Data and RR Data
        Dim Files As String(), myFile As String

        ExportDirRoot = myPath
        If Not System.IO.Directory.Exists(ExportDirRoot) Then System.IO.Directory.CreateDirectory(ExportDirRoot)
        If ClearAll Then
            Files = System.IO.Directory.GetFiles(ExportDirRoot)
            For Each myFile In Files
                System.IO.File.Delete(myFile)
            Next
        End If

        If MakeSobekSubDirs Then
            ExportDirLogs = myPath & "\Logs"
            If Not System.IO.Directory.Exists(ExportDirLogs) Then System.IO.Directory.CreateDirectory(ExportDirLogs)

        End If

    End Sub


End Class

Imports GemBox.Spreadsheet

Public Class clsExcelBook
    Public Path As String
    Public Sheets As Collection 'of clsExcelBook
    Dim oExcel As ExcelFile


    Public Sub New()
        ' TODO: If using GemBox.Spreadsheet Professional, put your serial key below.
        ' Otherwise, if you are using GemBox.Spreadsheet Free, comment out the 
        ' following line (Free version doesn't have SetLicense method). 
        SpreadsheetInfo.SetLicense("EVIG-1Y89-FYME-DPUJ")
        oExcel = New ExcelFile
        Sheets = New Collection
    End Sub

    Public Sub New(myPath As String)
        ' TODO: If using GemBox.Spreadsheet Professional, put your serial key below.
        ' Otherwise, if you are using GemBox.Spreadsheet Free, comment out the 
        ' following line (Free version doesn't have SetLicense method). 
        SpreadsheetInfo.SetLicense("EVIG-1Y89-FYME-DPUJ")
        oExcel = New ExcelFile
        Sheets = New Collection
        Path = myPath
    End Sub


    Public Function GetAddSheet(ByVal myName As String) As clsExcelSheet
        If Not Sheets.Contains(myName.Trim.ToUpper) Then
            Dim mySheet As New clsExcelSheet(Me.oExcel, myName)
            Sheets.Add(mySheet, myName.Trim.ToUpper)
            Return mySheet
        Else
            Return Sheets(myName.Trim.ToUpper)
        End If
    End Function

    Public Sub Save(Optional ByVal Show As Boolean = True)
        If oExcel.Worksheets.Count > 0 Then
            oExcel.Save(Path, SaveOptions.XlsxDefault)
            If Show Then TryToDisplayGeneratedFile(Path)
        Else
            MsgBox("Error writing Excel file: contains no worksheets.")
        End If
    End Sub

    Sub TryToDisplayGeneratedFile(ByVal fileName As String)
        Try
            System.Diagnostics.Process.Start(fileName)
        Catch
            Console.WriteLine(fileName + " created in application folder.")
        End Try
    End Sub

End Class

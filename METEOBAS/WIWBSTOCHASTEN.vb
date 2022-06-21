Option Explicit On

Imports METEOBAS
Imports System

Module WIWBSTOCHASTEN

    'Copyright Siebe Bosch Hydroconsult, 2012
    'Lulofsstraat 55, unit 47 Den Haag, The Netherlands
    'this program extracts precipitation and/or evaporation data from the Meteobase database (PostgreSQL)

    'lokale variabelen
    Dim Setup As New METEOBAS.General.clsSetup
    Dim StochData As New METEOBAS.clsMBStochData(Setup)
    Sub Main()

        Dim RunOnServer As Boolean = True

        Dim myArg As String
        Console.WriteLine("This program reads precipitation and evaporation data from the Meteobase database")

        '----------------------------------------------------------------------------------------------------------------------------
        'system-dependent variables
        '----------------------------------------------------------------------------------------------------------------------------
        If RunOnServer Then
            StochData.DownloadURL = "http://85.214.197.176:8080/meteobase/downloads/"                                                 'laptop en server
            StochData.DownloadDIR = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-ApachePHP\apache\www\meteobase\downloads\"  'server
            StochData.FilesDir = "c:\Program Files (x86)\PostgreSQL\EnterpriseDB-ApachePHP\apache\www\meteobase\downloads\fixed\"  'server
        Else
            StochData.DownloadURL = "http://85.214.197.176:8080/meteobase/downloads/"                                          'laptop en server
            StochData.DownloadDIR = "c:\temp"                                                                                   'laptop
            StochData.FilesDir = "v:\PROJECTEN\H0069.Meteobase\01.Klimaatreeksen14\stochasten\"
        End If

        If My.Application.CommandLineArgs.Count = 0 Then
            Console.WriteLine("Export Huidig Vol? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.VOL_HUIDIG = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export Huidig OVF? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.OVF_HUIDIG = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2030 Vol? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.VOL_2030 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2030 OVF? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.OVF_2030 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2050 Vol? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.VOL_2050 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2050 OVF? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.OVF_2050 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2085 Vol? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.VOL_2085 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export 2085 OVF? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.OVF_2085 = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Export SHORT Durations? (TRUE/FALSE)")
            myArg = Console.ReadLine
            StochData.KORT = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Enter the session ID")
            myArg = Console.ReadLine
            StochData.SessionID = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Enter the order number")
            myArg = Console.ReadLine
            StochData.OrderNum = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Enter the name of the person who orders")
            myArg = Console.ReadLine
            StochData.Naam = myArg
            Setup.Log.CmdArgs.Add(myArg)
            Console.WriteLine("Enter their e-mailaddress")
            myArg = Console.ReadLine
            StochData.MailAdres = myArg
            Setup.Log.CmdArgs.Add(myArg)

        ElseIf My.Application.CommandLineArgs.Count <> 13 Then
            Console.WriteLine("Error: incorrect number of arguments presented")
            Setup.Log.AddError("Error: incorrect number of arguments presented to the executable.")
        Else
            StochData.VOL_HUIDIG = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(0))
            StochData.OVF_HUIDIG = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(1))
            StochData.VOL_2030 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(2))
            StochData.OVF_2030 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(3))
            StochData.VOL_2050 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(4))
            StochData.OVF_2050 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(5))
            StochData.VOL_2085 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(6))
            StochData.OVF_2085 = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(7))
            StochData.KORT = Setup.GeneralFunctions.GetBooleanFromString(My.Application.CommandLineArgs(8))
            StochData.SessionID = My.Application.CommandLineArgs(9)
            StochData.OrderNum = My.Application.CommandLineArgs(10)
            StochData.Naam = Setup.GeneralFunctions.RemoveBoundingQuotes(My.Application.CommandLineArgs(11))
            StochData.MailAdres = Setup.GeneralFunctions.RemoveBoundingQuotes(My.Application.CommandLineArgs(12))
        End If

        '----------------------------------------------------------------------------------------------------------------------------
        'query the database and start writing the data
        '----------------------------------------------------------------------------------------------------------------------------
        Try
            If StochData.Build() Then  'add the requested datafiles to a zipfile
                Call StochData.InitializeGoodMail("Stochastentabellen Meteobase")
                Call StochData.sendGoodEmail()
            Else
                Call StochData.InitializeBadMail("Stochastentabellen Meteobase")
                Call StochData.sendBadEmail()
            End If
        Catch ex As Exception
            Setup.Log.AddError(ex.Message)
            Call StochData.InitializeBadMail("Stochastentabellen Meteobase")
            Call StochData.sendBadEmail()
        End Try

    End Sub

End Module

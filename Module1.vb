Module Module1

    Sub Main()
        Console.WriteLine(vbCrLf)
        Console.ForegroundColor = ConsoleColor.DarkRed
        Console.WriteLine("OFFICE 365 CONNECT")
        Console.WriteLine(vbCrLf)
        Console.ForegroundColor = ConsoleColor.DarkGray
        Console.WriteLine("Welcome to 365Connect a tool to power up what you can do at your Office 365 Tenant")
        Console.WriteLine(vbCrLf)
        Console.WriteLine("Now we're going to load some modules needed in order to do our job.")
        Console.WriteLine(vbCrLf)
        Console.WriteLine("Checking for modules... (1 Min aprox)")
#Disable Warning
        Dim AzureAD = "Install-Module -Name AzureAD -force"""
        Dim MSOnline = "Install-Module -Name MSOnline -force"""
        Dim MSOnlineImporter = "Import-Module MSOnline"""

        Shell("powershell.exe -Command " + AzureAD)
        Threading.Thread.Sleep(20000)
        Console.WriteLine(vbCrLf)
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("1/3 Modules imported...")

        Shell("powershell.exe -Command " + MSOnline)
        Threading.Thread.Sleep(20000)
        Console.WriteLine(vbCrLf)
        Console.WriteLine("2/3 Modules imported...")

        Shell("powershell.exe -Command " + MSOnlineImporter)
        Threading.Thread.Sleep(20000)
        Console.WriteLine(vbCrLf)
        Console.WriteLine("2/3 Modules imported...")

        Console.WriteLine(vbCrLf)
        Console.ForegroundColor = ConsoleColor.DarkGray
        Console.WriteLine("Now we are going to set-up a few things")
        Console.WriteLine(vbCrLf)
        Console.WriteLine("Enter your Tenant UPN (For example: myuser@contoso.com) ")
        Dim Tenant = Console.ReadLine()
        Console.WriteLine(vbCrLf)
        Console.WriteLine("Enter your Tenant password")
        Dim Password = Console.ReadLine()

        Dim MSOnlineConnect = "Connect-MsolService"""
        Shell("powershell.exe -Command " + MSOnlineConnect)
        Console.WriteLine(vbCrLf)
        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Connected to Office 365 as: " + Tenant)
        Console.WriteLine(vbCrLf)

        'Let's build a Nice and powerfull Menu
        Console.Clear()
        Console.ForegroundColor = ConsoleColor.DarkRed
        Console.WriteLine("MENU")
        Console.ForegroundColor = ConsoleColor.DarkGray
        Console.WriteLine("______________________________________________________")
        Console.ForegroundColor = ConsoleColor.DarkCyan
        Console.WriteLine("1. List all Active users")
        Console.WriteLine("2. List all Deleted users")
        Console.WriteLine("3. List all Groups")
        Console.WriteLine("4. List all Shared mailboxes")
        Dim Options = Val(Console.ReadLine())

        Select Case Options
            Case 1
                Dim ListAllUsers = """"
                Shell("powershell.exe -Command " + ListAllUsers)
            Case Else
                Console.WriteLine("Option not available, choose one from above please.")
        End Select
    End Sub

End Module

Imports Lib3CXPhone

Module Module1

    Sub Main()
        Dim LaunchArg As String = Chr(34) & "example" & Chr(34) & " " & Chr(34) & "%CallerNumber%" & Chr(34)

        Lib3CXPhone.Lib3CXPhone.Set3CXLaunchProgram("C:\Scrips\action-client.exe", LaunchArg)

    End Sub



End Module

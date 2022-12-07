Imports System.IO

Public Class Lib3CXPhone
    Public Shared Sub Set3CXLaunchProgram(ByVal LaunchProgram As String, ByVal ProgramArguments As String)
        'Checks 3CX configuration file and applies WFM Action Client.
        'Program supports %CallerNumber% in arguments.
        Dim Config3CXFile As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\3CXPhone for Windows\3CXPhone.xml"
        If My.Computer.FileSystem.FileExists(Config3CXFile) = False Then
            Exit Sub
        End If

        If Not My.Computer.FileSystem.FileExists(LaunchProgram) Then
            Console.WriteLine("Missing launch program.")
            Exit Sub
        End If


        Dim FileData() As String = File.ReadAllLines(Config3CXFile)
        Dim FileChanged As Boolean = False
        Dim NewFileData As New List(Of String)
        For Each Line In FileData
            If Line.StartsWith("    <LaunchApplicationExecutablePath>") Then
                Dim ExePath As String = "    <LaunchApplicationExecutablePath>" & LaunchProgram & "</LaunchApplicationExecutablePath>"
                If Not Line = ExePath Then
                    Line = ExePath
                    FileChanged = True
                End If
            ElseIf Line.StartsWith("    <LaunchApplicationParameters>") Then
                Dim AppParams As String = "    <LaunchApplicationParameters>" & ProgramArguments & "</LaunchApplicationParameters>"
                If Not Line = AppParams Then
                    Line = AppParams
                    FileChanged = True
                End If
            ElseIf Line.StartsWith("    <LaunchExternalApplication>") Then
                Dim LaunchApp As String = "    <LaunchExternalApplication>True</LaunchExternalApplication>"
                If Not Line = LaunchApp Then
                    Line = LaunchApp
                    FileChanged = True
                End If
            ElseIf Line.StartsWith("    <LaunchApplicationNotifyWhen>") Then
                Dim LaunchType As String = "    <LaunchApplicationNotifyWhen>1</LaunchApplicationNotifyWhen>"
                If Not Line = LaunchType Then
                    Line = LaunchType
                    FileChanged = True
                End If
            End If
            NewFileData.Add(Line)
        Next
        If FileChanged = True Then

            If MsgBox("3CX Windows client requires a configuration update for integration." & Environment.NewLine & "Click Yes to continue (3CX Client will close) or no to abort.", MsgBoxStyle.YesNo, "3CX Integration update") = MsgBoxResult.No Then
                Exit Sub
            End If
            Dim KillProc As New Process()
            KillProc.StartInfo.FileName = "C:\windows\system32\taskkill.exe"
            KillProc.StartInfo.Arguments = "/f /im 3CXWin8Phone.exe"
            KillProc.Start()
            KillProc.WaitForExit()


            If My.Computer.FileSystem.FileExists(Config3CXFile & ".back") Then
                My.Computer.FileSystem.DeleteFile(Config3CXFile & ".back")
            End If
            My.Computer.FileSystem.MoveFile(Config3CXFile, Config3CXFile & ".back")
            Dim SR As New StreamWriter(Config3CXFile)
            For Each line In NewFileData
                SR.WriteLine(line)
            Next
            SR.Close()



        End If

    End Sub
End Class

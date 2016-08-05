Imports System.IO
Imports System.Text.RegularExpressions

Module Module1

    Private WithEvents Timer As System.Timers.Timer
    Private rand As Random
    Dim DestDir As String
    Dim SourceDir As String
    Dim bReversed As Boolean = False
    Dim bSorted As Boolean = False
    Dim orderRegex As String
    Dim bQuiet As Boolean = True

    Sub Main(ByVal Args() As String)
        Timer = New System.Timers.Timer
        AddHandler Timer.Elapsed, AddressOf Timer_Tick

        Dim ArgsOffset As Integer = 0
        Dim Interval As Integer
        Dim bNextArgRegex As Boolean = False
        If Args.Count < 3 Then
            ShowHelp()
            Return
        End If
        Try
            If Args(0).StartsWith("-") Then
                ArgsOffset += 1
                For i As Integer = 1 To Args(0).Count - 1
                        Select Case (Args(0).Chars(i))
                            Case "h"
                                ShowHelp()
                                Return
                            Case "c"
                                bSorted = True
                                bNextArgRegex = True
                            Case "v"
                                bQuiet = False
                            Case "r"
                                bReversed = True
                            Case Else
                                ShowHelp()
                                Return
                        End Select
                Next
            End If
            'Next
            If bNextArgRegex Then
                orderRegex = Args(ArgsOffset)
                bNextArgRegex = False
            End If
            Interval = Args(1 + ArgsOffset)
            DestDir = Args(2 + ArgsOffset)
            SourceDir = Args(3 + ArgsOffset)
        Catch e As Exception
            ShowHelp()
            Return
        End Try

        If Not Directory.Exists(DestDir) Then
            Console.Error.Write("ERROR: Destination directory " & DestDir & " cannot be found, quitting...")
            Return
        End If

        If Not Directory.Exists(SourceDir) Then
            Console.Error.Write("ERROR: Source directory " & DestDir & " cannot be found, quitting...")
            Return
        End If

        If Not DestDir.EndsWith("\") Then
            DestDir = DestDir & "\"
        End If

        If Not SourceDir.EndsWith("\") Then
            SourceDir = SourceDir & "\"
        End If

        Timer.Interval = 1000 * Interval

        Timer.Start()
        Console.WriteLine("Press any key to exit...")
        Console.ReadKey()
        Timer.Stop()
    End Sub

    Private Sub Timer_Tick(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs)
        Timer.Enabled = False
        Dim FilesToDrop As New SortedDictionary(Of String, String) 'Moved to timer function so if I drop more files into folder as it runs, it considers them.

        For Each f As String In Directory.GetFiles(SourceDir)
            Try
                Dim fi As FileInfo = New FileInfo(f)
                Dim thisKey As String
                If bSorted Then
                    thisKey = Regex.Match(fi.Name, orderRegex).Groups(1).Value
                Else
                    thisKey = fi.Name
                End If
                FilesToDrop.Add(thisKey, fi.FullName)
            Catch ex As Exception
                Console.WriteLine("Cant match group1 regex with " & orderRegex & " on " & f & vbNewLine & "Skipping...") ' & "Quitting...")
                'Return
            End Try
        Next

        Dim DropFile As KeyValuePair(Of String, String)
        If Not bReversed Then
            DropFile = FilesToDrop.First
        Else
            DropFile = FilesToDrop.Last
        End If

        Try
            File.Move(DropFile.Value, DestDir & "ESCN_-UKRECONData_AutoSSIS-" & New FileInfo(DropFile.Value).Name)
            FilesToDrop.Remove(DropFile.Key)
            If Not bQuiet Then Console.WriteLine("MOVED " & DropFile.Value & " to " & DestDir & DropFile.Key)
        Catch ex As Exception
            Console.WriteLine("Failed to copy " & DropFile.Value & " to " & DestDir & " continuing...")
        End Try

        If FilesToDrop.Any Then
            Timer.Enabled = True
        Else
            Console.WriteLine("No files to drop, quitting...")
        End If
    End Sub

    Private Sub ShowHelp()
        Console.WriteLine("TimedFileDrop [<-options>] <interval in seconds> <dest> <source>     ")
        Console.WriteLine("            drops 1 file from source to dest every interval          ")
        Console.WriteLine("            -h shows this help                                       ")
        Console.WriteLine("            -c <regex> orders the drops on group1's criteria (ascii) ")
        Console.WriteLine("            -r reverse the order                                     ")
        Console.WriteLine("            -v means run verbose                                     ")
    End Sub

End Module

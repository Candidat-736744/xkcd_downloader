Imports System.Net
Imports System.Windows.Forms
Imports Newtonsoft.Json

Module Module1

    Sub Main()
        Console.SetWindowSize(120, 30)
        Console.SetBufferSize(120, 30)
        Dim sbd As New FolderBrowserDialog
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write("You can write ")
        Console.ForegroundColor = ConsoleColor.Green
        Console.Write("DIALOG")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine(" to open a window dialog!")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.Write("Download path:")
        Console.ForegroundColor = ConsoleColor.White
        Dim valid As Boolean = False
        Dim outpath As String = Nothing
        Do Until valid
            outpath = Console.ReadLine()
            Select Case outpath
                Case Nothing
                    Environment.Exit(0)
                Case "DIALOG"
                    sbd.ShowDialog()
                    Console.SetCursorPosition(14, Console.CursorTop - 1)
                    Console.Write("      ")
                    Console.SetCursorPosition(14, Console.CursorTop)
                    Console.WriteLine(sbd.SelectedPath)
                    outpath = sbd.SelectedPath
            End Select
            If Not (outpath.EndsWith("/") Or outpath.EndsWith("\")) Then
                outpath = outpath & "/"
            End If
            outpath = outpath.Replace("\", "/")
            Try
                If Not IO.Directory.Exists(outpath) Then
                    valid = False
                Else
                    valid = True
                End If
            Catch
                valid = False
            End Try
            If Not valid Then
                Console.SetCursorPosition(0, Console.CursorTop - 1)
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("The folder """ & outpath & """ wasn't found!")
                Console.ForegroundColor = ConsoleColor.Gray
                Console.Write("Download path:")
            End If
        Loop
        Console.Write("Download webcomic id (or ""all""):")
        Console.ForegroundColor = ConsoleColor.White
        Dim id As String = Console.ReadLine()
        If id = "all" Then
            Dim max As Integer = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(New WebClient().DownloadString("https://xkcd.com/info.0.json")).Item("num")
            For i = 1 To max
                downloader(i, max, outpath)
            Next
        Else
            downloader(id, id, outpath)
        End If
    End Sub

    Sub downloader(id As Integer, max As Integer, output As String)
        Try
            Dim scan As String = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(New WebClient().DownloadString("http://xkcd.com/" & id & "/info.0.json")).Item("img")
            My.Computer.Network.DownloadFile(scan, output & scan.Split("/")(4))
            Console.WriteLine("Downloaded file [" & id & "/" & max & "]: " & scan.Split("/")(4).Split(".")(0))
        Catch
            My.Computer.FileSystem.WriteAllText(output & "_failed.txt", id & vbNewLine, True)
            Console.WriteLine("Failed to download file: " & id)
            Console.Title = id & "/" & max
        End Try
    End Sub

End Module

Imports System.IO
Imports System.Reflection
Imports InternetArchive.InternetArchiveCli.Commands
Imports InternetArchive.InternetArchiveCli.Core
Imports InternetArchive.InternetArchiveCli.Services

Module Program
    Private ReadOnly Commands As New List(Of CommandDefinition) From {
        New CommandDefinition("account", New String() {"ac"}),
        New CommandDefinition("configure", New String() {"co"}),
        New CommandDefinition("copy", New String() {"cp"}),
        New CommandDefinition("delete", New String() {"rm"}),
        New CommandDefinition("download", New String() {"do"}),
        New CommandDefinition("flag", New String() {"fl"}),
        New CommandDefinition("list", New String() {"ls"}),
        New CommandDefinition("metadata", New String() {"md"}),
        New CommandDefinition("move", New String() {"mv"}),
        New CommandDefinition("reviews", New String() {"re"}),
        New CommandDefinition("search", New String() {"se"}),
        New CommandDefinition("simplelists", New String() {"sl"}),
        New CommandDefinition("tasks", New String() {"ta"}),
        New CommandDefinition("upload", New String() {"up"})
    }

    Private ReadOnly CommandLookup As Dictionary(Of String, CommandDefinition) =
        BuildCommandLookup()

    Public Function Main(args As String()) As Integer
        Try
            If args.Length = 0 Then
                PrintHelp(Console.Error)
                Return 1
            End If

            Dim parsed As ParsedArguments = ParseArguments(New List(Of String)(args))
            If parsed.ShowVersion Then
                Console.WriteLine(GetCliVersion())
                Return 0
            End If

            Dim overrideConfig = BuildConfig(parsed)
            Dim mergedConfig = ConfigService.LoadConfig(overrideConfig, parsed.ConfigFile)
            Dim resolvedConfigFile = ConfigService.ResolveConfigFile(parsed.ConfigFile)
            Dim session As New ArchiveSession(mergedConfig, resolvedConfigFile)
            Return ExecuteCommand(parsed, session)
        Catch argEx As ArgumentException
            If argEx.Message.StartsWith("unrecognized arguments:", StringComparison.Ordinal) Then
                PrintGlobalUsageError(argEx.Message)
                Return 2
            End If
            Console.Error.WriteLine("error: " & argEx.Message)
            Return 2
        Catch notImpl As NotImplementedException
            Console.Error.WriteLine("error: " & notImpl.Message)
            Return 1
        Catch ex As Exception
            Console.Error.WriteLine("error: " & ex.Message)
            Return 1
        End Try
    End Function

    Private Function BuildCommandLookup() As Dictionary(Of String, CommandDefinition)
        Dim map As New Dictionary(Of String, CommandDefinition)(
            StringComparer.OrdinalIgnoreCase
        )
        For Each cmd In Commands
            map(cmd.Name) = cmd
            For Each aliasName In cmd.Aliases
                map(aliasName) = cmd
            Next
        Next
        Return map
    End Function

    Private Function BuildConfig(parsed As ParsedArguments) As Dictionary(Of String, Dictionary(Of String, String))
        Dim config As New Dictionary(Of String, Dictionary(Of String, String))(
            StringComparer.OrdinalIgnoreCase
        )

        If parsed.EnableLog Then
            config("logging") = New Dictionary(Of String, String) From {{"level", "INFO"}}
        ElseIf parsed.EnableDebug Then
            config("logging") = New Dictionary(Of String, String) From {{"level", "DEBUG"}}
        End If

        If parsed.Insecure Then
            config("general") = New Dictionary(Of String, String) From {{"secure", "False"}}
        End If

        If Not String.IsNullOrWhiteSpace(parsed.Host) Then
            If Not config.ContainsKey("general") Then
                config("general") = New Dictionary(Of String, String)(
                    StringComparer.OrdinalIgnoreCase
                )
            End If
            config("general")("host") = parsed.Host
        End If

        If Not String.IsNullOrWhiteSpace(parsed.UserAgentSuffix) Then
            If Not config.ContainsKey("general") Then
                config("general") = New Dictionary(Of String, String)(
                    StringComparer.OrdinalIgnoreCase
                )
            End If
            config("general")("user_agent_suffix") = parsed.UserAgentSuffix
        End If

        Return config
    End Function

    Private Function ExecuteCommand(parsed As ParsedArguments, session As ArchiveSession) As Integer
        If String.IsNullOrWhiteSpace(parsed.CommandName) Then
            Throw New ArgumentException("No command provided")
        End If

        Dim commandDef As CommandDefinition = Nothing
        If Not CommandLookup.TryGetValue(parsed.CommandName, commandDef) Then
            Throw New ArgumentException(String.Format("Unknown command: {0}", parsed.CommandName))
        End If

        If String.Equals(commandDef.Name, "configure", StringComparison.OrdinalIgnoreCase) Then
            Return ConfigureCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "search", StringComparison.OrdinalIgnoreCase) Then
            Return SearchCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "list", StringComparison.OrdinalIgnoreCase) Then
            Return ListCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "metadata", StringComparison.OrdinalIgnoreCase) Then
            Return MetadataCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "flag", StringComparison.OrdinalIgnoreCase) Then
            Return FlagCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "tasks", StringComparison.OrdinalIgnoreCase) Then
            Return TasksCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "reviews", StringComparison.OrdinalIgnoreCase) Then
            Return ReviewsCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "delete", StringComparison.OrdinalIgnoreCase) Then
            Return DeleteCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "copy", StringComparison.OrdinalIgnoreCase) Then
            Return CopyCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "move", StringComparison.OrdinalIgnoreCase) Then
            Return MoveCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "simplelists", StringComparison.OrdinalIgnoreCase) Then
            Return SimplelistsCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "account", StringComparison.OrdinalIgnoreCase) Then
            Return AccountCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "download", StringComparison.OrdinalIgnoreCase) Then
            Return DownloadCommand.Execute(session, parsed.CommandArgs)
        End If
        If String.Equals(commandDef.Name, "upload", StringComparison.OrdinalIgnoreCase) Then
            Return UploadCommand.Execute(session, parsed.CommandArgs)
        End If

        ' Placeholder dispatch to preserve the command surface while command logic
        ' is ported module-by-module from Python.
        Throw New NotImplementedException(
            String.Format("Command '{0}' is not yet ported.", commandDef.Name)
        )
    End Function

    Private Function ParseArguments(rawArgs As IList(Of String)) As ParsedArguments
        Dim parsed As New ParsedArguments With {
            .CommandArgs = New List(Of String)()
        }

        Dim i As Integer = 0
        While i < rawArgs.Count
            Dim current As String = rawArgs(i)
            Select Case current
                Case "-v", "--version"
                    parsed.ShowVersion = True
                    i += 1
                Case "-c", "--config-file"
                    If i + 1 >= rawArgs.Count Then
                        Throw New ArgumentException("Missing value for --config-file")
                    End If
                    parsed.ConfigFile = ValidateConfigPath(rawArgs(i + 1), rawArgs)
                    i += 2
                Case "-l", "--log"
                    parsed.EnableLog = True
                    i += 1
                Case "-d", "--debug"
                    parsed.EnableDebug = True
                    i += 1
                Case "-i", "--insecure"
                    parsed.Insecure = True
                    i += 1
                Case "-H", "--host"
                    If i + 1 >= rawArgs.Count Then
                        Throw New ArgumentException("Missing value for --host")
                    End If
                    parsed.Host = rawArgs(i + 1)
                    i += 2
                Case "--user-agent-suffix"
                    If i + 1 >= rawArgs.Count Then
                        Throw New ArgumentException("Missing value for --user-agent-suffix")
                    End If
                    parsed.UserAgentSuffix = rawArgs(i + 1)
                    i += 2
                Case Else
                    parsed.CommandName = current
                    For j As Integer = i + 1 To rawArgs.Count - 1
                        parsed.CommandArgs.Add(rawArgs(j))
                    Next
                    Exit While
            End Select
        End While

        Return parsed
    End Function

    Private Function GetCliVersion() As String
        Try
            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim fileVersion = asm.GetCustomAttribute(Of AssemblyFileVersionAttribute)()
            If fileVersion IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(fileVersion.Version) Then
                Return fileVersion.Version
            End If

            Dim version = asm.GetName().Version
            If version IsNot Nothing Then
                Return version.ToString()
            End If
        Catch ex As Exception
            Console.Error.WriteLine("warning: unable to determine CLI version from assembly metadata - " & ex.Message)
        End Try

        Return "0.0.0.0"
    End Function

    Private Sub PrintGlobalUsageError(message As String)
        Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " [-h] [-v] [-c FILE] [-l] [-d] [-i] [-H HOST]")
        Console.Error.WriteLine("             [--user-agent-suffix STRING]")
        Console.Error.WriteLine("             {command} ...")
        Console.Error.WriteLine(CliApp.ExecutableName() & ": error: " & message)
    End Sub

    Private Sub PrintHelp(errorOutput As TextWriter)
        errorOutput.WriteLine("A command line interface to Archive.org.")
        errorOutput.WriteLine()
        errorOutput.WriteLine("Documentation for 'ia' is available at:")
        errorOutput.WriteLine()
        errorOutput.WriteLine(vbTab & "https://archive.org/developers/internetarchive/cli.html")
        errorOutput.WriteLine()
        errorOutput.WriteLine("See 'ia {command} --help' for help on a specific command.")
        errorOutput.WriteLine()
        errorOutput.WriteLine("Global options:")
        errorOutput.WriteLine("  -v, --version")
        errorOutput.WriteLine("  -c, --config-file FILE")
        errorOutput.WriteLine("  -l, --log")
        errorOutput.WriteLine("  -d, --debug")
        errorOutput.WriteLine("  -i, --insecure")
        errorOutput.WriteLine("  -H, --host HOST")
        errorOutput.WriteLine("      --user-agent-suffix STRING")
        errorOutput.WriteLine()
        errorOutput.WriteLine("commands:")
        For Each cmd In Commands
            errorOutput.WriteLine(String.Format("  {0} ({1})", cmd.Name, String.Join(", ", cmd.Aliases)))
        Next
    End Sub

    Private Function ValidateConfigPath(configPath As String, rawArgs As IList(Of String)) As String
        Dim resolvedPath As String = Path.GetFullPath(configPath)
        Dim configureSeen As Boolean = False
        For Each value In rawArgs
            If String.Equals(value, "configure", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(value, "co", StringComparison.OrdinalIgnoreCase) Then
                configureSeen = True
                Exit For
            End If
        Next

        If Not configureSeen AndAlso Not File.Exists(resolvedPath) Then
            Throw New ArgumentException(String.Format("Config file does not exist: {0}", resolvedPath))
        End If

        Return resolvedPath
    End Function

    Private NotInheritable Class CommandDefinition
        Public Sub New(name As String, aliases As IEnumerable(Of String))
            Me.Name = name
            Me.Aliases = New List(Of String)(aliases)
        End Sub

        Public ReadOnly Property Aliases As List(Of String)
        Public ReadOnly Property Name As String
    End Class

    Private NotInheritable Class ParsedArguments
        Public Property CommandArgs As List(Of String)
        Public Property CommandName As String
        Public Property ConfigFile As String
        Public Property EnableDebug As Boolean
        Public Property EnableLog As Boolean
        Public Property Host As String
        Public Property Insecure As Boolean
        Public Property ShowVersion As Boolean
        Public Property UserAgentSuffix As String
    End Class
End Module

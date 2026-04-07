Imports System.Globalization
Imports System.Text.RegularExpressions
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class DeleteCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As DeleteArgs
            Try
                parsed = ParseArguments(args)
            Catch ex As ArgumentException
                PrintUsageAndError(ex.Message)
                Return 2
            End Try
            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If
            Dim verbose As Boolean = Not parsed.Quiet
            Dim metadata = session.GetItemMetadata(parsed.Identifier, Nothing)
            If metadata.Count = 0 Then
                Console.Error.WriteLine(parsed.Identifier & ": skipping, item doesn't exist.")
                Return 0
            End If

            Dim requestHeaders = ApiShared.ConvertHeaderValues(parsed.Headers)
            If Not requestHeaders.ContainsKey("x-archive-keep-old-version") AndAlso Not parsed.NoBackup Then
                requestHeaders("x-archive-keep-old-version") = "1"
            End If

            If verbose Then
                Console.Error.WriteLine("Deleting files from " & parsed.Identifier)
            End If

            Dim files = GetFilesToDelete(parsed, metadata)
            If files.Count = 0 Then
                Console.Error.WriteLine(" warning: no files found, nothing deleted.")
                Return 1
            End If

            Dim errors As Boolean = False
            For Each fileName In files
                If parsed.DryRun Then
                    If parsed.Cascade Then
                        Console.Error.WriteLine(
                            String.Format(
                                " will delete: {0}/{1} and all derivatives",
                                parsed.Identifier,
                                fileName
                            )
                        )
                    Else
                        Console.Error.WriteLine(
                            String.Format(" will delete: {0}/{1}", parsed.Identifier, fileName)
                        )
                    End If
                    Continue For
                End If

                If verbose Then
                    Dim msg As String = " deleting: " & fileName
                    If parsed.Cascade Then
                        msg &= " and all derivative files."
                    End If
                    Console.Error.WriteLine(msg)
                End If

                If Not requestHeaders.ContainsKey("x-archive-cascade-delete") Then
                    requestHeaders("x-archive-cascade-delete") = If(parsed.Cascade, "1", "0")
                End If

                Dim result = session.DeleteS3File(parsed.Identifier, fileName, requestHeaders, parsed.Retries)
                If result.StatusCode = 503 Then
                    Console.Error.WriteLine(" error: max retries exceeded for " & fileName)
                    errors = True
                    Continue For
                End If

                If result.StatusCode <> 204 Then
                    Dim msg As String = ApiShared.GetS3XmlText(result.Text)
                    Console.Error.WriteLine(
                        String.Format(" error: {0} ({1})", msg, result.StatusCode)
                    )
                    errors = True
                    Continue For
                End If
            Next

            Return If(errors, 1, 0)
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function GetFilesToDelete(parsed As DeleteArgs, metadata As Dictionary(Of String, Object)) As List(Of String)
            Dim allFiles As List(Of ArchiveFileEntry) = ApiShared.GetArchiveFileEntries(metadata)

            If parsed.DeleteAll OrElse parsed.Files.Count = 0 Then
                parsed.Cascade = True
            End If

            If Not String.IsNullOrWhiteSpace(parsed.GlobPattern) Then
                Dim patterns = parsed.GlobPattern.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim matches As New List(Of String)()
                For Each f In allFiles
                    Dim name As String = f.Name
                    For Each pattern In patterns
                        If GlobMatch(name, pattern) Then
                            matches.Add(name)
                            Exit For
                        End If
                    Next
                Next
                Return matches
            End If

            If parsed.Formats.Count > 0 Then
                Dim matches As New List(Of String)()
                For Each f In allFiles
                    Dim fmt As String = f.Format
                    For Each requested In parsed.Formats
                        If String.Equals(fmt, requested, StringComparison.OrdinalIgnoreCase) Then
                            matches.Add(f.Name)
                            Exit For
                        End If
                    Next
                Next
                Return matches
            End If

            Dim names As List(Of String)
            If parsed.Files.Count = 1 AndAlso parsed.Files(0) = "-" Then
                names = New List(Of String)()
                While True
                    Dim line As String = Console.In.ReadLine()
                    If line Is Nothing Then
                        Exit While
                    End If
                    Dim trimmed As String = line.Trim()
                    If trimmed.Length > 0 Then
                        names.Add(trimmed)
                    End If
                End While
            Else
                names = New List(Of String)(parsed.Files)
                For i As Integer = 0 To names.Count - 1
                    names(i) = names(i).Trim()
                Next
            End If

            If names.Count = 0 Then
                Dim allNames As New List(Of String)()
                For Each f In allFiles
                    allNames.Add(f.Name)
                Next
                Return allNames
            End If

            Dim available As New HashSet(Of String)(StringComparer.Ordinal)
            For Each f In allFiles
                available.Add(f.Name)
            Next
            Dim selected As New List(Of String)()
            For Each n In names
                If available.Contains(n) Then
                    selected.Add(n)
                End If
            Next
            Return selected
        End Function

        Private Shared Function GlobMatch(value As String, pattern As String) As Boolean
            Dim escaped = Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".")
            Return Regex.IsMatch(value, "^" & escaped & "$", RegexOptions.IgnoreCase)
        End Function

        Private Shared Sub MergeQueryString(destination As Dictionary(Of String, Object), raw As String, optionName As String)
            Dim normalized As String = raw
            If normalized.IndexOf("="c) < 0 AndAlso normalized.IndexOf(":"c) >= 0 Then
                Dim firstColon As Integer = normalized.IndexOf(":"c)
                normalized = normalized.Substring(0, firstColon) & "=" & normalized.Substring(firstColon + 1)
            End If
            Dim eqIndex As Integer = normalized.IndexOf("="c)
            If eqIndex <= 0 OrElse eqIndex = normalized.Length - 1 Then
                Throw New ArgumentException(
                    String.Format("{0} must be formatted as 'key=value' or 'key:value'", optionName)
                )
            End If
            Dim key As String = Uri.UnescapeDataString(normalized.Substring(0, eqIndex))
            Dim value As String = Uri.UnescapeDataString(normalized.Substring(eqIndex + 1))
            If destination.ContainsKey(key) Then
                Dim existing = destination(key)
                Dim list = If(TryCast(existing, List(Of String)), New List(Of String) From {Convert.ToString(existing, CultureInfo.InvariantCulture)})
                list.Add(value)
                destination(key) = list
            Else
                destination(key) = value
            End If
        End Sub

        Private Shared Function ParseArguments(args As IList(Of String)) As DeleteArgs
            Dim parsed As New DeleteArgs With {
                .Files = New List(Of String)(),
                .Headers = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Formats = New List(Of String)()
            }
            Dim i As Integer = 0
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-q", "--quiet"
                        parsed.Quiet = True
                    Case "-c", "--cascade"
                        parsed.Cascade = True
                    Case "-H", "--header"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.Headers, args(i), current)
                    Case "-a", "--all"
                        parsed.DeleteAll = True
                    Case "-d", "--dry-run"
                        parsed.DryRun = True
                    Case "-g", "--glob"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.GlobPattern = args(i)
                    Case "-f", "--format"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Formats.Add(args(i))
                    Case "-R", "--retries"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Retries = Integer.Parse(args(i), CultureInfo.InvariantCulture)
                    Case "--no-backup"
                        parsed.NoBackup = True
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            Throw New ArgumentException(String.Format("Unknown option: {0}", current))
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                            parsed.Identifier = current
                        Else
                            parsed.Files.Add(current)
                        End If
                End Select
                i += 1
            End While

            If String.IsNullOrWhiteSpace(parsed.Identifier) AndAlso Not parsed.ShowHelp Then
                Throw New ArgumentException("the following arguments are required: identifier")
            End If
            Return parsed
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " delete [-h] [-q] [-c] [-H KEY:VALUE] [-a] [-d] [-g PATTERN]")
            Console.Error.WriteLine("                    [-f FORMAT] [-R RETRIES] [--no-backup]")
            Console.Error.WriteLine("                    identifier [file ...]")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier for the item from which files are to be deleted.")
            Console.Error.WriteLine("  file                  Specific file(s) to delete.")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -q, --quiet           Suppress verbose output.")
            Console.Error.WriteLine("  -c, --cascade         Delete all associated files including derivatives and the original.")
            Console.Error.WriteLine("  -H KEY:VALUE, --header KEY:VALUE")
            Console.Error.WriteLine("                        S3 HTTP headers to send with your request.")
            Console.Error.WriteLine("  -a, --all             Delete all files in the given item.")
            Console.Error.WriteLine("  -d, --dry-run         Output files to be deleted but do not delete.")
            Console.Error.WriteLine("  -g PATTERN, --glob PATTERN")
            Console.Error.WriteLine("                        Only delete files matching the given glob pattern.")
            Console.Error.WriteLine("  -f FORMAT, --format FORMAT")
            Console.Error.WriteLine("                        Only delete files matching the specified format.")
            Console.Error.WriteLine("  -R RETRIES, --retries RETRIES")
            Console.Error.WriteLine("                        Number of retries on S3 503 SlowDown error.")
            Console.Error.WriteLine("  --no-backup           Turn off archive.org backups.")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " delete [-h] [-q] [-c] [-H KEY:VALUE] [-a] [-d] [-g PATTERN]")
            Console.Error.WriteLine("                    [-f FORMAT] [-R RETRIES] [--no-backup]")
            Console.Error.WriteLine("                    identifier [file ...]")
            Console.Error.WriteLine(CliApp.ExecutableName() & " delete: error: " & message)
        End Sub

        Private NotInheritable Class DeleteArgs
            Public Property Cascade As Boolean
            Public Property DeleteAll As Boolean
            Public Property DryRun As Boolean
            Public Property Files As List(Of String)
            Public Property Formats As List(Of String)
            Public Property GlobPattern As String
            Public Property Headers As Dictionary(Of String, Object)
            Public Property Identifier As String
            Public Property NoBackup As Boolean
            Public Property Quiet As Boolean
            Public Property Retries As Integer = 2
            Public Property ShowHelp As Boolean
        End Class
    End Class
End Namespace

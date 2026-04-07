Imports System.Globalization
Imports System.IO
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class DownloadCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As DownloadArgs
            Try
                parsed = ParseArguments(args)
                If parsed.ShowHelp Then
                    PrintHelp()
                    Return 0
                End If
                ValidateArgs(parsed)
            Catch ex As ArgumentException
                PrintUsageAndError(ex.Message)
                Return 2
            End Try

            Dim ids As List(Of String) = ResolveIdentifiers(session, parsed)
            If ids.Count = 0 Then
                If Not String.IsNullOrWhiteSpace(parsed.SearchQuery) Then
                    Console.Error.WriteLine(
                        String.Format(
                            "error: the query '{0}' returned no results",
                            parsed.SearchQuery
                        )
                    )
                End If
                Return 1
            End If

            Dim hasErrors As Boolean = False
            Dim checksumArchiveEntries As HashSet(Of String) = Nothing
            If parsed.ChecksumArchive Then
                checksumArchiveEntries = LoadChecksumArchiveEntries("_checksum_archive.txt")
            End If
            Using sharedClient As New HttpClient()
                If parsed.TimeoutSeconds.HasValue Then
                    sharedClient.Timeout = TimeSpan.FromSeconds(parsed.TimeoutSeconds.Value)
                End If

                For i As Integer = 0 To ids.Count - 1
                    Dim identifier As String = ids(i)
                    Dim itemMetadata As Dictionary(Of String, Object)
                    Try
                        itemMetadata = session.GetItemMetadata(identifier, Nothing)
                    Catch ex As Exception
                        Console.Error.WriteLine(identifier & ": failed to retrieve item metadata - errors")
                        If ex.Message.IndexOf("You are attempting to make an HTTPS", StringComparison.Ordinal) >= 0 Then
                            Console.Error.WriteLine()
                            Console.Error.WriteLine(ex.Message)
                            Return 1
                        End If
                        Continue For
                    End Try

                    If IsDarkItem(itemMetadata) Then
                        If Not parsed.Quiet Then
                            Console.Error.WriteLine(" skipping " & identifier & ", item is dark")
                        End If
                        Continue For
                    End If

                    If itemMetadata.Count = 0 Then
                        If Not parsed.Quiet Then
                            Console.Error.WriteLine(" skipping " & identifier & ", item does not exist.")
                        End If
                        Continue For
                    End If

                    Dim fileCandidates = GetCandidateFiles(identifier, itemMetadata, parsed)
                    If fileCandidates.Count = 0 Then
                        If Not parsed.Quiet Then
                            Console.Error.WriteLine(" skipping " & identifier & ", no matching files found.")
                        End If
                        Continue For
                    End If

                    If Not parsed.Quiet AndAlso Not parsed.DryRun AndAlso Not parsed.StdoutOutput Then
                        If ids.Count > 1 Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "{0} ({1}/{2}):",
                                    identifier,
                                    i + 1,
                                    ids.Count
                                )
                            )
                        Else
                            Console.Error.WriteLine(identifier & ":")
                        End If
                    End If

                    For Each fileEntry In fileCandidates
                        Dim fileName As String = fileEntry.Name
                        Dim fileUrl As String = BuildDownloadUrl(session, identifier, fileName, parsed.RequestParameters)

                        If parsed.DryRun Then
                            Console.WriteLine(fileUrl)
                            Continue For
                        End If

                        Dim relativePath As String
                        If parsed.NoDirectories Then
                            relativePath = fileName
                        Else
                            relativePath = identifier & "/" & fileName
                        End If
                        Dim destinationPath As String = ResolveDestinationPath(relativePath, parsed.DestDir)
                        If Not EnsurePathWithinBase(destinationPath, parsed.DestDir) Then
                            Console.Error.WriteLine(" error: Unsafe file path resolved outside destination directory: " & destinationPath)
                            hasErrors = True
                            Continue For
                        End If

                        If ShouldSkipExisting(destinationPath, fileEntry, parsed, checksumArchiveEntries) Then
                            Continue For
                        End If

                        Dim success As Boolean = DownloadSingleFile(
                            sharedClient,
                            fileUrl,
                            fileName,
                            destinationPath,
                            parsed.StdoutOutput,
                            Not parsed.Quiet AndAlso Not parsed.StdoutOutput,
                            parsed.Retries
                        )
                        If Not success Then
                            hasErrors = True
                            Continue For
                        End If

                        If Not parsed.NoChangeTimestamp Then
                            ApplyMtime(destinationPath, fileEntry)
                        End If
                    Next
                Next
            End Using

            Return If(hasErrors, 1, 0)
        End Function

        Private Shared Sub ApplyMtime(destinationPath As String, fileEntry As ArchiveFileEntry)
            If Not File.Exists(destinationPath) Then
                Return
            End If
            Dim mtime As Long
            If fileEntry.TryGetLong("mtime", mtime) Then
                Try
                    Dim dt = DateTimeOffset.FromUnixTimeSeconds(mtime).UtcDateTime
                    File.SetLastWriteTimeUtc(destinationPath, dt)
                Catch
                End Try
            End If
        End Sub

        Private Shared Function BuildDownloadUrl(session As ArchiveSession, identifier As String, fileName As String, requestParams As Dictionary(Of String, Object)) As String
            Dim encodedFile = EscapePathKeepSlash(fileName)
            Dim baseUrl As String = String.Format(
                "{0}//{1}/download/{2}/{3}",
                session.Protocol,
                session.Host,
                Uri.EscapeDataString(identifier),
                encodedFile
            )
            If requestParams Is Nothing OrElse requestParams.Count = 0 Then
                Return baseUrl
            End If
            Dim parts As New List(Of String)()
            For Each kvp In requestParams
                parts.Add(Uri.EscapeDataString(kvp.Key) & "=" & Uri.EscapeDataString(Convert.ToString(kvp.Value)))
            Next
            Return baseUrl & "?" & String.Join("&", parts)
        End Function

        Private Shared Function ComputeMd5(path As String) As String
            Using hasher As MD5 = MD5.Create()
                Using fs = File.OpenRead(path)
                    Dim hash = hasher.ComputeHash(fs)
                    Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
                End Using
            End Using
        End Function

        Private Shared Function DownloadSingleFile(client As HttpClient, url As String, displayName As String, destinationPath As String, stdoutOutput As Boolean,
                                                   showProgress As Boolean, retries As Integer) As Boolean

            Dim attempts As Integer = Math.Max(1, retries)
            For i As Integer = 1 To attempts
                Try
                    Using response As HttpResponseMessage = client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult()
                        If Not response.IsSuccessStatusCode Then
                            If i = attempts Then
                                Console.Error.WriteLine(
                                    String.Format(" error downloading file {0}, status {1}", destinationPath, CInt(response.StatusCode))
                                )
                                Return False
                            End If
                            Continue For
                        End If

                        Dim totalBytes As Long = 0
                        If response.Content.Headers.ContentLength.HasValue Then
                            totalBytes = response.Content.Headers.ContentLength.Value
                        End If
                        Dim progress = BuildDownloadProgressReporter(displayName, showProgress)

                        Using source As Stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
                            Dim target As Stream = Nothing
                            Dim disposeTarget As Boolean = False
                            Try
                                If stdoutOutput Then
                                    target = Console.OpenStandardOutput()
                                Else
                                    Dim parent As String = Path.GetDirectoryName(destinationPath)
                                    If Not String.IsNullOrWhiteSpace(parent) Then
                                        Directory.CreateDirectory(parent)
                                    End If
                                    target = New FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                    disposeTarget = True
                                End If

                                Dim buffer(1048575) As Byte
                                Dim bytesRead As Integer = 0
                                Dim bytesReceived As Long = 0
                                Do
                                    bytesRead = source.Read(buffer, 0, buffer.Length)
                                    If bytesRead <= 0 Then
                                        Exit Do
                                    End If
                                    target.Write(buffer, 0, bytesRead)
                                    bytesReceived += bytesRead
                                    progress?.Invoke(bytesReceived, totalBytes)
                                Loop
                                progress?.Invoke(bytesReceived, If(totalBytes > 0, totalBytes, bytesReceived))
                            Finally
                                If disposeTarget AndAlso target IsNot Nothing Then
                                    target.Dispose()
                                End If
                            End Try
                        End Using
                        Return True
                    End Using
                Catch ex As Exception
                    If i = attempts Then
                        Console.Error.WriteLine(" error downloading file " & destinationPath & ", exception raised: " & ex.Message)
                        Return False
                    End If
                End Try
            Next
            Return False
        End Function

        Private Shared Function BuildDownloadProgressReporter(remoteName As String, enabled As Boolean) As Action(Of Long, Long)
            If Not enabled Then
                Return Nothing
            End If

            Dim sw As Stopwatch = Stopwatch.StartNew()
            Dim useUnicodeBars As Boolean = UseUnicodeProgressBar()
            Dim barFillChar As Char = If(useUnicodeBars, "█"c, "#"c)
            Dim barEmptyChar As Char = If(useUnicodeBars, "░"c, "-"c)
            Dim minBarWidth As Integer = 10
            Dim maxBarWidth As Integer = 40
            Dim lastRenderedMs As Long = -1
            Dim lastLineLength As Integer = 0

            Return Sub(bytesReceived As Long, callbackTotalBytes As Long)
                       Dim total As Long = callbackTotalBytes
                       If total < 0 Then
                           total = 0
                       End If
                       If bytesReceived < 0 Then
                           bytesReceived = 0
                       End If
                       If total > 0 AndAlso bytesReceived > total Then
                           bytesReceived = total
                       End If

                       Dim isFinal As Boolean = (total > 0 AndAlso bytesReceived >= total)
                       Dim elapsedMs As Long = sw.ElapsedMilliseconds
                       If Not isFinal AndAlso lastRenderedMs >= 0 AndAlso (elapsedMs - lastRenderedMs) < 100 Then
                           Return
                       End If
                       lastRenderedMs = elapsedMs

                       Dim elapsedSeconds As Double = Math.Max(sw.Elapsed.TotalSeconds, 0.000001R)
                       Dim speedBytesPerSecond As Double = CDbl(bytesReceived) / elapsedSeconds
                       Dim rateText As String = FormatRate(speedBytesPerSecond)
                       Dim elapsedText As String = FormatElapsed(sw.Elapsed)
                       Dim countText As String = String.Format(
                           CultureInfo.InvariantCulture,
                           "{0}/{1}",
                           FormatBinaryAmount(bytesReceived),
                           If(total > 0, FormatBinaryAmount(total), "?")
                       )

                       Dim etaText As String = "?"
                       If total > 0 Then
                           Dim remaining As Long = Math.Max(0L, total - bytesReceived)
                           If remaining = 0 Then
                               etaText = "00:00"
                           ElseIf speedBytesPerSecond > 0.000001R Then
                               Dim etaSeconds As Integer = CInt(Math.Ceiling(CDbl(remaining) / speedBytesPerSecond))
                               etaText = FormatEta(etaSeconds)
                           End If
                       End If

                       Dim percentWhole As Integer = 0
                       If total > 0 Then
                           percentWhole = CInt(Math.Round((CDbl(bytesReceived) / CDbl(total)) * 100.0R, MidpointRounding.AwayFromZero))
                           If percentWhole < 0 Then percentWhole = 0
                           If percentWhole > 100 Then percentWhole = 100
                       End If

                       Dim width As Integer = 0
                       Try
                           width = Console.WindowWidth
                       Catch
                           width = 0
                       End Try

                       Dim suffixBeforeBar As String = String.Format(
                           CultureInfo.InvariantCulture,
                           " {0,3}%|",
                           percentWhole
                       )
                       Dim suffixAfterBar As String = String.Format(
                           CultureInfo.InvariantCulture,
                           "| {0} [{1}<{2}, {3}]",
                           countText,
                           elapsedText,
                           etaText,
                           rateText
                       )

                       Dim displayName As String = remoteName
                       If width > 0 Then
                           Dim reservedNoName As Integer =
                               " downloading ".Length + 2 + suffixBeforeBar.Length + suffixAfterBar.Length + minBarWidth
                           Dim maxNameLen As Integer = width - reservedNoName
                           If maxNameLen < 8 Then
                               maxNameLen = 8
                           End If
                           If displayName.Length > maxNameLen Then
                               displayName = displayName.Substring(0, maxNameLen - 1) & "~"
                           End If
                       ElseIf displayName.Length > 48 Then
                           displayName = displayName.Substring(0, 47) & "~"
                       End If

                       Dim barWidth As Integer = 24
                       If width > 0 Then
                           Dim reserved As Integer =
                               " downloading ".Length + displayName.Length + 2 + suffixBeforeBar.Length + suffixAfterBar.Length
                           Dim calculatedBar As Integer = width - reserved - 1
                           barWidth = Math.Max(minBarWidth, Math.Min(maxBarWidth, calculatedBar))
                       End If

                       Dim bar As String
                       If total > 0 Then
                           Dim percent As Double = (CDbl(bytesReceived) / CDbl(total)) * 100.0R
                           Dim filled As Integer = CInt(Math.Floor((percent / 100.0R) * barWidth))
                           If filled < 0 Then
                               filled = 0
                           ElseIf filled > barWidth Then
                               filled = barWidth
                           End If
                           bar = New String(barFillChar, filled) & New String(barEmptyChar, barWidth - filled)
                       Else
                           bar = New String(barEmptyChar, barWidth)
                       End If

                       Dim progressLine As String = String.Format(
                           CultureInfo.InvariantCulture,
                           " downloading {0}:{1}{2}{3}",
                           displayName,
                           suffixBeforeBar,
                           bar,
                           suffixAfterBar
                       )

                       Dim paddedLine As String = progressLine
                       If lastLineLength > progressLine.Length Then
                           paddedLine &= New String(" "c, lastLineLength - progressLine.Length)
                       End If
                       lastLineLength = progressLine.Length

                       Console.Error.Write(vbCr & paddedLine)
                       If isFinal Then
                           Console.Error.WriteLine()
                       End If
                   End Sub
        End Function

        Private Shared Function FormatBinaryAmount(value As Long) As String
            Dim size As Double = CDbl(Math.Max(0L, value))
            Dim units As String() = {"B", "KiB", "MiB", "GiB", "TiB"}
            Dim idx As Integer = 0
            While size >= 1024.0R AndAlso idx < units.Length - 1
                size /= 1024.0R
                idx += 1
            End While

            If idx = 0 Then
                Return String.Format(CultureInfo.InvariantCulture, "{0:0}{1}", size, units(idx))
            End If
            Return String.Format(CultureInfo.InvariantCulture, "{0:0.00}{1}", size, units(idx))
        End Function

        Private Shared Function FormatRate(bytesPerSecond As Double) As String
            If bytesPerSecond <= 0 Then
                Return "0.00B/s"
            End If
            Dim value As Double = bytesPerSecond
            Dim units As String() = {"B/s", "KiB/s", "MiB/s", "GiB/s"}
            Dim idx As Integer = 0
            While value >= 1024.0R AndAlso idx < units.Length - 1
                value /= 1024.0R
                idx += 1
            End While
            Return String.Format(CultureInfo.InvariantCulture, "{0:0.00}{1}", value, units(idx))
        End Function

        Private Shared Function FormatEta(totalSeconds As Integer) As String
            If totalSeconds < 0 Then
                totalSeconds = 0
            End If
            Dim span As TimeSpan = TimeSpan.FromSeconds(totalSeconds)
            If span.TotalHours >= 1.0R Then
                Dim totalHours As Integer = CInt(Math.Floor(span.TotalHours))
                Return String.Format(
                    CultureInfo.InvariantCulture,
                    "{0:00}:{1:00}:{2:00}",
                    totalHours,
                    span.Minutes,
                    span.Seconds
                )
            End If
            Return String.Format(
                CultureInfo.InvariantCulture,
                "{0:00}:{1:00}",
                CInt(Math.Floor(span.TotalMinutes)),
                span.Seconds
            )
        End Function

        Private Shared Function FormatElapsed(elapsed As TimeSpan) As String
            If elapsed.TotalHours >= 1.0R Then
                Dim totalHours As Integer = CInt(Math.Floor(elapsed.TotalHours))
                Return String.Format(
                    CultureInfo.InvariantCulture,
                    "{0:00}:{1:00}:{2:00}",
                    totalHours,
                    elapsed.Minutes,
                    elapsed.Seconds
                )
            End If

            Dim totalMinutes As Integer = CInt(Math.Floor(elapsed.TotalMinutes))
            Return String.Format(
                CultureInfo.InvariantCulture,
                "{0:00}:{1:00}",
                totalMinutes,
                elapsed.Seconds
            )
        End Function

        Private Shared Function UseUnicodeProgressBar() As Boolean
            Try
                Dim forceAsciiRaw As String = Environment.GetEnvironmentVariable("IA_PROGRESS_ASCII")
                If Not String.IsNullOrWhiteSpace(forceAsciiRaw) Then
                    Dim normalized As String = forceAsciiRaw.Trim().ToLowerInvariant()
                    If normalized = "1" OrElse normalized = "true" OrElse normalized = "yes" Then
                        Return False
                    End If
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException("Missing value for " & optionName)
            End If
        End Sub

        Private Shared Function EnsurePathWithinBase(candidatePath As String, destDir As String) As Boolean
            Dim baseDir As String = If(String.IsNullOrWhiteSpace(destDir), Directory.GetCurrentDirectory(), destDir)
            Dim baseFull As String = Path.GetFullPath(baseDir)
            Dim candidate As String = Path.GetFullPath(candidatePath)
            Return candidate.StartsWith(baseFull, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function EscapePathKeepSlash(path As String) As String
            Dim parts = path.Split("/"c)
            For i As Integer = 0 To parts.Length - 1
                parts(i) = Uri.EscapeDataString(parts(i))
            Next
            Return String.Join("/", parts)
        End Function

        Private Shared Function IsDarkItem(metadata As Dictionary(Of String, Object)) As Boolean
            If metadata Is Nothing Then
                Return False
            End If

            Dim value As Object = Nothing
            If metadata.TryGetValue("is_dark", value) Then
                Dim boolValue As Boolean
                If TypeOf value Is Boolean Then
                    Return CBool(value)
                End If
                If Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), boolValue) Then
                    Return boolValue
                End If
            End If
            Return False
        End Function

        Private Shared Function GetCandidateFiles(identifier As String, itemMetadata As Dictionary(Of String, Object), parsed As DownloadArgs) As List(Of ArchiveFileEntry)
            Dim files As List(Of ArchiveFileEntry) = ApiShared.GetArchiveFileEntries(itemMetadata)
            If parsed.OnTheFly Then
                files.Add(New ArchiveFileEntry(New Dictionary(Of String, Object) From {{"name", identifier & ".epub"}, {"format", "EPUB"}, {"source", "derivative"}}))
                files.Add(New ArchiveFileEntry(New Dictionary(Of String, Object) From {{"name", identifier & ".mobi"}, {"format", "MOBI"}, {"source", "derivative"}}))
                files.Add(New ArchiveFileEntry(New Dictionary(Of String, Object) From {{"name", identifier & "_daisy.zip"}, {"format", "DAISY"}, {"source", "derivative"}}))
            End If

            Dim filtered As New List(Of ArchiveFileEntry)()
            For Each f In files
                Dim name As String = f.Name
                Dim format As String = f.Format
                Dim source As String = f.Source

                If parsed.Files.Count > 0 AndAlso Not parsed.Files.Contains(name) Then
                    Continue For
                End If
                If parsed.Formats.Count > 0 AndAlso Not parsed.Formats.Contains(format) Then
                    Continue For
                End If
                If Not String.IsNullOrWhiteSpace(parsed.GlobPattern) AndAlso Not GlobMatches(name, parsed.GlobPattern) Then
                    Continue For
                End If
                If Not String.IsNullOrWhiteSpace(parsed.ExcludePattern) AndAlso GlobMatches(name, parsed.ExcludePattern) Then
                    Continue For
                End If
                If Not parsed.DownloadHistory AndAlso source.StartsWith("history/", StringComparison.Ordinal) Then
                    Continue For
                End If
                If parsed.Source.Count > 0 AndAlso Not parsed.Source.Contains(source) Then
                    Continue For
                End If
                If parsed.ExcludeSource.Count > 0 AndAlso parsed.ExcludeSource.Contains(source) Then
                    Continue For
                End If
                filtered.Add(f)
            Next
            Return filtered
        End Function

        Private Shared Function GlobMatches(value As String, patternGroup As String) As Boolean
            Dim patterns = patternGroup.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each pattern In patterns
                Dim escaped = Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".")
                If Regex.IsMatch(value, "^" & escaped & "$", RegexOptions.IgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Sub MergeQueryString(destination As Dictionary(Of String, Object), raw As String, optionName As String)
            Dim normalized As String = raw
            If normalized.IndexOf("="c) < 0 AndAlso normalized.IndexOf(":"c) >= 0 Then
                Dim firstColon As Integer = normalized.IndexOf(":"c)
                normalized = normalized.Substring(0, firstColon) & "=" & normalized.Substring(firstColon + 1)
            End If
            Dim eq As Integer = normalized.IndexOf("="c)
            If eq <= 0 OrElse eq = normalized.Length - 1 Then
                Throw New ArgumentException(
                    String.Format("{0} must be formatted as 'key=value' or 'key:value'", optionName)
                )
            End If
            Dim key As String = Uri.UnescapeDataString(normalized.Substring(0, eq))
            Dim value As String = Uri.UnescapeDataString(normalized.Substring(eq + 1))
            If destination.ContainsKey(key) Then
                Dim list = If(TryCast(destination(key), List(Of String)), New List(Of String) From {Convert.ToString(destination(key), CultureInfo.InvariantCulture)})
                list.Add(value)
                destination(key) = list
            Else
                destination(key) = value
            End If
        End Sub

        Private Shared Function ParseArguments(args As IList(Of String)) As DownloadArgs
            Dim parsed As New DownloadArgs With {
                .Files = New List(Of String)(),
                .SearchParameters = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Formats = New List(Of String)(),
                .RequestParameters = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Source = New List(Of String)(),
                .ExcludeSource = New List(Of String)()
            }

            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help" : parsed.ShowHelp = True
                    Case "-q", "--quiet" : parsed.Quiet = True
                    Case "-d", "--dry-run" : parsed.DryRun = True
                    Case "-i", "--ignore-existing" : parsed.IgnoreExisting = True
                    Case "-C", "--checksum" : parsed.Checksum = True
                    Case "--checksum-archive" : parsed.ChecksumArchive = True
                    Case "-R", "--retries"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.Retries = Integer.Parse(args(i), CultureInfo.InvariantCulture)
                    Case "-I", "--itemlist"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.ItemlistPath = args(i)
                    Case "-S", "--search"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.SearchQuery = args(i)
                    Case "-P", "--search-parameters"
                        i += 1 : EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.SearchParameters, args(i), current)
                    Case "-g", "--glob"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.GlobPattern = args(i)
                    Case "-e", "--exclude"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.ExcludePattern = args(i)
                    Case "-f", "--format"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.Formats.Add(args(i))
                    Case "--on-the-fly" : parsed.OnTheFly = True
                    Case "--no-directories" : parsed.NoDirectories = True
                    Case "--destdir"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.DestDir = args(i)
                    Case "-s", "--stdout" : parsed.StdoutOutput = True
                    Case "--no-change-timestamp" : parsed.NoChangeTimestamp = True
                    Case "-p", "--parameters"
                        i += 1 : EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.RequestParameters, args(i), current)
                    Case "-a", "--download-history" : parsed.DownloadHistory = True
                    Case "--source"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.Source.Add(args(i))
                    Case "--exclude-source"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.ExcludeSource.Add(args(i))
                    Case "-t", "--timeout"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.TimeoutSeconds = Double.Parse(args(i), CultureInfo.InvariantCulture)
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            unknown.Add(current)
                            i += 1
                            Continue While
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                            parsed.Identifier = current
                        Else
                            parsed.Files.Add(current)
                        End If
                End Select
                i += 1
            End While

            If Not String.IsNullOrWhiteSpace(parsed.DestDir) AndAlso Not Directory.Exists(parsed.DestDir) Then
                Throw New ArgumentException("'" & parsed.DestDir & "' is not a valid directory")
            End If
            If unknown.Count > 0 Then
                Throw New ArgumentException("unrecognized arguments: " & String.Join(" ", unknown))
            End If
            Return parsed
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " download [-h] [-q] [-d] [-i] [-C] [--checksum-archive]")
            Console.Error.WriteLine("                      [-R RETRIES] [-I ITEMLIST] [-S SEARCH] [-P KEY:VALUE]")
            Console.Error.WriteLine("                      [-g GLOB] [-e EXCLUDE] [-f FORMAT] [--on-the-fly]")
            Console.Error.WriteLine("                      [--no-directories] [--destdir DESTDIR] [-s]")
            Console.Error.WriteLine("                      [--no-change-timestamp] [-p KEY:VALUE] [-a]")
            Console.Error.WriteLine("                      [--source SOURCE] [--exclude-source EXCLUDE_SOURCE]")
            Console.Error.WriteLine("                      [-t TIMEOUT]")
            Console.Error.WriteLine("                      [identifier] [file ...]")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            PrintHelp()
            Console.Error.WriteLine(CliApp.ExecutableName() & " download: error: " & message)
        End Sub

        Private Shared Function ResolveDestinationPath(relativePath As String, destDir As String) As String
            Dim baseDir As String = If(String.IsNullOrWhiteSpace(destDir), Directory.GetCurrentDirectory(), destDir)
            Return Path.GetFullPath(Path.Combine(baseDir, relativePath.Replace("/"c, Path.DirectorySeparatorChar)))
        End Function

        Private Shared Function ResolveIdentifiers(session As ArchiveSession, parsed As DownloadArgs) As List(Of String)
            Dim ids As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(parsed.ItemlistPath) Then
                For Each line In File.ReadLines(parsed.ItemlistPath)
                    Dim trimmed As String = line.Trim()
                    If trimmed.Length > 0 Then
                        ids.Add(trimmed)
                    End If
                Next
                Return ids
            End If

            If Not String.IsNullOrWhiteSpace(parsed.SearchQuery) Then
                Dim results = session.SearchItemsSimple(parsed.SearchQuery, parsed.SearchParameters)
                For Each r In results
                    If r.ContainsKey("identifier") Then
                        ids.Add(Convert.ToString(r("identifier"), CultureInfo.InvariantCulture))
                    End If
                Next
                Return ids
            End If

            If parsed.Identifier = "-" Then
                While True
                    Dim line As String = Console.In.ReadLine()
                    If line Is Nothing Then
                        Exit While
                    End If
                    Dim trimmed As String = line.Trim()
                    If trimmed.Length > 0 Then
                        ids.Add(trimmed)
                    End If
                End While
                Return ids
            End If

            If Not String.IsNullOrWhiteSpace(parsed.Identifier) Then
                If parsed.Identifier.IndexOf("/"c) >= 0 Then
                    Dim parts = parsed.Identifier.Split(New Char() {"/"c}, 2)
                    ids.Add(parts(0))
                    parsed.Files = New List(Of String) From {parts(1)}
                Else
                    ids.Add(parsed.Identifier)
                End If
            End If

            Return ids
        End Function

        Private Shared Function LoadChecksumArchiveEntries(archiveFile As String) As HashSet(Of String)
            Dim entries As New HashSet(Of String)(StringComparer.Ordinal)
            If Not File.Exists(archiveFile) Then
                Return entries
            End If

            For Each line In File.ReadLines(archiveFile)
                If line.Length > 0 Then
                    entries.Add(line)
                End If
            Next
            Return entries
        End Function

        Private Shared Function ShouldSkipExisting(destinationPath As String, fileEntry As ArchiveFileEntry, parsed As DownloadArgs, checksumArchiveEntries As HashSet(Of String)) As Boolean
            If parsed.StdoutOutput Then
                Return False
            End If
            If Not File.Exists(destinationPath) Then
                Return False
            End If

            If parsed.ChecksumArchive Then
                If checksumArchiveEntries IsNot Nothing AndAlso checksumArchiveEntries.Contains(destinationPath) Then
                    If Not parsed.Quiet Then
                        Console.Error.WriteLine(" skipping " & destinationPath & ", file already exists based on checksum_archive.")
                    End If
                    Return True
                End If
            End If

            If parsed.IgnoreExisting Then
                If Not parsed.Quiet Then
                    Console.Error.WriteLine(" skipping " & destinationPath & ", file already exists.")
                End If
                Return True
            End If

            If parsed.Checksum OrElse parsed.ChecksumArchive Then
                Dim md5Remote As String = fileEntry.Md5
                If md5Remote.Length > 0 Then
                    Dim md5Local = ComputeMd5(destinationPath)
                    If String.Equals(md5Local, md5Remote, StringComparison.OrdinalIgnoreCase) Then
                        If Not parsed.Quiet Then
                            Console.Error.WriteLine(" skipping " & destinationPath & ", file already exists based on checksum.")
                        End If
                        If parsed.ChecksumArchive Then
                            If checksumArchiveEntries Is Nothing OrElse checksumArchiveEntries.Add(destinationPath) Then
                                File.AppendAllText("_checksum_archive.txt", destinationPath & Environment.NewLine)
                            End If
                        End If
                        Return True
                    End If
                End If
            End If

            Return False
        End Function

        Private Shared Sub ValidateArgs(parsed As DownloadArgs)
            If Not String.IsNullOrWhiteSpace(parsed.ItemlistPath) AndAlso
               Not String.IsNullOrWhiteSpace(parsed.SearchQuery) Then
                Throw New ArgumentException("--itemlist and --search cannot be used together")
            End If

            If Not String.IsNullOrWhiteSpace(parsed.ItemlistPath) OrElse
               Not String.IsNullOrWhiteSpace(parsed.SearchQuery) Then
                If Not String.IsNullOrWhiteSpace(parsed.Identifier) Then
                    Throw New ArgumentException("Cannot specify an identifier with --itemlist/--search")
                End If
                If parsed.Files.Count > 0 Then
                    Throw New ArgumentException("Cannot specify files with --itemlist/--search")
                End If
            Else
                If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                    Throw New ArgumentException("Identifier is required when not using --itemlist/--search")
                End If
            End If

            If Not String.IsNullOrWhiteSpace(parsed.ItemlistPath) Then
                If Not File.Exists(parsed.ItemlistPath) Then
                    Throw New ArgumentException("No such file: " & parsed.ItemlistPath)
                End If
                Dim nonEmpty As Boolean = False
                For Each line In File.ReadLines(parsed.ItemlistPath)
                    If line.Trim().Length > 0 Then
                        nonEmpty = True
                        Exit For
                    End If
                Next
                If Not nonEmpty Then
                    Throw New ArgumentException("--itemlist file is empty or contains only whitespace")
                End If
            End If
        End Sub

        Private NotInheritable Class DownloadArgs
            Public Property Checksum As Boolean
            Public Property ChecksumArchive As Boolean
            Public Property DestDir As String
            Public Property DownloadHistory As Boolean
            Public Property DryRun As Boolean
            Public Property ExcludePattern As String
            Public Property ExcludeSource As List(Of String)
            Public Property Files As List(Of String)
            Public Property Formats As List(Of String)
            Public Property GlobPattern As String
            Public Property Identifier As String
            Public Property IgnoreExisting As Boolean
            Public Property ItemlistPath As String
            Public Property NoChangeTimestamp As Boolean
            Public Property NoDirectories As Boolean
            Public Property OnTheFly As Boolean
            Public Property Quiet As Boolean
            Public Property RequestParameters As Dictionary(Of String, Object)
            Public Property Retries As Integer = 5
            Public Property SearchParameters As Dictionary(Of String, Object)
            Public Property SearchQuery As String
            Public Property ShowHelp As Boolean
            Public Property Source As List(Of String)
            Public Property StdoutOutput As Boolean
            Public Property TimeoutSeconds As Double?
        End Class
    End Class
End Namespace

Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports Microsoft.VisualBasic.FileIO
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class UploadCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As UploadArgs
            Try
                parsed = ParseArguments(args)
                If Not parsed.ShowHelp Then
                    CheckIfFileArgRequired(parsed)
                End If
            Catch ex As ArgumentException
                PrintUsageAndError(ex.Message)
                Return 2
            End Try
            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If

            Dim stdinUpload As Boolean = IsUploadingFromStdin(parsed)
            If stdinUpload AndAlso String.IsNullOrWhiteSpace(parsed.RemoteName) Then
                Throw New ArgumentException(
                    "When uploading from stdin, you must specify a remote filename with --remote-name"
                )
            End If

            If parsed.StatusCheck Then
                Dim overloaded As Boolean = session.S3IsOverloaded(parsed.Identifier)
                If overloaded Then
                    Console.Error.WriteLine(
                        String.Format(
                            "warning: {0} is over limit, and not accepting requests. Expect 503 SlowDown errors.",
                            parsed.Identifier
                        )
                    )
                    Return 1
                End If
                Console.Error.WriteLine(
                    String.Format("success: {0} is accepting requests.", parsed.Identifier)
                )
                Return 0
            End If

            Dim verbose As Boolean = Not parsed.Quiet
            Dim queueDerive As Boolean = Not parsed.NoDerive

            Dim headers As Dictionary(Of String, String) = ApiShared.ConvertHeaderValues(parsed.Headers)
            If Not String.IsNullOrWhiteSpace(parsed.SizeHint) Then
                headers("x-archive-size-hint") = parsed.SizeHint
            End If
            If Not headers.ContainsKey("x-archive-keep-old-version") AndAlso Not parsed.NoBackup Then
                headers("x-archive-keep-old-version") = "1"
            End If

            If Not String.IsNullOrWhiteSpace(parsed.SpreadsheetPath) Then
                Return ExecuteSpreadsheetUpload(session, parsed, headers, queueDerive, verbose)
            End If

            If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                Throw New ArgumentException("Identifier is required for upload.")
            End If

            Dim fileMetadataEntries = LoadFileMetadataEntries(parsed.FileMetadataPath)
            Dim targets = BuildUploadTargets(parsed, stdinUpload, fileMetadataEntries)
            If targets.Count = 0 Then
                Throw New FileNotFoundException("No valid file was found. Check your paths.")
            End If

            Dim errors As Boolean = False
            If verbose Then
                Console.Error.WriteLine(parsed.Identifier & ":")
            End If

            Dim remoteMd5Map As New Dictionary(Of String, String)(StringComparer.Ordinal)
            If parsed.Checksum Then
                Dim itemMetadata = session.GetItemMetadata(parsed.Identifier, Nothing)
                remoteMd5Map = BuildRemoteMd5Map(itemMetadata)
            End If
            Dim retries As Integer = If(parsed.Retries, 0)
            Dim retriesSleep As Integer = If(parsed.SleepSeconds, 30)

            For Each target In targets
                Dim contentBytes As Byte() = GetTargetContentBytes(target)
                Dim localMd5 As String = ComputeMd5(contentBytes)
                If parsed.Checksum AndAlso remoteMd5Map.ContainsKey(target.RemoteName) Then
                    If String.Equals(remoteMd5Map(target.RemoteName), localMd5, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                Dim response = session.UploadS3File(
                    parsed.Identifier,
                    target.RemoteName,
                    contentBytes,
                    headers,
                    parsed.Metadata,
                    target.FileMetadata,
                    queueDerive,
                    Math.Max(1, retries),
                    retriesSleep,
                    parsed.Debug,
                    BuildProgressReporter(target.RemoteName, contentBytes.LongLength, verbose)
                )

                If parsed.Debug Then
                    PrintDebug(response)
                    Continue For
                End If

                If response.StatusCode < 200 OrElse response.StatusCode >= 300 Then
                    Console.Error.WriteLine(
                        String.Format(
                            " error uploading {0}: HTTP {1}",
                            target.RemoteName,
                            response.StatusCode
                        )
                    )
                    errors = True
                    Continue For
                End If

                If parsed.Verify Then
                    Dim etag As String = ""
                    If response.Headers IsNot Nothing AndAlso response.Headers.ContainsKey("ETag") Then
                        etag = response.Headers("ETag").Trim(""""c)
                    End If
                    If etag.Length > 0 AndAlso Not String.Equals(etag, localMd5, StringComparison.OrdinalIgnoreCase) Then
                        Console.Error.WriteLine(
                            String.Format(
                                "error: verify failed for {0} (checksum mismatch).",
                                target.RemoteName
                            )
                        )
                        errors = True
                        Continue For
                    End If
                End If

                If parsed.DeleteAfterUpload AndAlso Not String.IsNullOrWhiteSpace(target.LocalPath) Then
                    Try
                        File.Delete(target.LocalPath)
                    Catch
                    End Try
                End If
            Next

            If parsed.OpenAfterUpload AndAlso Not parsed.Debug Then
                Try
                    Dim detailsUrl As String = String.Format(
                        "{0}//{1}/details/{2}",
                        session.Protocol,
                        session.Host,
                        parsed.Identifier
                    )
                    Dim psi As New ProcessStartInfo(detailsUrl) With {.UseShellExecute = True}
                    Process.Start(psi)
                Catch
                End Try
            End If

            Return If(errors, 1, 0)
        End Function

        Private Shared Function BuildRemoteMd5Map(itemMetadata As Dictionary(Of String, Object)) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.Ordinal)
            For Each entry In ApiShared.GetArchiveFileEntries(itemMetadata)
                Dim name As String = entry.Name
                Dim md5 As String = entry.Md5
                If name.Length > 0 AndAlso md5.Length > 0 Then
                    result(name) = md5
                End If
            Next
            Return result
        End Function

        Private Shared Function BuildUploadTargets(parsed As UploadArgs, stdinUpload As Boolean, fileMetadataEntries As List(Of Dictionary(Of String, Object))) As List(Of UploadTarget)
            Dim targets As New List(Of UploadTarget)()

            If stdinUpload Then
                Dim bytes = ReadAllStdinBytes()
                targets.Add(New UploadTarget With {
                    .LocalPath = "",
                    .RemoteName = parsed.RemoteName,
                    .FileMetadata = Nothing,
                    .ContentBytes = bytes
                })
                Return targets
            End If

            If fileMetadataEntries IsNot Nothing AndAlso fileMetadataEntries.Count > 0 Then
                For Each entry In fileMetadataEntries
                    Dim localPath As String = GetString(entry, "name")
                    If String.IsNullOrWhiteSpace(localPath) Then
                        Continue For
                    End If
                    Dim fileMetadata As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    For Each kvp In entry
                        If String.Equals(kvp.Key, "name", StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If
                        fileMetadata(kvp.Key) = kvp.Value
                    Next
                    Dim remoteName As String = Path.GetFileName(localPath)
                    targets.Add(New UploadTarget With {
                        .LocalPath = localPath,
                        .RemoteName = remoteName,
                        .FileMetadata = fileMetadata,
                        .ContentBytes = Nothing
                    })
                Next
                Return targets
            End If

            Dim expanded As New List(Of KeyValuePair(Of String, String))()
            For Each inputPath In parsed.Files
                If inputPath = "." Then
                    For Each entry In Directory.EnumerateFiles(Directory.GetCurrentDirectory())
                        expanded.Add(New KeyValuePair(Of String, String)(entry, ""))
                    Next
                    Continue For
                End If
                If Directory.Exists(inputPath) Then
                    Dim fullBase As String = Path.GetFullPath(inputPath).TrimEnd("\"c, "/"c)
                    Dim normalizedBase As String = fullBase.Replace("\"c, "/"c)
                    For Each entry As String In Directory.EnumerateFiles(
                        inputPath,
                        "*",
                        System.IO.SearchOption.AllDirectories
                    )
                        Dim fullEntry As String = Path.GetFullPath(entry)
                        Dim relative As String
                        If fullEntry.StartsWith(fullBase, StringComparison.OrdinalIgnoreCase) Then
                            relative = fullEntry.Substring(fullBase.Length)
                        Else
                            relative = Path.GetFileName(fullEntry)
                        End If
                        relative = relative.TrimStart("\"c, "/"c)
                        relative = relative.Replace("\"c, "/"c)
                        Dim remoteOverride As String = normalizedBase & "/" & relative
                        expanded.Add(New KeyValuePair(Of String, String)(entry, remoteOverride))
                    Next
                ElseIf File.Exists(inputPath) Then
                    expanded.Add(New KeyValuePair(Of String, String)(inputPath, ""))
                End If
            Next

            If parsed.RemoteName IsNot Nothing AndAlso expanded.Count > 0 Then
                expanded = New List(Of KeyValuePair(Of String, String)) From {expanded(0)}
            End If

            For Each entry In expanded
                Dim localPath As String = entry.Key
                Dim remoteName As String
                If Not String.IsNullOrWhiteSpace(parsed.RemoteName) Then
                    remoteName = parsed.RemoteName
                ElseIf Not String.IsNullOrWhiteSpace(entry.Value) Then
                    remoteName = entry.Value
                ElseIf parsed.KeepDirectories Then
                    remoteName = localPath.Replace("\"c, "/"c)
                Else
                    remoteName = Path.GetFileName(localPath)
                End If

                targets.Add(New UploadTarget With {
                    .LocalPath = localPath,
                    .RemoteName = remoteName,
                    .FileMetadata = Nothing,
                    .ContentBytes = Nothing
                })
            Next

            Return targets
        End Function

        Private Shared Sub CheckIfFileArgRequired(parsed As UploadArgs)
            Dim requiredIfNoFile As Boolean =
                Not String.IsNullOrWhiteSpace(parsed.SpreadsheetPath) OrElse
                Not String.IsNullOrWhiteSpace(parsed.FileMetadataPath) OrElse
                parsed.StatusCheck
            If parsed.Files.Count = 0 AndAlso Not requiredIfNoFile Then
                Throw New ArgumentException("You must specify a file to upload.")
            End If
        End Sub

        Private Shared Function ComputeMd5(bytes As Byte()) As String
            Using hasher As MD5 = MD5.Create()
                Dim hash = hasher.ComputeHash(bytes)
                Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Function GetTargetContentBytes(target As UploadTarget) As Byte()
            If target.ContentBytes IsNot Nothing Then
                Return target.ContentBytes
            End If
            If String.IsNullOrWhiteSpace(target.LocalPath) Then
                Return Array.Empty(Of Byte)()
            End If
            Return File.ReadAllBytes(target.LocalPath)
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException("Missing value for " & optionName)
            End If
        End Sub

        Private Shared Function BuildProgressReporter(
            remoteName As String,
            totalBytes As Long,
            verbose As Boolean
        ) As Action(Of Long, Long)
            If Not verbose Then
                Return Nothing
            End If

            Dim lineCompleted As Boolean = False
            Dim sw As Stopwatch = Stopwatch.StartNew()
            Dim minBarWidth As Integer = 10
            Dim maxBarWidth As Integer = 40
            Dim fixedNameLimit As Integer = 48
            Dim chunkSizeBytes As Long = 1024L * 1024L
            Dim useUnicodeBars As Boolean = UseUnicodeProgressBar()
            Dim barFillChar As Char = If(useUnicodeBars, "█"c, "#"c)
            Dim barEmptyChar As Char = If(useUnicodeBars, "░"c, "-"c)
            Dim lastRenderedMs As Long = -1
            Dim lastLineLength As Integer = 0
            If totalBytes <= 0 Then
                Console.Error.WriteLine(" uploading " & remoteName)
                lineCompleted = True
            End If

            Return Sub(bytesSent As Long, callbackTotalBytes As Long)
                       Dim total As Long = callbackTotalBytes
                       If total < 0 Then
                           total = 0
                       End If
                       If total = 0 Then
                           If Not lineCompleted Then
                               Console.Error.WriteLine(" uploading " & remoteName)
                               lineCompleted = True
                           End If
                           Return
                       End If

                       If bytesSent < 0 Then
                           bytesSent = 0
                       ElseIf bytesSent > total Then
                           bytesSent = total
                       End If

                       Dim isFinal As Boolean = (bytesSent >= total)
                       Dim elapsedMs As Long = sw.ElapsedMilliseconds
                       If Not isFinal AndAlso lastRenderedMs >= 0 AndAlso (elapsedMs - lastRenderedMs) < 100 Then
                           Return
                       End If
                       lastRenderedMs = elapsedMs

                       Dim percent As Double = (CDbl(bytesSent) / CDbl(total)) * 100.0R
                       Dim sentMiB As Double = CDbl(bytesSent) / 1048576.0R
                       Dim totalMiB As Double = CDbl(total) / 1048576.0R
                       Dim totalChunks As Long = CLng(Math.Ceiling(CDbl(total) / CDbl(chunkSizeBytes)))
                       If totalChunks < 1 Then
                           totalChunks = 1
                       End If
                       Dim sentChunks As Long = CLng(Math.Ceiling(CDbl(bytesSent) / CDbl(chunkSizeBytes)))
                       If sentChunks > totalChunks Then
                           sentChunks = totalChunks
                       End If
                       If bytesSent = 0 Then
                           sentChunks = 0
                       End If
                       Dim elapsedSeconds As Double = Math.Max(sw.Elapsed.TotalSeconds, 0.000001R)
                       Dim speedBytesPerSecond As Double = CDbl(bytesSent) / elapsedSeconds
                       Dim speedValue As Double = speedBytesPerSecond
                       Dim speedUnit As String = "B/s"
                       If speedBytesPerSecond >= 1024.0R Then
                           speedValue = speedBytesPerSecond / 1024.0R
                           speedUnit = "KiB/s"
                       End If
                       If speedBytesPerSecond >= 1048576.0R Then
                           speedValue = speedBytesPerSecond / 1048576.0R
                           speedUnit = "MiB/s"
                       End If
                       If speedBytesPerSecond >= 1073741824.0R Then
                           speedValue = speedBytesPerSecond / 1073741824.0R
                           speedUnit = "GiB/s"
                       End If

                       Dim remainingBytes As Long = Math.Max(0L, total - bytesSent)
                       Dim etaText As String = "--:--"
                       If speedBytesPerSecond > 0.000001R AndAlso remainingBytes > 0 Then
                           Dim etaTotalSeconds As Integer = CInt(Math.Ceiling(CDbl(remainingBytes) / speedBytesPerSecond))
                           Dim etaSpan As TimeSpan = TimeSpan.FromSeconds(etaTotalSeconds)
                           etaText = If(
                               etaSpan.TotalHours >= 1.0R,
                               etaSpan.ToString("hh\:mm\:ss", CultureInfo.InvariantCulture),
                               etaSpan.ToString("mm\:ss", CultureInfo.InvariantCulture)
                           )
                       ElseIf remainingBytes = 0 Then
                           etaText = "00:00"
                       End If

                       Dim elapsedSpan As TimeSpan = sw.Elapsed
                       Dim elapsedText As String = FormatElapsed(elapsedSpan)

                       Dim percentWhole As Double = Math.Round(percent, MidpointRounding.AwayFromZero)
                       If percentWhole < 0 Then
                           percentWhole = 0
                       ElseIf percentWhole > 100 Then
                           percentWhole = 100
                       End If

                       Dim suffixBeforeBar As String = String.Format(
                           CultureInfo.InvariantCulture,
                           " {0,3:0}%|",
                           percentWhole
                       )
                       Dim suffixAfterBar As String = String.Format(
                           CultureInfo.InvariantCulture,
                           "| {0}/{1} [{2}<{3}, {4:0.00}{5}]",
                           sentChunks,
                           totalChunks,
                           elapsedText,
                           etaText,
                           speedValue,
                           speedUnit
                       )

                       Dim width As Integer = 0
                       Try
                           width = Console.WindowWidth
                       Catch
                           width = 0
                       End Try

                       Dim displayName As String = remoteName
                       If width > 0 Then
                           Dim reservedNoName As Integer =
                               " uploading ".Length + 2 + suffixBeforeBar.Length + suffixAfterBar.Length + minBarWidth
                           Dim maxNameLen As Integer = width - reservedNoName
                           If maxNameLen < 8 Then
                               maxNameLen = 8
                           End If
                           If displayName.Length > maxNameLen Then
                               displayName = displayName.Substring(0, maxNameLen - 1) & "~"
                           End If
                       ElseIf displayName.Length > fixedNameLimit Then
                           displayName = displayName.Substring(0, fixedNameLimit - 1) & "~"
                       End If

                       Dim barWidth As Integer = 24
                       If width > 0 Then
                           Dim reserved As Integer =
                               " uploading ".Length + displayName.Length + 2 + suffixBeforeBar.Length + suffixAfterBar.Length
                           Dim calculatedBar As Integer = width - reserved - 1
                           barWidth = Math.Max(minBarWidth, Math.Min(maxBarWidth, calculatedBar))
                       End If

                       Dim filled As Integer = CInt(Math.Floor((percent / 100.0R) * barWidth))
                       If filled < 0 Then
                           filled = 0
                       ElseIf filled > barWidth Then
                           filled = barWidth
                       End If
                       Dim bar As String = New String(barFillChar, filled) & New String(barEmptyChar, barWidth - filled)

                       Dim progressLine As String = String.Format(
                           CultureInfo.InvariantCulture,
                           " uploading {0}:{1}{2}{3}",
                           displayName,
                           suffixBeforeBar,
                           bar,
                           suffixAfterBar
                       )
                       progressLine = progressLine.TrimEnd()

                       Dim paddedLine As String = progressLine
                       If lastLineLength > progressLine.Length Then
                           paddedLine &= New String(" "c, lastLineLength - progressLine.Length)
                       End If
                       lastLineLength = progressLine.Length

                       Console.Error.Write(vbCr & paddedLine)
                       If isFinal Then
                           Console.Error.WriteLine()
                           lineCompleted = True
                       End If
                   End Sub
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

                ' Match Python tqdm behavior: prefer unicode bars by default.
                ' Keep ASCII available via IA_PROGRESS_ASCII=1 for legacy terminals.
                Dim outputEncoding = Console.OutputEncoding
                If outputEncoding Is Nothing Then
                    Return False
                End If
                Return True
            Catch
                Return False
            End Try
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

        Private Shared Function ExecuteSpreadsheetUpload(
            session As ArchiveSession,
            parsed As UploadArgs,
            headers As Dictionary(Of String, String),
            queueDerive As Boolean,
            verbose As Boolean
        ) As Integer
            Dim rows = ReadCsvRows(parsed.SpreadsheetPath)
            Dim errors As Boolean = False
            Dim prevIdentifier As String = Nothing

            For Each row In rows
                For Each metadataKey In row.Keys
                    If Not IsValidMetadataKey(metadataKey) Then
                        Console.Error.WriteLine(
                            String.Format("error: '{0}' is not a valid metadata key.", metadataKey)
                        )
                        Return 1
                    End If
                Next

                Dim identifier As String = Nothing
                If row.ContainsKey("item") Then
                    identifier = row("item")
                End If
                If String.IsNullOrWhiteSpace(identifier) AndAlso row.ContainsKey("identifier") Then
                    identifier = row("identifier")
                End If
                If String.IsNullOrWhiteSpace(identifier) Then
                    If String.IsNullOrWhiteSpace(prevIdentifier) Then
                        Console.Error.WriteLine("error: no identifier column on spreadsheet.")
                        Return 1
                    End If
                    identifier = prevIdentifier
                End If

                Dim localFile As String = ""
                If row.ContainsKey("file") Then
                    localFile = row("file")
                End If
                If String.IsNullOrWhiteSpace(localFile) Then
                    Console.Error.WriteLine("error: no file column on spreadsheet.")
                    Return 1
                End If

                Dim remoteName As String = Nothing
                If row.ContainsKey("REMOTE_NAME") AndAlso Not String.IsNullOrWhiteSpace(row("REMOTE_NAME")) Then
                    remoteName = row("REMOTE_NAME")
                End If

                Dim rowMetadata As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each kvp In parsed.Metadata
                    rowMetadata(kvp.Key) = kvp.Value
                Next
                For Each kvp In row
                    Dim key As String = kvp.Key
                    If String.Equals(key, "file", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(key, "identifier", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(key, "item", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(key, "REMOTE_NAME", StringComparison.Ordinal) Then
                        Continue For
                    End If
                    If Not String.IsNullOrWhiteSpace(kvp.Value) Then
                        rowMetadata(key.ToLowerInvariant()) = kvp.Value
                    End If
                Next

                Dim tempParsed As New UploadArgs With {
                    .Identifier = identifier,
                    .Files = New List(Of String) From {localFile},
                    .RemoteName = remoteName,
                    .KeepDirectories = parsed.KeepDirectories
                }
                Dim targets = BuildUploadTargets(tempParsed, False, Nothing)
                Dim remoteMd5Map As New Dictionary(Of String, String)(StringComparer.Ordinal)
                If parsed.Checksum Then
                    Dim itemMetadata = session.GetItemMetadata(identifier, Nothing)
                    remoteMd5Map = BuildRemoteMd5Map(itemMetadata)
                End If
                Dim retries As Integer = If(parsed.Retries, 0)
                Dim retriesSleep As Integer = If(parsed.SleepSeconds, 30)

                If verbose AndAlso Not String.Equals(prevIdentifier, identifier, StringComparison.Ordinal) Then
                    Console.Error.WriteLine(identifier & ":")
                End If

                For Each target In targets
                    Dim contentBytes As Byte() = GetTargetContentBytes(target)
                    Dim localMd5 As String = ComputeMd5(contentBytes)
                    If parsed.Checksum AndAlso remoteMd5Map.ContainsKey(target.RemoteName) Then
                        If String.Equals(remoteMd5Map(target.RemoteName), localMd5, StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If
                    End If

                    Dim response = session.UploadS3File(
                        identifier,
                        target.RemoteName,
                        contentBytes,
                        headers,
                        rowMetadata,
                        target.FileMetadata,
                        queueDerive,
                        Math.Max(1, retries),
                        retriesSleep,
                        parsed.Debug,
                        BuildProgressReporter(target.RemoteName, contentBytes.LongLength, verbose)
                    )

                    If parsed.Debug Then
                        PrintDebug(response)
                        Exit For
                    End If

                    If response.StatusCode < 200 OrElse response.StatusCode >= 300 Then
                        Console.Error.WriteLine(
                            String.Format(
                                " error uploading {0}: HTTP {1}",
                                target.RemoteName,
                                response.StatusCode
                            )
                        )
                        errors = True
                        Continue For
                    End If

                    If parsed.Verify Then
                        Dim etag As String = ""
                        If response.Headers IsNot Nothing AndAlso response.Headers.ContainsKey("ETag") Then
                            etag = response.Headers("ETag").Trim(""""c)
                        End If
                        If etag.Length > 0 AndAlso Not String.Equals(etag, localMd5, StringComparison.OrdinalIgnoreCase) Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "error: verify failed for {0} (checksum mismatch).",
                                    target.RemoteName
                                )
                            )
                            errors = True
                            Continue For
                        End If
                    End If

                    If parsed.DeleteAfterUpload AndAlso Not String.IsNullOrWhiteSpace(target.LocalPath) Then
                        Try
                            File.Delete(target.LocalPath)
                        Catch
                        End Try
                    End If

                    If parsed.OpenAfterUpload Then
                        Try
                            Dim detailsUrl As String = String.Format(
                                "{0}//{1}/details/{2}",
                                session.Protocol,
                                session.Host,
                                identifier
                            )
                            Dim psi As New ProcessStartInfo(detailsUrl) With {.UseShellExecute = True}
                            Process.Start(psi)
                        Catch
                        End Try
                    End If
                Next

                prevIdentifier = identifier
            Next

            Return If(errors, 1, 0)
        End Function

        Private Shared Function ExtractMetadataEntries(
            value As Object
        ) As List(Of Dictionary(Of String, Object))
            Dim result As New List(Of Dictionary(Of String, Object))()
            Dim objArr = TryCast(value, Object())
            If objArr IsNot Nothing Then
                For Each entry In objArr
                    Dim dict = TryCast(entry, Dictionary(Of String, Object))
                    If dict IsNot Nothing Then
                        result.Add(dict)
                    End If
                Next
                Return result
            End If
            Dim arrList = TryCast(value, ArrayList)
            If arrList IsNot Nothing Then
                For Each entry In arrList
                    Dim dict = TryCast(entry, Dictionary(Of String, Object))
                    If dict IsNot Nothing Then
                        result.Add(dict)
                    End If
                Next
            End If
            Return result
        End Function

        Private Shared Function GetString(dict As Dictionary(Of String, Object), key As String) As String
            If dict IsNot Nothing AndAlso dict.ContainsKey(key) Then
                Return Convert.ToString(dict(key), CultureInfo.InvariantCulture)
            End If
            Return ""
        End Function

        Private Shared Function IsUploadingFromStdin(parsed As UploadArgs) As Boolean
            Return parsed.Files.Count = 1 AndAlso parsed.Files(0) = "-"
        End Function

        Private Shared Function IsValidMetadataKey(name As String) As Boolean
            Return Regex.IsMatch(name, "^[A-Za-z][.\-0-9A-Za-z_]+(?:\[[0-9]+\])?$")
        End Function

        Private Shared Function LoadFileMetadataEntries(
            fileMetadataPath As String
        ) As List(Of Dictionary(Of String, Object))
            If String.IsNullOrWhiteSpace(fileMetadataPath) Then
                Return Nothing
            End If

            Dim serializer As New JavaScriptSerializer()
            Dim text As String = File.ReadAllText(fileMetadataPath)
            Try
                Dim deserialized = serializer.DeserializeObject(text)
                Dim entries = ExtractMetadataEntries(deserialized)
                If entries.Count > 0 Then
                    Return entries
                End If
            Catch
            End Try

            Dim jsonlEntries As New List(Of Dictionary(Of String, Object))()
            For Each rawLine In File.ReadLines(fileMetadataPath)
                Dim line As String = rawLine.Trim()
                If line.Length = 0 Then
                    Continue For
                End If
                Dim obj = serializer.DeserializeObject(line)
                Dim dict = TryCast(obj, Dictionary(Of String, Object))
                If dict IsNot Nothing Then
                    jsonlEntries.Add(dict)
                End If
            Next
            Return jsonlEntries
        End Function

        Private Shared Sub MergeMetadata(
            destination As Dictionary(Of String, Object),
            raw As String,
            optionName As String
        )
            Dim normalized As String = raw
            If normalized.IndexOf(":"c) < 0 AndAlso normalized.IndexOf("="c) >= 0 Then
                Dim firstEq As Integer = normalized.IndexOf("="c)
                normalized = normalized.Substring(0, firstEq) & ":" & normalized.Substring(firstEq + 1)
            End If
            Dim sep As Integer = normalized.IndexOf(":"c)
            If sep <= 0 OrElse sep = normalized.Length - 1 Then
                Throw New ArgumentException(String.Format("{0} must be formatted as 'KEY:VALUE'", optionName))
            End If
            Dim key As String = normalized.Substring(0, sep)
            Dim value As String = normalized.Substring(sep + 1)
            If destination.ContainsKey(key) Then
                Dim list = If(TryCast(destination(key), List(Of String)), New List(Of String) From {Convert.ToString(destination(key), CultureInfo.InvariantCulture)})
                list.Add(value)
                destination(key) = list
            Else
                destination(key) = value
            End If
        End Sub

        Private Shared Sub MergeQueryString(
            destination As Dictionary(Of String, Object),
            raw As String,
            optionName As String
        )
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

        Private Shared Function ParseArguments(args As IList(Of String)) As UploadArgs
            Dim parsed As New UploadArgs With {
                .Files = New List(Of String)(),
                .Metadata = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Headers = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            }

            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help" : parsed.ShowHelp = True
                    Case "-q", "--quiet" : parsed.Quiet = True
                    Case "-d", "--debug" : parsed.Debug = True
                    Case "-r", "--remote-name"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.RemoteName = args(i)
                    Case "-m", "--metadata"
                        i += 1 : EnsureHasValue(args, i, current)
                        MergeMetadata(parsed.Metadata, args(i), current)
                    Case "--spreadsheet"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.SpreadsheetPath = args(i)
                    Case "--file-metadata"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.FileMetadataPath = args(i)
                    Case "-H", "--header"
                        i += 1 : EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.Headers, args(i), current)
                    Case "-c", "--checksum" : parsed.Checksum = True
                    Case "-v", "--verify" : parsed.Verify = True
                    Case "-n", "--no-derive" : parsed.NoDerive = True
                    Case "--size-hint"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.SizeHint = args(i)
                    Case "--delete" : parsed.DeleteAfterUpload = True
                    Case "-R", "--retries"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.Retries = Integer.Parse(args(i), CultureInfo.InvariantCulture)
                    Case "-s", "--sleep"
                        i += 1 : EnsureHasValue(args, i, current)
                        parsed.SleepSeconds = Integer.Parse(args(i), CultureInfo.InvariantCulture)
                    Case "--no-collection-check" : parsed.NoCollectionCheck = True
                    Case "-o", "--open-after-upload" : parsed.OpenAfterUpload = True
                    Case "--no-backup" : parsed.NoBackup = True
                    Case "--keep-directories" : parsed.KeepDirectories = True
                    Case "--status-check" : parsed.StatusCheck = True
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) AndAlso current <> "-" Then
                            unknown.Add(current)
                            i += 1
                            Continue While
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                            parsed.Identifier = current
                        Else
                            If current <> "-" AndAlso Not File.Exists(current) AndAlso Not Directory.Exists(current) Then
                                Throw New ArgumentException("'" & current & "' is not a valid file or directory")
                            End If
                            parsed.Files.Add(current)
                        End If
                End Select
                i += 1
            End While

            If unknown.Count > 0 Then
                Throw New ArgumentException("unrecognized arguments: " & String.Join(" ", unknown))
            End If
            Return parsed
        End Function

        Private Shared Sub PrintDebug(response As ApiCallResult)
            If String.IsNullOrWhiteSpace(response.RequestUrl) Then
                Return
            End If
            Console.Error.WriteLine("Endpoint:")
            Console.Error.WriteLine(" " & response.RequestUrl)
            Console.Error.WriteLine()
            Console.Error.WriteLine("HTTP Headers:")
            If response.Headers IsNot Nothing Then
                For Each kvp In response.Headers
                    Console.Error.WriteLine(" " & kvp.Key & ":" & kvp.Value)
                Next
            End If
        End Sub

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " upload [-h] [-q] [-d] [-r REMOTE_NAME] [-m KEY:VALUE]")
            Console.Error.WriteLine("                    [--spreadsheet SPREADSHEET]")
            Console.Error.WriteLine("                    [--file-metadata FILE_METADATA] [-H KEY:VALUE] [-c] [-v]")
            Console.Error.WriteLine("                    [-n] [--size-hint SIZE_HINT] [--delete] [-R RETRIES]")
            Console.Error.WriteLine("                    [-s SLEEP] [--no-collection-check] [-o] [--no-backup]")
            Console.Error.WriteLine("                    [--keep-directories] [--status-check]")
            Console.Error.WriteLine("                    [identifier] [file ...]")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            PrintHelp()
            Console.Error.WriteLine(CliApp.ExecutableName() & " upload: error: " & message)
        End Sub

        Private Shared Function ReadAllStdinBytes() As Byte()
            Using input = Console.OpenStandardInput()
                Using ms As New MemoryStream()
                    input.CopyTo(ms)
                    Return ms.ToArray()
                End Using
            End Using
        End Function

        Private Shared Function ReadCsvRows(path As String) As List(Of Dictionary(Of String, String))
            Dim rows As New List(Of Dictionary(Of String, String))()
            Using parser As New TextFieldParser(path, System.Text.Encoding.UTF8)
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                If parser.EndOfData Then
                    Return rows
                End If
                Dim headers = parser.ReadFields()
                If headers Is Nothing Then
                    Return rows
                End If

                While Not parser.EndOfData
                    Dim fields = parser.ReadFields()
                    If fields Is Nothing Then
                        Continue While
                    End If
                    Dim row As New Dictionary(Of String, String)(StringComparer.Ordinal)
                    For i As Integer = 0 To headers.Length - 1
                        Dim key As String = headers(i)
                        Dim value As String = ""
                        If i < fields.Length Then
                            value = fields(i)
                        End If
                        row(key) = value
                    Next
                    rows.Add(row)
                End While
            End Using
            Return rows
        End Function

        Private NotInheritable Class UploadArgs
            Public Property Checksum As Boolean
            Public Property Debug As Boolean
            Public Property DeleteAfterUpload As Boolean
            Public Property FileMetadataPath As String
            Public Property Files As List(Of String)
            Public Property Headers As Dictionary(Of String, Object)
            Public Property Identifier As String
            Public Property KeepDirectories As Boolean
            Public Property Metadata As Dictionary(Of String, Object)
            Public Property NoBackup As Boolean
            Public Property NoCollectionCheck As Boolean
            Public Property NoDerive As Boolean
            Public Property OpenAfterUpload As Boolean
            Public Property Quiet As Boolean
            Public Property RemoteName As String
            Public Property Retries As Integer?
            Public Property ShowHelp As Boolean
            Public Property SizeHint As String
            Public Property SleepSeconds As Integer?
            Public Property SpreadsheetPath As String
            Public Property StatusCheck As Boolean
            Public Property Verify As Boolean
        End Class

        Private NotInheritable Class UploadTarget
            Public Property ContentBytes As Byte()
            Public Property FileMetadata As Dictionary(Of String, Object)
            Public Property LocalPath As String
            Public Property RemoteName As String
        End Class
    End Class
End Namespace

Imports System.Globalization
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Friend NotInheritable Class CopyMoveShared
        Private Sub New()
        End Sub

        Friend Shared Function ExecuteCopyLike(session As ArchiveSession, parsed As CopyMoveArgs, cmd As String) As Tuple(Of ApiCallResult, String)
            If String.Equals(parsed.Source, parsed.Destination, StringComparison.Ordinal) Then
                Throw New ArgumentException("The source and destination files cannot be the same!")
            End If

            Dim sourceIdentifier As String = parsed.Source.Split("/"c)(0)
            Dim sourcePathParts = parsed.Source.Split(New Char() {"/"c}, 2)
            Dim sourceFilename As String = If(sourcePathParts.Length > 1, sourcePathParts(1), "")

            Dim sourceItem = session.GetItemMetadata(sourceIdentifier, Nothing)
            If sourceItem.Count = 0 Then
                Throw New ArgumentException(
                    String.Format(
                        "https://{0}/download/{1} does not exist. Please check the identifier and filepath and retry.",
                        session.Host,
                        parsed.Source
                    )
                )
            End If

            Dim sourceFileMetadata As ArchiveFileEntry = FindFileMetadata(sourceItem, sourceFilename)
            If sourceFileMetadata Is Nothing Then
                Throw New ArgumentException(
                    String.Format(
                        "https://{0}/download/{1} does not exist. Please check the identifier and filepath and retry.",
                        session.Host,
                        parsed.Source
                    )
                )
            End If

            Dim headers = ApiShared.ConvertHeaderValues(parsed.Headers)
            headers("x-amz-copy-source") = "/" & ApiShared.EscapeArchivePath(parsed.Source)
            If parsed.Metadata.Count > 0 OrElse parsed.ReplaceMetadata Then
                headers("x-amz-metadata-directive") = "REPLACE"
            Else
                headers("x-amz-metadata-directive") = "COPY"
            End If

            Dim sourceItemMd As Dictionary(Of String, Object) = GetNestedDictionary(sourceItem, "metadata")
            Dim mergedMetadata As Dictionary(Of String, Object)
            If parsed.ReplaceMetadata Then
                mergedMetadata = New Dictionary(Of String, Object)(parsed.Metadata, StringComparer.OrdinalIgnoreCase)
            Else
                mergedMetadata = MergeDictionaries(sourceItemMd, parsed.Metadata)
            End If

            Dim fileMd As Dictionary(Of String, Object) =
                If(parsed.IgnoreFileMetadata, Nothing, sourceFileMetadata.CloneRawFields())

            If Not headers.ContainsKey("x-archive-keep-old-version") AndAlso Not parsed.NoBackup Then
                headers("x-archive-keep-old-version") = "1"
            End If

            Dim queueDerive As Boolean = Not parsed.NoDerive
            Dim copyResp = session.CopyS3Object(
                parsed.Destination,
                headers,
                mergedMetadata,
                fileMd,
                queueDerive
            )
            If copyResp.StatusCode <> 200 Then
                Dim msg As String = ApiShared.GetS3XmlText(copyResp.Text)
                Throw New InvalidOperationException(
                    String.Format(
                        "failed to {0} '{1}' to '{2}' - {3}",
                        cmd,
                        parsed.Source,
                        parsed.Destination,
                        msg
                    )
                )
            End If

            Return Tuple.Create(copyResp, sourceFilename)
        End Function

        Friend Shared Function ParseCommonArgs(args As IList(Of String)) As CopyMoveArgs
            Dim parsed As New CopyMoveArgs With {
                .Metadata = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Headers = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            }

            Dim i As Integer = 0
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-m", "--metadata"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadata(parsed.Metadata, args(i), current)
                    Case "--replace-metadata"
                        parsed.ReplaceMetadata = True
                    Case "-H", "--header"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.Headers, args(i), current)
                    Case "--ignore-file-metadata"
                        parsed.IgnoreFileMetadata = True
                    Case "-n", "--no-derive"
                        parsed.NoDerive = True
                    Case "--no-backup"
                        parsed.NoBackup = True
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            Throw New ArgumentException(String.Format("Unknown option: {0}", current))
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Source) Then
                            parsed.Source = current
                        ElseIf String.IsNullOrWhiteSpace(parsed.Destination) Then
                            parsed.Destination = current
                        Else
                            Throw New ArgumentException("unrecognized arguments")
                        End If
                End Select
                i += 1
            End While

            If String.IsNullOrWhiteSpace(parsed.Source) OrElse String.IsNullOrWhiteSpace(parsed.Destination) Then
                Throw New ArgumentException("the following arguments are required: source destination")
            End If

            parsed.Source = ApiShared.NormalizeArchivePath(parsed.Source)
            parsed.Destination = ApiShared.NormalizeArchivePath(parsed.Destination)

            Return parsed
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function FindFileMetadata(itemMetadata As Dictionary(Of String, Object), filename As String) As ArchiveFileEntry
            For Each entry In ApiShared.GetArchiveFileEntries(itemMetadata)
                If String.Equals(entry.Name, filename, StringComparison.Ordinal) Then
                    Return entry
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function GetNestedDictionary(root As Dictionary(Of String, Object), key As String) As Dictionary(Of String, Object)
            Dim node As Object = Nothing
            If root IsNot Nothing AndAlso root.TryGetValue(key, node) Then
                Dim dict = TryCast(node, Dictionary(Of String, Object))
                If dict IsNot Nothing Then
                    Return New Dictionary(Of String, Object)(dict, StringComparer.OrdinalIgnoreCase)
                End If
            End If
            Return New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        End Function

        Private Shared Function MergeDictionaries(dict0 As Dictionary(Of String, Object), dict1 As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            Dim merged As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If dict0 IsNot Nothing Then
                For Each kvp In dict0
                    merged(kvp.Key) = kvp.Value
                Next
            End If
            If dict1 IsNot Nothing Then
                For Each kvp In dict1
                    merged(kvp.Key) = kvp.Value
                Next
            End If
            Return merged
        End Function

        Private Shared Sub MergeMetadata(destination As Dictionary(Of String, Object), raw As String, optionName As String)
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

        Friend NotInheritable Class CopyMoveArgs
            Public Property Destination As String
            Public Property Headers As Dictionary(Of String, Object)
            Public Property IgnoreFileMetadata As Boolean
            Public Property Metadata As Dictionary(Of String, Object)
            Public Property NoBackup As Boolean
            Public Property NoDerive As Boolean
            Public Property ReplaceMetadata As Boolean
            Public Property Source As String
        End Class
    End Class
End Namespace

Imports System.Globalization
Imports System.IO
Imports System.Web.Script.Serialization
Imports Microsoft.VisualBasic.FileIO
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class MetadataCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed = ParseArguments(args)
            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If
            If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                parsed.Identifier = "None"
            End If

            If Not String.IsNullOrWhiteSpace(parsed.SpreadsheetPath) Then
                Return ExecuteSpreadsheetWrites(session, parsed)
            End If

            Dim metadata = session.GetItemMetadata(parsed.Identifier, parsed.Parameters)

            Dim mutationRequested As Boolean =
                HasValues(parsed.Modify) OrElse
                HasValues(parsed.Remove) OrElse
                HasValues(parsed.AppendValues) OrElse
                HasValues(parsed.AppendList) OrElse
                HasValues(parsed.InsertValues)

            If parsed.ExistsCheck Then
                If metadata.Count > 0 Then
                    Console.Error.WriteLine(String.Format("{0} exists", parsed.Identifier))
                    Return 0
                End If
                Console.Error.WriteLine(String.Format("{0} does not exist", parsed.Identifier))
                Return 1
            End If

            If mutationRequested Then
                If metadata.Count = 0 Then
                    Console.Error.WriteLine(
                        String.Format(
                            "{0} - error: {0} cannot be located because it is dark or does not exist.",
                            parsed.Identifier
                        )
                    )
                    Return 1
                End If

                Dim mode As String = GetWriteMode(parsed)
                Dim inputMetadata As Dictionary(Of String, Object) = GetInputMetadata(parsed)
                Dim sourceMetadata As Dictionary(Of String, Object) = GetTargetMetadata(metadata, parsed.Target)
                Dim destinationMetadata As Dictionary(Of String, Object)

                Try
                    If String.Equals(mode, "remove", StringComparison.OrdinalIgnoreCase) Then
                        Dim simplelistRemovalExitCode As Integer?
                        If TryHandleCollectionSimplelistRemoval(
                            session,
                            parsed.Identifier,
                            parsed.Target,
                            sourceMetadata,
                            inputMetadata,
                            simplelistRemovalExitCode
                        ) Then
                            Return simplelistRemovalExitCode.Value
                        End If

                        Dim removeOutcome = ApplyRemove(sourceMetadata, inputMetadata)
                        If removeOutcome.AllCollectionsRemoved Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "{0} - error: all collections would be removed, not submitting task.",
                                    parsed.Identifier
                                )
                            )
                            Return 1
                        End If
                        If removeOutcome.NothingToRemove Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "{0} - warning: nothing needed to be removed.",
                                    parsed.Identifier
                                )
                            )
                            Return 0
                        End If
                        destinationMetadata = removeOutcome.Destination
                    Else
                        destinationMetadata = ApplyWriteMode(sourceMetadata, inputMetadata, mode)
                    End If
                Catch ex As ArgumentException
                    Console.Error.WriteLine(String.Format("{0} - error: {1}", parsed.Identifier, ex.Message))
                    Return 1
                End Try

                Dim patchOps = BuildPatchOps(sourceMetadata, destinationMetadata, parsed.ExpectValues)
                Dim patchSerializer As New JavaScriptSerializer()
                Dim patchJson As String = patchSerializer.Serialize(patchOps)

                Dim result = session.PostMetadataPatch(
                    parsed.Identifier,
                    parsed.Target,
                    patchJson,
                    ParsePriority(parsed.Priority),
                    ApiShared.ConvertHeaderValues(parsed.HeaderValues),
                    parsed.ReducedPriority,
                    ParseTimeout(parsed.Timeout)
                )

                Dim success As Boolean = False
                If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("success") Then
                    success = Convert.ToBoolean(result.JsonBody("success"), CultureInfo.InvariantCulture)
                End If

                If success Then
                    Dim logText As String = ""
                    If result.JsonBody.ContainsKey("log") Then
                        logText = Convert.ToString(result.JsonBody("log"), CultureInfo.InvariantCulture)
                    End If
                    Console.Error.WriteLine(String.Format("{0} - success: {1}", parsed.Identifier, logText))
                    Return 0
                End If

                Dim errorText As String = ""
                If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("error") Then
                    errorText = Convert.ToString(result.JsonBody("error"), CultureInfo.InvariantCulture)
                End If
                Dim isNoChanges As Boolean = result.Text.IndexOf(
                    "no changes",
                    StringComparison.OrdinalIgnoreCase
                ) >= 0
                Dim typeLabel As String = If(isNoChanges, "warning", "error")
                Console.Error.WriteLine(
                    String.Format(
                        "{0} - {1} ({2}): {3}",
                        parsed.Identifier,
                        typeLabel,
                        result.StatusCode,
                        errorText
                    )
                )
                If result.StatusCode = 200 OrElse isNoChanges Then
                    Return 0
                End If
                Return 1
            End If

            If parsed.FormatsOnly Then
                Dim formats As New HashSet(Of String)(StringComparer.Ordinal)
                For Each fileEntry In ApiShared.GetArchiveFileEntries(metadata)
                    If Not String.IsNullOrWhiteSpace(fileEntry.Format) Then
                        formats.Add(fileEntry.Format)
                    End If
                Next
                For Each formatName In formats
                    Console.WriteLine(formatName)
                Next
                Return 0
            End If

            Dim serializer As New JavaScriptSerializer()
            Console.WriteLine(serializer.Serialize(metadata))
            Return 0
        End Function

        Private Shared Function ApplyRemove(source As Dictionary(Of String, Object), removeValues As Dictionary(Of String, Object)) As RemoveOutcome
            Dim destination As Dictionary(Of String, Object) = DeepCopyDictionary(source)
            Dim changed As Boolean = False
            Dim allCollectionRemoved As Boolean = False

            For Each kvp In removeValues
                Dim key As String = kvp.Key
                Dim removeList As List(Of String) = ToStringList(kvp.Value)
                If removeList.Count = 0 Then
                    Continue For
                End If

                If Not destination.ContainsKey(key) Then
                    Continue For
                End If
                Dim current As Object = destination(key)
                Dim currentList As List(Of String) = ToStringList(current)

                If currentList.Count <= 1 Then
                    If removeList.Contains(ToStringValue(current), StringComparer.Ordinal) Then
                        If String.Equals(key, "collection", StringComparison.OrdinalIgnoreCase) Then
                            allCollectionRemoved = True
                            Exit For
                        End If
                        destination.Remove(key)
                        changed = True
                    End If
                Else
                    Dim nextList As New List(Of String)()
                    For Each value In currentList
                        If Not removeList.Contains(value, StringComparer.Ordinal) Then
                            nextList.Add(value)
                        End If
                    Next

                    If String.Equals(key, "collection", StringComparison.OrdinalIgnoreCase) AndAlso
                       nextList.Count = 0 Then
                        allCollectionRemoved = True
                        Exit For
                    End If

                    If nextList.Count <> currentList.Count Then
                        changed = True
                        destination(key) = nextList
                    End If
                End If
            Next

            Return New RemoveOutcome With {
                .Destination = destination,
                .NothingToRemove = Not changed,
                .AllCollectionsRemoved = allCollectionRemoved
            }
        End Function

        Private Shared Function TryHandleCollectionSimplelistRemoval(
            session As ArchiveSession,
            identifier As String,
            target As String,
            sourceMetadata As Dictionary(Of String, Object),
            inputMetadata As Dictionary(Of String, Object),
            ByRef exitCode As Integer?
        ) As Boolean
            exitCode = Nothing

            If target IsNot Nothing AndAlso target.StartsWith("files/", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            If inputMetadata Is Nothing OrElse Not inputMetadata.ContainsKey("collection") Then
                Return False
            End If

            If sourceMetadata IsNot Nothing AndAlso sourceMetadata.ContainsKey("collection") Then
                Return False
            End If

            Dim collections As List(Of String) = ToStringList(inputMetadata("collection"))
            If collections Is Nothing OrElse collections.Count = 0 Then
                Return False
            End If

            For Each collectionName In collections
                If String.IsNullOrWhiteSpace(collectionName) Then
                    Continue For
                End If

                Dim patch As New Dictionary(Of String, Object) From {
                    {"op", "delete"},
                    {"parent", collectionName},
                    {"list", "holdings"}
                }
                Dim response = session.PostSimplelistPatch(identifier, patch)
                Dim responseError As String = ""
                If response.JsonBody IsNot Nothing AndAlso response.JsonBody.ContainsKey("error") Then
                    responseError = Convert.ToString(response.JsonBody("error"), CultureInfo.InvariantCulture)
                End If

                Dim success As Boolean = response.JsonBody IsNot Nothing AndAlso
                    response.JsonBody.ContainsKey("success") AndAlso
                    Convert.ToBoolean(response.JsonBody("success"), CultureInfo.InvariantCulture)
                If success OrElse responseError.StartsWith("no row to delete for", StringComparison.OrdinalIgnoreCase) Then
                    Console.Error.WriteLine(
                        String.Format(
                            "{0} - success: {0} no longer in {1}",
                            identifier,
                            collectionName
                        )
                    )
                    exitCode = 0
                    Return True
                End If

                Console.Error.WriteLine(String.Format("{0} - error: {1}", identifier, responseError))
                exitCode = 1
                Return True
            Next

            Return False
        End Function

        Private Shared Function ApplyWriteMode(
            source As Dictionary(Of String, Object),
            updates As Dictionary(Of String, Object),
            mode As String
        ) As Dictionary(Of String, Object)
            Dim destination As Dictionary(Of String, Object) = DeepCopyDictionary(source)

            If String.Equals(mode, "insert", StringComparison.OrdinalIgnoreCase) Then
                For Each kvp In updates
                    If IsIndexedKey(kvp.Key) Then
                        Dim baseKey As String = GetBaseKey(kvp.Key)
                        Dim idx As Integer = GetIndex(kvp.Key)
                        Dim existing = ToStringList(If(destination.ContainsKey(baseKey), destination(baseKey), Nothing))
                        While existing.Count < idx
                            existing.Add(String.Empty)
                        End While
                        If existing.Contains(ToStringValue(kvp.Value)) Then
                            existing.Remove(ToStringValue(kvp.Value))
                        End If
                        If idx >= existing.Count Then
                            existing.Add(ToStringValue(kvp.Value))
                        Else
                            existing.Insert(idx, ToStringValue(kvp.Value))
                        End If
                        destination(baseKey) = existing
                    Else
                        destination(kvp.Key) = kvp.Value
                    End If
                Next
                Return destination
            End If

            For Each kvp In updates
                Dim key As String = kvp.Key
                Dim value As Object = kvp.Value
                If String.Equals(mode, "append", StringComparison.OrdinalIgnoreCase) Then
                    If destination.ContainsKey(key) Then
                        If TypeOf destination(key) Is ArrayList OrElse
                           TypeOf destination(key) Is Object() OrElse
                           TypeOf destination(key) Is List(Of String) Then
                            Throw New ArgumentException(
                                "cannot append string to list metadata with '--append'; use '--append-list' instead."
                            )
                        End If
                        destination(key) = ToStringValue(destination(key)) & " " & ToStringValue(value)
                    Else
                        destination(key) = value
                    End If
                ElseIf String.Equals(mode, "append-list", StringComparison.OrdinalIgnoreCase) Then
                    Dim list As List(Of String) = ToStringList(If(destination.ContainsKey(key), destination(key), Nothing))
                    list.AddRange(ToStringList(value))
                    destination(key) = list
                Else
                    destination(key) = value
                End If
            Next

            Return destination
        End Function

        Private Shared Function BuildExpectPath(key As String) As String
            If key.Contains("[") AndAlso key.Contains("]") Then
                Dim baseKey As String = GetBaseKey(key)
                Dim idx As Integer = GetIndex(key)
                Return "/" & EscapeJsonPointer(baseKey) & "/" & idx.ToString(CultureInfo.InvariantCulture)
            End If
            Return "/" & EscapeJsonPointer(key)
        End Function

        Private Shared Function BuildPatchOps(
            source As Dictionary(Of String, Object),
            destination As Dictionary(Of String, Object),
            expectValues As Dictionary(Of String, Object)
        ) As List(Of Dictionary(Of String, Object))
            Dim ops As New List(Of Dictionary(Of String, Object))()

            For Each kvp In expectValues
                ops.Add(New Dictionary(Of String, Object) From {
                    {"op", "test"},
                    {"path", BuildExpectPath(kvp.Key)},
                    {"value", kvp.Value}
                })
            Next

            Dim keys As New HashSet(Of String)(source.Keys, StringComparer.OrdinalIgnoreCase)
            For Each key In destination.Keys
                keys.Add(key)
            Next

            For Each key In keys
                Dim inSource As Boolean = source.ContainsKey(key)
                Dim inDestination As Boolean = destination.ContainsKey(key)

                If inSource AndAlso Not inDestination Then
                    ops.Add(New Dictionary(Of String, Object) From {
                        {"op", "remove"},
                        {"path", "/" & EscapeJsonPointer(key)}
                    })
                ElseIf Not inSource AndAlso inDestination Then
                    ops.Add(New Dictionary(Of String, Object) From {
                        {"op", "add"},
                        {"path", "/" & EscapeJsonPointer(key)},
                        {"value", destination(key)}
                    })
                ElseIf inSource AndAlso inDestination AndAlso
                    Not ObjectDeepEquals(source(key), destination(key)) Then
                    ops.Add(New Dictionary(Of String, Object) From {
                        {"op", "replace"},
                        {"path", "/" & EscapeJsonPointer(key)},
                        {"value", destination(key)}
                    })
                End If
            Next

            Return ops
        End Function

        Private Shared Function DeepCopyDictionary(
            source As Dictionary(Of String, Object)
        ) As Dictionary(Of String, Object)
            Dim serializer As New JavaScriptSerializer()
            Dim json As String = serializer.Serialize(source)
            Dim copy = serializer.Deserialize(Of Dictionary(Of String, Object))(json)
            Return New Dictionary(Of String, Object)(copy, StringComparer.OrdinalIgnoreCase)
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function EscapeJsonPointer(segment As String) As String
            Return segment.Replace("~", "~0").Replace("/", "~1")
        End Function

        Private Shared Function ExecuteSpreadsheetWrites(
            session As ArchiveSession,
            parsed As MetadataArgs
        ) As Integer
            If String.IsNullOrWhiteSpace(parsed.Priority) Then
                parsed.Priority = "-5"
            End If

            Dim hasFailures As Boolean = False
            Using parser As New TextFieldParser(parsed.SpreadsheetPath, System.Text.Encoding.UTF8)
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                If parser.EndOfData Then
                    Return 0
                End If

                Dim headers = parser.ReadFields()
                If headers Is Nothing Then
                    Return 0
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

                    Dim identifier As String = ""
                    If row.ContainsKey("identifier") Then
                        identifier = row("identifier")
                    End If
                    If String.IsNullOrWhiteSpace(identifier) Then
                        Continue While
                    End If

                    If row.ContainsKey("file") Then
                        row.Remove("file")
                    End If

                    Dim updates As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    For Each kvp In row
                        If Not String.IsNullOrWhiteSpace(kvp.Value) Then
                            updates(kvp.Key.ToLowerInvariant()) = kvp.Value
                        End If
                    Next

                    Dim itemMetadata = session.GetItemMetadata(identifier, Nothing)
                    If itemMetadata.Count = 0 Then
                        Console.Error.WriteLine(
                            String.Format(
                                "{0} - error: {0} cannot be located because it is dark or does not exist.",
                                identifier
                            )
                        )
                        hasFailures = True
                        Continue While
                    End If
                    Dim sourceMetadata As Dictionary(Of String, Object) =
                        GetTargetMetadata(itemMetadata, parsed.Target)
                    Dim destinationMetadata As Dictionary(Of String, Object) =
                        ApplyWriteMode(sourceMetadata, updates, "modify")

                    Dim patchOps = BuildPatchOps(sourceMetadata, destinationMetadata, parsed.ExpectValues)
                    Dim serializer As New JavaScriptSerializer()
                    Dim patchJson As String = serializer.Serialize(patchOps)

                    Dim result = session.PostMetadataPatch(
                        identifier,
                        parsed.Target,
                        patchJson,
                        ParsePriority(parsed.Priority),
                        ApiShared.ConvertHeaderValues(parsed.HeaderValues),
                        parsed.ReducedPriority,
                        ParseTimeout(parsed.Timeout)
                    )

                    Dim success As Boolean = False
                    If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("success") Then
                        success = Convert.ToBoolean(result.JsonBody("success"), CultureInfo.InvariantCulture)
                    End If

                    If success Then
                        Dim logText As String = ""
                        If result.JsonBody.ContainsKey("log") Then
                            logText = Convert.ToString(result.JsonBody("log"), CultureInfo.InvariantCulture)
                        End If
                        Console.Error.WriteLine(String.Format("{0} - success: {1}", identifier, logText))
                        Continue While
                    End If

                    Dim errorText As String = ""
                    If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("error") Then
                        errorText = Convert.ToString(result.JsonBody("error"), CultureInfo.InvariantCulture)
                    End If
                    Dim isNoChanges As Boolean = result.Text.IndexOf(
                        "no changes",
                        StringComparison.OrdinalIgnoreCase
                    ) >= 0
                    Dim typeLabel As String = If(isNoChanges, "warning", "error")
                    Console.Error.WriteLine(
                        String.Format(
                            "{0} - {1} ({2}): {3}",
                            identifier,
                            typeLabel,
                            result.StatusCode,
                            errorText
                        )
                    )
                    If Not (result.StatusCode = 200 OrElse isNoChanges) Then
                        hasFailures = True
                    End If
                End While
            End Using

            Return If(hasFailures, 1, 0)
        End Function

        Private Shared Function GetBaseKey(indexed As String) As String
            Dim idx As Integer = indexed.IndexOf("["c)
            If idx < 0 Then
                Return indexed
            End If
            Return indexed.Substring(0, idx)
        End Function

        Private Shared Function GetIndex(indexed As String) As Integer
            Dim startIdx As Integer = indexed.IndexOf("["c)
            Dim endIdx As Integer = indexed.IndexOf("]"c)
            If startIdx < 0 OrElse endIdx <= startIdx + 1 Then
                Return 0
            End If
            Dim n As String = indexed.Substring(startIdx + 1, endIdx - startIdx - 1)
            Return Integer.Parse(n, CultureInfo.InvariantCulture)
        End Function

        Private Shared Function GetInputMetadata(parsed As MetadataArgs) As Dictionary(Of String, Object)
            If HasValues(parsed.Remove) Then Return parsed.Remove
            If HasValues(parsed.AppendValues) Then Return parsed.AppendValues
            If HasValues(parsed.AppendList) Then Return parsed.AppendList
            If HasValues(parsed.InsertValues) Then Return parsed.InsertValues
            Return parsed.Modify
        End Function

        Private Shared Function GetTargetMetadata(
            itemMetadata As Dictionary(Of String, Object),
            target As String
        ) As Dictionary(Of String, Object)
            Dim actualTarget As String = ApiShared.NormalizeArchivePath(If(String.IsNullOrWhiteSpace(target), "metadata", target))
            If String.Equals(actualTarget, "metadata", StringComparison.OrdinalIgnoreCase) Then
                Dim mdNode As Object = Nothing
                If itemMetadata.TryGetValue("metadata", mdNode) Then
                    Dim mdDict = TryCast(mdNode, Dictionary(Of String, Object))
                    If mdDict IsNot Nothing Then
                        Return DeepCopyDictionary(mdDict)
                    End If
                End If
                Return New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            End If

            If actualTarget.StartsWith("files/", StringComparison.OrdinalIgnoreCase) Then
                Dim parts = actualTarget.Split("/"c)
                Dim filename As String = If(parts.Length > 1, parts(1), "")
                For Each fileEntry In ApiShared.GetArchiveFileEntries(itemMetadata)
                    If String.Equals(fileEntry.Name, filename, StringComparison.Ordinal) Then
                        Return DeepCopyDictionary(fileEntry.CloneRawFields())
                    End If
                Next
                Return New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            End If

            ' Fallback for other targets: use top-level metadata object.
            Return DeepCopyDictionary(itemMetadata)
        End Function

        Private Shared Function GetWriteMode(parsed As MetadataArgs) As String
            If HasValues(parsed.Remove) Then Return "remove"
            If HasValues(parsed.AppendValues) Then Return "append"
            If HasValues(parsed.AppendList) Then Return "append-list"
            If HasValues(parsed.InsertValues) Then Return "insert"
            Return "modify"
        End Function

        Private Shared Function HasValues(values As Dictionary(Of String, Object)) As Boolean
            Return values IsNot Nothing AndAlso values.Count > 0
        End Function

        Private Shared Function IsIndexedKey(key As String) As Boolean
            Return key.Contains("[") AndAlso key.Contains("]")
        End Function

        Private Shared Sub MergeMetadataOption(
            destination As Dictionary(Of String, Object),
            raw As String,
            optionName As String
        )
            Dim normalized As String = raw
            If normalized.IndexOf(":"c) < 0 AndAlso normalized.IndexOf("="c) >= 0 Then
                Dim firstEq As Integer = normalized.IndexOf("="c)
                normalized = normalized.Substring(0, firstEq) & ":" & normalized.Substring(firstEq + 1)
            End If
            Dim separator As Integer = normalized.IndexOf(":"c)
            If separator <= 0 OrElse separator = normalized.Length - 1 Then
                Throw New ArgumentException(
                    String.Format("{0} must be formatted as 'KEY:VALUE'", optionName)
                )
            End If
            Dim key As String = normalized.Substring(0, separator)
            Dim value As String = normalized.Substring(separator + 1)
            If destination.ContainsKey(key) Then
                Dim existingList = If(TryCast(destination(key), List(Of String)), New List(Of String) From {Convert.ToString(destination(key))})
                existingList.Add(value)
                destination(key) = existingList
            Else
                destination(key) = value
            End If
        End Sub

        Private Shared Sub MergeQueryStringOption(
            destination As Dictionary(Of String, Object),
            raw As String,
            optionName As String
        )
            Dim normalized As String = raw
            If normalized.IndexOf("="c) < 0 AndAlso normalized.IndexOf(":"c) >= 0 Then
                Dim firstColon As Integer = normalized.IndexOf(":"c)
                normalized = normalized.Substring(0, firstColon) & "=" & normalized.Substring(firstColon + 1)
            End If

            Dim pairs = ParseQueryString(normalized)
            If normalized.Length > 0 AndAlso pairs.Count = 0 Then
                Throw New ArgumentException(
                    String.Format("{0} must be formatted as 'key=value' or 'key:value'", optionName)
                )
            End If

            For Each pair In pairs
                If destination.ContainsKey(pair.Key) Then
                    Dim existing = destination(pair.Key)
                    Dim list = If(TryCast(existing, List(Of String)), New List(Of String) From {Convert.ToString(existing)})
                    list.Add(pair.Value)
                    destination(pair.Key) = list
                Else
                    destination(pair.Key) = pair.Value
                End If
            Next
        End Sub

        Private Shared Function ObjectDeepEquals(left As Object, right As Object) As Boolean
            Dim serializer As New JavaScriptSerializer()
            Return String.Equals(
                serializer.Serialize(left),
                serializer.Serialize(right),
                StringComparison.Ordinal
            )
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As MetadataArgs
            Dim parsed As New MetadataArgs With {
                .Parameters = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Modify = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Remove = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .AppendValues = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .AppendList = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .InsertValues = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .ExpectValues = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .HeaderValues = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            }

            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-m", "--modify"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.Modify, args(i), current)
                    Case "-r", "--remove"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.Remove, args(i), current)
                    Case "-a", "--append"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.AppendValues, args(i), current)
                    Case "-A", "--append-list"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.AppendList, args(i), current)
                    Case "-i", "--insert"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.InsertValues, args(i), current)
                    Case "-E", "--expect"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeMetadataOption(parsed.ExpectValues, args(i), current)
                    Case "-H", "--header"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryStringOption(parsed.HeaderValues, args(i), current)
                    Case "-t", "--target"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Target = args(i)
                    Case "-s", "--spreadsheet"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.SpreadsheetPath = Path.GetFullPath(args(i))
                    Case "-e", "--exists"
                        parsed.ExistsCheck = True
                    Case "-F", "--formats"
                        parsed.FormatsOnly = True
                    Case "-p", "--priority"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Priority = args(i)
                    Case "--timeout"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Timeout = args(i)
                    Case "-R", "--reduced-priority"
                        parsed.ReducedPriority = True
                    Case "-P", "--parameters"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryStringOption(parsed.Parameters, args(i), current)
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            unknown.Add(current)
                            i += 1
                            Continue While
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                            parsed.Identifier = current
                        Else
                            Throw New ArgumentException("unrecognized arguments")
                        End If
                End Select
                i += 1
            End While

            If unknown.Count > 0 Then
                Throw New ArgumentException("unrecognized arguments: " & String.Join(" ", unknown))
            End If
            Return parsed
        End Function

        Private Shared Function ParsePriority(priorityRaw As String) As Integer
            If String.IsNullOrWhiteSpace(priorityRaw) Then
                Return -5
            End If
            Return Integer.Parse(priorityRaw, CultureInfo.InvariantCulture)
        End Function

        Private Shared Function ParseQueryString(raw As String) As List(Of KeyValuePair(Of String, String))
            Dim result As New List(Of KeyValuePair(Of String, String))()
            Dim segments = raw.Split(New Char() {"&"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each segment In segments
                Dim eqIndex As Integer = segment.IndexOf("="c)
                If eqIndex <= 0 OrElse eqIndex = segment.Length - 1 Then
                    Continue For
                End If
                Dim key As String = Uri.UnescapeDataString(segment.Substring(0, eqIndex))
                Dim value As String = Uri.UnescapeDataString(segment.Substring(eqIndex + 1))
                result.Add(New KeyValuePair(Of String, String)(key, value))
            Next
            Return result
        End Function

        Private Shared Function ParseTimeout(timeoutRaw As String) As Double
            If String.IsNullOrWhiteSpace(timeoutRaw) Then
                Return 60
            End If
            Return Double.Parse(timeoutRaw, CultureInfo.InvariantCulture)
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " metadata [-h] [-m KEY:VALUE | -r KEY:VALUE | -a KEY:VALUE | -A KEY:VALUE | -i KEY:VALUE]")
            Console.Error.WriteLine("                      [-E KEY:VALUE] [-H KEY:VALUE] [-t target]")
            Console.Error.WriteLine("                      [-s metadata.csv] [-e] [-F] [-p priority]")
            Console.Error.WriteLine("                      [--timeout value] [-R] [-P KEY:VALUE]")
            Console.Error.WriteLine("                      [identifier]")
        End Sub

        Private Shared Function ToStringList(value As Object) As List(Of String)
            Dim result As New List(Of String)()
            If value Is Nothing Then
                Return result
            End If
            Dim arrList = TryCast(value, ArrayList)
            If arrList IsNot Nothing Then
                For Each entry In arrList
                    result.Add(ToStringValue(entry))
                Next
                Return result
            End If
            Dim objArray = TryCast(value, Object())
            If objArray IsNot Nothing Then
                For Each entry In objArray
                    result.Add(ToStringValue(entry))
                Next
                Return result
            End If
            Dim strList = TryCast(value, List(Of String))
            If strList IsNot Nothing Then
                result.AddRange(strList)
                Return result
            End If
            result.Add(ToStringValue(value))
            Return result
        End Function

        Private Shared Function ToStringValue(value As Object) As String
            Return If(value Is Nothing, String.Empty, Convert.ToString(value, CultureInfo.InvariantCulture))
        End Function

        Private NotInheritable Class MetadataArgs
            Public Property AppendList As Dictionary(Of String, Object)
            Public Property AppendValues As Dictionary(Of String, Object)
            Public Property ExistsCheck As Boolean
            Public Property ExpectValues As Dictionary(Of String, Object)
            Public Property FormatsOnly As Boolean
            Public Property HeaderValues As Dictionary(Of String, Object)
            Public Property Identifier As String
            Public Property InsertValues As Dictionary(Of String, Object)
            Public Property Modify As Dictionary(Of String, Object)
            Public Property Parameters As Dictionary(Of String, Object)
            Public Property Priority As String
            Public Property ReducedPriority As Boolean
            Public Property Remove As Dictionary(Of String, Object)
            Public Property ShowHelp As Boolean
            Public Property SpreadsheetPath As String
            Public Property Target As String = "metadata"
            Public Property Timeout As String
        End Class
        Private NotInheritable Class RemoveOutcome
            Public Property AllCollectionsRemoved As Boolean
            Public Property Destination As Dictionary(Of String, Object)
            Public Property NothingToRemove As Boolean
        End Class
    End Class
End Namespace

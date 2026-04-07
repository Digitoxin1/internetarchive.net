Imports System.Text.RegularExpressions
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class ListCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As ListArgs
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
            Dim metadata = session.GetItemMetadata(parsed.Identifier, Nothing)
            Dim files = ApiShared.GetArchiveFileEntries(metadata)

            SetupColumns(parsed, files)
            files = FilterFiles(parsed, files)

            Dim outputRows As New List(Of Dictionary(Of String, String))()
            For Each fileEntry In files
                Dim row As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each kvp In fileEntry.RawFields
                    If parsed.Columns.Contains(kvp.Key) Then
                        Dim valueString As String
                        If kvp.Key = "name" AndAlso parsed.Location Then
                            Dim remotePath As String = ApiShared.EscapeArchivePath(parsed.Identifier & "/" & fileEntry.Name)
                            valueString = String.Format(
                                "https://{0}/download/{1}",
                                session.Host,
                                remotePath
                            )
                        Else
                            valueString = AsJoinedString(kvp.Value)
                        End If
                        row(kvp.Key) = valueString
                    End If
                Next
                outputRows.Add(row)
            Next

            Dim allEmpty As Boolean = True
            For Each row In outputRows
                If row.Count > 0 Then
                    allEmpty = False
                    Exit For
                End If
            Next
            If allEmpty Then
                Return 1
            End If

            If parsed.Verbose Then
                Console.Out.WriteLine(String.Join(vbTab, parsed.Columns))
            End If
            For Each row In outputRows
                Dim values As New List(Of String)()
                For Each col In parsed.Columns
                    If row.ContainsKey(col) Then
                        values.Add(row(col))
                    Else
                        values.Add(String.Empty)
                    End If
                Next
                Console.Out.WriteLine(String.Join(vbTab, values))
            Next
            Return 0
        End Function

        Private Shared Function AsJoinedString(value As Object) As String
            If value Is Nothing Then
                Return String.Empty
            End If

            Dim arrList = TryCast(value, ArrayList)
            If arrList IsNot Nothing Then
                Dim parts As New List(Of String)()
                For Each entry In arrList
                    parts.Add(AsString(entry))
                Next
                Return String.Join(";", parts)
            End If

            Dim objArray = TryCast(value, Object())
            If objArray IsNot Nothing Then
                Dim parts As New List(Of String)()
                For Each entry In objArray
                    parts.Add(AsString(entry))
                Next
                Return String.Join(";", parts)
            End If

            Return AsString(value)
        End Function

        Private Shared Function AsString(value As Object) As String
            Return If(value Is Nothing, String.Empty, Convert.ToString(value))
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function FilterFiles(parsed As ListArgs, files As List(Of ArchiveFileEntry)) As List(Of ArchiveFileEntry)
            If Not String.IsNullOrWhiteSpace(parsed.GlobPattern) Then
                Dim patterns = parsed.GlobPattern.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim filtered As New List(Of ArchiveFileEntry)()
                For Each fileEntry In files
                    For Each pattern In patterns
                        If GlobMatch(fileEntry.Name, pattern) Then
                            filtered.Add(fileEntry)
                            Exit For
                        End If
                    Next
                Next
                Return filtered
            End If

            If parsed.Formats.Count > 0 Then
                Dim filtered As New List(Of ArchiveFileEntry)()
                For Each fileEntry In files
                    For Each requestedFormat In parsed.Formats
                        If String.Equals(
                            fileEntry.Format,
                            requestedFormat,
                            StringComparison.OrdinalIgnoreCase
                        ) Then
                            filtered.Add(fileEntry)
                            Exit For
                        End If
                    Next
                Next
                Return filtered
            End If

            Return files
        End Function

        Private Shared Function GlobMatch(value As String, pattern As String) As Boolean
            Dim escaped = Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".")
            Return Regex.IsMatch(value, "^" & escaped & "$", RegexOptions.IgnoreCase)
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As ListArgs
            Dim parsed As New ListArgs With {
                .Columns = New List(Of String)(),
                .Formats = New List(Of String)()
            }
            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-v", "--verbose"
                        parsed.Verbose = True
                    Case "-a", "--all"
                        parsed.AllFields = True
                    Case "-l", "--location"
                        parsed.Location = True
                    Case "-c", "--columns"
                        i += 1
                        EnsureHasValue(args, i, current)
                        For Each value In SplitComma(args(i))
                            parsed.Columns.Add(value)
                        Next
                    Case "-g", "--glob"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.GlobPattern = args(i)
                    Case "-f", "--format"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Formats.Add(args(i))
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

            If String.IsNullOrWhiteSpace(parsed.Identifier) AndAlso Not parsed.ShowHelp Then
                Throw New ArgumentException("the following arguments are required: identifier")
            End If
            If unknown.Count > 0 Then
                Throw New ArgumentException("unrecognized arguments: " & String.Join(" ", unknown))
            End If
            Return parsed
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " list [-h] [-v] [-a] [-l] [-c COLUMNS] [-g PATTERN] [-f FORMAT]")
            Console.Error.WriteLine("                  identifier")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier of the item")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -v, --verbose")
            Console.Error.WriteLine("  -a, --all")
            Console.Error.WriteLine("  -l, --location")
            Console.Error.WriteLine("  -c COLUMNS, --columns COLUMNS")
            Console.Error.WriteLine("  -g PATTERN, --glob PATTERN")
            Console.Error.WriteLine("  -f FORMAT, --format FORMAT")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " list [-h] [-v] [-a] [-l] [-c COLUMNS] [-g PATTERN] [-f FORMAT]")
            Console.Error.WriteLine("                  identifier")
            Console.Error.WriteLine(CliApp.ExecutableName() & " list: error: " & message)
        End Sub

        Private Shared Sub SetupColumns(parsed As ListArgs, files As List(Of ArchiveFileEntry))
            If parsed.Columns.Count = 0 Then
                parsed.Columns.Add("name")
            End If

            If parsed.AllFields Then
                Dim colSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each fileEntry In files
                    For Each key In fileEntry.RawFields.Keys
                        colSet.Add(key)
                    Next
                Next
                parsed.Columns = New List(Of String)(colSet)
            End If

            If parsed.Columns.Contains("name") Then
                parsed.Columns.Remove("name")
                parsed.Columns.Insert(0, "name")
            End If
        End Sub

        Private Shared Function SplitComma(value As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim parts = value.Split(New Char() {","c}, StringSplitOptions.None)
            For Each part In parts
                result.Add(part)
            Next
            Return result
        End Function

        Private NotInheritable Class ListArgs
            Public Property AllFields As Boolean
            Public Property Columns As List(Of String)
            Public Property Formats As List(Of String)
            Public Property GlobPattern As String
            Public Property Identifier As String
            Public Property Location As Boolean
            Public Property ShowHelp As Boolean
            Public Property Verbose As Boolean
        End Class
    End Class
End Namespace

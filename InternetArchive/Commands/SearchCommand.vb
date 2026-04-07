Imports System.Globalization
Imports System.Net.Http
Imports System.Text
Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Core
Imports InternetArchive.InternetArchiveCli.Exceptions

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class SearchCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As SearchArgs
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
            Dim serializer As New JavaScriptSerializer()

            Try
                Dim preparedFields As List(Of String) = PrepareValues(parsed.Fields)
                Dim preparedSorts As List(Of String) = PrepareValues(parsed.Sorts)

                If parsed.NumFound Then
                    Dim count As Integer = GetNumFound(
                        session,
                        parsed.Query,
                        parsed.Parameters,
                        parsed.Headers,
                        parsed.TimeoutSeconds,
                        parsed.Fts,
                        parsed.DslFts
                    )
                    Console.WriteLine(count.ToString(CultureInfo.InvariantCulture))
                    Return 0
                End If

                Dim results As IEnumerable(Of Dictionary(Of String, Object)) = ExecuteSearch(
                    session,
                    parsed.Query,
                    parsed.Parameters,
                    parsed.Headers,
                    preparedFields,
                    preparedSorts,
                    parsed.TimeoutSeconds,
                    parsed.Fts,
                    parsed.DslFts
                )

                For Each result In results
                    If parsed.ItemList Then
                        If parsed.Fts OrElse parsed.DslFts Then
                            Dim fieldsNode As Object = Nothing
                            If result.TryGetValue("fields", fieldsNode) Then
                                Dim fieldDict = TryCast(fieldsNode, Dictionary(Of String, Object))
                                If fieldDict IsNot Nothing AndAlso fieldDict.ContainsKey("identifier") Then
                                    Dim identifiers = AsStringList(fieldDict("identifier"))
                                    For Each id In identifiers
                                        Console.WriteLine(id)
                                    Next
                                End If
                            End If
                        Else
                            Dim identifier As Object = Nothing
                            If result.TryGetValue("identifier", identifier) Then
                                Console.WriteLine(Convert.ToString(identifier))
                            Else
                                Console.WriteLine(String.Empty)
                            End If
                        End If
                    Else
                        Console.WriteLine(serializer.Serialize(result))
                        If result.ContainsKey("error") Then
                            Return 1
                        End If
                    End If
                Next
                Return 0
            Catch ex As ValueErrorException
                Console.Error.WriteLine("error: " & ex.Message)
                Return 1
            Catch ex As AuthenticationError
                Console.Error.WriteLine("error: " & ex.Message)
                Return 1
            Catch ex As TaskCanceledException
                Console.Error.WriteLine(
                    "error: Request timed out. Increase the --timeout and try again."
                )
                Return 1
            Catch ex As TimeoutException
                Console.Error.WriteLine(
                    "error: The server timed out and failed to return all search results, please try again"
                )
                Return 1
            End Try
        End Function

        Private Shared Function AdvancedSearch(
            session As ArchiveSession,
            params As Dictionary(Of String, Object),
            headers As Dictionary(Of String, Object),
            fields As List(Of String),
            sorts As List(Of String),
            timeoutSeconds As Double
        ) As IEnumerable(Of Dictionary(Of String, Object))
            Dim p As New Dictionary(Of String, Object)(params, StringComparer.OrdinalIgnoreCase)
            Dim fieldList As New List(Of String)()
            If fields IsNot Nothing Then
                fieldList.AddRange(fields)
            End If
            If Not fieldList.Contains("identifier") Then
                fieldList.Add("identifier")
            End If

            For i As Integer = 0 To fieldList.Count - 1
                p(String.Format("fl[{0}]", i)) = fieldList(i)
            Next
            If sorts IsNot Nothing Then
                For i As Integer = 0 To sorts.Count - 1
                    p(String.Format("sort[{0}]", i)) = sorts(i)
                Next
            End If
            p("output") = "json"

            Dim url As String = session.Protocol & "//" & session.Host & "/advancedsearch.php"
            Dim json = SendRequestAndDeserialize(
                session,
                HttpMethod.Get,
                url,
                p,
                Nothing,
                headers,
                timeoutSeconds
            )

            Dim results As New List(Of Dictionary(Of String, Object))()
            If json.ContainsKey("error") Then
                results.Add(json)
            End If

            Dim responseObj As Object = Nothing
            If json.TryGetValue("response", responseObj) Then
                Dim responseDict = TryCast(responseObj, Dictionary(Of String, Object))
                If responseDict IsNot Nothing AndAlso responseDict.ContainsKey("docs") Then
                    For Each docObj In ApiShared.ToObjectSequence(responseDict("docs"))
                        Dim doc = TryCast(docObj, Dictionary(Of String, Object))
                        If doc IsNot Nothing Then
                            results.Add(doc)
                        End If
                    Next
                End If
            End If

            Return results
        End Function

        Private Shared Sub ApplyHeaders(
            request As HttpRequestMessage,
            headers As Dictionary(Of String, Object)
        )
            For Each kvp In headers
                Dim values As List(Of String) = AsStringList(kvp.Value)
                For Each value In values
                    request.Headers.TryAddWithoutValidation(kvp.Key, value)
                Next
            Next
        End Sub

        Private Shared Function AsStringList(value As Object) As List(Of String)
            Dim result As New List(Of String)()
            If value Is Nothing Then
                Return result
            End If
            Dim list = TryCast(value, List(Of String))
            If list IsNot Nothing Then
                result.AddRange(list)
                Return result
            End If
            Dim arr = TryCast(value, ArrayList)
            If arr IsNot Nothing Then
                For Each v In arr
                    result.Add(Convert.ToString(v))
                Next
                Return result
            End If
            Dim objArray = TryCast(value, Object())
            If objArray IsNot Nothing Then
                For Each v In objArray
                    result.Add(Convert.ToString(v))
                Next
                Return result
            End If
            result.Add(Convert.ToString(value))
            Return result
        End Function

        Private Shared Function BuildDefaultSearchParams(
            query As String,
            original As Dictionary(Of String, Object)
        ) As Dictionary(Of String, Object)
            Dim merged As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                {"q", query}
            }

            For Each kvp In original
                merged(kvp.Key) = kvp.Value
            Next

            If Not merged.ContainsKey("page") Then
                If merged.ContainsKey("rows") Then
                    merged("page") = "1"
                Else
                    merged("count") = "10000"
                End If
            Else
                merged("output") = "json"
            End If

            If merged.ContainsKey("index") Then
                merged("scope") = merged("index")
                merged.Remove("index")
            End If

            Return merged
        End Function

        Private Shared Function BuildUrl(
            baseUrl As String,
            params As Dictionary(Of String, Object)
        ) As String
            If params Is Nothing OrElse params.Count = 0 Then
                Return baseUrl
            End If
            Dim sb As New StringBuilder()
            sb.Append(baseUrl)
            sb.Append("?")
            Dim first As Boolean = True
            For Each kvp In params
                Dim values = AsStringList(kvp.Value)
                For Each value In values
                    If Not first Then
                        sb.Append("&")
                    End If
                    sb.Append(Uri.EscapeDataString(kvp.Key))
                    sb.Append("=")
                    sb.Append(Uri.EscapeDataString(value))
                    first = False
                Next
            Next
            Return sb.ToString()
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function ExecuteSearch(
            session As ArchiveSession,
            query As String,
            params As Dictionary(Of String, Object),
            headers As Dictionary(Of String, Object),
            fields As List(Of String),
            sorts As List(Of String),
            timeoutSeconds As Double,
            fts As Boolean,
            dslFts As Boolean
        ) As IEnumerable(Of Dictionary(Of String, Object))
            If fts OrElse dslFts Then
                Return FullTextSearch(session, query, params, headers, timeoutSeconds, dslFts)
            End If

            Dim queryText As String = query
            Dim searchParams = BuildDefaultSearchParams(queryText, params)
            If searchParams.ContainsKey("page") Then
                Return AdvancedSearch(session, searchParams, headers, fields, sorts, timeoutSeconds)
            End If
            Return ScrapeSearch(session, searchParams, headers, fields, sorts, timeoutSeconds)
        End Function

        Private Shared Function ExtractAdvancedNumFound(
            payload As Dictionary(Of String, Object)
        ) As Integer
            Dim responseObj As Object = Nothing
            If payload.TryGetValue("response", responseObj) Then
                Dim response = TryCast(responseObj, Dictionary(Of String, Object))
                If response IsNot Nothing AndAlso response.ContainsKey("numFound") Then
                    Return Convert.ToInt32(response("numFound"), CultureInfo.InvariantCulture)
                End If
            End If
            Return 0
        End Function

        Private Shared Function FullTextSearch(
            session As ArchiveSession,
            query As String,
            params As Dictionary(Of String, Object),
            headers As Dictionary(Of String, Object),
            timeoutSeconds As Double,
            dslFts As Boolean
        ) As IEnumerable(Of Dictionary(Of String, Object))
            Dim queryText As String = query
            If Not dslFts Then
                queryText = "!L " & queryText
            End If

            Dim body As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                {"q", queryText},
                {"size", "10000"},
                {"from", "0"},
                {"scroll", "true"}
            }
            If params.ContainsKey("scope") Then
                body("scope") = params("scope")
            End If
            If params.ContainsKey("index") Then
                body("scope") = params("index")
            End If
            If params.ContainsKey("size") Then
                body("size") = params("size")
                body("scroll") = "false"
            End If

            Dim url As String = session.Protocol & "//be-api.us.archive.org/ia-pub-fts-api"
            Dim serializer As New JavaScriptSerializer()
            Dim results As New List(Of Dictionary(Of String, Object))()
            While True
                Dim payload As String = serializer.Serialize(body)
                Dim json = SendRequestAndDeserialize(
                    session,
                    HttpMethod.Post,
                    url,
                    Nothing,
                    payload,
                    headers,
                    timeoutSeconds
                )
                Dim hitsObj As Object = Nothing
                If Not json.TryGetValue("hits", hitsObj) Then
                    Exit While
                End If
                Dim hitsDict = TryCast(hitsObj, Dictionary(Of String, Object))
                If hitsDict Is Nothing OrElse Not hitsDict.ContainsKey("hits") Then
                    Exit While
                End If

                Dim innerHits = TryCast(hitsDict("hits"), ArrayList)
                Dim hitSeq = ApiShared.ToObjectSequence(hitsDict("hits"))
                If hitSeq.Count = 0 Then
                    Exit While
                End If
                For Each hitObj In hitSeq
                    Dim hit = TryCast(hitObj, Dictionary(Of String, Object))
                    If hit IsNot Nothing Then
                        results.Add(hit)
                    End If
                Next

                Dim scroll As String = Convert.ToString(body("scroll"), CultureInfo.InvariantCulture)
                If String.Equals(scroll, "false", StringComparison.OrdinalIgnoreCase) Then
                    Exit While
                End If
                If json.ContainsKey("_scroll_id") Then
                    body("scroll_id") = Convert.ToString(json("_scroll_id"))
                Else
                    Exit While
                End If
            End While
            Return results
        End Function

        Private Shared Function GetNumFound(
            session As ArchiveSession,
            query As String,
            params As Dictionary(Of String, Object),
            headers As Dictionary(Of String, Object),
            timeoutSeconds As Double,
            fts As Boolean,
            dslFts As Boolean
        ) As Integer
            If fts OrElse dslFts Then
                Dim queryText As String = query
                If fts AndAlso Not dslFts Then
                    queryText = "!L " & queryText
                End If
                Dim p = BuildDefaultSearchParams(queryText, params)
                Dim url As String = session.Protocol & "//be-api.us.archive.org/ia-pub-fts-api"
                Dim json = SendRequestAndDeserialize(
                    session,
                    HttpMethod.Get,
                    url,
                    p,
                    Nothing,
                    headers,
                    timeoutSeconds
                )
                Dim hitsObj As Object = Nothing
                If json.TryGetValue("hits", hitsObj) Then
                    Dim hits = TryCast(hitsObj, Dictionary(Of String, Object))
                    If hits IsNot Nothing AndAlso hits.ContainsKey("total") Then
                        Return Convert.ToInt32(hits("total"), CultureInfo.InvariantCulture)
                    End If
                End If
                Return 0
            End If

            Dim searchParams = BuildDefaultSearchParams(query, params)
            If searchParams.ContainsKey("page") Then
                searchParams("output") = "json"
                Dim url As String = session.Protocol & "//" & session.Host & "/advancedsearch.php"
                Dim json = SendRequestAndDeserialize(
                    session,
                    HttpMethod.Get,
                    url,
                    searchParams,
                    Nothing,
                    headers,
                    timeoutSeconds
                )
                Return ExtractAdvancedNumFound(json)
            End If

            searchParams("total_only") = "true"
            Dim scrapeUrl As String = session.Protocol & "//" & session.Host & "/services/search/v1/scrape"
            Dim scrape = SendRequestAndDeserialize(
                session,
                HttpMethod.Post,
                scrapeUrl,
                searchParams,
                Nothing,
                headers,
                timeoutSeconds
            )
            HandleScrapeError(scrape)
            If scrape.ContainsKey("total") Then
                Return Convert.ToInt32(scrape("total"), CultureInfo.InvariantCulture)
            End If
            Return 0
        End Function

        Private Shared Sub HandleScrapeError(payload As Dictionary(Of String, Object))
            If Not payload.ContainsKey("error") Then
                Return
            End If
            Dim message As String = Convert.ToString(payload("error"))
            Dim lower As String = message.ToLowerInvariant()
            If lower.Contains("invalid") AndAlso lower.Contains("secret") Then
                If Not message.EndsWith(".", StringComparison.Ordinal) Then
                    message &= "."
                End If
                Throw New ValueErrorException(message & " Try running 'ia configure' and retrying.")
            End If
            Throw New ValueErrorException(message)
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

        Private Shared Function ParseArguments(args As IList(Of String)) As SearchArgs
            Dim parsed As New SearchArgs With {
                .Parameters = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Headers = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Sorts = New List(Of String)(),
                .Fields = New List(Of String)()
            }

            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-p", "--parameters"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryStringOption(parsed.Parameters, args(i), current)
                    Case "-H", "--header"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryStringOption(parsed.Headers, args(i), current)
                    Case "-s", "--sort"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Sorts.Add(args(i))
                    Case "-i", "--itemlist"
                        parsed.ItemList = True
                    Case "-f", "--field"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Fields.Add(args(i))
                    Case "-n", "--num-found"
                        parsed.NumFound = True
                    Case "-F", "--fts"
                        parsed.Fts = True
                    Case "-D", "--dsl-fts"
                        parsed.DslFts = True
                    Case "-t", "--timeout"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.TimeoutSeconds = Double.Parse(args(i), CultureInfo.InvariantCulture)
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            unknown.Add(current)
                            i += 1
                            Continue While
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Query) Then
                            parsed.Query = current
                        Else
                            Throw New ArgumentException("unrecognized arguments")
                        End If
                End Select
                i += 1
            End While

            If String.IsNullOrWhiteSpace(parsed.Query) AndAlso Not parsed.ShowHelp Then
                Throw New ArgumentException("the following arguments are required: query")
            End If
            If unknown.Count > 0 Then
                Throw New ArgumentException("unrecognized arguments: " & String.Join(" ", unknown))
            End If

            Return parsed
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

        Private Shared Function PrepareValues(values As List(Of String)) As List(Of String)
            If values Is Nothing OrElse values.Count = 0 Then
                Return Nothing
            End If

            Dim prepared As New List(Of String)()
            For Each entry In values
                Dim parts = entry.Split(New Char() {","c}, StringSplitOptions.None)
                For Each part In parts
                    prepared.Add(part)
                Next
            Next
            Return prepared
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " search [-h] [-p KEY:VALUE] [-H KEY:VALUE] [-s SORT] [-i]")
            Console.Error.WriteLine("                    [-f FIELD] [-n] [-F] [-D] [-t TIMEOUT]")
            Console.Error.WriteLine("                    query")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  query                 Search query or queries.")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -p KEY:VALUE, --parameters KEY:VALUE")
            Console.Error.WriteLine("  -H KEY:VALUE, --header KEY:VALUE")
            Console.Error.WriteLine("  -s SORT, --sort SORT")
            Console.Error.WriteLine("  -i, --itemlist")
            Console.Error.WriteLine("  -f FIELD, --field FIELD")
            Console.Error.WriteLine("  -n, --num-found")
            Console.Error.WriteLine("  -F, --fts")
            Console.Error.WriteLine("  -D, --dsl-fts")
            Console.Error.WriteLine("  -t TIMEOUT, --timeout TIMEOUT")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " search [-h] [-p KEY:VALUE] [-H KEY:VALUE] [-s SORT] [-i]")
            Console.Error.WriteLine("                    [-f FIELD] [-n] [-F] [-D] [-t TIMEOUT]")
            Console.Error.WriteLine("                    query")
            Console.Error.WriteLine(CliApp.ExecutableName() & " search: error: " & message)
        End Sub

        Private Shared Function ScrapeSearch(
            session As ArchiveSession,
            params As Dictionary(Of String, Object),
            headers As Dictionary(Of String, Object),
            fields As List(Of String),
            sorts As List(Of String),
            timeoutSeconds As Double
        ) As IEnumerable(Of Dictionary(Of String, Object))
            Dim p As New Dictionary(Of String, Object)(params, StringComparer.OrdinalIgnoreCase)
            If fields IsNot Nothing AndAlso fields.Count > 0 Then
                p("fields") = String.Join(",", fields)
            End If
            If sorts IsNot Nothing AndAlso sorts.Count > 0 Then
                p("sorts") = String.Join(",", sorts)
            End If

            Dim url As String = session.Protocol & "//" & session.Host & "/services/search/v1/scrape"
            Dim results As New List(Of Dictionary(Of String, Object))()
            Dim seen As Integer = 0
            Dim total As Integer? = Nothing

            While True
                Dim json = SendRequestAndDeserialize(
                    session,
                    HttpMethod.Post,
                    url,
                    p,
                    Nothing,
                    headers,
                    timeoutSeconds
                )

                If json.ContainsKey("error") Then
                    results.Add(json)
                    Return results
                End If
                HandleScrapeError(json)

                If Not total.HasValue AndAlso json.ContainsKey("total") Then
                    total = Convert.ToInt32(json("total"), CultureInfo.InvariantCulture)
                End If

                Dim itemsObj As Object = Nothing
                If json.TryGetValue("items", itemsObj) Then
                    For Each itemObj In ApiShared.ToObjectSequence(itemsObj)
                        Dim item = TryCast(itemObj, Dictionary(Of String, Object))
                        If item IsNot Nothing Then
                            seen += 1
                            results.Add(item)
                        End If
                    Next
                End If

                If json.ContainsKey("cursor") Then
                    p("cursor") = Convert.ToString(json("cursor"))
                Else
                    If total.HasValue AndAlso seen <> total.Value Then
                        Throw New TimeoutException("server did not return all results")
                    End If
                    Exit While
                End If
            End While
            Return results
        End Function

        Private Shared Function SendRequestAndDeserialize(
            session As ArchiveSession,
            method As HttpMethod,
            url As String,
            params As Dictionary(Of String, Object),
            jsonBody As String,
            headers As Dictionary(Of String, Object),
            timeoutSeconds As Double
        ) As Dictionary(Of String, Object)
            Dim finalUrl As String = BuildUrl(url, params)
            Using handler As New HttpClientHandler()
                Using client As New HttpClient(handler)
                    client.Timeout = TimeSpan.FromSeconds(timeoutSeconds)
                    Dim request As New HttpRequestMessage(method, finalUrl)
                    ApplyHeaders(request, headers)
                    ApiShared.ApplyLowAuth(request, session.AccessKey, session.SecretKey)
                    If jsonBody IsNot Nothing Then
                        request.Content = New StringContent(jsonBody, Encoding.UTF8, "application/json")
                    End If

                    Dim response = client.SendAsync(request).GetAwaiter().GetResult()
                    Dim payload As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                    Dim serializer As New JavaScriptSerializer()
                    Dim decoded = serializer.DeserializeObject(payload)
                    Dim result = TryCast(decoded, Dictionary(Of String, Object))
                    If result Is Nothing Then
                        Return New Dictionary(Of String, Object)()
                    End If
                    Return result
                End Using
            End Using
        End Function

        Private NotInheritable Class SearchArgs
            Public Property DslFts As Boolean
            Public Property Fields As List(Of String)
            Public Property Fts As Boolean
            Public Property Headers As Dictionary(Of String, Object)
            Public Property ItemList As Boolean
            Public Property NumFound As Boolean
            Public Property Parameters As Dictionary(Of String, Object)
            Public Property Query As String
            Public Property ShowHelp As Boolean
            Public Property Sorts As List(Of String)
            Public Property TimeoutSeconds As Double = 300
        End Class
    End Class

    Public Class ValueErrorException
        Inherits Exception
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub
    End Class
End Namespace

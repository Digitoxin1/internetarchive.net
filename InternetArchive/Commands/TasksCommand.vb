Imports System.Globalization
Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class TasksCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As TasksArgs
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

            If Not String.IsNullOrWhiteSpace(parsed.Cmd) Then
                If parsed.GetRateLimit Then
                    Dim rateLimit = session.GetTasksApiRateLimit(parsed.Cmd)
                    Console.WriteLine(serializer.Serialize(rateLimit))
                    Return 0
                End If

                parsed.Data("args") = parsed.TaskArgs
                Dim priority As Integer = 0
                If parsed.Data.ContainsKey("priority") Then
                    priority = Convert.ToInt32(parsed.Data("priority"), CultureInfo.InvariantCulture)
                End If

                Dim result = session.SubmitTask(
                    parsed.Identifier,
                    parsed.Cmd,
                    parsed.Comment,
                    priority,
                    parsed.Data,
                    parsed.ReducedPriority
                )

                Dim resultJson As Dictionary(Of String, Object) = result.JsonBody
                If resultJson IsNot Nothing AndAlso resultJson.ContainsKey("success") AndAlso
                   Convert.ToBoolean(resultJson("success"), CultureInfo.InvariantCulture) Then
                    Dim valueNode As Object = Nothing
                    Dim taskLogUrl As String = Nothing
                    If resultJson.TryGetValue("value", valueNode) Then
                        Dim valueDict = TryCast(valueNode, Dictionary(Of String, Object))
                        If valueDict IsNot Nothing AndAlso valueDict.ContainsKey("log") Then
                            taskLogUrl = Convert.ToString(valueDict("log"), CultureInfo.InvariantCulture)
                        End If
                    End If
                    Console.Error.WriteLine("success: " & taskLogUrl)
                    Return 0
                End If

                Dim err As String = ""
                If resultJson IsNot Nothing AndAlso resultJson.ContainsKey("error") Then
                    err = Convert.ToString(resultJson("error"), CultureInfo.InvariantCulture)
                End If
                If err.IndexOf("already queued/running", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Console.Error.WriteLine(
                        String.Format("success: {0} task already queued/running", parsed.Cmd)
                    )
                    Return 0
                End If
                Console.Error.WriteLine("error: " & err)
                Return 1
            End If

            If Not String.IsNullOrWhiteSpace(parsed.Identifier) Then
                Dim p As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"identifier", parsed.Identifier},
                    {"catalog", 1},
                    {"history", 1}
                }
                For Each kvp In parsed.Parameters
                    p(kvp.Key) = kvp.Value
                Next
                parsed.Parameters = p
            ElseIf Not String.IsNullOrWhiteSpace(parsed.GetTaskLog) Then
                Dim logText As String = session.GetTaskLog(parsed.GetTaskLog)
                Console.WriteLine(logText)
                Return 0
            End If

            Dim queryableParams As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "identifier", "task_id", "server", "cmd", "args", "submitter",
                "priority", "wait_admin", "submittime"
            }

            If String.IsNullOrWhiteSpace(parsed.Identifier) AndAlso
               Not parsed.Parameters.ContainsKey("task_id") Then
                Dim p As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"catalog", 1},
                    {"history", 0}
                }
                For Each kvp In parsed.Parameters
                    p(kvp.Key) = kvp.Value
                Next
                parsed.Parameters = p
            End If

            Dim hasQueryable As Boolean = False
            For Each key In parsed.Parameters.Keys
                If queryableParams.Contains(key) Then
                    hasQueryable = True
                    Exit For
                End If
            Next
            If Not hasQueryable Then
                Dim p As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"submitter", session.UserEmail},
                    {"catalog", 1},
                    {"history", 0},
                    {"summary", 0}
                }
                For Each kvp In parsed.Parameters
                    p(kvp.Key) = kvp.Value
                Next
                parsed.Parameters = p
            End If

            If parsed.TabOutput Then
                Console.Error.WriteLine(
                    "tab-delimited output will be removed in a future release. Please switch to the default JSON output."
                )
            End If

            Dim tasks = session.GetTasks(parsed.Parameters)
            For Each t In tasks
                If parsed.TabOutput Then
                    Dim color As String = "done"
                    If t.ContainsKey("color") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(t("color"))) Then
                        color = Convert.ToString(t("color"))
                    End If
                    Dim taskArgsText As String = ""
                    If t.ContainsKey("args") Then
                        taskArgsText = JoinArgs(t("args"))
                    End If
                    Dim columns As New List(Of String) From {
                        GetOrEmpty(t, "identifier"),
                        GetOrEmpty(t, "task_id"),
                        GetOrEmpty(t, "server"),
                        GetOrEmpty(t, "submittime"),
                        GetOrEmpty(t, "cmd"),
                        color,
                        GetOrEmpty(t, "submitter"),
                        taskArgsText
                    }
                    columns = columns.FindAll(Function(x) x IsNot Nothing AndAlso x.Length > 0)
                    Console.WriteLine(String.Join(vbTab, columns))
                    Console.Out.Flush()
                Else
                    Console.WriteLine(serializer.Serialize(t))
                    Console.Out.Flush()
                End If
            Next

            Return 0
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function GetOrEmpty(
            task As Dictionary(Of String, Object),
            key As String
        ) As String
            If task.ContainsKey(key) Then
                Return Convert.ToString(task(key), CultureInfo.InvariantCulture)
            End If
            Return ""
        End Function

        Private Shared Function JoinArgs(value As Object) As String
            Dim dict = TryCast(value, Dictionary(Of String, Object))
            If dict IsNot Nothing Then
                Dim pieces As New List(Of String)()
                For Each kvp In dict
                    pieces.Add(String.Format("{0}={1}", kvp.Key, Convert.ToString(kvp.Value)))
                Next
                Return String.Join(vbTab, pieces)
            End If
            Dim arr = TryCast(value, ArrayList)
            If arr IsNot Nothing Then
                Dim pieces As New List(Of String)()
                For Each entry In arr
                    pieces.Add(Convert.ToString(entry))
                Next
                Return String.Join(vbTab, pieces)
            End If
            Return Convert.ToString(value)
        End Function

        Private Shared Sub MergePostData(
            destination As Dictionary(Of String, Object),
            raw As String,
            optionName As String
        )
            Dim serializer As New JavaScriptSerializer()
            Try
                Dim parsedJson = serializer.DeserializeObject(raw)
                Dim dict = TryCast(parsedJson, Dictionary(Of String, Object))
                If dict IsNot Nothing Then
                    For Each kvp In dict
                        destination(kvp.Key) = kvp.Value
                    Next
                    Return
                End If
                Throw New ArgumentException(
                    String.Format("{0} JSON must be an object, not {1}", optionName, parsedJson.GetType().Name)
                )
            Catch ex As ArgumentException
                If ex.Message.IndexOf("JSON must be an object", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Throw
                End If
                ' fall through to key:value mode
            Catch
                ' fall through to key:value mode
            End Try

            If raw.Contains(":") Then
                Dim parts = raw.Split(New Char() {":"c}, 2)
                destination(parts(0)) = parts(1)
            ElseIf raw.Contains("=") Then
                Dim parts = raw.Split(New Char() {"="c}, 2)
                destination(parts(0)) = parts(1)
            Else
                Throw New ArgumentException(
                    String.Format("{0} must be a JSON object or 'key:value' format", optionName)
                )
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

            Dim pairs = ParseQueryString(normalized)
            If normalized.Length > 0 AndAlso pairs.Count = 0 Then
                Throw New ArgumentException(
                    String.Format("{0} must be formatted as 'key=value' or 'key:value'", optionName)
                )
            End If

            For Each pair In pairs
                If destination.ContainsKey(pair.Key) Then
                    Dim existing = destination(pair.Key)
                    Dim list = If(TryCast(existing, List(Of String)), New List(Of String) From {Convert.ToString(existing, CultureInfo.InvariantCulture)})
                    list.Add(pair.Value)
                    destination(pair.Key) = list
                Else
                    destination(pair.Key) = pair.Value
                End If
            Next
        End Sub

        Private Shared Function ParseArguments(args As IList(Of String)) As TasksArgs
            Dim parsed As New TasksArgs With {
                .TaskIds = New List(Of String)(),
                .Parameters = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .TaskArgs = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                .Data = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            }

            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-t", "--task"
                        i += 1
                        While i < args.Count AndAlso Not args(i).StartsWith("-", StringComparison.Ordinal)
                            parsed.TaskIds.Add(args(i))
                            i += 1
                        End While
                        i -= 1
                    Case "-G", "--get-task-log"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.GetTaskLog = args(i)
                    Case "-p", "--parameter"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.Parameters, args(i), current)
                    Case "-T", "--tab-output"
                        parsed.TabOutput = True
                    Case "-c", "--cmd"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Cmd = args(i)
                    Case "-C", "--comment"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Comment = args(i)
                    Case "-a", "--task-args"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergeQueryString(parsed.TaskArgs, args(i), current)
                    Case "-d", "--data"
                        i += 1
                        EnsureHasValue(args, i, current)
                        MergePostData(parsed.Data, args(i), current)
                    Case "-r", "--reduced-priority"
                        parsed.ReducedPriority = True
                    Case "-l", "--get-rate-limit"
                        parsed.GetRateLimit = True
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

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " tasks [-h] [-t [TASK ...]] [-G GET_TASK_LOG] [-p KEY:VALUE] [-T]")
            Console.Error.WriteLine("                   [-c CMD] [-C COMMENT] [-a KEY:VALUE] [-d DATA] [-r] [-l]")
            Console.Error.WriteLine("                   [identifier]")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier for tasks specific operations.")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -t [TASK ...], --task [TASK ...]")
            Console.Error.WriteLine("  -G GET_TASK_LOG, --get-task-log GET_TASK_LOG")
            Console.Error.WriteLine("  -p KEY:VALUE, --parameter KEY:VALUE")
            Console.Error.WriteLine("  -T, --tab-output")
            Console.Error.WriteLine("  -c CMD, --cmd CMD")
            Console.Error.WriteLine("  -C COMMENT, --comment COMMENT")
            Console.Error.WriteLine("  -a KEY:VALUE, --task-args KEY:VALUE")
            Console.Error.WriteLine("  -d DATA, --data DATA")
            Console.Error.WriteLine("  -r, --reduced-priority")
            Console.Error.WriteLine("  -l, --get-rate-limit")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " tasks [-h] [-t [TASK ...]] [-G GET_TASK_LOG] [-p KEY:VALUE] [-T]")
            Console.Error.WriteLine("                   [-c CMD] [-C COMMENT] [-a KEY:VALUE] [-d DATA] [-r] [-l]")
            Console.Error.WriteLine("                   [identifier]")
            Console.Error.WriteLine(CliApp.ExecutableName() & " tasks: error: " & message)
        End Sub

        Private NotInheritable Class TasksArgs
            Public Property Cmd As String
            Public Property Comment As String
            Public Property Data As Dictionary(Of String, Object)
            Public Property GetRateLimit As Boolean
            Public Property GetTaskLog As String
            Public Property Identifier As String
            Public Property Parameters As Dictionary(Of String, Object)
            Public Property ReducedPriority As Boolean
            Public Property ShowHelp As Boolean
            Public Property TabOutput As Boolean
            Public Property TaskArgs As Dictionary(Of String, Object)
            Public Property TaskIds As List(Of String)
        End Class
    End Class
End Namespace

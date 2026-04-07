Imports System.Globalization
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class ReviewsCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed = ParseArguments(args)
            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If

            If parsed.IndexReview Then
                Dim r = session.IndexReview(
                    parsed.Identifier,
                    parsed.Username,
                    parsed.Screenname,
                    parsed.Itemname
                )
                If IsSuccess(r.JsonBody) Then
                    Console.Error.WriteLine(parsed.Identifier & " - success: review indexed")
                    Return 0
                End If
                Console.Error.WriteLine(
                    String.Format(
                        "{0} - error: failed to index review{1}",
                        parsed.Identifier,
                        FormatApiErrorSuffix(r)
                    )
                )
                Return 1
            ElseIf parsed.NoIndexReview Then
                Dim r = session.NoIndexReview(
                    parsed.Identifier,
                    parsed.Username,
                    parsed.Screenname,
                    parsed.Itemname
                )
                If IsSuccess(r.JsonBody) Then
                    Console.Error.WriteLine(parsed.Identifier & " - success: review removed from index")
                    Return 0
                End If
                Console.Error.WriteLine(
                    String.Format(
                        "{0} - error: failed to remove review from index{1}",
                        parsed.Identifier,
                        FormatApiErrorSuffix(r)
                    )
                )
                Return 1
            End If

            Dim result As ApiCallResult
            If parsed.DeleteReview Then
                result = session.DeleteReview(
                    parsed.Identifier,
                    parsed.Username,
                    parsed.Screenname,
                    parsed.Itemname
                )
            ElseIf String.IsNullOrWhiteSpace(parsed.Body) AndAlso String.IsNullOrWhiteSpace(parsed.Title) Then
                result = session.GetReview(parsed.Identifier)
                If result.StatusCode = 404 Then
                    Return 0
                End If
                If result.StatusCode >= 400 Then
                    Throw New InvalidOperationException(String.Format(
                        "HTTP {0} while getting review.",
                        result.StatusCode
                    ))
                End If
                Console.WriteLine(result.Text)
                Return 0
            Else
                If (Not String.IsNullOrWhiteSpace(parsed.Title) AndAlso String.IsNullOrWhiteSpace(parsed.Body)) OrElse
                   (Not String.IsNullOrWhiteSpace(parsed.Body) AndAlso String.IsNullOrWhiteSpace(parsed.Title)) Then
                    PrintUsageAndError("both --title and --body must be provided")
                    Return 2
                End If
                result = session.SubmitReview(
                    parsed.Identifier,
                    parsed.Title,
                    parsed.Body,
                    parsed.Stars
                )
            End If

            Dim j = result.JsonBody
            Dim err As String = ""
            If j IsNot Nothing AndAlso j.ContainsKey("error") Then
                err = Convert.ToString(j("error"), CultureInfo.InvariantCulture)
            End If
            If IsSuccess(j) OrElse err.ToLowerInvariant().Contains("no change detected") Then
                Dim taskId As String = Nothing
                If j IsNot Nothing AndAlso j.ContainsKey("value") Then
                    Dim val = TryCast(j("value"), Dictionary(Of String, Object))
                    If val IsNot Nothing AndAlso val.ContainsKey("task_id") Then
                        taskId = Convert.ToString(val("task_id"), CultureInfo.InvariantCulture)
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(taskId) Then
                    Console.Error.WriteLine(String.Format(
                        "{0} - success: https://catalogd.archive.org/log/{1}",
                        parsed.Identifier,
                        taskId
                    ))
                Else
                    Console.Error.WriteLine(parsed.Identifier & " - warning: no changes detected!")
                End If
                Return 0
            End If

            Console.Error.WriteLine(String.Format("{0} - error: {1}", parsed.Identifier, err))
            Return 1
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function IsSuccess(body As Dictionary(Of String, Object)) As Boolean
            If body Is Nothing OrElse Not body.ContainsKey("success") Then
                Return False
            End If
            Return Convert.ToBoolean(body("success"), CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatApiErrorSuffix(result As ApiCallResult) As String
            If result Is Nothing Then
                Return String.Empty
            End If

            If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("error") Then
                Dim err As String = Convert.ToString(result.JsonBody("error"), CultureInfo.InvariantCulture)
                If Not String.IsNullOrWhiteSpace(err) Then
                    Return ": " & err
                End If
            End If

            If Not String.IsNullOrWhiteSpace(result.Text) Then
                Return ": " & result.Text
            End If

            If result.StatusCode > 0 Then
                Return String.Format(" (HTTP {0})", result.StatusCode)
            End If

            Return String.Empty
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As ReviewsArgs
            Dim parsed As New ReviewsArgs()
            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-d", "--delete"
                        parsed.DeleteReview = True
                    Case "-t", "--title"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Title = args(i)
                    Case "-b", "--body"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Body = args(i)
                    Case "-s", "--stars"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Stars = Integer.Parse(args(i), CultureInfo.InvariantCulture)
                    Case "-i", "--index"
                        parsed.IndexReview = True
                    Case "-n", "--noindex"
                        parsed.NoIndexReview = True
                    Case "-u", "--username"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Username = args(i)
                    Case "-S", "--screenname"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Screenname = args(i)
                    Case "-I", "--itemname"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Itemname = args(i)
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
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " reviews [-h] [-d] [-t TITLE] [-b BODY] [-s STARS] [-i] [-n]")
            Console.Error.WriteLine("                     [-u USERNAME] [-S SCREENNAME] [-I ITEMNAME]")
            Console.Error.WriteLine("                     identifier")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier of the item")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -d, --delete          Delete your review")
            Console.Error.WriteLine("  -t TITLE, --title TITLE")
            Console.Error.WriteLine("  -b BODY, --body BODY")
            Console.Error.WriteLine("  -s STARS, --stars STARS")
            Console.Error.WriteLine("  -i, --index")
            Console.Error.WriteLine("  -n, --noindex")
            Console.Error.WriteLine("  -u USERNAME, --username USERNAME")
            Console.Error.WriteLine("  -S SCREENNAME, --screenname SCREENNAME")
            Console.Error.WriteLine("  -I ITEMNAME, --itemname ITEMNAME")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " reviews [-h] [-d] [-t TITLE] [-b BODY] [-s STARS] [-i] [-n]")
            Console.Error.WriteLine("                     [-u USERNAME] [-S SCREENNAME] [-I ITEMNAME]")
            Console.Error.WriteLine("                     identifier")
            Console.Error.WriteLine(CliApp.ExecutableName() & " reviews: error: " & message)
        End Sub

        Private NotInheritable Class ReviewsArgs
            Public Property Body As String
            Public Property DeleteReview As Boolean
            Public Property Identifier As String
            Public Property IndexReview As Boolean
            Public Property Itemname As String
            Public Property NoIndexReview As Boolean
            Public Property Screenname As String
            Public Property ShowHelp As Boolean
            Public Property Stars As Integer?
            Public Property Title As String
            Public Property Username As String
        End Class
    End Class
End Namespace

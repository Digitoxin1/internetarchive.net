Imports System.Globalization
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class FlagCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As FlagArgs
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
            Dim flagUser As String = parsed.User
            If String.IsNullOrWhiteSpace(flagUser) Then
                flagUser = GetScreenname(session)
            End If
            If Not flagUser.StartsWith("@", StringComparison.Ordinal) Then
                flagUser = "@" & flagUser
            End If

            If Not String.IsNullOrWhiteSpace(parsed.AddFlag) Then
                Dim result = session.AddFlag(parsed.Identifier, parsed.AddFlag, flagUser)
                If result.StatusCode >= 400 Then
                    Console.Error.WriteLine(String.Format("error: {0} - {1}", parsed.Identifier, DescribeApiFailure(result)))
                    Return 1
                End If
                Dim status As String = GetResultStatus(result)
                If String.Equals(status, "success", StringComparison.OrdinalIgnoreCase) Then
                    Console.WriteLine(
                        String.Format(
                            "success: added '{0}' flag by {1} to {2}",
                            parsed.AddFlag,
                            flagUser,
                            parsed.Identifier
                        )
                    )
                    Return 0
                End If
                Console.Error.WriteLine(String.Format("error: {0} - {1}", parsed.Identifier, DescribeApiFailure(result)))
                Return 1
            End If

            If Not String.IsNullOrWhiteSpace(parsed.DeleteFlag) Then
                Dim result = session.DeleteFlag(parsed.Identifier, parsed.DeleteFlag, flagUser)
                If result.StatusCode >= 400 Then
                    Console.Error.WriteLine(String.Format("error: {0} - {1}", parsed.Identifier, DescribeApiFailure(result)))
                    Return 1
                End If
                Dim status As String = GetResultStatus(result)
                If String.Equals(status, "success", StringComparison.OrdinalIgnoreCase) Then
                    Console.WriteLine(
                        String.Format(
                            "success: deleted '{0}' flag by {1} from {2}",
                            parsed.DeleteFlag,
                            flagUser,
                            parsed.Identifier
                        )
                    )
                    Return 0
                End If
                Console.Error.WriteLine(String.Format("error: {0} - {1}", parsed.Identifier, DescribeApiFailure(result)))
                Return 1
            End If

            Dim flags = session.GetFlags(parsed.Identifier)
            If flags.StatusCode >= 400 Then
                Console.Error.WriteLine(String.Format("error: {0} - {1}", parsed.Identifier, DescribeApiFailure(flags)))
                Return 1
            End If
            Console.WriteLine(flags.Text)
            Return 0
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function GetResultStatus(result As ApiCallResult) As String
            If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("status") Then
                Return Convert.ToString(result.JsonBody("status"), CultureInfo.InvariantCulture)
            End If
            Return ""
        End Function

        Private Shared Function DescribeApiFailure(result As ApiCallResult) As String
            If result Is Nothing Then
                Return "request failed"
            End If

            If result.JsonBody IsNot Nothing AndAlso result.JsonBody.ContainsKey("error") Then
                Dim errorText As String = Convert.ToString(result.JsonBody("error"), CultureInfo.InvariantCulture)
                If Not String.IsNullOrWhiteSpace(errorText) Then
                    Return errorText
                End If
            End If

            If Not String.IsNullOrWhiteSpace(result.Text) Then
                Return result.Text
            End If

            If result.StatusCode > 0 Then
                Return String.Format("HTTP {0}", result.StatusCode)
            End If

            Return "request failed"
        End Function

        Private Shared Function GetScreenname(session As ArchiveSession) As String
            Dim general As Dictionary(Of String, String) = Nothing
            If session.Config IsNot Nothing AndAlso
               session.Config.TryGetValue("general", general) AndAlso
               general IsNot Nothing AndAlso
               general.ContainsKey("screenname") Then
                Return Convert.ToString(general("screenname"), CultureInfo.InvariantCulture)
            End If
            Return ""
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As FlagArgs
            Dim parsed As New FlagArgs()
            Dim i As Integer = 0
            Dim unknown As New List(Of String)()
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-u", "--user"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.User = args(i)
                    Case "-a", "--add-flag"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.AddFlag = args(i)
                    Case "-d", "--delete-flag"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.DeleteFlag = args(i)
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

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " flag [-h] [-u USER] [-a CATEGORY] [-d CATEGORY] [identifier]")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier of the item")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -u USER, --user USER  User associated with the flag")
            Console.Error.WriteLine()
            Console.Error.WriteLine("Add flag operations:")
            Console.Error.WriteLine("  -a CATEGORY, --add-flag CATEGORY")
            Console.Error.WriteLine("                        Add a flag to the item")
            Console.Error.WriteLine()
            Console.Error.WriteLine("Delete flag operations:")
            Console.Error.WriteLine("  -d CATEGORY, --delete-flag CATEGORY")
            Console.Error.WriteLine("                        Delete a flag from the item")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " flag [-h] [-u USER] [-a CATEGORY] [-d CATEGORY] [identifier]")
            Console.Error.WriteLine(CliApp.ExecutableName() & " flag: error: " & message)
        End Sub

        Private NotInheritable Class FlagArgs
            Public Property AddFlag As String
            Public Property DeleteFlag As String
            Public Property Identifier As String
            Public Property ShowHelp As Boolean
            Public Property User As String
        End Class
    End Class
End Namespace

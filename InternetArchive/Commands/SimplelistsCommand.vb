Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Core
Imports InternetArchive.InternetArchiveCli.Exceptions

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class SimplelistsCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As SimplelistsArgs
            Try
                parsed = ParseArguments(args)
            Catch ex As ArgumentException
                PrintUsageAndError(ex.Message)
                Return 2
            End Try
            Dim serializer As New JavaScriptSerializer()

            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If

            If parsed.ListParents Then
                Try
                    Dim item = session.GetItemMetadata(parsed.Identifier, Nothing)
                    Dim simplelistsObj As Object = Nothing
                    If item.TryGetValue("simplelists", simplelistsObj) AndAlso simplelistsObj IsNot Nothing Then
                        Console.WriteLine(serializer.Serialize(simplelistsObj))
                    End If
                    Return 0
                Catch ex As AuthenticationError
                    Console.Error.WriteLine("error: " & ex.Message)
                    Return 1
                Catch ex As Exception
                    Console.Error.WriteLine("error: " & ex.Message)
                    Return 1
                End Try
            End If

            If parsed.ListChildren Then
                Try
                    Dim listName As String = If(String.IsNullOrWhiteSpace(parsed.ListName), "catchall", parsed.ListName)
                    Dim queryIdentifier As String = If(String.IsNullOrWhiteSpace(parsed.Identifier), "*", parsed.Identifier)
                    Dim query As String = String.Format("simplelists__{0}:{1}", listName, queryIdentifier)
                    Dim results = session.SearchItemsSimple(query, Nothing)
                    For Each result In results
                        Console.WriteLine(serializer.Serialize(result))
                    Next
                    Return 0
                Catch ex As AuthenticationError
                    Console.Error.WriteLine("error: " & ex.Message)
                    Return 1
                Catch ex As Exception
                    Console.Error.WriteLine("error: " & ex.Message)
                    Return 1
                End Try
            End If

            If Not String.IsNullOrWhiteSpace(parsed.SetParent) Then
                Return HandlePatchOperation(session, parsed, "set")
            End If
            If Not String.IsNullOrWhiteSpace(parsed.RemoveParent) Then
                Return HandlePatchOperation(session, parsed, "delete")
            End If

            PrintHelp()
            Return 1
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function HandlePatchOperation(
            session As ArchiveSession,
            parsed As SimplelistsArgs,
            operation As String
        ) As Integer
            If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                PrintUsageAndError("Missing required identifier argument")
                Return 2
            End If
            If String.IsNullOrWhiteSpace(parsed.ListName) Then
                PrintUsageAndError("Must specify list name with -l/--list-name")
                Return 2
            End If

            Dim patch As New Dictionary(Of String, Object) From {
                {"op", operation},
                {"parent", If(operation = "set", parsed.SetParent, parsed.RemoveParent)},
                {"list", parsed.ListName}
            }
            If Not String.IsNullOrWhiteSpace(parsed.Notes) Then
                patch("notes") = parsed.Notes
            End If

            Try
                Dim r = session.PostSimplelistPatch(parsed.Identifier, patch)
                If r.StatusCode >= 200 AndAlso r.StatusCode < 300 Then
                    Console.WriteLine("success: " & parsed.Identifier)
                    Return 0
                End If
                Console.Error.WriteLine(
                    String.Format(
                        "error: {0} - HTTP {1}",
                        parsed.Identifier,
                        r.StatusCode
                    )
                )
                Return 1
            Catch ex As Exception
                Console.Error.WriteLine(
                    String.Format("error: {0} - {1}", parsed.Identifier, ex.Message)
                )
                Return 1
            End Try
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As SimplelistsArgs
            Dim parsed As New SimplelistsArgs()
            Dim i As Integer = 0
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-p", "--list-parents"
                        parsed.ListParents = True
                    Case "-c", "--list-children"
                        parsed.ListChildren = True
                    Case "-l", "--list-name"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.ListName = args(i)
                    Case "-s", "--set-parent"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.SetParent = args(i)
                    Case "-n", "--notes"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Notes = args(i)
                    Case "-r", "--remove-parent"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.RemoveParent = args(i)
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            Throw New ArgumentException(String.Format("Unknown option: {0}", current))
                        End If
                        If String.IsNullOrWhiteSpace(parsed.Identifier) Then
                            parsed.Identifier = current
                        Else
                            Throw New ArgumentException("unrecognized arguments")
                        End If
                End Select
                i += 1
            End While
            Return parsed
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " simplelists [-h] [-p] [-c] [-l LIST_NAME] [-s PARENT] [-n NOTES]")
            Console.Error.WriteLine("                         [-r PARENT]")
            Console.Error.WriteLine("                         [identifier]")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  identifier            Identifier of the item")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine()
            Console.Error.WriteLine("List operations:")
            Console.Error.WriteLine("  -p, --list-parents    List parent lists for the given identifier")
            Console.Error.WriteLine("  -c, --list-children   List children in parent list")
            Console.Error.WriteLine("  -l LIST_NAME, --list-name LIST_NAME")
            Console.Error.WriteLine("                        Name of the list to operate on")
            Console.Error.WriteLine()
            Console.Error.WriteLine("Modification operations:")
            Console.Error.WriteLine("  -s PARENT, --set-parent PARENT")
            Console.Error.WriteLine("                        Add identifier to specified parent list")
            Console.Error.WriteLine("  -n NOTES, --notes NOTES")
            Console.Error.WriteLine("                        Notes to attach to the list membership")
            Console.Error.WriteLine("  -r PARENT, --remove-parent PARENT")
            Console.Error.WriteLine("                        Remove identifier from specified parent list")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine(
                "usage: " & CliApp.ExecutableName() & " simplelists [-h] [-p] [-c] [-l LIST_NAME] [-s PARENT] [-n NOTES]"
            )
            Console.Error.WriteLine("                         [-r PARENT]")
            Console.Error.WriteLine("                         [identifier]")
            Console.Error.WriteLine(CliApp.ExecutableName() & " simplelists: error: " & message)
        End Sub

        Private NotInheritable Class SimplelistsArgs
            Public Property Identifier As String
            Public Property ListChildren As Boolean
            Public Property ListName As String
            Public Property ListParents As Boolean
            Public Property Notes As String
            Public Property RemoveParent As String
            Public Property SetParent As String
            Public Property ShowHelp As Boolean
        End Class
    End Class
End Namespace

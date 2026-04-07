Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class AccountCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed = ParseArguments(args)
            Dim serializer As New JavaScriptSerializer()

            If parsed.ShowHelp Then
                PrintHelp()
                Return 0
            End If
            If Not String.IsNullOrWhiteSpace(parsed.ParseError) Then
                PrintUsageAndError(parsed.ParseError)
                Return 2
            End If

            Dim identifierType As String
            If parsed.User.StartsWith("@", StringComparison.Ordinal) Then
                identifierType = "itemname"
            ElseIf Not IsValidEmail(parsed.User) Then
                identifierType = "screenname"
            Else
                identifierType = "email"
            End If

            Dim lookup = session.LookupAccount(identifierType, parsed.User)
            Dim values = GetValuesOrError(lookup, serializer)
            If values Is Nothing Then
                Return 1
            End If

            If parsed.GetEmail Then
                Console.WriteLine(GetString(values, "canonical_email"))
                Return 0
            End If
            If parsed.GetScreenname Then
                Console.WriteLine(GetString(values, "screenname"))
                Return 0
            End If
            If parsed.GetItemname Then
                Console.WriteLine(GetString(values, "itemname"))
                Return 0
            End If
            If parsed.IsLocked Then
                Dim lockedObj As Object = Nothing
                If values.TryGetValue("locked", lockedObj) Then
                    Console.WriteLine(Convert.ToString(lockedObj, CultureInfo.InvariantCulture))
                Else
                    Console.WriteLine("False")
                End If
                Return 0
            End If
            If parsed.LockAccount OrElse parsed.UnlockAccount Then
                Dim itemname As String = GetString(values, "itemname")
                Dim r = session.LockUnlockAccount(itemname, parsed.LockAccount, parsed.Comment)
                Console.WriteLine(r.Text)
                Return 0
            End If

            Console.WriteLine(serializer.Serialize(values))
            Return 0
        End Function

        Private Shared Sub EnsureHasValue(args As IList(Of String), index As Integer, optionName As String)
            If index >= args.Count Then
                Throw New ArgumentException(String.Format("Missing value for {0}", optionName))
            End If
        End Sub

        Private Shared Function GetString(values As Dictionary(Of String, Object), key As String) As String
            If values IsNot Nothing AndAlso values.ContainsKey(key) Then
                Return Convert.ToString(values(key), CultureInfo.InvariantCulture)
            End If
            Return ""
        End Function

        Private Shared Function GetValuesOrError(lookup As ApiCallResult, serializer As JavaScriptSerializer) As Dictionary(Of String, Object)
            Dim j = lookup.JsonBody
            If j Is Nothing Then
                Console.WriteLine("null")
                Return Nothing
            End If

            If j.ContainsKey("error") OrElse Not j.ContainsKey("values") Then
                If j.Count = 0 Then
                    Console.WriteLine("null")
                Else
                    Console.WriteLine(serializer.Serialize(j))
                End If
                Return Nothing
            End If
            Dim values = TryCast(j("values"), Dictionary(Of String, Object))
            If values Is Nothing Then
                If j.Count = 0 Then
                    Console.WriteLine("null")
                Else
                    Console.WriteLine(serializer.Serialize(j))
                End If
                Return Nothing
            End If
            Return values
        End Function

        Private Shared Function IsValidEmail(email As String) As Boolean
            Return Regex.IsMatch(email, "^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$")
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As AccountArgs
            Dim parsed As New AccountArgs()
            Dim selectedExclusiveOption As String = ""
            Dim i As Integer = 0
            While i < args.Count
                Dim current As String = args(i)
                Select Case current
                    Case "-h", "--help"
                        parsed.ShowHelp = True
                    Case "-g", "--get-email"
                        If Not TrySetExclusive(selectedExclusiveOption, "-g/--get-email", parsed) Then Return parsed
                        parsed.GetEmail = True
                    Case "-s", "--get-screenname"
                        If Not TrySetExclusive(selectedExclusiveOption, "-s/--get-screenname", parsed) Then Return parsed
                        parsed.GetScreenname = True
                    Case "-i", "--get-itemname"
                        If Not TrySetExclusive(selectedExclusiveOption, "-i/--get-itemname", parsed) Then Return parsed
                        parsed.GetItemname = True
                    Case "-l", "--is-locked"
                        If Not TrySetExclusive(selectedExclusiveOption, "-l/--is-locked", parsed) Then Return parsed
                        parsed.IsLocked = True
                    Case "-L", "--lock"
                        If Not TrySetExclusive(selectedExclusiveOption, "-L/--lock", parsed) Then Return parsed
                        parsed.LockAccount = True
                    Case "-u", "--unlock"
                        If Not TrySetExclusive(selectedExclusiveOption, "-u/--unlock", parsed) Then Return parsed
                        parsed.UnlockAccount = True
                    Case "-c", "--comment"
                        i += 1
                        EnsureHasValue(args, i, current)
                        parsed.Comment = args(i)
                    Case Else
                        If current.StartsWith("-", StringComparison.Ordinal) Then
                            Throw New ArgumentException(String.Format("Unknown option: {0}", current))
                        End If
                        If String.IsNullOrWhiteSpace(parsed.User) Then
                            parsed.User = current
                        Else
                            Throw New ArgumentException("unrecognized arguments")
                        End If
                End Select
                i += 1
            End While

            If String.IsNullOrWhiteSpace(parsed.User) Then
                parsed.ParseError = "the following arguments are required: user"
                Return parsed
            End If

            Return parsed
        End Function

        Private Shared Sub PrintHelp()
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " account [-h] [-g | -s | -i | -l | -L | -u] [-c COMMENT] user")
            Console.Error.WriteLine()
            Console.Error.WriteLine("positional arguments:")
            Console.Error.WriteLine("  user                  Email address, screenname, or itemname for an archive.org account")
            Console.Error.WriteLine()
            Console.Error.WriteLine("options:")
            Console.Error.WriteLine("  -h, --help            show this help message and exit")
            Console.Error.WriteLine("  -g, --get-email       Print the email address associated with the user and exit")
            Console.Error.WriteLine("  -s, --get-screenname  Print the screenname associated with the user and exit")
            Console.Error.WriteLine("  -i, --get-itemname    Print the itemname associated with the user and exit")
            Console.Error.WriteLine("  -l, --is-locked       Check if an account is locked")
            Console.Error.WriteLine("  -L, --lock            Lock an account")
            Console.Error.WriteLine("  -u, --unlock          Unlock an account")
            Console.Error.WriteLine("  -c COMMENT, --comment COMMENT")
            Console.Error.WriteLine("                        Comment to include with lock/unlock action")
        End Sub

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " account [-h] [-g | -s | -i | -l | -L | -u] [-c COMMENT] user")
            Console.Error.WriteLine(CliApp.ExecutableName() & " account: error: " & message)
        End Sub

        Private Shared Function TrySetExclusive(ByRef selectedExclusiveOption As String, currentOption As String, parsed As AccountArgs) As Boolean
            If selectedExclusiveOption.Length = 0 Then
                selectedExclusiveOption = currentOption
                Return True
            End If
            If String.Equals(selectedExclusiveOption, currentOption, StringComparison.Ordinal) Then
                Return True
            End If
            parsed.ParseError = String.Format(
                "argument {0}: not allowed with argument {1}",
                currentOption,
                selectedExclusiveOption
            )
            Return False
        End Function

        Private NotInheritable Class AccountArgs
            Public Property Comment As String
            Public Property GetEmail As Boolean
            Public Property GetItemname As Boolean
            Public Property GetScreenname As Boolean
            Public Property IsLocked As Boolean
            Public Property LockAccount As Boolean
            Public Property ParseError As String
            Public Property ShowHelp As Boolean
            Public Property UnlockAccount As Boolean
            Public Property User As String
        End Class
    End Class
End Namespace

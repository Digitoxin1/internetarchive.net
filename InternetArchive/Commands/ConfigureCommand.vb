Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Core
Imports InternetArchive.InternetArchiveCli.Exceptions
Imports InternetArchive.InternetArchiveCli.Services

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class ConfigureCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As ConfigureArgs = ParseArguments(args)

            If parsed.PrintAuthHeader Then
                Dim secret As String = GetConfigValue(session.Config, "s3", "secret")
                Dim access As String = GetConfigValue(session.Config, "s3", "access")
                If String.IsNullOrWhiteSpace(secret) OrElse String.IsNullOrWhiteSpace(access) Then
                    If String.IsNullOrWhiteSpace(access) Then
                        Console.Error.WriteLine(
                            "error: 'access' key not found in config file, try reconfiguring."
                        )
                    ElseIf String.IsNullOrWhiteSpace(secret) Then
                        Console.Error.WriteLine(
                            "error: 'secret' key not found in config file, try reconfiguring."
                        )
                    End If
                    Return 1
                End If
                Console.WriteLine(String.Format("Authorization: LOW {0}:{1}", access, secret))
                Return 0
            End If

            If parsed.PrintCookies Then
                Dim user As String = GetConfigValue(session.Config, "cookies", "logged-in-user")
                Dim sig As String = GetConfigValue(session.Config, "cookies", "logged-in-sig")
                If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(sig) Then
                    If String.IsNullOrWhiteSpace(user) AndAlso String.IsNullOrWhiteSpace(sig) Then
                        Console.Error.WriteLine(
                            "error: 'logged-in-user' and 'logged-in-sig' cookies not found in config file, try reconfiguring."
                        )
                    ElseIf String.IsNullOrWhiteSpace(user) Then
                        Console.Error.WriteLine(
                            "error: 'logged-in-user' cookie not found in config file, try reconfiguring."
                        )
                    Else
                        Console.Error.WriteLine(
                            "error: 'logged-in-sig' cookie not found in config file, try reconfiguring."
                        )
                    End If
                    Return 1
                End If
                Console.WriteLine(String.Format("logged-in-user={0}; logged-in-sig={1}", user, sig))
                Return 0
            End If

            If parsed.Show Then
                Dim configCopy = DeepCopy(session.Config)
                If configCopy.ContainsKey("s3") AndAlso configCopy("s3").ContainsKey("secret") Then
                    configCopy("s3")("secret") = "REDACTED"
                End If
                If configCopy.ContainsKey("cookies") AndAlso
                   configCopy("cookies").ContainsKey("logged-in-sig") Then
                    configCopy("cookies")("logged-in-sig") = "REDACTED"
                End If

                Dim serializer As New JavaScriptSerializer()
                Console.WriteLine(serializer.Serialize(configCopy))
                Return 0
            End If

            If parsed.WhoAmI Then
                Dim serializer As New JavaScriptSerializer()
                Console.WriteLine(serializer.Serialize(session.WhoAmI()))
                Return 0
            End If

            If parsed.Check Then
                Dim who = session.WhoAmI()
                Dim successObj As Object = Nothing
                If who.TryGetValue("success", successObj) AndAlso Convert.ToBoolean(successObj) Then
                    Dim valueNode As Object = Nothing
                    who.TryGetValue("value", valueNode)
                    Dim valueDict = TryCast(valueNode, Dictionary(Of String, Object))
                    Dim user As String = ""
                    If valueDict IsNot Nothing AndAlso valueDict.ContainsKey("username") Then
                        user = Convert.ToString(valueDict("username"))
                    End If
                    Console.WriteLine(String.Format("The credentials for ""{0}"" are valid", user))
                    Return 0
                End If
                Console.WriteLine("Your credentials are invalid, check your configuration and try again")
                Return 1
            End If

            Try
                If parsed.UseNetrc Then
                    Console.Error.WriteLine("Configuring 'ia' with netrc file...")
                    Dim credentials = ReadNetrcArchiveCredentials()
                    Dim netrcAuthConfig = ConfigService.GetAuthConfig(
                        credentials.Item1,
                        credentials.Item2,
                        session.Host
                    )
                    ConfigService.WriteConfigFile(netrcAuthConfig, session.ConfigFile)
                    Console.Error.WriteLine(
                        String.Format("Config saved to: {0}", ConfigService.ResolveConfigFile(session.ConfigFile))
                    )
                    Return 0
                End If

                If String.IsNullOrWhiteSpace(parsed.Username) OrElse
                   String.IsNullOrWhiteSpace(parsed.Password) Then
                    Console.WriteLine("Enter your Archive.org credentials below to configure 'ia'.")
                    Console.WriteLine()
                End If

                Dim username As String = parsed.Username
                Dim password As String = parsed.Password
                If String.IsNullOrWhiteSpace(username) Then
                    Console.Write("Email address: ")
                    username = Console.ReadLine()
                End If
                If String.IsNullOrWhiteSpace(password) Then
                    Console.Write("Password: ")
                    password = ReadPassword()
                    Console.WriteLine()
                End If

                Dim authConfig = ConfigService.GetAuthConfig(username, password, session.Host)
                ConfigService.WriteConfigFile(authConfig, session.ConfigFile)
                Dim savedMsg As String = String.Format(
                    "Config saved to: {0}", ConfigService.ResolveConfigFile(session.ConfigFile)
                )
                If String.IsNullOrWhiteSpace(parsed.Username) OrElse
                   String.IsNullOrWhiteSpace(parsed.Password) Then
                    savedMsg = Environment.NewLine & savedMsg
                End If
                Console.WriteLine(savedMsg)
                Return 0
            Catch ex As AuthenticationError
                Console.Error.WriteLine()
                Console.Error.WriteLine("error: " & ex.Message)
                Return 1
            Catch ex As Exception
                If parsed.UseNetrc Then
                    Console.Error.WriteLine("error: " & ex.Message)
                    Return 1
                End If
                Throw
            End Try
        End Function

        Private Shared Function DeepCopy(source As Dictionary(Of String, Dictionary(Of String, String))) As Dictionary(Of String, Dictionary(Of String, String))
            Dim result As New Dictionary(Of String, Dictionary(Of String, String))(
                StringComparer.OrdinalIgnoreCase
            )
            For Each section In source
                result(section.Key) = New Dictionary(Of String, String)(
                    section.Value, StringComparer.OrdinalIgnoreCase
                )
            Next
            Return result
        End Function

        Private Shared Function GetConfigValue(config As Dictionary(Of String, Dictionary(Of String, String)), section As String, key As String) As String
            Dim sectionMap As Dictionary(Of String, String) = Nothing
            If config IsNot Nothing AndAlso config.TryGetValue(section, sectionMap) Then
                Dim value As String = Nothing
                If sectionMap.TryGetValue(key, value) Then
                    Return value
                End If
            End If
            Return Nothing
        End Function

        Private Shared Function ParseArguments(args As IList(Of String)) As ConfigureArgs
            Dim parsed As New ConfigureArgs()
            Dim i As Integer = 0
            While i < args.Count
                Select Case args(i)
                    Case "--username", "-u"
                        i += 1
                        If i >= args.Count Then
                            Throw New ArgumentException("Missing value for --username")
                        End If
                        parsed.Username = args(i)
                    Case "--password", "-p"
                        i += 1
                        If i >= args.Count Then
                            Throw New ArgumentException("Missing value for --password")
                        End If
                        parsed.Password = args(i)
                    Case "--netrc", "-n"
                        parsed.UseNetrc = True
                    Case "--show", "-s"
                        parsed.Show = True
                    Case "--check", "-C"
                        parsed.Check = True
                    Case "--whoami", "-w"
                        parsed.WhoAmI = True
                    Case "--print-cookies", "-c"
                        parsed.PrintCookies = True
                    Case "--print-auth-header", "-a"
                        parsed.PrintAuthHeader = True
                    Case Else
                        Throw New ArgumentException(String.Format("Unknown option: {0}", args(i)))
                End Select
                i += 1
            End While

            Dim exclusiveCount As Integer = 0
            If parsed.Show Then exclusiveCount += 1
            If parsed.Check Then exclusiveCount += 1
            If parsed.WhoAmI Then exclusiveCount += 1
            If exclusiveCount > 1 Then
                Throw New ArgumentException("--show, --check, and --whoami are mutually exclusive")
            End If

            Return parsed
        End Function

        Private Shared Function ReadNetrcArchiveCredentials() As Tuple(Of String, String)
            Dim home As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim netrcPath As String = IO.Path.Combine(home, ".netrc")
            If Not IO.File.Exists(netrcPath) Then
                Throw New Exception(".netrc file not found.")
            End If

            Dim text As String = IO.File.ReadAllText(netrcPath)
            Dim tokens As String() = SplitNetrcTokens(text)
            For i As Integer = 0 To tokens.Length - 1
                If String.Equals(tokens(i), "machine", StringComparison.OrdinalIgnoreCase) AndAlso
                   i + 1 < tokens.Length AndAlso
                   String.Equals(tokens(i + 1), "archive.org", StringComparison.OrdinalIgnoreCase) Then
                    Dim login As String = ""
                    Dim password As String = ""
                    Dim j As Integer = i + 2
                    While j < tokens.Length AndAlso
                          Not String.Equals(tokens(j), "machine", StringComparison.OrdinalIgnoreCase)
                        If String.Equals(tokens(j), "login", StringComparison.OrdinalIgnoreCase) AndAlso
                           j + 1 < tokens.Length Then
                            login = tokens(j + 1)
                            j += 2
                            Continue While
                        End If
                        If String.Equals(tokens(j), "password", StringComparison.OrdinalIgnoreCase) AndAlso
                           j + 1 < tokens.Length Then
                            password = tokens(j + 1)
                            j += 2
                            Continue While
                        End If
                        j += 1
                    End While

                    If String.IsNullOrWhiteSpace(login) Then
                        Throw New Exception("netrc.netrc() cannot parse your .netrc file.")
                    End If
                    Return Tuple.Create(login, If(password, ""))
                End If
            Next

            Throw New Exception("netrc.netrc() cannot parse your .netrc file.")
        End Function

        Private Shared Function ReadPassword() As String
            Dim chars As New List(Of Char)()
            While True
                Dim key As ConsoleKeyInfo = Console.ReadKey(intercept:=True)
                If key.Key = ConsoleKey.Enter Then
                    Exit While
                End If
                If key.Key = ConsoleKey.Backspace Then
                    If chars.Count > 0 Then
                        chars.RemoveAt(chars.Count - 1)
                    End If
                    Continue While
                End If
                chars.Add(key.KeyChar)
            End While
            Return New String(chars.ToArray())
        End Function

        Private Shared Function SplitNetrcTokens(content As String) As String()
            Dim cleaned As String = content.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ")
            Return cleaned.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        End Function

        Private NotInheritable Class ConfigureArgs
            Public Property Check As Boolean
            Public Property Password As String
            Public Property PrintAuthHeader As Boolean
            Public Property PrintCookies As Boolean
            Public Property Show As Boolean
            Public Property UseNetrc As Boolean
            Public Property Username As String
            Public Property WhoAmI As Boolean
        End Class
    End Class
End Namespace

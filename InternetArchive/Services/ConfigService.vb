Imports System.IO
Imports System.Text
Imports System.Web.Script.Serialization
Imports InternetArchive.InternetArchiveCli.Exceptions

Namespace InternetArchiveCli.Services
    Public NotInheritable Class ConfigService
        Private Sub New()
        End Sub

        Public Shared Function GetAuthConfig(email As String, password As String, host As String) As Dictionary(Of String, Dictionary(Of String, String))
            Dim url As String = String.Format("https://{0}/services/xauthn/?op=login", host)
            Dim formPairs As New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)("email", email),
                New KeyValuePair(Of String, String)("password", password)
            }

            Using client As New Net.Http.HttpClient()
                Using content As New Net.Http.FormUrlEncodedContent(formPairs)
                    Using response As Net.Http.HttpResponseMessage = client.PostAsync(url, content).GetAwaiter().GetResult()
                        Dim body As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Threading.Thread.Sleep(2000)

                        Dim serializer As New JavaScriptSerializer()
                        Dim payload = serializer.Deserialize(Of Dictionary(Of String, Object))(body)
                        Dim successObj As Object = Nothing
                        If payload.TryGetValue("success", successObj) AndAlso Convert.ToBoolean(successObj) Then
                            Dim values = CType(payload("values"), Dictionary(Of String, Object))
                            Dim s3 = CType(values("s3"), Dictionary(Of String, Object))
                            Dim cookies = CType(values("cookies"), Dictionary(Of String, Object))

                            Return New Dictionary(Of String, Dictionary(Of String, String))(
                                StringComparer.OrdinalIgnoreCase
                            ) From {
                                {
                                    "s3",
                                    New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                                        {"access", Convert.ToString(s3("access"))},
                                        {"secret", Convert.ToString(s3("secret"))}
                                    }
                                },
                                {
                                    "cookies",
                                    New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                                        {"logged-in-user", Convert.ToString(cookies("logged-in-user"))},
                                        {"logged-in-sig", Convert.ToString(cookies("logged-in-sig"))}
                                    }
                                },
                                {
                                    "general",
                                    New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                                        {"screenname", Convert.ToString(values("screenname"))}
                                    }
                                }
                            }
                        End If

                        Dim message As String = "Authentication failed"
                        Dim valuesNode As Object = Nothing
                        If payload.TryGetValue("values", valuesNode) Then
                            Dim values = TryCast(valuesNode, Dictionary(Of String, Object))
                            If values IsNot Nothing AndAlso values.ContainsKey("reason") Then
                                message = Convert.ToString(values("reason"))
                            End If
                        ElseIf payload.ContainsKey("error") Then
                            message = Convert.ToString(payload("error"))
                        End If

                        If String.Equals(message, "account_not_found", StringComparison.Ordinal) Then
                            message = "Account not found, check your email and try again."
                        ElseIf String.Equals(message, "account_bad_password", StringComparison.Ordinal) Then
                            message = "Incorrect password, try again."
                        Else
                            message = String.Format("Authentication failed: {0}", message)
                        End If
                        Throw New AuthenticationError(message)
                    End Using
                End Using
            End Using
        End Function

        Public Shared Function LoadConfig(overrideConfig As Dictionary(Of String, Dictionary(Of String, String)), configFile As String) As Dictionary(Of String, Dictionary(Of String, String))
            Dim parsed = ParseConfigFilePath(configFile)
            Dim resolvedConfigFile As String = parsed.Item1
            Dim fileConfig = ReadIni(resolvedConfigFile)
            EnsureSections(fileConfig)

            Dim envAccess As String = Environment.GetEnvironmentVariable("IA_ACCESS_KEY_ID")
            Dim envSecret As String = Environment.GetEnvironmentVariable("IA_SECRET_ACCESS_KEY")
            Dim onlyOneSet As Boolean =
                (Not String.IsNullOrWhiteSpace(envAccess) AndAlso String.IsNullOrWhiteSpace(envSecret)) OrElse
                (String.IsNullOrWhiteSpace(envAccess) AndAlso Not String.IsNullOrWhiteSpace(envSecret))
            If onlyOneSet Then
                Throw New ArgumentException(
                    "Both IA_ACCESS_KEY_ID and IA_SECRET_ACCESS_KEY environment variables must be set together, or neither should be set."
                )
            End If

            If Not String.IsNullOrWhiteSpace(envAccess) AndAlso Not String.IsNullOrWhiteSpace(envSecret) Then
                fileConfig("s3")("access") = envAccess
                fileConfig("s3")("secret") = envSecret
            End If

            If overrideConfig IsNot Nothing Then
                For Each section In overrideConfig
                    If Not fileConfig.ContainsKey(section.Key) Then
                        fileConfig(section.Key) = New Dictionary(Of String, String)(
                            StringComparer.OrdinalIgnoreCase
                        )
                    End If
                    For Each kvp In section.Value
                        fileConfig(section.Key)(kvp.Key) = kvp.Value
                    Next
                Next
            End If

            Return fileConfig
        End Function

        Public Shared Function ParseConfigFilePath(configFile As String) As Tuple(Of String, Boolean)
            Dim isXdg As Boolean = False
            If String.IsNullOrWhiteSpace(configFile) Then
                Dim candidates As New List(Of String)()
                Dim envConfig As String = Environment.GetEnvironmentVariable("IA_CONFIG_FILE")
                If Not String.IsNullOrWhiteSpace(envConfig) Then
                    candidates.Add(envConfig)
                End If

                Dim xdgConfigHome As String = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME")
                If String.IsNullOrWhiteSpace(xdgConfigHome) OrElse Not Path.IsPathRooted(xdgConfigHome) Then
                    xdgConfigHome = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                        ".config"
                    )
                End If
                Dim xdgConfigFile As String = Path.Combine(
                    xdgConfigHome, "internetarchive", "ia.ini"
                )
                candidates.Add(xdgConfigFile)
                candidates.Add(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".config", "ia.ini"
                ))
                candidates.Add(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".ia"
                ))

                Dim selected As String = Nothing
                For Each candidate In candidates
                    If File.Exists(candidate) Then
                        selected = candidate
                        Exit For
                    End If
                Next

                If String.IsNullOrWhiteSpace(selected) Then
                    selected = If(String.IsNullOrWhiteSpace(envConfig), xdgConfigFile, envConfig)
                End If
                configFile = selected
                If String.Equals(configFile, xdgConfigFile, StringComparison.OrdinalIgnoreCase) Then
                    isXdg = True
                End If
            End If
            Return Tuple.Create(configFile, isXdg)
        End Function
        Public Shared Function ResolveConfigFile(configFile As String) As String
            Return ParseConfigFilePath(configFile).Item1
        End Function

        Public Shared Sub WriteConfigFile(authConfig As Dictionary(Of String, Dictionary(Of String, String)), configFile As String)
            Dim parsed = ParseConfigFilePath(configFile)
            Dim resolved As String = parsed.Item1
            Dim isXdg As Boolean = parsed.Item2
            Dim merged = ReadIni(resolved)
            EnsureSections(merged)

            Dim s3Section = GetSection(authConfig, "s3")
            merged("s3")("access") = GetValue(s3Section, "access")
            merged("s3")("secret") = GetValue(s3Section, "secret")

            Dim cookieSection = GetSection(authConfig, "cookies")
            merged("cookies")("logged-in-user") = GetValue(cookieSection, "logged-in-user")
            merged("cookies")("logged-in-sig") = GetValue(cookieSection, "logged-in-sig")

            Dim generalSection = GetSection(authConfig, "general")
            merged("general")("screenname") = GetValue(generalSection, "screenname")

            Dim configDirectory As String = Path.GetDirectoryName(resolved)
            If isXdg AndAlso Not Directory.Exists(configDirectory) Then
                Dim parentDirectory As String = Path.GetDirectoryName(configDirectory)
                If Not String.IsNullOrWhiteSpace(parentDirectory) Then
                    Directory.CreateDirectory(parentDirectory)
                End If
                Directory.CreateDirectory(configDirectory)
            ElseIf Not String.IsNullOrWhiteSpace(configDirectory) AndAlso
                   Not Directory.Exists(configDirectory) Then
                Directory.CreateDirectory(configDirectory)
            End If

            WriteIni(resolved, merged)
        End Sub
        Private Shared Sub EnsureSections(config As Dictionary(Of String, Dictionary(Of String, String)))
            If Not config.ContainsKey("s3") Then
                config("s3") = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            End If
            If Not config.ContainsKey("cookies") Then
                config("cookies") = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            End If
            If Not config.ContainsKey("general") Then
                config("general") = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            End If

            If Not config("s3").ContainsKey("access") Then
                config("s3")("access") = Nothing
            End If
            If Not config("s3").ContainsKey("secret") Then
                config("s3")("secret") = Nothing
            End If
            If Not config("cookies").ContainsKey("logged-in-user") Then
                config("cookies")("logged-in-user") = Nothing
            End If
            If Not config("cookies").ContainsKey("logged-in-sig") Then
                config("cookies")("logged-in-sig") = Nothing
            End If
            If Not config("general").ContainsKey("screenname") Then
                config("general")("screenname") = Nothing
            End If
        End Sub

        Private Shared Function GetSection(config As Dictionary(Of String, Dictionary(Of String, String)), name As String) As Dictionary(Of String, String)
            If config Is Nothing Then
                Return New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            End If
            Dim section As Dictionary(Of String, String) = Nothing
            If config.TryGetValue(name, section) Then
                Return section
            End If
            Return New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetValue(section As Dictionary(Of String, String), key As String) As String
            If section Is Nothing Then
                Return Nothing
            End If
            Dim value As String = Nothing
            If section.TryGetValue(key, value) Then
                Return value
            End If
            Return Nothing
        End Function
        Private Shared Function ReadIni(path As String) As Dictionary(Of String, Dictionary(Of String, String))
            Dim result As New Dictionary(Of String, Dictionary(Of String, String))(
                StringComparer.OrdinalIgnoreCase
            )
            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                Return result
            End If

            Dim currentSection As String = Nothing
            For Each rawLine In File.ReadAllLines(path)
                Dim line As String = rawLine.Trim()
                If line.Length = 0 OrElse line.StartsWith("#", StringComparison.Ordinal) OrElse
                   line.StartsWith(";", StringComparison.Ordinal) Then
                    Continue For
                End If

                If line.StartsWith("[", StringComparison.Ordinal) AndAlso
                   line.EndsWith("]", StringComparison.Ordinal) Then
                    currentSection = line.Substring(1, line.Length - 2).Trim()
                    If Not result.ContainsKey(currentSection) Then
                        result(currentSection) = New Dictionary(Of String, String)(
                            StringComparer.OrdinalIgnoreCase
                        )
                    End If
                    Continue For
                End If

                Dim sepIndex As Integer = line.IndexOf("="c)
                If sepIndex < 0 OrElse String.IsNullOrWhiteSpace(currentSection) Then
                    Continue For
                End If

                Dim key As String = line.Substring(0, sepIndex).Trim()
                Dim value As String = line.Substring(sepIndex + 1).Trim()
                result(currentSection)(key) = value
            Next
            Return result
        End Function

        Private Shared Sub WriteIni(path As String, config As Dictionary(Of String, Dictionary(Of String, String)))
            Dim sb As New StringBuilder()
            For Each sectionName In config.Keys
                sb.AppendLine(String.Format("[{0}]", sectionName))
                For Each kvp In config(sectionName)
                    sb.AppendLine(String.Format("{0} = {1}", kvp.Key, If(kvp.Value, "")))
                Next
                sb.AppendLine()
            Next
            File.WriteAllText(path, sb.ToString())
        End Sub
    End Class
End Namespace

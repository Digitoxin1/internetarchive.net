Imports System.Globalization
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Web.Script.Serialization

Namespace InternetArchiveCli.Core
    Public Class ArchiveSession
        Public Sub New(config As Dictionary(Of String, Dictionary(Of String, String)), configFile As String)
            Me.Config = config
            Me.ConfigFile = configFile

            Dim generalSection As Dictionary(Of String, String) = GetSection("general")
            Dim secureRaw As String = GetValue(generalSection, "secure")
            Dim secure As Boolean = True
            If Not String.IsNullOrWhiteSpace(secureRaw) Then
                Boolean.TryParse(secureRaw, secure)
            End If

            Dim hostValue As String = GetValue(generalSection, "host")
            If String.IsNullOrWhiteSpace(hostValue) Then
                hostValue = "archive.org"
            End If
            If hostValue.IndexOf("archive.org", StringComparison.OrdinalIgnoreCase) < 0 Then
                hostValue &= ".archive.org"
            End If

            Host = hostValue
            Protocol = If(secure, "https:", "http:")

            Dim s3Section As Dictionary(Of String, String) = GetSection("s3")
            AccessKey = GetValue(s3Section, "access")
            SecretKey = GetValue(s3Section, "secret")

            Dim cookiesSection As Dictionary(Of String, String) = GetSection("cookies")
            Dim loggedInUser As String = GetValue(cookiesSection, "logged-in-user")
            If Not String.IsNullOrWhiteSpace(loggedInUser) Then
                Dim cookieEmail As String = loggedInUser
                Dim semiIdx As Integer = cookieEmail.IndexOf(";"c)
                If semiIdx >= 0 Then
                    cookieEmail = cookieEmail.Substring(0, semiIdx)
                End If
                UserEmail = Uri.UnescapeDataString(cookieEmail)
            Else
                UserEmail = Nothing
            End If
        End Sub

        Public ReadOnly Property AccessKey As String
        Public ReadOnly Property Config As Dictionary(Of String, Dictionary(Of String, String))
        Public ReadOnly Property ConfigFile As String
        Public ReadOnly Property Host As String
        Public ReadOnly Property Protocol As String
        Public ReadOnly Property SecretKey As String
        Public ReadOnly Property UserEmail As String

        Public Function AddFlag(identifier As String, category As String, user As String) As ApiCallResult
            Dim url As String = String.Format(
                "{0}//{1}/services/flags/admin.php",
                Protocol,
                Host
            )
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier},
                {"category", category},
                {"user", user}
            }
            Return SendFlagRequest(HttpMethod.Put, url, params)
        End Function

        Public Function CopyS3Object(destination As String,
                                     headers As Dictionary(Of String, String),
                                     itemMetadata As Dictionary(Of String, Object),
                                     fileMetadata As Dictionary(Of String, Object),
                                     queueDerive As Boolean) As ApiCallResult

            Dim destinationEncoded As String = ApiShared.EscapeArchivePath(destination)
            Dim url As String = String.Format(
                "{0}//s3.us.archive.org/{1}",
                Protocol,
                destinationEncoded
            )

            Dim requestHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If headers IsNot Nothing Then
                For Each kvp In headers
                    requestHeaders(kvp.Key) = kvp.Value
                Next
            End If
            If Not requestHeaders.ContainsKey("x-archive-auto-make-bucket") Then
                requestHeaders("x-archive-auto-make-bucket") = "1"
            End If
            requestHeaders("x-archive-queue-derive") = If(queueDerive, "1", "0")

            AddArchiveMetadataHeaders(requestHeaders, itemMetadata, "meta")
            AddArchiveMetadataHeaders(requestHeaders, fileMetadata, "filemeta")

            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Put, url)
                    AddS3AuthHeader(request)
                    For Each kvp In requestHeaders
                        request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                    Next

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object) = Nothing
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("copy S3 object response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function DeleteFlag(identifier As String, category As String, user As String) As ApiCallResult
            Dim url As String = String.Format(
                "{0}//{1}/services/flags/admin.php",
                Protocol,
                Host
            )
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier},
                {"category", category},
                {"user", user}
            }
            Return SendFlagRequest(HttpMethod.Delete, url, params)
        End Function

        Public Function DeleteReview(identifier As String, username As String, screenname As String, itemname As String) As ApiCallResult
            Dim payload As Dictionary(Of String, String) = Nothing
            If Not String.IsNullOrWhiteSpace(username) Then
                payload = New Dictionary(Of String, String) From {{"username", username}}
            ElseIf Not String.IsNullOrWhiteSpace(screenname) Then
                payload = New Dictionary(Of String, String) From {{"screenname", screenname}}
            ElseIf Not String.IsNullOrWhiteSpace(itemname) Then
                payload = New Dictionary(Of String, String) From {{"itemname", itemname}}
            End If
            Dim url As String = String.Format("{0}//{1}/services/reviews.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendReviewRequest(HttpMethod.Delete, url, params, payload, Nothing, True)
        End Function

        Public Function DeleteS3File(identifier As String, fileName As String, headers As Dictionary(Of String, String), retries As Integer) As ApiCallResult
            Dim encodedPath As String = ApiShared.EscapeArchivePath(identifier & "/" & fileName)
            Dim url As String = String.Format("{0}//s3.us.archive.org/{1}", Protocol, encodedPath)

            Dim attempts As Integer = Math.Max(0, retries) + 1
            Dim lastResult As ApiCallResult = Nothing
            Using client As New HttpClient()
                Dim serializer As New JavaScriptSerializer()
                For attempt As Integer = 1 To attempts
                    Using request As New HttpRequestMessage(HttpMethod.Delete, url)
                        AddS3AuthHeader(request)
                        If headers IsNot Nothing Then
                            For Each kvp In headers
                                request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                            Next
                        End If

                        Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                            Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                            Dim body As Dictionary(Of String, Object) = Nothing
                            Try
                                body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                            Catch ex As Exception
                                WarnJsonParseFallback("delete S3 file response", ex)
                                body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                            End Try
                            lastResult = New ApiCallResult With {
                                .StatusCode = CInt(response.StatusCode),
                                .Text = text,
                                .JsonBody = body
                            }
                        End Using
                    End Using

                    If lastResult.StatusCode <> 503 Then
                        Return lastResult
                    End If
                Next
            End Using

            Return lastResult
        End Function

        Public Function GetFlags(identifier As String) As ApiCallResult
            Dim url As String = String.Format(
                "{0}//{1}/services/flags/admin.php",
                Protocol,
                Host
            )
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendFlagRequest(HttpMethod.Get, url, params)
        End Function

        Public Function GetItemMetadata(identifier As String, params As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            Dim normalizedIdentifier As String = ApiShared.EscapeArchivePath(identifier)
            Dim baseUrl As String = String.Format("{0}//{1}/metadata/{2}", Protocol, Host, normalizedIdentifier)
            Dim url As String = BuildUrl(baseUrl, params)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Get, url)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim content As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim result = serializer.DeserializeObject(content)
                        Dim asDict = TryCast(result, Dictionary(Of String, Object))
                        If asDict Is Nothing Then
                            Return New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End If
                        Return asDict
                    End Using
                End Using
            End Using
        End Function

        Public Function GetReview(identifier As String) As ApiCallResult
            Dim url As String = String.Format("{0}//{1}/services/reviews.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendReviewRequest(HttpMethod.Get, url, params, Nothing, Nothing, False)
        End Function

        Public Function GetTaskLog(taskId As String) As String
            Dim hostForLog As String = Host
            If String.Equals(Host, "archive.org", StringComparison.OrdinalIgnoreCase) Then
                hostForLog = "catalogd.archive.org"
            End If
            Dim url As String = String.Format("{0}//{1}/services/tasks.php", Protocol, hostForLog)
            Dim params As New Dictionary(Of String, Object) From {
                {"task_log", taskId}
            }
            Dim finalUrl As String = BuildUrl(url, params)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Get, finalUrl)
                    AddS3AuthHeader(request)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim rawBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
                        If CInt(response.StatusCode) >= 400 Then
                            Throw New InvalidOperationException(String.Format(
                                "HTTP {0} while retrieving task log.",
                                CInt(response.StatusCode)
                            ))
                        End If
                        Return Encoding.UTF8.GetString(rawBytes)
                    End Using
                End Using
            End Using
        End Function

        Public Function GetTasks(params As Dictionary(Of String, Object)) As List(Of Dictionary(Of String, Object))
            Dim merged As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If params IsNot Nothing Then
                For Each kvp In params
                    merged(kvp.Key) = kvp.Value
                Next
            End If
            If Not merged.ContainsKey("history") Then
                merged("history") = 1
            End If
            If Not merged.ContainsKey("catalog") Then
                merged("catalog") = 1
            End If
            If Not merged.ContainsKey("limit") Then
                merged("limit") = 0
            End If
            If Not merged.ContainsKey("summary") Then
                merged("summary") = 0
            End If

            Dim url As String = String.Format("{0}//{1}/services/tasks.php", Protocol, Host)
            Dim finalUrl As String = BuildUrl(url, merged)

            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Get, finalUrl)
                    AddS3AuthHeader(request)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim tasks As New List(Of Dictionary(Of String, Object))()
                        Dim serializer As New JavaScriptSerializer()
                        Dim skippedNonJsonLines As Integer = 0
                        If CInt(response.StatusCode) >= 400 Then
                            Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                            ThrowForTasksError(response, text)
                        End If

                        Using stream As IO.Stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
                            Using reader As New IO.StreamReader(stream, Encoding.UTF8)
                                While True
                                    Dim line As String = reader.ReadLine()
                                    If line Is Nothing Then
                                        Exit While
                                    End If

                                    Dim trimmed As String = line.Trim()
                                    If trimmed.Length = 0 Then
                                        Continue While
                                    End If

                                    Try
                                        Dim obj = serializer.DeserializeObject(trimmed)
                                        Dim dict = TryCast(obj, Dictionary(Of String, Object))
                                        If dict IsNot Nothing Then
                                            tasks.Add(dict)
                                        End If
                                    Catch
                                        skippedNonJsonLines += 1
                                    End Try
                                End While
                            End Using
                        End Using

                        If skippedNonJsonLines > 0 Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "warning: skipped {0} non-JSON line(s) while parsing tasks response.",
                                    skippedNonJsonLines
                                )
                            )
                        End If

                        Dim invalidSubmitTimeCount As Integer = 0
                        For Each task In tasks
                            Dim parsedSubmitTime As DateTime
                            If Not TryGetSubmitTime(task, parsedSubmitTime) Then
                                invalidSubmitTimeCount += 1
                            End If
                        Next
                        If invalidSubmitTimeCount > 0 Then
                            Console.Error.WriteLine(
                                String.Format(
                                    "warning: {0} task(s) had missing or unparseable submittime values.",
                                    invalidSubmitTimeCount
                                )
                            )
                        End If

                        tasks.Sort(AddressOf CompareTaskByDateDesc)
                        Return tasks
                    End Using
                End Using
            End Using
        End Function

        Public Function GetTasksApiRateLimit(cmd As String) As Dictionary(Of String, Object)
            Dim url As String = String.Format("{0}//{1}/services/tasks.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"rate_limits", 1},
                {"cmd", cmd}
            }
            Dim finalUrl As String = BuildUrl(url, params)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Get, finalUrl)
                    AddS3AuthHeader(request)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        ThrowForTasksError(response, text)
                        Dim serializer As New JavaScriptSerializer()
                        Return serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                    End Using
                End Using
            End Using
        End Function

        Public Function IndexReview(identifier As String, username As String, screenname As String, itemname As String) As ApiCallResult
            Dim payload As New Dictionary(Of String, String) From {
                {"noindex", "0"}
            }
            If Not String.IsNullOrWhiteSpace(username) Then
                payload("username") = username
            ElseIf Not String.IsNullOrWhiteSpace(screenname) Then
                payload("screenname") = screenname
            ElseIf Not String.IsNullOrWhiteSpace(itemname) Then
                payload("itemname") = itemname
            End If
            Dim url As String = String.Format("{0}//{1}/services/reviews.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendReviewRequest(HttpMethod.Put, url, params, payload, Nothing, True)
        End Function

        Public Function LockUnlockAccount(itemname As String, isLock As Boolean, comment As String) As ApiCallResult
            Dim url As String = String.Format("{0}//{1}/services/xauthn/", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"op", "lock_unlock"}
            }
            Dim form As New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)("itemname", itemname),
                New KeyValuePair(Of String, String)("is_lock", If(isLock, "1", "0"))
            }
            If Not String.IsNullOrWhiteSpace(comment) Then
                form.Add(New KeyValuePair(Of String, String)("comments", comment))
            End If

            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Post, BuildUrl(url, params))
                    AddS3AuthHeader(request)
                    request.Content = New FormUrlEncodedContent(form)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object)
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("lock/unlock account response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        If CInt(response.StatusCode) >= 400 Then
                            Throw New InvalidOperationException(String.Format(
                                "HTTP {0} while performing account lock/unlock.",
                                CInt(response.StatusCode)
                            ))
                        End If
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function LookupAccount(identifierType As String, identifier As String) As ApiCallResult
            Dim url As String = String.Format("{0}//{1}/services/xauthn/", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"op", "info"}
            }
            Dim form As New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)(identifierType, identifier)
            }

            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Post, BuildUrl(url, params))
                    request.Content = New FormUrlEncodedContent(form)
                    request.Headers.TryAddWithoutValidation("Content-Type", "application/x-www-form-urlencoded")
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object)
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("lookup account response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function NoIndexReview(identifier As String, username As String, screenname As String, itemname As String) As ApiCallResult
            Dim payload As New Dictionary(Of String, String) From {
                {"noindex", "1"}
            }
            If Not String.IsNullOrWhiteSpace(username) Then
                payload("username") = username
            ElseIf Not String.IsNullOrWhiteSpace(screenname) Then
                payload("screenname") = screenname
            ElseIf Not String.IsNullOrWhiteSpace(itemname) Then
                payload("itemname") = itemname
            End If
            Dim url As String = String.Format("{0}//{1}/services/reviews.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendReviewRequest(HttpMethod.Put, url, params, payload, Nothing, True)
        End Function

        Public Function PostMetadataPatch(identifier As String,
                                          target As String,
                                          patchJson As String,
                                          priority As Integer,
                                          headers As Dictionary(Of String, String),
                                          reducedPriority As Boolean,
                                          timeoutSeconds As Double) As MetadataWriteResult

            Dim normalizedIdentifier As String = ApiShared.EscapeArchivePath(identifier)
            Dim url As String = String.Format("{0}//{1}/metadata/{2}", Protocol, Host, normalizedIdentifier)
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(timeoutSeconds)
                Using request As New HttpRequestMessage(HttpMethod.Post, url)
                    ApiShared.ApplyLowAuth(request, AccessKey, SecretKey)
                    If reducedPriority Then
                        request.Headers.TryAddWithoutValidation("X-Accept-Reduced-Priority", "1")
                    End If
                    If headers IsNot Nothing Then
                        For Each kvp In headers
                            request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                        Next
                    End If

                    Dim form As New List(Of KeyValuePair(Of String, String)) From {
                        New KeyValuePair(Of String, String)("-patch", patchJson),
                        New KeyValuePair(Of String, String)("-target", target),
                        New KeyValuePair(Of String, String)("priority", priority.ToString())
                    }
                    request.Content = New FormUrlEncodedContent(form)

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim responseText As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object) = Nothing
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(responseText)
                        Catch ex As Exception
                            WarnJsonParseFallback("metadata patch response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try

                        Return New MetadataWriteResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = responseText,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function PostSimplelistPatch(identifier As String, patch As Dictionary(Of String, Object)) As ApiCallResult
            Dim normalizedIdentifier As String = ApiShared.EscapeArchivePath(identifier)
            Dim url As String = String.Format(
                "{0}//{1}/metadata/{2}",
                Protocol,
                Host,
                normalizedIdentifier
            )
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Post, url)
                    AddS3AuthHeader(request)

                    Dim serializer As New JavaScriptSerializer()
                    Dim form As New List(Of KeyValuePair(Of String, String)) From {
                        New KeyValuePair(Of String, String)("-patch", serializer.Serialize(patch)),
                        New KeyValuePair(Of String, String)("-target", "simplelists")
                    }
                    request.Content = New FormUrlEncodedContent(form)

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim body As Dictionary(Of String, Object) = Nothing
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("simplelists patch response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function S3IsOverloaded(identifier As String) As Boolean
            Dim url As String = String.Format("{0}//s3.us.archive.org", Protocol)
            Dim normalizedIdentifier As String = ApiShared.NormalizeArchivePath(identifier)
            Dim params As New Dictionary(Of String, Object) From {
                {"check_limit", 1},
                {"accesskey", AccessKey},
                {"bucket", normalizedIdentifier}
            }
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(12)
                Using request As New HttpRequestMessage(HttpMethod.Get, BuildUrl(url, params))
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Try
                            Dim j = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                            If j IsNot Nothing AndAlso j.ContainsKey("over_limit") Then
                                Return Convert.ToInt32(j("over_limit"), CultureInfo.InvariantCulture) <> 0
                            End If
                            Console.Error.WriteLine(
                                String.Format(
                                    "warning: unable to determine overload status for '{0}' from S3 response; treating as overloaded.",
                                    normalizedIdentifier
                                )
                            )
                        Catch ex As Exception
                            Console.Error.WriteLine(
                                String.Format(
                                    "warning: failed to parse S3 overload status for '{0}': {1}; treating as overloaded.",
                                    normalizedIdentifier,
                                    ex.Message
                                )
                            )
                            Return True
                        End Try
                        Return True
                    End Using
                End Using
            End Using
        End Function

        Public Function SearchItemsSimple(query As String, extraParams As Dictionary(Of String, Object)) As List(Of Dictionary(Of String, Object))
            Dim params As New Dictionary(Of String, Object) From {
                {"q", query},
                {"count", "10000"}
            }
            If extraParams IsNot Nothing Then
                For Each kvp In extraParams
                    params(kvp.Key) = kvp.Value
                Next
            End If

            Dim url As String = String.Format(
                "{0}//{1}/services/search/v1/scrape",
                Protocol,
                Host
            )

            Dim results As New List(Of Dictionary(Of String, Object))()
            Using client As New HttpClient()
                Dim serializer As New JavaScriptSerializer()
                While True
                    Dim finalUrl As String = BuildUrl(url, params)
                    Using request As New HttpRequestMessage(HttpMethod.Post, finalUrl)
                        AddS3AuthHeader(request)
                        Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                            Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

                            Dim payload = serializer.DeserializeObject(text)
                            Dim json = TryCast(payload, Dictionary(Of String, Object))
                            If json Is Nothing Then
                                Return results
                            End If
                            If json.ContainsKey("error") Then
                                results.Add(json)
                                Return results
                            End If

                            Dim itemsObj As Object = Nothing
                            If json.TryGetValue("items", itemsObj) Then
                                For Each entry In ApiShared.ToObjectSequence(itemsObj)
                                    Dim dict = TryCast(entry, Dictionary(Of String, Object))
                                    If dict IsNot Nothing Then
                                        results.Add(dict)
                                    End If
                                Next
                            End If

                            If json.ContainsKey("cursor") Then
                                params("cursor") = Convert.ToString(json("cursor"), CultureInfo.InvariantCulture)
                            Else
                                Exit While
                            End If
                        End Using
                    End Using
                End While
            End Using

            Return results
        End Function

        Public Function SubmitReview(identifier As String, title As String, body As String, stars As Integer?) As ApiCallResult
            Dim payload As New Dictionary(Of String, Object) From {
                {"title", title},
                {"body", body}
            }
            If stars.HasValue Then
                payload("stars") = stars.Value
            End If
            Dim url As String = String.Format("{0}//{1}/services/reviews.php", Protocol, Host)
            Dim params As New Dictionary(Of String, Object) From {
                {"identifier", identifier}
            }
            Return SendReviewRequest(HttpMethod.Post, url, params, Nothing, payload, True)
        End Function

        Public Function SubmitTask(identifier As String,
                                   cmd As String,
                                   comment As String,
                                   priority As Integer,
                                   data As Dictionary(Of String, Object),
                                   reducedPriority As Boolean) As ApiCallResult

            Dim payload As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If data IsNot Nothing Then
                For Each kvp In data
                    payload(kvp.Key) = kvp.Value
                Next
            End If

            payload("cmd") = cmd
            payload("identifier") = identifier
            If Not String.IsNullOrWhiteSpace(comment) Then
                Dim argsObj As Object = Nothing
                If payload.TryGetValue("args", argsObj) Then
                    Dim argsDict = If(TryCast(argsObj, Dictionary(Of String, Object)), New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase))
                    argsDict("comment") = comment
                    payload("args") = argsDict
                Else
                    payload("args") = New Dictionary(Of String, Object) From {{"comment", comment}}
                End If
            End If
            If priority <> 0 Then
                payload("priority") = priority
            End If

            Dim url As String = String.Format("{0}//{1}/services/tasks.php", Protocol, Host)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(HttpMethod.Post, url)
                    AddS3AuthHeader(request)
                    If reducedPriority Then
                        request.Headers.TryAddWithoutValidation("X-Accept-Reduced-Priority", "1")
                    End If
                    Dim serializer As New JavaScriptSerializer()
                    Dim jsonBody As String = serializer.Serialize(payload)
                    request.Content = New StringContent(jsonBody, Encoding.UTF8, "application/json")

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim parsed As Dictionary(Of String, Object)
                        Try
                            parsed = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("submit task response", ex)
                            parsed = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try

                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = parsed
                        }
                    End Using
                End Using
            End Using
        End Function

        Public Function UploadS3File(identifier As String,
                                     remoteName As String,
                                     content As Byte(),
                                     Headers As Dictionary(Of String, String),
                                     itemMetadata As Dictionary(Of String, Object),
                                     fileMetadata As Dictionary(Of String, Object),
                                     queueDerive As Boolean,
                                     retries As Integer,
                                     retriesSleepSeconds As Integer,
                                     Debug As Boolean,
                                     Optional progressCallback As Action(Of Long, Long) = Nothing) As ApiCallResult

            Dim remotePath As String = identifier & "/" & remoteName
            Dim encodedPath As String = ApiShared.EscapeArchivePath(remotePath)
            Dim url As String = String.Format("{0}//s3.us.archive.org/{1}", Protocol, encodedPath)

            Dim requestHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If Headers IsNot Nothing Then
                For Each kvp In Headers
                    requestHeaders(kvp.Key) = kvp.Value
                Next
            End If
            If Not requestHeaders.ContainsKey("x-archive-auto-make-bucket") Then
                requestHeaders("x-archive-auto-make-bucket") = "1"
            End If
            If Not requestHeaders.ContainsKey("x-archive-size-hint") Then
                requestHeaders("x-archive-size-hint") =
                    content.Length.ToString(CultureInfo.InvariantCulture)
            End If
            requestHeaders("x-archive-queue-derive") = If(queueDerive, "1", "0")
            AddArchiveMetadataHeaders(requestHeaders, itemMetadata, "meta")
            AddArchiveMetadataHeaders(requestHeaders, fileMetadata, "filemeta")

            If Debug Then
                Dim debugHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Using request As New HttpRequestMessage(HttpMethod.Put, url)
                    request.Content = New ByteArrayContent(content)
                    request.Content.Headers.ContentLength = content.Length
                    AddS3AuthHeader(request)
                    For Each kvp In requestHeaders
                        If Not request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value) Then
                            request.Content.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                        End If
                    Next
                    For Each h In request.Headers
                        debugHeaders(h.Key) = String.Join(",", h.Value)
                    Next
                    For Each h In request.Content.Headers
                        debugHeaders(h.Key) = String.Join(",", h.Value)
                    Next
                End Using
                Return New ApiCallResult With {
                    .StatusCode = -1,
                    .Text = "",
                    .JsonBody = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase),
                    .RequestUrl = url,
                    .Headers = debugHeaders,
                    .WasSkipped = False
                }
            End If

            Dim attempts As Integer = Math.Max(1, If(retries > 0, retries, 1))
            Dim last As ApiCallResult = Nothing
            Using client As New HttpClient()
                Dim serializer As New JavaScriptSerializer()
                For attempt As Integer = 1 To attempts
                    Using request As New HttpRequestMessage(HttpMethod.Put, url)
                        If progressCallback Is Nothing Then
                            request.Content = New ByteArrayContent(content)
                        Else
                            request.Content = New ProgressByteArrayContent(content, progressCallback)
                        End If
                        request.Content.Headers.ContentLength = content.Length
                        AddS3AuthHeader(request)
                        For Each kvp In requestHeaders
                            If Not request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value) Then
                                request.Content.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                            End If
                        Next

                        Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                            Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                            Dim body As Dictionary(Of String, Object) = Nothing
                            Try
                                body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                            Catch ex As Exception
                                WarnJsonParseFallback("upload S3 file response", ex)
                                body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                            End Try
                            Dim responseHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                            For Each h In response.Headers
                                responseHeaders(h.Key) = String.Join(",", h.Value)
                            Next
                            For Each h In response.Content.Headers
                                responseHeaders(h.Key) = String.Join(",", h.Value)
                            Next

                            last = New ApiCallResult With {
                                .StatusCode = CInt(response.StatusCode),
                                .Text = text,
                                .JsonBody = body,
                                .RequestUrl = url,
                                .Headers = responseHeaders,
                                .WasSkipped = False
                            }
                        End Using
                    End Using

                    If last.StatusCode = 503 Then
                        Threading.Thread.Sleep(Math.Max(0, retriesSleepSeconds) * 1000)
                        Continue For
                    End If
                    Return last
                Next
            End Using

            Return last
        End Function

        Public Function WhoAmI() As Dictionary(Of String, Object)
            Using client As New HttpClient()
                Dim requestUri As String = "https://archive.org/services/user.php?op=whoami"
                Using request As New HttpRequestMessage(HttpMethod.Get, requestUri)
                    ApiShared.ApplyLowAuth(request, AccessKey, SecretKey)

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim content As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Return serializer.Deserialize(Of Dictionary(Of String, Object))(content)
                    End Using
                End Using
            End Using
        End Function

        Private Shared Sub AddArchiveMetadataHeaders(headers As Dictionary(Of String, String), metadata As Dictionary(Of String, Object), metaType As String)
            If metadata Is Nothing Then
                Return
            End If
            Dim idxMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each kvp In metadata
                Dim values As List(Of String) = ToMetadataValueList(kvp.Value)
                If values.Count = 0 Then
                    Continue For
                End If
                If Not idxMap.ContainsKey(kvp.Key) Then
                    idxMap(kvp.Key) = 0
                End If
                For Each value In values
                    If String.IsNullOrWhiteSpace(value) Then
                        Continue For
                    End If
                    Dim idx As Integer = idxMap(kvp.Key)
                    Dim headerKey As String = String.Format(
                        "x-archive-{0}{1:00}-{2}",
                        metaType,
                        idx,
                        kvp.Key.Replace("_", "--")
                    )
                    Dim headerValue As String = value
                    If NeedsQuote(value) Then
                        headerValue = "uri(" & Uri.EscapeDataString(value) & ")"
                    End If
                    headers(headerKey) = headerValue
                    idxMap(kvp.Key) = idx + 1
                Next
            Next
        End Sub

        Private Shared Function BuildUrl(baseUrl As String, params As Dictionary(Of String, Object)) As String
            If params Is Nothing OrElse params.Count = 0 Then
                Return baseUrl
            End If
            Dim parts As New List(Of String)()
            For Each kvp In params
                parts.Add(Uri.EscapeDataString(kvp.Key) & "=" & Uri.EscapeDataString(Convert.ToString(kvp.Value)))
            Next
            Return baseUrl & "?" & String.Join("&", parts)
        End Function

        Private Shared Function CompareTaskByDateDesc(a As Dictionary(Of String, Object), b As Dictionary(Of String, Object)) As Integer
            Dim da As DateTime = ParseSubmitTime(a)
            Dim db As DateTime = ParseSubmitTime(b)
            Return db.CompareTo(da)
        End Function

        Private Shared Function GetSection(config As Dictionary(Of String, Dictionary(Of String, String)), name As String) As Dictionary(Of String, String)
            Dim value As Dictionary(Of String, String) = Nothing
            If config IsNot Nothing AndAlso config.TryGetValue(name, value) Then
                Return value
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

        Private Shared Function NeedsQuote(value As String) As Boolean
            For Each c As Char In value
                If AscW(c) > 127 Then
                    Return True
                End If
                If Char.IsWhiteSpace(c) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Function ParseSubmitTime(task As Dictionary(Of String, Object)) As DateTime
            Dim dt As DateTime
            If TryGetSubmitTime(task, dt) Then
                Return dt
            End If
            Return DateTime.MinValue
        End Function

        Private Shared Function TryGetSubmitTime(task As Dictionary(Of String, Object), ByRef submitTime As DateTime) As Boolean
            submitTime = DateTime.MinValue
            If task Is Nothing Then
                Return False
            End If
            If task.ContainsKey("category") AndAlso
               String.Equals(Convert.ToString(task("category")), "summary", StringComparison.OrdinalIgnoreCase) Then
                submitTime = DateTime.Now
                Return True
            End If
            If Not task.ContainsKey("submittime") Then
                Return False
            End If
            Dim raw As String = Convert.ToString(task("submittime"), CultureInfo.InvariantCulture)
            If DateTime.TryParseExact(
                raw,
                "yyyy-MM-dd HH:mm:ss.ffffff",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                submitTime
            ) Then
                Return True
            End If
            If DateTime.TryParseExact(
                raw,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                submitTime
            ) Then
                Return True
            End If
            Return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, submitTime)
        End Function

        Private Shared Sub ThrowForTasksError(response As HttpResponseMessage, text As String)
            If CInt(response.StatusCode) < 400 Then
                Return
            End If
            Dim serializer As New JavaScriptSerializer()
            Try
                Dim obj = serializer.DeserializeObject(text)
                Dim dict = TryCast(obj, Dictionary(Of String, Object))
                If dict IsNot Nothing AndAlso dict.ContainsKey("error") Then
                    Throw New InvalidOperationException(Convert.ToString(dict("error"), CultureInfo.InvariantCulture))
                End If
            Catch ex As InvalidOperationException
                Throw
            Catch ex As Exception
                Console.Error.WriteLine(
                    String.Format(
                        "warning: unable to parse tasks error response body: {0}",
                        ex.Message
                    )
                )
            End Try
            Throw New InvalidOperationException(String.Format(
                "HTTP {0} from Tasks API.",
                CInt(response.StatusCode)
            ))
        End Sub

        Private Shared Function ToMetadataValueList(value As Object) As List(Of String)
            Dim result As New List(Of String)()
            If value Is Nothing Then
                Return result
            End If
            Dim arr = TryCast(value, ArrayList)
            If arr IsNot Nothing Then
                For Each entry In arr
                    result.Add(Convert.ToString(entry, CultureInfo.InvariantCulture))
                Next
                Return result
            End If
            Dim objArray = TryCast(value, Object())
            If objArray IsNot Nothing Then
                For Each entry In objArray
                    result.Add(Convert.ToString(entry, CultureInfo.InvariantCulture))
                Next
                Return result
            End If
            Dim strList = TryCast(value, List(Of String))
            If strList IsNot Nothing Then
                result.AddRange(strList)
                Return result
            End If
            result.Add(Convert.ToString(value, CultureInfo.InvariantCulture))
            Return result
        End Function

        Private Shared Sub WarnJsonParseFallback(context As String, ex As Exception)
            Console.Error.WriteLine(
                String.Format(
                    "warning: failed to parse {0} as JSON dictionary: {1}",
                    context,
                    ex.Message
                )
            )
        End Sub

        Private Sub AddS3AuthHeader(request As HttpRequestMessage)
            ApiShared.ApplyLowAuth(request, AccessKey, SecretKey)
        End Sub

        Private Function GetSection(name As String) As Dictionary(Of String, String)
            Return GetSection(Config, name)
        End Function

        Private Function SendFlagRequest(method As HttpMethod, url As String, params As Dictionary(Of String, Object)) As ApiCallResult
            Dim finalUrl As String = BuildUrl(url, params)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(method, finalUrl)
                    request.Headers.TryAddWithoutValidation("Accept", "text/json")
                    AddS3AuthHeader(request)
                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        Dim serializer As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object) = Nothing
                        Try
                            body = serializer.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("flag request response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Private Function SendReviewRequest(method As HttpMethod,
                                           url As String,
                                           params As Dictionary(Of String, Object),
                                           formData As Dictionary(Of String, String),
                                           jsonData As Dictionary(Of String, Object),
                                           throwOnHttpError As Boolean) As ApiCallResult

            Dim finalUrl As String = BuildUrl(url, params)
            Using client As New HttpClient()
                Using request As New HttpRequestMessage(method, finalUrl)
                    AddS3AuthHeader(request)

                    If jsonData IsNot Nothing Then
                        Dim serializer As New JavaScriptSerializer()
                        request.Content = New StringContent(
                            serializer.Serialize(jsonData),
                            Encoding.UTF8,
                            "application/json"
                        )
                    ElseIf formData IsNot Nothing Then
                        Dim pairs As New List(Of KeyValuePair(Of String, String))()
                        For Each kvp In formData
                            pairs.Add(New KeyValuePair(Of String, String)(kvp.Key, kvp.Value))
                        Next
                        request.Content = New FormUrlEncodedContent(pairs)
                    End If

                    Using response As HttpResponseMessage = client.SendAsync(request).GetAwaiter().GetResult()
                        Dim text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        If throwOnHttpError AndAlso CInt(response.StatusCode) >= 400 Then
                            Throw New InvalidOperationException(String.Format(
                                "HTTP {0} while calling reviews API.",
                                CInt(response.StatusCode)
                            ))
                        End If

                        Dim serializer2 As New JavaScriptSerializer()
                        Dim body As Dictionary(Of String, Object) = Nothing
                        Try
                            body = serializer2.Deserialize(Of Dictionary(Of String, Object))(text)
                        Catch ex As Exception
                            WarnJsonParseFallback("review request response", ex)
                            body = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        End Try
                        Return New ApiCallResult With {
                            .StatusCode = CInt(response.StatusCode),
                            .Text = text,
                            .JsonBody = body
                        }
                    End Using
                End Using
            End Using
        End Function

        Private NotInheritable Class ProgressByteArrayContent
            Inherits HttpContent

            Private ReadOnly _chunkSize As Integer
            Private ReadOnly _content As Byte()
            Private ReadOnly _progress As Action(Of Long, Long)
            Public Sub New(content As Byte(), progress As Action(Of Long, Long))
                _content = If(content, Array.Empty(Of Byte)())
                _progress = progress
                _chunkSize = 1024 * 1024
            End Sub

            Protected Overrides Async Function SerializeToStreamAsync(stream As IO.Stream, context As TransportContext) As Task
                Dim total As Long = _content.LongLength
                Dim sent As Long = 0

                If total = 0 Then
                    _progress?.Invoke(0, 0)
                    Return
                End If

                While sent < total
                    Dim toWrite As Integer = CInt(Math.Min(_chunkSize, total - sent))
                    Await stream.WriteAsync(_content, CInt(sent), toWrite).ConfigureAwait(False)
                    sent += toWrite
                    _progress?.Invoke(sent, total)
                End While
            End Function

            Protected Overrides Function TryComputeLength(ByRef length As Long) As Boolean
                length = _content.LongLength
                Return True
            End Function
        End Class
    End Class
End Namespace

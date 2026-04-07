Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class MoveCommand
        Private Sub New()
        End Sub

        Public Shared Function Execute(session As ArchiveSession, args As IList(Of String)) As Integer
            Dim parsed As CopyMoveShared.CopyMoveArgs
            Try
                parsed = CopyMoveShared.ParseCommonArgs(args)
            Catch ex As ArgumentException
                PrintUsageAndError(ex.Message)
                Return 2
            End Try

            Dim copyResult As Tuple(Of ApiCallResult, String)
            Try
                copyResult = CopyMoveShared.ExecuteCopyLike(session, parsed, "move")
            Catch ex As InvalidOperationException
                Console.Error.WriteLine(
                    String.Format(
                        "error: failed to move '{0}' to '{1}' - {2}",
                        parsed.Source,
                        parsed.Destination,
                        ex.Message
                    )
                )
                Return 1
            End Try
            Dim sourceFilename As String = copyResult.Item2
            If String.IsNullOrWhiteSpace(sourceFilename) Then
                Console.Error.WriteLine(String.Format("error: {0} does not exist", parsed.Source))
                Return 1
            End If

            Dim headerMap = ApiShared.ConvertHeaderValues(parsed.Headers)
            If Not headerMap.ContainsKey("x-archive-keep-old-version") AndAlso Not parsed.NoBackup Then
                headerMap("x-archive-keep-old-version") = "1"
            End If
            headerMap("x-archive-cascade-delete") = "1"

            Dim sourceIdentifier As String = ApiShared.NormalizeArchivePath(parsed.Source).Split("/"c)(0)
            Dim deleteResult = session.DeleteS3File(sourceIdentifier, sourceFilename, headerMap, 2)
            If deleteResult.StatusCode = 204 Then
                Console.Error.WriteLine(
                    String.Format(
                        "success: moved '{0}' to '{1}'",
                        parsed.Source,
                        parsed.Destination
                    )
                )
                Return 0
            End If

            Console.Error.WriteLine("error: " & deleteResult.Text)
            Return 1
        End Function

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " move [-h] [-H KEY:VALUE] [--metadata KEY:VALUE]")
            Console.Error.WriteLine("                  [--source-overwrite] [--replace] [--dry-run] [--no-backup]")
            Console.Error.WriteLine("                  source destination")
            Console.Error.WriteLine(CliApp.ExecutableName() & " move: error: " & message)
        End Sub
    End Class
End Namespace

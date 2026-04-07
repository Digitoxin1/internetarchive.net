Imports InternetArchive.InternetArchiveCli.Core

Namespace InternetArchiveCli.Commands
    Public NotInheritable Class CopyCommand
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

            Try
                Dim copyResult = CopyMoveShared.ExecuteCopyLike(session, parsed, "copy")
            Catch ex As InvalidOperationException
                Console.Error.WriteLine(
                    String.Format(
                        "error: failed to copy '{0}' to '{1}' - {2}",
                        parsed.Source,
                        parsed.Destination,
                        ex.Message
                    )
                )
                Return 1
            End Try
            Console.Error.WriteLine(
                String.Format(
                    "success: copied '{0}' to '{1}'.",
                    parsed.Source,
                    parsed.Destination
                )
            )
            Return 0
        End Function

        Private Shared Sub PrintUsageAndError(message As String)
            Console.Error.WriteLine("usage: " & CliApp.ExecutableName() & " copy [-h] [-H KEY:VALUE] [--metadata KEY:VALUE]")
            Console.Error.WriteLine("                  [--source-overwrite] [--replace] [--dry-run] [--no-backup]")
            Console.Error.WriteLine("                  source destination")
            Console.Error.WriteLine(CliApp.ExecutableName() & " copy: error: " & message)
        End Sub
    End Class
End Namespace

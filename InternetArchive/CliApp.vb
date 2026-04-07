Imports System.IO

Namespace InternetArchiveCli
    Public Module CliApp
        Public Function ExecutableName() As String
            Try
                Dim currentProcess As Process = Process.GetCurrentProcess()
                If currentProcess IsNot Nothing AndAlso currentProcess.MainModule IsNot Nothing Then
                    Dim fileName As String = currentProcess.MainModule.FileName
                    Dim executable As String = Path.GetFileName(fileName)
                    If Not String.IsNullOrWhiteSpace(executable) Then
                        Return executable
                    End If
                End If
            Catch ex As Exception
                Console.Error.WriteLine("warning: failed to resolve executable name; using default ia.exe (" & ex.Message & ")")
            End Try

            Return "ia.exe"
        End Function
    End Module
End Namespace

Public Module CliApp
    Public Function ExecutableName() As String
        Return InternetArchiveCli.CliApp.ExecutableName()
    End Function
End Module

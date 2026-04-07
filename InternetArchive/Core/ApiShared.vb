Imports System.Globalization
Imports System.Net.Http
Imports System.Xml

Namespace InternetArchiveCli.Core
    Friend NotInheritable Class ApiShared
        Private Sub New()
        End Sub

        Friend Shared Sub ApplyLowAuth(request As HttpRequestMessage, accessKey As String, secretKey As String)
            If request Is Nothing Then
                Return
            End If
            If String.IsNullOrWhiteSpace(accessKey) OrElse String.IsNullOrWhiteSpace(secretKey) Then
                Return
            End If
            request.Headers.TryAddWithoutValidation(
                "Authorization",
                String.Format("LOW {0}:{1}", accessKey, secretKey)
            )
        End Sub

        Friend Shared Function ConvertHeaderValues(headers As Dictionary(Of String, Object)) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If headers Is Nothing Then
                Return result
            End If
            For Each kvp In headers
                Dim list = TryCast(kvp.Value, List(Of String))
                If list IsNot Nothing Then
                    result(kvp.Key) = String.Join(",", list)
                Else
                    result(kvp.Key) = Convert.ToString(kvp.Value, CultureInfo.InvariantCulture)
                End If
            Next
            Return result
        End Function

        Friend Shared Function GetS3XmlText(xmlText As String) As String
            Try
                Dim doc As New XmlDocument()
                doc.LoadXml(xmlText)
                Dim message As String = ""
                Dim resource As String = ""
                For Each node As XmlNode In doc.GetElementsByTagName("Message")
                    message &= node.InnerText
                Next
                For Each node As XmlNode In doc.GetElementsByTagName("Resource")
                    resource &= node.InnerText
                Next
                If message.Length = 0 AndAlso resource.Length = 0 Then
                    Return xmlText
                End If
                If resource.Length > 0 Then
                    Return message & ": " & resource
                End If
                Return message
            Catch
                Return xmlText
            End Try
        End Function

        Friend Shared Function ToObjectSequence(value As Object) As List(Of Object)
            Dim result As New List(Of Object)()
            If value Is Nothing Then
                Return result
            End If
            Dim arr = TryCast(value, ArrayList)
            If arr IsNot Nothing Then
                For Each entry In arr
                    result.Add(entry)
                Next
                Return result
            End If
            Dim objArray = TryCast(value, Object())
            If objArray IsNot Nothing Then
                result.AddRange(objArray)
            End If
            Return result
        End Function

        Friend Shared Function GetArchiveFileEntries(itemMetadata As Dictionary(Of String, Object)) As List(Of ArchiveFileEntry)
            Dim result As New List(Of ArchiveFileEntry)()
            If itemMetadata Is Nothing Then
                Return result
            End If

            Dim filesNode As Object = Nothing
            If Not itemMetadata.TryGetValue("files", filesNode) Then
                Return result
            End If

            For Each entry In ToObjectSequence(filesNode)
                Dim dict = TryCast(entry, Dictionary(Of String, Object))
                If dict IsNot Nothing Then
                    result.Add(New ArchiveFileEntry(dict))
                End If
            Next

            Return result
        End Function
    End Class
End Namespace

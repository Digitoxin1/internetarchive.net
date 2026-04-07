Namespace InternetArchiveCli.Core
    Public NotInheritable Class ApiCallResult
        Public Property Headers As Dictionary(Of String, String)
        Public Property JsonBody As Dictionary(Of String, Object)
        Public Property RequestUrl As String
        Public Property StatusCode As Integer
        Public Property Text As String
        Public Property WasSkipped As Boolean
    End Class
End Namespace

Imports System.Globalization

Namespace InternetArchiveCli.Core
    Friend NotInheritable Class ArchiveFileEntry
        Public Sub New(rawFields As Dictionary(Of String, Object))
            Dim normalized As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If rawFields IsNot Nothing Then
                For Each kvp In rawFields
                    normalized(kvp.Key) = kvp.Value
                Next
            End If

            Me.RawFields = normalized
            Me.Name = GetStringValue("name")
            Me.Format = GetStringValue("format")
            Me.Source = GetStringValue("source")
            Me.Md5 = GetStringValue("md5")
            Me.Mtime = GetStringValue("mtime")
        End Sub

        Public ReadOnly Property Format As String
        Public ReadOnly Property Md5 As String
        Public ReadOnly Property Mtime As String
        Public ReadOnly Property Name As String
        Public ReadOnly Property RawFields As Dictionary(Of String, Object)
        Public ReadOnly Property Source As String

        Public Function CloneRawFields() As Dictionary(Of String, Object)
            Dim clone As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            For Each kvp In RawFields
                clone(kvp.Key) = kvp.Value
            Next
            Return clone
        End Function

        Public Function GetStringValue(key As String) As String
            If RawFields Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return ""
            End If
            Dim value As Object = Nothing
            If Not RawFields.TryGetValue(key, value) Then
                Return ""
            End If
            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

        Public Function TryGetLong(key As String, ByRef value As Long) As Boolean
            value = 0
            If RawFields Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return False
            End If

            Dim raw As Object = Nothing
            If Not RawFields.TryGetValue(key, raw) OrElse raw Is Nothing Then
                Return False
            End If

            If TypeOf raw Is Long Then
                value = CLng(raw)
                Return True
            End If
            If TypeOf raw Is Integer Then
                value = CLng(CInt(raw))
                Return True
            End If
            If TypeOf raw Is Short Then
                value = CLng(CShort(raw))
                Return True
            End If

            Return Long.TryParse(
                Convert.ToString(raw, CultureInfo.InvariantCulture),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                value
            )
        End Function
    End Class
End Namespace

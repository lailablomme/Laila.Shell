Imports System.ComponentModel
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml.Serialization

Namespace Helpers
    Public Class FolderViewState
        Inherits Behaviors.GridViewExtBehavior.GridViewStateData

        Private Shared _cache As Dictionary(Of String, FolderViewState) = New Dictionary(Of String, FolderViewState)()

        Public Property ViewName As String
        Public Property ActiveView As Guid?

        Public Sub New()

        End Sub

        Public Sub New(folder As Folder)
            Me.ActiveView = If(folder?.ActiveView, folder.DefaultView)
            Me.SortPropertyName = If(folder?.ItemsSortPropertyName, "ItemNameDisplaySortValue")
            Me.SortDirection = If(folder?.ItemsSortDirection, ListSortDirection.Ascending)
            Me.GroupByPropertyName = If(folder?.ItemsGroupByPropertyName, Nothing)
        End Sub

        Public Shared Function FromFolder(folder As Folder) As FolderViewState
            Return FromViewName(folder)
        End Function

        Public Shared Function FromViewName(folder As Folder) As FolderViewState
            Dim result As FolderViewState = Nothing

            Dim dbPath As String = GetStateDBPath()
            Dim viewId As String = getMD5Hash(folder.FullPath)

            If _cache.ContainsKey(viewId) Then
                result = _cache(viewId)
            Else
                If File.Exists(Path.Combine(dbPath, viewId)) Then
                    Using stream As FileStream = New FileStream(Path.Combine(dbPath, viewId), FileMode.Open, FileAccess.Read)
                        Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
                        Try
                            result = s.Deserialize(stream)
                            If Not result.ActiveView.HasValue Then result.ActiveView = folder.DefaultView
                            If Not folder.Views.Exists(Function(v) v.Guid.Equals(result.ActiveView)) Then result.ActiveView = folder.DefaultView
                        Catch ex As Exception
                        End Try
                    End Using
                End If
            End If

            If result Is Nothing Then
                result = New FolderViewState(folder)
            End If

            result.ViewName = folder.FullPath
            If folder.ActiveView.HasValue AndAlso Not folder.ActiveView.Value.Equals(Guid.Empty) Then result.ActiveView = folder.ActiveView
            If Not String.IsNullOrWhiteSpace(folder.ItemsSortPropertyName) Then result.SortPropertyName = folder.ItemsSortPropertyName
            If Not String.IsNullOrWhiteSpace(folder.ItemsSortDirection) Then result.SortDirection = folder.ItemsSortDirection
            If Not String.IsNullOrWhiteSpace(folder.ItemsGroupByPropertyName) Then result.GroupByPropertyName = folder.ItemsGroupByPropertyName
            _cache(viewId) = result

            Return result
        End Function

        Public Sub Persist()
            Dim dbPath As String = GetStateDBPath()
            Dim viewId As String = getMD5Hash(ViewName)

            Using stream As FileStream = New FileStream(Path.Combine(dbPath, viewId), FileMode.Create, FileAccess.Write)
                Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
                s.Serialize(stream, Me)
            End Using
        End Sub

        Private Shared Function getMD5Hash(input As String) As String
            Using md5 As MD5 = md5.Create()
                ' Convert the input string to a byte array and compute the hash
                Dim inputBytes As Byte() = Encoding.UTF8.GetBytes(input)
                Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)

                ' Convert the byte array to a hexadecimal string
                Dim sb As New StringBuilder()
                For Each b As Byte In hashBytes
                    sb.Append(b.ToString("x2"))
                Next

                Return sb.ToString()
            End Using
        End Function

        Public Shared Function GetStateDBPath() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell", "FolderViewState")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return path
        End Function
    End Class
End Namespace
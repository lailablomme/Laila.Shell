Imports System.IO
Imports System.Threading
Imports System.Windows.Media
Imports Laila.AutoCompleteTextBox

Namespace Helpers
    Public Class AddressBarSuggestionProvider
        Implements ISuggestionProviderAsync

        Friend _lock As SemaphoreSlim = New SemaphoreSlim(1, 1)
        Friend _items As List(Of Item)

        Public Async Function GetSuggestions(filter As String) As Task(Of IEnumerable) Implements ISuggestionProviderAsync.GetSuggestions
            Dim func As Func(Of Task(Of IEnumerable)) =
                Async Function() As Task(Of IEnumerable)
                    filter = Environment.ExpandEnvironmentVariables(filter).Trim()
                    Dim folderName As String
                    Dim fileName As String
                    If filter.StartsWith("\\") Then
                        If filter.Substring(2).LastIndexOf(IO.Path.DirectorySeparatorChar) = -1 Then
                            folderName = "\\"
                            fileName = filter
                        Else
                            folderName = filter.Substring(0, filter.LastIndexOf(IO.Path.DirectorySeparatorChar))
                            fileName = filter.Substring(folderName.Length - 1)
                        End If
                    Else
                        folderName = IO.Path.GetDirectoryName(filter)
                        fileName = IO.Path.GetFileName(filter)
                    End If
                    Dim folder As Folder
                    If String.IsNullOrWhiteSpace(folderName) Then
                        If filter.Equals(Path.GetPathRoot(filter)) Then
                            folderName = filter
                            fileName = ""
                        End If
                    End If
                    If Not String.IsNullOrWhiteSpace(folderName) Then
                        If folder Is Nothing Then folder = Item.FromParsingName(folderName, Nothing)
                        If folder Is Nothing Then folder = Await Item.FromParsingNameDeepGetAsync(folderName)
                    End If
                    Dim items As List(Of Item)
                    If Not folder Is Nothing Then
                        items = If(filter.EndsWith(IO.Path.DirectorySeparatorChar), New List(Of Item) From {folder}, New List(Of Item)).ToList() _
                            .Union((Await folder.GetItemsAsync()).Where(Function(f) _
                                If(If(f.FullPath.StartsWith("\\"), f.FullPath.Substring(2), f.FullPath).TrimEnd(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar)) _
                                       .ToLower().StartsWith(fileName.ToLower()) _
                                OrElse (Not f.DisplayName Is Nothing AndAlso f.DisplayName.ToLower().StartsWith(fileName.ToLower()))).ToList()).ToList()
                    Else
                        items = Shell.SpecialFolders.Values.Where(Function(f) f.DisplayName.ToLower().StartsWith(fileName.ToLower()) _
                                                                      OrElse f.FullPath.ToLower().StartsWith(fileName.ToLower())) _
                                                           .OrderBy(Function(f) f.DisplayName).Cast(Of Item).ToList()
                    End If
                    Await _lock.WaitAsync()
                    Try
                        If Not _items Is Nothing Then
                            For Each item In _items
                                item.MaybeDispose()
                            Next
                        End If
                        _items = New List(Of Item)()
                        _items.AddRange(items)
                    Finally
                        _lock.Release()
                    End Try
                    Return items
                End Function

            Return Await Task.Run(func)
        End Function
    End Class
End Namespace
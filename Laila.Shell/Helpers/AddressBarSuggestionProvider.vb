Imports System.IO
Imports System.Threading
Imports System.Windows.Media
Imports Laila.AutoCompleteTextBox

Namespace Helpers
    Public Class AddressBarSuggestionProvider
        Implements ISuggestionProviderAsync

        Friend _lock As SemaphoreSlim = New SemaphoreSlim(1, 1)
        Friend _folder As Folder

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
                        If folder Is Nothing Then folder = Item.FromParsingName(folderName, Nothing, False)
                        If folder Is Nothing Then folder = Await Item.FromParsingNameDeepGetAsync(folderName)
                    End If
                    Dim items As List(Of Item)
                    If Not folder Is Nothing Then
                        folder.IsVisibleInAddressBar = True
                        Dim byPath As List(Of Item) =
                            (Await folder.GetItemsAsync()).Where(Function(f) _
                                If(If(f.FullPath.StartsWith("\\"), f.FullPath.Substring(2), f.FullPath).TrimEnd(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar)) _
                                       .ToLower().StartsWith(fileName.ToLower())).ToList()
                        For Each f In byPath
                            f.AddressBarDisplayName =
                                If(If(f.FullPath.StartsWith("\\"), f.FullPath.Substring(2), f.FullPath).TrimEnd(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar))
                        Next
                        Dim byDisplayName As List(Of Item) =
                            (Await folder.GetItemsAsync()).Where(Function(f) _
                                Not f.DisplayName Is Nothing AndAlso f.DisplayName.ToLower().StartsWith(fileName.ToLower())).ToList()
                        For Each f In byDisplayName
                            f.AddressBarDisplayName = f.DisplayName
                        Next
                        items = byDisplayName.Union(byPath).Distinct().OrderBy(Function(f) f.AddressBarDisplayPath).ToList()
                        For Each f In items
                            f.AddressBarRoot = folderName
                        Next
                    ElseIf String.IsNullOrWhiteSpace(folderName) Then
                        items = Shell.GetSpecialFolders().Values.Where(Function(f) f.DisplayName.ToLower().StartsWith(fileName.ToLower()) _
                                                                       OrElse f.FullPath.ToLower().StartsWith(fileName.ToLower())) _
                                                                .OrderBy(Function(f) f.DisplayName).Cast(Of Item).ToList()
                    Else
                        items = New List(Of Item)()
                    End If

                    If Not _folder Is Nothing Then
                        _folder.IsVisibleInAddressBar = False
                    End If
                    _folder = folder
                    Return items
                End Function

            Return Await Task.Run(func)
        End Function
    End Class
End Namespace
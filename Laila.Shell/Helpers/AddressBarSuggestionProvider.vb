Imports System.IO
Imports System.Windows.Media
Imports Laila.AutoCompleteTextBox

Namespace Helpers
    Public Class AddressBarSuggestionProvider
        Implements ISuggestionProviderAsync

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
                    If folderName Is Nothing Then
                        If Not filter.EndsWith(IO.Path.DirectorySeparatorChar) Then
                            folder = Shell.SpecialFolders("This computer")
                            fileName = filter
                        Else
                            folderName = filter
                        End If
                    ElseIf String.IsNullOrWhiteSpace(folderName) AndAlso filter = IO.Path.GetPathRoot(fileName) Then
                        folderName = filter
                    End If
                    If folder Is Nothing Then folder = Item.FromParsingName(folderName, Nothing)
                    If Not folder Is Nothing Then
                        Return If(filter.EndsWith(IO.Path.DirectorySeparatorChar), New List(Of Item) From {folder}, New List(Of Item)).ToList() _
                            .Union((Await folder.GetItemsAsync()).Where(Function(f) _
                                If(If(f.FullPath.StartsWith("\\"), f.FullPath.Substring(2), f.FullPath).TrimEnd(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                   f.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar)) _
                                       .ToLower().StartsWith(fileName.ToLower()) _
                                OrElse (Not f.DisplayName Is Nothing AndAlso f.DisplayName.ToLower().StartsWith(fileName.ToLower()))).ToList())
                    Else
                        Return Shell.SpecialFolders.Values.Where(Function(f) f.DisplayName.ToLower().StartsWith(fileName.ToLower()) _
                                                                      OrElse f.FullPath.ToLower().StartsWith(fileName.ToLower())) _
                                                          .OrderBy(Function(f) f.DisplayName).ToList()
                    End If
                End Function

            Return Await Task.Run(func)
        End Function
    End Class
End Namespace
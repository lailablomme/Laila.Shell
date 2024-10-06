Imports System.IO
Imports System.Windows.Media
Imports Laila.AutoCompleteTextBox

Namespace Helpers
    Public Class AddressBarSuggestionProvider
        Implements ISuggestionProviderAsync

        Public Async Function GetSuggestions(filter As String) As Task(Of IEnumerable) Implements ISuggestionProviderAsync.GetSuggestions
            Dim func As Func(Of Task(Of IEnumerable)) =
                Async Function() As Task(Of IEnumerable)
                    filter = Environment.ExpandEnvironmentVariables(filter)
                    Dim folderName As String = IO.Path.GetDirectoryName(filter)
                    Dim fileName As String = IO.Path.GetFileName(filter)
                    Dim folder As Folder
                    If String.IsNullOrWhiteSpace(folderName) Then
                        fileName = filter
                        If fileName.StartsWith("\\") Then
                            folder = Shell.SpecialFolders("Network")
                        ElseIf fileName = IO.Path.GetPathRoot(fileName) Then
                            folder = Shell.SpecialFolders("This computer")
                            Return (Await folder.GetItems()) _
                                .Where(Function(f) IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower().StartsWith(fileName.ToLower()) _
                                    OrElse f.DisplayName.ToLower.StartsWith(fileName.ToLower())).ToList()
                        End If
                    Else
                        folder = Await Item.FromParsingNameDeepGet(folderName, Nothing)
                    End If
                    If Not folder Is Nothing Then
                        Return (Await folder.GetItems()) _
                            .Where(Function(f) IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower().StartsWith(fileName.ToLower()) _
                                    OrElse f.DisplayName.ToLower.StartsWith(fileName.ToLower())).ToList()
                    Else
                        Return Shell.SpecialFolders.Values.Where(Function(f) f.DisplayName.ToLower().StartsWith(fileName.ToLower()) _
                                                                      OrElse f.FullPath.ToLower().StartsWith(fileName.ToLower())) _
                                                          .OrderBy(Function(f) f.DisplayName).ToList()
                    End If

                    'If Not folder Is Nothing Then
                    '    'Dim fileName As String = IO.Path.GetFileName(filter)

                    'End If
                    '' resolve environment variable?
                    'filter = Environment.ExpandEnvironmentVariables(filter)

                    'Dim path As String = filter.Trim()

                    '' get parts of path
                    'Dim isNetworkPath As Boolean = path.StartsWith("\\")
                    'If isNetworkPath Then Path = Path.Substring(2)

                    'Dim parts As List(Of String) = New List(Of String)()
                    'While Not String.IsNullOrWhiteSpace(path)
                    '    Debug.WriteLine(path)
                    '    If path = IO.Path.GetPathRoot(path) Then
                    '        parts.Add(path)
                    '    Else
                    '        parts.Add(IO.Path.GetFileName(path))
                    '    End If
                    '    path = IO.Path.GetDirectoryName(path)
                    'End While
                    'parts.Reverse()

                    'If parts.Count > 0 Then
                    '    If isNetworkPath Then
                    '        Dim folder As Folder = Shell.SpecialFolders("Network")
                    '        Dim i As Integer

                    '        ' find folder
                    '        For i = 0 To parts.Count - 1
                    '            If i = parts.Count - 1 Then Exit For
                    '            Dim subFolder As Folder = (Await folder.GetItems()).FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(i).ToLower())
                    '            folder.Dispose()
                    '            folder = subFolder
                    '            If folder Is Nothing Then Exit For
                    '        Next

                    '        If i = parts.Count - 1 Then
                    '            ' found folders all the way down -- return matching files in folder
                    '            Try
                    '                Return (Await folder.GetItems()).Where(Function(f) f.DisplayName.ToLower().StartsWith(parts(parts.Count - 1).ToLower())).ToList()
                    '            Finally
                    '                folder.Dispose()
                    '            End Try
                    '        Else
                    '            Return Nothing
                    '        End If
                    '    ElseIf parts(0) = IO.Path.GetPathRoot(filter) Then
                    '        ' this is a path on disk
                    '        Dim folder As Folder = Shell.SpecialFolders("This computer")
                    '        Dim j As Integer = 0

                    '        ' find folder
                    '        For j = 0 To parts.Count - 1
                    '            If j = parts.Count - 1 Then Exit For
                    '            Dim subFolder As Folder
                    '            If j = 0 Then
                    '                subFolder = (Await folder.GetItems()).FirstOrDefault(Function(f) IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower() = parts(j).ToLower())
                    '            Else
                    '                subFolder = (Await folder.GetItems()).FirstOrDefault(Function(f) IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower() = parts(j).ToLower())
                    '            End If
                    '            folder.Dispose()
                    '            folder = subFolder
                    '            If folder Is Nothing Then Exit For
                    '        Next

                    '        If j = parts.Count - 1 Then
                    '            ' found folders all the way down -- return matching files in folder
                    '            Try
                    '                If j = 0 Then
                    '                    Return (Await folder.GetItems()).Where(Function(f) IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower().StartsWith(parts(parts.Count - 1).ToLower())).ToList()
                    '                Else
                    '                    Return (Await folder.GetItems()).Where(Function(f) IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower().StartsWith(parts(parts.Count - 1).ToLower())).ToList()
                    '                End If
                    '            Finally
                    '                folder.Dispose()
                    '            End Try
                    '        Else
                    '            Return Nothing
                    '        End If
                    '    Else
                    '        ' root must be some special folder
                    '        Dim folder As Folder = Shell.SpecialFolders.Values.FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(0).ToLower() _
                    '                                                                               OrElse f.FullPath.ToLower() = parts(0).ToLower())
                    '        Dim i As Integer

                    '        ' find folder
                    '        If Not folder Is Nothing AndAlso parts.Count > 1 Then
                    '            For i = 1 To parts.Count - 1
                    '                If i = parts.Count - 1 Then Exit For
                    '                Dim subFolder As Folder = (Await folder.GetItems()).FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(i).ToLower())
                    '                folder.Dispose()
                    '                folder = subFolder
                    '                If folder Is Nothing Then Exit For
                    '            Next
                    '        End If

                    '        If i = parts.Count - 1 Then
                    '            ' found folders all the way down -- return matching files in folder
                    '            If folder Is Nothing AndAlso i = 0 Then
                    '                Return Shell.SpecialFolders.Values.Where(Function(f) f.DisplayName.ToLower().StartsWith(parts(0).ToLower()) _
                    '                                                              OrElse f.FullPath.ToLower().StartsWith(parts(0).ToLower())) _
                    '                                                  .OrderBy(Function(f) f.DisplayName).ToList()
                    '            ElseIf Not folder Is Nothing AndAlso i = 0 Then
                    '                Return {folder}
                    '            Else
                    '                Try
                    '                    Return (Await folder.GetItems()).Where(Function(f) f.DisplayName.ToLower().StartsWith(parts(parts.Count - 1).ToLower())).ToList()
                    '                Finally
                    '                    folder.Dispose()
                    '                End Try
                    '            End If
                    '        Else
                    '            Return Nothing
                    '        End If
                    '    End If
                    'Else
                    '    ' TODO: return recent items
                    '    Return Nothing
                    'End If
                End Function

            Return Await Task.Run(func)
        End Function
    End Class
End Namespace
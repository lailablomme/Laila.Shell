﻿Imports System.IO
Imports Laila.AutoCompleteTextBox

Namespace Helpers
    Public Class AddressBarSuggestionProvider
        Implements ISuggestionProviderAsync

        Friend _folder As Folder
        Friend _items As List(Of Item)
        Friend _lock As Object = New Object()

        Public Async Function GetSuggestions(filter As String) As Task(Of IEnumerable) Implements ISuggestionProviderAsync.GetSuggestions
            Dim func As Func(Of Task(Of IEnumerable)) =
                Async Function() As Task(Of IEnumerable)
                    ' expand enviroment variables
                    filter = Environment.ExpandEnvironmentVariables(filter).Trim()

                    ' parse foldername and filename
                    Dim folderName As String = String.Empty
                    Dim fileName As String = String.Empty
                    Dim folder As Folder
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
                    If String.IsNullOrWhiteSpace(folderName) Then
                        If filter.Equals(Path.GetPathRoot(filter)) Then
                            folderName = filter
                        End If
                    End If

                    ' get folder
                    If Not String.IsNullOrWhiteSpace(folderName) Then
                        If folder Is Nothing Then
                            If folderName.StartsWith(IO.Path.DirectorySeparatorChar) Then
                                folder = Item.FromParsingName(IO.Path.Combine(
                                        (Await Shell.GetSpecialFolder("This computer").GetItemsAsync()).FirstOrDefault()?.FullPath,
                                        folderName.TrimStart(IO.Path.DirectorySeparatorChar)), Nothing, False)
                            Else
                                folder = Item.FromParsingName(folderName, Nothing, False)
                            End If
                        End If
                        If folder Is Nothing Then folder = Await Item.FromParsingNameDeepGetAsync(folderName)
                    End If

                    ' get items
                    Dim items As List(Of Item)
                    If Not folder Is Nothing Then
                        If String.IsNullOrWhiteSpace(fileName) Then
                            items = {folder}.Cast(Of Item).ToList()
                        Else
                            items = New List(Of Item)()
                        End If
                        items = items.Union((Await folder.GetItemsAsync()).OrderBy(Function(f) f.AddressBarDisplayPath)).ToList()
                    ElseIf String.IsNullOrWhiteSpace(folderName) Then
                        items = AddressBarHistory.GetHistory(fileName).Union(
                            Shell.GetSpecialFolders().Values.ToList().OrderBy(Function(f) f.AddressBarDisplayPath)) _
                                .Cast(Of Item).ToList()
                    Else
                        items = New List(Of Item)()
                    End If

                    ' filter items
                    Dim byPath As List(Of Item) =
                        items.Where(Function(f) _
                            If(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                               f.FullPath.Trim(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                               f.FullPath.Trim(IO.Path.DirectorySeparatorChar)) _
                                   .ToLower().StartsWith(fileName.ToLower())).ToList()
                    If Not folder Is Nothing Then
                        For Each f In byPath
                            f.AddressBarDisplayName =
                                If(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                   f.FullPath.Trim(IO.Path.DirectorySeparatorChar).Substring(f.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                   f.FullPath.Trim(IO.Path.DirectorySeparatorChar))
                        Next
                    End If
                    Dim byDisplayName As List(Of Item) =
                        items.Where(Function(f) _
                            Not f.DisplayName Is Nothing AndAlso f.DisplayName.ToLower().StartsWith(fileName.ToLower())).ToList()
                    If Not folder Is Nothing Then
                        For Each f In byDisplayName
                            f.AddressBarDisplayName = f.DisplayName
                        Next
                    End If
                    items = byDisplayName.Union(byPath).ToList()

                    ' set root
                    If Not folder Is Nothing Then
                        For Each f In items
                            If Not f.Equals(folder) Then
                                f.AddressBarRoot = folderName
                            Else
                                f.AddressBarRoot = Nothing
                                f.AddressBarDisplayName = folderName
                            End If
                        Next
                    End If

                    ' get distinct
                    Dim pidls As List(Of String) = New List(Of String)()
                    For Each item In items
                        pidls.Add(item.Pidl.ToString())
                    Next
                    items = pidls.Distinct().Select(Function(p) items.FirstOrDefault(Function(i) i.Pidl.ToString().Equals(p))).ToList()

                    SyncLock _lock
                        ' release previous folder
                        If Not _folder Is Nothing AndAlso Not items.Contains(_folder) AndAlso Not _folder.Equals(folder) Then
                            _folder.IsVisibleInAddressBar = False
                        End If

                        ' remember current folder and prevent it from getting disposed
                        _folder = folder
                        If Not _folder Is Nothing Then
                            _folder.IsVisibleInAddressBar = True
                        End If

                        ' release previous items
                        If Not _items Is Nothing Then
                            For Each item In _items.Where(Function(i) Not items.Contains(i) AndAlso Not i.Equals(_folder))
                                item.IsVisibleInAddressBar = False
                            Next
                        End If

                        ' remember current items and prevent them from getting disposed
                        _items = items.ToList()
                        For Each item In _items
                            item.IsVisibleInAddressBar = True
                        Next

                        ' take 15
                        Return items.Take(15)
                    End SyncLock
                End Function

            Return Await Task.Run(func)
        End Function
    End Class
End Namespace
Imports System.IO
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers

Namespace ViewModels
    Public Class TreeViewModel
        Inherits NotifyPropertyChangedBase

        Private _view As TreeView
        Private _folders1 As List(Of TreeViewFolder)
        Private _folders2 As List(Of TreeViewFolder)
        Private _folders3 As List(Of TreeViewFolder)
        Private _selectionHelper1 As SelectionHelper(Of TreeViewFolder) = Nothing
        Private _selectionHelper2 As SelectionHelper(Of TreeViewFolder) = Nothing
        Private _selectionHelper3 As SelectionHelper(Of TreeViewFolder) = Nothing

        Public Sub New(view As TreeView)
            _view = view

            ' home and galery
            _folders1 = New List(Of TreeViewFolder) From
            {
                TreeViewFolder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing),
                TreeViewFolder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing)
            }

            ' all the special folders under quick launch
            _folders2 = New List(Of TreeViewFolder)()
            For Each f In CType(Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing), Folder) _
                .Items.Where(Function(i) TypeOf i Is Folder AndAlso
                (i.Parent Is Nothing OrElse Path.GetDirectoryName(i.FullPath) = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)))
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) Then
                    _folders2.Add(TreeViewFolder.FromParsingName(f.FullPath, Nothing))
                End If
            Next

            ' 5 first regular folders under quicklaunch
            For Each f In CType(Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing), Folder) _
                .Items.Where(Function(i) TypeOf i Is Folder AndAlso
                (Not i.Parent Is Nothing AndAlso Path.GetDirectoryName(i.FullPath) <> Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))) _
                .Take(5)
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) Then
                    _folders2.Add(TreeViewFolder.FromParsingName(f.FullPath, Nothing))
                End If
            Next

            ' all special folders under user profile that we'rent added yet
            For Each f In CType(Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Nothing), Folder) _
                .Items.Where(Function(i) TypeOf i Is Folder AndAlso
                    Not _folders1.Exists(Function(f2) f2.FullPath = i.FullPath))
                Dim fpure As Folder = Folder.FromParsingName(f.FullPath, Nothing)
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    (fpure.Parent Is Nothing OrElse fpure.Parent.Parent Is Nothing) Then
                    _folders1.Insert(2, TreeViewFolder.FromParsingName(f.FullPath, Nothing))
                End If
            Next

            ' this computer & network
            _folders3 = New List(Of TreeViewFolder) From
            {
                TreeViewFolder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing),
                TreeViewFolder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing)
            }

            Dim isWorking As Boolean
            AddHandler _view.Loaded,
                Sub(s As Object, e As EventArgs)
                    _selectionHelper1 = New SelectionHelper(Of TreeViewFolder)(_view.treeView1)
                    _selectionHelper1.SelectionChanged =
                        Sub()
                            If Not isWorking Then
                                isWorking = True
                                _selectionHelper2.SetSelectedItems({})
                                _selectionHelper3.SetSelectedItems({})
                                Me.OnSelectionChanged()
                                NotifyOfPropertyChange("SelectedItem")
                                NotifyOfPropertyChange("SelectedItems")
                                isWorking = False
                            End If
                        End Sub
                    _selectionHelper2 = New SelectionHelper(Of TreeViewFolder)(_view.treeView2)
                    _selectionHelper2.SelectionChanged =
                        Sub()
                            If Not isWorking Then
                                isWorking = True
                                _selectionHelper1.SetSelectedItems({})
                                _selectionHelper3.SetSelectedItems({})
                                Me.OnSelectionChanged()
                                NotifyOfPropertyChange("SelectedItem")
                                NotifyOfPropertyChange("SelectedItems")
                                isWorking = False
                            End If
                        End Sub
                    _selectionHelper3 = New SelectionHelper(Of TreeViewFolder)(_view.treeView3)
                    _selectionHelper3.SelectionChanged =
                        Sub()
                            If Not isWorking Then
                                isWorking = True
                                _selectionHelper1.SetSelectedItems({})
                                _selectionHelper2.SetSelectedItems({})
                                Me.OnSelectionChanged()
                                NotifyOfPropertyChange("SelectedItem")
                                NotifyOfPropertyChange("SelectedItems")
                                isWorking = False
                            End If
                        End Sub
                End Sub
        End Sub

        Public ReadOnly Property Folders1 As List(Of TreeViewFolder)
            Get
                Return _folders1
            End Get
        End Property

        Public ReadOnly Property Folders2 As List(Of TreeViewFolder)
            Get
                Return _folders2
            End Get
        End Property

        Public ReadOnly Property Folders3 As List(Of TreeViewFolder)
            Get
                Return _folders3
            End Get
        End Property

        Protected Overridable Sub OnSelectionChanged()
            If Not Me.SelectedItem Is Nothing Then
                _view.SelectedFolderName = Me.SelectedItem.FullPath
            Else
                _view.SelectedFolderName = Nothing
            End If
        End Sub

        Public ReadOnly Property SelectedItem As Item
            Get
                If _selectionHelper1.SelectedItems.Count = 1 Then
                    Return _selectionHelper1.SelectedItems(0)
                ElseIf _selectionHelper2.SelectedItems.Count = 1 Then
                    Return _selectionHelper2.SelectedItems(0)
                ElseIf _selectionHelper3.SelectedItems.Count = 1 Then
                    Return _selectionHelper3.SelectedItems(0)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Sub SetSelectedItem(value As Item)
            If value Is Nothing Then
                _selectionHelper1.SetSelectedItems(New TreeViewFolder() {})
                _selectionHelper2.SetSelectedItems(New TreeViewFolder() {})
                _selectionHelper3.SetSelectedItems(New TreeViewFolder() {})
            Else
                _selectionHelper1.SetSelectedItems(New TreeViewFolder() {value})
                If Not _selectionHelper1.SelectedItems.Count = 1 Then
                    _selectionHelper2.SetSelectedItems(New TreeViewFolder() {value})
                    If Not _selectionHelper2.SelectedItems.Count = 1 Then
                        _selectionHelper3.SetSelectedItems(New TreeViewFolder() {value})
                    End If
                End If
            End If
        End Sub

        Public Sub SetSelectedFolder(path As String)
            Dim list As List(Of Folder) = New List(Of Folder)()
            Dim f As Folder = Folder.FromParsingName(path, Nothing)
            While Not f.Parent Is Nothing AndAlso
                Not _folders1.Exists(Function(f1) f1.FullPath = f.FullPath) AndAlso
                Not _folders2.Exists(Function(f1) f1.FullPath = f.FullPath) AndAlso
                Not _folders3.Exists(Function(f1) f1.FullPath = f.FullPath)
                list.Add(f)
                f = f.Parent
            End While

            Dim tf As TreeViewFolder
            If _folders1.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                tf = _folders1.First(Function(f1) f1.FullPath = f.FullPath)
            ElseIf _folders2.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                tf = _folders2.First(Function(f1) f1.FullPath = f.FullPath)
            ElseIf _folders3.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                tf = _folders3.First(Function(f1) f1.FullPath = f.FullPath)
            Else
                tf = Nothing
            End If

            If Not tf Is Nothing Then
                list.Reverse()
                For Each item In list
                    If Not tf.IsExpanded Then tf._folders = Nothing
                    Dim tf2 As TreeViewFolder = tf.Folders.FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                    If Not tf2 Is Nothing Then
                        tf.IsExpanded = True
                        'Await Me.SetSelectedItem(tf2)
                        tf = tf2
                    Else
                        Exit For
                    End If
                Next
                Me.SetSelectedItem(tf)
            End If
        End Sub
    End Class
End Namespace
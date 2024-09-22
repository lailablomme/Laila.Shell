Imports System.Collections.ObjectModel
Imports System.ComponentModel.Design
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers

Namespace ViewModels
    Public Class DetailsListViewModel
        Inherits NotifyPropertyChangedBase

        Private _view As DetailsListView
        Private _folderName As String
        Private _folder As Folder
        Private _gridView As GridView
        Private _columnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
        Private _isLoading As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _scrollState As Dictionary(Of String, ScrollState) = New Dictionary(Of String, ScrollState)()
        Private _skipSavingScrollState As Boolean = False

        Public Sub New(view As DetailsListView)
            _view = view

            Dim listViewItemStyle As Style = New Style()
            listViewItemStyle.TargetType = GetType(ListViewItem)
            listViewItemStyle.BasedOn = _view.TryFindResource(GetType(ListViewItem))
            listViewItemStyle.Setters.Add(New Setter(ListViewItem.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch))
            listViewItemStyle.Setters.Add(New EventSetter(ListViewItem.MouseDoubleClickEvent, New MouseButtonEventHandler(AddressOf OnListViewItemDoubleClick)))
            _view.Resources.Add(GetType(ListViewItem), listViewItemStyle)

            AddHandler _view.Loaded,
                Sub(s As Object, e As EventArgs)
                    _selectionHelper = New SelectionHelper(Of Item)(_view.listView)
                    _selectionHelper.SelectionChanged =
                        Async Function() As Task
                            Await Me.OnSelectionChanged()

                            NotifyOfPropertyChange("SelectedItem")
                            NotifyOfPropertyChange("SelectedItems")
                        End Function

                    AddHandler _view.listView.PreviewMouseRightButtonDown, AddressOf OnListViewItemRightClick
                End Sub
        End Sub

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Set(value As Boolean)
                SetValue(_isLoading, value)

                If Not _isLoading Then
                    If Not Me.Folder Is Nothing AndAlso _scrollState.ContainsKey(Me.Folder.FullPath) Then
                        Application.Current.Dispatcher.Invoke(
                            Sub()
                                loadScrollState()
                            End Sub)
                    End If
                Else
                    If Not Me.Folder Is Nothing AndAlso Not _skipSavingScrollState Then
                        Application.Current.Dispatcher.Invoke(
                            Sub()
                                saveScrollState()
                            End Sub)
                    End If
                    _skipSavingScrollState = False
                End If
            End Set
        End Property

        Private Sub saveScrollState()
            If _scrollState.ContainsKey(Me.Folder.FullPath) Then
                _scrollState.Remove(Me.Folder.FullPath)
            End If
            _scrollState.Add(
                Me.Folder.FullPath,
                New ScrollState() With
                {
                    .OffsetY = getScrollViewer().VerticalOffset,
                    .SelectedItems = Me.SelectedItems.ToList()
                })
        End Sub

        Private Sub loadScrollState()
            If _scrollState.ContainsKey(Me.Folder.FullPath) Then
                Dim state As ScrollState = _scrollState(Me.Folder.FullPath)
                getScrollViewer().ScrollToVerticalOffset(state.OffsetY)
                Me.SetSelectedItems(state.SelectedItems)

                Dim index As Integer = _scrollState.Keys.ToList().IndexOf(Me.Folder.FullPath)
                While _scrollState.Count > index
                    _scrollState.Remove(_scrollState.Keys.ElementAt(_scrollState.Count - 1))
                End While
            End If
        End Sub

        Public Property FolderName As String
            Get
                Return _folderName
            End Get
            Set(value As String)
                If Not (String.IsNullOrWhiteSpace(Me.FolderName) AndAlso String.IsNullOrWhiteSpace(value)) AndAlso
                    Not EqualityComparer(Of String).Default.Equals(Me.FolderName, value) Then
                    If Not Me.Folder Is Nothing Then
                        saveScrollState()
                    End If

                    SetValue(_folderName, value)
                    If Not String.IsNullOrWhiteSpace(Me.FolderName) Then
                        _skipSavingScrollState = True
                        If Not Me.Folder Is Nothing Then Me.Folder.Dispose()
                        Me.Folder = Folder.FromParsingName(FolderName, _view.LogicalParent, Sub(val) Me.IsLoading = val)
                        Me.ColumnsIn = buildColumnsIn()
                    End If
                    _view.FolderName = _folderName
                End If
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return _folder
            End Get
            Set(value As Folder)
                SetValue(_folder, value)
                Me.NotifyOfPropertyChange("Items")
            End Set
        End Property

        Public Property ColumnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
            Get
                Return _columnsIn
            End Get
            Set(value As Behaviors.GridViewExtBehavior.ColumnsInData)
                SetValue(_columnsIn, value)
            End Set
        End Property

        Public Function buildColumnsIn() As Behaviors.GridViewExtBehavior.ColumnsInData
            Dim d As Behaviors.GridViewExtBehavior.ColumnsInData = New Behaviors.GridViewExtBehavior.ColumnsInData()
            d.ViewName = Me.Folder.FullPath
            d.PrimarySortProperties = "PrimarySort"
            d.Items = New List(Of GridViewColumn)()

            For Each column In Me.Folder.Columns.Where(Function(c) Not String.IsNullOrWhiteSpace(c.DisplayName))
                Dim [property] As [Property] = [Property].FromCanonicalName(column.CanonicalName)

                Dim gvc As GridViewColumn = New GridViewColumn()
                gvc.Header = column.DisplayName
                gvc.CellTemplate = getCellTemplate(column, [property])
                gvc.SetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty, CBool(column.State And CM_STATE.VISIBLE))
                gvc.SetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty, String.Format("Properties[{0}].Value", column.CanonicalName))
                If column.CanonicalName = "System.ItemNameDisplay" Then
                    gvc.SetValue(Behaviors.GridViewExtBehavior.SortPropertyNameProperty, "ItemNameDisplaySortValue")
                End If
                gvc.SetValue(Behaviors.GridViewExtBehavior.GroupByPropertyNameProperty, String.Format("Properties[{0}].Text", column.CanonicalName))
                d.Items.Add(gvc)
            Next

            Return d
        End Function

        Private Function getCellTemplate(column As Column, [property] As [Property]) As DataTemplate
            Dim template As DataTemplate = New DataTemplate()

            Dim gridFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(Grid))
            Dim columnDefinition1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition1.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition1)
            Dim columnDefinition2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition2.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition2)
            Dim columnDefinition3 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition3.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Star))
            gridFactory.AppendChild(columnDefinition3)

            Dim getIconFactory As Func(Of String, FrameworkElementFactory) =
                Function(bindTo As String) As FrameworkElementFactory
                    Dim imageFactory1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                    imageFactory1.SetValue(Grid.ColumnProperty, 1)
                    imageFactory1.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                    imageFactory1.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                    imageFactory1.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                    imageFactory1.SetValue(Image.SourceProperty, New Binding() With {
                        .Path = New PropertyPath(bindTo),
                        .Mode = BindingMode.OneWay
                    })
                    Dim style As Style = New Style(GetType(Image))
                    style.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Collapsed))
                    style.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                    Dim dataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding(String.Format("ColumnIndexFor[Properties[{0}].Value]", column.CanonicalName)) With
                                   {
                                       .ElementName = "ext",
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = 0
                    }
                    dataTrigger1.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
                    style.Triggers.Add(dataTrigger1)
                    Dim dataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsHidden") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger2.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger2)
                    imageFactory1.SetValue(Image.StyleProperty, style)
                    Return imageFactory1
                End Function
            gridFactory.AppendChild(getIconFactory("Icon16"))
            gridFactory.AppendChild(getIconFactory("Overlay16"))

            If [property].HasIcon Then
                Dim imageFactory2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                imageFactory2.SetValue(Grid.ColumnProperty, 1)
                imageFactory2.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                imageFactory2.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                imageFactory2.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                imageFactory2.SetValue(Image.SourceProperty, New Binding() With {
                    .Path = New PropertyPath(String.Format("Properties[{0}].Icon16", column.CanonicalName)),
                    .Mode = BindingMode.OneWay
                })
                gridFactory.AppendChild(imageFactory2)
            End If

            Dim textBlockFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(TextBlock))
            textBlockFactory.SetValue(Grid.ColumnProperty, 2)
            textBlockFactory.SetValue(TextBlock.TextAlignmentProperty, column.Alignment)
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
            textBlockFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis)
            textBlockFactory.SetValue(TextBlock.TextProperty, New Binding() With {
                .Path = New PropertyPath(String.Format("Properties[{0}].Text", column.CanonicalName)),
                .Mode = BindingMode.OneWay
            })
            gridFactory.AppendChild(textBlockFactory)

            template.VisualTree = gridFactory
            Return template
        End Function

        Private Sub OnListViewItemDoubleClick(sender As Object, e As MouseButtonEventArgs)
            If Not sender Is Nothing AndAlso TypeOf sender Is ListViewItem Then
                Dim item As Item = CType(sender, ListViewItem).DataContext
                If Not item Is Nothing Then
                    If TypeOf item Is Folder Then
                        _view.LogicalParent = Me.Folder
                        Me.FolderName = item.FullPath
                    Else
                        Dim iContextMenu As IContextMenu, defaultId As String
                        Dim contextMenu As ContextMenu = Me.Folder.GetContextMenu(Me.SelectedItems, iContextMenu, defaultId)
                        Me.Folder.InvokeCommand(iContextMenu, Me.SelectedItems, defaultId)
                    End If
                End If
            End If
        End Sub

        Private Sub OnListViewItemRightClick(sender As Object, e As MouseButtonEventArgs)
            If Not e.OriginalSource Is Nothing Then
                Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                Dim contextItems As IEnumerable(Of Item)
                If Not listViewItem Is Nothing Then
                    If Not listViewItem.IsSelected Then Me.SetSelectedItem(listViewItem.DataContext)
                    contextItems = Me.SelectedItems
                End If

                Dim iContextMenu As IContextMenu, defaultId As String
                Dim contextMenu As ContextMenu = Me.Folder.GetContextMenu(contextItems, iContextMenu, defaultId)
                Dim wireItems As Action(Of ItemCollection) =
                        Sub(items As ItemCollection)
                            For Each c As Control In items
                                If TypeOf c Is MenuItem AndAlso CType(c, MenuItem).Items.Count = 0 Then
                                    Dim menuItem As MenuItem = c
                                    AddHandler menuItem.Click,
                                        Sub(s2 As Object, e2 As EventArgs)
                                            Dim isHandled As Boolean = False

                                            Select Case menuItem.Tag.ToString().Split(vbTab)(1)
                                                Case "open"
                                                    If Not Me.SelectedItem Is Nothing AndAlso TypeOf Me.SelectedItem Is Folder Then
                                                        Me.FolderName = Me.SelectedItem.FullPath
                                                        isHandled = True
                                                    End If
                                            End Select

                                            If Not isHandled Then
                                                Me.Folder.InvokeCommand(iContextMenu, Me.SelectedItems, menuItem.Tag)
                                            End If
                                        End Sub
                                ElseIf TypeOf c Is MenuItem Then
                                    wireItems(CType(c, MenuItem).Items)
                                End If
                            Next
                        End Sub
                wireItems(contextMenu.Items)

                _view.listView.ContextMenu = contextMenu
            End If
        End Sub

        Private Function getScrollViewer() As ScrollViewer
            Return UIHelper.FindVisualChildren(Of ScrollViewer)(_view.listView)(0)
        End Function

        Protected Overridable Async Function OnSelectionChanged() As Task

        End Function

        Public Overridable ReadOnly Property SelectedItems As IEnumerable(Of Item)
            Get
                If Not _selectionHelper Is Nothing Then
                    Return _selectionHelper.SelectedItems
                Else
                    Return {}
                End If
            End Get
        End Property

        Public Sub SetSelectedItems(value As IEnumerable(Of Item))
            If Not _selectionHelper Is Nothing Then
                _selectionHelper.SetSelectedItems(value)
            End If
        End Sub

        Public ReadOnly Property SelectedItem As Item
            Get
                Dim selectedItems As IEnumerable(Of Item) = Me.SelectedItems
                Return If(Not selectedItems Is Nothing AndAlso selectedItems.Count = 1, selectedItems(0), Nothing)
            End Get
        End Property

        Public Sub SetSelectedItem(value As Item)
            If value Is Nothing Then
                Me.SetSelectedItems(New Item() {})
            Else
                Me.SetSelectedItems(New Item() {value})
            End If
        End Sub

        Private Class ScrollState
            Public Property OffsetY As Double
            Public Property SelectedItems As IEnumerable(Of Item)
        End Class
    End Class
End Namespace
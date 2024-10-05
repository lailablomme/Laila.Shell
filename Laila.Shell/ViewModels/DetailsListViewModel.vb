Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers

Namespace ViewModels
    Public Class DetailsListViewModel
        Inherits NotifyPropertyChangedBase

        Private Const LOADINGSPINNER_ANIMATION_DURATION = 100

        Friend _view As DetailsListView
        Private _folderName As String
        Private _folder As Folder
        Private _gridView As GridView
        Private _columnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
        Private _isLoading As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _scrollState As Dictionary(Of String, ScrollState) = New Dictionary(Of String, ScrollState)()
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As ContextMenu = New ContextMenu()

        Public Sub New(view As DetailsListView)
            _view = view

            Dim listViewItemStyle As Style = New Style()
            listViewItemStyle.TargetType = GetType(ListViewItem)
            listViewItemStyle.BasedOn = _view.TryFindResource(GetType(ListViewItem))
            listViewItemStyle.Setters.Add(New Setter(ListViewItem.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch))
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

                    _dropTarget = New ListViewDropTarget(Me)
                    WpfDragTargetProxy.RegisterDragDrop(_view.listView, _dropTarget)
                End Sub

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(_view.listView)
                End Sub

            AddHandler _view.listView.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler _view.listView.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler _view.PreviewKeyDown, AddressOf OnListViewKeyDown
        End Sub

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Set(value As Boolean)
                SetValue(_isLoading, value)

                If Not _isLoading Then
                    If Not Me.Folder Is Nothing AndAlso _scrollState.ContainsKey(Me.Folder.FullPath) Then
                        UIHelper.OnUIThread(
                            Sub()
                                loadScrollState()
                            End Sub)
                    End If
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

        Public ReadOnly Property FolderDisplayName
            Get
                Return If(String.IsNullOrWhiteSpace(Me.Folder?.DisplayName), "(no folder selected)", Me.Folder?.DisplayName)
            End Get
        End Property

        Private Sub folder_LoadingStateChanged(isLoading As Boolean)
            UIHelper.OnUIThread(
                Sub()
                    Me.IsLoading = isLoading
                End Sub)
        End Sub

        Public Property Folder As Folder
            Get
                Return _folder
            End Get
            Set(value As Folder)
                If Not Me.Folder Is Nothing Then
                    saveScrollState()
                    RemoveHandler Me.Folder.LoadingStateChanged, AddressOf folder_LoadingStateChanged
                    UIHelper.OnUIThread(
                        Sub()
                            saveScrollState()
                        End Sub)
                End If

                If Not value Is Nothing AndAlso Not value._items Is Nothing AndAlso value._items.Count = 1 _
                    AndAlso TypeOf value._items(0) Is DummyFolder Then
                    value._items = Nothing
                End If

                If Not value Is Nothing Then
                    AddHandler value.LoadingStateChanged, AddressOf folder_LoadingStateChanged
                End If
                SetValue(_folder, value)
                Me.NotifyOfPropertyChange("Items")
                Me.NotifyOfPropertyChange("FolderDisplayName")

                If Not value Is Nothing Then
                    Me.ColumnsIn = buildColumnsIn()
                End If

                _view.Folder = Me.Folder
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
                gvc.SetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                If column.CanonicalName = "System.ItemNameDisplay" Then
                    gvc.SetValue(Behaviors.GridViewExtBehavior.SortPropertyNameProperty, "ItemNameDisplaySortValue")
                End If
                gvc.SetValue(Behaviors.GridViewExtBehavior.GroupByPropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString()))
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
                        .Binding = New Binding(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())) With
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
                    Dim dataTrigger3 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger3.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger3)
                    imageFactory1.SetValue(Image.StyleProperty, style)
                    Return imageFactory1
                End Function
            gridFactory.AppendChild(getIconFactory("Icon[16]"))
            gridFactory.AppendChild(getIconFactory("OverlaySmall"))

            If [property].HasIcon Then
                Dim imageFactory2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                imageFactory2.SetValue(Grid.ColumnProperty, 1)
                imageFactory2.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                imageFactory2.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                imageFactory2.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                imageFactory2.SetValue(Image.SourceProperty, New Binding() With {
                    .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Icon16", column.PROPERTYKEY.ToString())),
                    .Mode = BindingMode.OneWay
                })
                Dim imageStyle As Style = New Style(GetType(Image))
                imageStyle.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                Dim imageDataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                imageDataTrigger1.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                imageStyle.Triggers.Add(imageDataTrigger1)
                imageFactory2.SetValue(Image.StyleProperty, imageStyle)
                gridFactory.AppendChild(imageFactory2)
            End If

            Dim textBlockFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(TextBlock))
            textBlockFactory.SetValue(Grid.ColumnProperty, 2)
            textBlockFactory.SetValue(TextBlock.TextAlignmentProperty, column.Alignment)
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
            textBlockFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis)
            textBlockFactory.SetValue(TextBlock.TextProperty, New Binding() With {
                .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString())),
                .Mode = BindingMode.OneWay
            })
            Dim textBlockStyle As Style = New Style(GetType(TextBlock))
            textBlockStyle.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Black))
            textBlockStyle.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(1)))
            Dim textBlockDataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger1.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(0.5)))
            textBlockStyle.Triggers.Add(textBlockDataTrigger1)
            Dim textBlockDataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCompressed") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger2.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Blue))
            textBlockStyle.Triggers.Add(textBlockDataTrigger2)
            textBlockFactory.SetValue(Image.StyleProperty, textBlockStyle)
            gridFactory.AppendChild(textBlockFactory)

            template.VisualTree = gridFactory
            Return template
        End Function

        Private Sub OnListViewKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.C AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) AndAlso Me.SelectedItems.Count > 0 Then
                Clipboard.CopyFiles(Me.SelectedItems)
            ElseIf e.Key = Key.X AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) AndAlso Me.SelectedItems.Count > 0 Then
                Clipboard.CutFiles(Me.SelectedItems)
            End If
        End Sub

        Private Sub OnListViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            If Not _view.selection.IsSelecting Then
                If Not _mouseItemDown Is Nothing AndAlso Me.SelectedItems.Count > 0 AndAlso
                            (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                    Dim currentPointDown As Point = e.GetPosition(_view)
                    If Math.Abs(currentPointDown.X - _mousePointDown.X) > 5 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 5 Then
                        Drag.Start(Me.SelectedItems, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                    End If
                End If
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            If Not _view.selection.IsSelecting Then
                _mousePointDown = e.GetPosition(_view)

                ' this prevents a multiple selection getting replaced by the single clicked item
                If Not e.OriginalSource Is Nothing Then
                    Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                    Dim clickedItem As Item = listViewItem?.DataContext
                    _mouseItemDown = clickedItem
                    If clickedItem Is Nothing Then
                        _view.listView.Focus()
                    End If
                    If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 AndAlso Me.SelectedItems.Contains(clickedItem) Then
                        If TypeOf clickedItem Is Folder Then
                            CType(clickedItem, Folder).IsSelected = True
                            Me.Folder = clickedItem
                        Else
                            _menu = New ContextMenu()
                            _menu.GetContextMenu(Me.Folder, Me.SelectedItems, False)
                            _menu.InvokeCommand(_menu.DefaultId)
                        End If
                    ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                        If Me.SelectedItems.Count > 0 AndAlso Me.SelectedItems.Contains(clickedItem) Then
                            e.Handled = True
                        End If
                    ElseIf e.RightButton = MouseButtonState.Pressed AndAlso
                        UIHelper.GetParentOfType(Of Primitives.ScrollBar)(e.OriginalSource) Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of GridViewHeaderRowPresenter)(e.OriginalSource) Is Nothing Then

                        If (Me.SelectedItems.Count = 0 OrElse Not Me.SelectedItems.Contains(clickedItem)) _
                            AndAlso Not clickedItem Is Nothing Then Me.SetSelectedItem(clickedItem)

                        _menu = New ContextMenu()
                        AddHandler _menu.Click,
                            Sub(id As Integer, verb As String, ByRef isHandled As Boolean)
                                Select Case verb
                                    Case "open"
                                        If Not Me.SelectedItem Is Nothing AndAlso TypeOf Me.SelectedItem Is Folder Then
                                            Me.Folder = Me.SelectedItem
                                            isHandled = True
                                        End If
                                End Select
                            End Sub

                        _view.listView.ContextMenu = _menu.GetContextMenu(Me.Folder, Me.SelectedItems, False)
                        e.Handled = True
                    ElseIf clickedItem Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of System.Windows.Controls.Primitives.ScrollBar)(e.OriginalSource) Is Nothing Then
                        Me.SetSelectedItem(Nothing)
                    End If
                Else
                    _mouseItemDown = Nothing
                End If
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
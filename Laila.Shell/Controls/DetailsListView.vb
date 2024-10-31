Imports System.ComponentModel
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography.X509Certificates
Imports System.Threading
Imports System.Windows
Imports System.Windows.Annotations
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shell
Imports System.Xml
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.SevenZip

Namespace Controls
    Public Class DetailsListView
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(DetailsListView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(Behaviors.GridViewExtBehavior.ColumnsInData), GetType(DetailsListView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(DetailsListView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend PART_ListView As ListView
        Private PART_Selection As Laila.Shell.Behaviors.SelectionBehavior
        Private PART_Grid As Grid
        Private _columnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
        Private _isLoading As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _scrollState As Dictionary(Of String, ScrollState) = New Dictionary(Of String, ScrollState)()
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As Laila.Shell.ContextMenu
        Private _timeSpentTimer As Timer

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(DetailsListView), New FrameworkPropertyMetadata(GetType(DetailsListView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_ListView = Template.FindName("PART_ListView", Me)
            PART_Selection = Template.FindName("PART_Selection", Me)
            PART_Grid = Template.FindName("PART_Grid", Me)

            Dim listViewItemStyle As Style = New Style()
            listViewItemStyle.TargetType = GetType(ListViewItem)
            listViewItemStyle.BasedOn = PART_ListView.TryFindResource(GetType(ListViewItem))
            listViewItemStyle.Setters.Add(New Setter(ListViewItem.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch))
            PART_ListView.Resources.Add(GetType(ListViewItem), listViewItemStyle)

            AddHandler PART_ListView.Loaded,
                Sub(s As Object, e As EventArgs)
                    _selectionHelper = New SelectionHelper(Of Item)(PART_ListView)
                    _selectionHelper.SelectionChanged =
                        Async Function() As Task
                            'Await Me.OnSelectionChanged()
                        End Function

                    _dropTarget = New ListViewDropTarget(Me)
                    WpfDragTargetProxy.RegisterDragDrop(PART_ListView, _dropTarget)
                End Sub

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(PART_ListView)
                End Sub

            AddHandler PART_ListView.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler PART_ListView.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler PART_ListView.PreviewMouseUp, AddressOf OnListViewPreviewMouseButtonUp
            AddHandler PART_ListView.MouseLeave, AddressOf OnListViewMouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf OnListViewKeyDown
        End Sub

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
                    gvc.SetValue(Behaviors.GridViewExtBehavior.CanHideProperty, False)
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
            textBlockFactory.SetValue(TextBlock.PaddingProperty, New Thickness(0, 0, 2, 0))
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
            If Not PART_Selection.IsSelecting Then
                If Not _mouseItemDown Is Nothing AndAlso Me.SelectedItems.Count > 0 AndAlso
                            (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                    Dim currentPointDown As Point = e.GetPosition(Me)
                    If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                        Drag.Start(Me.SelectedItems, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                    End If
                End If
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me)

            If Not PART_Selection.IsSelecting Then
                ' this prevents a multiple selection getting replaced by the single clicked item
                If Not e.OriginalSource Is Nothing Then
                    Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                    Dim clickedItem As Item = listViewItem?.DataContext
                    _mouseItemDown = clickedItem
                    If clickedItem Is Nothing Then
                        PART_ListView.Focus()
                    Else
                        listViewItem.Focus()
                    End If
                    If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 AndAlso Me.SelectedItems.Contains(clickedItem) Then
                        If TypeOf clickedItem Is Folder Then
                            Me.Folder = clickedItem
                        Else
                            _menu = New Laila.Shell.ContextMenu()
                            _menu.GetContextMenu(Me.Folder, Me.SelectedItems, False)
                            _menu.InvokeCommand(_menu.DefaultId)
                        End If
                    ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                        If Me.SelectedItems.Count > 0 AndAlso Me.SelectedItems.Contains(clickedItem) _
                            AndAlso Keyboard.Modifiers = ModifierKeys.None Then
                            e.Handled = True
                        End If
                    ElseIf e.RightButton = MouseButtonState.Pressed AndAlso
                        UIHelper.GetParentOfType(Of Primitives.ScrollBar)(e.OriginalSource) Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of GridViewHeaderRowPresenter)(e.OriginalSource) Is Nothing Then

                        If (Me.SelectedItems.Count = 0 OrElse Not Me.SelectedItems.Contains(clickedItem)) _
                            AndAlso Not clickedItem Is Nothing Then
                            Me.SetSelectedItem(clickedItem)
                        ElseIf clickedItem Is Nothing Then
                            Me.SetSelectedItem(Nothing)
                        End If

                        _menu = New Laila.Shell.ContextMenu()
                        AddHandler _menu.CommandInvoked,
                            Sub(s As Object, e2 As CommandInvokedEventArgs)
                                Select Case e2.Verb
                                    Case "open"
                                        If Not Me.SelectedItem Is Nothing AndAlso TypeOf Me.SelectedItem Is Folder Then
                                            Me.Folder = clickedItem
                                            e2.IsHandled = True
                                        End If
                                    Case "rename"
                                        Dim column As Column = Me.Folder.Columns("System.ItemNameDisplay")
                                        If Not column Is Nothing Then
                                            Dim headers As IEnumerable(Of GridViewColumnHeader) =
                                                UIHelper.FindVisualChildren(Of GridViewColumnHeader)(Me.PART_ListView)
                                            Dim header As GridViewColumnHeader =
                                                headers.FirstOrDefault(Function(h) Not h.Column Is Nothing _
                                                    AndAlso h.Column.GetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty) _
                                                        = String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                                            If Not header Is Nothing Then
                                                Dim width As Double = header.ActualWidth
                                                Dim ptLeft As Point = Me.PointFromScreen(header.PointToScreen(New Point(0, 0)))
                                                If header.Column.GetValue(Behaviors.GridViewExtBehavior.ColumnIndexProperty) = 0 Then
                                                    ptLeft.X += 20
                                                    width -= 20
                                                End If
                                                Dim ptTop As Point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
                                                _menu.DoRename(New Point(ptLeft.X + 5, ptTop.Y + 1), width - 5, clickedItem, Me.PART_Grid)
                                            End If
                                        End If
                                        e2.IsHandled = True
                                    Case "laila.shell.(un)pin"
                                        If e2.IsChecked Then
                                            PinnedItems.PinItem(clickedItem.FullPath)
                                        Else
                                            PinnedItems.UnpinItem(clickedItem.FullPath)
                                        End If
                                        e2.IsHandled = True
                                End Select
                            End Sub

                        Dim contextMenu As Controls.ContextMenu = _menu.GetContextMenu(Me.Folder, Me.SelectedItems, False)
                        PART_ListView.ContextMenu = contextMenu
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

        Public Sub OnListViewPreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Public Sub OnListViewMouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
        End Sub

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

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Public Property IsLoading As Boolean
            Get
                Return GetValue(IsLoadingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsLoadingProperty, value)
            End Set
        End Property

        Public Property ColumnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
            Get
                Return GetValue(ColumnsInProperty)
            End Get
            Set(ByVal value As Behaviors.GridViewExtBehavior.ColumnsInData)
                SetCurrentValue(ColumnsInProperty, value)
            End Set
        End Property

        Private Function getScrollViewer() As ScrollViewer
            Return UIHelper.FindVisualChildren(Of ScrollViewer)(PART_ListView)(0)
        End Function

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

        Friend Async Sub OnFolderChangedLocal(oldValue As Folder, newValue As Folder)
            If Not _timeSpentTimer Is Nothing Then _timeSpentTimer.Dispose()

            If Not oldValue Is Nothing Then
                RemoveHandler oldValue.PropertyChanged, AddressOf folder_PropertyChanged
                Dim view As CollectionView = CollectionViewSource.GetDefaultView(oldValue.Items)
                view.SortDescriptions.Clear()
            End If

            If Not newValue Is Nothing Then
                Await Task.Delay(45)
                Dim func As Func(Of Task) =
                    Async Function() As Task
                        FrequentFolders.Track(newValue)
                    End Function
                Task.Run(func)
                _timeSpentTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                FrequentFolders.RecordTimeSpent(Me.Folder, 2)
                            End Sub)
                    End Sub), Nothing, 1000 * 60 * 2, 1000 * 60 * 2)
                BindingOperations.ClearBinding(Me.PART_ListView, ListView.ItemsSourceProperty)
                CType(Me.PART_ListView.View, GridView).Columns.Clear()
                Await newValue.GetItemsAsync()
                Me.ColumnsIn = buildColumnsIn()
                BindingOperations.SetBinding(Me.PART_ListView, ListView.ItemsSourceProperty, New Binding("Folder.Items") With {.Source = Me})
                AddHandler newValue.PropertyChanged, AddressOf folder_PropertyChanged
            End If

            Me.IsLoading = False
        End Sub

        Private Sub folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "IsRefreshingItems"
                    UIHelper.OnUIThread(
                        Sub()
                            Me.IsLoading = CType(s, Folder).IsRefreshingItems
                        End Sub)
            End Select
        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As DetailsListView = TryCast(d, DetailsListView)
            dlv.IsLoading = True
            dlv.OnFolderChangedLocal(e.OldValue, e.NewValue)
        End Sub

        Private Class ScrollState
            Public Property OffsetY As Double
            Public Property SelectedItems As IEnumerable(Of Item)
        End Class
    End Class
End Namespace
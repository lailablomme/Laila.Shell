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
    Public MustInherit Class BaseFolderView
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(Behaviors.GridViewExtBehavior.ColumnsInData), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend PART_ListView As System.Windows.Controls.ListView
        Private PART_Grid As Grid
        Private PART_StackPanel As Panel
        Private _columnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
        Private _isLoading As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _mouseItemOver As Item
        Private _menu As Laila.Shell.ContextMenu
        Private _timeSpentTimer As Timer
        Private _scrollViewer As ScrollViewer
        Private _lastScrollOffset As Point
        Private _lastScrollSize As Size
        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseFolderView), New FrameworkPropertyMetadata(GetType(BaseFolderView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ListView = Template.FindName("PART_ListView", Me)
            Me.PART_Grid = Template.FindName("PART_Grid", Me)

            If Not Me.Folder Is Nothing Then
                Me.MakeBinding()
            End If

            AddHandler PART_ListView.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        Me.PART_StackPanel = UIHelper.FindVisualChildren(Of Panel)(Me.PART_ListView).FirstOrDefault(Function(c) c.Name = "PART_StackPanel")
                        Me.PART_StackPanel.Visibility = Visibility.Hidden

                        _selectionHelper = New SelectionHelper(Of Item)(Me.PART_ListView)
                        _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListView)(0)

                        AddHandler _scrollViewer.ScrollChanged,
                        Sub(s2 As Object, e2 As ScrollChangedEventArgs)
                            If Not Me.Folder Is Nothing Then
                                _lastScrollOffset = New Point(_scrollViewer.HorizontalOffset, _scrollViewer.VerticalOffset)
                                _lastScrollSize = New Size(_scrollViewer.ScrollableWidth, _scrollViewer.ScrollableHeight)
                            End If
                            UIHelper.OnUIThreadAsync(
                                Sub()
                                    GC.Collect()
                                End Sub)
                        End Sub
                    End If
                End Sub

            AddHandler Me.PART_ListView.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler Me.PART_ListView.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler Me.PART_ListView.PreviewMouseUp, AddressOf OnListViewPreviewMouseButtonUp
            AddHandler Me.PART_ListView.MouseLeave, AddressOf OnListViewMouseLeave
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
                        .Mode = BindingMode.OneWay,
                        .IsAsync = True
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
            gridFactory.AppendChild(getIconFactory("IconAsync[16]"))
            gridFactory.AppendChild(getIconFactory("OverlaySmallAsync"))

            If [property].HasIcon Then
                Dim imageFactory2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                imageFactory2.SetValue(Grid.ColumnProperty, 1)
                imageFactory2.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                imageFactory2.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                imageFactory2.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                imageFactory2.SetValue(Image.SourceProperty, New Binding() With {
                    .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Icon16Async", column.PROPERTYKEY.ToString())),
                    .Mode = BindingMode.OneWay,
                    .IsAsync = True
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
            Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
            Dim overItem As Item = TryCast(listViewItem?.DataContext, Item)
            If Not overItem Is Nothing AndAlso Not overItem.Equals(_mouseItemOver) Then
                Dim toolTip As String = overItem.InfoTip
                listViewItem.ToolTip = If(String.IsNullOrWhiteSpace(toolTip), Nothing, toolTip)
                _mouseItemOver = overItem
            End If

            If Not _mouseItemDown Is Nothing AndAlso Me.SelectedItems.Count > 0 AndAlso
                (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(Me)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start(Me.SelectedItems, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me)

            ' this prevents a multiple selection getting replaced by the single clicked item
            If Not e.OriginalSource Is Nothing Then
                Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listViewItem?.DataContext, Item)
                _mouseItemDown = clickedItem
                If clickedItem Is Nothing Then
                    Me.PART_ListView.Focus()
                Else
                    listViewItem.Focus()
                End If
                If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 AndAlso Me.SelectedItems.Contains(clickedItem) Then
                    If TypeOf clickedItem Is Folder Then
                        CType(clickedItem, Folder).LastScrollOffset = New Point()
                        Me.Folder = clickedItem
                    Else
                        _menu = New Laila.Shell.ContextMenu()
                        _menu.GetContextMenu(Me.Folder, Me.SelectedItems, True)
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
                            If e2.Verb.StartsWith("laila.shell.view.") Then
                                Me.Host.View = e2.Verb.Substring("laila.shell.view.".Length)
                                e2.IsHandled = True
                            Else
                                Select Case e2.Verb
                                    Case "open"
                                        If Not Me.SelectedItem Is Nothing AndAlso TypeOf Me.SelectedItem Is Folder Then
                                            Me.Folder = clickedItem
                                            e2.IsHandled = True
                                        End If
                                    Case "rename"
                                        Dim point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double
                                        Me.GetItemNameCoordinates(listViewItem, textAlignment, point, size, fontSize)
                                        _menu.DoRename(point, size, textAlignment, fontSize, clickedItem, Me.PART_Grid)
                                        e2.IsHandled = True
                                    Case "laila.shell.(un)pin"
                                        If e2.IsChecked Then
                                            PinnedItems.PinItem(clickedItem.FullPath)
                                        Else
                                            PinnedItems.UnpinItem(clickedItem.FullPath)
                                        End If
                                        e2.IsHandled = True
                                End Select
                            End If
                        End Sub

                    Dim contextMenu As Controls.ContextMenu = _menu.GetContextMenu(Me.Folder, Me.SelectedItems, False)
                    Me.PART_ListView.ContextMenu = contextMenu
                    e.Handled = True
                ElseIf clickedItem Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of System.Windows.Controls.Primitives.ScrollBar)(e.OriginalSource) Is Nothing Then
                    Me.SetSelectedItem(Nothing)
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Protected MustOverride Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                          ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)

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

        Public Property Host As FolderView

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

        Friend Async Sub OnFolderChangedLocal(oldValue As Folder, newValue As Folder)
            If Not _timeSpentTimer Is Nothing Then _timeSpentTimer.Dispose()

            If Not oldValue Is Nothing Then
                RemoveHandler oldValue.PropertyChanged, AddressOf folder_PropertyChanged
                oldValue.LastScrollOffset = _lastScrollOffset
                oldValue.LastScrollSize = _lastScrollSize
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
                ClearBinding()
                Await newValue.GetItemsAsync()
                Me.MakeBinding()
                AddHandler newValue.PropertyChanged, AddressOf folder_PropertyChanged
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        Await Task.Delay(5)
                        _lastScrollOffset = newValue.LastScrollOffset
                        _lastScrollSize = newValue.LastScrollSize
                        _scrollViewer.ScrollToHorizontalOffset(If(_lastScrollSize.Width = 0, 0, _lastScrollOffset.X * _scrollViewer.ScrollableWidth / _lastScrollSize.Width))
                        _scrollViewer.ScrollToVerticalOffset(If(_lastScrollSize.Height = 0, 0, _lastScrollOffset.Y * _scrollViewer.ScrollableHeight / _lastScrollSize.Height))
                        If Not Me.PART_StackPanel Is Nothing Then
                            Me.PART_StackPanel.Visibility = Visibility.Visible
                        End If
                    End Sub, Threading.DispatcherPriority.ContextIdle)
            End If

            Me.IsLoading = False
        End Sub

        Protected Overridable Sub ClearBinding()
            If Not Me.PART_ListView Is Nothing Then
                BindingOperations.ClearBinding(Me.PART_ListView, System.Windows.Controls.ListView.ItemsSourceProperty)
            End If
            If Not Me.PART_StackPanel Is Nothing Then
                Me.PART_StackPanel.Visibility = Visibility.Hidden
            End If
        End Sub

        Protected Overridable Sub MakeBinding()
            If Not Me.PART_ListView Is Nothing Then
                BindingOperations.SetBinding(Me.PART_ListView, System.Windows.Controls.ListView.ItemsSourceProperty, New Binding("Folder.Items") With {.Source = Me})
            End If
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
            Dim dlv As BaseFolderView = TryCast(d, BaseFolderView)
            dlv.IsLoading = True
            dlv.OnFolderChangedLocal(e.OldValue, e.NewValue)
        End Sub

        Private Class ScrollState
            Public Property OffsetY As Double
            Public Property SelectedItems As IEnumerable(Of Item)
        End Class
    End Class
End Namespace
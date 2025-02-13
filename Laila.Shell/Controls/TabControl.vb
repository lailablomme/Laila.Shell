Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports Laila.Shell.Helpers

Namespace Controls
    <TemplatePart(Name:="PART_ItemsHolder", Type:=GetType(Panel))>
    Public Class TabControl
        Inherits System.Windows.Controls.TabControl
        Implements IDisposable

        Private PART_ItemsHolder As Panel
        Private PART_LeftButton As RepeatButton
        Private PART_RightButton As RepeatButton
        Private PART_ScrollViewer As ScrollViewer
        Private PART_PanelTop As Panel
        Private PART_AddTabButton As Button
        Private PART_TabGrid As Grid
        Private PART_TabColumn As ColumnDefinition
        Private _items As ObservableCollection(Of TabData) = New ObservableCollection(Of TabData)()
        Private _dropTarget As IDropTarget
        Private _isLoaded As Boolean
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TabControl), New FrameworkPropertyMetadata(GetType(TabControl)))
        End Sub

        Public Sub New()
            Shell.AddToControlCache(Me)

            AddHandler ItemContainerGenerator.StatusChanged, AddressOf ItemContainerGeneratorStatusChanged

            _items.Add(New TabData())
            Me.ItemsSource = _items

            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        setDesiredTabWidth()

                        focusChildren()

                        _dropTarget = New TabControlDropTarget(Me)
                        WpfDragTargetProxy.RegisterDragDrop(Me, _dropTarget)
                    End If
                End Sub
        End Sub

        Protected Overrides Sub OnRenderSizeChanged(sizeInfo As SizeChangedInfo)
            MyBase.OnRenderSizeChanged(sizeInfo)

            setDesiredTabWidth()
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_LeftButton = GetTemplateChild("TabLeftButtonTop")
            PART_RightButton = GetTemplateChild("TabRightButtonTop")
            PART_ScrollViewer = GetTemplateChild("TabScrollViewerTop")
            PART_PanelTop = GetTemplateChild("HeaderPanel")
            PART_AddTabButton = GetTemplateChild("PART_AddTabButton")
            PART_ItemsHolder = GetTemplateChild("PART_ItemsHolder")
            PART_TabGrid = GetTemplateChild("PART_TabGrid")
            PART_TabColumn = GetTemplateChild("PART_TabColumn")

            AddHandler PART_LeftButton.Click, AddressOf tabLeftButton_Click
            AddHandler PART_RightButton.Click, AddressOf tabRightButton_Click
            AddHandler PART_ScrollViewer.Loaded,
                Sub(sender As Object, e As EventArgs)
                    updateScrollButtonsAvailability()
                End Sub
            AddHandler PART_ScrollViewer.ScrollChanged,
                Sub(sender As Object, e As EventArgs)
                    updateScrollButtonsAvailability()
                End Sub
            AddHandler Me.SelectionChanged,
                Sub(sender As Object, e As EventArgs)
                    scrollToSelectedItem()
                End Sub
            AddHandler Me.PART_AddTabButton.Click,
                Sub(s As Object, e As EventArgs)
                    _items.Add(New TabData())
                    Me.SelectedItem = _items(_items.Count - 1)
                End Sub

            updateSelectedItem()

            Dim tabs As List(Of TabItem) = UIHelper.FindVisualChildren(Of TabItem)(Me).ToList()

            For Each t In tabs
                createChildContentPresenter(t)
            Next
        End Sub

        Private Sub ItemContainerGeneratorStatusChanged(sender As Object, e As EventArgs)
            If ItemContainerGenerator.Status = Primitives.GeneratorStatus.ContainersGenerated Then
                RemoveHandler ItemContainerGenerator.StatusChanged, AddressOf ItemContainerGeneratorStatusChanged
                updateSelectedItem()
            End If
        End Sub

        Protected Overrides Sub OnItemsChanged(e As Specialized.NotifyCollectionChangedEventArgs)
            MyBase.OnItemsChanged(e)

            If PART_ItemsHolder Is Nothing Then
                Return
            End If

            Select Case e.Action
                Case Specialized.NotifyCollectionChangedAction.Reset
                    PART_ItemsHolder.Children.Clear()

                    If Me.SelectedItem Is Nothing AndAlso Not Me.Items Is Nothing AndAlso Me.Items.Count > 0 Then
                        Me.SelectedItem = Me.Items(0)
                    End If

                    updateSelectedItem()
                Case Specialized.NotifyCollectionChangedAction.Add, Specialized.NotifyCollectionChangedAction.Remove
                    If Not e.OldItems Is Nothing Then
                        For Each item In e.OldItems
                            Dim cp As ContentPresenter = findChildContentPresenter(item)
                            If Not cp Is Nothing Then
                                For Each treeView In UIHelper.FindVisualChildren(Of TreeView)(cp)
                                    treeView.Dispose()
                                Next
                                For Each folderView In UIHelper.FindVisualChildren(Of FolderView)(cp)
                                    folderView.Dispose()
                                Next
                                For Each addressBar In UIHelper.FindVisualChildren(Of AddressBar)(cp)
                                    addressBar.Dispose()
                                Next
                                PART_ItemsHolder.Children.Remove(cp)
                            End If
                        Next
                    End If
                    setDesiredTabWidth()
                    updateSelectedItem()
                Case Specialized.NotifyCollectionChangedAction.Replace
                    Throw New NotImplementedException()
            End Select
        End Sub

        Private Sub focusChildren()
            Dim item As TabItem = GetSelectedTabItem()
            If Not item Is Nothing Then
                Dim cp As ContentPresenter = findChildContentPresenter(item)
                Dim folderView As FolderView = UIHelper.FindVisualChildren(Of FolderView)(cp)?(0)
                If Not folderView Is Nothing _
                    AndAlso Not folderView.ActiveView Is Nothing _
                    AndAlso Not folderView.ActiveView.PART_ListBox Is Nothing Then
                    folderView.ActiveView.PART_ListBox.Focus()
                Else
                    Dim treeView As TreeView = UIHelper.FindVisualChildren(Of TreeView)(cp)?(0)
                    If Not treeView Is Nothing _
                        AndAlso Not treeView.PART_ListBox Is Nothing Then
                        treeView.PART_ListBox.Focus()
                    End If
                End If
            End If
        End Sub

        Protected Overrides Sub OnSelectionChanged(e As SelectionChangedEventArgs)
            MyBase.OnSelectionChanged(e)
            updateSelectedItem()
        End Sub

        Private Sub updateSelectedItem()
            If PART_ItemsHolder Is Nothing Then
                Return
            End If

            Dim item As TabItem = GetSelectedTabItem()
            If Not item Is Nothing Then
                createChildContentPresenter(item)
            End If

            For Each child As ContentPresenter In PART_ItemsHolder.Children
                child.Visibility = IIf(CType(child.Tag, TabItem).IsSelected, Visibility.Visible, Visibility.Hidden)
            Next

            focusChildren()
        End Sub

        Private Function createChildContentPresenter(item As Object) As ContentPresenter
            If item Is Nothing Then
                Return Nothing
            End If

            Dim cp As ContentPresenter = findChildContentPresenter(item)

            If Not cp Is Nothing Then
                Return cp
            End If

            Dim tabItem As TabItem = item
            cp = New ContentPresenter() With {
                .Content = IIf(Not tabItem Is Nothing, tabItem.Content, item),
                .ContentTemplate = Me.SelectedContentTemplate,
                .ContentTemplateSelector = Me.SelectedContentTemplateSelector,
                .ContentStringFormat = Me.SelectedContentStringFormat,
                .Visibility = Visibility.Hidden,
                .Tag = IIf(Not tabItem Is Nothing, tabItem, Me.ItemContainerGenerator.ContainerFromItem(item))
            }
            PART_ItemsHolder.Children.Add(cp)

            AddHandler tabItem.Loaded,
                Sub(s As Object, e As EventArgs)
                    setDesiredTabWidth()
                End Sub

            Return cp
        End Function

        Private Function findChildContentPresenter(data As Object) As ContentPresenter
            If TypeOf data Is TabItem Then
                data = CType(data, TabItem).Content
            End If

            If data Is Nothing Then
                Return Nothing
            End If

            If PART_ItemsHolder Is Nothing Then
                Return Nothing
            End If

            For Each cp As ContentPresenter In PART_ItemsHolder.Children
                If cp.Content.Equals(data) Then
                    Return cp
                End If
            Next

            Return Nothing
        End Function

        Protected Function GetSelectedTabItem() As TabItem
            If Me.SelectedItem Is Nothing Then
                Return Nothing
            End If

            If TypeOf Me.SelectedItem Is TabItem Then
                Return Me.SelectedItem
            Else
                Return Me.ItemContainerGenerator.ContainerFromIndex(Me.SelectedIndex)
            End If
        End Function

        Private Sub tabLeftButton_Click(sender As Object, e As RoutedEventArgs)
            PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset - 25)
        End Sub

        Private Sub tabRightButton_Click(sender As Object, e As RoutedEventArgs)
            PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + 25)
        End Sub

        Private Sub scrollToSelectedItem()
            Dim item As TabItem = Me.ItemContainerGenerator.ContainerFromItem(MyBase.SelectedItem)
            scrollToItem(item)
        End Sub

        Private Sub scrollToItem(item As TabItem)
            If Not item Is Nothing Then
                Dim tabItems As List(Of TabItem) = Me.Items.Cast(Of Object).Select(Function(i) CType(Me.ItemContainerGenerator.ContainerFromItem(i), TabItem)).ToList()
                Dim leftItems As List(Of TabItem) = tabItems.Where(Function(i) Not i Is Nothing).TakeWhile(Function(i) Not i.Equals(item)).ToList()
                Dim leftItemsWidth As Double = leftItems.Sum(Function(i) i.ActualWidth)
                If leftItemsWidth + item.ActualWidth > PART_ScrollViewer.HorizontalOffset + PART_ScrollViewer.ViewportWidth Then
                    Dim currentHorizontalOffset As Double = leftItemsWidth + item.ActualWidth - PART_ScrollViewer.ViewportWidth
                    Dim hMargin As Double = If(Not leftItems.Any(Function(i) i.IsSelected) AndAlso Not item.IsSelected, PART_PanelTop.Margin.Left + PART_PanelTop.Margin.Right, 0)
                    currentHorizontalOffset += hMargin
                    PART_ScrollViewer.ScrollToHorizontalOffset(currentHorizontalOffset)
                ElseIf leftItemsWidth < PART_ScrollViewer.HorizontalOffset Then
                    Dim currentHorizontalOffset As Double = leftItemsWidth
                    Dim hMargin As Double = If(leftItems.Any(Function(i) i.IsSelected), PART_PanelTop.Margin.Left + PART_PanelTop.Margin.Right, 0)
                    currentHorizontalOffset += hMargin
                    PART_ScrollViewer.ScrollToHorizontalOffset(currentHorizontalOffset)
                End If
            End If
        End Sub

        Private Sub updateScrollButtonsAvailability()
            Dim hOffset As Double = PART_ScrollViewer.HorizontalOffset
            hOffset = Math.Max(hOffset, 0)

            Dim scrWidth As Double = PART_ScrollViewer.ScrollableWidth
            scrWidth = Math.Max(scrWidth, 0)

            PART_LeftButton.Visibility = If(scrWidth = 0, Visibility.Collapsed, Visibility.Visible)
            PART_LeftButton.IsEnabled = hOffset > 0

            PART_RightButton.Visibility = If(scrWidth = 0, Visibility.Collapsed, Visibility.Visible)
            PART_RightButton.IsEnabled = hOffset < scrWidth
        End Sub

        Private Sub setDesiredTabWidth()
            Dim allTabsMaxWidth As Double = Me.ActualWidth - 175 - Me.PART_LeftButton.Width - Me.PART_RightButton.Width - Me.PART_AddTabButton.Width
            allTabsMaxWidth = If(allTabsMaxWidth < 50, 50, allTabsMaxWidth)

            Dim singleTabWidth As Double = allTabsMaxWidth / _items.Count
            singleTabWidth = If(singleTabWidth < 100, 100, singleTabWidth)
            singleTabWidth = If(singleTabWidth > 200, 200, singleTabWidth)
            Dim allTabsWidth As Double = 0

            ' size single tabs
            For i = 0 To _items.Count - 1
                Dim tabItem As TabItem = Me.ItemContainerGenerator.ContainerFromIndex(i)
                If Not tabItem Is Nothing Then
                    For Each item In UIHelper.FindVisualChildren(Of Grid)(tabItem)
                        If item.Name = "PART_Grid" Then
                            item.Width = singleTabWidth
                        End If
                    Next
                End If
            Next

            Me.PART_PanelTop.InvalidateArrange()
            Me.PART_PanelTop.InvalidateMeasure()
            Me.PART_PanelTop.UpdateLayout()
            Me.PART_PanelTop.Measure(New Size(Me.ActualWidth, Me.ActualHeight))
            allTabsWidth = Me.PART_PanelTop.DesiredSize.Width + 4 + Me.PART_PanelTop.Margin.Left + Me.PART_PanelTop.Margin.Right

            ' size all tabs
            Me.PART_TabColumn.Width = New GridLength(If(allTabsWidth > allTabsMaxWidth, allTabsMaxWidth, allTabsWidth) + 5)
        End Sub

        Public Class TabData
            Inherits NotifyPropertyChangedBase

            Private _folder As Folder
            Private _selectedItems As IEnumerable(Of Item)

            Public Sub New()
                Me.Folder = Shell.GetSpecialFolder(SpecialFolders.Home).Clone()
            End Sub

            Public Property Folder As Folder
                Get
                    Return _folder
                End Get
                Set(value As Folder)
                    SetValue(_folder, value)
                End Set
            End Property

            Public Property SelectedItems As IEnumerable(Of Item)
                Get
                    Return _selectedItems
                End Get
                Set(value As IEnumerable(Of Item))
                    SetValue(_selectedItems, value)
                End Set
            End Property
        End Class

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    WpfDragTargetProxy.RevokeDragDrop(Me)
                    Shell.RemoveFromControlCache(Me)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace
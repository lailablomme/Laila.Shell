Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop.DragDrop

Namespace Controls
    <TemplatePart(Name:="PART_ItemsHolder", Type:=GetType(Panel))>
    Public Class TabControl
        Inherits System.Windows.Controls.TabControl
        Implements IDisposable

        Public Shared ReadOnly ModelTypeProperty As DependencyProperty = DependencyProperty.Register("ModelType", GetType(Type), GetType(TabControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_ItemsHolder As Panel
        Private PART_LeftButton As RepeatButton
        Private PART_RightButton As RepeatButton
        Private PART_ScrollViewer As ScrollViewer
        Private PART_PanelTop As Panel
        Private PART_AddTabButton As Button
        Private PART_TabGrid As Grid
        Private PART_TabColumn As ColumnDefinition
        Private _items As ObservableCollection(Of Object) = New ObservableCollection(Of Object)()
        Private _dropTarget As IDropTarget
        Private _isLoaded As Boolean
        Private _isOnce As Boolean = True
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TabControl), New FrameworkPropertyMetadata(GetType(TabControl)))
        End Sub

        Public Sub New()
            Shell.AddToControlCache(Me)

            Me.ItemsSource = _items

            AddHandler ItemContainerGenerator.StatusChanged, AddressOf ItemContainerGeneratorStatusChanged

            EventManager.RegisterClassHandler(GetType(FolderView), FolderView.ViewActivatedEvent,
                                              New RoutedEventHandler(AddressOf folderView_ViewActivated))

            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        setDesiredTabWidth()

                        _dropTarget = New TabControlDropTarget(Me)
                        WpfDragTargetProxy.RegisterDragDrop(Me, _dropTarget)

                        AddHandler Window.GetWindow(Me).Closed,
                            Sub(s2 As Object, e2 As EventArgs)
                                Me.Dispose()
                            End Sub
                    End If
                End Sub
        End Sub

        Private Sub folderView_ViewActivated(s As Object, e As RoutedEventArgs)
            If Not _isOnce Then Return
            _isOnce = False

            Dim item As TabItem = GetSelectedTabItem()
            If Not item Is Nothing Then
                Dim cp As Grid = findChildContentPresenter(item)
                Dim folderView As FolderView = UIHelper.FindVisualChildren(Of FolderView)(cp)?(0)
                If Not folderView Is Nothing _
                    AndAlso Not folderView.ActiveView Is Nothing Then
                    folderView.ActiveView.Focus()
                Else
                    Dim treeView As TreeView = UIHelper.FindVisualChildren(Of TreeView)(cp)?(0)
                    If Not treeView Is Nothing _
                        AndAlso Not treeView.PART_ListBox Is Nothing Then
                        treeView.PART_ListBox.Focus()
                    End If
                End If
            End If
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
                    Using Shell.OverrideCursor(Cursors.Wait)
                        Dim newTab As Object = Activator.CreateInstance(Me.ModelType)
                        _items.Add(newTab)
                        Me.SelectedItem = _items(_items.Count - 1)
                    End Using
                End Sub

            Dim firstTab As Object = Activator.CreateInstance(Me.ModelType)
            _items.Add(firstTab)

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

            _automationPeer = Nothing

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
                            Dim cp As Grid = findChildContentPresenter(item)
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

            For Each child As Grid In PART_ItemsHolder.Children
                child.Visibility = IIf(child.Tag.Equals(Me.SelectedItem), Visibility.Visible, Visibility.Hidden)
            Next
        End Sub

        Class DummyContentPresenter
            Inherits ContentPresenter

            Public Sub RemoveVisualChildPublic(child As UIElement)
                Me.RemoveVisualChild(child)
            End Sub
        End Class

        Private Function createChildContentPresenter(item As Object) As Grid
            If item Is Nothing Then Return Nothing
            If PART_ItemsHolder Is Nothing Then Return Nothing

            Dim cp As Grid = findChildContentPresenter(item)
            If Not cp Is Nothing Then Return cp

            Dim tabItem As TabItem = TryCast(item, TabItem)
            Dim data As Object = If(tabItem IsNot Nothing, tabItem.Content, item)

            ' Step 1: create a dummy ContentPresenter to apply the template
            Dim dummyPresenter As New DummyContentPresenter() With {
                .Content = data,
                .ContentTemplate = Me.SelectedContentTemplate,
                .ContentTemplateSelector = Me.SelectedContentTemplateSelector,
                .ContentStringFormat = Me.SelectedContentStringFormat
            }

            dummyPresenter.ApplyTemplate()
            dummyPresenter.UpdateLayout()

            ' Step 2: Extract the generated visuals from the template
            Dim visuals As List(Of UIElement) = UIHelper.FindVisualChildren(Of UIElement)(dummyPresenter, False).ToList()

            ' Step 3: Rehost them in a Grid
            Dim grid As New Grid()
            For Each child As UIElement In visuals
                dummyPresenter.RemoveVisualChildPublic(child)
                grid.Children.Add(child)
            Next

            ' Step 4: Create final ContentPresenter (or better: a ContentControl or your own wrapper)
            grid.Visibility = Visibility.Hidden
            grid.Tag = data
            grid.DataContext = data

            PART_ItemsHolder.Children.Add(grid)

            AddHandler tabItem.Loaded,
                Sub(s As Object, e As EventArgs)
                    setDesiredTabWidth()
                End Sub

            Return cp
        End Function

        Private _automationPeer As AutomationPeer = Nothing

        Protected Overrides Function OnCreateAutomationPeer() As AutomationPeer
            If _automationPeer Is Nothing Then
                _automationPeer = New TabControlAutomationPeer(Me)
            End If
            Return _automationPeer
        End Function

        Friend Function findChildContentPresenter(data As Object) As Grid
            If TypeOf data Is TabItem Then
                data = CType(data, TabItem).Content
            End If

            If data Is Nothing Then
                Return Nothing
            End If

            If PART_ItemsHolder Is Nothing Then
                Return Nothing
            End If

            For Each cp As Grid In PART_ItemsHolder.Children
                If cp.Tag.Equals(data) Then
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

        Public Property ModelType As Type
            Get
                Return GetValue(ModelTypeProperty)
            End Get
            Set(value As Type)
                SetValue(ModelTypeProperty, value)
            End Set
        End Property

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

        Public Class TabControlAutomationPeer
            Inherits System.Windows.Automation.Peers.TabControlAutomationPeer

            Public Sub New(owner As Laila.Shell.Controls.TabControl)
                MyBase.New(owner)
            End Sub

            Protected Overrides Function GetChildrenCore() As List(Of AutomationPeer)
                Dim ownerControl = CType(Me.Owner, Laila.Shell.Controls.TabControl)
                Dim peers As New List(Of AutomationPeer)

                For Each item In ownerControl.Items
                    Dim tabItem = CType(ownerControl.ItemContainerGenerator.ContainerFromItem(item), TabItem)
                    If tabItem IsNot Nothing Then
                        Dim tabItemPeer = New TabItemAutomationPeer(tabItem, Me)
                        If Not tabItemPeer Is Nothing Then
                            peers.Add(tabItemPeer)

                            ' Get content presenter inside your PART_ItemsHolder
                            Dim cp = ownerControl.findChildContentPresenter(item)
                            If Not cp Is Nothing Then
                                Dim gridPeer = New GridAutomationPeer(cp)
                                If Not gridPeer Is Nothing Then
                                    tabItemPeer.GridPeer = gridPeer
                                End If
                            End If
                        End If
                    End If
                Next

                Return peers
            End Function

            Public Class GridAutomationPeer
                Inherits FrameworkElementAutomationPeer

                Public Property Children As List(Of AutomationPeer) = New List(Of AutomationPeer)()
                Public Property Name As String

                Public Sub New(owner As Grid)
                    MyBase.New(owner)
                End Sub

                Protected Overrides Function GetAutomationControlTypeCore() As AutomationControlType
                    Return AutomationControlType.Pane
                End Function

                Protected Overrides Function GetChildrenCore() As List(Of AutomationPeer)
                    Return If(MyBase.GetChildrenCore(), New List(Of AutomationPeer)) _
                        .Union(If(Me.Children, New List(Of AutomationPeer)())).ToList()
                End Function

                Protected Overrides Function GetNameCore() As String
                    Return Me.Name
                End Function
            End Class

            Public Class TabItemAutomationPeer
                Inherits System.Windows.Automation.Peers.TabItemAutomationPeer

                Public Property GridPeer As GridAutomationPeer

                Public ReadOnly Property Children As List(Of AutomationPeer)
                    Get
                        Return GridPeer?.GetChildren()
                    End Get
                End Property

                Public Sub New(owner As TabItem, parent As TabControlAutomationPeer)
                    MyBase.New(owner, parent)
                End Sub

                Protected Overrides Function GetChildrenCore() As List(Of AutomationPeer)
                    Return New List(Of AutomationPeer) From {
                    New GridAutomationPeer(New Grid()) With {.Name = "Tab", .Children = If(MyBase.GetChildrenCore(), New List(Of AutomationPeer))},
                    New GridAutomationPeer(New Grid()) With {.Name = "Content", .Children = Me.Children}
                }
                End Function
            End Class
        End Class
    End Class
End Namespace
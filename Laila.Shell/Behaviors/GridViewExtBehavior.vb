Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Xml.Serialization
Imports Laila.Shell.Adorners
Imports Laila.Shell.Helpers
Imports LiteDB
Imports Microsoft.Xaml.Behaviors

Namespace Behaviors
    Public Class GridViewExtBehavior
        Inherits Behavior(Of ListView)
        Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Const COLUMN_MARGIN As Double = 5
        Private Const MARGIN_LEFT As Double = 20

        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(ColumnsInData), GetType(GridViewExtBehavior), New FrameworkPropertyMetadata(Nothing, AddressOf OnColumnsInChanged))

        Public Property ColumnsIn() As ColumnsInData
            Get
                Return GetValue(ColumnsInProperty)
            End Get
            Set(ByVal value As ColumnsInData)
                SetValue(ColumnsInProperty, value)
            End Set
        End Property

        Shared Sub OnColumnsInChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            If Not e.NewValue Is Nothing Then
                CType(d, GridViewExtBehavior).loadColumns()
            End If
        End Sub

        Public Shared ReadOnly ColumnIndexProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("ColumnIndex", GetType(Integer), GetType(GridViewExtBehavior), New UIPropertyMetadata(-1))

        Public Shared Function GetColumnIndex(obj As DependencyObject) As Integer
            Return obj.GetValue(ColumnIndexProperty)
        End Function

        Public Shared Sub SetColumnIndex(obj As DependencyObject, value As Integer)
            obj.SetValue(ColumnIndexProperty, value)
        End Sub

        Public Shared ReadOnly MaxAutoSizeWidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("MaxAutoSizeWidth", GetType(Double), GetType(GridViewExtBehavior), New UIPropertyMetadata(Double.NaN))

        Public Shared Function GetMaxAutoSizeWidth(obj As DependencyObject) As Double
            Return obj.GetValue(MaxAutoSizeWidthProperty)
        End Function

        Public Shared Sub SetMaxAutoSizeWidth(obj As DependencyObject, value As Double)
            obj.SetValue(MaxAutoSizeWidthProperty, value)
        End Sub

        Public Shared ReadOnly MinAutoSizeWidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("MinAutoSizeWidth", GetType(Double), GetType(GridViewExtBehavior), New UIPropertyMetadata(Double.NaN))

        Public Shared Function GetMinAutoSizeWidth(obj As DependencyObject) As Double
            Return obj.GetValue(MinAutoSizeWidthProperty)
        End Function

        Public Shared Sub SetMinAutoSizeWidth(obj As DependencyObject, value As Double)
            obj.SetValue(MinAutoSizeWidthProperty, value)
        End Sub

        Public Shared ReadOnly CanHideProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("CanHide", GetType(Boolean), GetType(GridViewExtBehavior), New UIPropertyMetadata(True))

        Public Shared Function GetCanHide(obj As DependencyObject) As String
            Return obj.GetValue(CanHideProperty)
        End Function

        Public Shared Sub SetCanHide(obj As DependencyObject, value As String)
            obj.SetValue(CanHideProperty, value)
        End Sub

        Public Shared ReadOnly IsVisibleProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("IsVisible", GetType(Boolean), GetType(GridViewExtBehavior), New UIPropertyMetadata(True))

        Public Shared Function GetIsVisible(obj As DependencyObject) As String
            Return obj.GetValue(IsVisibleProperty)
        End Function

        Public Shared Sub SetIsVisible(obj As DependencyObject, value As String)
            obj.SetValue(IsVisibleProperty, value)
        End Sub

        Public Shared ReadOnly IsHeaderVisibleProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("IsHeaderVisible", GetType(Boolean), GetType(GridViewExtBehavior), New UIPropertyMetadata(True))

        Public Shared Function GetIsHeaderVisible(obj As DependencyObject) As String
            Return obj.GetValue(IsHeaderVisibleProperty)
        End Function

        Public Shared Sub SetIsVHeaderisible(obj As DependencyObject, value As String)
            obj.SetValue(IsHeaderVisibleProperty, value)
        End Sub

        Public Shared ReadOnly PropertyNameProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("PropertyName", GetType(String), GetType(GridViewExtBehavior), New UIPropertyMetadata(Nothing))

        Public Shared Function GetPropertyName(obj As DependencyObject) As String
            Return obj.GetValue(PropertyNameProperty)
        End Function

        Public Shared Sub SetPropertyName(obj As DependencyObject, value As String)
            obj.SetValue(PropertyNameProperty, value)
        End Sub

        Public Shared ReadOnly GroupByPropertyNameProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("GroupByPropertyName", GetType(String), GetType(GridViewExtBehavior), New UIPropertyMetadata(Nothing))

        Public Shared Function GetGroupByPropertyName(obj As DependencyObject) As String
            Return obj.GetValue(GroupByPropertyNameProperty)
        End Function

        Public Shared Sub SetGroupByPropertyName(obj As DependencyObject, value As String)
            obj.SetValue(GroupByPropertyNameProperty, value)
        End Sub

        Public Shared ReadOnly SortPropertyNameProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("SortPropertyName", GetType(String), GetType(GridViewExtBehavior), New UIPropertyMetadata(Nothing))

        Public Shared Function GetSortPropertyName(obj As DependencyObject) As String
            Return obj.GetValue(SortPropertyNameProperty)
        End Function

        Public Shared Sub SetSortPropertyName(obj As DependencyObject, value As String)
            obj.SetValue(SortPropertyNameProperty, value)
        End Sub

        Public Shared ReadOnly HeaderTextProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("HeaderText", GetType(String), GetType(GridViewExtBehavior), New UIPropertyMetadata(Nothing))

        Public Shared Function GetHeaderText(obj As DependencyObject) As String
            Return obj.GetValue(HeaderTextProperty)
        End Function

        Public Shared Sub SetHeaderText(obj As DependencyObject, value As String)
            obj.SetValue(HeaderTextProperty, value)
        End Sub

        Private _listView As ListView
        Private _gridView As GridView
        Private _headerRowPresenter As GridViewHeaderRowPresenter
        Private _isLoaded As Boolean
        Private _activeColumns As List(Of ActiveColumnStateData)
        Private _dontWrite As Boolean = False
        Private _clickedColumn As GridViewColumn
        Private _groupByPropertyNames As List(Of String) = New List(Of String)()
        Private _scrollViewer As ScrollViewer
        Private _skipResize As Boolean
        Private _isInitialResize As Boolean = True
        Private _resizeTimer As Timer

        Public ReadOnly Property ColumnIndexFor(propertyName As String) As Integer
            Get
                Dim header As GridViewColumnHeader =
                    UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter) _
                        .FirstOrDefault(Function(h) Not h.Column Is Nothing _
                            AndAlso (EqualityComparer(Of String).Default.Equals(GetPropertyName(h.Column), propertyName)))
                Return If(Not header Is Nothing, GetColumnIndex(header.Column), -1)
            End Get
        End Property

        Protected Overrides Sub OnAttached()
            MyBase.OnAttached()

            _listView = Me.AssociatedObject
            _gridView = _listView.View

            _listView.AddHandler(GridViewColumnHeader.ClickEvent, New RoutedEventHandler(AddressOf ColumnHeader_Click))

            ' auto resize column when rebound/item added
            AddHandler _listView.ItemContainerGenerator.StatusChanged, AddressOf icgStatusChanged
            TypeDescriptor.GetProperties(_listView)("ItemsSource") _
                .AddValueChanged(_listView,
                    Sub()
                        _isInitialResize = True
                        _skipResize = False
                        resizeVisibleRows()

                        'If Not _listView.ItemsSource Is Nothing AndAlso TypeOf _listView.ItemsSource Is INotifyCollectionChanged Then
                        '    AddHandler CType(_listView.ItemsSource, INotifyCollectionChanged).CollectionChanged,
                        '        Sub(sender2 As Object, e2 As NotifyCollectionChangedEventArgs)
                        '            Select Case e2.Action
                        '                Case NotifyCollectionChangedAction.Add
                        '                    _skipResize = False
                        '            End Select
                        '        End Sub
                        'End If

                        If _isLoaded Then regroup()
                    End Sub)
            AddHandler _listView.Loaded,
               Sub(sender As Object, e As EventArgs)
                   If _isLoaded Then Return

                   _headerRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(_listView)(0)
                   Dim headerRowScrollViewer As ScrollViewer = _headerRowPresenter.Parent
                   Dim headerRowGrid As Grid = New Grid()
                   headerRowGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(MARGIN_LEFT, GridUnitType.Pixel)})
                   headerRowGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1, GridUnitType.Star)})
                   headerRowScrollViewer.Content = headerRowGrid
                   _headerRowPresenter.SetValue(Grid.ColumnProperty, 1)
                   headerRowGrid.Children.Add(_headerRowPresenter)
                   Dim headerRowMarginPresenter As GridViewHeaderRowPresenter = New GridViewHeaderRowPresenter()
                   headerRowMarginPresenter.SetValue(Grid.ColumnProperty, 0)
                   headerRowGrid.Children.Add(headerRowMarginPresenter)

                   _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_listView)(0)
                   UIHelper.FindVisualChildren(Of VirtualizingStackPanel)(_listView)(0).Margin = New Thickness(MARGIN_LEFT, 0, 0, 0)
                   _isLoaded = True

                   If Not Me.ColumnsIn Is Nothing Then
                       loadColumns()
                   End If

                   ' hook reorder event
                   AddHandler _gridView.Columns.CollectionChanged,
                       Async Sub(s2 As Object, e2 As NotifyCollectionChangedEventArgs)
                           If e2.Action = NotifyCollectionChangedAction.Move Then
                               ' move
                               Dim actualOldIndex As Integer = -1
                               For i = 0 To e2.OldStartingIndex
                                   actualOldIndex += 1
                                   If actualOldIndex + 1 > _activeColumns.Count Then
                                       Exit For
                                   End If
                                   If Not _activeColumns(actualOldIndex).IsVisible Then
                                       i = i - 1
                                   End If
                               Next
                               Dim actualNewIndex As Integer = -1
                               For i = 0 To e2.NewStartingIndex
                                   actualNewIndex += 1
                                   If actualNewIndex + 1 > _activeColumns.Count Then
                                       Exit For
                                   End If
                                   If Not _activeColumns(actualNewIndex).IsVisible Then
                                       i = i - 1
                                   End If
                               Next

                               Dim column As ActiveColumnStateData = _activeColumns(actualOldIndex)
                               _activeColumns.RemoveAt(actualOldIndex)
                               _activeColumns.Insert(actualNewIndex, column)

                               ' renumber columns
                               Dim x As Integer = 0
                               For Each col In _activeColumns.Where(Function(c) c.IsVisible)
                                   SetColumnIndex(col.Column, x)
                                   x += 1
                               Next
                               Me.NotifyOfPropertyChange("ColumnIndexFor")

                               ' write state
                               writeState()
                           End If
                       End Sub
               End Sub
        End Sub

        Private Sub loadColumns()
            Debug.WriteLine("loadColumns()")
            If _isLoaded Then
                _activeColumns = New List(Of ActiveColumnStateData)()

                Dim gridViewState As GridViewStateData = readState(Me.ColumnsIn.ViewName)

                If Not gridViewState Is Nothing Then
                    ' restore column order from state
                    For Each columnState In gridViewState.Columns
                        Dim column As GridViewColumn = Me.ColumnsIn.Items.SingleOrDefault(Function(c) getColumnName(c) = columnState.Name)
                        If Not column Is Nothing Then
                            _activeColumns.Add(New ActiveColumnStateData() With {
                                .Column = column, .IsVisible = columnState.IsVisible, .Width = columnState.Width
                            })
                            column.Width = 0
                        End If
                    Next
                End If

                ' add columns that aren't in the state 
                Dim idx As Integer = 0
                For Each column In Me.ColumnsIn.Items
                    If Not _activeColumns.Exists(Function(ac) ac.Column.Equals(column)) Then
                        ' insert column after it's predecessor
                        If idx > 0 Then
                            Dim previousName As String = getColumnName(Me.ColumnsIn.Items(idx - 1))
                            Dim previousColumn As ActiveColumnStateData = _activeColumns.SingleOrDefault(Function(ac) getColumnName(ac.Column) = previousName)
                            _activeColumns.Insert(_activeColumns.IndexOf(previousColumn) + 1,
                                New ActiveColumnStateData() With {.Column = column, .IsVisible = GetIsVisible(column), .Width = column.GetValue(GridViewColumn.WidthProperty)})
                        Else
                            _activeColumns.Insert(0,
                                New ActiveColumnStateData() With {.Column = column, .IsVisible = GetIsVisible(column), .Width = column.GetValue(GridViewColumn.WidthProperty)})
                        End If
                    End If
                    idx += 1
                Next

                ' verify column names
                Dim columnNames As List(Of String) = New List(Of String)
                For Each column In Me.ColumnsIn.Items
                    ' get name
                    Dim name As String = getColumnName(column)

                    If Not columnNames.Contains(name) Then
                        columnNames.Add(name)
                    Else
                        Throw New InvalidOperationException("GridViewExtBehavior requires all columns have unique names.")
                    End If
                Next

                showHideColumns()

                ' set sort from state
                Dim view As ICollectionView = _listView.Items
                Using view.DeferRefresh()
                    resetSortDescriptions(view)
                    If Not gridViewState Is Nothing AndAlso Not String.IsNullOrEmpty(gridViewState.SortPropertyName) Then
                        view.SortDescriptions.Add(New SortDescription() With {
                        .PropertyName = gridViewState.SortPropertyName,
                        .Direction = gridViewState.SortDirection
                    })
                    Else
                        Dim initialSortPropertyName As String = GetSortPropertyName(_activeColumns(0).Column)
                        If String.IsNullOrWhiteSpace(initialSortPropertyName) Then
                            initialSortPropertyName = GetPropertyName(_activeColumns(0).Column)
                        End If
                        view.SortDescriptions.Add(New SortDescription() With {
                        .PropertyName = initialSortPropertyName,
                        .Direction = ListSortDirection.Ascending
                    })
                    End If
                End Using

                fixSortGlyphs()

                For Each column In _activeColumns
                    ' hook column resize event
                    AddHandler CType(column.Column, INotifyPropertyChanged).PropertyChanged,
                        Sub(sender As Object, e As PropertyChangedEventArgs)
                            If e.PropertyName = "Width" Then
                                If Not _dontWrite Then
                                    ' keep track of autosize columns
                                    column.Width = column.Column.Width

                                    ' write state
                                    writeState()
                                End If
                            End If
                        End Sub
                Next

                ' set initial group by
                If Not String.IsNullOrWhiteSpace(Me.ColumnsIn.InitialGroupByPropertyNames) Then
                    _groupByPropertyNames = New List(Of String)(Me.ColumnsIn.InitialGroupByPropertyNames.Split(","c))
                    regroup()
                End If
            End If
        End Sub

        Private Sub resetSortDescriptions(view As ICollectionView)
            view.SortDescriptions.Clear()
            For Each firstSortProperty In Me.ColumnsIn.PrimarySortProperties.Split(",")
                Dim firstSortProperty1() As String = firstSortProperty.Split(" ")
                Dim firstSortDirection As ListSortDirection = ListSortDirection.Ascending
                If firstSortProperty1.Count >= 2 Then
                    Select Case firstSortProperty1(1).ToUpper()
                        Case "ASC" : firstSortDirection = ListSortDirection.Ascending
                        Case "DESC" : firstSortDirection = ListSortDirection.Descending
                        Case Else
                            Throw New IndexOutOfRangeException()
                    End Select
                End If
                view.SortDescriptions.Add(New SortDescription() With {
                    .PropertyName = firstSortProperty1(0),
                    .Direction = firstSortDirection
                })
            Next
        End Sub

        Private Sub showHideColumns()
            Debug.WriteLine("showHideColumns()")
            ' clear columns
            _gridView.Columns.Clear()

            ' re-add columns, restoring properties from state in the process
            Dim i As Integer = 0
            For Each column In _activeColumns.Where(Function(c) c.IsVisible)
                ' init width and visibility
                column.Column.Width = If(column.Width.Equals(Double.NaN), 0, column.Width)

                SetColumnIndex(column.Column, i)
                column.OriginalIndex = i

                _gridView.Columns.Add(column.Column)

                ' re-hook headers
                Dim colHeader As GridViewColumnHeader =
                        UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter) _
                            .FirstOrDefault(Function(h) Not h.Column Is Nothing AndAlso h.Column.Equals(column.Column))
                If GetIsHeaderVisible(column.Column) Then
                    AddHandler colHeader.MouseRightButtonDown, AddressOf ColumnHeader_MouseRightButtonDown
                Else
                    colHeader.Visibility = Visibility.Hidden
                End If

                i += 1
            Next

            _skipResize = False
            resizeVisibleRows()
        End Sub

        Private Sub fixSortGlyphs()
            ' fix sort glyphs
            Dim hcs As List(Of GridViewColumnHeader) = UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter).ToList()
            Dim view As ICollectionView = _listView.Items
            If view.SortDescriptions.Count > Me.ColumnsIn.PrimarySortProperties.Split(",").Count Then
                Dim currentSort As SortDescription = view.SortDescriptions(view.SortDescriptions.Count - 1)
                Dim currentSortedColumnHeader As GridViewColumnHeader = getCurrentSortedColumnHeader()
                If Not currentSortedColumnHeader Is Nothing Then
                    If currentSort.Direction = ListSortDirection.Ascending Then
                        GridViewColumnHeaderGlyphAdorner.Add(currentSortedColumnHeader, "GridViewExtBehavior.Sort", 1,
                           "pack://application:,,,/Laila.Shell;component/Images/sortasc.png", HorizontalAlignment.Center)
                    Else
                        GridViewColumnHeaderGlyphAdorner.Add(currentSortedColumnHeader, "GridViewExtBehavior.Sort", 1,
                           "pack://application:,,,/Laila.Shell;component/Images/sortdesc.png", HorizontalAlignment.Center)
                    End If
                End If
            End If
        End Sub

        Private Function getCurrentSortedColumnHeader() As GridViewColumnHeader
            Dim view As ICollectionView = _listView.Items
            Dim currentSort As SortDescription
            If view.SortDescriptions.Count > Me.ColumnsIn.PrimarySortProperties.Split(",").Count Then
                currentSort = view.SortDescriptions(view.SortDescriptions.Count - 1)

                Dim result As GridViewColumnHeader =
                    UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter) _
                        .FirstOrDefault(Function(h) Not h.Column Is Nothing _
                            AndAlso (EqualityComparer(Of String).Default.Equals(GetSortPropertyName(h.Column), currentSort.PropertyName)))
                If result Is Nothing Then
                    result =
                    UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter) _
                        .FirstOrDefault(Function(h) Not h.Column Is Nothing _
                            AndAlso (EqualityComparer(Of String).Default.Equals(GetPropertyName(h.Column), currentSort.PropertyName)))
                End If
                Return result
            Else
                Return Nothing
            End If
        End Function

        Private Sub ColumnHeader_Click(sender As Object, e As RoutedEventArgs)
            If TypeOf e.OriginalSource Is GridViewColumnHeader Then
                Dim headerClicked As GridViewColumnHeader = e.OriginalSource
                If Not headerClicked Is Nothing AndAlso Not headerClicked.Column Is Nothing Then
                    Dim propertyName As String = GetSortPropertyName(headerClicked.Column)
                    If String.IsNullOrEmpty(propertyName) Then
                        propertyName = GetPropertyName(headerClicked.Column)
                    End If
                    If Not String.IsNullOrEmpty(propertyName) Then
                        If Not _listView Is Nothing Then
                            ' apply sort
                            applySort(_listView.Items, propertyName, -1, _listView, headerClicked)
                            writeState()

                            ' set focus back to item
                            _listView.ScrollIntoView(_listView.SelectedItem)
                            _listView.UpdateLayout()
                            Dim item As ListViewItem = _listView.ItemContainerGenerator.ContainerFromItem(_listView.SelectedItem)
                            If Not item Is Nothing AndAlso Not item.IsKeyboardFocusWithin Then
                                item.Focus()
                            End If
                        End If
                    End If
                End If
            End If
        End Sub

        Private Sub ColumnHeader_MouseRightButtonDown(sender As Object, e As RoutedEventArgs)
            Dim menu As System.Windows.Controls.ContextMenu = New System.Windows.Controls.ContextMenu()

            For Each column In _activeColumns
                ' add menu item
                Dim headerText As String = column.Column.GetValue(GridViewExtBehavior.HeaderTextProperty)
                If String.IsNullOrWhiteSpace(headerText) AndAlso TypeOf column.Column.Header Is String Then
                    headerText = column.Column.Header
                End If
                Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = headerText,
                    .IsCheckable = True,
                    .StaysOpenOnClick = True,
                    .IsChecked = column.IsVisible,
                    .IsEnabled = GetCanHide(column.Column)
                }
                AddHandler menuItem.Checked,
                    Sub(s2 As Object, e2 As EventArgs)
                        column.IsVisible = True
                        showHideColumns()
                        fixSortGlyphs()

                        ' write state
                        writeState()
                    End Sub
                AddHandler menuItem.Unchecked,
                    Sub(s2 As Object, e2 As EventArgs)
                        Dim countChecked As Integer = 0
                        For i = 2 To menu.Items.Count - 1
                            If CType(menu.Items(i), MenuItem).IsChecked Then countChecked += 1
                        Next
                        If countChecked = 0 Then
                            ' always one column visible
                            menuItem.IsChecked = True
                        Else
                            column.IsVisible = False
                            showHideColumns()
                            fixSortGlyphs()

                            ' write state
                            writeState()

                            Dim view As ICollectionView = _listView.Items
                            Dim propertyName As String = GetSortPropertyName(column.Column)
                            If String.IsNullOrEmpty(propertyName) Then
                                propertyName = GetPropertyName(column.Column)
                            End If
                            If view.SortDescriptions.Count > Me.ColumnsIn.PrimarySortProperties.Split(",").Count Then
                                Dim currentSort As SortDescription = view.SortDescriptions(view.SortDescriptions.Count - 1)
                                If currentSort.PropertyName = propertyName Then
                                    resetSortDescriptions(view)

                                    Dim currentSortedColumnHeader As GridViewColumnHeader = getCurrentSortedColumnHeader()
                                    If Not currentSortedColumnHeader Is Nothing Then
                                        GridViewColumnHeaderGlyphAdorner.Remove(currentSortedColumnHeader, "GridViewSort")
                                    End If
                                End If
                            End If
                        End If
                    End Sub
                menu.Items.Add(menuItem)
            Next

            If TypeOf sender Is GridViewColumnHeader Then
                Dim headerClicked As GridViewColumnHeader = sender
                If Not headerClicked Is Nothing AndAlso Not headerClicked.Column Is Nothing Then
                    ' make "group by" menu item
                    Dim groupMenuItem As MenuItem = New MenuItem() With {.Header = "Groepeer op deze kolom", .IsCheckable = True, .IsEnabled = True, .IsChecked = False}
                    Dim propertyName As String = GetGroupByPropertyName(headerClicked.Column)
                    If String.IsNullOrEmpty(propertyName) Then
                        propertyName = GetPropertyName(headerClicked.Column)
                    End If
                    AddHandler groupMenuItem.Checked,
                        Sub(s2 As Object, e2 As EventArgs)
                            If Not _clickedColumn Is Nothing Then
                                _groupByPropertyNames.Add(propertyName)
                                regroup()
                            End If
                        End Sub
                    AddHandler groupMenuItem.Unchecked,
                        Sub(s2 As Object, e2 As EventArgs)
                            If Not _clickedColumn Is Nothing Then
                                _groupByPropertyNames.Remove(propertyName)
                                regroup()
                            End If
                        End Sub
                    menu.Items.Insert(0, groupMenuItem)
                    menu.Items.Insert(1, New Separator())

                    ' assign menu
                    _gridView.ColumnHeaderContextMenu = menu

                    _clickedColumn = Nothing
                    If Not String.IsNullOrEmpty(propertyName) Then
                        groupMenuItem.IsEnabled = True
                        groupMenuItem.IsChecked = _groupByPropertyNames.Contains(propertyName)
                    Else
                        groupMenuItem.IsEnabled = False
                        groupMenuItem.IsChecked = False
                    End If
                    _clickedColumn = headerClicked.Column
                End If
            End If
        End Sub

        Private Sub applySort(view As ICollectionView, propertyName As String, direction As ListSortDirection, listView As ListView, sortedColumnHeader As GridViewColumnHeader)
            Dim newDirection As ListSortDirection = ListSortDirection.Ascending
            If view.SortDescriptions.Count > Me.ColumnsIn.PrimarySortProperties.Split(",").Count Then
                Dim currentSort As SortDescription = view.SortDescriptions(view.SortDescriptions.Count - 1)
                If currentSort.PropertyName = propertyName Then
                    If currentSort.Direction = ListSortDirection.Ascending Then
                        newDirection = ListSortDirection.Descending
                    Else
                        newDirection = ListSortDirection.Ascending
                    End If
                End If

                Dim currentSortedColumnHeader As GridViewColumnHeader = getCurrentSortedColumnHeader()
                If Not currentSortedColumnHeader Is Nothing Then
                    GridViewColumnHeaderGlyphAdorner.Remove(currentSortedColumnHeader, "GridViewExtBehavior.Sort")
                End If

                view.SortDescriptions.Remove(view.SortDescriptions(view.SortDescriptions.Count - 1))
            End If
            If direction <> -1 Then
                newDirection = direction
            End If
            If Not String.IsNullOrEmpty(propertyName) Then
                view.SortDescriptions.Add(New SortDescription(propertyName, newDirection))
                If newDirection = ListSortDirection.Ascending Then
                    GridViewColumnHeaderGlyphAdorner.Add(sortedColumnHeader, "GridViewExtBehavior.Sort", 1,
                       "pack://application:,,,/Laila.Shell;component/Images/sortasc.png", HorizontalAlignment.Center)
                Else
                    GridViewColumnHeaderGlyphAdorner.Add(sortedColumnHeader, "GridViewExtBehavior.Sort", 1,
                       "pack://application:,,,/Laila.Shell;component/Images/sortdesc.png", HorizontalAlignment.Center)
                End If
            End If
        End Sub
        Private Sub regroup()
            Dim view As CollectionView = CollectionViewSource.GetDefaultView(_listView.ItemsSource)
            If Not view Is Nothing Then
                If view.GroupDescriptions.Count > 0 Then
                    view.GroupDescriptions.Clear()
                End If
                If _groupByPropertyNames.Count > 0 Then
                    If _listView.GroupStyle.Count = 0 Then
                        _listView.GroupStyle.Add(_listView.FindResource("groupByStyle"))
                    End If
                Else
                    If _listView.GroupStyle.Count > 0 Then
                        _listView.GroupStyle.Clear()
                    End If
                End If
                For Each item In _groupByPropertyNames
                    Dim groupDescription As PropertyGroupDescription = New PropertyGroupDescription(item)
                    view.GroupDescriptions.Add(groupDescription)
                Next
            End If
            _headerRowPresenter.Margin = New Thickness(_groupByPropertyNames.Count * 21, 0, 0, 0)
        End Sub

        Private Sub icgStatusChanged(sender As Object, e As EventArgs)
            If _listView.ItemContainerGenerator.Status = Primitives.GeneratorStatus.ContainersGenerated Then
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        resizeVisibleRows()
                    End Sub, Threading.DispatcherPriority.ContextIdle)
            End If
        End Sub

        Private Sub resizeVisibleRows()
            If Not _headerRowPresenter Is Nothing Then
                Dim rows As List(Of GridViewRowPresenter) = UIHelper.FindVisualChildren(Of GridViewRowPresenter)(_listView).ToList()

                If Not _skipResize Then
                    Dim isFullGrid As Boolean = rows.Sum(Function(r) r.DesiredSize.Height) <= _listView.ActualHeight
                    If rows.Count > 0 AndAlso (_isInitialResize OrElse isFullGrid) Then
                        resizeForRows(rows.Select(Function(r) r.DataContext).ToList(), False)

                        If isFullGrid OrElse rows.Count = _listView.Items.Count Then _isInitialResize = False
                        _skipResize = True
                    End If
                End If

                Dim headers As List(Of GridViewColumnHeader) =
                                            UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter).ToList()
                For Each header In headers.Where(Function(h) h.Column Is Nothing).ToList()
                    headers.Remove(header)
                Next
                For Each item In rows.Select(Function(r) r.DataContext).ToList()
                    Dim lvi As ListViewItem = _listView.ItemContainerGenerator.ContainerFromItem(item)
                    If Not lvi Is Nothing Then
                        lvi.HorizontalAlignment = HorizontalAlignment.Left
                        lvi.Width = headers.Sum(Function(h) h.Width)
                    End If
                Next
            End If
        End Sub

        Private Sub resizeForRows(list As List(Of Object), minimum As Boolean)
            If Not _headerRowPresenter Is Nothing Then
                Dim hcs As List(Of GridViewColumnHeader) = UIHelper.FindVisualChildren(Of GridViewColumnHeader)(_headerRowPresenter).ToList()

                _dontWrite = True
                If Not _activeColumns Is Nothing Then
                    For Each activeCol In _activeColumns.Where(Function(c) (Double.IsNaN(c.Width) OrElse c.Width = 0) AndAlso c.IsVisible)
                        Dim width As Double = 0
                        If minimum Then
                            width = activeCol.Column.ActualWidth
                        End If

                        ' measure header
                        Dim hc As GridViewColumnHeader = hcs.First(Function(i) activeCol.Column.Equals(i.Column))
                        hc.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
                        If Math.Ceiling(hc.DesiredSize.Width + COLUMN_MARGIN) > width Then
                            width = Math.Ceiling(hc.DesiredSize.Width + COLUMN_MARGIN)
                        End If

                        If list.Count > 0 Then
                            ' measure available rows
                            For Each item In list
                                Dim lvi As ListViewItem = _listView.ItemContainerGenerator.ContainerFromItem(item)
                                If Not lvi Is Nothing Then
                                    Dim rp As GridViewRowPresenter = UIHelper.FindVisualChildren(Of GridViewRowPresenter)(lvi).FirstOrDefault()
                                    If Not rp Is Nothing Then
                                        Dim c As GridViewColumn = activeCol.Column
                                        Dim el As UIElement = VisualTreeHelper.GetChild(rp, activeCol.OriginalIndex)
                                        el.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
                                        If Math.Ceiling(el.DesiredSize.Width + COLUMN_MARGIN) > width Then
                                            width = Math.Ceiling(el.DesiredSize.Width + COLUMN_MARGIN)
                                        End If
                                    End If
                                End If
                            Next
                        End If

                        If Not Double.IsNaN(GetMaxAutoSizeWidth(activeCol.Column)) Then
                            If width > GetMaxAutoSizeWidth(activeCol.Column) Then
                                width = GetMaxAutoSizeWidth(activeCol.Column)
                            End If
                        End If

                        If Not Double.IsNaN(GetMinAutoSizeWidth(activeCol.Column)) Then
                            If width < GetMinAutoSizeWidth(activeCol.Column) Then
                                width = GetMinAutoSizeWidth(activeCol.Column)
                            End If
                        End If

                        activeCol.Column.Width = width
                        hcs.FirstOrDefault(Function(hc1) Not hc1.Column Is Nothing AndAlso hc1.Column.Equals(activeCol.Column)).Width = width
                    Next
                End If
                _dontWrite = False
            End If
        End Sub

        Private Sub writeState()
            ' build state
            Dim gridViewState As GridViewStateData = New GridViewStateData()
            gridViewState.Columns = New List(Of ColumnStateData)()
            For Each column In _activeColumns
                Dim columnState As ColumnStateData = New ColumnStateData()
                columnState.Name = getColumnName(column.Column)
                columnState.IsVisible = column.IsVisible
                columnState.Width = column.Width
                gridViewState.Columns.Add(columnState)
            Next

            Dim view As ICollectionView = _listView.Items
            If view.SortDescriptions.Count > Me.ColumnsIn.PrimarySortProperties.Split(",").Count Then
                Dim currentSort As SortDescription = view.SortDescriptions(view.SortDescriptions.Count - 1)
                gridViewState.SortPropertyName = currentSort.PropertyName
                gridViewState.SortDirection = currentSort.Direction
            End If

            Me.WriteState(Me.ColumnsIn.ViewName, gridViewState)
        End Sub

        Protected Overridable Function readState(viewName As String) As GridViewStateData
            Dim dbFileName As String = getStateDBFileName()
            Dim viewId As String = Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(viewName))

            If Not File.Exists(dbFileName) Then
                Return Nothing
            End If

            Using mem As MemoryStream = New MemoryStream()
                Using db = New LiteDatabase(dbFileName)
                    Dim info As LiteFileInfo(Of String) = db.FileStorage.FindById(viewId)
                    If Not info Is Nothing Then
                        info.CopyTo(mem)
                        mem.Seek(0, SeekOrigin.Begin)
                        Dim s As XmlSerializer = New XmlSerializer(GetType(GridViewStateData))
                        Return s.Deserialize(mem)
                    End If
                End Using
            End Using
        End Function

        Protected Overridable Sub WriteState(viewName As String, state As GridViewStateData)
            Dim dbFileName As String = getStateDBFileName()
            Dim viewId As String = Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(viewName))

            If Not Directory.Exists(Path.GetDirectoryName(dbFileName)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(dbFileName))
            End If

            Dim s As XmlSerializer = New XmlSerializer(GetType(GridViewStateData))
            Using mem As MemoryStream = New MemoryStream()
                s.Serialize(mem, state)
                mem.Seek(0, SeekOrigin.Begin)

                Using db = New LiteDatabase(dbFileName)
                    db.FileStorage.Upload(viewId, "state.xml", mem)
                End Using
            End Using
        End Sub

        Private Function getStateDBFileName() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return IO.Path.Combine(path, "GridViewState.db")
        End Function

        Private Function getColumnName(column As GridViewColumn) As String
            Dim name As String =
                column.GetValue(GridViewExtBehavior.PropertyNameProperty)
            If String.IsNullOrWhiteSpace(name) Then
                name = column.GetValue(GridViewExtBehavior.HeaderTextProperty)
            End If
            If String.IsNullOrWhiteSpace(name) And TypeOf column.Header Is String Then
                name = column.Header
            End If
            If String.IsNullOrWhiteSpace(name) Then
                Throw New InvalidDataException("GridViewExtBehavior requires that each column has a name.")
            Else
                Return name
            End If
        End Function

        Public Sub NotifyOfPropertyChange(propertyName As String)
            UIHelper.OnUIThread(
                Sub()
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
                End Sub)
        End Sub

        Public Class GridViewStateData
            Public Property Columns As List(Of ColumnStateData)
            Public Property SortPropertyName As String
            Public Property SortDirection As ListSortDirection
        End Class

        Public Class ColumnStateData
            Public Property Name As String
            Public Property IsVisible As Boolean
            Public Property Width As Double
        End Class

        Public Class ActiveColumnStateData
            Public Property Column As GridViewColumn
            Public Property IsVisible As Boolean
            Public Property Width As Double
            Public Property OriginalIndex As Integer
        End Class

        Public Class ColumnsInData
            Public Property ViewName As String
            Public Property Items As List(Of GridViewColumn)
            Public Property InitialGroupByPropertyNames As String
            Public Property PrimarySortProperties As String = ""
        End Class
    End Class
End Namespace
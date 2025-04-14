Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.Helpers
Imports WpfToolkit.Controls

Namespace Controls
    Public Class SlidingExpander
        Inherits Expander

        Public Shared ReadOnly ToggleButtonStyleProperty As DependencyProperty = DependencyProperty.Register("ToggleButtonStyle", GetType(Style), GetType(SlidingExpander), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly OrientationProperty As DependencyProperty = DependencyProperty.Register("Orientation", GetType(Orientation), GetType(SlidingExpander), New FrameworkPropertyMetadata(Orientation.Horizontal, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ToggleButtonVisibilityProperty As DependencyProperty = DependencyProperty.Register("ToggleButtonVisibility", GetType(Visibility), GetType(SlidingExpander), New FrameworkPropertyMetadata(Visibility.Visible, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoSelectAllBehaviorProperty As DependencyProperty = DependencyProperty.Register("DoSelectAllBehavior", GetType(Boolean), GetType(SlidingExpander), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SlidingExpander), New FrameworkPropertyMetadata(GetType(SlidingExpander)))
            SlidingExpander.IsExpandedProperty.OverrideMetadata(GetType(SlidingExpander), New FrameworkPropertyMetadata(True, AddressOf IsExpandedChanged))
        End Sub

        Private Const ANIMATIONMS As Integer = 150
        Private Const ANIMATIONSTEPS As Integer = 25

        Private PART_ContentContainer As ScrollViewer
        Private PART_ContentPresenter As ContentPresenter
        Private PART_Title As ContentPresenter
        Private PART_TitleLabel As Button
        Private PART_ToggleButton As ToggleButton
        Private _timer As DispatcherTimer
        Private _currentStep As Double = 0
        Private _totalHeight As Double = 0
        Private _doIgnoreExpandCollapse As Boolean
        Private _doFocus As Boolean = False

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ContentContainer = Template.FindName("PART_ContentContainer", Me)
            Me.PART_ContentPresenter = Template.FindName("PART_ContentPresenter", Me)
            Me.PART_Title = Template.FindName("PART_Title", Me)
            Me.PART_TitleLabel = Template.FindName("PART_TitleLabel", Me)
            Me.PART_ToggleButton = Template.FindName("PART_ToggleButton", Me)

            _timer = New DispatcherTimer()

            AddHandler Me.PART_TitleLabel.PreviewKeyDown,
                Sub(s As Object, e As KeyEventArgs)
                    If e.Key = Key.Space Then
                        Me.PART_ToggleButton.IsChecked = Not Me.PART_ToggleButton.IsChecked
                    ElseIf e.Key = Key.Left Then
                        Me.PART_ToggleButton.IsChecked = False
                    ElseIf e.Key = Key.Right Then
                        Me.PART_ToggleButton.IsChecked = True
                    End If
                End Sub

            AddHandler Me.PART_TitleLabel.PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    If e.ClickCount = 2 Then
                        _doFocus = False
                        Me.PART_ToggleButton.IsChecked = Not Me.PART_ToggleButton.IsChecked
                        Me.PART_TitleLabel.Focus()
                    Else
                        _doFocus = True
                        Dim f As Func(Of Task) =
                            Async Function() As Task
                                Await Task.Delay(200)
                                If Not _doFocus OrElse _timer.IsEnabled Then Return
                                UIHelper.OnUIThread(
                                    Sub()
                                        Me.Focus()
                                    End Sub)
                            End Function
                        Task.Run(f)
                    End If
                    e.Handled = True
                End Sub

            AddHandler Me.PART_Title.PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                End Sub

            AddHandler Me.PART_ToggleButton.Unchecked,
                Sub(s As Object, e As RoutedEventArgs)
                    _doIgnoreExpandCollapse = True
                    Me.IsExpanded = False
                    _doIgnoreExpandCollapse = False
                    _totalHeight = Me.PART_ContentPresenter.ActualHeight
                    _currentStep = 1.0
                    RemoveHandler _timer.Tick, AddressOf timer_TickExpand
                    RemoveHandler _timer.Tick, AddressOf timer_TickCollapse
                    AddHandler _timer.Tick, AddressOf timer_TickCollapse
                    _timer.Interval = New TimeSpan(0, 0, 0, 0, ANIMATIONMS / ANIMATIONSTEPS)
                    _timer.Start()
                End Sub
            AddHandler Me.PART_ToggleButton.Checked,
                Sub(s As Object, e As RoutedEventArgs)
                    _doIgnoreExpandCollapse = True
                    Me.IsExpanded = True
                    _doIgnoreExpandCollapse = False
                    _totalHeight = Me.PART_ContentPresenter.ActualHeight
                    _currentStep = 0.0
                    RemoveHandler _timer.Tick, AddressOf timer_TickExpand
                    RemoveHandler _timer.Tick, AddressOf timer_TickCollapse
                    AddHandler _timer.Tick, AddressOf timer_TickExpand
                    _timer.Interval = New TimeSpan(0, 0, 0, 0, ANIMATIONMS / ANIMATIONSTEPS)
                    _timer.Start()
                End Sub
        End Sub

        Private Sub timer_TickCollapse(s As Object, e As EventArgs)
            _currentStep -= 1.0 / ANIMATIONSTEPS
            If _currentStep <= 0.0 Then
                _timer.Stop()
                Me.PART_ContentContainer.Visibility = Visibility.Collapsed
            Else
                Me.PART_ContentContainer.Height = Math.Max(0, _totalHeight * _currentStep)
                Me.PART_ContentContainer.ScrollToBottom()
            End If
        End Sub

        Private Sub timer_TickExpand(s As Object, e As EventArgs)
            Me.PART_ContentContainer.Visibility = Visibility.Visible
            _currentStep += 1.0 / ANIMATIONSTEPS
            If _currentStep >= 1.0 Then
                _timer.Stop()
                Me.PART_ContentContainer.Height = Double.NaN
            Else
                Me.PART_ContentContainer.Height = Math.Min(_totalHeight, _totalHeight * _currentStep)
                Me.PART_ContentContainer.ScrollToBottom()
            End If
        End Sub

        Public Overloads Sub Focus()
            'Dim vwp As VirtualizingPanel = UIHelper.FindVisualChildren(Of VirtualizingPanel)(Me.PART_ContentContainer).ToList()(0)
            Dim group As CollectionViewGroup = Me.DataContext
            Dim listBox As ListBox = UIHelper.GetParentOfType(Of ListBox)(Me)
            If Me.DoSelectAllBehavior AndAlso Not Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then
                If Not Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then listBox.SelectedItems.Clear()
                For Each item In group.Items
                    If Not listBox.SelectedItems.Contains(item) Then
                        listBox.SelectedItems.Add(item)
                    End If
                Next
            End If
            Me.PART_TitleLabel.Focus()
        End Sub

        Public Property Orientation As Orientation
            Get
                Return GetValue(OrientationProperty)
            End Get
            Set(value As Orientation)
                SetValue(OrientationProperty, value)
            End Set
        End Property

        Public Property ToggleButtonStyle As Style
            Get
                Return GetValue(ToggleButtonStyleProperty)
            End Get
            Set(value As Style)
                SetValue(ToggleButtonStyleProperty, value)
            End Set
        End Property

        Public Property ToggleButtonVisibility As Visibility
            Get
                Return GetValue(ToggleButtonVisibilityProperty)
            End Get
            Set(value As Visibility)
                SetValue(ToggleButtonVisibilityProperty, value)
            End Set
        End Property

        Public Property DoSelectAllBehavior As Boolean
            Get
                Return GetValue(DoSelectAllBehaviorProperty)
            End Get
            Set(value As Boolean)
                SetValue(DoSelectAllBehaviorProperty, value)
            End Set
        End Property

        Private Shared Sub IsExpandedChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim se As SlidingExpander = d
            If Not se._doIgnoreExpandCollapse AndAlso Not se.PART_ContentContainer Is Nothing Then
                If e.NewValue Then
                    se.PART_ContentContainer.Height = Double.NaN
                    se.PART_ContentContainer.Visibility = Visibility.Visible
                Else
                    se.PART_ContentContainer.Height = 0
                    se.PART_ContentContainer.Visibility = Visibility.Collapsed
                End If
            End If
        End Sub
    End Class
End Namespace
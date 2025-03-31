Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
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

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SlidingExpander), New FrameworkPropertyMetadata(GetType(SlidingExpander)))
            SlidingExpander.IsExpandedProperty.OverrideMetadata(GetType(SlidingExpander), New FrameworkPropertyMetadata(True, AddressOf IsExpandedChanged))
        End Sub

        Private Const ANIMATIONMS As Integer = 150
        Private Const ANIMATIONSTEPS As Integer = 25

        Private PART_ContentContainer As ScrollViewer
        Private PART_ContentPresenter As ContentPresenter
        Private PART_Title As ContentPresenter
        Private PART_ToggleButton As ToggleButton
        Private _timer As DispatcherTimer
        Private _currentStep As Double = 0
        Private _totalHeight As Double = 0
        Private _doIgnoreExpandCollapse As Boolean
        Private _vwp As VirtualizingWrapPanel
        Private _lastExpandCollapse As DateTime = DateTime.Now

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ContentContainer = Template.FindName("PART_ContentContainer", Me)
            Me.PART_ContentPresenter = Template.FindName("PART_ContentPresenter", Me)
            Me.PART_Title = Template.FindName("PART_Title", Me)
            Me.PART_ToggleButton = Template.FindName("PART_ToggleButton", Me)

            _timer = New DispatcherTimer()

            AddHandler Me.PART_Title.PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    getWrapPanel()
                    Dim listBox As ListBox = _vwp.ItemsControl
                    If _vwp.Items.Count > listBox.SelectedItems.Count AndAlso Not Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then
                        If Not Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then listBox.SelectedItems.Clear()
                        For Each item In _vwp.Items
                            If Not listBox.SelectedItems.Contains(item) Then
                                listBox.SelectedItems.Add(item)
                            End If
                        Next
                    End If
                    Me.PART_ToggleButton.Focus()
                End Sub

            AddHandler Me.PART_ToggleButton.Unchecked,
                Sub(s As Object, e As RoutedEventArgs)
                    getWrapPanel()
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
                    getWrapPanel()
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

        Private Sub getWrapPanel()
            If Not _vwp Is Nothing Then Return
            _vwp = UIHelper.FindVisualChildren(Of VirtualizingWrapPanel)(Me.PART_ContentContainer).ToList()(0)
            If Not _vwp Is Nothing Then
                _vwp.Margin = CType(CType(Me.PART_ContentContainer.Content, ContentPresenter).Content, FrameworkElement).Margin
                Me.PART_ContentContainer.Content = _vwp
                Me.PART_ContentContainer.CanContentScroll = True
                _vwp.ScrollOwner = Me.PART_ContentContainer
            End If
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

        Private Shared Sub IsExpandedChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim se As SlidingExpander = d
            se._lastExpandCollapse = DateTime.Now
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
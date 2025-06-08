Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Laila.Shell.Controls
Imports Laila.Shell.Converters

Namespace Themes
    Public Class StandardColors
        Inherits DependencyObject

        Public Shared ReadOnly DoUseLightThemeProperty As DependencyProperty = DependencyProperty.Register("DoUseLightTheme", GetType(Boolean), GetType(StandardColors), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoUseLightThemeOverrideProperty As DependencyProperty = DependencyProperty.Register("DoUseLightThemeOverride", GetType(Boolean?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoUseLightThemeOverrideChanged))
        Public Shared ReadOnly AccentProperty As DependencyProperty = DependencyProperty.Register("Accent", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentColorProperty As DependencyProperty = DependencyProperty.Register("AccentColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentColorOverrideProperty As DependencyProperty = DependencyProperty.Register("AccentColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnAccentColorOverrideChanged))
        Public Shared ReadOnly AccentDarkerProperty As DependencyProperty = DependencyProperty.Register("AccentDarker", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentDarkerColorProperty As DependencyProperty = DependencyProperty.Register("AccentDarkerColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentLighterProperty As DependencyProperty = DependencyProperty.Register("AccentLighter", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentLighterColorProperty As DependencyProperty = DependencyProperty.Register("AccentLighterColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ForegroundProperty As DependencyProperty = DependencyProperty.Register("Foreground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Black, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ForegroundColorProperty As DependencyProperty = DependencyProperty.Register("ForegroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.Black, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ForegroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ForegroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnForegroundColorOverrideChanged))
        Public Shared ReadOnly GrayForegroundProperty As DependencyProperty = DependencyProperty.Register("GrayForeground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Black, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly GrayForegroundColorProperty As DependencyProperty = DependencyProperty.Register("GrayForegroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.Black, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly GrayForegroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("GrayForegroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnGrayForegroundColorOverrideChanged))
        Public Shared ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly BackgroundColorProperty As DependencyProperty = DependencyProperty.Register("BackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly BackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("BackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnBackgroundColorOverrideChanged))
        Public Shared ReadOnly BackgroundLighterProperty As DependencyProperty = DependencyProperty.Register("BackgroundLighter", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly BackgroundLighterColorProperty As DependencyProperty = DependencyProperty.Register("BackgroundLighterColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly BackgroundLighterColorOverrideProperty As DependencyProperty = DependencyProperty.Register("BackgroundLighterColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnBackgroundLighterColorOverrideChanged))
        Public Shared ReadOnly SelectedBackgroundProperty As DependencyProperty = DependencyProperty.Register("SelectedBackground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedBackgroundColorProperty As DependencyProperty = DependencyProperty.Register("SelectedBackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedBackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("SelectedBackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedBackgroundColorOverrideChanged))
        Public Shared ReadOnly MouseOverBackgroundProperty As DependencyProperty = DependencyProperty.Register("MouseOverBackground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MouseOverBackgroundColorProperty As DependencyProperty = DependencyProperty.Register("MouseOverBackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MouseOverBackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("MouseOverBackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnMouseOverBackgroundColorOverrideChanged))
        Public Shared ReadOnly ItemMouseOverBackgroundProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBackground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemMouseOverBackgroundColorProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemMouseOverBackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemMouseOverBackgroundColorOverrideChanged))
        Public Shared ReadOnly ItemMouseOverBorderProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBorder", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemMouseOverBorderColorProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBorderColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemMouseOverBorderColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemMouseOverBorderColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemMouseOverBorderColorOverrideChanged))
        Public Shared ReadOnly ItemSelectedInactiveBackgroundProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBackground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedInactiveBackgroundColorProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedInactiveBackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemSelectedInactiveBackgroundColorOverrideChanged))
        Public Shared ReadOnly ItemSelectedInactiveBorderProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBorder", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedInactiveBorderColorProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBorderColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedInactiveBorderColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedInactiveBorderColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemSelectedInactiveBorderColorOverrideChanged))
        Public Shared ReadOnly ItemSelectedActiveBackgroundProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBackground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedActiveBackgroundColorProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBackgroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedActiveBackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBackgroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemSelectedActiveBackgroundColorOverrideChanged))
        Public Shared ReadOnly ItemSelectedActiveBorderProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBorder", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedActiveBorderColorProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBorderColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemSelectedActiveBorderColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ItemSelectedActiveBorderColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnItemSelectedActiveBorderColorOverrideChanged))

        Public Sub New()
            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoUseLightTheme"
                            setDoUseLightTheme()
                        Case "WindowsAccentColor"
                            setAccentColor()
                    End Select
                End Sub
            setDoUseLightTheme()
            setAccentColor()
        End Sub

        Public Property DoUseLightTheme As Boolean
            Get
                Return GetValue(DoUseLightThemeProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoUseLightThemeProperty, value)
            End Set
        End Property

        Private Sub setDoUseLightTheme()
            If Me.DoUseLightThemeOverride.HasValue Then
                Me.DoUseLightTheme = Me.DoUseLightThemeOverride.Value
            Else
                Me.DoUseLightTheme = Shell.Settings.DoUseLightTheme
            End If
            setForegroundColor()
            setGrayForegroundColor()
            setBackgroundColor()
            setBackgroundLighterColor()
            setSelectedBackgroundColor()
            setMouseOverBackgroundColor()
            setItemMouseOverBackgroundColor()
            setItemMouseOverBorderColor()
            setItemSelectedInactiveBackgroundColor()
            setItemSelectedInactiveBorderColor()
            setItemSelectedActiveBackgroundColor()
            setItemSelectedActiveBorderColor()
        End Sub

        Public Property DoUseLightThemeOverride As Boolean?
            Get
                Return GetValue(DoUseLightThemeOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoUseLightThemeOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoUseLightThemeOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setDoUseLightTheme()
        End Sub

        Public Property Accent As Brush
            Get
                Return GetValue(AccentProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentProperty, value)
            End Set
        End Property

        Public Property AccentColor As Color
            Get
                Return GetValue(AccentColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentColorProperty, value)
            End Set
        End Property

        Public Property AccentDarker As Brush
            Get
                Return GetValue(AccentDarkerProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentDarkerProperty, value)
            End Set
        End Property

        Public Property AccentDarkerColor As Color
            Get
                Return GetValue(AccentDarkerColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentDarkerColorProperty, value)
            End Set
        End Property

        Public Property AccentLighter As Brush
            Get
                Return GetValue(AccentLighterProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentLighterProperty, value)
            End Set
        End Property

        Public Property AccentLighterColor As Color
            Get
                Return GetValue(AccentLighterColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentLighterColorProperty, value)
            End Set
        End Property

        Private Sub setAccentColor()
            If Me.AccentColorOverride.HasValue Then
                Me.AccentColor = Me.AccentColorOverride.Value
            Else
                Me.AccentColor = Shell.Settings.WindowsAccentColor
            End If
            Me.Accent = New SolidColorBrush(Me.AccentColor)
            Me.AccentDarkerColor = New LightnessColorConverter().Convert(Me.AccentColor, Nothing, 0.8, Nothing)
            Me.AccentDarker = New SolidColorBrush(Me.AccentDarkerColor)
            Me.AccentLighterColor = New LightnessColorConverter().Convert(Me.AccentColor, Nothing, 1.2, Nothing)
            Me.AccentLighter = New SolidColorBrush(Me.AccentLighterColor)
        End Sub

        Public Property AccentColorOverride As Color?
            Get
                Return GetValue(AccentColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(AccentColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnAccentColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setAccentColor()
        End Sub

        Public Property Foreground As Brush
            Get
                Return GetValue(ForegroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ForegroundProperty, value)
            End Set
        End Property

        Public Property ForegroundColor As Color
            Get
                Return GetValue(ForegroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ForegroundColorProperty, value)
            End Set
        End Property

        Private Sub setForegroundColor()
            If Me.ForegroundColorOverride.HasValue Then
                Me.ForegroundColor = Me.ForegroundColorOverride.Value
            Else
                Me.ForegroundColor = If(Me.DoUseLightTheme, Colors.Black, Colors.White)
            End If
            Me.Foreground = New SolidColorBrush(Me.ForegroundColor)
        End Sub

        Public Property ForegroundColorOverride As Color?
            Get
                Return GetValue(ForegroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ForegroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnForegroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setForegroundColor()
        End Sub

        Public Property GrayForeground As Brush
            Get
                Return GetValue(GrayForegroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(GrayForegroundProperty, value)
            End Set
        End Property

        Public Property GrayForegroundColor As Color
            Get
                Return GetValue(GrayForegroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(GrayForegroundColorProperty, value)
            End Set
        End Property

        Private Sub setGrayForegroundColor()
            If Me.GrayForegroundColorOverride.HasValue Then
                Me.GrayForegroundColor = Me.GrayForegroundColorOverride.Value
            Else
                Me.GrayForegroundColor = If(Me.DoUseLightTheme, Colors.Gray, Colors.Silver)
            End If
            Me.GrayForeground = New SolidColorBrush(Me.GrayForegroundColor)
        End Sub

        Public Property GrayForegroundColorOverride As Color?
            Get
                Return GetValue(GrayForegroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(GrayForegroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnGrayForegroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setGrayForegroundColor()
        End Sub

        Public Property Background As Brush
            Get
                Return GetValue(BackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(BackgroundProperty, value)
            End Set
        End Property

        Public Property BackgroundColor As Color
            Get
                Return GetValue(BackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(BackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setBackgroundColor()
            If Me.BackgroundColorOverride.HasValue Then
                Me.BackgroundColor = Me.BackgroundColorOverride.Value
            Else
                Me.BackgroundColor = If(Me.DoUseLightTheme, Colors.White, Colors.Black)
            End If
            Me.Background = New SolidColorBrush(Me.BackgroundColor)
        End Sub

        Public Property BackgroundColorOverride As Color?
            Get
                Return GetValue(BackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(BackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setBackgroundColor()
        End Sub

        Public Property BackgroundLighter As Brush
            Get
                Return GetValue(BackgroundLighterProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(BackgroundLighterProperty, value)
            End Set
        End Property

        Public Property BackgroundLighterColor As Color
            Get
                Return GetValue(BackgroundLighterColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(BackgroundLighterColorProperty, value)
            End Set
        End Property

        Private Sub setBackgroundLighterColor()
            If Me.BackgroundLighterColorOverride.HasValue Then
                Me.BackgroundLighterColor = Me.BackgroundLighterColorOverride.Value
            Else
                Me.BackgroundLighterColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#CCFFFFFF"), ColorConverter.ConvertFromString("#33FFFFFF"))
            End If
            Me.BackgroundLighter = New SolidColorBrush(Me.BackgroundLighterColor)
        End Sub

        Public Property BackgroundLighterColorOverride As Color?
            Get
                Return GetValue(BackgroundLighterColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(BackgroundLighterColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnBackgroundLighterColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setBackgroundLighterColor()
        End Sub

        Public Property SelectedBackground As Brush
            Get
                Return GetValue(SelectedBackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(SelectedBackgroundProperty, value)
            End Set
        End Property

        Public Property SelectedBackgroundColor As Color
            Get
                Return GetValue(SelectedBackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(SelectedBackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setSelectedBackgroundColor()
            If Me.SelectedBackgroundColorOverride.HasValue Then
                Me.SelectedBackgroundColor = Me.SelectedBackgroundColorOverride.Value
            Else
                Me.SelectedBackgroundColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#EFEFEF"), ColorConverter.ConvertFromString("#333333"))
            End If
            Me.SelectedBackground = New SolidColorBrush(Me.SelectedBackgroundColor)
        End Sub

        Public Property SelectedBackgroundColorOverride As Color?
            Get
                Return GetValue(SelectedBackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(SelectedBackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnSelectedBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setSelectedBackgroundColor()
        End Sub

        Public Property MouseOverBackground As Brush
            Get
                Return GetValue(MouseOverBackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(MouseOverBackgroundProperty, value)
            End Set
        End Property

        Public Property MouseOverBackgroundColor As Color
            Get
                Return GetValue(MouseOverBackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(MouseOverBackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setMouseOverBackgroundColor()
            If Me.MouseOverBackgroundColorOverride.HasValue Then
                Me.MouseOverBackgroundColor = Me.MouseOverBackgroundColorOverride.Value
            Else
                Me.MouseOverBackgroundColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#EFEFEF"), ColorConverter.ConvertFromString("#222222"))
            End If
            Me.MouseOverBackground = New SolidColorBrush(Me.MouseOverBackgroundColor)
        End Sub

        Public Property MouseOverBackgroundColorOverride As Color?
            Get
                Return GetValue(MouseOverBackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(MouseOverBackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnMouseOverBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setMouseOverBackgroundColor()
        End Sub

        Public Property ItemMouseOverBackground As Brush
            Get
                Return GetValue(ItemMouseOverBackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemMouseOverBackgroundProperty, value)
            End Set
        End Property

        Public Property ItemMouseOverBackgroundColor As Color
            Get
                Return GetValue(ItemMouseOverBackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemMouseOverBackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setItemMouseOverBackgroundColor()
            If Me.ItemMouseOverBackgroundColorOverride.HasValue Then
                Me.ItemMouseOverBackgroundColor = Me.ItemMouseOverBackgroundColorOverride.Value
            Else
                Me.ItemMouseOverBackgroundColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#1F26A0DA"), ColorConverter.ConvertFromString("#444444"))
            End If
            Me.ItemMouseOverBackground = New SolidColorBrush(Me.ItemMouseOverBackgroundColor)
        End Sub

        Public Property ItemMouseOverBackgroundColorOverride As Color?
            Get
                Return GetValue(ItemMouseOverBackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemMouseOverBackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemMouseOverBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemMouseOverBackgroundColor()
        End Sub

        Public Property ItemMouseOverBorder As Brush
            Get
                Return GetValue(ItemMouseOverBorderProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemMouseOverBorderProperty, value)
            End Set
        End Property

        Public Property ItemMouseOverBorderColor As Color
            Get
                Return GetValue(ItemMouseOverBorderColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemMouseOverBorderColorProperty, value)
            End Set
        End Property

        Private Sub setItemMouseOverBorderColor()
            If Me.ItemMouseOverBorderColorOverride.HasValue Then
                Me.ItemMouseOverBorderColor = Me.ItemMouseOverBorderColorOverride.Value
            Else
                Me.ItemMouseOverBorderColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#1F26A0DA"), ColorConverter.ConvertFromString("#444444"))
            End If
            Me.ItemMouseOverBorder = New SolidColorBrush(Me.ItemMouseOverBorderColor)
        End Sub

        Public Property ItemMouseOverBorderColorOverride As Color?
            Get
                Return GetValue(ItemMouseOverBorderColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemMouseOverBorderColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemMouseOverBorderColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemMouseOverBorderColor()
        End Sub

        Public Property ItemSelectedInactiveBackground As Brush
            Get
                Return GetValue(ItemSelectedInactiveBackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemSelectedInactiveBackgroundProperty, value)
            End Set
        End Property

        Public Property ItemSelectedInactiveBackgroundColor As Color
            Get
                Return GetValue(ItemSelectedInactiveBackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemSelectedInactiveBackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setItemSelectedInactiveBackgroundColor()
            If Me.ItemSelectedInactiveBackgroundColorOverride.HasValue Then
                Me.ItemSelectedInactiveBackgroundColor = Me.ItemSelectedInactiveBackgroundColorOverride.Value
            Else
                Me.ItemSelectedInactiveBackgroundColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#cfcfcf"), ColorConverter.ConvertFromString("#333333"))
            End If
            Me.ItemSelectedInactiveBackground = New SolidColorBrush(Me.ItemSelectedInactiveBackgroundColor)
        End Sub

        Public Property ItemSelectedInactiveBackgroundColorOverride As Color?
            Get
                Return GetValue(ItemSelectedInactiveBackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemSelectedInactiveBackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemSelectedInactiveBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemSelectedInactiveBackgroundColor()
        End Sub

        Public Property ItemSelectedInactiveBorder As Brush
            Get
                Return GetValue(ItemSelectedInactiveBorderProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemSelectedInactiveBorderProperty, value)
            End Set
        End Property

        Public Property ItemSelectedInactiveBorderColor As Color
            Get
                Return GetValue(ItemSelectedInactiveBorderColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemSelectedInactiveBorderColorProperty, value)
            End Set
        End Property

        Private Sub setItemSelectedInactiveBorderColor()
            If Me.ItemSelectedInactiveBorderColorOverride.HasValue Then
                Me.ItemSelectedInactiveBorderColor = Me.ItemSelectedInactiveBorderColorOverride.Value
            Else
                Me.ItemSelectedInactiveBorderColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#cfcfcf"), ColorConverter.ConvertFromString("#333333"))
            End If
            Me.ItemSelectedInactiveBorder = New SolidColorBrush(Me.ItemSelectedInactiveBorderColor)
        End Sub

        Public Property ItemSelectedInactiveBorderColorOverride As Color?
            Get
                Return GetValue(ItemSelectedInactiveBorderColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemSelectedInactiveBorderColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemSelectedInactiveBorderColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemSelectedInactiveBorderColor()
        End Sub

        Public Property ItemSelectedActiveBackground As Brush
            Get
                Return GetValue(ItemSelectedActiveBackgroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemSelectedActiveBackgroundProperty, value)
            End Set
        End Property

        Public Property ItemSelectedActiveBackgroundColor As Color
            Get
                Return GetValue(ItemSelectedActiveBackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemSelectedActiveBackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setItemSelectedActiveBackgroundColor()
            If Me.ItemSelectedActiveBackgroundColorOverride.HasValue Then
                Me.ItemSelectedActiveBackgroundColor = Me.ItemSelectedActiveBackgroundColorOverride.Value
            Else
                Me.ItemSelectedActiveBackgroundColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#3D26A0DA"), ColorConverter.ConvertFromString("#5F5F5F"))
            End If
            Me.ItemSelectedActiveBackground = New SolidColorBrush(Me.ItemSelectedActiveBackgroundColor)
        End Sub

        Public Property ItemSelectedActiveBackgroundColorOverride As Color?
            Get
                Return GetValue(ItemSelectedActiveBackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemSelectedActiveBackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemSelectedActiveBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemSelectedActiveBackgroundColor()
        End Sub

        Public Property ItemSelectedActiveBorder As Brush
            Get
                Return GetValue(ItemSelectedActiveBorderProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ItemSelectedActiveBorderProperty, value)
            End Set
        End Property

        Public Property ItemSelectedActiveBorderColor As Color
            Get
                Return GetValue(ItemSelectedActiveBorderColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ItemSelectedActiveBorderColorProperty, value)
            End Set
        End Property

        Private Sub setItemSelectedActiveBorderColor()
            If Me.ItemSelectedActiveBorderColorOverride.HasValue Then
                Me.ItemSelectedActiveBorderColor = Me.ItemSelectedActiveBorderColorOverride.Value
            Else
                Me.ItemSelectedActiveBorderColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#FF26A0DA"), ColorConverter.ConvertFromString("#ffffff"))
            End If
            Me.ItemSelectedActiveBorder = New SolidColorBrush(Me.ItemSelectedActiveBorderColor)
        End Sub

        Public Property ItemSelectedActiveBorderColorOverride As Color?
            Get
                Return GetValue(ItemSelectedActiveBorderColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ItemSelectedActiveBorderColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnItemSelectedActiveBorderColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setItemSelectedActiveBorderColor()
        End Sub
    End Class
End Namespace
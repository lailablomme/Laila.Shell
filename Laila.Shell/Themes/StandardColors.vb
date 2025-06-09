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
        Public Shared ReadOnly ThreeDBorderLightProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderLight", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderLightColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderLightColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderLightColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderLightColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBorderLightColorOverrideChanged))
        Public Shared ReadOnly ThreeDBorderMediumProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderMedium", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderMediumColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderMediumColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderMediumColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderMediumColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBorderMediumColorOverrideChanged))
        Public Shared ReadOnly ThreeDBorderDarkProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderDark", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderDarkColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderDarkColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBorderDarkColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBorderDarkColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBorderDarkColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundLightProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLight", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundLightColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundMediumProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMedium", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundMediumColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundLightMouseOverProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightMouseOver", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightMouseOverColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightMouseOverColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightMouseOverColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightMouseOverColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundLightMouseOverColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundMediumMouseOverProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumMouseOver", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumMouseOverColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumMouseOverColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumMouseOverColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumMouseOverColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundMediumMouseOverColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundLightPressedProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightPressed", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightPressedColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightPressedColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundLightPressedColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundLightPressedColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundLightPressedColorOverrideChanged))
        Public Shared ReadOnly ThreeDBackgroundMediumPressedProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumPressed", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumPressedColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumPressedColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDBackgroundMediumPressedColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDBackgroundMediumPressedColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDBackgroundMediumPressedColorOverrideChanged))
        Public Shared ReadOnly ThreeDForegroundProperty As DependencyProperty = DependencyProperty.Register("ThreeDForeground", GetType(Brush), GetType(StandardColors), New FrameworkPropertyMetadata(Brushes.Silver, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDForegroundColorProperty As DependencyProperty = DependencyProperty.Register("ThreeDForegroundColor", GetType(Color), GetType(StandardColors), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ThreeDForegroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ThreeDForegroundColorOverride", GetType(Color?), GetType(StandardColors), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnThreeDForegroundColorOverrideChanged))

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
            setThreeDBackgroundLightColor()
            setThreeDBackgroundLightMouseOverColor()
            setThreeDBackgroundLightPressedColor()
            setThreeDBackgroundMediumColor()
            setThreeDBackgroundMediumMouseOverColor()
            setThreeDBackgroundMediumPressedColor()
            setThreeDBorderLightColor()
            setThreeDBorderMediumColor()
            setThreeDBorderDarkColor()
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
                Me.BackgroundColor = If(Me.DoUseLightTheme, Colors.White, ColorConverter.ConvertFromString("#181818"))
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
                Me.ItemSelectedActiveBorderColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#FF26A0DA"), Colors.Silver)
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

        Public Property ThreeDBorderLight As Brush
            Get
                Return GetValue(ThreeDBorderLightProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBorderLightProperty, value)
            End Set
        End Property

        Public Property ThreeDBorderLightColor As Color
            Get
                Return GetValue(ThreeDBorderLightColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBorderLightColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBorderLightColor()
            If Me.ThreeDBorderLightColorOverride.HasValue Then
                Me.ThreeDBorderLightColor = Me.ThreeDBorderLightColorOverride.Value
            Else
                Me.ThreeDBorderLightColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#EEEEEE"), ColorConverter.ConvertFromString("#EEEEEE"))
            End If
            Me.ThreeDBorderLight = New SolidColorBrush(Me.ThreeDBorderLightColor)
        End Sub

        Public Property ThreeDBorderLightColorOverride As Color?
            Get
                Return GetValue(ThreeDBorderLightColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBorderLightColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBorderLightColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBorderLightColor()
        End Sub

        Public Property ThreeDBorderMedium As Brush
            Get
                Return GetValue(ThreeDBorderMediumProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBorderMediumProperty, value)
            End Set
        End Property

        Public Property ThreeDBorderMediumColor As Color
            Get
                Return GetValue(ThreeDBorderMediumColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBorderMediumColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBorderMediumColor()
            If Me.ThreeDBorderMediumColorOverride.HasValue Then
                Me.ThreeDBorderMediumColor = Me.ThreeDBorderMediumColorOverride.Value
            Else
                Me.ThreeDBorderMediumColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#666666"), ColorConverter.ConvertFromString("#999999"))
            End If
            Me.ThreeDBorderMedium = New SolidColorBrush(Me.ThreeDBorderMediumColor)
        End Sub

        Public Property ThreeDBorderMediumColorOverride As Color?
            Get
                Return GetValue(ThreeDBorderMediumColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBorderMediumColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBorderMediumColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBorderMediumColor()
        End Sub

        Public Property ThreeDBorderDark As Brush
            Get
                Return GetValue(ThreeDBorderDarkProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBorderDarkProperty, value)
            End Set
        End Property

        Public Property ThreeDBorderDarkColor As Color
            Get
                Return GetValue(ThreeDBorderDarkColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBorderDarkColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBorderDarkColor()
            If Me.ThreeDBorderDarkColorOverride.HasValue Then
                Me.ThreeDBorderDarkColor = Me.ThreeDBorderDarkColorOverride.Value
            Else
                Me.ThreeDBorderDarkColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#666666"), ColorConverter.ConvertFromString("#999999"))
            End If
            Me.ThreeDBorderDark = New SolidColorBrush(Me.ThreeDBorderDarkColor)
        End Sub

        Public Property ThreeDBorderDarkColorOverride As Color?
            Get
                Return GetValue(ThreeDBorderDarkColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBorderDarkColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBorderDarkColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBorderDarkColor()
        End Sub

        Public Property ThreeDBackgroundLight As Brush
            Get
                Return GetValue(ThreeDBackgroundLightProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundLightProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundLightColor As Color
            Get
                Return GetValue(ThreeDBackgroundLightColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundLightColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundLightColor()
            If Me.ThreeDBackgroundLightColorOverride.HasValue Then
                Me.ThreeDBackgroundLightColor = Me.ThreeDBackgroundLightColorOverride.Value
            Else
                Me.ThreeDBackgroundLightColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#DDDDDD"), ColorConverter.ConvertFromString("#444444"))
            End If
            Me.ThreeDBackgroundLight = New SolidColorBrush(Me.ThreeDBackgroundLightColor)
        End Sub

        Public Property ThreeDBackgroundLightColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundLightColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundLightColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundLightColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundLightColor()
        End Sub

        Public Property ThreeDBackgroundMedium As Brush
            Get
                Return GetValue(ThreeDBackgroundMediumProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundMediumProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundMediumColor As Color
            Get
                Return GetValue(ThreeDBackgroundMediumColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundMediumColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundMediumColor()
            If Me.ThreeDBackgroundMediumColorOverride.HasValue Then
                Me.ThreeDBackgroundMediumColor = Me.ThreeDBackgroundMediumColorOverride.Value
            Else
                Me.ThreeDBackgroundMediumColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#DDDDDD"), ColorConverter.ConvertFromString("#444444"))
            End If
            Me.ThreeDBackgroundMedium = New SolidColorBrush(Me.ThreeDBackgroundMediumColor)
        End Sub

        Public Property ThreeDBackgroundMediumColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundMediumColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundMediumColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundMediumColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundMediumColor()
        End Sub

        Public Property ThreeDBackgroundLightMouseOver As Brush
            Get
                Return GetValue(ThreeDBackgroundLightMouseOverProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundLightMouseOverProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundLightMouseOverColor As Color
            Get
                Return GetValue(ThreeDBackgroundLightMouseOverColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundLightMouseOverColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundLightMouseOverColor()
            If Me.ThreeDBackgroundLightMouseOverColorOverride.HasValue Then
                Me.ThreeDBackgroundLightMouseOverColor = Me.ThreeDBackgroundLightMouseOverColorOverride.Value
            Else
                Me.ThreeDBackgroundLightMouseOverColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#BBBBBB"), ColorConverter.ConvertFromString("#333333"))
            End If
            Me.ThreeDBackgroundLightMouseOver = New SolidColorBrush(Me.ThreeDBackgroundLightMouseOverColor)
        End Sub

        Public Property ThreeDBackgroundLightMouseOverColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundLightMouseOverColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundLightMouseOverColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundLightMouseOverColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundLightMouseOverColor()
        End Sub

        Public Property ThreeDBackgroundMediumMouseOver As Brush
            Get
                Return GetValue(ThreeDBackgroundMediumMouseOverProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundMediumMouseOverProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundMediumMouseOverColor As Color
            Get
                Return GetValue(ThreeDBackgroundMediumMouseOverColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundMediumMouseOverColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundMediumMouseOverColor()
            If Me.ThreeDBackgroundMediumMouseOverColorOverride.HasValue Then
                Me.ThreeDBackgroundMediumMouseOverColor = Me.ThreeDBackgroundMediumMouseOverColorOverride.Value
            Else
                Me.ThreeDBackgroundMediumMouseOverColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#BBBBBB"), ColorConverter.ConvertFromString("#333333"))
            End If
            Me.ThreeDBackgroundMediumMouseOver = New SolidColorBrush(Me.ThreeDBackgroundMediumMouseOverColor)
        End Sub

        Public Property ThreeDBackgroundMediumMouseOverColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundMediumMouseOverColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundMediumMouseOverColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundMediumMouseOverColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundMediumMouseOverColor()
        End Sub

        Public Property ThreeDBackgroundLightPressed As Brush
            Get
                Return GetValue(ThreeDBackgroundLightPressedProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundLightPressedProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundLightPressedColor As Color
            Get
                Return GetValue(ThreeDBackgroundLightPressedColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundLightPressedColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundLightPressedColor()
            If Me.ThreeDBackgroundLightPressedColorOverride.HasValue Then
                Me.ThreeDBackgroundLightPressedColor = Me.ThreeDBackgroundLightPressedColorOverride.Value
            Else
                Me.ThreeDBackgroundLightPressedColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#AAAAAA"), ColorConverter.ConvertFromString("#222222"))
            End If
            Me.ThreeDBackgroundLightPressed = New SolidColorBrush(Me.ThreeDBackgroundLightPressedColor)
        End Sub

        Public Property ThreeDBackgroundLightPressedColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundLightPressedColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundLightPressedColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundLightPressedColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundLightPressedColor()
        End Sub

        Public Property ThreeDBackgroundMediumPressed As Brush
            Get
                Return GetValue(ThreeDBackgroundMediumPressedProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDBackgroundMediumPressedProperty, value)
            End Set
        End Property

        Public Property ThreeDBackgroundMediumPressedColor As Color
            Get
                Return GetValue(ThreeDBackgroundMediumPressedColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDBackgroundMediumPressedColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDBackgroundMediumPressedColor()
            If Me.ThreeDBackgroundMediumPressedColorOverride.HasValue Then
                Me.ThreeDBackgroundMediumPressedColor = Me.ThreeDBackgroundMediumPressedColorOverride.Value
            Else
                Me.ThreeDBackgroundMediumPressedColor = If(Me.DoUseLightTheme, ColorConverter.ConvertFromString("#AAAAAA"), ColorConverter.ConvertFromString("#222222"))
            End If
            Me.ThreeDBackgroundMediumPressed = New SolidColorBrush(Me.ThreeDBackgroundMediumPressedColor)
        End Sub

        Public Property ThreeDBackgroundMediumPressedColorOverride As Color?
            Get
                Return GetValue(ThreeDBackgroundMediumPressedColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDBackgroundMediumPressedColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDBackgroundMediumPressedColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDBackgroundMediumPressedColor()
        End Sub

        Public Property ThreeDForeground As Brush
            Get
                Return GetValue(ThreeDForegroundProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(ThreeDForegroundProperty, value)
            End Set
        End Property

        Public Property ThreeDForegroundColor As Color
            Get
                Return GetValue(ThreeDForegroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ThreeDForegroundColorProperty, value)
            End Set
        End Property

        Private Sub setThreeDForegroundColor()
            If Me.ThreeDForegroundColorOverride.HasValue Then
                Me.ThreeDForegroundColor = Me.ThreeDForegroundColorOverride.Value
            Else
                Me.ThreeDForegroundColor = If(Me.DoUseLightTheme, Colors.Black, Colors.Silver)
            End If
            Me.ThreeDForeground = New SolidColorBrush(Me.ThreeDForegroundColor)
        End Sub

        Public Property ThreeDForegroundColorOverride As Color?
            Get
                Return GetValue(ThreeDForegroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ThreeDForegroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnThreeDForegroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As StandardColors = d
            bfv.setThreeDForegroundColor()
        End Sub
    End Class
End Namespace
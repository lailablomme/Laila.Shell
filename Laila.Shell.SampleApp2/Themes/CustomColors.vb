Imports System.ComponentModel
Imports Laila.Shell.Themes

Namespace Themes
    Public Class CustomColors
        Inherits StandardColors

        Public Shared ReadOnly BackgroundFileNameProperty As DependencyProperty = DependencyProperty.Register("BackgroundFileName", GetType(String), GetType(CustomColors), New FrameworkPropertyMetadata("Images/dorabg.png", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Protected Overrides Sub SetDoUseLightTheme()
            MyBase.SetDoUseLightTheme()

            If Me.DoUseLightTheme Then
                Me.BackgroundFileName = "Images/dorabg_light.png"
            Else
                Me.BackgroundFileName = "Images/dorabg.png"
            End If
        End Sub

        Public Property BackgroundFileName As String
            Get
                Return CType(GetValue(BackgroundFileNameProperty), String)
            End Get
            Set(value As String)
                SetValue(BackgroundFileNameProperty, value)
            End Set
        End Property
    End Class
End Namespace
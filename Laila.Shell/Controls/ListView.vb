﻿Imports System.ComponentModel
Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports Laila.Shell.Helpers
Imports WpfToolkit.Controls

Namespace Controls
    Public Class ListView
        Inherits BaseFolderView

        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ListView), New FrameworkPropertyMetadata(GetType(ListView)))
        End Sub

        Protected Overrides Async Sub OnBeforeRestoreScrollOffset()
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                .FirstOrDefault(Function(b) b.Name = "PART_DisplayName")
            point = Me.PointFromScreen(textBlock.PointToScreen(New Point(0, 0)))
            point.X += -2
            point.Y += -1
            size.Width = textBlock.ActualWidth + 4
            size.Height = textBlock.ActualHeight + 2
            textAlignment = TextAlignment.Left
            fontSize = textBlock.FontSize
        End Sub
    End Class
End Namespace
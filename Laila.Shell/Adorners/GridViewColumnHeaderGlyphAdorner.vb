Imports System.Transactions
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media

Namespace Adorners
    Public Class GridViewColumnHeaderGlyphAdorner
        Inherits Adorner

        Public Shared Sub Add(columnHeader As GridViewColumnHeader, code As String, index As Integer, imageSource As String, alignment As HorizontalAlignment)
            Dim adornerLayer As AdornerLayer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(columnHeader)
            Dim adorners As Adorner() = adornerLayer.GetAdorners(columnHeader)
            Dim adorner As GridViewColumnHeaderGlyphAdorner
            If Not adorners Is Nothing AndAlso adorners.Count = 1 AndAlso TypeOf adorners(0) Is GridViewColumnHeaderGlyphAdorner Then
                adorner = adorners(0)
            Else
                adorner = New GridViewColumnHeaderGlyphAdorner(columnHeader)
                adornerLayer.Add(adorner)
            End If
            Dim isc As ImageSourceConverter = New ImageSourceConverter()
            adorner.add(code, New GridViewColumnHeaderGlyph(index, isc.ConvertFromInvariantString(imageSource), alignment))
        End Sub

        Public Shared Sub Remove(columnHeader As GridViewColumnHeader, code As String)
            Dim adornerLayer As AdornerLayer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(columnHeader)
            Dim adorners As Adorner() = adornerLayer.GetAdorners(columnHeader)
            Dim adorner As GridViewColumnHeaderGlyphAdorner
            If Not adorners Is Nothing AndAlso adorners.Count = 1 AndAlso TypeOf adorners(0) Is GridViewColumnHeaderGlyphAdorner Then
                adorner = adorners(0)
                adorner.remove(code)
            End If
        End Sub

        Private _columnHeader As GridViewColumnHeader
        Private _glyphs As Dictionary(Of String, GridViewColumnHeaderGlyph) = New Dictionary(Of String, GridViewColumnHeaderGlyph)

        Public Sub New(columnHeader As GridViewColumnHeader)
            MyBase.New(columnHeader)
            Me.IsHitTestVisible = False

            _columnHeader = columnHeader
        End Sub

        Private Sub add(code As String, glyph As GridViewColumnHeaderGlyph)
            If _glyphs.Keys.Contains(code) Then
                _glyphs(code) = glyph
            Else
                _glyphs.Add(code, glyph)
            End If
            update()
        End Sub

        Private Sub remove(code As String)
            _glyphs.Remove(code)
            update()
        End Sub

        Private Sub update()
            Me.InvalidateVisual()

            Dim c As Integer = _glyphs.Keys.Count, left As Double = 0, right As Double = 0
            If Not _glyphs.Values.FirstOrDefault(Function(g) g.Alignment = HorizontalAlignment.Right) Is Nothing Then
                right = _glyphs.Values.Where(Function(g) g.Alignment = HorizontalAlignment.Right).Count * 12
            End If
            If Not _glyphs.Values.FirstOrDefault(Function(g) g.Alignment = HorizontalAlignment.Left) Is Nothing Then
                left = _glyphs.Values.Where(Function(g) g.Alignment = HorizontalAlignment.Left).Count * 12
            End If
            Dim v As Thickness = _columnHeader.Padding
            _columnHeader.Padding = New Thickness(left, v.Top, right, v.Bottom)
        End Sub

        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            MyBase.OnRender(drawingContext)

            Dim z As Integer = 0
            For Each g In _glyphs.Values.Where(Function(gl) gl.Alignment = HorizontalAlignment.Right).OrderBy(Function(i) i.Index).ToList()
                z += 1
                Dim y As Double = (_columnHeader.ActualHeight - g.Glyph.Height) / 2
                Dim x As Double = _columnHeader.ActualWidth - z * 12
                Dim rect As Rect = New Rect(x, y, g.Glyph.Width, g.Glyph.Height)
                drawingContext.DrawImage(g.Glyph, rect)
            Next
            z = 0
            For Each g In _glyphs.Values.Where(Function(gl) gl.Alignment = HorizontalAlignment.Left).OrderBy(Function(i) i.Index).ToList()
                z += 1
                Dim y As Double = (_columnHeader.ActualHeight - g.Glyph.Height) / 2
                Dim x As Double = (z - 1) * 12
                Dim rect As Rect = New Rect(x, y, g.Glyph.Width, g.Glyph.Height)
                drawingContext.DrawImage(g.Glyph, rect)
            Next
            z = 0
            For Each g In _glyphs.Values.Where(Function(gl) gl.Alignment = HorizontalAlignment.Center).OrderBy(Function(i) i.Index).ToList()
                z += 1
                Dim y As Double = 0
                Dim x As Double = (_columnHeader.ActualWidth - _glyphs.Values.Where(Function(gl) gl.Alignment = HorizontalAlignment.Center).Sum(Function(gl) gl.Glyph.Width)) / 2 + _glyphs.Values.Take(z - 1).Sum(Function(gl) gl.Glyph.Width)
                Dim rect As Rect = New Rect(x, y, g.Glyph.Width, g.Glyph.Height)
                drawingContext.DrawImage(g.Glyph, rect)
            Next
        End Sub

        Private Class GridViewColumnHeaderGlyph
            Public Property Index As Integer
            Public Property Glyph As ImageSource
            Public Property Alignment As HorizontalAlignment

            Public Sub New(index As Integer, glyph As ImageSource, alignment As HorizontalAlignment)
                Me.Index = index
                Me.Glyph = glyph
                Me.Alignment = alignment
            End Sub
        End Class
    End Class
End Namespace

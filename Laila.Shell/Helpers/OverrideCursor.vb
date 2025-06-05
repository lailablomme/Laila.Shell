Imports System.Windows.Input
Imports System.Windows.Threading

Namespace Helpers
    Public Class OverrideCursor
        Implements IDisposable

        Private Shared _hist As Dictionary(Of Integer, Cursor) = New Dictionary(Of Integer, Cursor)()

        Private _id As Integer

        Public Sub New()
            Me.New(Cursors.Wait)
        End Sub

        Public Sub New(cursor As Cursor)
            Dim app As System.Windows.Application = System.Windows.Application.Current
            If Not app Is Nothing Then
                app.Dispatcher.Invoke(
                    Sub()
                        If _hist.Keys.Count > 0 Then
                            _id = _hist.Keys.Max() + 1
                        Else
                            _id = 0
                        End If
                        _hist(_id) = cursor
                        Mouse.OverrideCursor = cursor
                    End Sub, DispatcherPriority.Send)
            End If
        End Sub

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                Me.disposedValue = True

                If disposing Then
                    Dim app As System.Windows.Application = System.Windows.Application.Current
                    If Not app Is Nothing Then
                        app.Dispatcher.Invoke(
                            Sub()
                                Dim c As Cursor = Nothing
                                _hist.Remove(_id)
                                If _hist.Keys.Count > 0 Then
                                    Dim max As Integer = _hist.Keys.Max()
                                    c = _hist(max)
                                End If
                                Mouse.OverrideCursor = c
                            End Sub, DispatcherPriority.Send)
                    End If
                End If
            End If
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace
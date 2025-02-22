Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Namespace Helpers
    Public Class NotifyPropertyChangedBase
        Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Property IsNotifying As Boolean = True

        Public Sub NotifyOfPropertyChange(propertyName As String)
            If Me.IsNotifying Then
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
            End If
        End Sub

        Protected Sub SetValue(Of T)(ByRef member As T, value As T, <CallerMemberName> Optional propertyName As String = "")
            member = value
            NotifyOfPropertyChange(propertyName)
        End Sub
    End Class
End Namespace
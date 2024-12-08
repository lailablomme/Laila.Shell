Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private Const HIDEKNOWNFILEEXTENSIONS_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const HIDEKNOWNFILEEXTENSIONS_VALUENAME As String = "HideFileExt"

    Public Property DoHideKnownFileExtensions As Boolean
        Get
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(HIDEKNOWNFILEEXTENSIONS_KEY)
                If Not key Is Nothing Then
                    Dim hideFileExt As Object = key.GetValue(HIDEKNOWNFILEEXTENSIONS_VALUENAME)
                    If Not hideFileExt Is Nothing AndAlso TypeOf hideFileExt Is Integer Then
                        Return CType(hideFileExt, Integer) = 1
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            End Using
        End Get
        Set(value As Boolean)
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(HIDEKNOWNFILEEXTENSIONS_KEY)
                If Not key Is Nothing Then
                    key.SetValue(HIDEKNOWNFILEEXTENSIONS_VALUENAME, If(value, 1, 0))
                Else
                    Using newKey As RegistryKey = Registry.CurrentUser.CreateSubKey(HIDEKNOWNFILEEXTENSIONS_KEY)
                        newKey.SetValue(HIDEKNOWNFILEEXTENSIONS_VALUENAME, If(value, 1, 0))
                    End Using
                End If
            End Using
            Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")
        End Set
    End Property
End Class

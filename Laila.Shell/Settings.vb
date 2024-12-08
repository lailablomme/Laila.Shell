Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private Const HIDEKNOWNFILEEXTENSIONS_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const HIDEKNOWNFILEEXTENSIONS_VALUENAME As String = "HideFileExt"
    Private Const SHOWCHECKBOXESTOSELECT_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const SHOWCHECKBOXESTOSELECT_VALUENAME As String = "AutoCheckSelect"

    Public Property DoHideKnownFileExtensions As Boolean
        Get
            Return Settings.GetRegistryBoolean(HIDEKNOWNFILEEXTENSIONS_KEY, HIDEKNOWNFILEEXTENSIONS_VALUENAME)
        End Get
        Set(value As Boolean)
            Settings.SetRegistryBoolean(HIDEKNOWNFILEEXTENSIONS_KEY, HIDEKNOWNFILEEXTENSIONS_VALUENAME, value)
            Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")
        End Set
    End Property

    Public Property DoShowCheckBoxesToSelect As Boolean
        Get
            Return Settings.GetRegistryBoolean(SHOWCHECKBOXESTOSELECT_KEY, SHOWCHECKBOXESTOSELECT_VALUENAME)
        End Get
        Set(value As Boolean)
            Settings.SetRegistryBoolean(SHOWCHECKBOXESTOSELECT_KEY, SHOWCHECKBOXESTOSELECT_VALUENAME, value)
            Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")
        End Set
    End Property

    Private Shared Function GetRegistryBoolean(key As String, valueName As String) As Boolean
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                Dim hideFileExt As Object = registryKey.GetValue(valueName)
                If Not hideFileExt Is Nothing AndAlso TypeOf hideFileExt Is Integer Then
                    Return CType(hideFileExt, Integer) <> 0
                Else
                    Return True
                End If
            Else
                Return True
            End If
        End Using
    End Function

    Private Shared Sub SetRegistryBoolean(key As String, valueName As String, value As Boolean)
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                registryKey.SetValue(valueName, If(value, 1, 0))
            Else
                Using newRegistryKey As RegistryKey = Registry.CurrentUser.CreateSubKey(key)
                    newRegistryKey.SetValue(valueName, If(value, 1, 0))
                End Using
            End If
        End Using
    End Sub
End Class

Imports System.Runtime.InteropServices
Imports System.Windows.Documents
Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private Const HIDEKNOWNFILEEXTENSIONS_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const HIDEKNOWNFILEEXTENSIONS_VALUENAME As String = "HideFileExt"
    Private Const SHOWCHECKBOXESTOSELECT_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const SHOWCHECKBOXESTOSELECT_VALUENAME As String = "AutoCheckSelect"

    Public Property DoHideKnownFileExtensions As Boolean
        Get
            Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
            Dim val As SSF
            Functions.SHGetSetSettings(val, mask, False)
            Return Not val.HasFlag(SSF.SSF_SHOWEXTENSIONS)
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
            Dim val As SSF = If(value, 0, SSF.SSF_SHOWEXTENSIONS)
            Functions.SHGetSetSettings(val, mask, True)
            Dim h As HRESULT = Marshal.GetLastWin32Error()
            Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")

            Dim list As List(Of Item)
            SyncLock Shell._itemsCacheLock
                list = Shell.ItemsCache.Select(Function(i) i.Item1).ToList()
            End SyncLock
            For Each item In list
                item._displayName = Nothing
                item.ShellItem2.Update(IntPtr.zero)
                item.NotifyOfPropertyChange("DisplayName")
            Next
        End Set
    End Property

    Public Property DoShowCheckBoxesToSelect As Boolean
        Get
            Return Settings.GetRegistryBoolean(SHOWCHECKBOXESTOSELECT_KEY, SHOWCHECKBOXESTOSELECT_VALUENAME)
        End Get
        Set(value As Boolean)
            Settings.SetRegistryBoolean(SHOWCHECKBOXESTOSELECT_KEY, SHOWCHECKBOXESTOSELECT_VALUENAME, value)
            Me.NotifyOfPropertyChange("DoShowCheckBoxesToSelect")
        End Set
    End Property

    Private Shared Function GetRegistryBoolean(key As String, valueName As String) As Boolean
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                Dim hideFileExt As Object = registryKey.GetValue(valueName)
                If Not hideFileExt Is Nothing AndAlso TypeOf hideFileExt Is Integer Then
                    Return CType(hideFileExt, Integer) <> 0
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Using
    End Function

    Private Shared Sub SetRegistryBoolean(key As String, valueName As String, value As Boolean)
        Using registryKey As RegistryKey = Registry.CurrentUser.CreateSubKey(key, True)
            If Not key Is Nothing Then
                registryKey.SetValue(valueName, If(value, 1, 0))
            End If
        End Using
    End Sub
End Class

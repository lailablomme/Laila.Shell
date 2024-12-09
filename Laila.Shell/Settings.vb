Imports System.Runtime.InteropServices
Imports System.Windows.Documents
Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Public Property DoHideKnownFileExtensions As Boolean
        Get
            Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
            Dim val As SHELLSTATE
            Functions.SHGetSetSettings(val, mask, False)
            Return Not val.Data1 = SSF.SSF_SHOWEXTENSIONS
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
            Dim val As SHELLSTATE
            val.Data1 = If(value, 0, SSF.SSF_SHOWEXTENSIONS)
            Functions.SHGetSetSettings(val, mask, True)
            onDoHideKnownFileExtensionsChanged()
        End Set
    End Property

    Private Sub onDoHideKnownFileExtensionsChanged()
        Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")

        Dim list As List(Of Item)
        SyncLock Shell._itemsCacheLock
            list = Shell.ItemsCache.Select(Function(i) i.Item1).ToList()
        End SyncLock
        For Each item In list
            item._displayName = Nothing
            item.ShellItem2.Update(IntPtr.Zero)
            item.NotifyOfPropertyChange("DisplayName")
        Next
    End Sub

    Public Property DoShowCheckBoxesToSelect As Boolean
        Get
            Dim mask As SSF = SSF.SSF_AUTOCHECKSELECT
            Dim val As SHELLSTATE
            Functions.SHGetSetSettings(val, mask, False)
            Return val.Data10 = 8
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_AUTOCHECKSELECT
            Dim val As SHELLSTATE
            val.Data10 = If(value, 8, 0)
            Functions.SHGetSetSettings(val, mask, True)
            Me.NotifyOfPropertyChange("DoShowCheckBoxesToSelect")
        End Set
    End Property

    Friend Sub OnSettingChange()
        onDoHideKnownFileExtensionsChanged()
        Me.NotifyOfPropertyChange("DoShowCheckBoxesToSelect")
    End Sub

    Private Shared Function GetRegistryBoolean(key As String, valueName As String) As Boolean
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                Dim val As Object = registryKey.GetValue(valueName)
                If Not val Is Nothing AndAlso TypeOf val Is Integer Then
                    Return CType(val, Integer) <> 0
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

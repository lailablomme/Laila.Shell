Imports System.Windows.Input
Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private _doHideKnownFileExtensions As Boolean
    Private _doShowProtectedOperatingSystemFiles As Boolean
    Private _doShowCheckBoxesToSelect As Boolean
    Private _doShowHiddenFilesAndFolders As Boolean

    Public Sub New()
        Me.OnSettingChange(False)
    End Sub

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
        End Set
    End Property

    Public Property DoShowProtectedOperatingSystemFiles As Boolean
        Get
            Dim mask As SSF = SSF.SSF_SHOWSUPERHIDDEN
            Dim val As SHELLSTATE
            Functions.SHGetSetSettings(val, mask, False)
            Return val.Data2 = 128
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWSUPERHIDDEN
            Dim val As SHELLSTATE
            val.Data2 = If(value, 128, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Public Property DoShowHiddenFilesAndFolders As Boolean
        Get
            Dim mask As SSF = SSF.SSF_SHOWALLOBJECTS
            Dim val As SHELLSTATE
            Functions.SHGetSetSettings(val, mask, False)
            Return val.Data1 = 1
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWALLOBJECTS
            Dim val As SHELLSTATE
            val.Data1 = If(value, 1, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

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
        End Set
    End Property

    Friend Sub OnSettingChange(Optional doNotify As Boolean = True)
        Using Shell.OverrideCursor(Cursors.Wait)
            Dim b As Boolean
            b = Me.DoHideKnownFileExtensions
            If Not b = _doHideKnownFileExtensions Then
                _doHideKnownFileExtensions = b
                If doNotify Then Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")
            End If
            b = Me.DoShowCheckBoxesToSelect
            If Not b = _doShowCheckBoxesToSelect Then
                _doShowCheckBoxesToSelect = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowCheckBoxesToSelect")
            End If
            b = Me.DoShowProtectedOperatingSystemFiles
            If Not b = _doShowProtectedOperatingSystemFiles Then
                _doShowProtectedOperatingSystemFiles = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowProtectedOperatingSystemFiles")
            End If
            b = Me.DoShowHiddenFilesAndFolders
            If Not b = _doShowHiddenFilesAndFolders Then
                _doShowHiddenFilesAndFolders = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowHiddenFilesAndFolders")
            End If
        End Using
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

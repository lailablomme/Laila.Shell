Imports System.Drawing.Text
Imports System.Threading
Imports System.Windows.Input
Imports Laila.Shell.Helpers
Imports Microsoft.Win32

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private Const EXPLORER_KEYPATH As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    Private Const EXPLORER_ADVANCED_KEYPATH As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME As String = "ShowEncryptCompressedColor"
    Private Const UNDERLINEITEMONHOVER_VALUENAME As String = "IconUnderline"

    Private _doHideKnownFileExtensions As Boolean
    Private _doShowProtectedOperatingSystemFiles As Boolean
    Private _doShowCheckBoxesToSelect As Boolean
    Private _doShowHiddenFilesAndFolders As Boolean
    Private _doShowEncryptedOrCompressedFilesInColor As Boolean
    Private _isDoubleClickToOpenItem As Boolean
    Private _isUnderlineItemOnHover As Boolean

    Public Sub New()
        Me.OnSettingChange(False)

        ' some settings don't produce notifications, so we need to poll
        Dim doNotify As Boolean
        Dim t As Thread = New Thread(
            Sub()
                While Not Shell.ShuttingDownToken.IsCancellationRequested
                    Dim b As Boolean = Me.IsUnderlineItemOnHover
                    If Not b = _isUnderlineItemOnHover Then
                        _isUnderlineItemOnHover = b
                        If doNotify Then
                            UIHelper.OnUIThread(
                                Sub()
                                    Me.NotifyOfPropertyChange("IsUnderlineItemOnHover")
                                End Sub)
                        End If
                    End If

                    doNotify = True
                    Thread.Sleep(1000)
                End While
            End Sub)
        t.SetApartmentState(ApartmentState.STA)
        t.IsBackground = True
        t.Start()
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

    Public Property DoShowEncryptedOrCompressedFilesInColor As Boolean
        Get
            Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME, True)
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME, value)
            Dim mask As SSF = SSF.SSF_SHOWCOMPCOLOR
            Dim val As SHELLSTATE
            val.Data1 = If(value, 16, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Public Property IsDoubleClickToOpenItem As Boolean
        Get
            Dim mask As SSF = SSF.SSF_DOUBLECLICKINWEBVIEW
            Dim val As SHELLSTATE
            Functions.SHGetSetSettings(val, mask, False)
            Return val.Data1 = 32
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_DOUBLECLICKINWEBVIEW
            Dim val As SHELLSTATE
            val.Data1 = If(value, 32, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Public Property IsUnderlineItemOnHover As Boolean
        Get
            Return GetRegistryDWord(EXPLORER_KEYPATH, UNDERLINEITEMONHOVER_VALUENAME, 2) = 2
        End Get
        Set(value As Boolean)
            SetRegistryDWord(EXPLORER_KEYPATH, UNDERLINEITEMONHOVER_VALUENAME, If(value, 2, 3))
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
            b = Me.DoShowEncryptedOrCompressedFilesInColor
            If Not b = _doShowEncryptedOrCompressedFilesInColor Then
                _doShowEncryptedOrCompressedFilesInColor = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowEncryptedOrCompressedFilesInColor")
            End If
            b = Me.IsDoubleClickToOpenItem
            If Not b = _isDoubleClickToOpenItem Then
                _isDoubleClickToOpenItem = b
                If doNotify Then Me.NotifyOfPropertyChange("IsDoubleClickToOpenItem")
            End If
        End Using
    End Sub

    Private Shared Function GetRegistryBoolean(key As String, valueName As String, defaultValue As Boolean) As Boolean
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                Dim val As Object = registryKey.GetValue(valueName)
                If Not val Is Nothing AndAlso TypeOf val Is Integer Then
                    Return CType(val, Integer) <> 0
                Else
                    Return defaultValue
                End If
            Else
                Return defaultValue
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

    Private Shared Function GetRegistryDWord(key As String, valueName As String, defaultValue As Integer) As Integer
        Using registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key)
            If Not key Is Nothing Then
                Dim val As Object = registryKey.GetValue(valueName)
                If Not val Is Nothing AndAlso TypeOf val Is Integer Then
                    Return CType(val, Integer)
                Else
                    Return defaultValue
                End If
            Else
                Return defaultValue
            End If
        End Using
    End Function

    Private Shared Sub SetRegistryDWord(key As String, valueName As String, value As Integer)
        Using registryKey As RegistryKey = Registry.CurrentUser.CreateSubKey(key, True)
            If Not key Is Nothing Then
                registryKey.SetValue(valueName, value)
            End If
        End Using
    End Sub
End Class

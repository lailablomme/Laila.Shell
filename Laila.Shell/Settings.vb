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
    Private Const SHOWFOLDERCONTENTSININFOTIP_VALUENAME As String = "FolderContentsInfoTip"
    Private Const COMPACTMODE_VALUENAME As String = "UseCompactMode"
    Private Const SHOWDRIVELETTERS_VALUENAME As String = "ShowDriveLettersFirst"

    Private _threads As List(Of Thread) = New List(Of Thread)()
    Private _doHideKnownFileExtensions As Boolean
    Private _doShowProtectedOperatingSystemFiles As Boolean
    Private _doShowCheckBoxesToSelect As Boolean
    Private _doShowHiddenFilesAndFolders As Boolean
    Private _doShowEncryptedOrCompressedFilesInColor As Boolean
    Private _isDoubleClickToOpenItem As Boolean
    Private _isUnderlineItemOnHover As Boolean
    Private _doShowIconsOnly As Boolean
    Private _doShowTypeOverlay As Boolean
    Private _doShowFolderContentsInInfoTip As Boolean
    Private _doShowInfoTips As Boolean
    Private _isCompactMode As Boolean
    Private _doShowDriveLetters As Boolean
    Private _doShowStatusBar As Boolean

    Public Sub New()
        Me.OnSettingChange(False)

        ' some settings don't produce notifications, so we need to monitor the registry
        monitorRegistryKey(EXPLORER_KEYPATH,
            Sub()
                Dim b As Boolean
                b = readIsUnderlineItemOnHover()
                If Not b = _isUnderlineItemOnHover Then
                    _isUnderlineItemOnHover = b
                    Me.NotifyOfPropertyChange("IsUnderlineItemOnHover")
                End If
                b = readDoShowDriveLetters()
                If Not b = _doShowDriveLetters Then
                    _doShowDriveLetters = b
                    Me.NotifyOfPropertyChange("DoShowDriveLetters")
                End If
            End Sub)
        monitorRegistryKey(EXPLORER_ADVANCED_KEYPATH,
            Sub()
                Dim b As Boolean
                b = readDoShowFolderContentsInInfoTip()
                If Not b = _doShowFolderContentsInInfoTip Then
                    _doShowFolderContentsInInfoTip = b
                    Me.NotifyOfPropertyChange("DoShowFolderContentsInInfoTip")
                End If
                b = readIsCompactMode()
                If Not b = _isCompactMode Then
                    _isCompactMode = b
                    Me.NotifyOfPropertyChange("IsCompactMode")
                End If
            End Sub)
    End Sub

    Private Sub monitorRegistryKey(keyPath As String, onChange As Action)
        Dim t As Thread = New Thread(
            Sub()
                Using registryKey = Registry.CurrentUser.OpenSubKey(keyPath, writable:=False)
                    Dim hKey As IntPtr = registryKey.Handle.DangerousGetHandle()
                    While Not Shell.ShuttingDownToken.IsCancellationRequested
                        Dim result As HRESULT = Functions.RegNotifyChangeKeyValue(
                            hKey,
                            False,
                            REG_NOTIFY_CHANGE.LAST_SET,
                            IntPtr.Zero,
                            False
                        )

                        If result = HRESULT.S_OK Then
                            UIHelper.OnUIThread(
                                Sub()
                                    onChange()
                                End Sub)
                        End If
                    End While
                End Using
            End Sub)
        t.SetApartmentState(ApartmentState.STA)
        t.IsBackground = True
        t.Start()
        _threads.Add(t)
    End Sub

    Private Function readDoHideKnownFileExtensions() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return Not val.Data1 = SSF.SSF_SHOWEXTENSIONS
    End Function

    Public Property DoHideKnownFileExtensions As Boolean
        Get
            Return _doHideKnownFileExtensions
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWEXTENSIONS
            Dim val As SHELLSTATE
            val.Data1 = If(value, 0, SSF.SSF_SHOWEXTENSIONS)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowProtectedOperatingSystemFiles() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWSUPERHIDDEN
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data2 = 128
    End Function

    Public Property DoShowProtectedOperatingSystemFiles As Boolean
        Get
            Return _doShowProtectedOperatingSystemFiles
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWSUPERHIDDEN
            Dim val As SHELLSTATE
            val.Data2 = If(value, 128, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowHiddenFilesAndFolders() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWALLOBJECTS
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data1 = 1
    End Function

    Public Property DoShowHiddenFilesAndFolders As Boolean
        Get
            Return _doShowHiddenFilesAndFolders
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWALLOBJECTS
            Dim val As SHELLSTATE
            val.Data1 = If(value, 1, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowCheckBoxesToSelect() As Boolean
        Dim mask As SSF = SSF.SSF_AUTOCHECKSELECT
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data10 = 8
    End Function

    Public Property DoShowCheckBoxesToSelect As Boolean
        Get
            Return _doShowCheckBoxesToSelect
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_AUTOCHECKSELECT
            Dim val As SHELLSTATE
            val.Data10 = If(value, 8, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowEncryptedOrCompressedFilesInColor() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME, True)
    End Function

    Public Property DoShowEncryptedOrCompressedFilesInColor As Boolean
        Get
            Return _doShowEncryptedOrCompressedFilesInColor
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME, value)
            Dim mask As SSF = SSF.SSF_SHOWCOMPCOLOR
            Dim val As SHELLSTATE
            val.Data1 = If(value, 16, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readIsDoubleClickToOpenItem() As Boolean
        Dim mask As SSF = SSF.SSF_DOUBLECLICKINWEBVIEW
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data1 = 32
    End Function

    Public Property IsDoubleClickToOpenItem As Boolean
        Get
            Return _isDoubleClickToOpenItem
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_DOUBLECLICKINWEBVIEW
            Dim val As SHELLSTATE
            val.Data1 = If(value, 32, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readIsUnderlineItemOnHover() As Boolean
        Return GetRegistryDWord(EXPLORER_KEYPATH, UNDERLINEITEMONHOVER_VALUENAME, 2) = 2
    End Function

    Public Property IsUnderlineItemOnHover As Boolean
        Get
            Return _isUnderlineItemOnHover
        End Get
        Set(value As Boolean)
            SetRegistryDWord(EXPLORER_KEYPATH, UNDERLINEITEMONHOVER_VALUENAME, If(value, 2, 3))
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowIconsOnly() As Boolean
        Dim mask As SSF = SSF.SSF_ICONSONLY
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data10 = 16
    End Function

    Public Property DoShowIconsOnly As Boolean
        Get
            Return _doShowIconsOnly
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_ICONSONLY
            Dim val As SHELLSTATE
            val.Data10 = If(value, 16, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowTypeOverlay() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWTYPEOVERLAY
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data10 = 32
    End Function

    Public Property DoShowTypeOverlay As Boolean
        Get
            Return _doShowTypeOverlay
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWTYPEOVERLAY
            Dim val As SHELLSTATE
            val.Data10 = If(value, 32, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readDoShowFolderContentsInInfoTip() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWFOLDERCONTENTSININFOTIP_VALUENAME, True)
    End Function

    Public Property DoShowFolderContentsInInfoTip As Boolean
        Get
            Return _doShowFolderContentsInInfoTip
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, SHOWFOLDERCONTENTSININFOTIP_VALUENAME, value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowInfoTips() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWINFOTIP
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data2 = 8
    End Function

    Public Property DoShowInfoTips As Boolean
        Get
            Return _doShowInfoTips
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWINFOTIP
            Dim val As SHELLSTATE
            val.Data2 = If(value, 8, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Private Function readIsCompactMode() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, COMPACTMODE_VALUENAME, True)
    End Function

    Public Property IsCompactMode As Boolean
        Get
            Return _isCompactMode
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, COMPACTMODE_VALUENAME, value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowDriveLetters() As Boolean
        Return GetRegistryDWord(EXPLORER_KEYPATH, SHOWDRIVELETTERS_VALUENAME, 0) = 0
    End Function

    Public Property DoShowDriveLetters As Boolean
        Get
            Return _doShowDriveLetters
        End Get
        Set(value As Boolean)
            SetRegistryDWord(EXPLORER_KEYPATH, SHOWDRIVELETTERS_VALUENAME, If(value, 0, 2))
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowStatusBar() As Boolean
        Dim mask As SSF = SSF.SSF_SHOWSTATUSBAR
        Dim val As SHELLSTATE
        Functions.SHGetSetSettings(val, mask, False)
        Return val.Data10 = 64
    End Function

    Public Property DoShowStatusBar As Boolean
        Get
            Return _doShowStatusBar
        End Get
        Set(value As Boolean)
            Dim mask As SSF = SSF.SSF_SHOWSTATUSBAR
            Dim val As SHELLSTATE
            val.Data10 = If(value, 64, 0)
            Functions.SHGetSetSettings(val, mask, True)
        End Set
    End Property

    Public Sub Touch()
        ' cause Windows Explorer to pick up the changes after we've modified the registry directly
        Dim b As Boolean = Me.IsDoubleClickToOpenItem
        Me.IsDoubleClickToOpenItem = Not b
        Me.IsDoubleClickToOpenItem = b
        'Dim result As IntPtr
        'Functions.SendMessageTimeout(CType(&HFFFF, IntPtr), WM.SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero, 0, 1000, result)
    End Sub

    Friend Sub OnSettingChange(Optional doNotify As Boolean = True)
        Using Shell.OverrideCursor(Cursors.Wait)
            Dim b As Boolean
            b = readDoHideKnownFileExtensions()
            If Not b = _doHideKnownFileExtensions Then
                _doHideKnownFileExtensions = b
                If doNotify Then Me.NotifyOfPropertyChange("DoHideKnownFileExtensions")
            End If
            b = readDoShowCheckBoxesToSelect()
            If Not b = _doShowCheckBoxesToSelect Then
                _doShowCheckBoxesToSelect = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowCheckBoxesToSelect")
            End If
            b = readDoShowProtectedOperatingSystemFiles()
            If Not b = _doShowProtectedOperatingSystemFiles Then
                _doShowProtectedOperatingSystemFiles = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowProtectedOperatingSystemFiles")
            End If
            b = readDoShowHiddenFilesAndFolders()
            If Not b = _doShowHiddenFilesAndFolders Then
                _doShowHiddenFilesAndFolders = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowHiddenFilesAndFolders")
            End If
            b = readDoShowEncryptedOrCompressedFilesInColor()
            If Not b = _doShowEncryptedOrCompressedFilesInColor Then
                _doShowEncryptedOrCompressedFilesInColor = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowEncryptedOrCompressedFilesInColor")
            End If
            b = readIsDoubleClickToOpenItem()
            If Not b = _isDoubleClickToOpenItem Then
                _isDoubleClickToOpenItem = b
                If doNotify Then Me.NotifyOfPropertyChange("IsDoubleClickToOpenItem")
            End If
            b = readDoShowIconsOnly()
            If Not b = _doShowIconsOnly Then
                _doShowIconsOnly = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowIconsOnly")
            End If
            b = readDoShowTypeOverlay()
            If Not b = _doShowTypeOverlay Then
                _doShowTypeOverlay = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowTypeOverlay")
            End If
            b = readDoShowInfoTips()
            If Not b = _doShowInfoTips Then
                _doShowInfoTips = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowInfoTips")
            End If
            b = readDoShowStatusBar()
            If Not b = _doShowStatusBar Then
                _doShowStatusBar = b
                If doNotify Then Me.NotifyOfPropertyChange("DoShowStatusBar")
            End If

            _isUnderlineItemOnHover = readIsUnderlineItemOnHover()
            _doShowFolderContentsInInfoTip = readDoShowFolderContentsInInfoTip()
            _isCompactMode = readIsCompactMode()
            _doShowDriveLetters = readDoShowDriveLetters()
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

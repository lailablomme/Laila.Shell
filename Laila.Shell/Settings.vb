Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Public Class Settings
    Inherits NotifyPropertyChangedBase

    Private Const EXPLORER_KEYPATH As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    Private Const EXPLORER_ADVANCED_KEYPATH As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Private Const LIBRARIES_KEYPATH As String = "Software\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}"
    Private Const SHOWENCRYPTEDORCOMPRESSEDFILESINCOLOR_VALUENAME As String = "ShowEncryptCompressedColor"
    Private Const UNDERLINEITEMONHOVER_VALUENAME As String = "IconUnderline"
    Private Const SHOWFOLDERCONTENTSININFOTIP_VALUENAME As String = "FolderContentsInfoTip"
    Private Const COMPACTMODE_VALUENAME As String = "UseCompactMode"
    Private Const SHOWDRIVELETTERS_VALUENAME As String = "ShowDriveLettersFirst"
    Private Const TYPETOSELECT_VALUENAME As String = "TypeAhead"
    Private Const NAVPANESHOWALLFOLDERS_VALUENAME As String = "NavPaneShowAllFolders"
    Private Const NAVPANESHOWALLCLOUDSTATES_VALUENAME As String = "NavPaneShowAllCloudStates"
    Private Const NAVPANEEXPANDTOCURRENTFOLDER_VALUENAME As String = "NavPaneExpandToCurrentFolder"
    Private Const SHOWLIBRARIES_VALUENAME As String = "System.IsPinnedToNameSpaceTree"

    Public Shared Property DpiScaleX As Double = 1
    Public Shared Property DpiScaleY As Double = 1

    Private _threads As List(Of Thread) = New List(Of Thread)()
    Private _isMonitoring As Boolean
    Private _stopped1 As TaskCompletionSource
    Private _stopped2 As TaskCompletionSource
    Private _stopped3 As TaskCompletionSource
    Private _cancel As CancellationTokenSource
    Private _doHideKnownFileExtensions As Boolean = False
    Private _doShowProtectedOperatingSystemFiles As Boolean = True
    Private _doShowCheckBoxesToSelect As Boolean = True
    Private _doShowHiddenFilesAndFolders As Boolean = True
    Private _doShowEncryptedOrCompressedFilesInColor As Boolean = True
    Private _isDoubleClickToOpenItem As Boolean = True
    Private _isUnderlineItemOnHover As Boolean = True
    Private _doShowIconsOnly As Boolean = False
    Private _doShowTypeOverlay As Boolean = True
    Private _doShowFolderContentsInInfoTip As Boolean = True
    Private _doShowInfoTips As Boolean = True
    Private _isCompactMode As Boolean = True
    Private _doShowDriveLetters As Boolean = True
    Private _doShowStatusBar As Boolean = True
    Private _doTypeToSelect As Boolean = True
    Private _doShowAllFoldersInTreeView As Boolean = False
    Private _doShowAvailabilityStatusInTreeView As Boolean = True
    Private _doExpandTreeViewToCurrentFolder As Boolean = True
    Private _doShowLibrariesInTreeView As Boolean = False

    Public Sub New()
        Me.StartMonitoring()

        If Not Settings.IsWindows7OrLower Then
            Dim maxDpiX As UInteger = 96
            Dim maxDpiY As UInteger = 96
            For Each s In System.Windows.Forms.Screen.AllScreens
                Dim hMonitor As IntPtr = Functions.MonitorFromPoint(New WIN32POINT() With {.x = s.Bounds.X, .y = s.Bounds.Y}, 2)
                Dim dpiX As UInteger = 0
                Dim dpiY As UInteger = 0

                Dim result = Functions.GetDpiForMonitor(hMonitor, MonitorDpiType.MDT_EFFECTIVE_DPI, dpiX, dpiY)
                If result = HRESULT.S_OK Then ' S_OK
                    maxDpiX = Math.Max(maxDpiX, dpiX)
                    maxDpiY = Math.Max(maxDpiY, dpiY)
                End If
            Next
            Settings.DpiScaleX = maxDpiX / 96
            Settings.DpiScaleY = maxDpiY / 96
        End If
    End Sub

    ''' <summary>
    ''' Some settings don't produce notifications, so we need to monitor the registry
    ''' </summary>
    Public Sub StartMonitoring()
        If Not _isMonitoring Then
            ' read current settings
            Me.OnSettingChange(False)

            ' prepare tokens
            _stopped1 = New TaskCompletionSource()
            _stopped2 = New TaskCompletionSource()
            _stopped3 = New TaskCompletionSource()
            _cancel = New CancellationTokenSource()

            ' start monitoring
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
                End Sub, _cancel.Token, _stopped1)
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
                    b = readDoTypeToSelect()
                    If Not b = _doTypeToSelect Then
                        _doTypeToSelect = b
                        Me.NotifyOfPropertyChange("DoTypeToSelect")
                    End If
                    b = readDoShowAllFoldersInTreeView()
                    If Not b = _doShowAllFoldersInTreeView Then
                        _doShowAllFoldersInTreeView = b
                        Me.NotifyOfPropertyChange("DoShowAllFoldersInTreeView")
                    End If
                    b = readDoShowAvailabilityStatusInTreeView()
                    If Not b = _doShowAvailabilityStatusInTreeView Then
                        _doShowAvailabilityStatusInTreeView = b
                        Me.NotifyOfPropertyChange("DoShowAvailabilityStatusInTreeView")
                    End If
                    b = readDoExpandTreeViewToCurrentFolder()
                    If Not b = _doExpandTreeViewToCurrentFolder Then
                        _doExpandTreeViewToCurrentFolder = b
                        Me.NotifyOfPropertyChange("DoExpandTreeViewToCurrentFolder")
                    End If
                End Sub, _cancel.Token, _stopped2)
            monitorRegistryKey(LIBRARIES_KEYPATH,
                Sub()
                    Dim b As Boolean
                    b = readDoShowLibrariesInTreeView()
                    If Not b = _doShowLibrariesInTreeView Then
                        _doShowLibrariesInTreeView = b
                        Me.NotifyOfPropertyChange("DoShowLibrariesInTreeView")
                    End If
                End Sub, _cancel.Token, _stopped3)

            ' mark
            _isMonitoring = True
        Else
            Throw New Exception("We're already monitoring.")
        End If
    End Sub

    Public Sub StopMonitoring()
        If _isMonitoring Then
            ' make loops end
            _cancel.Cancel()

            ' make us drop out of RegNotifyChangeKeyValue
            Dim b As Boolean
            b = GetRegistryBoolean(EXPLORER_KEYPATH, "Laila_Shell_Monitor", False)
            SetRegistryBoolean(EXPLORER_KEYPATH, "Laila_Shell_Monitor", Not b)
            b = GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, "Laila_Shell_Monitor", False)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, "Laila_Shell_Monitor", Not b)
            b = GetRegistryBoolean(LIBRARIES_KEYPATH, "Laila_Shell_Monitor", False)
            SetRegistryBoolean(LIBRARIES_KEYPATH, "Laila_Shell_Monitor", Not b)

            ' wait for threads to end
            _stopped1.Task.Wait()
            _stopped2.Task.Wait()
            _stopped3.Task.Wait()

            ' mark
            _isMonitoring = False
        Else
            Throw New Exception("We're not monitoring now.")
        End If
    End Sub

    Private Sub monitorRegistryKey(keyPath As String, onChange As Action,
                                   cancellationToken As CancellationToken, stopped As TaskCompletionSource)
        Dim tcs As TaskCompletionSource = New TaskCompletionSource()
        Dim t As Thread = New Thread(
            Sub()
                Dim handle As IntPtr
                Try
                    Dim h As HRESULT = Functions.RegCreateKeyEx(HKEY.CURRENT_USER, keyPath, 0, Nothing, 0, REGKEY.NOTIFY Or REGKEY.WRITE Or REGKEY.WOW64_64KEY, IntPtr.Zero, handle, IntPtr.Zero)
                    If h = HRESULT.S_OK Then
                        tcs.SetResult()
                        While Not cancellationToken.IsCancellationRequested
                            Dim result As HRESULT = Functions.RegNotifyChangeKeyValue(
                                handle,
                                False,
                                REG_NOTIFY_CHANGE.LAST_SET,
                                IntPtr.Zero,
                                False
                            )

                            If result = HRESULT.S_OK AndAlso Not cancellationToken.IsCancellationRequested Then
                                UIHelper.OnUIThread(
                                    Sub()
                                        onChange()
                                    End Sub)
                            End If
                        End While
                    Else
                        tcs.SetException(If(Marshal.GetExceptionForHR(h), New Exception("HRESULT " & h.ToString())))
                        stopped.SetResult()
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(handle) Then
                        Functions.RegCloseKey(handle)
                    End If
                    stopped.SetResult()
                End Try
            End Sub)
        t.SetApartmentState(ApartmentState.STA)
        t.IsBackground = True
        t.Start()
        _threads.Add(t)
        tcs.Task.Wait()
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

    Private Function readDoTypeToSelect() As Boolean
        Return Not GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, TYPETOSELECT_VALUENAME, False)
    End Function

    Public Property DoTypeToSelect As Boolean
        Get
            Return _doTypeToSelect
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, TYPETOSELECT_VALUENAME, Not value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowAllFoldersInTreeView() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANESHOWALLFOLDERS_VALUENAME, False)
    End Function

    Public Property DoShowAllFoldersInTreeView As Boolean
        Get
            Return _doShowAllFoldersInTreeView
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANESHOWALLFOLDERS_VALUENAME, value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowAvailabilityStatusInTreeView() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANESHOWALLCLOUDSTATES_VALUENAME, True)
    End Function

    Public Property DoShowAvailabilityStatusInTreeView As Boolean
        Get
            Return _doShowAvailabilityStatusInTreeView
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANESHOWALLCLOUDSTATES_VALUENAME, value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoExpandTreeViewToCurrentFolder() As Boolean
        Return GetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANEEXPANDTOCURRENTFOLDER_VALUENAME, True)
    End Function

    Public Property DoExpandTreeViewToCurrentFolder As Boolean
        Get
            Return _doExpandTreeViewToCurrentFolder
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(EXPLORER_ADVANCED_KEYPATH, NAVPANEEXPANDTOCURRENTFOLDER_VALUENAME, value)
            Me.Touch()
        End Set
    End Property

    Private Function readDoShowLibrariesInTreeView() As Boolean
        Return GetRegistryBoolean(LIBRARIES_KEYPATH, SHOWLIBRARIES_VALUENAME, False)
    End Function

    Public Property DoShowLibrariesInTreeView As Boolean
        Get
            Return _doShowLibrariesInTreeView
        End Get
        Set(value As Boolean)
            SetRegistryBoolean(LIBRARIES_KEYPATH, SHOWLIBRARIES_VALUENAME, value)
            Me.Touch()
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

            If Not doNotify Then ' only read this values here on program start
                _isUnderlineItemOnHover = readIsUnderlineItemOnHover()
                _doShowFolderContentsInInfoTip = readDoShowFolderContentsInInfoTip()
                _isCompactMode = readIsCompactMode()
                _doShowDriveLetters = readDoShowDriveLetters()
                _doTypeToSelect = readDoTypeToSelect()
                _doShowAllFoldersInTreeView = readDoShowAllFoldersInTreeView()
                _doShowAvailabilityStatusInTreeView = readDoShowAvailabilityStatusInTreeView()
                _doExpandTreeViewToCurrentFolder = readDoExpandTreeViewToCurrentFolder()
                _doShowLibrariesInTreeView = readDoShowLibrariesInTreeView()
            End If
        End Using
    End Sub

    Private Shared Function GetRegistryBoolean(keyPath As String, valueName As String, defaultValue As Boolean) As Boolean
        Return GetRegistryDWord(keyPath, valueName, If(defaultValue, 1, 0)) <> 0
    End Function

    Private Shared Sub SetRegistryBoolean(keyPath As String, valueName As String, value As Boolean)
        SetRegistryDWord(keyPath, valueName, If(value, 1, 0))
    End Sub

    Private Shared Function GetRegistryDWord(keyPath As String, valueName As String, defaultValue As Integer) As Int32?
        Dim handle As IntPtr
        Try
            Dim h As HRESULT = Functions.RegOpenKeyEx(HKEY.CURRENT_USER, keyPath, 0, REGKEY.QUERY_VALUE Or REGKEY.WOW64_64KEY, handle)
            If h = HRESULT.S_OK Then
                Dim data As Int32

                h = Functions.RegQueryValueEx(handle, valueName, 0, 0, data, Marshal.SizeOf(Of Int32))
                If h = HRESULT.S_OK Then
                    Return data
                Else
                    Return defaultValue
                End If
            Else
                Return defaultValue
            End If
        Finally
            If Not IntPtr.Zero.Equals(handle) Then
                Functions.RegCloseKey(handle)
            End If
        End Try
        Return Nothing
    End Function

    Private Shared Sub SetRegistryDWord(keyPath As String, valueName As String, value As Integer)
        Dim handle As IntPtr
        Try
            Dim h As HRESULT = Functions.RegCreateKeyEx(HKEY.CURRENT_USER, keyPath, 0, Nothing, 0, REGKEY.WRITE Or REGKEY.WOW64_64KEY, IntPtr.Zero, handle, IntPtr.Zero)
            If h = HRESULT.S_OK Then
                Dim dataBytes As Byte() = BitConverter.GetBytes(value)
                Dim dataSize As Integer = dataBytes.Length

                h = Functions.RegSetValueEx(handle, valueName, 0, REG.DWORD, dataBytes, dataSize)
                If Not h = HRESULT.S_OK Then
                    Throw Marshal.GetExceptionForHR(h)
                End If
            Else
                Throw Marshal.GetExceptionForHR(h)
            End If
        Finally
            If Not IntPtr.Zero.Equals(handle) Then
                Functions.RegCloseKey(handle)
            End If
        End Try
    End Sub


    Public Shared ReadOnly Property IsWindows7OrLower() As Boolean
        Get
            ' Windows 7 has version number 6.1
            Dim osVersion As Version = Environment.OSVersion.Version
            Return osVersion.Major < 6 OrElse (osVersion.Major = 6 AndAlso osVersion.Minor <= 1)
        End Get
    End Property
End Class

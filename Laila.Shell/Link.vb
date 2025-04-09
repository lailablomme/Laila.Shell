Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Items

Public Class Link
    Inherits Item

    Private _threadId As Integer
    Private _shellLinkW As IShellLinkW
    Private _targetPidl As Pidl
    Private _targetFullPath As String
    Private _targetItem As Item = Nothing

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, threadId As Integer?, Optional pidl As Pidl = Nothing)
        MyBase.New(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, pidl)

        _threadId = Shell.GlobalThreadPool.GetNextFreeThreadId()
    End Sub

    Protected ReadOnly Property ShellLink As IShellLinkW
        Get
            If _shellLinkW Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                SyncLock _shellItemLock
                    If _shellLinkW Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                        CType(ShellItem2, IShellItem2ForShellLink).BindToHandler(Nothing, Guids.BHID_SFUIObject, GetType(IShellLinkW).GUID, _shellLinkW)
                    End If
                End SyncLock
            End If

            Return _shellLinkW
        End Get
    End Property

    Public Sub Resolve(flags As SLR_FLAGS)
        Shell.GlobalThreadPool.Run(
            Sub()
                Try
                    Me.ShellLink?.Resolve(IntPtr.Zero, flags)
                Catch ex As Exception
                End Try
            End Sub,, _threadId)
    End Sub

    Public ReadOnly Property TargetPidl As Pidl
        Get
            If _targetPidl Is Nothing Then
                _targetPidl = Shell.GlobalThreadPool.Run(
                    Function() As Pidl
                        Dim pidl As IntPtr
                        Me.ShellLink?.GetIDList(pidl)
                        If Not IntPtr.Zero.Equals(pidl) Then
                            Return New Pidl(pidl)
                        End If
                        Return Nothing
                    End Function,, _threadId)
            End If
            Return _targetPidl
        End Get
    End Property

    'Public ReadOnly Property TargetFullPath As String
    '    Get
    '        If String.IsNullOrWhiteSpace(_targetFullPath) Then
    '            Try
    '                Dim fullPath As StringBuilder = New StringBuilder(New String(Chr(0), 2048 + 2))
    '                _shellLinkW.GetPath(fullPath, 2048, Nothing, SLGP_FLAGS.RAWPATH)
    '                _targetFullPath = Environment.ExpandEnvironmentVariables(fullPath.ToString())
    '            Catch ex As Exception
    '            End Try
    '        End If

    '        Return _targetFullPath
    '    End Get
    'End Property

    Public ReadOnly Property TargetItem As Item
        Get
            If Not disposedValue Then
                If Not _targetItem Is Nothing Then
                    SyncLock _targetItem._shellItemLock
                        ' if still alive...
                        If Not _targetItem.disposedValue Then
                            ' extend lifetime
                            Shell.RemoveFromItemsCache(_targetItem)
                            Shell.AddToItemsCache(_targetItem)
                        Else
                            _targetItem = Nothing
                        End If
                    End SyncLock
                End If

                If _targetItem Is Nothing Then
                    _targetItem = Item.FromPidl(Me.TargetPidl, Nothing, _doKeepAlive, True)
                End If
            End If

            Return _targetItem
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        SyncLock _shellItemLock
            If Not disposedValue Then
                MyBase.Dispose(disposing)
            End If
        End SyncLock

        Shell.DisposerThreadPool.Add(
            Sub()
                If Not _shellLinkW Is Nothing Then
                    Marshal.ReleaseComObject(_shellLinkW)
                    _shellLinkW = Nothing
                End If

                If Not _targetPidl Is Nothing Then
                    _targetPidl.Dispose()
                    _targetPidl = Nothing
                End If
            End Sub)
    End Sub
End Class

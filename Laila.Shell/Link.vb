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
        CType(shellItem2, IShellItem2ForShellLink).BindToHandler(Nothing, Guids.BHID_SFUIObject, GetType(IShellLinkW).GUID, _shellLinkW)
    End Sub

    Protected ReadOnly Property ShellLink As IShellLinkW
        Get
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
                Dim pidl As IntPtr
                Me.ShellLink?.GetIDList(pidl)
                If Not IntPtr.Zero.Equals(pidl) Then
                    _targetPidl = New Pidl(pidl)
                End If
            End If
            Return _targetPidl
        End Get
    End Property

    Public ReadOnly Property TargetFullPath As String
        Get
            If String.IsNullOrWhiteSpace(_targetFullPath) Then
                Try
                    _targetFullPath = Me.TargetItem?.FullPath
                Catch ex As Exception
                End Try
            End If

            Return _targetFullPath
        End Get
    End Property

    Public ReadOnly Property TargetItem As Item
        Get
            If Not disposedValue Then
                If _targetItem Is Nothing Then
                    _targetItem = Item.FromPidl(Me.TargetPidl, Nothing, True, True)
                End If
            End If

            Return _targetItem
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            MyBase.Dispose(disposing)
        End If

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

                If Not _targetItem Is Nothing Then
                    _targetItem.Dispose()
                    _targetItem = Nothing
                End If
            End Sub)
    End Sub
End Class

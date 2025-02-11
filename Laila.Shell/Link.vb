Imports System.Runtime.InteropServices
Imports System.Text

Public Class Link
    Inherits Item

    Private _shellLinkW As IShellLinkW
    Private _targetPidl As Pidl
    Private _targetFullPath As String

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, Optional pidl As Pidl = Nothing)
        MyBase.New(shellItem2, logicalParent, doKeepAlive, doHookUpdates, pidl)

        'CType(shellItem2, IShellItem2ForShellLink).BindToHandler(Nothing, Guids.BHID_SFUIObject, GetType(IShellLinkW).GUID, _shellLinkW)
    End Sub

    'Public ReadOnly Property TargetPidl As Pidl
    '    Get
    '        If _targetPidl Is Nothing Then
    '            Try
    '                Dim pidl As IntPtr
    '                _shellLinkW.GetIDList(pidl)
    '                If Not IntPtr.Zero.Equals(pidl) Then
    '                    _targetPidl = New Pidl(pidl)
    '                End If
    '            Catch ex As Exception
    '            End Try
    '        End If

    '        Return _targetPidl
    '    End Get
    'End Property

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

    'Public Function GetTarget(parent As Folder, Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True) As Item
    '    Return Item.FromPidl(Me.TargetPidl, parent, doKeepAlive, doHookUpdates)
    'End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        SyncLock _shellItemLock
            If Not disposedValue Then
                MyBase.Dispose(disposing)

                If Not _shellLinkW Is Nothing Then
                    Marshal.ReleaseComObject(_shellLinkW)
                    _shellLinkW = Nothing
                End If

                If Not _targetPidl Is Nothing Then
                    _targetPidl.Dispose()
                    _targetPidl = Nothing
                End If
            End If
        End SyncLock
    End Sub
End Class

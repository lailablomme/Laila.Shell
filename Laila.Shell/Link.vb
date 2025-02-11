Imports System.Runtime.InteropServices

Public Class Link
    Inherits Item

    Private _shellLinkW As IShellLinkW
    Private _targetPidl As Pidl

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, Optional pidl As IntPtr? = Nothing)
        MyBase.New(shellItem2, logicalParent, doKeepAlive, doHookUpdates, pidl)

        CType(shellItem2, IShellItem2ForShellLink).BindToHandler(Nothing, Guids.BHID_SFUIObject, GetType(IShellLinkW).GUID, _shellLinkW)
    End Sub

    Public ReadOnly Property TargetPidl As Pidl
        Get
            If _targetPidl Is Nothing Then
                Try
                    Dim pidl As IntPtr
                    _shellLinkW.GetIDList(pidl)
                    If Not IntPtr.Zero.Equals(pidl) Then
                        _targetPidl = New Pidl(pidl)
                    End If
                Catch ex As Exception
                End Try
            End If

            Return _targetPidl
        End Get
    End Property

    Public Function GetTarget(parent As Folder, Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True) As Item
        Return Item.FromPidl(Me.TargetPidl.AbsolutePIDL, parent, doKeepAlive, doHookUpdates)
    End Function

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

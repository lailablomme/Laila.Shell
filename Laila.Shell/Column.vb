Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Windows

Public Class Column
    Private _columnInfo As CM_COLUMNINFO
    Friend _columnManager As IColumnManager
    Private _propertyKey As PROPERTYKEY
    Private _index As Integer
    Friend _propertyDescription As IPropertyDescription
    Private _viewFlags As PROPDESC_VIEW_FLAGS
    Private _canonicalName As String

    Friend Sub New(propertyKey As PROPERTYKEY, columnManager As IColumnManager, index As Integer)
        _propertyKey = propertyKey
        _columnManager = columnManager
        _index = index

        Dim ptr2 As IntPtr
        Try
            Functions.PSGetPropertyDescription(propertyKey, GetType(IPropertyDescription).GUID, ptr2)
            If Not IntPtr.Zero.Equals(ptr2) Then
                _propertyDescription = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IPropertyDescription))
                _propertyDescription.GetViewFlags(_viewFlags)
            End If
        Finally
            If Not IntPtr.Zero.Equals(ptr2) Then
                Marshal.Release(ptr2)
            End If
        End Try
    End Sub

    Public ReadOnly Property CanonicalName As String
        Get
            If String.IsNullOrWhiteSpace(_canonicalName) Then
                _propertyDescription.GetCanonicalName(_canonicalName)
            End If

            Return _canonicalName
        End Get
    End Property

    Public ReadOnly Property Index As Integer
        Get
            Return _index
        End Get
    End Property

    Public ReadOnly Property Alignment As TextAlignment
        Get
            Return If(CBool(_viewFlags And PROPDESC_VIEW_FLAGS.PDVF_CENTERALIGN), TextAlignment.Center,
                   If(CBool(_viewFlags And PROPDESC_VIEW_FLAGS.PDVF_RIGHTALIGN), TextAlignment.Right,
                   TextAlignment.Left))
        End Get
    End Property

    Friend ReadOnly Property PROPERTYKEY As PROPERTYKEY
        Get
            Return _propertyKey
        End Get
    End Property

    Private ReadOnly Property CM_COLUMNINFO As CM_COLUMNINFO
        Get
            If _columnInfo.dwMask = 0 Then
                _columnInfo.dwMask = CM_MASK.CM_MASK_NAME Or CM_MASK.CM_MASK_DEFAULTWIDTH Or CM_MASK.CM_MASK_IDEALWIDTH Or CM_MASK.CM_MASK_STATE Or CM_MASK.CM_MASK_WIDTH
                _columnInfo.cbSize = Marshal.SizeOf(Of CM_COLUMNINFO)
                _columnManager.GetColumnInfo(_propertyKey, _columnInfo)
            End If
            Return _columnInfo
        End Get
    End Property

    Public ReadOnly Property State As CM_STATE
        Get
            Return Me.CM_COLUMNINFO.dwState
        End Get
    End Property

    Public ReadOnly Property Width As Integer
        Get
            Return Me.CM_COLUMNINFO.uWidth
        End Get
    End Property

    Public ReadOnly Property DefaultWidth As UInt32
        Get
            Return Me.CM_COLUMNINFO.uDefaultWidth
        End Get
    End Property

    Public ReadOnly Property IdealWidth As UInt32
        Get
            Return Me.CM_COLUMNINFO.uIdealWidth
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayName As String
        Get
            Return Me.CM_COLUMNINFO.wszName
        End Get
    End Property
End Class

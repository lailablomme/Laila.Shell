Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows

Public Class Column
    Inherits NotifyPropertyChangedBase

    Private _columnInfo As CM_COLUMNINFO
    Private _propertyKey As PROPERTYKEY
    Private _index As Integer
    Friend _propertyDescription As IPropertyDescription
    Private _viewFlags As PROPDESC_VIEW_FLAGS
    Protected _canonicalName As String
    Private _isVisible As Boolean
    Protected _displayName As String

    Friend Sub New(propertyKey As PROPERTYKEY, columnInfo As CM_COLUMNINFO, index As Integer)
        _propertyKey = propertyKey
        _columnInfo = columnInfo
        _index = index
        _isVisible = Me.State.HasFlag(CM_STATE.VISIBLE)

        Functions.PSGetPropertyDescription(propertyKey, GetType(IPropertyDescription).GUID, _propertyDescription)
        If Not _propertyDescription Is Nothing Then
            _propertyDescription.GetViewFlags(_viewFlags)
        End If
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
            Return If(_viewFlags.HasFlag(PROPDESC_VIEW_FLAGS.PDVF_CENTERALIGN), TextAlignment.Center,
                   If(_viewFlags.HasFlag(PROPDESC_VIEW_FLAGS.PDVF_RIGHTALIGN), TextAlignment.Right,
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
            If String.IsNullOrWhiteSpace(_displayName) Then
                If String.IsNullOrWhiteSpace(Me.CM_COLUMNINFO.wszName?.Replace(Chr(0), "")) Then
                    Dim dn As StringBuilder = New StringBuilder()
                    _propertyDescription.GetDisplayName(dn)
                    _displayName = dn.ToString()
                Else
                    _displayName = Me.CM_COLUMNINFO.wszName
                End If
            End If
            Return _displayName
        End Get
    End Property

    Public Property IsVisible As Boolean
        Get
            Return _isVisible
        End Get
        Set(value As Boolean)
            SetValue(_isVisible, value)
        End Set
    End Property
End Class

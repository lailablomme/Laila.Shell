Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class System_StorageProviderUIStatusProperty
    Inherits [Property]

    Public Shared Property System_StorageProviderUIStatusKey As PROPERTYKEY = New PROPERTYKEY("e77e90df-6271-4f5b-834f-2dd1f245dda4", 2)

    Private _system_StorageProviderStateKey As PROPERTYKEY = New PROPERTYKEY("e77e90df-6271-4f5b-834f-2dd1f245dda4", 3)
    Private _system_StorageProviderCustomStatesKey As PROPERTYKEY = New PROPERTYKEY("fceff153-e839-4cf3-a9e7-ea22832094b8", 120)
    Private _system_ItemCustomState_ValuesKey As PROPERTYKEY = New PROPERTYKEY("fceff153-e839-4cf3-a9e7-ea22832094b8", 122)
    Private _system_ItemCustomState_IconReferencesKey As PROPERTYKEY = New PROPERTYKEY("fceff153-e839-4cf3-a9e7-ea22832094b8", 123)

    Private _system_StorageProviderStateProperty As [Property]
    Private _system_StorageProviderCustomStatesProperty As [Property]
    Private _system_ItemCustomState_ValuesProperty As [Property]
    Private _system_ItemCustomState_IconReferencesProperty As [Property]
    Private _propertyStore1 As IPropertyStore
    Private _propertyStore2 As IPropertyStore

    Private _didReadData As Boolean

    Public Sub New()
        MyBase.New(System_StorageProviderUIStatusKey, CType(Nothing, IPropertyStore))
    End Sub

    Public Sub New(propertyStore As IPropertyStore)
        MyBase.New(System_StorageProviderUIStatusKey, propertyStore)
        If Not propertyStore Is Nothing Then
            readData()
        End If
    End Sub

    Public Sub New(shellItem2 As IShellItem2)
        MyBase.New(System_StorageProviderUIStatusKey, shellItem2)
        If Not shellItem2 Is Nothing Then
            readData()
        End If
    End Sub

    Private Sub readData()
        _didReadData = True

        Dim ptr As IntPtr, persistSerializedPropStorage As IPersistSerializedPropStorage
        If _propertyStore1 Is Nothing AndAlso Me.RawValue.vt > 0 Then
            Try
                Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                _propertyStore1 = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                persistSerializedPropStorage = _propertyStore1
                persistSerializedPropStorage.SetFlags(0)
                persistSerializedPropStorage.SetPropertyStorage(
                    MyBase.RawValue.union.bstrblobVal.pData, MyBase.RawValue.union.bstrblobVal.cbSize)

                If _system_StorageProviderStateProperty Is Nothing Then
                    _system_StorageProviderStateProperty = New [Property](_system_StorageProviderStateKey, _propertyStore1)
                End If
                If _system_StorageProviderCustomStatesProperty Is Nothing Then
                    _system_StorageProviderCustomStatesProperty = New [Property](_system_StorageProviderCustomStatesKey, _propertyStore1)
                End If
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
            End Try
        End If

        If _propertyStore2 Is Nothing AndAlso Not _propertyStore1 Is Nothing Then
            Try
                Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                _propertyStore2 = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                persistSerializedPropStorage = _propertyStore2
                persistSerializedPropStorage.SetFlags(0)
                persistSerializedPropStorage.SetPropertyStorage(
                    _system_StorageProviderCustomStatesProperty.RawValue.union.bstrblobVal.pData,
                    _system_StorageProviderCustomStatesProperty.RawValue.union.bstrblobVal.cbSize)

                If _system_ItemCustomState_ValuesProperty Is Nothing Then
                    _system_ItemCustomState_ValuesProperty = New [Property](_system_ItemCustomState_ValuesKey, _propertyStore2)
                End If
                If _system_ItemCustomState_IconReferencesProperty Is Nothing Then
                    _system_ItemCustomState_IconReferencesProperty = New [Property](_system_ItemCustomState_IconReferencesKey, _propertyStore2)
                End If
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
                If Not _propertyStore1 Is Nothing Then
                    Marshal.ReleaseComObject(_propertyStore1)
                    _propertyStore1 = Nothing
                End If
                If Not _propertyStore2 Is Nothing Then
                    Marshal.ReleaseComObject(_propertyStore2)
                    _propertyStore2 = Nothing
                End If
            End Try
        End If
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            If Not _system_StorageProviderStateProperty Is Nothing Then
                Return _system_StorageProviderStateProperty.DisplayName
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property ActivityDisplayName As String
        Get
            If Not _system_StorageProviderCustomStatesProperty Is Nothing Then
                Return _system_StorageProviderCustomStatesProperty.DisplayName()
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property Text As String
        Get
            If Not disposedValue AndAlso String.IsNullOrWhiteSpace(_text) _
                AndAlso Not _system_StorageProviderStateProperty Is Nothing Then
                _text = _system_StorageProviderStateProperty.Text
            End If

            Return _text
        End Get
    End Property

    Public ReadOnly Property ActivityText As String
        Get
            If Not _system_ItemCustomState_ValuesProperty Is Nothing Then
                Return _system_ItemCustomState_ValuesProperty.Text
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property HasIcon As Boolean
        Get
            Return Not _didReadData OrElse Me.RawValue.vt > 0
        End Get
    End Property

    Public Overrides ReadOnly Property Icons16 As ImageSource()
        Get
            Dim icons As List(Of ImageSource) = New List(Of ImageSource)()

            If Not _system_StorageProviderStateProperty Is Nothing Then
                Dim mainIcons As ImageSource() = _system_StorageProviderStateProperty.Icons16
                If Not mainIcons Is Nothing Then
                    icons.AddRange(mainIcons)
                End If
            End If

            If Not _system_ItemCustomState_IconReferencesProperty Is Nothing Then
                If Not String.IsNullOrWhiteSpace(_system_ItemCustomState_IconReferencesProperty.Text) Then
                    For Each iconReference In _system_ItemCustomState_IconReferencesProperty.Text.Split(";")
                        icons.Add(ImageHelper.ExtractIcon(iconReference))
                    Next
                End If
            End If

            Return icons.ToArray()
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)

        If Not _system_StorageProviderStateProperty Is Nothing Then
            _system_StorageProviderStateProperty.Dispose()
            _system_StorageProviderStateProperty = Nothing
        End If
        If Not _system_StorageProviderCustomStatesProperty Is Nothing Then
            _system_StorageProviderCustomStatesProperty.Dispose()
            _system_StorageProviderCustomStatesProperty = Nothing
        End If
        If Not _system_ItemCustomState_ValuesProperty Is Nothing Then
            _system_ItemCustomState_ValuesProperty.Dispose()
            _system_ItemCustomState_ValuesProperty = Nothing
        End If
        If Not _system_ItemCustomState_IconReferencesProperty Is Nothing Then
            _system_ItemCustomState_IconReferencesProperty.Dispose()
            _system_ItemCustomState_IconReferencesProperty = Nothing
        End If
    End Sub
End Class

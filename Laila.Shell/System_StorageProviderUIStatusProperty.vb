Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class System_StorageProviderUIStatusProperty
    Inherits [Property]

    Public Overloads Shared Property Key As PROPERTYKEY = New PROPERTYKEY("e77e90df-6271-4f5b-834f-2dd1f245dda4", 2)
    Public Shared Shadows Property CanonicalName As String = "System.StorageProviderUIStatus"

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
    Private _hasData As Boolean

    Private _imageReference16 As String()

    Public Sub New()
        MyBase.New([Property].getDescription(Key))
    End Sub

    Public Sub New(propertyStore As IPropertyStore)
        MyBase.New([Property].GetDescription(Key), propertyStore)
        If Not propertyStore Is Nothing Then _hasData = True
    End Sub

    Public Sub New(shellItem2 As IShellItem2)
        MyBase.New([Property].GetDescription(Key), shellItem2)
        If Not shellItem2 Is Nothing Then _hasData = True
    End Sub

    Private Sub readData()
        If Not _didReadData Then
            _didReadData = True

            Dim persistSerializedPropStorage As IPersistSerializedPropStorage
            If _propertyStore1 Is Nothing AndAlso Me.RawValue.vt > 0 Then
                Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, _propertyStore1)
                persistSerializedPropStorage = _propertyStore1
                persistSerializedPropStorage.SetFlags(0)
                persistSerializedPropStorage.SetPropertyStorage(
                    MyBase.RawValue.union.bstrblobVal.pData, MyBase.RawValue.union.bstrblobVal.cbSize)

                If _system_StorageProviderStateProperty Is Nothing Then
                    _system_StorageProviderStateProperty = New [Property]([Property].GetDescription(_system_StorageProviderStateKey), _propertyStore1)
                End If
                If _system_StorageProviderCustomStatesProperty Is Nothing Then
                    _system_StorageProviderCustomStatesProperty = New [Property]([Property].GetDescription(_system_StorageProviderCustomStatesKey), _propertyStore1)
                End If
            End If

            If _propertyStore2 Is Nothing AndAlso Not _propertyStore1 Is Nothing Then
                Try
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, _propertyStore2)
                    persistSerializedPropStorage = _propertyStore2
                    persistSerializedPropStorage.SetFlags(0)
                    persistSerializedPropStorage.SetPropertyStorage(
                        _system_StorageProviderCustomStatesProperty.RawValue.union.bstrblobVal.pData,
                        _system_StorageProviderCustomStatesProperty.RawValue.union.bstrblobVal.cbSize)

                    If _system_ItemCustomState_ValuesProperty Is Nothing Then
                        _system_ItemCustomState_ValuesProperty = New [Property]([Property].GetDescription(_system_ItemCustomState_ValuesKey), _propertyStore2)
                    End If
                    If _system_ItemCustomState_IconReferencesProperty Is Nothing Then
                        _system_ItemCustomState_IconReferencesProperty = New [Property]([Property].GetDescription(_system_ItemCustomState_IconReferencesKey), _propertyStore2)
                    End If
                Finally
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
        End If
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            readData()
            If Not _system_StorageProviderStateProperty Is Nothing Then
                Return _system_StorageProviderStateProperty.DisplayName
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property ActivityDisplayName As String
        Get
            readData()
            If Not _system_StorageProviderCustomStatesProperty Is Nothing Then
                Return _system_StorageProviderCustomStatesProperty.DisplayName()
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property Text As String
        Get
            readData()
            If Not disposedValue AndAlso String.IsNullOrWhiteSpace(_text) _
                AndAlso Not _system_StorageProviderStateProperty Is Nothing Then
                _text = _system_StorageProviderStateProperty.Text
            End If

            Return _text
        End Get
    End Property

    Public Overrides ReadOnly Property Value As Object
        Get
            readData()
            If Not _system_StorageProviderStateProperty Is Nothing Then
                Return _system_StorageProviderStateProperty.Value
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property ActivityText As String
        Get
            readData()
            If Not _system_ItemCustomState_ValuesProperty Is Nothing Then
                Return _system_ItemCustomState_ValuesProperty.Text
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property HasIcon As Boolean
        Get
            Return Not _hasData OrElse Me.RawValue.vt > 0
        End Get
    End Property

    Public Overrides ReadOnly Property ImageReferences16 As String()
        Get
            readData()
            If _imageReference16 Is Nothing Then
                Dim imageReferences As List(Of String) = New List(Of String)()

                If Not _system_StorageProviderStateProperty Is Nothing Then
                    Dim mainReferences As String() = _system_StorageProviderStateProperty.ImageReferences16
                    If Not mainReferences Is Nothing Then
                        imageReferences.AddRange(mainReferences)
                    End If
                End If

                If Not _system_ItemCustomState_IconReferencesProperty Is Nothing Then
                    If Not String.IsNullOrWhiteSpace(_system_ItemCustomState_IconReferencesProperty.Text) Then
                        For Each iconReference In _system_ItemCustomState_IconReferencesProperty.Text.Split(";")
                            imageReferences.Add(iconReference)
                        Next
                    End If
                End If

                _imageReference16 = imageReferences.ToArray()
            End If

            Return _imageReference16
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

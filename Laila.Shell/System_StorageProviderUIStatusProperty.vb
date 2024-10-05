Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class System_StorageProviderUIStatusProperty
    Inherits [Property]

    Public Shared Property PropertyKey = New PROPERTYKEY("e77e90df-6271-4f5b-834f-2dd1f245dda4", 2)

    Private _system_StorageProviderStateDescription As IPropertyDescription
    Private _system_StorageProviderStateKey As PROPERTYKEY
    Private _System_StorageProviderCustomStatesDescription As IPropertyDescription
    Private _System_StorageProviderCustomStatesKey As PROPERTYKEY
    Private _System_ItemCustomState_StateListDescription As IPropertyDescription
    Private _System_ItemCustomState_StateListKey As PROPERTYKEY


    Public Sub New(item As Item)
        MyBase.New(PropertyKey, item)

        Functions.PSGetPropertyDescriptionByName("System.StorageProviderState", GetType(IPropertyDescription).GUID, _system_StorageProviderStateDescription)
        _system_StorageProviderStateDescription.GetPropertyKey(_system_StorageProviderStateKey)

        Functions.PSGetPropertyDescriptionByName("System.StorageProviderCustomStates", GetType(IPropertyDescription).GUID, _System_StorageProviderCustomStatesDescription)
        _System_StorageProviderCustomStatesDescription.GetPropertyKey(_System_StorageProviderCustomStatesKey)

        Functions.PSGetPropertyDescriptionByName("System.ItemCustomState.StateList", GetType(IPropertyDescription).GUID, _System_ItemCustomState_StateListDescription)
        _System_ItemCustomState_StateListDescription.GetPropertyKey(_System_ItemCustomState_StateListKey)
    End Sub

    Friend Overrides ReadOnly Property RawValue As PROPVARIANT
        Get
            Dim result As PROPVARIANT = MyBase.RawValue

            Dim ptr As IntPtr, propertyStore As IPropertyStore, persistSerializedPropStorage As IPersistSerializedPropStorage
            Try
                Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                persistSerializedPropStorage = propertyStore
                persistSerializedPropStorage.SetFlags(0)
                persistSerializedPropStorage.SetPropertyStorage(result.union.bstrblobVal.pData, result.union.bstrblobVal.cbSize)
                result.Dispose()
                propertyStore.GetValue(_system_StorageProviderStateKey, result)
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
                If Not propertyStore Is Nothing Then
                    Marshal.ReleaseComObject(propertyStore)
                End If
            End Try

            Return result
        End Get
    End Property

    Public ReadOnly Property Value As Object
        Get
            Using rawValue As PROPVARIANT = Me.RawValue
                If Me.DisplayType = PropertyDisplayType.Enumerated Then
                    Dim index As UInt32
                    Dim propertyEnumType As IPropertyEnumType = getSelectedPropertyEnumType(rawValue, _system_StorageProviderStateDescription, index)
                    If Not propertyEnumType Is Nothing Then
                        Return index
                    End If
                Else
                    Return getValue(rawValue)
                End If
            End Using

            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property Text As String
        Get
            If String.IsNullOrWhiteSpace(_text) Then
                Dim rawValue As PROPVARIANT = Me.RawValue
                Dim buffer As StringBuilder = New StringBuilder()
                buffer.Append(New String(" ", 2050))
                Functions.PSFormatForDisplay(_system_StorageProviderStateKey, rawValue, PropertyDescriptionFormatOptions.None, buffer, 2048)
                _text = buffer.ToString()
                rawValue.Dispose()

                rawValue = MyBase.RawValue

                Dim ptr As IntPtr, propertyStore As IPropertyStore, persistSerializedPropStorage As IPersistSerializedPropStorage

                Try
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                    propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                    persistSerializedPropStorage = propertyStore
                    persistSerializedPropStorage.SetFlags(0)
                    persistSerializedPropStorage.SetPropertyStorage(rawValue.union.bstrblobVal.pData, rawValue.union.bstrblobVal.cbSize)
                    rawValue.Dispose()
                    propertyStore.GetValue(_System_StorageProviderCustomStatesKey, rawValue)
                Finally
                    If Not IntPtr.Zero.Equals(ptr) Then
                        Marshal.Release(ptr)
                    End If
                    If Not propertyStore Is Nothing Then
                        Marshal.ReleaseComObject(propertyStore)
                    End If
                End Try

                Try
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                    propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                    persistSerializedPropStorage = propertyStore
                    persistSerializedPropStorage.SetFlags(0)
                    persistSerializedPropStorage.SetPropertyStorage(rawValue.union.bstrblobVal.pData, rawValue.union.bstrblobVal.cbSize)
                    rawValue.Dispose()
                    propertyStore.GetValue(_System_ItemCustomState_StateListKey, rawValue)
                Finally
                    If Not IntPtr.Zero.Equals(ptr) Then
                        Marshal.Release(ptr)
                    End If
                    If Not propertyStore Is Nothing Then
                        Marshal.ReleaseComObject(propertyStore)
                    End If
                End Try

                Dim customStates As List(Of String) = getValue(rawValue)
                rawValue.Dispose()

                If Not customStates Is Nothing Then
                    _text &= " (" & String.Join(", ", customStates) & ")"
                End If
            End If

            Return _text
        End Get
    End Property

    Public Overrides ReadOnly Property DisplayType As PropertyDisplayType
        Get
            Dim dt As PropertyDisplayType
            _system_StorageProviderStateDescription.GetDisplayType(dt)
            Return dt
        End Get
    End Property

    Public Overrides ReadOnly Property HasIcon As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property Icon16 As ImageSource
        Get
            If Me.DisplayType = PropertyDisplayType.Enumerated Then
                Using rawValue As PROPVARIANT = Me.RawValue
                    Dim imageReference As String, icon As IntPtr
                    Dim index As UInt32
                    Dim propertyEnumType2 As IPropertyEnumType2 = getSelectedPropertyEnumType(rawValue, _system_StorageProviderStateDescription, index)
                    propertyEnumType2.GetImageReference(imageReference)
                    If Not String.IsNullOrWhiteSpace(imageReference) Then
                        Dim s() As String = Split(imageReference, ",")
                        Functions.ExtractIconEx(s(0), s(1), Nothing, icon, 1)
                        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    End If
                End Using
            End If

            Return Nothing
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)

        If Not _system_StorageProviderStateDescription Is Nothing Then
            Marshal.ReleaseComObject(_system_StorageProviderStateDescription)
        End If
    End Sub
End Class

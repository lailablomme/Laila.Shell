Imports System.ComponentModel.DataAnnotations
Imports System.Drawing.Text
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports MS.WindowsAPICodePack.Internal

Public Class [Property]
    Private _canonicalName As String
    Private _item As Item
    Private _originalPropertyDescription As IPropertyDescription
    Private _originalPropertyKey As PROPERTYKEY
    Private _propertyDescription As IPropertyDescription
    Private _propertyKey As PROPERTYKEY

    Public Shared Function FromCanonicalName(canonicalName As String, Optional item As Item = Nothing) As [Property]
        Dim propertyDescription As IPropertyDescription
        Functions.PSGetPropertyDescriptionByName(canonicalName, GetType(IPropertyDescription).GUID, propertyDescription)
        If Not propertyDescription Is Nothing Then
            Return New [Property](canonicalName, propertyDescription, item)
        Else
            Throw New Exception(String.Format("Property '{0}' not found.", canonicalName))
        End If
    End Function

    Public Sub New(canonicalName As String, propertyDescription As IPropertyDescription, item As Item)
        _canonicalName = canonicalName
        _item = item
        _originalPropertyDescription = propertyDescription
        _originalPropertyDescription.GetPropertyKey(_originalPropertyKey)

        Select Case _canonicalName
            Case "System.StorageProviderUIStatus"
                Functions.PSGetPropertyDescriptionByName("System.StorageProviderState", GetType(IPropertyDescription).GUID, _propertyDescription)
                _propertyDescription.GetPropertyKey(_propertyKey)
            Case Else
                _propertyDescription = _originalPropertyDescription
                _propertyKey = _originalPropertyKey
        End Select
    End Sub

    Friend ReadOnly Property RawValue As PROPVARIANT
        Get
            Dim result As PROPVARIANT

            Select Case _canonicalName
                Case "System.StorageProviderUIStatus"
                    _item._shellItem2.GetProperty(_originalPropertyKey, result)
                    Dim ptr As IntPtr, propertyStore As IPropertyStore, persistSerializedPropStorage As IPersistSerializedPropStorage
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
                    propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
                    persistSerializedPropStorage = propertyStore
                    persistSerializedPropStorage.SetFlags(0)
                    persistSerializedPropStorage.SetPropertyStorage(result.union.bstrblobVal.pData, result.union.bstrblobVal.cbSize)
                    result.Dispose()
                    propertyStore.GetValue(_propertyKey, result)
                Case Else
                    _item._shellItem2.GetProperty(_propertyKey, result)
            End Select

            Return result
        End Get
    End Property

    Public ReadOnly Property Value As Object
        Get
            Using rawValue As PROPVARIANT = Me.RawValue
                If Me.DisplayType = PropertyDisplayType.Enumerated Then
                    Dim index As UInt32
                    Dim propertyEnumType As IPropertyEnumType = getSelectedPropertyEnumType(rawValue, index)
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

    Public ReadOnly Property Text As String
        Get
            Using rawValue As PROPVARIANT = Me.RawValue
                Dim buffer As StringBuilder = New StringBuilder()
                Functions.PSFormatForDisplay(_propertyKey, rawValue, PropertyDescriptionFormatOptions.None, buffer, 2048)
                Return Buffer.ToString()
            End Using
        End Get
    End Property

    Public ReadOnly Property DisplayType As PropertyDisplayType
        Get
            Dim dt As PropertyDisplayType
            _propertyDescription.GetDisplayType(dt)
            Return dt
        End Get
    End Property

    Public ReadOnly Property RelativeDescriptionType As RelativeDescriptionType
        Get
            Dim rdt As RelativeDescriptionType
            _propertyDescription.GetRelativeDescriptionType(rdt)
            Return rdt
        End Get
    End Property

    Public ReadOnly Property HasIcon As Boolean
        Get
            If Me.DisplayType = PropertyDisplayType.Enumerated Then
                Dim propertyEnumTypeList As IPropertyEnumTypeList
                _propertyDescription.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                Dim count As UInt32
                propertyEnumTypeList.GetCount(count)
                Dim propertyEnumType As IPropertyEnumType
                For x As UInt32 = 0 To count - 1
                    propertyEnumTypeList.GetAt(x, GetType(IPropertyEnumType).GUID, propertyEnumType)
                    Dim imageReference As String
                    Dim propertyEnumType2 As IPropertyEnumType2 = propertyEnumType
                    propertyEnumType2.GetImageReference(imageReference)
                    If Not String.IsNullOrWhiteSpace(imageReference) Then
                        Return True
                    End If
                Next
            End If

            Return False
        End Get
    End Property

    Public ReadOnly Property Icon16 As ImageSource
        Get
            If Me.DisplayType = PropertyDisplayType.Enumerated Then
                Using rawValue As PROPVARIANT = Me.RawValue
                    Dim imageReference As String, icon As IntPtr
                    Dim index As UInt32
                    Dim propertyEnumType2 As IPropertyEnumType2 = getSelectedPropertyEnumType(rawValue, index)
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

    Private Function getValue(value As PROPVARIANT) As Object
        Dim ve As VarEnum = value.vt
        Select Case ve
            Case VarEnum.VT_EMPTY
                Return Nothing
            Case VarEnum.VT_BSTR
                Return Marshal.PtrToStringBSTR(value.union.ptr)
            Case VarEnum.VT_LPWSTR
                Return Marshal.PtrToStringUni(value.union.ptr)
            Case VarEnum.VT_FILETIME
                Return DateTime.FromFileTime(value.union.hVal)
            Case VarEnum.VT_I1
                Return Chr(value.union.bVal)
            Case VarEnum.VT_I2
                Return value.union.iVal
            Case VarEnum.VT_UI2
                Return value.union.uiVal
            Case VarEnum.VT_I4
                Return value.union.intVal
            Case VarEnum.VT_UI4
                Return value.union.uintVal
            Case VarEnum.VT_UI8
                Return value.union.uhVal
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function getSelectedPropertyEnumType(value As PROPVARIANT, ByRef index As UInt32) As IPropertyEnumType
        If Me.DisplayType = PropertyDisplayType.Enumerated Then
            Dim propertyEnumTypeList As IPropertyEnumTypeList
            _propertyDescription.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
            Dim count As UInt32
            propertyEnumTypeList.GetCount(count)
            Dim propertyEnumType As IPropertyEnumType
            Dim obj As Object = getValue(value)
            For x As UInt32 = 0 To count - 1
                propertyEnumTypeList.GetAt(x, GetType(IPropertyEnumType).GUID, propertyEnumType)
                If x = Val(obj) Then
                    index = x
                    Return propertyEnumType
                End If
            Next
        End If

        Return Nothing
    End Function
End Class

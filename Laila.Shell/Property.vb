Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports System.Drawing.Text
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports MS.WindowsAPICodePack.Internal

Public Class [Property]
    Implements IDisposable

    Protected _canonicalName As String
    Protected _item As Item
    Protected _propertyDescription As IPropertyDescription
    Protected _propertyKey As PROPERTYKEY
    Protected _text As String
    Private _hasIcon As Boolean?
    Protected disposedValue As Boolean
    Private _rawValue As PROPVARIANT
    Private _icon16 As ImageSource

    Public Shared Function FromCanonicalName(canonicalName As String, Optional item As Item = Nothing) As [Property]
        If canonicalName = "System.StorageProviderUIStatus" Then
            Return New System_StorageProviderUIStatusProperty(item)
        Else
            Dim propertyDescription As IPropertyDescription
            Functions.PSGetPropertyDescriptionByName(canonicalName, GetType(IPropertyDescription).GUID, propertyDescription)
            If Not propertyDescription Is Nothing Then
                Return New [Property](canonicalName, propertyDescription, item)
            Else
                Return Nothing
            End If
        End If
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional item As Item = Nothing) As [Property]
        If propertyKey.Equals(System_StorageProviderUIStatusProperty.PropertyKey) Then
            Return New System_StorageProviderUIStatusProperty(item)
        Else
            Return New [Property](propertyKey, item)
        End If
    End Function

    Public Sub New(canonicalName As String, propertyDescription As IPropertyDescription, item As Item)
        _canonicalName = canonicalName
        propertyDescription.GetPropertyKey(_propertyKey)
        _item = item
    End Sub

    Public Sub New(propertyKey As PROPERTYKEY, item As Item)
        _propertyKey = propertyKey
        _item = item
    End Sub

    Public ReadOnly Property CanonicalName As String
        Get
            If String.IsNullOrEmpty(_canonicalName) Then
                Functions.PSGetNameFromPropertyKey(_propertyKey, _canonicalName)
            End If
            Return _canonicalName
        End Get
    End Property

    Public ReadOnly Property Key As PROPERTYKEY
        Get
            Return _propertyKey
        End Get
    End Property

    Public ReadOnly Property DescriptionDisplayName As String
        Get
            Dim displayName As StringBuilder = New StringBuilder()
            Me.Description.GetDisplayName(displayName)
            Return displayName.ToString()
        End Get
    End Property

    Public ReadOnly Property DescriptionDisplayNameWithColon As String
        Get
            Dim displayName As StringBuilder = New StringBuilder()
            Me.Description.GetDisplayName(displayName)
            Return displayName.ToString() & ":"
        End Get
    End Property

    Public ReadOnly Property Description As IPropertyDescription
        Get
            If _propertyDescription Is Nothing Then
                Functions.PSGetPropertyDescriptionByName(Me.CanonicalName, GetType(IPropertyDescription).GUID, _propertyDescription)
            End If
            Return _propertyDescription
        End Get
    End Property

    Friend Overridable ReadOnly Property RawValue As PROPVARIANT
        Get
            Dim result As PROPVARIANT
            Try
                If Not _item.disposedValue Then
                    Dim shellItem2 As IShellItem2 = _item.ShellItem22
                    Try
                        shellItem2.GetProperty(_propertyKey, result)
                    Finally
                        If Not shellItem2 Is Nothing Then
                            Marshal.ReleaseComObject(shellItem2)
                        End If
                    End Try
                End If
            Catch ex As COMException
            End Try
            Return result
        End Get
    End Property

    Public ReadOnly Property Value As Object
        Get
            Using rawValue As PROPVARIANT = Me.RawValue
                If Me.DisplayType = PropertyDisplayType.Enumerated Then
                    Dim index As UInt32
                    Dim propertyEnumType As IPropertyEnumType = getSelectedPropertyEnumType(rawValue, Me.Description, index)
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

    Public Overridable ReadOnly Property Text As String
        Get
            If String.IsNullOrWhiteSpace(_text) Then
                Using rawValue As PROPVARIANT = Me.RawValue
                    If rawValue.vt > 0 Then
                        Dim buffer As StringBuilder = New StringBuilder()
                        buffer.Append(New String(" ", 2050))
                        Functions.PSFormatForDisplay(_propertyKey, rawValue, PropertyDescriptionFormatOptions.None, buffer, 2048)
                        _text = buffer.ToString()
                    End If
                End Using
            End If
            Return _text
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayType As PropertyDisplayType
        Get
            Dim dt As PropertyDisplayType
            Me.Description.GetDisplayType(dt)
            Return dt
        End Get
    End Property

    Public ReadOnly Property RelativeDescriptionType As RelativeDescriptionType
        Get
            Dim rdt As RelativeDescriptionType
            Me.Description.GetRelativeDescriptionType(rdt)
            Return rdt
        End Get
    End Property

    Public Overridable ReadOnly Property HasIcon As Boolean
        Get
            If Not _hasIcon.HasValue Then
                _hasIcon = False

                If Me.DisplayType = PropertyDisplayType.Enumerated Then
                    Dim propertyEnumTypeList As IPropertyEnumTypeList
                    Me.Description.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                    Dim count As UInt32
                    propertyEnumTypeList.GetCount(count)
                    Dim propertyEnumType As IPropertyEnumType
                    For x As UInt32 = 0 To count - 1
                        propertyEnumTypeList.GetAt(x, GetType(IPropertyEnumType).GUID, propertyEnumType)
                        Dim imageReference As String
                        Dim propertyEnumType2 As IPropertyEnumType2 = propertyEnumType
                        propertyEnumType2.GetImageReference(imageReference)
                        If Not String.IsNullOrWhiteSpace(imageReference) Then
                            _hasIcon = True
                        End If
                    Next
                End If
            End If

            Return _hasIcon.Value
        End Get
    End Property

    Public Overridable ReadOnly Property Icon16 As ImageSource
        Get
            If _icon16 Is Nothing Then
                If Me.DisplayType = PropertyDisplayType.Enumerated Then
                    Using rawValue As PROPVARIANT = Me.RawValue
                        Dim imageReference As String, icon As IntPtr
                        Dim index As UInt32
                        Dim propertyEnumType2 As IPropertyEnumType2 = getSelectedPropertyEnumType(rawValue, Me.Description, index)
                        propertyEnumType2.GetImageReference(imageReference)
                        If Not String.IsNullOrWhiteSpace(imageReference) Then
                            Dim s() As String = Split(imageReference, ",")
                            Functions.ExtractIconEx(s(0), s(1), Nothing, icon, 1)
                            _icon16 = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        End If
                    End Using
                End If
            End If

            Return _icon16
        End Get
    End Property

    Protected Function getValue(value As PROPVARIANT) As Object
        Dim ve As VarEnum = value.vt
        If ve.HasFlag(VarEnum.VT_VECTOR) Then
            ve = ve Xor VarEnum.VT_VECTOR
            Select Case ve
                Case VarEnum.VT_UNKNOWN
                    Dim count As Integer = value.union.bstrblobVal.cbSize
                    Dim resultList As New List(Of IntPtr)()
                    Dim pArray As IntPtr = value.union.bstrblobVal.pData
                    For i As Integer = 0 To count - 1
                        Dim pUnknown As IntPtr = Marshal.ReadIntPtr(pArray, i * IntPtr.Size)
                        'Dim obj As Object = Marshal.GetObjectForIUnknown(pUnknown)
                        resultList.Add(pUnknown)
                    Next
                    Return resultList
                Case VarEnum.VT_LPWSTR
                    Dim count As Integer = value.union.bstrblobVal.cbSize
                    Dim resultList As New List(Of String)()
                    Dim pData As IntPtr = value.union.bstrblobVal.pData
                    For i As Integer = 0 To count - 1
                        Dim ptr As IntPtr = Marshal.ReadIntPtr(pData, i * IntPtr.Size)
                        Dim str As String = Marshal.PtrToStringUni(ptr)
                        resultList.Add(str)
                    Next
                    Return resultList
                Case VarEnum.VT_UI1
                    Dim count As Integer = value.union.bstrblobVal.cbSize
                    Dim byteArray(count - 1) As Byte
                    Marshal.Copy(value.union.bstrblobVal.pData, byteArray, 0, count)
                    Return byteArray
            End Select
        Else
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
                    Debug.WriteLine("Unknown " & ve.ToString())
                    Return Nothing
            End Select
        End If
    End Function

    Protected Function getSelectedPropertyEnumType(value As PROPVARIANT, description As IPropertyDescription, ByRef index As UInt32) As IPropertyEnumType
        If Me.DisplayType = PropertyDisplayType.Enumerated Then
            Dim propertyEnumTypeList As IPropertyEnumTypeList
            description.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
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

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' dispose managed state (managed objects)
            End If

            ' free unmanaged resources (unmanaged objects) and override finalizer
            ' set large fields to null
            If Not _propertyDescription Is Nothing Then
                Marshal.ReleaseComObject(_propertyDescription)
            End If
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

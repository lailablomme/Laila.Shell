Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class [Property]
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Private Shared _descriptionsByCanonicalName As Dictionary(Of String, IPropertyDescription) = New Dictionary(Of String, IPropertyDescription)
    Private Shared _descriptionsByPropertyKey As Dictionary(Of String, IPropertyDescription) = New Dictionary(Of String, IPropertyDescription)
    Private Shared _descriptionsLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _hasIcon As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)()
    Private Shared _hasIconLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _imageReferences16 As Dictionary(Of String, String()) = New Dictionary(Of String, String())()
    Private Shared _imageReferences16Lock As SemaphoreSlim = New SemaphoreSlim(1, 1)

    Protected _canonicalName As String
    Protected _propertyDescription As IPropertyDescription
    Protected _propertyKey As PROPERTYKEY
    Protected _text As String
    Protected disposedValue As Boolean
    Private _rawValue As PROPVARIANT
    Private _displayType As PropertyDisplayType = -1

    Private Shared Function getDescription(canonicalName As String) As IPropertyDescription
        _descriptionsLock.Wait()
        Dim propertyDescription As IPropertyDescription
        Try
            If Not _descriptionsByCanonicalName.TryGetValue(canonicalName, propertyDescription) Then
                Dim pkey As PROPERTYKEY
                Functions.PSGetPropertyDescriptionByName(canonicalName, GetType(IPropertyDescription).GUID, propertyDescription)
                If propertyDescription Is Nothing Then Return Nothing
                _descriptionsByCanonicalName.Add(canonicalName, propertyDescription)
            End If
        Finally
            _descriptionsLock.Release()
        End Try
        Return propertyDescription
    End Function

    Private Shared Function getDescription(pkey As PROPERTYKEY) As IPropertyDescription
        _descriptionsLock.Wait()
        Dim propertyDescription As IPropertyDescription
        Try
            If Not _descriptionsByPropertyKey.TryGetValue(pkey.ToString(), propertyDescription) Then
                Dim canonicalName As String
                Functions.PSGetPropertyDescription(pkey, GetType(IPropertyDescription).GUID, propertyDescription)
                If propertyDescription Is Nothing Then Return Nothing
                _descriptionsByPropertyKey.Add(pkey.ToString(), propertyDescription)
            End If
        Finally
            _descriptionsLock.Release()
        End Try
        Return propertyDescription
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String) As [Property]
        If canonicalName = "System.StorageProviderUIStatus" Then
            Return New System_StorageProviderUIStatusProperty()
        Else
            Dim desc As IPropertyDescription = [Property].getDescription(canonicalName)
            If Not desc Is Nothing Then
                Return New [Property](canonicalName)
            Else
                Return Nothing
            End If
        End If
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String, propertyStore As IPropertyStore) As [Property]
        If canonicalName = "System.StorageProviderUIStatus" Then
            Return New System_StorageProviderUIStatusProperty(propertyStore)
        Else
            Return New [Property](canonicalName, propertyStore)
        End If
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String, shellItem2 As IShellItem2) As [Property]
        If canonicalName = "System.StorageProviderUIStatus" Then
            Return New System_StorageProviderUIStatusProperty(shellItem2)
        Else
            Return New [Property](canonicalName, shellItem2)
        End If
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional propertyStore As IPropertyStore = Nothing) As [Property]
        If propertyKey.Equals(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey) Then
            Return New System_StorageProviderUIStatusProperty(propertyStore)
        Else
            Return New [Property](propertyKey, propertyStore)
        End If
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional shellItem2 As IShellItem2 = Nothing) As [Property]
        If propertyKey.Equals(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey) Then
            Return New System_StorageProviderUIStatusProperty(shellItem2)
        Else
            Return New [Property](propertyKey, shellItem2)
        End If
    End Function

    Public Sub New(canonicalName As String, Optional propertyStore As IPropertyStore = Nothing)
        _canonicalName = canonicalName
        _propertyDescription = [Property].getDescription(canonicalName)
        If Not _propertyDescription Is Nothing Then
            _propertyDescription.GetPropertyKey(_propertyKey)
            If Not propertyStore Is Nothing Then
                propertyStore.GetValue(_propertyKey, _rawValue)
            End If
        End If
    End Sub

    Public Sub New(canonicalName As String, shellItem2 As IShellItem2)
        _canonicalName = canonicalName
        _propertyDescription = [Property].getDescription(canonicalName)
        If Not _propertyDescription Is Nothing Then
            _propertyDescription.GetPropertyKey(_propertyKey)
            If Not shellItem2 Is Nothing Then
                shellItem2.GetProperty(_propertyKey, _rawValue)
            End If
        End If
    End Sub

    Public Sub New(propertyKey As PROPERTYKEY, propertyStore As IPropertyStore)
        _propertyKey = propertyKey
        If Not propertyStore Is Nothing Then
            propertyStore.GetValue(_propertyKey, _rawValue)
        End If
    End Sub

    Public Sub New(propertyKey As PROPERTYKEY, shellItem2 As IShellItem2)
        _propertyKey = propertyKey
        If Not shellItem2 Is Nothing Then
            shellItem2.GetProperty(_propertyKey, _rawValue)
        End If
    End Sub

    Public ReadOnly Property CanonicalName As String
        Get
            If String.IsNullOrWhiteSpace(_canonicalName) Then
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

    Public Overridable ReadOnly Property DisplayName As String
        Get
            Dim name As StringBuilder = New StringBuilder()
            Me.Description.GetDisplayName(name)
            Return name.ToString()
        End Get
    End Property

    Public ReadOnly Property DisplayNameWithColon As String
        Get
            Return Me.DisplayName.ToString() & ":"
        End Get
    End Property

    Public ReadOnly Property Description As IPropertyDescription
        Get
            If _propertyDescription Is Nothing Then
                _propertyDescription = [Property].getDescription(_propertyKey)
            End If
            Return _propertyDescription
        End Get
    End Property

    Friend Overridable ReadOnly Property RawValue As PROPVARIANT
        Get
            Return _rawValue
        End Get
    End Property

    Public Overridable ReadOnly Property Value As Object
        Get
            Return Me.RawValue.GetValue()
        End Get
    End Property

    Public Overridable ReadOnly Property Text As String
        Get
            If _text Is Nothing Then
                Dim buffer As StringBuilder = New StringBuilder()
                buffer.Append(New String(" ", 2050))
                Functions.PSFormatForDisplay(_propertyKey, Me.RawValue, PropertyDescriptionFormatOptions.None, buffer, 2048)
                _text = buffer.ToString()
            End If
            Return _text
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayType As PropertyDisplayType
        Get
            If _displayType = -1 Then
                Me.Description.GetDisplayType(_displayType)
            End If
            Return _displayType
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
            Dim hi As Boolean

            _hasIconLock.Wait()
            Try
                If Not _hasIcon.TryGetValue(_propertyKey.ToString(), hi) Then
                    Dim result As Boolean = False
                    If Me.DisplayType = PropertyDisplayType.Enumerated Then
                        Dim propertyEnumTypeList As IPropertyEnumTypeList
                        Try
                            Me.Description.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                            Dim count As UInt32
                            propertyEnumTypeList.GetCount(count)
                            For x As UInt32 = 0 To count - 1
                                Dim propertyEnumType2 As IPropertyEnumType2
                                Try
                                    propertyEnumTypeList.GetAt(x, GetType(IPropertyEnumType2).GUID, propertyEnumType2)
                                    Dim imageReference As String
                                    propertyEnumType2.GetImageReference(imageReference)
                                    If Not String.IsNullOrWhiteSpace(imageReference) Then
                                        result = True
                                        Exit For
                                    End If
                                Finally
                                    If Not propertyEnumType2 Is Nothing Then
                                        Marshal.ReleaseComObject(propertyEnumType2)
                                        propertyEnumType2 = Nothing
                                    End If
                                End Try
                            Next
                        Finally
                            If Not propertyEnumTypeList Is Nothing Then
                                Marshal.ReleaseComObject(propertyEnumTypeList)
                            End If
                        End Try
                    End If
                    _hasIcon.Add(_propertyKey.ToString(), hi)
                End If
            Finally
                _hasIconLock.Release()
            End Try

            Return hi
        End Get
    End Property

    Public Overridable ReadOnly Property HasIconAsync As Boolean
        Get
            Dim tcs As New TaskCompletionSource(Of Boolean)

            Shell.PriorityTaskQueue.Add(
                Sub()
                    Try
                        tcs.SetResult(Me.HasIcon)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            tcs.Task.Wait(Shell.ShuttingDownToken)
            If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                Return tcs.Task.Result
            Else
                Return False
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property ImageReferences16 As String()
        Get
            If Me.DisplayType = PropertyDisplayType.Enumerated Then
                _imageReferences16Lock.Wait()
                Dim i16 As String()
                Try
                    If Not _imageReferences16.TryGetValue(String.Format("{0}_{1}", _propertyKey, Me.RawValue.GetValue()), i16) Then
                        Dim propertyEnumType2 As IPropertyEnumType2
                        Try
                            propertyEnumType2 = getSelectedPropertyEnumType(Me.RawValue, Me.Description)
                            Dim imageReference As String
                            propertyEnumType2.GetImageReference(imageReference)
                            If Not String.IsNullOrWhiteSpace(imageReference) Then
                                i16 = {imageReference}
                            End If
                        Finally
                            If Not propertyEnumType2 Is Nothing Then
                                Marshal.ReleaseComObject(propertyEnumType2)
                            End If
                        End Try
                        _imageReferences16.Add(String.Format("{0}_{1}", _propertyKey, Me.RawValue.GetValue()), i16)
                    End If
                Finally
                    _imageReferences16Lock.Release()
                End Try
                Return i16
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property Icons16 As ImageSource()
        Get
            Return Me.ImageReferences16?.Select(Function(i) ImageHelper.ExtractIcon(i)).ToArray()
        End Get
    End Property

    Public Overridable ReadOnly Property Icons16Async As ImageSource()
        Get
            Dim tcs As New TaskCompletionSource(Of String())

            Shell.PriorityTaskQueue.Add(
                Sub()
                    Try
                        tcs.SetResult(Me.ImageReferences16)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            tcs.Task.Wait(Shell.ShuttingDownToken)
            If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                Dim imageReferences16 As String() = tcs.Task.Result
                If Not imageReferences16 Is Nothing Then
                    Return imageReferences16.Select(Function(i) ImageHelper.ExtractIcon(i)).ToArray()
                End If
            End If
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property FirstIcon16 As ImageSource
        Get
            Dim icons16() As ImageSource = Me.Icons16
            Return If(Not icons16 Is Nothing AndAlso icons16.Count > 0, icons16(0), Nothing)
        End Get
    End Property

    Public Overridable ReadOnly Property FirstIcon16Async As ImageSource
        Get
            Dim icons16() As ImageSource = Me.Icons16Async
            Return If(Not icons16 Is Nothing AndAlso icons16.Count > 0, icons16(0), Nothing)
        End Get
    End Property

    Protected Function getSelectedPropertyEnumType(value As PROPVARIANT, description As IPropertyDescription) As IPropertyEnumType
        If Me.DisplayType = PropertyDisplayType.Enumerated Then
            Dim propertyEnumTypeList As IPropertyEnumTypeList, propertyEnumType As IPropertyEnumType
            Try
                description.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                Dim index As UInt32 = value.GetValue()
                propertyEnumTypeList.GetAt(index, GetType(IPropertyEnumType).GUID, propertyEnumType)
                Return propertyEnumType
            Finally
                If Not propertyEnumTypeList Is Nothing Then
                    Marshal.ReleaseComObject(propertyEnumTypeList)
                End If
            End Try
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
            _rawValue.Dispose()
            disposedValue = True
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
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

﻿Imports System.ComponentModel
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Public Class [Property]
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Private Shared _cachedDescriptionsByCanonicalName As Dictionary(Of String, CachedPropertyDescription) = New Dictionary(Of String, CachedPropertyDescription)
    Private Shared _cachedDescriptionsByPropertyKey As Dictionary(Of String, CachedPropertyDescription) = New Dictionary(Of String, CachedPropertyDescription)
    Private Shared _cachedDescriptionsLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _hasIcon As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)()
    Private Shared _hasIconLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _imageReferences16 As Dictionary(Of String, String()) = New Dictionary(Of String, String())()
    Private Shared _imageReferences16Lock As SemaphoreSlim = New SemaphoreSlim(1, 1)

    Protected _canonicalName As String
    Protected _propertyDescription As IPropertyDescription
    Protected _propertyKey As PROPERTYKEY
    Protected _text As String
    Protected disposedValue As Boolean
    Friend _rawValue As PROPVARIANT
    Protected _displayType As PropertyDisplayType = -1
    Protected _val As Object
    Protected _dataType As VarEnum
    Protected _displayName As String
    Protected _isCustom As Boolean

    Public Shared Function FromCanonicalName(canonicalName As String) As [Property]
        Return makeProperty(Nothing, canonicalName,
            Function(cachedPropertyDescription) New [Property](cachedPropertyDescription),
            Function(type) Activator.CreateInstance(type))
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String, propertyStore As IPropertyStore) As [Property]
        Return makeProperty(Nothing, canonicalName,
            Function(cachedPropertyDescription) New [Property](cachedPropertyDescription, propertyStore),
            Function(type) Activator.CreateInstance(type, {propertyStore}))
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String, shellItem2 As IShellItem2) As [Property]
        Return makeProperty(Nothing, canonicalName,
            Function(cachedPropertyDescription) New [Property](cachedPropertyDescription, shellItem2),
            Function(type) Activator.CreateInstance(type, {shellItem2}))
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional propertyStore As IPropertyStore = Nothing) As [Property]
        Return makeProperty(propertyKey, Nothing,
            Function(cachedPropertyDescription) New [Property](cachedPropertyDescription, propertyStore),
            Function(type) Activator.CreateInstance(type, {propertyStore}))
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional shellItem2 As IShellItem2 = Nothing) As [Property]
        Return makeProperty(propertyKey, Nothing,
            Function(cachedPropertyDescription) New [Property](cachedPropertyDescription, shellItem2),
            Function(type) Activator.CreateInstance(type, {shellItem2}))
    End Function

    Protected Sub New(cachedPropertyDescription As CachedPropertyDescription)
        If Not cachedPropertyDescription Is Nothing Then
            _propertyKey = cachedPropertyDescription.PropertyKey
            _canonicalName = cachedPropertyDescription.CanonicalName
            _propertyDescription = cachedPropertyDescription.PropertyDescription
        End If
    End Sub

    Protected Sub New(cachedPropertyDescription As CachedPropertyDescription, Optional propertyStore As IPropertyStore = Nothing)
        Me.New(cachedPropertyDescription)
        If Not propertyStore Is Nothing Then
            propertyStore.GetValue(_propertyKey, _rawValue)
        End If
    End Sub

    Protected Sub New(cachedPropertyDescription As CachedPropertyDescription, shellItem2 As IShellItem2)
        Me.New(cachedPropertyDescription)
        If Not shellItem2 Is Nothing Then
            shellItem2.GetProperty(_propertyKey, _rawValue)
        End If
    End Sub

    Private Shared Function makeProperty(propertyKey As PROPERTYKEY?, canonicalName As String,
                                         func1 As Func(Of CachedPropertyDescription, [Property]),
                                         func2 As Func(Of Type, [Property])) As [Property]
        'Return Shell.GlobalThreadPool.Run(
        '    Function() As [Property]
        Dim t As Type
                If propertyKey.HasValue Then
                    t = Shell.GetCustomProperty(propertyKey.Value)
                Else
                    t = Shell.GetCustomProperty(canonicalName)
                End If
                If Not t Is Nothing Then
                    Return func2(t)
                Else
                    Dim cachedPropertyDescription As CachedPropertyDescription
                    If propertyKey.HasValue Then
                        cachedPropertyDescription = [Property].GetDescription(propertyKey)
                    Else
                        cachedPropertyDescription = [Property].GetDescription(canonicalName)
                    End If
                    If Not cachedPropertyDescription Is Nothing Then
                        Return func1(cachedPropertyDescription)
                    Else
                        Return Nothing
                    End If
                End If
        'End Function)
    End Function

    Protected Shared Function GetDescription(canonicalName As String) As CachedPropertyDescription
        Dim result As CachedPropertyDescription = Nothing
        Try
            _cachedDescriptionsLock.Wait()
            If Not _cachedDescriptionsByCanonicalName.TryGetValue(canonicalName, result) Then
                _cachedDescriptionsLock.Release()
                Dim propertyDescription As IPropertyDescription = Nothing
                Functions.PSGetPropertyDescriptionByName(canonicalName, GetType(IPropertyDescription).GUID, propertyDescription)
                If Not propertyDescription Is Nothing Then
                    Dim pkey As PROPERTYKEY
                    propertyDescription.GetPropertyKey(pkey)
                    _cachedDescriptionsLock.Wait()
                    If Not _cachedDescriptionsByPropertyKey.ContainsKey(pkey.ToString().ToLower()) Then
                        result = New CachedPropertyDescription() With {
                            .PropertyDescription = propertyDescription,
                            .PropertyKey = pkey,
                            .CanonicalName = canonicalName
                        }
                        Try
                            _cachedDescriptionsByCanonicalName.Add(canonicalName, result)
                            _cachedDescriptionsByPropertyKey.Add(pkey.ToString().ToLower(), result)
                        Catch ex As Exception
                        End Try
                    End If
                    _cachedDescriptionsLock.Release()
                End If
            End If
        Finally
            If _cachedDescriptionsLock.CurrentCount = 0 Then
                _cachedDescriptionsLock.Release()
            End If
        End Try
        Return result
    End Function

    Protected Shared Function GetDescription(pkey As PROPERTYKEY) As CachedPropertyDescription
        Dim result As CachedPropertyDescription = Nothing
        Try
            _cachedDescriptionsLock.Wait()
            If Not _cachedDescriptionsByPropertyKey.TryGetValue(pkey.ToString().ToLower(), result) Then
                _cachedDescriptionsLock.Release()
                Dim propertyDescription As IPropertyDescription = Nothing
                Functions.PSGetPropertyDescription(pkey, GetType(IPropertyDescription).GUID, propertyDescription)
                If Not propertyDescription Is Nothing Then
                    Dim canonicalName As String = Nothing
                    Functions.PSGetNameFromPropertyKey(pkey, canonicalName)
                    _cachedDescriptionsLock.Wait()
                    If Not _cachedDescriptionsByPropertyKey.ContainsKey(pkey.ToString().ToLower()) Then
                        result = New CachedPropertyDescription() With {
                            .PropertyDescription = propertyDescription,
                            .PropertyKey = pkey,
                            .CanonicalName = canonicalName
                        }
                        Try
                            _cachedDescriptionsByCanonicalName.Add(canonicalName, result)
                            _cachedDescriptionsByPropertyKey.Add(pkey.ToString().ToLower(), result)
                        Catch ex As Exception
                        End Try
                    End If
                    _cachedDescriptionsLock.Release()
                End If
            End If
        Finally
            If _cachedDescriptionsLock.CurrentCount = 0 Then
                _cachedDescriptionsLock.Release()
            End If
        End Try
        Return result
    End Function

    Public ReadOnly Property CanonicalName As String
        Get
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
            If String.IsNullOrWhiteSpace(_displayName) Then
                Dim name As StringBuilder = New StringBuilder()
                Me.Description?.GetDisplayName(name)
                _displayName = If(String.IsNullOrWhiteSpace(name.ToString()), Nothing, name.ToString())
            End If
            Return _displayName
        End Get
    End Property

    Public ReadOnly Property DisplayNameWithColon As String
        Get
            Return Me.DisplayName?.ToString() & ":"
        End Get
    End Property

    Public ReadOnly Property Description As IPropertyDescription
        Get
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
            If _val Is Nothing Then
                _val = Me.RawValue.GetValue()
            End If
            Return _val
        End Get
    End Property

    Public Overridable ReadOnly Property Text As String
        Get
            If _text Is Nothing Then
                Dim buffer As StringBuilder = New StringBuilder()
                buffer.Append(New String(" ", 2050))
                Functions.PSFormatForDisplay(_propertyKey, Me.RawValue, PropertyDescriptionFormatOptions.None, buffer, 2048)
                _text = If(String.IsNullOrWhiteSpace(buffer.ToString()), Nothing, buffer.ToString())
            End If
            Return _text
        End Get
    End Property

    Public ReadOnly Property DataType As VarEnum
        Get
            If _dataType = 0 Then
                Me.Description?.GetPropertyType(_dataType)
            End If
            Return _dataType
        End Get
    End Property

    Public Overridable ReadOnly Property GroupByText As String
        Get
            If Me.DataType = VarEnum.VT_FILETIME Then
                Dim dt As DateTime = Me.Value

                Dim thisWeek As Integer, thisYear As Integer
                Dim lastWeek As Integer, lastYear As Integer
                Dim fileWeek As Integer, fileYear As Integer
                getISO8601WeekOfYear(DateTime.Now.Date, thisWeek, thisYear)
                getISO8601WeekOfYear(DateTime.Now.Date.AddDays(-7), lastWeek, lastYear)
                getISO8601WeekOfYear(dt, fileWeek, fileYear)
                Dim lastMonth As DateTime = DateTime.Now.Date.AddMonths(-1)

                If dt.Date = DateTime.Now.Date Then
                    Return My.Resources.Property_Date_Today
                ElseIf dt.Date = DateTime.Now.Date.AddDays(-1) Then
                    Return My.Resources.Property_Date_Yesterday
                ElseIf thisWeek = fileWeek AndAlso thisYear = fileYear Then
                    Return My.Resources.Property_Date_EarlierThisWeek
                ElseIf lastWeek = fileWeek And lastYear = fileYear Then
                    Return My.Resources.Property_Date_LastWeek
                ElseIf dt.Month = DateTime.Now.Month AndAlso dt.Year = DateTime.Now.Year Then
                    Return My.Resources.Property_Date_EarlierThisMonth
                ElseIf dt.Month = lastMonth.Month AndAlso dt.Year = lastMonth.Year Then
                    Return My.Resources.Property_Date_LastMonth
                Else
                    Return My.Resources.Property_Date_LongAgo
                End If
            ElseIf Me.DataType = VarEnum.VT_UI8 Then
                If Me.RawValue.vt = 0 Then
                    Return My.Resources.Property_Size_Unknown 
                ElseIf Me.Value <= 16 * 1024 Then
                    Return My.Resources.Property_Size_VerySmall
                ElseIf Me.Value <= 1024 * 1024 Then
                    Return My.Resources.Property_Size_Small
                ElseIf Me.Value <= 128 * 1024 * 1024 Then
                    Return My.Resources.Property_Size_Normal
                ElseIf Me.Value <= 1024 * 1024 * 1024 Then
                    Return My.Resources.Property_Size_Large
                ElseIf Me.Value <= 4L * 1024 * 1024 * 1024 Then
                    Return My.Resources.Property_Size_VeryLarge
                Else
                    Return My.Resources.Property_Size_Gigantic
                End If
            Else
                Return If(String.IsNullOrWhiteSpace(Me.Text), Nothing, Me.Text)
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayType As PropertyDisplayType
        Get
            If _displayType = -1 Then
                Me.Description?.GetDisplayType(_displayType)
            End If
            Return _displayType
        End Get
    End Property

    Public ReadOnly Property RelativeDescriptionType As RelativeDescriptionType
        Get
            Dim rdt As RelativeDescriptionType
            Me.Description?.GetRelativeDescriptionType(rdt)
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
                        Dim propertyEnumTypeList As IPropertyEnumTypeList = Nothing
                        Try
                            Me.Description?.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                            If propertyEnumTypeList Is Nothing Then Return False
                            Dim count As UInt32
                            propertyEnumTypeList.GetCount(count)
                            For x As UInt32 = 0 To count - 1
                                Dim propertyEnumType2 As IPropertyEnumType2 = Nothing
                                Try
                                    propertyEnumTypeList.GetAt(x, GetType(IPropertyEnumType2).GUID, propertyEnumType2)
                                    Dim imageReference As String = Nothing
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
                                propertyEnumTypeList = Nothing
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
            Return Shell.GlobalThreadPool.Run(
                Function() As Boolean
                    Return Me.HasIcon
                End Function)
        End Get
    End Property

    Public Overridable ReadOnly Property ImageReferences16 As String()
        Get
            If Me.DisplayType = PropertyDisplayType.Enumerated Then
                _imageReferences16Lock.Wait()
                Dim i16 As String() = Nothing
                Try
                    If Not _imageReferences16.TryGetValue(String.Format("{0}_{1}", _propertyKey, Me.RawValue.GetValue()), i16) Then
                        Dim propertyEnumType2 As IPropertyEnumType2 = Nothing
                        Try
                            propertyEnumType2 = getSelectedPropertyEnumType(Me.RawValue, Me.Description)
                            Dim imageReference As String = Nothing
                            propertyEnumType2.GetImageReference(imageReference)
                            If Not String.IsNullOrWhiteSpace(imageReference) Then
                                i16 = {imageReference}
                            End If
                        Finally
                            If Not propertyEnumType2 Is Nothing Then
                                Marshal.ReleaseComObject(propertyEnumType2)
                                propertyEnumType2 = Nothing
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

    Public Overridable ReadOnly Property Icons16 As BitmapSource()
        Get
            Return Me.ImageReferences16?.Select(Function(i) ImageHelper.ExtractIcon(i, True)).ToArray()
        End Get
    End Property

    Public Overridable ReadOnly Property Icons16Async As BitmapSource()
        Get
            Dim imageReferences16 As String() = Shell.GlobalThreadPool.Run(
                Function() As String()
                    Return Me.ImageReferences16
                End Function)

            If Not imageReferences16 Is Nothing Then
                Return imageReferences16.Select(Function(i) ImageHelper.ExtractIcon(i, True)).ToArray()
            End If
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property FirstIcon16 As BitmapSource
        Get
            Dim icons16() As BitmapSource = Me.Icons16
            Return If(Not icons16 Is Nothing AndAlso icons16.Count > 0, icons16(0), Nothing)
        End Get
    End Property

    Public Overridable ReadOnly Property FirstIcon16Async As BitmapSource
        Get
            Dim icons16() As BitmapSource = Me.Icons16Async
            Return If(Not icons16 Is Nothing AndAlso icons16.Count > 0, icons16(0), Nothing)
        End Get
    End Property

    Public Property IsCustom As Boolean
        Get
            Return _isCustom
        End Get
        Set(value As Boolean)
            _isCustom = value
        End Set
    End Property

    Protected Function getSelectedPropertyEnumType(value As PROPVARIANT, description As IPropertyDescription) As IPropertyEnumType
        If Me.DisplayType = PropertyDisplayType.Enumerated Then
            Dim propertyEnumTypeList As IPropertyEnumTypeList = Nothing, propertyEnumType As IPropertyEnumType = Nothing
            Try
                description.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, propertyEnumTypeList)
                Dim index As UInt32 = value.GetValue()
                propertyEnumTypeList.GetAt(index, GetType(IPropertyEnumType).GUID, propertyEnumType)
                Return propertyEnumType
            Finally
                If Not propertyEnumTypeList Is Nothing Then
                    Marshal.ReleaseComObject(propertyEnumTypeList)
                    propertyEnumTypeList = Nothing
                End If
            End Try
        End If

        Return Nothing
    End Function

    ' This presumes that weeks start with Monday.
    ' Week 1 Is the 1st week of the year with a Thursday in it.
    Private Sub getISO8601WeekOfYear(dt As DateTime, ByRef weekNumber As Integer, ByRef year As Integer)
        ' Seriously cheat.  If its Monday, Tuesday Or Wednesday, then it'll 
        ' be the same week# as whatever Thursday, Friday Or Saturday are,
        ' And we always get those right
        Dim day As DayOfWeek = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(dt)
        If (day >= DayOfWeek.Monday AndAlso day <= DayOfWeek.Wednesday) Then
            dt = dt.AddDays(3)
        End If

        ' Return the week of our adjusted day
        weekNumber = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
        If weekNumber = 1 AndAlso dt.Month = 12 Then
            year = dt.Year + 1
        ElseIf weekNumber > 45 AndAlso dt.Month = 1 Then
            year = dt.Year - 1
        Else
            year = dt.Year
        End If
    End Sub

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

    Protected Class CachedPropertyDescription
        Public Property PropertyDescription As IPropertyDescription
        Public Property PropertyKey As PROPERTYKEY
        Public Property CanonicalName As String
    End Class
End Class

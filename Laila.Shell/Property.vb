Imports System.ComponentModel
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
    Friend _rawValue As PROPVARIANT
    Protected _displayType As PropertyDisplayType = -1
    Protected _val As Object
    Protected _dataType As VarEnum
    Protected _displayName As String
    Protected _isCustom As Boolean

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
        Dim t As Type = Shell.GetCustomProperty(canonicalName)
        If Not t Is Nothing Then
            Return Activator.CreateInstance(t)
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
        Dim t As Type = Shell.GetCustomProperty(canonicalName)
        If Not t Is Nothing Then
            Return Activator.CreateInstance(t, {propertyStore})
        Else
            Return New [Property](canonicalName, propertyStore)
        End If
    End Function

    Public Shared Function FromCanonicalName(canonicalName As String, shellItem2 As IShellItem2) As [Property]
        Dim t As Type = Shell.GetCustomProperty(canonicalName)
        If Not t Is Nothing Then
            Return Activator.CreateInstance(t, {shellItem2})
        Else
            Return New [Property](canonicalName, shellItem2)
        End If
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional propertyStore As IPropertyStore = Nothing) As [Property]
        Dim t As Type = Shell.GetCustomProperty(propertyKey)
        If Not t Is Nothing Then
            Return Activator.CreateInstance(t, {propertyStore})
        Else
            Return New [Property](propertyKey, propertyStore)
        End If
    End Function

    Public Shared Function FromKey(propertyKey As PROPERTYKEY, Optional shellItem2 As IShellItem2 = Nothing) As [Property]
        Dim t As Type = Shell.GetCustomProperty(propertyKey)
        If Not t Is Nothing Then
            Return Activator.CreateInstance(t, {shellItem2})
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
            If String.IsNullOrWhiteSpace(_displayName) Then
                Dim name As StringBuilder = New StringBuilder()
                Me.Description.GetDisplayName(name)
                _displayName = name.ToString()
            End If
            Return _displayName
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
                _text = buffer.ToString()
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
                    Return "Today"
                ElseIf dt.Date = DateTime.Now.Date.AddDays(-1) Then
                    Return "Yesterday"
                ElseIf thisWeek = fileWeek AndAlso thisYear = fileYear Then
                    Return "Earlier this week"
                ElseIf lastWeek = fileWeek And lastYear = fileYear Then
                    Return "Last week"
                ElseIf dt.Month = DateTime.Now.Month AndAlso dt.Year = DateTime.Now.Year Then
                    Return "Earlier this month"
                ElseIf dt.Month = lastMonth.Month AndAlso dt.Year = lastMonth.Year Then
                    Return "Last month"
                Else
                    Return "Long ago"
                End If
            ElseIf Me.DataType = VarEnum.VT_UI8 Then
                If Me.RawValue.vt = 0 Then
                    Return "Unknown"
                ElseIf Me.Value <= 16 * 1024 Then
                    Return "Very small (0 - 16 KB)"
                ElseIf Me.Value <= 1024 * 1024 Then
                    Return "Small (16 KB - 1 MB)"
                ElseIf Me.Value <= 128 * 1024 * 1024 Then
                    Return "Normal (1 - 128 MB)"
                ElseIf Me.Value <= 1024 * 1024 * 1024 Then
                    Return "Large (128 MB - 1 GB)"
                ElseIf Me.Value <= 4L * 1024 * 1024 * 1024 Then
                    Return "Very large (1 - 4 GB)"
                Else
                    Return "Gigantic (> 4 GB)"
                End If
            Else
                Return Me.Text
            End If
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
            Shell.RunOnSTAThread(
                Function() As Boolean
                    Return Me.HasIcon
                End Function)
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

    Public Overridable ReadOnly Property Icons16 As ImageSource()
        Get
            Return Me.ImageReferences16?.Select(Function(i) ImageHelper.ExtractIcon(i)).ToArray()
        End Get
    End Property

    Public Overridable ReadOnly Property Icons16Async As ImageSource()
        Get
            Dim imageReferences16 As String() = Shell.RunOnSTAThread(
                Function() As String()
                    Return Me.ImageReferences16
                End Function)

            If Not imageReferences16 Is Nothing Then
                Return imageReferences16.Select(Function(i) ImageHelper.ExtractIcon(i)).ToArray()
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

    Public ReadOnly Property IsCustom As Boolean
        Get
            Return _isCustom
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
End Class

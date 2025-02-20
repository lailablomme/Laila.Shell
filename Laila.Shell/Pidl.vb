Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Laila.Shell.Helpers

Public Class Pidl
    Implements IDisposable

    Private _pidl As IntPtr
    Private _lastId As IntPtr
    Private disposedValue As Boolean

    Public Shared Function FromCustomBytes(bytes As Byte()) As Pidl
        ' Compute total size: 2 bytes (size) + data length + 2 bytes (NULL terminator)
        Dim cb As Integer = 2 + bytes.Length + 2

        ' Allocate memory for PIDL using CoTaskMemAlloc
        Dim pidlPtr As IntPtr = Marshal.AllocCoTaskMem(cb)

        If pidlPtr = IntPtr.Zero Then
            Throw New OutOfMemoryException("Failed to allocate memory for PIDL")
        End If

        ' Write cb (total size including NULL terminator)
        Marshal.WriteInt16(pidlPtr, CType(cb, Short))

        ' Copy data into the PIDL
        Marshal.Copy(bytes, 0, IntPtr.Add(pidlPtr, 2), bytes.Length)

        ' Write the NULL terminator (final 2 bytes)
        Marshal.WriteInt16(IntPtr.Add(pidlPtr, 2 + bytes.Length), 0)

        Return New Pidl(pidlPtr)
    End Function

    Public Sub New(pidl As IntPtr)
        _pidl = pidl
        If IntPtr.Zero.Equals(pidl) Then
            Throw New InvalidOperationException("Pidl can't be zero")
        End If
    End Sub

    Public Sub New(pidl As String)
        Dim bytes As Byte() = pidl.Split("-"c).Select(Function(hex) Convert.ToByte(hex, 16)).ToArray()
        _pidl = Marshal.AllocCoTaskMem(bytes.Length)
        Marshal.Copy(bytes, 0, _pidl, bytes.Length)
    End Sub

    Public Sub New(bytes As Byte())
        _pidl = Marshal.AllocCoTaskMem(bytes.Length)
        Marshal.Copy(bytes, 0, _pidl, bytes.Length)
    End Sub

    Public Shared Function CreateShellIDListArray(items As IEnumerable(Of Item)) As IntPtr
        Dim pidls As List(Of Pidl) = New List(Of Pidl)()
        For Each item In items
            pidls.Add(item.Pidl)
        Next

        Dim mem As MemoryStream = New MemoryStream()
        Dim writer As New System.IO.BinaryWriter(mem)

        ' write count
        writer.Write(Convert.ToUInt32(pidls.Count))

        ' write offsets
        Dim offset As Integer
        offset = Marshal.SizeOf(GetType(UInt32)) * (pidls.Count + 2)
        writer.Write(Convert.ToUInt32(offset)) ' count + parent pidl + number of pidls
        offset += 2 ' skip parent pidl, all pidls must be absolute
        For i = 0 To pidls.Count - 1
            writer.Write(Convert.ToUInt32(offset))
            offset += pidls(i).Bytes.Length
        Next

        ' write bytes
        writer.Write(Convert.ToUInt16(0)) ' parent
        For i = 0 To pidls.Count - 1
            writer.Write(pidls(i).Bytes)
        Next

        ' get pointer to data
        mem.Seek(0, SeekOrigin.Begin)
        Dim bytes() As Byte = mem.GetBuffer()
        Dim hGlobal As IntPtr = Marshal.AllocHGlobal(bytes.Length)
        Marshal.Copy(bytes, 0, hGlobal, bytes.Length)
        mem.Dispose()

        Return hGlobal
    End Function

    Public Shared Function GetItemsFromShellIDListArray(ptr As IntPtr) As List(Of Item)
        Dim result As List(Of Item) = New List(Of Item)()
        Dim start As IntPtr = ptr

        ' read count
        Dim count As UInt32 = Convert.ToUInt32(Marshal.ReadInt32(ptr)) : ptr = IntPtr.Add(ptr, Marshal.SizeOf(Of UInt32))
        If count > 1000 Then Return Nothing

        ' read parent
        Dim offset As UInt32 = Convert.ToUInt32(Marshal.ReadInt32(ptr)) : ptr = IntPtr.Add(ptr, Marshal.SizeOf(Of UInt32)) ' parent
        Dim parentShellFolder As IShellFolder
        Dim parentFolder As Folder = Nothing

        ' read items
        Shell.GlobalThreadPool.Run(
            Sub()
                If Convert.ToUInt16(Marshal.ReadInt16(IntPtr.Add(start, offset))) <> 0 Then
                    Dim parentPidl As Pidl = New Pidl(IntPtr.Add(start, offset))
                    parentFolder = CType(Item.FromPidl(parentPidl, Nothing, True), Folder)
                    parentShellFolder = parentFolder.ShellFolder
                End If

                For i = 0 To count - 1
                    offset = Convert.ToUInt32(Marshal.ReadInt32(ptr)) : ptr = IntPtr.Add(ptr, Marshal.SizeOf(Of UInt32))
                    Dim itemPidl As Pidl = New Pidl(IntPtr.Add(start, offset))
                    result.Add(Item.FromPidl(itemPidl, parentFolder, True, True))
                Next
            End Sub, 1)

        Return result
    End Function

    Public ReadOnly Property ItemIDListSize As Integer
        Get
            Dim i As Integer = ItemIDSize
            Dim b As Integer = Marshal.ReadByte(_pidl, i) + (Marshal.ReadByte(_pidl, i + 1) * 256)
            Do While b > 0
                i += b
                b = Marshal.ReadByte(_pidl, i) + (Marshal.ReadByte(_pidl, i + 1) * 256)
            Loop
            Return i
        End Get
    End Property

    Public ReadOnly Property ItemIDSize As Integer
        Get
            Dim b(1) As Byte
            Marshal.Copy(_pidl, b, 0, 2)
            Return b(1) * 256 + b(0)
        End Get
    End Property

    Public ReadOnly Property Bytes As Byte()
        Get
            Dim result() As Byte
            Dim iilSize As Integer = ItemIDListSize
            If iilSize > 0 Then
                ReDim result(iilSize + 1)
                Marshal.Copy(_pidl, result, 0, iilSize)
            Else
                ReDim result(1)
            End If
            result(result.Length - 2) = 0
            result(result.Length - 1) = 0
            Return result
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Dim result As Integer = 0
            Dim i As Integer = 0
            Dim b As Integer = Marshal.ReadByte(_pidl, i) + (Marshal.ReadByte(_pidl, i + 1) * 256)
            Do While b > 0
                result += 1
                i += b
                b = Marshal.ReadByte(_pidl, i) + (Marshal.ReadByte(_pidl, i + 1) * 256)
            Loop
            Return result
        End Get
    End Property

    Public ReadOnly Property AbsolutePIDL As IntPtr
        Get
            Return _pidl
        End Get
    End Property

    Public ReadOnly Property RelativePIDL As IntPtr
        Get
            If IntPtr.Zero.Equals(_lastId) Then
                _lastId = Functions.ILFindLastID(_pidl)
            End If
            Return _lastId
        End Get
    End Property

    Public ReadOnly Property DekstopRelativePIDL As IntPtr
        Get
            Return Functions.ILFindChild(Shell.Desktop.Pidl.AbsolutePIDL, _pidl)
        End Get
    End Property

    Public Shadows Function Equals(pidl As Pidl) As Boolean
        If pidl Is Nothing Then Return False
        Return Functions.ILIsEqual(Me.AbsolutePIDL, pidl.AbsolutePIDL)
    End Function

    Public Overrides Function ToString() As String
        Return BitConverter.ToString(Me.Bytes)
    End Function

    Public Function Clone() As Pidl
        Return New Pidl(Functions.ILClone(Me.AbsolutePIDL))
    End Function

    Public Function GetParent() As Pidl
        Dim parentPidl As IntPtr = Functions.ILClone(Me.AbsolutePIDL)
        Functions.ILRemoveLastID(parentPidl)
        Return New Pidl(parentPidl)
    End Function

    Public ReadOnly Property IsZero As Boolean
        Get
            Return IntPtr.Zero.Equals(Me.AbsolutePIDL)
        End Get
    End Property

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            disposedValue = True

            If disposing Then
                ' dispose managed state (managed objects)
            End If

            ' free unmanaged resources (unmanaged objects) and override finalizer
            If Not IntPtr.Zero.Equals(_pidl) Then
                Marshal.FreeCoTaskMem(_pidl)
                _pidl = IntPtr.Zero
            End If
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '    Dispose(disposing:=False)
    '    MyBase.Finalize()
    'End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

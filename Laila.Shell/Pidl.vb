Imports System.IO
Imports System.Runtime.InteropServices

Public Class Pidl
    Private _pidl As IntPtr

    Public Sub New(pidl As IntPtr)
        _pidl = pidl
    End Sub

    Public Shared Function CreateShellIDListArray(items As IEnumerable(Of Item)) As IntPtr
        Dim pidls As List(Of Pidl) = New List(Of Pidl)()
        For Each item In items
            Dim pidl As IntPtr = IntPtr.Zero
            Dim punk As IntPtr = Marshal.GetIUnknownForObject(item.ShellItem2)
            Functions.SHGetIDListFromObject(punk, pidl)
            pidls.Add(New Pidl(pidl))
            Marshal.Release(punk)
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
        If Convert.ToUInt16(Marshal.ReadInt16(IntPtr.Add(start, offset))) = 0 Then
            parentShellFolder = Nothing
        Else
            parentShellFolder = New Folder(Item.GetIShellItem2FromPidl(IntPtr.Add(start, offset), Nothing), Nothing).ShellFolder
        End If

        ' read items
        Try
            For i = 0 To count - 1
                offset = Convert.ToUInt32(Marshal.ReadInt32(ptr)) : ptr = IntPtr.Add(ptr, Marshal.SizeOf(Of UInt32))
                Dim shellItem2 As IShellItem2 = Item.GetIShellItem2FromPidl(IntPtr.Add(start, offset), parentShellFolder)
                Try
                    result.Add(Item.FromParsingName(Item.GetFullPathFromShellItem2(shellItem2), Nothing))
                Finally
                    If Not shellItem2 Is Nothing Then
                        Marshal.ReleaseComObject(shellItem2)
                    End If
                End Try
            Next
        Finally
            If Not parentShellFolder Is Nothing Then
                Marshal.ReleaseComObject(parentShellFolder)
            End If
        End Try

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
End Class

Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Windows

Public Class OurDataObject
    Implements System.Windows.IDataObject

    <DllImport("kernel32.dll", ExactSpelling:=True)>
    Shared Function GlobalLock(ByVal handle As IntPtr) As IntPtr
    End Function

    <DllImport("ole32.dll")>
    Public Shared Function CreateILockBytesOnHGlobal(ByVal hGlobal As IntPtr,
                          ByVal fDeleteOnRelease As Boolean,
                          <Out()> ByRef ppLkbyt As ILockBytes) _
                          As Integer
    End Function

    <DllImport("OLE32.DLL", CharSet:=CharSet.Auto, PreserveSig:=False)>
    Public Shared Function GetHGlobalFromILockBytes(pLockBytes As ILockBytes) As IntPtr
    End Function

    <DllImport("OLE32.DLL", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function StgCreateDocfileOnILockBytes(plkbyt As ILockBytes, grfMode As UInteger, reserved As UInteger) As IStorage
    End Function

    <ComImport(), Guid("0000000b-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IStorage
        Sub CreateStream(ByVal pwcsName As String, ByVal grfMode As UInteger, ByVal reserved1 As UInteger, ByVal reserved2 As UInteger, ByRef ppstm As IStream)
        Sub OpenStream(ByVal pwcsName As String, ByVal reserved1 As IntPtr, ByVal grfMode As UInteger, ByVal reserved2 As UInteger, ByRef ppstm As IStream)
        Sub CreateStorage(ByVal pwcsName As String, ByVal grfMode As UInteger, ByVal reserved1 As UInteger, ByVal reserved2 As UInteger, ByRef ppstg As IStorage)
        Sub OpenStorage(ByVal pwcsName As String, ByVal pstgPriority As IStorage, ByVal grfMode As UInteger, ByVal snbExclude As IntPtr, ByVal reserved As UInteger, ByRef ppstg As IStorage)
        Sub CopyTo(ByVal ciidExclude As UInteger, ByVal rgiidExclude() As Guid, ByVal snbExclude As IntPtr, ByVal pstgDest As IStorage)
        Sub MoveElementTo(ByVal pwcsName As String,
                  ByVal pstgDest As IStorage,
                  ByVal pwcsNewName As String,
                  ByVal grfFlags As UInteger)
        Sub Commit(ByVal grfCommitFlags As UInteger)
        Sub Revert()
        Sub EnumElements(ByVal reserved1 As UInteger, ByVal reserved2 As IntPtr, ByVal reserved3 As UInteger, ByRef ppenum As Object)
        Sub DestroyElement(ByVal pwcsName As String)
        Sub RenameElement(ByVal pwcsOldName As String, ByVal pwcsNewName As String)
        Sub SetElementTimes(ByVal pwcsName As String, ByVal pctime As ComTypes.FILETIME, ByVal patime As ComTypes.FILETIME, ByVal pmtime As ComTypes.FILETIME)
        Sub SetClass(ByVal clsid As Guid)
        Sub SetStateBits(ByVal grfStateBits As UInteger, ByVal grfMask As UInteger)
        Sub Stat(ByRef pstatstg As ComTypes.STATSTG, ByVal grfStatFlag As UInteger)
    End Interface

    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000a-0000-0000-C000-000000000046")>
    Public Interface ILockBytes
        Sub ReadAt(ByVal ulOffset As Long, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal pv As Byte(), ByVal cv As Integer, <Out()> ByRef pcbRead As Integer)
        Sub WriteAt(ByVal ulOffset As Long, ByVal pv As IntPtr, ByVal cb As Integer, <Out()> ByRef pcbWritten As Integer)
        Sub Flush()
        Sub SetSize(ByVal cb As Long)
        Sub LockRegion(ByVal libOffset As Long, ByVal cb As Long, ByVal dwLockType As Integer)
        Sub UnlockRegion(ByVal libOffset As Long, ByVal cb As Long, ByVal dwLockType As Integer)
        Sub Stat(ByRef pstatstg As ComTypes.STATSTG, ByVal grfStatFlag As Integer)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure POINTL
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SIZEL
        Public cx As Integer
        Public cy As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure FILEGROUPDESCRIPTORA
        Public cItems As UInteger
        Public fgd As FILEDESCRIPTORA
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure FILEDESCRIPTORA
        Public dwFlags As UInteger
        Public clsid As Guid
        Public sizel As SIZEL
        Public pointl As POINTL
        Public dwFileAttributes As UInteger
        Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure FILEGROUPDESCRIPTORW
        Public cItems As UInteger
        Public fgd As FILEDESCRIPTORW
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure FILEDESCRIPTORW
        Public dwFlags As UInteger
        Public clsid As Guid
        Public sizel As SIZEL
        Public pointl As POINTL
        Public dwFileAttributes As UInteger
        Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
    End Structure

    Private underlyingDataObject As System.Windows.IDataObject
    Private comUnderlyingDataObject As System.Runtime.InteropServices.ComTypes.IDataObject
    Private oleUnderlyingDataObject As System.Windows.IDataObject
    Private getDataFromHGLOBALMethod As MethodInfo

    Public Sub New(underlyingDataObject As System.Windows.IDataObject)
        ' get the underlying dataobject And its ComType IDataObject interface to it
        Me.underlyingDataObject = underlyingDataObject
        Me.comUnderlyingDataObject = CType(Me.underlyingDataObject, System.Runtime.InteropServices.ComTypes.IDataObject)

        ' get the internal ole dataobject And its GetDataFromHGLOBAL so it can be called later
        Dim innerDataField As FieldInfo = Me.underlyingDataObject.GetType().GetField("_innerData", BindingFlags.NonPublic Or BindingFlags.Instance)
        Me.oleUnderlyingDataObject = CType(innerDataField.GetValue(Me.underlyingDataObject), System.Windows.IDataObject)
        Me.getDataFromHGLOBALMethod = Me.oleUnderlyingDataObject.GetType().GetMethod("GetDataFromHGLOBAL", BindingFlags.NonPublic Or BindingFlags.Instance)

    End Sub

    Public Sub SetData(data As Object) Implements System.Windows.IDataObject.SetData
        Me.underlyingDataObject.SetData(data)
    End Sub

    Public Sub SetData(format As String, data As Object) Implements System.Windows.IDataObject.SetData
        Me.underlyingDataObject.SetData(format, data)
    End Sub

    Public Sub SetData(format As Type, data As Object) Implements System.Windows.IDataObject.SetData
        Me.underlyingDataObject.SetData(format, data)
    End Sub

    Public Sub SetData(format As String, data As Object, autoConvert As Boolean) Implements System.Windows.IDataObject.SetData
        Me.underlyingDataObject.SetData(format, data, autoConvert)
    End Sub

    Public Function GetData(format As String) As Object Implements System.Windows.IDataObject.GetData
        Return Me.GetData(format, True)
    End Function

    Public Function GetData(format As Type) As Object Implements System.Windows.IDataObject.GetData
        Return Me.GetData(format.FullName)
    End Function

    Public Function GetData(format As String, autoConvert As Boolean) As Object Implements System.Windows.IDataObject.GetData
        ' handle the "FileGroupDescriptor" And "FileContents" format request in this class otherwise pass through to underlying IDataObject 
        Select Case format
            Case "FileGroupDescriptor"
                ' override the default handling of FileGroupDescriptor which returns a
                ' MemoryStream And instead return a string array of file names
                Dim fileGroupDescriptorAPointer As IntPtr = IntPtr.Zero
                Try
                    ' use the underlying IDataObject to get the FileGroupDescriptor as a MemoryStream
                    Dim fileGroupDescriptorStream As MemoryStream = CType(Me.underlyingDataObject.GetData("FileGroupDescriptor", autoConvert), MemoryStream)
                    Dim fileGroupDescriptorBytes(fileGroupDescriptorStream.Length) As Byte
                    fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                    fileGroupDescriptorStream.Close()

                    ' copy the file group descriptor into unmanaged memory 
                    fileGroupDescriptorAPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                    Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorAPointer, fileGroupDescriptorBytes.Length)

                    ' marshal the unmanaged memory to to FILEGROUPDESCRIPTORA struct
                    ' FIX FROM - https://stackoverflow.com/questions/27173844/accessviolationexception-after-copying-a-file-from-inside-a-zip-archive-to-the-c
                    Dim ITEMCOUNT As Integer = Marshal.ReadInt32(fileGroupDescriptorAPointer)

                    ' create a New array to store file names in of the number of items in the file group descriptor
                    Dim fileNames(ITEMCOUNT - 1) As String

                    ' get the pointer to the first file descriptor
                    Dim fileDescriptorPointer As IntPtr = CType((CType(fileGroupDescriptorAPointer, Long) + Marshal.SizeOf(ITEMCOUNT)), IntPtr)

                    ' loop for the number of files acording to the file group descriptor
                    For fileDescriptorIndex = 0 To ITEMCOUNT - 1
                        ' marshal the pointer top the file descriptor as a FILEDESCRIPTORA struct And get the file name
                        Dim fileDescriptor As FILEDESCRIPTORA = CType(Marshal.PtrToStructure(fileDescriptorPointer, GetType(FILEDESCRIPTORA)), FILEDESCRIPTORA)
                        fileNames(fileDescriptorIndex) = fileDescriptor.cFileName

                        ' move the file descriptor pointer to the next file descriptor
                        fileDescriptorPointer = CType((CType(fileDescriptorPointer, Long) + Marshal.SizeOf(fileDescriptor)), IntPtr)
                    Next

                    ' return the array of filenames
                    Return fileNames
                Finally
                    ' free unmanaged memory pointer
                    Marshal.FreeHGlobal(fileGroupDescriptorAPointer)
                End Try

            Case "FileGroupDescriptorW"
                ' override the default handling of FileGroupDescriptorW which returns a
                ' MemoryStream And instead return a string array of file names
                Dim fileGroupDescriptorWPointer As IntPtr = IntPtr.Zero
                Try
                    ' use the underlying IDataObject to get the FileGroupDescriptorW as a MemoryStream
                    Dim fileGroupDescriptorStream As MemoryStream = CType(Me.underlyingDataObject.GetData("FileGroupDescriptorW"), MemoryStream)
                    Dim fileGroupDescriptorBytes(fileGroupDescriptorStream.Length) As Byte
                    fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                    fileGroupDescriptorStream.Close()

                    ' copy the file group descriptor into unmanaged memory
                    fileGroupDescriptorWPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                    Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorWPointer, fileGroupDescriptorBytes.Length)

                    ' marshal the unmanaged memory to to FILEGROUPDESCRIPTORW struct
                    ' FIX FROM - https://stackoverflow.com/questions/27173844/accessviolationexception-after-copying-a-file-from-inside-a-zip-archive-to-the-c
                    Dim ITEMCOUNT As Integer = Marshal.ReadInt32(fileGroupDescriptorWPointer)

                    ' create a New array to store file names in of the number of items in the file group descriptor
                    Dim fileNames(ITEMCOUNT - 1) As String

                    ' get the pointer to the first file descriptor
                    Dim fileDescriptorPointer As IntPtr = CType((CType(fileGroupDescriptorWPointer, Long) + Marshal.SizeOf(ITEMCOUNT)), IntPtr)

                    ' loop for the number of files acording to the file group descriptor
                    For fileDescriptorIndex = 0 To ITEMCOUNT - 1
                        ' marshal the pointer top the file descriptor as a FILEDESCRIPTORW struct And get the file name
                        Dim fileDescriptor As FILEDESCRIPTORW = CType(Marshal.PtrToStructure(fileDescriptorPointer, GetType(FILEDESCRIPTORW)), FILEDESCRIPTORW)
                        fileNames(fileDescriptorIndex) = fileDescriptor.cFileName

                        ' move the file descriptor pointer to the next file descriptor
                        fileDescriptorPointer = CType((CType(fileDescriptorPointer, Long) + Marshal.SizeOf(fileDescriptor)), IntPtr)
                    Next

                    ' return the array of filenames
                    Return fileNames
                Finally
                    ' free unmanaged memory pointer
                    Marshal.FreeHGlobal(fileGroupDescriptorWPointer)
                End Try

            Case "FileContents"
                ' override the default handling of FileContents which returns the
                ' contents of the first file as a memory stream And instead return                    
                ' a array of MemoryStreams containing the data to each file dropped                    
                ' 
                '  FILECONTENTS requires a companion FILEGROUPDESCRIPTOR to be                     
                '  available so we bail out if we don't find one in the data object.

                Dim fgdFormatName As String
                If GetDataPresent("FileGroupDescriptorW") Then
                    fgdFormatName = "FileGroupDescriptorW"
                ElseIf (GetDataPresent("FileGroupDescriptor")) Then
                    fgdFormatName = "FileGroupDescriptor"
                Else
                    Return Nothing
                End If
                ' get the array of filenames which lets us know how many file contents exist                    
                Dim fileContentNames() As String = CType(Me.GetData(fgdFormatName), String())

                ' create a MemoryStream array to store the file contents
                Dim fileContents(fileContentNames.Length - 1) As MemoryStream

                ' loop for the number of files acording to the file names
                For fileIndex = 0 To fileContentNames.Length - 1
                    ' get the data at the file index And store in array
                    fileContents(fileIndex) = Me.GetData(format, fileIndex)
                Next

                ' return array of MemoryStreams containing file contents
                Return fileContents
        End Select

        ' use underlying IDataObject to handle getting of data
        Return Me.underlyingDataObject.GetData(format, autoConvert)
    End Function

    Public Function GetData(format As String, index As Integer) As MemoryStream
        ' create a FORMATETC struct to request the data with
        Dim FORMATETC As FORMATETC = New FORMATETC()
        Debug.WriteLine("FORMATETC.cfFormat = " & DataFormats.GetDataFormat(format).Id)
        Dim id As Integer = DataFormats.GetDataFormat(format).Id
        If id > Short.MaxValue Then id -= 65536
        FORMATETC.cfFormat = CType(id, Short)
        FORMATETC.dwAspect = DVASPECT.DVASPECT_CONTENT
        FORMATETC.lindex = index
        FORMATETC.ptd = New IntPtr(0)
        FORMATETC.tymed = TYMED.TYMED_ISTREAM Or TYMED.TYMED_ISTORAGE Or TYMED.TYMED_HGLOBAL

        ' create STGMEDIUM to output request results into
        Dim medium As STGMEDIUM = New STGMEDIUM()

        ' using the Com IDataObject interface get the data using the defined FORMATETC
        Me.comUnderlyingDataObject.GetData(FORMATETC, medium)

        ' retrieve the data depending on the returned store type
        Select Case medium.tymed
            Case TYMED.TYMED_ISTORAGE
                ' to handle a IStorage it needs to be written into a second unmanaged
                ' memory mapped storage And then the data can be read from memory into
                ' a managed byte And returned as a MemoryStream

                Dim IStorage As IStorage = Nothing
                Dim iStorage2 As IStorage = Nothing
                Dim ILockBytes As ILockBytes = Nothing
                Dim iLockBytesStat As System.Runtime.InteropServices.ComTypes.STATSTG
                Try
                    ' marshal the returned pointer to a IStorage object
                    IStorage = CType(Marshal.GetObjectForIUnknown(medium.unionmember), IStorage)
                    Marshal.Release(medium.unionmember)

                    ' create a ILockBytes (unmanaged byte array) And then create a IStorage using the byte array as a backing store
                    CreateILockBytesOnHGlobal(IntPtr.Zero, True, ILockBytes)
                    iStorage2 = StgCreateDocfileOnILockBytes(ILockBytes, &H1012, 0)

                    ' copy the returned IStorage into the New IStorage
                    IStorage.CopyTo(0, Nothing, IntPtr.Zero, iStorage2)
                    ILockBytes.Flush()
                    iStorage2.Commit(0)

                    ' get the STATSTG of the ILockBytes to determine how many bytes were written to it
                    iLockBytesStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                    ILockBytes.Stat(iLockBytesStat, 1)
                    Dim iLockBytesSize As Integer = iLockBytesStat.cbSize

                    ' read the data from the ILockBytes (unmanaged byte array) into a managed byte array
                    Dim iLockBytesContent(iLockBytesSize - 1) As Byte
                    ILockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, Nothing)

                    ' wrapped the managed byte array into a memory stream And return it
                    Return New MemoryStream(iLockBytesContent)
                Finally
                    ' release all unmanaged objects
                    Marshal.ReleaseComObject(iStorage2)
                    Marshal.ReleaseComObject(ILockBytes)
                    Marshal.ReleaseComObject(IStorage)
                End Try

            Case TYMED.TYMED_ISTREAM
                ' to handle a IStream it needs to be read into a managed byte And
                ' returned as a MemoryStream

                Dim IStream As IStream = Nothing
                Dim iStreamStat As System.Runtime.InteropServices.ComTypes.STATSTG
                Try
                    ' marshal the returned pointer to a IStream object
                    IStream = CType(Marshal.GetObjectForIUnknown(medium.unionmember), IStream)
                    Marshal.Release(medium.unionmember)

                    ' get the STATSTG of the IStream to determine how many bytes are in it
                    iStreamStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                    IStream.Stat(iStreamStat, 0)
                    Dim iStreamSize As Integer = iStreamStat.cbSize

                    ' read the data from the IStream into a managed byte array
                    Dim iStreamContent(iStreamSize - 1) As Byte
                    IStream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero)

                    ' wrapped the managed byte array into a memory stream And return it
                    Return New MemoryStream(iStreamContent)
                Finally
                    ' release all unmanaged objects
                    Marshal.ReleaseComObject(IStream)
                End Try

            Case TYMED.TYMED_HGLOBAL
                ' to handle a HGlobal the exisitng "GetDataFromHGLOBAL" method Is invoked via
                ' reflection

                Return CType(Me.getDataFromHGLOBALMethod.Invoke(Me.oleUnderlyingDataObject, New Object() {DataFormats.GetDataFormat(CType(FORMATETC.cfFormat, Short)).Name, medium.unionmember}), MemoryStream)
        End Select

        Return Nothing
    End Function

    Public Function GetDataPresent(format As String) As Boolean Implements System.Windows.IDataObject.GetDataPresent
        Return Me.underlyingDataObject.GetDataPresent(format)
    End Function

    Public Function GetDataPresent(format As Type) As Boolean Implements System.Windows.IDataObject.GetDataPresent
        Return Me.underlyingDataObject.GetDataPresent(format)
    End Function

    Public Function GetDataPresent(format As String, autoConvert As Boolean) As Boolean Implements System.Windows.IDataObject.GetDataPresent
        Return Me.underlyingDataObject.GetDataPresent(format, autoConvert)
    End Function

    Public Function GetFormats() As String() Implements System.Windows.IDataObject.GetFormats
        Return Me.underlyingDataObject.GetFormats()
    End Function

    Public Function GetFormats(autoConvert As Boolean) As String() Implements System.Windows.IDataObject.GetFormats
        Return Me.underlyingDataObject.GetFormats(autoConvert)
    End Function

End Class

' NEEDED FOR 7Z ARCHIVE CREATION -- TEMP DISABLED
'' Implement IArchiveUpdateCallback to specify which files to compress
'Imports System.IO
'Imports System.Runtime.InteropServices
'Imports Laila.Shell.Interop.Properties
'Imports Laila.Shell.Interop.SevenZip

'Namespace SevenZip
'    Public Class ArchiveUpdateCallback
'        Implements IArchiveUpdateCallback

'        Private files As List(Of String)
'        Private _existingFileCount As UInt32

'        ' Constructor that accepts a list of files to be compressed
'        Public Sub New(fileList As List(Of String), existingFileCount As UInt32)
'            files = fileList
'            _existingFileCount = existingFileCount
'        End Sub

'        ' This method provides file properties such as the path and size
'        Public Function GetProperty(index As UInt32, propID As ITEMPROPID, ByRef value As PROPVARIANT) As Integer Implements IArchiveUpdateCallback.GetProperty
'            Dim filePath As String = files(CInt(index) - _existingFileCount)

'            Select Case propID
'                Case ITEMPROPID.kpidPath
'                    value = New PROPVARIANT()
'                    Dim fileInfo As New FileInfo(filePath)
'                    value.SetValue(fileInfo.Name)
'                Case ITEMPROPID.kpidSize
'                    Dim fileInfo As New FileInfo(filePath)
'                    value = New PROPVARIANT()
'                    value.SetValue(CType(fileInfo.Length, ULong))
'                Case ITEMPROPID.kpidAttrib
'                    Dim fileInfo As New FileInfo(filePath)
'                    value = New PROPVARIANT()
'                    value.SetValue(CType(fileInfo.Attributes, UInt32))
'                Case ITEMPROPID.kpidMTime
'                    Dim fileInfo As New FileInfo(filePath)
'                    value = New PROPVARIANT()
'                    value.SetValue(fileInfo.LastWriteTimeUtc)
'                Case ITEMPROPID.kpidIsDir
'                    value = New PROPVARIANT()
'                    value.SetValue(False)
'            End Select

'            Return 0 ' S_OK
'        End Function

'        ' Provides the stream of the file to be compressed
'        Public Function GetStream(index As UInt32, ByRef outStream As ISequentialInStream) As Integer Implements IArchiveUpdateCallback.GetStream
'            Dim filePath As String = files(CInt(index) - _existingFileCount)
'            If File.Exists(filePath) Then
'                outStream = New SuquentialInStream(filePath)
'                Return 0 ' S_OK
'            Else
'                outStream = Nothing
'                Return 1 ' File not found (S_FALSE or error code)
'            End If
'        End Function

'        ' Other methods of IArchiveUpdateCallback can be left unimplemented if not needed
'        Public Function GetUpdateItemInfo(index As UInt32, ByRef newData As Integer, ByRef newProperties As Integer, ByRef indexInArchive As UInt32) As Integer Implements IArchiveUpdateCallback.GetUpdateItemInfo
'            If index < _existingFileCount Then
'                newData = 0 ' Always send new data
'                newProperties = 0 ' Always send new properties
'                indexInArchive = index
'            Else
'                newData = 1 ' Always send new data
'                newProperties = 1 ' Always send new properties
'                indexInArchive = UInt32.MaxValue ' New item
'            End If
'            Return 0 ' S_OK
'        End Function

'        Public Function SetOperationResult(operationResult As Integer) As Integer Implements IArchiveUpdateCallback.SetOperationResult
'            Return 0 ' S_OK
'        End Function

'        Public Function SetTotal(bytes As ULong) As Integer Implements IArchiveUpdateCallback.SetTotal
'            Debug.WriteLine("SetTotal  bytes=" & bytes)
'        End Function

'        Public Function SetCompleted(bytes As IntPtr) As Integer Implements IArchiveUpdateCallback.SetCompleted
'            Dim i As UInt64 = Marshal.ReadInt64(bytes)
'            Debug.WriteLine("SetCompleted  bytes=" & i)
'        End Function

'        Public Function EnumProperties(enumerator As IntPtr) As Long Implements IArchiveUpdateCallback.EnumProperties
'            Return -1
'        End Function
'    End Class
'End Namespace

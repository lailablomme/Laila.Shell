'Case "laila.shell.createzip"
'' create new zip
'Dim outArchive As IOutArchive
'Dim h As HRESULT = SevenZip.Functions.CreateObject(Guids.CLSID_OutArchiveZIP, GetType(IOutArchive).GUID, outArchive)
'Dim propKeys() As String = New String() {"m", "x", "mt"}
'Dim propCompressionMethod As PROPVARIANT
'PROPVARIANT.SetValue(propCompressionMethod, "Deflate")
'Dim propCompressionLevel As PROPVARIANT
'PROPVARIANT.SetValue(propCompressionLevel, Convert.ToUInt32(5))
'Dim propThreadCount As PROPVARIANT
'PROPVARIANT.SetValue(propThreadCount, Convert.ToUInt32(1))
'Dim propValues() As PROPVARIANT = New PROPVARIANT() {
'   propCompressionMethod, propCompressionLevel, propThreadCount
'}
'Dim vptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of PROPVARIANT) * propValues.Count)
'For i = 0 To propValues.Count - 1
'Marshal.StructureToPtr(propValues(i), IntPtr.Add(vptr, Marshal.SizeOf(Of PROPVARIANT) * i), False)
'Next
'h = CType(outArchive, ISetProperties).SetProperties _
'    (propKeys, vptr, Convert.ToUInt32(propKeys.Length))
'Dim outputStream As New FileOutStream("c:\map\test3.zip")
'Dim updateCallback As New ArchiveUpdateCallback(New List(Of String) From {"c:\map\file1.txt"}, 0)
'outArchive.UpdateItems(outputStream, CUInt(1), updateCallback)
'outputStream.Close()

'' add to existing zip
'Dim inArchive As IInArchive
'h = SevenZip.Functions.CreateObject(Guids.CLSID_InArchiveZIP, GetType(IInArchive).GUID, inArchive)
'Dim inStream As InStream = New InStream("c:\map\test3.zip")
'Dim maxptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of UInt64))
'Marshal.WriteInt64(maxptr, -1)
'Dim oac As ArchiveOpenCallback = New ArchiveOpenCallback()
'h = inArchive.Open(inStream, maxptr, oac)
'Dim filecount As UInt32 = inArchive.GetNumberOfItems()
'outArchive = inArchive
'outputStream = New FileOutStream("c:\map\test4.zip")
'updateCallback = New ArchiveUpdateCallback(New List(Of String) From {"c:\map\file2.txt"}, filecount)
'outArchive.UpdateItems(outputStream, CUInt(1) + filecount, updateCallback)
'outputStream.Close()

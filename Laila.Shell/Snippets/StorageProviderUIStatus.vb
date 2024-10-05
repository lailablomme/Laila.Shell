'Case "System.StorageProviderUIStatus"
'Dim ptr As IntPtr, propertyStore As IPropertyStore, persistSerializedPropStorage As IPersistSerializedPropStorage
'Try
'Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
'propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
'persistSerializedPropStorage = propertyStore
'persistSerializedPropStorage.SetFlags(0)
'persistSerializedPropStorage.SetPropertyStorage(result.union.bstrblobVal.pData, result.union.bstrblobVal.cbSize)
'result.Dispose()
'Dim count As Integer, c2 As Integer, pk As PROPERTYKEY, name As String, rv As PROPVARIANT
'propertyStore.GetCount(count)
'For i = 0 To count - 1
'propertyStore.GetAt(i, pk)
'Functions.PSGetNameFromPropertyKey(pk, name)
'propertyStore.GetValue(pk, rv)
'If Not name = "System.StorageProviderCustomStates" Then
'Debug.WriteLine(name & " = " & getValue(rv))
'Else
'Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
'propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
'persistSerializedPropStorage = propertyStore
'persistSerializedPropStorage.SetFlags(0)
'persistSerializedPropStorage.SetPropertyStorage(rv.union.bstrblobVal.pData, rv.union.bstrblobVal.cbSize)
'propertyStore.GetCount(count)
'For ij = 0 To count - 1
'propertyStore.GetAt(ij, pk)
'Functions.PSGetNameFromPropertyKey(pk, name)
'propertyStore.GetValue(pk, rv)
'Debug.WriteLine(name & " = " & getValue(rv))
'Next
'Exit For
'End If
'Next
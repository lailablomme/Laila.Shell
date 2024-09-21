Imports System.Windows.Forms

Public Enum VARTYPE
    VT_EMPTY = 0    ' A Property With a type indicator Of VT_EMPTY has no data associated With it; that Is, the size Of the value Is zero.
    VT_NULL = 1 ' This Is Like a pointer To NULL.
    VT_I1 = 16  ' 1-Byte signed Integer.
    VT_UI1 = 17  ' 1-Byte unsigned Integer.
    VT_I2 = 2   ' Two bytes representing a 2-Byte signed Integer value.
    VT_UI2 = 18 ' 2-Byte unsigned Integer.
    VT_I4 = 3    ' 4-Byte signed Integer value.
    VT_UI4 = 19  ' 4-Byte unsigned Integer.
    VT_INT = 22  ' 4-Byte signed Integer value (equivalent To VT_I4).
    VT_UINT = 23     '	4-Byte unsigned Integer (equivalent To VT_UI4).
    VT_I8 = 20   ' 8-Byte signed Integer.
    VT_UI8 = 21  ' 8-Byte unsigned Integer.
    VT_R4 = 4    ' 32-bit IEEE floating point value.
    VT_R8 = 5    ' 64-bit IEEE floating point value.
    VT_BOOL = 11     '  (bool In earlier designs)	Boolean value, a WORD that contains 0 (False) Or -1 (True).
    VT_ERROR = 10    ' A DWORD that contains a status code.
    VT_CY = 6    ' 8-Byte two's complement integer (scaled by 10,000). This type is commonly used for currency amounts.
    VT_DATE = 7  ' A 64-bit floating point number representing the number Of days (Not seconds) since December 31, 1899. For example, January 1, 1900, Is 2.0, January 2, 1900, Is 3.0, And so On). This Is stored In the same representation As VT_R8.
    VT_FILETIME = 64 ' 64-bit FILETIME Structure As defined by Win32. It Is recommended that all times be stored In Universal Coordinate Time (UTC).
    VT_CLSID = 72    ' Pointer To a Class identifier (CLSID) (Or other globally unique identifier (GUID)).
    VT_CF = 71   ' Pointer To a CLIPDATA Structure, described above.
    VT_BSTR = 8  ' Pointer To a null-terminated Unicode String. The String Is immediately preceded by a DWORD representing the Byte count, but bstrVal points past this DWORD To the first character Of the String. BSTRs must be allocated And freed Using the Automation SysAllocString And SysFreeString calls.
    VT_BSTR_BLOB = &HFFF ' bstrblobVal	For system use only.
    VT_BLOB = 65     ' DWORD count Of bytes, followed by that many bytes Of data. The Byte count does Not include the four bytes For the length Of the count itself; an empty blob member would have a count Of zero, followed by zero bytes. This Is similar To the value VT_BSTR, but does Not guarantee a null Byte at the End Of the data.
    VT_BLOBOBJECT = 70   ' A blob member that contains a serialized Object In the same representation that would appear In VT_STREAMED_OBJECT. That Is, a DWORD Byte count (where the Byte count does Not include the size Of itself) which Is In the format Of a Class identifier followed by initialization data For that Class.
    ' The only significant difference between VT_BLOB_OBJECT And VT_STREAMED_OBJECT Is that the former does Not have the system-level storage overhead that the latter would have, And Is therefore more suitable For scenarios involving numbers Of small objects.
    VT_LPSTR = 30   ' A pointer To a null-terminated ANSI String In the system Default code page.
    VT_LPWSTR = 31   ' A pointer To a null-terminated Unicode String In the user Default locale.
    VT_UNKNOWN = 13  '	New.
    VT_DISPATCH = 9  ' New.
    VT_STREAM = 66   '	A pointer To an IStream Interface that represents a stream which Is a sibling To the "Contents" stream.
    VT_STREAMED_OBJECT = 68 ' As In VT_STREAM, but indicates that the stream contains a serialized Object, which Is a CLSID followed by initialization data For the Class. The stream Is a sibling To the "Contents" stream that contains the Property Set.
    VT_STORAGE = 67  ' A pointer To an IStorage Interface, representing a storage Object that Is a sibling To the "Contents" stream.
    VT_STORED_OBJECT = 69 ' As In VT_STORAGE, but indicates that the designated IStorage contains a loadable Object.
    VT_VERSIONED_STREAM = 73 ' A stream With a GUID version.
    VT_DECIMAL = 14  ' A Decimal Structure.
    VT_VECTOR = &H1000   ' If the type indicator Is combined With VT_VECTOR by Using an Or Operator, the value Is one Of the counted array values. This creates a DWORD count Of elements, followed by a pointer To the specified repetitions Of the value.
    'For example, a type indicator of VT_LPSTR|VT_VECTOR has a DWORD element count, followed by a pointer To an array Of LPSTR elements.

    'VT_VECTOR can be combined by an Or Operator With the following types: VT_I1, VT_UI1, VT_I2, VT_UI2, VT_BOOL, VT_I4, VT_UI4, VT_R4, VT_R8, VT_ERROR, VT_I8, VT_UI8, VT_CY, VT_DATE, VT_FILETIME, VT_CLSID, VT_CF, VT_BSTR, VT_LPSTR, VT_LPWSTR, And VT_VARIANT. VT_VECTOR can also be combined by an Or operation With VT_BSTR_BLOB, however it Is For system use only.

    VT_ARRAY = &H2000    ' If the type indicator Is combined With VT_ARRAY by an Or Operator, the value Is a pointer To a SAFEARRAY. VT_ARRAY can use the Or With the following data types VT_I1, VT_UI1, VT_I2, VT_UI2, VT_I4, VT_UI4, VT_INT, VT_UINT, VT_R4, VT_R8, VT_BOOL, VT_DECIMAL, VT_ERROR, VT_CY, VT_DATE, VT_BSTR, VT_DISPATCH, VT_UNKNOWN, And VT_VARIANT. VT_ARRAY cannot use Or With VT_VECTOR.
    VT_BYREF = &H4000    ' If the type indicator Is combined With VT_BYREF by an Or Operator, the value Is a reference. Reference types are interpreted As a reference To data, similar To the reference type In C++ (For example, "int&").
    ' VT_BYREF can use Or With the following types: VT_I1, VT_UI1, VT_I2, VT_UI2, VT_I4, VT_UI4, VT_INT, VT_UINT, VT_R4, VT_R8, VT_BOOL, VT_DECIMAL, VT_ERROR, VT_CY, VT_DATE, VT_BSTR, VT_UNKNOWN, VT_DISPATCH, VT_ARRAY, And VT_VARIANT.

    VT_VARIANT = 12  ' A DWORD type indicator followed by the corresponding value. VT_VARIANT can be used only With VT_VECTOR Or VT_BYREF.
    VT_TYPEMASK = &HFFF
End Enum

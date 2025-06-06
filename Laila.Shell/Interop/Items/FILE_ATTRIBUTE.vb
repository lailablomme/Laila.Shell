﻿Namespace Interop.Items
    <Flags()>
    Public Enum FILE_ATTRIBUTE As UInteger
        FILE_ATTRIBUTE_READONLY = &H1
        FILE_ATTRIBUTE_HIDDEN = &H2
        FILE_ATTRIBUTE_SYSTEM = &H4
        FILE_ATTRIBUTE_DIRECTORY = &H10
        FILE_ATTRIBUTE_ARCHIVE = &H20
        FILE_ATTRIBUTE_DEVICE = &H40
        FILE_ATTRIBUTE_NORMAL = &H80
        FILE_ATTRIBUTE_TEMPORARY = &H100
        FILE_ATTRIBUTE_SPARSE_FILE = &H200
        FILE_ATTRIBUTE_REPARSE_POINT = &H400
        FILE_ATTRIBUTE_COMPRESSED = &H800
        FILE_ATTRIBUTE_OFFLINE = &H1000
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
        FILE_ATTRIBUTE_ENCRYPTED = &H4000
        FILE_ATTRIBUTE_INTEGRITY_STREAM = &H8000
        FILE_ATTRIBUTE_VIRTUAL = &H10000
        FILE_ATTRIBUTE_NO_SCRUB_DATA = &H20000
        FILE_ATTRIBUTE_EA = &H40000
        FILE_ATTRIBUTE_PINNED = &H80000
        FILE_ATTRIBUTE_UNPINNED = &H100000
        FILE_ATTRIBUTE_RECALL_ON_OPEN = &H40000
        FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = &H400000

        ' Special cases for invalid or no attributes
        INVALID_FILE_ATTRIBUTES = &HFFFFFFFFUI
    End Enum
End Namespace

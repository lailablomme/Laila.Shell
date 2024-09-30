<Flags()>
Public Enum SHCONTF As UInt32
    EMPTY = 0                      ' used to zero a SHCONTF variable
    FOLDERS = &H20                 ' only want folders enumerated (FOLDER)
    NONFOLDERS = &H40              ' include non folders
    INCLUDEHIDDEN = &H80           ' show items normally hidden
    INIT_ON_FIRST_NEXT = &H100     ' allow EnumObject() to return before validating enum
    NETPRINTERSRCH = &H200         ' hint that client is looking for printers
    SHAREABLE = &H400              ' hint that client is looking sharable resources (remote shares)
    STORAGE = &H800                ' include all items with accessible storage and their ancestors
    NAVIGATION_ENUM = &H1000
    FASTITEMS = &H2000
    FLATLIST = &H4000
    ENABLE_ASYNC = &H8000
    INCLUDESUPERHIDDEN = &H10000
End Enum
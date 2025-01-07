Public Enum HRESULT
    ' Success Codes
    S_OK = &H0                         ' Operation successful
    S_FALSE = &H1                      ' Operation successful, with a false result

    ' Generic Failure Codes
    E_UNEXPECTED = &H8000FFFF          ' Unexpected failure
    E_NOTIMPL = &H80004001             ' Not implemented
    E_OUTOFMEMORY = &H8007000E         ' Ran out of memory
    E_INVALIDARG = &H80070057          ' One or more arguments are invalid
    E_NOINTERFACE = &H80004002         ' No such interface supported
    E_POINTER = &H80004003             ' Invalid pointer
    E_HANDLE = &H80070006              ' Invalid handle
    E_ABORT = &H80004004               ' Operation aborted
    E_FAIL = &H80004005                ' Unspecified failure
    E_ACCESSDENIED = &H80070005        ' General access denied error
    E_NOT_SUPPORTED = &H80070032       ' Function not supported
    E_ALREADY_INITIALIZED = &H800704DF ' Already initialized
    E_OPERATION_ABORTED = &H800703E3   ' Operation aborted
    E_PENDING = &H8000000A             ' The data necessary to complete this operation is not yet available
    E_RETRY = &H80004005               ' Retry the operation
    E_IO_PENDING = &H800703E5          ' Overlapped I/O operation is in progress
    E_TIMEOUT = &H800705B4             ' The wait operation timed out
    E_ILLEGAL_METHOD_CALL = &H8000000E ' Illegal method call
    E_UNAUTHORIZED_ACCESS = &H80070005 ' Unauthorized access

    ' Windows Specific Errors
    ERROR_FILE_NOT_FOUND2 = &H2
    ERROR_FILE_NOT_FOUND = &H80070002  ' File not found
    ERROR_PATH_NOT_FOUND = &H80070003  ' Path not found
    ERROR_ACCESS_DENIED = &H80070005   ' Access denied
    ERROR_INVALID_PARAMETER = &H80070057 ' Invalid parameter
    ERROR_NOT_SUPPORTED = &H80070032   ' The request is not supported
    ERROR_DISK_FULL = &H80070070       ' The disk is full
    ERROR_INVALID_DATA = &H8007000D    ' The data is invalid
    ERROR_BUSY = &H800700AA            ' The requested resource is in use
    ERROR_INVALID_NAME = &H8007007B    ' The filename, directory name, or volume label syntax is incorrect

    ' COM Errors
    CO_E_INIT_TLS = &H80004006         ' Failed to initialize thread local storage
    CO_E_INIT_SHARED_ALLOCATOR = &H80004007 ' Failed to initialize shared allocator
    CO_E_INIT_MEMORY_ALLOCATOR = &H80004008 ' Failed to initialize memory allocator
    CO_E_INIT_CLASS_CACHE = &H80004009 ' Failed to initialize class cache
    CO_E_INIT_RPC_CHANNEL = &H8000400A ' Failed to initialize RPC channel
    CO_E_INIT_TLS_SET_CHANNEL_CONTROL = &H8000400B ' Failed to set thread local storage channel control
    CO_E_INIT_TLS_CHANNEL_CONTROL = &H8000400C ' Failed to initialize thread local storage channel control
    CO_E_CLASS_CREATE_FAILED = &H80040154 ' Class creation failed
    CO_E_ALREADY_INITIALIZED = &H800401F5 ' Object is already initialized
    CO_E_OBJ_NOT_CONNECTED = &H800401FD ' Object is not connected to the server
    CO_E_OBJ_IS_REGULAR_OBJECT = &H800401FE ' Object is a regular object, not a proxy
    CO_E_SCM_ERROR = &H800401F0        ' Failure in the SCM
    CO_E_THREADPOOL_STOPPING = &H80040200 ' Threadpool is stopping
    CO_E_WRONG_SERVER_IDENTITY = &H80040203 ' Wrong server identity

    ' Shell Errors
    E_FILE_NOT_FOUND = &H80070002      ' File not found (same as ERROR_FILE_NOT_FOUND)
    E_PATH_NOT_FOUND = &H80070003      ' Path not found (same as ERROR_PATH_NOT_FOUND)
    E_ACCESS_DENIED = &H80070005       ' Access denied (same as ERROR_ACCESS_DENIED)
    E_INVALID_NAMESPACE = &H80041010   ' Invalid namespace
    E_ITEM_NOT_FOUND = &H80070490      ' Item not found
    E_NO_MATCH = &H800704A0           ' No match found
    E_CANNOT_MAKE = &H80070052        ' Cannot create directory or file
    E_NOT_A_DIRECTORY = &H800700B3    ' Not a directory
    E_DIRECTORY_NOT_EMPTY = &H80070091 ' Directory not empty
    E_VOLUME_NOT_MOUNTED = &H8007006E ' Volume not mounted
    E_UNKNOWN_PROPERTY = &H80070492   ' Unknown property
    E_REQUIRES_INTERACTIVE_WINDOWSTATION = &H80070568 ' Requires interactive window station
    E_BAD_FORMAT = &H8007000B          ' Bad format
    E_NOT_READY = &H80070015           ' Device not ready
    E_MEDIA_CHANGED = &H800704B0       ' Media changed

    ' Network Errors
    INET_E_DOWNLOAD_FAILURE = &H800C0008 ' Download failed
    INET_E_RESOURCE_NOT_FOUND = &H800C0005 ' Resource not found
    INET_E_DATA_NOT_AVAILABLE = &H800C0007 ' Data not available
    INET_E_CONNECTION_TIMEOUT = &H800C000B ' Connection timeout
    INET_E_SECURITY_PROBLEM = &H800C000E ' Security problem
    INET_E_CANNOT_CONNECT = &H800C000D   ' Cannot connect to the server
    INET_E_INVALID_CERTIFICATE = &H800C0019 ' Invalid certificate
    INET_E_AUTHENTICATION_REQUIRED = &H800C001B ' Authentication required
    INET_E_REDIRECT_FAILED = &H800C0016  ' Redirection failed
    INET_E_INVALID_URL = &H800C0002      ' Invalid URL
    ERROR_NETWORK_UNREACHABLE = &H800704CF ' The network location cannot be reached
    ERROR_HOST_UNREACHABLE = &H800704D0  ' Host unreachable
    ERROR_PROTOCOL_UNREACHABLE = &H800704D1 ' Protocol unreachable
    ERROR_CONNECTION_ABORTED = &H800704D4 ' Connection aborted
    ERROR_CONNECTION_REFUSED = &H800704D5 ' Connection refused
    ERROR_NAME_NOT_RESOLVED = &H80072AF9 ' DNS name cannot be resolved
    ERROR_INTERNET_DISCONNECTED = &H80072EFD ' Internet connection has been lost
    ERROR_NO_NET_OR_BAD_PATH = &H80070035 ' Network path not found

    ' File and Disk Errors
    ERROR_FILE_EXISTS = &H80070050           ' File already exists
    ERROR_WRITE_PROTECT = &H80070013        ' The media is write-protected
    ERROR_BAD_UNIT = &H80070014             ' The device is not ready
    ERROR_NOT_READY = &H80070015            ' The device is not ready
    ERROR_BAD_COMMAND = &H80070016          ' The device does not recognize the command
    ERROR_CRC = &H80070017                  ' Data error (cyclic redundancy check)
    ERROR_SEEK = &H80070019                ' The drive cannot locate a specific area or track on the disk
    ERROR_WRITE_FAULT = &H8007001D         ' The system cannot write to the specified device
    ERROR_READ_FAULT = &H8007001E          ' The system cannot read from the specified device
    ERROR_GEN_FAILURE = &H8007001F         ' A device attached to the system is not functioning
    ERROR_LOCK_VIOLATION = &H80070021      ' The process cannot access the file because another process has locked a portion of the file
    ERROR_SHARING_VIOLATION = &H80070020   ' The process cannot access the file because it is being used by another process
    ERROR_NOT_DOS_DISK = &H80070026        ' The specified disk or diskette cannot be accessed
    ERROR_SECTOR_NOT_FOUND = &H80070027    ' The drive cannot find the sector requested
    ERROR_OUT_OF_PAPER = &H80070028        ' The printer is out of paper
    ERROR_WRITE_PROTECT_MEDIA = &H80070029 ' The media is write protected
    ERROR_DRIVE_LOCKED = &H8007002C        ' The disk is locked
    ERROR_WRONG_DISK = &H80070022         ' The wrong diskette is in the drive
    ERROR_SHARING_BUFFER_EXCEEDED = &H80070024 ' Too many files opened for sharing
    ERROR_FILE_CORRUPT = &H80070570       ' The file or directory is corrupted and unreadable
    ERROR_DISK_CORRUPT = &H80070571       ' The disk structure is corrupted and unreadable
    ERROR_DISK_OPERATION_FAILED = &H800703F2 ' The disk operation failed
    ERROR_DIR_NOT_EMPTY = &H80070091      ' The directory is not empty
    ERROR_EAS_DIDNT_FIT = &H8007007A      ' The extended attributes did not fit in the buffer
    ERROR_NO_MORE_FILES = &H80070012      ' There are no more files
    ERROR_FILE_TOO_LARGE = &H800700DF     ' The file size exceeds the limit allowed
    ERROR_CANNOT_MAKE = &H80070052        ' The directory or file cannot be created
    ERROR_CURRENT_DIRECTORY = &H80070094 ' The directory cannot be removed
    ERROR_DISK_RESET_FAILED = &H800703F3 ' The disk reset operation failed
    ERROR_NO_MEDIA_IN_DRIVE = &H800703F0  ' No media in drive
    ERROR_MEDIA_CHANGED = &H800704B0      ' The media has been changed
    ERROR_UNRECOGNIZED_MEDIA = &H800704C4 ' Unrecognized media

    RPC_E_WRONG_THREAD = &H8001010E
End Enum
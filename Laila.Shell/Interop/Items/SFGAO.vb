Namespace Interop.Items
    <Flags()>
    Public Enum SFGAO
        CANCOPY = &H1                    ' Objects can be copied    
        CANMOVE = &H2                    ' Objects can be moved     
        CANLINK = &H4                    ' Objects can be linked    
        STORAGE = &H8                    ' supports BindToObject(IID_IStorage)
        CANRENAME = &H10                 ' Objects can be renamed
        CANDELETE = &H20                 ' Objects can be deleted
        HASPROPSHEET = &H40              ' Objects have property sheets
        DROPTARGET = &H100               ' Objects are drop target
        CAPABILITYMASK = &H177           ' This flag is a mask for the capability flags.
        ENCRYPTED = &H2000               ' object is encrypted (use alt color)
        ISSLOW = &H4000                  ' 'slow' object
        GHOSTED = &H8000                 ' ghosted icon
        LINK = &H10000                   ' Shortcut (link)
        SHARE = &H20000                  ' shared
        RDONLY = &H40000               ' read-only
        HIDDEN = &H80000                 ' hidden object
        DISPLAYATTRMASK = &HFC000        ' This flag is a mask for the display attributes.
        FILESYSANCESTOR = &H10000000     ' may contain children with FILESYSTEM
        FOLDER = &H20000000              ' support BindToObject(IID_IShellFolder)
        FILESYSTEM = &H40000000          ' is a win32 file system object (file/folder/root)
        HASSUBFOLDER = &H80000000        ' may contain children with FOLDER
        CONTENTSMASK = &H80000000        ' This flag is a mask for the contents attributes.
        VALIDATE = &H1000000             ' invalidate cached information
        REMOVABLE = &H2000000            ' is this removeable media?
        COMPRESSED = &H4000000           ' Object is compressed (use alt color)
        BROWSABLE = &H8000000            ' supports IShellFolder but only implements CreateViewObject() (non-folder view)
        NONENUMERATED = &H100000         ' is a non-enumerated object
        NEWCONTENT = &H200000            ' should show bold in explorer tree
        CANMONIKER = &H400000            ' defunct
        HASSTORAGE = &H400000            ' defunct
        STREAM = &H400000                ' supports BindToObject(IID_IStream)
        STORAGEANCESTOR = &H800000       ' may contain children with STORAGE or STREAM
        STORAGECAPMASK = &H70C50008      ' for determining storage capabilities ie for open/save semantics
    End Enum
End Namespace

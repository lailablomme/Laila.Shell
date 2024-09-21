<Flags()>
Public Enum SHGDN
    ''' <summary>
    ''' When Not combined with another flag, return the parent-relative name that identifies the item, suitable for displaying to the user. This 
    ''' name often does Not include extra information such As the file name extension And does Not need To be unique. This name might include 
    ''' information that identifies the folder that contains the item. For instance, this flag could cause IShellFolder:GetDisplayNameOf to
    ''' Return the String "username (on Machine)" For a particular user's folder.
    ''' </summary>
    NORMAL = 0

    ''' <summary>
    ''' The name Is relative To the folder from which the request was made. This Is the name display To the user When used In the context
    ''' Of the folder. For example, it Is used In the view And In the address bar path segment For the folder. This name should Not include
    ''' disambiguation information—For instance "username" instead Of "username (on Machine)" For a particular user's folder.
    ''' 
    ''' Use this flag In combinations With SHGDN_FORPARSING And SHGDN_FOREDITING.
    ''' </summary>
    INFOLDER = 1

    ''' <summary>
    ''' The name Is used For In-place editing When the user renames the item.
    ''' </summary>
    FOREDITING = &H1000

    ''' <summary>
    ''' The name Is displayed In an address bar combo box.
    ''' </summary>
    FORADDRESSBAR = &H4000

    ''' <summary>
    ''' The name Is used For parsing. That Is, it can be passed To IShellFolder:ParseDisplayName to recover the object's PIDL. The form this name
    ''' takes depends On the particular Object. When SHGDN_FORPARSING Is used alone, the name Is relative To the desktop. When combined With 
    ''' SHGDN_INFOLDER, the name Is relative To the folder from which the request was made.
    ''' </summary>
    FORPARSING = &H8000
End Enum

Namespace Interop.Items
    <Flags>
    Public Enum SHGFI As UInteger
        ' Retrieves the handle to the icon that represents the file.
        SHGFI_ICON = &H100

        ' Retrieves the index of the icon in the system image list.
        SHGFI_DISPLAYNAME = &H200

        ' Retrieves the string that describes the file's type.
        SHGFI_TYPENAME = &H400

        ' Retrieves the item attributes.
        SHGFI_ATTRIBUTES = &H800

        ' Retrieves the name of the file as it appears in a Windows Explorer window.
        SHGFI_ICONLOCATION = &H1000

        ' Retrieves the index of the icon in the system image list.
        SHGFI_EXETYPE = &H2000

        ' Retrieves the value that indicates whether the file is in the recycle bin.
        SHGFI_SYSICONINDEX = &H4000

        ' Modifies SHGFI_ICON, causing the function to retrieve the file's large icon.
        SHGFI_LINKOVERLAY = &H8000

        ' Retrieves the specified file's attributes.
        SHGFI_SELECTED = &H10000

        ' Modifies SHGFI_ICON, causing the function to retrieve the file's open icon.
        SHGFI_ATTR_SPECIFIED = &H20000

        ' Modifies SHGFI_ICON, causing the function to retrieve the file's small icon.
        SHGFI_LARGEICON = &H0

        ' Modifies SHGFI_ICON, causing the function to retrieve the file's small icon.
        SHGFI_SMALLICON = &H1

        ' Retrieves the file's information in the system image list.
        SHGFI_OPENICON = &H2

        ' Retrieves the file's information in the system image list.
        SHGFI_SHELLICONSIZE = &H4

        ' Retrieves the specified file's attributes.
        SHGFI_PIDL = &H8

        ' Retrieves the specified file's attributes.
        SHGFI_USEFILEATTRIBUTES = &H10

        ' Modifies SHGFI_ICON, causing the function to add the link overlay to the file's icon.
        SHGFI_ADDOVERLAYS = &H20

        ' Modifies SHGFI_ICON, causing the function to add the selected overlay to the file's icon.
        SHGFI_OVERLAYINDEX = &H40
    End Enum
End Namespace

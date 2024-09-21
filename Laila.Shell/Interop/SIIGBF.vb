<Flags>
Public Enum SIIGBF
    SIIGBF_RESIZETOFIT = 0  ' Shrink the bitmap As necessary To fit, preserving its aspect ratio.
    SIIGBF_BIGGERSIZEOK = 1 ' Passed by callers If they want To stretch the returned image themselves.
    SIIGBF_MEMORYONLY = 2   'Return the item only If it Is already In memory. Note that this only returns an already-cached icon And can fall back To a per-Class icon If an item has a per-instance icon that has Not been cached.
    SIIGBF_ICONONLY = 4 'Return only the icon, never the thumbnail.
    SIIGBF_THUMBNAILONLY = 8    'Return only the thumbnail, never the icon. Note that Not all items have thumbnails, so SIIGBF_THUMBNAILONLY will cause the method To fail In these cases.
    SIIGBF_INCACHEONLY = 16 'Allows access To the disk, but only To retrieve a cached item. This returns a cached thumbnail If it Is available. If no cached thumbnail Is available, it returns a cached per-instance icon but does Not extract a thumbnail Or icon.
    SIIGBF_CROPTOSQUARE = 32    'Introduced In Windows 8. If necessary, crop the bitmap To a square.
    SIIGBF_WIDETHUMBNAILS = 64  'Introduced In Windows 8. Stretch And crop the bitmap To a 0.7 aspect ratio.
    SIIGBF_ICONBACKGROUND = 128 'Introduced In Windows 8. If returning an icon, paint a background Using the associated app's registered background color.
    SIIGBF_SCALEUP = 256    'Introduced In Windows 8. If necessary, stretch the bitmap so that the height And width fit the given size.
End Enum

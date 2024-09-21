Public Enum PropertyAggregationType
    '/// <summary>
    '   /// The string "Multiple Values" Is displayed.
    '   /// </summary>
    [Default] = 0

    '/// <summary>
    '/// The first value in the selection Is displayed.
    '/// </summary>
    First = 1

    '/// <summary>
    '/// The sum of the selected values Is displayed. This flag Is never returned 
    '/// for data types VT_LPWSTR, VT_BOOL, And VT_FILETIME.
    '/// </summary>
    Sum = 2

    '/// <summary>
    '/// The numerical average of the selected values Is displayed. This flag 
    '/// Is never returned for data types VT_LPWSTR, VT_BOOL, And VT_FILETIME.
    '/// </summary>
    Average = 3

    '/// <summary>
    '/// The date range of the selected values Is displayed. This flag Is only 
    '/// returned for values of the VT_FILETIME data type.
    '/// </summary>
    DateRange = 4

    '/// <summary>
    '/// A concatenated string of all the values Is displayed. The order of 
    '/// individual values in the string Is undefined. The concatenated 
    '/// string omits duplicate values; if a value occurs more than once, 
    '/// it only appears a single time in the concatenated string.
    '/// </summary>
    Union = 5

    '/// <summary>
    '/// The highest of the selected values Is displayed.
    '/// </summary>
    Max = 6

    '/// <summary>
    '/// The lowest of the selected values Is displayed.
    '/// </summary>
    Min = 7
End Enum

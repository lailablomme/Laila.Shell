﻿Namespace Interop.Properties
    Public Enum PropertyGroupingRange
        '/// <summary>
        '   /// The individual values.
        '   /// </summary>
        Discrete = 0

        '/// <summary>
        '/// The static alphanumeric ranges.
        '/// </summary>
        Alphanumeric = 1

        '/// <summary>
        '/// The static size ranges.
        '/// </summary>
        Size = 2

        '/// <summary>
        '/// The dynamically-created ranges.
        '/// </summary>
        Dynamic = 3

        '/// <summary>
        '/// The month And year groups.
        '/// </summary>
        [Date] = 4

        '/// <summary>
        '/// The percent groups.
        '/// </summary>
        Percent = 5

        '/// <summary>
        '/// The enumerated groups.
        '/// </summary>
        Enumerated = 6
    End Enum
End Namespace


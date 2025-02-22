Namespace Interop.Properties
    <Flags>
    Public Enum PropertyDescriptionFormatOptions
        '/// <summary>
        '  /// The format settings specified in the property's .propdesc file.
        '  /// </summary>
        None = 0

        '/// <summary>
        '/// The value preceded with the property's display name.
        '/// </summary>
        '/// <remarks>
        '/// This flag Is ignored when the <c>hideLabelPrefix</c> attribute of the <c>labelInfo</c> element 
        '/// in the property's .propinfo file is set to true.
        '/// </remarks>
        PrefixName = &H1

        '/// <summary>
        '/// The string treated as a file name.
        '/// </summary>
        FileName = &H2

        '/// <summary>
        '/// The sizes displayed in kilobytes (KB), regardless of size. 
        '/// </summary>
        '/// <remarks>
        '/// This flag applies to properties of <c>Integer</c> types And aligns the values in the column. 
        '/// </remarks>
        AlwaysKB = &H4

        '/// <summary>
        '/// Reserved.
        '/// </summary>
        RightToLeft = &H8

        '/// <summary>
        '/// The time displayed as 'hh:mm am/pm'.
        '/// </summary>
        ShortTime = &H10

        '/// <summary>
        '/// The time displayed as 'hh:mm:ss am/pm'.
        '/// </summary>
        LongTime = &H20

        '/// <summary>
        '/// The time portion of date/time hidden.
        '/// </summary>
        HideTime = 64

        '/// <summary>
        '/// The date displayed as 'MM/DD/YY'. For example, '3/21/04'.
        '/// </summary>
        ShortDate = &H80

        '/// <summary>
        '/// The date displayed as 'DayOfWeek Month day, year'. 
        '/// For example, 'Monday, March 21, 2004'.
        '/// </summary>
        LongDate = &H100

        '/// <summary>
        '/// The date portion of date/time hidden.
        '/// </summary>
        HideDate = &H200

        '/// <summary>
        '/// The friendly date descriptions, such as "Yesterday".
        '/// </summary>
        RelativeDate = &H400

        '/// <summary>
        '/// The text displayed in a text box as a cue for the user, such as 'Enter your name'.
        '/// </summary>
        '/// <remarks>
        '/// The invitation text Is returned if formatting failed Or the value was empty. 
        '/// Invitation text Is text displayed in a text box as a cue for the user, 
        '/// Formatting can fail if the data entered 
        '/// Is Not of an expected type, such as putting alpha characters in 
        '/// a phone number field.
        '/// </remarks>
        UseEditInvitation = &H800

        '/// <summary>
        '/// This flag requires UseEditInvitation to also be specified. When the 
        '/// formatting flags are ReadOnly | UseEditInvitation And the algorithm 
        '/// would have shown invitation text, a string Is returned that indicates 
        '/// the value Is "Unknown" instead of the invitation text.
        '/// </summary>
        [ReadOnly] = &H1000

        '/// <summary>
        '/// The detection of the reading order Is Not automatic. Useful when converting 
        '/// to ANSI to omit the Unicode reading order characters.
        '/// </summary>
        NoAutoReadingOrder = &H2000

        '/// <summary>
        '/// Smart display of DateTime values
        '/// </summary>
        SmartDateTime = &H4000
    End Enum
End Namespace

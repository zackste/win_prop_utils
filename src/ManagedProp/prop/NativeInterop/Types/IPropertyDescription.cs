using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace prop.NativeInterop.Types;

using HRESULT = int;

[Guid("6f79d558-3e96-4549-a1d1-7d75d2288814")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[GeneratedComInterface]
[System.Diagnostics.CodeAnalysis.SuppressMessage(
    "ComInterfaceGenerator",
    "SYSLIB1092:The return value in the managed definition will be converted to an additional 'out' parameter at the end of the parameter list when calling the unmanaged COM method.",
    Justification = "The behavior described by the warning is understood and intentional.")]
public partial interface IPropertyDescription
{
    PROPERTYKEY GetPropertyKey();

    [return: MarshalAs(UnmanagedType.LPWStr)]
    string? GetCanonicalName();

    /*VARENUM*/
    ushort GetPropertyType();

    [return: MarshalAs(UnmanagedType.LPWStr)]
    string? GetDisplayName();

    [return: MarshalAs(UnmanagedType.LPWStr)]
    string? GetEditInvitation();

    PROPDESC_TYPE_FLAGS GetTypeFlags(PROPDESC_TYPE_FLAGS mask = PROPDESC_TYPE_FLAGS.PDTF_DEFAULT);

    PROPDESC_VIEW_FLAGS GetViewFlags();

    uint GetDefaultColumnWidth();

    PROPDESC_DISPLAYTYPE GetDisplayType();

    /*SHCOLSTATEF*/
    uint GetColumnState();

    PROPDESC_GROUPING_RANGE GetGroupingRange();

    PROPDESC_RELATIVEDESCRIPTION_TYPE GetRelativeDescriptionType();

    [PreserveSig]
    HRESULT GetRelativeDescription(
        /*REFPROPVARIANT*/ nint propvar1,
        /*REFPROPVARIANT*/ nint propvar2,
        [MarshalAs(UnmanagedType.LPWStr)] out string? ppszDesc1,
        [MarshalAs(UnmanagedType.LPWStr)] out string? ppszDesc2);

    PROPDESC_SORTDESCRIPTION GetSortDescription();

    [return: MarshalAs(UnmanagedType.LPWStr)]
    string? GetSortDescriptionLabel([MarshalAs(UnmanagedType.Bool)] bool fDescending);

    PROPDESC_AGGREGATION_TYPE GetAggregationType();

    [PreserveSig]
    HRESULT GetConditionType(
        /*PROPDESC_CONDITION_TYPE*/ out int pcontype,
        /*CONDITION_OPERATION*/ out nint popDefault);

    [PreserveSig]
    HRESULT GetEnumTypeList(
        Guid riid,
        /*IPropertyEnumTypeList */ out nint ppv);

    [PreserveSig]
    HRESULT CoerceToCanonicalValue(/*PROPVARIANT*/ ref nint ppropvar);

    [return: MarshalAs(UnmanagedType.LPWStr)]
    string? FormatForDisplay(
        /*REFPROPVARIANT*/ nint propvar,
        /*PROPDESC_FORMAT_FLAGS*/ int pdfFlags);

    [PreserveSig]
    HRESULT IsValueCanonical(/*REFPROPVARIANT*/ nint propvar);
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

public class UpdateFolderTemplateNativeHelper
{
    [ComImport,
    Guid("1F9FC1D0-C39B-4B26-817F-011967D3440E"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyDescriptionList
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetCount();
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        IPropertyDescription GetAt([In] uint iElem, [In] ref Guid riid);
    }

    [ComImport,
    Guid("6F79D558-3E96-4549-A1D1-7D75D2288814"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyDescription
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetPropertyKey(out IntPtr pkey);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetCanonicalName();
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetPropertyType(out IntPtr pvartype);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime),
        PreserveSig]
        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetDisplayName();
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetEditInvitation(out IntPtr ppszInvite);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetTypeFlags([In] IntPtr mask, out IntPtr ppdtFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetViewFlags(out IntPtr ppdvFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetDefaultColumnWidth(out uint pcxChars);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetDisplayType(out IntPtr pdisplaytype);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetColumnState(out IntPtr pcsFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetGroupingRange(out IntPtr pgr);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRelativeDescriptionType(out IntPtr prdt);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRelativeDescription([In] IntPtr propvar1, [In] IntPtr propvar2, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc1, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc2);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetSortDescription(out IntPtr psd);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetSortDescriptionLabel([In] bool fDescending, out IntPtr ppszDescription);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetAggregationType(out IntPtr paggtype);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetConditionType(out IntPtr pcontype, out IntPtr popDefault);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint GetEnumTypeList([In] ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out IntPtr ppv);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void CoerceToCanonicalValue([In, Out] IntPtr propvar);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] // Note: this method signature may be wrong, but it is not used.
        uint FormatForDisplay([In] IntPtr propvar, [In] ref IntPtr pdfFlags, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplay);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint IsValueCanonical([In] IntPtr propvar);
    }

    [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private extern static int PSEnumeratePropertyDescriptions(
        int filterOn,
        Guid riid,
        out IPropertyDescriptionList ppv);

    internal static PropertyDescriptionList PSEnumeratePropertyDescriptions(int filterOn)
    {
        int result = PSEnumeratePropertyDescriptions(
            filterOn,
            typeof(IPropertyDescriptionList).GUID,
            out IPropertyDescriptionList list);
        Marshal.ThrowExceptionForHR(result);
        return new PropertyDescriptionList(list);
    }

    internal class PropertyDescriptionList : IReadOnlyList<IPropertyDescription>
    {
        private readonly IPropertyDescriptionList _NativeList;

        public int Count
        {
            get
            {
                return (int)_NativeList.GetCount();
            }
        }

        public IPropertyDescription this[int index]
        {
            get
            {
                Guid guid = typeof(IPropertyDescription).GUID;
                return _NativeList.GetAt((uint)index, ref guid);
            }
        }

        internal PropertyDescriptionList(IPropertyDescriptionList list)
        {
            this._NativeList = list;
        }

        public IEnumerator<IPropertyDescription> GetEnumerator()
        {
            var size = Count;
            for (int i = 0; i < size; i++)
            {
                yield return this[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public static IEnumerable<string> GetAvailableColumns(int filterOn)
    {
        var propertyDescriptions = PSEnumeratePropertyDescriptions(filterOn);
        foreach (var propDesc in propertyDescriptions)
        {
            yield return propDesc.GetCanonicalName();
        }
    }
}
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ReferenceTestNet35
{
    // This file is specifically .Net 3.5 compliant implementations of the main implementation in this solution.
    // This is a minimal implementation designed to only make the needed functionality available.

    [Guid("6f79d558-3e96-4549-a1d1-7d75d2288814")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public partial interface IPropertyDescription
    {
        IntPtr GetPropertyKey();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetCanonicalName();
        ushort GetPropertyType();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetDisplayName();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetEditInvitation();
        IntPtr GetTypeFlags(IntPtr mask);
        IntPtr GetViewFlags();
        uint GetDefaultColumnWidth();
        IntPtr GetDisplayType();
        uint GetColumnState();
        IntPtr GetGroupingRange();
        IntPtr GetRelativeDescriptionType();
        [PreserveSig] int GetRelativeDescription(IntPtr propvar1, IntPtr propvar2, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc1, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc2);
        IntPtr GetSortDescription();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetSortDescriptionLabel([MarshalAs(UnmanagedType.Bool)] bool fDescending);
        IntPtr GetAggregationType();
        [PreserveSig] int GetConditionType(out int pcontype, out IntPtr popDefault);
        [PreserveSig] int GetEnumTypeList(Guid riid, out IntPtr ppv);
        [PreserveSig] int CoerceToCanonicalValue(ref IntPtr ppropvar);
        [return: MarshalAs(UnmanagedType.LPWStr)] string FormatForDisplay(IntPtr propvar, int pdfFlags);
        [PreserveSig] int IsValueCanonical(IntPtr propvar);
    }

    [Guid("1f9fc1d0-c39b-4b26-817f-011967d3440e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public partial interface IPropertyDescriptionList
    {
        uint GetCount();
        IPropertyDescription GetAt(uint iElem, ref Guid riid);
    }

    internal class UpdateFolderTemplateNativeHelper
    {
        [DllImport("Propsys.dll", CallingConvention = CallingConvention.StdCall)]
        private static extern int PSEnumeratePropertyDescriptions(
            int filterOn,
            in Guid riid,
            out IPropertyDescriptionList ppv);

        public static IEnumerable<string> GetAvailableColumns(int filterOn)
        {
            Guid guid = typeof(IPropertyDescriptionList).GUID;
            int result = PSEnumeratePropertyDescriptions(filterOn, guid, out IPropertyDescriptionList list);
            Marshal.ThrowExceptionForHR(result);
            var size = (int)list.GetCount();
            for (int i = 0; i < size; i++)
            {
                guid = typeof(IPropertyDescription).GUID;
                var item = list.GetAt((uint)i, ref guid);
                yield return item.GetCanonicalName();
            }
        }
    }
}

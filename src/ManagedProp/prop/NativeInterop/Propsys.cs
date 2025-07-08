using prop.NativeInterop.Types;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace prop.NativeInterop;

using HRESULT = int;

public static partial class Propsys
{
    [LibraryImport("Propsys.dll")]
    [UnmanagedCallConv(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static partial HRESULT PSEnumeratePropertyDescriptions(
        PROPDESC_ENUMFILTER filterOn,
        in Guid riid,
        out IPropertyDescriptionList ppv);

    public static PropertyDescriptionList PSEnumeratePropertyDescriptions(PROPDESC_ENUMFILTER filterOn)
    {
        HRESULT result = PSEnumeratePropertyDescriptions(
            filterOn,
            typeof(IPropertyDescriptionList).GUID,
            out IPropertyDescriptionList list);
        Marshal.ThrowExceptionForHR(result);
        return new PropertyDescriptionList(list);
    }
}

using prop.NativeInterop.Types;
using System.Runtime.InteropServices;

namespace prop.NativeInterop;

using HRESULT = int;

public enum DisplayNameFallback
{
    Null = 0,
    Canonical
}

public static class Extensions
{
    // https://learn.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
    private const HRESULT E_FAIL = unchecked((int)0x80004005);

    public static string? GetDisplayName(this IPropertyDescription propertyDescription, DisplayNameFallback fallback)
    {
        try
        {
            return propertyDescription.GetDisplayName();
        }
        catch (COMException come) when (come.HResult == E_FAIL)
        {
            return fallback == DisplayNameFallback.Null ? (string?)null : propertyDescription.GetCanonicalName();
        }
    }
}

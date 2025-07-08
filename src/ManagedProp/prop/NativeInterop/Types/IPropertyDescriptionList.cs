using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace prop.NativeInterop.Types;

[Guid("1f9fc1d0-c39b-4b26-817f-011967d3440e")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[GeneratedComInterface]
public partial interface IPropertyDescriptionList
{
    uint GetCount();

    IPropertyDescription GetAt(uint iElem, ref Guid riid);
}

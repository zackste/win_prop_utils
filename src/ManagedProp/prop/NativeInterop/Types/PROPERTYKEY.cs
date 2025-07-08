using System.Runtime.InteropServices;

namespace prop.NativeInterop.Types;

[StructLayout(LayoutKind.Sequential, Pack = 4)]
public readonly struct PROPERTYKEY(Guid fmtid, int pid)
{
    public readonly Guid fmtid = fmtid;
    public readonly int pid = pid;

    public override readonly string ToString() =>
        $"{fmtid:B}, {pid}";
}

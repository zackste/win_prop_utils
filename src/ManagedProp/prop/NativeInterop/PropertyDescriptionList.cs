using prop.NativeInterop.Types;
using System.Collections;

namespace prop.NativeInterop;

public record PropertyDescriptionList : IReadOnlyList<IPropertyDescription>
{
    private readonly IPropertyDescriptionList NativeList;

    public int Count => (int)NativeList.GetCount();

    public IPropertyDescription this[int index]
    {
        get
        {
            Guid guid = typeof(IPropertyDescription).GUID;
            return NativeList.GetAt((uint)index, ref guid);
        }
    }

    internal PropertyDescriptionList(IPropertyDescriptionList list)
    {
        NativeList = list;
    }

    public IEnumerator<IPropertyDescription> GetEnumerator()
    {
        var size = Count;
        for (int i = 0; i < size; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

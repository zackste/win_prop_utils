using prop.NativeInterop;
using prop.NativeInterop.Types;
using System.Runtime.InteropServices;

var start = DateTime.Now;

Console.WriteLine("Hello, World!");

PropertyDescriptionList list = Propsys.PSEnumeratePropertyDescriptions(PROPDESC_ENUMFILTER.PDEF_COLUMN);

int listSize = list.Count;
Console.WriteLine($"Count: {listSize}");

//for (int i = 0; i < 100;  i++)
foreach (var item in list)
{
    Console.WriteLine($"{item.GetCanonicalName()}\t{item.GetDisplayName(DisplayNameFallback.Null)}");
}

Console.WriteLine($"Done! {DateTime.Now - start}");
using System;

namespace ReferenceTestNet35
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var results = UpdateFolderTemplateNativeHelper.GetAvailableColumns(/*PDEF_COLUMN*/ 6);

            foreach (var column in results)
            {
                Console.WriteLine(column[0] + "\t" + column[1]);
            }
        }
    }
}

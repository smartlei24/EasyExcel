using System.Collections.Generic;
using System.Linq;

namespace EasyExcel.Extension
{
    public static class EnumerableExtension
    {
        public static bool IsNullOrEmpty<T>(this IEnumerable<T> source)
        {
            return source == null || !source.Any();
        }
    }
}
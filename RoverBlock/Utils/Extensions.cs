using System.Collections.Generic;
using System.Linq;

namespace RoverBlock.Utils
{
    public static class Extensions
    {
        public static IEnumerable<T> PadLeft<T>(this IEnumerable<T> enumerable, int length, T padWith)
        {
            var list = enumerable.ToList();

            while(list.Count < length)
            {
                list.Insert(0, padWith);
            }

            return list;
        } 
    }
}

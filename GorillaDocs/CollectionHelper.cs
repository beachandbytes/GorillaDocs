using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public static class CollectionHelper
    {
        public static T ReplaceAndReturn<T>(this IList<T> collection, T find, T replace)
        {
            int index = collection.IndexOf(find);
            find = replace;
            collection[index] = find;
            return find;
        }

        public static int RemoveAll<T>(this IList<T> coll, Func<T, bool> condition)
        {
            var itemsToRemove = coll.Where(condition).ToList();
            foreach (var itemToRemove in itemsToRemove)
                coll.Remove(itemToRemove);
            return itemsToRemove.Count;
        }

        public static T FirstOrCreateIfEmpty<T>(this IList<T> list) where T : new()
        {
            if (list.Any())
                return list.First();
            var item = new T();
            list.Add(item);
            return item;
        }
    }
}

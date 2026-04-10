using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;

namespace ERP_Management_Infrastructor.Cache
{
    public static class CacheManager
    {
        private static readonly ObjectCache cache = MemoryCache.Default;

        /// <summary>
        /// Adds an item to the cache with absolute expiration.
        /// </summary>
        public static void Add(string key, object value, int? expirationMinutes = null)
        {
            if (value == null) return;

            CacheItemPolicy policy = new CacheItemPolicy();

            if (expirationMinutes.HasValue)
            {
                policy.AbsoluteExpiration = DateTimeOffset.Now.AddMinutes(expirationMinutes.Value);
            }
            // Else, no expiration set (item stays until app shutdown or manual removal)

            cache.Set(key, value, policy);
        }

        /// <summary>
        /// Retrieves an item from the cache.
        /// </summary>
        public static T Get<T>(string key)
        {
            var cachedItem = cache.Get(key);
            if (cachedItem != null)
            {
                return (T)cachedItem;
            }

            return default(T);
        }

        /// <summary>
        /// Removes an item from the cache.
        /// </summary>
        public static void Remove(string key)
        {
            if (cache.Contains(key))
            {
                cache.Remove(key);
            }
        }

        /// <summary>
        /// Checks if an item exists in the cache.
        /// </summary>
        public static bool Exists(string key)
        {
            return cache.Contains(key);
        }

        /// <summary>
        /// Clears all items from the cache.
        /// </summary>
        public static void ClearCache()
        {
            foreach (var item in cache)
            {
                cache.Remove(item.Key);
            }
        }
    }
}

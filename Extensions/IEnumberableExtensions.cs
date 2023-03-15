using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;

namespace FT_ADDON
{
    static class IEnumberableExtensions
    {
        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, object> selector, object value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }

        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, string> selector, string value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }

        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, int> selector, int value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }

        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, bool> selector, bool value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }

        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, double> selector, double value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }

        public static bool Contains<TSource>(this IEnumerable sources, Func<TSource, float> selector, float value)
        {
            return sources.OfType<TSource>().Select(selector).Where(each => each == value).Any();
        }
    }
}

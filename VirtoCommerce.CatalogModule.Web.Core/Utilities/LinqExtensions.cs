using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace VirtoCommerce.CatalogModule.Web.Utilities
{
    public static class LinqExtensions
    {
        public static PropertyInfo GetProperty<TType, TProperty>(this Expression<Func<TType, TProperty>> propertyExpression)
        {
            var memberExpression = propertyExpression.Body as MemberExpression;
            if (memberExpression == null)
            {
                // Convert expression?
                var unaryExpression = propertyExpression.Body as UnaryExpression;
                if (unaryExpression != null)
                {
                    memberExpression = unaryExpression.Operand as MemberExpression;
                }
            }
            if (memberExpression == null) throw new ArgumentException("Illegal property selection expression.");

            var property = memberExpression.Member as PropertyInfo;
            if (property == null) throw new ArgumentException("Illegal property selection expression.");

            return property;
        }

        public static string ToCommaDelimitedString<TSource>(this IEnumerable<TSource> source)
        {
            return source.ToDelimitedString(",");
        }

        public static string ToDelimitedString<TSource>(this IEnumerable<TSource> source, string delimiter)
        {
            return source.ToDelimitedString(e => e == null ? null : e.ToString(), delimiter);
        }

        public static string ToDelimitedString<TSource>(this IEnumerable<TSource> source, Func<TSource, string> func,
            string delimiter)
        {
            var builder = new StringBuilder();
            source.ToDelimitedString(builder, func, delimiter);
            return builder.ToString();
        }

        public static void ToDelimitedString<TSource>(this IEnumerable<TSource> source, StringBuilder builder,
            Func<TSource, string> func, string delimiter)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (builder == null) throw new ArgumentNullException("builder");

            // Execute action for each item and add the delimiter.
            var hasItems = false;
            foreach (var item in source)
            {
                builder.Append(func(item));
                builder.Append(delimiter);
                hasItems = true;
            }
            if (hasItems) builder.Remove(builder.Length - delimiter.Length, delimiter.Length);
        }
    }
}
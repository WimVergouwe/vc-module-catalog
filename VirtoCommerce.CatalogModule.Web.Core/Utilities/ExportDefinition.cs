using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using VirtoCommerce.Domain.Catalog.Model;

namespace VirtoCommerce.CatalogModule.Web.Utilities
{
    public class ExportDefinition<T> : IEnumerable<ColumnExportDefinition<T>>
    {
        private readonly SortedDictionary<string, ColumnExportDefinition<T>> _columns
            = new SortedDictionary<string, ColumnExportDefinition<T>>(StringComparer.OrdinalIgnoreCase);

        public void Add(Expression<Func<T, object>> propertyExpression)
        {
            if (propertyExpression == null) throw new ArgumentNullException(nameof(propertyExpression));

            var property = propertyExpression.GetProperty();
            var column = new ColumnExportDefinition<T>(property.Name, propertyExpression.Compile());
            _columns[column.Name] = column;
        }

        public void Add(ColumnExportDefinition<T> columnExportDefinition)
        {
            _columns[columnExportDefinition.Name] = columnExportDefinition;
        }

        public IEnumerator<ColumnExportDefinition<T>> GetEnumerator()
        {
            return _columns.Select(x => x.Value).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public class ColumnExportDefinition<T>
    {
        private readonly Func<T, object> _fieldAccessor;
        public string Name { get; }

        public ColumnExportDefinition(string name, Func<T, object> fieldAccessor)
        {
            if (fieldAccessor == null) throw new ArgumentNullException(nameof(fieldAccessor));
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));

            Name = name;
            _fieldAccessor = fieldAccessor;
        }

        public string GetValue(T source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            return GetFormattedValue(_fieldAccessor(source));
        }

        private string GetFormattedValue(object value)
        {
            if (value == null) return "";
            
            var values = value as IEnumerable<PropertyValue>;
            return values == null ? GetFormattedString(value) : string.Join(",", values.Select(x => GetFormattedString(x.Value)));
            
        }

        private string GetFormattedString(object value)
        {
            var inv = CultureInfo.InvariantCulture;

            var formattable = value as IFormattable;
            return formattable?.ToString(null, inv) ?? value.ToString();
        }
    }

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
    }
}

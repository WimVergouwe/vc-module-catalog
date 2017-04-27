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
    public class ImportDefinition<T> : IEnumerable<ColumnImportDefinition<T>>
        where T : CatalogProduct
    {
        private readonly SortedDictionary<string, ColumnImportDefinition<T>> _columns = new SortedDictionary<string, ColumnImportDefinition<T>>(StringComparer.OrdinalIgnoreCase);

        public void Add(Expression<Func<T, object>> propertyExpression)
        {
            if (propertyExpression == null) throw new ArgumentNullException(nameof(propertyExpression));

            var property = propertyExpression.GetProperty();
            var column = new ColumnImportDefinition<T>(property);
            _columns[column.Name] = column;
        }

        // Use this add to map properties from a different name.
        public void Add(string name, Expression<Func<T, object>> propertyExpression)
        {
            if (propertyExpression == null) throw new ArgumentNullException(nameof(propertyExpression));
            
            var property = propertyExpression.GetProperty();
            var column = new ColumnImportDefinition<T>(name, property);
            _columns[name] = column;
        }

        public IEnumerator<ColumnImportDefinition<T>> GetEnumerator()
        {
            return _columns.Select(x => x.Value).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public class ColumnImportDefinition<T>
    {
        private readonly PropertyInfo _property;

        public string Name { get; protected set; }

        public ColumnImportDefinition(PropertyInfo property)
        {
            _property = property;
            Name = property?.Name;
        }

        public ColumnImportDefinition(string name, PropertyInfo property)
        {
            Name = name;
            _property = property;
        }

        public virtual void SetValue(T source, object value)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            
            if (string.IsNullOrEmpty((string)value))
                value = null;

            bool booleanValue;
            if (bool.TryParse((string)value, out booleanValue))
            {
                _property.SetValue(source, booleanValue);
                return;
            }

            // TODO: tryformat to _property.PropertyType
            // TODO: support lists etc.
            _property.SetValue(source, value);
        }
    }

    public class ExportDefinition<T> : IEnumerable<ColumnExportDefinition<T>>
    {
        private readonly SortedDictionary<string, ColumnExportDefinition<T>> _columns
            = new SortedDictionary<string, ColumnExportDefinition<T>>(StringComparer.OrdinalIgnoreCase);

        public void Add(Expression<Func<T, object>> propertyExpression)
        {
            if (propertyExpression == null) throw new ArgumentNullException(nameof(propertyExpression));

            var property = propertyExpression.GetProperty();
            var column = new ColumnExportDefinition<T>(property.Name, property.PropertyType, propertyExpression.Compile());
            _columns[column.Name] = column;
        }

        public void Add(string name, Expression<Func<T, object>> propertyExpression)
        {
            if (propertyExpression == null) throw new ArgumentNullException(nameof(propertyExpression));

            var property = propertyExpression.GetProperty();
            var column = new ColumnExportDefinition<T>(name, property.PropertyType, propertyExpression.Compile());
            _columns[name] = column;
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
        public Type PropertyType { get; set; }

        public ColumnExportDefinition(string name, Type propertyType, Func<T, object> fieldAccessor)
        {
            if (fieldAccessor == null) throw new ArgumentNullException(nameof(fieldAccessor));
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));

            Name = name;
            PropertyType = propertyType;
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
}

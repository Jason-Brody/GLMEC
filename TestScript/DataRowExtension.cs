﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace TestScript
{
    public static class DataRowExtension
    {
        public static T ToEntity<T>(this DataRow dr) where T : class, new()
        {
            if (dr == null)
                return null;
            T item = new T();
            var table = dr.Table;
            Dictionary<string, string> cols = new Dictionary<string, string>();
            foreach (DataColumn col in table.Columns)
            {
                cols.Add(col.ColumnName.ToLower(), col.ColumnName);
            }
            foreach (var prop in typeof(T).GetProperties())
            {
                var alias = prop.GetCustomAttribute<AliasAttribute>();
                var lowerProp = prop.Name.ToLower();
                if (alias != null)
                    lowerProp = alias.Name.ToLower();
                
                if (cols.ContainsKey(lowerProp))
                {
                    object value = Convert.ChangeType(dr[cols[lowerProp]], prop.PropertyType);
                    prop.SetValue(item, value);
                }
            }
            return item;
        }
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class AliasAttribute : Attribute
    {
        public AliasAttribute(string Name)
        {
            this.Name = Name;
        }

        public string Name { get; set; }
    }
}
